"""
baptismal/image_matching.py — Cautious Image-to-Record Matching & Admin Confirmation
==================================================================================
NEVER automatically assumes a filename's page number is authoritative.
Example: Record indicates Book 2, Page 743 while image filename says `Abelardo Galban BC_744.jpeg` ->
flagged strictly as `POSSIBLE_MATCH_NEEDS_CONFIRMATION` until Master Admin confirms.
"""

from datetime import datetime, timezone
import re
import secrets
from typing import Any, Dict, List, Optional

from .audit_logs import AuditLogger
from .authorization import AuthorizationService
from .database import DatabaseManager


def extract_page_hint_from_filename(filename: str) -> str:
    """Extracts a numeric page hint from a scan filename (e.g., 'BC_744.jpeg' -> '744', BOOK 9 Epson scans -> '4146'..'4164')."""
    from .config import BOOK9_EPSON_SCAN_MAP
    raw_fn = str(filename or "").strip()
    if raw_fn in BOOK9_EPSON_SCAN_MAP:
        return BOOK9_EPSON_SCAN_MAP[raw_fn]["page_number"]
    for k, v in BOOK9_EPSON_SCAN_MAP.items():
        if k.lower() == raw_fn.lower():
            return v["page_number"]
    stem = raw_fn.rsplit(".", 1)[0]
    m = re.search(r"(?:bc[_\-\s]*|page[_\-\s]*|p[_\-\s]*)(\d{1,5})\b", stem, flags=re.I)
    if m:
        val = m.group(1)
        return str(int(val)) if val.isdigit() else val
    m_end = re.search(r"(\d{1,5})$", stem.strip())
    if m_end:
        val = m_end.group(1)
        return str(int(val)) if val.isdigit() else val
    return ""


class ImageMatchingService:
    """Manages candidate image matching and Master Admin confirmation."""

    def __init__(self, db: DatabaseManager, authz: AuthorizationService, audit: Optional[AuditLogger] = None):
        self.db = db
        self.authz = authz
        self.audit = audit or AuditLogger(db)

    def suggest_possible_match(
        self,
        record_id: str,
        book_folder: str,
        image_filename: str,
        drive_file_id: str = "",
        notes: str = "",
    ) -> Dict[str, Any]:
        """Creates a candidate image association marked as POSSIBLE_MATCH_NEEDS_CONFIRMATION
        whenever filename page differs from record Page Number (or before admin review)."""
        rec = self.db.fetch_one(
            "SELECT record_id, child_name, book_number, page_number FROM historical_records WHERE record_id = %s",
            (record_id,),
        )
        rec_page = str((rec or {}).get("page_number") or "").strip()
        fn_hint = extract_page_hint_from_filename(image_filename)
        page_mismatch = bool(rec_page and fn_hint and rec_page != fn_hint)

        assoc_id = f"img_assoc_{secrets.token_hex(6)}"
        now_iso = datetime.now(timezone.utc).isoformat()
        status = "POSSIBLE_MATCH_NEEDS_CONFIRMATION"
        auto_note = (
            f"Page mismatch: Record Page '{rec_page}' vs Filename hint '{fn_hint}'. Master Admin confirmation required."
            if page_mismatch
            else "Awaiting Master Admin confirmation."
        )
        final_notes = f"{auto_note} {notes}".strip()

        existing = self.db.fetch_one(
            """
            SELECT association_id FROM historical_image_associations
            WHERE record_id = %s AND book_folder = %s AND image_filename = %s
            """,
            (record_id, book_folder, image_filename),
        )
        if existing:
            assoc_id = existing["association_id"]
        else:
            self.db.execute(
                """
                INSERT INTO historical_image_associations (
                    association_id, record_id, book_folder, image_filename, drive_file_id,
                    extracted_page_hint, record_page_number, match_status, confirmed_by,
                    confirmed_at, notes, created_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NULL, NULL, %s, %s)
                """,
                (
                    assoc_id,
                    record_id,
                    book_folder,
                    image_filename,
                    drive_file_id,
                    fn_hint,
                    rec_page,
                    status,
                    final_notes,
                    now_iso,
                ),
            )
        return {
            "association_id": assoc_id,
            "record_id": record_id,
            "book_folder": book_folder,
            "image_filename": image_filename,
            "record_page_number": rec_page,
            "extracted_page_hint": fn_hint,
            "page_mismatch": page_mismatch,
            "match_status": status,
            "notes": final_notes,
        }

    def confirm_image_association(
        self,
        actor_account_id: str,
        association_id: str,
        approve: bool = True,
        admin_notes: str = "",
    ) -> str:
        """Master Admin confirmation or rejection of an image-to-record association."""
        self.authz.require_master_admin(actor_account_id, action="confirm_image_association")
        assoc = self.db.fetch_one(
            "SELECT * FROM historical_image_associations WHERE association_id = %s",
            (association_id,),
        )
        if not assoc:
            raise ValueError(f"Image association '{association_id}' not found.")
        new_status = "CONFIRMED" if approve else "REJECTED"
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE historical_image_associations
            SET match_status = %s, confirmed_by = %s, confirmed_at = %s, notes = %s
            WHERE association_id = %s
            """,
            (new_status, actor_account_id, now_iso, admin_notes or assoc.get("notes", ""), association_id),
        )
        self.audit.log_event(
            actor_account_id,
            "IMAGE_MAPPING_CONFIRMED" if approve else "IMAGE_MAPPING_REJECTED",
            "SUCCESS",
            related_record_id=assoc["record_id"],
            metadata={"association_id": association_id, "image_filename": assoc["image_filename"]},
        )
        return new_status

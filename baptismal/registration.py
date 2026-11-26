"""
baptismal/registration.py — New Baptism Registration & Google Drive + Excel Sync
================================================================================
- Validates all required fields among the 27 canonical fields.
- Generates a stable unique internal Record ID (BAP-YYYY-XXXXXXXX), NEVER a person's name.
- Prevents duplicate registrations via deterministic fingerprint.
- Uploads registration images to private `NEW BAPTISM IMAGES/<year>/`.
- Safely appends approved record to `NEW BAPTISMS - <year>` in the designated Excel workbook
  after verifying Editor permission, creating a backup, and checking Drive revisions.
"""

from datetime import datetime, timezone
import hashlib
import json
import secrets
from typing import Any, Dict, List, Optional, Tuple

from .audit_logs import AuditLogger
from .authorization import AuthorizationService
from .backup_manager import BackupManager
from .config import HISTORICAL_27_COLUMNS
from .database import DatabaseManager
from .excel_manager import append_new_baptism_to_workbook_bytes, get_yearly_worksheet_name
from .google_drive import GoogleDriveManager
from .validation import validate_uploaded_file


def generate_stable_record_id(baptism_year: str) -> str:
    yr = str(baptism_year or datetime.now(timezone.utc).year).strip()
    return f"BAP-{yr}-{secrets.token_hex(4).upper()}"


def compute_registration_fingerprint(data: Dict[str, Any]) -> str:
    parts = [
        str(data.get("Name of the Child", "")).strip().lower(),
        str(data.get("Date of Baptism", "")).strip().lower(),
        str(data.get("Date of Birth", "")).strip().lower(),
        str(data.get("Name of the Father", "")).strip().lower(),
        str(data.get("Name of the Mother", "")).strip().lower(),
    ]
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()


class RegistrationService:
    """Coordinates validated new baptism registrations, Drive uploads, and Excel workbook appends."""

    REQUIRED_FIELDS = [
        "Parish",
        "Name of the Child",
        "Date of Baptism",
        "Year of Baptism",
    ]

    def __init__(
        self,
        db: DatabaseManager,
        authz: AuthorizationService,
        audit: Optional[AuditLogger] = None,
        drive: Optional[GoogleDriveManager] = None,
        backup: Optional[BackupManager] = None,
    ):
        self.db = db
        self.authz = authz
        self.audit = audit or AuditLogger(db)
        self.drive = drive
        self.backup = backup or BackupManager(db, audit=self.audit)

    def validate_registration_payload(self, data: Dict[str, Any]) -> Tuple[bool, List[str]]:
        missing = [f for f in self.REQUIRED_FIELDS if not str(data.get(f, "")).strip()]
        yr = str(data.get("Year of Baptism", "")).strip()
        if yr and (not yr.isdigit() or len(yr) != 4):
            missing.append("Year of Baptism (must be 4 digits, e.g. 2026)")
        return len(missing) == 0, missing

    def register_new_baptism(
        self,
        actor_account_id: str,
        payload: Dict[str, Any],
        uploaded_images: Optional[List[Dict[str, Any]]] = None,
        sync_to_drive_excel: bool = True,
    ) -> Dict[str, Any]:
        self.authz.require_master_admin(actor_account_id, action="register_new_baptism")

        clean_payload = {col: str(payload.get(col, "")).strip() for col in HISTORICAL_27_COLUMNS}
        ok_val, missing = self.validate_registration_payload(clean_payload)
        if not ok_val:
            raise ValueError(f"Missing or invalid required fields: {', '.join(missing)}")

        bap_year = clean_payload["Year of Baptism"]
        ws_name = get_yearly_worksheet_name(bap_year)
        fingerprint = compute_registration_fingerprint(clean_payload)

        dup = self.db.fetch_one(
            "SELECT record_id FROM new_baptism_registrations WHERE dedup_fingerprint = %s",
            (fingerprint,),
        )
        if dup:
            raise ValueError(
                f"Duplicate registration prevented: matching record already exists ({dup['record_id']})."
            )

        validated_imgs: List[Dict[str, Any]] = []
        for img_item in uploaded_images or []:
            ok_img, msg_img, safe_fn = validate_uploaded_file(
                img_item.get("bytes", b""),
                img_item.get("filename", ""),
                img_item.get("mime_type", "image/jpeg"),
            )
            if not ok_img:
                raise ValueError(f"Image validation failed for '{img_item.get('filename', '')}': {msg_img}")
            validated_imgs.append({
                "bytes": img_item["bytes"],
                "filename": safe_fn,
                "mime_type": img_item.get("mime_type") or "image/jpeg",
            })

        record_id = generate_stable_record_id(bap_year)
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            INSERT INTO new_baptism_registrations (
                record_id, baptism_year, worksheet_name, registration_status,
                upload_status, workbook_update_status, parish, child_name,
                date_of_baptism, father_name, mother_name, book_number,
                page_number, dedup_fingerprint, payload_json, created_by,
                created_at, updated_at
            ) VALUES (
                %s, %s, %s, 'REGISTERED', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
            """,
            (
                record_id,
                bap_year,
                ws_name,
                "PENDING" if validated_imgs else "NONE",
                "PENDING" if sync_to_drive_excel else "LOCAL_ONLY",
                clean_payload["Parish"],
                clean_payload["Name of the Child"],
                clean_payload["Date of Baptism"],
                clean_payload["Name of the Father"],
                clean_payload["Name of the Mother"],
                clean_payload["Book Number"],
                clean_payload["Page Number"],
                fingerprint,
                json.dumps(clean_payload, ensure_ascii=False),
                actor_account_id,
                now_iso,
                now_iso,
            ),
        )
        self.audit.log_event(
            actor_account_id,
            "RECORD_REGISTERED",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"worksheet": ws_name, "baptism_year": bap_year},
        )

        uploaded_drive_ids: List[str] = []
        workbook_status = "LOCAL_ONLY"
        if sync_to_drive_excel and self.drive is not None:
            folder_tree = self.drive.ensure_baptismal_folder_tree()
            img_folder_id = folder_tree.get(f"NEW BAPTISM IMAGES/{bap_year}") or folder_tree.get("NEW BAPTISM IMAGES", "")
            for v_img in validated_imgs:
                created_f = self.drive.upload_private_file(
                    v_img["bytes"],
                    v_img["filename"],
                    v_img["mime_type"],
                    img_folder_id,
                )
                d_fid = str(created_f["id"])
                uploaded_drive_ids.append(d_fid)
                self.db.execute(
                    """
                    INSERT INTO drive_file_mappings (
                        mapping_id, record_id, category, original_filename,
                        drive_file_id, drive_folder_id, year_folder, mime_type,
                        file_size_bytes, upload_status, uploaded_by, created_at
                    ) VALUES (%s, %s, 'NEW_BAPTISM_IMAGE', %s, %s, %s, %s, %s, %s, 'UPLOADED', %s, %s)
                    """,
                    (
                        f"drv_{secrets.token_hex(6)}",
                        record_id,
                        v_img["filename"],
                        d_fid,
                        img_folder_id,
                        bap_year,
                        v_img["mime_type"],
                        len(v_img["bytes"]),
                        actor_account_id,
                        now_iso,
                    ),
                )
                self.audit.log_event(
                    actor_account_id,
                    "IMAGE_UPLOADED",
                    "SUCCESS",
                    related_record_id=record_id,
                    metadata={"drive_file_id": d_fid, "year": bap_year},
                )

            wb_file_id = str(self.drive.config.get("workbook_file_id") or "").strip()
            if wb_file_id:
                meta = self.drive.verify_workbook_write_access(wb_file_id)
                before_rev = str(meta.get("headRevisionId") or "")
                wb_bytes = self.drive.download_file_bytes(wb_file_id)
                self.backup.save_local_workbook_backup(wb_bytes, label=f"before_{record_id}", actor_id=actor_account_id)
                bk_folder_id = folder_tree.get("BACKUPS", "")
                bk_drive_id = ""
                if bk_folder_id:
                    bk_res = self.drive.upload_private_file(
                        wb_bytes,
                        f"BACKUP_before_{record_id}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        bk_folder_id,
                    )
                    bk_drive_id = str(bk_res.get("id", ""))

                updated_wb_bytes, used_ws = append_new_baptism_to_workbook_bytes(
                    wb_bytes, clean_payload, record_id, bap_year
                )
                upd_res = self.drive.update_existing_file_bytes(
                    wb_file_id,
                    updated_wb_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    expected_revision_id=before_rev,
                )
                after_rev = str(upd_res.get("headRevisionId") or "")
                workbook_status = "SYNCED"
                self.db.execute(
                    """
                    INSERT INTO workbook_update_logs (
                        update_id, record_id, worksheet_name, drive_file_id,
                        before_revision_id, after_revision_id, backup_drive_file_id,
                        status, error_message, actor_account_id, created_at
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, 'SUCCESS', '', %s, %s)
                    """,
                    (
                        f"wb_upd_{secrets.token_hex(6)}",
                        record_id,
                        used_ws,
                        wb_file_id,
                        before_rev,
                        after_rev,
                        bk_drive_id,
                        actor_account_id,
                        now_iso,
                    ),
                )
                self.audit.log_event(
                    actor_account_id,
                    "WORKBOOK_UPDATE_SUCCESS",
                    "SUCCESS",
                    related_record_id=record_id,
                    metadata={"worksheet": used_ws, "drive_file_id": wb_file_id},
                )

        self.db.execute(
            """
            UPDATE new_baptism_registrations
            SET upload_status = %s, workbook_update_status = %s, updated_at = %s
            WHERE record_id = %s
            """,
            (
                "UPLOADED" if uploaded_drive_ids else ("NONE" if not validated_imgs else "PENDING"),
                workbook_status,
                datetime.now(timezone.utc).isoformat(),
                record_id,
            ),
        )
        return {
            "record_id": record_id,
            "worksheet_name": ws_name,
            "uploaded_drive_file_ids": uploaded_drive_ids,
            "workbook_update_status": workbook_status,
        }

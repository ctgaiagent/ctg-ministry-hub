"""
baptismal/registration.py — Stage-Then-Confirm Parish Baptism Registration & Drive Sync
=======================================================================================
Workflow:
1. When a Parish Leader Admin submits a new baptism + uploads baptismal images/files:
   - The record and uploaded files are staged in the database (`Pending Verification`,
     `workbook_update_status = 'AWAITING_ADMIN_CONFIRMATION'`).
   - It does NOT write to the Google Drive Master Excel or Drive folder yet!
2. When the Master Admin reviews the submission and clicks Confirm (`confirm_and_sync_submission_to_drive`):
   - The submission is marked `Approved`.
   - Staged images/files are uploaded to private Google Drive (`NEW BAPTISM IMAGES/<year>`).
   - The 27-column record is encoded/appended into `NEW BAPTISMS - <year>` inside
     `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` on Google Drive!
"""

import base64
from datetime import datetime, timezone
import hashlib
import json
import secrets
from typing import Any, Dict, List, Optional, Tuple

from .audit_logs import AuditLogger
from .authorization import AuthorizationError, AuthorizationService
from .backup_manager import BackupManager
from .config import (
    EDITABLE_REGISTRATION_STATUSES,
    HISTORICAL_27_COLUMNS,
    REGISTRATION_VERIFICATION_STATUSES,
    is_excluded_parish,
    normalize_official_nine_parish,
)
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
    """Coordinates staged Parish Leader Admin submissions and Master Admin confirmation & Google Drive encoding."""

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

    def _resolve_authoritative_parish(self, actor_account_id: str, requested_parish: str = "") -> str:
        acc = self.authz.get_live_account(actor_account_id)
        if not acc:
            raise AuthorizationError("Access denied: unauthenticated or inactive account.")

        if acc["role"] == "Parish Leader Admin":
            return self.authz.get_authenticated_parish(actor_account_id)

        if acc["role"] == "Master Admin":
            raw_p = str(requested_parish or "").strip()
            if is_excluded_parish(raw_p):
                raise ValueError("Hong Kong Parish is excluded from the CTG Baptismal System.")
            if not raw_p:
                raise ValueError("Parish is required.")
            try:
                return normalize_official_nine_parish(raw_p)
            except ValueError:
                return raw_p

        raise AuthorizationError("Access denied: account is not authorized to register baptisms.")

    def validate_registration_payload(self, data: Dict[str, Any]) -> Tuple[bool, List[str]]:
        missing = [f for f in self.REQUIRED_FIELDS if not str(data.get(f, "")).strip()]
        yr = str(data.get("Year of Baptism", "")).strip()
        if yr and (not yr.isdigit() or len(yr) != 4):
            missing.append("Year of Baptism (must be 4 digits, e.g. 2026)")
        if is_excluded_parish(data.get("Parish", "")):
            missing.append("Parish (Hong Kong Parish is excluded)")
        return len(missing) == 0, missing

    def _push_images_and_encode_to_drive(
        self,
        actor_account_id: str,
        record_id: str,
        bap_year: str,
        ws_name: str,
        clean_payload: Dict[str, str],
        validated_imgs: List[Dict[str, Any]],
    ) -> Tuple[List[str], str]:
        """Uploads images/files to Google Drive `NEW BAPTISM IMAGES/<year>` and encodes the 27-col row into `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`."""
        uploaded_drive_ids: List[str] = []
        workbook_status = "LOCAL_ONLY"
        if self.drive is None:
            return uploaded_drive_ids, workbook_status

        now_iso = datetime.now(timezone.utc).isoformat()
        try:
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
        except Exception as sync_ex:
            workbook_status = "SYNC_FAILED"
            self.db.execute(
                """
                INSERT INTO workbook_update_logs (
                    update_id, record_id, worksheet_name, drive_file_id,
                    before_revision_id, after_revision_id, backup_drive_file_id,
                    status, error_message, actor_account_id, created_at
                ) VALUES (%s, %s, %s, '', '', '', '', 'FAILED', %s, %s, %s)
                """,
                (
                    f"wb_upd_{secrets.token_hex(6)}",
                    record_id,
                    ws_name,
                    str(sync_ex)[:300],
                    actor_account_id,
                    now_iso,
                ),
            )
            self.audit.log_event(
                actor_account_id,
                "WORKBOOK_UPDATE_FAILED",
                "FAILURE",
                related_record_id=record_id,
                metadata={"error": str(sync_ex)[:180]},
            )
        return uploaded_drive_ids, workbook_status

    def register_new_baptism(
        self,
        actor_account_id: str,
        payload: Dict[str, Any],
        uploaded_images: Optional[List[Dict[str, Any]]] = None,
        sync_to_drive_excel: bool = True,
    ) -> Dict[str, Any]:
        authoritative_parish = self._resolve_authoritative_parish(actor_account_id, str(payload.get("Parish", "")))
        clean_payload = {col: str(payload.get(col, "")).strip() for col in HISTORICAL_27_COLUMNS}
        clean_payload["Parish"] = authoritative_parish

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

        is_master = self.authz.is_master_admin(actor_account_id)
        ver_status = "Approved" if is_master else "Pending Verification"
        record_id = generate_stable_record_id(bap_year)
        now_iso = datetime.now(timezone.utc).isoformat()

        # Store staged files in DB so Master Admin can review & push them upon confirmation
        for v_img in validated_imgs:
            self.db.execute(
                """
                INSERT INTO staged_registration_files (
                    staged_id, record_id, filename, mime_type, file_b64,
                    file_size_bytes, uploaded_by, is_pushed_to_drive, drive_file_id, created_at
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, 0, NULL, %s)
                """,
                (
                    f"stg_{secrets.token_hex(6)}",
                    record_id,
                    v_img["filename"],
                    v_img["mime_type"],
                    base64.b64encode(v_img["bytes"]).decode("ascii"),
                    len(v_img["bytes"]),
                    actor_account_id,
                    now_iso,
                ),
            )

        initial_up_status = "STAGED_FOR_ADMIN_CONFIRMATION" if (validated_imgs and not is_master) else ("PENDING" if validated_imgs else "NONE")
        initial_wb_status = "AWAITING_ADMIN_CONFIRMATION" if not is_master else ("PENDING" if sync_to_drive_excel else "LOCAL_ONLY")

        self.db.execute(
            """
            INSERT INTO new_baptism_registrations (
                record_id, baptism_year, worksheet_name, registration_status,
                verification_status, verified_by, verified_at, verification_notes,
                upload_status, workbook_update_status, parish, child_name,
                date_of_baptism, father_name, mother_name, book_number,
                page_number, dedup_fingerprint, payload_json, created_by,
                created_at, updated_at
            ) VALUES (
                %s, %s, %s, 'REGISTERED', %s, %s, %s, '', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
            """,
            (
                record_id,
                bap_year,
                ws_name,
                ver_status,
                actor_account_id if is_master else None,
                now_iso if is_master else None,
                initial_up_status,
                initial_wb_status,
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
            metadata={"parish": clean_payload["Parish"], "worksheet": ws_name, "verification_status": ver_status},
        )

        # ONLY encode immediately to Google Drive if submitted by Master Admin!
        # Parish Leader Admin submissions wait for Master Admin confirmation.
        uploaded_drive_ids: List[str] = []
        workbook_status = initial_wb_status
        if is_master and sync_to_drive_excel:
            uploaded_drive_ids, workbook_status = self._push_images_and_encode_to_drive(
                actor_account_id, record_id, bap_year, ws_name, clean_payload, validated_imgs
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
            "parish": clean_payload["Parish"],
            "verification_status": ver_status,
            "worksheet_name": ws_name,
            "uploaded_drive_file_ids": uploaded_drive_ids,
            "workbook_update_status": workbook_status,
        }

    def get_staged_files_for_record(self, record_id: str) -> List[Dict[str, Any]]:
        rows = self.db.fetch_all(
            "SELECT * FROM staged_registration_files WHERE record_id = %s ORDER BY created_at ASC",
            (record_id,),
        )
        out = []
        for r in rows:
            item = dict(r)
            try:
                item["bytes"] = base64.b64decode(str(r.get("file_b64", "")).encode("ascii"))
            except Exception:
                item["bytes"] = b""
            out.append(item)
        return out

    def confirm_and_sync_submission_to_drive(
        self,
        admin_account_id: str,
        record_id: str,
        notes: str = "",
        sync_to_drive: bool = True,
    ) -> Dict[str, Any]:
        """Master-Admin-only confirmation: approves a Parish Leader Admin submission,
        uploads all staged images/files to Google Drive, and encodes the 27-col record into
        `NEW BAPTISMS - YYYY` inside `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` on Google Drive."""
        self.authz.require_master_admin(admin_account_id, action="confirm_and_sync_submission_to_drive")
        row = self.db.fetch_one("SELECT * FROM new_baptism_registrations WHERE record_id = %s", (record_id,))
        if not row:
            raise ValueError(f"Submission '{record_id}' not found.")

        clean_payload = json.loads(row["payload_json"])
        bap_year = str(row["baptism_year"])
        ws_name = str(row["worksheet_name"])
        staged_files = [
            {"bytes": sf["bytes"], "filename": sf["filename"], "mime_type": sf["mime_type"], "staged_id": sf["staged_id"]}
            for sf in self.get_staged_files_for_record(record_id)
            if not int(sf.get("is_pushed_to_drive", 0)) and sf.get("bytes")
        ]

        uploaded_ids: List[str] = []
        wb_status = "APPROVED_LOCAL"
        if sync_to_drive and self.drive is not None:
            uploaded_ids, wb_status = self._push_images_and_encode_to_drive(
                admin_account_id, record_id, bap_year, ws_name, clean_payload, staged_files
            )
            if uploaded_ids:
                self.db.execute(
                    "UPDATE staged_registration_files SET is_pushed_to_drive = 1 WHERE record_id = %s",
                    (record_id,),
                )

        now_iso = datetime.now(timezone.utc).isoformat()
        up_st = "UPLOADED" if uploaded_ids else ("NONE" if not staged_files else row.get("upload_status", "STAGED"))
        self.db.execute(
            """
            UPDATE new_baptism_registrations
            SET verification_status = 'Approved', verified_by = %s, verified_at = %s,
                verification_notes = %s, upload_status = %s, workbook_update_status = %s,
                updated_at = %s
            WHERE record_id = %s
            """,
            (
                admin_account_id,
                now_iso,
                notes.strip() or "Confirmed and encoded to Google Drive by Master Admin.",
                up_st,
                wb_status,
                now_iso,
                record_id,
            ),
        )
        self.audit.log_event(
            admin_account_id,
            "SUBMISSION_CONFIRMED_AND_ENCODED_TO_DRIVE",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"parish": row["parish"], "workbook_status": wb_status, "uploaded_files": len(uploaded_ids)},
        )
        return {
            "record_id": record_id,
            "parish": row["parish"],
            "verification_status": "Approved",
            "uploaded_drive_file_ids": uploaded_ids,
            "workbook_update_status": wb_status,
        }

    def update_own_submission(
        self,
        actor_account_id: str,
        record_id: str,
        updated_payload: Dict[str, Any],
    ) -> Dict[str, Any]:
        row = self.db.fetch_one("SELECT * FROM new_baptism_registrations WHERE record_id = %s", (record_id,))
        if not row:
            raise ValueError(f"Submission '{record_id}' not found.")

        is_master = self.authz.is_master_admin(actor_account_id)
        if not is_master:
            actor_parish = self.authz.get_authenticated_parish(actor_account_id)
            if row["created_by"] != actor_account_id or row["parish"] != actor_parish:
                self.audit.log_event(
                    actor_account_id,
                    "ACCESS_DENIED",
                    "FAILURE",
                    related_record_id=record_id,
                    metadata={"reason": "edit_not_owner_or_wrong_parish"},
                )
                raise AuthorizationError("Access denied: you can only edit submissions belonging to your own parish account.")
            if str(row.get("verification_status", "")) not in EDITABLE_REGISTRATION_STATUSES:
                raise AuthorizationError(
                    f"Submission '{record_id}' can no longer be edited because its status is '{row.get('verification_status')}'."
                )

        authoritative_parish = row["parish"] if not is_master else self._resolve_authoritative_parish(actor_account_id, updated_payload.get("Parish", row["parish"]))
        clean_payload = {col: str(updated_payload.get(col, "")).strip() for col in HISTORICAL_27_COLUMNS}
        clean_payload["Parish"] = authoritative_parish

        ok_val, missing = self.validate_registration_payload(clean_payload)
        if not ok_val:
            raise ValueError(f"Missing or invalid required fields: {', '.join(missing)}")

        new_fp = compute_registration_fingerprint(clean_payload)
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE new_baptism_registrations
            SET child_name = %s, date_of_baptism = %s, father_name = %s, mother_name = %s,
                book_number = %s, page_number = %s, dedup_fingerprint = %s,
                payload_json = %s, updated_at = %s
            WHERE record_id = %s
            """,
            (
                clean_payload["Name of the Child"],
                clean_payload["Date of Baptism"],
                clean_payload["Name of the Father"],
                clean_payload["Name of the Mother"],
                clean_payload["Book Number"],
                clean_payload["Page Number"],
                new_fp,
                json.dumps(clean_payload, ensure_ascii=False),
                now_iso,
                record_id,
            ),
        )
        self.audit.log_event(actor_account_id, "RECORD_UPDATED", "SUCCESS", related_record_id=record_id)
        return {"record_id": record_id, "parish": authoritative_parish, "updated_at": now_iso}

    def verify_registration(
        self,
        admin_account_id: str,
        record_id: str,
        new_status: str,
        notes: str = "",
    ) -> None:
        self.authz.require_master_admin(admin_account_id, action="verify_registration")
        if new_status not in REGISTRATION_VERIFICATION_STATUSES:
            raise ValueError(f"Invalid verification status '{new_status}'. Allowed: {REGISTRATION_VERIFICATION_STATUSES}")
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE new_baptism_registrations
            SET verification_status = %s, verified_by = %s, verified_at = %s,
                verification_notes = %s, updated_at = %s
            WHERE record_id = %s
            """,
            (new_status, admin_account_id, now_iso, notes.strip(), now_iso, record_id),
        )
        self.audit.log_event(
            admin_account_id,
            "REGISTRATION_VERIFIED",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"verification_status": new_status},
        )

    def list_submissions_for_actor(self, actor_account_id: str, parish_filter: str = "") -> List[Dict[str, Any]]:
        if self.authz.is_master_admin(actor_account_id):
            if parish_filter and parish_filter != "ALL":
                norm_p = normalize_official_nine_parish(parish_filter)
                return self.db.fetch_all(
                    "SELECT * FROM new_baptism_registrations WHERE parish = %s ORDER BY created_at DESC",
                    (norm_p,),
                )
            return self.db.fetch_all(
                "SELECT * FROM new_baptism_registrations WHERE LOWER(parish) NOT LIKE '%hong kong%' AND LOWER(parish) NOT LIKE '%hongkong%' ORDER BY created_at DESC"
            )

        actor_parish = self.authz.get_authenticated_parish(actor_account_id)
        return self.db.fetch_all(
            "SELECT * FROM new_baptism_registrations WHERE parish = %s ORDER BY created_at DESC",
            (actor_parish,),
        )

    def list_uploaded_images_for_actor(self, actor_account_id: str) -> List[Dict[str, Any]]:
        if self.authz.is_master_admin(actor_account_id):
            return self.db.fetch_all(
                """
                SELECT m.*, r.parish, r.child_name
                FROM drive_file_mappings m
                LEFT JOIN new_baptism_registrations r ON m.record_id = r.record_id
                WHERE m.category = 'NEW_BAPTISM_IMAGE'
                ORDER BY m.created_at DESC
                """
            )
        actor_parish = self.authz.get_authenticated_parish(actor_account_id)
        return self.db.fetch_all(
            """
            SELECT m.*, r.parish, r.child_name
            FROM drive_file_mappings m
            INNER JOIN new_baptism_registrations r ON m.record_id = r.record_id
            WHERE m.category = 'NEW_BAPTISM_IMAGE' AND r.parish = %s
            ORDER BY m.created_at DESC
            """,
            (actor_parish,),
        )

    def get_parish_dashboard_statistics(self, actor_account_id: str) -> Dict[str, Any]:
        actor_parish = self.authz.get_authenticated_parish(actor_account_id)
        tot_my_regs = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM new_baptism_registrations WHERE parish = %s",
            (actor_parish,),
        ) or {}).get("c", 0))
        pend_ver = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM new_baptism_registrations WHERE parish = %s AND verification_status IN ('Pending Verification', 'Additional Information Required', 'PENDING_VERIFICATION')",
            (actor_parish,),
        ) or {}).get("c", 0))
        appr_regs = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM new_baptism_registrations WHERE parish = %s AND verification_status IN ('Approved', 'APPROVED')",
            (actor_parish,),
        ) or {}).get("c", 0))
        rej_regs = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM new_baptism_registrations WHERE parish = %s AND verification_status IN ('Rejected', 'REJECTED')",
            (actor_parish,),
        ) or {}).get("c", 0))
        pend_cert_reqs = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM certificate_requests WHERE assigned_parish = %s AND UPPER(status) IN ('PENDING', 'UNDER REVIEW', 'ADDITIONAL INFORMATION REQUIRED')",
            (actor_parish,),
        ) or {}).get("c", 0))
        appr_certs = int((self.db.fetch_one(
            "SELECT COUNT(*) AS c FROM generated_certificates WHERE (assigned_parish = %s OR authorized_account_id = %s) AND is_released_to_requester = 1",
            (actor_parish, actor_account_id),
        ) or {}).get("c", 0))

        return {
            "assigned_parish": actor_parish,
            "my_new_registrations": tot_my_regs,
            "pending_verification": pend_ver,
            "approved_registrations": appr_regs,
            "rejected_registrations": rej_regs,
            "pending_certificate_requests": pend_cert_reqs,
            "approved_certificates": appr_certs,
        }

"""
baptismal/temporary_access.py — Certificate Requests & Time-Limited Record-Specific Access
=========================================================================================
- Manages certificate requests (PENDING, APPROVED, REJECTED, IN PROGRESS, COMPLETED, EXPIRED, REVOKED).
- Master Admin can approve a request for ONE specific record_id with granular permissions
  (`view_record`, `view_image`, `generate_certificate`, `download_certificate`, `print_certificate`),
  a strict expiration timestamp, and immediate revocation at any time.
"""

from datetime import datetime, timedelta, timezone
import json
import secrets
from typing import Any, Dict, List, Optional

from .audit_logs import AuditLogger
from .authorization import AuthorizationService
from .config import REQUEST_STATUSES, TEMP_ACCESS_PERMISSIONS
from .database import DatabaseManager


class TemporaryAccessService:
    """Handles certificate request lifecycle and time-limited record-specific grants."""

    def __init__(self, db: DatabaseManager, authz: AuthorizationService, audit: Optional[AuditLogger] = None):
        self.db = db
        self.authz = authz
        self.audit = audit or AuditLogger(db)

    def submit_request(
        self,
        requester_account_id: str,
        requester_username: str,
        requested_person_name: str,
        purpose: str,
        target_record_id: str = "",
    ) -> str:
        if not requested_person_name.strip() or not purpose.strip():
            raise ValueError("Requested person name and purpose are required.")
        req_id = f"REQ-{secrets.token_hex(4).upper()}"
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            INSERT INTO certificate_requests (
                request_id, requester_account_id, requester_username, target_record_id,
                requested_person_name, purpose, status, reviewed_by, reviewed_at,
                review_notes, created_at, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, 'PENDING', NULL, NULL, '', %s, %s)
            """,
            (
                req_id,
                requester_account_id,
                requester_username,
                target_record_id.strip(),
                requested_person_name.strip(),
                purpose.strip(),
                now_iso,
                now_iso,
            ),
        )
        self.audit.log_event(
            requester_account_id,
            "CERTIFICATE_REQUEST_SUBMITTED",
            "SUCCESS",
            related_record_id=target_record_id or None,
            metadata={"request_id": req_id},
        )
        return req_id

    def approve_request_and_grant_access(
        self,
        admin_account_id: str,
        request_id: str,
        record_id: str,
        allowed_permissions: List[str],
        duration_hours: float = 24.0,
        review_notes: str = "",
    ) -> str:
        self.authz.require_master_admin(admin_account_id, action="approve_temporary_access")
        req = self.db.fetch_one("SELECT * FROM certificate_requests WHERE request_id = %s", (request_id,))
        if not req:
            raise ValueError(f"Request '{request_id}' not found.")
        valid_perms = [p for p in allowed_permissions if p in TEMP_ACCESS_PERMISSIONS]
        if not valid_perms:
            raise ValueError("At least one valid permission must be selected.")

        now = datetime.now(timezone.utc)
        expires_at = (now + timedelta(hours=float(duration_hours))).isoformat()
        grant_id = f"GRANT-{secrets.token_hex(5).upper()}"

        self.db.execute(
            """
            UPDATE certificate_requests
            SET status = 'APPROVED', target_record_id = %s, reviewed_by = %s,
                reviewed_at = %s, review_notes = %s, updated_at = %s
            WHERE request_id = %s
            """,
            (record_id, admin_account_id, now.isoformat(), review_notes.strip(), now.isoformat(), request_id),
        )
        self.db.execute(
            """
            INSERT INTO temporary_access_grants (
                grant_id, request_id, account_id, record_id, permissions_json,
                expires_at, is_revoked, revoked_by, revoked_at, granted_by, created_at
            ) VALUES (%s, %s, %s, %s, %s, %s, 0, NULL, NULL, %s, %s)
            """,
            (
                grant_id,
                request_id,
                req["requester_account_id"],
                record_id,
                json.dumps(valid_perms),
                expires_at,
                admin_account_id,
                now.isoformat(),
            ),
        )
        self.audit.log_event(
            admin_account_id,
            "REQUEST_APPROVED",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"request_id": request_id, "grant_id": grant_id, "expires_at": expires_at},
        )
        self.audit.log_event(
            admin_account_id,
            "TEMPORARY_ACCESS_GRANTED",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"grant_id": grant_id, "target_account_id": req["requester_account_id"], "permissions": valid_perms},
        )
        return grant_id

    def reject_request(self, admin_account_id: str, request_id: str, review_notes: str = "") -> None:
        self.authz.require_master_admin(admin_account_id, action="reject_certificate_request")
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE certificate_requests
            SET status = 'REJECTED', reviewed_by = %s, reviewed_at = %s, review_notes = %s, updated_at = %s
            WHERE request_id = %s
            """,
            (admin_account_id, now_iso, review_notes.strip(), now_iso, request_id),
        )
        self.audit.log_event(
            admin_account_id,
            "REQUEST_REJECTED",
            "SUCCESS",
            metadata={"request_id": request_id},
        )

    def update_request_status(self, admin_account_id: str, request_id: str, new_status: str) -> None:
        self.authz.require_master_admin(admin_account_id, action="update_request_status")
        if new_status not in REQUEST_STATUSES:
            raise ValueError(f"Invalid request status: {new_status}")
        self.db.execute(
            "UPDATE certificate_requests SET status = %s, updated_at = %s WHERE request_id = %s",
            (new_status, datetime.now(timezone.utc).isoformat(), request_id),
        )

    def revoke_grant_immediately(self, admin_account_id: str, grant_id: str) -> None:
        self.authz.require_master_admin(admin_account_id, action="revoke_temporary_access")
        grant = self.db.fetch_one("SELECT * FROM temporary_access_grants WHERE grant_id = %s", (grant_id,))
        if not grant:
            raise ValueError(f"Grant '{grant_id}' not found.")
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE temporary_access_grants
            SET is_revoked = 1, revoked_by = %s, revoked_at = %s
            WHERE grant_id = %s
            """,
            (admin_account_id, now_iso, grant_id),
        )
        if grant.get("request_id"):
            self.db.execute(
                "UPDATE certificate_requests SET status = 'REVOKED', updated_at = %s WHERE request_id = %s",
                (now_iso, grant["request_id"]),
            )
        self.audit.log_event(
            admin_account_id,
            "TEMPORARY_ACCESS_REVOKED",
            "SUCCESS",
            related_record_id=grant.get("record_id"),
            metadata={"grant_id": grant_id},
        )

    def list_requests(self, requester_account_id: str = "") -> List[Dict[str, Any]]:
        if requester_account_id:
            return self.db.fetch_all(
                "SELECT * FROM certificate_requests WHERE requester_account_id = %s ORDER BY created_at DESC",
                (requester_account_id,),
            )
        return self.db.fetch_all("SELECT * FROM certificate_requests ORDER BY created_at DESC")

    def list_active_grants(self, account_id: str = "") -> List[Dict[str, Any]]:
        if account_id:
            return self.db.fetch_all(
                "SELECT * FROM temporary_access_grants WHERE account_id = %s ORDER BY created_at DESC",
                (account_id,),
            )
        return self.db.fetch_all("SELECT * FROM temporary_access_grants ORDER BY created_at DESC")

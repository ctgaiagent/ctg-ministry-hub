"""
baptismal/temporary_access.py — Historical Record Requests, Certified Certificate Requests & Temporary Grants
=============================================================================================================
Implements:
1. Section 6: Historical Record Requests (Parish Leader Admins submit requests with child name,
   father's name, mother's name, approx baptism/birth dates, parish, book/page hint, reason,
   and optional supporting document; ONLY Master Admin searches and verifies the archive).
2. Section 7: Certified Baptismal Certificate Requests (7-step workflow where Parish Admin submits,
   Master Admin reviews, verifies historical record, approves/rejects, generates, officially issues,
   and releases the certificate specifically to the requesting account).
3. Section 8: Master Admin Temporary Access (user-specific, record-specific, action-specific,
   time-limited, immediately revocable; NEVER grants full archive access).
"""

from datetime import datetime, timedelta, timezone
import json
import secrets
from typing import Any, Dict, List, Optional

from .audit_logs import AuditLogger
from .authorization import AuthorizationError, AuthorizationService
from .config import (
    CANONICAL_REQUEST_STATUSES,
    REQUEST_STATUSES,
    TEMP_ACCESS_PERMISSIONS,
    is_excluded_parish,
    normalize_official_nine_parish,
)
from .database import DatabaseManager


class TemporaryAccessService:
    """Handles historical record requests, certified certificate requests, and time-limited record-specific grants."""

    def __init__(self, db: DatabaseManager, authz: AuthorizationService, audit: Optional[AuditLogger] = None):
        self.db = db
        self.authz = authz
        self.audit = audit or AuditLogger(db)

    def _resolve_requester_parish(self, requester_account_id: str, explicit_parish: str = "") -> str:
        acc = self.authz.get_live_account(requester_account_id)
        if not acc:
            raise AuthorizationError("Access denied: unauthenticated account.")
        if acc["role"] == "Parish Leader Admin":
            return self.authz.get_authenticated_parish(requester_account_id)
        if explicit_parish.strip():
            if is_excluded_parish(explicit_parish):
                raise ValueError("Hong Kong Parish is excluded from the CTG Baptismal System.")
            return normalize_official_nine_parish(explicit_parish)
        return str(acc.get("assigned_parish") or "")

    def submit_request(
        self,
        requester_account_id: str,
        requester_username: str,
        requested_person_name: str,
        purpose: str,
        target_record_id: str = "",
        request_type: str = "CERTIFIED_CERTIFICATE",
        father_name: str = "",
        mother_name: str = "",
        approx_baptism_date: str = "",
        approx_birth_date: str = "",
        parish: str = "",
        book_hint: str = "",
        page_hint: str = "",
        supporting_doc_filename: str = "",
        supporting_doc_drive_id: str = "",
    ) -> str:
        if not requested_person_name.strip() or not purpose.strip():
            raise ValueError("Name of baptized person and reason/purpose for request are required.")

        resolved_parish = self._resolve_requester_parish(requester_account_id, parish)
        req_type_clean = "HISTORICAL_RECORD" if "HIST" in str(request_type).upper() else "CERTIFIED_CERTIFICATE"
        req_id = f"REQ-{secrets.token_hex(4).upper()}"
        now_iso = datetime.now(timezone.utc).isoformat()

        self.db.execute(
            """
            INSERT INTO certificate_requests (
                request_id, request_type, requester_account_id, requester_username,
                assigned_parish, target_record_id, requested_person_name,
                father_name, mother_name, approx_baptism_date, approx_birth_date,
                book_hint, page_hint, purpose, supporting_doc_filename,
                supporting_doc_drive_id, approved_certificate_id, status,
                reviewed_by, reviewed_at, review_notes, created_at, updated_at
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NULL,
                'Pending', NULL, NULL, '', %s, %s
            )
            """,
            (
                req_id,
                req_type_clean,
                requester_account_id,
                requester_username,
                resolved_parish,
                target_record_id.strip(),
                requested_person_name.strip(),
                father_name.strip(),
                mother_name.strip(),
                approx_baptism_date.strip(),
                approx_birth_date.strip(),
                book_hint.strip(),
                page_hint.strip(),
                purpose.strip(),
                supporting_doc_filename.strip(),
                supporting_doc_drive_id.strip(),
                now_iso,
                now_iso,
            ),
        )
        self.audit.log_event(
            requester_account_id,
            "HISTORICAL_RECORD_REQUESTED" if req_type_clean == "HISTORICAL_RECORD" else "CERTIFICATE_REQUEST_SUBMITTED",
            "SUCCESS",
            related_record_id=target_record_id or None,
            metadata={"request_id": req_id, "parish": resolved_parish, "type": req_type_clean},
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
        """Master-Admin-only approval of a request with a time-limited, record-specific temporary access grant."""
        self.authz.require_master_admin(admin_account_id, action="approve_temporary_access")
        req = self.db.fetch_one("SELECT * FROM certificate_requests WHERE request_id = %s", (request_id,))
        if not req:
            raise ValueError(f"Request '{request_id}' not found.")
        if req["requester_account_id"] == admin_account_id and not self.authz.is_master_admin(admin_account_id):
            raise AuthorizationError("Users cannot approve their own requests.")

        valid_perms = [p for p in allowed_permissions if p in TEMP_ACCESS_PERMISSIONS]
        if not valid_perms:
            raise ValueError("At least one valid permission must be selected.")

        now = datetime.now(timezone.utc)
        expires_at = (now + timedelta(hours=float(duration_hours))).isoformat()
        grant_id = f"GRANT-{secrets.token_hex(5).upper()}"

        self.db.execute(
            """
            UPDATE certificate_requests
            SET status = 'Approved', target_record_id = %s, reviewed_by = %s,
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
            SET status = 'Rejected', reviewed_by = %s, reviewed_at = %s, review_notes = %s, updated_at = %s
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

    def update_request_status(
        self,
        admin_account_id: str,
        request_id: str,
        new_status: str,
        review_notes: str = "",
        target_record_id: str = "",
    ) -> None:
        self.authz.require_master_admin(admin_account_id, action="update_request_status")
        if new_status not in REQUEST_STATUSES and new_status not in CANONICAL_REQUEST_STATUSES:
            raise ValueError(f"Invalid request status: {new_status}")
        now_iso = datetime.now(timezone.utc).isoformat()
        if target_record_id.strip():
            self.db.execute(
                """
                UPDATE certificate_requests
                SET status = %s, target_record_id = %s, reviewed_by = %s,
                    reviewed_at = %s, review_notes = COALESCE(NULLIF(%s, ''), review_notes), updated_at = %s
                WHERE request_id = %s
                """,
                (new_status, target_record_id.strip(), admin_account_id, now_iso, review_notes.strip(), now_iso, request_id),
            )
        else:
            self.db.execute(
                """
                UPDATE certificate_requests
                SET status = %s, reviewed_by = %s, reviewed_at = %s,
                    review_notes = COALESCE(NULLIF(%s, ''), review_notes), updated_at = %s
                WHERE request_id = %s
                """,
                (new_status, admin_account_id, now_iso, review_notes.strip(), now_iso, request_id),
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
                "UPDATE certificate_requests SET status = 'Revoked', updated_at = %s WHERE request_id = %s",
                (now_iso, grant["request_id"]),
            )
        self.audit.log_event(
            admin_account_id,
            "TEMPORARY_ACCESS_REVOKED",
            "SUCCESS",
            related_record_id=grant.get("record_id"),
            metadata={"grant_id": grant_id},
        )

    def list_requests(self, requester_account_id: str = "", actor_account_id: str = "") -> List[Dict[str, Any]]:
        """Filters requests server-side: Master Admin sees all 9 parishes; Parish Leader Admin sees ONLY their own parish & account."""
        eff_actor = actor_account_id or requester_account_id
        if eff_actor and not self.authz.is_master_admin(eff_actor):
            acc = self.authz.get_live_account(eff_actor)
            if not acc:
                return []
            if acc["role"] == "Parish Leader Admin":
                parish = self.authz.get_authenticated_parish(eff_actor)
                return self.db.fetch_all(
                    "SELECT * FROM certificate_requests WHERE requester_account_id = %s AND assigned_parish = %s ORDER BY created_at DESC",
                    (eff_actor, parish),
                )
            return self.db.fetch_all(
                "SELECT * FROM certificate_requests WHERE requester_account_id = %s ORDER BY created_at DESC",
                (eff_actor,),
            )

        if requester_account_id:
            return self.db.fetch_all(
                "SELECT * FROM certificate_requests WHERE requester_account_id = %s ORDER BY created_at DESC",
                (requester_account_id,),
            )
        return self.db.fetch_all(
            "SELECT * FROM certificate_requests WHERE LOWER(COALESCE(assigned_parish, '')) NOT LIKE '%hong kong%' ORDER BY created_at DESC"
        )

    def list_active_grants(self, account_id: str = "") -> List[Dict[str, Any]]:
        if account_id:
            return self.db.fetch_all(
                "SELECT * FROM temporary_access_grants WHERE account_id = %s ORDER BY created_at DESC",
                (account_id,),
            )
        return self.db.fetch_all("SELECT * FROM temporary_access_grants ORDER BY created_at DESC")

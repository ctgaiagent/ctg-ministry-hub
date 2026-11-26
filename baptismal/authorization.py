"""
baptismal/authorization.py — Server-Side Role, Parish & Record-Specific Authorization
=====================================================================================
Enforces:
1. Master Admin has exclusive permanent access to the complete historical archive,
   BOOK 1-9 visual archive, global search, registration verification, certificate
   generation/approval/issuance, temporary access grants, and audit logs.
2. Each of the 9 Parish Leader Admins is locked server-side to their fixed `assigned_parish`
   (from `OFFICIAL_NINE_PARISHES`). They CANNOT:
   - Access BOOK 1-9
   - Search the complete historical baptismal database
   - Access any other parish's submissions, images, requests, or certificates
   - Generate, approve, or issue official certificates independently
   - Access Master Admin settings or global audit logs
3. Hong Kong Parish is strictly excluded from all operational dashboards and permissions.
"""

from datetime import datetime, timezone
import json
from typing import Any, Dict, Optional

from .audit_logs import AuditLogger
from .config import is_excluded_parish, normalize_official_nine_parish
from .database import DatabaseManager


class AuthorizationError(PermissionError):
    """Raised when a user attempts an unauthorized action, cross-parish access, or archive access."""


class AuthorizationService:
    """Enforces server-side RBAC, fixed parish isolation, and temporary record-specific access checks."""

    def __init__(self, db: DatabaseManager, audit: Optional[AuditLogger] = None):
        self.db = db
        self.audit = audit or AuditLogger(db)

    def get_live_account(self, account_id: str) -> Optional[Dict[str, Any]]:
        if not account_id:
            return None
        acc = self.db.fetch_one(
            "SELECT account_id, username, full_name, role, assigned_parish, is_active FROM baptismal_accounts WHERE account_id = %s",
            (account_id,),
        )
        if not acc or not int(acc.get("is_active", 0)):
            return None
        if is_excluded_parish(acc.get("assigned_parish", "")):
            return None
        return acc

    def is_master_admin(self, account_id: str) -> bool:
        acc = self.get_live_account(account_id)
        return bool(acc and acc.get("role") == "Master Admin")

    def is_parish_leader_admin(self, account_id: str) -> bool:
        acc = self.get_live_account(account_id)
        return bool(acc and acc.get("role") == "Parish Leader Admin")

    def get_authenticated_parish(self, account_id: str) -> str:
        """Returns the authoritative server-side assigned parish for a Parish Leader Admin."""
        acc = self.get_live_account(account_id)
        if not acc:
            raise AuthorizationError("Access denied: unauthenticated or inactive account.")
        if acc["role"] != "Parish Leader Admin":
            raise AuthorizationError("Account is not a Parish Leader Admin.")
        return normalize_official_nine_parish(acc.get("assigned_parish", ""))

    def require_master_admin(self, account_id: str, action: str = "admin_action") -> Dict[str, Any]:
        acc = self.get_live_account(account_id)
        if not acc or acc.get("role") != "Master Admin":
            self.audit.log_event(
                account_id or "anonymous",
                "ACCESS_DENIED",
                "FAILURE",
                metadata={"action": action, "reason": "master_admin_required"},
            )
            raise AuthorizationError("Access denied: Master Admin authorization is required.")
        return acc

    def require_parish_access(self, account_id: str, target_parish: str, action: str = "parish_action") -> str:
        """Verifies that the user is either Master Admin or a Parish Leader Admin assigned to `target_parish`.
        Rejects Hong Kong Parish and cross-parish access attempts."""
        if is_excluded_parish(target_parish):
            self.audit.log_event(
                account_id or "anonymous",
                "ACCESS_DENIED",
                "FAILURE",
                metadata={"action": action, "reason": "excluded_hong_kong_parish"},
            )
            raise AuthorizationError("Access denied: Hong Kong Parish is excluded from the Baptismal System.")

        norm_target = normalize_official_nine_parish(target_parish)
        acc = self.get_live_account(account_id)
        if not acc:
            self.audit.log_event("anonymous", "ACCESS_DENIED", "FAILURE", metadata={"action": action})
            raise AuthorizationError("Access denied: unauthenticated account.")

        if acc["role"] == "Master Admin":
            return norm_target

        if acc["role"] == "Parish Leader Admin":
            assigned = normalize_official_nine_parish(acc.get("assigned_parish", ""))
            if assigned == norm_target:
                return assigned

        self.audit.log_event(
            account_id,
            "CROSS_PARISH_ACCESS_DENIED",
            "FAILURE",
            metadata={"action": action, "assigned_parish": acc.get("assigned_parish", ""), "target_parish": norm_target},
        )
        raise AuthorizationError(
            f"Access denied: account for '{acc.get('assigned_parish', 'Unassigned')}' cannot access '{norm_target}'."
        )

    def has_temporary_record_permission(
        self,
        account_id: str,
        record_id: str,
        required_permission: str,
    ) -> bool:
        """Checks whether a non-Master-Admin user has a valid, non-expired, non-revoked
        grant for ONE specific record_id and specific permission."""
        if self.is_master_admin(account_id):
            return True
        acc = self.get_live_account(account_id)
        if not acc:
            return False

        now = datetime.now(timezone.utc)
        grants = self.db.fetch_all(
            """
            SELECT grant_id, permissions_json, expires_at, is_revoked
            FROM temporary_access_grants
            WHERE account_id = %s AND record_id = %s AND is_revoked = 0
            ORDER BY created_at DESC
            """,
            (account_id, record_id),
        )
        for g in grants:
            try:
                exp_dt = datetime.fromisoformat(str(g["expires_at"]))
                if exp_dt <= now:
                    continue
                perms = json.loads(g["permissions_json"])
                if isinstance(perms, list) and required_permission in perms:
                    return True
            except Exception:
                continue
        return False

    def require_record_permission(
        self,
        account_id: str,
        record_id: str,
        required_permission: str,
    ) -> None:
        if not self.has_temporary_record_permission(account_id, record_id, required_permission):
            self.audit.log_event(
                account_id or "anonymous",
                "ACCESS_DENIED",
                "FAILURE",
                related_record_id=record_id,
                metadata={"required_permission": required_permission},
            )
            raise AuthorizationError(
                f"Access denied: missing permission '{required_permission}' for record '{record_id}'."
            )

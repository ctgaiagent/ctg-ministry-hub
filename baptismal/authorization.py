"""
baptismal/authorization.py — Server-Side Role & Record-Specific Authorization
===========================================================================
- Master Admin has permanent full access to the Baptismal System.
- Existing CTG Ministry Hub login does NOT grant baptismal access.
- Temporary Requesters NEVER get general search, full workbook, or BOOK 1-9 archive access.
  They can ONLY access a specific approved record_id with explicitly granted permissions
  within a non-expired, non-revoked time window.
"""

from datetime import datetime, timezone
import json
from typing import Any, Dict, Optional

from .audit_logs import AuditLogger
from .database import DatabaseManager


class AuthorizationError(PermissionError):
    """Raised when a user attempts an unauthorized action or record access."""


class AuthorizationService:
    """Enforces server-side RBAC and temporary record-specific access checks."""

    def __init__(self, db: DatabaseManager, audit: Optional[AuditLogger] = None):
        self.db = db
        self.audit = audit or AuditLogger(db)

    def _get_live_account(self, account_id: str) -> Optional[Dict[str, Any]]:
        if not account_id:
            return None
        acc = self.db.fetch_one(
            "SELECT account_id, username, role, is_active FROM baptismal_accounts WHERE account_id = %s",
            (account_id,),
        )
        if not acc or not int(acc.get("is_active", 0)):
            return None
        return acc

    def is_master_admin(self, account_id: str) -> bool:
        acc = self._get_live_account(account_id)
        return bool(acc and acc.get("role") == "Master Admin")

    def require_master_admin(self, account_id: str, action: str = "admin_action") -> Dict[str, Any]:
        acc = self._get_live_account(account_id)
        if not acc or acc.get("role") != "Master Admin":
            self.audit.log_event(
                account_id or "anonymous",
                "ACCESS_DENIED",
                "FAILURE",
                metadata={"action": action, "reason": "master_admin_required"},
            )
            raise AuthorizationError("Access denied: Master Admin authorization is required.")
        return acc

    def has_temporary_record_permission(
        self,
        account_id: str,
        record_id: str,
        required_permission: str,
    ) -> bool:
        """Checks whether a non-Master-Admin user has a valid, non-expired, non-revoked
        grant for the specific record_id and specific permission."""
        if self.is_master_admin(account_id):
            return True
        acc = self._get_live_account(account_id)
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

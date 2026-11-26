"""
baptismal/authentication.py — Server-Side PBKDF2 Authentication & 10-Account Provisioning
=========================================================================================
Supports the 10 canonical account types:
1. One Master Admin (`role = 'Master Admin'`, `assigned_parish = 'ALL_NINE_PARISHES'`)
2. Nine Parish Leader Admins (`role = 'Parish Leader Admin'`, fixed `assigned_parish` from `OFFICIAL_NINE_PARISHES`)
   - Hong Kong Parish is strictly rejected.
- PBKDF2-HMAC-SHA256 (310,000 iterations, 16-byte random salt)
- Rate limiting & 15-minute lockout after 5 failed attempts
- Server-side session re-validation reloads `role` and `assigned_parish` from the database on every request.
"""

import base64
from datetime import datetime, timedelta, timezone
import hashlib
import hmac
import secrets
from typing import Any, Dict, List, Optional, Tuple

from .audit_logs import AuditLogger
from .config import (
    DEFAULT_NINE_PARISH_ADMIN_USERNAMES,
    LOCKOUT_DURATION_SECONDS,
    MAX_FAILED_LOGIN_ATTEMPTS,
    OFFICIAL_NINE_PARISHES,
    PBKDF2_ITERATIONS,
    SESSION_TIMEOUT_SECONDS,
    get_config_value,
    is_excluded_parish,
    normalize_official_nine_parish,
)
from .database import DatabaseManager

VALID_ROLES = ("Master Admin", "Parish Leader Admin", "Temporary Requester")


def hash_password_pbkdf2(plain_password: str, iterations: int = PBKDF2_ITERATIONS) -> str:
    if not plain_password or len(plain_password) < 8:
        raise ValueError("Password must be at least 8 characters long.")
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", plain_password.encode("utf-8"), salt, int(iterations))
    salt_b64 = base64.b64encode(salt).decode("ascii")
    dk_b64 = base64.b64encode(dk).decode("ascii")
    return f"pbkdf2_sha256${int(iterations)}${salt_b64}${dk_b64}"


def verify_password_pbkdf2(plain_password: str, stored_hash: str) -> bool:
    if not plain_password or not stored_hash or not stored_hash.startswith("pbkdf2_sha256$"):
        return False
    try:
        parts = stored_hash.split("$")
        if len(parts) != 4:
            return False
        _, iter_s, salt_b64, dk_b64 = parts
        iterations = int(iter_s)
        salt = base64.b64decode(salt_b64.encode("ascii"))
        expected_dk = base64.b64decode(dk_b64.encode("ascii"))
        actual_dk = hashlib.pbkdf2_hmac("sha256", plain_password.encode("utf-8"), salt, iterations)
        return hmac.compare_digest(actual_dk, expected_dk)
    except Exception:
        return False


class AuthService:
    """Handles server-side authentication, lockout, session validation, and 10-account provisioning."""

    def __init__(self, db: DatabaseManager, audit: Optional[AuditLogger] = None):
        self.db = db
        self.audit = audit or AuditLogger(db)

    def has_master_admin(self) -> bool:
        row = self.db.fetch_one(
            "SELECT account_id FROM baptismal_accounts WHERE role = %s AND is_active = 1 LIMIT 1",
            ("Master Admin",),
        )
        return row is not None

    def ensure_initial_master_admin_from_env(self) -> bool:
        if self.has_master_admin():
            return True
        init_user = get_config_value("BAPTISMAL_INITIAL_ADMIN_USER")
        init_hash = get_config_value("BAPTISMAL_INITIAL_ADMIN_PASSWORD_HASH")
        init_plain = get_config_value("BAPTISMAL_INITIAL_ADMIN_PASSWORD")
        if not init_user or (not init_hash and not init_plain):
            return False
        pw_hash = init_hash if (init_hash and init_hash.startswith("pbkdf2_sha256$")) else hash_password_pbkdf2(init_plain)
        self.create_account(
            username=init_user,
            full_name="CTG Baptismal Master Admin",
            role="Master Admin",
            assigned_parish="ALL_NINE_PARISHES",
            password_hash=pw_hash,
            email=get_config_value("BAPTISMAL_INITIAL_ADMIN_EMAIL"),
        )
        return True

    def create_account(
        self,
        username: str,
        full_name: str,
        role: str,
        plain_password: str = "",
        password_hash: str = "",
        email: str = "",
        assigned_parish: str = "",
    ) -> str:
        uname = str(username or "").strip().lower()
        if not uname:
            raise ValueError("Username is required.")
        if role not in VALID_ROLES:
            raise ValueError(f"Invalid role '{role}'. Allowed: {VALID_ROLES}")

        if is_excluded_parish(assigned_parish):
            raise ValueError("Hong Kong Parish is excluded from the CTG Baptismal System.")

        if role == "Master Admin":
            norm_parish = "ALL_NINE_PARISHES"
        elif role == "Parish Leader Admin":
            norm_parish = normalize_official_nine_parish(assigned_parish)
        else:
            norm_parish = normalize_official_nine_parish(assigned_parish) if assigned_parish.strip() else ""

        final_hash = password_hash if password_hash.startswith("pbkdf2_sha256$") else hash_password_pbkdf2(plain_password)
        account_id = f"bap_acc_{secrets.token_hex(8)}"
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            INSERT INTO baptismal_accounts (
                account_id, username, full_name, email, role, assigned_parish,
                password_hash, is_active, failed_attempts, locked_until,
                must_change_password, created_at, updated_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, 1, 0, NULL, 0, %s, %s)
            """,
            (
                account_id,
                uname,
                str(full_name or uname).strip(),
                str(email or "").strip(),
                role,
                norm_parish,
                final_hash,
                now_iso,
                now_iso,
            ),
        )
        self.audit.log_event(
            account_id,
            "ACCOUNT_CREATED",
            "SUCCESS",
            metadata={"role": role, "assigned_parish": norm_parish},
        )
        return account_id

    def provision_nine_parish_leader_admins(
        self,
        actor_account_id: str,
        initial_password: str,
        custom_usernames: Optional[Dict[str, str]] = None,
    ) -> List[Dict[str, str]]:
        """Creates or updates the 9 separate Parish Leader Admin accounts (one for each of the 9 official parishes;
        Hong Kong Parish is never included)."""
        created_or_existing: List[Dict[str, str]] = []
        uname_map = {**DEFAULT_NINE_PARISH_ADMIN_USERNAMES, **(custom_usernames or {})}
        for parish in OFFICIAL_NINE_PARISHES:
            uname = str(uname_map[parish]).strip().lower()
            existing = self.db.fetch_one(
                "SELECT account_id, username, assigned_parish FROM baptismal_accounts WHERE username = %s OR (role = 'Parish Leader Admin' AND assigned_parish = %s)",
                (uname, parish),
            )
            if existing:
                created_or_existing.append({
                    "account_id": existing["account_id"],
                    "username": existing["username"],
                    "assigned_parish": parish,
                    "status": "EXISTING",
                })
            else:
                aid = self.create_account(
                    username=uname,
                    full_name=f"{parish} Baptismal Leader Admin",
                    role="Parish Leader Admin",
                    assigned_parish=parish,
                    plain_password=initial_password,
                )
                created_or_existing.append({
                    "account_id": aid,
                    "username": uname,
                    "assigned_parish": parish,
                    "status": "CREATED",
                })
        self.audit.log_event(
            actor_account_id,
            "NINE_PARISH_ADMINS_PROVISIONED",
            "SUCCESS",
            metadata={"count": len(created_or_existing)},
        )
        return created_or_existing

    def list_parish_admin_accounts(self) -> List[Dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT account_id, username, full_name, email, role, assigned_parish,
                   is_active, failed_attempts, locked_until, created_at
            FROM baptismal_accounts
            WHERE role = 'Parish Leader Admin'
            ORDER BY assigned_parish ASC
            """
        )

    def authenticate(self, username: str, plain_password: str) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """Verifies credentials server-side with lockout protection and generic error messages."""
        uname = str(username or "").strip().lower()
        generic_err = "Invalid credentials or account temporarily locked."
        if not uname or not plain_password:
            return False, None, generic_err

        acc = self.db.fetch_one(
            "SELECT * FROM baptismal_accounts WHERE username = %s",
            (uname,),
        )
        if not acc or not int(acc.get("is_active", 0)):
            self.audit.log_event("anonymous", "LOGIN_FAILED", "FAILURE", metadata={"reason": "unknown_or_inactive"})
            return False, None, generic_err

        if is_excluded_parish(acc.get("assigned_parish", "")):
            self.audit.log_event(acc["account_id"], "LOGIN_BLOCKED_EXCLUDED_PARISH", "FAILURE")
            return False, None, generic_err

        now = datetime.now(timezone.utc)
        locked_until_raw = acc.get("locked_until")
        if locked_until_raw:
            try:
                locked_dt = datetime.fromisoformat(str(locked_until_raw))
                if locked_dt > now:
                    self.audit.log_event(acc["account_id"], "LOGIN_LOCKED_ATTEMPT", "FAILURE")
                    return False, None, generic_err
            except Exception:
                pass

        if not verify_password_pbkdf2(plain_password, str(acc.get("password_hash", ""))):
            new_fails = int(acc.get("failed_attempts", 0)) + 1
            locked_str = None
            if new_fails >= MAX_FAILED_LOGIN_ATTEMPTS:
                locked_str = (now + timedelta(seconds=LOCKOUT_DURATION_SECONDS)).isoformat()
                self.audit.log_event(acc["account_id"], "ACCOUNT_LOCKED", "FAILURE", metadata={"failed_attempts": new_fails})
            else:
                self.audit.log_event(acc["account_id"], "LOGIN_FAILED", "FAILURE", metadata={"failed_attempts": new_fails})
            self.db.execute(
                "UPDATE baptismal_accounts SET failed_attempts = %s, locked_until = %s, updated_at = %s WHERE account_id = %s",
                (new_fails, locked_str, now.isoformat(), acc["account_id"]),
            )
            return False, None, generic_err

        self.db.execute(
            "UPDATE baptismal_accounts SET failed_attempts = 0, locked_until = NULL, updated_at = %s WHERE account_id = %s",
            (now.isoformat(), acc["account_id"]),
        )
        safe_user = {
            "account_id": acc["account_id"],
            "username": acc["username"],
            "full_name": acc["full_name"],
            "email": acc.get("email", ""),
            "role": acc["role"],
            "assigned_parish": acc.get("assigned_parish", ""),
            "authenticated_at": now.isoformat(),
            "session_token": secrets.token_urlsafe(32),
        }
        self.audit.log_event(
            acc["account_id"],
            "LOGIN_SUCCESS",
            "SUCCESS",
            metadata={"role": acc["role"], "assigned_parish": acc.get("assigned_parish", "")},
        )
        return True, safe_user, "OK"

    def validate_session(self, session_user: Optional[Dict[str, Any]]) -> Tuple[bool, Optional[Dict[str, Any]]]:
        """Re-verifies session expiration and reloads role & assigned_parish directly from the database."""
        if not session_user or not isinstance(session_user, dict):
            return False, None
        acc_id = str(session_user.get("account_id", "")).strip()
        auth_at = str(session_user.get("authenticated_at", "")).strip()
        if not acc_id or not auth_at:
            return False, None
        try:
            auth_dt = datetime.fromisoformat(auth_at)
            if (datetime.now(timezone.utc) - auth_dt).total_seconds() > SESSION_TIMEOUT_SECONDS:
                return False, None
        except Exception:
            return False, None

        acc = self.db.fetch_one(
            "SELECT account_id, username, full_name, email, role, assigned_parish, is_active FROM baptismal_accounts WHERE account_id = %s",
            (acc_id,),
        )
        if not acc or not int(acc.get("is_active", 0)):
            return False, None
        if is_excluded_parish(acc.get("assigned_parish", "")):
            return False, None
        return True, {
            **session_user,
            "role": acc["role"],
            "assigned_parish": acc.get("assigned_parish", ""),
            "username": acc["username"],
            "full_name": acc["full_name"],
        }

    def reset_password(self, account_id: str, new_plain_password: str, actor_account_id: str) -> None:
        new_hash = hash_password_pbkdf2(new_plain_password)
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            "UPDATE baptismal_accounts SET password_hash = %s, failed_attempts = 0, locked_until = NULL, updated_at = %s WHERE account_id = %s",
            (new_hash, now_iso, account_id),
        )
        self.audit.log_event(actor_account_id, "PASSWORD_RESET", "SUCCESS", metadata={"target_account_id": account_id})

    def logout(self, session_user: Optional[Dict[str, Any]]) -> None:
        if session_user and session_user.get("account_id"):
            self.audit.log_event(str(session_user["account_id"]), "LOGOUT", "SUCCESS")

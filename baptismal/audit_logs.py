"""
baptismal/audit_logs.py — Tamper-Aware Baptismal Audit Trail
============================================================
Records all authentication, authorization, historical view, registration,
Drive upload, image matching, certificate, temporary access, and workbook events.
Never logs passwords, hashes, OAuth tokens, API keys, or full confidential records.
"""

from datetime import datetime, timezone
import json
import re
import secrets
from typing import Any, Dict, List, Optional

from .database import DatabaseManager

_SECRET_SCRUB_PATTERNS = [
    re.compile(r"pbkdf2_sha256\$[0-9]+\$[A-Za-z0-9+/=]+\$[A-Za-z0-9+/=]+"),
    re.compile(r"\b1//[0-9A-Za-z_\-]{20,}\b"),
    re.compile(r"\bya29\.[0-9A-Za-z_\-\.]{20,}\b"),
    re.compile(r"\bAIza[0-9A-Za-z_\-]{25,}\b"),
    re.compile(r"(?i)(password|passwd|secret|token|api_key|hash)\s*[:=]\s*['\"]?[^\s,;'\"]+"),
]


def _scrub_sensitive_metadata(raw_meta: Dict[str, Any]) -> str:
    safe_dict: Dict[str, Any] = {}
    blocked_keys = {"password", "password_hash", "token", "refresh_token", "access_token", "api_key", "secret"}
    for k, v in (raw_meta or {}).items():
        if str(k).lower() in blocked_keys:
            safe_dict[k] = "[REDACTED]"
            continue
        s = str(v if v is not None else "")
        for pat in _SECRET_SCRUB_PATTERNS:
            s = pat.sub("[REDACTED]", s)
        safe_dict[str(k)] = s[:240]
    return json.dumps(safe_dict, ensure_ascii=False)


class AuditLogger:
    """Appends structured non-sensitive audit events to baptismal_audit_logs."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    def log_event(
        self,
        account_id: str,
        event_type: str,
        outcome: str = "SUCCESS",
        related_record_id: Optional[str] = None,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> str:
        event_id = f"bap_evt_{int(datetime.now(timezone.utc).timestamp() * 1000)}_{secrets.token_hex(4)}"
        ts = datetime.now(timezone.utc).isoformat()
        meta_json = _scrub_sensitive_metadata(metadata or {})
        self.db.execute(
            """
            INSERT INTO baptismal_audit_logs (
                event_id, account_id, event_type, related_record_id, timestamp, outcome, metadata_json
            ) VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            (
                event_id,
                str(account_id or "anonymous"),
                str(event_type).upper(),
                related_record_id,
                ts,
                str(outcome).upper(),
                meta_json,
            ),
        )
        return event_id

    def list_events(self, event_type: str = "", limit: int = 200) -> List[Dict[str, Any]]:
        if event_type and event_type != "ALL":
            return self.db.fetch_all(
                """
                SELECT event_id, account_id, event_type, related_record_id, timestamp, outcome, metadata_json
                FROM baptismal_audit_logs
                WHERE event_type = %s
                ORDER BY timestamp DESC
                LIMIT %s
                """,
                (event_type, int(limit)),
            )
        return self.db.fetch_all(
            """
            SELECT event_id, account_id, event_type, related_record_id, timestamp, outcome, metadata_json
            FROM baptismal_audit_logs
            ORDER BY timestamp DESC
            LIMIT %s
            """,
            (int(limit),),
        )

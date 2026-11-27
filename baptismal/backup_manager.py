"""
baptismal/backup_manager.py — Encrypted Workbook & Metadata Backup and Recovery
===============================================================================
Creates encrypted, owner-only (chmod 0600) timestamped local backups before workbook updates
so sensitive sacramental records never leak in plaintext on disk.
"""

from datetime import datetime, timezone
import json
import os
from typing import Any, Dict, Optional

from .audit_logs import AuditLogger
from .database import DatabaseManager
from .security_crypto import (
    decrypt_sensitive_token_to_bytes,
    encrypt_sensitive_bytes_to_token,
    harden_directory_permissions,
    harden_file_permissions,
)


class BackupManager:
    """Manages encrypted pre-update Excel backups and encrypted database JSON snapshots."""

    def __init__(self, db: DatabaseManager, backup_dir: str = "baptismal_backups", audit: Optional[AuditLogger] = None):
        self.db = db
        self.backup_dir = backup_dir
        self.audit = audit or AuditLogger(db)

    def save_local_workbook_backup(self, workbook_bytes: bytes, label: str = "pre_update", actor_id: str = "system") -> str:
        os.makedirs(self.backup_dir, exist_ok=True)
        harden_directory_permissions(self.backup_dir)
        ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.backup_dir, f"baptismal_workbook_{label}_{ts}.enc.xlsx")
        enc_token = encrypt_sensitive_bytes_to_token(workbook_bytes)
        with open(path, "w", encoding="utf-8") as f:
            f.write(enc_token)
        harden_file_permissions(path)
        self.audit.log_event(actor_id, "ENCRYPTED_BACKUP_CREATED", "SUCCESS", metadata={"backup_path": os.path.basename(path)})
        return path

    @staticmethod
    def read_local_workbook_backup(path: str) -> bytes:
        with open(path, "rb") as f:
            raw = f.read()
        if raw.startswith(b"CTG_ENC_V1:"):
            return decrypt_sensitive_token_to_bytes(raw.decode("ascii"))
        return raw

    def export_metadata_snapshot(self, actor_id: str = "system") -> Dict[str, Any]:
        snapshot = {
            "exported_at": datetime.now(timezone.utc).isoformat(),
            "new_baptism_registrations": self.db.fetch_all("SELECT * FROM new_baptism_registrations"),
            "historical_image_associations": self.db.fetch_all("SELECT * FROM historical_image_associations"),
            "generated_certificates": self.db.fetch_all("SELECT * FROM generated_certificates"),
            "certificate_requests": self.db.fetch_all("SELECT * FROM certificate_requests"),
        }
        os.makedirs(self.backup_dir, exist_ok=True)
        harden_directory_permissions(self.backup_dir)
        ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        snap_path = os.path.join(self.backup_dir, f"baptismal_metadata_snapshot_{ts}.enc.json")
        raw_json_bytes = json.dumps(snapshot, ensure_ascii=False, indent=2).encode("utf-8")
        enc_token = encrypt_sensitive_bytes_to_token(raw_json_bytes)
        with open(snap_path, "w", encoding="utf-8") as f:
            f.write(enc_token)
        harden_file_permissions(snap_path)
        self.audit.log_event(actor_id, "ENCRYPTED_METADATA_SNAPSHOT_EXPORTED", "SUCCESS", metadata={"snapshot": os.path.basename(snap_path)})
        return snapshot

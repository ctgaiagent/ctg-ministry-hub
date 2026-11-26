"""
baptismal/backup_manager.py — Workbook & Metadata Backup and Recovery
=====================================================================
Creates timestamped local/Drive backups before workbook updates and supports
restoring historical metadata snapshots safely.
"""

from datetime import datetime, timezone
import json
import os
from typing import Any, Dict, Optional

from .audit_logs import AuditLogger
from .database import DatabaseManager


class BackupManager:
    """Manages pre-update Excel backups and database JSON snapshots."""

    def __init__(self, db: DatabaseManager, backup_dir: str = "baptismal_backups", audit: Optional[AuditLogger] = None):
        self.db = db
        self.backup_dir = backup_dir
        self.audit = audit or AuditLogger(db)

    def save_local_workbook_backup(self, workbook_bytes: bytes, label: str = "pre_update", actor_id: str = "system") -> str:
        os.makedirs(self.backup_dir, exist_ok=True)
        ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.backup_dir, f"baptismal_workbook_{label}_{ts}.xlsx")
        with open(path, "wb") as f:
            f.write(workbook_bytes)
        self.audit.log_event(actor_id, "BACKUP_CREATED", "SUCCESS", metadata={"backup_path": os.path.basename(path)})
        return path

    def export_metadata_snapshot(self, actor_id: str = "system") -> Dict[str, Any]:
        snapshot = {
            "exported_at": datetime.now(timezone.utc).isoformat(),
            "new_baptism_registrations": self.db.fetch_all("SELECT * FROM new_baptism_registrations"),
            "historical_image_associations": self.db.fetch_all("SELECT * FROM historical_image_associations"),
            "generated_certificates": self.db.fetch_all("SELECT * FROM generated_certificates"),
            "certificate_requests": self.db.fetch_all("SELECT * FROM certificate_requests"),
        }
        os.makedirs(self.backup_dir, exist_ok=True)
        ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        snap_path = os.path.join(self.backup_dir, f"baptismal_metadata_snapshot_{ts}.json")
        with open(snap_path, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, ensure_ascii=False, indent=2)
        self.audit.log_event(actor_id, "METADATA_SNAPSHOT_EXPORTED", "SUCCESS", metadata={"snapshot": os.path.basename(snap_path)})
        return snapshot

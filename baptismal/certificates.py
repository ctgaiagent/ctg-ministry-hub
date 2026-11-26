"""
baptismal/certificates.py — Certificate Generation, Archive & Reissue Tracking
==============================================================================
Generates PDF certificates (checking server-side permission), archives metadata
(Certificate ID, Record ID, Issued By, Issuance Date, Original/Reissued status,
Reissue Reason, PDF SHA-256, and Drive File ID), and uploads private PDFs to
`GENERATED CERTIFICATES/<year>/`.
"""

from datetime import datetime, timezone
import hashlib
import secrets
from typing import Any, Dict, List, Optional

from templates.baptismal_certificate import render_baptismal_certificate_pdf

from .audit_logs import AuditLogger
from .authorization import AuthorizationService
from .database import DatabaseManager
from .google_drive import GoogleDriveManager
from .historical_records import HistoricalRecordsService


class CertificateService:
    """Manages authorized Certificate of Baptism generation and archival."""

    def __init__(
        self,
        db: DatabaseManager,
        authz: AuthorizationService,
        audit: Optional[AuditLogger] = None,
        drive: Optional[GoogleDriveManager] = None,
    ):
        self.db = db
        self.authz = authz
        self.audit = audit or AuditLogger(db)
        self.drive = drive
        self.records_svc = HistoricalRecordsService(db)

    def generate_certificate(
        self,
        actor_account_id: str,
        record_id: str,
        orientation: str = "portrait",
        is_reissue: bool = False,
        reissue_reason: str = "",
        custom_heading: str = "CERTIFICATE OF BAPTISM",
        custom_wording: str = "This is to certify that the person named below received the Sacrament of Christian Baptism according to the rites of the Church of the Temple of God.",
        upload_to_drive: bool = True,
    ) -> Dict[str, Any]:
        self.authz.require_record_permission(actor_account_id, record_id, "generate_certificate")

        record = self.records_svc.get_record_by_id(record_id)
        if not record:
            raise ValueError(f"Baptismal record '{record_id}' not found.")

        if is_reissue and not str(reissue_reason or "").strip():
            raise ValueError("Reissue reason is required when reissuing a certificate.")

        now = datetime.now(timezone.utc)
        cert_id = f"CTG-CERT-{now.strftime('%Y%m%d')}-{secrets.token_hex(3).upper()}"
        issuance_date = now.strftime("%Y-%m-%d")
        issuance_status = "REISSUED" if is_reissue else "ORIGINAL"

        pdf_bytes = render_baptismal_certificate_pdf(
            record=record,
            certificate_id=cert_id,
            issuance_date=issuance_date,
            orientation=orientation,
            custom_heading=custom_heading,
            custom_wording=custom_wording,
        )
        pdf_sha256 = hashlib.sha256(pdf_bytes).hexdigest()

        drive_file_id = ""
        if upload_to_drive and self.drive is not None:
            yr = str(record.get("Year of Baptism") or now.year).strip()
            if yr not in ("2026", "2027", "2028"):
                yr = str(now.year) if str(now.year) in ("2026", "2027", "2028") else "2026"
            folder_tree = self.drive.ensure_baptismal_folder_tree()
            cert_folder_id = folder_tree.get(f"GENERATED CERTIFICATES/{yr}") or folder_tree.get("GENERATED CERTIFICATES", "")
            if cert_folder_id:
                up_res = self.drive.upload_private_file(
                    pdf_bytes,
                    f"{cert_id}_{record_id}.pdf",
                    "application/pdf",
                    cert_folder_id,
                )
                drive_file_id = str(up_res.get("id", ""))

        self.db.execute(
            """
            INSERT INTO generated_certificates (
                certificate_id, record_id, issued_by, issuance_date,
                issuance_status, reissue_reason, drive_file_id, pdf_sha256,
                layout_orientation, created_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                cert_id,
                record_id,
                actor_account_id,
                issuance_date,
                issuance_status,
                reissue_reason.strip(),
                drive_file_id,
                pdf_sha256,
                orientation.lower(),
                now.isoformat(),
            ),
        )
        self.audit.log_event(
            actor_account_id,
            "CERTIFICATE_REISSUED" if is_reissue else "CERTIFICATE_GENERATED",
            "SUCCESS",
            related_record_id=record_id,
            metadata={"certificate_id": cert_id, "status": issuance_status},
        )
        return {
            "certificate_id": cert_id,
            "record_id": record_id,
            "issuance_date": issuance_date,
            "issuance_status": issuance_status,
            "drive_file_id": drive_file_id,
            "pdf_sha256": pdf_sha256,
            "pdf_bytes": pdf_bytes,
        }

    def list_certificates(self, record_id: str = "") -> List[Dict[str, Any]]:
        if record_id:
            return self.db.fetch_all(
                "SELECT * FROM generated_certificates WHERE record_id = %s ORDER BY created_at DESC",
                (record_id,),
            )
        return self.db.fetch_all("SELECT * FROM generated_certificates ORDER BY created_at DESC")

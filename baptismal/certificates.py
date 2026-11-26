"""
baptismal/certificates.py — Master-Admin-Controlled Certificate Issuance, Release & Direct User Sharing
=======================================================================================================
Supports:
1. Official Certificate Generation & Print (A4 Portrait / Landscape)
2. 7-Step Request Approval & Release (`approve_and_release_certificate_for_request`)
3. 1-Click Direct Sharing to Any Username (`share_record_and_certificate_with_username`) so
   Master Admin can share any record & printable certificate directly with any Parish Leader Admin
   username for them to view, download, and print.
"""

from datetime import datetime, timedelta, timezone
import hashlib
import json
import secrets
from typing import Any, Dict, List, Optional

from templates.baptismal_certificate import render_baptismal_certificate_pdf

from .audit_logs import AuditLogger
from .authorization import AuthorizationError, AuthorizationService
from .database import DatabaseManager
from .google_drive import GoogleDriveManager
from .historical_records import HistoricalRecordsService, resolve_quick_record_id


class CertificateService:
    """Manages Master-Admin-controlled Certificate of Baptism generation, issuance, release, and username sharing."""

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
        self.records_svc = HistoricalRecordsService(db, authz=authz)

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
        request_id: str = "",
        authorized_account_id: str = "",
        assigned_parish: str = "",
        release_to_requester: bool = False,
    ) -> Dict[str, Any]:
        resolved_rec_id = resolve_quick_record_id(record_id)
        if self.authz.is_parish_leader_admin(actor_account_id):
            self.audit.log_event(
                actor_account_id,
                "ACCESS_DENIED",
                "FAILURE",
                related_record_id=resolved_rec_id,
                metadata={"reason": "parish_admin_cannot_generate_or_issue_certificates"},
            )
            raise AuthorizationError(
                "Access denied: Parish Leader Admins cannot generate or issue official certificates independently."
            )

        self.authz.require_record_permission(actor_account_id, resolved_rec_id, "generate_certificate")

        record = self.records_svc.get_record_by_id(resolved_rec_id)
        if not record:
            raise ValueError(f"Baptismal record '{resolved_rec_id}' not found.")

        if is_reissue and not str(reissue_reason or "").strip():
            raise ValueError("Reissue reason is required when reissuing a certificate.")

        now = datetime.now(timezone.utc)
        cert_id = f"CTG-CERT-{now.strftime('%Y%m%d')}-{secrets.token_hex(3).upper()}"
        issuance_date = now.strftime("%Y-%m-%d")
        issuance_status = "REISSUED" if is_reissue else "ORIGINAL"
        is_master = self.authz.is_master_admin(actor_account_id)

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
                    f"{cert_id}_{resolved_rec_id}.pdf",
                    "application/pdf",
                    cert_folder_id,
                )
                drive_file_id = str(up_res.get("id", ""))

        rel_flag = 1 if (is_master and release_to_requester) else 0
        eff_parish = assigned_parish.strip() or str(record.get("Parish") or "").strip()

        self.db.execute(
            """
            INSERT INTO generated_certificates (
                certificate_id, record_id, request_id, authorized_account_id,
                assigned_parish, issued_by, issuance_date, issuance_status,
                approval_status, approved_by, approved_at, is_released_to_requester,
                reissue_reason, drive_file_id, pdf_sha256, layout_orientation, created_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                cert_id,
                resolved_rec_id,
                request_id.strip() or None,
                authorized_account_id.strip() or None,
                eff_parish,
                actor_account_id,
                issuance_date,
                issuance_status,
                "APPROVED_AND_ISSUED" if is_master else "GENERATED",
                actor_account_id if is_master else None,
                now.isoformat() if is_master else None,
                rel_flag,
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
            related_record_id=resolved_rec_id,
            metadata={"certificate_id": cert_id, "status": issuance_status, "released": bool(rel_flag)},
        )
        return {
            "certificate_id": cert_id,
            "record_id": resolved_rec_id,
            "issuance_date": issuance_date,
            "issuance_status": issuance_status,
            "is_released_to_requester": bool(rel_flag),
            "drive_file_id": drive_file_id,
            "pdf_sha256": pdf_sha256,
            "pdf_bytes": pdf_bytes,
        }

    def share_record_and_certificate_with_username(
        self,
        admin_account_id: str,
        record_id: str,
        target_username: str,
        duration_hours: float = 72.0,
        orientation: str = "portrait",
        upload_to_drive: bool = False,
    ) -> Dict[str, Any]:
        """Allows Master Admin to directly share/release any record and its official printable
        certificate with another username (such as any Parish Leader Admin) so they can view & print it."""
        self.authz.require_master_admin(admin_account_id, action="share_certificate_with_username")
        resolved_rec_id = resolve_quick_record_id(record_id)
        uname_clean = str(target_username or "").strip().lower()
        target_acc = self.db.fetch_one(
            "SELECT account_id, username, assigned_parish FROM baptismal_accounts WHERE username = %s AND is_active = 1",
            (uname_clean,),
        )
        if not target_acc:
            raise ValueError(f"Target username '{target_username}' not found or inactive.")

        cert_res = self.generate_certificate(
            actor_account_id=admin_account_id,
            record_id=resolved_rec_id,
            orientation=orientation,
            upload_to_drive=upload_to_drive,
            authorized_account_id=target_acc["account_id"],
            assigned_parish=str(target_acc.get("assigned_parish") or ""),
            release_to_requester=True,
        )

        now = datetime.now(timezone.utc)
        expires_at = (now + timedelta(hours=float(duration_hours))).isoformat()
        grant_id = f"GRANT-{secrets.token_hex(5).upper()}"
        perms = ["view_record", "view_image", "download_certificate", "print_certificate"]
        self.db.execute(
            """
            INSERT INTO temporary_access_grants (
                grant_id, request_id, account_id, record_id, permissions_json,
                expires_at, is_revoked, revoked_by, revoked_at, granted_by, created_at
            ) VALUES (%s, NULL, %s, %s, %s, %s, 0, NULL, NULL, %s, %s)
            """,
            (
                grant_id,
                target_acc["account_id"],
                resolved_rec_id,
                json.dumps(perms),
                expires_at,
                admin_account_id,
                now.isoformat(),
            ),
        )
        self.audit.log_event(
            admin_account_id,
            "CERTIFICATE_SHARED_WITH_USERNAME",
            "SUCCESS",
            related_record_id=resolved_rec_id,
            metadata={
                "certificate_id": cert_res["certificate_id"],
                "grant_id": grant_id,
                "target_username": target_acc["username"],
            },
        )
        return {
            **cert_res,
            "grant_id": grant_id,
            "target_username": target_acc["username"],
            "expires_at": expires_at,
        }

    def approve_and_release_certificate_for_request(
        self,
        admin_account_id: str,
        request_id: str,
        record_id: str,
        orientation: str = "portrait",
        review_notes: str = "",
        upload_to_drive: bool = False,
    ) -> Dict[str, Any]:
        self.authz.require_master_admin(admin_account_id, action="approve_and_release_certificate")
        req = self.db.fetch_one("SELECT * FROM certificate_requests WHERE request_id = %s", (request_id,))
        if not req:
            raise ValueError(f"Certificate request '{request_id}' not found.")

        resolved_rec_id = resolve_quick_record_id(record_id)
        cert_res = self.generate_certificate(
            actor_account_id=admin_account_id,
            record_id=resolved_rec_id,
            orientation=orientation,
            upload_to_drive=upload_to_drive,
            request_id=request_id,
            authorized_account_id=req["requester_account_id"],
            assigned_parish=req.get("assigned_parish", ""),
            release_to_requester=True,
        )
        now_iso = datetime.now(timezone.utc).isoformat()
        self.db.execute(
            """
            UPDATE certificate_requests
            SET status = 'Approved', target_record_id = %s, approved_certificate_id = %s,
                reviewed_by = %s, reviewed_at = %s, review_notes = %s, updated_at = %s
            WHERE request_id = %s
            """,
            (
                resolved_rec_id,
                cert_res["certificate_id"],
                admin_account_id,
                now_iso,
                review_notes.strip() or f"Certificate {cert_res['certificate_id']} officially approved and released.",
                now_iso,
                request_id,
            ),
        )
        self.audit.log_event(
            admin_account_id,
            "CERTIFICATE_APPROVED_AND_RELEASED",
            "SUCCESS",
            related_record_id=resolved_rec_id,
            metadata={
                "request_id": request_id,
                "certificate_id": cert_res["certificate_id"],
                "released_to": req["requester_account_id"],
                "parish": req.get("assigned_parish", ""),
            },
        )
        return cert_res

    def download_released_certificate(self, actor_account_id: str, certificate_id: str) -> bytes:
        cert = self.db.fetch_one(
            "SELECT * FROM generated_certificates WHERE certificate_id = %s",
            (certificate_id,),
        )
        if not cert:
            raise ValueError(f"Certificate '{certificate_id}' not found.")

        if not self.authz.is_master_admin(actor_account_id):
            acc = self.authz.get_live_account(actor_account_id)
            if not acc:
                raise AuthorizationError("Access denied: unauthenticated account.")
            is_released = int(cert.get("is_released_to_requester", 0)) == 1
            is_target_user = cert.get("authorized_account_id") == actor_account_id
            if acc["role"] == "Parish Leader Admin":
                actor_parish = self.authz.get_authenticated_parish(actor_account_id)
                is_same_parish = cert.get("assigned_parish") == actor_parish
                if not (is_released and is_target_user and is_same_parish):
                    self.audit.log_event(
                        actor_account_id,
                        "ACCESS_DENIED",
                        "FAILURE",
                        related_record_id=cert.get("record_id"),
                        metadata={"certificate_id": certificate_id, "reason": "unreleased_or_wrong_parish_certificate"},
                    )
                    raise AuthorizationError("Access denied: certificate has not been approved and released to your parish account.")
            else:
                self.authz.require_record_permission(actor_account_id, cert["record_id"], "download_certificate")

        record = self.records_svc.get_record_by_id(cert["record_id"])
        if not record:
            raise ValueError("Source record for certificate not found.")
        pdf_bytes = render_baptismal_certificate_pdf(
            record=record,
            certificate_id=cert["certificate_id"],
            issuance_date=cert["issuance_date"],
            orientation=cert.get("layout_orientation", "portrait"),
        )
        self.audit.log_event(
            actor_account_id,
            "CERTIFICATE_DOWNLOADED",
            "SUCCESS",
            related_record_id=cert["record_id"],
            metadata={"certificate_id": certificate_id},
        )
        return pdf_bytes

    def list_certificates(self, record_id: str = "", actor_account_id: str = "") -> List[Dict[str, Any]]:
        if actor_account_id and not self.authz.is_master_admin(actor_account_id):
            if self.authz.is_parish_leader_admin(actor_account_id):
                parish = self.authz.get_authenticated_parish(actor_account_id)
                return self.db.fetch_all(
                    """
                    SELECT * FROM generated_certificates
                    WHERE authorized_account_id = %s AND assigned_parish = %s AND is_released_to_requester = 1
                    ORDER BY created_at DESC
                    """,
                    (actor_account_id, parish),
                )
            return []
        if record_id:
            return self.db.fetch_all(
                "SELECT * FROM generated_certificates WHERE record_id = %s ORDER BY created_at DESC",
                (resolve_quick_record_id(record_id),),
            )
        return self.db.fetch_all("SELECT * FROM generated_certificates ORDER BY created_at DESC")

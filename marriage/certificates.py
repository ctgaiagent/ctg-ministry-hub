"""
marriage/certificates.py — Master-Admin-Only Official Marriage Certificate Generation & Archiving
=================================================================================================
"""

from datetime import datetime, timezone
import json
import os
import uuid
from typing import Any, Dict, List

from baptismal.audit_logs import AuditLogger
from baptismal.database import DatabaseManager
from marriage.config import DEFAULT_MARRIAGE_ROOT_FOLDER_ID
from marriage.records import assert_master_admin_for_marriage, ensure_marriage_tables
from templates.marriage_certificate import render_marriage_certificate_pdf


def generate_and_archive_marriage_certificate(
    db: DatabaseManager,
    actor: Dict[str, Any],
    record: Dict[str, Any],
    orientation: str = "Portrait",
    custom_heading: str = "CERTIFICATE OF MARRIAGE",
    custom_wording: str = "",
    drive_mgr: Any = None,
    output_dir: str = "",
) -> Dict[str, Any]:
    """Generates an official CTG Certificate of Marriage PDF, saves locally/Drive, and records in DB."""
    assert_master_admin_for_marriage(actor, db=db)
    ensure_marriage_tables(db)

    now = datetime.now(timezone.utc)
    cert_id = f"CTG-MAR-CERT-{now.strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}"
    issuance_date = now.strftime("%B %d, %Y")

    kwargs: Dict[str, Any] = {
        "record": record,
        "certificate_id": cert_id,
        "issuance_date": issuance_date,
        "orientation": orientation,
        "custom_heading": custom_heading or "CERTIFICATE OF MARRIAGE",
    }
    if custom_wording and custom_wording.strip():
        kwargs["custom_wording"] = custom_wording.strip()

    pdf_bytes = render_marriage_certificate_pdf(**kwargs)

    husband = str(record.get("Husband Name") or "Groom").strip()
    wife = str(record.get("Wife Name") or "Bride").strip()
    rec_id = str(record.get("Record ID") or "MAR-00000").strip()
    parish = str(record.get("Parish") or "Caba Parish").strip()
    dom = str(record.get("Date of Marriage") or "").strip()

    local_pdf_path = ""
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        safe_fn = f"{cert_id}_{rec_id}.pdf".replace(" ", "_")
        local_pdf_path = os.path.join(output_dir, safe_fn)
        try:
            with open(local_pdf_path, "wb") as fp:
                fp.write(pdf_bytes)
        except Exception:
            local_pdf_path = ""

    drive_file_id = ""
    if drive_mgr and hasattr(drive_mgr, "upload_private_file"):
        try:
            up_res = drive_mgr.upload_private_file(
                file_bytes=pdf_bytes,
                filename=f"{cert_id}_{husband}_{wife}.pdf",
                mime_type="application/pdf",
                folder_id=DEFAULT_MARRIAGE_ROOT_FOLDER_ID,
            )
            if isinstance(up_res, dict) and up_res.get("id"):
                drive_file_id = str(up_res["id"])
        except Exception:
            drive_file_id = ""

    db.execute(
        """
        INSERT INTO generated_marriage_certificates (
            certificate_id, record_id, husband_name, wife_name,
            parish, date_of_marriage, layout_orientation,
            drive_file_id, local_pdf_path, generated_by_username,
            generated_at, metadata_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            cert_id,
            rec_id,
            husband,
            wife,
            parish,
            dom,
            orientation,
            drive_file_id,
            local_pdf_path,
            str(actor.get("username") or "ctg.baptismal.admin"),
            now.isoformat(),
            json.dumps({"record_id": rec_id, "mc_number": record.get("MC Number", "")}, ensure_ascii=False),
        ),
    )

    AuditLogger(db).log_event(
        account_id=str(actor.get("account_id") or actor.get("username") or "master_admin"),
        event_type="GENERATE_MARRIAGE_CERTIFICATE",
        outcome="SUCCESS",
        related_record_id=rec_id,
        metadata={"certificate_id": cert_id, "husband": husband, "wife": wife, "orientation": orientation},
    )

    return {
        "ok": True,
        "certificate_id": cert_id,
        "issuance_date": issuance_date,
        "pdf_bytes": pdf_bytes,
        "drive_file_id": drive_file_id,
        "local_pdf_path": local_pdf_path,
    }


def list_generated_marriage_certificates(
    db: DatabaseManager,
    actor: Dict[str, Any],
    limit: int = 100,
) -> List[Dict[str, Any]]:
    """Lists generated Marriage Certificates (strictly Master Admin only)."""
    assert_master_admin_for_marriage(actor, db=db)
    ensure_marriage_tables(db)
    return db.fetch_all(
        "SELECT * FROM generated_marriage_certificates ORDER BY generated_at DESC LIMIT ?",
        (int(limit),),
    )

"""
baptismal_app.py — CTG Baptismal Records & Certification System (Role-Based Dashboards)
======================================================================================
Implements:
1. Weekly-Report-Style Horizontal Tabs (`st.tabs`) for fast 1-click page switching.
2. Quick Record ID / Last-Digits Jump (`HIST-` + last number `1..4026` with Prev/Next buttons)
   plus full Record ID & 27-column keyword search.
3. Side-by-Side Record + Book Image Viewer across ALL Books (`BOOK 1`–`BOOK 9`) right inside
   the search tab, with 1-Click Print (`window.print()` + PDF) and 1-Click Share/Release to
   any username (`📤 Ibigay sa Ibang Username para Ma-Print Din Nila`).
4. Distinct 9 Parish Leader Admin Portals (Hong Kong Excluded):
   - Parish Leader Admins upload new baptized members & images/files -> staged in Master Admin Confirmation Queue.
   - Master Admin clicks `✅ Confirm & I-encode sa Google Drive Excel + I-save ang Image/File sa Google Drive`.
   - Parish Leader Admins can submit requests to Master Admin and print/download released certificates & records.
"""

import base64
import inspect
import json
import re
from typing import Any, Dict, List
import streamlit as st
import streamlit.components.v1 as components

from baptismal.audit_logs import AuditLogger
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.backup_manager import BackupManager
from baptismal.certificates import CertificateService
from baptismal.config import (
    BOOK9_EPSON_SCAN_MAP,
    CANONICAL_REQUEST_STATUSES,
    KNOWN_BAPTISMAL_SUBFOLDER_IDS,
    DEFAULT_BAPTISMAL_ROOT_FOLDER_ID,
    DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID,
    EDITABLE_REGISTRATION_STATUSES,
    HISTORICAL_27_COLUMNS,
    HISTORICAL_BOOKS,
    OFFICIAL_NINE_PARISHES,
    REGISTRATION_VERIFICATION_STATUSES,
    SUPPORTED_NEW_YEARS,
    TEMP_ACCESS_PERMISSIONS,
    THEME_GOLD,
    THEME_NAVY,
    THEME_NAVY_DARK,
    THEME_WHITE,
)
from baptismal.database import DatabaseManager
from baptismal.google_drive import GoogleDriveManager, GoogleDrivePermissionError
from baptismal.historical_records import HistoricalRecordsService, resolve_quick_record_id
from baptismal.image_archive import ImageArchiveService, render_image_preview_bytes
from baptismal.image_matching import ImageMatchingService, extract_page_hint_from_filename
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService
from baptismal.temporary_access import TemporaryAccessService
from baptismal.validation import escape_html_text, validate_uploaded_file


_BAP_SERVICES_VERSION = "2026-11-27-v3-book9"


@st.cache_resource
def _get_baptismal_services(_version: str = _BAP_SERVICES_VERSION) -> Dict[str, Any]:
    db = DatabaseManager()
    run_migrations(db)
    audit = AuditLogger(db)
    auth = AuthService(db, audit=audit)
    auth.ensure_initial_master_admin_from_env()
    authz = AuthorizationService(db, audit=audit)
    backup = BackupManager(db, audit=audit)
    drive = GoogleDriveManager()
    hist = HistoricalRecordsService(db, authz=authz)
    img_arc = ImageArchiveService(db, authz=authz)
    img_match = ImageMatchingService(db, authz=authz, audit=audit)
    reg = RegistrationService(db, authz=authz, audit=audit, drive=drive, backup=backup)
    certs = CertificateService(db, authz=authz, audit=audit, drive=drive)
    temp_acc = TemporaryAccessService(db, authz=authz, audit=audit)
    return {
        "db": db,
        "audit": audit,
        "auth": auth,
        "authz": authz,
        "backup": backup,
        "drive": drive,
        "hist": hist,
        "img_arc": img_arc,
        "img_match": img_match,
        "reg": reg,
        "certs": certs,
        "temp_acc": temp_acc,
    }


def _render_printable_certificate_or_image_modal(pdf_bytes: bytes = b"", image_bytes: bytes = b"", title: str = "CTG Baptismal Print") -> None:
    """Renders an inline PDF/Image viewer with a 1-click browser Print dialog button."""
    if pdf_bytes:
        b64_pdf = base64.b64encode(pdf_bytes).decode("ascii")
        components.html(
            f"""
            <div style="margin-bottom:8px;">
                <button onclick="const w=window.open('');w.document.write('<iframe src=\\'data:application/pdf;base64,{b64_pdf}\\' style=\\'width:100%;height:100vh;border:none;\\'></iframe>');setTimeout(()=>w.print(),700);"
                        style="background:#0F2C59;color:#FFFFFF;border:2px solid #D4AF37;padding:8px 18px;border-radius:8px;font-weight:700;cursor:pointer;">
                    🖨️ Print Official Certificate Now ({escape_html_text(title)})
                </button>
            </div>
            <iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="580" style="border:1px solid #CBD5E1;border-radius:8px;"></iframe>
            """,
            height=640,
        )
    elif image_bytes:
        b64_img = base64.b64encode(image_bytes).decode("ascii")
        components.html(
            f"""
            <div style="margin-bottom:8px;">
                <button onclick="const w=window.open('');w.document.write('<img src=\\'data:image/jpeg;base64,{b64_img}\\' style=\\'max-width:100%;\\'/>');setTimeout(()=>w.print(),400);"
                        style="background:#0F2C59;color:#FFFFFF;border:2px solid #D4AF37;padding:8px 18px;border-radius:8px;font-weight:700;cursor:pointer;">
                    🖨️ Print Scanned Book Page Now ({escape_html_text(title)})
                </button>
            </div>
            """,
            height=55,
        )


def _find_candidate_drive_images_for_record(record: Dict[str, Any], drive_mgr: GoogleDriveManager) -> List[Dict[str, Any]]:
    """Searches the record's Book folder and all BOOK 1-9 folders in Google Drive for scanned images
    matching the child's name, record ID, or page number (including BOOK 9 Epson scans and +/- 2 page mismatch checks)."""
    rec_id = str(record.get("Record ID") or "").strip().upper()
    book_num = str(record.get("Book Number") or "").strip()
    page_num = str(record.get("Page Number") or "").strip()
    child_name = str(record.get("Name of the Child") or "").strip().lower()
    name_tokens = [t for t in re.split(r"[^a-z0-9]+", child_name) if len(t) >= 3]

    candidates: List[Dict[str, Any]] = []
    seen_ids = set()

    # Check BOOK 9 verified Epson scan map first (covers HIST-04008..HIST-04026 / Pages 4146..4164)
    for b9_fn, b9_meta in BOOK9_EPSON_SCAN_MAP.items():
        b9_rec_id = b9_meta.get("record_id", "")
        b9_pg = b9_meta.get("page_number", "")
        b9_child = b9_meta.get("child_name", "").lower()
        b9_score = 0
        if rec_id and rec_id == b9_rec_id:
            b9_score += 20
        if page_num and page_num == b9_pg:
            b9_score += 15
        if name_tokens and all(tok in b9_child for tok in name_tokens):
            b9_score += 12
        if b9_score >= 12:
            drv_id = b9_meta["drive_id"]
            seen_ids.add(drv_id)
            candidates.append({
                "id": drv_id,
                "name": f"{b9_fn} — {b9_meta['child_name']}",
                "raw_name": b9_fn,
                "book_folder": "BOOK 9",
                "fn_page_hint": b9_pg,
                "match_score": b9_score,
                "page_mismatch": bool(page_num and b9_pg and page_num != b9_pg),
            })

    book_folders = drive_mgr.discover_book_folders_in_drive()
    target_books: List[str] = []
    if book_num and f"BOOK {book_num}" in book_folders:
        target_books.append(f"BOOK {book_num}")
    if "BOOK 9" in book_folders and "BOOK 9" not in target_books:
        target_books.append("BOOK 9")
    for bk in sorted(book_folders.keys()):
        if bk not in target_books:
            target_books.append(bk)

    for bk_name in target_books:
        fid = book_folders.get(bk_name, "")
        if not fid:
            continue
        cache_k = f"_drv_book_files_{bk_name}"
        if cache_k not in st.session_state and bk_name in (f"BOOK {book_num}", "BOOK 9"):
            try:
                st.session_state[cache_k] = drive_mgr.list_files_in_folder(fid, page_size=1000)
            except Exception:
                st.session_state[cache_k] = []
        files_in_bk = st.session_state.get(cache_k, [])
        for f_it in files_in_bk:
            f_id = str(f_it.get("id") or "")
            if f_id in seen_ids:
                continue
            fn = str(f_it.get("name") or "")
            fn_low = fn.lower()
            fn_page = extract_page_hint_from_filename(fn)
            b9_info = BOOK9_EPSON_SCAN_MAP.get(fn)
            b9_child_low = b9_info["child_name"].lower() if b9_info else ""
            score = 0
            if name_tokens and (all(tok in fn_low for tok in name_tokens) or (b9_child_low and all(tok in b9_child_low for tok in name_tokens))):
                score += 10
            elif name_tokens and (any(tok in fn_low for tok in name_tokens) or (b9_child_low and any(tok in b9_child_low for tok in name_tokens))):
                score += 4
            if page_num and fn_page:
                if fn_page == page_num:
                    score += 8
                elif page_num.isdigit() and fn_page.isdigit() and abs(int(page_num) - int(fn_page)) <= 2:
                    score += 5
            if score >= 5 or (bk_name == f"BOOK {book_num}" and score >= 4):
                seen_ids.add(f_id)
                disp_name = f"{fn} — {b9_info['child_name']}" if b9_info else fn
                candidates.append({
                    **f_it,
                    "name": disp_name,
                    "raw_name": fn,
                    "book_folder": bk_name,
                    "fn_page_hint": fn_page,
                    "match_score": score,
                    "page_mismatch": bool(page_num and fn_page and page_num != fn_page),
                })
    candidates.sort(key=lambda x: x["match_score"], reverse=True)
    return candidates


def _render_parish_leader_admin_dashboard(
    account_id: str,
    current_user: Dict[str, Any],
    svc: Dict[str, Any],
) -> None:
    """Renders the distinct Parish Leader Admin Portal using horizontal tabs (like Weekly Report)."""
    auth_svc: AuthService = svc["auth"]
    authz_svc: AuthorizationService = svc["authz"]
    hist_svc: HistoricalRecordsService = svc["hist"]
    reg_svc: RegistrationService = svc["reg"]
    cert_svc: CertificateService = svc["certs"]
    temp_svc: TemporaryAccessService = svc["temp_acc"]

    assigned_parish = authz_svc.get_authenticated_parish(account_id)

    st.markdown(
        f"<div class='ctg-bap-banner'>"
        f"<h2>⛪ {escape_html_text(assigned_parish)} — Parish Leader Admin Baptismal Portal</h2>"
        f"<p>Upload New Baptized Members for Master Admin Confirmation • Request Historical Records & Certified Certificates • Print Released Certificates</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    p_stats = reg_svc.get_parish_dashboard_statistics(account_id)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("My New Baptisms", p_stats["my_new_registrations"])
    k2.metric("Pending Admin Confirm", p_stats["pending_verification"])
    k3.metric("Confirmed & Encoded", p_stats["approved_registrations"])
    k4.metric("Rejected", p_stats["rejected_registrations"])
    k5.metric("Pending Requests", p_stats["pending_certificate_requests"])
    k6.metric("Ready to Print", p_stats["approved_certificates"])

    # Horizontal Tabs just like Weekly Report & Parish Monitoring
    ptab_upload, ptab_subs, ptab_req_hist, ptab_req_cert, ptab_print, ptab_acc = st.tabs([
        "📤 1. Mag-upload ng Bagong Baptized (Ipasa sa Master Admin)",
        "📂 2. My Submissions & Uploaded Images",
        "📜 3. Request Historical Record",
        "🏅 4. Request Certified Certificate",
        "🖨️ 5. My Approved Certificates & Records (Print / Download)",
        "🔐 6. My Account",
    ])

    with ptab_upload:
        st.subheader(f"📤 Mag-upload ng Bagong Baptized Member — {assigned_parish}")
        st.info(
            "📌 **Workflow:** Kapag nag-submit ka rito kasama ang Baptismal Image/File, mapupunta muna ito sa **Master Admin Confirmation Queue**. "
            "Kapag na-confirm na ng Master Admin, awtomatiko itong ie-encode sa `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` at ise-save ang image/file sa Google Drive!"
        )
        with st.form("parish_new_baptism_form"):
            payload: Dict[str, str] = {}
            cols = st.columns(3)
            for idx, col_name in enumerate(HISTORICAL_27_COLUMNS):
                with cols[idx % 3]:
                    if col_name == "Parish":
                        st.text_input("Parish (Locked Server-Side):", value=assigned_parish, disabled=True)
                        payload["Parish"] = assigned_parish
                    elif col_name == "Year of Baptism":
                        payload[col_name] = st.selectbox(col_name, SUPPORTED_NEW_YEARS)
                    else:
                        payload[col_name] = st.text_input(col_name, key=f"p_reg_{col_name}")
            up_imgs = st.file_uploader(
                "📸 I-upload ang Baptismal Image(s) o Supporting File(s) (JPG, PNG, WEBP, PDF):",
                type=["jpg", "jpeg", "png", "webp", "pdf"],
                accept_multiple_files=True,
            )
            if st.form_submit_button("📤 Ipasa sa Master Admin para I-confirm at I-encode sa Google Drive", type="primary"):
                img_list = [
                    {"bytes": uf.read(), "filename": uf.name, "mime_type": uf.type or "image/jpeg"}
                    for uf in (up_imgs or [])
                ]
                try:
                    res = reg_svc.register_new_baptism(
                        account_id, payload, uploaded_images=img_list, sync_to_drive_excel=False
                    )
                    st.success(
                        f"✅ Naipasa na sa Master Admin! Record ID: `{res['record_id']}` • Parish: `{res['parish']}` • "
                        f"Status: `{res['verification_status']}` (`{res['workbook_update_status']}`)"
                    )
                    st.rerun()
                except Exception as ex:
                    st.error(str(ex))

    with ptab_subs:
        st.subheader(f"📂 My Submissions & Staged Images — {assigned_parish}")
        subs = reg_svc.list_submissions_for_actor(account_id)
        if not subs:
            st.info("Wala pang naipapasang new baptism submission.")
        else:
            st.dataframe(subs, use_container_width=True, hide_index=True)
            editable_subs = [s for s in subs if s.get("verification_status") in EDITABLE_REGISTRATION_STATUSES]
            if editable_subs:
                with st.expander("✏️ I-edit ang Submission Habang Pending Verification Pa"):
                    sel_edit_id = st.selectbox("Piliin ang Pending Record ID:", [s["record_id"] for s in editable_subs])
                    target_sub = next(s for s in editable_subs if s["record_id"] == sel_edit_id)
                    cur_payload = json.loads(target_sub["payload_json"])
                    with st.form("edit_pending_submission_form"):
                        upd_payload: Dict[str, str] = {}
                        cols_e = st.columns(3)
                        for idx, col_name in enumerate(HISTORICAL_27_COLUMNS):
                            with cols_e[idx % 3]:
                                if col_name == "Parish":
                                    st.text_input("Parish (Locked):", value=assigned_parish, disabled=True)
                                    upd_payload["Parish"] = assigned_parish
                                elif col_name == "Year of Baptism":
                                    cur_y = cur_payload.get("Year of Baptism", SUPPORTED_NEW_YEARS[0])
                                    y_idx = SUPPORTED_NEW_YEARS.index(cur_y) if cur_y in SUPPORTED_NEW_YEARS else 0
                                    upd_payload[col_name] = st.selectbox(col_name, SUPPORTED_NEW_YEARS, index=y_idx)
                                else:
                                    upd_payload[col_name] = st.text_input(
                                        col_name, value=str(cur_payload.get(col_name, "")), key=f"ed_f_{col_name}"
                                    )
                        if st.form_submit_button("💾 I-save ang Pagbabago sa Pending Submission", type="primary"):
                            try:
                                reg_svc.update_own_submission(account_id, sel_edit_id, upd_payload)
                                st.success(f"Updated pending submission `{sel_edit_id}`.")
                                st.rerun()
                            except Exception as ex:
                                st.error(str(ex))

        imgs = reg_svc.list_uploaded_images_for_actor(account_id)
        if imgs:
            st.markdown("#### ☁️ Uploaded Images Confirmed in Google Drive")
            st.dataframe(imgs, use_container_width=True, hide_index=True)

    with ptab_req_hist:
        st.subheader(f"📜 Request Historical Baptismal Record sa Master Admin — {assigned_parish}")
        with st.form("req_hist_record_form"):
            c_h1, c_h2 = st.columns(2)
            with c_h1:
                child_nm = st.text_input("Name of Baptized Person (Required):")
                father_nm = st.text_input("Father's Name (if known):")
                mother_nm = st.text_input("Mother's Name (if known):")
                approx_bap = st.text_input("Approximate Baptism Date:")
            with c_h2:
                approx_dob = st.text_input("Approximate Birth Date:")
                st.text_input("Parish (Locked to Your Parish):", value=assigned_parish, disabled=True)
                bk_hint = st.text_input("Book Number (if known):")
                pg_hint = st.text_input("Page Number (if known):")
            reason_txt = st.text_area("Reason for Request (Required):")
            sup_doc = st.file_uploader("Supporting Document (Optional — JPG, PNG, PDF):", type=["jpg", "jpeg", "png", "pdf"])
            if st.form_submit_button("📨 Ipadala ang Historical Record Request sa Master Admin", type="primary"):
                sup_fn = ""
                if sup_doc is not None:
                    ok_f, msg_f, safe_fn = validate_uploaded_file(sup_doc.read(), sup_doc.name)
                    if not ok_f:
                        st.error(msg_f)
                        st.stop()
                    sup_fn = safe_fn
                try:
                    rid = temp_svc.submit_request(
                        requester_account_id=account_id,
                        requester_username=current_user["username"],
                        requested_person_name=child_nm,
                        purpose=reason_txt,
                        request_type="HISTORICAL_RECORD",
                        father_name=father_nm,
                        mother_name=mother_nm,
                        approx_baptism_date=approx_bap,
                        approx_birth_date=approx_dob,
                        parish=assigned_parish,
                        book_hint=bk_hint,
                        page_hint=pg_hint,
                        supporting_doc_filename=sup_fn,
                    )
                    st.success(f"✅ Naipadala ang Historical Record Request sa Master Admin! Request ID: `{rid}`")
                except Exception as ex:
                    st.error(str(ex))

    with ptab_req_cert:
        st.subheader(f"🏅 Request Certified Baptismal Certificate sa Master Admin — {assigned_parish}")
        with st.form("req_cert_copy_form"):
            c_c1, c_c2 = st.columns(2)
            with c_c1:
                c_child = st.text_input("Name of Baptized Person (Required):")
                c_father = st.text_input("Father's Name (if known):")
                c_mother = st.text_input("Mother's Name (if known):")
                c_bap_dt = st.text_input("Approximate Baptism Date:")
            with c_c2:
                c_dob_dt = st.text_input("Approximate Birth Date:")
                st.text_input("Parish (Locked):", value=assigned_parish, disabled=True)
                c_rec_id = st.text_input("Known Record ID (Optional, e.g. HIST-00743 or BAP-2026-...):")
                c_bk_pg = st.text_input("Book / Page Number (if known):")
            c_purpose = st.text_area("Purpose / Reason for Certified Certificate (Required):")
            c_sup = st.file_uploader("Supporting Document (Optional):", type=["jpg", "jpeg", "png", "pdf"])
            if st.form_submit_button("📨 Ipadala ang Request para sa Official Certificate", type="primary"):
                sup_fn = ""
                if c_sup is not None:
                    ok_f, msg_f, safe_fn = validate_uploaded_file(c_sup.read(), c_sup.name)
                    if not ok_f:
                        st.error(msg_f)
                        st.stop()
                    sup_fn = safe_fn
                try:
                    rid = temp_svc.submit_request(
                        requester_account_id=account_id,
                        requester_username=current_user["username"],
                        requested_person_name=c_child,
                        purpose=c_purpose,
                        target_record_id=c_rec_id,
                        request_type="CERTIFIED_CERTIFICATE",
                        father_name=c_father,
                        mother_name=c_mother,
                        approx_baptism_date=c_bap_dt,
                        approx_birth_date=c_dob_dt,
                        parish=assigned_parish,
                        book_hint=c_bk_pg,
                        supporting_doc_filename=sup_fn,
                    )
                    st.success(f"✅ Naipadala ang Certified Certificate Request! Request ID: `{rid}`")
                except Exception as ex:
                    st.error(str(ex))

        st.divider()
        st.markdown("#### 📋 My Request History & Status")
        my_reqs = temp_svc.list_requests(actor_account_id=account_id)
        if my_reqs:
            st.dataframe(my_reqs, use_container_width=True, hide_index=True)
        else:
            st.info("Wala pang naipadalang request.")

    with ptab_print:
        st.subheader(f"🖨️ My Approved Certificates & Released Records (Ready to Print) — {assigned_parish}")
        rel_certs = cert_svc.list_certificates(actor_account_id=account_id)
        if not rel_certs:
            st.info("Wala pang opisyal na certificate na na-release o naibigay sa iyong username.")
        else:
            st.dataframe(rel_certs, use_container_width=True, hide_index=True)
            sel_cid = st.selectbox("Piliin ang Released Certificate para I-print o I-download:", [c["certificate_id"] for c in rel_certs])
            if sel_cid:
                try:
                    pdf_bytes = cert_svc.download_released_certificate(account_id, sel_cid)
                    st.download_button(
                        "⬇️ Download Official Certified Baptismal Certificate (PDF)",
                        data=pdf_bytes,
                        file_name=f"{sel_cid}.pdf",
                        mime="application/pdf",
                        type="primary",
                    )
                    _render_printable_certificate_or_image_modal(pdf_bytes=pdf_bytes, title=sel_cid)
                except Exception as ex:
                    st.error(str(ex))

        my_grants = temp_svc.list_active_grants(account_id=account_id)
        if my_grants:
            st.divider()
            st.markdown("#### 📄 Mga Record na Ibinigay ng Master Admin sa Iyong Username")
            for g in my_grants:
                rec_id = g["record_id"]
                if authz_svc.has_temporary_record_permission(account_id, rec_id, "view_record"):
                    rec_data = hist_svc.get_record_by_id(rec_id, actor_account_id=account_id)
                    if rec_data:
                        st.success(f"Released Record `{rec_id}` — **{rec_data.get('Name of the Child', '')}** (Expires: {g['expires_at']})")
                        st.json(rec_data)

    with ptab_acc:
        st.subheader(f"🔐 My Parish Admin Account — {assigned_parish}")
        st.markdown(
            f"- **Username:** `{escape_html_text(current_user['username'])}`\n"
            f"- **Role:** `{escape_html_text(current_user['role'])}`\n"
            f"- **Assigned Parish (Server-Locked):** `{escape_html_text(assigned_parish)}`"
        )
        with st.form("parish_admin_pw_change"):
            np1 = st.text_input("New Password (minimum 8 characters):", type="password")
            np2 = st.text_input("Confirm New Password:", type="password")
            if st.form_submit_button("Update My Password"):
                if len(np1) < 8:
                    st.error("Password must be at least 8 characters.")
                elif np1 != np2:
                    st.error("Passwords do not match.")
                else:
                    auth_svc.reset_password(account_id, np1, actor_account_id=account_id)
                    st.success("Password updated successfully.")


def render_baptismal_portal(embedded_in_hub: bool = False) -> None:
    """Renders the complete Baptismal Records & Certification System."""
    st.markdown(
        f"""
        <style>
        .ctg-bap-banner {{
            background: linear-gradient(135deg, {THEME_NAVY_DARK} 0%, {THEME_NAVY} 100%);
            border-left: 6px solid {THEME_GOLD};
            border-radius: 12px;
            padding: 18px 22px;
            color: {THEME_WHITE};
            margin-bottom: 16px;
            box-shadow: 0 4px 14px rgba(15, 44, 89, 0.16);
        }}
        .ctg-bap-banner h2 {{
            margin: 0;
            color: {THEME_WHITE};
            font-size: 1.45rem;
        }}
        .ctg-bap-banner p {{
            margin: 4px 0 0 0;
            color: #E2E8F0;
            font-size: 0.92rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    svc = _get_baptismal_services(_version=_BAP_SERVICES_VERSION)
    auth_svc: AuthService = svc["auth"]
    authz_svc: AuthorizationService = svc["authz"]
    audit_svc: AuditLogger = svc["audit"]
    hist_svc: HistoricalRecordsService = svc["hist"]
    img_arc_svc: ImageArchiveService = svc["img_arc"]
    img_match_svc: ImageMatchingService = svc["img_match"]
    reg_svc: RegistrationService = svc["reg"]
    cert_svc: CertificateService = svc["certs"]
    temp_svc: TemporaryAccessService = svc["temp_acc"]
    drive_mgr: GoogleDriveManager = svc["drive"]
    backup_mgr: BackupManager = svc["backup"]

    if not auth_svc.has_master_admin():
        st.markdown(
            "<div class='ctg-bap-banner'>"
            "<h2>⛪ Church of the Temple of God — Baptismal Records & Certification System</h2>"
            "<p>Initial Master Admin Provisioning</p>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.warning("Initial Setup Required: No Baptismal Master Admin account exists yet. Create the first Master Admin account below.")
        with st.form("initial_master_admin_setup"):
            setup_u = st.text_input("Baptismal Master Admin Username:")
            setup_fn = st.text_input("Full Name:", value="CTG Baptismal Master Admin")
            setup_em = st.text_input("Email (Optional):")
            setup_p1 = st.text_input("Password (minimum 8 characters):", type="password")
            setup_p2 = st.text_input("Confirm Password:", type="password")
            if st.form_submit_button("Create Master Admin Account", type="primary"):
                if not setup_u.strip() or len(setup_p1) < 8:
                    st.error("Username and a password of at least 8 characters are required.")
                elif setup_p1 != setup_p2:
                    st.error("Passwords do not match.")
                else:
                    auth_svc.create_account(
                        username=setup_u,
                        full_name=setup_fn,
                        role="Master Admin",
                        assigned_parish="ALL_NINE_PARISHES",
                        plain_password=setup_p1,
                        email=setup_em,
                    )
                    st.success("Baptismal Master Admin account provisioned securely. Please sign in.")
                    st.rerun()
        return

    valid_sess, current_user = auth_svc.validate_session(st.session_state.get("bap_auth_user"))
    if not valid_sess or not current_user:
        st.session_state.pop("bap_auth_user", None)
        st.markdown(
            "<div class='ctg-bap-banner'>"
            "<h2>⛪ Church of the Temple of God — Baptismal Records & Certification System</h2>"
            "<p>1 Central Master Admin Archive • 9 Official Parish Leader Admin Dashboards (Hong Kong Excluded)</p>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.info("🔒 **Separate Baptismal Authorization Required:** Mag-sign in gamit ang iyong **Master Admin** o **Parish Leader Admin (1 of 9 Parishes)** account.")
        with st.form("baptismal_login_form"):
            in_u = st.text_input("Baptismal Username:")
            in_p = st.text_input("Baptismal Password:", type="password")
            if st.form_submit_button("Sign In to Baptismal System", type="primary"):
                ok, user_obj, msg = auth_svc.authenticate(in_u, in_p)
                if ok and user_obj:
                    st.session_state["bap_auth_user"] = user_obj
                    st.rerun()
                else:
                    st.error(msg)
        return

    account_id = str(current_user["account_id"])
    is_master_admin = authz_svc.is_master_admin(account_id)
    is_parish_admin = authz_svc.is_parish_leader_admin(account_id)

    if not is_master_admin:
        top_c1, top_c2 = st.columns([3.6, 1.0])
        with top_c1:
            st.caption(
                f"👤 Signed in as **{escape_html_text(current_user['full_name'])}** "
                f"(`{escape_html_text(current_user['role'])}` • `{escape_html_text(current_user.get('assigned_parish', ''))}`)"
            )
        with top_c2:
            if st.button("🚪 Sign Out (Baptismal)", use_container_width=True, key="bap_parish_sign_out"):
                auth_svc.logout(current_user)
                st.session_state.pop("bap_auth_user", None)
                st.rerun()

        if is_parish_admin:
            _render_parish_leader_admin_dashboard(account_id, current_user, svc)
            return

        st.info("🔒 **Temporary Requester Portal:** You only have access to view your specifically approved record.")
        my_grants = temp_svc.list_active_grants(account_id=account_id)
        for g in my_grants:
            rec_id = g["record_id"]
            if authz_svc.has_temporary_record_permission(account_id, rec_id, "view_record"):
                rec_data = hist_svc.get_record_by_id(rec_id, actor_account_id=account_id)
                if rec_data:
                    st.success(f"Approved Record: `{rec_id}` (Expires: {g['expires_at']})")
                    st.json(rec_data)
        return

    # =================================================================
    # MASTER ADMIN CENTRAL DASHBOARD (HORIZONTAL TABS LIKE WEEKLY REPORT)
    # =================================================================
    st.markdown(
        "<div class='ctg-bap-banner'>"
        "<h2>⛪ Church of the Temple of God — Master Admin Baptismal Archive & Certification Hub</h2>"
        "<p>Central Control Over All 9 Official Parishes • BOOK 1–9 Visual Archive • Historical Verification & Official Certificate Issuance (Hong Kong Excluded)</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    hdr_c1, hdr_c2, hdr_c3 = st.columns([2.4, 1.4, 0.9])
    with hdr_c1:
        st.caption(
            f"👑 Signed in as **{escape_html_text(current_user['full_name'])}** (`Master Admin`) • "
            f"📂 [CTG BAPTISMAL RECORDS Drive Folder](https://drive.google.com/drive/folders/{DEFAULT_BAPTISMAL_ROOT_FOLDER_ID}) • "
            f"📗 [Master Baptismal Excel](https://drive.google.com/file/d/{DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID}/view)"
        )
    with hdr_c2:
        if st.button("🔄 Sync 4,093 Records Mula Google Drive", use_container_width=True, type="primary", key="bap_top_sync_drive_btn"):
            with st.spinner("Binabasa ang OVERALL BAPTISMAL CERTIFICATE LIST.xlsx mula sa Google Drive..."):
                ok_s, cnt_s, msg_s = hist_svc.auto_sync_from_drive_or_disk(drive_mgr, force=True)
                if ok_s:
                    audit_svc.log_event(account_id, "HISTORICAL_DRIVE_SYNC", "SUCCESS", metadata={"records": cnt_s})
                    st.success(f"✅ {msg_s}")
                    st.rerun()
                else:
                    st.warning(msg_s)
    with hdr_c3:
        if st.button("🚪 Sign Out (Baptismal)", use_container_width=True, key="bap_sign_out_btn"):
            auth_svc.logout(current_user)
            st.session_state.pop("bap_auth_user", None)
            st.rerun()

    if not st.session_state.get("_bap_auto_sync_attempted"):
        st.session_state["_bap_auto_sync_attempted"] = True
        cur_stats_check = hist_svc.get_dashboard_statistics(actor_account_id=account_id)
        if cur_stats_check["total_historical"] == 0:
            with st.spinner("Awtomatikong kino-connect at binabasa ang historical records mula sa Google Drive (`OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`)..."):
                ok_auto, cnt_auto, _ = hist_svc.auto_sync_from_drive_or_disk(drive_mgr, force=False)
                if ok_auto and cnt_auto > 0:
                    audit_svc.log_event(account_id, "HISTORICAL_AUTO_SYNC", "SUCCESS", metadata={"records": cnt_auto})

    # Horizontal st.tabs navigation (just like Weekly Report & Parish Monitoring)
    (
        tab_dash,
        tab_search_verify,
        tab_books,
        tab_confirm_queue,
        tab_reqs_share,
        tab_new_reg,
        tab_cert_gen,
        tab_accounts_drive,
    ) = st.tabs([
        "📊 1. Dashboard",
        "🔍 2. Search ID, Book Image & Print/Share",
        "🖼️ 3. BOOK 1–9 Archive",
        "📥 4. Confirm Leaders' Uploads (Encode to Drive)",
        "📨 5. Requests & Release",
        "✍️ 6. New Baptism (Master)",
        "📜 7. Certificates Archive",
        "👥 8. 9 Parish Leaders & Drive/Audit",
    ])

    with tab_dash:
        stats = hist_svc.get_dashboard_statistics(actor_account_id=account_id)
        if stats["total_historical"] == 0:
            st.warning(
                "📥 **Wala pang naka-load na Historical Records sa local database.** Pindutin ang **'🔄 Sync 4,093 Records Mula Google Drive'** sa itaas "
                "o i-upload ang `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` sa ibaba!"
            )
            dash_up = st.file_uploader("O i-upload nang direkta ang OVERALL BAPTISMAL CERTIFICATE LIST.xlsx:", type=["xlsx"], key="dash_quick_xlsx_up")
            if dash_up and st.button("📥 I-load ang Historical Records Ngayon", type="primary"):
                cnt = hist_svc.import_sheet1_from_excel_bytes(dash_up.read())
                st.success(f"✅ Na-load ang {cnt} historical records!")
                st.rerun()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Historical Records (9 Parishes Scope)", stats["total_historical"])
        c2.metric("New Submissions (9 Parishes)", stats["new_registrations"])
        c3.metric("Generated Certificates", stats["generated_certificates"])
        c4.metric("Pending Requests", stats["pending_requests"])

        c5, c6 = st.columns(2)
        c5.metric("Matched Images", stats["matched_images"])
        c6.metric("Unmatched / Pending Images", stats["unmatched_images"])

        col_b, col_p, col_y = st.columns(3)
        with col_b:
            st.subheader("Records by Book")
            st.dataframe(stats["by_book"], use_container_width=True, hide_index=True)
        with col_p:
            st.subheader("Records by Parish (Hong Kong Excluded)")
            st.dataframe(stats["by_parish"], use_container_width=True, hide_index=True)
        with col_y:
            st.subheader("Records by Baptism Year")
            st.dataframe(stats["by_year"], use_container_width=True, hide_index=True)

    with tab_search_verify:
        st.subheader("🔍 Search Record ID (Palitan Lang ang Huling Number), Verify Book Image, Print & Ibigay sa Ibang Username")

        # Quick Record ID Jump where HIST- is automatic and user only changes the last number!
        if "bap_quick_num" not in st.session_state:
            st.session_state["bap_quick_num"] = 1

        q_c1, q_c2, q_c3, q_c4, q_c5 = st.columns([1.3, 0.55, 0.55, 1.4, 1.2])
        with q_c1:
            quick_num_val = st.number_input(
                "🔢 Quick ID Number (Huling Number Lang — 1 to 4093 → HIST-00XXX):",
                min_value=1,
                max_value=5000,
                value=int(st.session_state["bap_quick_num"]),
                step=1,
                key="bap_quick_num_input",
            )
            st.session_state["bap_quick_num"] = int(quick_num_val)
        with q_c2:
            st.write("")
            if st.button("⬅️ Prev ID", use_container_width=True):
                st.session_state["bap_quick_num"] = max(1, int(st.session_state["bap_quick_num"]) - 1)
                st.session_state["hist_rid_filter"] = str(st.session_state["bap_quick_num"])
                st.rerun()
        with q_c3:
            st.write("")
            if st.button("➡️ Next ID", use_container_width=True):
                st.session_state["bap_quick_num"] = min(5000, int(st.session_state["bap_quick_num"]) + 1)
                st.session_state["hist_rid_filter"] = str(st.session_state["bap_quick_num"])
                st.rerun()
        with q_c4:
            rid_filter = st.text_input(
                "🆔 Search by Record ID / Last Digits (hal. 743 o HIST-00743):",
                key="hist_rid_filter",
                placeholder="Type last digits e.g. 743",
            )
        with q_c5:
            st.write("")
            if st.button(f"🎯 Buksan Agad ang HIST-{int(quick_num_val):05d}", type="primary", use_container_width=True):
                st.session_state["hist_rid_filter"] = str(int(quick_num_val))
                st.rerun()

        fc1, fc2, fc3, fc4, fc5 = st.columns([1.6, 1.0, 0.8, 0.8, 0.7])
        kw = fc1.text_input("🔎 Keyword Search (Pangalan, Tatay, Nanay, Ninong/Ninang, Pari, o ID):", key="hist_kw")
        f_parish = fc2.text_input("Parish:", key="hist_f_parish")
        f_book = fc3.selectbox("Book Number (All Books):", ["ALL"] + [str(i) for i in range(1, 10)], key="hist_f_book")
        f_year = fc4.text_input("Year of Baptism:", key="hist_f_year")
        with fc5:
            st.write("")
            if st.button("🧹 Clear Filters", use_container_width=True):
                for k in ("hist_kw", "hist_f_parish", "hist_f_year", "hist_rid_filter"):
                    st.session_state[k] = ""
                st.session_state["hist_f_book"] = "ALL"
                st.rerun()

        page_num = int(st.number_input("Page (50 records per page):", min_value=1, value=1, step=1))
        book_flt_val = "" if f_book == "ALL" else f_book
        _search_params = inspect.signature(hist_svc.search_records).parameters
        if "record_id_query" in _search_params:
            rows, total_rows = hist_svc.search_records(
                keyword=kw,
                filters={"Parish": f_parish, "Book Number": book_flt_val, "Year of Baptism": f_year},
                page=page_num,
                page_size=50,
                actor_account_id=account_id,
                record_id_query=rid_filter,
            )
        else:
            _fallback_kw = resolve_quick_record_id(rid_filter) if rid_filter.strip() else kw
            rows, total_rows = hist_svc.search_records(
                keyword=_fallback_kw,
                filters={"Parish": f_parish, "Book Number": book_flt_val, "Year of Baptism": f_year},
                page=page_num,
                page_size=50,
                actor_account_id=account_id,
            )
        st.caption(f"Showing {len(rows)} of {total_rows} matching records (Record ID `HIST-00001`..`HIST-04026`).")

        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)
            record_options = [
                f"{r['Record ID']} — {r.get('Name of the Child', '')} (Book {r.get('Book Number', '?')}, Page {r.get('Page Number', '?')} • {r.get('Date of Baptism', '')})"
                for r in rows
            ]
            sel_rec_lbl = st.selectbox("📌 Pumili ng Record para Makita ang Buong 27 Details + Scanned Book Image + Print / Share:", record_options)
            sel_id = sel_rec_lbl.split(" — ")[0].strip()
            detail = hist_svc.get_record_by_id(sel_id, actor_account_id=account_id)

            if detail:
                audit_svc.log_event(account_id, "HISTORICAL_RECORD_VIEWED", "SUCCESS", related_record_id=sel_id)
                st.divider()
                left_rec_col, right_img_col = st.columns([1.15, 1.15])

                with left_rec_col:
                    st.markdown(
                        f"#### 📋 Record Details: `{sel_id}` — **{escape_html_text(detail.get('Name of the Child', ''))}** "
                        f"(Book `{escape_html_text(detail.get('Book Number', ''))}`, Page `{escape_html_text(detail.get('Page Number', ''))}`)"
                    )
                    st.dataframe(
                        [{"Field": k, "Value": str(detail.get(k, ""))} for k in (["Record ID"] + HISTORICAL_27_COLUMNS)],
                        use_container_width=True,
                        hide_index=True,
                        height=440,
                    )

                    # 1-Click Print & Download Certificate right on the selected record
                    pr_c1, pr_c2 = st.columns(2)
                    with pr_c1:
                        pr_orient = st.radio("Certificate Layout:", ["portrait", "landscape"], horizontal=True, key=f"pr_or_{sel_id}")
                    with pr_c2:
                        if st.button(f"🖨️ Generate & Print Certificate ({sel_id})", type="primary", use_container_width=True, key=f"btn_gen_pr_{sel_id}"):
                            cert_now = cert_svc.generate_certificate(
                                account_id, sel_id, orientation=pr_orient, upload_to_drive=False
                            )
                            st.session_state["_active_printable_cert"] = cert_now

                    if st.session_state.get("_active_printable_cert") and st.session_state["_active_printable_cert"].get("record_id") == sel_id:
                        ac_cert = st.session_state["_active_printable_cert"]
                        st.download_button(
                            f"⬇️ Download PDF Certificate ({ac_cert['certificate_id']})",
                            data=ac_cert["pdf_bytes"],
                            file_name=f"{ac_cert['certificate_id']}_{sel_id}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )
                        _render_printable_certificate_or_image_modal(pdf_bytes=ac_cert["pdf_bytes"], title=f"{sel_id} - {detail.get('Name of the Child', '')}")

                with right_img_col:
                    rec_bk_num = str(detail.get("Book Number") or "1").strip()
                    rec_pg_num = str(detail.get("Page Number") or "").strip()
                    st.markdown(
                        f"#### 🖼️ Scanned Book Image Verification (Book `{escape_html_text(rec_bk_num)}`, Page `{escape_html_text(rec_pg_num)}`)"
                    )
                    st.caption("Suriin kung tugma ang scanned image mula sa Google Drive BOOK 1–9 sa nakasulat sa record sa kaliwa.")

                    # Check both linked image associations and Google Drive BOOK folders across ALL books
                    linked_assocs = img_arc_svc.list_associations_for_record(sel_id, actor_account_id=account_id)
                    cand_drive_imgs = _find_candidate_drive_images_for_record(detail, drive_mgr)

                    book_folders_all = drive_mgr.discover_book_folders_in_drive()
                    default_bk_name = f"BOOK {rec_bk_num}" if f"BOOK {rec_bk_num}" in book_folders_all else "BOOK 1"
                    sel_bk_for_rec = st.selectbox(
                        "📂 Book Folder sa Google Drive (Pwede sa Lahat ng BOOK 1–9):",
                        HISTORICAL_BOOKS,
                        index=HISTORICAL_BOOKS.index(default_bk_name) if default_bk_name in HISTORICAL_BOOKS else 0,
                        key=f"rec_bk_sel_{sel_id}",
                    )
                    bk_drive_id = book_folders_all.get(sel_bk_for_rec, "")

                    if cand_drive_imgs:
                        st.success(f"✨ Nakahanap ng {len(cand_drive_imgs)} tugmang scanned image(s) sa Google Drive para sa `{detail.get('Name of the Child', '')}` / Page `{rec_pg_num}`:")
                    elif bk_drive_id:
                        if st.button(f"🔍 I-scan ang {sel_bk_for_rec} sa Google Drive para sa Record na Ito", key=f"btn_scan_bk_{sel_id}"):
                            try:
                                st.session_state[f"_drv_book_files_{sel_bk_for_rec}"] = drive_mgr.list_files_in_folder(bk_drive_id, page_size=500)
                                st.rerun()
                            except Exception as ex:
                                st.error(str(ex))

                    all_bk_files = cand_drive_imgs + [
                        {**f, "book_folder": sel_bk_for_rec, "fn_page_hint": extract_page_hint_from_filename(f.get("name", "")), "page_mismatch": False}
                        for f in st.session_state.get(f"_drv_book_files_{sel_bk_for_rec}", [])
                        if f.get("id") not in {c.get("id") for c in cand_drive_imgs}
                    ]

                    if all_bk_files:
                        img_choices = [
                            f"[{it.get('book_folder', sel_bk_for_rec)}] {it['name']} (Page Hint: {it.get('fn_page_hint') or '?'}) :: {it['id']}"
                            for it in all_bk_files
                        ]
                        chosen_img_lbl = st.selectbox("Piliin ang Scanned Book Page para Makita at Ikumpara sa Record:", img_choices, key=f"sel_img_cmp_{sel_id}")
                        chosen_drv_id = chosen_img_lbl.split(" :: ")[-1].strip()
                        chosen_obj = next((x for x in all_bk_files if x["id"] == chosen_drv_id), all_bk_files[0])

                        if chosen_obj.get("page_mismatch"):
                            st.warning(
                                f"⚠️ **Page Number Notice:** Ang Record ay nasa **Page `{rec_pg_num}`**, habang ang filename ng image (`{chosen_obj['name']}`) ay **`{chosen_obj.get('fn_page_hint')}`**. "
                                f"Tingnan ang image sa ibaba at pindutin ang **✅ Confirm Tama ang Image sa Record** kung ito ang tamang pahina."
                            )

                        v_c1, v_c2 = st.columns(2)
                        with v_c1:
                            if st.button("👁️ I-load at Tingnan ang Scanned Image", type="primary", use_container_width=True, key=f"ld_im_{sel_id}"):
                                try:
                                    raw_ib = drive_mgr.download_file_bytes(chosen_drv_id)
                                    st.session_state[f"_loaded_rec_img_{sel_id}"] = (chosen_obj["name"], chosen_drv_id, chosen_obj.get("book_folder", sel_bk_for_rec), raw_ib)
                                except Exception as ex:
                                    st.error(f"Hindi ma-download ang image: {ex}")
                        with v_c2:
                            if st.button("✅ Confirm Tama ang Image sa Record", use_container_width=True, key=f"cnf_im_{sel_id}"):
                                sug = img_match_svc.suggest_possible_match(
                                    sel_id,
                                    chosen_obj.get("book_folder", sel_bk_for_rec),
                                    chosen_obj["name"],
                                    drive_file_id=chosen_drv_id,
                                )
                                img_match_svc.confirm_image_association(account_id, sug["association_id"], approve=True, admin_notes="Verified side-by-side by Master Admin")
                                st.success(f"✅ Na-confirm na tugma ang `{chosen_obj['name']}` sa `{sel_id}`!")

                    if st.session_state.get(f"_loaded_rec_img_{sel_id}"):
                        im_nm, im_id, im_bk, im_bytes = st.session_state[f"_loaded_rec_img_{sel_id}"]
                        rot_v = st.selectbox("Rotate Image:", [0, 90, 180, 270], key=f"rot_s_{sel_id}")
                        zoom_v = st.slider("Zoom Width (px):", min_value=400, max_value=2200, value=1000, step=100, key=f"zm_s_{sel_id}")
                        disp_b = render_image_preview_bytes(im_bytes, rotation_degrees=rot_v, max_width=zoom_v)
                        st.image(disp_b, caption=f"{im_bk} / {im_nm} (Record {sel_id})")
                        st.download_button("⬇️ Download Scanned Book Page", data=disp_b, file_name=im_nm, mime="image/jpeg", use_container_width=True)
                        _render_printable_certificate_or_image_modal(image_bytes=disp_b, title=f"{sel_id} - {im_nm}")

                    if linked_assocs:
                        st.caption("🔗 Naka-link na Scanned Images sa Record na Ito:")
                        st.dataframe(linked_assocs, use_container_width=True, hide_index=True)

                # =========================================================
                # 📤 IBIGAY SA IBANG USERNAME / PARISH LEADER ADMIN PARA MA-PRINT DIN NILA
                # =========================================================
                st.divider()
                st.markdown(f"#### 📤 Ibigay ang Record at Printable Certificate (`{sel_id}`) sa Ibang Username / Parish Leader Admin")
                st.caption("Pumili ng Parish Leader Admin username (o i-type ang username) upang agad nilang makita, ma-download, at **ma-print** ang Official Certificate at Record na ito sa kanilang dashboard.")
                p_admins_all = auth_svc.list_parish_admin_accounts()
                uname_choices = [f"{a['username']} ({a['assigned_parish']})" for a in p_admins_all]
                sh_c1, sh_c2, sh_c3, sh_c4 = st.columns([1.5, 1.1, 0.9, 1.3])
                with sh_c1:
                    picked_u_lbl = st.selectbox(
                        "Piliin ang Parish Leader Admin Username:",
                        ["(Mag-type ng Username sa Kanan)"] + uname_choices,
                        key=f"sh_sel_u_{sel_id}",
                    )
                with sh_c2:
                    default_u_txt = picked_u_lbl.split(" (")[0].strip() if picked_u_lbl != "(Mag-type ng Username sa Kanan)" else ""
                    target_uname_in = st.text_input("Target Username:", value=default_u_txt, key=f"sh_txt_u_{sel_id}")
                with sh_c3:
                    share_hrs = st.number_input("Valid Hours:", min_value=1.0, max_value=720.0, value=72.0, key=f"sh_hrs_{sel_id}")
                with sh_c4:
                    st.write("")
                    if st.button("📤 Ibigay & I-release para Ma-Print Nila", type="primary", use_container_width=True, key=f"btn_sh_u_{sel_id}"):
                        if not target_uname_in.strip():
                            st.error("Pumili o mag-type muna ng target username.")
                        else:
                            try:
                                sh_res = cert_svc.share_record_and_certificate_with_username(
                                    admin_account_id=account_id,
                                    record_id=sel_id,
                                    target_username=target_uname_in.strip(),
                                    duration_hours=share_hrs,
                                )
                                st.success(
                                    f"✅ Naibigay na ang Record `{sel_id}` at Official Printable Certificate `{sh_res['certificate_id']}` "
                                    f"kay **`{sh_res['target_username']}`**! Maaari na niya itong i-print sa kanyang `🖨️ My Approved Certificates & Records` tab."
                                )
                            except Exception as ex:
                                st.error(str(ex))

    with tab_books:
        st.subheader("🖼️ BOOK 1–9 Visual Archive & Google Drive Scanned Pages (Kumpleto BOOK 1 Hanggang BOOK 9)")
        book_folders = drive_mgr.discover_book_folders_in_drive()
        bk_summary_rows = [
            {
                "Book Folder": bk,
                "Google Drive Folder ID": book_folders.get(bk, ""),
                "Google Drive Link": f"https://drive.google.com/drive/folders/{book_folders.get(bk, '')}" if book_folders.get(bk) else "",
                "Cached Scanned Files": len(st.session_state.get(f"_drv_book_files_{bk}", [])),
            }
            for bk in HISTORICAL_BOOKS
        ]
        st.dataframe(bk_summary_rows, use_container_width=True, hide_index=True)
        if st.button("🔄 I-sync at I-load ang Listahan ng Lahat ng Scanned Images sa BOOK 1 Hanggang BOOK 9 Mula sa Google Drive", type="primary"):
            loaded_counts = {}
            for bk in HISTORICAL_BOOKS:
                fid_b = book_folders.get(bk, "")
                if fid_b:
                    try:
                        imgs_b = drive_mgr.list_files_in_folder(fid_b, page_size=1000)
                        st.session_state[f"_drv_book_files_{bk}"] = imgs_b
                        loaded_counts[bk] = len(imgs_b)
                    except Exception as ex:
                        loaded_counts[bk] = f"Error: {ex}"
            st.success(f"✅ Na-load ang BOOK 1–9 mula sa Google Drive: {loaded_counts}")
            st.rerun()
        sel_book = st.selectbox("Select Book Folder:", HISTORICAL_BOOKS)
        b_fid = book_folders.get(sel_book, "")
        if not b_fid:
            st.info(f"Walang nakitang `{sel_book}` folder sa Google Drive (kasalukuyang naka-upload sa Drive: {', '.join(sorted(book_folders.keys()))}).")
        else:
            st.caption(f"📂 Google Drive Folder ID (`{sel_book}`): `{b_fid}` — [Buksan sa Google Drive](https://drive.google.com/drive/folders/{b_fid})")
            if st.button(f"🔍 I-load ang Listahan ng Scanned Images sa {sel_book} Mula sa Google Drive", key=f"btn_list_drv_{sel_book}"):
                try:
                    drv_imgs = drive_mgr.list_files_in_folder(b_fid, page_size=500)
                    st.session_state[f"_drv_book_files_{sel_book}"] = drv_imgs
                except Exception as ex:
                    st.error(f"Google Drive Folder List Error: {ex}")
            cached_drv_imgs = st.session_state.get(f"_drv_book_files_{sel_book}", [])
            if cached_drv_imgs:
                st.dataframe(cached_drv_imgs, use_container_width=True, hide_index=True)
                sel_drv_img = st.selectbox(
                    "Pumili ng Scanned Image mula sa Google Drive para I-preview / I-rotate / I-print:",
                    [f"{it['name']} ({it['id']})" for it in cached_drv_imgs],
                    key=f"sel_drv_img_{sel_book}",
                )
                if sel_drv_img and st.button("👁️ I-download at I-preview ang Scanned Image na Ito", key=f"btn_prev_drv_{sel_book}"):
                    chosen_id = sel_drv_img.rsplit("(", 1)[-1].rstrip(")")
                    chosen_nm = sel_drv_img.rsplit(" (", 1)[0]
                    try:
                        img_b = drive_mgr.download_file_bytes(chosen_id)
                        st.session_state["_active_drive_preview_img"] = (chosen_nm, chosen_id, img_b)
                    except Exception as ex:
                        st.error(f"Hindi ma-download ang image: {ex}")

            if st.session_state.get("_active_drive_preview_img"):
                p_nm, p_id, p_bytes = st.session_state["_active_drive_preview_img"]
                rot_d = st.selectbox("Rotate Drive Image (Degrees):", [0, 90, 180, 270], key="rot_drv_preview")
                zw_d = st.slider("Zoom Width (px):", min_value=400, max_value=2200, value=1100, step=100, key="zoom_drv_preview")
                rendered_b = render_image_preview_bytes(p_bytes, rotation_degrees=rot_d, max_width=zw_d)
                st.image(rendered_b, caption=f"{p_nm} (Drive ID: {p_id})")
                st.download_button("⬇️ Download Scanned Image", data=rendered_b, file_name=p_nm, mime="image/jpeg")
                _render_printable_certificate_or_image_modal(image_bytes=rendered_b, title=p_nm)

    with tab_confirm_queue:
        st.subheader("📥 Parish Leader Admins' New Baptism Uploads — Master Admin Confirmation & Google Drive Encoding Queue")
        st.info(
            "📌 Dito pumupunta ang lahat ng **New Baptized Members at Images/Files** na in-upload ng **9 Parish Leader Admins**. "
            "Kapag pinindot mo ang **`✅ Confirm & I-encode sa Google Drive Master Excel + I-save ang Image/File sa Google Drive`**, "
            "awtomatiko na itong ie-encode sa `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` (`NEW BAPTISMS - YYYY`) at ise-save ang uploaded images/files sa Google Drive!"
        )
        p_filter = st.selectbox("Filter by Parish:", ["ALL"] + OFFICIAL_NINE_PARISHES)
        all_subs = reg_svc.list_submissions_for_actor(account_id, parish_filter=p_filter)
        if not all_subs:
            st.info("Wala pang bagong baptism submissions mula sa 9 Parish Leader Admins.")
        else:
            st.dataframe(all_subs, use_container_width=True, hide_index=True)
            sel_sub_id = st.selectbox("Piliin ang Submission Record ID para I-review at I-confirm:", [s["record_id"] for s in all_subs])
            sel_sub_obj = next(s for s in all_subs if s["record_id"] == sel_sub_id)
            staged_files = reg_svc.get_staged_files_for_record(sel_sub_id)

            sc1, sc2 = st.columns([1.2, 1.0])
            with sc1:
                st.markdown(f"#### 📋 Submitted Details (`{sel_sub_id}` — `{sel_sub_obj['parish']}`)")
                st.json(json.loads(sel_sub_obj["payload_json"]))
            with sc2:
                st.markdown(f"#### 📸 Uploaded Baptismal Images/Files ({len(staged_files)} File(s))")
                for sf in staged_files:
                    st.caption(f"📎 `{sf['filename']}` ({sf['file_size_bytes']} bytes • Pushed to Drive: `{bool(sf['is_pushed_to_drive'])}`)")
                    if sf.get("bytes") and not sf["filename"].lower().endswith(".pdf"):
                        try:
                            st.image(sf["bytes"], caption=sf["filename"], use_container_width=True)
                        except Exception:
                            pass

            ver_notes = st.text_input("Master Admin Confirmation / Review Notes:", key="adm_conf_notes")
            bc1, bc2, bc3 = st.columns(3)
            with bc1:
                if st.button("✅ Confirm & I-encode sa Google Drive Master Excel + I-save ang Image/File sa Drive", type="primary", use_container_width=True):
                    try:
                        c_res = reg_svc.confirm_and_sync_submission_to_drive(account_id, sel_sub_id, notes=ver_notes, sync_to_drive=True)
                        st.success(
                            f"✅ Na-confirm ang `{sel_sub_id}`! Workbook Status: `{c_res['workbook_update_status']}` • "
                            f"Uploaded Files to Google Drive: `{len(c_res['uploaded_drive_file_ids'])}`"
                        )
                        st.rerun()
                    except Exception as ex:
                        st.error(str(ex))
            with bc2:
                if st.button("📝 Humingi ng Additional Information (Allow Leader to Edit)", use_container_width=True):
                    reg_svc.verify_registration(account_id, sel_sub_id, "Additional Information Required", ver_notes)
                    st.info(f"Marked `{sel_sub_id}` as Additional Information Required.")
                    st.rerun()
            with bc3:
                if st.button("❌ Reject Submission", use_container_width=True):
                    reg_svc.verify_registration(account_id, sel_sub_id, "Rejected", ver_notes)
                    st.warning(f"Rejected `{sel_sub_id}`.")
                    st.rerun()

    with tab_reqs_share:
        st.subheader("📨 Historical Record & Certified Certificate Requests from the 9 Parish Leader Admins")
        reqs = temp_svc.list_requests(actor_account_id=account_id)
        if not reqs:
            st.info("Wala pang requests mula sa 9 Parish Leader Admins.")
        else:
            st.dataframe(reqs, use_container_width=True, hide_index=True)
            sel_req_id = st.selectbox("Select Request ID:", [r["request_id"] for r in reqs])
            sel_req_obj = next(r for r in reqs if r["request_id"] == sel_req_id)
            st.caption(
                f"**Type:** `{sel_req_obj.get('request_type')}` • **Parish:** `{sel_req_obj.get('assigned_parish')}` • "
                f"**Requester Username:** `{sel_req_obj.get('requester_username')}` • **Person:** `{sel_req_obj.get('requested_person_name')}` • "
                f"**Current Status:** `{sel_req_obj.get('status')}`"
            )
            verified_rec_id = st.text_input(
                "Verified Record ID o Huling Number (hal. 743 o HIST-00743 o BAP-2026-...):",
                value=str(sel_req_obj.get("target_record_id") or ""),
            )
            new_req_status = st.selectbox("Update Request Status:", CANONICAL_REQUEST_STATUSES)
            rev_notes = st.text_input("Master Admin Review Notes:")
            orient_rel = st.radio("Certificate Layout (if issuing & releasing):", ["portrait", "landscape"], horizontal=True)

            b_st, b_rel, b_rj = st.columns(3)
            if b_st.button("🔄 Update Status Only", use_container_width=True):
                temp_svc.update_request_status(account_id, sel_req_id, new_req_status, rev_notes, resolve_quick_record_id(verified_rec_id))
                st.success(f"Request `{sel_req_id}` updated to `{new_req_status}`.")
                st.rerun()
            if b_rel.button("🏅 Approve, Generate & Release Official Printable Certificate sa Requester", type="primary", use_container_width=True):
                if not verified_rec_id.strip():
                    st.error("Ilagay muna ang verified Record ID (o huling number, hal. 743).")
                else:
                    try:
                        cert_out = cert_svc.approve_and_release_certificate_for_request(
                            admin_account_id=account_id,
                            request_id=sel_req_id,
                            record_id=resolve_quick_record_id(verified_rec_id.strip()),
                            orientation=orient_rel,
                            review_notes=rev_notes,
                        )
                        st.success(
                            f"✅ Official Certificate `{cert_out['certificate_id']}` generated, approved, and released to **`{sel_req_obj.get('requester_username')}`** para ma-print na nila!"
                        )
                        st.rerun()
                    except Exception as ex:
                        st.error(str(ex))
            if b_rj.button("❌ Reject Request", use_container_width=True):
                temp_svc.reject_request(account_id, sel_req_id, rev_notes)
                st.warning(f"Request `{sel_req_id}` rejected.")
                st.rerun()

    with tab_new_reg:
        st.subheader("✍️ Direct New Baptism Registration by Master Admin (9 Official Parishes — Hong Kong Excluded)")
        with st.form("new_baptism_reg_form"):
            payload: Dict[str, str] = {}
            cols = st.columns(3)
            for idx, col_name in enumerate(HISTORICAL_27_COLUMNS):
                with cols[idx % 3]:
                    if col_name == "Parish":
                        payload["Parish"] = st.selectbox("Parish (9 Official Parishes):", OFFICIAL_NINE_PARISHES)
                    elif col_name == "Year of Baptism":
                        payload[col_name] = st.selectbox(col_name, SUPPORTED_NEW_YEARS)
                    else:
                        payload[col_name] = st.text_input(col_name, key=f"reg_f_{col_name}")
            up_imgs = st.file_uploader("Upload Baptismal Images (Private Drive Year Folder):", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
            sync_drive = st.checkbox("Direktang I-encode sa Google Drive Workbook (`NEW BAPTISMS - YYYY`)", value=True)
            if st.form_submit_button("Register & Encode Baptism to Google Drive", type="primary"):
                img_list = [{"bytes": uf.read(), "filename": uf.name, "mime_type": uf.type or "image/jpeg"} for uf in (up_imgs or [])]
                try:
                    res = reg_svc.register_new_baptism(account_id, payload, uploaded_images=img_list, sync_to_drive_excel=sync_drive)
                    st.success(f"Registered! Record ID: `{res['record_id']}` • Parish: `{res['parish']}` • Worksheet: `{res['worksheet_name']}`")
                except Exception as ex:
                    st.error(str(ex))

    with tab_cert_gen:
        st.subheader("📜 Official CTG Certificate Generator, Print & Archive")
        rec_id_in = st.text_input("Enter Record ID o Huling Number (hal. 743 o HIST-00743 o BAP-2026-...):")
        orient = st.radio("Layout Orientation:", ["portrait", "landscape"], horizontal=True)
        heading_in = st.text_input("Certificate Heading:", value="CERTIFICATE OF BAPTISM")
        wording_in = st.text_area(
            "Certificate Wording:",
            value="This is to certify that the person named below received the Sacrament of Christian Baptism according to the rites of the Church of the Temple of God.",
        )
        is_reissue = st.checkbox("Mark as Reissued Certificate")
        reissue_why = st.text_input("Reissue Reason (Required if Reissued):") if is_reissue else ""
        up_cert_drive = st.checkbox("Upload Private PDF to Google Drive (`GENERATED CERTIFICATES/<year>/`)", value=False)

        if st.button("🖨️ Generate & Print Official Certificate PDF", type="primary") and rec_id_in.strip():
            try:
                cert_res = cert_svc.generate_certificate(
                    account_id,
                    resolve_quick_record_id(rec_id_in.strip()),
                    orientation=orient,
                    is_reissue=is_reissue,
                    reissue_reason=reissue_why,
                    custom_heading=heading_in,
                    custom_wording=wording_in,
                    upload_to_drive=up_cert_drive,
                )
                st.success(f"Generated Certificate `{cert_res['certificate_id']}` ({cert_res['issuance_status']})")
                st.download_button(
                    "⬇️ Download Certificate PDF",
                    data=cert_res["pdf_bytes"],
                    file_name=f"{cert_res['certificate_id']}.pdf",
                    mime="application/pdf",
                )
                _render_printable_certificate_or_image_modal(pdf_bytes=cert_res["pdf_bytes"], title=cert_res["certificate_id"])
            except Exception as ex:
                st.error(str(ex))

        st.divider()
        st.markdown("#### 🗂️ Generated Certificates Archive")
        certs_list = cert_svc.list_certificates(actor_account_id=account_id)
        st.dataframe(certs_list, use_container_width=True, hide_index=True)

    with tab_accounts_drive:
        st.subheader("👥 Nine Separate Parish Leader Admin Accounts (Hong Kong Excluded)")
        st.caption("Official 9 Parishes: " + " • ".join(OFFICIAL_NINE_PARISHES))
        with st.expander("⚡ 1-Click Provision All 9 Official Parish Leader Admin Accounts", expanded=True):
            init_pw_9 = st.text_input("Set Initial Password for Newly Provisioned Parish Admin Accounts (min 8 chars):", type="password")
            if st.button("Provision / Verify All 9 Parish Leader Admin Accounts", type="primary"):
                if len(init_pw_9) < 8:
                    st.error("Initial password must be at least 8 characters.")
                else:
                    res_9 = auth_svc.provision_nine_parish_leader_admins(account_id, init_pw_9)
                    st.success("All 9 Official Parish Leader Admin accounts are ready!")
                    st.dataframe(res_9, use_container_width=True, hide_index=True)

        existing_p_admins = auth_svc.list_parish_admin_accounts()
        if existing_p_admins:
            st.markdown("#### Active Parish Leader Admin Accounts (Magkakaibang Username Bawat Parokya)")
            st.dataframe(existing_p_admins, use_container_width=True, hide_index=True)
            with st.expander("🔑 Reset Password for a Specific Parish Leader Admin"):
                sel_pa = st.selectbox(
                    "Select Parish Admin Account:",
                    [f"{a['assigned_parish']} — {a['username']} ({a['account_id']})" for a in existing_p_admins],
                )
                new_pa_pw = st.text_input("New Password (min 8 chars):", type="password", key="rst_pa_pw")
                if st.button("Reset Selected Parish Admin Password"):
                    target_aid = sel_pa.rsplit("(", 1)[-1].rstrip(")")
                    try:
                        auth_svc.reset_password(target_aid, new_pa_pw, actor_account_id=account_id)
                        st.success("Password reset successfully.")
                    except Exception as ex:
                        st.error(str(ex))

        st.divider()
        st.subheader("☁️ Google Drive Private Storage & Audit Logs")
        wb_id = str(drive_mgr.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
        root_id = str(drive_mgr.config.get("root_folder_id") or DEFAULT_BAPTISMAL_ROOT_FOLDER_ID).strip()
        st.markdown(
            f"- **Root Folder (`CTG BAPTISMAL RECORDS`):** `{root_id}` ([Buksan sa Google Drive](https://drive.google.com/drive/folders/{root_id}))\n"
            f"- **Master Workbook (`OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`):** `{wb_id}` ([Buksan sa Google Drive](https://drive.google.com/file/d/{wb_id}/view))"
        )
        if st.button("Verify Workbook Editor Permission & Private Folder Tree"):
            try:
                meta = drive_mgr.verify_workbook_write_access(wb_id)
                tree = drive_mgr.ensure_baptismal_folder_tree()
                st.success(f"Verified Editor access on '{meta.get('name')}' (Revision: {meta.get('headRevisionId')})")
                st.json(tree)
            except GoogleDrivePermissionError as pex:
                st.error(f"Permission Verification Failed: {pex}")
            except Exception as ex:
                st.error(f"Google Drive Check Error: {ex}")

        ev_rows = audit_svc.list_events(limit=150)
        st.dataframe(ev_rows, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    st.set_page_config(
        page_title="CTG Baptismal Records & Certification System",
        page_icon="⛪",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_baptismal_portal(embedded_in_hub=False)

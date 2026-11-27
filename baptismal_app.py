"""
baptismal_app.py — CTG BAPTISMAL RECORDS & Certification System
===============================================================
Clean, Role-Based Workflow:
1. 9 Parish Leader Admins (Lipay Norte, Malanas, Poblacion, Duplas, Santiago, Caba, Manila,
   Andabuen, Apayao — Hong Kong Excluded):
   - No extra Baptismal login gate: opens directly to their assigned Parish Leader Admin
     Baptismal Dashboard from their main app session.
   - Strictly 3 tabs only:
     1) Direct New Baptism Registration (staged for Master Admin confirmation & Google Drive encoding)
     2) Historical Record & Certified Certificate Requests (request old copy from Master Admin)
     3) Approved Copies Mula kay Master Admin (Print / Download released certificates)
   - Zero access to Master Admin Archive, BOOK 1–9, or other parishes.
2. Master Admin (`CTG BAPTISMAL RECORDS`):
   - Dedicated `CTG BAPTISMAL RECORDS` Master Admin login gate (`ctg.Baptismal.admin`, verified
     strictly against PBKDF2-HMAC-SHA256 hash — plaintext password is never stored in code).
   - 7 clean horizontal tabs (no UI log dumps) covering Overview, Quick Record ID Search + Side-by-Side
     BOOK 1–9 Image Viewer + 1-Click Print/Share, BOOK 1–9 Visual Archive, Confirm Leaders' Uploads
     (Encode to Google Drive), Requests & Release, Direct New Baptism, and Certificates Archive.
"""

import base64
from datetime import datetime, timezone
import inspect
import json
import re
import time as _time_mod
from typing import Any, Dict, List
import streamlit as st
import streamlit.components.v1 as components

from baptismal.audit_logs import AuditLogger
from baptismal.authentication import AuthService, verify_password_pbkdf2
from baptismal.authorization import AuthorizationService
from baptismal.backup_manager import BackupManager
from baptismal.certificates import CertificateService
from baptismal.config import (
    BOOK9_EPSON_SCAN_MAP,
    CANONICAL_REQUEST_STATUSES,
    DEFAULT_BAPTISMAL_ROOT_FOLDER_ID,
    DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID,
    DEFAULT_NINE_PARISH_ADMIN_USERNAMES,
    DEFAULT_PARISH_LEADER_INTERNAL_PW_HASH,
    EDITABLE_REGISTRATION_STATUSES,
    HISTORICAL_27_COLUMNS,
    HISTORICAL_BOOKS,
    MASTER_ADMIN_BAPTISMAL_PW_HASH,
    MASTER_ADMIN_BAPTISMAL_USERNAME,
    OFFICIAL_NINE_PARISHES,
    SUPPORTED_NEW_YEARS,
    THEME_GOLD,
    THEME_NAVY,
    THEME_NAVY_DARK,
    THEME_WHITE,
    VERIFIED_RECORD_DRIVE_IMAGE_MAP,
    is_excluded_parish,
    normalize_official_nine_parish,
)
from baptismal.database import DatabaseManager
from baptismal.google_drive import GoogleDriveManager
from baptismal.historical_records import HistoricalRecordsService, resolve_quick_record_id
from baptismal.image_archive import ImageArchiveService, render_image_preview_bytes
from baptismal.image_matching import ImageMatchingService, extract_page_hint_from_filename
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService
from baptismal.security_crypto import (
    enforce_data_leak_protection_on_workspace,
    redact_sensitive_text,
)
from baptismal.temporary_access import TemporaryAccessService
from baptismal.validation import escape_html_text, validate_uploaded_file

_BAP_SERVICES_VERSION = "2026-11-27-v6-clean-workflow"


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
    enforce_data_leak_protection_on_workspace(".")
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
        "_provisioned_cache": {},
    }


_PROVISIONED_USERS_MEMO: Dict[str, Dict[str, Any]] = {}


def _resolve_or_provision_hub_synced_user(auth_svc: AuthService, is_master: bool, parish_name: str = "") -> Dict[str, Any]:
    """Ensures the 1 Master Admin (`ctg.Baptismal.admin` with hidden PBKDF2-SHA256 hash) and 9 Official
    Parish Leader Admin accounts exist in the database, memoized in memory for 0-overhead reruns."""
    now_iso = datetime.now(timezone.utc).isoformat()
    cache_key = "MASTER" if is_master else f"PARISH::{parish_name.strip().upper()}"
    if cache_key in _PROVISIONED_USERS_MEMO:
        cached = dict(_PROVISIONED_USERS_MEMO[cache_key])
        cached["authenticated_at"] = now_iso
        return cached
    master_row = auth_svc.db.fetch_one(
        "SELECT * FROM baptismal_accounts WHERE role = 'Master Admin' AND is_active = 1 LIMIT 1"
    )
    if not master_row:
        m_id = auth_svc.create_account(
            username=MASTER_ADMIN_BAPTISMAL_USERNAME,
            full_name="CTG Baptismal Master Admin",
            role="Master Admin",
            assigned_parish="ALL_NINE_PARISHES",
            password_hash=MASTER_ADMIN_BAPTISMAL_PW_HASH,
        )
        master_row = auth_svc.db.fetch_one("SELECT * FROM baptismal_accounts WHERE account_id = %s", (m_id,))
    else:
        if master_row.get("username") != MASTER_ADMIN_BAPTISMAL_USERNAME or master_row.get("password_hash") != MASTER_ADMIN_BAPTISMAL_PW_HASH:
            auth_svc.db.execute(
                "UPDATE baptismal_accounts SET username = %s, password_hash = %s, failed_attempts = 0, locked_until = NULL, updated_at = %s WHERE account_id = %s",
                (MASTER_ADMIN_BAPTISMAL_USERNAME, MASTER_ADMIN_BAPTISMAL_PW_HASH, now_iso, master_row["account_id"]),
            )
            master_row = auth_svc.db.fetch_one("SELECT * FROM baptismal_accounts WHERE account_id = %s", (master_row["account_id"],))

    for parish in OFFICIAL_NINE_PARISHES:
        uname = str(DEFAULT_NINE_PARISH_ADMIN_USERNAMES[parish]).strip().lower()
        existing = auth_svc.db.fetch_one(
            "SELECT account_id FROM baptismal_accounts WHERE role = 'Parish Leader Admin' AND assigned_parish = %s",
            (parish,),
        )
        if not existing:
            auth_svc.create_account(
                username=uname,
                full_name=f"{parish} Baptismal Leader Admin",
                role="Parish Leader Admin",
                assigned_parish=parish,
                password_hash=DEFAULT_PARISH_LEADER_INTERNAL_PW_HASH,
            )

    _PROVISIONED_USERS_MEMO["MASTER"] = {
        "account_id": master_row["account_id"],
        "username": master_row["username"],
        "full_name": master_row["full_name"],
        "email": master_row.get("email", ""),
        "role": "Master Admin",
        "assigned_parish": "ALL_NINE_PARISHES",
        "authenticated_at": now_iso,
        "session_token": "hub_synced_master_admin",
    }
    if is_master:
        return dict(_PROVISIONED_USERS_MEMO["MASTER"])

    norm_p = normalize_official_nine_parish(parish_name)
    p_row = auth_svc.db.fetch_one(
        "SELECT * FROM baptismal_accounts WHERE role = 'Parish Leader Admin' AND assigned_parish = %s AND is_active = 1 LIMIT 1",
        (norm_p,),
    )
    res_p = {
        "account_id": p_row["account_id"],
        "username": p_row["username"],
        "full_name": p_row["full_name"],
        "email": p_row.get("email", ""),
        "role": "Parish Leader Admin",
        "assigned_parish": norm_p,
        "authenticated_at": now_iso,
        "session_token": f"hub_synced_{p_row['username']}",
    }
    _PROVISIONED_USERS_MEMO[cache_key] = res_p
    _PROVISIONED_USERS_MEMO[f"PARISH::{norm_p.upper()}"] = res_p
    return dict(res_p)


def _render_printable_certificate_or_image_modal(pdf_bytes: bytes = b"", image_bytes: bytes = b"", title: str = "CTG Baptismal Print") -> None:
    """Renders an inline PDF/Image viewer with a 1-click browser Print dialog button."""
    if pdf_bytes:
        b64_pdf = base64.b64encode(pdf_bytes).decode("ascii")
        components.html(
            f"""
            <div style="margin-bottom:8px;">
                <button onclick="const w=window.open('');w.document.write('<iframe src=\'data:application/pdf;base64,{b64_pdf}\' style=\'width:100%;height:100vh;border:none;\'></iframe>');setTimeout(()=>w.print(),700);"
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
                <button onclick="const w=window.open('');w.document.write('<img src=\'data:image/jpeg;base64,{b64_img}\' style=\'max-width:100%;\'/>');setTimeout(()=>w.print(),400);"
                        style="background:#0F2C59;color:#FFFFFF;border:2px solid #D4AF37;padding:8px 18px;border-radius:8px;font-weight:700;cursor:pointer;">
                    🖨️ Print Scanned Book Page Now ({escape_html_text(title)})
                </button>
            </div>
            """,
            height=55,
        )


def _render_drive_file_embed_fallback(drive_file_id: str, filename: str, book_folder: str) -> None:
    """Renders a live Google Drive image preview iframe + direct open link whenever raw byte download is unavailable."""
    if not drive_file_id:
        return
    preview_url = f"https://drive.google.com/file/d/{drive_file_id}/preview"
    view_url = f"https://drive.google.com/file/d/{drive_file_id}/view"
    st.markdown(
        f"🖼️ **{escape_html_text(book_folder)} / `{escape_html_text(filename)}`** — "
        f"[🔗 Buksan at I-print sa Google Drive]({view_url})"
    )
    components.html(
        f'<iframe src="{preview_url}" width="100%" height="520" style="border:2px solid #D4AF37;border-radius:10px;background:#0F2C59;" allow="autoplay"></iframe>',
        height=535,
    )


def _find_candidate_drive_images_for_record(record: Dict[str, Any], drive_mgr: GoogleDriveManager) -> List[Dict[str, Any]]:
    """Searches pre-verified mappings, BOOK 9 Epson scans, and the record's Book folder in Google Drive
    for scanned images matching the child's name, record ID, or page number."""
    rec_id = str(record.get("Record ID") or "").strip().upper()
    book_num = str(record.get("Book Number") or "1").strip()
    page_num = str(record.get("Page Number") or "").strip()
    child_name = str(record.get("Name of the Child") or "").strip().lower()
    name_tokens = [t for t in re.split(r"[^a-z0-9]+", child_name) if len(t) >= 3]

    candidates: List[Dict[str, Any]] = []
    seen_ids = set()

    if rec_id in VERIFIED_RECORD_DRIVE_IMAGE_MAP:
        vm = VERIFIED_RECORD_DRIVE_IMAGE_MAP[rec_id]
        drv_id = vm["drive_id"]
        seen_ids.add(drv_id)
        candidates.append({
            "id": drv_id,
            "name": vm["filename"],
            "raw_name": vm["filename"],
            "book_folder": vm["book_folder"],
            "fn_page_hint": vm["page_number"],
            "match_score": 25,
            "page_mismatch": bool(page_num and vm["page_number"] and page_num != vm["page_number"]),
        })

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
            if drv_id not in seen_ids:
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
    primary_bk = f"BOOK {book_num}" if f"BOOK {book_num}" in book_folders else "BOOK 1"
    fid = book_folders.get(primary_bk, "")

    files_to_check: List[Any] = []
    cache_k = f"_drv_book_files_{primary_bk}"
    if cache_k in st.session_state and st.session_state[cache_k]:
        files_to_check.extend([(primary_bk, f) for f in st.session_state[cache_k]])
    elif fid:
        rec_search_cache_k = f"_drv_rec_search_{rec_id}"
        if rec_search_cache_k in st.session_state:
            files_to_check.extend([(primary_bk, f) for f in st.session_state.get(rec_search_cache_k, [])])

    for bk_other in sorted(book_folders.keys()):
        if bk_other != primary_bk and st.session_state.get(f"_drv_book_files_{bk_other}"):
            files_to_check.extend([(bk_other, f) for f in st.session_state[f"_drv_book_files_{bk_other}"]])

    for bk_name, f_it in files_to_check:
        f_id = str(f_it.get("id") or "")
        if not f_id or f_id in seen_ids:
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
        if score >= 4:
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
    """Renders the strictly isolated Parish Leader Admin Dashboard for 1 of the 9 Official Parishes (Hong Kong Excluded).
    No extra login required for Parish Leaders — opens directly to their 3 allowed tabs:
    1. Direct New Baptism Registration (ipasa sa Master Admin para i-confirm at i-encode sa Google Drive).
    2. Historical Record & Certified Certificate Requests (humingi ng dating copy sa Master Admin).
    3. Approved Copies Mula kay Master Admin (Print / Download released certificates).
    """
    authz_svc: AuthorizationService = svc["authz"]
    hist_svc: HistoricalRecordsService = svc["hist"]
    reg_svc: RegistrationService = svc["reg"]
    cert_svc: CertificateService = svc["certs"]
    temp_svc: TemporaryAccessService = svc["temp_acc"]

    assigned_parish = authz_svc.get_authenticated_parish(account_id)

    st.markdown(
        f"<div class='ctg-bap-banner'>"
        f"<h2>⛪ {escape_html_text(assigned_parish)} — Parish Leader Admin Baptismal Dashboard</h2>"
        f"<p>1. Magpasa ng Bagong Na-Baptize sa Master Admin • 2. Mag-Request ng Dating Copy / Certified Certificate sa Master Admin • 3. I-print ang Approved Copy</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    p_stats = reg_svc.get_parish_dashboard_statistics(account_id)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("My Submitted New Baptisms", p_stats["my_new_registrations"])
    k2.metric("Pending Master Admin Confirm", p_stats["pending_verification"])
    k3.metric("Confirmed & Encoded sa Drive", p_stats["approved_registrations"])
    k4.metric("Pending Copy Requests", p_stats["pending_certificate_requests"])
    k5.metric("Approved Copies (Ready to Print)", p_stats["approved_certificates"])

    ptab_new_bap, ptab_req_copy, ptab_approved_print = st.tabs([
        "✍️ 1. Direct New Baptism Registration (Ipasa sa Master Admin)",
        "📨 2. Historical Record & Certified Certificate Requests (Humingi ng Dating Copy sa Master Admin)",
        "🖨️ 3. Approved Copies Mula kay Master Admin (Print / Download)",
    ])

    with ptab_new_bap:
        st.subheader(f"✍️ Direct New Baptism Registration — {assigned_parish}")
        st.info(
            f"📌 **Para sa {assigned_parish} Leader Admin:** I-fill up ang bagong na-baptize at i-upload ang Baptismal Image/File. "
            "Mapupunta ito sa **Master Admin Confirmation Queue** upang i-confirm ng Master Admin, i-encode sa `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` sa Google Drive, at i-save ang image/file sa Google Drive."
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
                "📸 I-upload ang Baptismal Image(s) o File(s) (JPG, PNG, WEBP, PDF):",
                type=["jpg", "jpeg", "png", "webp", "pdf"],
                accept_multiple_files=True,
            )
            if st.form_submit_button("📤 Ipasa ang Bagong Baptized sa Master Admin para I-confirm at I-encode sa Google Drive", type="primary"):
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
                    st.error(redact_sensitive_text(ex))

        st.divider()
        st.markdown(f"#### 📂 Mga Naipasa Kong Bagong Baptized sa Master Admin ({assigned_parish})")
        subs = reg_svc.list_submissions_for_actor(account_id)
        if not subs:
            st.info("Wala pang naipapasang new baptism submission mula sa iyong parokya.")
        else:
            st.dataframe(subs, use_container_width=True, hide_index=True)
            editable_subs = [s for s in subs if s.get("verification_status") in EDITABLE_REGISTRATION_STATUSES]
            if editable_subs:
                with st.expander("✏️ I-edit ang Submission Habang Pending Verification Pa sa Master Admin"):
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
                                st.error(redact_sensitive_text(ex))

    with ptab_req_copy:
        st.subheader(f"📨 Historical Record & Certified Certificate Requests — Humingi ng Dating Copy sa Master Admin ({assigned_parish})")
        st.info(
            "📌 Gamitin ang form sa ibaba upang humingi o mag-request ng dating copy / official certified baptismal certificate mula sa Master Admin."
        )
        with st.form("unified_parish_leader_request_copy_form"):
            req_kind = st.radio(
                "Uri ng Hinihinging Copy Mula kay Master Admin:",
                [
                    "CERTIFIED_CERTIFICATE — Official Certified Baptismal Certificate Copy (Dating Copy para Ma-Print)",
                    "HISTORICAL_RECORD — Historical Baptismal Record Verification & Copy",
                ],
                horizontal=True,
            )
            c_h1, c_h2 = st.columns(2)
            with c_h1:
                child_nm = st.text_input("Name of Baptized Person (Required):")
                father_nm = st.text_input("Father's Name (if known):")
                mother_nm = st.text_input("Mother's Name (if known):")
                approx_bap = st.text_input("Approximate Baptism Date / Year:")
            with c_h2:
                approx_dob = st.text_input("Approximate Birth Date:")
                st.text_input("Requesting Parish Leader Admin (Locked):", value=assigned_parish, disabled=True)
                bk_hint = st.text_input("Book Number (if known):")
                pg_hint = st.text_input("Page Number / Record ID (if known):")
            reason_txt = st.text_area("Reason / Purpose for Requesting Copy from Master Admin (Required):")
            sup_doc = st.file_uploader("Supporting Document / ID (Optional — JPG, PNG, PDF):", type=["jpg", "jpeg", "png", "pdf"])
            if st.form_submit_button("📨 Ipadala ang Request ng Dating Copy / Certificate sa Master Admin", type="primary"):
                sup_fn = ""
                if sup_doc is not None:
                    ok_f, msg_f, safe_fn = validate_uploaded_file(sup_doc.read(), sup_doc.name)
                    if not ok_f:
                        st.error(msg_f)
                        st.stop()
                    sup_fn = safe_fn
                selected_req_type = "CERTIFIED_CERTIFICATE" if req_kind.startswith("CERTIFIED_CERTIFICATE") else "HISTORICAL_RECORD"
                try:
                    rid = temp_svc.submit_request(
                        requester_account_id=account_id,
                        requester_username=current_user["username"],
                        requested_person_name=child_nm,
                        purpose=reason_txt,
                        request_type=selected_req_type,
                        father_name=father_nm,
                        mother_name=mother_nm,
                        approx_baptism_date=approx_bap,
                        approx_birth_date=approx_dob,
                        parish=assigned_parish,
                        book_hint=bk_hint,
                        page_hint=pg_hint,
                        supporting_doc_filename=sup_fn,
                    )
                    st.success(f"✅ Naipadala na sa Master Admin ang iyong request! Request ID: `{rid}`")
                    st.rerun()
                except Exception as ex:
                    st.error(redact_sensitive_text(ex))

        st.divider()
        st.markdown(f"#### 📋 My Historical Record & Certified Certificate Requests ({assigned_parish})")
        my_reqs = temp_svc.list_requests(actor_account_id=account_id)
        if my_reqs:
            st.dataframe(my_reqs, use_container_width=True, hide_index=True)
        else:
            st.info("Wala pang naipadalang request sa Master Admin.")

    with ptab_approved_print:
        st.subheader(f"🖨️ Approved Copies & Certificates Mula kay Master Admin (Print / Download) — {assigned_parish}")
        st.caption("Dito lalabas ang lahat ng dating copy at official baptismal certificates na in-approve o ibinigay ng Master Admin sa iyong parokya.")
        rel_certs = cert_svc.list_certificates(actor_account_id=account_id)
        if not rel_certs:
            st.info("Wala pang opisyal na certificate o dating copy na na-release ng Master Admin sa iyong parokya.")
        else:
            st.dataframe(rel_certs, use_container_width=True, hide_index=True)
            sel_cid = st.selectbox("Piliin ang Approved Certificate Copy para I-print o I-download:", [c["certificate_id"] for c in rel_certs])
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
                    st.error(redact_sensitive_text(ex))

        my_grants = temp_svc.list_active_grants(account_id=account_id)
        if my_grants:
            st.divider()
            st.markdown("#### 📄 Mga Historical Record Details na Ibinigay ng Master Admin sa Iyong Parokya")
            for g in my_grants:
                rec_id = g["record_id"]
                if authz_svc.has_temporary_record_permission(account_id, rec_id, "view_record"):
                    rec_data = hist_svc.get_record_by_id(rec_id, actor_account_id=account_id)
                    if rec_data:
                        st.success(f"Released Record `{rec_id}` — **{rec_data.get('Name of the Child', '')}** (Expires: {g['expires_at']})")
                        st.json(rec_data)


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

    # Ensure 1 Master Admin (`ctg.Baptismal.admin`) + 9 Parish Leader Admin accounts exist in DB
    _resolve_or_provision_hub_synced_user(auth_svc, is_master=True)

    hub_is_authenticated = bool(embedded_in_hub and st.session_state.get("authenticated", False))
    hub_true_login_role = str(st.session_state.get("login_role") or "")
    hub_active_view_role = str(st.session_state.get("user_role") or hub_true_login_role or "")
    hub_active_parish = str(st.session_state.get("assigned_parish") or st.session_state.get("login_parish") or "").strip()

    is_hub_view_master = (not hub_is_authenticated) or (hub_active_view_role in ("Master Admin", "Super Admin"))
    is_hub_true_master = hub_true_login_role in ("Master Admin", "Super Admin")

    # =========================================================================
    # 1. 9 PARISH LEADER ADMINS -> NO EXTRA LOGIN GATE (DIRECT TO PARISH DASHBOARD)
    # =========================================================================
    if hub_is_authenticated and not is_hub_view_master:
        if is_excluded_parish(hub_active_parish):
            st.markdown(
                "<div class='ctg-bap-banner'>"
                "<h2>⛪ CTG BAPTISMAL RECORDS</h2>"
                "<p>Official 9 Philippine Parishes Only (Hong Kong Excluded)</p>"
                "</div>",
                unsafe_allow_html=True,
            )
            st.error(
                "⛔ **Hong Kong Parish is Excluded:** Ang Baptismal System ay para lamang sa **9 Official Philippine Parishes** "
                f"({', '.join(OFFICIAL_NINE_PARISHES)})."
            )
            if is_hub_true_master:
                if st.button("👑 Bumalik sa Master Admin View", type="primary"):
                    st.session_state["user_role"] = "Master Admin"
                    st.session_state["assigned_parish"] = "Lahat ng Parokya"
                    st.rerun()
            return

        try:
            norm_hub_p = normalize_official_nine_parish(hub_active_parish)
        except Exception:
            norm_hub_p = OFFICIAL_NINE_PARISHES[0]

        parish_user_obj = _resolve_or_provision_hub_synced_user(auth_svc, is_master=False, parish_name=norm_hub_p)
        if is_hub_true_master:
            adm_c1, adm_c2 = st.columns([3.0, 1.2])
            with adm_c1:
                st.info(f"👁️ **Master Admin Preview:** Tinitingnan mo ang **Parish Leader Admin Dashboard** ng **`{norm_hub_p}`**.")
            with adm_c2:
                if st.button("👑 Bumalik sa Master Admin View", type="primary", use_container_width=True):
                    st.session_state["user_role"] = "Master Admin"
                    st.session_state["assigned_parish"] = "Lahat ng Parokya"
                    st.rerun()

        _render_parish_leader_admin_dashboard(str(parish_user_obj["account_id"]), parish_user_obj, svc)
        return

    # =========================================================================
    # 2. MASTER ADMIN ONLY -> DEDICATED "CTG BAPTISMAL RECORDS" LOGIN GATE
    # =========================================================================
    now_epoch = _time_mod.time()
    last_bap_act = float(st.session_state.get("_bap_last_activity_ts", now_epoch))
    if st.session_state.get("bap_master_authenticated") and (now_epoch - last_bap_act) > 1800:
        st.session_state.pop("bap_master_authenticated", None)
        st.session_state.pop("bap_auth_user", None)
    st.session_state["_bap_last_activity_ts"] = now_epoch

    valid_sess, current_user = auth_svc.validate_session(st.session_state.get("bap_auth_user"))
    if not st.session_state.get("bap_master_authenticated") or not valid_sess or not current_user or current_user.get("role") != "Master Admin":
        st.session_state.pop("bap_master_authenticated", None)
        st.session_state.pop("bap_auth_user", None)
        st.markdown(
            "<div class='ctg-bap-banner'>"
            "<h2>⛪ CTG BAPTISMAL RECORDS</h2>"
            "<p>Master Admin Archival Vault & Certification Hub</p>"
            "</div>",
            unsafe_allow_html=True,
        )
        with st.form("ctg_baptismal_records_master_login_form"):
            in_u = st.text_input("User Name:")
            in_p = st.text_input("Password:", type="password")
            if st.form_submit_button("Sign In to CTG BAPTISMAL RECORDS", type="primary"):
                u_clean = str(in_u or "").strip().lower()
                if u_clean == MASTER_ADMIN_BAPTISMAL_USERNAME and verify_password_pbkdf2(in_p, MASTER_ADMIN_BAPTISMAL_PW_HASH):
                    master_obj = _resolve_or_provision_hub_synced_user(auth_svc, is_master=True)
                    st.session_state["bap_auth_user"] = master_obj
                    st.session_state["bap_master_authenticated"] = True
                    st.session_state["_bap_last_activity_ts"] = _time_mod.time()
                    st.rerun()
                else:
                    ok_db, db_user, msg_db = auth_svc.authenticate(u_clean, in_p)
                    if ok_db and db_user and db_user.get("role") == "Master Admin":
                        st.session_state["bap_auth_user"] = db_user
                        st.session_state["bap_master_authenticated"] = True
                        st.session_state["_bap_last_activity_ts"] = _time_mod.time()
                        st.rerun()
                    else:
                        st.error("Invalid Master Admin User Name or Password.")
        return

    account_id = str(current_user["account_id"])

    # =================================================================
    # MASTER ADMIN CENTRAL DASHBOARD (CLEAN HORIZONTAL TABS, NO LOG DUMPS)
    # =================================================================
    st.markdown(
        "<div class='ctg-bap-banner'>"
        "<h2>⛪ CTG BAPTISMAL RECORDS — Master Admin Archive & Certification Hub</h2>"
        "<p>Central Control Over All 9 Official Parishes • BOOK 1–9 Visual Archive • Historical Verification & Official Certificate Issuance (Hong Kong Excluded)</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    hdr_c1, hdr_c2, hdr_c3 = st.columns([2.4, 1.4, 0.9])
    with hdr_c1:
        st.caption(
            f"👑 Signed in as **ctg.Baptismal.admin** (`Master Admin`) • "
            f"📂 [CTG BAPTISMAL RECORDS Drive Folder](https://drive.google.com/drive/folders/{DEFAULT_BAPTISMAL_ROOT_FOLDER_ID}) • "
            f"📗 [Master Baptismal Excel](https://drive.google.com/file/d/{DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID}/view)"
        )
    with hdr_c2:
        if st.button("🔄 Sync 4,093 Records Mula Google Drive", use_container_width=True, type="primary", key="bap_top_sync_drive_btn"):
            with st.spinner("Binabasa ang OVERALL BAPTISMAL CERTIFICATE LIST.xlsx mula sa Google Drive..."):
                ok_s, cnt_s, msg_s = hist_svc.auto_sync_from_drive_or_disk(drive_mgr, force=True)
                if ok_s:
                    st.success(f"✅ {msg_s}")
                    st.rerun()
                else:
                    st.warning(redact_sensitive_text(msg_s))
    with hdr_c3:
        if st.button("🔒 Sign Out (CTG BAPTISMAL RECORDS)", use_container_width=True, key="bap_sign_out_btn"):
            auth_svc.logout(current_user)
            st.session_state.pop("bap_master_authenticated", None)
            st.session_state.pop("bap_auth_user", None)
            st.rerun()

    preview_options = ["👑 Master Admin Central Hub (All 9 Parishes & BOOK 1–9 Archive)"] + [
        f"⛪ {p} — Parish Leader Admin Dashboard (Leader View)" for p in OFFICIAL_NINE_PARISHES
    ]
    picked_preview = st.selectbox(
        "🔀 Switch Baptismal Dashboard View (Master Admin Hub vs 9 Parish Leader Admin Dashboards):",
        preview_options,
        key="bap_top_role_switcher",
    )
    if not picked_preview.startswith("👑"):
        chosen_p_name = picked_preview.replace("⛪ ", "").split(" — ")[0].strip()
        p_user_obj = _resolve_or_provision_hub_synced_user(auth_svc, is_master=False, parish_name=chosen_p_name)
        _render_parish_leader_admin_dashboard(str(p_user_obj["account_id"]), p_user_obj, svc)
        return

    if not st.session_state.get("_bap_auto_sync_attempted"):
        st.session_state["_bap_auto_sync_attempted"] = True
        cur_stats_check = hist_svc.get_dashboard_statistics(actor_account_id=account_id)
        if cur_stats_check["total_historical"] == 0:
            # Fast local-disk-only check on initial login so login never hangs waiting for network
            hist_svc.auto_sync_from_drive_or_disk(drive_mgr=None, force=False)

    (
        tab_dash,
        tab_search_verify,
        tab_books,
        tab_confirm_queue,
        tab_reqs_share,
        tab_new_reg,
        tab_cert_gen,
    ) = st.tabs([
        "📊 1. Dashboard",
        "🔍 2. Search ID, Book Image & Print/Share",
        "🖼️ 3. BOOK 1–9 Archive",
        "📥 4. Confirm Leaders' Uploads (Encode to Drive)",
        "📨 5. Requests & Release",
        "✍️ 6. Direct New Baptism (Master)",
        "📜 7. Certificates Archive",
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
        st.subheader("🔍 Search Record ID (Palitan Lang ang Huling Number), Verify Book Image, Print & Ibigay sa Parokya/Username")

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
                    if bk_drive_id:
                        st.caption(
                            f"📂 [Buksan ang `{sel_bk_for_rec}` Folder sa Google Drive](https://drive.google.com/drive/folders/{bk_drive_id})"
                        )

                    if cand_drive_imgs:
                        st.success(f"✨ Nakahanap ng {len(cand_drive_imgs)} tugmang scanned image(s) sa Google Drive para sa `{detail.get('Name of the Child', '')}` / Page `{rec_pg_num}`:")
                    if bk_drive_id:
                        sc_b1, sc_b2 = st.columns(2)
                        with sc_b1:
                            if st.button(f"🔍 Hanapin ang Image ng {sel_id} sa {sel_bk_for_rec}", key=f"btn_qfind_{sel_id}", use_container_width=True):
                                try:
                                    s_terms = []
                                    if rec_pg_num and rec_pg_num.isdigit():
                                        pi = int(rec_pg_num)
                                        s_terms.extend([f"BC_{pi}.", f"BC_{pi:03d}.", f"BC_{pi:04d}."])
                                    c_toks = [t for t in re.split(r"[^a-zA-Z0-9]+", str(detail.get("Name of the Child") or "")) if len(t) >= 3]
                                    s_terms.extend(c_toks[:2])
                                    st.session_state[f"_drv_rec_search_{sel_id}"] = drive_mgr.search_files_in_folder(bk_drive_id, s_terms, page_size=40)
                                    st.rerun()
                                except Exception:
                                    st.info(f"📂 [Buksan ang {sel_bk_for_rec} sa Google Drive](https://drive.google.com/drive/folders/{bk_drive_id})")
                        with sc_b2:
                            if st.button(f"📂 I-load Lahat ng Pahina sa {sel_bk_for_rec}", key=f"btn_scan_bk_{sel_id}", use_container_width=True):
                                try:
                                    st.session_state[f"_drv_book_files_{sel_bk_for_rec}"] = drive_mgr.list_files_in_folder(bk_drive_id, page_size=1000)
                                    st.rerun()
                                except Exception:
                                    st.info(f"📂 [Buksan ang {sel_bk_for_rec} sa Google Drive](https://drive.google.com/drive/folders/{bk_drive_id})")

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
                                except Exception:
                                    st.session_state[f"_embed_rec_img_{sel_id}"] = (chosen_obj["name"], chosen_drv_id, chosen_obj.get("book_folder", sel_bk_for_rec))
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
                        elif st.session_state.get(f"_embed_rec_img_{sel_id}"):
                            em_nm, em_id, em_bk = st.session_state[f"_embed_rec_img_{sel_id}"]
                            _render_drive_file_embed_fallback(em_id, em_nm, em_bk)

                    if linked_assocs:
                        st.caption("🔗 Naka-link na Scanned Images sa Record na Ito:")
                        st.dataframe(linked_assocs, use_container_width=True, hide_index=True)

                # =========================================================
                # 📤 IBIGAY SA PAROKYA / PARISH LEADER ADMIN PARA MA-PRINT DIN NILA
                # =========================================================
                st.divider()
                st.markdown(f"#### 📤 Ibigay ang Record at Printable Certificate (`{sel_id}`) sa Parish Leader Admin para Ma-Print Nila")
                p_admins_all = auth_svc.list_parish_admin_accounts()
                uname_choices = [f"{a['assigned_parish']} — {a['username']}" for a in p_admins_all]
                sh_c1, sh_c2, sh_c3 = st.columns([2.0, 1.0, 1.4])
                with sh_c1:
                    picked_u_lbl = st.selectbox(
                        "Piliin ang Tatanggap na Parish Leader Admin:",
                        uname_choices,
                        key=f"sh_sel_u_{sel_id}",
                    )
                with sh_c2:
                    share_hrs = st.number_input("Valid Hours:", min_value=1.0, max_value=720.0, value=72.0, key=f"sh_hrs_{sel_id}")
                with sh_c3:
                    st.write("")
                    if st.button("📤 Ibigay & I-release para Ma-Print Nila", type="primary", use_container_width=True, key=f"btn_sh_u_{sel_id}"):
                        target_uname = picked_u_lbl.split(" — ")[-1].strip()
                        try:
                            sh_res = cert_svc.share_record_and_certificate_with_username(
                                admin_account_id=account_id,
                                record_id=sel_id,
                                target_username=target_uname,
                                duration_hours=share_hrs,
                            )
                            st.success(
                                f"✅ Naibigay na ang Record `{sel_id}` at Printable Certificate `{sh_res['certificate_id']}` "
                                f"sa **`{picked_u_lbl.split(' — ')[0]}`**! Makikita at mapi-print na nila ito sa `🖨️ 3. Approved Copies Mula kay Master Admin` tab."
                            )
                        except Exception as ex:
                            st.error(redact_sensitive_text(ex))

    with tab_books:
        st.subheader("🖼️ BOOK 1–9 Visual Archive & Google Drive Scanned Pages")
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
                    except Exception:
                        loaded_counts[bk] = f"Direct Link Ready ({fid_b})"
            st.success(f"✅ Na-load ang BOOK 1–9 mula sa Google Drive: {loaded_counts}")
            st.rerun()
        sel_book = st.selectbox("Select Book Folder:", HISTORICAL_BOOKS)
        b_fid = book_folders.get(sel_book, "")
        if not b_fid:
            st.info(f"Walang nakitang `{sel_book}` folder sa Google Drive.")
        else:
            st.caption(f"📂 Google Drive Folder ID (`{sel_book}`): `{b_fid}` — [Buksan sa Google Drive](https://drive.google.com/drive/folders/{b_fid})")
            if st.button(f"🔍 I-load ang Listahan ng Scanned Images sa {sel_book} Mula sa Google Drive", key=f"btn_list_drv_{sel_book}"):
                try:
                    drv_imgs = drive_mgr.list_files_in_folder(b_fid, page_size=1000)
                    st.session_state[f"_drv_book_files_{sel_book}"] = drv_imgs
                except Exception:
                    st.info(f"📂 [Buksan ang `{sel_book}` Direkta sa Google Drive](https://drive.google.com/drive/folders/{b_fid})")
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
                    except Exception:
                        _render_drive_file_embed_fallback(chosen_id, chosen_nm, sel_book)

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
                        st.error(redact_sensitive_text(ex))
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
                f"**Person:** `{sel_req_obj.get('requested_person_name')}` • "
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
                            f"✅ Official Certificate `{cert_out['certificate_id']}` generated, approved, and released to **`{sel_req_obj.get('assigned_parish')}`** para ma-print na nila!"
                        )
                        st.rerun()
                    except Exception as ex:
                        st.error(redact_sensitive_text(ex))
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
                    st.error(redact_sensitive_text(ex))

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
                st.error(redact_sensitive_text(ex))

        st.divider()
        st.markdown("#### 🗂️ Generated Certificates Archive")
        certs_list = cert_svc.list_certificates(actor_account_id=account_id)
        st.dataframe(certs_list, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    st.set_page_config(
        page_title="CTG BAPTISMAL RECORDS",
        page_icon="⛪",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_baptismal_portal(embedded_in_hub=False)

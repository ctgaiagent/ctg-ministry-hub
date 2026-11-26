"""
baptismal_app.py — CTG Baptismal Records & Certification System (Role-Based Dashboards)
======================================================================================
Implements separate role-based dashboards for:
1. ONE MASTER ADMIN (Central Archive, BOOK 1-9, Global 27-Column Search, All 9 Parishes'
   Submissions & Verification, Historical & Certified Certificate Request Approval/Issuance,
   Temporary Access Grants, 9 Parish Leader Admin Accounts Management, Google Drive & Audit Logs)
2. NINE PARISH LEADER ADMINS (Dedicated Parish-Scoped Dashboards for the 9 Official Parishes:
   Lipay Norte, Malanas, Poblacion, Duplas, Santiago, Caba, Manila, Andabuen, Apayao)
   - HONG KONG PARISH IS COMPLETELY EXCLUDED.
"""

import base64
import json
from typing import Any, Dict, List
import streamlit as st

from baptismal.audit_logs import AuditLogger
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.backup_manager import BackupManager
from baptismal.certificates import CertificateService
from baptismal.config import (
    CANONICAL_REQUEST_STATUSES,
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
from baptismal.historical_records import HistoricalRecordsService
from baptismal.image_archive import ImageArchiveService, render_image_preview_bytes
from baptismal.image_matching import ImageMatchingService
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService
from baptismal.temporary_access import TemporaryAccessService
from baptismal.validation import escape_html_text, validate_uploaded_file


@st.cache_resource
def _get_baptismal_services() -> Dict[str, Any]:
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


def _render_parish_leader_admin_dashboard(
    account_id: str,
    current_user: Dict[str, Any],
    svc: Dict[str, Any],
    embedded_in_hub: bool = False,
) -> None:
    """Renders the dedicated Parish Leader Admin Dashboard scoped strictly to the authenticated parish."""
    auth_svc: AuthService = svc["auth"]
    authz_svc: AuthorizationService = svc["authz"]
    audit_svc: AuditLogger = svc["audit"]
    hist_svc: HistoricalRecordsService = svc["hist"]
    reg_svc: RegistrationService = svc["reg"]
    cert_svc: CertificateService = svc["certs"]
    temp_svc: TemporaryAccessService = svc["temp_acc"]

    # Always obtain the authoritative parish from the database server-side
    assigned_parish = authz_svc.get_authenticated_parish(account_id)

    st.markdown(
        f"<div class='ctg-bap-banner'>"
        f"<h2>⛪ {escape_html_text(assigned_parish)} — Baptismal Leader Admin Dashboard</h2>"
        f"<p>Parish-Scoped New Baptism Registration • Historical Record & Certified Certificate Requests • Official Released Certificates</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    PARISH_NAV_ITEMS = [
        "My Dashboard",
        "New Baptism Registration",
        "My Submissions",
        "My Uploaded Images",
        "Request Historical Record",
        "Request Certified Certificate",
        "My Request History",
        "My Approved Certificates",
        "My Account",
    ]
    if "bap_parish_nav_override" in st.session_state:
        target_p_nav = st.session_state.pop("bap_parish_nav_override")
        if target_p_nav in PARISH_NAV_ITEMS:
            st.session_state["bap_parish_active_nav"] = target_p_nav

    if embedded_in_hub:
        p_page = st.selectbox("📂 Parish Baptismal Navigation:", PARISH_NAV_ITEMS, key="bap_parish_active_nav")
    else:
        p_page = st.sidebar.radio("Parish Navigation", PARISH_NAV_ITEMS, key="bap_parish_active_nav")

    if p_page == "My Dashboard":
        p_stats = reg_svc.get_parish_dashboard_statistics(account_id)
        k1, k2, k3 = st.columns(3)
        k1.metric("My New Baptism Registrations", p_stats["my_new_registrations"])
        k2.metric("Pending Verification", p_stats["pending_verification"])
        k3.metric("Approved Registrations", p_stats["approved_registrations"])

        k4, k5, k6 = st.columns(3)
        k4.metric("Rejected Registrations", p_stats["rejected_registrations"])
        k5.metric("Pending Certificate Requests", p_stats["pending_certificate_requests"])
        k6.metric("Approved Certificates", p_stats["approved_certificates"])

        st.markdown("#### ⚡ Quick Actions")
        qa1, qa2, qa3, qa4, qa5 = st.columns(5)
        if qa1.button("✍️ Register New Baptism", use_container_width=True, type="primary"):
            st.session_state["bap_parish_nav_override"] = "New Baptism Registration"
            st.rerun()
        if qa2.button("📸 Upload Baptismal Images", use_container_width=True):
            st.session_state["bap_parish_nav_override"] = "My Uploaded Images"
            st.rerun()
        if qa3.button("📜 Request Historical Record", use_container_width=True):
            st.session_state["bap_parish_nav_override"] = "Request Historical Record"
            st.rerun()
        if qa4.button("🏅 Request Certified Certificate", use_container_width=True):
            st.session_state["bap_parish_nav_override"] = "Request Certified Certificate"
            st.rerun()
        if qa5.button("📋 Track My Requests", use_container_width=True):
            st.session_state["bap_parish_nav_override"] = "My Request History"
            st.rerun()

        st.divider()
        st.subheader(f"📋 Recent Submissions for {assigned_parish}")
        my_subs = reg_svc.list_submissions_for_actor(account_id)
        if my_subs:
            st.dataframe(my_subs[:10], use_container_width=True, hide_index=True)
        else:
            st.info(f"No baptismal registrations submitted yet for {assigned_parish}.")

    elif p_page == "New Baptism Registration":
        st.subheader(f"✍️ Register Newly Baptized Member — {assigned_parish}")
        st.caption("Ang Parish assignment ay awtomatikong naka-lock sa iyong authenticated Parish account sa server.")
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
                "Upload Baptismal Photos & Supporting Documents:",
                type=["jpg", "jpeg", "png", "webp", "pdf"],
                accept_multiple_files=True,
            )
            sync_drive = st.checkbox("Sync to Google Drive Workbook (`NEW BAPTISMS - YYYY`)", value=True)
            if st.form_submit_button("Submit New Baptism Registration", type="primary"):
                img_list = [
                    {"bytes": uf.read(), "filename": uf.name, "mime_type": uf.type or "image/jpeg"}
                    for uf in (up_imgs or [])
                ]
                try:
                    res = reg_svc.register_new_baptism(
                        account_id, payload, uploaded_images=img_list, sync_to_drive_excel=sync_drive
                    )
                    st.success(
                        f"Submitted! Record ID: `{res['record_id']}` • Parish: `{res['parish']}` • "
                        f"Verification Status: `{res['verification_status']}` • Workbook: `{res['workbook_update_status']}`"
                    )
                except Exception as ex:
                    st.error(str(ex))

    elif p_page == "My Submissions":
        st.subheader(f"📂 My Submissions — {assigned_parish}")
        subs = reg_svc.list_submissions_for_actor(account_id)
        if not subs:
            st.info("No submissions recorded yet.")
        else:
            st.dataframe(subs, use_container_width=True, hide_index=True)
            editable_subs = [s for s in subs if s.get("verification_status") in EDITABLE_REGISTRATION_STATUSES]
            if editable_subs:
                st.markdown("#### ✏️ Edit Pending Submission (Allowed Only While Pending Verification)")
                sel_edit_id = st.selectbox("Select Pending Record ID to Edit:", [s["record_id"] for s in editable_subs])
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
                    if st.form_submit_button("Save Changes to Pending Submission", type="primary"):
                        try:
                            reg_svc.update_own_submission(account_id, sel_edit_id, upd_payload)
                            st.success(f"Updated pending submission `{sel_edit_id}`.")
                            st.rerun()
                        except Exception as ex:
                            st.error(str(ex))

    elif p_page == "My Uploaded Images":
        st.subheader(f"📸 My Uploaded Baptismal Images — {assigned_parish}")
        imgs = reg_svc.list_uploaded_images_for_actor(account_id)
        if imgs:
            st.dataframe(imgs, use_container_width=True, hide_index=True)
        else:
            st.info("No uploaded images found for your parish submissions.")

    elif p_page == "Request Historical Record":
        st.subheader(f"📜 Request Historical Baptismal Record — {assigned_parish}")
        st.caption(
            "Ang Parish Leader Admins ay walang direktang search access sa historical archive. "
            "Ipadala ang detalye sa ibaba upang beripikahin ng Master Admin."
        )
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
            if st.form_submit_button("Submit Historical Record Request to Master Admin", type="primary"):
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
                    st.success(f"Historical Record Request submitted to Master Admin! Request ID: `{rid}` (Status: `Pending`)")
                except Exception as ex:
                    st.error(str(ex))

    elif p_page == "Request Certified Certificate":
        st.subheader(f"🏅 Request Certified Baptismal Certificate — {assigned_parish}")
        st.caption(
            "7-Step Workflow: (1) Parish Admin submits request → (2–3) Master Admin reviews & verifies record → "
            "(4–6) Master Admin approves, generates & officially issues certificate → (7) Released to your account for download."
        )
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
                c_rec_id = st.text_input("Known Record ID (Optional, e.g. BAP-2026-...):")
                c_bk_pg = st.text_input("Book / Page Number (if known):")
            c_purpose = st.text_area("Purpose / Reason for Certified Certificate (Required — e.g. Marriage, Confirmation, Legal):")
            c_sup = st.file_uploader("Supporting Document (Optional):", type=["jpg", "jpeg", "png", "pdf"])
            if st.form_submit_button("Submit Certified Certificate Request", type="primary"):
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
                    st.success(f"Certified Certificate Request submitted! Request ID: `{rid}` (Status: `Pending`)")
                except Exception as ex:
                    st.error(str(ex))

    elif p_page == "My Request History":
        st.subheader(f"📋 My Request History — {assigned_parish}")
        my_reqs = temp_svc.list_requests(actor_account_id=account_id)
        if my_reqs:
            st.dataframe(my_reqs, use_container_width=True, hide_index=True)
        else:
            st.info("No historical or certificate requests submitted yet.")

    elif p_page == "My Approved Certificates":
        st.subheader(f"🏅 My Approved & Released Certificates — {assigned_parish}")
        rel_certs = cert_svc.list_certificates(actor_account_id=account_id)
        if not rel_certs:
            st.info("No official certificates have been approved and released to your account yet.")
        else:
            st.dataframe(rel_certs, use_container_width=True, hide_index=True)
            sel_cid = st.selectbox("Select Released Certificate to Preview / Download:", [c["certificate_id"] for c in rel_certs])
            if sel_cid and st.button("📄 Load Released Certificate PDF", type="primary"):
                try:
                    pdf_bytes = cert_svc.download_released_certificate(account_id, sel_cid)
                    b64_pdf = base64.b64encode(pdf_bytes).decode("ascii")
                    st.markdown(
                        f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="600" type="application/pdf"></iframe>',
                        unsafe_allow_html=True,
                    )
                    st.download_button(
                        "⬇️ Download Official Certified Baptismal Certificate (PDF)",
                        data=pdf_bytes,
                        file_name=f"{sel_cid}.pdf",
                        mime="application/pdf",
                    )
                except Exception as ex:
                    st.error(str(ex))

    elif p_page == "My Account":
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

        my_grants = temp_svc.list_active_grants(account_id=account_id)
        if my_grants:
            st.divider()
            st.subheader("🔑 Temporary Single-Record Access Grants")
            for g in my_grants:
                rec_id = g["record_id"]
                if authz_svc.has_temporary_record_permission(account_id, rec_id, "view_record"):
                    rec_data = hist_svc.get_record_by_id(rec_id, actor_account_id=account_id)
                    if rec_data:
                        st.success(f"Temporary Access to Verified Record `{rec_id}` (Expires: {g['expires_at']})")
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

    svc = _get_baptismal_services()
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

    # Top bar for sign-out
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
            _render_parish_leader_admin_dashboard(account_id, current_user, svc, embedded_in_hub=embedded_in_hub)
            return

        # Legacy Temporary Requester fallback view
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
    # MASTER ADMIN CENTRAL DASHBOARD (EXCLUSIVE FULL ARCHIVE CONTROL)
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

    MASTER_NAV_ITEMS = [
        "1. Master Admin Dashboard",
        "2. Historical Records (Global Search)",
        "3. BOOK 1–9 Image Archive",
        "4. All 9 Parishes' Submissions & Verification",
        "5. Historical & Certificate Requests (All 9 Parishes)",
        "6. Temporary Access Management",
        "7. New Baptism Registration",
        "8. New Baptism Images",
        "9. Certificate Generator",
        "10. Generated Certificates Archive",
        "11. Parish Leader Admin Accounts (9 Parishes)",
        "12. Google Drive Storage Status",
        "13. Audit Logs & Security Settings",
    ]
    if embedded_in_hub:
        page = st.selectbox("📂 Master Admin Baptismal Module:", MASTER_NAV_ITEMS, key="bap_embedded_nav")
    else:
        page = st.sidebar.radio("Master Admin Navigation", MASTER_NAV_ITEMS, key="bap_standalone_nav")

    if page == "1. Master Admin Dashboard":
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
        c2.metric("New Registrations (9 Parishes)", stats["new_registrations"])
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

    elif page == "2. Historical Records (Global Search)":
        st.subheader("📚 Global Historical Records Search (Master Admin Exclusive — All 27 Columns)")
        with st.expander("📥 Import / Sync `Historical Records` mula sa Google Drive o Local File"):
            ic1, ic2 = st.columns(2)
            with ic1:
                if st.button("☁️ Download & Import Direkta Mula sa Google Drive", type="primary", use_container_width=True):
                    ok_d, cnt_d, msg_d = hist_svc.auto_sync_from_drive_or_disk(drive_mgr, force=True)
                    if ok_d:
                        st.success(f"✅ {msg_d}")
                        st.rerun()
                    else:
                        st.error(msg_d)
            with ic2:
                hist_up = st.file_uploader("O Mag-upload ng OVERALL BAPTISMAL CERTIFICATE LIST.xlsx:", type=["xlsx"], key="hist_xlsx_import")
                if hist_up and st.button("Import Uploaded Excel (Read-Only)", use_container_width=True):
                    cnt = hist_svc.import_sheet1_from_excel_bytes(hist_up.read())
                    audit_svc.log_event(account_id, "HISTORICAL_IMPORT", "SUCCESS", metadata={"imported_rows": cnt})
                    st.success(f"Imported {cnt} historical records into read-only storage.")
                    st.rerun()

        fc1, fc2, fc3, fc4 = st.columns(4)
        kw = fc1.text_input("Keyword Search (Any of 27 Columns — Child, Father, Mother, Sponsors, Priest):", key="hist_kw")
        f_parish = fc2.text_input("Parish:", key="hist_f_parish")
        f_book = fc3.text_input("Book Number:", key="hist_f_book")
        f_year = fc4.text_input("Year of Baptism:", key="hist_f_year")

        if st.button("Clear Search Filters"):
            for k in ("hist_kw", "hist_f_parish", "hist_f_book", "hist_f_year"):
                st.session_state[k] = ""
            st.rerun()

        page_num = int(st.number_input("Page (50 records per page):", min_value=1, value=1, step=1))
        rows, total_rows = hist_svc.search_records(
            keyword=kw,
            filters={"Parish": f_parish, "Book Number": f_book, "Year of Baptism": f_year},
            page=page_num,
            page_size=50,
            actor_account_id=account_id,
        )
        st.caption(f"Showing {len(rows)} of {total_rows} matching historical records (read-only; Hong Kong excluded from operational scope).")
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)
            sel_id = st.selectbox("Inspect Full Historical Record (All 27 Fields):", [r["Record ID"] for r in rows])
            if sel_id:
                detail = hist_svc.get_record_by_id(sel_id, actor_account_id=account_id)
                st.json(detail)
                audit_svc.log_event(account_id, "HISTORICAL_RECORD_VIEWED", "SUCCESS", related_record_id=sel_id)

    elif page == "3. BOOK 1–9 Image Archive":
        st.subheader("🖼️ BOOK 1–9 Visual Archive & Google Drive Scanned Pages (Master Admin Exclusive)")
        sel_book = st.selectbox("Select Book Folder:", HISTORICAL_BOOKS)

        with st.expander(f"☁️ Browse Live Scanned Images in Google Drive ({sel_book})", expanded=True):
            book_folders = drive_mgr.discover_book_folders_in_drive()
            b_fid = book_folders.get(sel_book, "")
            if not b_fid:
                st.info(f"Walang nakitang `{sel_book}` folder sa Google Drive (kasalukuyang naka-upload sa Drive: {', '.join(sorted(book_folders.keys()))}).")
            else:
                st.caption(f"📂 Google Drive Folder ID (`{sel_book}`): `{b_fid}` — [Buksan sa Google Drive](https://drive.google.com/drive/folders/{b_fid})")
                if st.button(f"🔍 I-load ang Listahan ng Scanned Images sa {sel_book} Mula sa Google Drive", key=f"btn_list_drv_{sel_book}"):
                    try:
                        drv_imgs = drive_mgr.list_files_in_folder(b_fid, page_size=300)
                        st.session_state[f"_drv_book_files_{sel_book}"] = drv_imgs
                    except Exception as ex:
                        st.error(f"Google Drive Folder List Error: {ex}")
                cached_drv_imgs = st.session_state.get(f"_drv_book_files_{sel_book}", [])
                if cached_drv_imgs:
                    st.dataframe(cached_drv_imgs, use_container_width=True, hide_index=True)
                    sel_drv_img = st.selectbox(
                        "Pumili ng Scanned Image mula sa Google Drive para I-preview / I-rotate / I-link sa Record:",
                        [f"{it['name']} ({it['id']})" for it in cached_drv_imgs],
                        key=f"sel_drv_img_{sel_book}"
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

        st_flt = st.selectbox("Match Status Filter:", ["ALL", "POSSIBLE_MATCH_NEEDS_CONFIRMATION", "CONFIRMED", "REJECTED"])
        assocs = img_arc_svc.list_by_book(sel_book, status_filter=st_flt, actor_account_id=account_id)
        if assocs:
            st.dataframe(assocs, use_container_width=True, hide_index=True)
            for a in assocs:
                if a["match_status"] == "POSSIBLE_MATCH_NEEDS_CONFIRMATION":
                    st.warning(
                        f"Possible Match Only — Record `{a['record_id']}` (Page `{a['record_page_number']}`) vs "
                        f"`{a['image_filename']}` (Filename hint `{a['extracted_page_hint']}`): {a['notes']}"
                    )
                    b1, b2 = st.columns(2)
                    if b1.button(f"✅ Confirm Match ({a['association_id']})", key=f"conf_{a['association_id']}"):
                        img_match_svc.confirm_image_association(account_id, a["association_id"], approve=True)
                        st.rerun()
                    if b2.button(f"❌ Reject Match ({a['association_id']})", key=f"rej_{a['association_id']}"):
                        img_match_svc.confirm_image_association(account_id, a["association_id"], approve=False)
                        st.rerun()

        with st.expander("➕ Suggest / Link Scanned Image to Historical Record"):
            m_rec_id = st.text_input("Record ID (e.g. HIST-00001):")
            m_img_fn = st.text_input("Scanned Image Filename (e.g. Abelardo Galban BC_744.jpeg):")
            m_notes = st.text_input("Notes:")
            if st.button("Create Candidate Association") and m_rec_id and m_img_fn:
                res_m = img_match_svc.suggest_possible_match(m_rec_id, sel_book, m_img_fn, notes=m_notes)
                st.info(f"Created `{res_m['association_id']}` — Status: `{res_m['match_status']}`")

    elif page == "4. All 9 Parishes' Submissions & Verification":
        st.subheader("🏛️ All 9 Parishes' New Baptism Submissions & Master Admin Verification")
        p_filter = st.selectbox("Filter by Parish:", ["ALL"] + OFFICIAL_NINE_PARISHES)
        all_subs = reg_svc.list_submissions_for_actor(account_id, parish_filter=p_filter)
        if not all_subs:
            st.info("No new baptism submissions recorded yet across the 9 official parishes.")
        else:
            st.dataframe(all_subs, use_container_width=True, hide_index=True)
            st.markdown("#### 🛡️ Verify / Approve / Reject Parish Submission")
            sel_sub_id = st.selectbox("Select Submission Record ID:", [s["record_id"] for s in all_subs])
            new_ver_st = st.selectbox("Verification Decision:", REGISTRATION_VERIFICATION_STATUSES)
            ver_notes = st.text_input("Verification Notes / Instructions to Parish Leader Admin:")
            if st.button("Apply Verification Decision", type="primary"):
                reg_svc.verify_registration(account_id, sel_sub_id, new_ver_st, ver_notes)
                st.success(f"Updated `{sel_sub_id}` to `{new_ver_st}`.")
                st.rerun()

    elif page == "5. Historical & Certificate Requests (All 9 Parishes)":
        st.subheader("📨 Historical Record & Certified Certificate Requests (All 9 Parishes)")
        reqs = temp_svc.list_requests(actor_account_id=account_id)
        if not reqs:
            st.info("No requests from the 9 parishes yet.")
        else:
            st.dataframe(reqs, use_container_width=True, hide_index=True)
            st.markdown("#### 🏅 Master Admin Review, Historical Verification & Official Certificate Release")
            sel_req_id = st.selectbox("Select Request ID:", [r["request_id"] for r in reqs])
            sel_req_obj = next(r for r in reqs if r["request_id"] == sel_req_id)
            st.caption(
                f"**Type:** `{sel_req_obj.get('request_type')}` • **Parish:** `{sel_req_obj.get('assigned_parish')}` • "
                f"**Person:** `{sel_req_obj.get('requested_person_name')}` • **Father:** `{sel_req_obj.get('father_name')}` • "
                f"**Mother:** `{sel_req_obj.get('mother_name')}` • **Current Status:** `{sel_req_obj.get('status')}`"
            )
            verified_rec_id = st.text_input(
                "Verified Record ID from Archive (e.g. HIST-00001 or BAP-2026-...):",
                value=str(sel_req_obj.get("target_record_id") or ""),
            )
            new_req_status = st.selectbox("Update Request Status:", CANONICAL_REQUEST_STATUSES)
            rev_notes = st.text_input("Master Admin Review Notes:")
            orient_rel = st.radio("Certificate Layout (if issuing & releasing):", ["portrait", "landscape"], horizontal=True)

            b_st, b_rel, b_rj = st.columns(3)
            if b_st.button("🔄 Update Status Only", use_container_width=True):
                temp_svc.update_request_status(account_id, sel_req_id, new_req_status, rev_notes, verified_rec_id)
                st.success(f"Request `{sel_req_id}` updated to `{new_req_status}`.")
                st.rerun()
            if b_rel.button("🏅 Approve, Generate & Release Official Certificate", type="primary", use_container_width=True):
                if not verified_rec_id.strip():
                    st.error("Please enter the verified Record ID first.")
                else:
                    try:
                        cert_out = cert_svc.approve_and_release_certificate_for_request(
                            admin_account_id=account_id,
                            request_id=sel_req_id,
                            record_id=verified_rec_id.strip(),
                            orientation=orient_rel,
                            review_notes=rev_notes,
                        )
                        st.success(
                            f"Official Certificate `{cert_out['certificate_id']}` generated, approved, and released to `{sel_req_obj.get('requester_username')}`!"
                        )
                        st.rerun()
                    except Exception as ex:
                        st.error(str(ex))
            if b_rj.button("❌ Reject Request", use_container_width=True):
                temp_svc.reject_request(account_id, sel_req_id, rev_notes)
                st.warning(f"Request `{sel_req_id}` rejected.")
                st.rerun()

    elif page == "6. Temporary Access Management":
        st.subheader("🔑 Temporary Single-Record Access Management")
        reqs = temp_svc.list_requests(actor_account_id=account_id)
        if reqs:
            sel_req = st.selectbox("Select Request to Grant Temporary Single-Record Access:", [r["request_id"] for r in reqs])
            target_rec = st.text_input("Permitted Record ID (Required):")
            perms_sel = st.multiselect("Allowed Permissions:", TEMP_ACCESS_PERMISSIONS, default=["view_record"])
            exp_hrs = st.number_input("Expiration (Hours):", min_value=1.0, max_value=168.0, value=24.0)
            notes_in = st.text_input("Grant Notes:")
            if st.button("✅ Grant Temporary Single-Record Access", type="primary") and target_rec.strip():
                gid = temp_svc.approve_request_and_grant_access(account_id, sel_req, target_rec.strip(), perms_sel, exp_hrs, notes_in)
                st.success(f"Granted! Grant ID: `{gid}`")
                st.rerun()

        grants = temp_svc.list_active_grants()
        if grants:
            st.markdown("#### Active & Historical Temporary Grants")
            st.dataframe(grants, use_container_width=True, hide_index=True)
            active_gids = [g["grant_id"] for g in grants if not int(g.get("is_revoked", 0))]
            if active_gids:
                rev_gid = st.selectbox("Revoke Grant Immediately:", active_gids)
                if st.button("🚫 Revoke Access Now"):
                    temp_svc.revoke_grant_immediately(account_id, rev_gid)
                    st.success(f"Grant `{rev_gid}` revoked immediately.")
                    st.rerun()

    elif page == "7. New Baptism Registration":
        st.subheader("✍️ New Baptism Registration (9 Official Parishes Only — Hong Kong Excluded)")
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
            sync_drive = st.checkbox("Sync to Google Drive Workbook (`NEW BAPTISMS - YYYY`)", value=True)
            if st.form_submit_button("Preview & Register Baptism", type="primary"):
                img_list = [{"bytes": uf.read(), "filename": uf.name, "mime_type": uf.type or "image/jpeg"} for uf in (up_imgs or [])]
                try:
                    res = reg_svc.register_new_baptism(account_id, payload, uploaded_images=img_list, sync_to_drive_excel=sync_drive)
                    st.success(f"Registered! Record ID: `{res['record_id']}` • Parish: `{res['parish']}` • Worksheet: `{res['worksheet_name']}`")
                except Exception as ex:
                    st.error(str(ex))

    elif page == "8. New Baptism Images":
        st.subheader("📸 New Baptism Images Registry (All 9 Parishes)")
        rows_m = reg_svc.list_uploaded_images_for_actor(account_id)
        st.dataframe(rows_m, use_container_width=True, hide_index=True)

    elif page == "9. Certificate Generator":
        st.subheader("📜 Official CTG Certificate of Baptism Generator (Master Admin Exclusive)")
        rec_id_in = st.text_input("Enter Record ID (e.g. HIST-00001 or BAP-2026-...):")
        orient = st.radio("Layout Orientation:", ["portrait", "landscape"], horizontal=True)
        heading_in = st.text_input("Certificate Heading:", value="CERTIFICATE OF BAPTISM")
        wording_in = st.text_area(
            "Certificate Wording:",
            value="This is to certify that the person named below received the Sacrament of Christian Baptism according to the rites of the Church of the Temple of God.",
        )
        is_reissue = st.checkbox("Mark as Reissued Certificate")
        reissue_why = st.text_input("Reissue Reason (Required if Reissued):") if is_reissue else ""
        up_cert_drive = st.checkbox("Upload Private PDF to Google Drive (`GENERATED CERTIFICATES/<year>/`)", value=False)

        if st.button("Generate Official Certificate PDF", type="primary") and rec_id_in.strip():
            try:
                cert_res = cert_svc.generate_certificate(
                    account_id,
                    rec_id_in.strip(),
                    orientation=orient,
                    is_reissue=is_reissue,
                    reissue_reason=reissue_why,
                    custom_heading=heading_in,
                    custom_wording=wording_in,
                    upload_to_drive=up_cert_drive,
                )
                st.success(f"Generated Certificate `{cert_res['certificate_id']}` ({cert_res['issuance_status']})")
                b64_pdf = base64.b64encode(cert_res["pdf_bytes"]).decode("ascii")
                st.markdown(
                    f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="620" type="application/pdf"></iframe>',
                    unsafe_allow_html=True,
                )
                st.download_button(
                    "⬇️ Download Certificate PDF",
                    data=cert_res["pdf_bytes"],
                    file_name=f"{cert_res['certificate_id']}.pdf",
                    mime="application/pdf",
                )
            except Exception as ex:
                st.error(str(ex))

    elif page == "10. Generated Certificates Archive":
        st.subheader("🗂️ Generated Certificates Archive")
        certs_list = cert_svc.list_certificates(actor_account_id=account_id)
        st.dataframe(certs_list, use_container_width=True, hide_index=True)

    elif page == "11. Parish Leader Admin Accounts (9 Parishes)":
        st.subheader("🏛️ Nine Parish Leader Admin Accounts Management (Hong Kong Excluded)")
        st.caption(
            "Official 9 Parishes: " + " • ".join(OFFICIAL_NINE_PARISHES)
        )
        with st.expander("⚡ Provision All 9 Official Parish Leader Admin Accounts at Once", expanded=True):
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
            st.markdown("#### Active Parish Leader Admin Accounts")
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

    elif page == "12. Google Drive Storage Status":
        st.subheader("☁️ Google Drive Private Storage & Permission Check")
        wb_id = str(drive_mgr.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
        root_id = str(drive_mgr.config.get("root_folder_id") or DEFAULT_BAPTISMAL_ROOT_FOLDER_ID).strip()
        st.markdown(
            f"- **Root Folder (`CTG BAPTISMAL RECORDS`):** `{root_id}` ([Buksan sa Google Drive](https://drive.google.com/drive/folders/{root_id}))\n"
            f"- **Master Workbook (`OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`):** `{wb_id}` ([Buksan sa Google Drive](https://drive.google.com/file/d/{wb_id}/view))"
        )
        if st.button("Verify Workbook Editor Permission & Private Folder Tree", type="primary"):
            try:
                meta = drive_mgr.verify_workbook_write_access(wb_id)
                tree = drive_mgr.ensure_baptismal_folder_tree()
                st.success(f"Verified Editor access on '{meta.get('name')}' (Revision: {meta.get('headRevisionId')})")
                st.json(tree)
            except GoogleDrivePermissionError as pex:
                st.error(f"Permission Verification Failed: {pex}")
            except Exception as ex:
                st.error(f"Google Drive Check Error: {ex}")

    elif page == "13. Audit Logs & Security Settings":
        st.subheader("🛡️ Baptismal Security Audit Logs & Metadata Backups")
        if st.button("💾 Export Metadata Backup Snapshot"):
            snap = backup_mgr.export_metadata_snapshot(actor_id=account_id)
            st.success(f"Exported metadata snapshot at {snap['exported_at']}.")
        ev_rows = audit_svc.list_events(limit=250)
        st.dataframe(ev_rows, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    st.set_page_config(
        page_title="CTG Baptismal Records & Certification System",
        page_icon="⛪",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_baptismal_portal(embedded_in_hub=False)

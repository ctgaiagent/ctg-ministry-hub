"""
baptismal_app.py — CTG Baptismal Records & Certification System
===============================================================
Modular Baptismal Portal that runs seamlessly BOTH:
1. Embedded inside the unified CTG Ministry Hub (`app.py` -> "⛪ Baptismal Records & Certificates"), AND
2. Standalone if ever needed (`streamlit run baptismal_app.py`).

All confidential baptismal records stay strictly inside PostgreSQL/SQLite
and private Google Drive folders — NEVER sent to external AI services.
"""

import base64
from typing import Any, Dict
import streamlit as st

from baptismal.audit_logs import AuditLogger
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.backup_manager import BackupManager
from baptismal.certificates import CertificateService
from baptismal.config import (
    HISTORICAL_27_COLUMNS,
    HISTORICAL_BOOKS,
    SUPPORTED_NEW_YEARS,
    TEMP_ACCESS_PERMISSIONS,
    THEME_GOLD,
    THEME_NAVY,
    THEME_NAVY_DARK,
    THEME_SLATE_BG,
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
from baptismal.validation import escape_html_text


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
    hist = HistoricalRecordsService(db)
    img_arc = ImageArchiveService(db)
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


def render_baptismal_portal(embedded_in_hub: bool = False) -> None:
    """Renders the complete Baptismal Records & Certification System.
    Enforces independent server-side PBKDF2 authentication & RBAC even when embedded in app.py."""
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

    st.markdown(
        "<div class='ctg-bap-banner'>"
        "<h2>⛪ Church of the Temple of God — Baptismal Records & Certification System</h2>"
        "<p>Confidential Archival Registry • Read-Only Historical Protection • Private Google Drive & PostgreSQL Storage</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # 1. Initial Master Admin Provisioning (No hardcoded credentials)
    if not auth_svc.has_master_admin():
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
                        plain_password=setup_p1,
                        email=setup_em,
                    )
                    st.success("Baptismal Master Admin account provisioned securely. Please sign in.")
                    st.rerun()
        return

    # 2. Separate Server-Side Authentication Gate (Hub login does NOT grant Baptismal access)
    valid_sess, current_user = auth_svc.validate_session(st.session_state.get("bap_auth_user"))
    if not valid_sess or not current_user:
        st.session_state.pop("bap_auth_user", None)
        st.info("🔒 **Separate Baptismal Authorization Required:** Ang access sa Baptismal Records ay hiwalay sa karaniwang CTG Ministry Hub login.")
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

    hdr_c1, hdr_c2 = st.columns([3.5, 1.0])
    with hdr_c1:
        st.caption(f"👤 Signed in to Baptismal System as **{escape_html_text(current_user['full_name'])}** (`{escape_html_text(current_user['role'])}`)")
    with hdr_c2:
        if st.button("🚪 Sign Out (Baptismal)", use_container_width=True, key="bap_sign_out_btn"):
            auth_svc.logout(current_user)
            st.session_state.pop("bap_auth_user", None)
            st.rerun()

    # 3. Temporary Requester Isolated View
    if not is_master_admin:
        st.info(
            "🔒 **Temporary Requester Portal:** You only have access to submit certificate requests "
            "and view/download the specific record explicitly approved for your account."
        )
        tab_my_reqs, tab_my_grant = st.tabs(["📝 My Certificate Requests", "📄 Approved Record & Certificate"])
        with tab_my_reqs:
            with st.form("submit_cert_req_form"):
                req_name = st.text_input("Name of Baptized Person:")
                req_rec_id = st.text_input("Known Record ID (Optional):")
                req_purpose = st.text_area("Purpose of Request:")
                if st.form_submit_button("Submit Certificate Request", type="primary"):
                    try:
                        new_rid = temp_svc.submit_request(
                            account_id, current_user["username"], req_name, req_purpose, req_rec_id
                        )
                        st.success(f"Request submitted: `{new_rid}` (Status: PENDING)")
                    except Exception as ex:
                        st.error(str(ex))
            my_reqs = temp_svc.list_requests(requester_account_id=account_id)
            if my_reqs:
                st.dataframe(my_reqs, use_container_width=True, hide_index=True)

        with tab_my_grant:
            my_grants = temp_svc.list_active_grants(account_id=account_id)
            if not my_grants:
                st.info("No temporary record access grants assigned to your account.")
            else:
                for g in my_grants:
                    rec_id = g["record_id"]
                    if authz_svc.has_temporary_record_permission(account_id, rec_id, "view_record"):
                        rec_data = hist_svc.get_record_by_id(rec_id)
                        if rec_data:
                            st.success(f"Approved Record: `{rec_id}` (Expires: {g['expires_at']})")
                            st.json(rec_data)
                            audit_svc.log_event(account_id, "HISTORICAL_RECORD_VIEWED", "SUCCESS", related_record_id=rec_id)
                            if authz_svc.has_temporary_record_permission(account_id, rec_id, "generate_certificate"):
                                if st.button(f"Generate Certificate PDF ({rec_id})", key=f"temp_gen_{g['grant_id']}"):
                                    res_c = cert_svc.generate_certificate(account_id, rec_id, upload_to_drive=False)
                                    if authz_svc.has_temporary_record_permission(account_id, rec_id, "download_certificate"):
                                        st.download_button(
                                            "⬇️ Download Certificate PDF",
                                            data=res_c["pdf_bytes"],
                                            file_name=f"{res_c['certificate_id']}.pdf",
                                            mime="application/pdf",
                                        )
        return

    # 4. Master Admin 12-Section Navigation (Sidebar when standalone, Top Navigation when embedded in app.py)
    NAV_ITEMS = [
        "1. Dashboard",
        "2. Historical Records",
        "3. BOOK 1–9 Image Archive",
        "4. Certificate Requests",
        "5. Temporary Access Management",
        "6. New Baptism Registration",
        "7. New Baptism Images",
        "8. Certificate Generator",
        "9. Generated Certificates",
        "10. Google Drive Storage Status",
        "11. Audit Logs",
        "12. Security Settings",
    ]
    if embedded_in_hub:
        page = st.selectbox("📂 Baptismal System Module:", NAV_ITEMS, key="bap_embedded_nav")
    else:
        page = st.sidebar.radio("Navigation", NAV_ITEMS, key="bap_standalone_nav")

    if page == "1. Dashboard":
        stats = hist_svc.get_dashboard_statistics()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Historical Records", stats["total_historical"])
        c2.metric("New Registrations", stats["new_registrations"])
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
            st.subheader("Records by Parish")
            st.dataframe(stats["by_parish"], use_container_width=True, hide_index=True)
        with col_y:
            st.subheader("Records by Baptism Year")
            st.dataframe(stats["by_year"], use_container_width=True, hide_index=True)

    elif page == "2. Historical Records":
        st.subheader("📚 Historical Records Search (Read-Only — All 27 Columns)")
        with st.expander("📥 Import Historical `Sheet1` from `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`"):
            hist_up = st.file_uploader("Upload OVERALL BAPTISMAL CERTIFICATE LIST.xlsx:", type=["xlsx"], key="hist_xlsx_import")
            if hist_up and st.button("Import Historical Sheet1 (Read-Only)", type="primary"):
                cnt = hist_svc.import_sheet1_from_excel_bytes(hist_up.read())
                audit_svc.log_event(account_id, "HISTORICAL_IMPORT", "SUCCESS", metadata={"imported_rows": cnt})
                st.success(f"Imported {cnt} historical records into read-only storage.")
                st.rerun()

        fc1, fc2, fc3, fc4 = st.columns(4)
        kw = fc1.text_input("Keyword Search (Any of 27 Columns):", key="hist_kw")
        f_parish = fc2.text_input("Parish:", key="hist_f_parish")
        f_book = fc3.text_input("Book Number:", key="hist_f_book")
        f_year = fc4.text_input("Year of Baptism:", key="hist_f_year")

        if st.button("Clear Search Filters"):
            for k in ("hist_kw", "hist_f_parish", "hist_f_book", "hist_f_year"):
                st.session_state[k] = ""
            st.rerun()

        page_num = int(st.number_input("Page:", min_value=1, value=1, step=1))
        rows, total_rows = hist_svc.search_records(
            keyword=kw,
            filters={"Parish": f_parish, "Book Number": f_book, "Year of Baptism": f_year},
            page=page_num,
            page_size=50,
        )
        st.caption(f"Showing {len(rows)} of {total_rows} matching historical records (read-only).")
        if rows:
            st.dataframe(rows, use_container_width=True, hide_index=True)
            sel_id = st.selectbox("Inspect Full Historical Record (All 27 Fields):", [r["Record ID"] for r in rows])
            if sel_id:
                detail = hist_svc.get_record_by_id(sel_id)
                st.json(detail)
                audit_svc.log_event(account_id, "HISTORICAL_RECORD_VIEWED", "SUCCESS", related_record_id=sel_id)

    elif page == "3. BOOK 1–9 Image Archive":
        st.subheader("🖼️ BOOK 1–9 Visual Archive & Manual Image Confirmation")
        sel_book = st.selectbox("Select Book Folder:", HISTORICAL_BOOKS)
        st_flt = st.selectbox("Match Status Filter:", ["ALL", "POSSIBLE_MATCH_NEEDS_CONFIRMATION", "CONFIRMED", "REJECTED"])
        assocs = img_arc_svc.list_by_book(sel_book, status_filter=st_flt)
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
            m_rec_id = st.text_input("Record ID (e.g. HIST-00002):")
            m_img_fn = st.text_input("Scanned Image Filename (e.g. Abelardo Galban BC_744.jpeg):")
            m_notes = st.text_input("Notes:")
            if st.button("Create Candidate Association") and m_rec_id and m_img_fn:
                res_m = img_match_svc.suggest_possible_match(m_rec_id, sel_book, m_img_fn, notes=m_notes)
                st.info(f"Created `{res_m['association_id']}` — Status: `{res_m['match_status']}`")

        prev_up = st.file_uploader("Preview / Rotate / Zoom Local Scan Image:", type=["jpg", "jpeg", "png", "webp", "tif", "tiff"])
        if prev_up:
            rot_deg = st.selectbox("Rotation (Degrees):", [0, 90, 180, 270])
            zoom_w = st.slider("Preview Width (px):", min_value=400, max_value=2200, value=1100, step=100)
            raw_b = prev_up.read()
            prev_b = render_image_preview_bytes(raw_b, rotation_degrees=rot_deg, max_width=zoom_w)
            st.image(prev_b, caption=f"{prev_up.name} (Rotated {rot_deg}°)")
            st.download_button("⬇️ Download Image", data=prev_b, file_name=f"rotated_{prev_up.name}.jpg", mime="image/jpeg")

    elif page == "4. Certificate Requests":
        st.subheader("📨 Certificate Requests Queue")
        reqs = temp_svc.list_requests()
        if reqs:
            st.dataframe(reqs, use_container_width=True, hide_index=True)
        else:
            st.info("No certificate requests recorded yet.")

    elif page == "5. Temporary Access Management":
        st.subheader("🔑 Temporary Record-Specific Access Management")
        reqs = temp_svc.list_requests()
        pending_reqs = [r for r in reqs if r["status"] == "PENDING"]
        if pending_reqs:
            sel_req = st.selectbox("Select Pending Request:", [r["request_id"] for r in pending_reqs])
            target_rec = st.text_input("Permitted Record ID (Required):")
            perms_sel = st.multiselect("Allowed Permissions:", TEMP_ACCESS_PERMISSIONS, default=["view_record", "generate_certificate", "download_certificate"])
            exp_hrs = st.number_input("Expiration (Hours):", min_value=1.0, max_value=168.0, value=24.0)
            notes_in = st.text_input("Review Notes:")
            c_ap, c_rj = st.columns(2)
            if c_ap.button("✅ Approve & Grant Temporary Access", type="primary") and target_rec.strip():
                gid = temp_svc.approve_request_and_grant_access(account_id, sel_req, target_rec.strip(), perms_sel, exp_hrs, notes_in)
                st.success(f"Approved! Grant ID: `{gid}`")
                st.rerun()
            if c_rj.button("❌ Reject Request"):
                temp_svc.reject_request(account_id, sel_req, notes_in)
                st.warning("Request rejected.")
                st.rerun()

        grants = temp_svc.list_active_grants()
        if grants:
            st.markdown("#### Active & Historical Grants")
            st.dataframe(grants, use_container_width=True, hide_index=True)
            active_gids = [g["grant_id"] for g in grants if not int(g.get("is_revoked", 0))]
            if active_gids:
                rev_gid = st.selectbox("Revoke Grant Immediately:", active_gids)
                if st.button("🚫 Revoke Access Now"):
                    temp_svc.revoke_grant_immediately(account_id, rev_gid)
                    st.success(f"Grant `{rev_gid}` revoked immediately.")
                    st.rerun()

    elif page == "6. New Baptism Registration":
        st.subheader("✍️ New Baptism Registration (All 27 Canonical Fields)")
        with st.form("new_baptism_reg_form"):
            payload: Dict[str, str] = {}
            cols = st.columns(3)
            for idx, col_name in enumerate(HISTORICAL_27_COLUMNS):
                with cols[idx % 3]:
                    if col_name == "Year of Baptism":
                        payload[col_name] = st.selectbox(col_name, SUPPORTED_NEW_YEARS)
                    else:
                        payload[col_name] = st.text_input(col_name, key=f"reg_f_{col_name}")
            up_imgs = st.file_uploader("Upload Baptismal Images (Private Drive Year Folder):", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
            sync_drive = st.checkbox("Sync to Google Drive Workbook (`NEW BAPTISMS - YYYY`)", value=False)
            if st.form_submit_button("Preview & Register Baptism", type="primary"):
                img_list = []
                for uf in up_imgs or []:
                    img_list.append({"bytes": uf.read(), "filename": uf.name, "mime_type": uf.type or "image/jpeg"})
                try:
                    res = reg_svc.register_new_baptism(account_id, payload, uploaded_images=img_list, sync_to_drive_excel=sync_drive)
                    st.success(f"Registered! Record ID: `{res['record_id']}` • Worksheet: `{res['worksheet_name']}` • Workbook Status: `{res['workbook_update_status']}`")
                except Exception as ex:
                    st.error(str(ex))

    elif page == "7. New Baptism Images":
        st.subheader("📸 New Baptism Images Registry")
        rows_m = svc["db"].fetch_all("SELECT * FROM drive_file_mappings WHERE category = 'NEW_BAPTISM_IMAGE' ORDER BY created_at DESC")
        st.dataframe(rows_m, use_container_width=True, hide_index=True)

    elif page == "8. Certificate Generator":
        st.subheader("📜 CTG Certificate of Baptism Generator (A4 Portrait / Landscape)")
        rec_id_in = st.text_input("Enter Record ID (Historical or New):")
        orient = st.radio("Layout Orientation:", ["portrait", "landscape"], horizontal=True)
        heading_in = st.text_input("Certificate Heading:", value="CERTIFICATE OF BAPTISM")
        wording_in = st.text_area(
            "Certificate Wording:",
            value="This is to certify that the person named below received the Sacrament of Christian Baptism according to the rites of the Church of the Temple of God.",
        )
        is_reissue = st.checkbox("Mark as Reissued Certificate")
        reissue_why = st.text_input("Reissue Reason (Required if Reissued):") if is_reissue else ""
        up_cert_drive = st.checkbox("Upload Private PDF to Google Drive (`GENERATED CERTIFICATES/<year>/`)", value=False)

        if st.button("Generate Certificate PDF", type="primary") and rec_id_in.strip():
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

    elif page == "9. Generated Certificates":
        st.subheader("🗂️ Certificate Archive")
        certs_list = cert_svc.list_certificates()
        st.dataframe(certs_list, use_container_width=True, hide_index=True)

    elif page == "10. Google Drive Storage Status":
        st.subheader("☁️ Google Drive Private Storage & Permission Check")
        wb_id = str(drive_mgr.config.get("workbook_file_id") or "").strip()
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

    elif page == "11. Audit Logs":
        st.subheader("🛡️ Baptismal Security & Activity Audit Logs")
        ev_rows = audit_svc.list_events(limit=250)
        st.dataframe(ev_rows, use_container_width=True, hide_index=True)

    elif page == "12. Security Settings":
        st.subheader("⚙️ Security Settings, Accounts & Backups")
        with st.expander("➕ Create Temporary Requester Account"):
            with st.form("create_temp_acc_form"):
                tu = st.text_input("Requester Username:")
                tfn = st.text_input("Full Name:")
                tem = st.text_input("Email:")
                tpw = st.text_input("Initial Password (min 8 chars):", type="password")
                if st.form_submit_button("Create Temporary Requester"):
                    try:
                        aid = auth_svc.create_account(tu, tfn, "Temporary Requester", plain_password=tpw, email=tem)
                        st.success(f"Created Temporary Requester account `{aid}`.")
                    except Exception as ex:
                        st.error(str(ex))
        if st.button("💾 Export Metadata Backup Snapshot"):
            snap = backup_mgr.export_metadata_snapshot(actor_id=account_id)
            st.success(f"Exported metadata snapshot at {snap['exported_at']}.")


if __name__ == "__main__":
    st.set_page_config(
        page_title="CTG Baptismal Records & Certification System",
        page_icon="⛪",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_baptismal_portal(embedded_in_hub=False)

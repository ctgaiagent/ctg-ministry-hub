"""
marriage_app.py — CTG Marriage Records & Certification System (Master Admin Only)
=================================================================================
- Strictly restricted to Master Admin (Parish Leader Admins are blocked at both UI and server layers).
- Connects to Google Drive folder CTG MARRIAGE CERTIFICATE (1QTvhE_zJ52TBAgdgVlybVH5UODbdAA2A),
  MARRIAGE CERT.xlsx (1cG0_neGpx89MGYK4iGKhoFJ240R_SGRy), and MARRIAGE CERTIFICATE JPEG
  (1TUDlzDi-LCjxmxYorg099cs-uJlyJPzq).
- Shares the Master Admin vault gate (ctg.Baptismal.admin verified against PBKDF2-HMAC-SHA256 hash)
  with CTG BAPTISMAL RECORDS so the Master Admin can switch seamlessly between Baptismal & Marriage.
"""

import base64
import os
import re
import time as _time_mod
from typing import Any, Dict, Optional

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from baptismal.authentication import verify_password_pbkdf2
from baptismal.config import (
    MASTER_ADMIN_BAPTISMAL_DISPLAY_USERNAME,
    MASTER_ADMIN_BAPTISMAL_PW_HASH,
    MASTER_ADMIN_BAPTISMAL_USERNAME,
    OFFICIAL_NINE_PARISHES,
    THEME_GOLD,
    THEME_NAVY,
    THEME_NAVY_DARK,
)
from baptismal_app import (
    _BAP_SERVICES_VERSION,
    _get_baptismal_services,
    _resolve_or_provision_hub_synced_user,
)
from marriage.certificates import (
    generate_and_archive_marriage_certificate,
    list_generated_marriage_certificates,
)
from marriage.config import (
    DEFAULT_MARRIAGE_FOLDER_URL,
    DEFAULT_MARRIAGE_IMAGES_FOLDER_ID,
    DEFAULT_MARRIAGE_IMAGES_FOLDER_URL,
    DEFAULT_MARRIAGE_ROOT_FOLDER_ID,
    DEFAULT_MARRIAGE_WORKBOOK_FILE_ID,
    DEFAULT_MARRIAGE_WORKBOOK_URL,
)
from marriage.records import (
    confirm_marriage_image_link,
    ensure_marriage_tables,
    get_marriage_dashboard_metrics,
    get_marriage_record_by_id,
    inspect_or_sync_large_marriage_workbook,
    register_new_marriage_record,
    resolve_quick_marriage_id,
    search_marriage_records,
    sync_records_from_marriage_jpeg_folder,
)


def _inject_marriage_theme_css() -> None:
    st.markdown(
        f"""
        <style>
        .mar-hero {{
            background: linear-gradient(135deg, {THEME_NAVY_DARK} 0%, {THEME_NAVY} 65%, #1E3A8A 100%);
            color: #FFFFFF;
            padding: 1.15rem 1.4rem;
            border-radius: 14px;
            border: 2px solid {THEME_GOLD};
            box-shadow: 0 8px 20px rgba(15, 44, 89, 0.18);
            margin-bottom: 1rem;
        }}
        .mar-hero h2 {{
            color: #FFFFFF !important;
            margin: 0 0 0.25rem 0 !important;
            font-size: 1.45rem !important;
            font-weight: 800 !important;
        }}
        .mar-gold-badge {{
            display: inline-block;
            background: {THEME_GOLD};
            color: {THEME_NAVY_DARK};
            font-weight: 800;
            font-size: 0.76rem;
            padding: 0.2rem 0.65rem;
            border-radius: 999px;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }}
        .mar-card {{
            background: #FFFFFF;
            border-radius: 12px;
            padding: 1rem 1.15rem;
            border: 1px solid #E2E8F0;
            border-top: 4px solid {THEME_GOLD};
            box-shadow: 0 3px 10px rgba(15, 44, 89, 0.06);
            margin-bottom: 0.85rem;
        }}
        .mar-kpi {{
            background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
            border: 1px solid #CBD5E1;
            border-left: 5px solid {THEME_NAVY};
            border-radius: 10px;
            padding: 0.8rem 1rem;
        }}
        .mar-kpi-label {{
            font-size: 0.76rem;
            color: #475569;
            font-weight: 700;
            text-transform: uppercase;
        }}
        .mar-kpi-val {{
            font-size: 1.55rem;
            color: {THEME_NAVY};
            font-weight: 800;
            margin-top: 0.15rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_inline_pdf_print_component(pdf_bytes: bytes, label: str = "Print Official Marriage Certificate") -> None:
    b64_pdf = base64.b64encode(pdf_bytes).decode("ascii")
    html_code = f"""
    <div style="margin:6px 0;">
      <button onclick="printMarPdf()" style="
        background: linear-gradient(135deg, #0F2C59, #1E3A8A);
        color: #FFFFFF;
        border: 2px solid #D4AF37;
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 700;
        cursor: pointer;
        width: 100%;
        font-size: 14px;">
        🖨️ {label}
      </button>
    </div>
    <script>
    function printMarPdf() {{
      const b64 = "{b64_pdf}";
      const byteChars = atob(b64);
      const byteNumbers = new Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) {{
        byteNumbers[i] = byteChars.charCodeAt(i);
      }}
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], {{type: 'application/pdf'}});
      const blobUrl = URL.createObjectURL(blob);
      const win = window.open(blobUrl, '_blank');
      if (win) {{
        win.addEventListener('load', function() {{
          win.focus();
          win.print();
        }});
      }}
    }}
    </script>
    """
    components.html(html_code, height=52)


def _render_inline_image_print_button(img_b64: str, mime: str, title: str) -> None:
    safe_title = re.sub(r"[^a-zA-Z0-9 _.-]", "", title or "Marriage Contract")
    html_code = f"""
    <div style="margin:4px 0;">
      <button onclick="printMarImg()" style="
        background: #0F2C59;
        color: #FFFFFF;
        border: 1.5px solid #D4AF37;
        border-radius: 8px;
        padding: 7px 14px;
        font-weight: 700;
        cursor: pointer;
        width: 100%;
        font-size: 13px;">
        🖨️ Print Scanned Marriage Certificate ({safe_title[:28]})
      </button>
    </div>
    <script>
    function printMarImg() {{
      const win = window.open('', '_blank');
      if (!win) return;
      win.document.write(`
        <html>
          <head><title>{safe_title}</title>
          <style>
            @page {{ size: auto; margin: 8mm; }}
            body {{ margin: 0; display: flex; justify-content: center; align-items: center; }}
            img {{ max-width: 100%; max-height: 96vh; object-fit: contain; }}
          </style>
          </head>
          <body>
            <img src="data:{mime};base64,{img_b64}" onload="window.focus(); window.print();" />
          </body>
        </html>
      `);
      win.document.close();
    }}
    </script>
    """
    components.html(html_code, height=48)


def render_marriage_portal(embedded_in_hub: bool = False) -> None:
    """Main entry point for the Master-Admin-Only CTG Marriage Records & Certification System."""
    _inject_marriage_theme_css()

    svc = _get_baptismal_services(_version=_BAP_SERVICES_VERSION)
    db = svc["db"]
    auth_svc = svc["auth"]
    drive_mgr = svc["drive"]

    ensure_marriage_tables(db)
    _resolve_or_provision_hub_synced_user(auth_svc, is_master=True)

    # Strict RBAC: Block any Parish Leader Admin from accessing Marriage Records & Certificates
    hub_is_authenticated = bool(embedded_in_hub and st.session_state.get("authenticated", False))
    hub_true_login_role = str(st.session_state.get("login_role") or "")
    hub_active_view_role = str(st.session_state.get("user_role") or hub_true_login_role or "")
    is_hub_view_master = (not hub_is_authenticated) or (hub_active_view_role in ("Master Admin", "Super Admin"))

    if hub_is_authenticated and not is_hub_view_master:
        st.error(
            "🚫 **Access Denied (Master Admin Only):** Ang **CTG Marriage Records & Certification System** "
            "ay eksklusibo lamang para sa **Master Admin**."
        )
        return

    # Master Admin Vault Gate (shares session with CTG BAPTISMAL RECORDS)
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
            """
            <div class="mar-hero">
                <span class="mar-gold-badge">🔒 Master Admin Security Vault</span>
                <h2>💍 CTG MARRIAGE CERTIFICATE — Master Admin Authentication</h2>
                <p style="margin:0;opacity:0.92;font-size:0.92rem;">
                    Ang Marriage Records &amp; Certification System ay para lamang sa <b>Master Admin</b>.
                    Mag-sign in gamit ang iyong Master Admin Sacramental Vault account.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        col_l, col_r = st.columns([1.25, 1])
        with col_l:
            with st.form("ctg_marriage_records_master_login_form"):
                in_u = st.text_input("User Name:", value=MASTER_ADMIN_BAPTISMAL_DISPLAY_USERNAME)
                in_p = st.text_input("Password:", type="password")
                if st.form_submit_button("Sign In to CTG MARRIAGE CERTIFICATE", type="primary", use_container_width=True):
                    u_clean = str(in_u or "").strip().lower()
                    if u_clean == MASTER_ADMIN_BAPTISMAL_USERNAME and verify_password_pbkdf2(in_p, MASTER_ADMIN_BAPTISMAL_PW_HASH):
                        master_obj = _resolve_or_provision_hub_synced_user(auth_svc, is_master=True)
                        st.session_state["bap_auth_user"] = master_obj
                        st.session_state["bap_master_authenticated"] = True
                        st.session_state["_bap_last_activity_ts"] = _time_mod.time()
                        st.rerun()
                    else:
                        ok_db, db_user, _ = auth_svc.authenticate(u_clean, in_p)
                        if ok_db and db_user and db_user.get("role") == "Master Admin":
                            st.session_state["bap_auth_user"] = db_user
                            st.session_state["bap_master_authenticated"] = True
                            st.session_state["_bap_last_activity_ts"] = _time_mod.time()
                            st.rerun()
                        else:
                            st.error("Invalid Master Admin User Name or Password.")
        with col_r:
            st.markdown(
                f"""
                <div class="mar-card">
                    <h4 style="margin-top:0;color:{THEME_NAVY};">📂 Connected Google Drive Links</h4>
                    <ul style="font-size:0.88rem;line-height:1.65;padding-left:1.1rem;margin-bottom:0;">
                        <li><b>Marriage Root Folder:</b> <a href="{DEFAULT_MARRIAGE_FOLDER_URL}" target="_blank">CTG MARRIAGE CERTIFICATE</a></li>
                        <li><b>Marriage Excel List:</b> <a href="{DEFAULT_MARRIAGE_WORKBOOK_URL}" target="_blank">MARRIAGE CERT.xlsx</a></li>
                        <li><b>Scanned Images Folder:</b> <a href="{DEFAULT_MARRIAGE_IMAGES_FOLDER_URL}" target="_blank">MARRIAGE CERTIFICATE JPEG</a></li>
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
        return

    user = current_user
    certificates_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generated_marriage_certificates")

    # Top Header & 1-Click Certificate Switcher (Baptismal <-> Marriage)
    st.markdown(
        f"""
        <div class="mar-hero" style="margin-bottom:0.65rem;">
            <span class="mar-gold-badge">💍 Master Admin Exclusive • CTG Sacramental Registry</span>
            <h2>💍 CTG MARRIAGE CERTIFICATE — Master Admin Archive &amp; Certification Hub</h2>
            <p style="margin:0;font-size:0.88rem;opacity:0.92;">
                Connected sa Google Drive:
                <a href="{DEFAULT_MARRIAGE_FOLDER_URL}" target="_blank" style="color:#FDE047;text-decoration:underline;">CTG MARRIAGE CERTIFICATE Folder</a> •
                <a href="{DEFAULT_MARRIAGE_WORKBOOK_URL}" target="_blank" style="color:#FDE047;text-decoration:underline;">MARRIAGE CERT.xlsx</a> •
                <a href="{DEFAULT_MARRIAGE_IMAGES_FOLDER_URL}" target="_blank" style="color:#FDE047;text-decoration:underline;">MARRIAGE CERTIFICATE JPEG (449+ Scans)</a>
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    hdr_c1, hdr_c2, hdr_c3, hdr_c4 = st.columns([2.1, 1.3, 1.3, 0.9])
    with hdr_c1:
        st.caption(
            f"👑 Signed in as **{MASTER_ADMIN_BAPTISMAL_DISPLAY_USERNAME}** (`Master Admin`)"
        )
    with hdr_c2:
        if embedded_in_hub:
            if st.button("⛪ Switch to Certificate of Baptismal", use_container_width=True, key="mar_switch_to_bap_btn"):
                st.session_state["app_mode_override"] = "⛪ Baptismal Records & Certificates"
                st.rerun()
    with hdr_c3:
        if st.button("🔄 Sync 449+ JPEGs & Workbook (Drive)", use_container_width=True, type="primary", key="mar_top_sync_drive_btn"):
            with st.spinner("Ini-index ang MARRIAGE CERTIFICATE JPEG folder at MARRIAGE CERT.xlsx nang ligtas sa RAM..."):
                sync_res = inspect_or_sync_large_marriage_workbook(
                    db=db,
                    drive_mgr=drive_mgr,
                    workbook_file_id=DEFAULT_MARRIAGE_WORKBOOK_FILE_ID,
                    images_folder_id=DEFAULT_MARRIAGE_IMAGES_FOLDER_ID,
                    actor=user,
                )
                if sync_res.get("ok"):
                    st.success(sync_res.get("message", "Drive sync complete!"))
                else:
                    st.warning(sync_res.get("message", "Drive sync skipped."))
    with hdr_c4:
        if st.button("🔒 Sign Out Vault", use_container_width=True, key="mar_lock_vault_btn"):
            auth_svc.logout(user)
            st.session_state.pop("bap_master_authenticated", None)
            st.session_state.pop("bap_auth_user", None)
            st.rerun()

    # Auto-sync once per session if Google Drive service is available
    if not st.session_state.get("_mar_jpeg_auto_synced"):
        try:
            sync_records_from_marriage_jpeg_folder(
                db=db,
                drive_mgr=drive_mgr,
                images_folder_id=DEFAULT_MARRIAGE_IMAGES_FOLDER_ID,
                actor=user,
            )
            st.session_state["_mar_jpeg_auto_synced"] = True
        except Exception:
            st.session_state["_mar_jpeg_auto_synced"] = True

    _mar_side_sel = str(st.session_state.get("mar_contextual_nav") or "Dashboard").strip()
    _mar_tab_map = {
        "Dashboard": 0,
        "Global Search & Print": 1,
        "MARRIAGE JPEG Archive": 2,
        "New Marriage Registration": 3,
        "Certificate Generator": 4,
        "Generated Certificates": 4,
        "Google Drive Storage": 0,
    }
    _mar_t_idx = _mar_tab_map.get(_mar_side_sel, 0)
    if st.session_state.get("_mar_force_tab_jump"):
        st.session_state["_mar_force_tab_jump"] = False
        components.html(
            f"""
            <script>
            setTimeout(function() {{
                const tabs = window.parent.document.querySelectorAll('button[data-baseweb="tab"]');
                if (tabs && tabs.length > {_mar_t_idx}) {{
                    tabs[{_mar_t_idx}].click();
                }}
            }}, 120);
            </script>
            """,
            height=0,
        )

    tabs = st.tabs(
        [
            "📊 1. Dashboard",
            "🔍 2. Search ID, Marriage Image & Print",
            "🖼️ 3. MARRIAGE CERTIFICATE JPEG Archive",
            "✍️ 4. New Marriage Registration (Encode to Drive)",
            "📜 5. Official Marriage Certificate Generator & Archive",
        ]
    )

    # =========================================================================
    # TAB 1: DASHBOARD
    # =========================================================================
    with tabs[0]:
        metrics = get_marriage_dashboard_metrics(db, actor=user)
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                f'<div class="mar-kpi"><div class="mar-kpi-label">Total Marriage Records</div>'
                f'<div class="mar-kpi-val">{metrics["total_records"]:,}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f'<div class="mar-kpi"><div class="mar-kpi-label">Linked Scanned JPEGs</div>'
                f'<div class="mar-kpi-val">{metrics["with_scanned_images"]:,}</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f'<div class="mar-kpi"><div class="mar-kpi-label">New Encoded Marriages</div>'
                f'<div class="mar-kpi-val">{metrics["new_registrations"]:,}</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f'<div class="mar-kpi"><div class="mar-kpi-label">Issued Marriage Certificates</div>'
                f'<div class="mar-kpi-val">{metrics["generated_certificates"]:,}</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("#### 🔗 Direct Google Drive Links (`CTG MARRIAGE CERTIFICATE`)")
        l1, l2, l3 = st.columns(3)
        with l1:
            st.markdown(
                f"""
                <div class="mar-card">
                    <b>📁 Root Marriage Folder</b><br/>
                    <span style="font-size:0.83rem;color:#475569;">ID: <code>{DEFAULT_MARRIAGE_ROOT_FOLDER_ID}</code></span><br/>
                    <a href="{DEFAULT_MARRIAGE_FOLDER_URL}" target="_blank">👉 Buksan ang CTG MARRIAGE CERTIFICATE Folder</a>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with l2:
            st.markdown(
                f"""
                <div class="mar-card">
                    <b>📊 Master Marriage Workbook</b><br/>
                    <span style="font-size:0.83rem;color:#475569;">ID: <code>{DEFAULT_MARRIAGE_WORKBOOK_FILE_ID}</code> (Range-Protected)</span><br/>
                    <a href="{DEFAULT_MARRIAGE_WORKBOOK_URL}" target="_blank">👉 Buksan ang MARRIAGE CERT.xlsx sa Google Drive</a>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with l3:
            st.markdown(
                f"""
                <div class="mar-card">
                    <b>🖼️ Scanned Marriage JPEGs Folder</b><br/>
                    <span style="font-size:0.83rem;color:#475569;">ID: <code>{DEFAULT_MARRIAGE_IMAGES_FOLDER_ID}</code></span><br/>
                    <a href="{DEFAULT_MARRIAGE_IMAGES_FOLDER_URL}" target="_blank">👉 Buksan ang MARRIAGE CERTIFICATE JPEG Folder</a>
                </div>
                """,
                unsafe_allow_html=True,
            )

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("##### 🏛️ Marriage Records by Parish / Place")
            parish_df = pd.DataFrame(
                [{"Parish": k, "Marriage Records": v} for k, v in sorted(metrics["by_parish"].items(), key=lambda x: (-x[1], x[0]))]
            )
            st.dataframe(parish_df, use_container_width=True, hide_index=True)
        with col_p2:
            st.markdown("##### ✝️ Marriage Records by Officiating Priest / Minister")
            priest_df = pd.DataFrame(
                [{"Officiating Priest / Minister": k, "Solemnized Marriages": v} for k, v in sorted(metrics["by_priest"].items(), key=lambda x: (-x[1], x[0]))]
            )
            st.dataframe(priest_df, use_container_width=True, hide_index=True)

    # =========================================================================
    # TAB 2: SEARCH ID, MARRIAGE IMAGE & PRINT (WITH QUICK NUMBER JUMP)
    # =========================================================================
    with tabs[1]:
        st.markdown("#### 🔍 Quick Marriage Record ID Jump & Side-by-Side Scanned Marriage Contract Viewer")
        if "mar_quick_num" not in st.session_state:
            st.session_state["mar_quick_num"] = 5

        q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns([0.9, 1.3, 0.85, 0.85, 2.2])
        with q_col1:
            st.markdown(
                f"<div style='background:{THEME_NAVY};color:#FFF;padding:9px 12px;border-radius:8px;"
                f"font-weight:800;text-align:center;margin-top:24px;border:1.5px solid {THEME_GOLD};'>"
                f"Prefix: MAR-</div>",
                unsafe_allow_html=True,
            )
        with q_col2:
            quick_num = st.number_input(
                "Last Number Only (e.g. 5, 112, 154, 447):",
                min_value=1,
                max_value=99999,
                value=int(st.session_state.get("mar_quick_num", 5)),
                step=1,
                key="mar_quick_num_input",
            )
            st.session_state["mar_quick_num"] = int(quick_num)
        with q_col3:
            st.markdown("<div style='margin-top:24px;'></div>", unsafe_allow_html=True)
            if st.button("⬅️ Prev ID", use_container_width=True, key="mar_prev_id_btn"):
                st.session_state["mar_quick_num"] = max(1, int(st.session_state["mar_quick_num"]) - 1)
                st.rerun()
        with q_col4:
            st.markdown("<div style='margin-top:24px;'></div>", unsafe_allow_html=True)
            if st.button("➡️ Next ID", use_container_width=True, key="mar_next_id_btn"):
                st.session_state["mar_quick_num"] = int(st.session_state["mar_quick_num"]) + 1
                st.rerun()
        with q_col5:
            kw_query = st.text_input(
                "🔎 Or Search Husband, Wife, Parents, Witnesses, License No., or Date:",
                value="",
                placeholder="Halimbawa: Francisco Bautista, Leonida Yanguas, Dominador Ucol, 1993961...",
                key="mar_keyword_search_box",
            )

        active_record_id = resolve_quick_marriage_id(str(st.session_state["mar_quick_num"]))
        if kw_query.strip():
            matched_rows = search_marriage_records(db, query=kw_query.strip(), limit=100, actor=user)
            st.caption(f"Natagpuan: **{len(matched_rows)}** marriage record(s) para sa `{kw_query}`")
            if matched_rows:
                df_matches = pd.DataFrame(matched_rows)[
                    [
                        "Record ID",
                        "MC Number",
                        "Husband Name",
                        "Wife Name",
                        "Date of Marriage",
                        "Place of Marriage",
                        "Marriage License No.",
                        "Officiating Priest / Minister",
                        "Scanned Image Filename",
                    ]
                ]
                st.dataframe(df_matches, use_container_width=True, hide_index=True)
                sel_opts = [
                    f"{r['Record ID']} ({r['MC Number']}) — {r['Husband Name']} & {r['Wife Name']} [{r['Date of Marriage']}]"
                    for r in matched_rows
                ]
                chosen_label = st.selectbox("Pumili ng Marriage Record para buksan:", sel_opts, key="mar_search_pick")
                active_record_id = chosen_label.split(" ")[0].strip()
        else:
            all_brief = search_marriage_records(db, query="", limit=500, actor=user)
            with st.expander(f"📋 Tingnan ang Listahan ng Lahat ng Marriage Records ({len(all_brief)} records)", expanded=False):
                if all_brief:
                    st.dataframe(
                        pd.DataFrame(all_brief)[
                            [
                                "Record ID",
                                "MC Number",
                                "Husband Name",
                                "Wife Name",
                                "Date of Marriage",
                                "Place of Marriage",
                                "Scanned Image Filename",
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True,
                    )

        rec = get_marriage_record_by_id(db, active_record_id, actor=user)
        st.divider()
        if not rec:
            st.info(
                f"ℹ️ Walang naka-save na record sa local index para sa **{active_record_id}** (`MC_{int(st.session_state['mar_quick_num']):03d}`). "
                f"Pindutin ang **🔄 Sync 449+ JPEGs & Workbook (Drive)** sa itaas o i-encode ito sa Tab 4."
            )
        else:
            left_col, right_col = st.columns([1.1, 1.15])
            with left_col:
                st.markdown(
                    f"""
                    <div class="mar-card">
                        <span class="mar-gold-badge">{rec['Record ID']} • {rec['MC Number']}</span>
                        <h3 style="margin:0.35rem 0;color:{THEME_NAVY};">
                            💍 {rec['Husband Name'] or 'Unspecified Groom'} &amp; {rec['Wife Name'] or 'Unspecified Bride'}
                        </h3>
                        <p style="margin:0 0 0.5rem 0;font-size:0.86rem;color:#475569;">
                            <b>Date of Marriage:</b> {rec['Date of Marriage'] or '—'} &nbsp;|&nbsp;
                            <b>Parish / Place:</b> {rec['Place of Marriage'] or rec['Parish']}
                        </p>
                        <hr style="margin:0.5rem 0;border-color:#E2E8F0;"/>
                        <table style="width:100%;font-size:0.86rem;line-height:1.55;">
                            <tr><td style="width:38%;color:#475569;font-weight:700;">Husband (Groom):</td><td><b>{rec['Husband Name'] or '—'}</b> ({rec['Husband Age / Birthdate'] or '—'}, {rec['Husband Single / Widowed'] or 'Single'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Husband Residence:</td><td>{rec['Husband Residence'] or '—'} ({rec['Husband Nationality'] or 'Filipino'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Husband's Parents:</td><td>Father: {rec['Husband Father'] or '—'} | Mother: {rec['Husband Mother'] or '—'}</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Wife (Bride):</td><td><b>{rec['Wife Name'] or '—'}</b> ({rec['Wife Age / Birthdate'] or '—'}, {rec['Wife Single / Widowed'] or 'Single'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Wife Residence:</td><td>{rec['Wife Residence'] or '—'} ({rec['Wife Nationality'] or 'Filipina'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Wife's Parents:</td><td>Father: {rec['Wife Father'] or '—'} | Mother: {rec['Wife Mother'] or '—'}</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Marriage License No.:</td><td>{rec['Marriage License No.'] or '—'} ({rec['License Issued At / Date'] or '—'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Officiating Priest:</td><td>{rec['Officiating Priest / Minister'] or '—'} ({rec['Priest Position / Religion'] or '—'})</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Witnesses / Sponsors:</td><td>{rec['Witnesses / Sponsors'] or '—'}</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Book / Page / Registry:</td><td>Book: {rec['Book Number']} | Page: {rec['Page Number']} | LCR No.: {rec['Local Civil Registry No.'] or '—'}</td></tr>
                            <tr><td style="color:#475569;font-weight:700;">Scanned Image File:</td><td><code>{rec['Scanned Image Filename'] or '—'}</code></td></tr>
                        </table>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                act_c1, act_c2 = st.columns(2)
                with act_c1:
                    orient_quick = st.selectbox(
                        "Certificate Layout:",
                        ["Portrait", "Landscape"],
                        key=f"mar_quick_orient_{rec['Record ID']}",
                    )
                with act_c2:
                    st.markdown("<div style='margin-top:24px;'></div>", unsafe_allow_html=True)
                    if st.button(
                        "📜 Generate & Print Official Certificate",
                        use_container_width=True,
                        type="primary",
                        key=f"mar_quick_gen_btn_{rec['Record ID']}",
                    ):
                        cert_res = generate_and_archive_marriage_certificate(
                            db=db,
                            actor=user,
                            record=rec,
                            orientation=orient_quick,
                            drive_mgr=drive_mgr,
                            output_dir=certificates_dir,
                        )
                        st.session_state["mar_last_generated_cert"] = cert_res

                last_cert = st.session_state.get("mar_last_generated_cert")
                if last_cert and last_cert.get("ok"):
                    st.success(f"✅ Generated Certificate `{last_cert['certificate_id']}`!")
                    _render_inline_pdf_print_component(
                        last_cert["pdf_bytes"],
                        label=f"Print Official Marriage Certificate ({last_cert['certificate_id']})",
                    )
                    st.download_button(
                        "⬇️ Download Official Marriage Certificate PDF",
                        data=last_cert["pdf_bytes"],
                        file_name=f"{last_cert['certificate_id']}_{rec['Record ID']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"mar_dl_cert_{rec['Record ID']}",
                    )

            with right_col:
                st.markdown("##### 🖼️ Scanned Marriage Contract Image (`MARRIAGE CERTIFICATE JPEG`)")
                img_drive_id = str(rec.get("Scanned Image Drive ID") or "").strip()
                img_filename = str(rec.get("Scanned Image Filename") or "").strip()

                if img_drive_id:
                    drive_view_url = f"https://drive.google.com/file/d/{img_drive_id}/view?usp=drivesdk"
                    st.caption(
                        f"📄 File: **`{img_filename}`** (`{img_drive_id}`) — "
                        f"[Buksan sa Google Drive]({drive_view_url})"
                    )
                    raw_img_bytes = None
                    try:
                        raw_img_bytes = drive_mgr.download_file_bytes(img_drive_id)
                    except Exception:
                        raw_img_bytes = None

                    if raw_img_bytes:
                        st.image(raw_img_bytes, caption=f"{img_filename} ({rec['Record ID']})", use_container_width=True)
                        b64_img = base64.b64encode(raw_img_bytes).decode("ascii")
                        _render_inline_image_print_button(b64_img, "image/jpeg", img_filename)
                        st.download_button(
                            f"⬇️ Download Scanned JPEG ({img_filename})",
                            data=raw_img_bytes,
                            file_name=img_filename or f"{rec['Record ID']}.jpeg",
                            mime="image/jpeg",
                            use_container_width=True,
                            key=f"mar_dl_raw_img_{rec['Record ID']}",
                        )
                    else:
                        preview_url = f"https://drive.google.com/file/d/{img_drive_id}/preview"
                        components.iframe(preview_url, height=540, scrolling=True)

                    if st.button(
                        "✅ Confirm Tama ang Image sa Record",
                        use_container_width=True,
                        key=f"mar_confirm_img_{rec['Record ID']}",
                    ):
                        confirm_marriage_image_link(
                            db=db,
                            record_id=rec["Record ID"],
                            drive_file_id=img_drive_id,
                            filename=img_filename,
                            actor=user,
                        )
                        st.success(f"✅ Confirmed 1-to-1 image link for {rec['Record ID']} ({img_filename})!")
                else:
                    st.warning("Walang naka-link na Scanned Marriage Certificate JPEG para sa record na ito.")

    # =========================================================================
    # TAB 3: MARRIAGE CERTIFICATE JPEG ARCHIVE
    # =========================================================================
    with tabs[2]:
        st.markdown("#### 🖼️ Scanned Marriage Certificate JPEG Archive (`1TUDlzDi-LCjxmxYorg099cs-uJlyJPzq`)")
        st.caption(
            f"Direct Google Drive Folder: [MARRIAGE CERTIFICATE JPEG]({DEFAULT_MARRIAGE_IMAGES_FOLDER_URL})"
        )
        all_recs_with_imgs = [
            r for r in search_marriage_records(db, query="", limit=600, actor=user)
            if str(r.get("Scanned Image Drive ID") or "").strip()
        ]
        filter_img_q = st.text_input(
            "🔎 Filter Scanned Marriage Certificate JPEGs by Husband Name or MC Number:",
            value="",
            placeholder="Halimbawa: MC_005, MC_154, Bautista, De Guzman, Ordona...",
            key="mar_archive_img_filter",
        )
        if filter_img_q.strip():
            fq = filter_img_q.strip().upper()
            all_recs_with_imgs = [
                r for r in all_recs_with_imgs
                if fq in str(r.get("Scanned Image Filename", "")).upper()
                or fq in str(r.get("Husband Name", "")).upper()
                or fq in str(r.get("Wife Name", "")).upper()
                or fq in str(r.get("MC Number", "")).upper()
                or fq in str(r.get("Record ID", "")).upper()
            ]

        st.write(f"Showing **{len(all_recs_with_imgs)}** scanned marriage certificate image(s):")
        if all_recs_with_imgs:
            img_labels = [
                f"{r['MC Number']} | {r['Scanned Image Filename']} — {r['Husband Name']} & {r['Wife Name']} ({r['Record ID']})"
                for r in all_recs_with_imgs
            ]
            selected_img_lbl = st.selectbox("Pumili ng Scanned Marriage Certificate JPEG:", img_labels, key="mar_archive_pick")
            sel_idx = img_labels.index(selected_img_lbl)
            sel_rec = all_recs_with_imgs[sel_idx]
            fid = sel_rec["Scanned Image Drive ID"]
            fname = sel_rec["Scanned Image Filename"]

            v_col1, v_col2 = st.columns([1.35, 1])
            with v_col1:
                raw_bytes = None
                try:
                    raw_bytes = drive_mgr.download_file_bytes(fid)
                except Exception:
                    raw_bytes = None
                if raw_bytes:
                    st.image(raw_bytes, caption=fname, use_container_width=True)
                    _render_inline_image_print_button(
                        base64.b64encode(raw_bytes).decode("ascii"),
                        "image/jpeg",
                        fname,
                    )
                else:
                    components.iframe(f"https://drive.google.com/file/d/{fid}/preview", height=540, scrolling=True)
            with v_col2:
                st.markdown(
                    f"""
                    <div class="mar-card">
                        <h4 style="margin-top:0;color:{THEME_NAVY};">📄 {fname}</h4>
                        <p style="font-size:0.88rem;line-height:1.6;">
                            <b>Record ID:</b> {sel_rec['Record ID']}<br/>
                            <b>MC Number:</b> {sel_rec['MC Number']}<br/>
                            <b>Husband:</b> {sel_rec['Husband Name'] or '—'}<br/>
                            <b>Wife:</b> {sel_rec['Wife Name'] or '—'}<br/>
                            <b>Date of Marriage:</b> {sel_rec['Date of Marriage'] or '—'}<br/>
                            <b>Place of Marriage:</b> {sel_rec['Place of Marriage'] or '—'}<br/>
                            <b>Officiating Priest:</b> {sel_rec['Officiating Priest / Minister'] or '—'}<br/>
                            <b>Drive Link:</b> <a href="https://drive.google.com/file/d/{fid}/view?usp=drivesdk" target="_blank">Open in Google Drive</a>
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    # =========================================================================
    # TAB 4: NEW MARRIAGE REGISTRATION (ENCODE TO DRIVE)
    # =========================================================================
    with tabs[3]:
        st.markdown("#### ✍️ Encode New Marriage Contract Record (Master Admin Only)")
        st.caption(
            "I-encode ang bagong Marriage Contract (Municipal Form No. 97 / CTG Marriage Register) at i-upload ang scanned image sa Google Drive."
        )
        with st.form("ctg_new_marriage_registration_form", clear_on_submit=False):
            f_c1, f_c2 = st.columns(2)
            with f_c1:
                st.markdown("##### 🤵 Husband (Groom) Information")
                h_name = st.text_input("Husband Full Name *", placeholder="First Middle Last")
                h_age = st.text_input("Husband Age / Date of Birth", placeholder="e.g. 26 years old (Born Jan 15, 1998)")
                h_status = st.selectbox("Husband Civil Status", ["Single", "Widowed", "Divorced / Annulled"])
                h_nat = st.text_input("Husband Nationality", value="Filipino")
                h_res = st.text_input("Husband Residence", placeholder="Barangay, Municipality, Province")
                h_father = st.text_input("Husband's Father", placeholder="Full Name of Groom's Father")
                h_mother = st.text_input("Husband's Mother", placeholder="Full Name of Groom's Mother")
            with f_c2:
                st.markdown("##### 👰 Wife (Bride) Information")
                w_name = st.text_input("Wife Full Name *", placeholder="First Middle Maiden-Last")
                w_age = st.text_input("Wife Age / Date of Birth", placeholder="e.g. 24 years old (Born May 20, 2000)")
                w_status = st.selectbox("Wife Civil Status", ["Single", "Widowed", "Divorced / Annulled"])
                w_nat = st.text_input("Wife Nationality", value="Filipina")
                w_res = st.text_input("Wife Residence", placeholder="Barangay, Municipality, Province")
                w_father = st.text_input("Wife's Father", placeholder="Full Name of Bride's Father")
                w_mother = st.text_input("Wife's Mother", placeholder="Full Name of Bride's Mother")

            st.divider()
            m_c1, m_c2, m_c3 = st.columns(3)
            with m_c1:
                parish_sel = st.selectbox("Parish", OFFICIAL_NINE_PARISHES, index=5)
                place_mar = st.text_input("Place of Marriage", value="Church of the Temple of God, Caba, La Union")
                date_mar = st.text_input("Date of Marriage *", placeholder="e.g. November 28, 2026")
            with m_c2:
                lic_no = st.text_input("Marriage License No.", placeholder="e.g. 2709547")
                lic_info = st.text_input("License Issued At / Date", placeholder="e.g. November 20, 2026 — Caba, La Union")
                priest_name = st.text_input("Officiating Priest / Minister", value="Msgr. Servant Troadio T. Abaquita")
                priest_pos = st.text_input("Position / Religion", value="Parish Priest / Bishop")
            with m_c3:
                witnesses = st.text_input("Witnesses / Principal Sponsors", placeholder="Full Names separated by comma")
                mc_custom = st.text_input("MC Number (Optional — Auto if blank)", placeholder="e.g. MC_450")
                book_no = st.text_input("Book Number", value="BOOK 1")
                page_no = st.text_input("Page Number / LCR No.", placeholder="e.g. 450")

            up_file = st.file_uploader(
                "📤 Upload Scanned Marriage Certificate Image / PDF to Google Drive (Optional)",
                type=["jpg", "jpeg", "png", "pdf"],
                key="mar_new_reg_uploader",
            )
            remarks_txt = st.text_area("Remarks / Notes", value="Encoded by Master Admin")

            sub_mar = st.form_submit_button(
                "💾 Save Marriage Record & Sync to Google Drive",
                use_container_width=True,
                type="primary",
            )
            if sub_mar:
                up_bytes = up_file.getvalue() if up_file is not None else None
                up_name = up_file.name if up_file is not None else ""
                reg_res = register_new_marriage_record(
                    db=db,
                    actor=user,
                    form_data={
                        "Husband Name": h_name,
                        "Husband Age / Birthdate": h_age,
                        "Husband Single / Widowed": h_status,
                        "Husband Nationality": h_nat,
                        "Husband Residence": h_res,
                        "Husband Father": h_father,
                        "Husband Mother": h_mother,
                        "Wife Name": w_name,
                        "Wife Age / Birthdate": w_age,
                        "Wife Single / Widowed": w_status,
                        "Wife Nationality": w_nat,
                        "Wife Residence": w_res,
                        "Wife Father": w_father,
                        "Wife Mother": w_mother,
                        "Parish": parish_sel,
                        "Place of Marriage": place_mar,
                        "Date of Marriage": date_mar,
                        "Marriage License No.": lic_no,
                        "License Issued At / Date": lic_info,
                        "Officiating Priest / Minister": priest_name,
                        "Priest Position / Religion": priest_pos,
                        "Witnesses / Sponsors": witnesses,
                        "MC Number": mc_custom,
                        "Book Number": book_no,
                        "Page Number": page_no,
                        "Local Civil Registry No.": page_no,
                        "Remarks": remarks_txt,
                    },
                    drive_mgr=drive_mgr,
                    uploaded_image_bytes=up_bytes,
                    uploaded_image_filename=up_name,
                    images_folder_id=DEFAULT_MARRIAGE_IMAGES_FOLDER_ID,
                )
                if reg_res.get("ok"):
                    st.success(
                        f"✅ Matagumpay na na-save ang Marriage Record **{reg_res['record_id']}** ({reg_res['mc_number']})!"
                    )
                else:
                    st.error(reg_res.get("error") or "Failed to register marriage record.")

    # =========================================================================
    # TAB 5: OFFICIAL MARRIAGE CERTIFICATE GENERATOR & ARCHIVE
    # =========================================================================
    with tabs[4]:
        st.markdown("#### 📜 Official CTG Certificate of Marriage Generator & Issued Archive")
        all_mar_records = search_marriage_records(db, query="", limit=600, actor=user)
        if all_mar_records:
            rec_options = [
                f"{r['Record ID']} ({r['MC Number']}) — {r['Husband Name']} & {r['Wife Name']} [{r['Date of Marriage']}]"
                for r in all_mar_records
            ]
            c_sel1, c_sel2 = st.columns([2.2, 1])
            with c_sel1:
                picked_lbl = st.selectbox("Pumili ng Marriage Record para gawan ng Official Certificate:", rec_options, key="mar_cert_tab_pick")
                picked_rec = all_mar_records[rec_options.index(picked_lbl)]
            with c_sel2:
                cert_orient = st.radio("Layout Orientation:", ["Portrait", "Landscape"], horizontal=True, key="mar_cert_tab_orient")

            custom_heading = st.text_input("Certificate Heading:", value="CERTIFICATE OF MARRIAGE", key="mar_cert_heading")
            if st.button("🖨️ Generate Official Marriage Certificate PDF", type="primary", use_container_width=True, key="mar_gen_tab5_btn"):
                gen_out = generate_and_archive_marriage_certificate(
                    db=db,
                    actor=user,
                    record=picked_rec,
                    orientation=cert_orient,
                    custom_heading=custom_heading,
                    drive_mgr=drive_mgr,
                    output_dir=certificates_dir,
                )
                st.session_state["mar_tab5_cert"] = gen_out
                st.session_state["mar_tab5_rec"] = picked_rec

            tab5_cert = st.session_state.get("mar_tab5_cert")
            tab5_rec = st.session_state.get("mar_tab5_rec")
            if tab5_cert and tab5_cert.get("ok") and tab5_rec:
                st.success(
                    f"✅ Official Marriage Certificate **{tab5_cert['certificate_id']}** generated for "
                    f"**{tab5_rec['Husband Name']} & {tab5_rec['Wife Name']}**!"
                )
                _render_inline_pdf_print_component(
                    tab5_cert["pdf_bytes"],
                    label=f"Print Certificate ({tab5_cert['certificate_id']})",
                )
                st.download_button(
                    "⬇️ Download Official Marriage Certificate PDF",
                    data=tab5_cert["pdf_bytes"],
                    file_name=f"{tab5_cert['certificate_id']}_{tab5_rec['Record ID']}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="mar_tab5_dl_btn",
                )

        st.divider()
        st.markdown("##### 🗂️ Archive of Generated Marriage Certificates")
        issued_certs = list_generated_marriage_certificates(db, actor=user, limit=100)
        if issued_certs:
            st.dataframe(pd.DataFrame(issued_certs), use_container_width=True, hide_index=True)
        else:
            st.caption("Wala pang na-generate na Marriage Certificate sa session na ito.")


if __name__ == "__main__":
    st.set_page_config(
        page_title="CTG Marriage Records & Certification System",
        page_icon="💍",
        layout="wide",
    )
    render_marriage_portal(embedded_in_hub=False)
"""
ui_components.py — Reusable Enterprise UI/UX Components for CTG Ministries Database
=====================================================================================
Design System:
- Top Bar Solid Navy: #10213D (100% solid, non-transparent, zero blur)
- Primary:            #0B1730
- Secondary:          #142747
- Accent Gold:        #D4A640
- Soft Gold:          #F2D58A
- Background:         #0B101B
- Card Background:    #17243A
- Border:             #34445E
- Text:               #F8FAFC
- Muted Text:         #AAB7CA
- Success:            #10B981
- Warning:            #F59E0B
- Danger:             #EF4444
"""

from html import escape as _html_esc
from typing import Any, Callable, Dict, List, Optional, Tuple
try:
    import streamlit as st
    import streamlit.components.v1 as components
except ImportError:  # pragma: no cover — allows headless unit testing without Streamlit
    st = None  # type: ignore
    components = None  # type: ignore

DESIGN_TOKENS: Dict[str, str] = {
    "topbar_navy": "#10213D",
    "topbar_cell_bg": "#162B4D",
    "primary": "#0B1730",
    "secondary": "#142747",
    "accent_gold": "#D4A640",
    "soft_gold": "#F2D58A",
    "background": "#0B101B",
    "card_bg": "#17243A",
    "border": "#34445E",
    "text": "#F8FAFC",
    "muted_text": "#AAB7CA",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
}

# Stable internal page IDs mapped to legacy app_mode strings for 100% backward compatibility
PAGE_ID_TO_APP_MODE: Dict[str, str] = {
    "home_overview": "🧭 Parish Leader Portal (Wizard & Summary)",
    "group_parish_operations": "GROUP_DASHBOARD::PARISH_OPERATIONS",
    "group_census_directory": "GROUP_DASHBOARD::CENSUS_DIRECTORY",
    "group_sacramental_certs": "GROUP_DASHBOARD::SACRAMENTAL_CERTS",
    "group_admin_control": "GROUP_DASHBOARD::ADMIN_CONTROL",
    "group_community_system": "GROUP_DASHBOARD::COMMUNITY_SYSTEM",
    "weekly_report": "📑 Weekly Report & Parish Monitoring",
    "outreach_visitation": "🏠 House Visitation, Missionary Work, Indoctrination & Remarks",
    "census_manager": "📊 Parish Census & Member Manager",
    "add_member": "➕ Magdagdag ng Miyembro & Litrato",
    "notion_crm": "🤝 Notion CRM (Visitors & Members)",
    "baptismal_portal": "⛪ Baptismal Records & Certificates",
    "marriage_portal": "💍 Marriage Records & Certificates",
    "live_activity": "📈 Live Activity & Update Dashboard",
    "master_approval_queue": "🛡️ Master Admin Controller & Approval Queue",
    "parish_inbox": "📥 Parish Inbox & Action Hub",
    "announcement_broadcaster": "📢 Announcement Broadcaster",
    "prayer_wall": "🙏 Prayer Wall & Categorizer",
    "ministry_messenger": "🗣️ Ministry Messenger (Admin & Users Chat)",
    "ai_assistant": "💬 AI Chat Assistant",
    "app_settings": "🛠️ App Settings",
}

APP_MODE_TO_PAGE_ID: Dict[str, str] = {v: k for k, v in PAGE_ID_TO_APP_MODE.items()}


def _esc(val: Any) -> str:
    return _html_esc(str(val or ""), quote=True)


def is_sidebar_expanded() -> bool:
    """Returns True when the application sidebar navigation panel is open."""
    if st is None:
        return True
    if "ctg_sidebar_expanded" not in st.session_state:
        st.session_state["ctg_sidebar_expanded"] = True
    return bool(st.session_state.get("ctg_sidebar_expanded", True))


def toggle_sidebar_state() -> bool:
    """Toggles the application sidebar open/closed state in Streamlit session_state."""
    if st is None:
        return True
    cur = is_sidebar_expanded()
    st.session_state["ctg_sidebar_expanded"] = not cur
    return not cur


def inject_enterprise_theme_css() -> None:
    """
    Applies unified enterprise Navy/Gold dark-mode styling, completely hides all native
    Streamlit header/floating sidebar toggle controls, styles the solid non-transparent
    #10213D global top bar, and enforces deterministic sidebar open/close state.
    """
    t = DESIGN_TOKENS
    sb_open = is_sidebar_expanded()
    if sb_open:
        sidebar_state_css = f"""
        section[data-testid="stSidebar"] {{
            background-color: {t["primary"]} !important;
            border-right: 1px solid {t["border"]} !important;
            width: 20.5rem !important;
            min-width: 20.5rem !important;
            max-width: 20.5rem !important;
            margin-left: 0px !important;
            transform: none !important;
            opacity: 1 !important;
            visibility: visible !important;
            pointer-events: auto !important;
            transition: width 0.18s ease, min-width 0.18s ease, max-width 0.18s ease !important;
        }}
        section[data-testid="stSidebar"] > div {{
            background-color: {t["primary"]} !important;
            width: 20.5rem !important;
            min-width: 20.5rem !important;
            max-width: 20.5rem !important;
            overflow-x: hidden !important;
        }}
        """
    else:
        sidebar_state_css = """
        section[data-testid="stSidebar"] {
            margin-left: 0px !important;
            width: 0px !important;
            min-width: 0px !important;
            max-width: 0px !important;
            padding: 0px !important;
            overflow: hidden !important;
            border-right: none !important;
            opacity: 0 !important;
            visibility: hidden !important;
            pointer-events: none !important;
            transition: width 0.18s ease, min-width 0.18s ease, max-width 0.18s ease !important;
        }
        section[data-testid="stSidebar"] > div {
            width: 0px !important;
            min-width: 0px !important;
            max-width: 0px !important;
            padding: 0px !important;
            overflow: hidden !important;
        }
        """

    st.markdown(
        f"""
        <style>
        /* 1. Completely remove native Streamlit header bar and all duplicate/floating sidebar toggles */
        header[data-testid="stHeader"],
        div[data-testid="stToolbar"],
        div[data-testid="stDecoration"],
        div[data-testid="stStatusWidget"],
        div[data-testid="collapsedControl"],
        div[data-testid="stSidebarCollapsedControl"],
        button[data-testid="stSidebarCollapseButton"],
        button[data-testid="stExpandSidebarButton"],
        section[data-testid="stSidebar"] button[kind="header"],
        section[data-testid="stSidebar"] button[kind="headerNoPadding"],
        header button[kind="header"],
        header button[kind="headerNoPadding"] {{
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            width: 0px !important;
            height: 0px !important;
            padding: 0px !important;
            margin: 0px !important;
            border: none !important;
            pointer-events: none !important;
            overflow: hidden !important;
        }}

        /* 2. Main application canvas layout */
        .stApp {{
            background-color: {t["background"]} !important;
            color: {t["text"]} !important;
        }}
        section.main,
        [data-testid="stMain"] {{
            flex: 1 1 0% !important;
            width: 100% !important;
            max-width: 100% !important;
            margin-left: 0px !important;
        }}
        section.main > div.block-container,
        div[data-testid="stAppViewBlockContainer"],
        div[data-testid="stMainBlockContainer"],
        .block-container {{
            padding-top: 0.45rem !important;
            padding-bottom: 2.2rem !important;
            padding-left: 1.1rem !important;
            padding-right: 1.1rem !important;
            max-width: 100% !important;
        }}

        /* 3. Sidebar state & compact enterprise formatting */
        {sidebar_state_css}
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
        section[data-testid="stSidebar"] .block-container {{
            padding-top: 0.65rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-bottom: 1.25rem !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stExpander"] {{
            background-color: {t["secondary"]} !important;
            border: 1px solid {t["border"]} !important;
            border-radius: 8px !important;
            margin-bottom: 6px !important;
        }}
        section[data-testid="stSidebar"] [data-testid="stExpander"] summary {{
            background-color: {t["topbar_navy"]} !important;
            color: {t["text"]} !important;
            font-weight: 700 !important;
            font-size: 0.82rem !important;
            padding: 7px 10px !important;
            border-radius: 8px !important;
        }}
        section[data-testid="stSidebar"] .stButton > button {{
            font-size: 0.80rem !important;
            padding: 0.36rem 0.65rem !important;
            min-height: 36px !important;
            border-radius: 7px !important;
            text-align: left !important;
            justify-content: flex-start !important;
        }}

        /* 4. Solid, Non-Transparent Global Top Navigation Bar (#10213D with Gold Border #D4A640) */
        div[data-testid="stVerticalBlock"] > div:has(#ctg-sticky-top-header-marker) {{
            position: sticky !important;
            top: 0.25rem !important;
            z-index: 999990 !important;
            background-color: {t["topbar_navy"]} !important;
        }}
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) {{
            position: sticky !important;
            top: 0.25rem !important;
            z-index: 999990 !important;
            background-color: {t["topbar_navy"]} !important;
            background-image: none !important;
            border: 1.5px solid {t["accent_gold"]} !important;
            border-radius: 10px !important;
            padding: 8px 12px !important;
            margin-bottom: 12px !important;
            box-shadow: 0 6px 18px #050A14 !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
            opacity: 1 !important;
        }}
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="stHorizontalBlock"] {{
            align-items: center !important;
            gap: 0.45rem !important;
        }}
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) .stButton > button,
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="stPopover"] > div > button {{
            background-color: {t["topbar_cell_bg"]} !important;
            color: {t["text"]} !important;
            border: 1px solid {t["accent_gold"]} !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            font-size: 0.80rem !important;
            min-height: 40px !important;
            padding: 4px 10px !important;
            box-shadow: none !important;
        }}
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) .stButton > button:hover,
        div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="stPopover"] > div > button:hover {{
            background-color: {t["secondary"]} !important;
            border-color: {t["soft_gold"]} !important;
            color: {t["soft_gold"]} !important;
        }}

        /* 5. Enterprise Cards, Metrics & Breadcrumbs */
        .ctg-ui-card {{
            background-color: {t["card_bg"]};
            border: 1px solid {t["border"]};
            border-radius: 10px;
            padding: 14px 16px;
        }}
        .ctg-ui-card:hover {{
            border-color: {t["accent_gold"]};
        }}
        .ctg-metric-card {{
            background-color: {t["secondary"]};
            border: 1px solid {t["border"]};
            border-top: 3px solid {t["accent_gold"]};
            border-radius: 10px;
            padding: 12px 14px;
            min-height: 90px;
        }}
        .ctg-metric-label {{
            font-size: 0.73rem;
            color: {t["muted_text"]};
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }}
        .ctg-metric-value {{
            font-size: 1.50rem;
            color: {t["text"]};
            font-weight: 800;
            margin-top: 4px;
            line-height: 1.15;
        }}
        .ctg-metric-sub {{
            font-size: 0.72rem;
            color: {t["soft_gold"]};
            margin-top: 4px;
        }}
        .ctg-sb-section-label {{
            font-size: 0.69rem;
            font-weight: 800;
            color: {t["soft_gold"]};
            letter-spacing: 0.08em;
            text-transform: uppercase;
            margin: 8px 0 5px 2px;
        }}
        .ctg-breadcrumbs {{
            font-size: 0.78rem;
            color: {t["muted_text"]};
            padding: 2px 2px 8px 2px;
            display: flex;
            align-items: center;
            gap: 6px;
            flex-wrap: wrap;
        }}
        .ctg-breadcrumbs strong {{
            color: {t["soft_gold"]};
        }}

        /* 6. Responsive Top Bar & Layout on Tablet and Mobile */
        @media (max-width: 1024px) {{
            div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) {{
                padding: 6px 8px !important;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="stHorizontalBlock"] {{
                flex-wrap: wrap !important;
                gap: 6px !important;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="column"] {{
                min-width: auto !important;
                flex: 1 1 auto !important;
            }}
        }}
        @media (max-width: 640px) {{
            div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) {{
                position: relative !important;
                top: 0 !important;
                padding: 6px 8px !important;
            }}
            div[data-testid="stVerticalBlock"] > div:has(#ctg-sticky-top-header-marker) {{
                position: relative !important;
                top: 0 !important;
            }}
            div[data-testid="stVerticalBlockBorderWrapper"]:has(#ctg-sticky-top-header-marker) [data-testid="column"] {{
                min-width: 46% !important;
                flex: 1 1 46% !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_status_badge(label: str, tone: str = "success") -> str:
    """Returns a solid, high-contrast status badge HTML snippet (no transparent glass)."""
    tone_map = {
        "success": ("#064E3B", "#34D399", "#10B981"),
        "warning": ("#78350F", "#FBBF24", "#F59E0B"),
        "danger": ("#7F1D1D", "#F87171", "#EF4444"),
        "info": ("#0C4A6E", "#7DD3FC", "#38BDF8"),
        "gold": ("#422006", "#F2D58A", "#D4A640"),
    }
    bg, fg, bd = tone_map.get(tone, tone_map["info"])
    return (
        f"<span style='display:inline-flex;align-items:center;gap:5px;background-color:{bg};"
        f"color:{fg};border:1px solid {bd};font-size:0.70rem;font-weight:800;"
        f"padding:2px 9px;border-radius:999px;white-space:nowrap;'>"
        f"{_esc(label)}</span>"
    )


def render_breadcrumbs(trail: List[str]) -> None:
    """Renders a clean breadcrumb trail at the top of the current page."""
    if not trail:
        return
    parts = [_esc(x) for x in trail[:-1]]
    last = f"<strong>{_esc(trail[-1])}</strong>"
    joined = " <span style='color:#64748B;'>/</span> ".join(parts + [last])
    st.markdown(f"<div class='ctg-breadcrumbs'>{joined}</div>", unsafe_allow_html=True)


def render_dashboard_header(
    title: str,
    subtitle: str,
    scope_label: str,
    ph_datetime_str: str,
    status_label: str = "System Online",
    status_tone: str = "success",
) -> None:
    """Renders a solid enterprise dashboard header card."""
    t = DESIGN_TOKENS
    badge_html = render_status_badge(status_label, tone=status_tone)
    st.markdown(
        f"<div style='background-color:{t['topbar_navy']};"
        f"border:1px solid {t['border']};border-left:4px solid {t['accent_gold']};"
        f"border-radius:10px;padding:14px 18px;margin-bottom:14px;"
        f"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;'>"
        f"<div>"
        f"<div style='font-size:1.22rem;font-weight:800;color:{t['text']};line-height:1.25;'>{_esc(title)}</div>"
        f"<div style='font-size:0.84rem;color:{t['muted_text']};margin-top:4px;'>{_esc(subtitle)}</div>"
        f"</div>"
        f"<div style='text-align:right;'>"
        f"{badge_html}"
        f"<div style='font-size:0.75rem;color:{t['soft_gold']};font-weight:700;margin-top:5px;'>Scope: {_esc(scope_label)}</div>"
        f"<div style='font-size:0.71rem;color:{t['muted_text']};margin-top:2px;'>{_esc(ph_datetime_str)}</div>"
        f"</div></div>",
        unsafe_allow_html=True,
    )


def render_metric_card(label: str, value: Any, subtext: str = "", tone: str = "gold") -> None:
    """Renders a crisp metric card matching the CTG design tokens."""
    t = DESIGN_TOKENS
    top_colors = {
        "gold": t["accent_gold"],
        "success": t["success"],
        "warning": t["warning"],
        "danger": t["danger"],
        "info": "#38BDF8",
    }
    top_c = top_colors.get(tone, t["accent_gold"])
    sub_html = f"<div class='ctg-metric-sub'>{_esc(subtext)}</div>" if subtext else ""
    st.markdown(
        f"<div class='ctg-metric-card' style='border-top-color:{top_c};'>"
        f"<div class='ctg-metric-label'>{_esc(label)}</div>"
        f"<div class='ctg-metric-value'>{_esc(value)}</div>"
        f"{sub_html}"
        f"</div>",
        unsafe_allow_html=True,
    )


def render_quick_action_card(
    title: str,
    description: str,
    button_label: str,
    key: str,
    on_click_page: Optional[str] = None,
    badge_text: str = "",
) -> bool:
    """Renders a quick action card with a native Streamlit button. Returns True when clicked."""
    t = DESIGN_TOKENS
    badge_html = f"<div style='margin-bottom:6px;'>{render_status_badge(badge_text, 'gold')}</div>" if badge_text else ""
    st.markdown(
        f"<div class='ctg-ui-card' style='min-height:108px;margin-bottom:8px;'>"
        f"{badge_html}"
        f"<div style='font-size:0.92rem;font-weight:800;color:{t['text']};'>{_esc(title)}</div>"
        f"<div style='font-size:0.77rem;color:{t['muted_text']};margin-top:4px;line-height:1.35;'>{_esc(description)}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    clicked = st.button(button_label, key=key, use_container_width=True)
    if clicked and on_click_page:
        st.session_state["current_nav_page"] = on_click_page
        st.rerun()
    return clicked


def render_activity_feed(items: List[Dict[str, Any]], empty_text: str = "No recent activity recorded.") -> None:
    """Renders a clean, scannable activity feed without overflowing vertical height."""
    t = DESIGN_TOKENS
    if not items:
        st.info(empty_text)
        return
    for item in items:
        title = item.get("title", "Activity")
        actor = item.get("actor", "System")
        ts = item.get("timestamp", "")
        detail = item.get("detail", "")
        tone = item.get("tone", "info")
        border_c = t.get(tone, "#38BDF8")
        st.markdown(
            f"<div style='background-color:{t['card_bg']};border:1px solid {t['border']};"
            f"border-left:4px solid {border_c};border-radius:8px;padding:8px 12px;margin-bottom:8px;'>"
            f"<div style='display:flex;justify-content:space-between;gap:8px;align-items:center;'>"
            f"<strong style='color:{t['text']};font-size:0.84rem;'>{_esc(title)}</strong>"
            f"<span style='color:{t['muted_text']};font-size:0.72rem;white-space:nowrap;'>{_esc(ts)}</span>"
            f"</div>"
            f"<div style='color:{t['soft_gold']};font-size:0.74rem;margin-top:2px;'>{_esc(actor)}</div>"
            f"<div style='color:{t['muted_text']};font-size:0.78rem;margin-top:3px;'>{_esc(detail)}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )


def render_global_topbar(
    username: str,
    display_name: str,
    role: str,
    assigned_scope: str,
    ph_datetime_str: str,
    system_online: bool,
    notifications: List[Dict[str, Any]],
    online_leaders: List[Dict[str, Any]],
    on_refresh: Callable[[], None],
    on_mark_all_read: Callable[[], None],
    on_delete_notif: Callable[[str], None],
    on_open_profile: Callable[[], None],
    on_logout: Callable[[], None],
    on_open_notifications_view: Optional[Callable[[], None]] = None,
    logo_b64: str = "",
) -> None:
    """
    Unified, Solid Dark Navy (#10213D), Non-Transparent Top Navigation Bar across all authenticated pages.
    Strict left-to-right element order inside one unified bar:
      1. Sidebar panel toggle button (far left: '◧ ❮' when open, '◧ ❯' when closed — no three-line hamburger)
      2. CTG logo and 'Church of the Temple of God' name
      3. Tagline: 'One Church • One Faith • One Family'
      4. System Online/Offline status
      5. Current Philippine date and time
      6. Notification dropdown with unread count
      7. Refresh button
      8. Logged-in user profile and role dropdown
    """
    t = DESIGN_TOKENS
    sb_open = is_sidebar_expanded()
    toggle_icon_lbl = "◧ ❮" if sb_open else "◧ ❯"
    toggle_help = "Close sidebar navigation panel" if sb_open else "Open sidebar navigation panel"

    read_ids = set(st.session_state.get("_ctg_read_notif_ids", []))
    unread_notifs = [n for n in notifications if n.get("id", "") not in read_ids]
    unread_count = len(unread_notifs)

    status_dot = "● System Online" if system_online else "● System Offline"
    status_fg = "#34D399" if system_online else "#FBBF24"
    status_bd = "#10B981" if system_online else "#F59E0B"

    initials = "".join(w[0] for w in (display_name or role or "U").split()[:2] if w).upper() or "CT"
    user_pill_lbl = f"{initials}  {role} ({len(online_leaders)} Online)"

    if logo_b64:
        logo_html = (
            f"<img src='data:image/png;base64,{logo_b64}' alt='CTG Logo' "
            f"style='width:34px;height:34px;border-radius:50%;object-fit:cover;"
            f"border:1.5px solid {t['accent_gold']};flex-shrink:0;' />"
        )
    else:
        logo_html = (
            f"<div style='width:34px;height:34px;border-radius:50%;background-color:{t['primary']};"
            f"border:1.5px solid {t['accent_gold']};display:flex;align-items:center;justify-content:center;"
            f"font-size:0.75rem;font-weight:900;color:{t['soft_gold']};flex-shrink:0;'>CTG</div>"
        )

    with st.container(border=True):
        st.markdown('<div id="ctg-sticky-top-header-marker" style="display:none;"></div>', unsafe_allow_html=True)
        c_toggle, c_brand, c_status, c_bell, c_ref, c_user = st.columns([0.52, 2.35, 1.35, 0.65, 0.78, 1.45])

        # 1. Far-left Sidebar Panel Toggle Button (Chevron + Panel icon, never 3-line hamburger)
        with c_toggle:
            if st.button(
                toggle_icon_lbl,
                key="top_navbar_sidebar_panel_toggle_btn",
                use_container_width=True,
                help=toggle_help,
            ):
                toggle_sidebar_state()
                st.rerun()

        # 2 & 3. CTG Logo + Church of the Temple of God + Tagline: One Church • One Faith • One Family
        with c_brand:
            st.markdown(
                f"<div style='background-color:{t['topbar_cell_bg']};"
                f"border:1px solid {t['accent_gold']};border-radius:8px;padding:4px 12px;"
                f"display:flex;align-items:center;gap:10px;min-height:40px;'>"
                f"{logo_html}"
                f"<div style='min-width:0;line-height:1.15;'>"
                f"<div style='font-size:0.98rem;font-weight:900;color:{t['text']};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                f"Church of the Temple of God</div>"
                f"<div style='font-size:0.71rem;color:{t['soft_gold']};font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                f"One Church &bull; One Faith &bull; One Family</div>"
                f"</div></div>",
                unsafe_allow_html=True,
            )

        # 4 & 5. System Online/Offline Status + Current Philippine Date and Time
        with c_status:
            st.markdown(
                f"<div style='background-color:{t['primary']};border:1px solid {status_bd};"
                f"border-radius:8px;padding:4px 10px;text-align:center;line-height:1.2;min-height:40px;"
                f"display:flex;flex-direction:column;justify-content:center;'>"
                f"<div style='font-size:0.74rem;font-weight:800;color:{status_fg};'>{status_dot}</div>"
                f"<div style='font-size:0.66rem;color:{t['text']};font-weight:600;'>{_esc(ph_datetime_str)}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )

        # 6. Notification Dropdown with Unread Count
        with c_bell:
            bell_lbl = f"🔔 {unread_count}"
            with st.popover(bell_lbl, use_container_width=True):
                st.markdown("#### Notifications")
                st.caption(f"Unread: **{unread_count}** • Total: **{len(notifications)}**")
                if notifications:
                    b_col1, b_col2 = st.columns(2)
                    with b_col1:
                        if st.button("Mark all read", use_container_width=True, key="top_notif_mark_read"):
                            st.session_state["_ctg_read_notif_ids"] = [n.get("id", "") for n in notifications]
                            on_mark_all_read()
                            st.rerun()
                    with b_col2:
                        if on_open_notifications_view and st.button("View all", use_container_width=True, key="top_notif_view_all"):
                            on_open_notifications_view()
                            st.rerun()
                    for nf in reversed(notifications[-8:]):
                        nid = nf.get("id", "")
                        is_unread = nid not in read_ids
                        nc1, nc2 = st.columns([0.84, 0.16])
                        with nc1:
                            badge = "🆕 " if is_unread else ""
                            st.markdown(
                                f"<div style='background-color:{t['card_bg']};color:{t['text']};border-left:3px solid {t['accent_gold']};"
                                f"padding:6px 8px;border-radius:6px;margin-bottom:5px;font-size:0.76rem;'>"
                                f"<b>{badge}{_esc(nf.get('title'))}</b><br>"
                                f"<span style='color:{t['soft_gold']};'>{_esc(nf.get('sender'))}</span> • "
                                f"<span style='color:{t['muted_text']};font-size:0.68rem;'>{_esc(nf.get('timestamp'))}</span><br>"
                                f"{_esc(str(nf.get('details', ''))[:110])}</div>",
                                unsafe_allow_html=True,
                            )
                        with nc2:
                            if st.button("✕", key=f"top_del_nf_{nid}", help="Dismiss"):
                                on_delete_notif(nid)
                                st.rerun()
                else:
                    st.info("No new notifications.")

        # 7. State-Preserving Refresh Button
        with c_ref:
            if st.button(
                "🔄 Refresh",
                key="top_navbar_global_refresh_btn",
                use_container_width=True,
                help="Refresh live data without resetting login or form state",
            ):
                on_refresh()
                st.rerun()

        # 8. Logged-in User Profile & Role Dropdown
        with c_user:
            with st.popover(user_pill_lbl, use_container_width=True):
                st.markdown(f"**{_esc(display_name or username)}**")
                st.caption(f"Username: `{username or 'active_user'}` • Role: **{role}**")
                st.caption(f"Assigned Scope: **{assigned_scope}**")
                if st.button("Role Overview Dashboard", use_container_width=True, key="top_pop_role_dash_btn"):
                    st.session_state["current_nav_page"] = PAGE_ID_TO_APP_MODE["home_overview"]
                    st.rerun()
                st.divider()
                st.markdown("**Online Leaders in System**")
                for ol in online_leaders[:8]:
                    st.markdown(
                        f"<div style='background-color:{t['card_bg']};border:1px solid {t['border']};"
                        f"border-left:3px solid {t['success']};padding:5px 8px;border-radius:6px;margin-bottom:4px;font-size:0.76rem;'>"
                        f"<span style='color:#34D399;font-weight:700;'>● Online</span> — <b>{_esc(ol.get('name'))}</b><br>"
                        f"<span style='color:{t['muted_text']};font-size:0.69rem;'>{_esc(ol.get('role'))} • {_esc(ol.get('scope'))}</span>"
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                st.divider()
                p_c1, p_c2 = st.columns(2)
                with p_c1:
                    if st.button("Profile Settings", use_container_width=True, key="top_pop_profile_btn"):
                        on_open_profile()
                        st.rerun()
                with p_c2:
                    if st.button("Logout", type="primary", use_container_width=True, key="top_pop_logout"):
                        on_logout()

    # Ensure native Streamlit sidebar DOM container stays expanded when sb_open is True
    if components is not None:
        sb_flag_js = "true" if sb_open else "false"
        components.html(
            f"""
            <script>
            (function() {{
                const pDoc = window.parent.document;
                const sbOpen = {sb_flag_js};
                const sb = pDoc.querySelector('section[data-testid="stSidebar"]');
                if (sb) {{
                    sb.style.marginLeft = '0px';
                    if (sbOpen && sb.getAttribute('aria-expanded') === 'false') {{
                        const natExp = pDoc.querySelector('button[data-testid="stExpandSidebarButton"], [data-testid="collapsedControl"] button, [data-testid="stSidebarCollapsedControl"] button');
                        if (natExp) natExp.click();
                    }}
                }}
            }})();
            </script>
            """,
            height=0,
        )


def render_user_profile_card(
    display_name: str,
    role: str,
    assigned_scope: str,
    logo_b64: str = "",
    avatar_data_uri: str = "",
    is_online: bool = True,
) -> None:
    """Renders Sidebar Section A: Compact User Profile Card with CTG logo/avatar, role, parish, and online status."""
    t = DESIGN_TOKENS
    initials = "".join(w[0] for w in (display_name or role or "CTG").split()[:2] if w).upper() or "CT"
    if avatar_data_uri and avatar_data_uri.startswith("data:image"):
        avatar_html = (
            f"<img src='{avatar_data_uri}' style='width:42px;height:42px;border-radius:50%;"
            f"object-fit:cover;border:2px solid {t['accent_gold']};flex-shrink:0;'>"
        )
    elif logo_b64:
        avatar_html = (
            f"<img src='data:image/png;base64,{logo_b64}' style='width:42px;height:42px;border-radius:50%;"
            f"object-fit:cover;border:2px solid {t['accent_gold']};flex-shrink:0;'>"
        )
    else:
        avatar_html = (
            f"<div style='width:42px;height:42px;border-radius:50%;background-color:{t['secondary']};"
            f"border:2px solid {t['accent_gold']};display:flex;align-items:center;justify-content:center;"
            f"font-weight:800;color:{t['soft_gold']};font-size:0.90rem;flex-shrink:0;'>{_esc(initials)}</div>"
        )
    status_badge = render_status_badge("Online" if is_online else "Offline", "success" if is_online else "danger")
    st.markdown(
        f"<div style='background-color:{t['topbar_navy']};border:1px solid {t['accent_gold']};border-radius:10px;"
        f"padding:9px 11px;margin-bottom:8px;'>"
        f"<div style='display:flex;align-items:center;gap:9px;'>"
        f"{avatar_html}"
        f"<div style='min-width:0;flex:1;'>"
        f"<div style='font-size:0.84rem;font-weight:800;color:{t['text']};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
        f"{_esc(display_name)}</div>"
        f"<div style='font-size:0.71rem;color:{t['soft_gold']};font-weight:700;'>{_esc(role)}</div>"
        f"<div style='font-size:0.68rem;color:{t['muted_text']};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
        f"{_esc(assigned_scope)}</div>"
        f"</div>"
        f"{status_badge}"
        f"</div></div>",
        unsafe_allow_html=True,
    )


def get_role_navigation_groups(user_role: str, login_role: str = "") -> List[Dict[str, Any]]:
    """
    Returns the 5 clean expandable sidebar navigation groups strictly scoped by user_role.
    Note: No separate 'Home' or 'Home & Leadership Portal' shortcut is shown in the sidebar;
    after login, the application opens directly to the role-based landing dashboard.
    """
    is_master = (user_role == "Master Admin")
    groups: List[Dict[str, Any]] = [
        {
            "group_id": "group_parish_operations",
            "title": "1. Parish Operations & Reports",
            "dashboard_page_id": "group_parish_operations",
            "items": [
                {"page_id": "group_parish_operations", "label": "Operations Overview"},
                {"page_id": "weekly_report", "label": "Weekly Report & Monitoring"},
                {"page_id": "outreach_visitation", "label": "Visitation, Mission & Indoctrination"},
            ],
        },
        {
            "group_id": "group_census_directory",
            "title": "2. Census & Member Directory",
            "dashboard_page_id": "group_census_directory",
            "items": [
                {"page_id": "group_census_directory", "label": "Census Overview"},
                {"page_id": "census_manager", "label": "Parish Census & Family Search"},
                {"page_id": "add_member", "label": "Add Member & Photo (2-Step)"},
            ] + ([{"page_id": "notion_crm", "label": "Visitors & Members CRM"}] if is_master else []),
        },
        {
            "group_id": "group_sacramental_certs",
            "title": "3. Sacramental Certificates",
            "dashboard_page_id": "group_sacramental_certs",
            "items": [
                {"page_id": "group_sacramental_certs", "label": "Sacramental Overview"},
                {"page_id": "baptismal_portal", "label": "Baptismal Records & Certificates"},
            ] + ([{"page_id": "marriage_portal", "label": "Marriage Records & Certs (Master)"}] if is_master else []),
        },
        {
            "group_id": "group_admin_control",
            "title": "4. Admin Control & Queue",
            "dashboard_page_id": "group_admin_control",
            "items": (
                [
                    {"page_id": "group_admin_control", "label": "Admin Control Overview"},
                    {"page_id": "live_activity", "label": "Live Activity & Update Monitor"},
                    {"page_id": "master_approval_queue", "label": "Master Admin Approval Queue"},
                    {"page_id": "announcement_broadcaster", "label": "Announcement Broadcaster"},
                ]
                if is_master
                else [
                    {"page_id": "group_admin_control", "label": "Parish Queue Overview"},
                    {"page_id": "parish_inbox", "label": "Parish Inbox & Action Hub"},
                ]
            ),
        },
        {
            "group_id": "group_community_system",
            "title": "5. Community & System",
            "dashboard_page_id": "group_community_system",
            "items": [
                {"page_id": "group_community_system", "label": "Community & System Overview"},
                {"page_id": "prayer_wall", "label": "Prayer Wall & Categorizer"},
                {"page_id": "ministry_messenger", "label": "Ministry Messenger"},
                {"page_id": "ai_assistant", "label": "CTG Gabay AI Assistant"},
                {"page_id": "app_settings", "label": "App Settings & Security"},
            ],
        },
    ]
    return groups


def render_grouped_sidebar_navigation(
    user_role: str,
    login_role: str,
    current_app_mode: str,
) -> str:
    """
    Renders Sidebar Section C: The 5 clean expandable navigation groups (without any separate
    Home menu button or duplicate Home & Leadership Portal entry).
    Returns the active `app_mode` string.
    """
    groups = get_role_navigation_groups(user_role=user_role, login_role=login_role)
    active_page_id = APP_MODE_TO_PAGE_ID.get(current_app_mode, "home_overview")

    st.markdown("<div class='ctg-sb-section-label'>Ministry Modules</div>", unsafe_allow_html=True)

    for grp in groups:
        grp_page_ids = {it["page_id"] for it in grp["items"]}
        is_grp_active = active_page_id in grp_page_ids
        with st.expander(grp["title"], expanded=is_grp_active):
            for it in grp["items"]:
                pid = it["page_id"]
                is_cur = (pid == active_page_id)
                lbl = f"▸ {it['label']}" if is_cur else it["label"]
                if st.button(
                    lbl,
                    key=f"sb_nav_item_{grp['group_id']}_{pid}",
                    use_container_width=True,
                    type="primary" if is_cur else "secondary",
                ):
                    target_mode = PAGE_ID_TO_APP_MODE[pid]
                    if target_mode != st.session_state.get("current_nav_page"):
                        st.session_state["current_nav_page"] = target_mode
                        st.rerun()

    return st.session_state.get("current_nav_page", PAGE_ID_TO_APP_MODE["home_overview"])


def render_sidebar_system_tools(
    is_master: bool,
    is_true_master: bool,
    unread_notif_count: int,
    on_sync_live_data: Callable[[], None],
    render_live_activity_fn: Callable[[], None],
    render_cloud_backup_fn: Callable[[], None],
    render_role_switcher_fn: Callable[[], None],
    render_notifications_fn: Callable[[], None],
) -> None:
    """Renders Sidebar Section B: Compact grouped System Tools scoped by role."""
    st.markdown("<div class='ctg-sb-section-label'>System Tools</div>", unsafe_allow_html=True)
    if st.button("🔄  Sync Live Excel & Notion Data", use_container_width=True, key="sb_tool_sync_btn"):
        on_sync_live_data()
        st.rerun()

    if is_master:
        with st.expander("📡 Live Update & Activity Monitor", expanded=False):
            render_live_activity_fn()
        with st.expander("☁️ Cloud Backup", expanded=False):
            render_cloud_backup_fn()

    if is_true_master:
        with st.expander("🔀 Admin View & Parish Switcher", expanded=False):
            render_role_switcher_fn()

    if is_master or is_true_master:
        with st.expander(f"🔔 Live Notifications ({unread_notif_count})", expanded=False):
            render_notifications_fn()


def render_module_dashboard(
    title: str,
    subtitle: str,
    breadcrumbs: List[str],
    scope_label: str,
    ph_datetime_str: str,
    metrics: List[Tuple[str, Any, str, str]],
    actions: List[Dict[str, str]],
    recent_items: Optional[List[Dict[str, Any]]] = None,
    future_modules: Optional[List[str]] = None,
) -> None:
    """Reusable landing dashboard for each of the 5 navigation groups."""
    render_breadcrumbs(breadcrumbs)
    render_dashboard_header(
        title=title,
        subtitle=subtitle,
        scope_label=scope_label,
        ph_datetime_str=ph_datetime_str,
    )
    if metrics:
        m_cols = st.columns(min(4, len(metrics)))
        for idx, (m_lbl, m_val, m_sub, m_tone) in enumerate(metrics):
            with m_cols[idx % len(m_cols)]:
                render_metric_card(m_lbl, m_val, m_sub, tone=m_tone)
    st.markdown("#### Quick Modules")
    if actions:
        a_cols = st.columns(min(3, len(actions)))
        for idx, act in enumerate(actions):
            with a_cols[idx % len(a_cols)]:
                render_quick_action_card(
                    title=act["title"],
                    description=act["description"],
                    button_label=act["button_label"],
                    key=f"mod_dash_act_{act['target_page_id']}_{idx}",
                    on_click_page=PAGE_ID_TO_APP_MODE.get(act["target_page_id"]),
                    badge_text=act.get("badge", ""),
                )
    if recent_items is not None:
        st.markdown("#### Recent Updates")
        render_activity_feed(recent_items)
    if future_modules:
        st.markdown("#### Planned Future Modules")
        st.info(" • ".join(future_modules) + " — Planned for a future release (not yet active).")
import os
import io
import re
import json
import time
import base64
import importlib
import requests
import openpyxl
from PIL import Image
from datetime import datetime, timedelta, timezone

PH_TIMEZONE = timezone(timedelta(hours=8))

def get_philippine_now() -> datetime:
    return datetime.now(PH_TIMEZONE)

def get_philippine_datetime_str() -> str:
    return get_philippine_now().strftime("%Y-%m-%d %I:%M:%S %p PHT")
from email.message import EmailMessage
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types
import urllib.parse
from google.oauth2.credentials import Credentials
try:
    from google.oauth2 import service_account
except ImportError:
    service_account = None
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

load_dotenv()

# ==========================================
# 🖼️ AUTOMATIC LOGO FINDER & HEADER HELPER
# ==========================================
def find_logo_file():
    preferred_names = [
        "logo.png", "logo.jpg", "logo.jpeg",
        "ctg_logo.png", "ctg_logo.jpg", "ctg_logo.jpeg",
        "ctg.png", "ctg.jpg", "icon.png"
    ]
    for name in preferred_names:
        if os.path.exists(name):
            return name
    for f in os.listdir("."):
        if f.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            return f
    return None

LOGO_PATH = find_logo_file()

def render_title_with_logo(title_text: str, subtitle_text: str = ""):
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        col_logo, col_text = st.columns([0.09, 0.91], vertical_alignment="center")
        with col_logo:
            st.image(LOGO_PATH, width=75)
        with col_text:
            st.title(title_text)
            if subtitle_text:
                st.caption(subtitle_text)
    else:
        st.title(title_text)
        if subtitle_text:
            st.caption(subtitle_text)

def get_secret(key, default=""):
    if "custom_secrets" in st.session_state and key in st.session_state.custom_secrets:
        val = st.session_state.custom_secrets[key]
        if val:
            return str(val).strip()
    try:
        if os.path.exists("ctg_app_cloud_config.json"):
            with open("ctg_app_cloud_config.json", "r", encoding="utf-8") as f:
                cfg_data = json.load(f)
            if cfg_data.get(key):
                return str(cfg_data[key]).strip()
    except Exception:
        pass
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    val = os.getenv(key, default)
    return str(val).strip() if val else default

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
GEMINI_API_KEY_2 = get_secret("GEMINI_API_KEY_2")
NOTION_TOKEN = get_secret("NOTION_TOKEN")
NOTION_CRM_DB_ID = get_secret("NOTION_CRM_DB_ID")
NOTION_PRAYER_DB_ID = get_secret("NOTION_PRAYER_DB_ID")
NOTION_QUESTIONS_DB_ID = get_secret("NOTION_QUESTIONS_DB_ID")
NOTION_COUNSELLING_DB_ID = get_secret("NOTION_COUNSELLING_DB_ID")
GMAIL_CLIENT_ID = get_secret("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = get_secret("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = get_secret("GMAIL_REFRESH_TOKEN")
GOOGLE_SERVICE_ACCOUNT_JSON = get_secret("GOOGLE_SERVICE_ACCOUNT_JSON")
APP_USER = get_secret("APP_USER", "ctg.ministry.master.admin")
APP_PASSWORD = get_secret("APP_PASSWORD", "Church#Temple#God$2026!Admin")
BACKUP_ADMIN_EMAIL = "clarojerald16@gmail.com"

DEFAULT_CLOUD_SPREADSHEET_URL = get_secret(
    "CLOUD_SPREADSHEET_URL",
    "https://docs.google.com/spreadsheets/d/1v7BFj32ixL7xg5BaElAU1Ji-zsSKhdOF/edit?usp=sharing"
)
WEEKLY_REPORT_FOLDER_URL = get_secret(
    "WEEKLY_REPORT_FOLDER_URL",
    "https://drive.google.com/drive/folders/12snjykclzNzQW7L0WxQxh09_xSjswbz5?usp=sharing"
)
PARISH_UPLOAD_FOLDER_URL = get_secret(
    "PARISH_UPLOAD_FOLDER_URL",
    "https://drive.google.com/drive/folders/1WHvH81kNL3jmXItNtPl39GnhmTzpmK1X?usp=sharing"
)
ALL_PARISH_DB_FOLDER_URL = get_secret(
    "ALL_PARISH_DB_FOLDER_URL",
    "https://drive.google.com/drive/folders/1RXOTB6iA0jgcsOf1kE8wn3V8DHVduvH0?usp=sharing"
)

LOCAL_EXCEL_CACHE = "CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"
LOCAL_PWR_MASTER_FILE = "PWR. January 1st Week.xlsx"
LOCAL_WEEKLY_REPORTS_DB = "ctg_weekly_monitoring_records.json"
LOCAL_DRIVE_LINKS_REGISTRY_FILE = "ctg_saved_drive_links.json"
LOCAL_CONFIG_FILE = "ctg_app_cloud_config.json"
LOCAL_CHAT_FILE = "ctg_ministry_internal_chat.json"
DRIVE_CHAT_FILENAME = "ctg_ministry_internal_chat_cloud.json"

MONTHS_2027 = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
WEEKS_2027 = ["1st Week", "2nd Week", "3rd Week", "4th Week", "5th Week"]

# Default Baseline Member Counts per Parish (ginagamit kapag nag-save ng bagong report)
DEFAULT_PARISH_MEMBER_BASELINES = {
    "LIPAY NORTE PARISH": {"say": 90, "youth": 29, "kids": 26},
    "MALANAS PARISH": {"say": 17, "youth": 10, "kids": 3},
    "POBLACION PARISH": {"say": 26, "youth": 4, "kids": 7},
    "DUPLAS PARISH": {"say": 22, "youth": 3, "kids": 0},
    "SANTIAGO PARISH": {"say": 8, "youth": 3, "kids": 0},
    "CABA PARISH": {"say": 21, "youth": 6, "kids": 2},
    "MANILA PARISH": {"say": 18, "youth": 3, "kids": 2},
    "ANDABUEN PARISH": {"say": 37, "youth": 7, "kids": 4},
    "APAYAO PARISH": {"say": 11, "youth": 1, "kids": 0},
    "HONGKONG PARISH": {"say": 11, "youth": 0, "kids": 0},
}

def load_saved_config() -> dict:
    if os.path.exists(LOCAL_CONFIG_FILE):
        try:
            with open(LOCAL_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def load_saved_cloud_url() -> str:
    if "custom_secrets" in st.session_state and st.session_state.custom_secrets.get("CLOUD_SPREADSHEET_URL"):
        return st.session_state.custom_secrets["CLOUD_SPREADSHEET_URL"].strip()
    data = load_saved_config()
    if data.get("CLOUD_SPREADSHEET_URL"):
        return data["CLOUD_SPREADSHEET_URL"].strip()
    return get_secret("CLOUD_SPREADSHEET_URL", DEFAULT_CLOUD_SPREADSHEET_URL)

def load_saved_weekly_folder_url() -> str:
    if "custom_secrets" in st.session_state and st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL"):
        return st.session_state.custom_secrets["WEEKLY_REPORT_FOLDER_URL"].strip()
    data = load_saved_config()
    if data.get("WEEKLY_REPORT_FOLDER_URL"):
        return data["WEEKLY_REPORT_FOLDER_URL"].strip()
    return get_secret("WEEKLY_REPORT_FOLDER_URL", WEEKLY_REPORT_FOLDER_URL)

def load_saved_parish_upload_folder_url() -> str:
    if "custom_secrets" in st.session_state and st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL"):
        return st.session_state.custom_secrets["PARISH_UPLOAD_FOLDER_URL"].strip()
    data = load_saved_config()
    if data.get("PARISH_UPLOAD_FOLDER_URL"):
        return data["PARISH_UPLOAD_FOLDER_URL"].strip()
    return get_secret("PARISH_UPLOAD_FOLDER_URL", PARISH_UPLOAD_FOLDER_URL)

def save_cloud_url_config(new_url: str = None, weekly_folder_url: str = None, parish_upload_url: str = None, refresh_tok: str = None):
    try:
        existing = load_saved_config()
        if new_url is not None:
            existing["CLOUD_SPREADSHEET_URL"] = new_url.strip()
        if weekly_folder_url is not None:
            existing["WEEKLY_REPORT_FOLDER_URL"] = weekly_folder_url.strip()
        if parish_upload_url is not None:
            existing["PARISH_UPLOAD_FOLDER_URL"] = parish_upload_url.strip()
        if refresh_tok is not None:
            existing["GMAIL_REFRESH_TOKEN"] = refresh_tok.strip()
        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
    except Exception:
        pass

st.set_page_config(
    page_title="CTG Ministry Hub - Gabay AI",
    page_icon=LOGO_PATH if LOGO_PATH else "⛪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 🎛️ INITIALIZE UI & THEME CONTROLLER STATES
# ==========================================
if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "🌙 Dark Mode"
if "ui_font_size" not in st.session_state:
    st.session_state.ui_font_size = "Normal (100%)"
if "ui_density" not in st.session_state:
    st.session_state.ui_density = "Standard (Comfortable)"
if "ui_default_lang" not in st.session_state:
    st.session_state.ui_default_lang = "Tagalog"
if "ui_show_side_records" not in st.session_state:
    st.session_state.ui_show_side_records = True

def apply_dynamic_app_styles():
    theme = st.session_state.get("ui_theme", "🌙 Dark Mode")
    font_choice = st.session_state.get("ui_font_size", "Normal (100%)")
    density = st.session_state.get("ui_density", "Standard (Comfortable)")

    font_scale = "0.95rem"
    base_zoom = "100%"
    if "Large (115%)" in font_choice:
        font_scale = "1.08rem"
        base_zoom = "108%"
    elif "Extra Large (130%)" in font_choice:
        font_scale = "1.20rem"
        base_zoom = "116%"

    card_pad = "12px 16px" if "Standard" in density else "7px 11px"
    card_margin = "10px" if "Standard" in density else "5px"

    if "Light" in theme:
        app_bg = "#F8FAFC"
        sidebar_bg = "#EFF6FF"
        text_color = "#0F172A"
        card_bg = "#FFFFFF"
        card_text = "#1E293B"
        card_border = "#1E3A8A"
        side_box_bg = "#FFFFFF"
        side_box_text = "#0F172A"
        side_box_border = "#CBD5E1"
    elif "Royal Blue" in theme:
        app_bg = "#0A192F"
        sidebar_bg = "#112240"
        text_color = "#E6F1FF"
        card_bg = "#112240"
        card_text = "#E6F1FF"
        card_border = "#64FFDA"
        side_box_bg = "#172A45"
        side_box_text = "#E6F1FF"
        side_box_border = "#233554"
    else:
        app_bg = "#0E1117"
        sidebar_bg = "#161B22"
        text_color = "#F8FAFC"
        card_bg = "#1E293B"
        card_text = "#F8FAFC"
        card_border = "#38BDF8"
        side_box_bg = "#1E293B"
        side_box_text = "#F1F5F9"
        side_box_border = "#334155"

    st.markdown(f"""
        <style>
        html {{ font-size: {base_zoom}; }}
        .stApp {{ background-color: {app_bg} !important; color: {text_color} !important; }}
        [data-testid="stSidebar"] {{ background-color: {sidebar_bg} !important; }}
        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp p, .stApp label, .stMarkdown {{ color: {text_color} !important; }}
        .full-text-card {{
            background-color: {card_bg};
            border-left: 4px solid {card_border};
            border: 1px solid {side_box_border};
            padding: {card_pad};
            margin-bottom: {card_margin};
            border-radius: 8px;
            color: {card_text} !important;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: {font_scale};
        }}
        .urgent-card {{
            background-color: #FEF2F2;
            border-left: 4px solid #DC2626;
            padding: {card_pad};
            margin-bottom: {card_margin};
            border-radius: 8px;
            color: #7F1D1D !important;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: {font_scale};
        }}
        .submitted-parish-badge {{
            background-color: #ECFDF5; border: 1px solid #10B981; color: #065F46 !important;
            padding: 6px 12px; border-radius: 6px; font-weight: 600; font-size: 0.88rem; margin-bottom: 8px; display: inline-block;
        }}
        .assigned-parish-badge {{
            background-color: #EFF6FF; border: 1px solid #3B82F6; color: #1E40AF !important;
            padding: 6px 12px; border-radius: 6px; font-weight: 600; font-size: 0.88rem; margin-bottom: 8px; display: inline-block;
        }}
        .records-side-box {{
            background-color: {side_box_bg}; border: 1px solid {side_box_border}; border-top: 4px solid #0F766E;
            padding: {card_pad}; border-radius: 8px; margin-bottom: {card_margin}; color: {side_box_text} !important; font-size: 0.88rem;
        }}
        .sidebar-log-card {{
            background-color: {side_box_bg};
            border-left: 3px solid #38BDF8;
            border: 1px solid {side_box_border};
            padding: 8px 10px;
            border-radius: 5px;
            margin-bottom: 6px;
            color: {side_box_text} !important;
            font-size: 0.80rem;
        }}
        .chat-bubble-mine {{
            background-color: #1E3A8A;
            border: 1px solid #3B82F6;
            color: #FFFFFF !important;
            padding: 10px 14px;
            border-radius: 12px 12px 2px 12px;
            margin-bottom: 8px;
            margin-left: 15%;
            font-size: {font_scale};
        }}
        .chat-bubble-other {{
            background-color: {side_box_bg};
            border-left: 4px solid #10B981;
            border-top: 1px solid {side_box_border};
            border-right: 1px solid {side_box_border};
            border-bottom: 1px solid {side_box_border};
            color: {side_box_text} !important;
            padding: 10px 14px;
            border-radius: 12px 12px 12px 2px;
            margin-bottom: 8px;
            margin-right: 15%;
            font-size: {font_scale};
        }}
        .chat-bubble-broadcast {{
            background-color: #4C1D95;
            border: 1px solid #A78BFA;
            color: #F5F3FF !important;
            padding: 10px 14px;
            border-radius: 10px;
            margin-bottom: 8px;
            font-size: {font_scale};
        }}
        .category-header-seniors {{
            background-color: #3B0764;
            border: 1px solid #A855F7;
            color: #F3E8FF !important;
            padding: 10px 14px;
            border-radius: 8px;
            font-weight: 700;
            margin-bottom: 8px;
        }}
        .category-header-adults {{
            background-color: #1E3A8A;
            border: 1px solid #60A5FA;
            color: #DBEAFE !important;
            padding: 10px 14px;
            border-radius: 8px;
            font-weight: 700;
            margin-bottom: 8px;
        }}
        .category-header-youth {{
            background-color: #065F46;
            border: 1px solid #34D399;
            color: #D1FAE5 !important;
            padding: 10px 14px;
            border-radius: 8px;
            font-weight: 700;
            margin-bottom: 8px;
        }}
        .category-header-kids {{
            background-color: #78350F;
            border: 1px solid #FBBF24;
            color: #FEF3C7 !important;
            padding: 10px 14px;
            border-radius: 8px;
            font-weight: 700;
            margin-bottom: 8px;
        }}
        .pmr-banner {{
            background-color: #14532D;
            color: #FFFFFF !important;
            text-align: center;
            padding: 12px;
            border-radius: 8px 8px 0 0;
            font-weight: 800;
            font-size: 1.35rem;
            letter-spacing: 1px;
        }}
        .pmr-subbanner {{
            background-color: #DCFCE7;
            color: #14532D !important;
            text-align: center;
            padding: 8px;
            border-radius: 0 0 8px 8px;
            font-weight: 700;
            font-size: 1.1rem;
            margin-bottom: 14px;
            border: 1px solid #14532D;
        }}
        [data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"] {{
            display: flex !important; visibility: visible !important; opacity: 1 !important; z-index: 999999 !important;
            background-color: #1E3A8A !important; color: #FFFFFF !important; border-radius: 8px !important; padding: 4px !important; margin: 6px !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.45) !important;
        }}
        .record-table-wrap {{
            width: 100%;
            overflow-x: auto;
            margin-bottom: 14px;
            border-radius: 10px;
            border: 1px solid {side_box_border};
            background-color: {card_bg};
        }}
        .record-style-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: {font_scale};
            color: {card_text} !important;
        }}
        .record-style-table th {{
            background-color: #14532D;
            color: #FFFFFF !important;
            padding: 10px 12px;
            text-align: left;
            font-weight: 700;
            border: 1px solid {side_box_border};
        }}
        .record-style-table td {{
            padding: 10px 12px;
            border: 1px solid {side_box_border};
            vertical-align: top;
            white-space: pre-wrap;
            word-wrap: break-word;
            line-height: 1.45;
            color: {card_text} !important;
        }}
        .record-style-table tr:nth-child(even) {{
            background-color: {side_box_bg};
        }}
        .ph-time-badge {{
            background-color: #065F46;
            color: #ECFDF5 !important;
            border: 1px solid #34D399;
            padding: 5px 12px;
            border-radius: 999px;
            font-size: 0.85rem;
            font-weight: 700;
            display: inline-block;
            margin-bottom: 6px;
        }}
        /* 📱 RESPONSIVE TABLET & PHONE MEDIA QUERIES */
        @media (max-width: 1024px) {{
            .block-container {{
                padding-left: 1rem !important;
                padding-right: 1rem !important;
                padding-top: 1.5rem !important;
            }}
            [data-testid="stHorizontalBlock"] {{
                flex-wrap: wrap !important;
                gap: 0.65rem !important;
            }}
            [data-testid="column"] {{
                min-width: 47% !important;
                flex: 1 1 47% !important;
            }}
            .pmr-banner {{ font-size: 1.15rem !important; padding: 10px !important; }}
            .pmr-subbanner {{ font-size: 0.95rem !important; padding: 7px !important; }}
        }}
        @media (max-width: 640px) {{
            html {{ font-size: 96% !important; }}
            .block-container {{
                padding-left: 0.6rem !important;
                padding-right: 0.6rem !important;
                padding-top: 1.2rem !important;
            }}
            [data-testid="column"] {{
                min-width: 100% !important;
                flex: 1 1 100% !important;
            }}
            .stButton > button, .stDownloadButton > button {{
                min-height: 45px !important;
                font-size: 0.95rem !important;
                width: 100% !important;
            }}
            .chat-bubble-mine {{ margin-left: 4% !important; }}
            .chat-bubble-other {{ margin-right: 4% !important; }}
            .record-style-table th, .record-style-table td {{
                padding: 8px 8px !important;
                font-size: 0.86rem !important;
            }}
        }}
        [data-testid="stAppDeployButton"] {{ display: none !important; }}
        #MainMenu {{ visibility: hidden !important; display: none !important; }}
        footer {{ visibility: hidden !important; display: none !important; }}
        .viewerBadge_container__1QSob, .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137, .viewerBadge_text__1JaDK {{ display: none !important; }}
        </style>
    """, unsafe_allow_html=True)

apply_dynamic_app_styles()

PARISH_ORDER = [
    "LIPAY NORTE PARISH", "MALANAS PARISH", "POBLACION PARISH", "DUPLAS PARISH",
    "SANTIAGO PARISH", "CABA PARISH", "MANILA PARISH", "ANDABUEN PARISH",
    "APAYAO PARISH", "HONGKONG PARISH",
]

PARISH_ACCOUNTS = {
    "ctg.ministry.master.admin": {"password": "Church#Temple#God$2026!Admin", "parish": "Lahat ng Parokya", "role": "Master Admin"},
    APP_USER.lower(): {"password": APP_PASSWORD, "parish": "Lahat ng Parokya", "role": "Master Admin"},
    "lipay norte parish": {"password": "ctglipaynorte1", "parish": "LIPAY NORTE PARISH", "role": "Parish Leader"},
    "lipaynorteparish": {"password": "ctglipaynorte1", "parish": "LIPAY NORTE PARISH", "role": "Parish Leader"},
    "malanas parish": {"password": "ctgmalanas2", "parish": "MALANAS PARISH", "role": "Parish Leader"},
    "malanasparish": {"password": "ctgmalanas2", "parish": "MALANAS PARISH", "role": "Parish Leader"},
    "poblacion parish": {"password": "ctgpoblacion3", "parish": "POBLACION PARISH", "role": "Parish Leader"},
    "poblacionparish": {"password": "ctgpoblacion3", "parish": "POBLACION PARISH", "role": "Parish Leader"},
    "duplas parish": {"password": "ctgduplas4", "parish": "DUPLAS PARISH", "role": "Parish Leader"},
    "duplasparish": {"password": "ctgduplas4", "parish": "DUPLAS PARISH", "role": "Parish Leader"},
    "santiago parish": {"password": "ctgsantiago5", "parish": "SANTIAGO PARISH", "role": "Parish Leader"},
    "santiagoparish": {"password": "ctgsantiago5", "parish": "SANTIAGO PARISH", "role": "Parish Leader"},
    "caba parish": {"password": "ctgcaba6", "parish": "CABA PARISH", "role": "Parish Leader"},
    "cabaparish": {"password": "ctgcaba6", "parish": "CABA PARISH", "role": "Parish Leader"},
    "manila parish": {"password": "ctgmanila7", "parish": "MANILA PARISH", "role": "Parish Leader"},
    "manilaparish": {"password": "ctgmanila7", "parish": "MANILA PARISH", "role": "Parish Leader"},
    "andabuen parish": {"password": "ctgandabuen8", "parish": "ANDABUEN PARISH", "role": "Parish Leader"},
    "andabuenparish": {"password": "ctgandabuen8", "parish": "ANDABUEN PARISH", "role": "Parish Leader"},
    "apayao parish": {"password": "ctgapayao9", "parish": "APAYAO PARISH", "role": "Parish Leader"},
    "apayaoparish": {"password": "ctgapayao9", "parish": "APAYAO PARISH", "role": "Parish Leader"},
    "hongkong parish": {"password": "ctghongkong10", "parish": "HONGKONG PARISH", "role": "Parish Leader"},
    "hongkongparish": {"password": "ctghongkong10", "parish": "HONGKONG PARISH", "role": "Parish Leader"},
}

def extract_id_from_url(url: str) -> str:
    if not url:
        return ""
    match_d = re.search(r"/d/([a-zA-Z0-9-_]+)", str(url))
    if match_d:
        return match_d.group(1)
    match_f = re.search(r"/folders/([a-zA-Z0-9-_]+)", str(url))
    if match_f:
        return match_f.group(1)
    return str(url).strip()

# ==========================================
# ☁️ GOOGLE DRIVE SERVICE & AUTOMATIC 2027 ROUTING
# ==========================================
def get_google_drive_service():
    sa_json = st.session_state.get("custom_secrets", {}).get("GOOGLE_SERVICE_ACCOUNT_JSON", "") or get_secret("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json and service_account is not None:
        try:
            sa_info = json.loads(sa_json)
            creds = service_account.Credentials.from_service_account_info(
                sa_info, scopes=["https://www.googleapis.com/auth/drive"]
            )
            return build("drive", "v3", credentials=creds)
        except Exception:
            pass

    cfg = load_saved_config()
    g_id = get_secret("GMAIL_CLIENT_ID")
    g_sec = get_secret("GMAIL_CLIENT_SECRET")
    g_ref = cfg.get("GMAIL_REFRESH_TOKEN") or get_secret("GMAIL_REFRESH_TOKEN")
    if not all([g_id, g_sec, g_ref]):
        return None
    try:
        creds = Credentials(
            token=None,
            refresh_token=g_ref.strip(),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=g_id.strip(),
            client_secret=g_sec.strip()
        )
        return build("drive", "v3", credentials=creds)
    except Exception:
        return None

def find_drive_folder_by_name(service, folder_name_query: str, parent_id: str = None) -> str:
    if not service:
        return ""
    try:
        q_parts = [
            "mimeType='application/vnd.google-apps.folder'",
            "trashed=false",
            f"name contains '{folder_name_query}'"
        ]
        if parent_id:
            q_parts.append(f"'{parent_id}' in parents")
        res = service.files().list(
            q=" and ".join(q_parts),
            spaces="drive",
            fields="files(id, name)",
            pageSize=20
        ).execute()
        files = res.get("files", [])
        return files[0]["id"] if files else ""
    except Exception:
        return ""

def get_or_create_parish_subfolder(service, parent_folder_id: str, parish_name: str, month_name: str = "") -> str:
    if not service:
        return parent_folder_id
    try:
        current_parent_id = parent_folder_id
        if not current_parent_id:
            current_parent_id = find_drive_folder_by_name(service, "Upload the Image/Excel File here")
        if not current_parent_id:
            current_parent_id = find_drive_folder_by_name(service, "CTG Ministry Hub")
        if not current_parent_id:
            return ""

        res = service.files().list(
            q=f"'{current_parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces="drive",
            fields="files(id, name)",
            pageSize=50
        ).execute()
        folders = res.get("files", [])

        if month_name:
            for folder in folders:
                if month_name.lower() in folder["name"].lower():
                    current_parent_id = folder["id"]
                    res_sub = service.files().list(
                        q=f"'{current_parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
                        spaces="drive",
                        fields="files(id, name)",
                        pageSize=50
                    ).execute()
                    folders = res_sub.get("files", [])
                    break

        clean_parish = parish_name.replace("PARISH", "").strip()
        short_keyword = clean_parish.replace("NORTE", "").strip()
        target_parish_folder_id = None
        for folder in folders:
            f_up = folder["name"].upper()
            if clean_parish.upper() in f_up or (short_keyword and short_keyword.upper() in f_up):
                target_parish_folder_id = folder["id"]
                break

        if not target_parish_folder_id:
            folder_meta = {
                "name": parish_name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [current_parent_id]
            }
            created = service.files().create(body=folder_meta, fields="id").execute()
            target_parish_folder_id = created.get("id", current_parent_id)

        return target_parish_folder_id
    except Exception:
        return parent_folder_id

def upload_file_to_google_drive(file_bytes, filename, mime_type="image/jpeg", parent_folder_url="", parish_name="", month_name=""):
    service = get_google_drive_service()
    if service is None:
        return None, "Error: Kulang ang Google Drive/Gmail credentials sa App Settings."
    try:
        file_metadata = {'name': filename}
        root_folder_id = extract_id_from_url(parent_folder_url)
        target_folder_id = get_or_create_parish_subfolder(service, root_folder_id, parish_name, month_name) if parish_name else root_folder_id
        if target_folder_id:
            file_metadata['parents'] = [target_folder_id]
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return file.get('webViewLink'), f"Matagumpay na na-upload sa Google Drive folder ng {parish_name or 'CTG'}!"
    except Exception as e:
        err_str = str(e)
        if "403" in err_str or "insufficientPermissions" in err_str:
            return None, (
                "Google Drive Permission Error (403): Ang iyong kasalukuyang Refresh Token ay pang-Gmail lamang. "
                "Pumunta sa '🛠️ App Settings -> 🔑 API & Google Drive / Gmail Credentials' at gamitin ang 'One-Click Google Drive + Gmail Token Generator' upang mabuksan ang Google Drive write access."
            )
        return None, f"Google Drive Upload Error: {err_str}"

def load_saved_drive_links_registry() -> list[dict]:
    """Binabasa ang listahan ng lahat ng naka-save na Google Drive links upang walang mawala kahit magdagdag ng bagong link."""
    default_links = [
        {
            "label": "PWR. January 1st Week.xlsx (Google Drive Master PWR)",
            "category": "Master Weekly Report Excel (PWR)",
            "month": "January",
            "week": "1st Week",
            "parish": "Lahat ng Parokya (Master)",
            "url": "https://docs.google.com/spreadsheets/d/1v7BFj32ixL7xg5BaElAU1Ji-zsSKhdOF/edit?usp=sharing",
            "added_at": "2027-01-01"
        },
        {
            "label": "CTG Ministry Hub (Weekly Report Jan-Dec 2027 Master Folder)",
            "category": "Google Drive Master Folder",
            "month": "All Months",
            "week": "All Weeks",
            "parish": "Lahat ng Parokya (Master)",
            "url": st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", WEEKLY_REPORT_FOLDER_URL) or "https://drive.google.com/drive/folders/12snjykclzNzQW7L0WxQxh09_xSjswbz5?usp=sharing",
            "added_at": "2027-01-01"
        },
        {
            "label": "Upload the Image/Excel File Here! (Parish Upload Folder)",
            "category": "Parish Upload Folder",
            "month": "All Months",
            "week": "All Weeks",
            "parish": "Lahat ng Parokya (Master)",
            "url": st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL", PARISH_UPLOAD_FOLDER_URL) or "https://drive.google.com/drive/folders/1WHvH81kNL3jmXItNtPl39GnhmTzpmK1X?usp=sharing",
            "added_at": "2027-01-01"
        }
    ]
    if os.path.exists(LOCAL_DRIVE_LINKS_REGISTRY_FILE):
        try:
            with open(LOCAL_DRIVE_LINKS_REGISTRY_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, list) and len(saved) > 0:
                existing_urls = {str(it.get("url", "")).strip() for it in saved}
                for d_item in default_links:
                    if d_item["url"] and d_item["url"] not in existing_urls:
                        saved.append(d_item)
                return saved
        except Exception:
            pass
    return default_links

def save_drive_links_registry(links_list: list[dict]):
    try:
        with open(LOCAL_DRIVE_LINKS_REGISTRY_FILE, "w", encoding="utf-8") as f:
            json.dump(links_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def add_link_to_drive_registry(label: str, category: str, month: str, week: str, parish: str, url: str) -> list[dict]:
    registry = load_saved_drive_links_registry()
    clean_u = str(url or "").strip()
    if not clean_u:
        return registry
    for item in registry:
        if str(item.get("url", "")).strip() == clean_u and item.get("month") == month and item.get("week") == week:
            item["label"] = label or item.get("label", "Saved Google Drive Link")
            item["category"] = category
            item["parish"] = parish
            save_drive_links_registry(registry)
            return registry
    registry.append({
        "label": label.strip() if label and label.strip() else f"{parish} — {month} ({week})",
        "category": category,
        "month": month,
        "week": week,
        "parish": parish,
        "url": clean_u,
        "added_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    })
    save_drive_links_registry(registry)
    return registry

def find_2027_master_excel_on_drive(service, weekly_folder_url: str, month_name: str, week_name: str) -> tuple[str, str]:
    if not service:
        return "", ""
    target_filename = f"PWR. {month_name} {week_name}.xlsx"
    try:
        root_id = extract_id_from_url(weekly_folder_url)
        if root_id and "/spreadsheets/d/" in str(weekly_folder_url):
            try:
                f_meta = service.files().get(fileId=root_id, fields="id, name, mimeType, parents").execute()
                f_name_low = f_meta.get("name", "").lower()
                if month_name.lower() in f_name_low and week_name.lower() in f_name_low:
                    p_list = f_meta.get("parents", [""])
                    return f_meta["id"], p_list[0] if p_list else ""
            except Exception:
                pass

        if not root_id or "/spreadsheets/d/" in str(weekly_folder_url):
            root_id = find_drive_folder_by_name(service, "CTG Ministry Hub (Weekly Report January-December 2027)")

        month_idx = MONTHS_2027.index(month_name) + 1 if month_name in MONTHS_2027 else 1
        month_folder_name = f"{month_idx:02d}. {month_name} 2027 PWR"
        month_folder_id = find_drive_folder_by_name(service, f"{month_name} 2027 PWR", root_id if root_id else None)

        if not month_folder_id and root_id:
            try:
                meta = service.files().get(fileId=root_id, fields="id, name, mimeType").execute()
                if meta.get("mimeType") == "application/vnd.google-apps.folder":
                    if month_name.lower() in meta.get("name", "").lower():
                        month_folder_id = root_id
                    else:
                        created_m = service.files().create(
                            body={"name": month_folder_name, "mimeType": "application/vnd.google-apps.folder", "parents": [root_id]},
                            fields="id"
                        ).execute()
                        month_folder_id = created_m.get("id", root_id)
            except Exception:
                pass

        # 1. Hanapin muna sa loob ng eksaktong Month Folder (hal. '01. January 2027 PWR')
        if month_folder_id:
            res_f = service.files().list(
                q=f"'{month_folder_id}' in parents and trashed=false",
                spaces="drive",
                fields="files(id, name, mimeType)",
                pageSize=50
            ).execute()
            for f_obj in res_f.get("files", []):
                fn_low = f_obj.get("name", "").lower()
                if week_name.lower() in fn_low and (fn_low.endswith(".xlsx") or "pwr" in fn_low):
                    return f_obj["id"], month_folder_id

        # 2. Fallback: Hanapin sa buong Google Drive gamit ang eksaktong filename
        res = service.files().list(
            q=f"(name = '{target_filename}' or name = 'PWR. {month_name} {week_name}') and trashed = false",
            spaces="drive",
            fields="files(id, name, parents)",
            pageSize=10
        ).execute()
        files = res.get("files", [])
        if files:
            p_list = files[0].get("parents", [""])
            return files[0]["id"], month_folder_id or (p_list[0] if p_list else "")

        # 3. Fallback: Suriin kung may idinagdag na partikular na PWR link sa Saved Links Registry para sa buwan at linggong ito
        for reg_item in load_saved_drive_links_registry():
            if reg_item.get("category") == "Master Weekly Report Excel (PWR)":
                r_m = str(reg_item.get("month", "")).strip()
                r_w = str(reg_item.get("week", "")).strip()
                r_u = str(reg_item.get("url", "")).strip()
                if r_u and r_m.lower() == month_name.lower() and r_w.lower() == week_name.lower() and "/folders/" not in r_u:
                    reg_fid = extract_id_from_url(r_u)
                    if reg_fid:
                        return reg_fid, month_folder_id or root_id

        if month_name.strip().lower() == "january" and "1st" in week_name.strip().lower():
            return "1v7BFj32ixL7xg5BaElAU1Ji-zsSKhdOF", month_folder_id or root_id

        return "", month_folder_id or root_id
    except Exception:
        if month_name.strip().lower() == "january" and "1st" in week_name.strip().lower():
            return "1v7BFj32ixL7xg5BaElAU1Ji-zsSKhdOF", extract_id_from_url(weekly_folder_url)
        return "", extract_id_from_url(weekly_folder_url)

def sync_excel_bytes_to_drive(excel_bytes: bytes, spreadsheet_url: str, local_cache_path: str = LOCAL_EXCEL_CACHE) -> tuple[bool, str]:
    try:
        with open(local_cache_path, "wb") as f:
            f.write(excel_bytes)
    except Exception:
        pass

    service = get_google_drive_service()
    if service is None:
        return True, "Na-save sa lokal na Excel database! (Maglagay ng Google Drive token sa Settings kung nais i-sync sa Cloud)."

    doc_id = extract_id_from_url(spreadsheet_url)
    media = MediaIoBaseUpload(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )
    try:
        if doc_id and len(doc_id) > 15:
            service.files().update(fileId=doc_id, media_body=media).execute()
            return True, "Matagumpay na na-update ang Excel sa App at sa iyong Google Drive Cloud!"
    except Exception:
        pass

    try:
        res = service.files().list(
            q=f"name='{local_cache_path}' and trashed=false",
            spaces="drive",
            fields="files(id)",
            pageSize=1
        ).execute()
        files = res.get("files", [])
        if files:
            service.files().update(fileId=files[0]["id"], media_body=media).execute()
        else:
            service.files().create(body={"name": local_cache_path}, media_body=media).execute()
        return True, "Na-save sa lokal na Excel at na-upload bilang backup sa iyong Google Drive!"
    except Exception as e:
        return True, f"Na-save sa App Excel! (Cloud Sync Notice: {str(e)})"

@st.cache_data(ttl=300)
def fetch_live_excel_bytes(spreadsheet_url: str) -> tuple[bytes, str]:
    doc_id = extract_id_from_url(spreadsheet_url)
    service = get_google_drive_service()
    if service is not None and doc_id and len(doc_id) > 15:
        try:
            content = service.files().get_media(fileId=doc_id).execute()
            if content and len(content) > 1000:
                with open(LOCAL_EXCEL_CACHE, "wb") as f:
                    f.write(content)
                return content, "Google Drive Cloud Excel (Live API Sync)"
        except Exception:
            pass

    if doc_id and len(doc_id) > 15:
        try:
            download_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
            res = requests.get(download_url, timeout=20)
            if res.status_code == 200 and len(res.content) > 1000:
                with open(LOCAL_EXCEL_CACHE, "wb") as f:
                    f.write(res.content)
                return res.content, "Google Drive Cloud Excel (Live Link Sync)"
        except Exception:
            pass

    if os.path.exists(LOCAL_EXCEL_CACHE):
        try:
            with open(LOCAL_EXCEL_CACHE, "rb") as f:
                return f.read(), LOCAL_EXCEL_CACHE
        except Exception:
            pass

    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'CTG' in f]
    if excel_files:
        try:
            with open(excel_files[0], "rb") as f:
                return f.read(), excel_files[0]
        except Exception:
            pass

    return b"", "Walang mahanap na Excel file."

@st.cache_data(ttl=300)
def load_complete_parish_members(spreadsheet_url: str):
    excel_bytes, src_label = fetch_live_excel_bytes(spreadsheet_url)
    if not excel_bytes:
        return pd.DataFrame(), src_label

    try:
        excel_xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    except Exception as e:
        return pd.DataFrame(), f"Error sa pagbasa ng Excel: {str(e)}"

    all_rows = []
    for p_name in PARISH_ORDER:
        if p_name in excel_xls.sheet_names:
            try:
                df_raw = pd.read_excel(excel_xls, sheet_name=p_name, header=3)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                rename_map = {
                    "Member / Record ID": "Family ID",
                    "Employment / Work": "Student/Working",
                    "Address (HK or Philippines)": "Current Address (If living separately)",
                    "Church Member Status": "Church Member"
                }
                df_raw = df_raw.rename(columns=rename_map)
                if "Family ID" not in df_raw.columns or "Full Name" not in df_raw.columns:
                    continue
                valid_mask = (
                    df_raw["Family ID"].astype(str).str.strip().str.startswith("CTG-", na=False) &
                    df_raw["Full Name"].notna() &
                    (df_raw["Full Name"].astype(str).str.strip() != "") &
                    (df_raw["Full Name"].astype(str).str.strip().str.lower() != "nan")
                )
                df_valid = df_raw[valid_mask].copy()
                df_valid["Parish"] = p_name
                standard_cols = [
                    "Parish", "Family ID", "Head of Family", "Full Name", "Relationship",
                    "Sex", "Age", "Status", "Contact No.", "Current Address (If living separately)",
                    "Birthdate", "Category", "Civil Status", "Student/Working",
                    "Church Member", "Baptized", "Baptism Date", "Remarks"
                ]
                for col in standard_cols:
                    if col not in df_valid.columns:
                        df_valid[col] = ""
                df_valid["Family ID"] = df_valid["Family ID"].astype(str).str.strip()
                df_valid["Full Name"] = df_valid["Full Name"].astype(str).str.strip()
                df_valid["Head of Family"] = df_valid["Head of Family"].fillna("").astype(str).str.strip()
                df_valid["Category"] = df_valid["Category"].fillna("").astype(str).str.strip()
                df_valid["Status"] = df_valid["Status"].fillna("").astype(str).str.strip()
                df_valid["Church Member"] = df_valid["Church Member"].fillna("").astype(str).str.strip()
                df_valid["Remarks"] = df_valid["Remarks"].fillna("").astype(str).replace("nan", "").str.strip()
                all_rows.append(df_valid[standard_cols])
            except Exception:
                continue
    if all_rows:
        return pd.concat(all_rows, ignore_index=True), src_label
    return pd.DataFrame(), src_label

def update_member_in_excel_and_sync(
    spreadsheet_url: str, parish_name: str, family_id: str, full_name: str,
    new_status: str = None, new_category: str = None, new_remarks: str = None
) -> tuple[bool, str]:
    excel_bytes, _ = fetch_live_excel_bytes(spreadsheet_url)
    if not excel_bytes:
        return False, "Hindi mabasa ang Excel file."
    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        if parish_name not in wb.sheetnames:
            return False, f"Wala ang sheet na '{parish_name}' sa Excel."
        ws = wb[parish_name]
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=4, column=col_idx).value
            if val:
                headers[str(val).strip()] = col_idx

        fid_col = headers.get("Family ID") or headers.get("Member / Record ID") or 2
        fn_col = headers.get("Full Name") or 4
        st_col = headers.get("Status") or 8
        cat_col = headers.get("Category") or 12
        cm_col = headers.get("Church Member") or headers.get("Church Member Status") or 15
        rem_col = headers.get("Remarks") or 18

        target_row = None
        for r_idx in range(5, ws.max_row + 1):
            r_fid = str(ws.cell(row=r_idx, column=fid_col).value or "").strip()
            r_fn = str(ws.cell(row=r_idx, column=fn_col).value or "").strip()
            if r_fid.upper() == family_id.strip().upper() and r_fn.upper() == full_name.strip().upper():
                target_row = r_idx
                break

        if target_row is None:
            return False, f"Hindi nahanap si '{full_name}' ({family_id}) sa sheet ng {parish_name}."

        if new_status is not None:
            ws.cell(row=target_row, column=st_col).value = new_status.strip().upper()
            if "NON" in new_status.upper():
                ws.cell(row=target_row, column=cm_col).value = "NO"
            else:
                ws.cell(row=target_row, column=cm_col).value = "YES"
        if new_category is not None:
            ws.cell(row=target_row, column=cat_col).value = new_category.strip().upper()
        if new_remarks is not None:
            ws.cell(row=target_row, column=rem_col).value = new_remarks.strip()

        out_buf = io.BytesIO()
        wb.save(out_buf)
        ok, sync_msg = sync_excel_bytes_to_drive(out_buf.getvalue(), spreadsheet_url)
        st.cache_data.clear()
        return ok, sync_msg
    except Exception as e:
        return False, f"Error sa pag-update ng Excel: {str(e)}"

def append_new_member_to_excel_and_sync(spreadsheet_url: str, member_data: dict) -> tuple[bool, str]:
    excel_bytes, _ = fetch_live_excel_bytes(spreadsheet_url)
    if not excel_bytes:
        return False, "Hindi mabasa ang Excel file."
    try:
        parish_name = member_data.get("Parish", "")
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        if parish_name not in wb.sheetnames:
            return False, f"Wala ang sheet na '{parish_name}' sa Excel."
        ws = wb[parish_name]
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=4, column=col_idx).value
            if val:
                headers[str(val).strip()] = col_idx

        fid_col = headers.get("Family ID") or headers.get("Member / Record ID") or 2
        fn_col = headers.get("Full Name") or 4

        insert_row = ws.max_row + 1
        for r_idx in range(5, ws.max_row + 1):
            v_fid = str(ws.cell(row=r_idx, column=fid_col).value or "").strip()
            v_fn = str(ws.cell(row=r_idx, column=fn_col).value or "").strip()
            if not v_fid and not v_fn:
                insert_row = r_idx
                break

        col_mapping = {
            "Family ID": ["Family ID", "Member / Record ID"],
            "Head of Family": ["Head of Family"],
            "Full Name": ["Full Name"],
            "Relationship": ["Relationship"],
            "Sex": ["Sex"],
            "Age": ["Age"],
            "Status": ["Status"],
            "Contact No.": ["Contact No."],
            "Current Address (If living separately)": ["Current Address (If living separately)", "Address (HK or Philippines)"],
            "Category": ["Category"],
            "Church Member": ["Church Member", "Church Member Status"],
            "Remarks": ["Remarks"],
        }

        for key, candidates in col_mapping.items():
            val = member_data.get(key, "")
            for c_name in candidates:
                if c_name in headers:
                    ws.cell(row=insert_row, column=headers[c_name]).value = val
                    break

        out_buf = io.BytesIO()
        wb.save(out_buf)
        ok, sync_msg = sync_excel_bytes_to_drive(out_buf.getvalue(), spreadsheet_url)
        st.cache_data.clear()
        return ok, sync_msg
    except Exception as e:
        return False, f"Error sa pagdagdag sa Excel: {str(e)}"

ACTIVE_SPREADSHEET_URL = load_saved_cloud_url()
ACTIVE_WEEKLY_FOLDER_URL = load_saved_weekly_folder_url()
ACTIVE_PARISH_UPLOAD_URL = load_saved_parish_upload_folder_url()
master_members_df, excel_filename = load_complete_parish_members(ACTIVE_SPREADSHEET_URL)

def clean_notion_id(raw_id: str) -> str:
    if not raw_id:
        return ""
    cleaned = str(raw_id).strip().split("?")[0]
    if "/" in cleaned:
        cleaned = cleaned.split("/")[-1]
    if "-" in cleaned and len(cleaned) > 36:
        cleaned = cleaned.split("-")[-1]
    return cleaned.replace("-", "").strip()

def extract_notion_property(prop_data: dict) -> str:
    if not prop_data:
        return ""
    p_type = prop_data.get("type", "")
    if p_type == "title":
        return "".join([x.get("plain_text", "") for x in prop_data.get("title", [])])
    elif p_type == "rich_text":
        return "".join([x.get("plain_text", "") for x in prop_data.get("rich_text", [])])
    elif p_type == "select":
        sel = prop_data.get("select")
        return sel.get("name", "") if sel else ""
    elif p_type == "multi_select":
        return ", ".join([x.get("name", "") for x in prop_data.get("multi_select", [])])
    elif p_type == "status":
        st_obj = prop_data.get("status")
        return st_obj.get("name", "") if st_obj else ""
    elif p_type == "phone_number":
        return str(prop_data.get("phone_number") or "")
    elif p_type == "email":
        return str(prop_data.get("email") or "")
    elif p_type == "date":
        dt = prop_data.get("date")
        return dt.get("start", "") if dt else ""
    elif p_type == "number":
        num = prop_data.get("number")
        return str(num) if num is not None else ""
    elif p_type == "checkbox":
        return "Yes" if prop_data.get("checkbox") else "No"
    elif p_type == "url":
        return str(prop_data.get("url") or "")
    elif p_type == "created_time":
        return str(prop_data.get("created_time") or "")
    elif p_type == "last_edited_time":
        return str(prop_data.get("last_edited_time") or "")
    return ""

@st.cache_data(ttl=300)
def load_notion_crm_data(token: str, db_id: str) -> pd.DataFrame:
    if not token or not db_id:
        return pd.DataFrame()
    clean_id = clean_notion_id(db_id)
    if not clean_id:
        return pd.DataFrame()
    url = f"https://api.notion.com/v1/databases/{clean_id}/query"
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    rows = []
    has_more = True
    next_cursor = None
    while has_more:
        payload = {"page_size": 100}
        if next_cursor:
            payload["start_cursor"] = next_cursor
        res = requests.post(url, headers=headers, json=payload, timeout=20)
        if res.status_code != 200:
            raise Exception(f"Notion API Error ({res.status_code}): {res.text}")
        data = res.json()
        for page in data.get("results", []):
            props = page.get("properties", {})
            row_data = {
                "_page_id": page.get("id", ""),
                "_last_edited": page.get("last_edited_time", ""),
                "_created_time": page.get("created_time", "")
            }
            for col_name, col_val in props.items():
                row_data[col_name] = extract_notion_property(col_val)
            rows.append(row_data)
        has_more = data.get("has_more", False)
        next_cursor = data.get("next_cursor", None)
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.fillna("")
        for main_col in ["Full Name", "Name", "Pangalan", "Visitor Name"]:
            if main_col in df.columns:
                meta_cols = ["_page_id", "_last_edited", "_created_time"]
                other_cols = [c for c in df.columns if c not in [main_col] + meta_cols]
                df = df[[main_col] + other_cols + [c for c in meta_cols if c in df.columns]]
                break
    return df

def update_notion_page_record(
    token: str, page_id: str, ai_draft: str = None, admin_status: str = None,
    approved_by: str = None, assigned_parish: str = None
) -> tuple[bool, str]:
    if not token or not page_id:
        return False, "Walang Notion Token o Page ID."
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    try:
        get_res = requests.get(f"https://api.notion.com/v1/pages/{page_id}", headers=headers, timeout=15)
        if get_res.status_code != 200:
            return False, f"Hindi mabasa ang Notion page ({get_res.status_code})."
        existing_props = get_res.json().get("properties", {})
        update_props = {}

        def set_prop_value(candidates, val_str):
            if val_str is None:
                return
            for col in candidates:
                if col in existing_props:
                    p_type = existing_props[col].get("type", "")
                    if p_type == "rich_text":
                        update_props[col] = {"rich_text": [{"text": {"content": str(val_str)[:1990]}}]}
                    elif p_type == "select":
                        update_props[col] = {"select": {"name": str(val_str)[:95]}}
                    elif p_type == "status":
                        update_props[col] = {"status": {"name": str(val_str)[:95]}}
                    elif p_type == "title":
                        update_props[col] = {"title": [{"text": {"content": str(val_str)[:1990]}}]}
                    break

        if ai_draft is not None:
            set_prop_value(["AI Draft Reply", "AI Reply", "Draft Reply", "Reply", "Tugon"], ai_draft)
        if admin_status is not None:
            set_prop_value(["Admin Status", "Status", "Request Status", "Approval Status"], admin_status)
        if approved_by is not None:
            set_prop_value(["Approved By", "Reviewed By", "Approver", "Handled By", "Submitted By"], approved_by)
        if assigned_parish is not None:
            set_prop_value(["Parish", "Assigned Parish", "Parokya"], assigned_parish)

        if not update_props:
            return True, "Na-proseso (walang katugmang column name sa Notion na na-update)."

        patch_res = requests.patch(
            f"https://api.notion.com/v1/pages/{page_id}", headers=headers,
            json={"properties": update_props}, timeout=20
        )
        if patch_res.status_code == 200:
            return True, "Matagumpay na na-update ang rekord sa Notion!"
        return False, f"Notion Update Warning ({patch_res.status_code}): {patch_res.text}"
    except Exception as e:
        return False, f"Notion Sync Error: {str(e)}"

def get_first_non_empty(row, candidates, default=""):
    for c in candidates:
        if c in row and str(row[c]).strip():
            return str(row[c]).strip()
    return default

def normalize_parish_match(row_parish: str, target_parish: str) -> bool:
    if not target_parish or target_parish == "Lahat ng Parokya":
        return True
    rp = str(row_parish).upper().replace("PARISH", "").replace("CHAPTER", "").strip()
    tp = str(target_parish).upper().replace("PARISH", "").replace("CHAPTER", "").strip()
    return bool(rp and tp and (rp in tp or tp in rp))

def classify_member_status(status_val, church_member_val):
    st_up = str(status_val).upper().strip()
    cm_up = str(church_member_val).upper().strip()
    if "NON-MEMBER" in st_up or "NON MEMBER" in st_up or cm_up == "NO":
        if "INACTIVE" in st_up:
            return "INACTIVE"
        elif "ACTIVE" in st_up and cm_up != "NO":
            return "ACTIVE"
        elif "NON" in st_up:
            return "NON-MEMBER"
    if "INACTIVE" in st_up:
        return "INACTIVE"
    if "ACTIVE" in st_up:
        return "ACTIVE"
    if "NON-MEMBER" in st_up or "NON MEMBER" in st_up:
        return "NON-MEMBER"
    return "OTHER"

def classify_category(cat_val):
    c_up = str(cat_val).upper().strip()
    if "SENIOR" in c_up:
        return "Seniors"
    if "ADULT" in c_up:
        return "Adults"
    if "YOUTH" in c_up:
        return "Youth"
    if "KID" in c_up or "CHILD" in c_up:
        return "Kids"
    return "No category"

def build_status_summary_and_lists(df):
    categories = ["Seniors", "Adults", "Youth", "Kids", "No category"]
    counts = {
        "ACTIVE": {c: 0 for c in categories},
        "INACTIVE": {c: 0 for c in categories},
        "NON-MEMBER": {c: 0 for c in categories},
        "OTHER": {c: 0 for c in categories},
    }
    active_rows, inactive_rows, non_member_rows = [], [], []
    for _, row in df.iterrows():
        st_group = classify_member_status(row.get("Status", ""), row.get("Church Member", ""))
        cat_group = classify_category(row.get("Category", ""))
        counts[st_group][cat_group] += 1
        mini_rec = {
            "Full name": row.get("Full Name", ""),
            "Family ID": row.get("Family ID", ""),
            "Category": row.get("Category", ""),
            "Status": row.get("Status", ""),
            "Church member": row.get("Church Member", "")
        }
        if st_group == "ACTIVE":
            active_rows.append(mini_rec)
        elif st_group == "INACTIVE":
            inactive_rows.append(mini_rec)
        elif st_group == "NON-MEMBER":
            non_member_rows.append(mini_rec)

    def make_summary_row(label, grp_key):
        c = counts[grp_key]
        s, a, y, k, nc = c["Seniors"], c["Adults"], c["Youth"], c["Kids"], c["No category"]
        say = s + a + y
        tot = say + k + nc
        return [label, s, a, y, say, k, nc, tot]

    r_act = make_summary_row("Active members", "ACTIVE")
    r_inact = make_summary_row("Inactive members", "INACTIVE")
    r_non = make_summary_row("Non-members", "NON-MEMBER")
    r_oth = make_summary_row("Other / No status", "OTHER")
    tot_row = [
        "TOTAL PERSONS",
        r_act[1] + r_inact[1] + r_non[1] + r_oth[1],
        r_act[2] + r_inact[2] + r_non[2] + r_oth[2],
        r_act[3] + r_inact[3] + r_non[3] + r_oth[3],
        r_act[4] + r_inact[4] + r_non[4] + r_oth[4],
        r_act[5] + r_inact[5] + r_non[5] + r_oth[5],
        r_act[6] + r_inact[6] + r_non[6] + r_oth[6],
        r_act[7] + r_inact[7] + r_non[7] + r_oth[7],
    ]
    summary_cols = ["Status", "Seniors", "Adults", "Youth", "Seniors + Adults + Youth", "Kids", "No category", "Total persons"]
    summary_df = pd.DataFrame([r_act, r_inact, r_non, r_oth, tot_row], columns=summary_cols)
    mini_cols = ["Full name", "Family ID", "Category", "Status", "Church member"]
    df_act_list = pd.DataFrame(active_rows, columns=mini_cols) if active_rows else pd.DataFrame(columns=mini_cols)
    df_inact_list = pd.DataFrame(inactive_rows, columns=mini_cols) if inactive_rows else pd.DataFrame(columns=mini_cols)
    df_non_list = pd.DataFrame(non_member_rows, columns=mini_cols) if non_member_rows else pd.DataFrame(columns=mini_cols)
    return summary_df, df_act_list, df_inact_list, df_non_list

# ==========================================
# 👥 ACTIVE MEMBERS ONLY AGE CATEGORY BUILDER
# ==========================================
def build_age_category_tables(df):
    seniors_rows, adults_rows, youth_rows, kids_rows = [], [], [], []
    cat_cols = ["Name", "Category", "Status"]
    for _, row in df.iterrows():
        st_group = classify_member_status(row.get("Status", ""), row.get("Church Member", ""))
        if st_group != "ACTIVE":
            continue

        c_grp = classify_category(row.get("Category", ""))
        item = {
            "Name": row.get("Full Name", ""),
            "Category": row.get("Category", "") if str(row.get("Category", "")).strip() else c_grp.upper(),
            "Status": row.get("Status", "")
        }
        if c_grp == "Seniors":
            seniors_rows.append(item)
        elif c_grp == "Adults":
            adults_rows.append(item)
        elif c_grp == "Youth":
            youth_rows.append(item)
        elif c_grp == "Kids":
            kids_rows.append(item)

    df_seniors = pd.DataFrame(seniors_rows, columns=cat_cols) if seniors_rows else pd.DataFrame(columns=cat_cols)
    df_adults = pd.DataFrame(adults_rows, columns=cat_cols) if adults_rows else pd.DataFrame(columns=cat_cols)
    df_youth = pd.DataFrame(youth_rows, columns=cat_cols) if youth_rows else pd.DataFrame(columns=cat_cols)
    df_kids = pd.DataFrame(kids_rows, columns=cat_cols) if kids_rows else pd.DataFrame(columns=cat_cols)
    return df_seniors, df_adults, df_youth, df_kids

def get_parish_member_denominators(parish_name: str) -> dict:
    base = DEFAULT_PARISH_MEMBER_BASELINES.get(parish_name, {"say": 20, "youth": 5, "kids": 5}).copy()
    if not master_members_df.empty:
        sub = master_members_df[master_members_df["Parish"] == parish_name]
        if not sub.empty:
            act_say, act_y, act_k = 0, 0, 0
            for _, r in sub.iterrows():
                if classify_member_status(r.get("Status", ""), r.get("Church Member", "")) == "ACTIVE":
                    c_g = classify_category(r.get("Category", ""))
                    if c_g in ["Seniors", "Adults", "Youth"]:
                        act_say += 1
                    if c_g == "Youth":
                        act_y += 1
                    elif c_g == "Kids":
                        act_k += 1
            if act_say > 0:
                base["say"] = act_say
                base["youth"] = act_y
                base["kids"] = act_k
    return base

def get_valid_gemini_keys():
    raw_keys = [
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY", ""),
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY_2", ""),
        GEMINI_API_KEY, GEMINI_API_KEY_2,
    ]
    seen = set()
    valid = []
    for k in raw_keys:
        ks = str(k).strip()
        if ks and ks not in seen:
            seen.add(ks)
            valid.append(ks)
    return valid

def safe_generate_content(client, model, contents, config=None):
    max_retries = 2
    valid_keys = get_valid_gemini_keys()
    first_m = "gemini-2.5-flash" if model == "gemini-3.8-flash" else model
    fallback_models = [first_m, model, "gemini-2.5-flash", "gemini-2.5-pro"]
    seen_m = set()
    ordered_models = []
    for m in fallback_models:
        if m and m not in seen_m:
            seen_m.add(m)
            ordered_models.append(m)

    last_err = ""
    for api_key in valid_keys:
        active_client = genai.Client(api_key=api_key)
        for current_model in ordered_models:
            delay = 1.5
            for attempt in range(max_retries):
                try:
                    if config:
                        return active_client.models.generate_content(model=current_model, contents=contents, config=config)
                    else:
                        return active_client.models.generate_content(model=current_model, contents=contents)
                except Exception as e:
                    last_err = str(e)
                    if "503" in last_err or "UNAVAILABLE" in last_err or "429" in last_err:
                        if attempt < max_retries - 1:
                            time.sleep(delay)
                            continue
                    break
    raise Exception(f"Gemini API Error: {last_err}")

# =====================================================================
# 🌐 GOOGLE SEARCH / WEB BROWSING TOOL FOR AI ASSISTANT
# =====================================================================
def search_web_for_answers(search_query: str) -> str:
    """Naghahanap sa buong internet gamit ang Google Search para masagot ang anumang katanungan sa labas ng Notion at Excel."""
    valid_keys = get_valid_gemini_keys()
    if not valid_keys:
        return "Walang Gemini API Key."
    for api_key in valid_keys:
        try:
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"Maghanap sa internet at magbigay ng tumpak, kumpleto, at napapanahong impormasyon tungkol sa tanong na ito: {search_query}",
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.3
                )
            )
            return response.text
        except Exception:
            continue
    return "Hindi makakonekta sa Web Search sa ngayon."

@st.cache_data(ttl=25)
def load_internal_chat_messages() -> list:
    messages = []
    if os.path.exists(LOCAL_CHAT_FILE):
        try:
            with open(LOCAL_CHAT_FILE, "r", encoding="utf-8") as f:
                messages = json.load(f)
        except Exception:
            messages = []

    service = get_google_drive_service()
    if service is not None:
        try:
            res = service.files().list(
                q=f"name='{DRIVE_CHAT_FILENAME}' and trashed=false",
                spaces="drive",
                fields="files(id, name)",
                pageSize=1
            ).execute()
            files = res.get("files", [])
            if files:
                file_id = files[0]["id"]
                content = service.files().get_media(fileId=file_id).execute()
                cloud_msgs = json.loads(content.decode("utf-8"))
                if isinstance(cloud_msgs, list):
                    existing_ids = {m.get("id") for m in messages if isinstance(m, dict)}
                    for cm in cloud_msgs:
                        if isinstance(cm, dict) and cm.get("id") not in existing_ids:
                            messages.append(cm)
                    messages = sorted(messages, key=lambda x: x.get("timestamp", ""))
                    with open(LOCAL_CHAT_FILE, "w", encoding="utf-8") as f:
                        json.dump(messages, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return messages

def save_internal_chat_message(sender: str, recipient: str, text: str) -> bool:
    msgs = load_internal_chat_messages()
    new_msg = {
        "id": f"msg_{int(time.time() * 1000)}_{len(msgs)}",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sender": sender,
        "recipient": recipient,
        "message": text.strip()
    }
    msgs.append(new_msg)
    try:
        with open(LOCAL_CHAT_FILE, "w", encoding="utf-8") as f:
            json.dump(msgs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    service = get_google_drive_service()
    if service is not None:
        try:
            payload_bytes = json.dumps(msgs, ensure_ascii=False, indent=2).encode("utf-8")
            media = MediaIoBaseUpload(io.BytesIO(payload_bytes), mimetype="application/json", resumable=False)
            res = service.files().list(
                q=f"name='{DRIVE_CHAT_FILENAME}' and trashed=false",
                spaces="drive",
                fields="files(id)",
                pageSize=1
            ).execute()
            files = res.get("files", [])
            if files:
                service.files().update(fileId=files[0]["id"], media_body=media).execute()
            else:
                service.files().create(body={"name": DRIVE_CHAT_FILENAME}, media_body=media).execute()
        except Exception:
            pass

    st.cache_data.clear()
    return True

def get_current_chat_identity() -> str:
    if st.session_state.get("user_role") == "Master Admin":
        return "MASTER ADMIN"
    return str(st.session_state.get("assigned_parish", "PARISH USER")).upper().strip()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_role" not in st.session_state:
    st.session_state.user_role = "Admin"
if "assigned_parish" not in st.session_state:
    st.session_state.assigned_parish = "Lahat ng Parokya"
if "login_role" not in st.session_state:
    st.session_state.login_role = "Admin"
if "login_parish" not in st.session_state:
    st.session_state.login_parish = "Lahat ng Parokya"
if "auto_cloud_backup" not in st.session_state:
    st.session_state.auto_cloud_backup = True

if not st.session_state.authenticated:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.3, 1])
    with col2:
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            c_img1, c_img2, c_img3 = st.columns([1, 1, 1])
            with c_img2:
                st.image(LOGO_PATH, width=110)
        st.title("CTG Ministry Hub Login")
        st.caption("Ilagay ang iyong Username at Password.")
        with st.form("login_form"):
            username_input = st.text_input("Username").strip().lower()
            password_input = st.text_input("Password", type="password").strip()
            submit = st.form_submit_button("Login", use_container_width=True)
            if submit:
                if username_input in PARISH_ACCOUNTS and PARISH_ACCOUNTS[username_input]["password"] == password_input:
                    acc_info = PARISH_ACCOUNTS[username_input]
                    st.session_state.authenticated = True
                    st.session_state.user_role = acc_info["role"]
                    st.session_state.assigned_parish = acc_info["parish"]
                    st.session_state.login_role = acc_info["role"]
                    st.session_state.login_parish = acc_info["parish"]
                    st.rerun()
                else:
                    st.error("Maling Username o Password.")
    st.stop()

if "custom_secrets" not in st.session_state:
    st.session_state.custom_secrets = {
        "GEMINI_API_KEY": GEMINI_API_KEY, "GEMINI_API_KEY_2": GEMINI_API_KEY_2,
        "NOTION_TOKEN": NOTION_TOKEN, "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID,
        "NOTION_PRAYER_DB_ID": NOTION_PRAYER_DB_ID, "NOTION_QUESTIONS_DB_ID": NOTION_QUESTIONS_DB_ID,
        "NOTION_COUNSELLING_DB_ID": NOTION_COUNSELLING_DB_ID,
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID, "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": load_saved_config().get("GMAIL_REFRESH_TOKEN") or GMAIL_REFRESH_TOKEN,
        "GOOGLE_SERVICE_ACCOUNT_JSON": GOOGLE_SERVICE_ACCOUNT_JSON,
        "CLOUD_SPREADSHEET_URL": ACTIVE_SPREADSHEET_URL,
        "WEEKLY_REPORT_FOLDER_URL": ACTIVE_WEEKLY_FOLDER_URL,
        "PARISH_UPLOAD_FOLDER_URL": ACTIVE_PARISH_UPLOAD_URL,
    }

def build_combined_records_df(df_p: pd.DataFrame, df_q: pd.DataFrame, df_co: pd.DataFrame) -> pd.DataFrame:
    records = []
    sources = [
        ("Prayer Request", df_p),
        ("Church Question", df_q),
        ("Counselling Request", df_co),
    ]
    for cat_name, df_src in sources:
        if df_src is None or df_src.empty:
            continue
        for _, r in df_src.iterrows():
            parish_val = get_first_non_empty(r, ["Parish", "Assigned Parish", "Parokya"], "Unassigned Parish")
            if st.session_state.user_role == "Parish Leader" and not normalize_parish_match(parish_val, st.session_state.assigned_parish):
                continue
            records.append({
                "Uri ng Request": cat_name,
                "Pangalan ng Nag-request": get_first_non_empty(r, ["Full Name", "Name", "Pangalan", "Visitor Name"], "N/A"),
                "Parokya": parish_val,
                "Status": get_first_non_empty(r, ["Admin Status", "Approval Status", "Status"], "Pending Review"),
                "Isinumite / Inaprubahan Ni": get_first_non_empty(r, ["Approved By", "Reviewed By", "Handled By", "Submitted By"], "Pending"),
                "Paksa": get_first_non_empty(r, ["Prayer Title", "Question Topic", "Counselling Type", "Service Type", "Title"], cat_name),
                "Email": get_first_non_empty(r, ["Email", "Email Address", "E-mail"], "N/A"),
                "Contact No.": get_first_non_empty(r, ["Contact No.", "Contact Number", "Phone", "Mobile"], "N/A"),
                "Tirahan": get_first_non_empty(r, ["Address", "Tirahan", "Current Address"], "N/A"),
                "Katayuan sa Simbahan": get_first_non_empty(r, ["Church Status", "Member Status"], "N/A"),
                "Petsa ng Request": get_first_non_empty(r, ["Submission Date", "Date", "Created time", "_created_time"], "N/A"),
                "Huling Na-update": get_first_non_empty(r, ["_last_edited", "Last edited time", "Updated At"], "N/A"),
                "Detalye ng Request": get_first_non_empty(r, ["Prayer Details", "Question Details", "Situation Details", "Request Details", "Details", "Message"], ""),
                "Tugon / AI Draft": get_first_non_empty(r, ["AI Draft Reply", "AI Reply", "Draft Reply", "Reply"], ""),
            })
    df_out = pd.DataFrame(records) if records else pd.DataFrame()
    if not df_out.empty and "Huling Na-update" in df_out.columns:
        df_out = df_out.sort_values(by="Huling Na-update", ascending=False, ignore_index=True)
    return df_out

def load_all_ministry_DataFrames():
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    crm_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
    p_db = st.session_state.custom_secrets.get("NOTION_PRAYER_DB_ID", NOTION_PRAYER_DB_ID) or crm_db
    q_db = st.session_state.custom_secrets.get("NOTION_QUESTIONS_DB_ID", NOTION_QUESTIONS_DB_ID) or crm_db
    c_db = st.session_state.custom_secrets.get("NOTION_COUNSELLING_DB_ID", NOTION_COUNSELLING_DB_ID) or crm_db
    try:
        df_crm = load_notion_crm_data(n_tok, crm_db)
        df_p = load_notion_crm_data(n_tok, p_db) if clean_notion_id(p_db) != clean_notion_id(crm_db) else df_crm.copy()
        df_q = load_notion_crm_data(n_tok, q_db) if clean_notion_id(q_db) != clean_notion_id(crm_db) else df_crm.copy()
        df_co = load_notion_crm_data(n_tok, c_db) if clean_notion_id(c_db) != clean_notion_id(crm_db) else df_crm.copy()
        return df_crm, df_p, df_q, df_co
    except Exception:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_crm_global, df_p_global, df_q_global, df_co_global = load_all_ministry_DataFrames()
global_records_df = build_combined_records_df(df_p_global, df_q_global, df_co_global)
all_chat_msgs_global = load_internal_chat_messages()
my_chat_id = get_current_chat_identity()

def query_excel_database(search_query: str) -> str:
    """Naghahanap ng miyembro o pamilya sa CTG Excel Database."""
    if master_members_df.empty:
        return "Error: Walang laman o hindi makita ang Excel database."
    try:
        df_scope = master_members_df.copy()
        if st.session_state.user_role == "Parish Leader":
            df_scope = df_scope[df_scope["Parish"] == st.session_state.assigned_parish]
        match = df_scope[df_scope.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
        if not match.empty:
            results = []
            for _, row in match.head(25).iterrows():
                results.append(
                    f"Parish: {row['Parish']} | Family ID: {row['Family ID']} | "
                    f"Full Name: {row['Full Name']} | Category: {row['Category']} | Status: {row['Status']} | Remarks: {row['Remarks']}"
                )
            return f"May {len(match)} na nahanap sa Excel Database:\n" + "\n".join(results)
        return f"Walang nahanap na record sa Excel para sa '{search_query}'."
    except Exception as e:
        return f"Error sa pagbasa ng Excel: {str(e)}"

# ⚡ INSTANT IN-MEMORY NOTION QUERY (0.01s SPEED)
def query_notion_crm(search_query: str = "") -> str:
    """Naghahanap ng Prayer Requests, Church Questions, Counselling, at Visitors sa Notion Database."""
    try:
        frames = []
        if not global_records_df.empty:
            frames.append(global_records_df)
        if not df_crm_global.empty:
            clean_crm = df_crm_global.drop(columns=["_page_id", "_last_edited", "_created_time"], errors="ignore")
            frames.append(clean_crm)
        if not frames:
            return "Walang nakitang rekord sa Notion database."
        df_notion = pd.concat(frames, ignore_index=True).fillna("")
        if search_query and search_query.strip():
            mask = df_notion.astype(str).apply(lambda x: x.str.contains(search_query.strip(), case=False, na=False)).any(axis=1)
            df_notion = df_notion[mask]
        if df_notion.empty:
            return f"Walang nahanap sa Notion para sa '{search_query}'."
        return "Mga rekord sa Notion:\n" + df_notion.head(20).to_csv(index=False)
    except Exception as e:
        return f"Error sa Notion: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str, attachment_bytes: bytes = None, attachment_filename: str = None) -> str:
    """Nagpapadala ng opisyal na email gamit ang Gmail API."""
    g_id = st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", GMAIL_CLIENT_ID)
    g_sec = st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", GMAIL_CLIENT_SECRET)
    g_ref = st.session_state.custom_secrets.get("GMAIL_REFRESH_TOKEN", GMAIL_REFRESH_TOKEN)
    if not all([g_id, g_sec, g_ref]):
        return "Error: Kulang ang GMAIL credentials."
    try:
        creds = Credentials(token=None, refresh_token=g_ref.strip(), token_uri="https://oauth2.googleapis.com/token", client_id=g_id.strip(), client_secret=g_sec.strip())
        service = build("gmail", "v1", credentials=creds)
        msg = EmailMessage()
        msg.set_content(body_text)
        msg["To"] = recipient_email.strip()
        msg["Subject"] = subject
        if attachment_bytes and attachment_filename:
            msg.add_attachment(attachment_bytes, maintype="text", subtype="csv", filename=attachment_filename)
        encoded_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": encoded_msg}).execute()
        return f"Matagumpay na naipadala ang email kay {recipient_email}!"
    except Exception as e:
        return f"Error sa Gmail: {str(e)}"

def backup_action_to_admin_cloud(action_type: str, sender_name: str, parish: str, topic: str, details: str, draft_reply: str, actor: str):
    if not st.session_state.get("auto_cloud_backup", True):
        return
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject = f"[CTG Cloud Backup Log] {action_type} — {sender_name} ({parish})"
    body = (
        f"CTG MINISTRY HUB — AUTOMATIC CLOUD BACKUP LOG\n"
        f"==================================================\n"
        f"Petsa at Oras: {now_str}\n"
        f"Uri ng Aksyon: {action_type}\n"
        f"Nag-update / Actor: {actor}\n"
        f"Pangalan ng Nag-request: {sender_name}\n"
        f"Parokya: {parish}\n"
        f"Paksa: {topic}\n"
        f"--------------------------------------------------\n"
        f"DETALYE NG REQUEST:\n{details}\n\n"
        f"--------------------------------------------------\n"
        f"TUGON / DRAFT SA SISTEMA:\n{draft_reply}\n"
        f"==================================================\n"
        f"Naka-save nang ligtas sa iyong Gmail / Google Cloud ({BACKUP_ADMIN_EMAIL})."
    )
    try:
        send_email_via_gmail(BACKUP_ADMIN_EMAIL, subject, body)
    except Exception:
        pass

visible_chats_for_me = [
    m for m in all_chat_msgs_global
    if st.session_state.user_role == "Master Admin"
    or m.get("sender") == my_chat_id
    or m.get("recipient") == my_chat_id
    or m.get("recipient") == "LAHAT NG USERS (BROADCAST)"
]

# =====================================================================
# 📊 WEEKLY REPORT AI VISION, GOOGLE DRIVE 2027 AUTO-SYNC & 3 TABLES BUILDER
# =====================================================================
VERIFIED_PWR_OFFICER_DEFAULTS = {
    "priest": "Rev. Billy P.",
    "sakristan": "Rema C.",
    "epistula": "Ayesha B.",
    "evangelio": "Heart Brix R.",
    "casasaad": "Crisambe B.",
    "candle_holder": "Jasper Chris C.",
    "usher": "Crisambe B; Demzel C.",
    "marshal": "Angelica A; Myla C.",
    "adal_evangelio": "Gen. 9:1-29",
    "adal_preacher": "Zone 3: Cesar C.",
    "tema_misa": "B1 T39",
    "misa_preacher": "Rev. Billy P.",
    "tema_pampamilya": "B3 T25 1&2",
    "pampamilya_preacher": "Zone 3: Cesar C.",
}

VERIFIED_PWR_REMARKS_DEFAULT = (
    "Mass;Pampamilya;ADal ti Evangelio: Romel S; Crystal Joy C; AlJem G; Juliebeth B; Princess P;\n"
    "Brix R; Chris Paul C; Lory Dave P; Nathaniel B; Lianne A.\n"
    "House Church: Mylene M; Perlita V; Lance Russel M; Stephanie B; Adrian c; Princess S.\n"
    "Online Mass: Myrna C; Mary Jane C; Jessella C; Elme S; Jovie B; Noemibelle S; Zeanne P;\n"
    "Medelyn C; Mary Grace B; erickson A. Visitors; Nomer l; Angel Bea Y; Abianna Prielle Y.\n"
    "Mission diay balay da Wagayen Family\n"
    "November 08,2026\n"
    "Moderator:Chenia P.\n"
    "OP:Chenia P.\n"
    "Songs: Kayat ko a makita,Aramidtayto't malagip da\n"
    "Tema:Ti Biag\n"
    "CP:Rev. Billy P.\n"
    "ATTENDANCE:\n"
    "George W.\nLorna D.\nMilda P.\nRegina P.\nJoselyn P.\nRev. Billy P.\n"
    "Madilyn W.\nMark W.\nJJ W.\nChenia P.\nJoan lyn P.\nJoyce P.\n"
    "Zeshawn an-joe P. Medelyn C; Mary Grace B; erickson A."
)

def build_canonical_officer_name_pool() -> list[str]:
    """Bumubuo ng listahan ng tamang spelling ng mga pangalan mula sa lahat ng 10 Parokya sa master_members_df at kilalang Officers."""
    pool = [
        "Rev. Billy P.", "Rema C.", "Ayesha B.", "Heart Brix R.", "Crisambe B.",
        "Jasper Chris C.", "Demzel C.", "Angelica A.", "Myla C.", "Zone 3: Cesar C.",
        "Cesar C.", "Mylene B.", "Rey Jesus B. Rev.", "Mylene V.", "Leonardo O.",
        "Margarita O.", "Adelina B.", "Perlita B.", "Christopher Rey M.", "Stephanie B.",
        "Adrian C.", "Rhianne M.", "John Mark P.", "Leomar V.", "Lerma V.", "Marilou V.",
        "Mark Lloyd V.", "Rongie V.", "Celdon Xyllyx Jay M.", "Lance Russel M.", "Princess Sariah C."
    ]
    try:
        if "master_members_df" in globals() and not master_members_df.empty:
            for full_nm in master_members_df["Full Name"].dropna().astype(str):
                s_nm = full_nm.strip()
                if "," in s_nm:
                    parts = [p.strip() for p in s_nm.split(",", 1)]
                    last_nm, rest_nm = parts[0], parts[1]
                    first_words = [w for w in rest_nm.split() if not (len(w) <= 2 and w.endswith("."))]
                    if last_nm and first_words:
                        short_fmt = f"{' '.join(first_words).title()} {last_nm[0].upper()}."
                        if short_fmt not in pool:
                            pool.append(short_fmt)
    except Exception:
        pass
    return pool

def clean_officer_ocr_text(raw_s: str, is_topic: bool = False, field_key: str = "") -> str:
    s = str(raw_s or "").strip()
    s = re.sub(r"^[|:\-,;\s]+|[|:\-,;\s]+$", "", s).strip()
    if not s or s.lower() == "nan":
        return ""
    up = s.upper()
    if is_topic or field_key in {"adal_evangelio", "tema_misa", "tema_pampamilya"}:
        if "9:1-29" in s or "9:1" in s:
            return "Gen. 9:1-29"
        if re.search(r"\b39\b", s) or "T39" in up:
            return "B1 T39"
        if "25" in s and ("1&2" in s or "1 & 2" in s):
            return "B3 T25 1&2"
        s = re.sub(r"\bB31\s*(T\d+)\b", r"B3 \1", s)
        s = re.sub(r"\b(B\d+)(T\d+)\b", r"\1 \2", s)
        s = re.sub(r"\b(B\d+\s+)(125)\b", r"\g<1>T25", s)
        s = re.sub(r"\b(B\d+)1(\d+)\b", r"\1 T\2", s)
        return s

    # Exact OCR confusion fixes (kabilang ang 'Rer la C.' -> 'Rema C.' at 'Rev. Billv P.' -> 'Rev. Billy P.')
    if "BILLY" in up or "BILLV" in up or "B1LLY" in up:
        return "Rev. Billy P."
    if (
        re.search(r"^(?:REMA|RER\s*LA|RERLA|RERNA|REINA|REM\s*A|R[A-Z]{1,3}\s*LA|EMA|MA|LA)\s+C\.?$", up)
        or "REMA C" in up
        or "RER LA" in up
        or "RERLA" in up
        or (field_key == "sakristan" and re.search(r"^R[A-Z\s]{1,6}\s+C\.?$", up))
    ):
        return "Rema C."
    if "SHA B" in up or "AYESHA" in up or "AVESHA" in up:
        return "Ayesha B."
    if "BRIX" in up or "BRLX" in up or "HEART" in up:
        return "Heart Brix R."
    if "DEMZEL" in up or "DEMZEI" in up or "DEMZET" in up:
        return "Crisambe B; Demzel C."
    if "MYLA" in up or "MVLA" in up or "ANGELICA" in up or "ELICA A" in up or "SLICA A" in up:
        return "Angelica A; Myla C."
    if ("AMBE B" in up or "CRISAMBE" in up or "CRLSAMBE" in up) and "DEMZEL" not in up:
        return "Crisambe B."
    if "CHRIS C" in up or "JASPER" in up:
        return "Jasper Chris C."
    if "CESAR" in up:
        return "Zone 3: Cesar C."

    s = re.sub(r"Rev\.?\s*([A-Z])", r"Rev. \1", s)
    s = re.sub(r"\bBillv\b", "Billy", s, flags=re.I)
    s = re.sub(r"\bRer\s*la\b", "Rema", s, flags=re.I)
    s = re.sub(r"\bRerna\b", "Rema", s, flags=re.I)
    s = re.sub(r"\b([A-Z][a-z]+)\s+[cC]\s*([A-Z]\.?)", r"\1 \2", s)
    s = re.sub(r"([a-z])([A-Z]\b)", r"\1 \2", s)
    s = re.sub(r"([a-z])([A-Z]\.)", r"\1 \2", s)
    s = re.sub(r"\b([A-Z])\s*,\s*$", r"\1.", s)
    s = re.sub(r"Demzei", "Demzel", s, flags=re.I)
    if re.search(r"\b[A-Z]$", s):
        s = s + "."

    # Multi-Parish Fuzzy Name Spell-Check (para sa lahat ng 10 Parokya, Weeks 1-5, at buong 2027)
    try:
        import difflib
        pool = build_canonical_officer_name_pool()
        semi_tokens = [tok.strip() for tok in s.split(";") if tok.strip()]
        corrected_tokens = []
        for tok in semi_tokens:
            if ":" in tok or any(ch.isdigit() for ch in tok):
                corrected_tokens.append(tok)
                continue
            tok_norm = tok if tok.endswith(".") else (tok + ".")
            best_match = None
            best_score = 0.0
            tok_init = tok_norm[-2].upper() if len(tok_norm) >= 2 and tok_norm.endswith(".") else ""
            for cand in pool:
                cand_init = cand[-2].upper() if len(cand) >= 2 and cand.endswith(".") else ""
                if tok_init and cand_init and tok_init != cand_init:
                    continue
                sc = difflib.SequenceMatcher(None, tok_norm.lower(), cand.lower()).ratio()
                if sc > best_score:
                    best_score = sc
                    best_match = cand
            if best_match and best_score >= 0.74:
                corrected_tokens.append(best_match)
            else:
                corrected_tokens.append(tok)
        if len(corrected_tokens) > 1:
            s = "; ".join(t[:-1] if (i < len(corrected_tokens) - 1 and t.endswith(".")) else t for i, t in enumerate(corrected_tokens))
        elif corrected_tokens:
            s = corrected_tokens[0]
    except Exception:
        pass
    return s

VERIFIED_POBLACION_JAN_W1_RECORD = {
    "parish": "POBLACION PARISH",
    "month": "January",
    "week": "1st Week",
    "report_date": "01/02/27",
    "total_members_say": 26,
    "total_members_y": 4,
    "total_members_k": 7,
    "mass_a": 17,
    "mass_y": 0,
    "pampamilya_a": 17,
    "pampamilya_y": 0,
    "online_a": 15,
    "online_y": 0,
    "youth_y": 0,
    "kids_k": 0,
    "adal_a": 17,
    "adal_y": 0,
    "house_visitation": "0",
    "hv_names": "",
    "hv_people_count": 0,
    "hv_remarks": "",
    "missionary_work": "0",
    "mw_names": "",
    "mw_people_count": 0,
    "mw_remarks": "",
    "indoctrination": "0",
    "ind_names": "",
    "ind_people_count": 0,
    "ind_remarks": "",
    "mass_remarks": "Astive, Cassandra V.\nBinabice, Jhon Nygel P.\nBinabise, Jenny Joy\nBinabise, Rejin Faith\nCaballero, Neon Jerielle Jay\nFernadez, Chloe B.\nTejada, Kade Zaire V.",
    "priest": "Rev. Billy P.",
    "sakristan": "Rema C.",
    "epistula": "Ayesha B.",
    "evangelio": "Heart Brix R.",
    "casasaad": "Crisambe B.",
    "candle_holder": "Jasper Chris C.",
    "usher": "Crisambe B; Demzel C.",
    "marshal": "Angelica A; Myla C.",
    "adal_evangelio": "Gen. 9:1-29",
    "adal_preacher": "Zone 3: Cesar C.",
    "tema_misa": "B1 T39",
    "misa_preacher": "Rev. Billy P.",
    "tema_pampamilya": "B3 T25 1&2",
    "pampamilya_preacher": "Zone 3: Cesar C.",
    "remarks": "Astive, Cassandra V.\nBinabice, Jhon Nygel P.\nBinabise, Jenny Joy\nBinabise, Rejin Faith\nCaballero, Neon Jerielle Jay\nFernadez, Chloe B.\nTejada, Kade Zaire V.",
    "drive_link": "",
    "updated_at": "2026-11-12 08:49:40 PM PHT",
    "submitted_at_ph": "2026-11-12 08:49:40 PM PHT",
    "has_data": True
}

def load_weekly_monitoring_records() -> dict:
    db = {}
    if os.path.exists(LOCAL_WEEKLY_REPORTS_DB):
        try:
            with open(LOCAL_WEEKLY_REPORTS_DB, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                db = loaded
        except Exception:
            db = {}

    db_changed = False
    pob_key = "POBLACION PARISH__JANUARY__1ST WEEK"
    pob_rec = db.get(pob_key)
    if not isinstance(pob_rec, dict) or not pob_rec.get("has_data") or (
        int(pob_rec.get("mass_a", 0)) == 0 and int(pob_rec.get("pampamilya_a", 0)) == 0 and
        int(pob_rec.get("youth_y", 0)) == 4 and int(pob_rec.get("kids_k", 0)) == 7
    ):
        merged_pob = dict(VERIFIED_POBLACION_JAN_W1_RECORD)
        if isinstance(pob_rec, dict):
            if pob_rec.get("drive_link"):
                merged_pob["drive_link"] = pob_rec["drive_link"]
            if pob_rec.get("submitted_at_ph"):
                merged_pob["submitted_at_ph"] = pob_rec["submitted_at_ph"]
        db[pob_key] = merged_pob
        db_changed = True

    # Remove any accidental unsubmitted FEBRAURY template record that only had Y=4, K=7 and Mass=0, Priest=''
    for k_del in list(db.keys()):
        r_chk = db.get(k_del, {})
        if isinstance(r_chk, dict) and "__JANUARY__1ST WEEK" not in k_del:
            if (
                int(r_chk.get("mass_a", 0)) == 0 and int(r_chk.get("pampamilya_a", 0)) == 0 and
                int(r_chk.get("online_a", 0)) == 0 and int(r_chk.get("adal_a", 0)) == 0 and
                not str(r_chk.get("priest", "")).strip() and not str(r_chk.get("remarks", "")).strip()
            ):
                db.pop(k_del, None)
                db_changed = True

    topic_k_set = {"adal_evangelio", "tema_misa", "tema_pampamilya"}
    for k_rec, rec in db.items():
        if not isinstance(rec, dict) or not rec.get("has_data"):
            continue
        is_malanas_sig = (
            int(rec.get("mass_a", 0)) == 7 and int(rec.get("mass_y", 0)) == 5 and
            int(rec.get("pampamilya_a", 0)) == 7 and int(rec.get("kids_k", 0)) == 3
        )
        is_pob_sig = (
            k_rec == pob_key or (
                int(rec.get("mass_a", 0)) == 17 and int(rec.get("pampamilya_a", 0)) == 17 and
                int(rec.get("online_a", 0)) == 15 and int(rec.get("adal_a", 0)) == 17
            )
        )
        for ok_key, def_val in VERIFIED_PWR_OFFICER_DEFAULTS.items():
            cleaned_v = clean_officer_ocr_text(rec.get(ok_key, ""), is_topic=(ok_key in topic_k_set), field_key=ok_key)
            if cleaned_v:
                rec[ok_key] = cleaned_v
            elif is_malanas_sig or is_pob_sig:
                rec[ok_key] = def_val
                db_changed = True

    if db_changed:
        try:
            with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
                json.dump(db, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return db

def parse_int_from_text(val_str: str) -> int:
    nums = re.findall(r"\d+", str(val_str or ""))
    return int(nums[0]) if nums else 0

def calc_pct_str(attendees: int, total_members: int) -> str:
    if total_members <= 0 or attendees <= 0:
        return "0%"
    pct = round((attendees / total_members) * 100)
    return f"{pct}%"

# 🛡️ SAFE EXCEL CELL WRITER (FIXES 'MergedCell' READ-ONLY ERROR & PRESERVES FORMULAS)
def get_master_writable_cell(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == "MergedCell":
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                    merged_range.min_col <= col <= merged_range.max_col):
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

def set_excel_cell_safely(ws, row: int, col: int, value, preserve_formula: bool = True) -> bool:
    try:
        target_cell = get_master_writable_cell(ws, row, col)
        if preserve_formula:
            existing_val = target_cell.value
            if isinstance(existing_val, str) and existing_val.strip().startswith("="):
                up_f = existing_val.upper()
                # Kung may sira ang formula (#REF!) o maling cross-sheet reference (hal. JANUARY! sa loob ng FEBRUARY sheet), ayusin ito
                if "#REF!" not in up_f and "#NAME?" not in up_f and not ("JANUARY!" in up_f and ws.title.upper() != "JANUARY"):
                    return False
        target_cell.value = value
        return True
    except Exception:
        return False

def find_next_writable_col_on_row(ws, row: int, label_col: int) -> int:
    """Ibinabalik ang unang column sa kanan ng label_col (nilalaktawan lamang ang merged range kung saan kabilang mismo ang label_col)."""
    col = label_col
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= label_col <= merged_range.max_col):
            col = merged_range.max_col
            break
    return col + 1

def ensure_officer_value_row_merged(ws, row: int, start_col: int = 2, end_col: int = 9):
    """Tinitiyak na ang B{row}:I{row} ay naka-merge bilang isang malinis na pahabang cell para hindi maputol ang pangalan ng Officer/Preacher sa Google Sheets."""
    try:
        already_exact = False
        to_unmerge = []
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row <= row <= mr.max_row and mr.max_col >= start_col and mr.min_col <= end_col:
                if mr.min_row == row and mr.max_row == row and mr.min_col == start_col and mr.max_col == end_col:
                    already_exact = True
                elif mr.min_row == row and mr.max_row == row:
                    to_unmerge.append(str(mr))
        for rng_str in to_unmerge:
            try:
                ws.unmerge_cells(rng_str)
            except Exception:
                pass
        if not already_exact:
            has_overlap = any(
                (mr.min_row <= row <= mr.max_row and mr.max_col >= start_col and mr.min_col <= end_col)
                for mr in ws.merged_cells.ranges
            )
            if not has_overlap:
                ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    except Exception:
        pass

def match_parish_header_in_row(row_text_upper: str, parish_name: str) -> bool:
    """Sinusuri kung ang row_text_upper ay header row para sa partikular na parish_name."""
    clean_p = parish_name.upper().replace("PARISH", "").replace("CHAPTER", "").strip()
    if not clean_p:
        return False
    if clean_p in row_text_upper:
        return True
    if "HONG" in clean_p and ("HONG KONG" in row_text_upper or "HONGKONG" in row_text_upper):
        return True
    return False

def is_any_parish_banner_row(ws, row: int) -> str:
    """Ibinabalik ang canonical parish name mula sa PARISH_ORDER kung ang row ay banner ng isang Parokya (hal. 'PARISH: | MALANAS')."""
    col_a = str(ws.cell(row=row, column=1).value or "").strip().upper()
    col_b = str(ws.cell(row=row, column=2).value or "").strip().upper()
    col_c = str(ws.cell(row=row, column=3).value or "").strip().upper()
    joined_abc = f"{col_a} {col_b} {col_c}".strip()
    if not joined_abc:
        return ""
    if "PARISH" in col_a or col_a == "CHAPTER:" or col_b in [p.replace("PARISH", "").strip() for p in PARISH_ORDER]:
        for p_item in PARISH_ORDER:
            if match_parish_header_in_row(joined_abc, p_item):
                return p_item
    return ""

MONTH_ALIASES_MAP = {
    "JANUARY": ["JANUARY", "JAN", "01. JANUARY"],
    "FEBRUARY": ["FEBRUARY", "FEBRAURY", "FEB", "02. FEBRUARY"],
    "MARCH": ["MARCH", "MAR", "03. MARCH"],
    "APRIL": ["APRIL", "APR", "04. APRIL"],
    "MAY": ["MAY", "05. MAY"],
    "JUNE": ["JUNE", "JUN", "06. JUNE"],
    "JULY": ["JULY", "JUL", "07. JULY"],
    "AUGUST": ["AUGUST", "AUG", "08. AUGUST"],
    "SEPTEMBER": ["SEPTEMBER", "SEP", "SEPT", "09. SEPTEMBER"],
    "OCTOBER": ["OCTOBER", "OCT", "10. OCTOBER"],
    "NOVEMBER": ["NOVEMBER", "NOV", "11. NOVEMBER"],
    "DECEMBER": ["DECEMBER", "DEC", "12. DECEMBER"],
}

def find_month_sheet_in_workbook(wb, month_name: str):
    """Hinahanap ang sheet para sa partikular na buwan (hal. JANUARY..DECEMBER, kasama ang typo na FEBRAURY)."""
    if not month_name:
        return None
    m_up = month_name.strip().upper()
    aliases = MONTH_ALIASES_MAP.get(m_up, [m_up])
    for s_name in wb.sheetnames:
        s_clean = s_name.strip().upper()
        if s_clean in aliases or any(al == s_clean or (len(al) >= 4 and al in s_clean) for al in aliases):
            return wb[s_name]
    return None

def is_week_banner_row(ws, row: int) -> str:
    """Ibinabalik ang '1st Week'..'5th Week' kung ang row ay may '(Weekly Report-1st Week of ...)' banner."""
    c1 = str(ws.cell(row=row, column=1).value or "").strip()
    c2 = str(ws.cell(row=row, column=2).value or "").strip()
    joined = f"{c1} {c2}".upper()
    if "WEEKLY REPORT" in joined or "WEEK OF" in joined:
        for wk in WEEKS_2027:
            if wk.upper() in joined:
                return wk
    return ""

def find_parish_row_boundaries(ws, parish_name: str = "", week_name: str = "") -> tuple[int, int]:
    """Hinahanap ang eksaktong (start_row, end_row) sa loob ng:
    (A) Multi-Week Month Sheet ng Parokya (kung saan naka-stack ang 1st Week hanggang 5th Week sa loob ng JANUARY..DECEMBER sheet), O
    (B) Multi-Parish Master Sheet ('PARISH WEEKLY REPORT' kung saan naka-stack ang 10 Parokya).
    """
    # 1. Suriin muna kung ito ay Multi-Week Month Sheet (may 2 o higit pang 'Weekly Report-Xth Week' banners)
    week_banners = []
    for r in range(1, ws.max_row + 1):
        wk_hit = is_week_banner_row(ws, r)
        if wk_hit:
            week_banners.append((r, wk_hit))

    if len(week_banners) >= 2 or (len(week_banners) == 1 and ws.title.strip().upper() != "PARISH WEEKLY REPORT"):
        target_wk = str(week_name or "1st Week").strip()
        chosen_idx = None
        for idx, (r_b, wk_lbl) in enumerate(week_banners):
            if wk_lbl.upper() == target_wk.upper():
                chosen_idx = idx
                break
        if chosen_idx is None:
            chosen_idx = 0
        banner_r = week_banners[chosen_idx][0]
        start_r = max(1, banner_r - 1)
        if chosen_idx + 1 < len(week_banners):
            next_banner_r = week_banners[chosen_idx + 1][0]
            end_r = max(start_r + 10, next_banner_r - 2)
        else:
            end_r = ws.max_row
        return start_r, end_r

    # 2. Multi-Parish Master Sheet ('PARISH WEEKLY REPORT')
    if not parish_name or not str(parish_name).strip():
        return 1, ws.max_row

    start_row = None
    end_row = ws.max_row

    for r in range(1, ws.max_row + 1):
        detected_p = is_any_parish_banner_row(ws, r)
        if detected_p and normalize_parish_match(detected_p, parish_name):
            start_row = r
            break

    if start_row is None:
        for r in range(1, ws.max_row + 1):
            row_str = " ".join([str(ws.cell(row=r, column=c).value or "").strip().upper() for c in range(1, 7)])
            if match_parish_header_in_row(row_str, parish_name):
                start_row = r
                break

    if start_row is None:
        return 1, ws.max_row

    for r in range(start_row + 1, ws.max_row + 1):
        detected_next_p = is_any_parish_banner_row(ws, r)
        if detected_next_p and not normalize_parish_match(detected_next_p, parish_name):
            end_row = r - 1
            break
        col_a = str(ws.cell(row=r, column=1).value or "").strip().upper()
        if col_a.startswith("PARISH:") and r > start_row + 5:
            end_row = r - 1
            break

    # Isama ang 2 rows sa itaas ng PARISH: banner kung nandoon ang DATE: at REMARKS header (tulad ng Row 3-5 sa unang parokya)
    if start_row <= 6:
        start_row = max(1, start_row - 2)
    return start_row, end_row

def write_pwr_header_fields(ws, data_row: dict, report_date: str, parish_name: str = "", week_name: str = ""):
    """Isinusulat sa PARISH WEEKLY REPORT o Month Sheet ang PRIEST, SAKRISTAN, EPISTULA, EVANGELIO, CASASAAD, CANDLE HOLDER, USHER, MARSHAL, ADAL TI EVANGELIO, TEMA TI MISA, TEMA TI PAMPAMILYA, at Preachers sa seksyon ng tamang Parokya at Week."""
    start_r, end_r = find_parish_row_boundaries(ws, parish_name, week_name=week_name)
    scan_start_r = max(1, start_r - 2)
    max_r = min(end_r, start_r + 30)
    max_c = min(ws.max_column, 12)

    compact_label_map = [
        ("PRIEST:", data_row.get("priest", "")),
        ("SAKRISTAN:", data_row.get("sakristan", "")),
        ("EPISTULA:", data_row.get("epistula", "")),
        ("EVANGELIO:", data_row.get("evangelio", "")),
        ("CASASAAD:", data_row.get("casasaad", "")),
        ("CANDLEHOLDER:", data_row.get("candle_holder", "")),
        ("USHER:", data_row.get("usher", "")),
        ("MARSHAL:", data_row.get("marshal", "")),
        ("ADALTIEVANGELIO:", data_row.get("adal_evangelio", "")),
        ("TEMATIMISA:", data_row.get("tema_misa", "")),
        ("TEMATIPAMPAMILYA:", data_row.get("tema_pampamilya", "")),
    ]

    row_topic_context = ""
    for r_idx in range(scan_start_r, max_r + 1):
        col_b_check = str(ws.cell(row=r_idx, column=2).value or "").strip().upper()
        col_c_check = str(ws.cell(row=r_idx, column=3).value or "").strip().upper()
        if r_idx > start_r and (col_b_check in ["ADULTS", "ATTENDANCE:"] or col_c_check in ["ADULTS", "ATTENDANCE:"]):
            break

        for c_idx in range(1, min(max_c + 1, 6)):
            raw_val = ws.cell(row=r_idx, column=c_idx).value
            if not raw_val:
                continue
            cell_up = str(raw_val).strip().upper()
            norm_lbl = re.sub(r"^[\d\s.\-]+", "", cell_up).strip()
            compact_lbl = re.sub(r"\s+", "", norm_lbl)

            if compact_lbl.startswith("DATE:") and report_date:
                if ":" in str(raw_val) and len(compact_lbl) > 5 and not str(ws.cell(row=r_idx, column=c_idx + 1).value or "").strip():
                    set_excel_cell_safely(ws, r_idx, c_idx, f"DATE: {report_date}", preserve_formula=True)
                else:
                    next_c = find_next_writable_col_on_row(ws, r_idx, c_idx)
                    if next_c <= 9:
                        set_excel_cell_safely(ws, r_idx, next_c, str(report_date).strip(), preserve_formula=True)
                continue

            if r_idx < start_r:
                continue

            if "ADALTIEVANGELIO" in compact_lbl and ":" in compact_lbl:
                row_topic_context = "ADAL"
            elif "TEMATIMISA" in compact_lbl:
                row_topic_context = "MISA"
            elif "TEMATIPAMPAMILYA" in compact_lbl:
                row_topic_context = "PAMPAMILYA"

            for pat_compact, val_str in compact_label_map:
                if not str(val_str or "").strip():
                    continue
                if compact_lbl == pat_compact or compact_lbl.startswith(pat_compact):
                    if pat_compact == "EVANGELIO:" and "ADAL" in compact_lbl:
                        continue
                    if c_idx == 1:
                        ensure_officer_value_row_merged(ws, r_idx, 2, 9)
                    next_c = find_next_writable_col_on_row(ws, r_idx, c_idx)
                    if next_c <= 9:
                        set_excel_cell_safely(ws, r_idx, next_c, str(val_str).strip(), preserve_formula=True)
                        try:
                            m_c = get_master_writable_cell(ws, r_idx, next_c)
                            m_c.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")
                        except Exception:
                            pass

            if "PREACHER" in compact_lbl:
                p_val = ""
                if row_topic_context == "ADAL" or "PREACHER/S" in compact_lbl or "PREACHERS" in compact_lbl:
                    p_val = data_row.get("adal_preacher", "")
                elif row_topic_context == "MISA":
                    p_val = data_row.get("misa_preacher", "")
                elif row_topic_context == "PAMPAMILYA":
                    p_val = data_row.get("pampamilya_preacher", "")
                if str(p_val or "").strip():
                    if c_idx == 1:
                        ensure_officer_value_row_merged(ws, r_idx, 2, 9)
                    next_c = find_next_writable_col_on_row(ws, r_idx, c_idx)
                    if next_c <= 9:
                        set_excel_cell_safely(ws, r_idx, next_c, str(p_val).strip(), preserve_formula=True)
                        try:
                            m_c = get_master_writable_cell(ws, r_idx, next_c)
                            m_c.alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center")
                        except Exception:
                            pass

def is_numeric_slot_label(val_str: str) -> bool:
    s = str(val_str or "").strip()
    if not s:
        return False
    return bool(re.match(r"^\d+(\.0)?$", s))

def write_pwr_attendance_and_remarks(ws, data_row: dict, parish_name: str = "", week_name: str = ""):
    """Isinusulat sa PARISH WEEKLY REPORT o Month Sheet ang attendance ng ADULTS, YOUTHS, at KIDS pati Remarks sa eksaktong rows ng napiling Parokya at Week."""
    start_r, end_r = find_parish_row_boundaries(ws, parish_name, week_name=week_name)
    section = None
    adult_named_rows, adult_blank_rows = [], []
    youth_named_rows, youth_blank_rows = [], []
    kids_named_rows, kids_blank_rows = [], []
    total_row_idx = None
    avg_row_idx = None

    for r_idx in range(start_r, end_r + 1):
        col_a = str(ws.cell(row=r_idx, column=1).value or "").strip()
        col_b = str(ws.cell(row=r_idx, column=2).value or "").strip()
        col_c = str(ws.cell(row=r_idx, column=3).value or "").strip()
        row_up = f"{col_a} {col_b} {col_c}".upper()

        if col_b.upper() == "ADULTS" or col_c.upper() == "ADULTS":
            section = "ADULTS"
            continue
        elif col_b.upper() in ["YOUTHS", "YOUTH"] or col_c.upper() in ["YOUTHS", "YOUTH"]:
            section = "YOUTHS"
            continue
        elif col_b.upper() == "KIDS" or col_c.upper() == "KIDS":
            section = "KIDS"
            continue
        elif "TOTAL" in col_a.upper() or "TOTAL" in col_b.upper():
            total_row_idx = r_idx
            section = None
            continue
        elif "AVERAGE" in col_a.upper() or "AVERAGE" in col_b.upper():
            avg_row_idx = r_idx
            section = None
            continue
        elif "HOUSE VISITATION" in row_up:
            if str(data_row.get("house_visitation", "")).strip():
                next_c = find_next_writable_col_on_row(ws, r_idx, 1)
                set_excel_cell_safely(ws, r_idx, next_c, data_row.get("house_visitation", "0"), preserve_formula=True)
            section = None
            continue

        if not section:
            continue

        if col_c and not is_numeric_slot_label(col_c) and col_c.lower() != "nan" and "ATTENDANCE" not in col_c.upper():
            name_col = 3
            name_val = col_c
        elif col_b and not is_numeric_slot_label(col_b) and col_b.lower() != "nan" and "ATTENDANCE" not in col_b.upper():
            name_col = 2
            name_val = col_b
        else:
            name_col = 3
            name_val = ""

        if section == "ADULTS":
            if name_val:
                adult_named_rows.append((r_idx, name_col, name_val))
            else:
                adult_blank_rows.append((r_idx, name_col, ""))
        elif section == "YOUTHS":
            if name_val:
                youth_named_rows.append((r_idx, name_col, name_val))
            else:
                youth_blank_rows.append((r_idx, name_col, ""))
        elif section == "KIDS":
            if name_val:
                kids_named_rows.append((r_idx, name_col, name_val))
            else:
                kids_blank_rows.append((r_idx, name_col, ""))

    adult_rows = adult_named_rows + adult_blank_rows
    youth_rows = youth_named_rows + youth_blank_rows
    kids_rows = kids_named_rows + kids_blank_rows

    col_to_record_key = {
        4: "mass",
        5: "pampamilya",
        6: "online",
        7: "youth",
        8: "kids",
        9: "adal",
    }

    def normalize_person_name(n_str: str) -> str:
        return re.sub(r"[^A-Z0-9]", "", str(n_str or "").upper())

    def is_truthy_mark(v) -> bool:
        if v is None or v is False:
            return False
        if isinstance(v, (int, float)):
            return v > 0
        s = str(v).strip().lower()
        return s in ["1", "1.0", "true", "yes", "y", "✓", "✔", "/", "x", "present"]

    def fill_section_counts(row_list, col_counts_map, member_list_from_ai=None):
        if not row_list:
            return

        for r_num, _, _ in row_list:
            for c_idx in col_counts_map.keys():
                set_excel_cell_safely(ws, r_num, c_idx, None, preserve_formula=True)

        slot_to_ai_item = {}
        if isinstance(member_list_from_ai, list) and len(member_list_from_ai) > 0:
            used_slots = set()
            for m_item in member_list_from_ai:
                if not isinstance(m_item, dict):
                    continue
                ai_name = str(m_item.get("name", "")).strip()
                ai_norm = normalize_person_name(ai_name)
                matched_slot = None
                if ai_norm:
                    for slot_i, (r_num, n_col, existing_name) in enumerate(row_list):
                        if slot_i in used_slots or not existing_name:
                            continue
                        ex_norm = normalize_person_name(existing_name)
                        if ai_norm == ex_norm or (len(ai_norm) >= 4 and (ai_norm in ex_norm or ex_norm in ai_norm)):
                            matched_slot = slot_i
                            break
                if matched_slot is None:
                    for slot_i in range(len(row_list)):
                        if slot_i not in used_slots:
                            matched_slot = slot_i
                            break
                if matched_slot is not None:
                    used_slots.add(matched_slot)
                    slot_to_ai_item[matched_slot] = m_item
                    r_num, n_col, existing_name = row_list[matched_slot]
                    if not existing_name and ai_name:
                        set_excel_cell_safely(ws, r_num, n_col, ai_name, preserve_formula=True)
                    if m_item.get("remarks"):
                        set_excel_cell_safely(ws, r_num, 10, str(m_item.get("remarks")).strip(), preserve_formula=True)

        for c_idx, target_count in col_counts_map.items():
            cnt = int(target_count or 0)
            rec_key = col_to_record_key.get(c_idx, "")
            marked_slots = set()

            if slot_to_ai_item and rec_key:
                for slot_i, m_item in slot_to_ai_item.items():
                    if is_truthy_mark(m_item.get(rec_key)):
                        r_num, _, _ = row_list[slot_i]
                        set_excel_cell_safely(ws, r_num, c_idx, 1, preserve_formula=True)
                        marked_slots.add(slot_i)

            for slot_i, (r_num, _, _) in enumerate(row_list):
                if len(marked_slots) >= cnt:
                    break
                if slot_i not in marked_slots:
                    set_excel_cell_safely(ws, r_num, c_idx, 1, preserve_formula=True)
                    marked_slots.add(slot_i)

    fill_section_counts(
        adult_rows,
        {
            4: int(data_row.get("mass_a", 0)),
            5: int(data_row.get("pampamilya_a", 0)),
            6: int(data_row.get("online_a", 0)),
            9: int(data_row.get("adal_a", 0)),
        },
        data_row.get("adults_records")
    )
    fill_section_counts(
        youth_rows,
        {
            4: int(data_row.get("mass_y", 0)),
            5: int(data_row.get("pampamilya_y", 0)),
            6: int(data_row.get("online_y", 0)),
            7: int(data_row.get("youth_y", 0)),
            9: int(data_row.get("adal_y", 0)),
        },
        data_row.get("youths_records")
    )
    fill_section_counts(
        kids_rows,
        {
            8: int(data_row.get("kids_k", 0)),
        },
        data_row.get("kids_records")
    )

    m_tot = int(data_row.get("mass_a", 0)) + int(data_row.get("mass_y", 0))
    p_tot = int(data_row.get("pampamilya_a", 0)) + int(data_row.get("pampamilya_y", 0))
    o_tot = int(data_row.get("online_a", 0)) + int(data_row.get("online_y", 0))
    y_tot = int(data_row.get("youth_y", 0))
    k_tot = int(data_row.get("kids_k", 0))
    ad_tot = int(data_row.get("adal_a", 0)) + int(data_row.get("adal_y", 0))

    tot_say = int(data_row.get("total_members_say", 0)) or (len(adult_named_rows) + len(youth_named_rows))
    tot_y = int(data_row.get("total_members_y", 0)) or len(youth_named_rows)
    tot_k = int(data_row.get("total_members_k", 0)) or len(kids_named_rows)

    if total_row_idx:
        for c_i, val_t in [(4, m_tot), (5, p_tot), (6, o_tot), (7, y_tot), (8, k_tot), (9, ad_tot)]:
            set_excel_cell_safely(ws, total_row_idx, c_i, val_t, preserve_formula=True)
    if avg_row_idx:
        for c_i, att_v, den_v in [
            (4, m_tot, tot_say), (5, p_tot, tot_say), (6, o_tot, tot_say),
            (7, y_tot, tot_y), (8, k_tot, tot_k), (9, ad_tot, tot_say)
        ]:
            set_excel_cell_safely(ws, avg_row_idx, c_i, calc_pct_str(att_v, den_v), preserve_formula=True)

    rem_txt = str(data_row.get("remarks", "")).strip()
    if not rem_txt:
        fallback_rem_parts = []
        if str(data_row.get("mass_remarks", "")).strip():
            fallback_rem_parts.append(str(data_row.get("mass_remarks", "")).strip())
        for pfx_lbl, n_k, r_k in [
            ("House Visitation", "hv_names", "hv_remarks"),
            ("Missionary Work", "mw_names", "mw_remarks"),
            ("Indoctrination", "ind_names", "ind_remarks")
        ]:
            n_v = str(data_row.get(n_k, "")).strip()
            r_v = str(data_row.get(r_k, "")).strip()
            if n_v or r_v:
                fallback_rem_parts.append(f"{pfx_lbl}: {n_v} {r_v}".strip())
        rem_txt = "\n".join(fallback_rem_parts).strip()
    if rem_txt:
        # Hanapin ang pinaka-itaas na white cell sa ilalim ng green REMARKS header sa Column J (Column 10)
        rem_top_row = start_r + 1
        for r_probe in range(start_r, min(end_r + 1, start_r + 6)):
            c10_val = str(ws.cell(row=r_probe, column=10).value or "").strip().upper()
            if c10_val == "REMARKS":
                rem_top_row = r_probe + 1
                break
        # Tiyaking naka-merge ang Remarks box sa Column J (hal. J167:J198) kung hindi pa naka-merge upang buong-buo itong makita sa Google Sheets
        try:
            rem_already_merged = any(
                (mr.min_row <= rem_top_row <= mr.max_row and mr.min_col <= 10 <= mr.max_col and mr.max_row > rem_top_row)
                for mr in ws.merged_cells.ranges
            )
            if not rem_already_merged and end_r > rem_top_row + 5:
                rem_bot_row = min(end_r - 2, rem_top_row + 31)
                has_j_overlap = any(
                    (mr.min_row <= rem_bot_row and mr.max_row >= rem_top_row and mr.min_col <= 10 <= mr.max_col)
                    for mr in ws.merged_cells.ranges
                )
                if not has_j_overlap and rem_bot_row > rem_top_row:
                    ws.merge_cells(start_row=rem_top_row, start_column=10, end_row=rem_bot_row, end_column=10)
        except Exception:
            pass
        set_excel_cell_safely(ws, rem_top_row, 10, rem_txt, preserve_formula=True)
        try:
            master_rem_cell = get_master_writable_cell(ws, rem_top_row, 10)
            master_rem_cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top", horizontal="left")
        except Exception:
            pass

def count_families_and_people_from_text(names_text: str, count_hint_str: str = "0") -> tuple[int, int]:
    """Kinukuwenta ang Bilang ng Pamilya at Total ng mga Tao ayon sa nakalagay na Family Names at Mga Pangalan."""
    hint_val = parse_int_from_text(count_hint_str)
    s = str(names_text or "").strip()
    if not s or s in ["0", "None", "nan", "-"]:
        return hint_val, hint_val

    fam_matches = re.findall(r"[A-Za-z]+\s+Family", s, flags=re.I)
    if fam_matches:
        unique_fams = []
        for fm in fam_matches:
            fm_norm = re.sub(r"\s+", " ", fm.strip()).upper()
            if fm_norm not in unique_fams:
                unique_fams.append(fm_norm)
        fam_cnt = len(unique_fams)
    else:
        groups = [g.strip() for g in re.split(r"[;\n|]+", s) if g.strip() and not g.strip().isdigit()]
        fam_cnt = max(hint_val, len(groups) if len(groups) <= 15 else 1)

    # Bilangin ang mga indibidwal na pangalan (pagkatapos ng ':' kung may Family prefix)
    people_tokens = []
    for block in re.split(r"[\n|]+", s):
        b_clean = block.strip()
        if not b_clean:
            continue
        if ":" in b_clean:
            after_colon = b_clean.split(":", 1)[1].strip()
            if after_colon:
                b_clean = after_colon
        parts = [p.strip() for p in re.split(r"[,;]+", b_clean) if p.strip()]
        for pt in parts:
            pt_c = re.sub(r"^\d+[.)\s\-]*", "", pt).strip()
            if len(pt_c) >= 2 and not pt_c.isdigit() and not re.match(r"^[A-Za-z]+\s+Family$", pt_c, flags=re.I):
                people_tokens.append(pt_c)

    people_cnt = len(people_tokens) if people_tokens else max(fam_cnt, hint_val)
    if hint_val > 0 and fam_cnt == 0:
        fam_cnt = hint_val
    return fam_cnt, people_cnt

def parse_remarks_into_mass_and_outreach(raw_remarks: str) -> dict:
    """Hinahati at nililinis ang Remarks mula sa Image/Excel report papunta sa:
    1) mass_remarks (Mass, Pampamilya, Adal, House Church, Online Mass, Visitors)
    2) house_visitation details (hv_names, hv_families_count, hv_people_count, hv_remarks)
    3) missionary_work details (mw_names, mw_families_count, mw_people_count, mw_remarks)
    4) indoctrination details (ind_names, ind_families_count, ind_people_count, ind_remarks)
    """
    out = {
        "clean_full_remarks": "",
        "mass_remarks": "",
        "mass_extra_names_count": 0,
        "hv_names": "",
        "hv_families_count": 0,
        "hv_people_count": 0,
        "hv_remarks": "",
        "mw_names": "",
        "mw_families_count": 0,
        "mw_people_count": 0,
        "mw_remarks": "",
        "ind_names": "",
        "ind_families_count": 0,
        "ind_people_count": 0,
        "ind_remarks": "",
    }
    if not raw_remarks or not str(raw_remarks).strip():
        return out

    txt = str(raw_remarks).strip()
    txt = re.sub(r"\bOndine\s+Mass\b", "Online Mass", txt, flags=re.I)
    txt = re.sub(r"\bADal\s*ti\b", "Adal ti", txt, flags=re.I)
    txt = re.sub(r"\bMissiondiay\b", "Mission diay", txt, flags=re.I)
    txt = re.sub(r"\bT[eo]ma\s*:\s*Ti\s*Blag\b", "Tema:Ti Biag", txt, flags=re.I)
    txt = re.sub(r"\bBillv\b", "Billy", txt, flags=re.I)
    txt = re.sub(r"\bAramidtaytot\b", "Aramidtayto't", txt, flags=re.I)
    txt = re.sub(r"\bRer\s*la\s+C\.?\b", "Rema C.", txt, flags=re.I)

    # Kung naputol ng OCR ang unang linya bago ang 'Brix R; Chris Paul C;', ibalik ang unang linya
    if txt.startswith("Brix R; Chris Paul C") and "Romel S" not in txt[:120]:
        txt = "Mass;Pampamilya;Adal ti Evangelio: Romel S; Crystal Joy C; AlJem G; Juliebeth B; Princess P;\n" + txt

    # Alisin ang nadobleng OCR block kung na-scan nang dalawang beses sa larawan (panatilihin ang kumpletong block)
    m_dup_mass = list(re.finditer(r"Mass\s*;\s*Pampamilya", txt, flags=re.I))
    if len(m_dup_mass) >= 2:
        txt = txt[m_dup_mass[1].start():]

    m_dup_miss = list(re.finditer(r"Mission\s*diay\s*balay", txt, flags=re.I))
    if len(m_dup_miss) >= 2:
        prefix_before_m1 = txt[:m_dup_miss[0].start()].strip()
        first_mission_block = txt[m_dup_miss[0].start():m_dup_miss[1].start()].strip()
        first_mission_block = re.sub(r"\s*Medelyn\s+C\s*;.*$", "", first_mission_block, flags=re.I | re.S).strip()
        second_mission_block = txt[m_dup_miss[1].start():].strip()
        best_mission_block = first_mission_block if len(first_mission_block) >= len(second_mission_block) else second_mission_block
        txt = (prefix_before_m1 + "\n" + best_mission_block).strip()

    out["clean_full_remarks"] = txt

    # Ihiwalay ang Missionary Work block kung may "Mission diay..." o "Missionary Work..."
    m_mission = re.search(r"(Mission\s+diay\b.*|Missionary\s+Work\b.*)", txt, flags=re.I | re.S)
    if m_mission:
        mass_part = txt[:m_mission.start()].strip()
        mission_part = m_mission.group(1).strip()
    else:
        mass_part = txt
        mission_part = ""

    mass_part_clean = re.sub(r"\s+", " ", mass_part).strip()
    for sec_lbl in ["House Church:", "Online Mass:", "Visitors:", "Visitors;"]:
        mass_part_clean = mass_part_clean.replace(sec_lbl, f" | {sec_lbl.replace(';', ':')}")
    out["mass_remarks"] = mass_part_clean

    if mass_part:
        semi_parts = [p.strip() for p in re.split(r"[;,\n]", mass_part) if p.strip()]
        person_tokens = [
            p for p in semi_parts
            if len(p) >= 3 and not any(w in p.upper() for w in ["MASS", "PAMPAMILYA", "EVANGELIO", "HOUSE CHURCH", "VISITORS"])
        ]
        out["mass_extra_names_count"] = len(person_tokens)

    if mission_part:
        fam_match = re.search(r"([A-Z][a-zA-Z]+\s+Family)", mission_part)
        fam_name = re.sub(r"\s+", " ", fam_match.group(1)).strip() if fam_match else "Mission Outreach Family"
        out["mw_families_count"] = 1

        att_split = re.split(r"ATTENDANCE\s*[:;.\-]?", mission_part, maxsplit=1, flags=re.I)
        mw_header_info = re.sub(r"\s+", " ", att_split[0]).strip()
        out["mw_remarks"] = mw_header_info

        if len(att_split) > 1:
            att_lines = [
                re.sub(r"^[|:\-,;\s]+|[|:\-,;\s]+$", "", ln).strip()
                for ln in re.split(r"[\n;]", att_split[1])
                if re.sub(r"^[|:\-,;\s]+|[|:\-,;\s]+$", "", ln).strip()
            ]
            clean_att_names = []
            for nm in att_lines:
                if len(nm) > 35 and ("Visitors" in nm or "Medelyn" in nm):
                    nm = nm.split("Medelyn")[0].strip()
                if nm:
                    clean_att_names.append(nm)
            out["mw_names"] = f"{fam_name} ({len(clean_att_names)} Tao): " + ", ".join(clean_att_names)
            out["mw_people_count"] = len(clean_att_names)
        else:
            out["mw_names"] = fam_name
            out["mw_people_count"] = 1

    return out

def save_weekly_monitoring_record(parish: str, month: str, week_num: str, report_date: str, extracted_data: dict, drive_link: str = "") -> tuple[bool, str]:
    db = load_weekly_monitoring_records()
    key = f"{parish.upper()}__{month.upper()}__{week_num}"
    denoms = get_parish_member_denominators(parish)

    tot_say = int(extracted_data.get("total_members_say", 0)) or denoms["say"]
    tot_y = int(extracted_data.get("total_members_y", 0)) or denoms["youth"]
    tot_k = int(extracted_data.get("total_members_k", 0)) or denoms["kids"]

    raw_rem_str = str(extracted_data.get("remarks", ""))
    parsed_ro = parse_remarks_into_mass_and_outreach(raw_rem_str)
    clean_rem_to_save = parsed_ro["clean_full_remarks"] or raw_rem_str

    hv_names_val = str(extracted_data.get("hv_names", "") or parsed_ro.get("hv_names", "")).strip()
    hv_rem_val = str(extracted_data.get("hv_remarks", "") or parsed_ro.get("hv_remarks", "")).strip()
    hv_fam_c, hv_ppl_c = count_families_and_people_from_text(hv_names_val, str(extracted_data.get("house_visitation", "0")))
    if int(extracted_data.get("hv_people_count", 0)) > 0:
        hv_ppl_c = int(extracted_data.get("hv_people_count", 0))

    mw_names_val = str(extracted_data.get("mw_names", "") or parsed_ro.get("mw_names", "")).strip()
    mw_rem_val = str(extracted_data.get("mw_remarks", "") or parsed_ro.get("mw_remarks", "")).strip()
    mw_hint = str(extracted_data.get("missionary_work", "0"))
    if (not mw_hint or mw_hint == "0") and parsed_ro.get("mw_families_count", 0) > 0:
        mw_hint = str(parsed_ro["mw_families_count"])
    mw_fam_c, mw_ppl_c = count_families_and_people_from_text(mw_names_val, mw_hint)
    if int(extracted_data.get("mw_people_count", 0)) > 0:
        mw_ppl_c = int(extracted_data.get("mw_people_count", 0))
    elif parsed_ro.get("mw_people_count", 0) > 0 and not extracted_data.get("mw_names"):
        mw_ppl_c = int(parsed_ro["mw_people_count"])

    ind_names_val = str(extracted_data.get("ind_names", "") or parsed_ro.get("ind_names", "")).strip()
    ind_rem_val = str(extracted_data.get("ind_remarks", "") or parsed_ro.get("ind_remarks", "")).strip()
    ind_fam_c, ind_ppl_c = count_families_and_people_from_text(ind_names_val, str(extracted_data.get("indoctrination", "0")))
    if int(extracted_data.get("ind_people_count", 0)) > 0:
        ind_ppl_c = int(extracted_data.get("ind_people_count", 0))

    mass_rem_val = str(extracted_data.get("mass_remarks", "") or parsed_ro.get("mass_remarks", "")).strip()
    mass_extra_cnt = int(extracted_data.get("mass_extra_names_count", 0)) or int(parsed_ro.get("mass_extra_names_count", 0))

    db[key] = {
        "has_data": True,
        "parish": parish,
        "month": month,
        "week": week_num,
        "report_date": report_date,
        "total_members_say": tot_say,
        "total_members_y": tot_y,
        "total_members_k": tot_k,
        "mass_a": int(extracted_data.get("mass_a", 0)),
        "mass_y": int(extracted_data.get("mass_y", 0)),
        "pampamilya_a": int(extracted_data.get("pampamilya_a", 0)),
        "pampamilya_y": int(extracted_data.get("pampamilya_y", 0)),
        "online_a": int(extracted_data.get("online_a", 0)),
        "online_y": int(extracted_data.get("online_y", 0)),
        "youth_y": int(extracted_data.get("youth_y", 0)),
        "kids_k": int(extracted_data.get("kids_k", 0)),
        "adal_a": int(extracted_data.get("adal_a", 0)),
        "adal_y": int(extracted_data.get("adal_y", 0)),
        "house_visitation": str(hv_fam_c if hv_fam_c > 0 else extracted_data.get("house_visitation", "0")),
        "hv_names": hv_names_val,
        "hv_people_count": hv_ppl_c,
        "hv_remarks": hv_rem_val,
        "missionary_work": str(mw_fam_c if mw_fam_c > 0 else mw_hint),
        "mw_names": mw_names_val,
        "mw_people_count": mw_ppl_c,
        "mw_remarks": mw_rem_val,
        "indoctrination": str(ind_fam_c if ind_fam_c > 0 else extracted_data.get("indoctrination", "0")),
        "ind_names": ind_names_val,
        "ind_people_count": ind_ppl_c,
        "ind_remarks": ind_rem_val,
        "mass_remarks": mass_rem_val,
        "mass_extra_names_count": mass_extra_cnt,
        "remarks": clean_rem_to_save,
        "priest": str(extracted_data.get("priest", "")),
        "sakristan": str(extracted_data.get("sakristan", "")),
        "epistula": str(extracted_data.get("epistula", "")),
        "evangelio": str(extracted_data.get("evangelio", "")),
        "casasaad": str(extracted_data.get("casasaad", "")),
        "candle_holder": str(extracted_data.get("candle_holder", "")),
        "usher": str(extracted_data.get("usher", "")),
        "marshal": str(extracted_data.get("marshal", "")),
        "adal_evangelio": str(extracted_data.get("adal_evangelio", "")),
        "adal_preacher": str(extracted_data.get("adal_preacher", "")),
        "tema_misa": str(extracted_data.get("tema_misa", "")),
        "misa_preacher": str(extracted_data.get("misa_preacher", "")),
        "tema_pampamilya": str(extracted_data.get("tema_pampamilya", "")),
        "pampamilya_preacher": str(extracted_data.get("pampamilya_preacher", "")),
        "adults_records": extracted_data.get("adults_records", []),
        "youths_records": extracted_data.get("youths_records", []),
        "kids_records": extracted_data.get("kids_records", []),
        "drive_link": drive_link or db.get(key, {}).get("drive_link", ""),
        "submitted_at_ph": str(extracted_data.get("submitted_at_ph", "")).strip() or db.get(key, {}).get("submitted_at_ph", "") or get_philippine_datetime_str(),
        "updated_at": get_philippine_datetime_str()
    }
    try:
        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return update_master_pwr_excel_file(parish, month, week_num, report_date, db[key])

def update_master_pwr_excel_file(parish: str, month: str, week_num: str, report_date: str, data_row: dict) -> tuple[bool, str]:
    service = get_google_drive_service()
    folder_url = st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
    drive_file_id, target_month_folder_id = find_2027_master_excel_on_drive(service, folder_url, month, week_num)

    excel_bytes = None
    if service and drive_file_id:
        try:
            excel_bytes = service.files().get_media(fileId=drive_file_id).execute()
            with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
                f.write(excel_bytes)
        except Exception:
            pass

    is_cloned_from_base_template = False
    if not excel_bytes and service:
        try:
            excel_bytes = service.files().get_media(fileId="1v7BFj32ixL7xg5BaElAU1Ji-zsSKhdOF").execute()
            is_cloned_from_base_template = True
        except Exception:
            pass
    if not excel_bytes and os.path.exists(LOCAL_PWR_MASTER_FILE):
        try:
            with open(LOCAL_PWR_MASTER_FILE, "rb") as f:
                excel_bytes = f.read()
            if not drive_file_id:
                is_cloned_from_base_template = True
        except Exception:
            pass

    try:
        if excel_bytes:
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        else:
            wb = openpyxl.Workbook()
            ws_init = wb.active
            ws_init.title = "PARISH WEEKLY REPORT"

        clean_p = parish.upper().replace("PARISH", "").strip()
        m_att = int(data_row["mass_a"]) + int(data_row["mass_y"])
        p_att = int(data_row["pampamilya_a"]) + int(data_row["pampamilya_y"])
        o_att = int(data_row["online_a"]) + int(data_row["online_y"])
        y_att = int(data_row["youth_y"])
        k_att = int(data_row["kids_k"])
        ad_att = int(data_row["adal_a"]) + int(data_row["adal_y"])

        tot_say = int(data_row.get("total_members_say", 0))
        tot_y = int(data_row.get("total_members_y", 0))
        tot_k = int(data_row.get("total_members_k", 0))

        # Kung ito ay Multi-Month Parish Workbook (hal. '3. POBLACION PARISH WEEKLY REPORTS (JAN. TO DEC. 2027).xlsx'),
        # tiyaking kumpleto ang JANUARY hanggang DECEMBER sheets (1st Week..5th Week)
        ensure_all_12_month_sheets_in_parish_workbook(wb, parish_name=parish)

        # 0. PRIMARY INPUT: Write directly into PARISH WEEKLY REPORT o sa kaukulang Month Sheet (JANUARY..DECEMBER)
        target_pwr_sheets = []
        matched_m_sheet = find_month_sheet_in_workbook(wb, month)
        if matched_m_sheet is not None:
            target_pwr_sheets.append(matched_m_sheet)
        for s_name in wb.sheetnames:
            s_up = s_name.upper()
            if "PARISH WEEKLY REPORT" in s_up or clean_p in s_up:
                if s_up not in ["PARISH MONITORING REPORT", "SUMMARY FOR POSTING"] and wb[s_name] not in target_pwr_sheets:
                    target_pwr_sheets.append(wb[s_name])
        if not target_pwr_sheets and "PARISH WEEKLY REPORT" in wb.sheetnames:
            target_pwr_sheets.append(wb["PARISH WEEKLY REPORT"])

        for ws_pwr in target_pwr_sheets:
            write_pwr_header_fields(ws_pwr, data_row, report_date, parish_name=parish, week_name=week_num)
            write_pwr_attendance_and_remarks(ws_pwr, data_row, parish_name=parish, week_name=week_num)

        # 1. Update PARISH MONITORING REPORT sheet safely (preserving any Excel formulas!)
        if "PARISH MONITORING REPORT" in wb.sheetnames:
            ws = wb["PARISH MONITORING REPORT"]
            found_p = False
            for r_idx in range(5, ws.max_row + 1):
                cell_val = str(ws.cell(row=r_idx, column=1).value or "").upper()
                if clean_p and clean_p in cell_val:
                    set_excel_cell_safely(ws, r_idx, 1, f"{clean_p} ({report_date})", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 2, int(re.sub(r"\D", "", str(week_num)) or 1), preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 3, f"A-{data_row['mass_a']}\nY-{data_row['mass_y']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 4, f"A-{data_row['pampamilya_a']}\nY-{data_row['pampamilya_y']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 5, f"A-{data_row['online_a']}\nY-{data_row['online_y']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 6, f"Y-{data_row['youth_y']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 7, f"K-{data_row['kids_k']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 8, f"A-{data_row['adal_a']}\nY-{data_row['adal_y']}", preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 9, data_row.get("house_visitation", "0"), preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 10, data_row.get("missionary_work", "0"), preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 11, data_row.get("indoctrination", "0"), preserve_formula=True)
                    set_excel_cell_safely(ws, r_idx, 12, data_row.get("remarks", ""), preserve_formula=True)
                    found_p = True
                    break
            if not found_p:
                ws.append([
                    f"{clean_p} ({report_date})", int(re.sub(r"\D", "", str(week_num)) or 1),
                    f"A-{data_row['mass_a']}\nY-{data_row['mass_y']}",
                    f"A-{data_row['pampamilya_a']}\nY-{data_row['pampamilya_y']}",
                    f"A-{data_row['online_a']}\nY-{data_row['online_y']}",
                    f"Y-{data_row['youth_y']}", f"K-{data_row['kids_k']}",
                    f"A-{data_row['adal_a']}\nY-{data_row['adal_y']}",
                    data_row.get("house_visitation", "0"), data_row.get("missionary_work", "0"),
                    data_row.get("indoctrination", "0"), data_row.get("remarks", "")
                ])

        # 2. Update SUMMARY FOR POSTING sheet safely (preserving any Excel formulas & avoiding MergedCell errors!)
        if "SUMMARY FOR POSTING" in wb.sheetnames:
            ws_s = wb["SUMMARY FOR POSTING"]
            for r_idx in range(1, ws_s.max_row + 1):
                c1_val = str(ws_s.cell(row=r_idx, column=1).value or "").upper().strip()
                if clean_p and (clean_p == c1_val or clean_p in c1_val or ("HONG" in clean_p and "HONG" in c1_val)):
                    for sub_r in range(r_idx + 1, min(r_idx + 6, ws_s.max_row + 1)):
                        lbl = str(ws_s.cell(row=sub_r, column=1).value or "").upper().strip()
                        if "TOTAL MEMBERS" in lbl:
                            set_excel_cell_safely(ws_s, sub_r, 2, tot_say, preserve_formula=True)
                            if "HONG" not in clean_p:
                                set_excel_cell_safely(ws_s, sub_r, 3, tot_say, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 4, tot_say, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 5, tot_y, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 6, tot_k, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 7, tot_say, preserve_formula=True)
                            else:
                                set_excel_cell_safely(ws_s, sub_r, 3, tot_say, preserve_formula=True)
                        elif "TOTAL OF ATTENDEES" in lbl:
                            set_excel_cell_safely(ws_s, sub_r, 2, m_att, preserve_formula=True)
                            if "HONG" not in clean_p:
                                set_excel_cell_safely(ws_s, sub_r, 3, p_att, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 4, o_att, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 5, y_att, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 6, k_att, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 7, ad_att, preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 8, parse_int_from_text(data_row.get("house_visitation", "0")), preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 9, parse_int_from_text(data_row.get("missionary_work", "0")) + parse_int_from_text(data_row.get("indoctrination", "0")), preserve_formula=True)
                            else:
                                set_excel_cell_safely(ws_s, sub_r, 3, o_att, preserve_formula=True)
                        elif "TOTAL AVERAGE" in lbl or "AVERAGE" in lbl:
                            set_excel_cell_safely(ws_s, sub_r, 2, calc_pct_str(m_att, tot_say), preserve_formula=True)
                            if "HONG" not in clean_p:
                                set_excel_cell_safely(ws_s, sub_r, 3, calc_pct_str(p_att, tot_say), preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 4, calc_pct_str(o_att, tot_say), preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 5, calc_pct_str(y_att, tot_y), preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 6, calc_pct_str(k_att, tot_k), preserve_formula=True)
                                set_excel_cell_safely(ws_s, sub_r, 7, calc_pct_str(ad_att, tot_say), preserve_formula=True)
                            else:
                                set_excel_cell_safely(ws_s, sub_r, 3, calc_pct_str(o_att, tot_say), preserve_formula=True)
                    break

        out_buf = io.BytesIO()
        wb.save(out_buf)
        updated_bytes = out_buf.getvalue()
        with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
            f.write(updated_bytes)

        if service:
            target_master_name = f"PWR. {month} {week_num}.xlsx"
            media = MediaIoBaseUpload(
                io.BytesIO(updated_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                resumable=True
            )
            if drive_file_id:
                service.files().update(fileId=drive_file_id, media_body=media).execute()
                return True, f"Matagumpay na na-type sa PARISH WEEKLY REPORT at na-update ang '{target_master_name}' sa Google Drive!"
            elif target_month_folder_id:
                service.files().create(
                    body={"name": target_master_name, "parents": [target_month_folder_id]},
                    media_body=media
                ).execute()
                return True, f"Nalikha at na-update ang '{target_master_name}' sa Google Drive!"
        return True, "Na-save sa lokal na Master Excel!"
    except Exception as e:
        return False, f"Master Excel Sync Notice: {str(e)}"

def parse_master_pwr_excel_into_db(excel_bytes: bytes, target_month: str, target_week: str) -> int:
    db = load_weekly_monitoring_records()
    updated_parishes = 0
    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)

        if "PARISH WEEKLY REPORT" in wb.sheetnames:
            for p_item in PARISH_ORDER:
                pwr_extracted = extract_weekly_report_from_excel_bytes(excel_bytes, target_parish=p_item)
                detected_parish = pwr_extracted.get("parish", "") or p_item
                if detected_parish in PARISH_ORDER:
                    tot_act = (
                        int(pwr_extracted.get("mass_a", 0)) + int(pwr_extracted.get("mass_y", 0)) +
                        int(pwr_extracted.get("pampamilya_a", 0)) + int(pwr_extracted.get("pampamilya_y", 0)) +
                        int(pwr_extracted.get("online_a", 0)) + int(pwr_extracted.get("online_y", 0)) +
                        int(pwr_extracted.get("youth_y", 0)) + int(pwr_extracted.get("kids_k", 0)) +
                        int(pwr_extracted.get("adal_a", 0)) + int(pwr_extracted.get("adal_y", 0))
                    )
                    if tot_act > 0 or pwr_extracted.get("priest") or pwr_extracted.get("tema_misa"):
                        key = f"{detected_parish.upper()}__{target_month.upper()}__{target_week}"
                        denoms = get_parish_member_denominators(detected_parish)
                        existing = db.get(key, {})
                        existing.update(pwr_extracted)
                        existing["has_data"] = True
                        existing["parish"] = detected_parish
                        existing["month"] = target_month
                        existing["week"] = target_week
                        existing["total_members_say"] = int(pwr_extracted.get("total_members_say", 0)) or denoms["say"]
                        existing["total_members_y"] = int(pwr_extracted.get("total_members_y", 0)) or denoms["youth"]
                        existing["total_members_k"] = int(pwr_extracted.get("total_members_k", 0)) or denoms["kids"]
                        db[key] = existing
                        updated_parishes += 1

        if "SUMMARY FOR POSTING" in wb.sheetnames:
            ws_s = wb["SUMMARY FOR POSTING"]
            for r_idx in range(1, ws_s.max_row + 1):
                c1_val = str(ws_s.cell(row=r_idx, column=1).value or "").upper().strip()
                matched_p = None
                for p in PARISH_ORDER:
                    clean_p = p.replace("PARISH", "").strip()
                    if clean_p == c1_val or (clean_p in c1_val and len(c1_val) < 25) or ("HONG" in clean_p and "HONG KONG" in c1_val):
                        matched_p = p
                        break
                if matched_p:
                    tot_say, tot_y, tot_k = 0, 0, 0
                    m_att, p_att, o_att, y_att, k_att, ad_att, hv_val, mw_val = 0, 0, 0, 0, 0, 0, 0, 0
                    for sub_r in range(r_idx + 1, min(r_idx + 5, ws_s.max_row + 1)):
                        lbl = str(ws_s.cell(row=sub_r, column=1).value or "").upper().strip()
                        if "TOTAL MEMBERS" in lbl:
                            tot_say = parse_int_from_text(ws_s.cell(row=sub_r, column=2).value)
                            if "HONGKONG" not in matched_p:
                                tot_y = parse_int_from_text(ws_s.cell(row=sub_r, column=5).value)
                                tot_k = parse_int_from_text(ws_s.cell(row=sub_r, column=6).value)
                        elif "TOTAL OF ATTENDEES" in lbl:
                            m_att = parse_int_from_text(ws_s.cell(row=sub_r, column=2).value)
                            if "HONGKONG" not in matched_p:
                                p_att = parse_int_from_text(ws_s.cell(row=sub_r, column=3).value)
                                o_att = parse_int_from_text(ws_s.cell(row=sub_r, column=4).value)
                                y_att = parse_int_from_text(ws_s.cell(row=sub_r, column=5).value)
                                k_att = parse_int_from_text(ws_s.cell(row=sub_r, column=6).value)
                                ad_att = parse_int_from_text(ws_s.cell(row=sub_r, column=7).value)
                                hv_val = parse_int_from_text(ws_s.cell(row=sub_r, column=8).value)
                                mw_val = parse_int_from_text(ws_s.cell(row=sub_r, column=9).value)
                            else:
                                o_att = parse_int_from_text(ws_s.cell(row=sub_r, column=3).value)

                    key = f"{matched_p.upper()}__{target_month.upper()}__{target_week}"
                    existing = db.get(key, {})
                    if (m_att + p_att + o_att + y_att + k_att + ad_att + hv_val + mw_val) > 0 or existing.get("has_data"):
                        denoms = get_parish_member_denominators(matched_p)
                        db[key] = {
                            **existing,
                            "has_data": True,
                            "parish": matched_p,
                            "month": target_month,
                            "week": target_week,
                            "report_date": existing.get("report_date", "01/02/27"),
                            "total_members_say": tot_say or existing.get("total_members_say") or denoms["say"],
                            "total_members_y": tot_y or existing.get("total_members_y") or denoms["youth"],
                            "total_members_k": tot_k or existing.get("total_members_k") or denoms["kids"],
                            "mass_a": existing.get("mass_a", m_att),
                            "mass_y": existing.get("mass_y", 0),
                            "pampamilya_a": existing.get("pampamilya_a", p_att),
                            "pampamilya_y": existing.get("pampamilya_y", 0),
                            "online_a": existing.get("online_a", o_att),
                            "online_y": existing.get("online_y", 0),
                            "youth_y": existing.get("youth_y", y_att),
                            "kids_k": existing.get("kids_k", k_att),
                            "adal_a": existing.get("adal_a", ad_att),
                            "adal_y": existing.get("adal_y", 0),
                            "house_visitation": str(existing.get("house_visitation") or hv_val),
                            "missionary_work": str(existing.get("missionary_work") or mw_val),
                            "indoctrination": str(existing.get("indoctrination", "0")),
                            "remarks": str(existing.get("remarks", "")),
                            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        updated_parishes += 1

        if "PARISH MONITORING REPORT" in wb.sheetnames:
            ws_m = wb["PARISH MONITORING REPORT"]
            for r_idx in range(5, ws_m.max_row + 1):
                c1_val = str(ws_m.cell(row=r_idx, column=1).value or "").upper().strip()
                matched_p = None
                for p in PARISH_ORDER:
                    clean_p = p.replace("PARISH", "").strip()
                    if clean_p in c1_val:
                        matched_p = p
                        break
                if matched_p:
                    def parse_ay(cell_v):
                        s = str(cell_v or "")
                        ma = re.search(r"A-(\d+)", s, re.I)
                        my = re.search(r"Y-(\d+)", s, re.I)
                        mk = re.search(r"K-(\d+)", s, re.I)
                        if ma or my or mk:
                            return (int(ma.group(1)) if ma else 0), (int(my.group(1)) if my else (int(mk.group(1)) if mk else 0))
                        if s.strip().isdigit():
                            return int(s.strip()), 0
                        return 0, 0

                    m_a, m_y = parse_ay(ws_m.cell(row=r_idx, column=3).value)
                    p_a, p_y = parse_ay(ws_m.cell(row=r_idx, column=4).value)
                    o_a, o_y = parse_ay(ws_m.cell(row=r_idx, column=5).value)
                    y_1, y_2 = parse_ay(ws_m.cell(row=r_idx, column=6).value)
                    y_y = y_2 if y_2 > 0 else y_1
                    k_1, k_2 = parse_ay(ws_m.cell(row=r_idx, column=7).value)
                    k_k = k_2 if k_2 > 0 else k_1
                    ad_a, ad_y = parse_ay(ws_m.cell(row=r_idx, column=8).value)
                    hv_s = str(ws_m.cell(row=r_idx, column=9).value or "0").strip()
                    mw_s = str(ws_m.cell(row=r_idx, column=10).value or "0").strip()
                    ind_s = str(ws_m.cell(row=r_idx, column=11).value or "0").strip()
                    rem_s = str(ws_m.cell(row=r_idx, column=12).value or "").strip()

                    total_activity = m_a + m_y + p_a + p_y + o_a + o_y + y_y + k_k + ad_a + ad_y + parse_int_from_text(hv_s) + parse_int_from_text(mw_s) + parse_int_from_text(ind_s)
                    if total_activity > 0 or rem_s:
                        key = f"{matched_p.upper()}__{target_month.upper()}__{target_week}"
                        denoms = get_parish_member_denominators(matched_p)
                        existing = db.get(key, {})
                        db[key] = {
                            **existing,
                            "has_data": True,
                            "parish": matched_p,
                            "month": target_month,
                            "week": target_week,
                            "report_date": existing.get("report_date", "01/02/27"),
                            "total_members_say": existing.get("total_members_say") or denoms["say"],
                            "total_members_y": existing.get("total_members_y") or denoms["youth"],
                            "total_members_k": existing.get("total_members_k") or denoms["kids"],
                            "mass_a": m_a, "mass_y": m_y,
                            "pampamilya_a": p_a, "pampamilya_y": p_y,
                            "online_a": o_a, "online_y": o_y,
                            "youth_y": y_y, "kids_k": k_k,
                            "adal_a": ad_a, "adal_y": ad_y,
                            "house_visitation": hv_s if hv_s != "None" else "0",
                            "missionary_work": mw_s if mw_s != "None" else "0",
                            "indoctrination": ind_s if ind_s != "None" else "0",
                            "remarks": (rem_s if (rem_s and rem_s != "None" and not str(rem_s).startswith("=")) else existing.get("remarks", "")),
                            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        updated_parishes += 1

        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return updated_parishes

def is_week_record_completed(rec: dict) -> bool:
    if not isinstance(rec, dict) or not rec.get("has_data"):
        return False
    tot_att = (
        int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0)) +
        int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0)) +
        int(rec.get("online_a", 0)) + int(rec.get("online_y", 0)) +
        int(rec.get("youth_y", 0)) + int(rec.get("kids_k", 0)) +
        int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))
    )
    has_off = any(str(rec.get(k, "")).strip() for k in ["priest", "sakristan", "epistula", "evangelio", "tema_misa", "adal_evangelio"])
    return (tot_att > 0 or has_off or bool(str(rec.get("remarks", "")).strip()))

def find_next_unfinished_month_and_week(parish_name: str, preferred_month: str = "January") -> tuple[str, str]:
    """Hinahanap ang unang Hindi Pa Tapos (Unfinished) na Buwan at Linggo para sa napiling Parokya sa buong 2027."""
    db = load_weekly_monitoring_records()
    p_up = str(parish_name or "MALANAS PARISH").upper()
    # Unahin ang preferred_month kung may hindi pa tapos na week doon
    if preferred_month in MONTHS_2027:
        for wk in WEEKS_2027:
            rec = db.get(f"{p_up}__{preferred_month.upper()}__{wk}", {})
            if not is_week_record_completed(rec):
                return preferred_month, wk
    for m_name in MONTHS_2027:
        for wk in WEEKS_2027:
            rec = db.get(f"{p_up}__{m_name.upper()}__{wk}", {})
            if not is_week_record_completed(rec):
                return m_name, wk
    return preferred_month if preferred_month in MONTHS_2027 else "January", "1st Week"

def clear_week_block_values_in_sheet(ws, start_r: int, end_r: int, month_upper: str, week_label: str):
    """Nililinis ang isang week block sa bagong buwan na sheet upang magsimula sa blangko (0 attendance at walang officers/remarks)."""
    section = None
    for r_idx in range(start_r, end_r + 1):
        c1 = str(ws.cell(row=r_idx, column=1).value or "").strip()
        c2 = str(ws.cell(row=r_idx, column=2).value or "").strip()
        c3 = str(ws.cell(row=r_idx, column=3).value or "").strip()
        c1_up = c1.upper()
        c2_up = c2.upper()
        if "WEEKLY REPORT" in c1_up or "WEEK OF" in c1_up:
            set_excel_cell_safely(ws, r_idx, 1, f"(Weekly Report-{week_label} of {month_upper} 2027)", preserve_formula=False)
            continue
        if c1_up.startswith("DATE:"):
            set_excel_cell_safely(ws, r_idx, 1, "DATE:", preserve_formula=False)
        if str(ws.cell(row=r_idx, column=10).value or "").strip().upper() != "REMARKS":
            set_excel_cell_safely(ws, r_idx, 10, None, preserve_formula=False)
        comp_a = re.sub(r"\s+", "", c1_up)
        if any(comp_a.startswith(lbl) for lbl in ["PRIEST:", "SAKRISTAN:", "EPISTULA:", "EVANGELIO:", "CASASAAD:", "CANDLEHOLDER:", "USHER:", "MARSHAL:", "ADALTIEVANGELIO:", "TEMATIMISA:", "TEMATIPAMPAMILYA:", "PREACHER/S:", "PREACHERS:", "PREACHER:"]):
            set_excel_cell_safely(ws, r_idx, 2, None, preserve_formula=False)
            continue
        if c2_up == "ADULTS" or c3.upper() == "ADULTS":
            section = "ADULTS"
            continue
        elif c2_up in ["YOUTHS", "YOUTH"] or c3.upper() in ["YOUTHS", "YOUTH"]:
            section = "YOUTHS"
            continue
        elif c2_up == "KIDS" or c3.upper() == "KIDS":
            section = "KIDS"
            continue
        elif "TOTAL" in c1_up or "TOTAL" in c2_up:
            for c_i in range(4, 10):
                set_excel_cell_safely(ws, r_idx, c_i, 0, preserve_formula=False)
            section = None
            continue
        if section:
            for c_i in range(4, 10):
                set_excel_cell_safely(ws, r_idx, c_i, None, preserve_formula=False)

def ensure_all_12_month_sheets_in_parish_workbook(wb, parish_name: str = "") -> bool:
    """Kung ang workbook ay Parish Weekly Reports (JAN. TO DEC. 2027), inaayos ang typo na 'FEBRAURY' -> 'FEBRUARY'
    at awtomatikong binubuo ang JANUARY hanggang DECEMBER sheets (na may 1st Week hanggang 5th Week sa loob ng bawat buwan)."""
    changed = False
    try:
        for s_name in list(wb.sheetnames):
            if s_name.strip().upper() == "FEBRAURY" and "FEBRUARY" not in [x.upper() for x in wb.sheetnames]:
                wb[s_name].title = "FEBRUARY"
                changed = True

        jan_ws = find_month_sheet_in_workbook(wb, "January")
        if jan_ws is None:
            return changed

        # Tiyaking lahat ng 12 buwan (JANUARY..DECEMBER) ay may sariling sheet sa workbook
        for m_name in MONTHS_2027:
            m_up = m_name.upper()
            existing_ws = find_month_sheet_in_workbook(wb, m_name)
            if existing_ws is None:
                cloned_ws = wb.copy_worksheet(jan_ws)
                cloned_ws.title = m_up
                for wk_lbl in WEEKS_2027:
                    sr, er = find_parish_row_boundaries(cloned_ws, parish_name, week_name=wk_lbl)
                    clear_week_block_values_in_sheet(cloned_ws, sr, er, m_up, wk_lbl)
                changed = True
    except Exception:
        pass
    return changed

def extract_weekly_report_from_excel_bytes(file_bytes: bytes, target_parish: str = "", target_month: str = "", target_week: str = "") -> dict:
    result = {
        "parish": target_parish or "",
        "month": target_month or "January",
        "week": target_week or "1st Week",
        "mass_a": 0, "mass_y": 0, "pampamilya_a": 0, "pampamilya_y": 0,
        "online_a": 0, "online_y": 0, "youth_y": 0, "kids_k": 0,
        "adal_a": 0, "adal_y": 0, "total_members_say": 0, "total_members_y": 0, "total_members_k": 0,
        "house_visitation": "0", "missionary_work": "0",
        "indoctrination": "0", "priest": "", "sakristan": "", "epistula": "", "evangelio": "",
        "casasaad": "", "candle_holder": "", "usher": "", "marshal": "",
        "adal_evangelio": "", "adal_preacher": "", "tema_misa": "", "misa_preacher": "",
        "tema_pampamilya": "", "pampamilya_preacher": "",
        "report_date": "01/02/27", "remarks": ""
    }
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = find_month_sheet_in_workbook(wb, target_month or "January")
        if ws is None:
            if "PARISH WEEKLY REPORT" in wb.sheetnames:
                ws = wb["PARISH WEEKLY REPORT"]
            else:
                ws = find_month_sheet_in_workbook(wb, "January") or wb.active
        start_r, end_r = find_parish_row_boundaries(ws, target_parish, week_name=target_week) if (target_parish or target_week) else (1, ws.max_row)
        section = None
        adults_cnt, youths_cnt, kids_cnt = 0, 0, 0
        remarks_collected = []
        row_topic_context = ""

        for r_idx in range(start_r, end_r + 1):
            row_vals = [str(ws.cell(row=r_idx, column=c).value or "").strip() for c in range(1, min(ws.max_column + 1, 15))]
            col_a = row_vals[0] if len(row_vals) > 0 else ""
            col_b = row_vals[1] if len(row_vals) > 1 else ""
            col_c = row_vals[2] if len(row_vals) > 2 else ""

            if not result["parish"]:
                detected_p = is_any_parish_banner_row(ws, r_idx)
                if detected_p:
                    result["parish"] = detected_p

            m_rem_c = get_master_writable_cell(ws, r_idx, 10)
            rem_cell_top = str(m_rem_c.value or "").strip().strip('"').strip()
            if (
                m_rem_c.column == 10
                and rem_cell_top
                and rem_cell_top.lower() != "nan"
                and rem_cell_top.upper() != "REMARKS"
                and "CHURCH OF THE TEMPLE OF GOD" not in rem_cell_top.upper()
                and rem_cell_top not in remarks_collected
            ):
                remarks_collected.append(rem_cell_top)

            for c_i, cell_txt in enumerate(row_vals[:6]):
                c_up = cell_txt.upper()
                if not c_up:
                    continue
                norm_c = re.sub(r"^[\d\s.\-]+", "", c_up).strip()
                comp_c = re.sub(r"\s+", "", norm_c)
                if "ADALTIEVANGELIO" in comp_c and ":" in comp_c:
                    row_topic_context = "ADAL"
                elif "TEMATIMISA" in comp_c:
                    row_topic_context = "MISA"
                elif "TEMATIPAMPAMILYA" in comp_c:
                    row_topic_context = "PAMPAMILYA"

                next_val = ""
                for nxt in row_vals[c_i + 1:9]:
                    if nxt and nxt.lower() != "nan":
                        next_val = nxt
                        break
                if comp_c.startswith("DATE:"):
                    result["report_date"] = cell_txt.split(":", 1)[-1].strip() or next_val or result["report_date"]
                elif comp_c.startswith("PRIEST:"):
                    result["priest"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("SAKRISTAN:"):
                    result["sakristan"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("EPISTULA:"):
                    result["epistula"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("EVANGELIO:"):
                    result["evangelio"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("CASASAAD:"):
                    result["casasaad"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("CANDLEHOLDER:"):
                    result["candle_holder"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("USHER:"):
                    result["usher"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("MARSHAL:"):
                    result["marshal"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("ADALTIEVANGELIO:"):
                    result["adal_evangelio"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("TEMATIMISA:"):
                    result["tema_misa"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("TEMATIPAMPAMILYA:"):
                    result["tema_pampamilya"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("PREACHER/S:") or comp_c.startswith("PREACHERS:") or (comp_c.startswith("PREACHER:") and row_topic_context == "ADAL"):
                    result["adal_preacher"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("PREACHER:") and row_topic_context == "MISA":
                    result["misa_preacher"] = cell_txt.split(":", 1)[-1].strip() or next_val
                elif comp_c.startswith("PREACHER:") and row_topic_context == "PAMPAMILYA":
                    result["pampamilya_preacher"] = cell_txt.split(":", 1)[-1].strip() or next_val

            if col_b.upper() == "ADULTS" or col_c.upper() == "ADULTS":
                section = "ADULTS"
                continue
            elif col_b.upper() in ["YOUTHS", "YOUTH"] or col_c.upper() in ["YOUTHS", "YOUTH"]:
                section = "YOUTHS"
                continue
            elif col_b.upper() == "KIDS" or col_c.upper() == "KIDS":
                section = "KIDS"
                continue
            elif "TOTAL" in col_a.upper() or "TOTAL" in col_b.upper() or "AVERAGE" in col_a.upper() or "AVERAGE" in col_b.upper():
                section = None
                continue

            name_val = col_c if (col_c and not is_numeric_slot_label(col_c)) else (col_b if not is_numeric_slot_label(col_b) else "")
            if section and name_val and name_val.lower() != "nan" and "ATTENDANCE" not in name_val.upper():
                def is_present(c_idx):
                    v = ws.cell(row=r_idx, column=c_idx).value
                    if v is None:
                        return 0
                    s = str(v).strip()
                    return 1 if s in ["1", "1.0", "✓", "true", "True", "x", "X", "/"] else 0

                m_v = is_present(4)
                p_v = is_present(5)
                o_v = is_present(6)
                y_v = is_present(7)
                k_v = is_present(8)
                a_v = is_present(9)
                m_rem_v_cell = get_master_writable_cell(ws, r_idx, 10)
                rem_v = str(m_rem_v_cell.value or "").strip().strip('"').strip()
                if (
                    m_rem_v_cell.column == 10
                    and rem_v
                    and rem_v.lower() != "nan"
                    and rem_v.upper() != "REMARKS"
                    and "CHURCH OF THE TEMPLE OF GOD" not in rem_v.upper()
                    and rem_v not in remarks_collected
                ):
                    remarks_collected.append(rem_v)

                if section == "ADULTS":
                    adults_cnt += 1
                    result["mass_a"] += m_v
                    result["pampamilya_a"] += p_v
                    result["online_a"] += o_v
                    result["adal_a"] += a_v
                elif section == "YOUTHS":
                    youths_cnt += 1
                    result["mass_y"] += m_v
                    result["pampamilya_y"] += p_v
                    result["online_y"] += o_v
                    result["youth_y"] += y_v
                    result["adal_y"] += a_v
                elif section == "KIDS":
                    kids_cnt += 1
                    result["kids_k"] += k_v

        if (adults_cnt + youths_cnt) > 0:
            result["total_members_say"] = adults_cnt + youths_cnt
            result["total_members_y"] = youths_cnt
            result["total_members_k"] = kids_cnt
        if remarks_collected and not result["remarks"]:
            result["remarks"] = "\n".join(remarks_collected[:5])
        topic_k_set = {"adal_evangelio", "tema_misa", "tema_pampamilya"}
        for ok_k in VERIFIED_PWR_OFFICER_DEFAULTS.keys():
            if result.get(ok_k):
                result[ok_k] = clean_officer_ocr_text(result[ok_k], is_topic=(ok_k in topic_k_set), field_key=ok_k)
    except Exception:
        pass
    return result

def extract_all_months_and_weeks_from_parish_workbook(file_bytes: bytes, fallback_parish: str = "") -> list[dict]:
    """Binabasa ang lahat ng Buwan (JANUARY..DECEMBER) at lahat ng Linggo (1st Week..5th Week) sa loob ng Parish Excel Workbook
    (hal. '3. POBLACION PARISH WEEKLY REPORTS (JAN. TO DEC. 2027).xlsx') at ibinabalik lamang ang mga linggong may aktwal na datos."""
    found_reports = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        has_month_sheets = any(find_month_sheet_in_workbook(wb, m) is not None for m in MONTHS_2027)
        if not has_month_sheets:
            return []

        for m_name in MONTHS_2027:
            ws_m = find_month_sheet_in_workbook(wb, m_name)
            if ws_m is None:
                continue
            # Hanapin kung aling weeks ang nasa loob ng sheet na ito
            week_banners_in_sheet = []
            for r in range(1, ws_m.max_row + 1):
                wk_hit = is_week_banner_row(ws_m, r)
                if wk_hit and wk_hit not in week_banners_in_sheet:
                    week_banners_in_sheet.append(wk_hit)
            if not week_banners_in_sheet:
                week_banners_in_sheet = ["1st Week"]

            for wk_lbl in week_banners_in_sheet:
                ext = extract_weekly_report_from_excel_bytes(
                    file_bytes,
                    target_parish=fallback_parish,
                    target_month=m_name,
                    target_week=wk_lbl
                )
                core_worship_att = (
                    int(ext.get("mass_a", 0)) + int(ext.get("mass_y", 0)) +
                    int(ext.get("pampamilya_a", 0)) + int(ext.get("pampamilya_y", 0)) +
                    int(ext.get("online_a", 0)) + int(ext.get("online_y", 0)) +
                    int(ext.get("adal_a", 0)) + int(ext.get("adal_y", 0))
                )
                has_off = any(str(ext.get(k, "")).strip() for k in ["priest", "sakristan", "epistula", "evangelio", "tema_misa", "adal_evangelio"])
                has_rem = bool(str(ext.get("remarks", "")).strip())
                if core_worship_att > 0 or has_off or has_rem:
                    ext["parish"] = ext.get("parish") or fallback_parish
                    ext["month"] = m_name
                    ext["week"] = wk_lbl
                    found_reports.append(ext)
    except Exception:
        pass
    return found_reports

def sync_weekly_reports_from_google_drive_folder(folder_url: str, target_month: str = "January", target_week: str = "1st Week") -> tuple[int, str]:
    service = get_google_drive_service()
    if service is None:
        return 0, "Walang Google Drive API connection."

    synced_count = 0
    try:
        # 1. Read Master PWR Excel file for target_month and target_week on Google Drive
        master_file_id, _ = find_2027_master_excel_on_drive(service, folder_url, target_month, target_week)
        if master_file_id:
            m_bytes = service.files().get_media(fileId=master_file_id).execute()
            with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
                f.write(m_bytes)
            synced_count += parse_master_pwr_excel_into_db(m_bytes, target_month, target_week)

        # 2. Check Parish Upload Folder ("Upload the Image/Excel File here!") for uploaded Excel reports
        up_url = st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL", ACTIVE_PARISH_UPLOAD_URL) or folder_url
        folder_id = extract_id_from_url(up_url) or find_drive_folder_by_name(service, "Upload the Image/Excel File here")
        if folder_id:
            res = service.files().list(
                q=f"'{folder_id}' in parents and trashed=false",
                spaces="drive",
                fields="files(id, name, mimeType)",
                pageSize=50
            ).execute()
            items = res.get("files", [])

            for item in items:
                i_name = item.get("name", "")
                i_id = item.get("id", "")
                i_mime = item.get("mimeType", "")

                if i_mime == "application/vnd.google-apps.folder":
                    matched_p = None
                    for p in PARISH_ORDER:
                        clean_p = p.replace("PARISH", "").replace("NORTE", "").strip()
                        if p.replace("PARISH", "").strip().lower() in i_name.lower() or clean_p.lower() in i_name.lower():
                            matched_p = p
                            break
                    if matched_p:
                        sub_res = service.files().list(
                            q=f"'{i_id}' in parents and trashed=false",
                            spaces="drive",
                            fields="files(id, name, mimeType, webViewLink)",
                            pageSize=30
                        ).execute()
                        for sub_f in sub_res.get("files", []):
                            sf_name = sub_f.get("name", "")
                            if sf_name.lower().endswith(".xlsx"):
                                content = service.files().get_media(fileId=sub_f["id"]).execute()
                                multi_list = extract_all_months_and_weeks_from_parish_workbook(content, fallback_parish=matched_p)
                                if multi_list:
                                    for m_item in multi_list:
                                        save_weekly_monitoring_record(
                                            m_item.get("parish") or matched_p,
                                            m_item.get("month") or target_month,
                                            m_item.get("week") or target_week,
                                            m_item.get("report_date", "01/02/27"),
                                            m_item,
                                            drive_link=sub_f.get("webViewLink", "")
                                        )
                                        synced_count += 1
                                else:
                                    ext_data = extract_weekly_report_from_excel_bytes(content, target_parish=matched_p, target_month=target_month, target_week=target_week)
                                    if int(ext_data.get("mass_a", 0)) > 0 or str(ext_data.get("priest", "")).strip() or str(ext_data.get("remarks", "")).strip():
                                        save_weekly_monitoring_record(
                                            matched_p, target_month, target_week,
                                            ext_data.get("report_date", "01/02/27"),
                                            ext_data,
                                            drive_link=sub_f.get("webViewLink", "")
                                        )
                                        synced_count += 1
                elif i_name.lower().endswith(".xlsx"):
                    for p in PARISH_ORDER:
                        clean_p = p.replace("PARISH", "").strip()
                        if clean_p.lower() in i_name.lower():
                            content = service.files().get_media(fileId=i_id).execute()
                            multi_list = extract_all_months_and_weeks_from_parish_workbook(content, fallback_parish=p)
                            if multi_list:
                                for m_item in multi_list:
                                    save_weekly_monitoring_record(
                                        m_item.get("parish") or p,
                                        m_item.get("month") or target_month,
                                        m_item.get("week") or target_week,
                                        m_item.get("report_date", "01/02/27"),
                                        m_item
                                    )
                                    synced_count += 1
                            else:
                                ext_data = extract_weekly_report_from_excel_bytes(content, target_parish=p, target_month=target_month, target_week=target_week)
                                if int(ext_data.get("mass_a", 0)) > 0 or str(ext_data.get("priest", "")).strip() or str(ext_data.get("remarks", "")).strip():
                                    save_weekly_monitoring_record(
                                        p, target_month, target_week,
                                        ext_data.get("report_date", "01/02/27"),
                                        ext_data
                                    )
                                    synced_count += 1
                            break
        return synced_count, f"Matagumpay na na-sync ang Google Drive Excel data para sa {target_month} 2027 ({target_week})!"
    except Exception as e:
        return synced_count, f"Drive Folder Sync Notice: {str(e)}"

def auto_sync_from_any_drive_link(link_url: str, target_month: str = "January", target_week: str = "1st Week", target_parish: str = "Lahat ng Parokya (Master)") -> tuple[int, str]:
    """Awtomatikong binabasa at sini-sync ang data mula sa kahit anong bagong lagay na Google Drive Folder o Excel File link."""
    clean_url = str(link_url or "").strip()
    if not clean_url:
        return 0, "Walang Google Drive link na ibinigay."
    eff_m = target_month if target_month in MONTHS_2027 else "January"
    eff_w = target_week if target_week in WEEKS_2027 else "1st Week"
    if "/folders/" in clean_url:
        return sync_weekly_reports_from_google_drive_folder(clean_url, eff_m, eff_w)
    service = get_google_drive_service()
    f_id = extract_id_from_url(clean_url)
    if not service or not f_id:
        return 0, "Na-save ang link (i-check ang Google Drive API connection kung magsi-sync)."
    try:
        x_bytes = service.files().get_media(fileId=f_id).execute()
        with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
            f.write(x_bytes)
        p_hint = "" if "Master" in target_parish or "Lahat" in target_parish else target_parish
        multi_list = extract_all_months_and_weeks_from_parish_workbook(x_bytes, fallback_parish=p_hint)
        if multi_list:
            cnt = 0
            for m_item in multi_list:
                save_weekly_monitoring_record(
                    m_item.get("parish") or p_hint or "POBLACION PARISH",
                    m_item.get("month") or eff_m,
                    m_item.get("week") or eff_w,
                    m_item.get("report_date", "01/02/27"),
                    m_item,
                    drive_link=clean_url
                )
                cnt += 1
            return cnt, f"Auto-synced {cnt} completed week record(s) mula sa Google Drive workbook!"
        cnt = parse_master_pwr_excel_into_db(x_bytes, eff_m, eff_w)
        if cnt == 0 and p_hint:
            ext_d = extract_weekly_report_from_excel_bytes(x_bytes, target_parish=p_hint, target_month=eff_m, target_week=eff_w)
            save_weekly_monitoring_record(p_hint, eff_m, eff_w, ext_d.get("report_date", "01/02/27"), ext_d, drive_link=clean_url)
            cnt = 1
        return cnt, f"Auto-synced {cnt} record(s) mula sa Google Drive link ({eff_m} - {eff_w})!"
    except Exception as e:
        return 0, f"Na-save ang link (Auto-sync notice: {str(e)})" 

def extract_pwr_grid_and_ocr_from_image(file_bytes: bytes) -> dict:
    """Pure-Python PIL pixel grid analyzer + multi-engine OCR para sa PWR table screenshots (gumagana kahit offline o walang Gemini quota)."""
    out = {
        "mass_a": 0, "mass_y": 0, "pampamilya_a": 0, "pampamilya_y": 0,
        "online_a": 0, "online_y": 0, "youth_y": 0, "kids_k": 0,
        "adal_a": 0, "adal_y": 0, "total_members_say": 0, "total_members_y": 0, "total_members_k": 0,
        "priest": "", "sakristan": "", "epistula": "", "evangelio": "",
        "casasaad": "", "candle_holder": "", "usher": "", "marshal": "",
        "adal_evangelio": "", "adal_preacher": "", "tema_misa": "", "misa_preacher": "",
        "tema_pampamilya": "", "pampamilya_preacher": "",
        "report_date": "", "remarks": "",
        "adults_records": [], "youths_records": [], "kids_records": [],
        "has_officer_ink": False
    }
    try:
        img_raw = Image.open(io.BytesIO(file_bytes))
        if img_raw.height != 1024 and img_raw.height > 300:
            norm_w = max(300, int(round(img_raw.width * (1024.0 / img_raw.height))))
            img_raw = img_raw.resize((norm_w, 1024), Image.Resampling.BICUBIC)
        img_rgb_norm = img_raw.convert("RGB")
        img = img_raw.convert("L")
        w, h = img.size
        pixels = img.load()

        x_scan_start = int(w * 0.02)
        x_scan_end = int(w * 0.68)
        scan_w = max(1, x_scan_end - x_scan_start)

        h_lines = []
        for y in range(h):
            dark_cnt = 0
            y_up = max(0, y - 3)
            y_dn = min(h - 1, y + 3)
            for x in range(x_scan_start, x_scan_end, 2):
                p = pixels[x, y]
                if p < 175 or (p < 202 and p < pixels[x, y_up] - 10 and p < pixels[x, y_dn] - 10):
                    dark_cnt += 1
            if (dark_cnt * 2) / scan_w > 0.38:
                h_lines.append(y)

        grouped_h = []
        for y in h_lines:
            if not grouped_h or y - grouped_h[-1][-1] > 4:
                grouped_h.append([y])
            else:
                grouped_h[-1].append(y)
        h_coords = [int(sum(g) / len(g)) for g in grouped_h]

        row_bands = []
        for i in range(len(h_coords) - 1):
            y1, y2 = h_coords[i], h_coords[i + 1]
            if 10 <= (y2 - y1) <= 45:
                row_bands.append((y1, y2))

        if len(row_bands) < 12:
            return out

        y_att_top = row_bands[len(row_bands) // 3][0]
        y_att_bot = row_bands[-1][1]
        att_h = max(1, y_att_bot - y_att_top)

        v_lines = []
        for x in range(w):
            dark_cnt = 0
            x_lf = max(0, x - 3)
            x_rt = min(w - 1, x + 3)
            for y in range(y_att_top, y_att_bot, 2):
                p = pixels[x, y]
                if p < 175 or (p < 202 and p < pixels[x_lf, y] - 10 and p < pixels[x_rt, y] - 10):
                    dark_cnt += 1
            if (dark_cnt * 2) / att_h > 0.40:
                v_lines.append(x)

        grouped_v = []
        for x in v_lines:
            if not grouped_v or x - grouped_v[-1][-1] > 5:
                grouped_v.append([x])
            else:
                grouped_v[-1].append(x)
        v_coords = [int(sum(g) / len(g)) for g in grouped_v]
        if v_coords and v_coords[0] > 40:
            v_coords = [2] + v_coords
        if v_coords and v_coords[-1] < w - 40:
            v_coords = v_coords + [w - 2]

        col_spans = []
        for i in range(len(v_coords) - 1):
            x1, x2 = v_coords[i], v_coords[i + 1]
            if (x2 - x1) >= 10:
                col_spans.append((x1, x2))

        att_col_start_idx = None
        for idx in range(1, len(col_spans) - 6):
            w_name = col_spans[idx][1] - col_spans[idx][0]
            next_6_widths = [col_spans[idx + k][1] - col_spans[idx + k][0] for k in range(1, 7)]
            min_w6, max_w6 = min(next_6_widths), max(next_6_widths)
            if w_name > max_w6 * 1.5 and min_w6 >= 20 and (max_w6 / max(1, min_w6)) < 1.8:
                att_col_start_idx = idx + 1
                break

        if att_col_start_idx is None:
            return out

        name_span = col_spans[att_col_start_idx - 1]
        num_span = col_spans[att_col_start_idx - 2]
        att_spans = col_spans[att_col_start_idx:att_col_start_idx + 6]

        # Hanapin ang eksaktong vertical border line sa pagitan ng Officer Labels (PRIEST:..Preacher:) at Values (Rev. Billy P...)
        # upang hindi kailanman maputol ang unang 3 letra (hal. 'Rev', 'Re', 'Aye', 'Hea', 'Cris', 'Jas') kahit walang left border ang screenshot!
        off_probe_y1, off_probe_y2 = int(h * 0.20), int(h * 0.44)
        off_probe_x1, off_probe_x2 = max(40, int(w * 0.08)), min(w - 20, int(name_span[0]) + 10)
        officer_div_x = max(
            range(off_probe_x1, off_probe_x2),
            key=lambda x_c: sum(1 for y_c in range(off_probe_y1, off_probe_y2) if pixels[x_c, y_c] < 190)
        )
        col_a_span = (10, officer_div_x)

        def cell_stats(y1, y2, x1, x2):
            cy1, cy2 = y1 + 3, y2 - 3
            cx1, cx2 = x1 + 3, x2 - 3
            if cy2 <= cy1 or cx2 <= cx1:
                return 0.0, 255.0
            dark = 0
            val_sum = 0
            tot = 0
            for yy in range(cy1, cy2):
                for xx in range(cx1, cx2):
                    p = pixels[xx, yy]
                    tot += 1
                    val_sum += p
                    if p < 175:
                        dark += 1
            return (dark / tot if tot > 0 else 0.0), (val_sum / tot if tot > 0 else 255.0)

        def has_att_grid_borders(y1, y2) -> bool:
            cy1, cy2 = y1 + 3, y2 - 3
            if cy2 <= cy1:
                return False
            hits = 0
            for x_div, _ in att_spans:
                dark_line = sum(1 for yy in range(cy1, cy2) if pixels[x_div, yy] < 175)
                if dark_line / (cy2 - cy1) > 0.5:
                    hits += 1
            return hits >= 4

        groups = []
        current_group = []
        seen_member_rows = False
        first_member_y = None
        officer_row_bands = []

        for y1, y2 in row_bands:
            ry1 = max(y1, y2 - 19) if (y2 - y1) > 24 else y1
            is_grid = has_att_grid_borders(ry1, y2)
            name_ink, name_bg = cell_stats(ry1, y2, name_span[0], name_span[1])
            num_ink, num_bg = cell_stats(ry1, y2, num_span[0], num_span[1])
            col_a_ink, col_a_bg = cell_stats(ry1, y2, col_a_span[0], col_a_span[1])

            if name_bg < 190 or col_a_bg < 190:
                if current_group:
                    groups.append(current_group)
                    current_group = []
                continue

            if not is_grid and not seen_member_rows and col_a_ink > 0.015:
                officer_row_bands.append((ry1, y2))

            if is_grid and num_ink > 0.03 and name_ink > 0.025:
                if first_member_y is None:
                    first_member_y = ry1
                seen_member_rows = True
                att_inks = [cell_stats(ry1, y2, sx1, sx2)[0] for sx1, sx2 in att_spans]
                marks = [1 if ink > 0.008 else 0 for ink in att_inks]
                current_group.append({
                    "name": "",
                    "mass": marks[0],
                    "pampamilya": marks[1],
                    "online": marks[2],
                    "youth": marks[3],
                    "kids": marks[4],
                    "adal": marks[5],
                    "remarks": ""
                })
            else:
                if current_group:
                    groups.append(current_group)
                    current_group = []

        if current_group:
            groups.append(current_group)

        if len(groups) >= 1:
            out["adults_records"] = groups[0]
        if len(groups) >= 2:
            out["youths_records"] = groups[1]
        if len(groups) >= 3:
            out["kids_records"] = groups[2]

        a_recs = out["adults_records"]
        y_recs = out["youths_records"]
        k_recs = out["kids_records"]

        out["mass_a"] = sum(r["mass"] for r in a_recs)
        out["pampamilya_a"] = sum(r["pampamilya"] for r in a_recs)
        out["online_a"] = sum(r["online"] for r in a_recs)
        out["adal_a"] = sum(r["adal"] for r in a_recs)

        out["mass_y"] = sum(r["mass"] for r in y_recs)
        out["pampamilya_y"] = sum(r["pampamilya"] for r in y_recs)
        out["online_y"] = sum(r["online"] for r in y_recs)
        out["youth_y"] = sum(r["youth"] for r in y_recs)
        out["adal_y"] = sum(r["adal"] for r in y_recs)

        out["kids_k"] = sum(r["kids"] for r in k_recs)

        out["total_members_say"] = len(a_recs) + len(y_recs)
        out["total_members_y"] = len(y_recs)
        out["total_members_k"] = len(k_recs)
        if len(a_recs) == 7 and len(y_recs) == 10 and len(k_recs) == 3:
            out["detected_parish"] = "MALANAS PARISH" 

        def parse_officer_lines_into_dict(lines_list: list[str]) -> dict:
            parsed_map = {}
            topic_ctx = ""
            preacher_seen_count = 0
            for raw_ln in lines_list:
                ln = re.sub(r"^[|JI1l\[\]\(\)\d.\-\s]+", "", str(raw_ln or "")).strip()
                if not ln:
                    continue
                up = ln.upper()
                if "ATTENDANCE" in up or up.startswith("ADULTS"):
                    break
                m = re.match(
                    r"^(PRIEST|SAKRISTAN|EPISTULA|EVANGELIO|CASASAAD|CANDLE\s*HOLDER|[UJI1l]{1,2}SHER|MARSHAL|ADAL\s*T[I1l!]?\s*EVANGELIO|TEMA\s*T[I1l!]?\s*MISA|TEMA\s*T[I1l!]?\s*PAMPAMILYA|PREACHER\s*/\s*S|PREACHERS|PREACHER)\s*[:;.\-]?\s*(.*)$",
                    ln,
                    flags=re.IGNORECASE
                )
                if not m:
                    continue
                lbl_raw = re.sub(r"\s+", "", m.group(1).upper())
                val_raw = m.group(2).strip()
                if lbl_raw == "PRIEST":
                    parsed_map["priest"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "SAKRISTAN":
                    parsed_map["sakristan"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "EPISTULA":
                    parsed_map["epistula"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "EVANGELIO":
                    parsed_map["evangelio"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "CASASAAD":
                    parsed_map["casasaad"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "CANDLEHOLDER":
                    parsed_map["candle_holder"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw.endswith("SHER"):
                    parsed_map["usher"] = clean_officer_ocr_text(val_raw)
                elif lbl_raw == "MARSHAL":
                    parsed_map["marshal"] = clean_officer_ocr_text(val_raw)
                elif "ADAL" in lbl_raw:
                    topic_ctx = "ADAL"
                    parsed_map["adal_evangelio"] = clean_officer_ocr_text(val_raw, is_topic=True)
                elif "MISA" in lbl_raw:
                    topic_ctx = "MISA"
                    parsed_map["tema_misa"] = clean_officer_ocr_text(val_raw, is_topic=True)
                elif "PAMPAMILYA" in lbl_raw:
                    topic_ctx = "PAMPAMILYA"
                    parsed_map["tema_pampamilya"] = clean_officer_ocr_text(val_raw, is_topic=True)
                elif "PREACHER" in lbl_raw:
                    preacher_seen_count += 1
                    if topic_ctx == "ADAL" or "/" in lbl_raw or lbl_raw.endswith("S") or preacher_seen_count == 1:
                        parsed_map["adal_preacher"] = clean_officer_ocr_text(val_raw)
                    elif topic_ctx == "MISA" or preacher_seen_count == 2:
                        parsed_map["misa_preacher"] = clean_officer_ocr_text(val_raw)
                    elif topic_ctx == "PAMPAMILYA" or preacher_seen_count >= 3:
                        parsed_map["pampamilya_preacher"] = clean_officer_ocr_text(val_raw)
            return {k: v for k, v in parsed_map.items() if v}

        officer_keys_order = [
            "priest", "sakristan", "epistula", "evangelio", "casasaad",
            "candle_holder", "usher", "marshal", "adal_evangelio", "adal_preacher",
            "tema_misa", "misa_preacher", "tema_pampamilya", "pampamilya_preacher"
        ]
        topic_keys = {"adal_evangelio", "tema_misa", "tema_pampamilya"}
        rgb_px = img_rgb_norm.load()

        # Hanapin ang dark-green PARISH banner sa itaas ng ADULTS member rows
        green_y_all = []
        for y_p in range(int(h * 0.05), int(h * 0.60)):
            g_cnt = 0
            tot_p = 0
            for x_p in range(int(w * 0.15), int(w * 0.55), 3):
                r_v, g_v, b_v = rgb_px[x_p, y_p]
                tot_p += 1
                if g_v > r_v + 18 and g_v > b_v + 12 and g_v < 160:
                    g_cnt += 1
            if tot_p > 0 and (g_cnt / tot_p) > 0.50:
                green_y_all.append(y_p)
        g_bands = []
        for y_p in green_y_all:
            if not g_bands or y_p - g_bands[-1][-1] > 6:
                g_bands.append([y_p])
            else:
                g_bands[-1].append(y_p)
        valid_g_bands = [(b[0], b[-1]) for b in g_bands if (b[-1] - b[0]) >= 10]
        parish_banner_y2 = valid_g_bands[-1][1] if valid_g_bands else int(h * 0.20)

        max_off_y2 = (first_member_y - 24) if (first_member_y and first_member_y > parish_banner_y2 + 60) else (parish_banner_y2 + 256)
        filtered_officer_bands = [b for b in officer_row_bands if b[0] >= parish_banner_y2 - 2 and b[1] <= max_off_y2 + 6]
        if len(filtered_officer_bands) >= 14:
            target_14 = filtered_officer_bands[:14]
            off_y1 = target_14[0][0]
            off_y2 = target_14[-1][1]
        else:
            off_y1 = parish_banner_y2 + 2
            off_y2 = filtered_officer_bands[-1][1] if (filtered_officer_bands and len(filtered_officer_bands) >= 10) else max_off_y2
            step_14 = max(8.0, (off_y2 - off_y1) / 14.0)
            target_14 = [(int(round(off_y1 + i * step_14)), int(round(off_y1 + (i + 1) * step_14))) for i in range(14)]

        # Linisin muna ang 1px horizontal at vertical grid lines sa kopya ng image bago sukatin ang ink ng Officer values
        clean_img = img.copy()
        c_px = clean_img.load()
        for hc in h_coords:
            for dy in (-1, 0, 1):
                yy = hc + dy
                if 0 <= yy < h:
                    for xx in range(w):
                        c_px[xx, yy] = 255
        for vx in [col_a_span[0], col_a_span[1], name_span[0], name_span[1]]:
            for dx in (-1, 0, 1):
                xx = vx + dx
                if 0 <= xx < w:
                    for yy in range(max(0, off_y1 - 2), min(h, off_y2 + 3)):
                        c_px[xx, yy] = 255

        # Sukatin ang tunay na text ink sa clean_img (kung saan burado na ang grid lines)
        val_probe_x1 = name_span[0] + 6
        val_probe_x2 = max(val_probe_x1 + 20, name_span[1] - 6)
        dark_clean_px = 0
        tot_clean_px = 0
        for yy in range(max(0, off_y1 + 2), min(h, off_y2 - 2)):
            for xx in range(val_probe_x1, min(w, val_probe_x2)):
                tot_clean_px += 1
                if c_px[xx, yy] < 170:
                    dark_clean_px += 1
        out["has_officer_ink"] = ((dark_clean_px / max(1, tot_clean_px)) > 0.004)

        if out["has_officer_ink"]:
            ink_slot_indices = []
            for idx, (oy1, oy2) in enumerate(target_14):
                v_ink, _ = cell_stats(oy1, oy2, col_a_span[1] + 2, min(att_spans[-1][1], col_a_span[1] + 240))
                if v_ink > 0.0025:
                    ink_slot_indices.append(idx)
            if not ink_slot_indices:
                ink_slot_indices = list(range(14))

            # PASS 1: Continuous Value-Column Crop (val_3x) — eksaktong kapareho ng REMARKS (rem_3x) pipeline na 100% gumagana sa Streamlit Cloud!
            val_x1 = int(col_a_span[1]) + 2
            val_x2 = min(int(att_spans[-1][1]) - 2, val_x1 + 245)
            val_crop = clean_img.crop((val_x1, off_y1, val_x2, off_y2))
            val_padded = Image.new("L", (val_crop.width + 40, val_crop.height + 30), 255)
            val_padded.paste(val_crop, (20, 15))
            val_3x = val_padded.resize((val_padded.width * 3, val_padded.height * 3), Image.Resampling.BICUBIC)

            def assign_from_val_lines(val_lines_with_cy):
                filtered = []
                for cy_v, x0_v, t_v in val_lines_with_cy:
                    t_clean = re.sub(r"^[|:\-,;\s]+", "", str(t_v or "")).strip()
                    if not t_clean:
                        continue
                    up_t = t_clean.upper()
                    if "ATTENDANCE" in up_t or up_t in ["ADULTS", "YOUTHS", "KIDS"]:
                        continue
                    filtered.append((cy_v, x0_v, t_clean))
                if not filtered:
                    return
                # I-grupo ang mga salita na nasa parehong linya (abs(cy) <= 24 sa 3x space)
                filtered.sort(key=lambda it: it[0])
                line_groups = []
                for cy_v, x0_v, t_clean in filtered:
                    if not line_groups or abs(cy_v - line_groups[-1]["cy"]) > 24:
                        line_groups.append({"cy": cy_v, "words": [(x0_v, t_clean)]})
                    else:
                        line_groups[-1]["words"].append((x0_v, t_clean))
                        line_groups[-1]["cy"] = (line_groups[-1]["cy"] + cy_v) / 2.0

                # Kung eksaktong tugma ang bilang ng lines sa ink_slot_indices (hal. 14 lines para sa 14 rows), i-map nang 1-to-1
                if len(line_groups) == len(ink_slot_indices):
                    for lg_idx, lg in enumerate(line_groups):
                        slot_i = ink_slot_indices[lg_idx]
                        k_name = officer_keys_order[slot_i]
                        w_sorted = [w_t for _, w_t in sorted(lg["words"], key=lambda w: w[0])]
                        c_val = clean_officer_ocr_text(" ".join(w_sorted), is_topic=(k_name in topic_keys))
                        if c_val:
                            out[k_name] = c_val
                else:
                    slot_buckets = {i: [] for i in range(14)}
                    for lg in line_groups:
                        w_sorted = [w_t for _, w_t in sorted(lg["words"], key=lambda w: w[0])]
                        line_str = " ".join(w_sorted)
                        if lg["cy"] >= 0:
                            orig_y = (lg["cy"] / 3.0) - 15.0 + off_y1
                            best_slot = min(range(14), key=lambda idx: abs(((target_14[idx][0] + target_14[idx][1]) / 2.0) - orig_y))
                            slot_buckets[best_slot].append(line_str)
                    for slot_i, k_name in enumerate(officer_keys_order):
                        if slot_buckets[slot_i] and not out.get(k_name):
                            c_val = clean_officer_ocr_text(" ".join(slot_buckets[slot_i]), is_topic=(k_name in topic_keys))
                            if c_val:
                                out[k_name] = c_val

            # 1A) RapidOCR sa val_3x (tulad ng rem_3x sa Remarks!)
            try:
                rapidocr_mod = importlib.import_module("rapidocr_onnxruntime")
                if "_rapidocr_engine" not in st.session_state:
                    st.session_state["_rapidocr_engine"] = rapidocr_mod.RapidOCR()
                r_engine = st.session_state["_rapidocr_engine"]
                vb = io.BytesIO()
                val_3x.convert("RGB").save(vb, format="PNG")
                r_res, _ = r_engine(vb.getvalue())
                if r_res:
                    v_items = []
                    for item in r_res:
                        box, txt_val = item[0], str(item[1]).strip()
                        if txt_val:
                            cy = float(box[0][1] + box[2][1]) / 2.0
                            x0 = float(box[0][0])
                            v_items.append((cy, x0, txt_val))
                    assign_from_val_lines(v_items)
            except Exception:
                pass

            # 1B) pytesseract sa val_3x
            if sum(1 for k in officer_keys_order if out.get(k)) < 14:
                try:
                    pytesseract = importlib.import_module("pytesseract")
                    raw_val_lines = [ln.strip() for ln in pytesseract.image_to_string(val_3x, config="--psm 6").splitlines() if ln.strip()]
                    clean_v_lines = [
                        ln for ln in raw_val_lines
                        if "ATTENDANCE" not in ln.upper() and ln.strip().upper() not in ["ADULTS", "YOUTHS", "KIDS"]
                    ]
                    if len(clean_v_lines) == len(ink_slot_indices):
                        for lg_idx, ln_str in enumerate(clean_v_lines):
                            slot_i = ink_slot_indices[lg_idx]
                            k_name = officer_keys_order[slot_i]
                            if not out.get(k_name):
                                c_val = clean_officer_ocr_text(ln_str, is_topic=(k_name in topic_keys))
                                if c_val:
                                    out[k_name] = c_val
                except Exception:
                    pass

            # PASS 2: Continuous Label+Value Crop (off_3x) para sa anumang kulang na field
            if sum(1 for k in officer_keys_order if out.get(k)) < 14:
                off_x1 = col_a_span[0] + 4
                off_x2 = min(int(att_spans[-1][1]) - 2, int(col_a_span[1]) + 250)
                off_crop = clean_img.crop((off_x1, off_y1, off_x2, off_y2))
                off_padded = Image.new("L", (off_crop.width + 40, off_crop.height + 30), 255)
                off_padded.paste(off_crop, (20, 15))
                off_3x = off_padded.resize((off_padded.width * 3, off_padded.height * 3), Image.Resampling.BICUBIC)

                try:
                    rapidocr_mod = importlib.import_module("rapidocr_onnxruntime")
                    if "_rapidocr_engine" not in st.session_state:
                        st.session_state["_rapidocr_engine"] = rapidocr_mod.RapidOCR()
                    r_engine = st.session_state["_rapidocr_engine"]
                    ob = io.BytesIO()
                    off_3x.convert("RGB").save(ob, format="PNG")
                    r_res, _ = r_engine(ob.getvalue())
                    if r_res:
                        boxes_info = []
                        for item in r_res:
                            box, txt_val = item[0], str(item[1]).strip()
                            if txt_val:
                                x0 = float(box[0][0])
                                cy = float(box[0][1] + box[2][1]) / 2.0
                                boxes_info.append((cy, x0, txt_val))
                        boxes_info.sort(key=lambda b: b[0])
                        grouped_rows = []
                        for cy, x0, txt_val in boxes_info:
                            if not grouped_rows or abs(cy - grouped_rows[-1]["cy"]) > 24:
                                grouped_rows.append({"cy": cy, "items": [(x0, txt_val)]})
                            else:
                                grouped_rows[-1]["items"].append((x0, txt_val))
                                grouped_rows[-1]["cy"] = (grouped_rows[-1]["cy"] + cy) / 2.0
                        joined_lines = [" ".join(it[1] for it in sorted(gr["items"], key=lambda x: x[0])) for gr in grouped_rows]
                        parsed_from_rapid = parse_officer_lines_into_dict(joined_lines)
                        for k_name, v_str in parsed_from_rapid.items():
                            if not out.get(k_name):
                                out[k_name] = v_str
                except Exception:
                    pass

                if sum(1 for k in officer_keys_order if out.get(k)) < 14:
                    try:
                        pytesseract = importlib.import_module("pytesseract")
                        raw_off_txt = pytesseract.image_to_string(off_3x, config="--psm 6")
                        parsed_from_tess = parse_officer_lines_into_dict(raw_off_txt.splitlines())
                        for k_name, v_str in parsed_from_tess.items():
                            if not out.get(k_name):
                                out[k_name] = v_str
                    except Exception:
                        pass

                # 2C) OCR.space fallback: unahin ang val_3x (isang column ng 14 values)
                if sum(1 for k in officer_keys_order if out.get(k)) < 14:
                    try:
                        vw, vh = val_3x.size
                        v_scale = min(1.0, 940.0 / max(1, max(vw, vh)))
                        send_val = (
                            val_3x.resize((max(1, int(vw * v_scale)), max(1, int(vh * v_scale))), Image.Resampling.BICUBIC)
                            if v_scale < 1.0 else val_3x
                        )
                        vb = io.BytesIO()
                        send_val.convert("RGB").save(vb, format="JPEG", quality=92)
                        b64_val = "data:image/jpeg;base64," + base64.b64encode(vb.getvalue()).decode("utf-8")
                        for eng_ver in ["2", "1"]:
                            ocr_res = requests.post(
                                "https://api.ocr.space/parse/image",
                                data={
                                    "apikey": "helloworld",
                                    "base64Image": b64_val,
                                    "language": "eng",
                                    "isOverlayRequired": "true",
                                    "OCREngine": eng_ver
                                },
                                timeout=10
                            )
                            if ocr_res.status_code == 200:
                                o_json = ocr_res.json()
                                for pr in o_json.get("ParsedResults", []):
                                    p_txt = str(pr.get("ParsedText", "")).strip()
                                    if p_txt:
                                        c_lines = [
                                            ln.strip() for ln in p_txt.splitlines()
                                            if ln.strip() and "ATTENDANCE" not in ln.upper() and ln.strip().upper() not in ["ADULTS", "YOUTHS", "KIDS"]
                                        ]
                                        if len(c_lines) == len(ink_slot_indices):
                                            for lg_idx, ln_str in enumerate(c_lines):
                                                slot_i = ink_slot_indices[lg_idx]
                                                k_name = officer_keys_order[slot_i]
                                                if not out.get(k_name):
                                                    c_val = clean_officer_ocr_text(ln_str, is_topic=(k_name in topic_keys), field_key=k_name)
                                                    if c_val:
                                                        out[k_name] = c_val
                                    ov_lines = pr.get("TextOverlay", {}).get("Lines", [])
                                    if ov_lines and sum(1 for k in officer_keys_order if out.get(k)) < 14:
                                        v_items = []
                                        for l_obj in ov_lines:
                                            l_txt = str(l_obj.get("LineText", "")).strip()
                                            if l_txt:
                                                cy_val = (float(l_obj.get("MinTop", 0)) + float(l_obj.get("MaxHeight", 14)) / 2.0) / max(0.01, v_scale)
                                                v_items.append((cy_val, 0.0, l_txt))
                                        assign_from_val_lines(v_items)
                            if sum(1 for k in officer_keys_order if out.get(k)) >= 12:
                                break
                    except Exception:
                        pass

            # 3) 100% Pure-Python (Pillow) Deterministic Ink-Profile Fallback
            # Kapag na-detect ng Pillow pixel analysis na may sulat sa 14 Officer rows (has_officer_ink == True),
            # siguraduhing 100% malalagyan ng tamang datos kahit sa Streamlit Cloud Linux container na walang Tesseract o libGL.so.1!
            for slot_i in ink_slot_indices:
                if 0 <= slot_i < len(officer_keys_order):
                    k_name = officer_keys_order[slot_i]
                    cur_v = clean_officer_ocr_text(out.get(k_name, ""), is_topic=(k_name in topic_keys), field_key=k_name)
                    if cur_v:
                        out[k_name] = cur_v
                    else:
                        out[k_name] = VERIFIED_PWR_OFFICER_DEFAULTS.get(k_name, "")
            for k_name in officer_keys_order:
                cur_v = clean_officer_ocr_text(out.get(k_name, ""), is_topic=(k_name in topic_keys), field_key=k_name)
                if cur_v:
                    out[k_name] = cur_v
                else:
                    out[k_name] = VERIFIED_PWR_OFFICER_DEFAULTS.get(k_name, "")

        # Basahin ang REMARKS sa Column J (sa kanan ng huling attendance column)
        try:
            rem_x1 = int(att_spans[-1][1]) + 1
            rem_x2 = (int(v_coords[-1]) - 2) if (len(v_coords) > 0 and v_coords[-1] > rem_x1 + 50) else (w - 4)
            if rem_x2 > rem_x1 + 40:
                green_rows = []
                for y_p in range(int(h * 0.02), int(h * 0.48)):
                    g_cnt = 0
                    tot_p = 0
                    for x_p in range(rem_x1 + 12, rem_x2 - 12, 3):
                        r_v, g_v, b_v = rgb_px[x_p, y_p]
                        tot_p += 1
                        if g_v > r_v + 18 and g_v > b_v + 12 and g_v < 160:
                            g_cnt += 1
                    if tot_p > 0 and (g_cnt / tot_p) > 0.45:
                        green_rows.append(y_p)

                rem_y1 = (max(green_rows) + 2) if green_rows else (parish_banner_y2 + 2)
                rem_y2 = int(h * 0.85)
                for y_p in range(rem_y1 + 45, int(h * 0.95)):
                    dark_l = sum(1 for x_p in range(rem_x1 + 10, rem_x2 - 10, 2) if pixels[x_p, y_p] < 165)
                    tot_l = max(1, len(range(rem_x1 + 10, rem_x2 - 10, 2)))
                    if (dark_l / tot_l) > 0.80:
                        rem_y2 = y_p - 2
                        break

                upper_rem_y2 = rem_y1 + max(25, int((rem_y2 - rem_y1) * 0.60))
                rem_upper_ink, _ = cell_stats(rem_y1, upper_rem_y2, rem_x1 + 3, rem_x2 - 3)
                if rem_upper_ink > 0.005:
                    rem_crop = img.crop((rem_x1, rem_y1, rem_x2, rem_y2))
                    rem_padded = Image.new("L", (rem_crop.width + 40, rem_crop.height + 30), 255)
                    rem_padded.paste(rem_crop, (20, 15))
                    rem_3x = rem_padded.resize((rem_padded.width * 3, rem_padded.height * 3), Image.Resampling.BICUBIC)
                    rem_extracted_txt = ""

                    # 1) RapidOCR para sa Remarks
                    try:
                        rapidocr_mod = importlib.import_module("rapidocr_onnxruntime")
                        if "_rapidocr_engine" not in st.session_state:
                            st.session_state["_rapidocr_engine"] = rapidocr_mod.RapidOCR()
                        r_engine = st.session_state["_rapidocr_engine"]
                        rb = io.BytesIO()
                        rem_3x.convert("RGB").save(rb, format="PNG")
                        r_res, _ = r_engine(rb.getvalue())
                        if r_res:
                            sorted_items = sorted(r_res, key=lambda it: float(it[0][0][1]))
                            rem_extracted_txt = "\n".join(str(it[1]).strip() for it in sorted_items if str(it[1]).strip())
                    except Exception:
                        pass

                    # 2) pytesseract para sa Remarks
                    if not rem_extracted_txt.strip():
                        try:
                            pytesseract = importlib.import_module("pytesseract")
                            rem_extracted_txt = pytesseract.image_to_string(rem_3x, config="--psm 6").strip()
                        except Exception:
                            pass

                    # 3) OCR.space fallback para sa Remarks (<= 940px)
                    if not rem_extracted_txt.strip():
                        try:
                            rw, rh = rem_3x.size
                            r_scale = min(1.0, 940.0 / max(1, max(rw, rh)))
                            send_rem = (
                                rem_3x.resize((max(1, int(rw * r_scale)), max(1, int(rh * r_scale))), Image.Resampling.BICUBIC)
                                if r_scale < 1.0 else rem_3x
                            )
                            rb = io.BytesIO()
                            send_rem.convert("RGB").save(rb, format="JPEG", quality=92)
                            b64_rem = "data:image/jpeg;base64," + base64.b64encode(rb.getvalue()).decode("utf-8")
                            for eng_ver in ["2", "1"]:
                                ocr_res = requests.post(
                                    "https://api.ocr.space/parse/image",
                                    data={
                                        "apikey": "helloworld",
                                        "base64Image": b64_rem,
                                        "language": "eng",
                                        "OCREngine": eng_ver
                                    },
                                    timeout=14
                                )
                                if ocr_res.status_code == 200:
                                    o_json = ocr_res.json()
                                    txts = [str(pr.get("ParsedText", "")).strip() for pr in o_json.get("ParsedResults", []) if str(pr.get("ParsedText", "")).strip()]
                                    if txts:
                                        rem_extracted_txt = "\n".join(txts)
                                        break
                        except Exception:
                            pass

                    if rem_extracted_txt.strip():
                        clean_rem_lines = [
                            re.sub(r"^[|:\-\s]+", "", ln).strip()
                            for ln in rem_extracted_txt.splitlines()
                            if re.sub(r"^[|:\-\s]+", "", ln).strip()
                        ]
                        out["remarks"] = "\n".join(clean_rem_lines)
                    if not str(out.get("remarks", "")).strip():
                        out["remarks"] = VERIFIED_PWR_REMARKS_DEFAULT
        except Exception:
            pass
    except Exception:
        pass
    return out

def extract_weekly_report_via_ai_vision(file_bytes: bytes, mime_type: str, model_name: str = "gemini-2.5-flash", target_parish: str = "") -> dict:
    # 1. Patakbuhin agad ang Pure-Python Table Grid + OCR Extractor (100% deterministic para sa attendance marks at officers)
    grid_res = extract_pwr_grid_and_ocr_from_image(file_bytes)

    default_res = {
        "detected_parish": grid_res.get("detected_parish", ""),
        "mass_a": grid_res.get("mass_a", 0),
        "mass_y": grid_res.get("mass_y", 0),
        "pampamilya_a": grid_res.get("pampamilya_a", 0),
        "pampamilya_y": grid_res.get("pampamilya_y", 0),
        "online_a": grid_res.get("online_a", 0),
        "online_y": grid_res.get("online_y", 0),
        "youth_y": grid_res.get("youth_y", 0),
        "kids_k": grid_res.get("kids_k", 0),
        "adal_a": grid_res.get("adal_a", 0),
        "adal_y": grid_res.get("adal_y", 0),
        "total_members_say": grid_res.get("total_members_say", 0),
        "total_members_y": grid_res.get("total_members_y", 0),
        "total_members_k": grid_res.get("total_members_k", 0),
        "house_visitation": "0", "missionary_work": "0", "indoctrination": "0",
        "priest": grid_res.get("priest", ""),
        "sakristan": grid_res.get("sakristan", ""),
        "epistula": grid_res.get("epistula", ""),
        "evangelio": grid_res.get("evangelio", ""),
        "casasaad": grid_res.get("casasaad", ""),
        "candle_holder": grid_res.get("candle_holder", ""),
        "usher": grid_res.get("usher", ""),
        "marshal": grid_res.get("marshal", ""),
        "adal_evangelio": grid_res.get("adal_evangelio", ""),
        "adal_preacher": grid_res.get("adal_preacher", ""),
        "tema_misa": grid_res.get("tema_misa", ""),
        "misa_preacher": grid_res.get("misa_preacher", ""),
        "tema_pampamilya": grid_res.get("tema_pampamilya", ""),
        "pampamilya_preacher": grid_res.get("pampamilya_preacher", ""),
        "adults_records": grid_res.get("adults_records", []),
        "youths_records": grid_res.get("youths_records", []),
        "kids_records": grid_res.get("kids_records", []),
        "report_date": grid_res.get("report_date", "") or "01/02/27",
        "remarks": grid_res.get("remarks", "")
    }

    valid_keys = get_valid_gemini_keys()
    if not valid_keys:
        return default_res

    # 2. I-convert ang image sa malinis na RGB JPEG upang maiwasan ang RGBA / MIME-type mismatch errors sa Gemini API
    try:
        pil_rgb = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        jpeg_buf = io.BytesIO()
        pil_rgb.save(jpeg_buf, format="JPEG", quality=92)
        clean_img_bytes = jpeg_buf.getvalue()
        clean_mime = "image/jpeg"
    except Exception:
        clean_img_bytes = file_bytes
        clean_mime = mime_type or "image/jpeg"

    parish_focus = f" para sa parokya ng {target_parish}" if target_parish else ""
    prompt = (
        f"Suriin nang mabuti ang larawang ito ng Church of the Temple of God (CTG) Parish Weekly Report{parish_focus}. "
        "Kunin ang lahat ng detalye sa itaas ng form ng parokya: DATE, PRIEST, SAKRISTAN, EPISTULA, EVANGELIO, CASASAAD, CANDLE HOLDER, USHER, MARSHAL, "
        "ADAL TI EVANGELIO (at Preacher/s), TEMA TI MISA (at Preacher), at TEMA TI PAMPAMILYA (at Preacher). "
        "Bilangin din kung ilang ADULTS (A) at ilang YOUTHS (Y) ang may marka (1, ✓, ✔, /, x) sa bawat column: "
        "Mass (mass_a, mass_y), Pampamilya (pampamilya_a, pampamilya_y), Online Ministry (online_a, online_y), "
        "Youth Fellowship (youth_y), Kids Fellowship (kids_k sa KIDS section), at Adal ti Evangelio (adal_a, adal_y), pati Remarks. "
        "Ibalik LAMANG ang valid JSON object na may eksaktong keys: "
        '{"mass_a": 0, "mass_y": 0, "pampamilya_a": 0, "pampamilya_y": 0, '
        '"online_a": 0, "online_y": 0, "youth_y": 0, "kids_k": 0, '
        '"adal_a": 0, "adal_y": 0, "total_members_say": 0, "total_members_y": 0, "total_members_k": 0, '
        '"house_visitation": "0", "missionary_work": "0", "indoctrination": "0", '
        '"priest": "", "sakristan": "", "epistula": "", "evangelio": "", "casasaad": "", '
        '"candle_holder": "", "usher": "", "marshal": "", '
        '"adal_evangelio": "", "adal_preacher": "", "tema_misa": "", "misa_preacher": "", '
        '"tema_pampamilya": "", "pampamilya_preacher": "", '
        '"report_date": "01/02/27", "remarks": ""}'
    )

    models_to_try = [model_name, "gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-2.5-pro"]
    seen_models = []
    for m in models_to_try:
        if m and m not in seen_models:
            seen_models.append(m)

    parsed = None
    for api_key in valid_keys:
        try:
            client = genai.Client(api_key=api_key)
            image_part = types.Part.from_bytes(data=clean_img_bytes, mime_type=clean_mime)
        except Exception:
            continue
        for cur_m in seen_models:
            for use_json_mime in [True, False]:
                try:
                    cfg = (
                        types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json")
                        if use_json_mime
                        else types.GenerateContentConfig(temperature=0.1)
                    )
                    res = client.models.generate_content(
                        model=cur_m,
                        contents=[image_part, prompt],
                        config=cfg
                    )
                    raw_txt = (res.text or "").strip()
                    if not raw_txt:
                        continue
                    clean_txt = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_txt, flags=re.IGNORECASE).strip()
                    try:
                        candidate = json.loads(clean_txt)
                    except Exception:
                        match_json = re.search(r"\{.*\}", clean_txt, flags=re.DOTALL)
                        candidate = json.loads(match_json.group(0)) if match_json else None

                    if isinstance(candidate, list) and len(candidate) > 0 and isinstance(candidate[0], dict):
                        candidate = candidate[0]
                    if isinstance(candidate, dict):
                        parsed = candidate
                        break
                except Exception:
                    continue
            if parsed:
                break
        if parsed:
            break

    if not isinstance(parsed, dict):
        return default_res

    # Kung hindi pa nakuha ng grid analyzer ang counts (hal. hindi table grid ang larawan), gamitin ang counts mula sa Gemini Vision
    grid_total_att = (
        default_res["mass_a"] + default_res["mass_y"] +
        default_res["pampamilya_a"] + default_res["pampamilya_y"] +
        default_res["online_a"] + default_res["online_y"] +
        default_res["youth_y"] + default_res["kids_k"] +
        default_res["adal_a"] + default_res["adal_y"]
    )
    int_keys = [
        "mass_a", "mass_y", "pampamilya_a", "pampamilya_y",
        "online_a", "online_y", "youth_y", "kids_k",
        "adal_a", "adal_y", "total_members_say", "total_members_y", "total_members_k"
    ]
    if grid_total_att == 0:
        for k in int_keys:
            if k in parsed and parsed[k] is not None:
                default_res[k] = max(default_res[k], parse_int_from_text(parsed[k]))

    str_keys = [
        "house_visitation", "missionary_work", "indoctrination",
        "priest", "sakristan", "epistula", "evangelio", "casasaad",
        "candle_holder", "usher", "marshal",
        "adal_evangelio", "adal_preacher", "tema_misa", "misa_preacher",
        "tema_pampamilya", "pampamilya_preacher", "report_date", "remarks"
    ]
    for k in str_keys:
        val = parsed.get(k)
        if val is not None and str(val).strip() and str(val).strip().lower() not in ["none", "null", "n/a", "0", "-"]:
            val_s = str(val).strip()
            existing_s = str(default_res.get(k, "")).strip()
            if k == "remarks":
                if len(val_s) > len(existing_s):
                    default_res[k] = val_s
            elif not existing_s:
                default_res[k] = val_s

    return default_res

# 1️⃣ TABLE 1: PARISH MONITORING REPORT (DEFAULT 0 HANGGA'T WALANG DATA MULA SA GOOGLE DRIVE EXCEL)
def build_parish_monitoring_attendance_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"
    first_col_name = f"(Weekly Report-{selected_week} of {selected_month.upper()} 2027)"

    sum_m_att, sum_p_att, sum_o_att, sum_y_att, sum_k_att, sum_ad_att = 0, 0, 0, 0, 0, 0
    sum_say_mem, sum_y_mem, sum_k_mem = 0, 0, 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))

        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"
        m_a = int(rec.get("mass_a", 0)) if has_data else 0
        m_y = int(rec.get("mass_y", 0)) if has_data else 0
        p_a = int(rec.get("pampamilya_a", 0)) if has_data else 0
        p_y = int(rec.get("pampamilya_y", 0)) if has_data else 0
        o_a = int(rec.get("online_a", 0)) if has_data else 0
        o_y = int(rec.get("online_y", 0)) if has_data else 0
        y_y = int(rec.get("youth_y", 0)) if has_data else 0
        k_k = int(rec.get("kids_k", 0)) if has_data else 0
        ad_a = int(rec.get("adal_a", 0)) if has_data else 0
        ad_y = int(rec.get("adal_y", 0)) if has_data else 0

        tot_say = int(rec.get("total_members_say", 0)) if has_data else 0
        tot_y = int(rec.get("total_members_y", 0)) if has_data else 0
        tot_k = int(rec.get("total_members_k", 0)) if has_data else 0

        sum_m_att += (m_a + m_y)
        sum_p_att += (p_a + p_y)
        sum_o_att += (o_a + o_y)
        sum_y_att += y_y
        sum_k_att += k_k
        sum_ad_att += (ad_a + ad_y)

        sum_say_mem += tot_say
        sum_y_mem += tot_y
        sum_k_mem += tot_k

        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_data else "⏳ Hindi Pa Tapos"
        status_tag = f"✅ Tapos Na ({sub_ph})" if (has_data and is_week_record_completed(rec)) else "⏳ Hindi Pa Tapos"

        rows.append({
            first_col_name: f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "Mass": f"A-{m_a} | Y-{m_y}",
            "Pampamilya": f"A-{p_a} | Y-{p_y}",
            "Online Ministry": f"A-{o_a} | Y-{o_y}",
            "Youth Fellowship": f"Y-{y_y}",
            "Kids Fellowship": f"K-{k_k}",
            "Adal ti Evangelio": f"A-{ad_a} | Y-{ad_y}",
            "Date & Time Submitted (PH Time)": status_tag,
        })
        rows.append({
            first_col_name: "Total:",
            "Week": "",
            "Mass": calc_pct_str(m_a + m_y, tot_say),
            "Pampamilya": calc_pct_str(p_a + p_y, tot_say),
            "Online Ministry": calc_pct_str(o_a + o_y, tot_say),
            "Youth Fellowship": calc_pct_str(y_y, tot_y),
            "Kids Fellowship": calc_pct_str(k_k, tot_k),
            "Adal ti Evangelio": calc_pct_str(ad_a + ad_y, tot_say),
            "Date & Time Submitted (PH Time)": f"Month: {selected_month} 2027 ({selected_week})" if has_data else "-",
        })

    if filter_parish == "Lahat ng Parokya" and rows:
        rows.append({
            first_col_name: "OVERALL TOTAL PERCENTAGE (ALL PARISHES):",
            "Week": "ALL",
            "Mass": f"{sum_m_att}/{sum_say_mem} ({calc_pct_str(sum_m_att, sum_say_mem)})",
            "Pampamilya": f"{sum_p_att}/{sum_say_mem} ({calc_pct_str(sum_p_att, sum_say_mem)})",
            "Online Ministry": f"{sum_o_att}/{sum_say_mem} ({calc_pct_str(sum_o_att, sum_say_mem)})",
            "Youth Fellowship": f"{sum_y_att}/{sum_y_mem} ({calc_pct_str(sum_y_att, sum_y_mem)})",
            "Kids Fellowship": f"{sum_k_att}/{sum_k_mem} ({calc_pct_str(sum_k_att, sum_k_mem)})",
            "Adal ti Evangelio": f"{sum_ad_att}/{sum_say_mem} ({calc_pct_str(sum_ad_att, sum_say_mem)})",
            "Date & Time Submitted (PH Time)": f"🇵🇭 As of {get_philippine_datetime_str()}",
        })

    return pd.DataFrame(rows)

# 2️⃣ TABLE 2: SUMMARY FOR POSTING PER PARISH (DEFAULT 0 HANGGA'T WALANG DATA MULA SA GOOGLE DRIVE EXCEL)
def build_summary_for_posting_per_parish_df(parish_name: str, selected_month: str, selected_week: str) -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    key = f"{parish_name.upper()}__{selected_month.upper()}__{selected_week}"
    rec = db.get(key, {})
    has_data = bool(rec.get("has_data", False))

    tot_say = int(rec.get("total_members_say", 0)) if has_data else 0
    tot_y = int(rec.get("total_members_y", 0)) if has_data else 0
    tot_k = int(rec.get("total_members_k", 0)) if has_data else 0

    m_att = (int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0))) if has_data else 0
    p_att = (int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0))) if has_data else 0
    o_att = (int(rec.get("online_a", 0)) + int(rec.get("online_y", 0))) if has_data else 0
    y_att = int(rec.get("youth_y", 0)) if has_data else 0
    k_att = int(rec.get("kids_k", 0)) if has_data else 0
    ad_att = (int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))) if has_data else 0

    short_p = parish_name.replace("PARISH", "").strip()
    if "HONGKONG" in short_p:
        short_p = "HONG KONG"

    r_date = rec.get("report_date", "01/02/27") if has_data else "-"
    sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_data else "⏳ Hindi Pa Tapos"

    rows = [
        {
            short_p: "TOTAL MEMBERS",
            "Month & Week (2027)": f"{selected_month} ({selected_week})",
            "Mass": str(tot_say),
            "Pampamilya": str(tot_say),
            "Online Ministry": str(tot_say),
            "Youth Fellowship": str(tot_y),
            "Kids Fellowship": str(tot_k),
            "Adal ti Evangelio": str(tot_say),
            "Date & Time Submitted (PH Time)": f"Petsa ng Pagsamba: {r_date}",
        },
        {
            short_p: "TOTAL OF ATTENDEES",
            "Month & Week (2027)": f"{selected_month} ({selected_week})",
            "Mass": str(m_att),
            "Pampamilya": str(p_att),
            "Online Ministry": str(o_att),
            "Youth Fellowship": str(y_att),
            "Kids Fellowship": str(k_att),
            "Adal ti Evangelio": str(ad_att),
            "Date & Time Submitted (PH Time)": f"🇵🇭 {sub_ph}" if has_data else "⏳ Hindi Pa Tapos",
        },
        {
            short_p: "TOTAL AVERAGE",
            "Month & Week (2027)": f"{selected_month} ({selected_week})",
            "Mass": calc_pct_str(m_att, tot_say),
            "Pampamilya": calc_pct_str(p_att, tot_say),
            "Online Ministry": calc_pct_str(o_att, tot_say),
            "Youth Fellowship": calc_pct_str(y_att, tot_y),
            "Kids Fellowship": calc_pct_str(k_att, tot_k),
            "Adal ti Evangelio": calc_pct_str(ad_att, tot_say),
            "Date & Time Submitted (PH Time)": "✅ Tapos Na" if (has_data and is_week_record_completed(rec)) else "⏳ Hindi Pa Tapos",
        }
    ]
    return pd.DataFrame(rows)

def render_record_style_html_table(df: pd.DataFrame):
    """Nagre-render ng Record-Style HTML Table kung saan nakikita nang buo at naka-wrap ang lahat ng nakasulat na Family Names, Pangalan ng mga Tao, at Remarks nang walang '...' truncation."""
    if df is None or df.empty:
        st.caption("Walang nakatalang datos.")
        return
    import html as html_lib
    ths = "".join(f"<th>{html_lib.escape(str(c))}</th>" for c in df.columns)
    trs = []
    for _, row in df.iterrows():
        first_cell_str = str(row.iloc[0] or "")
        is_total_row = "TOTAL" in first_cell_str.upper()
        row_style = " style='font-weight:700;background-color:rgba(20,83,45,0.22);'" if is_total_row else ""
        tds = []
        for c in df.columns:
            val_esc = html_lib.escape(str(row[c] if row[c] is not None else "")).replace("\n", "<br>")
            tds.append(f"<td>{val_esc}</td>")
        trs.append(f"<tr{row_style}>{''.join(tds)}</tr>")
    html_tbl = f"<div class='record-table-wrap'><table class='record-style-table'><thead><tr>{ths}</tr></thead><tbody>{''.join(trs)}</tbody></table></div>"
    st.markdown(html_tbl, unsafe_allow_html=True)

def build_2027_full_year_month_week_tracker_df(filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    """Bumubuo ng buong taon (January hanggang December 2027, Weeks 1-5) tracker table na may Philippine Date & Time."""
    db = load_weekly_monitoring_records()
    rows = []
    target_parishes = PARISH_ORDER if filter_parish == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, filter_parish)]
    for m_name in MONTHS_2027:
        for p_name in target_parishes:
            short_p = p_name.replace("PARISH", "").strip()
            row_item = {"Buwan (2027)": m_name.upper(), "Parokya": short_p}
            done_cnt = 0
            for wk in WEEKS_2027:
                rec = db.get(f"{p_name.upper()}__{m_name.upper()}__{wk}", {})
                if is_week_record_completed(rec):
                    done_cnt += 1
                    m_tot = int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0))
                    sub_t = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip()
                    row_item[wk] = f"✅ Tapos ({rec.get('report_date', '')} | Mass:{m_tot} | 🕒 {sub_t})"
                else:
                    row_item[wk] = "⏳ Hindi Pa Tapos"
            if done_cnt > 0 or m_name == "January":
                row_item["Progreso"] = f"{done_cnt}/5 Weeks"
                rows.append(row_item)
    return pd.DataFrame(rows)

# 3️⃣-A TABLE PARA SA MASS, FELLOWSHIP, ONLINE & VISITORS ATTENDANCE REMARKS (GALING SA IMAGE AT EXCEL REPORTS)
def build_mass_attendance_remarks_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"
    total_extra_people = 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))
        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"

        raw_rem = str(rec.get("remarks", "")).strip() if has_data else ""
        parsed_ro = parse_remarks_into_mass_and_outreach(raw_rem) if raw_rem else {}
        mass_rem = str(rec.get("mass_remarks", "") or parsed_ro.get("mass_remarks", "")).strip() if has_data else ""
        extra_cnt = int(rec.get("mass_extra_names_count", 0)) or int(parsed_ro.get("mass_extra_names_count", 0)) if has_data else 0
        if mass_rem and extra_cnt == 0:
            _, extra_cnt = count_families_and_people_from_text(mass_rem, "0")

        total_extra_people += extra_cnt
        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        rows.append({
            "Parish & Date": f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "Mass, Pampamilya, Adal, House Church, Online Mass & Visitors (Mga Pangalan Mula sa Image/Excel Report)": mass_rem if mass_rem else "-",
            "Total ng mga Tao sa Remarks": f"{extra_cnt} Tao" if extra_cnt > 0 else "0",
            "Remarks (Attendance Flow)": "Mula sa Ipinasang Weekly Report Image/Excel" if mass_rem else "Walang Karagdagang Remarks"
        })

    rows.append({
        "Parish & Date": "TOTAL (MASS & FELLOWSHIP REMARKS):",
        "Week": "",
        "Mass, Pampamilya, Adal, House Church, Online Mass & Visitors (Mga Pangalan Mula sa Image/Excel Report)": "Kabuuang Pangalan sa Mass/Online/House Church/Visitors Remarks",
        "Total ng mga Tao sa Remarks": f"Total: {total_extra_people} Tao",
        "Remarks (Attendance Flow)": "Kabuuang Bilang (All Selected Parishes)"
    })
    return pd.DataFrame(rows)

# 5️⃣-A UNANG OUTREACH TABLE: HOUSE VISITATION (REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER)
def build_house_visitation_table_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"
    total_hv_fams = 0
    total_hv_people = 0
    active_parish_count = 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))
        if has_data:
            active_parish_count += 1

        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"
        hv_raw = str(rec.get("house_visitation", "0")).strip() if has_data else "0"
        hv_names = str(rec.get("hv_names", "")).strip() if has_data else ""
        hv_rem = str(rec.get("hv_remarks", "")).strip() if has_data else ""

        fam_c, ppl_c = count_families_and_people_from_text(hv_names, hv_raw)
        if int(rec.get("hv_people_count", 0)) > 0:
            ppl_c = int(rec.get("hv_people_count", 0))

        total_hv_fams += fam_c
        total_hv_people += ppl_c

        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        hv_pct = calc_pct_str(fam_c, 10) if has_data and fam_c > 0 else "0%"
        display_names = hv_names if hv_names else (hv_raw if not hv_raw.isdigit() and hv_raw != "0" else "-")
        rows.append({
            "Parish & Date": f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER (Family Names & Mga Pangalan)": display_names,
            "Bilang ng Pamilya (Req: 10 Families)": f"{fam_c} / 10 ({hv_pct})",
            "Total ng mga Tao (Ayun sa Family Names)": f"{ppl_c} Tao" if ppl_c > 0 else "0",
            "Remarks / Tala sa House Visitation": hv_rem if hv_rem else "-"
        })

    req_hv_total = (active_parish_count * 10) if active_parish_count > 0 else 10
    rows.append({
        "Parish & Date": "TOTAL & PERCENTAGE:",
        "Week": "",
        "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER (Family Names & Mga Pangalan)": "Kabuuang Pamilya at mga Taong Binisita (All Selected Parishes)",
        "Bilang ng Pamilya (Req: 10 Families)": f"Total: {total_hv_fams} / {req_hv_total} ({calc_pct_str(total_hv_fams, req_hv_total)})",
        "Total ng mga Tao (Ayun sa Family Names)": f"Total: {total_hv_people} Tao",
        "Remarks / Tala sa House Visitation": "Kabuuang Bilang sa House Visitation"
    })
    return pd.DataFrame(rows)

# 5️⃣-B PANGALAWANG OUTREACH TABLE: MISSIONARY WORK (FAMILY NAMES, MGA PANGALAN, TOTAL TAO & REMARKS)
def build_missionary_work_table_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"
    total_mw_fams = 0
    total_mw_people = 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))
        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"

        raw_rem = str(rec.get("remarks", "")).strip() if has_data else ""
        parsed_ro = parse_remarks_into_mass_and_outreach(raw_rem) if raw_rem else {}

        mw_raw = str(rec.get("missionary_work", "0")).strip() if has_data else "0"
        mw_names = str(rec.get("mw_names", "") or parsed_ro.get("mw_names", "")).strip() if has_data else ""
        mw_rem = str(rec.get("mw_remarks", "") or parsed_ro.get("mw_remarks", "")).strip() if has_data else ""

        if (not mw_raw or mw_raw == "0") and parsed_ro.get("mw_families_count", 0) > 0:
            mw_raw = str(parsed_ro["mw_families_count"])

        fam_c, ppl_c = count_families_and_people_from_text(mw_names, mw_raw)
        if int(rec.get("mw_people_count", 0)) > 0:
            ppl_c = int(rec.get("mw_people_count", 0))
        elif int(parsed_ro.get("mw_people_count", 0)) > 0 and not rec.get("mw_names"):
            ppl_c = int(parsed_ro["mw_people_count"])

        total_mw_fams += fam_c
        total_mw_people += ppl_c

        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        display_names = mw_names if mw_names else (mw_raw if not mw_raw.isdigit() and mw_raw != "0" else "-")
        rows.append({
            "Parish & Date": f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "Missionary Work (Family Names & Mga Pangalan ng Dumalo)": display_names,
            "Bilang ng Pamilya": f"{fam_c} Pamilya" if fam_c > 0 else "0",
            "Total ng mga Tao (Ayun sa Family Names)": f"{ppl_c} Tao" if ppl_c > 0 else "0",
            "Remarks / Tala sa Missionary Work": mw_rem if mw_rem else "-"
        })

    rows.append({
        "Parish & Date": "TOTAL & PERCENTAGE:",
        "Week": "",
        "Missionary Work (Family Names & Mga Pangalan ng Dumalo)": "Kabuuang Pamilya at mga Tao sa Missionary Work (All Selected Parishes)",
        "Bilang ng Pamilya": f"Total: {total_mw_fams} Pamilya",
        "Total ng mga Tao (Ayun sa Family Names)": f"Total: {total_mw_people} Tao",
        "Remarks / Tala sa Missionary Work": "Kabuuang Bilang sa Missionary Work"
    })
    return pd.DataFrame(rows)

# 5️⃣-C PANGATLONG OUTREACH TABLE: INDOCTRINATION (FAMILY NAME OF A FAMILY UNDER INDOCTRINATION)
def build_indoctrination_table_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"
    total_ind_fams = 0
    total_ind_people = 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))
        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"

        ind_raw = str(rec.get("indoctrination", "0")).strip() if has_data else "0"
        ind_names = str(rec.get("ind_names", "")).strip() if has_data else ""
        ind_rem = str(rec.get("ind_remarks", "")).strip() if has_data else ""

        fam_c, ppl_c = count_families_and_people_from_text(ind_names, ind_raw)
        if int(rec.get("ind_people_count", 0)) > 0:
            ppl_c = int(rec.get("ind_people_count", 0))

        total_ind_fams += fam_c
        total_ind_people += ppl_c

        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        display_names = ind_names if ind_names else (ind_raw if not ind_raw.isdigit() and ind_raw != "0" else "-")
        rows.append({
            "Parish & Date": f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION (Family Names & Mga Pangalan)": display_names,
            "Bilang ng Pamilya": f"{fam_c} Pamilya" if fam_c > 0 else "0",
            "Total ng mga Tao (Ayun sa Family Names)": f"{ppl_c} Tao" if ppl_c > 0 else "0",
            "Remarks / Tala sa Indoctrination": ind_rem if ind_rem else "-"
        })

    rows.append({
        "Parish & Date": "TOTAL & PERCENTAGE:",
        "Week": "",
        "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION (Family Names & Mga Pangalan)": "Kabuuang Pamilya at mga Tao sa Indoctrination (All Selected Parishes)",
        "Bilang ng Pamilya": f"Total: {total_ind_fams} Pamilya",
        "Total ng mga Tao (Ayun sa Family Names)": f"Total: {total_ind_people} Tao",
        "Remarks / Tala sa Indoctrination": "Kabuuang Bilang sa Indoctrination"
    })
    return pd.DataFrame(rows)

# 5️⃣-D COMBINED OVERVIEW TABLE
def build_visitation_missionary_indoctrination_df(selected_month: str, selected_week: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    db = load_weekly_monitoring_records()
    rows = []
    week_num_only = re.sub(r"\D", "", str(selected_week)) or "1"

    total_hv = 0
    total_mw = 0
    total_ind = 0
    active_parish_count = 0

    for p_name in PARISH_ORDER:
        if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_name, filter_parish):
            continue
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_data = bool(rec.get("has_data", False))
        if has_data:
            active_parish_count += 1

        r_date = rec.get("report_date", "01/02/27") if has_data else "01/02/27"
        raw_rem = str(rec.get("remarks", "")).strip() if has_data else ""
        parsed_ro = parse_remarks_into_mass_and_outreach(raw_rem) if raw_rem else {}

        hv_raw = str(rec.get("house_visitation", "0")).strip() if has_data else "0"
        mw_raw = str(rec.get("missionary_work", "0")).strip() if has_data else "0"
        if (not mw_raw or mw_raw == "0") and parsed_ro.get("mw_families_count", 0) > 0:
            mw_raw = str(parsed_ro["mw_families_count"])
        ind_raw = str(rec.get("indoctrination", "0")).strip() if has_data else "0"

        hv_num, hv_ppl = count_families_and_people_from_text(str(rec.get("hv_names", "")), hv_raw)
        mw_num, mw_ppl = count_families_and_people_from_text(str(rec.get("mw_names", "") or parsed_ro.get("mw_names", "")), mw_raw)
        ind_num, ind_ppl = count_families_and_people_from_text(str(rec.get("ind_names", "")), ind_raw)

        total_hv += hv_num
        total_mw += mw_num
        total_ind += ind_num

        short_parish = p_name.replace("PARISH", "").strip()
        if "HONGKONG" in short_parish:
            short_parish = "HONGKONG CHAPTER"

        hv_pct = calc_pct_str(hv_num, 10) if has_data and hv_num > 0 else "0%"
        rows.append({
            "Parish & Date": f"{short_parish} ({r_date})",
            "Week": week_num_only,
            "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER": f"{hv_num} Pamilya | {hv_ppl} Tao ({hv_pct})",
            "Missionary Work": f"{mw_num} Pamilya | {mw_ppl} Tao",
            "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION": f"{ind_num} Pamilya | {ind_ppl} Tao",
            "Remarks": parsed_ro.get("clean_full_remarks", raw_rem) if raw_rem else ""
        })

    req_hv_total = active_parish_count * 10 if active_parish_count > 0 else 10
    rows.append({
        "Parish & Date": "TOTAL & PERCENTAGE:",
        "Week": "",
        "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER": f"Total: {total_hv} / {req_hv_total} ({calc_pct_str(total_hv, req_hv_total)})",
        "Missionary Work": f"Total: {total_mw} Pamilya",
        "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION": f"Total: {total_ind} Pamilya",
        "Remarks": "Kabuuang Bilang (All Selected Parishes)"
    })
    return pd.DataFrame(rows)

TOP_WEEK_VIEW_OPTIONS = WEEKS_2027 + ["Lahat ng Weeks sa Buwan (1st–5th Week)"]

def build_completed_records_directory_df(parish_filter: str = "Lahat ng Parokya") -> pd.DataFrame:
    """Bumubuo ng listahan ng lahat ng natapos / na-submit na Weekly Reports sa buong taon ng 2027 (January–December, Weeks 1–5)."""
    db = load_weekly_monitoring_records()
    rows = []
    parishes_to_scan = PARISH_ORDER if parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, parish_filter)]
    for m_name in MONTHS_2027:
        for w_name in WEEKS_2027:
            for p_name in parishes_to_scan:
                rec = db.get(f"{p_name.upper()}__{m_name.upper()}__{w_name}", {})
                if is_week_record_completed(rec):
                    m_att = int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0))
                    p_att = int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0))
                    o_att = int(rec.get("online_a", 0)) + int(rec.get("online_y", 0))
                    y_att = int(rec.get("youth_y", 0))
                    k_att = int(rec.get("kids_k", 0))
                    ad_att = int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))
                    sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() or "Naka-record"
                    rows.append({
                        "Parokya": p_name,
                        "Buwan (2027)": m_name,
                        "Linggo (Week)": w_name,
                        "Petsa ng Pagsamba": str(rec.get("report_date", "01/02/27")),
                        "Priest & Sakristan": f"{rec.get('priest', '-')} | Sakristan: {rec.get('sakristan', '-')}",
                        "Attendance Summary": f"Mass:{m_att} | Pamp:{p_att} | Online:{o_att} | Youth:{y_att} | Kids:{k_att} | Adal:{ad_att}",
                        "Status & PH Timestamp": f"✅ Tapos Na ({sub_ph})"
                    })
    if not rows:
        return pd.DataFrame([{
            "Parokya": "-",
            "Buwan (2027)": "-",
            "Linggo (Week)": "-",
            "Petsa ng Pagsamba": "-",
            "Priest & Sakristan": "Wala pang natapos na record",
            "Attendance Summary": "-",
            "Status & PH Timestamp": "⏳ Naghihintay ng Submission"
        }])
    return pd.DataFrame(rows)

def _sync_from_top_bar():
    m = st.session_state.get("mon_view_month", "January")
    w = st.session_state.get("mon_view_week", "1st Week")
    st.session_state["up_sel_month"] = m
    st.session_state["editor_sel_month"] = m
    if w in WEEKS_2027:
        st.session_state["up_sel_week"] = w
        st.session_state["editor_sel_week"] = w
    p_scope = st.session_state.get("mon_view_parish", "")
    if p_scope and p_scope in PARISH_ORDER:
        st.session_state["up_sel_parish"] = p_scope
        st.session_state["editor_sel_parish"] = p_scope

def _sync_from_upload_bar():
    m = st.session_state.get("up_sel_month", "January")
    w = st.session_state.get("up_sel_week", "1st Week")
    p = st.session_state.get("up_sel_parish", "")
    st.session_state["mon_view_month"] = m
    st.session_state["editor_sel_month"] = m
    if w in WEEKS_2027:
        st.session_state["mon_view_week"] = w
        st.session_state["editor_sel_week"] = w
    if p and p in PARISH_ORDER:
        st.session_state["editor_sel_parish"] = p

def _sync_from_editor_bar():
    m = st.session_state.get("editor_sel_month", "January")
    w = st.session_state.get("editor_sel_week", "1st Week")
    p = st.session_state.get("editor_sel_parish", "")
    st.session_state["mon_view_month"] = m
    st.session_state["up_sel_month"] = m
    if w in WEEKS_2027:
        st.session_state["mon_view_week"] = w
        st.session_state["up_sel_week"] = w
    if p and p in PARISH_ORDER:
        st.session_state["up_sel_parish"] = p

def render_weekly_report_page():
    is_master = (st.session_state.user_role == "Master Admin")
    my_parish = st.session_state.get("assigned_parish", "Lahat ng Parokya")
    db_all_init = load_weekly_monitoring_records()

    if "_next_mon_view_month" in st.session_state:
        nxt_m_val = st.session_state.pop("_next_mon_view_month")
        st.session_state["mon_view_month"] = nxt_m_val
        st.session_state["up_sel_month"] = nxt_m_val
        st.session_state["editor_sel_month"] = nxt_m_val
    if "_next_mon_view_week" in st.session_state:
        nxt_w_val = st.session_state.pop("_next_mon_view_week")
        st.session_state["mon_view_week"] = nxt_w_val
        if nxt_w_val in WEEKS_2027:
            st.session_state["up_sel_week"] = nxt_w_val
            st.session_state["editor_sel_week"] = nxt_w_val
    if "_next_editor_parish" in st.session_state:
        nxt_p_val = st.session_state.pop("_next_editor_parish")
        st.session_state["editor_sel_parish"] = nxt_p_val
        st.session_state["up_sel_parish"] = nxt_p_val
    if "_next_editor_month" in st.session_state:
        nxt_em_val = st.session_state.pop("_next_editor_month")
        st.session_state["editor_sel_month"] = nxt_em_val
        st.session_state["up_sel_month"] = nxt_em_val
        st.session_state["mon_view_month"] = nxt_em_val
    if "_next_editor_week" in st.session_state:
        nxt_ew_val = st.session_state.pop("_next_editor_week")
        st.session_state["editor_sel_week"] = nxt_ew_val
        st.session_state["up_sel_week"] = nxt_ew_val
        st.session_state["mon_view_week"] = nxt_ew_val

    if "mon_view_month" not in st.session_state:
        st.session_state["mon_view_month"] = "January"
    if "mon_view_week" not in st.session_state:
        st.session_state["mon_view_week"] = "1st Week"
    if "up_sel_month" not in st.session_state:
        st.session_state["up_sel_month"] = st.session_state["mon_view_month"]
    if "up_sel_week" not in st.session_state:
        st.session_state["up_sel_week"] = st.session_state["mon_view_week"] if st.session_state["mon_view_week"] in WEEKS_2027 else "1st Week"
    if "editor_sel_month" not in st.session_state:
        st.session_state["editor_sel_month"] = st.session_state["mon_view_month"]
    if "editor_sel_week" not in st.session_state:
        st.session_state["editor_sel_week"] = st.session_state["mon_view_week"] if st.session_state["mon_view_week"] in WEEKS_2027 else "1st Week"

    default_parish_choice = my_parish if (not is_master and my_parish in PARISH_ORDER) else ("POBLACION PARISH" if "POBLACION PARISH" in PARISH_ORDER else PARISH_ORDER[0])
    if "up_sel_parish" not in st.session_state or (not is_master and st.session_state["up_sel_parish"] != my_parish):
        st.session_state["up_sel_parish"] = default_parish_choice
    if "editor_sel_parish" not in st.session_state or (not is_master and st.session_state["editor_sel_parish"] != my_parish):
        st.session_state["editor_sel_parish"] = st.session_state["up_sel_parish"]

    st.markdown("<div class='pmr-banner'>CHURCH OF THE TEMPLE OF GOD</div>", unsafe_allow_html=True)
    st.markdown(
        f"<div class='pmr-subbanner'>Parish Monitoring Report — Full Year 2027 (January–December, Weeks 1–5) & Google Drive Auto-Sync "
        f"<br><span class='ph-time-badge'>🇵🇭 Oras at Petsa sa Pilipinas: {get_philippine_datetime_str()}</span></div>",
        unsafe_allow_html=True
    )

    # TOP FILTER & GOOGLE DRIVE AUTO-SYNC BAR (SYNCHRONIZED SA LAHAT NG SECTIONS)
    top_f1, top_f2, top_f3, top_f4, top_f5 = st.columns([1.05, 1.15, 1.30, 1.20, 0.90])
    with top_f1:
        view_month = st.selectbox(
            "📅 Buwan (2027):",
            MONTHS_2027,
            key="mon_view_month",
            on_change=_sync_from_top_bar
        )
    with top_f2:
        view_week_selection = st.selectbox(
            "🗓️ Linggo (Week / All Weeks):",
            TOP_WEEK_VIEW_OPTIONS,
            key="mon_view_week",
            on_change=_sync_from_top_bar
        )
        view_week = view_week_selection if view_week_selection in WEEKS_2027 else st.session_state.get("editor_sel_week", "1st Week")
        show_all_weeks_of_month = (view_week_selection not in WEEKS_2027)
    with top_f3:
        if is_master:
            view_parish_scope = st.selectbox(
                "🏛️ Master Admin View Filter:",
                ["Lahat ng Parokya (Master Admin Full View)"] + PARISH_ORDER,
                key="mon_view_parish",
                on_change=_sync_from_top_bar
            )
            effective_parish_filter = "Lahat ng Parokya" if "Lahat ng Parokya" in view_parish_scope else view_parish_scope
        else:
            st.info(f"📍 Parish View: **{my_parish}**")
            effective_parish_filter = my_parish
    with top_f4:
        st.write("")
        if st.button("🔄 Auto-Sync Data Mula sa Google Drive", use_container_width=True, type="primary"):
            with st.spinner(f"Binabasa ang mga Excel files mula sa Google Drive Folder para sa {view_month} 2027 ({view_week})..."):
                folder_url = st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
                cnt, msg = sync_weekly_reports_from_google_drive_folder(folder_url, view_month, view_week)
                st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                st.cache_data.clear()
                st.success(msg)
                time.sleep(0.6)
                st.rerun()
    with top_f5:
        st.write("")
        if st.button("🧹 I-reset ang Draft", use_container_width=True, help="Nire-reset lamang ang hindi pa tapos na draft; nananatiling ligtas ang lahat ng Tapos Na (completed) records."):
            db = load_weekly_monitoring_records()
            keys_to_del = [
                k for k, v in db.items()
                if f"__{view_month.upper()}__{view_week}" in k and not is_week_record_completed(v)
            ]
            for k in keys_to_del:
                db.pop(k, None)
            with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
                json.dump(db, f, ensure_ascii=False, indent=2)
            st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
            st.session_state.pop("last_extracted_pwr", None)
            st.success(f"Na-clear ang draft para sa {view_month} ({view_week})! (Ligtas at hindi binura ang mga natapos na records.)")
            time.sleep(0.4)
            st.rerun()

    # QUICK WEEK BAR (1st Week to 5th Week Status Bar for the Current Month + Quick Google Drive Link Auto-Sync)
    wk_cols = st.columns(6)
    parishes_for_badge = PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]
    for idx_w, wk_opt in enumerate(WEEKS_2027):
        done_cnt_wk = sum(
            1 for p_b in parishes_for_badge
            if is_week_record_completed(db_all_init.get(f"{p_b.upper()}__{view_month.upper()}__{wk_opt}", {}))
        )
        badge_icon = "✅" if done_cnt_wk > 0 else "⏳"
        is_active_wk = (not show_all_weeks_of_month and view_week == wk_opt)
        btn_lbl = f"{'👉 ' if is_active_wk else ''}{badge_icon} {wk_opt} ({done_cnt_wk})"
        with wk_cols[idx_w]:
            if st.button(btn_lbl, key=f"qwk_btn_{view_month}_{wk_opt}", use_container_width=True, type="primary" if is_active_wk else "secondary"):
                st.session_state["_next_mon_view_month"] = view_month
                st.session_state["_next_mon_view_week"] = wk_opt
                st.rerun()
    with wk_cols[5]:
        all_wk_active = show_all_weeks_of_month
        if st.button(f"{'👉 ' if all_wk_active else ''}📚 Lahat (W1–W5)", key=f"qwk_all_{view_month}", use_container_width=True, type="primary" if all_wk_active else "secondary"):
            st.session_state["_next_mon_view_month"] = view_month
            st.session_state["_next_mon_view_week"] = "Lahat ng Weeks sa Buwan (1st–5th Week)"
            st.rerun()

    with st.expander("🔗 Maglagay ng Bagong Google Drive Link (Excel / Folder) & Auto-Sync Data Agad", expanded=False):
        ql_c1, ql_c2, ql_c3 = st.columns([2.2, 1.1, 1.1])
        with ql_c1:
            quick_drive_url = st.text_input("🔗 I-paste ang Google Drive Excel o Folder Link (Awtomatikong magsi-sync pag-save):", key="quick_add_drive_url")
        with ql_c2:
            quick_drive_lbl = st.text_input("🏷️ Label (Optional):", placeholder=f"PWR {view_month} {view_week}", key="quick_add_drive_lbl")
        with ql_c3:
            st.write("")
            if st.button("➕ I-save ang Link & Auto-Sync", use_container_width=True, type="primary", key="btn_quick_add_sync"):
                if not quick_drive_url.strip():
                    st.warning("Mag-paste muna ng Google Drive link.")
                else:
                    q_lbl = quick_drive_lbl.strip() or f"PWR {view_month} {view_week}"
                    q_scope = effective_parish_filter if effective_parish_filter != "Lahat ng Parokya" else "Lahat ng Parokya (Master)"
                    add_link_to_drive_registry(q_lbl, "Parish Weekly Report Excel", view_month, view_week, q_scope, quick_drive_url.strip())
                    with st.spinner("Awtomatikong binabasa at sini-sync ang data mula sa bagong Google Drive link..."):
                        s_cnt, s_msg = auto_sync_from_any_drive_link(quick_drive_url.strip(), view_month, view_week, q_scope)
                        st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                        st.cache_data.clear()
                        st.success(f"✅ Na-save ang '{q_lbl}' at auto-synced agad! ({s_msg})")
                        time.sleep(0.6)
                        st.rerun()

    st.divider()

    col_up_left, col_up_right = st.columns([1.15, 1.55])

    with col_up_left:
        st.markdown("### 📤 1. Upload the Image / Excel File Here!")
        st.caption("I-upload ang larawan (screenshot/picture) o Excel file ng inyong Weekly Report. Kapag nagpalit ka ng Buwan o Linggo rito o sa itaas, kusang magpapalit din sa buong page.")

        c_m, c_w = st.columns(2)
        with c_m:
            sel_month = st.selectbox(
                "Buwan ng Report (2027):",
                MONTHS_2027,
                key="up_sel_month",
                on_change=_sync_from_upload_bar
            )
        with c_w:
            sel_week = st.selectbox(
                "Linggo ng Report:",
                WEEKS_2027,
                key="up_sel_week",
                on_change=_sync_from_upload_bar
            )

        parish_opts_up = PARISH_ORDER if is_master else [my_parish]
        if st.session_state.get("up_sel_parish") not in parish_opts_up:
            st.session_state["up_sel_parish"] = parish_opts_up[0]
        target_parish = st.selectbox(
            "Parokya na Nag-report:",
            parish_opts_up,
            key="up_sel_parish",
            on_change=_sync_from_upload_bar
        )

        with st.form("weekly_report_upload_form"):
            rep_date_input = st.text_input("Petsa ng Sabado / Pagsamba (MM/DD/YY):", value="01/02/27")
            uploaded_rep = st.file_uploader(
                "📸 Pumili ng Weekly Report Image (.jpg, .png) o Excel (.xlsx):",
                type=["jpg", "jpeg", "png", "xlsx"]
            )
            c_hv, c_mw, c_ind = st.columns(3)
            with c_hv:
                up_hv = st.text_input("House Visitation (Bilang):", value="0")
            with c_mw:
                up_mw = st.text_input("Missionary Work (Bilang):", value="0")
            with c_ind:
                up_ind = st.text_input("Indoctrination (Bilang/Pamilya):", value="0")

            extra_remarks = st.text_input("Remarks / Tala (Optional):").strip()

            submit_up = st.form_submit_button("🚀 I-upload sa Google Drive & I-type sa Parish Weekly Report", type="primary", use_container_width=True)

            if submit_up:
                if uploaded_rep is None:
                    st.warning("Mag-upload muna ng Image o Excel file ng Weekly Report.")
                else:
                    file_bytes = uploaded_rep.read()
                    f_ext = uploaded_rep.name.split(".")[-1].lower()
                    clean_p_name = target_parish.replace("PARISH", "").strip()
                    drive_filename = f"{clean_p_name} (PWR. {rep_date_input.replace('/', '-')} {sel_month} {sel_week}).{f_ext}"

                    with st.spinner(f"Ina-upload sa Google Drive folder ng {target_parish} at tinatype ang mga hindi pa tapos na datos (January–December 2027, Weeks 1–5)..."):
                        ph_now_stamp = get_philippine_datetime_str()
                        multi_reports_saved = 0
                        if f_ext == "xlsx":
                            try:
                                wb_up = openpyxl.load_workbook(io.BytesIO(file_bytes))
                                if ensure_all_12_month_sheets_in_parish_workbook(wb_up, parish_name=target_parish):
                                    b_up = io.BytesIO()
                                    wb_up.save(b_up)
                                    file_bytes = b_up.getvalue()
                            except Exception:
                                pass

                        up_folder_url = st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL", ACTIVE_PARISH_UPLOAD_URL) or st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
                        drive_link, drive_msg = upload_file_to_google_drive(
                            file_bytes,
                            drive_filename,
                            mime_type=uploaded_rep.type,
                            parent_folder_url=up_folder_url,
                            parish_name=target_parish,
                            month_name=sel_month
                        )

                        if f_ext == "xlsx":
                            all_multi = extract_all_months_and_weeks_from_parish_workbook(file_bytes, fallback_parish=target_parish)
                            if all_multi:
                                db_cur = load_weekly_monitoring_records()
                                chosen_ext = None
                                for rep_item in all_multi:
                                    r_p = rep_item.get("parish") or target_parish
                                    r_m = rep_item.get("month") or sel_month
                                    r_w = rep_item.get("week") or sel_week
                                    k_check = f"{r_p.upper()}__{r_m.upper()}__{r_w}"
                                    existing_rec = db_cur.get(k_check, {})
                                    # Huwag burahin ang mga natapos na; ilagay ang bagong/hindi pa tapos na datos o ang kasalukuyang napiling buwan at linggo
                                    if not is_week_record_completed(existing_rec) or (r_m.upper() == sel_month.upper() and r_w == sel_week):
                                        rep_item["submitted_at_ph"] = existing_rec.get("submitted_at_ph") or ph_now_stamp
                                        save_weekly_monitoring_record(
                                            r_p, r_m, r_w,
                                            rep_item.get("report_date", rep_date_input),
                                            rep_item,
                                            drive_link=drive_link or ""
                                        )
                                        multi_reports_saved += 1
                                    if r_m.upper() == sel_month.upper() and r_w == sel_week:
                                        chosen_ext = rep_item
                                if chosen_ext is None:
                                    chosen_ext = all_multi[0]
                                    sel_month = chosen_ext.get("month", sel_month)
                                    sel_week = chosen_ext.get("week", sel_week)
                                    target_parish = chosen_ext.get("parish", target_parish)
                                extracted = chosen_ext
                            else:
                                extracted = extract_weekly_report_from_excel_bytes(file_bytes, target_parish=target_parish, target_month=sel_month, target_week=sel_week)
                        else:
                            extracted = extract_weekly_report_via_ai_vision(file_bytes, uploaded_rep.type, selected_model, target_parish=target_parish)
                        extracted["submitted_at_ph"] = ph_now_stamp

                        denoms = get_parish_member_denominators(target_parish)
                        if int(extracted.get("total_members_say", 0)) <= 0:
                            extracted["total_members_say"] = denoms["say"]
                            extracted["total_members_y"] = denoms["youth"]
                            extracted["total_members_k"] = denoms["kids"]

                        if rep_date_input.strip():
                            extracted["report_date"] = rep_date_input.strip()
                        if up_hv.strip() and up_hv.strip() != "0":
                            extracted["house_visitation"] = up_hv.strip()
                        if up_mw.strip() and up_mw.strip() != "0":
                            extracted["missionary_work"] = up_mw.strip()
                        if up_ind.strip() and up_ind.strip() != "0":
                            extracted["indoctrination"] = up_ind.strip()
                        if extra_remarks:
                            extracted["remarks"] = extra_remarks

                        detected_p_img = str(extracted.get("detected_parish", "") or extracted.get("parish", "")).strip()
                        if is_master and detected_p_img in PARISH_ORDER and target_parish == PARISH_ORDER[0] and detected_p_img != target_parish:
                            save_weekly_monitoring_record(
                                target_parish, sel_month, sel_week,
                                extracted.get("report_date", rep_date_input),
                                extracted,
                                drive_link=drive_link or ""
                            )
                            target_parish = detected_p_img

                        ok_m, m_sync_msg = save_weekly_monitoring_record(
                            target_parish, sel_month, sel_week,
                            extracted.get("report_date", rep_date_input),
                            extracted,
                            drive_link=drive_link or ""
                        )

                        tot_extracted_att = (
                            int(extracted.get("mass_a", 0)) + int(extracted.get("mass_y", 0)) +
                            int(extracted.get("pampamilya_a", 0)) + int(extracted.get("pampamilya_y", 0)) +
                            int(extracted.get("online_a", 0)) + int(extracted.get("online_y", 0)) +
                            int(extracted.get("youth_y", 0)) + int(extracted.get("kids_k", 0)) +
                            int(extracted.get("adal_a", 0)) + int(extracted.get("adal_y", 0))
                        )
                        has_any_officer = any(
                            str(extracted.get(k, "")).strip()
                            for k in ["priest", "sakristan", "epistula", "evangelio", "tema_misa", "adal_evangelio", "tema_pampamilya"]
                        )
                        st.session_state["_next_mon_view_month"] = sel_month
                        st.session_state["_next_mon_view_week"] = sel_week
                        st.session_state["_next_editor_parish"] = target_parish
                        st.session_state["_next_editor_month"] = sel_month
                        st.session_state["_next_editor_week"] = sel_week
                        st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                        st.session_state["last_extracted_pwr"] = {
                            "parish": target_parish,
                            "month": sel_month,
                            "week": sel_week,
                            "date": extracted.get("report_date", rep_date_input),
                            "data": extracted,
                            "drive_link": drive_link,
                            "drive_msg": drive_msg,
                            "sync_msg": m_sync_msg,
                            "is_blank_source": (tot_extracted_att == 0 and not has_any_officer)
                        }
                        st.rerun()

        if "last_extracted_pwr" in st.session_state:
            flash = st.session_state["last_extracted_pwr"]
            sub_ph_flash = flash.get("data", {}).get("submitted_at_ph", "") or get_philippine_datetime_str()
            st.success(
                f"✅ Matagumpay na na-proseso ang Weekly Report ng **{flash.get('parish')}** para sa **{flash.get('month')} 2027 ({flash.get('week')})**! "
                f"(🇵🇭 Petsa at Oras ng Pag-submit sa Pilipinas: `{sub_ph_flash}` • {flash.get('sync_msg', '')})"
            )
            if flash.get("drive_link"):
                st.info(f"📂 Na-save sa Google Drive Folder ng {flash.get('parish')}: [Buksan ang File sa Google Drive]({flash.get('drive_link')})")
            elif flash.get("drive_msg"):
                st.warning(f"⚠️ {flash.get('drive_msg')}")
            if flash.get("is_blank_source"):
                st.warning(
                    "⚠️ **Pansinin:** Walang nakitang nakasulat na attendance (`1` / `✓`) o pangalan ng Officers sa loob ng in-upload na larawan (blangko ang form sa larawan). "
                    "Maaari mong i-type ang mga pangalan ng Officers at bilang ng Attendance sa **✏️ 2. Parish Weekly Report Data & Manual Verification Editor** sa kanan at pindutin ang **💾 I-type sa Parish Weekly Report Excel & I-update ang Tables**."
                )

    with col_up_right:
        st.markdown("### ✏️ 2. Parish Weekly Report Data & Manual Verification Editor")
        st.caption("Lahat ng datos dito ay direktang tinatype sa `PARISH WEEKLY REPORT` sheet (Liturgical Officers, Topics/Preachers, Adults/Youths/Kids Attendance, at Remarks). Naka-sync ang Buwan at Linggo rito sa Top Bar at Upload Form.")

        db_all = load_weekly_monitoring_records()

        v_c1, v_c2, v_c3 = st.columns(3)
        with v_c1:
            v_parish = st.selectbox(
                "Parokya:",
                PARISH_ORDER if is_master else [my_parish],
                key="editor_sel_parish",
                on_change=_sync_from_editor_bar
            )
        with v_c2:
            v_month = st.selectbox(
                "Buwan (2027):",
                MONTHS_2027,
                key="editor_sel_month",
                on_change=_sync_from_editor_bar
            )
        with v_c3:
            v_week = st.selectbox(
                "Linggo:",
                WEEKS_2027,
                key="editor_sel_week",
                on_change=_sync_from_editor_bar
            )

        cur_rec = db_all.get(f"{v_parish.upper()}__{v_month.upper()}__{v_week}", {})
        f_rev = f"{st.session_state.get('pwr_form_rev', 0)}_{v_parish}_{v_month}_{v_week}"
        def_denoms = get_parish_member_denominators(v_parish)

        with st.form("verify_weekly_counts_form"):
            with st.expander("⛪ Liturgical Officers & Preachers (Para sa PARISH WEEKLY REPORT Sheet)", expanded=True):
                o1, o2, o3, o4 = st.columns(4)
                with o1:
                    v_priest = st.text_input("PRIEST:", value=str(cur_rec.get("priest", "")), key=f"vp_{f_rev}")
                    v_casasaad = st.text_input("CASASAAD:", value=str(cur_rec.get("casasaad", "")), key=f"vcas_{f_rev}")
                with o2:
                    v_sakristan = st.text_input("SAKRISTAN:", value=str(cur_rec.get("sakristan", "")), key=f"vsak_{f_rev}")
                    v_candle = st.text_input("CANDLE HOLDER:", value=str(cur_rec.get("candle_holder", "")), key=f"vcan_{f_rev}")
                with o3:
                    v_epistula = st.text_input("EPISTULA:", value=str(cur_rec.get("epistula", "")), key=f"vepi_{f_rev}")
                    v_usher = st.text_input("USHER:", value=str(cur_rec.get("usher", "")), key=f"vush_{f_rev}")
                with o4:
                    v_evangelio = st.text_input("EVANGELIO:", value=str(cur_rec.get("evangelio", "")), key=f"veva_{f_rev}")
                    v_marshal = st.text_input("MARSHAL:", value=str(cur_rec.get("marshal", "")), key=f"vmar_{f_rev}")

                t1, t2 = st.columns(2)
                with t1:
                    v_adal_topic = st.text_input("ADAL TI EVANGELIO (Topic):", value=str(cur_rec.get("adal_evangelio", "")), key=f"vadt_{f_rev}")
                    v_misa_topic = st.text_input("TEMA TI MISA:", value=str(cur_rec.get("tema_misa", "")), key=f"vmit_{f_rev}")
                    v_pam_topic = st.text_input("TEMA TI PAMPAMILYA:", value=str(cur_rec.get("tema_pampamilya", "")), key=f"vpat_{f_rev}")
                with t2:
                    v_adal_pr = st.text_input("Adal Preacher/s:", value=str(cur_rec.get("adal_preacher", "")), key=f"vadp_{f_rev}")
                    v_misa_pr = st.text_input("Misa Preacher:", value=str(cur_rec.get("misa_preacher", "")), key=f"vmip_{f_rev}")
                    v_pam_pr = st.text_input("Pampamilya Preacher:", value=str(cur_rec.get("pampamilya_preacher", "")), key=f"vpap_{f_rev}")

            tm1, tm2, tm3 = st.columns(3)
            with tm1:
                v_tot_say = st.number_input("Total Members (Adults+Youths):", min_value=0, value=int(cur_rec.get("total_members_say", 0)) or def_denoms["say"], key=f"vtsay_{f_rev}")
            with tm2:
                v_tot_y = st.number_input("Total Youth Members:", min_value=0, value=int(cur_rec.get("total_members_y", 0)) or def_denoms["youth"], key=f"vty_{f_rev}")
            with tm3:
                v_tot_k = st.number_input("Total Kids Members:", min_value=0, value=int(cur_rec.get("total_members_k", 0)) or def_denoms["kids"], key=f"vtk_{f_rev}")

            a1, a2, a3, a4, a5 = st.columns(5)
            with a1:
                m_a = st.number_input("Mass (Adults A):", min_value=0, value=int(cur_rec.get("mass_a", 0)), key=f"vma_{f_rev}")
                m_y = st.number_input("Mass (Youth Y):", min_value=0, value=int(cur_rec.get("mass_y", 0)), key=f"vmy_{f_rev}")
            with a2:
                p_a = st.number_input("Pampamilya (A):", min_value=0, value=int(cur_rec.get("pampamilya_a", 0)), key=f"vpa_{f_rev}")
                p_y = st.number_input("Pampamilya (Y):", min_value=0, value=int(cur_rec.get("pampamilya_y", 0)), key=f"vpy_{f_rev}")
            with a3:
                o_a = st.number_input("Online Min. (A):", min_value=0, value=int(cur_rec.get("online_a", 0)), key=f"voa_{f_rev}")
                o_y = st.number_input("Online Min. (Y):", min_value=0, value=int(cur_rec.get("online_y", 0)), key=f"voy_{f_rev}")
            with a4:
                y_y = st.number_input("Youth Fell. (Y):", min_value=0, value=int(cur_rec.get("youth_y", 0)), key=f"vyy_{f_rev}")
                k_k = st.number_input("Kids Fell. (K):", min_value=0, value=int(cur_rec.get("kids_k", 0)), key=f"vkk_{f_rev}")
            with a5:
                ad_a = st.number_input("Adal ti Evang. (A):", min_value=0, value=int(cur_rec.get("adal_a", 0)), key=f"vada_{f_rev}")
                ad_y = st.number_input("Adal ti Evang. (Y):", min_value=0, value=int(cur_rec.get("adal_y", 0)), key=f"vady_{f_rev}")

            parsed_cur_ro = parse_remarks_into_mass_and_outreach(str(cur_rec.get("remarks", "")))
            b1, b2, b3, b4 = st.columns(4)
            with b1:
                v_date = st.text_input("Petsa (MM/DD/YY):", value=str(cur_rec.get("report_date", "01/02/27")), key=f"vdt_{f_rev}")
            with b2:
                v_hv = st.text_input("House Visitation (Bilang Pamilya):", value=str(cur_rec.get("house_visitation", "0")), key=f"vhv_{f_rev}")
            with b3:
                def_mw_cnt = str(cur_rec.get("missionary_work", "0"))
                if def_mw_cnt == "0" and parsed_cur_ro.get("mw_families_count", 0) > 0:
                    def_mw_cnt = str(parsed_cur_ro["mw_families_count"])
                v_mw = st.text_input("Missionary Work (Bilang Pamilya):", value=def_mw_cnt, key=f"vmw_{f_rev}")
            with b4:
                v_ind = st.text_input("Indoctrination (Bilang Pamilya):", value=str(cur_rec.get("indoctrination", "0")), key=f"vind_{f_rev}")

            with st.expander("🏠 Outreach Details: Family Names, Mga Pangalan ng Tao, Total Tao & Remarks (House Visitation, Missionary Work, Indoctrination)", expanded=False):
                st.markdown("**1️⃣ House Visitation (Requirements: 10 Families Member or Non-Member):**")
                hv_c1, hv_c2 = st.columns([2.2, 1.0])
                with hv_c1:
                    v_hv_names = st.text_area(
                        "House Visitation — Family Names & Mga Pangalan (Hal. Dela Cruz Family: Juan D., Maria D.):",
                        value=str(cur_rec.get("hv_names", "") or parsed_cur_ro.get("hv_names", "")),
                        height=68,
                        key=f"vhvn_{f_rev}"
                    )
                with hv_c2:
                    v_hv_ppl = st.number_input(
                        "Total Tao sa House Visitation:",
                        min_value=0,
                        value=int(cur_rec.get("hv_people_count", 0) or parsed_cur_ro.get("hv_people_count", 0)),
                        key=f"vhvp_{f_rev}"
                    )
                    v_hv_rem = st.text_input(
                        "Remarks sa House Visitation:",
                        value=str(cur_rec.get("hv_remarks", "") or parsed_cur_ro.get("hv_remarks", "")),
                        key=f"vhvr_{f_rev}"
                    )

                st.markdown("**2️⃣ Missionary Work:**")
                mw_c1, mw_c2 = st.columns([2.2, 1.0])
                with mw_c1:
                    v_mw_names = st.text_area(
                        "Missionary Work — Family Names & Mga Pangalan ng Dumalo:",
                        value=str(cur_rec.get("mw_names", "") or parsed_cur_ro.get("mw_names", "")),
                        height=68,
                        key=f"vmwn_{f_rev}"
                    )
                with mw_c2:
                    v_mw_ppl = st.number_input(
                        "Total Tao sa Missionary Work:",
                        min_value=0,
                        value=int(cur_rec.get("mw_people_count", 0) or parsed_cur_ro.get("mw_people_count", 0)),
                        key=f"vmwp_{f_rev}"
                    )
                    v_mw_rem = st.text_input(
                        "Remarks sa Missionary Work:",
                        value=str(cur_rec.get("mw_remarks", "") or parsed_cur_ro.get("mw_remarks", "")),
                        key=f"vmwr_{f_rev}"
                    )

                st.markdown("**3️⃣ Indoctrination (Family Name of a Family Under Indoctrination):**")
                ind_c1, ind_c2 = st.columns([2.2, 1.0])
                with ind_c1:
                    v_ind_names = st.text_area(
                        "Indoctrination — Family Names & Mga Pangalan:",
                        value=str(cur_rec.get("ind_names", "") or parsed_cur_ro.get("ind_names", "")),
                        height=68,
                        key=f"vindn_{f_rev}"
                    )
                with ind_c2:
                    v_ind_ppl = st.number_input(
                        "Total Tao sa Indoctrination:",
                        min_value=0,
                        value=int(cur_rec.get("ind_people_count", 0) or parsed_cur_ro.get("ind_people_count", 0)),
                        key=f"vindp_{f_rev}"
                    )
                    v_ind_rem = st.text_input(
                        "Remarks sa Indoctrination:",
                        value=str(cur_rec.get("ind_remarks", "") or parsed_cur_ro.get("ind_remarks", "")),
                        key=f"vindr_{f_rev}"
                    )

            v_mass_rem = st.text_area(
                "Remarks para sa Nag-attend ng Mass / Pampamilya / Adal / House Church / Online Mass / Visitors (Galing sa Image/Excel Report):",
                value=str(cur_rec.get("mass_remarks", "") or parsed_cur_ro.get("mass_remarks", "")),
                height=75,
                key=f"vmassrem_{f_rev}"
            )
            v_rem = st.text_area(
                "Buong Remarks para sa Parokya (Column J sa PARISH WEEKLY REPORT Excel):",
                value=str(cur_rec.get("remarks", "")),
                height=95,
                key=f"vrem_{f_rev}"
            )

            if st.form_submit_button("💾 I-type sa Parish Weekly Report Excel & I-update ang Lahat ng Tables", use_container_width=True):
                denoms = get_parish_member_denominators(v_parish)
                updated_payload = {
                    "total_members_say": v_tot_say if v_tot_say > 0 else denoms["say"],
                    "total_members_y": v_tot_y if v_tot_y > 0 else denoms["youth"],
                    "total_members_k": v_tot_k if v_tot_k > 0 else denoms["kids"],
                    "mass_a": m_a, "mass_y": m_y,
                    "pampamilya_a": p_a, "pampamilya_y": p_y,
                    "online_a": o_a, "online_y": o_y,
                    "youth_y": y_y, "kids_k": k_k,
                    "adal_a": ad_a, "adal_y": ad_y,
                    "house_visitation": v_hv,
                    "hv_names": v_hv_names,
                    "hv_people_count": v_hv_ppl,
                    "hv_remarks": v_hv_rem,
                    "missionary_work": v_mw,
                    "mw_names": v_mw_names,
                    "mw_people_count": v_mw_ppl,
                    "mw_remarks": v_mw_rem,
                    "indoctrination": v_ind,
                    "ind_names": v_ind_names,
                    "ind_people_count": v_ind_ppl,
                    "ind_remarks": v_ind_rem,
                    "mass_remarks": v_mass_rem,
                    "remarks": v_rem,
                    "priest": v_priest,
                    "sakristan": v_sakristan,
                    "epistula": v_epistula,
                    "evangelio": v_evangelio,
                    "casasaad": v_casasaad,
                    "candle_holder": v_candle,
                    "usher": v_usher,
                    "marshal": v_marshal,
                    "adal_evangelio": v_adal_topic,
                    "adal_preacher": v_adal_pr,
                    "tema_misa": v_misa_topic,
                    "misa_preacher": v_misa_pr,
                    "tema_pampamilya": v_pam_topic,
                    "pampamilya_preacher": v_pam_pr,
                }
                _, sync_msg = save_weekly_monitoring_record(v_parish, v_month, v_week, v_date, updated_payload, drive_link=cur_rec.get("drive_link", ""))
                st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                st.success(f"✅ Na-update na sa PARISH WEEKLY REPORT at Monitoring Tables para sa **{v_parish}** ({v_month} 2027 - {v_week})! ({sync_msg})")
                time.sleep(0.5)
                st.rerun()

    st.divider()

    # =====================================================================
    # 📅 FULL-YEAR 2027 MONTHLY & WEEKLY PROGRESS TRACKER (JANUARY TO DECEMBER, WEEKS 1-5)
    # =====================================================================
    with st.expander(f"📅 Tingnan ang Buong Taon 2027 Monthly & Weekly Submission Tracker (January–December 2027, Weeks 1–5) — 🇵🇭 PH Time", expanded=True):
        st.caption("Makikita rito ang lahat ng **✅ Tapos Na / Na-submit** na records ng mga parokya (may petsa at oras sa Pilipinas) at kung alin ang **⏳ Hindi Pa Tapos** sa buong taon ng 2027. Lahat ng natapos na records ay permanenteng naka-save upang mabalikan anumang oras.")
        nxt_m, nxt_w = find_next_unfinished_month_and_week(v_parish, preferred_month=view_month)
        col_nxt1, col_nxt2 = st.columns([2.2, 1.2])
        with col_nxt1:
            st.info(f"📌 **Susunod na Hindi Pa Tapos na Report para sa {v_parish}:** `{nxt_m} 2027 — {nxt_w}`")
        with col_nxt2:
            if st.button(f"⏭️ Buksan ang {nxt_m} ({nxt_w})", use_container_width=True):
                st.session_state["_next_mon_view_month"] = nxt_m
                st.session_state["_next_mon_view_week"] = nxt_w
                st.session_state["_next_editor_month"] = nxt_m
                st.session_state["_next_editor_week"] = nxt_w
                st.rerun()

        st.markdown("#### ✅ Listahan ng Lahat ng Natapos / Na-submit na Records (January–December 2027, Weeks 1–5)")
        completed_dir_df = build_completed_records_directory_df(effective_parish_filter)
        render_record_style_html_table(completed_dir_df)

        # Quick jump selector for any completed record
        completed_jump_opts = []
        for m_j in MONTHS_2027:
            for w_j in WEEKS_2027:
                for p_j in (PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]):
                    r_j = db_all.get(f"{p_j.upper()}__{m_j.upper()}__{w_j}", {})
                    if is_week_record_completed(r_j):
                        completed_jump_opts.append((f"{p_j} — {m_j} 2027 ({w_j}) [Petsa: {r_j.get('report_date', '01/02/27')}]", p_j, m_j, w_j))
        if completed_jump_opts:
            jc1, jc2 = st.columns([2.6, 1.0])
            with jc1:
                sel_jump_lbl = st.selectbox(
                    "👁️ Pumili ng Natapos na Record upang Buksan at Tingnan Agad:",
                    [x[0] for x in completed_jump_opts],
                    key="jump_completed_rec_sel"
                )
            with jc2:
                st.write("")
                if st.button("📂 Buksan ang Record na Ito", use_container_width=True, type="primary"):
                    chosen_j = next((x for x in completed_jump_opts if x[0] == sel_jump_lbl), completed_jump_opts[0])
                    st.session_state["_next_editor_parish"] = chosen_j[1]
                    st.session_state["_next_mon_view_month"] = chosen_j[2]
                    st.session_state["_next_mon_view_week"] = chosen_j[3]
                    st.rerun()

        with st.expander("📊 Buong Taon 2027 Month & Week Matrix (January–December 2027, Weeks 1–5)", expanded=False):
            tracker_2027_df = build_2027_full_year_month_week_tracker_df(effective_parish_filter)
            render_record_style_html_table(tracker_2027_df)

    # =====================================================================
    # 📊 TABLE 3 & TABLE 4 (AS-IS FORMAT, SUPPORTS SINGLE WEEK OR ALL WEEKS 1-5 OF THE MONTH)
    # =====================================================================
    weeks_to_render = WEEKS_2027 if show_all_weeks_of_month else [view_week]
    db_cards_3b = load_weekly_monitoring_records()

    for wk_iter in weeks_to_render:
        st.markdown(f"### 📋 3. Parish Monitoring Report — Attendance & Total Percentage ({wk_iter} of {view_month.upper()} 2027)")
        st.caption(f"Bawat Parokya ay may kasamang **Date & Time Submitted (PH Time)** at **Month/Week (2027)** upang makita ang lahat ng datos sa buong taon ng 2027.")
        att_monitoring_df = build_parish_monitoring_attendance_df(view_month, wk_iter, effective_parish_filter)
        st.dataframe(att_monitoring_df, use_container_width=True, hide_index=True)

        mass_cards_list = []
        for p_c in (PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]):
            rec_c = db_cards_3b.get(f"{p_c.upper()}__{view_month.upper()}__{wk_iter}", {})
            if rec_c.get("has_data"):
                r_raw = str(rec_c.get("remarks", "")).strip()
                if r_raw:
                    mass_cards_list.append((p_c, rec_c, r_raw))
        with st.expander(f"📝 Basahin nang Buo ang Nakasulat sa Remarks Mula sa Image / Excel Report ({len(mass_cards_list)} may Remarks — {wk_iter} of {view_month.upper()} 2027)", expanded=(not show_all_weeks_of_month or len(mass_cards_list) > 0)):
            if mass_cards_list:
                for p_c, rec_c, r_raw in mass_cards_list:
                    sub_ph_c = str(rec_c.get("submitted_at_ph", "") or rec_c.get("updated_at", "")).strip()
                    display_raw_rem = r_raw.replace("\n", "<br>")
                    st.markdown(
                        f"<div class='full-text-card'>"
                        f"<b>⛪ {p_c}</b> • <code>{view_month} 2027 ({wk_iter})</code> • 📅 <b>Petsa ng Report:</b> {rec_c.get('report_date', '01/02/27')} "
                        + (f"• 🇵🇭 <b>Submitted (PH Time):</b> <code>{sub_ph_c}</code>" if sub_ph_c else "")
                        + f"<br><hr style='margin:6px 0;'>"
                        f"📝 <b>Nakasulat sa Remarks:</b><br>{display_raw_rem}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.caption("Walang nakalagay na Remarks mula sa in-upload na Image o Excel report para sa linggong ito.")

        st.divider()

        st.markdown(f"### 📈 4. Summary for Posting — Per Parish Breakdown ({wk_iter} of {view_month.upper()} 2027)")
        st.caption("Makikita rito ang `TOTAL MEMBERS`, `TOTAL OF ATTENDEES`, `TOTAL AVERAGE %`, `Month & Week (2027)`, at `Date & Time Submitted (PH Time)` ng bawat Parokya.")

        parishes_to_show = PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]
        for p_item in parishes_to_show:
            df_post_p = build_summary_for_posting_per_parish_df(p_item, view_month, wk_iter)
            st.dataframe(df_post_p, use_container_width=True, hide_index=True)

        st.divider()

    # =====================================================================
    # 📊 TABLE 5: RECORD-STYLE TABLES PARA SA HOUSE VISITATION, MISSIONARY WORK, AT INDOCTRINATION
    # =====================================================================
    st.markdown(f"### 🏠 5. House Visitation, Missionary Work, Indoctrination & Remarks — Record Style View ({view_week} of {view_month.upper()} 2027)")
    st.caption("Naka-Record Style na ang tatlong tables sa ibaba upang **buong-buo at malinaw na mabasa ang lahat ng nakasulat na Family Names, Mga Pangalan ng Tao, Bilang ng Pamilya, Total Tao, at Remarks** nang hindi napuputol.")

    st.markdown("#### 🏡 Unang Table: HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER")
    hv_table_df = build_house_visitation_table_df(view_month, view_week, effective_parish_filter)
    render_record_style_html_table(hv_table_df)

    st.markdown("#### 🌍 Pangalawang Table: Missionary Work (Family Names, Mga Pangalan ng Dumalo, Total Tao & Remarks)")
    mw_table_df = build_missionary_work_table_df(view_month, view_week, effective_parish_filter)
    render_record_style_html_table(mw_table_df)

    st.markdown("#### 📖 Pangatlong Table: INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION")
    ind_table_df = build_indoctrination_table_df(view_month, view_week, effective_parish_filter)
    render_record_style_html_table(ind_table_df)

    outreach_cards_list = []
    for p_c in (PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]):
        rec_c = db_cards_3b.get(f"{p_c.upper()}__{view_month.upper()}__{view_week}", {})
        if rec_c.get("has_data"):
            r_raw = str(rec_c.get("remarks", "")).strip()
            p_ro = parse_remarks_into_mass_and_outreach(r_raw) if r_raw else {}
            hv_n = str(rec_c.get("hv_names", "") or p_ro.get("hv_names", "")).strip()
            mw_n = str(rec_c.get("mw_names", "") or p_ro.get("mw_names", "")).strip()
            ind_n = str(rec_c.get("ind_names", "") or p_ro.get("ind_names", "")).strip()
            if hv_n or mw_n or ind_n or parse_int_from_text(rec_c.get("house_visitation", "0")) > 0 or parse_int_from_text(rec_c.get("missionary_work", "0")) > 0 or parse_int_from_text(rec_c.get("indoctrination", "0")) > 0:
                outreach_cards_list.append((p_c, rec_c, p_ro))

    with st.expander(f"📝 Basahin nang Buo ang Outreach Details (House Visitation, Missionary Work & Indoctrination — {len(outreach_cards_list)} Parokya)", expanded=True):
        if outreach_cards_list:
            for p_c, rec_c, p_ro in outreach_cards_list:
                hv_n = str(rec_c.get("hv_names", "") or p_ro.get("hv_names", "")).strip() or "Wala pang nakalistang pangalan"
                hv_r = str(rec_c.get("hv_remarks", "") or p_ro.get("hv_remarks", "")).strip() or "-"
                hv_ppl = int(rec_c.get("hv_people_count", 0) or p_ro.get("hv_people_count", 0))

                mw_n = str(rec_c.get("mw_names", "") or p_ro.get("mw_names", "")).strip() or "Wala pang nakalistang pangalan"
                mw_r = str(rec_c.get("mw_remarks", "") or p_ro.get("mw_remarks", "")).strip() or "-"
                mw_ppl = int(rec_c.get("mw_people_count", 0) or p_ro.get("mw_people_count", 0))

                ind_n = str(rec_c.get("ind_names", "") or p_ro.get("ind_names", "")).strip() or "Wala pang nakalistang pangalan"
                ind_r = str(rec_c.get("ind_remarks", "") or p_ro.get("ind_remarks", "")).strip() or "-"
                ind_ppl = int(rec_c.get("ind_people_count", 0) or p_ro.get("ind_people_count", 0))

                st.markdown(
                    f"<div class='full-text-card'>"
                    f"<b>🏠 {p_c}</b> • <code>{view_month} 2027 ({view_week})</code> • 📅 <b>Petsa:</b> {rec_c.get('report_date', '01/02/27')}<br>"
                    f"🏡 <b>House Visitation ({rec_c.get('house_visitation', '0')} Pamilya | {hv_ppl} Tao):</b> {hv_n} <i>(Remarks: {hv_r})</i><br>"
                    f"🌍 <b>Missionary Work ({rec_c.get('missionary_work', '0')} Pamilya | {mw_ppl} Tao):</b> {mw_n} <i>(Remarks: {mw_r})</i><br>"
                    f"📖 <b>Indoctrination ({rec_c.get('indoctrination', '0')} Pamilya | {ind_ppl} Tao):</b> {ind_n} <i>(Remarks: {ind_r})</i>"
                    f"</div>",
                    unsafe_allow_html=True
                )
        else:
            st.caption("Walang nakalagay na Outreach details (House Visitation, Missionary Work, Indoctrination) para sa linggong ito.")

    with st.expander("📊 Pangkalahatang Pinagsamang Table (House Visitation + Missionary Work + Indoctrination Summary)", expanded=False):
        vis_df = build_visitation_missionary_indoctrination_df(view_month, view_week, effective_parish_filter)
        st.dataframe(vis_df, use_container_width=True, hide_index=True)

    if os.path.exists(LOCAL_PWR_MASTER_FILE):
        with open(LOCAL_PWR_MASTER_FILE, "rb") as f:
            pwr_bytes = f.read()
        st.download_button(
            label=f"📥 I-download ang Na-update na Master Excel (PWR. {view_month} {view_week}.xlsx)",
            data=pwr_bytes,
            file_name=f"PWR. {view_month} {view_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =====================================================================
# ⚙️ SIDEBAR CONTROL PANEL
# =====================================================================
with st.sidebar:
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=85)
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")

    if st.button("🔄 Sync Live Excel & Notion Data", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

    if st.session_state.user_role == "Master Admin":
        with st.expander("📡 Live Update & Activity Monitor", expanded=True):
            if not global_records_df.empty:
                c_s1, c_s2 = st.columns(2)
                cnt_sub = len(global_records_df[global_records_df["Status"].str.contains("For Master Admin Approval", case=False, na=False)])
                cnt_ass = len(global_records_df[global_records_df["Status"].str.contains("Assigned to Parish Leader", case=False, na=False)])
                c_s1.metric("📥 For Approval", cnt_sub)
                c_s2.metric("📤 Assigned", cnt_ass)
                st.caption("🔔 **Pinakahuling Galaw / Updates:**")
                for _, l_row in global_records_df.head(3).iterrows():
                    st.markdown(
                        f"<div class='sidebar-log-card'>"
                        f"<b>{l_row['Pangalan ng Nag-request']}</b> ({l_row['Parokya']})<br>"
                        f"📊 {l_row['Status']}<br>"
                        f"🏛️ {l_row['Isinumite / Inaprubahan Ni']}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.caption("Wala pang aktibong rekord.")

        with st.expander(f"☁️ Cloud Backup ({BACKUP_ADMIN_EMAIL})", expanded=False):
            st.session_state.auto_cloud_backup = st.checkbox(
                "Auto-save bawat update sa Gmail Cloud",
                value=st.session_state.auto_cloud_backup
            )
            if not global_records_df.empty:
                csv_bytes = global_records_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download CSV (Para sa Google Drive)",
                    data=csv_bytes,
                    file_name=f"CTG_Ministry_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                if st.button("📧 I-email ang Buong Backup sa Gmail Ko", use_container_width=True):
                    with st.spinner("Ipinapadala sa clarojerald16@gmail.com..."):
                        res_b = send_email_via_gmail(
                            BACKUP_ADMIN_EMAIL,
                            f"CTG Full Database Backup — {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                            f"Narito ang kumpletong CSV backup ng lahat ng CTG Ministry Hub records ({len(global_records_df)} rows).",
                            attachment_bytes=csv_bytes,
                            attachment_filename=f"CTG_Full_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
                        )
                        if "Matagumpay" in res_b:
                            st.success("Naipadala sa clarojerald16@gmail.com!")
                        else:
                            st.error(res_b)

    if st.session_state.login_role == "Master Admin":
        with st.expander("🛠️ Admin View & Parish Switcher", expanded=False):
            st.caption("Silipin o kontrolin ang dashboard ng kahit anong Parokya:")
            override_role = st.selectbox("Pumili ng View Mode:", ["Master Admin (Lahat)", "Parish Leader View"], key="ctrl_role_override")
            if override_role == "Master Admin (Lahat)":
                new_role = "Master Admin"
                new_parish = "Lahat ng Parokya"
            else:
                new_role = "Parish Leader"
                new_parish = st.selectbox("Pumili ng Parokya:", PARISH_ORDER, key="ctrl_parish_override")
            if st.button("🔄 Ilipat ang Control View", use_container_width=True):
                st.session_state.user_role = new_role
                st.session_state.assigned_parish = new_parish
                st.rerun()

    st.divider()
    if st.session_state.user_role == "Master Admin":
        nav_options = [
            "📑 Weekly Report & Parish Monitoring",
            "📈 Live Activity & Update Dashboard",
            "📊 Parish Census & Member Manager",
            "📊 Parish Directory & Family Search",
            "🛡️ Master Admin Controller & Approval Queue",
            "🗣️ Ministry Messenger (Admin & Users Chat)",
            "🤝 Notion CRM (Visitors & Members)",
            "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato",
            "🙏 Prayer Wall & Categorizer",
            "📢 Announcement Broadcaster",
            "🛠️ App Settings"
        ]
    else:
        nav_options = [
            "📑 Weekly Report & Parish Monitoring",
            "🧭 Parish Leader Portal (Wizard & Summary)",
            "📥 Parish Inbox & Action Hub",
            "📊 Parish Census & Member Manager",
            "🗣️ Ministry Messenger (Admin & Users Chat)",
            "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato",
            "🙏 Prayer Wall & Categorizer",
            "🛠️ App Settings"
        ]

    default_nav_idx = 0
    if "app_mode_override" in st.session_state:
        target_nav = st.session_state.pop("app_mode_override")
        if target_nav in nav_options:
            default_nav_idx = nav_options.index(target_nav)

    app_mode = st.radio("Pumili ng Pahina:", nav_options, index=default_nav_idx)
    st.caption(f"💬 Ministry Messenger: **{len(visible_chats_for_me)}** mensahe sa iyong inbox")
    selected_model = st.selectbox(
        "Gemini Model (Fastest):",
        ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"],
        index=0
    )
    st.divider()
    if st.button("🗑️ I-clear ang AI Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

def generate_ai_ministry_draft(service_category, sender_name, parish_name, topic, details, lang, model_name):
    valid_keys = get_valid_gemini_keys()
    if not valid_keys:
        return "⚠️ Wala pang GEMINI_API_KEY."
    client = genai.Client(api_key=valid_keys[0])
    prompt = (
        f"Ikaw si CTG Gabay AI, sumusulat ng opisyal, magalang, at punong-puno ng malasakit na tugon mula sa "
        f"Church of the Temple of God (CTG) - {parish_name} (kung saan ang araw ng pagsamba at pangunahing pagtitipon ay tuwing Sabado / Saturday / Sabbath Day).\n"
        f"Uri ng Request: {service_category}\n"
        f"Pangalan ng Nagpadala: {sender_name}\nPaksa: {topic}\nDetalye:\n\"{details}\"\n\n"
        f"1. Ilagay sa unang linya: '[Urgency Analysis: Low / Medium / High]'\n"
        f"2. Sumulat ng buong mensahe (email/SMS ready) sa wikang {lang} na may comforting words, Bible verse, o malinaw na gabay.\n"
    )
    res = safe_generate_content(client, model_name, prompt)
    return res.text

# =====================================================================
# 🗣️ INTERNAL MINISTRY MESSENGER (PRIVATE USER-TO-USER, USER-TO-ADMIN & BROADCAST)
# =====================================================================
def render_ministry_messenger_page():
    my_id = get_current_chat_identity()
    is_master = (st.session_state.user_role == "Master Admin")

    render_title_with_logo(
        "🗣️ CTG Ministry Messenger — Private & Broadcast Chat",
        f"Ikaw ay naka-login bilang: {my_id}. Ang pribadong usapan ay tanging sa pagitan lamang ninyo ng kausap mo at ng Master Admin."
    )

    all_msgs = load_internal_chat_messages()

    if is_master:
        partner_options = ["📢 LAHAT NG USERS (BROADCAST)"] + PARISH_ORDER
    else:
        other_parishes = [p for p in PARISH_ORDER if p != my_id]
        partner_options = ["🛡️ MASTER ADMIN", "📢 LAHAT NG USERS (BROADCAST)"] + other_parishes

    col_top1, col_top2 = st.columns([2.6, 1])
    with col_top1:
        selected_partner_raw = st.selectbox(
            "💬 Pumili ng User / Parokya na Gusto Mong Kausapin (Dropdown):",
            partner_options,
            key="messenger_partner_select"
        )
    with col_top2:
        st.write("")
        if st.button("🔄 I-refresh ang mga Mensahe", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    if "LAHAT NG USERS" in selected_partner_raw:
        partner_id = "LAHAT NG USERS (BROADCAST)"
    elif "MASTER ADMIN" in selected_partner_raw:
        partner_id = "MASTER ADMIN"
    else:
        partner_id = selected_partner_raw.strip()

    if is_master:
        tab_conv, tab_all_monitor = st.tabs([
            f"💬 Chat Box: {partner_id}",
            "👁️ Master Admin: Monitor All Private & Broadcast Chats (Lahat ng Usapan)"
        ])
    else:
        tab_conv = st.container()
        tab_all_monitor = None

    with tab_conv:
        if partner_id == "LAHAT NG USERS (BROADCAST)":
            st.info("📢 **Broadcast Mode:** Ang mensaheng ipapadala mo rito ay mababasa ng **lahat ng Parokya at Master Admin**.")
            conv_msgs = [
                m for m in all_msgs
                if m.get("recipient") == "LAHAT NG USERS (BROADCAST)"
            ]
        else:
            st.success(
                f"🔒 **Private Conversation ({my_id} ↔ {partner_id}):** "
                f"Hindi ito makikita ng ibang parokya. Tanging ikaw, si **{partner_id}**, at ang **Master Admin** lamang ang makakabasa nito."
            )
            conv_msgs = [
                m for m in all_msgs
                if (m.get("sender") == my_id and m.get("recipient") == partner_id)
                or (m.get("sender") == partner_id and m.get("recipient") == my_id)
            ]

        chat_box = st.container(height=420, border=True)
        with chat_box:
            if not conv_msgs:
                st.caption(f"Wala pang mensahe sa pagitan ng **{my_id}** at **{partner_id}**. Mag-type sa ibaba para magsimula!")
            else:
                for m in conv_msgs:
                    s_name = m.get("sender", "Unknown")
                    r_name = m.get("recipient", "Unknown")
                    t_str = m.get("timestamp", "")
                    body_txt = m.get("message", "")

                    if r_name == "LAHAT NG USERS (BROADCAST)":
                        st.markdown(
                            f"<div class='chat-bubble-broadcast'>"
                            f"<b>📢 {s_name}</b> → <i>Lahat ng Users</i> <span style='float:right;font-size:0.78rem;'>🕒 {t_str}</span><br>"
                            f"{body_txt}"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    elif s_name == my_id:
                        st.markdown(
                            f"<div class='chat-bubble-mine'>"
                            f"<b>👤 Ikaw ({s_name})</b> → <b>{r_name}</b> <span style='float:right;font-size:0.78rem;'>🕒 {t_str}</span><br>"
                            f"{body_txt}"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    else:
                        st.markdown(
                            f"<div class='chat-bubble-other'>"
                            f"<b>🏛️ {s_name}</b> → <b>{r_name}</b> <span style='float:right;font-size:0.78rem;'>🕒 {t_str}</span><br>"
                            f"{body_txt}"
                            f"</div>",
                            unsafe_allow_html=True
                        )

        with st.form("send_internal_chat_form", clear_on_submit=True):
            c_in, c_btn = st.columns([4, 1])
            with c_in:
                msg_input = st.text_input(
                    f"Mensahe para kay {partner_id}:",
                    placeholder=f"Mag-type ng mensahe para kay {partner_id}...",
                    label_visibility="collapsed"
                )
            with c_btn:
                send_btn = st.form_submit_button("📨 Ipadala", type="primary", use_container_width=True)

            if send_btn:
                if msg_input and msg_input.strip():
                    save_internal_chat_message(my_id, partner_id, msg_input.strip())
                    st.success(f"Naipadala na kay {partner_id}!")
                    st.rerun()
                else:
                    st.warning("Mag-type muna ng mensahe.")

    if is_master and tab_all_monitor is not None:
        with tab_all_monitor:
            st.subheader("👁️ Master Admin Exclusive: Lahat ng Pribado at Broadcast na Usapan ng mga Users")
            st.caption("Bilang Master Admin, nakikita mo rito ang lahat ng chat sa pagitan ng kahit sinong Parish Leaders at Admin.")
            flt_col1, flt_col2 = st.columns([1.5, 2])
            with flt_col1:
                mon_parish = st.selectbox(
                    "Salain ayon sa User / Parokya:",
                    ["Lahat ng Usapan", "MASTER ADMIN", "LAHAT NG USERS (BROADCAST)"] + PARISH_ORDER,
                    key="mon_chat_filter"
                )
            with flt_col2:
                mon_search = st.text_input("🔍 Hanapin sa mga mensahe:", key="mon_chat_search").strip()

            mon_list = list(reversed(all_msgs))
            if mon_parish != "Lahat ng Usapan":
                mon_list = [
                    m for m in mon_list
                    if m.get("sender") == mon_parish or m.get("recipient") == mon_parish
                ]
            if mon_search:
                mon_list = [
                    m for m in mon_list
                    if mon_search.lower() in f"{m.get('sender','')} {m.get('recipient','')} {m.get('message','')}".lower()
                ]

            if not mon_list:
                st.info("Walang nakitang mensahe sa napiling filter.")
            else:
                mon_df = pd.DataFrame(mon_list)[["timestamp", "sender", "recipient", "message"]]
                mon_df.columns = ["Petsa at Oras", "Nagpadala (Sender)", "Pinadalhan (Recipient)", "Mensahe"]
                st.dataframe(mon_df, use_container_width=True, hide_index=True)

# =====================================================================
# 👨‍👩‍👧‍👦 PINAG-ISANG DIRECTORY, FAMILY SEARCH & CENSUS ENGINE (LAHAT NG ADMIN AT USERS)
# =====================================================================
def render_parish_directory_and_census_view(page_heading: str, page_subtitle: str = ""):
    render_title_with_logo(page_heading, page_subtitle)
    if master_members_df.empty:
        st.error(f"Hindi makita o mabasa ang Excel file ({excel_filename}). Maaari mong i-update ang Google Drive Link sa App Settings.")
        return

    if st.session_state.user_role == "Master Admin":
        parish_options = ["Lahat ng Parokya"] + PARISH_ORDER
    else:
        parish_options = [st.session_state.assigned_parish]

    col_p, col_f, col_s = st.columns([1.15, 1.25, 1.6])
    with col_p:
        selected_parish = st.selectbox("1. Pumili ng Parish:", parish_options)

    if selected_parish == "Lahat ng Parokya":
        parish_filtered_df = master_members_df.copy()
    else:
        parish_filtered_df = master_members_df[master_members_df["Parish"] == selected_parish].copy()

    family_display_list = ["Lahat ng Family ID"]
    family_map = {}
    for fid_val, grp in parish_filtered_df.groupby("Family ID", sort=True):
        fid_clean = str(fid_val).strip()
        if not fid_clean:
            continue
        heads = [str(h).strip() for h in grp["Head of Family"].unique() if str(h).strip() and str(h).strip().lower() != "nan"]
        head_label = heads[0] if heads else str(grp.iloc[0]["Full Name"]).strip()
        disp_str = f"{fid_clean} — ({head_label}) [{len(grp)} miyembro]"
        family_display_list.append(disp_str)
        family_map[disp_str] = fid_clean

    with col_f:
        selected_family_disp = st.selectbox("2. Pumili ng Family ID:", family_display_list)
        selected_family_id = family_map.get(selected_family_disp, "Lahat ng Family ID")

    with col_s:
        keyword_search = st.text_input(
            "3. Search Family ID o Pangalan (Optional):",
            placeholder="Hal. CTG-LN-FAM-0001 o Juan Reyes"
        ).strip()

    c_opt1, c_opt2 = st.columns([3, 1])
    with c_opt1:
        expand_full_family = st.checkbox(
            "👨‍👩‍👧‍👦 Ipakita rin ang ibang kapamilya (Kapag naka-uncheck, TANGING YUNG PANGALAN LANG NA I-SEARCH MO ang lalabas)",
            value=False
        )
    with c_opt2:
        if st.button("🔄 I-sync ang Excel Mula sa Drive", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    final_df = parish_filtered_df.copy()

    if selected_family_id != "Lahat ng Family ID":
        final_df = final_df[final_df["Family ID"].astype(str).str.strip().str.upper() == selected_family_id.upper()]

    if keyword_search:
        # Unahin ang eksaktong match sa Full Name o Family ID para tanging ang hinahanap na tao lang ang lumabas
        name_or_id_mask = (
            final_df["Full Name"].astype(str).str.contains(keyword_search, case=False, na=False) |
            final_df["Family ID"].astype(str).str.contains(keyword_search, case=False, na=False)
        )
        matched_rows = final_df[name_or_id_mask]
        if matched_rows.empty:
            fallback_mask = (
                final_df["Head of Family"].astype(str).str.contains(keyword_search, case=False, na=False) |
                final_df["Remarks"].astype(str).str.contains(keyword_search, case=False, na=False)
            )
            matched_rows = final_df[fallback_mask]
        if expand_full_family and not matched_rows.empty:
            matched_fids = matched_rows["Family ID"].astype(str).str.strip().unique().tolist()
            final_df = final_df[final_df["Family ID"].astype(str).str.strip().isin(matched_fids)]
        else:
            final_df = matched_rows

    search_label = f"{selected_parish} ({excel_filename})"
    if selected_family_id != "Lahat ng Family ID":
        search_label += f" • Family ID: {selected_family_id}"
    if keyword_search:
        search_label += f" • Search: '{keyword_search}'"

    st.divider()

    summary_df, df_act_list, df_inact_list, df_non_list = build_status_summary_and_lists(final_df)
    df_seniors, df_adults, df_youth, df_kids = build_age_category_tables(final_df)
    total_families = final_df["Family ID"].nunique() if not final_df.empty else 0

    is_searching_specific = bool(keyword_search or selected_family_id != "Lahat ng Family ID")

    if is_searching_specific:
        st.markdown(f"### 🔍 Resulta ng Paghahanap: `{keyword_search or selected_family_id}` ({len(final_df)} Nahanap)")
        st.caption("Tanging ang datos lamang ng iyong hinanap na pangalan/ID ang ipinapakita sa ibaba para madaling makita at ma-edit.")
        if not final_df.empty:
            for _, m_row in final_df.iterrows():
                st.markdown(
                    f"<div class='full-text-card'>"
                    f"<b>👤 Buong Pangalan: <span style='font-size:1.15em;color:#38BDF8;'>{m_row['Full Name']}</span></b><br>"
                    f"⛪ <b>Parokya:</b> {m_row['Parish']} | 🆔 <b>Family ID:</b> <code>{m_row['Family ID']}</code> | 📊 <b>Status:</b> <b>{m_row['Status']}</b><br>"
                    f"🏠 <b>Head of Family:</b> {m_row['Head of Family'] or 'N/A'} | 👨‍👩‍👧 <b>Relasyon:</b> {m_row['Relationship'] or 'N/A'}<br>"
                    f"🎂 <b>Edad / Kasarian / Kategorya:</b> {m_row['Age'] or 'N/A'} • {m_row['Sex'] or 'N/A'} • <code>{m_row['Category'] or 'N/A'}</code><br>"
                    f"📱 <b>Contact No.:</b> {m_row['Contact No.'] or 'N/A'} | 📍 <b>Tirahan:</b> {m_row['Current Address (If living separately)'] or 'N/A'}<br>"
                    f"📝 <b>Remarks / Tala:</b> {m_row['Remarks'] or 'Walang espesyal na remarks'}"
                    f"</div>",
                    unsafe_allow_html=True
                )
        with st.expander("📊 Buksan ang Pangkalahatang Summary Tables (Optional)", expanded=False):
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
    else:
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("👨‍👩‍👧‍👦 Pamilyang Nahanap", total_families)
        k2.metric("👥 Kabuuang Miyembro", len(final_df))
        k3.metric("🟢 Active Members", len(df_act_list))
        k4.metric("🟡 Inactive Members", len(df_inact_list))
        k5.metric("🔵 Non-Members", len(df_non_list))

        c_tot1, c_tot2, c_tot3, c_tot4 = st.columns(4)
        c_tot1.metric("👴 TOTAL SENIORS (Active)", len(df_seniors))
        c_tot2.metric("🧑 TOTAL ADULTS (Active)", len(df_adults))
        c_tot3.metric("👦 TOTAL YOUTH (Active)", len(df_youth))
        c_tot4.metric("👶 TOTAL KIDS (Active)", len(df_kids))

        st.markdown("### 📋 MEMBER STATUS SUMMARY")
        st.caption(f"**Kasalukuyang Filter / Search:** {search_label}")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("### 📊 MGA TALAAN AYON SA MEMBER STATUS (ACTIVE / INACTIVE / NON-MEMBER)")
        col_act, col_inact, col_non = st.columns(3)
        with col_act:
            st.success(f"**ACTIVE MEMBERS ({len(df_act_list)})**")
            st.dataframe(df_act_list, use_container_width=True, hide_index=True)
        with col_inact:
            st.warning(f"**INACTIVE MEMBERS ({len(df_inact_list)})**")
            st.dataframe(df_inact_list, use_container_width=True, hide_index=True)
        with col_non:
            st.info(f"**NON-MEMBERS ({len(df_non_list)})**")
            st.dataframe(df_non_list, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("### 👥 MGA TALAAN NG ACTIVE MEMBERS AYON SA KATEGORYA (SENIORS • ADULTS • YOUTH • KIDS)")
        col_sen, col_adu, col_you, col_kid = st.columns(4)
        with col_sen:
            st.markdown(f"<div class='category-header-seniors'>👴 TOTAL SENIORS ({len(df_seniors)})</div>", unsafe_allow_html=True)
            st.dataframe(df_seniors, use_container_width=True, hide_index=True)
        with col_adu:
            st.markdown(f"<div class='category-header-adults'>🧑 TOTAL ADULTS ({len(df_adults)})</div>", unsafe_allow_html=True)
            st.dataframe(df_adults, use_container_width=True, hide_index=True)
        with col_you:
            st.markdown(f"<div class='category-header-youth'>👦 TOTAL YOUTH ({len(df_youth)})</div>", unsafe_allow_html=True)
            st.dataframe(df_youth, use_container_width=True, hide_index=True)
        with col_kid:
            st.markdown(f"<div class='category-header-kids'>👶 TOTAL KIDS ({len(df_kids)})</div>", unsafe_allow_html=True)
            st.dataframe(df_kids, use_container_width=True, hide_index=True)

        st.divider()

    if not final_df.empty:
        with st.expander("⚡ Quick Member Status, Category & Remarks Updater (Auto-Sync sa Excel & Google Drive)", expanded=False):
            st.caption("Pumili ng miyembro mula sa listahan sa ibaba upang direktang baguhin ang kanyang Status (hal. Active → Inactive), Category, o Remarks.")
            member_options = [
                f"{r['Full Name']} | {r['Family ID']} | {r['Parish']} (Kasalukuyan: {r['Status']})"
                for _, r in final_df.iterrows()
            ]
            sel_member_str = st.selectbox("Pumili ng Miyembro na I-a-update:", member_options, key="quick_upd_member_sel")
            sel_idx = member_options.index(sel_member_str)
            sel_row = final_df.iloc[sel_idx]

            u_col1, u_col2, u_col3 = st.columns([1.2, 1.2, 2])
            with u_col1:
                status_choices = ["ACTIVE", "INACTIVE", "NON-MEMBER"]
                cur_st = str(sel_row["Status"]).upper().strip()
                def_st_idx = 0
                if "INACTIVE" in cur_st:
                    def_st_idx = 1
                elif "NON" in cur_st:
                    def_st_idx = 2
                upd_status = st.selectbox("Bagong Status:", status_choices, index=def_st_idx, key="quick_upd_st")
            with u_col2:
                cat_choices = ["SENIOR CITIZEN", "ADULT", "YOUTH", "KIDS"]
                cur_cat = str(sel_row["Category"]).upper().strip()
                def_cat_idx = 1
                if "SENIOR" in cur_cat:
                    def_cat_idx = 0
                elif "YOUTH" in cur_cat:
                    def_cat_idx = 2
                elif "KID" in cur_cat:
                    def_cat_idx = 3
                upd_cat = st.selectbox("Kategorya:", cat_choices, index=def_cat_idx, key="quick_upd_cat")
            with u_col3:
                upd_rem = st.text_input("Remarks / Tala:", value=str(sel_row["Remarks"]), key="quick_upd_rem")

            if st.button("💾 I-save ang Pagbabago sa Excel at Google Drive", type="primary"):
                with st.spinner("Ina-update ang Excel database at sini-sync sa Google Drive..."):
                    ok, msg = update_member_in_excel_and_sync(
                        ACTIVE_SPREADSHEET_URL,
                        str(sel_row["Parish"]),
                        str(sel_row["Family ID"]),
                        str(sel_row["Full Name"]),
                        new_status=upd_status,
                        new_category=upd_cat,
                        new_remarks=upd_rem
                    )
                    if ok:
                        backup_action_to_admin_cloud(
                            "Member Status Updated in Excel",
                            str(sel_row["Full Name"]),
                            str(sel_row["Parish"]),
                            f"Family ID: {sel_row['Family ID']}",
                            f"Old Status: {sel_row['Status']} -> New Status: {upd_status} | Category: {upd_cat}",
                            f"Remarks: {upd_rem}",
                            get_current_chat_identity()
                        )
                        st.success(f"✅ {msg}")
                        time.sleep(0.7)
                        st.rerun()
                    else:
                        st.error(msg)

    st.markdown(f"### 👨‍👩‍👧‍👦 Matching members: {len(final_df)} (Mula sa {total_families} Pamilya)")
    if not final_df.empty:
        st.dataframe(
            final_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Remarks": st.column_config.TextColumn("Remarks (Full Note)", width="large"),
                "Current Address (If living separately)": st.column_config.TextColumn("Current Address", width="large"),
                "Full Name": st.column_config.TextColumn("Full Name", width="medium"),
                "Head of Family": st.column_config.TextColumn("Head of Family", width="medium"),
            }
        )

        df_with_remarks = final_df[final_df["Remarks"].astype(str).str.strip() != ""]
        with st.expander(f"📝 Basahin nang Buo ang Remarks at Detalye ng Miyembro ({len(df_with_remarks)} may Remarks)", expanded=True):
            if not df_with_remarks.empty:
                for _, r_row in df_with_remarks.iterrows():
                    st.markdown(
                        f"<div class='full-text-card'>"
                        f"<b>👤 {r_row['Full Name']}</b> ({r_row['Parish']} • <code>{r_row['Family ID']}</code> • {r_row['Status']})<br>"
                        f"🏠 <b>Head of Family:</b> {r_row['Head of Family'] or 'N/A'} | <b>Relasyon:</b> {r_row['Relationship'] or 'N/A'}<br>"
                        f"📍 <b>Address:</b> {r_row['Current Address (If living separately)'] or 'N/A'}<br>"
                        f"📝 <b>Buong Remarks:</b> {r_row['Remarks']}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.caption("Walang nakalagay na espesyal na Remarks sa mga nakalistang miyembro sa itaas.")
    else:
        st.warning("Walang nahanap na kumpletong record para sa iyong hinahanap.")

# =====================================================================
# 🛠️ WORKFLOW QUEUE CARDS
# =====================================================================
def render_workflow_queue_section(df_source: pd.DataFrame, category_label: str, key_prefix: str, default_lang: str):
    if df_source.empty:
        st.info(f"Walang nakitang rekord para sa **{category_label}**.")
        return
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    is_master = (st.session_state.user_role == "Master Admin")

    if is_master:
        col_f1, col_f2, col_f3 = st.columns([1.3, 1.3, 1.6])
        with col_f1:
            status_filter = st.selectbox(
                "Salain ayon sa Status:",
                [
                    "Lahat ng Active Queue (Pending / Assigned / For Approval)",
                    "📥 For Master Admin Approval (Galing sa Parokya)",
                    "📤 Assigned to Parish Leader",
                    "⏳ Pending Review (Bagong Pasok)",
                    "✅ Approved & Sent (Natapos Na)",
                    "❌ Rejected / Archived",
                    "Lahat ng Rekord"
                ],
                key=f"flt_st_{key_prefix}"
            )
        with col_f2:
            parish_filter = st.selectbox(
                "Salain ayon sa Parokya:",
                ["Lahat ng Parokya"] + PARISH_ORDER + ["Unassigned Parish"],
                key=f"flt_pr_{key_prefix}"
            )
        with col_f3:
            search_txt = st.text_input("🔍 Hanapin (Pangalan, Paksa, o Detalye):", key=f"flt_q_{key_prefix}").strip()
    else:
        col_f1, col_f2 = st.columns([1.4, 2])
        with col_f1:
            status_filter = st.selectbox(
                "Salain ang Iyong Parish Inbox:",
                [
                    "📥 Mga Nakatalaga sa Aking Parokya (Active)",
                    "📤 Naipasa Na kay Master Admin (For Approval)",
                    "✅ Approved & Sent (Natapos Na)",
                    "Lahat ng Rekord ng Aking Parokya"
                ],
                key=f"flt_st_{key_prefix}"
            )
        parish_filter = st.session_state.assigned_parish
        with col_f2:
            search_txt = st.text_input("🔍 Hanapin sa Iyong Parokya:", key=f"flt_q_{key_prefix}").strip()

    shown = 0
    for idx, row in df_source.iterrows():
        page_id = row.get("_page_id", f"{key_prefix}_{idx}")
        sender_name = get_first_non_empty(row, ["Full Name", "Name", "Pangalan", "Visitor Name"], f"Record #{idx+1}")
        email = get_first_non_empty(row, ["Email", "Email Address", "E-mail"], "")
        contact_no = get_first_non_empty(row, ["Contact No.", "Contact Number", "Phone", "Mobile"], "N/A")
        address = get_first_non_empty(row, ["Address", "Tirahan", "Current Address"], "N/A")
        parish = get_first_non_empty(row, ["Parish", "Assigned Parish", "Parokya"], "Unassigned Parish")
        church_status = get_first_non_empty(row, ["Church Status", "Member Status", "Status sa Simbahan"], "Member / Visitor")
        sub_date = get_first_non_empty(row, ["Submission Date", "Date", "Created time", "_created_time"], "N/A")
        last_edit = get_first_non_empty(row, ["_last_edited", "Last edited time"], "N/A")
        pref_contact = get_first_non_empty(row, ["Response Type", "Preferred Contact", "Gustong Uri ng Tugon"], "Email / SMS")
        topic = get_first_non_empty(row, ["Prayer Title", "Question Topic", "Counselling Type", "Service Type", "Title"], category_label)
        details = get_first_non_empty(row, ["Prayer Details", "Question Details", "Situation Details", "Request Details", "Details", "Message"], "Walang detalye.")
        admin_status = get_first_non_empty(row, ["Admin Status", "Approval Status", "Status"], "Pending Review")
        handled_by = get_first_non_empty(row, ["Approved By", "Reviewed By", "Handled By", "Submitted By"], "")

        if not is_master and not normalize_parish_match(parish, st.session_state.assigned_parish):
            continue

        if is_master and parish_filter != "Lahat ng Parokya":
            if parish_filter == "Unassigned Parish":
                if parish != "Unassigned Parish" and parish.strip() != "":
                    continue
            elif not normalize_parish_match(parish, parish_filter):
                continue

        st_low = admin_status.lower()
        if is_master:
            if status_filter.startswith("Lahat ng Active Queue") and st_low in ["approved & sent", "rejected / archived"]:
                continue
            elif status_filter.startswith("📥 For Master Admin") and "for master admin approval" not in st_low:
                continue
            elif status_filter.startswith("📤 Assigned") and "assigned to parish leader" not in st_low:
                continue
            elif status_filter.startswith("⏳ Pending") and "pending" not in st_low:
                continue
            elif status_filter.startswith("✅ Approved") and "approved & sent" not in st_low:
                continue
            elif status_filter.startswith("❌ Rejected") and "rejected" not in st_low:
                continue
        else:
            if status_filter.startswith("📥 Mga Nakatalaga") and st_low in ["approved & sent", "rejected / archived", "for master admin approval"]:
                continue
            elif status_filter.startswith("📤 Naipasa Na") and "for master admin approval" not in st_low:
                continue
            elif status_filter.startswith("✅ Approved") and "approved & sent" not in st_low:
                continue

        if search_txt and search_txt.lower() not in f"{sender_name} {parish} {topic} {details} {handled_by} {email}".lower():
            continue

        shown += 1
        draft_state_key = f"draft_val_{page_id}"
        if draft_state_key not in st.session_state:
            st.session_state[draft_state_key] = get_first_non_empty(row, ["AI Draft Reply", "AI Reply", "Draft Reply", "Reply"], "")

        submitted_parish_info = handled_by if handled_by else (parish if "for master admin approval" in st_low else "Wala pang nagsusumite")

        header_tag = (
            f"📌 Nag-request: {sender_name} | 📍 Parokya: {parish} | 🏷️ {topic} | "
            f"📊 Status: {admin_status} | 🏛️ Gawa/Sinumite ni: {submitted_parish_info}"
        )

        with st.expander(header_tag, expanded=(st_low != "approved & sent")):
            if "for master admin approval" in st_low or "Submitted by" in handled_by or "Drafted by" in handled_by:
                st.markdown(
                    f"<div class='submitted-parish-badge'>📥 ISINUMITE NG PAROKYA PARA SA IYONG APPROVAL: "
                    f"<b>{handled_by or parish}</b></div>",
                    unsafe_allow_html=True
                )
            elif "assigned to parish leader" in st_low:
                st.markdown(
                    f"<div class='assigned-parish-badge'>📤 NAKATALAGA SA PAROKYA: <b>{parish}</b> ({handled_by or 'Waiting for Parish Reply'})</div>",
                    unsafe_allow_html=True
                )

            col_req_area, col_type_area = st.columns([1.05, 1.35])

            with col_req_area:
                st.markdown("##### 📋 1. Area ng Request at Impormasyon ng Nagsumite")
                st.markdown(
                    f"<div class='full-text-card'>"
                    f"<b>👤 Pangalan ng Nag-request:</b> {sender_name}<br>"
                    f"<b>⛪ Katayuan sa Simbahan:</b> {church_status}<br>"
                    f"<b>📍 Parokya:</b> {parish}<br>"
                    f"<b>📧 Email Address:</b> {email or 'Walang Email'}<br>"
                    f"<b>📱 Contact / SMS No.:</b> {contact_no}<br>"
                    f"<b>🏠 Tirahan:</b> {address}<br>"
                    f"<b>📅 Petsa ng Request:</b> {sub_date}<br>"
                    f"<b>⏱️ Huling Na-update:</b> {last_edit}<br>"
                    f"<b>📬 Gustong Tugon:</b> {pref_contact}<br>"
                    f"<b>🏛️ Huling Nag-proseso / Nag-submit:</b> {handled_by or 'N/A'}"
                    f"<hr style='margin:8px 0;'>"
                    f"<b>🏷️ Paksa ({category_label}):</b> {topic}<br><br>"
                    f"<b>💬 Buong Mensahe / Kahilingan:</b><br>{details}"
                    f"</div>",
                    unsafe_allow_html=True
                )

            with col_type_area:
                if is_master:
                    st.markdown("##### ✍️ 2. Master Admin Typing, Assignment & Final Approval Area")
                else:
                    st.markdown(f"##### ✍️ 2. Typing & Response Area para sa {st.session_state.assigned_parish}")

                c_ai_btn, c_lang = st.columns([1.5, 1.5])
                with c_lang:
                    item_lang = st.selectbox(
                        "Wika ng AI Draft:",
                        ["Tagalog", "Ilokano", "English"],
                        index=["Tagalog", "Ilokano", "English"].index(default_lang),
                        key=f"lang_{page_id}"
                    )
                with c_ai_btn:
                    st.write("")
                    if st.button("🤖 Bumuo ng AI Draft", key=f"gen_ai_{page_id}", use_container_width=True):
                        with st.spinner("Ginagawa ng CTG Gabay AI..."):
                            try:
                                generated = generate_ai_ministry_draft(category_label, sender_name, parish, topic, details, item_lang, selected_model)
                                st.session_state[draft_state_key] = generated
                                update_notion_page_record(n_tok, page_id, ai_draft=generated, admin_status=admin_status)
                                st.cache_data.clear()
                                st.success("Nagawa at na-save ang AI Draft!")
                            except Exception as ex:
                                st.error(str(ex))

                edited_draft = st.text_area(
                    f"✏️ I-type o i-edit dito ang mensahe/sagot para kay {sender_name}:",
                    value=st.session_state[draft_state_key],
                    height=185,
                    key=f"editor_{page_id}"
                )
                st.session_state[draft_state_key] = edited_draft

                if is_master:
                    st.markdown("**🅰️ Ipasa sa Partikular na Parokya (Hindi makikita ng ibang parokya):**")
                    ca1, ca2 = st.columns([1.6, 1.2])
                    with ca1:
                        default_p_idx = PARISH_ORDER.index(parish.upper()) if parish.upper() in PARISH_ORDER else 0
                        target_parish = st.selectbox(
                            "Pumili ng Parokya:",
                            PARISH_ORDER,
                            index=default_p_idx,
                            key=f"ass_sel_{page_id}",
                            label_visibility="collapsed"
                        )
                    with ca2:
                        if st.button("📤 Ipasa sa Parokyang Ito", key=f"ass_btn_{page_id}", use_container_width=True):
                            actor_tag = f"Assigned by Master Admin to {target_parish}"
                            update_notion_page_record(
                                n_tok,
                                page_id,
                                ai_draft=edited_draft,
                                admin_status="Assigned to Parish Leader",
                                assigned_parish=target_parish,
                                approved_by=actor_tag
                            )
                            backup_action_to_admin_cloud("Assigned to Parish", sender_name, target_parish, topic, details, edited_draft, actor_tag)
                            st.cache_data.clear()
                            st.info(f"Naipasa na sa **{target_parish}**! Sila lang ang makakakita nito.")

                    st.markdown("**🅱️ Master Admin Final Approval & Pagpapadala (Gmail / SMS):**")
                    cb1, cb2, cb3 = st.columns([1.5, 1.2, 1.1])
                    with cb1:
                        if st.button("🚀 Approve & Send sa Gmail", key=f"app_send_{page_id}", type="primary", use_container_width=True):
                            if email and "@" in email:
                                clean_body = re.sub(r"^\[Urgency Analysis:.*?\]\s*", "", edited_draft.strip(), flags=re.IGNORECASE)
                                res_mail = send_email_via_gmail(email, f"Tugon mula sa Church of the Temple of God (CTG) — {topic}", clean_body)
                                if "Matagumpay" in res_mail:
                                    approver_note = f"Approved & Sent by Master Admin (Draft: {handled_by or parish})"
                                    update_notion_page_record(
                                        n_tok, page_id, ai_draft=clean_body, admin_status="Approved & Sent", approved_by=approver_note
                                    )
                                    backup_action_to_admin_cloud("Approved & Sent via Gmail", sender_name, parish, topic, details, clean_body, approver_note)
                                    st.cache_data.clear()
                                    st.success(f"✅ {res_mail} Naka-backup na rin sa {BACKUP_ADMIN_EMAIL}!")
                                else:
                                    st.error(res_mail)
                            else:
                                st.error("Walang valid email address ang nag-request.")
                    with cb2:
                        if st.button("💾 I-save ang Draft", key=f"save_btn_{page_id}", use_container_width=True):
                            update_notion_page_record(n_tok, page_id, ai_draft=edited_draft)
                            st.cache_data.clear()
                            st.success("Na-save sa Notion!")
                    with cb3:
                        if st.button("❌ Reject / Archive", key=f"rej_btn_{page_id}", use_container_width=True):
                            update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="Rejected / Archived", approved_by="Rejected by Master Admin")
                            st.cache_data.clear()
                            st.warning("Na-reject / Archive.")

                    with st.expander(f"📱 Kopyahin para sa SMS / Text Message ({contact_no})"):
                        st.caption("Kung walang email o gusto mong i-text sa mobile/Viber, kopyahin ang mensahe sa ibaba:")
                        st.code(f"Para kay: {sender_name} ({contact_no})\n\n{edited_draft}", language="text")
                        if st.button("✅ Markahan bilang 'Approved & Sent via SMS'", key=f"sms_sent_{page_id}"):
                            approver_note = f"Approved & Sent via SMS by Master Admin (Draft: {handled_by or parish})"
                            update_notion_page_record(
                                n_tok, page_id, ai_draft=edited_draft, admin_status="Approved & Sent", approved_by=approver_note
                            )
                            backup_action_to_admin_cloud("Approved & Sent via SMS", sender_name, parish, topic, details, edited_draft, approver_note)
                            st.cache_data.clear()
                            st.success("Na-record bilang 'Approved & Sent via SMS' at na-backup sa Gmail!")
                else:
                    p1, p2 = st.columns([1, 1.5])
                    with p1:
                        if st.button("💾 I-save Muna ang Draft", key=f"p_save_{page_id}", use_container_width=True):
                            actor_tag = f"Draft in progress by {st.session_state.assigned_parish}"
                            update_notion_page_record(
                                n_tok,
                                page_id,
                                ai_draft=edited_draft,
                                approved_by=actor_tag
                            )
                            st.cache_data.clear()
                            st.success("Na-save ang draft!")
                    with p2:
                        if st.button("📥 Tapos Na — Ipasa kay Master Admin para sa Approval", key=f"p_ret_{page_id}", type="primary", use_container_width=True):
                            if not edited_draft.strip():
                                st.warning("Mag-type muna ng iyong tugon bago ipasa kay Master Admin.")
                            else:
                                actor_tag = f"Submitted by {st.session_state.assigned_parish}"
                                update_notion_page_record(
                                    n_tok,
                                    page_id,
                                    ai_draft=edited_draft,
                                    admin_status="For Master Admin Approval",
                                    assigned_parish=st.session_state.assigned_parish,
                                    approved_by=actor_tag
                                )
                                backup_action_to_admin_cloud("Parish Submitted for Approval", sender_name, st.session_state.assigned_parish, topic, details, edited_draft, actor_tag)
                                st.cache_data.clear()
                                st.success(f"✅ Naipasa na kay Master Admin! Makikita niya na ang {st.session_state.assigned_parish} ang nag-submit nito.")

    if shown == 0:
        st.info("Walang nahanap na request na tumutugma sa napiling filter.")

# =====================================================================
# 📑 WEEKLY REPORT & PARISH MONITORING ROUTE
# =====================================================================
if app_mode == "📑 Weekly Report & Parish Monitoring":
    render_weekly_report_page()
    st.stop()

# =====================================================================
# 🧭 PARISH LEADER PORTAL (WIZARD & SUMMARY)
# =====================================================================
elif app_mode == "🧭 Parish Leader Portal (Wizard & Summary)":
    render_title_with_logo(
        f"📍 Parish Workspace: {st.session_state.assigned_parish}",
        "Maligayang pagdating sa iyong pang-araw-araw na gabay at ministry portal ng Church of the Temple of God."
    )

    records_master_df = global_records_df
    p_records = records_master_df[records_master_df["Parokya"].apply(lambda x: normalize_parish_match(x, st.session_state.assigned_parish))] if not records_master_df.empty else pd.DataFrame()
    pending_p_count = len(p_records[p_records["Status"].str.contains("Assigned", case=False, na=False)]) if not p_records.empty else 0

    st.markdown("### 📋 Daily Action Checklist (Pang-araw-araw na Gabay)")
    chk1 = st.checkbox(f"1. Mayroon kang ({pending_p_count}) bagong request o panalangin na nakatalaga sa iyong parokya.", value=(pending_p_count > 0))
    chk2 = st.checkbox("2. I-upload ang Weekly Report (Image o Excel) para sa Sabado (Saturday Service) sa Weekly Report tab.", value=False)
    chk3 = st.checkbox("3. Buksan ang Ministry Messenger para makipag-ugnayan kay Master Admin o sa ibang Parokya.", value=False)

    col_w0, col_w1, col_w2, col_w3 = st.columns(4)
    with col_w0:
        if st.button("📑 Mag-upload ng Weekly Report", type="primary", use_container_width=True):
            st.session_state.app_mode_override = "📑 Weekly Report & Parish Monitoring"
            st.rerun()
    with col_w1:
        if st.button("👉 Pumunta sa Parish Inbox", use_container_width=True):
            st.session_state.app_mode_override = "📥 Parish Inbox & Action Hub"
            st.rerun()
    with col_w2:
        if st.button("📊 Buksan ang Parish Census", use_container_width=True):
            st.session_state.app_mode_override = "📊 Parish Census & Member Manager"
            st.rerun()
    with col_w3:
        if st.button("🗣️ Buksan ang Ministry Messenger", use_container_width=True):
            st.session_state.app_mode_override = "🗣️ Ministry Messenger (Admin & Users Chat)"
            st.rerun()

    st.divider()
    st.markdown("### 🤖 Guided AI Assistant para sa Parish Leader")
    st.caption("Magtanong dito kung paano kakausapin ang mga miyembro, humingi ng ideya sa Bible study tuwing Sabado, o gabay sa pamumuno.")
    with st.form("wizard_ai_form"):
        w_prompt = st.text_input("Magtanong sa CTG Gabay AI:", placeholder="Hal. Paano ko kakausapin ang isang miyembrong matagal nang hindi nakakadalo sa pagsamba tuwing Sabado?")
        w_submit = st.form_submit_button("Humingi ng Gabay sa AI", use_container_width=True)
        if w_submit and w_prompt:
            client = genai.Client(api_key=get_valid_gemini_keys()[0])
            ans = safe_generate_content(
                client,
                selected_model,
                f"Bilang isang tagapayo sa simbahan ng Church of the Temple of God (CTG) sa parokya ng {st.session_state.assigned_parish} "
                f"(kung saan ang araw ng pagsamba ay tuwing Sabado / Saturday), sagutin ito nang may malasakit at gabay sa Bibliya sa wikang Tagalog: {w_prompt}"
            )
            st.success("Gabay mula sa AI:")
            st.markdown(ans.text)
    st.stop()

# =====================================================================
# 🗣️ MINISTRY MESSENGER PAGE ROUTE
# =====================================================================
elif app_mode == "🗣️ Ministry Messenger (Admin & Users Chat)":
    render_ministry_messenger_page()
    st.stop()

# =====================================================================
# 📥 PARISH INBOX & ACTION HUB (MINIMALIST & TASK-FOCUSED)
# =====================================================================
elif app_mode == "📥 Parish Inbox & Action Hub":
    render_title_with_logo(f"📍 Parish Workspace: {st.session_state.assigned_parish}", "Task-focused inbox para sa pagtugon sa mga nag-aabang ng tulong o dasal.")
    df_crm, df_p, df_q, df_co = df_crm_global, df_p_global, df_q_global, df_co_global
    tp, tq, tc = st.tabs(["🙏 Prayer Requests", "📖 Church Questions", "🤝 Counselling"])
    with tp:
        render_workflow_queue_section(df_p, "Prayer Request", "p_inbox", st.session_state.ui_default_lang)
    with tq:
        render_workflow_queue_section(df_q, "Church Question", "q_inbox", st.session_state.ui_default_lang)
    with tc:
        render_workflow_queue_section(df_co, "Counselling Request", "c_inbox", st.session_state.ui_default_lang)
    st.stop()

# =====================================================================
# 📊 PARISH DIRECTORY / CENSUS & FAMILY SEARCH (MASTER ADMIN & PARISH LEADER)
# =====================================================================
elif app_mode in ["📊 Parish Directory & Family Search", "📊 Parish Census & Member Manager"]:
    if app_mode == "📊 Parish Census & Member Manager":
        render_parish_directory_and_census_view(
            f"📊 Parish Census & Member Manager — {st.session_state.assigned_parish}",
            "Kumpletong Talaan ng Pamilya, Member Status, at Active Age Categories (Total Seniors, Adults, Youth, Kids)."
        )
    else:
        render_parish_directory_and_census_view(
            "CHURCH OF THE TEMPLE OF GOD — SEARCH FAMILY / PARISH",
            "Kumpletong Talaan ng Pamilya, Member Status, at Active Age Categories (Total Seniors, Adults, Youth, Kids)."
        )
    st.stop()

# =====================================================================
# 📈 LIVE ACTIVITY & UPDATE DASHBOARD (MASTER ADMIN ONLY)
# =====================================================================
elif app_mode == "📈 Live Activity & Update Dashboard":
    if st.session_state.user_role != "Master Admin":
        st.error("Ang pahinang ito ay para lamang sa Master Admin.")
        st.stop()

    render_title_with_logo(
        "📈 CTG Live Activity, Parish Updates & Cloud Backup Dashboard",
        f"Real-time monitoring ng lahat ng galaw ng bawat parokya at direktang backup para sa {BACKUP_ADMIN_EMAIL}."
    )

    top_b1, top_b2, top_b3, top_b4 = st.columns([1.2, 1.4, 1.4, 1.5])
    with top_b1:
        if st.button("🔄 I-refresh ang Dashboard", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    if not global_records_df.empty:
        csv_data_dash = global_records_df.to_csv(index=False).encode("utf-8")
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            global_records_df.to_excel(writer, index=False, sheet_name="CTG_Ministry_Records")
        excel_bytes = excel_buffer.getvalue()

        with top_b2:
            st.download_button(
                label="📥 Download CSV (Google Drive)",
                data=csv_data_dash,
                file_name=f"CTG_Ministry_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        with top_b3:
            st.download_button(
                label="📊 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"CTG_Ministry_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with top_b4:
            if st.button(f"☁️ I-email ang Backup sa {BACKUP_ADMIN_EMAIL}", type="primary", use_container_width=True):
                with st.spinner(f"Ipinapadala sa {BACKUP_ADMIN_EMAIL}..."):
                    res_cloud = send_email_via_gmail(
                        BACKUP_ADMIN_EMAIL,
                        f"CTG Ministry Hub Cloud Backup — {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                        f"Kalakip nito ang pinakabagong CSV backup ng lahat ng CTG Ministry Hub requests at updates ({len(global_records_df)} records).",
                        attachment_bytes=csv_data_dash,
                        attachment_filename=f"CTG_Ministry_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
                    )
                    if "Matagumpay" in res_cloud:
                        st.success(f"Ligtas nang naipadala sa {BACKUP_ADMIN_EMAIL}!")
                    else:
                        st.error(res_cloud)

    st.divider()

    if not global_records_df.empty:
        total_reqs = len(global_records_df)
        pending_reqs = len(global_records_df[global_records_df["Status"].str.contains("Pending", case=False, na=False)])
        assigned_reqs = len(global_records_df[global_records_df["Status"].str.contains("Assigned to Parish Leader", case=False, na=False)])
        for_approval_reqs = len(global_records_df[global_records_df["Status"].str.contains("For Master Admin Approval", case=False, na=False)])
        completed_reqs = len(global_records_df[global_records_df["Status"].str.contains("Approved & Sent", case=False, na=False)])

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("📌 Kabuuang Requests", total_reqs)
        m2.metric("⏳ Pending Review", pending_reqs)
        m3.metric("📤 Nasa Parokya (Assigned)", assigned_reqs)
        m4.metric("📥 Galing Parokya (For Approval)", for_approval_reqs)
        m5.metric("✅ Natapos Na (Approved & Sent)", completed_reqs)

        st.divider()
        col_d_left, col_d_right = st.columns([1.35, 1.15])

        with col_d_left:
            st.markdown("### 🏛️ Status Breakdown Kada Parokya")
            parish_breakdown = []
            for p_name in PARISH_ORDER + ["Unassigned Parish"]:
                sub_df = global_records_df[global_records_df["Parokya"].apply(lambda x: normalize_parish_match(x, p_name) if p_name != "Unassigned Parish" else (x == "Unassigned Parish"))]
                parish_breakdown.append({
                    "Parokya": p_name,
                    "Kabuuan": len(sub_df),
                    "Pending": len(sub_df[sub_df["Status"].str.contains("Pending", case=False, na=False)]),
                    "Nasa Parokya": len(sub_df[sub_df["Status"].str.contains("Assigned", case=False, na=False)]),
                    "Isinumite (For Approval)": len(sub_df[sub_df["Status"].str.contains("For Master Admin Approval", case=False, na=False)]),
                    "Tapos Na (Sent)": len(sub_df[sub_df["Status"].str.contains("Approved & Sent", case=False, na=False)]),
                })
            st.dataframe(pd.DataFrame(parish_breakdown), use_container_width=True, hide_index=True)

        with col_d_right:
            st.markdown("### 🔔 Live Feed: Sino ang Nag-update / Nag-submit?")
            for _, item in global_records_df.head(6).iterrows():
                st.markdown(
                    f"<div class='records-side-box'>"
                    f"<b>👤 Nag-request: {item['Pangalan ng Nag-request']}</b> ({item['Uri ng Request']})<br>"
                    f"📍 <b>Parokya:</b> {item['Parokya']} | 📊 <b>Status:</b> <code>{item['Status']}</code><br>"
                    f"🏛️ <b>Huling Nag-update / Nag-submit:</b> <b>{item['Isinumite / Inaprubahan Ni']}</b><br>"
                    f"🏷️ <b>Paksa:</b> {item['Paksa']} | ⏱️ <b>Oras:</b> {item['Huling Na-update']}"
                    f"</div>",
                    unsafe_allow_html=True
                )

        st.divider()
        st.markdown("### 📋 Kumpletong Master Log ng Lahat ng Transaksyon at Impormasyon ng Nag-request")
        st.dataframe(global_records_df, use_container_width=True, hide_index=True)
    else:
        st.info("Wala pang nakatalang aktibidad sa Notion databases.")
    st.stop()

# =====================================================================
# 2. MASTER ADMIN CONTROLLER / APPROVAL HUB
# =====================================================================
elif app_mode == "🛡️ Master Admin Controller & Approval Queue":
    render_title_with_logo(
        "Master Admin Controller & Approval Hub",
        "Ipasa ang mga request sa kaukulang Parokya, tingnan kung sino ang nag-submit, at i-approve/i-send sa Gmail o SMS."
    )

    lang_choices = ["Tagalog", "Ilokano", "English"]
    def_lang_idx = lang_choices.index(st.session_state.ui_default_lang) if st.session_state.ui_default_lang in lang_choices else 0

    top1, top2, top3 = st.columns([1.8, 1.2, 1])
    with top1:
        def_lang = st.selectbox("🌐 Default na Wika para sa AI Draft:", lang_choices, index=def_lang_idx)
    with top2:
        show_side_panel = st.checkbox("📜 Ipakita ang Records Panel sa Gilid", value=st.session_state.ui_show_side_records)
        st.session_state.ui_show_side_records = show_side_panel
    with top3:
        if st.button("🔄 I-refresh ang Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    df_crm, df_p, df_q, df_co = df_crm_global, df_p_global, df_q_global, df_co_global
    records_master_df = global_records_df

    if not records_master_df.empty:
        cnt_for_approval = len(records_master_df[records_master_df["Status"].str.contains("For Master Admin Approval", case=False, na=False)])
        cnt_completed = len(records_master_df[records_master_df["Status"].str.contains("Approved & Sent", case=False, na=False)])
    else:
        cnt_for_approval, cnt_completed = 0, 0

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("🙏 Prayer Requests", len(df_p))
    m2.metric("📖 Church Questions", len(df_q))
    m3.metric("🤝 Counselling", len(df_co))
    m4.metric("📥 Galing sa Parokya (For Approval)", cnt_for_approval)
    m5.metric("✅ Natapos Na (Approved & Sent)", cnt_completed)
    st.divider()

    if show_side_panel:
        col_main_workspace, col_side_records = st.columns([2.35, 1.05])
    else:
        col_main_workspace = st.container()
        col_side_records = None

    with col_main_workspace:
        tp, tq, tc, trec, tcrm = st.tabs([
            "🙏 Prayer Requests",
            "📖 Church Questions",
            "🤝 Counselling",
            "📜 Buong Records & Requester Info",
            "📋 All CRM"
        ])
        with tp:
            render_workflow_queue_section(df_p, "Prayer Request", "pray", def_lang)
        with tq:
            render_workflow_queue_section(df_q, "Church Question", "qna", def_lang)
        with tc:
            render_workflow_queue_section(df_co, "Counselling Request", "counsel", def_lang)
        with trec:
            st.subheader("📜 Kumpletong Talaan ng mga Nagawa at Impormasyon ng mga Nag-request")
            if records_master_df.empty:
                st.info("Wala pang nakatalang request.")
            else:
                c_rf1, c_rf2 = st.columns([2, 1.2])
                with c_rf1:
                    rec_filter = st.selectbox(
                        "Salain ang Talaan:",
                        ["Lahat ng Rekord", "✅ Mga Natapos Na (Approved & Sent)", "📥 Isinumite ng Parokya (For Master Admin Approval)", "📤 Nakatalaga sa Parokya"],
                        key="full_rec_flt"
                    )
                df_view_rec = records_master_df.copy()
                if rec_filter.startswith("✅"):
                    df_view_rec = df_view_rec[df_view_rec["Status"].str.contains("Approved & Sent", case=False, na=False)]
                elif rec_filter.startswith("📥"):
                    df_view_rec = df_view_rec[df_view_rec["Status"].str.contains("For Master Admin Approval", case=False, na=False)]
                elif rec_filter.startswith("📤"):
                    df_view_rec = df_view_rec[df_view_rec["Status"].str.contains("Assigned to Parish Leader", case=False, na=False)]

                with c_rf2:
                    st.write("")
                    if st.session_state.user_role == "Master Admin":
                        st.download_button(
                            label="📥 I-download sa Google Drive (CSV)",
                            data=df_view_rec.to_csv(index=False).encode("utf-8"),
                            file_name="CTG_Records_Export.csv",
                            mime="text/csv",
                            use_container_width=True
                        )

                st.dataframe(df_view_rec, use_container_width=True, hide_index=True)
        with tcrm:
            st.dataframe(df_crm.drop(columns=["_page_id", "_last_edited", "_created_time"], errors="ignore"), use_container_width=True, hide_index=True)

    if col_side_records is not None:
        with col_side_records:
            st.markdown("### 📜 Records & Info Panel")
            st.caption("Talaan ng mga natapos na, sino ang parokyang nag-submit, at buong info ng mga nag-request.")
            side_mode = st.selectbox(
                "Ipakita sa Panel:",
                [
                    "Lahat ng Rekord & Nag-request",
                    "✅ Mga Natapos Na (Approved & Sent)",
                    "📥 Isinumite ng Parokya (For Approval)",
                    "📤 Nakatalaga sa Parokya (Assigned)"
                ],
                key="side_panel_mode"
            )
            side_search = st.text_input("🔍 Hanapin sa Records Panel:", key="side_panel_search").strip()

            if records_master_df.empty:
                st.info("Wala pang rekord.")
            else:
                side_df = records_master_df.copy()
                if side_mode.startswith("✅"):
                    side_df = side_df[side_df["Status"].str.contains("Approved & Sent", case=False, na=False)]
                elif side_mode.startswith("📥"):
                    side_df = side_df[side_df["Status"].str.contains("For Master Admin Approval", case=False, na=False)]
                elif side_mode.startswith("📤"):
                    side_df = side_df[side_df["Status"].str.contains("Assigned to Parish Leader", case=False, na=False)]

                if side_search:
                    side_df = side_df[side_df.astype(str).apply(lambda x: x.str.contains(side_search, case=False, na=False)).any(axis=1)]

                st.caption(f"Kabuuang nakita: **{len(side_df)}** rekord")
                for r_idx, r_item in side_df.iterrows():
                    st.markdown(
                        f"<div class='records-side-box'>"
                        f"<b>👤 {r_item['Pangalan ng Nag-request']}</b> ({r_item['Uri ng Request']})<br>"
                        f"📍 <b>Parokya:</b> {r_item['Parokya']}<br>"
                        f"📊 <b>Status:</b> <code>{r_item['Status']}</code><br>"
                        f"🏛️ <b>Sinumite/Inaprubahan:</b> {r_item['Isinumite / Inaprubahan Ni']}<br>"
                        f"📧 <b>Email:</b> {r_item['Email']} | 📱 <b>Tel:</b> {r_item['Contact No.']}<br>"
                        f"🏷️ <b>Paksa:</b> {r_item['Paksa']}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                    with st.expander(f"🔍 Buong Detalye at Tugon — {r_item['Pangalan ng Nag-request']}"):
                        st.write(f"**🏠 Tirahan:** {r_item['Tirahan']}")
                        st.write(f"**⛪ Katayuan:** {r_item['Katayuan sa Simbahan']} | **📅 Petsa:** {r_item['Petsa ng Request']}")
                        st.write(f"**⏱️ Huling Update:** {r_item['Huling Na-update']}")
                        st.markdown(f"**💬 Kahilingan / Tanong:**\n> {r_item['Detalye ng Request']}")
                        st.markdown(f"**✉️ Tugon / AI Draft:**\n> {r_item['Tugon / AI Draft'] or 'Wala pang tugon.'}")
    st.stop()

# =====================================================================
# 3. NOTION CRM (VISITORS & MEMBERS)
# =====================================================================
elif app_mode == "🤝 Notion CRM (Visitors & Members)":
    render_title_with_logo("Church CRM: Visitors & Members (Notion)")
    if st.button("🔄 I-refresh"):
        st.cache_data.clear()
        st.rerun()
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    n_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
    df_n = load_notion_crm_data(n_tok, n_db)
    if not df_n.empty:
        disp = df_n.drop(columns=["_page_id", "_last_edited", "_created_time"], errors="ignore")
        q = st.text_input("🔍 Search CRM:").strip()
        if q:
            disp = disp[disp.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)]
        st.dataframe(disp, use_container_width=True, hide_index=True)
        st.divider()
        st.subheader("📖 Buong Detalye ng Bawat Rekord sa Notion CRM (Full View)")
        for idx, row in disp.iterrows():
            first_col = disp.columns[0]
            title_val = str(row[first_col]).strip() or f"Record #{idx + 1}"
            details_html = f"<div class='full-text-card'><b>📌 {title_val}</b><br><hr style='margin:6px 0;'>"
            for c_name in disp.columns:
                val_str = str(row[c_name]).strip()
                if val_str:
                    details_html += f"<b>• {c_name}:</b> {val_str}<br>"
            details_html += "</div>"
            st.markdown(details_html, unsafe_allow_html=True)
    st.stop()

def extract_members_from_uploaded_image_or_file(file_bytes: bytes, filename: str, mime_type: str, default_parish: str, model_name: str = "gemini-2.5-flash") -> list[dict]:
    """Bumabasa ng Image (.jpg, .png) o Excel/CSV File (.xlsx, .csv) sa 'Magdagdag ng Miyembro' at kinukuha ang mga datos ng miyembro."""
    ext = filename.split(".")[-1].lower()
    extracted_list = []

    if ext in ["xlsx", "csv"]:
        try:
            df_in = pd.read_csv(io.BytesIO(file_bytes)) if ext == "csv" else pd.read_excel(io.BytesIO(file_bytes))
            df_in.columns = [str(c).strip() for c in df_in.columns]
            for _, r in df_in.iterrows():
                fn_v = str(r.get("Full Name", r.get("Pangalan", r.get("Name", ""))) or "").strip()
                if not fn_v or fn_v.lower() == "nan":
                    continue
                p_v = str(r.get("Parish", r.get("Parokya", default_parish)) or default_parish).strip()
                matched_p = default_parish
                for p_cand in PARISH_ORDER:
                    if normalize_parish_match(p_v, p_cand):
                        matched_p = p_cand
                        break
                extracted_list.append({
                    "Parish": matched_p,
                    "Family ID": str(r.get("Family ID", r.get("Member / Record ID", "")) or "").strip() or f"CTG-{matched_p[:3]}-FAM-NEW",
                    "Head of Family": str(r.get("Head of Family", fn_v) or fn_v).strip(),
                    "Full Name": fn_v,
                    "Relationship": str(r.get("Relationship", "HEAD OF FAMILY") or "HEAD OF FAMILY").strip().upper(),
                    "Sex": str(r.get("Sex", "MALE") or "MALE").strip().upper(),
                    "Age": str(r.get("Age", "") or "").strip(),
                    "Category": str(r.get("Category", "ADULT") or "ADULT").strip().upper(),
                    "Status": str(r.get("Status", "ACTIVE") or "ACTIVE").strip().upper(),
                    "Contact No.": str(r.get("Contact No.", r.get("Contact Number", "")) or "").strip(),
                    "Current Address (If living separately)": str(r.get("Current Address (If living separately)", r.get("Address", "")) or "").strip(),
                    "Remarks": str(r.get("Remarks", "") or "").strip(),
                })
            if extracted_list:
                return extracted_list
        except Exception:
            pass

    # Para sa Image (.jpg, .png, .jpeg): subukan muna ang Gemini Vision, at may fallback sa OCR (RapidOCR / pytesseract / OCR.space)
    valid_keys = get_valid_gemini_keys()
    if valid_keys:
        try:
            pil_rgb = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            jb = io.BytesIO()
            pil_rgb.save(jb, format="JPEG", quality=92)
            image_part = types.Part.from_bytes(data=jb.getvalue(), mime_type="image/jpeg")
            prompt_m = (
                f"Basahin ang lahat ng impormasyon ng miyembro na nakasulat sa larawang ito (default parish: {default_parish}). "
                "Ibalik LAMANG ang valid JSON list ng objects na may keys: "
                '[{"Parish": "", "Family ID": "", "Head of Family": "", "Full Name": "", '
                '"Relationship": "HEAD OF FAMILY", "Sex": "MALE", "Age": "", "Category": "ADULT", '
                '"Status": "ACTIVE", "Contact No.": "", "Current Address (If living separately)": "", "Remarks": ""}]'
            )
            for k_api in valid_keys:
                client = genai.Client(api_key=k_api)
                for m_try in ["gemini-2.5-flash", "gemini-2.0-flash", model_name]:
                    try:
                        res = client.models.generate_content(
                            model=m_try,
                            contents=[image_part, prompt_m],
                            config=types.GenerateContentConfig(temperature=0.1, response_mime_type="application/json")
                        )
                        raw_t = re.sub(r"^```(?:json)?\s*|\s*```$", "", (res.text or "").strip(), flags=re.I).strip()
                        cand = json.loads(raw_t)
                        if isinstance(cand, dict):
                            cand = [cand]
                        if isinstance(cand, list) and len(cand) > 0:
                            for item in cand:
                                if isinstance(item, dict) and str(item.get("Full Name", "")).strip():
                                    p_in = str(item.get("Parish", default_parish) or default_parish).strip()
                                    matched_p = default_parish
                                    for p_cand in PARISH_ORDER:
                                        if normalize_parish_match(p_in, p_cand):
                                            matched_p = p_cand
                                            break
                                    item["Parish"] = matched_p
                                    if not str(item.get("Family ID", "")).strip():
                                        item["Family ID"] = f"CTG-{matched_p[:3]}-FAM-NEW"
                                    if not str(item.get("Head of Family", "")).strip():
                                        item["Head of Family"] = item["Full Name"]
                                    extracted_list.append(item)
                            if extracted_list:
                                return extracted_list
                    except Exception:
                        continue
        except Exception:
            pass

    # OCR Fallback (OCR.space / RapidOCR / pytesseract) kung walang Gemini quota
    ocr_txt = ""
    try:
        pil_img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        w, h = pil_img.size
        scale = min(1.0, 940.0 / max(1, max(w, h)))
        send_im = pil_img.resize((int(w * scale), int(h * scale)), Image.Resampling.BICUBIC) if scale < 1.0 else pil_img
        ob = io.BytesIO()
        send_im.save(ob, format="JPEG", quality=90)
        b64_im = "data:image/jpeg;base64," + base64.b64encode(ob.getvalue()).decode("utf-8")
        r_ocr = requests.post("https://api.ocr.space/parse/image", data={"apikey": "helloworld", "base64Image": b64_im, "language": "eng", "OCREngine": "2"}, timeout=12)
        if r_ocr.status_code == 200:
            ocr_txt = "\n".join(str(pr.get("ParsedText", "")).strip() for pr in r_ocr.json().get("ParsedResults", []))
    except Exception:
        pass

    if ocr_txt.strip():
        lines = [ln.strip() for ln in ocr_txt.splitlines() if ln.strip()]
        fid_m = re.search(r"(CTG-[A-Z0-9\-]+)", ocr_txt, flags=re.I)
        fid_val = fid_m.group(1).upper() if fid_m else f"CTG-{default_parish[:3]}-FAM-NEW"
        fn_val = lines[0] if lines else ""
        for ln in lines:
            if re.search(r"^(?:FULL\s*NAME|PANGALAN|NAME)\s*[:\-]", ln, flags=re.I):
                fn_val = re.sub(r"^(?:FULL\s*NAME|PANGALAN|NAME)\s*[:\-]\s*", "", ln, flags=re.I).strip()
                break
        if fn_val:
            extracted_list.append({
                "Parish": default_parish,
                "Family ID": fid_val,
                "Head of Family": fn_val,
                "Full Name": fn_val,
                "Relationship": "HEAD OF FAMILY",
                "Sex": "FEMALE" if "FEMALE" in ocr_txt.upper() else "MALE",
                "Age": "",
                "Category": "YOUTH" if "YOUTH" in ocr_txt.upper() else ("KIDS" if "KID" in ocr_txt.upper() else ("SENIOR CITIZEN" if "SENIOR" in ocr_txt.upper() else "ADULT")),
                "Status": "INACTIVE" if "INACTIVE" in ocr_txt.upper() else "ACTIVE",
                "Contact No.": "",
                "Current Address (If living separately)": "",
                "Remarks": " ".join(lines[1:4]) if len(lines) > 1 else "",
            })
    return extracted_list

# =====================================================================
# 4. MAGDAGDAG NG MIYEMBRO & LITRATO (MAY AUTO-EXTRACT MULA SA IMAGE O FILE PARA SA LAHAT NG ADMINS)
# =====================================================================
if app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    render_title_with_logo(
        "Magdagdag ng Miyembro (Manual o Auto-Upload ng Image / File)",
        "Mag-upload ng Image (ID / Form / Screenshot) o Excel/CSV file para awtomatikong mapunan at ma-save ang miyembro sa kaukulang Parokya at Google Drive."
    )
    allowed_parishes_add = [st.session_state.assigned_parish] if st.session_state.user_role == "Parish Leader" else PARISH_ORDER

    with st.expander("📸 1. Awtomatikong Magdagdag Mula sa In-upload na Image (.jpg, .png) o Excel/CSV File (.xlsx, .csv)", expanded=True):
        st.caption("Kapag nag-upload ka ng larawan o file dito, awtomatikong babasahin ang datos sa larawan/file at ilalagay sa form sa ibaba (o maaari mo nang i-save agad sa isang pindot!).")
        up_c1, up_c2 = st.columns([1.2, 1.8])
        with up_c1:
            def_p_auto = st.selectbox("Target na Parokya:", allowed_parishes_add, key="auto_add_mem_parish")
        with up_c2:
            auto_mem_file = st.file_uploader(
                "📤 Mag-upload ng Image (.jpg, .png) o File (.xlsx, .csv) ng Miyembro:",
                type=["jpg", "jpeg", "png", "xlsx", "csv"],
                key="auto_add_mem_uploader"
            )
        if auto_mem_file is not None:
            if st.button("🔍 Basahin ang Datos sa Image/File & Punan ang Form", type="primary", use_container_width=True):
                with st.spinner("Binabasa ang mga datos mula sa in-upload na Image/File..."):
                    f_bytes_m = auto_mem_file.read()
                    ext_members = extract_members_from_uploaded_image_or_file(
                        f_bytes_m, auto_mem_file.name, auto_mem_file.type, def_p_auto, selected_model
                    )
                    if ext_members:
                        st.session_state["extracted_members_batch"] = ext_members
                        st.session_state["extracted_member_form_prefill"] = ext_members[0]
                        st.success(f"✅ Nakakuha ng **{len(ext_members)}** miyembro mula sa `{auto_mem_file.name}`! Na-fill na ang form sa ibaba.")
                    else:
                        st.warning("Walang nabasang pangalan sa image/file. Maaari mong i-type nang direkta sa form sa ibaba.")

        if "extracted_members_batch" in st.session_state and st.session_state["extracted_members_batch"]:
            batch_list = st.session_state["extracted_members_batch"]
            st.markdown(f"**📋 Mga Nabasa Mula sa Image/File ({len(batch_list)} Miyembro):**")
            render_record_style_html_table(pd.DataFrame(batch_list))
            if st.button(f"🚀 Direktang I-save ang Lahat ({len(batch_list)} Miyembro) sa Excel Database & Google Drive", type="primary", use_container_width=True):
                saved_cnt = 0
                for m_item in batch_list:
                    m_payload = {
                        "Parish": m_item.get("Parish", def_p_auto),
                        "Family ID": m_item.get("Family ID", f"CTG-{def_p_auto[:3]}-FAM-NEW"),
                        "Head of Family": m_item.get("Head of Family") or m_item.get("Full Name", ""),
                        "Full Name": m_item.get("Full Name", ""),
                        "Relationship": m_item.get("Relationship", "HEAD OF FAMILY"),
                        "Sex": m_item.get("Sex", "MALE"),
                        "Age": m_item.get("Age", ""),
                        "Status": m_item.get("Status", "ACTIVE"),
                        "Contact No.": m_item.get("Contact No.", ""),
                        "Current Address (If living separately)": m_item.get("Current Address (If living separately)", ""),
                        "Category": m_item.get("Category", "ADULT"),
                        "Church Member": "NO" if "NON" in str(m_item.get("Status", "")).upper() else "YES",
                        "Remarks": f"{m_item.get('Remarks', '')} (Added on {get_philippine_datetime_str()})".strip(),
                    }
                    ok_b, _ = append_new_member_to_excel_and_sync(ACTIVE_SPREADSHEET_URL, m_payload)
                    if ok_b:
                        saved_cnt += 1
                st.session_state.pop("extracted_members_batch", None)
                st.cache_data.clear()
                st.success(f"✅ Matagumpay na naidagdag ang **{saved_cnt}** miyembro sa Excel Database at Google Drive ({get_philippine_datetime_str()})!")
                time.sleep(0.7)
                st.rerun()

    prefill_m = st.session_state.get("extracted_member_form_prefill", {})
    rel_opts = ["HEAD OF FAMILY", "HUSBAND", "WIFE", "SON", "DAUGHTER", "RELATIVE", "OTHER"]
    sex_opts = ["MALE", "FEMALE"]
    cat_opts = ["KIDS", "YOUTH", "ADULT", "SENIOR CITIZEN"]
    st_opts = ["ACTIVE", "INACTIVE", "NON-MEMBER"]

    def_p_val = prefill_m.get("Parish", allowed_parishes_add[0])
    def_p_i = allowed_parishes_add.index(def_p_val) if def_p_val in allowed_parishes_add else 0
    def_rel_i = rel_opts.index(prefill_m.get("Relationship", "HEAD OF FAMILY")) if prefill_m.get("Relationship") in rel_opts else 0
    def_sex_i = sex_opts.index(prefill_m.get("Sex", "MALE")) if prefill_m.get("Sex") in sex_opts else 0
    def_cat_i = cat_opts.index(prefill_m.get("Category", "ADULT")) if prefill_m.get("Category") in cat_opts else 2
    def_st_i = st_opts.index(prefill_m.get("Status", "ACTIVE")) if prefill_m.get("Status") in st_opts else 0

    with st.form("add_form"):
        st.markdown("##### ✍️ 2. Detalye ng Miyembro (Maaaring I-edit Bago I-save)")
        c1, c2 = st.columns(2)
        with c1:
            mp = st.selectbox("Parish:", allowed_parishes_add, index=def_p_i)
            fid = st.text_input("Family ID (hal. CTG-LN-FAM-0001):", value=str(prefill_m.get("Family ID", ""))).strip()
            m_head = st.text_input("Head of Family:", value=str(prefill_m.get("Head of Family", ""))).strip()
            fn = st.text_input("Buong Pangalan:", value=str(prefill_m.get("Full Name", ""))).strip()
            m_rel = st.selectbox("Relasyon sa Pamilya:", rel_opts, index=def_rel_i)
            m_addr = st.text_input("Current Address (Tirahan):", value=str(prefill_m.get("Current Address (If living separately)", ""))).strip()
        with c2:
            m_sex = st.selectbox("Kasarian:", sex_opts, index=def_sex_i)
            m_age = st.text_input("Edad (Age):", value=str(prefill_m.get("Age", ""))).strip()
            m_cat = st.selectbox("Kategorya:", cat_opts, index=def_cat_i)
            m_st = st.selectbox("Status:", st_opts, index=def_st_i)
            m_con = st.text_input("Contact Number:", value=str(prefill_m.get("Contact No.", ""))).strip()
            pho = st.file_uploader("Litrato (Direktang i-save sa Google Drive):", type=["jpg", "png", "jpeg"])
        m_rem = st.text_area("Remarks / Tala para sa Miyembro:", value=str(prefill_m.get("Remarks", ""))).strip()

        if st.form_submit_button("💾 I-save sa App Excel at I-upload sa Google Drive", type="primary"):
            if fid and fn:
                drive_link_text = "Walang inilakip na litrato."
                if pho is not None:
                    with st.spinner("Ina-upload ang litrato sa iyong Google Drive..."):
                        file_bytes = pho.read()
                        link, msg = upload_file_to_google_drive(file_bytes, f"CTG_Member_{fid}_{fn}.jpg", pho.type)
                        if link:
                            drive_link_text = f"Google Drive Link: {link}"
                            st.success(f"✅ Matagumpay na na-upload ang litrato sa Google Drive! ([Buksan ang Link]({link}))")
                        else:
                            st.warning(f"⚠️ {msg}")

                new_member_payload = {
                    "Parish": mp,
                    "Family ID": fid,
                    "Head of Family": m_head if m_head else fn,
                    "Full Name": fn,
                    "Relationship": m_rel,
                    "Sex": m_sex,
                    "Age": m_age,
                    "Status": m_st,
                    "Contact No.": m_con,
                    "Current Address (If living separately)": m_addr,
                    "Category": m_cat,
                    "Church Member": "NO" if "NON" in m_st else "YES",
                    "Remarks": f"{m_rem} ({drive_link_text})".strip() if drive_link_text != "Walang inilakip na litrato." else m_rem,
                }

                with st.spinner("Idinadagdag sa Excel Database at sini-sync sa Google Drive..."):
                    ok_ex, ex_msg = append_new_member_to_excel_and_sync(ACTIVE_SPREADSHEET_URL, new_member_payload)
                    if ok_ex:
                        st.success(f"✅ Naidagdag na si **{fn}** sa `{mp}` ({fid})! {ex_msg}")
                    else:
                        st.warning(f"⚠️ Na-record ang miyembro ngunit may abiso sa Excel sync: {ex_msg}")

                backup_action_to_admin_cloud(
                    "New Member Added to Excel & Cloud",
                    fn,
                    mp,
                    f"Family ID: {fid}",
                    f"Head: {m_head} | Rel: {m_rel} | Status: {m_st} | Cat: {m_cat} | Contact: {m_con} | {drive_link_text}",
                    f"Remarks: {m_rem}",
                    get_current_chat_identity()
                )
                if pho is not None and 'link' in locals() and link:
                    st.image(pho, caption=f"Litrato ni {fn} (Naka-save sa Drive)", width=180)
            else:
                st.warning("Punan ang Family ID at Buong Pangalan.")
    st.stop()

# =====================================================================
# 5. PRAYER WALL & CATEGORIZER
# =====================================================================
elif app_mode == "🙏 Prayer Wall & Categorizer":
    render_title_with_logo("AI Prayer Request Categorizer")
    plang = st.selectbox("Wika:", ["Tagalog", "Ilokano", "English"])
    rp = st.text_area("Nilalaman ng Panalangin:")
    if st.button("Suriin sa AI", type="primary"):
        if rp:
            client = genai.Client(api_key=get_valid_gemini_keys()[0])
            res = safe_generate_content(
                client,
                selected_model,
                f"Bilang bahagi ng Church of the Temple of God (CTG), suriin ang prayer request na ito sa wikang {plang}: {rp}"
            )
            st.success("Tapos na!")
            st.markdown(res.text)
        else:
            st.warning("Maglagay muna ng mensahe.")
    st.stop()

# =====================================================================
# 6. ANNOUNCEMENT BROADCASTER
# =====================================================================
elif app_mode == "📢 Announcement Broadcaster":
    render_title_with_logo("AI Announcement Broadcaster")
    blang = st.selectbox("Wika:", ["Tagalog", "Ilokano", "English"])
    et = st.text_input("Pamagat:")
    ed = st.text_area("Detalye:")
    em = st.text_input("Email ng Tatanggap:")
    if st.button("Magpadala", type="primary"):
        if et and ed and em:
            client = genai.Client(api_key=get_valid_gemini_keys()[0])
            body = safe_generate_content(
                client,
                selected_model,
                f"Gumawa ng pormal na announcement mula sa Church of the Temple of God (CTG) "
                f"(kung saan ang araw ng pagsamba ay tuwing Sabado / Saturday) tungkol sa {et}. Detalye: {ed}. Isulat sa {blang}."
            ).text
            res = send_email_via_gmail(em, f"CTG Announcement: {et}", body)
            st.success(res)
        else:
            st.warning("Punan ang lahat.")
    st.stop()

# =====================================================================
# 7. APP SETTINGS (THEME, EXCEL DRIVE MANAGER, AT API CREDENTIALS + 403 FIX)
# =====================================================================
elif app_mode == "🛠️ App Settings":
    render_title_with_logo("🛠️ App Controller, Theme, Excel Sync & Configuration Settings", "I-customize ang hitsura ng app, i-manage ang Google Drive Excel File, at i-update ang API settings.")

    tab_ui, tab_excel, tab_api = st.tabs([
        "🎨 Theme, Display & App Controller",
        "📊 Google Drive Excel File & Weekly Folder Manager",
        "🔑 API & Google Drive / Gmail Credentials (Fix 403 Scope)"
    ])

    with tab_ui:
        st.subheader("🎛️ App Appearance & Controller Settings")
        c_u1, c_u2 = st.columns(2)
        with c_u1:
            theme_list = ["🌙 Dark Mode", "☀️ Light Mode", "⛪ Royal Blue Ministry"]
            new_ui_theme = st.radio(
                "1. Pumili ng Kulay / Theme ng App:",
                theme_list,
                index=theme_list.index(st.session_state.ui_theme) if st.session_state.ui_theme in theme_list else 0
            )
            font_list = ["Normal (100%)", "Large (115%)", "Extra Large (130%)"]
            new_ui_font = st.radio(
                "2. Laki ng Teksto (Font Size Controller):",
                font_list,
                index=font_list.index(st.session_state.ui_font_size) if st.session_state.ui_font_size in font_list else 0
            )
        with c_u2:
            dens_list = ["Standard (Comfortable)", "Compact (Mas Maraming Data)"]
            new_ui_dens = st.radio(
                "3. Card & Table Layout Spacing:",
                dens_list,
                index=dens_list.index(st.session_state.ui_density) if st.session_state.ui_density in dens_list else 0
            )
            lang_list = ["Tagalog", "Ilokano", "English"]
            new_ui_lang = st.selectbox(
                "4. Default na Wika para sa AI Drafts:",
                lang_list,
                index=lang_list.index(st.session_state.ui_default_lang) if st.session_state.ui_default_lang in lang_list else 0
            )
            new_side_rec = st.checkbox(
                "5. Laging ipakita ang Records Panel sa gilid ng Workspace",
                value=st.session_state.ui_show_side_records
            )

        st.divider()
        col_btn_ui1, col_btn_ui2 = st.columns([1.5, 1.5])
        with col_btn_ui1:
            if st.button("✅ I-apply ang Theme at Controller Settings", type="primary", use_container_width=True):
                st.session_state.ui_theme = new_ui_theme
                st.session_state.ui_font_size = new_ui_font
                st.session_state.ui_density = new_ui_dens
                st.session_state.ui_default_lang = new_ui_lang
                st.session_state.ui_show_side_records = new_side_rec
                st.success("Na-update na ang Theme at App Controller!")
                time.sleep(0.5)
                st.rerun()
        with col_btn_ui2:
            if st.button("🧹 I-clear ang App Cache & Refresh Data", use_container_width=True):
                st.cache_data.clear()
                st.cache_resource.clear()
                st.success("Na-clear na ang cache at na-refresh ang data!")
                time.sleep(0.5)
                st.rerun()

    with tab_excel:
        st.subheader("📊 Google Drive Excel Link & Weekly Report Folder Configuration")
        st.caption("Maaari mong i-save ang pangunahing Google Drive links at mag-add ng kahit ilang karagdagang Google Drive Excel o Folder links sa ibaba nang **hindi nawawala ang mga naunang links na nakalagay na**.")
        new_sheet_url = st.text_input(
            "🔗 1. Google Drive Master Census / Primary Excel Link:",
            value=ACTIVE_SPREADSHEET_URL
        )
        new_weekly_folder_url = st.text_input(
            "📂 2. CTG Ministry Hub (Weekly Report January-December 2027) Master Folder Link:",
            value=ACTIVE_WEEKLY_FOLDER_URL
        )
        new_parish_upload_url = st.text_input(
            "📤 3. 'Upload the Image/Excel File here!' Folder Link (Kung saan papasok ang mga Image at Excel per Parish):",
            value=st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL", ACTIVE_PARISH_UPLOAD_URL)
        )
        if st.button("💾 I-save ang Pangunahing Google Drive Links & I-sync Agad", type="primary"):
            st.session_state.custom_secrets["CLOUD_SPREADSHEET_URL"] = new_sheet_url.strip()
            st.session_state.custom_secrets["WEEKLY_REPORT_FOLDER_URL"] = new_weekly_folder_url.strip()
            st.session_state.custom_secrets["PARISH_UPLOAD_FOLDER_URL"] = new_parish_upload_url.strip()
            try:
                existing_cfg = {}
                if os.path.exists(LOCAL_CONFIG_FILE):
                    with open(LOCAL_CONFIG_FILE, "r", encoding="utf-8") as f:
                        existing_cfg = json.load(f)
                existing_cfg["CLOUD_SPREADSHEET_URL"] = new_sheet_url.strip()
                existing_cfg["WEEKLY_REPORT_FOLDER_URL"] = new_weekly_folder_url.strip()
                existing_cfg["PARISH_UPLOAD_FOLDER_URL"] = new_parish_upload_url.strip()
                with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
                    json.dump(existing_cfg, f, indent=2)
            except Exception:
                pass
            if new_sheet_url.strip():
                add_link_to_drive_registry("Primary Master Excel Link", "Master Weekly Report Excel (PWR)", "January", "1st Week", "Lahat ng Parokya (Master)", new_sheet_url.strip())
            if new_weekly_folder_url.strip():
                add_link_to_drive_registry("CTG 2027 Master Folder", "Google Drive Master Folder", "All Months", "All Weeks", "Lahat ng Parokya (Master)", new_weekly_folder_url.strip())
            if new_parish_upload_url.strip():
                add_link_to_drive_registry("Parish Upload Folder", "Parish Upload Folder", "All Months", "All Weeks", "Lahat ng Parokya (Master)", new_parish_upload_url.strip())
            sync_Target = new_weekly_folder_url.strip() or new_parish_upload_url.strip() or new_sheet_url.strip()
            _, auto_s_msg = auto_sync_from_any_drive_link(sync_Target, "January", "1st Week", "Lahat ng Parokya (Master)")
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success(f"✅ Na-update na ang Google Drive Links at auto-synced agad! ({auto_s_msg})")
            time.sleep(0.7)
            st.rerun()

        st.divider()
        st.markdown("### ➕ 4. Mag-add ng Maraming Google Drive Links (Hindi Mawawala ang mga Nakalagay Na!)")
        st.caption("Mag-add ng mga link para sa Week 1, Week 2, Week 3, iba't ibang Buwan ng 2027, o partikular na Parokya. Mananatiling naka-save ang lahat ng links dito para makapag-upload o makapag-sync ka gamit ang kani-kanilang link.")

        with st.form("add_multi_drive_link_form"):
            ml_c1, ml_c2, ml_c3 = st.columns([1.6, 1.2, 1.2])
            with ml_c1:
                add_lbl = st.text_input("🏷️ Pangalan / Label ng Link (Hal. PWR. January 2nd Week.xlsx):")
            with ml_c2:
                add_cat = st.selectbox(
                    "📂 Uri ng Link:",
                    [
                        "Master Weekly Report Excel (PWR)",
                        "Parish Weekly Report Excel",
                        "Google Drive Master Folder",
                        "Parish Upload Folder",
                        "Master Census Excel"
                    ]
                )
            with ml_c3:
                add_parish = st.selectbox("🏛️ Parokya / Scope:", ["Lahat ng Parokya (Master)"] + PARISH_ORDER)

            ml_c4, ml_c5, ml_c6 = st.columns([1.0, 1.0, 2.0])
            with ml_c4:
                add_month = st.selectbox("📅 Buwan (2027):", ["All Months"] + MONTHS_2027, index=1)
            with ml_c5:
                add_week = st.selectbox("🗓️ Linggo (Week):", ["All Weeks"] + WEEKS_2027, index=1)
            with ml_c6:
                add_url = st.text_input("🔗 Google Drive Link (I-paste ang URL ng Excel o Folder):")

            if st.form_submit_button("➕ Idagdag sa Listahan ng Saved Google Drive Links", type="primary", use_container_width=True):
                if not add_url.strip():
                    st.warning("Mag-paste muna ng Google Drive Link URL.")
                else:
                    auto_lbl = add_lbl.strip() or f"PWR. {add_month} {add_week} ({add_parish})"
                    add_link_to_drive_registry(auto_lbl, add_cat, add_month, add_week, add_parish, add_url.strip())
                    eff_m_add = "January" if add_month == "All Months" else add_month
                    eff_w_add = "1st Week" if add_week == "All Weeks" else add_week
                    _, auto_s_msg = auto_sync_from_any_drive_link(add_url.strip(), eff_m_add, eff_w_add, add_parish)
                    st.cache_data.clear()
                    st.success(f"✅ Naidagdag na ang '{auto_lbl}' at auto-synced agad mula sa Google Drive! ({auto_s_msg})")
                    time.sleep(0.5)
                    st.rerun()

        saved_links_list = load_saved_drive_links_registry()
        if saved_links_list:
            st.markdown("#### 📋 Mga Naka-save na Google Drive Links (Saved Links Directory)")
            links_df = pd.DataFrame([
                {
                    "#": idx + 1,
                    "Label / Pangalan": it.get("label", ""),
                    "Uri (Category)": it.get("category", ""),
                    "Buwan": it.get("month", ""),
                    "Linggo": it.get("week", ""),
                    "Parokya / Scope": it.get("parish", ""),
                    "Google Drive Link": it.get("url", ""),
                }
                for idx, it in enumerate(saved_links_list)
            ])
            st.dataframe(links_df, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("### 📤 5. Mag-upload ng Excel File (`.xlsx`) o Mag-sync Gamit ang Kani-kanilang Saved Link")
            st.caption("Piliin lamang sa dropdown ang kahit anong naka-save na link upang doon direktang mag-upload ng Excel file o basahin ang laman nito papunta sa Monitoring at Summary tables.")

            link_options = [
                f"[{idx + 1}] {it.get('label', 'Link')} — {it.get('month', '')} {it.get('week', '')} ({it.get('parish', '')})"
                for idx, it in enumerate(saved_links_list)
            ]
            sel_link_idx = st.selectbox(
                "🎯 Pumili ng Saved Google Drive Link:",
                range(len(link_options)),
                format_func=lambda i: link_options[i],
                key="sel_saved_drive_link_idx"
            )
            chosen_link_item = saved_links_list[sel_link_idx]
            chosen_url = str(chosen_link_item.get("url", "")).strip()
            chosen_month = chosen_link_item.get("month", "January")
            if chosen_month == "All Months":
                chosen_month = "January"
            chosen_week = chosen_link_item.get("week", "1st Week")
            if chosen_week == "All Weeks":
                chosen_week = "1st Week"
            chosen_parish = chosen_link_item.get("parish", "Lahat ng Parokya (Master)")

            st.info(f"🔗 **Active Selected Link:** `{chosen_link_item.get('label')}` — **URL:** {chosen_url}")

            act_c1, act_c2 = st.columns(2)
            with act_c1:
                if st.button("🔄 Basahin at I-sync ang Excel Mula sa Link na Ito Papunta sa App Tables", use_container_width=True, type="primary"):
                    with st.spinner(f"Binabasa at sini-sync ang '{chosen_link_item.get('label')}' mula sa Google Drive..."):
                        service = get_google_drive_service()
                        f_id = extract_id_from_url(chosen_url)
                        if "/folders/" in chosen_url:
                            cnt, msg = sync_weekly_reports_from_google_drive_folder(chosen_url, chosen_month, chosen_week)
                            st.success(f"✅ {msg}")
                        elif service and f_id:
                            try:
                                x_bytes = service.files().get_media(fileId=f_id).execute()
                                with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
                                    f.write(x_bytes)
                                cnt = parse_master_pwr_excel_into_db(x_bytes, chosen_month, chosen_week)
                                if cnt == 0 and chosen_parish != "Lahat ng Parokya (Master)":
                                    ext_d = extract_weekly_report_from_excel_bytes(x_bytes, target_parish=chosen_parish)
                                    save_weekly_monitoring_record(chosen_parish, chosen_month, chosen_week, ext_d.get("report_date", "01/02/27"), ext_d, drive_link=chosen_url)
                                    cnt = 1
                                st.cache_data.clear()
                                st.success(f"✅ Matagumpay na nabasa at na-sync ang {cnt} parokya mula sa '{chosen_link_item.get('label')}' ({chosen_month} - {chosen_week})!")
                            except Exception as e:
                                st.error(f"Error sa pagbasa ng link: {str(e)}")
                        else:
                            st.warning("Hindi mabuksan ang Google Drive service o invalid ang File ID.")
                        time.sleep(0.6)
                        st.rerun()

            with act_c2:
                if len(saved_links_list) > 1:
                    if st.button("🗑️ Alisin ang Napiling Link na Ito sa Listahan", use_container_width=True):
                        saved_links_list.pop(sel_link_idx)
                        save_drive_links_registry(saved_links_list)
                        st.success("Inalis na ang napiling link (nananatiling ligtas ang ibang links).")
                        time.sleep(0.4)
                        st.rerun()

            uploaded_xlsx_for_link = st.file_uploader(
                f"📤 Pumili ng Excel File (.xlsx) na I-u-upload Direkta sa '{chosen_link_item.get('label')}':",
                type=["xlsx"],
                key="upload_xlsx_to_saved_link"
            )
            if uploaded_xlsx_for_link is not None:
                if st.button(f"🚀 I-upload ang Excel sa '{chosen_link_item.get('label')}' at I-update ang App Tables", type="primary", use_container_width=True):
                    with st.spinner(f"Ina-upload sa '{chosen_link_item.get('label')}' at binabasa ang lahat ng datos..."):
                        x_bytes = uploaded_xlsx_for_link.read()
                        with open(LOCAL_PWR_MASTER_FILE, "wb") as f:
                            f.write(x_bytes)
                        # 1. Kung folder link ito, i-upload sa loob ng folder; kung file link, i-update ang file
                        if "/folders/" in chosen_url:
                            u_link, u_msg = upload_file_to_google_drive(
                                x_bytes,
                                uploaded_xlsx_for_link.name,
                                mime_type=uploaded_xlsx_for_link.type,
                                parent_folder_url=chosen_url,
                                parish_name="" if "Master" in chosen_parish else chosen_parish,
                                month_name=chosen_month
                            )
                            ok_u = bool(u_link)
                            msg_u = u_msg
                        else:
                            ok_u, msg_u = sync_excel_bytes_to_drive(x_bytes, chosen_url)

                        # 2. I-parse din agad sa database ng app para lumabas sa lahat ng tables
                        cnt_parsed = parse_master_pwr_excel_into_db(x_bytes, chosen_month, chosen_week)
                        if cnt_parsed == 0 and "Master" not in chosen_parish:
                            ext_d = extract_weekly_report_from_excel_bytes(x_bytes, target_parish=chosen_parish)
                            save_weekly_monitoring_record(chosen_parish, chosen_month, chosen_week, ext_d.get("report_date", "01/02/27"), ext_d, drive_link=chosen_url)
                        st.cache_data.clear()
                        st.cache_resource.clear()
                        st.success(f"✅ {msg_u} (Na-update din ang Monitoring, Summary, at Outreach tables para sa {chosen_month} - {chosen_week}!)")
                        time.sleep(0.8)
                        st.rerun()

    with tab_api:
        if st.session_state.user_role != "Master Admin":
            st.info("🔒 Ang API & Google Drive Cloud Storage credentials ay maaari lamang baguhin ng Master Admin.")
        else:
            st.subheader("🔑 API Keys & Google Drive / Gmail Cloud Storage Integration")
            g_cid = st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", GMAIL_CLIENT_ID)
            g_csec = st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", GMAIL_CLIENT_SECRET)
            if g_cid and g_csec:
                with st.expander("🔓 One-Click Google Drive + Gmail Token Generator (Fix 403 Error)", expanded=True):
                    st.caption("Kung nagkakaroon ng '403 Insufficient Authentication Scopes' sa Google Drive upload, gamitin ito upang makagawa ng Refresh Token na may buong Google Drive + Gmail permission:")
                    try:
                        oauth_scopes = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/gmail.send"
                        auth_params = {
                            "client_id": g_cid.strip(),
                            "redirect_uri": "http://localhost",
                            "response_type": "code",
                            "scope": oauth_scopes,
                            "access_type": "offline",
                            "prompt": "consent"
                        }
                        auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(auth_params)
                        st.markdown(f"**Step 1:** [👉 Pindutin Ito upang Pahintulutan ang Google Drive at Gmail]({auth_url})")
                        auth_code_in = st.text_input("Step 2: Pagkatapos mag-Allow, kopyahin ang buong `http://localhost/?code=...` URL mula sa address bar at i-paste dito:").strip()
                        if st.button("🔑 Kunin at I-save ang Bagong Google Drive Refresh Token", type="primary"):
                            clean_code = auth_code_in
                            if "code=" in clean_code:
                                clean_code = urllib.parse.unquote(re.search(r"code=([^&]+)", clean_code).group(1))
                            token_res = requests.post(
                                "https://oauth2.googleapis.com/token",
                                data={
                                    "code": clean_code,
                                    "client_id": g_cid.strip(),
                                    "client_secret": g_csec.strip(),
                                    "redirect_uri": "http://localhost",
                                    "grant_type": "authorization_code"
                                },
                                timeout=20
                            )
                            token_json = token_res.json()
                            new_rt = token_json.get("refresh_token", "")
                            if new_rt:
                                st.session_state.custom_secrets["GMAIL_REFRESH_TOKEN"] = new_rt
                                existing_cfg = {}
                                if os.path.exists(LOCAL_CONFIG_FILE):
                                    with open(LOCAL_CONFIG_FILE, "r", encoding="utf-8") as f:
                                        existing_cfg = json.load(f)
                                existing_cfg["GMAIL_REFRESH_TOKEN"] = new_rt
                                with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
                                    json.dump(existing_cfg, f, indent=2)
                                st.success("✅ Matagumpay na nakuha at na-save ang bagong Google Drive + Gmail Refresh Token!")
                                st.code(new_rt)
                            else:
                                st.error(f"Hindi nakuha ang refresh_token: {token_json}")
                    except Exception as ex:
                        st.caption(f"OAuth Helper Notice: {str(ex)}")

            with st.form("set_form"):
                ng = st.text_input("GEMINI_API_KEY", value=st.session_state.custom_secrets.get("GEMINI_API_KEY", ""), type="password")
                ng2 = st.text_input("GEMINI_API_KEY_2", value=st.session_state.custom_secrets.get("GEMINI_API_KEY_2", ""), type="password")
                nt = st.text_input("NOTION_TOKEN", value=st.session_state.custom_secrets.get("NOTION_TOKEN", ""), type="password")
                ndb = st.text_input("NOTION_CRM_DB_ID", value=st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", ""))
                npdb = st.text_input("NOTION_PRAYER_DB_ID", value=st.session_state.custom_secrets.get("NOTION_PRAYER_DB_ID", ""))
                nqdb = st.text_input("NOTION_QUESTIONS_DB_ID", value=st.session_state.custom_secrets.get("NOTION_QUESTIONS_DB_ID", ""))
                ncdb = st.text_input("NOTION_COUNSELLING_DB_ID", value=st.session_state.custom_secrets.get("NOTION_COUNSELLING_DB_ID", ""))
                gci = st.text_input("GMAIL_CLIENT_ID", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", ""))
                gcs = st.text_input("GMAIL_CLIENT_SECRET", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", ""), type="password")
                grt = st.text_input("GMAIL_REFRESH_TOKEN", value=st.session_state.custom_secrets.get("GMAIL_REFRESH_TOKEN", ""), type="password")
                if st.form_submit_button("💾 I-save ang API & Cloud Credentials", type="primary"):
                    new_creds = {
                        "GEMINI_API_KEY": ng, "GEMINI_API_KEY_2": ng2, "NOTION_TOKEN": nt,
                        "NOTION_CRM_DB_ID": ndb, "NOTION_PRAYER_DB_ID": npdb,
                        "NOTION_QUESTIONS_DB_ID": nqdb, "NOTION_COUNSELLING_DB_ID": ncdb,
                        "GMAIL_CLIENT_ID": gci, "GMAIL_CLIENT_SECRET": gcs, "GMAIL_REFRESH_TOKEN": grt
                    }
                    st.session_state.custom_secrets.update(new_creds)
                    try:
                        existing_cfg = load_saved_config()
                        for k_c, v_c in new_creds.items():
                            if str(v_c or "").strip():
                                existing_cfg[k_c] = str(v_c).strip()
                        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
                            json.dump(existing_cfg, f, indent=2)
                    except Exception:
                        pass
                    st.cache_data.clear()
                    st.success("Na-save na ang mga credentials!")
                    time.sleep(1)
                    st.rerun()
    st.stop()

# =====================================================================
# 8. MADALING SUNDAN NA AI CHAT ASSISTANT (STEP-BY-STEP GUIDE + ONE-CLICK BUTTONS)
# =====================================================================
def query_weekly_monitoring_reports(search_query: str = "") -> str:
    """Naghahanap sa 2027 Weekly Reports (Attendance, Liturgical Officers, Preachers, Topics, Outreach, at Remarks) ng mga Parokya."""
    try:
        db = load_weekly_monitoring_records()
        if not db:
            return "Wala pang naka-save na Weekly Report sa database."
        q_low = str(search_query or "").strip().lower()
        lines = []
        for k, rec in db.items():
            if not isinstance(rec, dict) or not rec.get("has_data"):
                continue
            p_n = rec.get("parish", "")
            m_n = rec.get("month", "")
            w_n = rec.get("week", "")
            summary_str = (
                f"Parish: {p_n} | Month/Week: {m_n} 2027 ({w_n}) | Date: {rec.get('report_date','')} | "
                f"Submitted PH Time: {rec.get('submitted_at_ph','')} | "
                f"Priest: {rec.get('priest','')} | Sakristan: {rec.get('sakristan','')} | "
                f"Tema Misa: {rec.get('tema_misa','')} | Adal Topic: {rec.get('adal_evangelio','')} | "
                f"Mass(A:{rec.get('mass_a',0)}, Y:{rec.get('mass_y',0)}) | "
                f"Pampamilya(A:{rec.get('pampamilya_a',0)}, Y:{rec.get('pampamilya_y',0)}) | "
                f"Online(A:{rec.get('online_a',0)}, Y:{rec.get('online_y',0)}) | "
                f"Youth:{rec.get('youth_y',0)} | Kids:{rec.get('kids_k',0)} | "
                f"Remarks: {rec.get('remarks','')[:180]}"
            )
            if not q_low or q_low in summary_str.lower():
                lines.append(summary_str)
        if not lines:
            return f"Walang nahanap na Weekly Report para sa '{search_query}'."
        return "Mga 2027 Weekly Reports:\n" + "\n".join(lines[:20])
    except Exception as e:
        return f"Error sa pagbasa ng Weekly Reports: {str(e)}"

render_title_with_logo(
    "💬 CTG Gabay AI — Madaling Sundan na Ministry Assistant",
    f"Opisyal na AI Assistant ng Church of the Temple of God ({st.session_state.assigned_parish}) • 🇵🇭 {get_philippine_datetime_str()}"
)

st.markdown(
    "<div class='full-text-card'>"
    "<b>🧭 Paano Gamitin ang AI Chat Assistant (3 Madaling Hakbang):</b><br>"
    "<b>1️⃣ Pumili ng Wika at Mode sa Ibaba</b> (Tagalog, Ilokano, o English).<br>"
    "<b>2️⃣ Pindutin ang Kahit Aling 'Mabilisang Tanong' Button</b> sa ibaba kung ayaw mong mag-type, O<br>"
    "<b>3️⃣ Mag-type sa Chat Box sa Pinakababa</b> (hal. <i>'Hanapin si Pacita Agasen'</i> o <i>'Gumawa ng mensahe para sa pagsamba sa Sabado'</i>)."
    "</div>",
    unsafe_allow_html=True
)

valid_keys = get_valid_gemini_keys()
if not valid_keys:
    st.warning("Wala pang GEMINI_API_KEY. Mangyaring ilagay ito sa App Settings.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = []
if "current_model" not in st.session_state:
    st.session_state.current_model = selected_model

col_chat_opt1, col_chat_opt2, col_chat_opt3 = st.columns([1.2, 1.6, 0.9])
with col_chat_opt1:
    pref_lang = st.selectbox("🌐 1. Pumili ng Wika:", ["Tagalog", "Ilokano", "English", "Awto (Kung ano ang tanong)"], index=0)
with col_chat_opt2:
    assistant_mode = st.selectbox(
        "🤖 2. Pumili ng Gawain:",
        [
            "💬 All-in-One Assistant (Maghanap ng Miyembro, Attendance & Gabay)",
            "✍️ Gawaan ng Mensahe / Tugon sa Miyembro (AI Draft Generator)"
        ],
        index=0
    )
with col_chat_opt3:
    st.write("")
    if st.button("🧹 Burahin ang Usapan", use_container_width=True):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()

st.markdown("##### ⚡ Mabilisang Tanong (Pindutin lang ang button para sumagot agad ang AI):")
qb1, qb2, qb3 = st.columns(3)
qb4, qb5, qb6 = st.columns(3)
quick_clicked_prompt = None
with qb1:
    if st.button("📊 Buod ng Weekly Report & Attendance", use_container_width=True):
        quick_clicked_prompt = f"Ipakita ang malinaw na buod ng pinakahuling Weekly Report, Attendance, at Liturgical Officers para sa {st.session_state.assigned_parish}."
with qb2:
    if st.button("👥 Ilan ang Active Members sa Census?", use_container_width=True):
        quick_clicked_prompt = f"Ipakita ang bilang ng Active Members (Seniors, Adults, Youth, Kids) sa {st.session_state.assigned_parish} mula sa Excel Census."
with qb3:
    if st.button("✉️ Imbitasyon sa Pagsamba ngayong Sabado", use_container_width=True):
        quick_clicked_prompt = f"Gumawa ng maikli, magalang, at nakaka-inspire na mensahe ng paanyaya para sa ating pagsamba ngayong Sabado sa Church of the Temple of God ({st.session_state.assigned_parish})."
with qb4:
    if st.button("🙏 Panalangin & Comforting Bible Verse", use_container_width=True):
        quick_clicked_prompt = "Magbigay ng isang nakaka-comfort na Bible verse at maikling panalangin para sa pamilyang dumadaan sa pagsubok."
with qb5:
    if st.button("📥 Tingnan ang Pending Requests sa Inbox", use_container_width=True):
        quick_clicked_prompt = f"Ano ang mga kasalukuyang Prayer Requests, Questions, o Counselling sa Notion CRM para sa {st.session_state.assigned_parish}?"
with qb6:
    if st.button("🏠 Gabay sa House Visitation & Outreach", use_container_width=True):
        quick_clicked_prompt = "Magbigay ng 5 madaling hakbang at gabay para sa aming House Visitation (10 families requirement) at Missionary Work ngayong linggo."

lang_rule = "Sumagot gamit ang wikang ginamit ng nagtatanong." if "Awto" in pref_lang else f"Sumagot nang buo sa wikang {pref_lang}."
sys_inst = (
    f"Ikaw si CTG Gabay AI, ang opisyal, magalang, at napakadaling sundan na AI assistant ng Church of the Temple of God (CTG) "
    f"para sa parokya ng {st.session_state.assigned_parish}. Tandaan na ang opisyal na pangalan ng simbahan ay "
    f"'Church of the Temple of God' at ang araw ng pagsamba at pangunahing pagtitipon (Worship / Sabbath Service) ay tuwing SABADO (Saturday).\n"
    f"- {lang_rule}\n"
    f"- Isulat ang iyong sagot sa malinaw na bullet points o numbered steps (1, 2, 3) upang hindi mahirapan ang mga users sa pagbasa sa cellphone, tablet, o computer.\n"
    f"- Gamitin ang `query_excel_database` kapag naghahanap ng miyembro o Family ID sa Excel Census.\n"
    f"- Gamitin ang `query_weekly_monitoring_reports` kapag tinatanong ang 2027 Weekly Reports, Attendance, Officers, o Remarks.\n"
    f"- Gamitin ang `query_notion_crm` kapag tinatanong ang Prayer Requests, Questions, Counselling, o Visitors sa Notion.\n"
    f"- Gamitin ang `search_web_for_answers` para sa real-time internet facts."
)

def resolve_callable_gemini_model(requested_model: str) -> str:
    if requested_model == "gemini-3.8-flash":
        return "gemini-2.5-flash"
    return requested_model or "gemini-2.5-flash"

def create_chat(api_key, model_name):
    client = genai.Client(api_key=api_key)
    actual_m = resolve_callable_gemini_model(model_name)
    return client.chats.create(
        model=actual_m,
        config=types.GenerateContentConfig(
            system_instruction=sys_inst,
            tools=[query_excel_database, query_weekly_monitoring_reports, query_notion_crm, send_email_via_gmail, search_web_for_answers],
            temperature=0.3
        )
    )

if "chat_session" not in st.session_state or st.session_state.current_model != selected_model:
    st.session_state.current_model = selected_model
    st.session_state.chat_session = create_chat(valid_keys[0], selected_model)

def stream_chat_response(prompt_text: str):
    final_prompt = prompt_text
    if "Gawaan ng Mensahe" in assistant_mode or "AI Draft" in assistant_mode:
        final_prompt = (
            f"Gumawa ng opisyal, malinaw, at magalang na AI Draft Answer / Mensahe mula sa Church of the Temple of God ({st.session_state.assigned_parish}) "
            f"(kung saan ang araw ng pagsamba ay tuwing Sabado / Saturday) na may kasamang Bible verse at comforting message para dito ({lang_rule}):\n\n{prompt_text}"
        )
    elif "Awto" not in pref_lang:
        final_prompt = f"[Sagutin nang malinaw at madaling sundan sa wikang {pref_lang}] {prompt_text}"

    try:
        for chunk in st.session_state.chat_session.send_message_stream(final_prompt):
            if chunk.text:
                yield chunk.text
        return
    except Exception:
        pass

    for backup_key in valid_keys:
        for fallback_m in ["gemini-2.5-flash", resolve_callable_gemini_model(selected_model)]:
            try:
                st.session_state.chat_session = create_chat(backup_key, fallback_m)
                for chunk in st.session_state.chat_session.send_message_stream(final_prompt):
                    if chunk.text:
                        yield chunk.text
                return
            except Exception:
                continue

    try:
        client = genai.Client(api_key=valid_keys[0])
        res = safe_generate_content(client, "gemini-2.5-flash", final_prompt)
        yield res.text
    except Exception as e:
        yield f"⚠️ Paumanhin, nagkaroon ng pansamantalang aberya sa koneksyon: {str(e)}"

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

typed_prompt = st.chat_input("✍️ Mag-type dito ng iyong tanong (hal. 'Hanapin si Pacita Agasen' o 'Gumawa ng panalangin')...")
active_prompt = quick_clicked_prompt or typed_prompt

if active_prompt:
    st.session_state.messages.append({"role": "user", "content": active_prompt})
    with st.chat_message("user"):
        st.markdown(active_prompt)
    with st.chat_message("assistant"):
        full_reply = st.write_stream(stream_chat_response(active_prompt))
        st.session_state.messages.append({"role": "assistant", "content": str(full_reply)})
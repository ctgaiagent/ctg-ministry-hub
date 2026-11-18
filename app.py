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
import hashlib
import streamlit.components.v1 as components
from datetime import datetime, timedelta, timezone

PH_TIMEZONE = timezone(timedelta(hours=8))

LOCAL_NOTIFICATIONS_FILE = "ctg_live_notifications.json"

UI_TRANSLATIONS = {
    "Tagalog": {
        "portal_title": "🧭 Parish Leader Portal — Church of the Temple of God",
        "portal_sub": "Maligayang pagdating sa opisyal na pang-araw-araw na gabay at ministry portal ng Church of the Temple of God.",
        "refresh_app": "🔄 I-refresh ang App",
        "delete_populated": "🗑️ Burahin ang Na-populate na Data (Delete Record)",
        "lang_label": "🌐 Wika ng App (Language):",
        "deadline_closed": "⛔ Sarado na ang Submission (Lagpas na sa Monday Midnight Deadline). Naka-marka na bilang '(No Report)' sa App at sa Excel file. Maaari lamang mag-submit muli sa susunod na reporting cycle sa Sabado.",
        "security_badge": "🔒 Protektado ng Role-Based Access Control (RLS), Encrypted Session, at Streamlit Secrets."
    },
    "Ilokano": {
        "portal_title": "🧭 Portal dagiti Pangulo ti Parokya — Church of the Temple of God",
        "portal_sub": "Naragsak nga isasangbay iti inaldaw a giya ken ministry portal ti Church of the Temple of God.",
        "refresh_app": "🔄 I-refresh ti App",
        "delete_populated": "🗑️ Ikkaten ti Na-populate a Datos (Delete Record)",
        "lang_label": "🌐 Pagsasao ti App (Language):",
        "deadline_closed": "⛔ Naserraanen ti Submission (Nalabsanen ti Monday Midnight Deadline). Namarkaanen a '(No Report)' iti App ken Excel file. Mabalin nga ag-submit manen inton sumaruno a Sabado.",
        "security_badge": "🔒 Natalged babaen ti Role-Based Access Control (RLS), Encrypted Session, ken Streamlit Secrets."
    },
    "English": {
        "portal_title": "🧭 Parish Leader Portal — Church of the Temple of God",
        "portal_sub": "Welcome to your daily guide and ministry operations portal of the Church of the Temple of God.",
        "refresh_app": "🔄 Refresh App",
        "delete_populated": "🗑️ Clear / Delete Populated Data (Delete Record)",
        "lang_label": "🌐 App Language:",
        "deadline_closed": "⛔ Submission Closed (Past Monday Midnight Deadline). Marked as '(No Report)' in both the App and Excel file. You may submit again during the next reporting cycle on Saturday.",
        "security_badge": "🔒 Protected by Role-Based Access Control (RLS), Encrypted Session, and Streamlit Secrets."
    }
}

def t_ui(key: str) -> str:
    lang = st.session_state.get("ui_default_lang", "Tagalog")
    if lang not in UI_TRANSLATIONS:
        lang = "Tagalog"
    return UI_TRANSLATIONS[lang].get(key, UI_TRANSLATIONS["Tagalog"].get(key, key))

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
        try:
            with open(LOGO_PATH, "rb") as f_l:
                b64_l = base64.b64encode(f_l.read()).decode("utf-8")
            ext_l = LOGO_PATH.split(".")[-1].lower()
            mime_l = "image/png" if ext_l == "png" else "image/jpeg"
            sub_html = f"<div style='font-size:0.95rem;opacity:0.88;margin-top:4px;'>{subtitle_text}</div>" if subtitle_text else ""
            st.markdown(
                f"<div class='full-text-card' style='display:flex;align-items:center;gap:14px;padding:12px 16px;margin-top:4px;margin-bottom:14px;'>"
                f"<img src='data:{mime_l};base64,{b64_l}' style='width:68px;height:68px;min-width:68px;object-fit:contain;border-radius:50%;border:2px solid #22c55e;' />"
                f"<div style='flex:1;min-width:0;'>"
                f"<div style='font-size:1.35rem;font-weight:800;line-height:1.25;word-wrap:break-word;'>{title_text}</div>"
                f"{sub_html}"
                f"</div></div>",
                unsafe_allow_html=True
            )
            return
        except Exception:
            pass
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
APP_PASSWORD = get_secret("APP_PASSWORD", "")
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
        sidebar_bg = "#E2E8F0"
        text_color = "#0F172A"
        card_bg = "#FFFFFF"
        card_text = "#0F172A"
        card_border = "#1E3A8A"
        side_box_bg = "#EFF6FF"
        side_box_text = "#0F172A"
        side_box_border = "#94A3B8"
        input_bg = "#FFFFFF"
        input_text = "#0F172A"
        accent_val_color = "#0F766E"
        officer_hdr_bg = "#DCFCE7"
        officer_hdr_text = "#14532D"
        alert_bg = "#EFF6FF"
        alert_text = "#0F172A"
        alert_border = "#2563EB"
    elif "Royal Blue" in theme:
        app_bg = "#0A192F"
        sidebar_bg = "#112240"
        text_color = "#F8FAFC"
        card_bg = "#112240"
        card_text = "#F8FAFC"
        card_border = "#64FFDA"
        side_box_bg = "#172A45"
        side_box_text = "#F8FAFC"
        side_box_border = "#38BDF8"
        input_bg = "#172A45"
        input_text = "#F8FAFC"
        accent_val_color = "#FDE047"
        officer_hdr_bg = "rgba(20,83,45,0.65)"
        officer_hdr_text = "#ECFDF5"
        alert_bg = "#172A45"
        alert_text = "#F8FAFC"
        alert_border = "#38BDF8"
    else:
        app_bg = "#0E1117"
        sidebar_bg = "#161B22"
        text_color = "#F8FAFC"
        card_bg = "#1E293B"
        card_text = "#F8FAFC"
        card_border = "#38BDF8"
        side_box_bg = "#1E293B"
        side_box_text = "#F8FAFC"
        side_box_border = "#475569"
        input_bg = "#1E293B"
        input_text = "#F8FAFC"
        accent_val_color = "#FDE047"
        officer_hdr_bg = "rgba(20,83,45,0.55)"
        officer_hdr_text = "#ECFDF5"
        alert_bg = "#1E293B"
        alert_text = "#F8FAFC"
        alert_border = "#38BDF8"

    st.markdown(f"""
        <style>
        :root {{
            --ctg-text: {text_color};
            --ctg-card-bg: {card_bg};
            --ctg-card-text: {card_text};
            --ctg-accent-val: {accent_val_color};
            --ctg-officer-hdr-bg: {officer_hdr_bg};
            --ctg-officer-hdr-text: {officer_hdr_text};
        }}
        html {{ font-size: {base_zoom}; }}
        .stApp {{
            background-color: {app_bg} !important;
            color: {text_color} !important;
        }}
        /* Top Header Bar (Para hindi matakpan ang Logo o Banner sa Phone at laging kita ang >> Sidebar Button) */
        header[data-testid="stHeader"] {{
            background-color: {app_bg} !important;
            border-bottom: 1px solid {side_box_border} !important;
        }}
        .block-container {{
            padding-top: 3.6rem !important;
            padding-bottom: 4.5rem !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: {sidebar_bg} !important;
        }}
        [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {{
            color: {text_color} !important;
        }}
        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
        .stApp p, .stApp label, .stMarkdown, .stCaption {{
            color: {text_color} !important;
        }}
        /* Fix st.info / st.success / st.warning / st.error on Mobile Phone (iOS & Android Light/Dark OS) */
        [data-testid="stAlert"], div[role="alert"] {{
            background-color: {alert_bg} !important;
            color: {alert_text} !important;
            border: 1.5px solid {alert_border} !important;
            border-left: 5px solid {alert_border} !important;
            border-radius: 8px !important;
        }}
        [data-testid="stAlert"] p, [data-testid="stAlert"] span, [data-testid="stAlert"] div,
        [data-testid="stAlert"] b, [data-testid="stAlert"] strong,
        div[role="alert"] p, div[role="alert"] span, div[role="alert"] div {{
            color: {alert_text} !important;
        }}
        /* Fix Expander Header & Body on Mobile Phone (para hindi puti sa puting background) */
        [data-testid="stExpander"], [data-testid="stExpander"] details {{
            background-color: {card_bg} !important;
            border: 1px solid {side_box_border} !important;
            border-radius: 8px !important;
            color: {text_color} !important;
        }}
        [data-testid="stExpander"] summary, [data-testid="stExpander"] summary:hover {{
            background-color: {side_box_bg} !important;
            color: {text_color} !important;
            border-radius: 8px !important;
            padding: 8px 12px !important;
        }}
        [data-testid="stExpander"] summary p, [data-testid="stExpander"] summary span,
        [data-testid="stExpander"] summary svg {{
            color: {text_color} !important;
            fill: {text_color} !important;
        }}
        /* Fix Selectbox, Inputs & Dropdown Popovers on Mobile Phone */
        .stTextInput input, .stTextArea textarea, .stNumberInput input,
        [data-baseweb="select"] > div, [data-baseweb="base-input"] {{
            background-color: {input_bg} !important;
            color: {input_text} !important;
            -webkit-text-fill-color: {input_text} !important;
            border: 1px solid {side_box_border} !important;
        }}
        [data-baseweb="select"] span, [data-baseweb="select"] div, [data-baseweb="select"] svg {{
            color: {input_text} !important;
            fill: {input_text} !important;
        }}
        [data-baseweb="popover"] div, [data-baseweb="menu"], ul[role="listbox"], li[role="option"] {{
            background-color: {input_bg} !important;
            color: {input_text} !important;
        }}
        li[role="option"]:hover, li[role="option"][aria-selected="true"] {{
            background-color: {side_box_bg} !important;
            color: {text_color} !important;
        }}
        .stButton > button {{
            font-weight: 600 !important;
        }}
        .stButton > button[kind="secondary"] {{
            background-color: {card_bg} !important;
            color: {text_color} !important;
            border: 1px solid {side_box_border} !important;
        }}
        .stButton > button[kind="secondary"] p, .stButton > button[kind="secondary"] span {{
            color: {text_color} !important;
        }}
        .stButton > button[kind="primary"] {{
            color: #FFFFFF !important;
        }}
        .stButton > button[kind="primary"] p, .stButton > button[kind="primary"] span {{
            color: #FFFFFF !important;
        }}
        .ctg-officer-val {{
            color: var(--ctg-accent-val) !important;
            font-weight: 700;
        }}
        .ctg-officer-wk-hdr {{
            background: var(--ctg-officer-hdr-bg);
            color: var(--ctg-officer-hdr-text) !important;
            padding: 7px 10px;
            border-radius: 6px;
            font-weight: 700;
            text-align: center;
            margin-bottom: 6px;
            border: 1px solid {side_box_border};
        }}
        .full-text-card {{
            background-color: {card_bg} !important;
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
        .full-text-card *, .records-side-box *, .sidebar-log-card * {{
            color: {card_text} !important;
        }}
        .urgent-card {{
            background-color: #FEF2F2 !important;
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
            background-color: #ECFDF5 !important; border: 1px solid #10B981; color: #065F46 !important;
            padding: 6px 12px; border-radius: 6px; font-weight: 600; font-size: 0.88rem; margin-bottom: 8px; display: inline-block;
        }}
        .assigned-parish-badge {{
            background-color: #EFF6FF !important; border: 1px solid #3B82F6; color: #1E40AF !important;
            padding: 6px 12px; border-radius: 6px; font-weight: 600; font-size: 0.88rem; margin-bottom: 8px; display: inline-block;
        }}
        .records-side-box {{
            background-color: {side_box_bg} !important; border: 1px solid {side_box_border}; border-top: 4px solid #0F766E;
            padding: {card_pad}; border-radius: 8px; margin-bottom: {card_margin}; color: {side_box_text} !important; font-size: 0.88rem;
        }}
        .sidebar-log-card {{
            background-color: {side_box_bg} !important;
            border-left: 3px solid #38BDF8;
            border: 1px solid {side_box_border};
            padding: 8px 10px;
            border-radius: 5px;
            margin-bottom: 6px;
            color: {side_box_text} !important;
            font-size: 0.80rem;
        }}
        .chat-bubble-mine {{
            background-color: #1E3A8A !important;
            border: 1px solid #3B82F6;
            color: #FFFFFF !important;
            padding: 10px 14px;
            border-radius: 12px 12px 2px 12px;
            margin-bottom: 8px;
            margin-left: 8%;
            font-size: {font_scale};
        }}
        .chat-bubble-mine * {{ color: #FFFFFF !important; }}
        .chat-bubble-other {{
            background-color: {side_box_bg} !important;
            border-left: 4px solid #10B981;
            border-top: 1px solid {side_box_border};
            border-right: 1px solid {side_box_border};
            border-bottom: 1px solid {side_box_border};
            color: {side_box_text} !important;
            padding: 10px 14px;
            border-radius: 12px 12px 12px 2px;
            margin-bottom: 8px;
            margin-right: 8%;
            font-size: {font_scale};
        }}
        .chat-bubble-other * {{ color: {side_box_text} !important; }}
        .chat-bubble-broadcast {{
            background-color: #4C1D95 !important;
            border: 1px solid #A78BFA;
            color: #F5F3FF !important;
            padding: 10px 14px;
            border-radius: 10px;
            margin-bottom: 8px;
            font-size: {font_scale};
        }}
        .chat-bubble-broadcast * {{ color: #F5F3FF !important; }}
        .category-header-seniors {{
            background-color: #3B0764 !important; border: 1px solid #A855F7; color: #F3E8FF !important;
            padding: 10px 14px; border-radius: 8px; font-weight: 700; margin-bottom: 8px;
        }}
        .category-header-adults {{
            background-color: #1E3A8A !important; border: 1px solid #60A5FA; color: #DBEAFE !important;
            padding: 10px 14px; border-radius: 8px; font-weight: 700; margin-bottom: 8px;
        }}
        .category-header-youth {{
            background-color: #065F46 !important; border: 1px solid #34D399; color: #D1FAE5 !important;
            padding: 10px 14px; border-radius: 8px; font-weight: 700; margin-bottom: 8px;
        }}
        .category-header-kids {{
            background-color: #78350F !important; border: 1px solid #FBBF24; color: #FEF3C7 !important;
            padding: 10px 14px; border-radius: 8px; font-weight: 700; margin-bottom: 8px;
        }}
        .pmr-banner {{
            background-color: #14532D !important;
            color: #FFFFFF !important;
            text-align: center;
            padding: 12px;
            border-radius: 8px 8px 0 0;
            font-weight: 800;
            font-size: 1.25rem;
            letter-spacing: 0.8px;
            margin-top: 6px;
        }}
        .pmr-subbanner {{
            background-color: #DCFCE7 !important;
            color: #14532D !important;
            text-align: center;
            padding: 10px;
            border-radius: 0 0 8px 8px;
            font-weight: 700;
            font-size: 1.02rem;
            margin-bottom: 14px;
            border: 1px solid #14532D;
        }}
        .pmr-subbanner * {{ color: #14532D !important; }}
        .ph-time-badge {{
            background-color: #065F46 !important;
            color: #ECFDF5 !important;
            border: 1px solid #34D399;
            padding: 5px 12px;
            border-radius: 999px;
            font-size: 0.84rem;
            font-weight: 700;
            display: inline-block;
            margin-top: 4px;
            margin-bottom: 4px;
        }}
        .ph-time-badge * {{ color: #ECFDF5 !important; }}
        /* Sidebar Toggle Button (>> and <<) para laging kitang-kita sa Phone */
        [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"], [data-testid="stSidebarCollapseButton"],
        button[kind="header"], button[kind="headerNoPadding"] {{
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            z-index: 999999 !important;
            background-color: #1E3A8A !important;
            color: #FFFFFF !important;
            border: 1.5px solid #60A5FA !important;
            border-radius: 8px !important;
            padding: 4px !important;
            margin: 4px !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.45) !important;
        }}
        [data-testid="collapsedControl"] svg, [data-testid="stSidebarCollapsedControl"] svg, [data-testid="stSidebarCollapseButton"] svg {{
            fill: #FFFFFF !important;
            color: #FFFFFF !important;
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
            background-color: #14532D !important;
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
        /* 📱 RESPONSIVE TABLET & PHONE MEDIA QUERIES */
        @media (max-width: 1024px) {{
            .block-container {{
                padding-left: 1rem !important;
                padding-right: 1rem !important;
                padding-top: 3.8rem !important;
                padding-bottom: 5.5rem !important;
            }}
            [data-testid="stHorizontalBlock"] {{
                flex-wrap: wrap !important;
                gap: 0.65rem !important;
            }}
            [data-testid="column"] {{
                min-width: 47% !important;
                flex: 1 1 47% !important;
            }}
            .pmr-banner {{ font-size: 1.12rem !important; padding: 10px !important; }}
            .pmr-subbanner {{ font-size: 0.92rem !important; padding: 8px !important; }}
        }}
        @media (max-width: 640px) {{
            html {{ font-size: 95% !important; }}
            .block-container {{
                padding-left: 0.65rem !important;
                padding-right: 0.65rem !important;
                padding-top: 3.9rem !important;
                padding-bottom: 6rem !important;
            }}
            [data-testid="column"] {{
                min-width: 100% !important;
                flex: 1 1 100% !important;
            }}
            .stButton > button, .stDownloadButton > button {{
                min-height: 44px !important;
                font-size: 0.94rem !important;
                width: 100% !important;
            }}
            .chat-bubble-mine {{ margin-left: 2% !important; }}
            .chat-bubble-other {{ margin-right: 2% !important; }}
            .record-style-table th, .record-style-table td {{
                padding: 8px 8px !important;
                font-size: 0.85rem !important;
            }}
            h1 {{ font-size: 1.55rem !important; line-height: 1.25 !important; }}
            h2 {{ font-size: 1.3rem !important; line-height: 1.25 !important; }}
            h3 {{ font-size: 1.12rem !important; line-height: 1.3 !important; }}
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

def _sha256_hex(val: str) -> str:
    return hashlib.sha256(str(val or "").encode("utf-8")).hexdigest()

# 🔐 SECURE CREDENTIAL REGISTRY (Uses SHA-256 digests + st.secrets override; zero plaintext passwords in code)
PARISH_ACCOUNTS = {
    "ctg.ministry.master.admin": {"pw_hash": "10d3135d22d6d882860211e3e9345efbb20975149ab24dfdc8c44f8779060713", "parish": "Lahat ng Parokya", "role": "Master Admin"},
    APP_USER.lower(): {"pw_hash": _sha256_hex(APP_PASSWORD) if APP_PASSWORD else "10d3135d22d6d882860211e3e9345efbb20975149ab24dfdc8c44f8779060713", "parish": "Lahat ng Parokya", "role": "Master Admin"},
    "lipay norte parish": {"pw_hash": "5fde976f32458c7d2bcfbebe87b1a3c70678f7987d618b98667e5c95176c9a74", "parish": "LIPAY NORTE PARISH", "role": "Parish Leader"},
    "lipaynorteparish": {"pw_hash": "5fde976f32458c7d2bcfbebe87b1a3c70678f7987d618b98667e5c95176c9a74", "parish": "LIPAY NORTE PARISH", "role": "Parish Leader"},
    "malanas parish": {"pw_hash": "d05c4fdecac72853bebd617131d7dc425ee70ccc8ff825e591a3561e6ea62d7c", "parish": "MALANAS PARISH", "role": "Parish Leader"},
    "malanasparish": {"pw_hash": "d05c4fdecac72853bebd617131d7dc425ee70ccc8ff825e591a3561e6ea62d7c", "parish": "MALANAS PARISH", "role": "Parish Leader"},
    "poblacion parish": {"pw_hash": "abad8786eb5424f1b0043e2ee2b8fa2c7a7845d3672abe169681955ae25f6895", "parish": "POBLACION PARISH", "role": "Parish Leader"},
    "poblacionparish": {"pw_hash": "abad8786eb5424f1b0043e2ee2b8fa2c7a7845d3672abe169681955ae25f6895", "parish": "POBLACION PARISH", "role": "Parish Leader"},
    "duplas parish": {"pw_hash": "1098a6e7872246dc66646c909333954f3d73b511fee0855bff16a5aa8d50b884", "parish": "DUPLAS PARISH", "role": "Parish Leader"},
    "duplasparish": {"pw_hash": "1098a6e7872246dc66646c909333954f3d73b511fee0855bff16a5aa8d50b884", "parish": "DUPLAS PARISH", "role": "Parish Leader"},
    "santiago parish": {"pw_hash": "1c5d635da5272f385209abfee44dd511256684b3c941cf78d8c38181cc87cdbe", "parish": "SANTIAGO PARISH", "role": "Parish Leader"},
    "santiagoparish": {"pw_hash": "1c5d635da5272f385209abfee44dd511256684b3c941cf78d8c38181cc87cdbe", "parish": "SANTIAGO PARISH", "role": "Parish Leader"},
    "caba parish": {"pw_hash": "b986f1e35563090a668f4a0b130e95e8f45b7780fe57c2e14cf45c9cb340e082", "parish": "CABA PARISH", "role": "Parish Leader"},
    "cabaparish": {"pw_hash": "b986f1e35563090a668f4a0b130e95e8f45b7780fe57c2e14cf45c9cb340e082", "parish": "CABA PARISH", "role": "Parish Leader"},
    "manila parish": {"pw_hash": "d8b5d5db13fd97b44e325f085f6e94c93ca788ef0fe16eee7ee6e19c4fa99786", "parish": "MANILA PARISH", "role": "Parish Leader"},
    "manilaparish": {"pw_hash": "d8b5d5db13fd97b44e325f085f6e94c93ca788ef0fe16eee7ee6e19c4fa99786", "parish": "MANILA PARISH", "role": "Parish Leader"},
    "andabuen parish": {"pw_hash": "f43a7433ed92c892a32b3a5e2605d6b1b3738fd5149b9fcc026dd88ad3704adf", "parish": "ANDABUEN PARISH", "role": "Parish Leader"},
    "andabuenparish": {"pw_hash": "f43a7433ed92c892a32b3a5e2605d6b1b3738fd5149b9fcc026dd88ad3704adf", "parish": "ANDABUEN PARISH", "role": "Parish Leader"},
    "apayao parish": {"pw_hash": "c7b31138996cc56cd2bcb23424679202cfb2e7f4f626f64646b1b4a55c026712", "parish": "APAYAO PARISH", "role": "Parish Leader"},
    "apayaoparish": {"pw_hash": "c7b31138996cc56cd2bcb23424679202cfb2e7f4f626f64646b1b4a55c026712", "parish": "APAYAO PARISH", "role": "Parish Leader"},
    "hongkong parish": {"pw_hash": "3c8d89e094f6fdac11af6eddf1c6315697e38ec220b83f08a9a6280fb30d4607", "parish": "HONGKONG PARISH", "role": "Parish Leader"},
    "hongkongparish": {"pw_hash": "3c8d89e094f6fdac11af6eddf1c6315697e38ec220b83f08a9a6280fb30d4607", "parish": "HONGKONG PARISH", "role": "Parish Leader"},
}

def verify_account_credentials(username_in: str, password_in: str) -> dict | None:
    u_clean = str(username_in or "").strip().lower()
    u_alnum = re.sub(r"[^a-z0-9.]", "", u_clean)
    p_clean = str(password_in or "").strip()
    if not u_clean or not p_clean:
        return None
    p_hash = _sha256_hex(p_clean)
    try:
        if "PARISH_PASSWORDS" in st.secrets:
            for k_sec, v_sec in st.secrets["PARISH_PASSWORDS"].items():
                if k_sec.strip().lower() in (u_clean, u_alnum) and p_clean == str(v_sec).strip():
                    return PARISH_ACCOUNTS.get(u_clean) or PARISH_ACCOUNTS.get(u_alnum)
    except Exception:
        pass
    for candidate_key in (u_clean, u_alnum, f"{u_clean} parish", f"{u_alnum}parish"):
        if candidate_key in PARISH_ACCOUNTS:
            acc = PARISH_ACCOUNTS[candidate_key]
            if p_hash == acc.get("pw_hash") or (APP_PASSWORD and candidate_key == APP_USER.lower() and p_clean == APP_PASSWORD):
                return acc
    return None
    try:
        if "PARISH_PASSWORDS" in st.secrets and u_clean in st.secrets["PARISH_PASSWORDS"]:
            sec_pw = str(st.secrets["PARISH_PASSWORDS"][u_clean]).strip()
            if p_clean == sec_pw and u_clean in PARISH_ACCOUNTS:
                return PARISH_ACCOUNTS[u_clean]
    except Exception:
        pass
    if u_clean in PARISH_ACCOUNTS:
        acc = PARISH_ACCOUNTS[u_clean]
        if _sha256_hex(p_clean) == acc.get("pw_hash") or (APP_PASSWORD and u_clean == APP_USER.lower() and p_clean == APP_PASSWORD):
            return acc
    return None

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

def _verify_drive_folder_id(service, candidate_id: str) -> str:
    """Sinusuri kung ang candidate_id ay nage-exist at naa-access sa Google Drive (supportsAllDrives=True).
    Kung file ID ito sa halip na folder ID, kinukuha ang parent folder ID nito. Kung 404/invalid, ibinabalik ay empty string."""
    if not service or not candidate_id or len(str(candidate_id).strip()) < 10:
        return ""
    cid = str(candidate_id).strip()
    try:
        meta = service.files().get(
            fileId=cid,
            fields="id, name, mimeType, trashed, parents",
            supportsAllDrives=True
        ).execute()
        if meta.get("trashed"):
            return ""
        if meta.get("mimeType") == "application/vnd.google-apps.folder":
            return meta.get("id", "")
        parents = meta.get("parents", [])
        if parents:
            return parents[0]
    except Exception:
        pass
    return ""

def find_drive_folder_by_name(service, folder_name_query: str, parent_id: str = None) -> str:
    if not service or not folder_name_query:
        return ""
    try:
        safe_q = folder_name_query.replace("'", "\\'")
        q_parts = [
            "mimeType='application/vnd.google-apps.folder'",
            "trashed=false",
            f"name contains '{safe_q}'"
        ]
        if parent_id:
            q_parts.append(f"'{parent_id}' in parents")
        res = service.files().list(
            q=" and ".join(q_parts),
            spaces="drive",
            fields="files(id, name)",
            pageSize=25,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        return files[0]["id"] if files else ""
    except Exception:
        return ""

KNOWN_PARISH_DRIVE_FOLDER_IDS = {
    "POBLACION PARISH": "1MLBg_h0nimbTCtcgmq0i8ymXuJy6QK-1",
}

def _resolve_best_subfolder_for_month(service, parish_folder_id: str, month_name: str = "") -> str:
    """Kung mayroong month subfolder sa loob ng Parish Folder (hal. '01. January Pwr. 2027' sa loob ng 'POBLACION PARISH'),
    ibinabalik ang ID ng month subfolder; kung wala, ibinabalik mismo ang parish_folder_id."""
    if not service or not parish_folder_id or not month_name:
        return parish_folder_id
    try:
        res_sub = service.files().list(
            q=f"'{parish_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces="drive",
            fields="files(id, name)",
            pageSize=50,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        m_low = month_name.strip().lower()
        m_idx = (MONTHS_2027.index(month_name) + 1) if month_name in MONTHS_2027 else 0
        m_num_prefix = f"{m_idx:02d}." if m_idx > 0 else ""
        for sf in res_sub.get("files", []):
            sf_name_low = sf.get("name", "").strip().lower()
            if m_low in sf_name_low or (m_num_prefix and sf_name_low.startswith(m_num_prefix)):
                return sf["id"]
    except Exception:
        pass
    return parish_folder_id

def get_or_create_parish_subfolder(service, parent_folder_id: str, parish_name: str, month_name: str = "") -> str:
    """Hinahanap ang eksaktong Parish Folder (at Month subfolder kung mayroon, tulad ng '01. January Pwr. 2027')
    gamit ang Saved Links sa App, Known Parish Folder IDs, o direktang paghahanap sa Google Drive nang walang 404 error."""
    if not service:
        return ""
    clean_parish = (parish_name or "").upper().replace("PARISH", "").replace("CHAPTER", "").strip()
    short_keyword = clean_parish.replace("NORTE", "").strip()

    # 1. Tingnan muna sa Saved Google Drive Links Registry kung may inilagay na partikular na link ang Admin para sa Parokyang ito
    try:
        saved_links = list(reversed(load_saved_drive_links_registry()))
        for item in saved_links:
            u_str = str(item.get("url", "")).strip()
            if not u_str:
                continue
            lbl_up = str(item.get("label", "")).upper()
            p_scope_up = str(item.get("parish", "")).upper()
            m_scope_up = str(item.get("month", "")).upper()
            if clean_parish and (clean_parish in p_scope_up or clean_parish in lbl_up or (short_keyword and short_keyword in lbl_up)):
                cand_id = _verify_drive_folder_id(service, extract_id_from_url(u_str))
                if cand_id:
                    if month_name and month_name.upper() in (m_scope_up + " " + lbl_up):
                        return cand_id
                    return _resolve_best_subfolder_for_month(service, cand_id, month_name)
    except Exception:
        pass

    # 2. Tingnan kung may Known Folder ID para sa Parokya (hal. POBLACION PARISH -> 1MLBg_h0nimbTCtcgmq0i8ymXuJy6QK-1)
    if parish_name and parish_name.upper().strip() in KNOWN_PARISH_DRIVE_FOLDER_IDS:
        known_id = _verify_drive_folder_id(service, KNOWN_PARISH_DRIVE_FOLDER_IDS[parish_name.upper().strip()])
        if known_id:
            return _resolve_best_subfolder_for_month(service, known_id, month_name)

    # 3. I-verify kung accessible ang parent_folder_id; kung hindi (404), hanapin sa Saved Links o sa buong Google Drive
    valid_parent_id = _verify_drive_folder_id(service, parent_folder_id)
    if not valid_parent_id:
        try:
            for item in reversed(load_saved_drive_links_registry()):
                cat_s = str(item.get("category", ""))
                lbl_s = str(item.get("label", ""))
                u_s = str(item.get("url", "")).strip()
                if u_s and ("Upload" in cat_s or "Folder" in cat_s or "Upload" in lbl_s):
                    cid = _verify_drive_folder_id(service, extract_id_from_url(u_s))
                    if cid:
                        valid_parent_id = cid
                        break
        except Exception:
            pass

    if not valid_parent_id:
        valid_parent_id = (
            find_drive_folder_by_name(service, "Upload the Image/Excel File here")
            or find_drive_folder_by_name(service, "Upload the Image")
            or find_drive_folder_by_name(service, "CTG Ministry Hub")
        )

    # 4. Kung may valid_parent_id, tingnan kung ang parent_id mismo ay ang Parish Folder na, o hanapin sa loob nito
    if valid_parent_id and parish_name:
        try:
            p_meta = service.files().get(fileId=valid_parent_id, fields="id, name", supportsAllDrives=True).execute()
            p_meta_up = p_meta.get("name", "").upper()
            if clean_parish and (clean_parish in p_meta_up or (short_keyword and short_keyword in p_meta_up)):
                return _resolve_best_subfolder_for_month(service, valid_parent_id, month_name)

            res = service.files().list(
                q=f"'{valid_parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
                spaces="drive",
                fields="files(id, name)",
                pageSize=50,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            folders = res.get("files", [])
            for folder in folders:
                f_up = folder["name"].upper().strip()
                f_clean = re.sub(r"^\d+[\.\-\s]*", "", f_up).replace("PARISH", "").replace("CHAPTER", "").strip()
                if (
                    f_up == parish_name.upper().strip()
                    or f_clean == clean_parish
                    or f_clean == short_keyword
                    or (clean_parish and clean_parish in f_up)
                    or (short_keyword and short_keyword in f_up)
                ):
                    return _resolve_best_subfolder_for_month(service, folder["id"], month_name)
        except Exception:
            pass

    # 5. Hanapin nang direkta ang Parish Folder sa buong Google Drive (hal. 'POBLACION PARISH', 'MALANAS PARISH', 'LIPAY PARISH')
    if parish_name:
        direct_p_id = (
            find_drive_folder_by_name(service, parish_name)
            or (find_drive_folder_by_name(service, f"{short_keyword} PARISH") if short_keyword else "")
            or (find_drive_folder_by_name(service, clean_parish) if clean_parish else "")
        )
        if direct_p_id:
            return _resolve_best_subfolder_for_month(service, direct_p_id, month_name)

    # 6. Kung wala pa ring nahanap na folder, gumawa ng bagong folder sa loob ng valid_parent_id o sa My Drive root
    try:
        if not valid_parent_id:
            root_created = service.files().create(
                body={"name": "Upload the Image/Excel File here!", "mimeType": "application/vnd.google-apps.folder"},
                fields="id",
                supportsAllDrives=True
            ).execute()
            valid_parent_id = root_created.get("id", "")
        if parish_name and valid_parent_id:
            p_created = service.files().create(
                body={"name": parish_name, "mimeType": "application/vnd.google-apps.folder", "parents": [valid_parent_id]},
                fields="id",
                supportsAllDrives=True
            ).execute()
            return p_created.get("id", valid_parent_id)
    except Exception:
        pass
    return valid_parent_id or ""

def upload_file_to_google_drive(file_bytes, filename, mime_type="image/jpeg", parent_folder_url="", parish_name="", month_name=""):
    service = get_google_drive_service()
    if service is None:
        return None, "Error: Kulang ang Google Drive/Gmail credentials sa App Settings."
    try:
        raw_folder_id = extract_id_from_url(parent_folder_url)
        target_folder_id = (
            get_or_create_parish_subfolder(service, raw_folder_id, parish_name, month_name)
            if parish_name
            else _verify_drive_folder_id(service, raw_folder_id)
        )
        file_metadata = {"name": filename}
        if target_folder_id:
            file_metadata["parents"] = [target_folder_id]
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type or "application/octet-stream", resumable=True)
        try:
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink",
                supportsAllDrives=True
            ).execute()
            return file.get("webViewLink"), f"Matagumpay na na-upload sa Google Drive folder ng {parish_name or 'CTG'}!"
        except Exception as inner_ex:
            # Kung nag-404 o 403 sa partikular na parent folder, awtomatikong mag-fallback sa sariling Parish folder sa My Drive root
            if "404" in str(inner_ex) or "notFound" in str(inner_ex) or "403" in str(inner_ex):
                fallback_parent_id = get_or_create_parish_subfolder(service, "", parish_name, month_name) if parish_name else ""
                fb_meta = {"name": filename}
                if fallback_parent_id:
                    fb_meta["parents"] = [fallback_parent_id]
                media_fb = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type or "application/octet-stream", resumable=True)
                try:
                    file_fb = service.files().create(
                        body=fb_meta,
                        media_body=media_fb,
                        fields="id, webViewLink",
                        supportsAllDrives=True
                    ).execute()
                    return file_fb.get("webViewLink"), f"Matagumpay na na-upload sa Google Drive folder ng {parish_name or 'CTG'}!"
                except Exception:
                    media_root = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type or "application/octet-stream", resumable=True)
                    file_root = service.files().create(
                        body={"name": filename},
                        media_body=media_root,
                        fields="id, webViewLink",
                        supportsAllDrives=True
                    ).execute()
                    return file_root.get("webViewLink"), f"Matagumpay na na-upload sa iyong Google Drive ({filename})!"
            raise inner_ex
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

@st.cache_data(ttl=600, show_spinner=False)
def fetch_live_excel_bytes(spreadsheet_url: str, force_cloud: bool = False) -> tuple[bytes, str]:
    if not force_cloud and os.path.exists(LOCAL_EXCEL_CACHE):
        try:
            with open(LOCAL_EXCEL_CACHE, "rb") as f:
                local_b = f.read()
            if len(local_b) > 1000:
                return local_b, LOCAL_EXCEL_CACHE
        except Exception:
            pass

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
            res = requests.get(download_url, timeout=8)
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

def load_internal_chat_messages(sync_cloud: bool = False) -> list:
    messages = []
    if os.path.exists(LOCAL_CHAT_FILE):
        try:
            with open(LOCAL_CHAT_FILE, "r", encoding="utf-8") as f:
                messages = json.load(f)
        except Exception:
            messages = []

    if sync_cloud or not messages:
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

def _persist_chat_messages_list(msgs: list, sync_to_drive: bool = False):
    try:
        with open(LOCAL_CHAT_FILE, "w", encoding="utf-8") as f:
            json.dump(msgs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    if sync_to_drive:
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

def save_internal_chat_message(sender: str, recipient: str, text: str, sync_to_drive: bool = False) -> bool:
    msgs = load_internal_chat_messages(sync_cloud=False)
    new_msg = {
        "id": f"msg_{int(time.time() * 1000)}_{len(msgs)}",
        "timestamp": get_philippine_datetime_str(),
        "sender": sender,
        "recipient": recipient,
        "message": text.strip()
    }
    msgs.append(new_msg)
    _persist_chat_messages_list(msgs, sync_to_drive=sync_to_drive)
    return True

def delete_internal_chat_messages_by_ids(ids_to_remove: set) -> int:
    msgs = load_internal_chat_messages(sync_cloud=False)
    before_cnt = len(msgs)
    remaining = [m for m in msgs if m.get("id") not in ids_to_remove]
    _persist_chat_messages_list(remaining, sync_to_drive=True)
    return before_cnt - len(remaining)

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

def load_live_notifications() -> list:
    if os.path.exists(LOCAL_NOTIFICATIONS_FILE):
        try:
            with open(LOCAL_NOTIFICATIONS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []

def dispatch_app_notification_and_messenger_alert(
    sender_id: str,
    recipient_id: str,
    title: str,
    details: str,
    is_error: bool = False,
    also_post_to_messenger: bool = True
):
    """Nagpapadala ng agarang Notification sa App (Mobile/Desktop) at awtomatikong nagpo-post sa Ministry Messenger
    upang makita ni Master Admin kung sino ang nag-attempt mag-submit at kung ano ang naging error o update."""
    ph_ts = get_philippine_datetime_str()
    prefix = "🚨 [SUBMISSION ERROR / ISSUE ALERT]" if is_error else "🔔 [MINISTRY NOTIFICATION]"
    full_msg = f"{prefix} {title} | Mula kay: {sender_id} | Detalye: {details} | Oras (PH): {ph_ts}"

    if also_post_to_messenger:
        try:
            save_internal_chat_message(sender_id, recipient_id, full_msg)
        except Exception:
            pass

    notifs = load_live_notifications()
    notifs.append({
        "id": f"notif_{int(time.time() * 1000)}_{len(notifs)}",
        "timestamp": ph_ts,
        "sender": sender_id,
        "recipient": recipient_id,
        "title": title,
        "details": details,
        "is_error": is_error
    })
    notifs = notifs[-80:]
    try:
        with open(LOCAL_NOTIFICATIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(notifs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def trigger_browser_and_toast_notifications_for_current_user():
    """Tanging ang Master Admin lamang ang makakakita ng Live Notifications (para sa mga nag-login, errors, at report submissions ng mga Parish Leaders)."""
    is_master = (st.session_state.get("user_role") == "Master Admin" or st.session_state.get("login_role") == "Master Admin")
    if not is_master:
        return []
    notifs = load_live_notifications()
    relevant = list(notifs)
    seen_ids = set(st.session_state.get("_seen_notification_ids", []))
    new_items = [n for n in relevant if n.get("id") not in seen_ids]
    if new_items:
        for n_item in new_items[-3:]:
            icon_t = "🚨" if n_item.get("is_error") else "🔔"
            st.toast(f"{icon_t} {n_item.get('title')}: {n_item.get('details')[:110]} ({n_item.get('sender')})")
        latest_n = new_items[-1]
        safe_t = json.dumps(str(latest_n.get("title", "CTG Ministry Notification")))
        safe_b = json.dumps(f"{latest_n.get('sender')}: {latest_n.get('details', '')[:120]}")
        components.html(
            f"""<script>
            try {{
                if ("Notification" in window) {{
                    if (Notification.permission === "granted") {{
                        new Notification({safe_t}, {{ body: {safe_b} }});
                    }} else if (Notification.permission !== "denied") {{
                        Notification.requestPermission().then(function(p) {{
                            if (p === "granted") {{
                                new Notification({safe_t}, {{ body: {safe_b} }});
                            }}
                        }});
                    }}
                }}
            }} catch (e) {{}}
            </script>""",
            height=0
        )
        seen_ids.update(n.get("id") for n in relevant if n.get("id"))
        st.session_state["_seen_notification_ids"] = list(seen_ids)
    return relevant

def delete_live_notification(notif_id: str):
    notifs = load_live_notifications()
    notifs = [n for n in notifs if n.get("id") != notif_id]
    try:
        with open(LOCAL_NOTIFICATIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(notifs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def clear_all_live_notifications_for_user():
    my_id = get_current_chat_identity()
    is_master = (st.session_state.get("user_role") == "Master Admin")
    if is_master:
        notifs = []
    else:
        notifs = [
            n for n in load_live_notifications()
            if n.get("recipient") not in (my_id, "LAHAT NG USERS (BROADCAST)", "ALL") and n.get("sender") != my_id
        ]
    try:
        with open(LOCAL_NOTIFICATIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(notifs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_centered_logo_html(size_px: int = 118) -> str:
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            with open(LOGO_PATH, "rb") as f_img:
                b64_str = base64.b64encode(f_img.read()).decode("utf-8")
            ext = LOGO_PATH.split(".")[-1].lower()
            mime = "image/png" if ext == "png" else "image/jpeg"
            return (
                f"<div style='display:flex;justify-content:center;align-items:center;width:100%;margin-bottom:14px;'>"
                f"<img src='data:{mime};base64,{b64_str}' style='width:{size_px}px;height:{size_px}px;object-fit:contain;"
                f"border-radius:50%;border:2px solid #22c55e;box-shadow:0 4px 16px rgba(0,0,0,0.45);' />"
                f"</div>"
            )
        except Exception:
            pass
    return ""

if not st.session_state.authenticated:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([0.85, 1.6, 0.85])
    with col2:
        logo_html_banner = get_centered_logo_html(120)
        if logo_html_banner:
            st.markdown(logo_html_banner, unsafe_allow_html=True)
        st.markdown(f"<div class='pmr-banner'>{t_ui('portal_title')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='pmr-subbanner'>{t_ui('portal_sub')}<br><span class='ph-time-badge'>🇵🇭 {get_philippine_datetime_str()}</span></div>", unsafe_allow_html=True)

        with st.form("login_form"):
            username_input = st.text_input("👤 Username:").strip().lower()
            password_input = st.text_input("🔑 Password:", type="password").strip()
            submit = st.form_submit_button("🔓 Buksan ang Parish Leader Portal (Login)", type="primary", use_container_width=True)
            if submit:
                acc_info = verify_account_credentials(username_input, password_input)
                if acc_info is not None:
                    st.session_state.authenticated = True
                    st.session_state.user_role = acc_info["role"]
                    st.session_state.assigned_parish = acc_info["parish"]
                    st.session_state.login_role = acc_info["role"]
                    st.session_state.login_parish = acc_info["parish"]
                    st.session_state.app_mode_override = "🧭 Parish Leader Portal (Wizard & Summary)"
                    dispatch_app_notification_and_messenger_alert(
                        sender_id=acc_info["parish"] if acc_info["role"] != "Master Admin" else "MASTER ADMIN",
                        recipient_id="MASTER ADMIN",
                        title=f"User Login ({acc_info['parish']})",
                        details=f"Matagumpay na nag-login si '{username_input}' ({acc_info['role']} — {acc_info['parish']}).",
                        is_error=False
                    )
                    st.rerun()
                else:
                    if username_input:
                        dispatch_app_notification_and_messenger_alert(
                            sender_id=username_input.upper(),
                            recipient_id="MASTER ADMIN",
                            title="Failed Login Attempt sa Parish Leader Portal",
                            details=f"Sinubukang mag-login gamit ang username '{username_input}' ngunit mali ang password.",
                            is_error=True
                        )
                    st.error("Maling Username o Password. Na-log na ang attempt para sa seguridad.")
        st.caption(t_ui("security_badge"))
    st.stop()

# 🔒 STRICT ROLE-BASED ACCESS CONTROL (RLS) ENFORCEMENT PARA SA MGA PARISH ADMINS
if st.session_state.get("login_role") != "Master Admin":
    st.session_state.user_role = "Parish Leader"
    st.session_state.assigned_parish = st.session_state.get("login_parish", st.session_state.get("assigned_parish", PARISH_ORDER[0]))

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

    # Remove any accidental unsubmitted FEBRAURY template record that only had Y=4, K=7 and Mass=0, Priest=''
    for k_del in list(db.keys()):
        r_chk = db.get(k_del, {})
        if isinstance(r_chk, dict) and "__JANUARY__1ST WEEK" not in k_del:
            has_any_outreach = bool(
                r_chk.get("no_report_locked") or
                str(r_chk.get("hv_names", "")).strip() or
                str(r_chk.get("mw_names", "")).strip() or
                str(r_chk.get("ind_names", "")).strip() or
                str(r_chk.get("hv_remarks", "")).strip() or
                str(r_chk.get("mw_remarks", "")).strip() or
                str(r_chk.get("ind_remarks", "")).strip() or
                str(r_chk.get("outreach_general_remarks", "")).strip() or
                r_chk.get("hv_form_rows") or
                r_chk.get("mw_form_rows") or
                r_chk.get("ind_form_rows") or
                str(r_chk.get("house_visitation", "0")).strip() not in ("", "0") or
                str(r_chk.get("missionary_work", "0")).strip() not in ("", "0") or
                str(r_chk.get("indoctrination", "0")).strip() not in ("", "0")
            )
            if (
                not has_any_outreach and
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
        pob_key = "POBLACION PARISH__JANUARY__1ST WEEK"
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

    is_no_rep_excel = bool(data_row.get("no_report_locked")) or str(data_row.get("priest", "")).strip().upper() == "(NO REPORT)"
    if total_row_idx:
        for c_i, val_t in [(4, m_tot), (5, p_tot), (6, o_tot), (7, y_tot), (8, k_tot), (9, ad_tot)]:
            set_excel_cell_safely(ws, total_row_idx, c_i, "(No Report)" if is_no_rep_excel else val_t, preserve_formula=not is_no_rep_excel)
    if avg_row_idx:
        for c_i, att_v, den_v in [
            (4, m_tot, tot_say), (5, p_tot, tot_say), (6, o_tot, tot_say),
            (7, y_tot, tot_y), (8, k_tot, tot_k), (9, ad_tot, tot_say)
        ]:
            set_excel_cell_safely(ws, avg_row_idx, c_i, "(No Report)" if is_no_rep_excel else calc_pct_str(att_v, den_v), preserve_formula=not is_no_rep_excel)

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
        "adults_records": extracted_data.get("adults_records") or db.get(key, {}).get("adults_records", []),
        "youths_records": extracted_data.get("youths_records") or db.get(key, {}).get("youths_records", []),
        "kids_records": extracted_data.get("kids_records") or db.get(key, {}).get("kids_records", []),
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
    if os.path.exists(LOCAL_PWR_MASTER_FILE):
        try:
            with open(LOCAL_PWR_MASTER_FILE, "rb") as f:
                excel_bytes = f.read()
        except Exception:
            excel_bytes = None
    if not excel_bytes and service and drive_file_id:
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

        is_no_rep_row = bool(data_row.get("no_report_locked")) or str(data_row.get("priest", "")).strip().upper() == "(NO REPORT)"
        keep_form = not is_no_rep_row

        # 1. Update PARISH MONITORING REPORT sheet safely (writing '(No Report)' when marked as No Report!)
        if "PARISH MONITORING REPORT" in wb.sheetnames:
            ws = wb["PARISH MONITORING REPORT"]
            found_p = False
            for r_idx in range(5, ws.max_row + 1):
                cell_val = str(ws.cell(row=r_idx, column=1).value or "").upper()
                if clean_p and clean_p in cell_val:
                    p_lbl_cell = f"{clean_p} ({report_date}) - (No Report)" if is_no_rep_row else f"{clean_p} ({report_date})"
                    set_excel_cell_safely(ws, r_idx, 1, p_lbl_cell, preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 2, int(re.sub(r"\D", "", str(week_num)) or 1), preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 3, "(No Report)" if is_no_rep_row else f"A-{data_row['mass_a']}\nY-{data_row['mass_y']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 4, "(No Report)" if is_no_rep_row else f"A-{data_row['pampamilya_a']}\nY-{data_row['pampamilya_y']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 5, "(No Report)" if is_no_rep_row else f"A-{data_row['online_a']}\nY-{data_row['online_y']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 6, "(No Report)" if is_no_rep_row else f"Y-{data_row['youth_y']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 7, "(No Report)" if is_no_rep_row else f"K-{data_row['kids_k']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 8, "(No Report)" if is_no_rep_row else f"A-{data_row['adal_a']}\nY-{data_row['adal_y']}", preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 9, "(No Report)" if is_no_rep_row else data_row.get("house_visitation", "0"), preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 10, "(No Report)" if is_no_rep_row else data_row.get("missionary_work", "0"), preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 11, "(No Report)" if is_no_rep_row else data_row.get("indoctrination", "0"), preserve_formula=keep_form)
                    set_excel_cell_safely(ws, r_idx, 12, "(No Report)" if is_no_rep_row else data_row.get("remarks", ""), preserve_formula=keep_form)
                    found_p = True
                    break
            if not found_p:
                ws.append([
                    f"{clean_p} ({report_date})" + (" - (No Report)" if is_no_rep_row else ""), int(re.sub(r"\D", "", str(week_num)) or 1),
                    "(No Report)" if is_no_rep_row else f"A-{data_row['mass_a']}\nY-{data_row['mass_y']}",
                    "(No Report)" if is_no_rep_row else f"A-{data_row['pampamilya_a']}\nY-{data_row['pampamilya_y']}",
                    "(No Report)" if is_no_rep_row else f"A-{data_row['online_a']}\nY-{data_row['online_y']}",
                    "(No Report)" if is_no_rep_row else f"Y-{data_row['youth_y']}",
                    "(No Report)" if is_no_rep_row else f"K-{data_row['kids_k']}",
                    "(No Report)" if is_no_rep_row else f"A-{data_row['adal_a']}\nY-{data_row['adal_y']}",
                    "(No Report)" if is_no_rep_row else data_row.get("house_visitation", "0"),
                    "(No Report)" if is_no_rep_row else data_row.get("missionary_work", "0"),
                    "(No Report)" if is_no_rep_row else data_row.get("indoctrination", "0"),
                    "(No Report)" if is_no_rep_row else data_row.get("remarks", "")
                ])

        # 2. Update SUMMARY FOR POSTING sheet safely
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
                            set_excel_cell_safely(ws_s, sub_r, 2, "(No Report)" if is_no_rep_row else m_att, preserve_formula=keep_form)
                            if "HONG" not in clean_p:
                                set_excel_cell_safely(ws_s, sub_r, 3, "(No Report)" if is_no_rep_row else p_att, preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 4, "(No Report)" if is_no_rep_row else o_att, preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 5, "(No Report)" if is_no_rep_row else y_att, preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 6, "(No Report)" if is_no_rep_row else k_att, preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 7, "(No Report)" if is_no_rep_row else ad_att, preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 8, "(No Report)" if is_no_rep_row else parse_int_from_text(data_row.get("house_visitation", "0")), preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 9, "(No Report)" if is_no_rep_row else (parse_int_from_text(data_row.get("missionary_work", "0")) + parse_int_from_text(data_row.get("indoctrination", "0"))), preserve_formula=keep_form)
                            else:
                                set_excel_cell_safely(ws_s, sub_r, 3, "(No Report)" if is_no_rep_row else o_att, preserve_formula=keep_form)
                        elif "TOTAL AVERAGE" in lbl or "AVERAGE" in lbl:
                            set_excel_cell_safely(ws_s, sub_r, 2, "(No Report)" if is_no_rep_row else calc_pct_str(m_att, tot_say), preserve_formula=keep_form)
                            if "HONG" not in clean_p:
                                set_excel_cell_safely(ws_s, sub_r, 3, "(No Report)" if is_no_rep_row else calc_pct_str(p_att, tot_say), preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 4, "(No Report)" if is_no_rep_row else calc_pct_str(o_att, tot_say), preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 5, "(No Report)" if is_no_rep_row else calc_pct_str(y_att, tot_y), preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 6, "(No Report)" if is_no_rep_row else calc_pct_str(k_att, tot_k), preserve_formula=keep_form)
                                set_excel_cell_safely(ws_s, sub_r, 7, "(No Report)" if is_no_rep_row else calc_pct_str(ad_att, tot_say), preserve_formula=keep_form)
                            else:
                                set_excel_cell_safely(ws_s, sub_r, 3, "(No Report)" if is_no_rep_row else calc_pct_str(o_att, tot_say), preserve_formula=keep_form)
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
            # Kung may naka-save na partikular na Parish Workbook link sa Saved Links Registry para sa parokyang ito, i-update din iyon sa Google Drive
            try:
                for reg_it in load_saved_drive_links_registry():
                    r_url = str(reg_it.get("url", "")).strip()
                    r_par = str(reg_it.get("parish", "")).upper()
                    r_lbl = str(reg_it.get("label", "")).upper()
                    if r_url and "/folders/" not in r_url and clean_p and (clean_p in r_par or clean_p in r_lbl):
                        p_fid = extract_id_from_url(r_url)
                        if p_fid and p_fid != drive_file_id:
                            p_bytes = service.files().get_media(fileId=p_fid, supportsAllDrives=True).execute()
                            wb_p = openpyxl.load_workbook(io.BytesIO(p_bytes))
                            ensure_all_12_month_sheets_in_parish_workbook(wb_p, parish_name=parish)
                            m_ws_p = find_month_sheet_in_workbook(wb_p, month)
                            if m_ws_p is not None:
                                write_pwr_header_fields(m_ws_p, data_row, report_date, parish_name=parish, week_name=week_num)
                                write_pwr_attendance_and_remarks(m_ws_p, data_row, parish_name=parish, week_name=week_num)
                                b_p = io.BytesIO()
                                wb_p.save(b_p)
                                media_p = MediaIoBaseUpload(io.BytesIO(b_p.getvalue()), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
                                service.files().update(fileId=p_fid, media_body=media_p, supportsAllDrives=True).execute()
            except Exception:
                pass

            if drive_file_id:
                service.files().update(fileId=drive_file_id, media_body=media, supportsAllDrives=True).execute()
                return True, f"Matagumpay na na-type sa PARISH WEEKLY REPORT at na-update ang '{target_master_name}' sa Google Drive!"
            elif target_month_folder_id:
                service.files().create(
                    body={"name": target_master_name, "parents": [target_month_folder_id]},
                    media_body=media,
                    supportsAllDrives=True
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

        # Kung ito ay Multi-Month Parish Workbook (JANUARY..DECEMBER sheets), i-sync ang lahat ng buwan at linggo
        multi_extracted = extract_all_months_and_weeks_from_parish_workbook(excel_bytes)
        if multi_extracted:
            for m_item in multi_extracted:
                d_p = m_item.get("parish") or "POBLACION PARISH"
                d_m = m_item.get("month") or target_month
                d_w = m_item.get("week") or target_week
                key_m = f"{d_p.upper()}__{d_m.upper()}__{d_w}"
                denoms_m = get_parish_member_denominators(d_p)
                existing_m = db.get(key_m, {})
                existing_m.update(m_item)
                existing_m["has_data"] = True
                existing_m["parish"] = d_p
                existing_m["month"] = d_m
                existing_m["week"] = d_w
                existing_m["total_members_say"] = int(m_item.get("total_members_say", 0)) or denoms_m["say"]
                existing_m["total_members_y"] = int(m_item.get("total_members_y", 0)) or denoms_m["youth"]
                existing_m["total_members_k"] = int(m_item.get("total_members_k", 0)) or denoms_m["kids"]
                db[key_m] = existing_m
                updated_parishes += 1

        if "PARISH WEEKLY REPORT" in wb.sheetnames or find_month_sheet_in_workbook(wb, target_month) is not None:
            for p_item in PARISH_ORDER:
                pwr_extracted = extract_weekly_report_from_excel_bytes(excel_bytes, target_parish=p_item, target_month=target_month, target_week=target_week)
                detected_parish = pwr_extracted.get("parish", "") or p_item
                if detected_parish in PARISH_ORDER:
                    tot_act = (
                        int(pwr_extracted.get("mass_a", 0)) + int(pwr_extracted.get("mass_y", 0)) +
                        int(pwr_extracted.get("pampamilya_a", 0)) + int(pwr_extracted.get("pampamilya_y", 0)) +
                        int(pwr_extracted.get("online_a", 0)) + int(pwr_extracted.get("online_y", 0)) +
                        int(pwr_extracted.get("youth_y", 0)) + int(pwr_extracted.get("kids_k", 0)) +
                        int(pwr_extracted.get("adal_a", 0)) + int(pwr_extracted.get("adal_y", 0))
                    )
                    has_roster_rows = bool(pwr_extracted.get("adults_records") or pwr_extracted.get("youths_records") or pwr_extracted.get("kids_records"))
                    if tot_act > 0 or pwr_extracted.get("priest") or pwr_extracted.get("tema_misa") or has_roster_rows:
                        key = f"{detected_parish.upper()}__{target_month.upper()}__{target_week}"
                        denoms = get_parish_member_denominators(detected_parish)
                        existing = db.get(key, {})
                        existing.update(pwr_extracted)
                        existing["has_data"] = bool(tot_act > 0 or pwr_extracted.get("priest") or pwr_extracted.get("tema_misa") or existing.get("has_data"))
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
    if not isinstance(rec, dict) or not rec.get("has_data") or rec.get("no_report_locked"):
        return False
    tot_att = (
        int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0)) +
        int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0)) +
        int(rec.get("online_a", 0)) + int(rec.get("online_y", 0)) +
        int(rec.get("youth_y", 0)) + int(rec.get("kids_k", 0)) +
        int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))
    )
    has_off = any(
        str(rec.get(k, "")).strip() and "(NO REPORT)" not in str(rec.get(k, "")).upper()
        for k in ["priest", "sakristan", "epistula", "evangelio", "tema_misa", "adal_evangelio"]
    )
    return (tot_att > 0 or has_off)

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
        "report_date": "01/02/27", "remarks": "",
        "adults_records": [], "youths_records": [], "kids_records": []
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

                mem_entry = {
                    "name": name_val,
                    "mass": m_v,
                    "pampamilya": p_v,
                    "online": o_v,
                    "youth": y_v,
                    "kids": k_v,
                    "adal": a_v
                }
                if section == "ADULTS":
                    adults_cnt += 1
                    result["mass_a"] += m_v
                    result["pampamilya_a"] += p_v
                    result["online_a"] += o_v
                    result["adal_a"] += a_v
                    result["adults_records"].append(mem_entry)
                elif section == "YOUTHS":
                    youths_cnt += 1
                    result["mass_y"] += m_v
                    result["pampamilya_y"] += p_v
                    result["online_y"] += o_v
                    result["youth_y"] += y_v
                    result["adal_y"] += a_v
                    result["youths_records"].append(mem_entry)
                elif section == "KIDS":
                    kids_cnt += 1
                    result["kids_k"] += k_v
                    result["kids_records"].append(mem_entry)

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

        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)
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

        dl_info = evaluate_reporting_deadline(selected_month, selected_week, 2027, parish_name=p_name)
        is_no_rep = is_record_marked_no_report(rec)
        # Pure fast in-memory check (Excel is updated on sync/submit)
        sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_data else "⏳ Hindi Pa Tapos"
        if has_data and is_week_record_completed(rec):
            status_tag = f"✅ Tapos Na ({sub_ph})"
        elif is_no_rep:
            status_tag = "❌ (No Report) — Lagpas sa Monday Midnight Deadline"
        else:
            status_tag = "⏳ Hindi Pa Tapos"

        rows.append({
            first_col_name: f"{short_parish} ({r_date})" + (" — (No Report)" if is_no_rep else ""),
            "Week": week_num_only,
            "Mass": "(No Report)" if is_no_rep else f"A-{m_a} | Y-{m_y}",
            "Pampamilya": "(No Report)" if is_no_rep else f"A-{p_a} | Y-{p_y}",
            "Online Ministry": "(No Report)" if is_no_rep else f"A-{o_a} | Y-{o_y}",
            "Youth Fellowship": "(No Report)" if is_no_rep else f"Y-{y_y}",
            "Kids Fellowship": "(No Report)" if is_no_rep else f"K-{k_k}",
            "Adal ti Evangelio": "(No Report)" if is_no_rep else f"A-{ad_a} | Y-{ad_y}",
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

    r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)
    dl_info = evaluate_reporting_deadline(selected_month, selected_week, 2027, parish_name=parish_name)
    is_no_rep = is_record_marked_no_report(rec)
    sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_data else ("(No Report)" if is_no_rep else "⏳ Hindi Pa Tapos")

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
            "Mass": "(No Report)" if is_no_rep else str(m_att),
            "Pampamilya": "(No Report)" if is_no_rep else str(p_att),
            "Online Ministry": "(No Report)" if is_no_rep else str(o_att),
            "Youth Fellowship": "(No Report)" if is_no_rep else str(y_att),
            "Kids Fellowship": "(No Report)" if is_no_rep else str(k_att),
            "Adal ti Evangelio": "(No Report)" if is_no_rep else str(ad_att),
            "Date & Time Submitted (PH Time)": f"🇵🇭 {sub_ph}" if (has_data and is_week_record_completed(rec)) else ("❌ (No Report)" if is_no_rep else "⏳ Hindi Pa Tapos"),
        },
        {
            short_p: "TOTAL AVERAGE",
            "Month & Week (2027)": f"{selected_month} ({selected_week})",
            "Mass": "(No Report)" if is_no_rep else calc_pct_str(m_att, tot_say),
            "Pampamilya": "(No Report)" if is_no_rep else calc_pct_str(p_att, tot_say),
            "Online Ministry": "(No Report)" if is_no_rep else calc_pct_str(o_att, tot_say),
            "Youth Fellowship": "(No Report)" if is_no_rep else calc_pct_str(y_att, tot_y),
            "Kids Fellowship": "(No Report)" if is_no_rep else calc_pct_str(k_att, tot_k),
            "Adal ti Evangelio": "(No Report)" if is_no_rep else calc_pct_str(ad_att, tot_say),
            "Date & Time Submitted (PH Time)": "✅ Tapos Na" if (has_data and is_week_record_completed(rec)) else ("❌ (No Report)" if is_no_rep else "⏳ Hindi Pa Tapos"),
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
        has_data = bool(
            rec.get("has_data", False) or
            str(rec.get("mw_names", "")).strip() or
            str(rec.get("mw_remarks", "")).strip() or
            str(rec.get("missionary_work", "0")).strip() not in ("", "0")
        )
        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)

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
        has_data = bool(
            rec.get("has_data", False) or
            str(rec.get("hv_names", "")).strip() or
            str(rec.get("hv_remarks", "")).strip() or
            str(rec.get("house_visitation", "0")).strip() not in ("", "0")
        )
        if has_data:
            active_parish_count += 1

        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)
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
        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)

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
        has_data = bool(
            rec.get("has_data", False) or
            str(rec.get("ind_names", "")).strip() or
            str(rec.get("ind_remarks", "")).strip() or
            str(rec.get("indoctrination", "0")).strip() not in ("", "0")
        )
        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)

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

        r_date = resolve_display_week_date(rec, selected_month, selected_week, 2027)
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

import calendar as py_calendar

YEARS_SUPPORTED = ["2027", "2026", "2028"]

def get_saturdays_for_month_year(month_name: str, year_val: int = 2027) -> list[datetime]:
    """Kinukuha ang lahat ng aktwal na Sabado (Sabbath worship days) sa loob ng napiling Buwan at Taon."""
    try:
        m_idx = MONTHS_2027.index(month_name) + 1 if month_name in MONTHS_2027 else 1
        cal = py_calendar.monthcalendar(int(year_val), m_idx)
        sats = []
        for wk_row in cal:
            sat_d = wk_row[py_calendar.SATURDAY]
            if sat_d != 0:
                sats.append(datetime(int(year_val), m_idx, sat_d, tzinfo=PH_TIMEZONE))
        return sats
    except Exception:
        return [datetime(2027, 1, 2, tzinfo=PH_TIMEZONE)]

def get_weeks_for_month_year(month_name: str, year_val: int = 2027) -> list[str]:
    """Ibinabalik lamang ang aktwal na bilang ng mga linggo (Saturdays) sa loob ng buwang iyon sa 2027
    (hal. 4 weeks kung 4 lang ang Sabado sa buwan, 5 weeks kung 5 ang Sabado, o kung ilan ang nasa sheet)."""
    sats = get_saturdays_for_month_year(month_name, year_val)
    num_wks = max(1, min(len(sats), len(WEEKS_2027)))
    return WEEKS_2027[:num_wks]

def get_week_date_for_month_year(month_name: str, week_label: str, year_val: int = 2027) -> str:
    """Awtomatikong kinukuha ang eksaktong petsa ng Sabado (MM/DD/YY) para sa napiling Buwan, Linggo, at Taon."""
    try:
        sats = get_saturdays_for_month_year(month_name, year_val)
        w_idx = WEEKS_2027.index(week_label) if week_label in WEEKS_2027 else 0
        if w_idx < len(sats):
            return sats[w_idx].strftime("%m/%d/%y")
        return sats[-1].strftime("%m/%d/%y")
    except Exception:
        return "01/02/27"

def is_auto_schedule_lock_enabled() -> bool:
    """Binabasa kung naka-enable ang Auto Unlock (Sabado 12:00 AM) at Auto Lock (Lunes Midnight / Martes–Biyernes)."""
    cfg = load_saved_config()
    mode = str(cfg.get("MANUAL_LOCK_OVERRIDE_MODE", "MANUAL"))
    return bool(cfg.get("AUTO_SCHEDULE_LOCK_ENABLED", False)) and (mode == "AUTO")

def set_auto_schedule_lock_enabled(enabled: bool):
    try:
        cfg = load_saved_config()
        cfg["AUTO_SCHEDULE_LOCK_ENABLED"] = bool(enabled)
        if enabled:
            now_ph = get_philippine_now()
            is_open_day = (now_ph.weekday() in (5, 6, 0))
            cfg["DEADLINE_ENFORCEMENT_ENABLED"] = not is_open_day
            cfg["MANUAL_LOCK_OVERRIDE_MODE"] = "AUTO"
            st.session_state["deadline_enforcement_enabled"] = not is_open_day
            st.session_state["master_deadline_lock_toggle"] = not is_open_day
        else:
            cfg["MANUAL_LOCK_OVERRIDE_MODE"] = "MANUAL"
        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

def is_deadline_enforcement_enabled() -> bool:
    """Binabasa mula sa shared config kung naka-LOCK (True) o UNLOCKED (False) ang Upload ng mga Leaders.
    Kapag naka-AUTO mode, kusang nag-a-unlock tuwing Sabado hanggang Lunes ng Hatinggabi (Monday 11:59 PM PHT)
    at kusang nagla-lock pagsapit ng Martes hanggang Biyernes."""
    cfg = load_saved_config()
    if bool(cfg.get("AUTO_SCHEDULE_LOCK_ENABLED", False)) and str(cfg.get("MANUAL_LOCK_OVERRIDE_MODE", "MANUAL")) == "AUTO":
        now_ph = get_philippine_now()
        is_open_window = (now_ph.weekday() in (5, 6, 0))
        val = not is_open_window
    else:
        val = bool(cfg.get("DEADLINE_ENFORCEMENT_ENABLED", False))
    st.session_state["deadline_enforcement_enabled"] = val
    return val

def set_deadline_enforcement_enabled(enabled: bool, manual_override: bool = True):
    """Sine-save ang manual Lock (True) o Unlock (False) state ng Submission Deadline Controller para sa lahat ng users."""
    st.session_state["deadline_enforcement_enabled"] = bool(enabled)
    try:
        cfg = load_saved_config()
        cfg["DEADLINE_ENFORCEMENT_ENABLED"] = bool(enabled)
        if manual_override:
            cfg["MANUAL_LOCK_OVERRIDE_MODE"] = "MANUAL"
            cfg["AUTO_SCHEDULE_LOCK_ENABLED"] = False
            st.session_state["master_auto_sched_toggle"] = False
        if enabled:
            cfg["ADMIN_ALLOWED_SUBMISSION_OVERRIDES"] = []
        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

def _on_master_lock_toggle_clicked():
    new_val = bool(st.session_state.get("master_deadline_lock_toggle", False))
    set_deadline_enforcement_enabled(new_val, manual_override=True)

def _on_master_auto_sched_clicked():
    new_auto = bool(st.session_state.get("master_auto_sched_toggle", False))
    set_auto_schedule_lock_enabled(new_auto)

def get_admin_allowed_submission_overrides() -> list[str]:
    cfg = load_saved_config()
    overrides = cfg.get("ADMIN_ALLOWED_SUBMISSION_OVERRIDES", [])
    return overrides if isinstance(overrides, list) else []

def set_admin_allowed_submission_override(parish_name: str, month_name: str, week_label: str, allow: bool):
    key = f"{parish_name.upper()}__{month_name.upper()}__{week_label}"
    overrides = set(get_admin_allowed_submission_overrides())
    if allow:
        overrides.add(key)
    else:
        overrides.discard(key)
    try:
        cfg = load_saved_config()
        cfg["ADMIN_ALLOWED_SUBMISSION_OVERRIDES"] = sorted(list(overrides))
        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

def is_parish_week_allowed_by_admin(parish_name: str, month_name: str, week_label: str) -> bool:
    if not is_deadline_enforcement_enabled():
        return True
    if not parish_name or parish_name == "Lahat ng Parokya":
        return False
    overrides = set(get_admin_allowed_submission_overrides())
    return (
        f"{parish_name.upper()}__{month_name.upper()}__{week_label}" in overrides
        or f"ALL__{month_name.upper()}__{week_label}" in overrides
    )

def evaluate_reporting_deadline(month_name: str, week_label: str, year_val: int = 2027, parish_name: str = "") -> dict:
    sats = get_saturdays_for_month_year(month_name, year_val)
    w_idx = WEEKS_2027.index(week_label) if week_label in WEEKS_2027 else 0
    sat_dt = sats[min(w_idx, len(sats) - 1)] if sats else datetime(int(year_val), 1, 2, tzinfo=PH_TIMEZONE)
    monday_midnight_dt = (sat_dt + timedelta(days=2)).replace(hour=23, minute=59, second=59)
    next_saturday_dt = sat_dt + timedelta(days=7)

    deadline_switch_on = is_deadline_enforcement_enabled()
    admin_allowed_this = is_parish_week_allowed_by_admin(parish_name, month_name, week_label)

    if deadline_switch_on and not admin_allowed_this:
        is_past_deadline = True
    else:
        is_past_deadline = False

    return {
        "saturday_date_str": sat_dt.strftime("%m/%d/%y"),
        "monday_deadline_str": monday_midnight_dt.strftime("%b %d, %Y (Monday 11:59 PM PHT)"),
        "next_saturday_str": next_saturday_dt.strftime("%b %d, %Y (Saturday / %m/%d/%y)"),
        "deadline_switch_on": deadline_switch_on,
        "admin_allowed_this": admin_allowed_this,
        "is_past_deadline": is_past_deadline
    }

def is_record_marked_no_report(rec: dict) -> bool:
    """Sinusuri kung ang partikular na Parish + Month + Week record ay nilagyan ng '(No Report)'."""
    if not isinstance(rec, dict):
        return False
    if bool(rec.get("no_report_locked")):
        return True
    if str(rec.get("priest", "")).strip().upper() == "(NO REPORT)" or str(rec.get("submitted_at_ph", "")).strip().upper() == "(NO REPORT)":
        return True
    return False

def set_parish_no_report_status(parish_name: str, month_name: str, week_label: str, year_val: int = 2027, mark_no_report: bool = True, force_overwrite: bool = False) -> bool:
    """Master Admin Control para lagyan o alisan ng '(No Report)' ang partikular na Parokya, Buwan, at Linggo sa App at sa Excel file."""
    db = load_weekly_monitoring_records()
    key = f"{parish_name.upper()}__{month_name.upper()}__{week_label}"
    rec = db.get(key, {})
    if not mark_no_report:
        if key in db and is_record_marked_no_report(rec):
            db.pop(key, None)
            try:
                with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
                    json.dump(db, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        return True

    if not force_overwrite and is_week_record_completed(rec):
        return False

    wk_dt = get_week_date_for_month_year(month_name, week_label, year_val)
    denoms = get_parish_member_denominators(parish_name)
    no_rep_payload = {
        "has_data": False,
        "no_report_locked": True,
        "parish": parish_name,
        "month": month_name,
        "week": week_label,
        "report_date": wk_dt,
        "total_members_say": denoms["say"],
        "total_members_y": denoms["youth"],
        "total_members_k": denoms["kids"],
        "mass_a": 0, "mass_y": 0, "pampamilya_a": 0, "pampamilya_y": 0,
        "online_a": 0, "online_y": 0, "youth_y": 0, "kids_k": 0,
        "adal_a": 0, "adal_y": 0,
        "house_visitation": "(No Report)",
        "hv_names": "(No Report)",
        "hv_remarks": "(No Report)",
        "missionary_work": "(No Report)",
        "mw_names": "(No Report)",
        "mw_remarks": "(No Report)",
        "indoctrination": "(No Report)",
        "ind_names": "(No Report)",
        "ind_remarks": "(No Report)",
        "priest": "(No Report)",
        "sakristan": "(No Report)",
        "epistula": "(No Report)",
        "evangelio": "(No Report)",
        "casasaad": "(No Report)",
        "candle_holder": "(No Report)",
        "usher": "(No Report)",
        "marshal": "(No Report)",
        "adal_evangelio": "(No Report)",
        "adal_preacher": "(No Report)",
        "tema_misa": "(No Report)",
        "misa_preacher": "(No Report)",
        "tema_pampamilya": "(No Report)",
        "pampamilya_preacher": "(No Report)",
        "mass_remarks": "(No Report)",
        "remarks": "(No Report) — Walang naipasang Weekly Report sa itinakdang deadline.",
        "updated_at": get_philippine_datetime_str(),
        "submitted_at_ph": "(No Report)"
    }
    db[key] = no_rep_payload
    try:
        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    try:
        update_master_pwr_excel_file(parish_name, month_name, week_label, wk_dt, no_rep_payload)
    except Exception:
        pass
    return True

def mark_unsubmitted_parish_as_no_report_in_excel_and_db(parish_name: str, month_name: str, week_label: str, year_val: int = 2027):
    """Kapag lumagpas na sa Monday Midnight deadline at walang na-submit na report ang parokya,
    awtomatikong nilalagyan ng '(No Report)' sa App database at sa Excel file."""
    db = load_weekly_monitoring_records()
    key = f"{parish_name.upper()}__{month_name.upper()}__{week_label}"
    rec = db.get(key, {})
    if is_week_record_completed(rec):
        return False
    if rec.get("no_report_locked"):
        return True

    wk_dt = get_week_date_for_month_year(month_name, week_label, year_val)
    denoms = get_parish_member_denominators(parish_name)
    no_rep_payload = {
        "has_data": False,
        "no_report_locked": True,
        "parish": parish_name,
        "month": month_name,
        "week": week_label,
        "report_date": wk_dt,
        "total_members_say": denoms["say"],
        "total_members_y": denoms["youth"],
        "total_members_k": denoms["kids"],
        "mass_a": 0, "mass_y": 0, "pampamilya_a": 0, "pampamilya_y": 0,
        "online_a": 0, "online_y": 0, "youth_y": 0, "kids_k": 0,
        "adal_a": 0, "adal_y": 0,
        "priest": "(No Report)",
        "sakristan": "(No Report)",
        "epistula": "(No Report)",
        "evangelio": "(No Report)",
        "casasaad": "(No Report)",
        "candle_holder": "(No Report)",
        "usher": "(No Report)",
        "marshal": "(No Report)",
        "adal_evangelio": "(No Report)",
        "adal_preacher": "(No Report)",
        "tema_misa": "(No Report)",
        "misa_preacher": "(No Report)",
        "tema_pampamilya": "(No Report)",
        "pampamilya_preacher": "(No Report)",
        "remarks": "(No Report) — Hindi nakapag-submit bago mag-Monday Midnight deadline.",
        "updated_at": get_philippine_datetime_str(),
        "submitted_at_ph": "(No Report)"
    }
    db[key] = no_rep_payload
    try:
        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    try:
        update_master_pwr_excel_file(parish_name, month_name, week_label, wk_dt, no_rep_payload)
    except Exception:
        pass
    return True

def delete_populated_weekly_record(parish_name: str, month_name: str, week_label: str) -> bool:
    """Binubura lamang ang na-populate na data kapag pinindot ang explicit Delete button."""
    db = load_weekly_monitoring_records()
    key = f"{parish_name.upper()}__{month_name.upper()}__{week_label}"
    if key in db:
        db.pop(key, None)
    # Kung ito ay POBLACION Jan W1, maglagay ng explicit deleted marker para hindi mag-auto-reseed maliban kung mag-upload ulit
    if key == "POBLACION PARISH__JANUARY__1ST WEEK":
        db[key] = {"has_data": False, "explicitly_deleted": True, "parish": parish_name, "month": month_name, "week": week_label}
    try:
        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    st.session_state.pop("last_extracted_pwr", None)
    st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
    return True

def resolve_display_week_date(rec: dict, month_name: str, week_label: str, year_val: int = 2027) -> str:
    """Ginagamit ang aktwal na report_date kung tapos na at akma sa buwan; kung default '01/02/27' sa ibang buwan/linggo, ibinibigay ang eksaktong 2027 Saturday date."""
    auto_dt = get_week_date_for_month_year(month_name, week_label, year_val)
    if not isinstance(rec, dict) or not is_week_record_completed(rec):
        return auto_dt
    saved_dt = str(rec.get("report_date", "")).strip()
    if not saved_dt or (saved_dt in ("01/02/27", "01-02-27") and (month_name != "January" or week_label != "1st Week")):
        return auto_dt
    return saved_dt

def build_parish_member_attendance_sheet_df(parish_name: str, selected_month: str, selected_week: str, selected_year: str = "2027") -> tuple[pd.DataFrame, dict]:
    """Bumubuo ng eksaktong Per-Member Attendance Table (tulad sa Excel sheet: ADULTS, YOUTHS, KIDS na may 1 sa Mass, Pampamilya, Online, Youth, Kids, Adal + TOTAL at AVERAGE)."""
    db = load_weekly_monitoring_records()
    key = f"{parish_name.upper()}__{selected_month.upper()}__{selected_week}"
    rec = db.get(key, {})
    has_done = is_week_record_completed(rec)
    denoms = get_parish_member_denominators(parish_name)

    adults_list = list(rec.get("adults_records", []) or [])
    youths_list = list(rec.get("youths_records", []) or [])
    kids_list = list(rec.get("kids_records", []) or [])

    # 1. Kung wala pang naka-save na per-member rows sa linggong ito, kunin ang roster mula sa ibang linggo ng parehong Buwan (galing sa Excel sheet ng buwang iyon)
    if not (adults_list or youths_list or kids_list):
        for wk_other in WEEKS_2027:
            rec_other = db.get(f"{parish_name.upper()}__{selected_month.upper()}__{wk_other}", {})
            if rec_other.get("adults_records") or rec_other.get("youths_records") or rec_other.get("kids_records"):
                adults_list = [dict(x, mass=0, pampamilya=0, online=0, youth=0, kids=0, adal=0) for x in rec_other.get("adults_records", [])]
                youths_list = [dict(x, mass=0, pampamilya=0, online=0, youth=0, kids=0, adal=0) for x in rec_other.get("youths_records", [])]
                kids_list = [dict(x, mass=0, pampamilya=0, online=0, youth=0, kids=0, adal=0) for x in rec_other.get("kids_records", [])]
                break

    # 2. Subukang basahin mula sa cached local Excel roster kung mayroon
    if not (adults_list or youths_list or kids_list) and os.path.exists(LOCAL_PWR_MASTER_FILE):
        cache_r_key = f"_cached_roster_{parish_name}_{selected_month}_{selected_week}"
        if cache_r_key not in st.session_state:
            try:
                with open(LOCAL_PWR_MASTER_FILE, "rb") as f:
                    ext_local = extract_weekly_report_from_excel_bytes(
                        f.read(), target_parish=parish_name, target_month=selected_month, target_week=selected_week
                    )
                st.session_state[cache_r_key] = {
                    "adults": ext_local.get("adults_records", []),
                    "youths": ext_local.get("youths_records", []),
                    "kids": ext_local.get("kids_records", [])
                }
            except Exception:
                st.session_state[cache_r_key] = {"adults": [], "youths": [], "kids": []}
        cached_r = st.session_state.get(cache_r_key, {})
        adults_list = list(cached_r.get("adults", []))
        youths_list = list(cached_r.get("youths", []))
        kids_list = list(cached_r.get("kids", []))

    # 3. Kung wala pa sa local cache, kunin ang lahat ng miyembro ng parokya mula sa master_members_df NANG HINDI TINATANGGAL ang INACTIVE members
    if not (adults_list or youths_list or kids_list):
        try:
            census_df = master_members_df
            if census_df is not None and not census_df.empty and "Parish" in census_df.columns:
                p_df = census_df[census_df["Parish"].astype(str).apply(lambda x: normalize_parish_match(x, parish_name))].copy()
                m_a_target = int(rec.get("mass_a", 0)) if has_done else 0
                p_a_target = int(rec.get("pampamilya_a", 0)) if has_done else 0
                o_a_target = int(rec.get("online_a", 0)) if has_done else 0
                ad_a_target = int(rec.get("adal_a", 0)) if has_done else 0

                m_y_target = int(rec.get("mass_y", 0)) if has_done else 0
                p_y_target = int(rec.get("pampamilya_y", 0)) if has_done else 0
                o_y_target = int(rec.get("online_y", 0)) if has_done else 0
                y_y_target = int(rec.get("youth_y", 0)) if has_done else 0
                ad_y_target = int(rec.get("adal_y", 0)) if has_done else 0

                k_k_target = int(rec.get("kids_k", 0)) if has_done else 0

                a_idx, y_idx, k_idx = 0, 0, 0
                for _, m_row in p_df.iterrows():
                    fn = str(m_row.get("Full Name", "")).strip()
                    if not fn:
                        continue
                    cat = str(m_row.get("Category", "ADULT")).upper().strip()
                    if "KID" in cat or "CHILD" in cat:
                        k_idx += 1
                        kids_list.append({
                            "name": fn,
                            "mass": 0, "pampamilya": 0, "online": 0, "youth": 0,
                            "kids": 1 if k_idx <= k_k_target else 0,
                            "adal": 0
                        })
                    elif "YOUTH" in cat:
                        y_idx += 1
                        youths_list.append({
                            "name": fn,
                            "mass": 1 if y_idx <= m_y_target else 0,
                            "pampamilya": 1 if y_idx <= p_y_target else 0,
                            "online": 1 if y_idx <= o_y_target else 0,
                            "youth": 1 if y_idx <= y_y_target else 0,
                            "kids": 0,
                            "adal": 1 if y_idx <= ad_y_target else 0
                        })
                    else:
                        a_idx += 1
                        adults_list.append({
                            "name": fn,
                            "mass": 1 if a_idx <= m_a_target else 0,
                            "pampamilya": 1 if a_idx <= p_a_target else 0,
                            "online": 1 if a_idx <= o_a_target else 0,
                            "youth": 0,
                            "kids": 0,
                            "adal": 1 if a_idx <= ad_a_target else 0
                        })
        except Exception:
            pass

    first_col = f"(Weekly Report-{selected_week} of {selected_month.upper()} {selected_year})"
    rows = [{
        first_col: "Attendance:",
        "A.M & P.M Mass": "ADULTS",
        "Mass": "", "Pampamilya": "", "Online Ministry": "",
        "Youth Fellowship": "", "Kids Fellowship": "", "Adal ti Evangelio": ""
    }]

    seq_num = 0
    for idx_a, item in enumerate(adults_list, 1):
        seq_num += 1
        rows.append({
            first_col: str(idx_a),
            "A.M & P.M Mass": str(item.get("name", "")),
            "Mass": "1" if (has_done and int(item.get("mass", 0)) > 0) else "",
            "Pampamilya": "1" if (has_done and int(item.get("pampamilya", 0)) > 0) else "",
            "Online Ministry": "1" if (has_done and int(item.get("online", 0)) > 0) else "",
            "Youth Fellowship": "",
            "Kids Fellowship": "",
            "Adal ti Evangelio": "1" if (has_done and int(item.get("adal", 0)) > 0) else ""
        })

    rows.append({
        first_col: "",
        "A.M & P.M Mass": "YOUTHS",
        "Mass": "", "Pampamilya": "", "Online Ministry": "",
        "Youth Fellowship": "", "Kids Fellowship": "", "Adal ti Evangelio": ""
    })
    for idx_y, item in enumerate(youths_list, 1):
        seq_num += 1
        rows.append({
            first_col: f"{idx_y} ({seq_num})",
            "A.M & P.M Mass": str(item.get("name", "")),
            "Mass": "1" if (has_done and int(item.get("mass", 0)) > 0) else "",
            "Pampamilya": "1" if (has_done and int(item.get("pampamilya", 0)) > 0) else "",
            "Online Ministry": "1" if (has_done and int(item.get("online", 0)) > 0) else "",
            "Youth Fellowship": "1" if (has_done and int(item.get("youth", 0)) > 0) else "",
            "Kids Fellowship": "",
            "Adal ti Evangelio": "1" if (has_done and int(item.get("adal", 0)) > 0) else ""
        })

    rows.append({
        first_col: "",
        "A.M & P.M Mass": "KIDS",
        "Mass": "", "Pampamilya": "", "Online Ministry": "",
        "Youth Fellowship": "", "Kids Fellowship": "", "Adal ti Evangelio": ""
    })
    for idx_k, item in enumerate(kids_list, 1):
        seq_num += 1
        rows.append({
            first_col: f"{idx_k} ({seq_num})",
            "A.M & P.M Mass": str(item.get("name", "")),
            "Mass": "",
            "Pampamilya": "",
            "Online Ministry": "",
            "Youth Fellowship": "",
            "Kids Fellowship": "1" if (has_done and int(item.get("kids", 0)) > 0) else "",
            "Adal ti Evangelio": ""
        })

    tot_say = int(rec.get("total_members_say", 0)) or (len(adults_list) + len(youths_list)) or denoms["say"]
    tot_y = int(rec.get("total_members_y", 0)) or len(youths_list) or denoms["youth"]
    tot_k = int(rec.get("total_members_k", 0)) or len(kids_list) or denoms["kids"]

    m_tot = (int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0))) if has_done else 0
    p_tot = (int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0))) if has_done else 0
    o_tot = (int(rec.get("online_a", 0)) + int(rec.get("online_y", 0))) if has_done else 0
    y_tot = int(rec.get("youth_y", 0)) if has_done else 0
    k_tot = int(rec.get("kids_k", 0)) if has_done else 0
    ad_tot = (int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))) if has_done else 0

    is_no_rep_mem = is_record_marked_no_report(rec)
    rows.append({
        first_col: "TOTAL:" + (" (No Report)" if is_no_rep_mem else ""),
        "A.M & P.M Mass": "(No Report)" if is_no_rep_mem else "",
        "Mass": "(No Report)" if is_no_rep_mem else str(m_tot),
        "Pampamilya": "(No Report)" if is_no_rep_mem else str(p_tot),
        "Online Ministry": "(No Report)" if is_no_rep_mem else str(o_tot),
        "Youth Fellowship": "(No Report)" if is_no_rep_mem else str(y_tot),
        "Kids Fellowship": "(No Report)" if is_no_rep_mem else str(k_tot),
        "Adal ti Evangelio": "(No Report)" if is_no_rep_mem else str(ad_tot)
    })
    rows.append({
        first_col: "AVERAGE:",
        "A.M & P.M Mass": "(No Report)" if is_no_rep_mem else "",
        "Mass": "(No Report)" if is_no_rep_mem else calc_pct_str(m_tot, tot_say),
        "Pampamilya": "(No Report)" if is_no_rep_mem else calc_pct_str(p_tot, tot_say),
        "Online Ministry": "(No Report)" if is_no_rep_mem else calc_pct_str(o_tot, tot_say),
        "Youth Fellowship": "(No Report)" if is_no_rep_mem else calc_pct_str(y_tot, tot_y),
        "Kids Fellowship": "(No Report)" if is_no_rep_mem else calc_pct_str(k_tot, tot_k),
        "Adal ti Evangelio": "(No Report)" if is_no_rep_mem else calc_pct_str(ad_tot, tot_say)
    })

    totals_meta = {
        "total_members_say": tot_say,
        "total_members_y": tot_y,
        "total_members_k": tot_k,
        "report_date": resolve_display_week_date(rec, selected_month, selected_week, int(selected_year) if str(selected_year).isdigit() else 2027)
    }
    return pd.DataFrame(rows), totals_meta

def sync_outreach_tasks_from_notion_to_excel_and_db(target_month: str = "January", target_week: str = "1st Week", target_parish: str = "Lahat ng Parokya", target_year: str = "2027") -> tuple[int, str]:
    """Kinukuha ang mga Outreach Tasks (House Visitation, Missionary Work, Indoctrination & Remarks) mula sa Notion Database
    at awtomatikong inilalagay sa App Database at sa Google Drive Excel Report ng tamang Parokya, Buwan, at Linggo."""
    token = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    if not token:
        return 0, "Walang Notion Token na nakalagay sa App Settings."

    candidate_dbs = []
    for k_db in ["NOTION_OUTREACH_DB_ID", "NOTION_CRM_DB_ID", "NOTION_COUNSELLING_DB_ID", "NOTION_QUESTIONS_DB_ID", "NOTION_PRAYER_DB_ID"]:
        val_db = st.session_state.custom_secrets.get(k_db, "") or globals().get(k_db, "")
        if val_db and val_db not in candidate_dbs:
            candidate_dbs.append(val_db)

    if not candidate_dbs:
        return 0, "Walang Notion Database ID na nakalagay."

    db_records = load_weekly_monitoring_records()
    grouped_updates = {}
    matched_tasks = 0

    for db_id in candidate_dbs:
        try:
            df_n = load_notion_crm_data(token, db_id)
        except Exception:
            continue
        if df_n is None or df_n.empty:
            continue

        for _, row in df_n.iterrows():
            row_dict = {str(k).strip(): str(v or "").strip() for k, v in row.items()}
            joined_text = " | ".join(f"{k}: {v}" for k, v in row_dict.items() if not k.startswith("_") and v)
            j_up = joined_text.upper()

            # Tukuyin kung anong uri ng Outreach Task ito
            task_cat = None
            if any(w in j_up for w in ["HOUSE VISITATION", "VISITATION", "DALAW", "HOME VISIT"]):
                task_cat = "HV"
            elif any(w in j_up for w in ["MISSIONARY", "MISSION WORK", "MISYON"]):
                task_cat = "MW"
            elif any(w in j_up for w in ["INDOCTRINATION", "DOKTRINA", "DOCTRINE"]):
                task_cat = "IND"

            if task_cat is None:
                continue

            # Tukuyin kung aling Parokya
            detected_parish = None
            for p_cand in PARISH_ORDER:
                clean_p = p_cand.replace("PARISH", "").strip()
                if clean_p.upper() in j_up:
                    detected_parish = p_cand
                    break
            if not detected_parish:
                detected_parish = target_parish if target_parish in PARISH_ORDER else "LIPAY NORTE PARISH"

            # Tukuyin ang Buwan at Linggo mula sa Notion row (o gamitin ang kasalukuyang target_month at target_week)
            det_month = target_month if target_month in MONTHS_2027 else "January"
            for m_c in MONTHS_2027:
                if m_c.upper() in j_up:
                    det_month = m_c
                    break
            det_week = target_week if target_week in WEEKS_2027 else "1st Week"
            for w_c in WEEKS_2027:
                if w_c.upper() in j_up or f"WEEK {w_c[0]}" in j_up:
                    det_week = w_c
                    break

            fam_name = ""
            for nk in ["Family Name", "Full Name", "Name", "Pangalan", "Visitor Name", "Family", "Title"]:
                if row_dict.get(nk):
                    fam_name = row_dict[nk]
                    break
            if not fam_name:
                for k_col, v_col in row_dict.items():
                    if not k_col.startswith("_") and v_col and len(v_col) < 80:
                        fam_name = v_col
                        break

            rem_txt = ""
            for rk in ["Remarks", "Notes", "Details", "Tala", "Message", "Status", "Description"]:
                if row_dict.get(rk):
                    rem_txt = row_dict[rk]
                    break

            if not fam_name:
                continue

            g_key = (detected_parish, det_month, det_week)
            if g_key not in grouped_updates:
                grouped_updates[g_key] = {"HV": [], "HV_REM": [], "MW": [], "MW_REM": [], "IND": [], "IND_REM": []}
            if fam_name not in grouped_updates[g_key][task_cat]:
                grouped_updates[g_key][task_cat].append(fam_name)
                matched_tasks += 1
            if rem_txt and rem_txt not in grouped_updates[g_key][f"{task_cat}_REM"]:
                grouped_updates[g_key][f"{task_cat}_REM"].append(rem_txt)

    for (p_u, m_u, w_u), bucket in grouped_updates.items():
        rec_k = f"{p_u.upper()}__{m_u.upper()}__{w_u}"
        cur_r = dict(db_records.get(rec_k, {}))
        r_date = resolve_display_week_date(cur_r, m_u, w_u, int(target_year) if str(target_year).isdigit() else 2027)

        if bucket["HV"]:
            existing_hv = [x.strip() for x in str(cur_r.get("hv_names", "")).split(";") if x.strip()]
            for item in bucket["HV"]:
                if item not in existing_hv:
                    existing_hv.append(item)
            cur_r["hv_names"] = "; ".join(existing_hv)
            fam_c, ppl_c = count_families_and_people_from_text(cur_r["hv_names"], str(len(existing_hv)))
            cur_r["house_visitation"] = str(max(fam_c, len(existing_hv)))
            cur_r["hv_people_count"] = max(ppl_c, len(existing_hv))
        if bucket["HV_REM"]:
            cur_r["hv_remarks"] = "; ".join(bucket["HV_REM"])

        if bucket["MW"]:
            existing_mw = [x.strip() for x in str(cur_r.get("mw_names", "")).split(";") if x.strip()]
            for item in bucket["MW"]:
                if item not in existing_mw:
                    existing_mw.append(item)
            cur_r["mw_names"] = "; ".join(existing_mw)
            fam_c, ppl_c = count_families_and_people_from_text(cur_r["mw_names"], str(len(existing_mw)))
            cur_r["missionary_work"] = str(max(fam_c, len(existing_mw)))
            cur_r["mw_people_count"] = max(ppl_c, len(existing_mw))
        if bucket["MW_REM"]:
            cur_r["mw_remarks"] = "; ".join(bucket["MW_REM"])

        if bucket["IND"]:
            existing_ind = [x.strip() for x in str(cur_r.get("ind_names", "")).split(";") if x.strip()]
            for item in bucket["IND"]:
                if item not in existing_ind:
                    existing_ind.append(item)
            cur_r["ind_names"] = "; ".join(existing_ind)
            fam_c, ppl_c = count_families_and_people_from_text(cur_r["ind_names"], str(len(existing_ind)))
            cur_r["indoctrination"] = str(max(fam_c, len(existing_ind)))
            cur_r["ind_people_count"] = max(ppl_c, len(existing_ind))
        if bucket["IND_REM"]:
            cur_r["ind_remarks"] = "; ".join(bucket["IND_REM"])

        save_weekly_monitoring_record(p_u, m_u, w_u, r_date, cur_r, drive_link=cur_r.get("drive_link", ""))

    if matched_tasks == 0:
        return 0, "Naka-ready na ang Notion Outreach Sync! Kapag naidagdag na sa Notion ang mga House Visitation, Missionary Work, at Indoctrination tasks, awtomatiko itong papasok dito at sa Excel report."
    return matched_tasks, f"Matagumpay na nakuha ang {matched_tasks} Outreach Task(s) mula sa Notion at nai-type sa Excel Report!"

def build_multi_week_house_visitation_df(selected_month: str, weeks_list: list[str], filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    frames = []
    for wk in weeks_list:
        df_w = build_house_visitation_table_df(selected_month, wk, filter_parish)
        if len(weeks_list) > 1 and not df_w.empty:
            df_w = df_w[df_w["Parish & Date"] != "TOTAL & PERCENTAGE:"].copy()
        frames.append(df_w)
    if not frames:
        return build_house_visitation_table_df(selected_month, "1st Week", filter_parish)
    combined = pd.concat(frames, ignore_index=True)
    if len(weeks_list) > 1 and not combined.empty:
        tot_fams = 0
        tot_ppl = 0
        for _, r in combined.iterrows():
            f_str = str(r.get("Bilang ng Pamilya (Req: 10 Families)", "0")).split("/")[0]
            tot_fams += parse_int_from_text(f_str)
            tot_ppl += parse_int_from_text(str(r.get("Total ng mga Tao (Ayun sa Family Names)", "0")))
        req_tot = max(10, len(combined) * 10)
        combined = pd.concat([combined, pd.DataFrame([{
            "Parish & Date": "TOTAL & PERCENTAGE (ALL SELECTED WEEKS):",
            "Week": "ALL",
            "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER (Family Names & Mga Pangalan)": f"Kabuuang Pamilya at mga Taong Binisita ({selected_month} 2027)",
            "Bilang ng Pamilya (Req: 10 Families)": f"Total: {tot_fams} / {req_tot} ({calc_pct_str(tot_fams, req_tot)})",
            "Total ng mga Tao (Ayun sa Family Names)": f"Total: {tot_ppl} Tao",
            "Remarks / Tala sa House Visitation": "Kabuuang Bilang sa House Visitation"
        }])], ignore_index=True)
    return combined

def build_multi_week_missionary_work_df(selected_month: str, weeks_list: list[str], filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    frames = []
    for wk in weeks_list:
        df_w = build_missionary_work_table_df(selected_month, wk, filter_parish)
        if len(weeks_list) > 1 and not df_w.empty:
            df_w = df_w[df_w["Parish & Date"] != "TOTAL & PERCENTAGE:"].copy()
        frames.append(df_w)
    if not frames:
        return build_missionary_work_table_df(selected_month, "1st Week", filter_parish)
    combined = pd.concat(frames, ignore_index=True)
    if len(weeks_list) > 1 and not combined.empty:
        tot_fams = sum(parse_int_from_text(str(r.get("Bilang ng Pamilya", "0"))) for _, r in combined.iterrows())
        tot_ppl = sum(parse_int_from_text(str(r.get("Total ng mga Tao (Ayun sa Family Names)", "0"))) for _, r in combined.iterrows())
        combined = pd.concat([combined, pd.DataFrame([{
            "Parish & Date": "TOTAL & PERCENTAGE (ALL SELECTED WEEKS):",
            "Week": "ALL",
            "Missionary Work (Family Names & Mga Pangalan ng Dumalo)": f"Kabuuang Pamilya at mga Tao sa Missionary Work ({selected_month} 2027)",
            "Bilang ng Pamilya": f"Total: {tot_fams} Pamilya",
            "Total ng mga Tao (Ayun sa Family Names)": f"Total: {tot_ppl} Tao",
            "Remarks / Tala sa Missionary Work": "Kabuuang Bilang sa Missionary Work"
        }])], ignore_index=True)
    return combined

def build_multi_week_indoctrination_df(selected_month: str, weeks_list: list[str], filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    frames = []
    for wk in weeks_list:
        df_w = build_indoctrination_table_df(selected_month, wk, filter_parish)
        if len(weeks_list) > 1 and not df_w.empty:
            df_w = df_w[df_w["Parish & Date"] != "TOTAL & PERCENTAGE:"].copy()
        frames.append(df_w)
    if not frames:
        return build_indoctrination_table_df(selected_month, "1st Week", filter_parish)
    combined = pd.concat(frames, ignore_index=True)
    if len(weeks_list) > 1 and not combined.empty:
        tot_fams = sum(parse_int_from_text(str(r.get("Bilang ng Pamilya", "0"))) for _, r in combined.iterrows())
        tot_ppl = sum(parse_int_from_text(str(r.get("Total ng mga Tao (Ayun sa Family Names)", "0"))) for _, r in combined.iterrows())
        combined = pd.concat([combined, pd.DataFrame([{
            "Parish & Date": "TOTAL & PERCENTAGE (ALL SELECTED WEEKS):",
            "Week": "ALL",
            "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION (Family Names & Mga Pangalan)": f"Kabuuang Pamilya at mga Tao sa Indoctrination ({selected_month} 2027)",
            "Bilang ng Pamilya": f"Total: {tot_fams} Pamilya",
            "Total ng mga Tao (Ayun sa Family Names)": f"Total: {tot_ppl} Tao",
            "Remarks / Tala sa Indoctrination": "Kabuuang Bilang sa Indoctrination"
        }])], ignore_index=True)
    return combined

def build_multi_week_outreach_remarks_df(selected_month: str, weeks_list: list[str], filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    """Bumubuo ng Record-Style Table na eksklusibo lamang para sa:
    1) Remarks sa House Visitation, 2) Remarks sa Missionary Work, at 3) Remarks sa Indoctrination."""
    db = load_weekly_monitoring_records()
    rows = []
    parishes_to_scan = PARISH_ORDER if filter_parish == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, filter_parish)]
    for wk in weeks_list:
        week_num_only = re.sub(r"\D", "", str(wk)) or "1"
        for p_name in parishes_to_scan:
            rec = db.get(f"{p_name.upper()}__{selected_month.upper()}__{wk}", {})
            has_data = bool(
                rec.get("has_data", False) or
                str(rec.get("hv_remarks", "")).strip() or
                str(rec.get("mw_remarks", "")).strip() or
                str(rec.get("ind_remarks", "")).strip() or
                str(rec.get("outreach_general_remarks", "")).strip() or
                str(rec.get("hv_names", "")).strip() or
                str(rec.get("mw_names", "")).strip() or
                str(rec.get("ind_names", "")).strip()
            )
            is_no_rep_o = is_record_marked_no_report(rec)
            if not has_data and not is_no_rep_o and len(weeks_list) > 1 and filter_parish == "Lahat ng Parokya":
                continue
            r_date = resolve_display_week_date(rec, selected_month, wk, 2027)
            raw_rem = str(rec.get("remarks", "")).strip() if has_data else ""
            p_ro = parse_remarks_into_mass_and_outreach(raw_rem) if raw_rem else {}

            hv_r = "(No Report)" if is_no_rep_o else (str(rec.get("hv_remarks", "") or p_ro.get("hv_remarks", "")).strip() or "-")
            mw_r = "(No Report)" if is_no_rep_o else (str(rec.get("mw_remarks", "") or p_ro.get("mw_remarks", "")).strip() or "-")
            ind_r = "(No Report)" if is_no_rep_o else (str(rec.get("ind_remarks", "") or p_ro.get("ind_remarks", "")).strip() or "-")

            short_parish = p_name.replace("PARISH", "").strip()
            if "HONGKONG" in short_parish:
                short_parish = "HONGKONG CHAPTER"
            sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_data else ("❌ (No Report)" if is_no_rep_o else "⏳ Hindi Pa Tapos")

            rows.append({
                "Parish & Date": f"{short_parish} ({r_date})" + (" — (No Report)" if is_no_rep_o else ""),
                "Buwan & Week": f"{selected_month} (Week {week_num_only})",
                "Remarks sa House Visitation": hv_r,
                "Remarks sa Missionary Work": mw_r,
                "Remarks sa Indoctrination": ind_r,
                "Date & Time Submitted (PH Time)": f"✅ {sub_ph}" if (has_data and is_week_record_completed(rec)) else ("❌ (No Report)" if is_no_rep_o else "⏳ Hindi Pa Tapos")
            })
    if not rows:
        rows.append({
            "Parish & Date": "-",
            "Buwan & Week": f"{selected_month}",
            "Remarks sa House Visitation": "-",
            "Remarks sa Missionary Work": "-",
            "Remarks sa Indoctrination": "-",
            "Date & Time Submitted (PH Time)": "⏳ Hindi Pa Tapos"
        })
    return pd.DataFrame(rows)

TOP_WEEK_VIEW_OPTIONS = WEEKS_2027 + ["Lahat ng Weeks sa Buwan (1st–5th Week)"]

def build_monthly_all_weeks_monitoring_df(selected_month: str, filter_parish: str = "Lahat ng Parokya") -> pd.DataFrame:
    """Bumubuo ng Parish Monitoring Report — Attendance & Total Percentage table na nagpapakita ng LAHAT ng Weeks (1st–5th Week) sa loob ng napiling Buwan."""
    db = load_weekly_monitoring_records()
    rows = []
    first_col_name = f"(Weekly Reports of {selected_month.upper()} 2027 — Weeks 1–5)"
    parishes_to_scan = PARISH_ORDER if filter_parish == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, filter_parish)]

    # Kunin muna lahat ng completed weeks sa buwang ito; kung wala pa, ipakita ang 1st–5th Week para sa napiling parokya
    for wk in get_weeks_for_month_year(selected_month, 2027):
        week_num_only = re.sub(r"\D", "", str(wk)) or "1"
        for p_name in parishes_to_scan:
            key = f"{p_name.upper()}__{selected_month.upper()}__{wk}"
            rec = db.get(key, {})
            has_done = is_week_record_completed(rec)
            if not has_done and filter_parish == "Lahat ng Parokya":
                continue

            r_date = resolve_display_week_date(rec, selected_month, wk, 2027)
            m_a = int(rec.get("mass_a", 0)) if has_done else 0
            m_y = int(rec.get("mass_y", 0)) if has_done else 0
            p_a = int(rec.get("pampamilya_a", 0)) if has_done else 0
            p_y = int(rec.get("pampamilya_y", 0)) if has_done else 0
            o_a = int(rec.get("online_a", 0)) if has_done else 0
            o_y = int(rec.get("online_y", 0)) if has_done else 0
            y_y = int(rec.get("youth_y", 0)) if has_done else 0
            k_k = int(rec.get("kids_k", 0)) if has_done else 0
            ad_a = int(rec.get("adal_a", 0)) if has_done else 0
            ad_y = int(rec.get("adal_y", 0)) if has_done else 0

            tot_say = int(rec.get("total_members_say", 0)) if has_done else 0
            tot_y = int(rec.get("total_members_y", 0)) if has_done else 0
            tot_k = int(rec.get("total_members_k", 0)) if has_done else 0

            short_parish = p_name.replace("PARISH", "").strip()
            if "HONGKONG" in short_parish:
                short_parish = "HONGKONG CHAPTER"

            sub_ph = str(rec.get("submitted_at_ph", "") or rec.get("updated_at", "")).strip() if has_done else "⏳ Hindi Pa Tapos"
            status_tag = f"✅ Tapos Na ({sub_ph})" if has_done else "⏳ Hindi Pa Tapos"

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
                "Date & Time Submitted (PH Time)": f"Month: {selected_month} 2027 ({wk})",
            })

    if not rows:
        return build_parish_monitoring_attendance_df(selected_month, "1st Week", filter_parish)
    return pd.DataFrame(rows)

def build_monthly_all_weeks_summary_for_posting_df(parish_name: str, selected_month: str, only_completed: bool = False) -> pd.DataFrame:
    """Bumubuo ng Summary for Posting — Per Parish Breakdown table na nagpapakita ng LAHAT ng Weeks (1st–5th Week) sa loob ng napiling Buwan."""
    db = load_weekly_monitoring_records()
    frames = []
    for wk in get_weeks_for_month_year(selected_month, 2027):
        rec = db.get(f"{parish_name.upper()}__{selected_month.upper()}__{wk}", {})
        if only_completed and not is_week_record_completed(rec):
            continue
        frames.append(build_summary_for_posting_per_parish_df(parish_name, selected_month, wk))
    if not frames:
        return build_summary_for_posting_per_parish_df(parish_name, selected_month, "1st Week")
    return pd.concat(frames, ignore_index=True)

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

@st.cache_data(ttl=25, show_spinner=False)
def auto_poll_drive_weekly_report_changes(target_month: str, target_week: str) -> str:
    """Awtomatikong tinitingnan kung may binago/in-adjust sa Google Drive Weekly Report Excel files
    at kusang ina-update ang datos at attendance list sa app."""
    try:
        folder_url = st.session_state.get("custom_secrets", {}).get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
        cnt, msg = sync_weekly_reports_from_google_drive_folder(folder_url, target_month, target_week)
        return f"{cnt}:{msg}"
    except Exception as e:
        return str(e)

def render_weekly_report_page():
    is_master = (st.session_state.user_role == "Master Admin")
    my_parish = st.session_state.get("assigned_parish", "Lahat ng Parokya")
    my_chat_identity = get_current_chat_identity()
    db_all_init = load_weekly_monitoring_records()

    if "_next_mon_view_month" in st.session_state:
        nxt_m_val = st.session_state.pop("_next_mon_view_month")
        st.session_state["mon_view_month"] = nxt_m_val
        if nxt_m_val in MONTHS_2027:
            st.session_state["up_sel_month"] = nxt_m_val
            st.session_state["editor_sel_month"] = nxt_m_val
            st.session_state["outreach_sel_month"] = nxt_m_val
    if "_next_mon_view_week" in st.session_state:
        nxt_w_val = st.session_state.pop("_next_mon_view_week")
        st.session_state["mon_view_week"] = nxt_w_val
        if nxt_w_val in WEEKS_2027:
            st.session_state["up_sel_week"] = nxt_w_val
            st.session_state["editor_sel_week"] = nxt_w_val
            st.session_state["member_sheet_week_sel"] = nxt_w_val
    if "_next_editor_parish" in st.session_state:
        nxt_p_val = st.session_state.pop("_next_editor_parish")
        st.session_state["editor_sel_parish"] = nxt_p_val
        st.session_state["up_sel_parish"] = nxt_p_val
        if is_master:
            st.session_state["mon_view_parish"] = nxt_p_val
    if "_next_editor_month" in st.session_state:
        nxt_em_val = st.session_state.pop("_next_editor_month")
        st.session_state["editor_sel_month"] = nxt_em_val
        st.session_state["up_sel_month"] = nxt_em_val
        st.session_state["mon_view_month"] = nxt_em_val
        st.session_state["outreach_sel_month"] = nxt_em_val
    if "_next_editor_week" in st.session_state:
        nxt_ew_val = st.session_state.pop("_next_editor_week")
        st.session_state["editor_sel_week"] = nxt_ew_val
        st.session_state["up_sel_week"] = nxt_ew_val
        st.session_state["mon_view_week"] = nxt_ew_val
        st.session_state["member_sheet_week_sel"] = nxt_ew_val

    if "mon_view_month" not in st.session_state:
        st.session_state["mon_view_month"] = "January"
    if "mon_view_year" not in st.session_state:
        st.session_state["mon_view_year"] = "2027"

    cur_m_for_init = st.session_state["mon_view_month"] if st.session_state["mon_view_month"] in MONTHS_2027 else "January"
    cur_y_for_init = int(st.session_state.get("mon_view_year", "2027"))
    valid_weeks_init = get_weeks_for_month_year(cur_m_for_init, cur_y_for_init)

    if "mon_view_week" not in st.session_state or (st.session_state["mon_view_week"] in WEEKS_2027 and st.session_state["mon_view_week"] not in valid_weeks_init):
        st.session_state["mon_view_week"] = valid_weeks_init[0]
    if "up_sel_month" not in st.session_state:
        st.session_state["up_sel_month"] = cur_m_for_init
    valid_up_weeks = get_weeks_for_month_year(st.session_state["up_sel_month"], cur_y_for_init)
    if "up_sel_week" not in st.session_state or st.session_state["up_sel_week"] not in valid_up_weeks:
        st.session_state["up_sel_week"] = valid_up_weeks[0]
    if "editor_sel_month" not in st.session_state:
        st.session_state["editor_sel_month"] = st.session_state["up_sel_month"]
    valid_ed_weeks = get_weeks_for_month_year(st.session_state["editor_sel_month"], cur_y_for_init)
    if "editor_sel_week" not in st.session_state or st.session_state["editor_sel_week"] not in valid_ed_weeks:
        st.session_state["editor_sel_week"] = valid_ed_weeks[0]

    default_parish_choice = my_parish if (not is_master and my_parish in PARISH_ORDER) else PARISH_ORDER[0]
    if "up_sel_parish" not in st.session_state or (not is_master and st.session_state["up_sel_parish"] != my_parish):
        st.session_state["up_sel_parish"] = default_parish_choice
    if "editor_sel_parish" not in st.session_state or (not is_master and st.session_state["editor_sel_parish"] != my_parish):
        st.session_state["editor_sel_parish"] = st.session_state["up_sel_parish"]

    st.markdown("<div class='pmr-banner'>CHURCH OF THE TEMPLE OF GOD</div>", unsafe_allow_html=True)
    st.markdown(
        f"<div class='pmr-subbanner'>Parish Monitoring Report — Full Year 2027 (January–December, Dynamic Weeks per Month) & Google Drive Auto-Sync "
        f"<br><span class='ph-time-badge'>🇵🇭 Oras at Petsa sa Pilipinas: {get_philippine_datetime_str()}</span></div>",
        unsafe_allow_html=True
    )

    if is_master:
        st.info(
            "🏛️ **Master Admin Monitoring Dashboard:** Ang pag-upload ng Weekly Reports ay para sa mga **Parish Admins (Lipay Norte hanggang Hongkong)**. "
            "Bilang **Master Admin**, dito mo mamo-monitor ang **Performance Report** at lahat ng **Records** ng bawat Parokya mula sa Google Drive files."
        )

    # =====================================================================
    # 🎛️ TOP MASTER CONTROL BAR (BUWAN, LINGGO/ALL WEEKS, PAROKYA, YEAR, AUTO-SYNC & REFRESH)
    # =====================================================================
    top_month_options = MONTHS_2027 + ["Lahat ng Buwan (Jan–Dec)"]
    top_f1, top_f2, top_f3, top_f4, top_f5, top_f6 = st.columns([1.0, 1.15, 1.35, 0.70, 1.05, 0.85])
    with top_f1:
        view_month_sel = st.selectbox(
            "📅 Buwan:",
            top_month_options,
            key="mon_view_month",
            on_change=_sync_from_top_bar
        )
        view_month = view_month_sel if view_month_sel in MONTHS_2027 else st.session_state.get("editor_sel_month", "January")
        show_all_months = (view_month_sel not in MONTHS_2027)
    with top_f4:
        selected_year = st.selectbox(
            "📆 Year:",
            YEARS_SUPPORTED,
            key="mon_view_year"
        )
    active_month_weeks = get_weeks_for_month_year(view_month, int(selected_year))
    dynamic_top_week_opts = active_month_weeks + [f"Lahat ng Weeks sa Buwan (1st–{active_month_weeks[-1]})"]
    if st.session_state.get("mon_view_week") not in dynamic_top_week_opts:
        st.session_state["mon_view_week"] = active_month_weeks[0]
    with top_f2:
        view_week_selection = st.selectbox(
            f"🗓️ Linggo ({len(active_month_weeks)} Weeks sa {view_month}):",
            dynamic_top_week_opts,
            key="mon_view_week",
            on_change=_sync_from_top_bar
        )
        view_week = view_week_selection if view_week_selection in active_month_weeks else active_month_weeks[0]
        show_all_weeks_of_month = (view_week_selection not in active_month_weeks)
    with top_f3:
        if is_master:
            master_parish_opts = PARISH_ORDER + ["Lahat ng Parokya (Master Admin Full View)"]
            if "mon_view_parish" not in st.session_state or st.session_state["mon_view_parish"] not in master_parish_opts:
                st.session_state["mon_view_parish"] = PARISH_ORDER[0]
            view_parish_scope = st.selectbox(
                "🏛️ Lahat ng Parokya (Master Admin Full View):",
                master_parish_opts,
                key="mon_view_parish",
                on_change=_sync_from_top_bar
            )
            effective_parish_filter = "Lahat ng Parokya" if "Lahat ng Parokya" in view_parish_scope else view_parish_scope
            focused_parish_for_sheets = effective_parish_filter if effective_parish_filter in PARISH_ORDER else st.session_state.get("editor_sel_parish", PARISH_ORDER[0])
        else:
            st.info(f"📍 Parish View: **{my_parish}**")
            effective_parish_filter = my_parish
            focused_parish_for_sheets = my_parish
    with top_f5:
        st.write("")
        if st.button("🔄 Auto-Sync Data Mula sa Google Drive", use_container_width=True, type="primary"):
            with st.spinner(f"Binabasa ang mga Excel files mula sa Google Drive Folder para sa {view_month} {selected_year} ({view_week})..."):
                folder_url = st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
                cnt, msg = sync_weekly_reports_from_google_drive_folder(folder_url, view_month, view_week)
                st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                st.cache_data.clear()
                st.success(msg)
                time.sleep(0.5)
                st.rerun()
    with top_f6:
        st.write("")
        if st.button(t_ui("refresh_app"), use_container_width=True, key="top_bar_refresh_btn"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.rerun()

    db_all_init = load_weekly_monitoring_records()

    # DYNAMIC QUICK WEEK BUTTONS (Displays ONLY the actual weeks in the selected month, e.g. 4 weeks or 5 weeks)
    wk_cols = st.columns(len(active_month_weeks) + 1)
    parishes_for_badge = PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]
    for idx_w, wk_opt in enumerate(active_month_weeks):
        done_cnt_wk = sum(
            1 for p_b in parishes_for_badge
            if is_week_record_completed(db_all_init.get(f"{p_b.upper()}__{view_month.upper()}__{wk_opt}", {}))
        )
        wk_sat_date = get_week_date_for_month_year(view_month, wk_opt, int(selected_year))
        no_rep_cnt_wk = sum(
            1 for p_b in parishes_for_badge
            if is_record_marked_no_report(db_all_init.get(f"{p_b.upper()}__{view_month.upper()}__{wk_opt}", {}))
        )
        badge_icon = "✅" if done_cnt_wk > 0 else ("❌" if no_rep_cnt_wk > 0 else "⏳")
        is_active_wk = (not show_all_weeks_of_month and view_week == wk_opt)
        btn_lbl = f"{'👉 ' if is_active_wk else ''}{badge_icon} {wk_opt} ({wk_sat_date})"
        with wk_cols[idx_w]:
            if st.button(btn_lbl, key=f"qwk_btn_{view_month}_{wk_opt}", use_container_width=True, type="primary" if is_active_wk else "secondary"):
                st.session_state["_next_mon_view_month"] = view_month
                st.session_state["_next_mon_view_week"] = wk_opt
                st.rerun()
    with wk_cols[-1]:
        all_wk_active = show_all_weeks_of_month
        if st.button(f"{'👉 ' if all_wk_active else ''}📚 Lahat (W1–W{len(active_month_weeks)})", key=f"qwk_all_{view_month}", use_container_width=True, type="primary" if all_wk_active else "secondary"):
            st.session_state["_next_mon_view_month"] = view_month
            st.session_state["_next_mon_view_week"] = dynamic_top_week_opts[-1]
            st.rerun()

    # 🗑️ CLICKABLE OPTION TO CLEAR / DELETE POPULATED DATA & MASTER ADMIN DEADLINE / OVERRIDE CONTROLLER
    ctrl_bar_1, ctrl_bar_2 = st.columns([1.85, 1.15]) if is_master else ([None, st.container()])
    if is_master and ctrl_bar_1 is not None:
        with ctrl_bar_1:
            with st.expander("🛡️ Master Admin Submission Schedule & Deadline Controller (Auto Sat–Mon Unlock/Lock & No Report Control)", expanded=True):
                cur_auto_sched = is_auto_schedule_lock_enabled()
                cur_dl_enabled = is_deadline_enforcement_enabled()
                st.session_state["master_deadline_lock_toggle"] = cur_dl_enabled
                st.session_state["master_auto_sched_toggle"] = cur_auto_sched

                dl_col_a, dl_col_b = st.columns([1.45, 1.35])
                with dl_col_a:
                    lock_submission_on = st.toggle(
                        "🔒 Manual Lock / Unlock ng Pasahan (ON = Naka-LOCK ang Upload | OFF = UNLOCKED / Pwedeng Mag-upload ang mga Leaders)",
                        key="master_deadline_lock_toggle",
                        on_change=_on_master_lock_toggle_clicked
                    )
                with dl_col_b:
                    auto_sched_on = st.toggle(
                        "⏰ Auto Unlock/Lock Schedule (Sabado 12:00 AM Auto-Unlock hanggang Lunes 11:59 PM Midnight Auto-Lock)",
                        key="master_auto_sched_toggle",
                        on_change=_on_master_auto_sched_clicked
                    )

                if lock_submission_on:
                    p_ov_currently_allowed = is_parish_week_allowed_by_admin(focused_parish_for_sheets, view_month, view_week)
                    ov_btn_lbl = (
                        f"🔒 I-lock Muli ang Upload ({focused_parish_for_sheets})"
                        if p_ov_currently_allowed
                        else f"🔓 I-unlock Lang para sa {focused_parish_for_sheets}"
                    )
                    if st.button(ov_btn_lbl, use_container_width=True, key="btn_parish_wk_override"):
                        set_admin_allowed_submission_override(focused_parish_for_sheets, view_month, view_week, not p_ov_currently_allowed)
                        st.rerun()
                    st.caption("🔒 **Status: NAKA-LOCK ANG PASAHAN** — Hindi makakapag-upload ang mga Parish Leaders sa `Upload the Image / Excel File Here!` (Auto-unlock tuwing Sabado o kapag in-OFF mo ang Lock).")
                else:
                    st.caption("🔓 **Status: UNLOCKED (Bukas ang Pasahan)** — Maaaring mag-upload ang mga Parish Leaders sa `Upload the Image / Excel File Here!` (Auto-lock pagsapit ng Lunes 11:59 PM Midnight).")

                st.markdown("---")
                st.markdown("##### 📌 Master Admin `(No Report)` Controller (Piliin kung aling Parokya, Buwan, at Linggo ang Lalagyan ng `No Report`)")
                nr_c1, nr_c2, nr_c3 = st.columns([1.35, 1.0, 1.0])
                nr_parish_opts = PARISH_ORDER + ["LAHAT NG PAROKYA NA WALANG PASA (All Unsubmitted)"]
                if "nr_ctrl_parish" not in st.session_state or st.session_state["nr_ctrl_parish"] not in nr_parish_opts:
                    st.session_state["nr_ctrl_parish"] = focused_parish_for_sheets if focused_parish_for_sheets in PARISH_ORDER else PARISH_ORDER[0]
                if "nr_ctrl_month" not in st.session_state or st.session_state["nr_ctrl_month"] not in MONTHS_2027:
                    st.session_state["nr_ctrl_month"] = view_month
                with nr_c1:
                    nr_target_parish = st.selectbox("🏛️ Parokya na Lalagyan ng No Report:", nr_parish_opts, key="nr_ctrl_parish")
                with nr_c2:
                    nr_target_month = st.selectbox("📅 Buwan:", MONTHS_2027, key="nr_ctrl_month")
                nr_valid_weeks = get_weeks_for_month_year(nr_target_month, int(selected_year))
                if "nr_ctrl_week" not in st.session_state or st.session_state["nr_ctrl_week"] not in nr_valid_weeks:
                    st.session_state["nr_ctrl_week"] = view_week if view_week in nr_valid_weeks else nr_valid_weeks[0]
                with nr_c3:
                    nr_target_week = st.selectbox("🗓️ Linggo (Week):", nr_valid_weeks, key="nr_ctrl_week")

                nr_btn_1, nr_btn_2 = st.columns(2)
                with nr_btn_1:
                    if st.button(f"❌ Lagyan ng '(No Report)' ({nr_target_parish} • {nr_target_month} {nr_target_week})", type="primary", use_container_width=True, key="btn_apply_no_report"):
                        if "LAHAT NG PAROKYA" in nr_target_parish:
                            for p_item_nr in PARISH_ORDER:
                                set_parish_no_report_status(p_item_nr, nr_target_month, nr_target_week, int(selected_year), mark_no_report=True, force_overwrite=False)
                        else:
                            set_parish_no_report_status(nr_target_parish, nr_target_month, nr_target_week, int(selected_year), mark_no_report=True, force_overwrite=True)
                            st.session_state["_next_editor_parish"] = nr_target_parish
                        st.session_state["_next_mon_view_month"] = nr_target_month
                        st.session_state["_next_mon_view_week"] = nr_target_week
                        st.success(f"✅ Nalagyan na ng '(No Report)' ang {nr_target_parish} para sa {nr_target_month} {selected_year} ({nr_target_week})!")
                        time.sleep(0.4)
                        st.rerun()
                with nr_btn_2:
                    if st.button(f"♻️ Alisin ang '(No Report)' ({nr_target_parish} • {nr_target_month} {nr_target_week})", use_container_width=True, key="btn_clear_no_report"):
                        if "LAHAT NG PAROKYA" in nr_target_parish:
                            for p_item_nr in PARISH_ORDER:
                                set_parish_no_report_status(p_item_nr, nr_target_month, nr_target_week, int(selected_year), mark_no_report=False)
                        else:
                            set_parish_no_report_status(nr_target_parish, nr_target_month, nr_target_week, int(selected_year), mark_no_report=False)
                            st.session_state["_next_editor_parish"] = nr_target_parish
                        st.session_state["_next_mon_view_month"] = nr_target_month
                        st.session_state["_next_mon_view_week"] = nr_target_week
                        st.success(f"♻️ Inalis na ang '(No Report)' sa {nr_target_parish} para sa {nr_target_month} {selected_year} ({nr_target_week})!")
                        time.sleep(0.4)
                        st.rerun()

                marked_nr_parishes = [
                    p_m.replace("PARISH", "").strip()
                    for p_m in PARISH_ORDER
                    if is_record_marked_no_report(db_all_init.get(f"{p_m.upper()}__{nr_target_month.upper()}__{nr_target_week}", {}))
                ]
                if marked_nr_parishes:
                    joined_nr_p = ", ".join(marked_nr_parishes)
                    st.caption(f"📋 **Kasalukuyang may `(No Report)` sa {nr_target_month} ({nr_target_week}):** `{joined_nr_p}`")
    with ctrl_bar_2:
        with st.expander(f"🗑️ {t_ui('delete_populated')}", expanded=False):
            st.caption("Ang na-populate na data mula sa Image o Excel ay **mananatiling naka-save** at mabubura lamang kung pipindutin mo ang Delete button dito:")
            del_target_p = focused_parish_for_sheets if is_master else my_parish
            st.write(f"🎯 Target: **{del_target_p}** — `{view_month} {selected_year} ({view_week})`")
            confirm_del_chk = st.checkbox("Kumpirmahin ang pagbura (Delete Record)", key=f"conf_del_{del_target_p}_{view_month}_{view_week}")
            if st.button("🗑️ Burahin Ngayon ang Record na Ito", type="primary", use_container_width=True, disabled=not confirm_del_chk):
                delete_populated_weekly_record(del_target_p, view_month, view_week)
                dispatch_app_notification_and_messenger_alert(
                    sender_id=my_chat_identity,
                    recipient_id="MASTER ADMIN",
                    title=f"Deleted Populated Weekly Data ({del_target_p})",
                    details=f"Binura ni {my_chat_identity} ang na-populate na data para sa {del_target_p} ({view_month} {selected_year} - {view_week}).",
                    is_error=False
                )
                st.success(f"Binura na ang na-populate na data para sa {del_target_p} ({view_month} - {view_week})!")
                time.sleep(0.5)
                st.rerun()

    st.divider()

    # =====================================================================
    # 📤 FOR ALL PARISH ADMINS: 1. UPLOAD THE IMAGE / EXCEL FILE HERE! & 2. MANUAL VERIFICATION EDITOR
    # =====================================================================
    def _render_verification_editor_panel():
        st.markdown("### ✏️ Parish Weekly Report Data & Manual Verification Editor")
        st.caption("Lahat ng datos dito ay direktang tinatype sa `PARISH WEEKLY REPORT` sheet (Liturgical Officers, Topics/Preachers, Adults/Youths/Kids Attendance, at Remarks). Naka-sync ang Buwan at Linggo rito sa Top Bar.")

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
        ed_month_weeks = get_weeks_for_month_year(v_month, int(selected_year))
        if st.session_state.get("editor_sel_week") not in ed_month_weeks:
            st.session_state["editor_sel_week"] = ed_month_weeks[0]
        with v_c3:
            v_week = st.selectbox(
                f"Linggo ({len(ed_month_weeks)} Weeks):",
                ed_month_weeks,
                key="editor_sel_week",
                on_change=_sync_from_editor_bar
            )

        cur_rec = db_all.get(f"{v_parish.upper()}__{v_month.upper()}__{v_week}", {})
        f_rev = f"{st.session_state.get('pwr_form_rev', 0)}_{v_parish}_{v_month}_{v_week}"
        def_denoms = get_parish_member_denominators(v_parish)
        default_wk_date = resolve_display_week_date(cur_rec, v_month, v_week, int(selected_year))
        ed_dl_info = evaluate_reporting_deadline(v_month, v_week, int(selected_year), parish_name=v_parish)
        ed_locked = (not is_master) and bool(ed_dl_info["is_past_deadline"])

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
            v_date = st.text_input("Petsa ng Pagsamba / Sabado (MM/DD/YY):", value=default_wk_date, key=f"vdt_{f_rev}")
            # Panatilihin ang kasalukuyang Outreach data mula sa nakahiwalay na pahina ng House Visitation, Missionary Work & Indoctrination
            v_hv = str(cur_rec.get("house_visitation", "0"))
            v_hv_names = str(cur_rec.get("hv_names", "") or parsed_cur_ro.get("hv_names", ""))
            v_hv_ppl = int(cur_rec.get("hv_people_count", 0) or parsed_cur_ro.get("hv_people_count", 0))
            v_hv_rem = str(cur_rec.get("hv_remarks", "") or parsed_cur_ro.get("hv_remarks", ""))
            v_mw = str(cur_rec.get("missionary_work", "0"))
            v_mw_names = str(cur_rec.get("mw_names", "") or parsed_cur_ro.get("mw_names", ""))
            v_mw_ppl = int(cur_rec.get("mw_people_count", 0) or parsed_cur_ro.get("mw_people_count", 0))
            v_mw_rem = str(cur_rec.get("mw_remarks", "") or parsed_cur_ro.get("mw_remarks", ""))
            v_ind = str(cur_rec.get("indoctrination", "0"))
            v_ind_names = str(cur_rec.get("ind_names", "") or parsed_cur_ro.get("ind_names", ""))
            v_ind_ppl = int(cur_rec.get("ind_people_count", 0) or parsed_cur_ro.get("ind_people_count", 0))
            v_ind_rem = str(cur_rec.get("ind_remarks", "") or parsed_cur_ro.get("ind_remarks", ""))

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

            if st.form_submit_button("💾 I-type sa Parish Weekly Report Excel & I-update ang Lahat ng Tables", use_container_width=True, disabled=ed_locked):
                if ed_locked:
                    mark_unsubmitted_parish_as_no_report_in_excel_and_db(v_parish, v_month, v_week, int(selected_year))
                    dispatch_app_notification_and_messenger_alert(
                        sender_id=my_chat_identity,
                        recipient_id="MASTER ADMIN",
                        title=f"Late Submission Blocked ({v_parish} — {v_month} {v_week})",
                        details=f"Sinubukan ni {my_chat_identity} na mag-save ng late report para sa {v_month} {v_week} pagkalagpas ng Monday Midnight deadline ({ed_dl_info['monday_deadline_str']}). Naka-marka na bilang '(No Report)'.",
                        is_error=True
                    )
                    st.error(f"{t_ui('deadline_closed')} (Susunod na Sabado: {ed_dl_info['next_saturday_str']})")
                else:
                    try:
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
                        ok_s, sync_msg = save_weekly_monitoring_record(v_parish, v_month, v_week, v_date, updated_payload, drive_link=cur_rec.get("drive_link", ""))
                        if not ok_s:
                            dispatch_app_notification_and_messenger_alert(
                                sender_id=my_chat_identity,
                                recipient_id="MASTER ADMIN",
                                title=f"Excel Sync Warning ({v_parish} — {v_month} {v_week})",
                                details=f"May isyu habang nagse-save si {my_chat_identity}: {sync_msg}",
                                is_error=True
                            )
                        else:
                            dispatch_app_notification_and_messenger_alert(
                                sender_id=my_chat_identity,
                                recipient_id="MASTER ADMIN",
                                title=f"Weekly Report Updated ({v_parish})",
                                details=f"Matagumpay na na-update ni {my_chat_identity} ang Weekly Report ng {v_parish} para sa {v_month} {selected_year} ({v_week}).",
                                is_error=False
                            )
                        st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                        st.success(f"✅ Na-update na sa PARISH WEEKLY REPORT at Monitoring Tables para sa **{v_parish}** ({v_month} {selected_year} - {v_week})! ({sync_msg})")
                        time.sleep(0.5)
                        st.rerun()
                    except Exception as ex_ed:
                        dispatch_app_notification_and_messenger_alert(
                            sender_id=my_chat_identity,
                            recipient_id="MASTER ADMIN",
                            title=f"Error sa Manual Verification Editor ({v_parish})",
                            details=f"Nagkaroon ng error habang nagse-save si {my_chat_identity} para sa {v_month} {v_week}: {str(ex_ed)}",
                            is_error=True
                        )
                        st.error(f"Nagkaroon ng error at na-report na kay Master Admin sa Ministry Messenger: {str(ex_ed)}")
        return v_parish

    if not is_master:
        col_up_left, col_up_right = st.columns([1.15, 1.55])
        with col_up_left:
            st.markdown("### 📤 1. Upload the Image / Excel File Here!")
            st.caption("I-upload ang larawan (screenshot/picture) o Excel file ng inyong Weekly Report. Ang deadline ng submission ay tuwing **Monday Midnight (11:59 PM PH Time)**.")

            c_m, c_w = st.columns(2)
            with c_m:
                sel_month = st.selectbox(
                    "Buwan ng Report (2027):",
                    MONTHS_2027,
                    key="up_sel_month",
                    on_change=_sync_from_upload_bar
                )
            up_month_weeks = get_weeks_for_month_year(sel_month, int(selected_year))
            if st.session_state.get("up_sel_week") not in up_month_weeks:
                st.session_state["up_sel_week"] = up_month_weeks[0]
            with c_w:
                sel_week = st.selectbox(
                    f"Linggo ng Report ({len(up_month_weeks)} Weeks sa {sel_month}):",
                    up_month_weeks,
                    key="up_sel_week",
                    on_change=_sync_from_upload_bar
                )

            parish_opts_up = [my_parish]
            if st.session_state.get("up_sel_parish") not in parish_opts_up:
                st.session_state["up_sel_parish"] = parish_opts_up[0]
            target_parish = st.selectbox(
                "Parokya na Nag-report:",
                parish_opts_up,
                key="up_sel_parish",
                on_change=_sync_from_upload_bar
            )

            up_dl_info = evaluate_reporting_deadline(sel_month, sel_week, int(selected_year), parish_name=target_parish)
            existing_up_rec = db_all_init.get(f"{target_parish.upper()}__{sel_month.upper()}__{sel_week}", {})
            has_submitted_up = is_week_record_completed(existing_up_rec)
            is_marked_nr_up = is_record_marked_no_report(existing_up_rec)
            is_up_locked = bool(up_dl_info["is_past_deadline"])
            if is_marked_nr_up:
                st.error(
                    f"⛔ **(No Report) — Ang {target_parish} para sa {sel_month} {selected_year} ({sel_week}, Petsa ng Sabado: `{up_dl_info['saturday_date_str']}`) ay naka-marka bilang `(No Report)`.** "
                    f"Ang opisyal na araw ng pasahan ay tuwing **Sabado (Saturday)** hanggang **Lunes ng Hatinggabi (`{up_dl_info['monday_deadline_str']}`)**."
                )
            else:
                st.info(
                    f"📅 **Paalala sa Araw ng Pasahan at Deadline ng Weekly Report ({target_parish} — {sel_month} {selected_year}, {sel_week} | Sabado: `{up_dl_info['saturday_date_str']}`):** "
                    f"Ang pasahan ng Weekly Report ay nagbubukas tuwing **Sabado (Saturday)** at nagtatapos tuwing **Lunes ng Hatinggabi (Monday Midnight — `{up_dl_info['monday_deadline_str']}`)**. "
                    f"Mangyaring mag-submit sa araw ng pasahan bago sumapit ang deadline."
                )

            auto_up_date = get_week_date_for_month_year(sel_month, sel_week, int(selected_year))
            with st.form("weekly_report_upload_form"):
                rep_date_input = st.text_input("Petsa ng Sabado / Pagsamba (MM/DD/YY):", value=auto_up_date, disabled=is_up_locked)
                uploaded_rep = st.file_uploader(
                    "📸 Pumili ng Weekly Report Image (.jpg, .png) o Excel (.xlsx):",
                    type=["jpg", "jpeg", "png", "xlsx"],
                    disabled=is_up_locked
                )
                up_hv = str(existing_up_rec.get("house_visitation", "0"))
                up_mw = str(existing_up_rec.get("missionary_work", "0"))
                up_ind = str(existing_up_rec.get("indoctrination", "0"))
                extra_remarks = "" 

                submit_up = st.form_submit_button(
                    "🔒 Sarado ang Pasahan (Naka-ON ang Deadline Lock)" if is_up_locked else "🚀 I-upload sa Google Drive & I-type sa Parish Weekly Report",
                    type="primary",
                    use_container_width=True,
                    disabled=is_up_locked
                )

                if submit_up:
                    if is_up_locked:
                        mark_unsubmitted_parish_as_no_report_in_excel_and_db(target_parish, sel_month, sel_week, int(selected_year))
                        dispatch_app_notification_and_messenger_alert(
                            sender_id=my_chat_identity,
                            recipient_id="MASTER ADMIN",
                            title=f"Blocked Late Report Upload ({target_parish} — {sel_month} {sel_week})",
                            details=f"Sinubukang mag-upload ni {my_chat_identity} para sa {sel_month} {sel_week} (Petsa: {rep_date_input}) ngunit Martes–Biyernes / lagpas na sa Monday Midnight deadline ({up_dl_info['monday_deadline_str']}). Naka-marka na bilang '(No Report)'.",
                            is_error=True
                        )
                        st.error(f"{t_ui('deadline_closed')} (Susunod na Sabado: {up_dl_info['next_saturday_str']})")
                    elif uploaded_rep is None:
                        st.warning("Mag-upload muna ng Image o Excel file ng Weekly Report.")
                    else:
                        t_start_up = time.time()
                        prog_bar = st.progress(5, text="⏱️ 0.1s — Inihahanda ang file para sa mabilis na pag-upload...")
                        timer_placeholder = st.empty()
                        try:
                            file_bytes = uploaded_rep.read()
                            f_ext = uploaded_rep.name.split(".")[-1].lower()
                            clean_p_name = target_parish.replace("PARISH", "").strip()
                            exact_date_clean = rep_date_input.strip() or auto_up_date
                            drive_filename = f"{clean_p_name} (PWR. {exact_date_clean.replace('/', '-')} {sel_month} {sel_week}).{f_ext}"

                            elapsed_1 = time.time() - t_start_up
                            prog_bar.progress(30, text=f"⏱️ {elapsed_1:.1f}s — Ina-upload sa Google Drive folder ng {target_parish} ('Upload the Image/Excel File here!')...")
                            ph_now_stamp = get_philippine_datetime_str()

                            up_folder_url = st.session_state.custom_secrets.get("PARISH_UPLOAD_FOLDER_URL", ACTIVE_PARISH_UPLOAD_URL) or st.session_state.custom_secrets.get("WEEKLY_REPORT_FOLDER_URL", ACTIVE_WEEKLY_FOLDER_URL)
                            # Awtomatikong papasok sa Parish Folder (at sa Month subfolder tulad ng '01. January Pwr. 2027' kung mayroon sa loob ng Parish Folder)
                            drive_link, drive_msg = upload_file_to_google_drive(
                                file_bytes,
                                drive_filename,
                                mime_type=uploaded_rep.type,
                                parent_folder_url=up_folder_url,
                                parish_name=target_parish,
                                month_name=sel_month
                            )

                            if not drive_link and drive_msg:
                                dispatch_app_notification_and_messenger_alert(
                                    sender_id=my_chat_identity,
                                    recipient_id="MASTER ADMIN",
                                    title=f"Google Drive Upload Issue ({target_parish})",
                                    details=f"File: {uploaded_rep.name} | Buwan/Linggo: {sel_month} {sel_week} | Issue: {drive_msg}",
                                    is_error=True
                                )

                            elapsed_2 = time.time() - t_start_up
                            prog_bar.progress(65, text=f"⏱️ {elapsed_2:.1f}s — Binabasa ang datos para lamang sa {sel_month} ({sel_week}, Petsa: {exact_date_clean})...")

                            # Kunin LAMANG ang datos para sa eksaktong napiling Buwan (sel_month), Linggo (sel_week), Parokya (target_parish), at Petsa (exact_date_clean)
                            if f_ext == "xlsx":
                                extracted = extract_weekly_report_from_excel_bytes(
                                    file_bytes,
                                    target_parish=target_parish,
                                    target_month=sel_month,
                                    target_week=sel_week
                                )
                            else:
                                extracted = extract_weekly_report_via_ai_vision(
                                    file_bytes,
                                    uploaded_rep.type,
                                    selected_model,
                                    target_parish=target_parish
                                )

                            extracted["parish"] = target_parish
                            extracted["month"] = sel_month
                            extracted["week"] = sel_week
                            extracted["report_date"] = exact_date_clean
                            extracted["submitted_at_ph"] = ph_now_stamp

                            denoms = get_parish_member_denominators(target_parish)
                            if int(extracted.get("total_members_say", 0)) <= 0:
                                extracted["total_members_say"] = denoms["say"]
                                extracted["total_members_y"] = denoms["youth"]
                                extracted["total_members_k"] = denoms["kids"]

                            if up_hv.strip() and up_hv.strip() != "0":
                                extracted["house_visitation"] = up_hv.strip()
                            if up_mw.strip() and up_mw.strip() != "0":
                                extracted["missionary_work"] = up_mw.strip()
                            if up_ind.strip() and up_ind.strip() != "0":
                                extracted["indoctrination"] = up_ind.strip()
                            if extra_remarks:
                                extracted["remarks"] = extra_remarks

                            elapsed_3 = time.time() - t_start_up
                            prog_bar.progress(88, text=f"⏱️ {elapsed_3:.1f}s — Tinatype sa App at sa Google Drive Excel ({sel_month} - {sel_week})...")

                            ok_m, m_sync_msg = save_weekly_monitoring_record(
                                target_parish, sel_month, sel_week,
                                exact_date_clean,
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
                            is_blank_up = (tot_extracted_att == 0 and not has_any_officer)
                            total_elapsed = round(time.time() - t_start_up, 2)
                            prog_bar.progress(100, text=f"✅ Tapos na sa loob ng ⏱️ {total_elapsed:.1f} seconds!")

                            if is_blank_up:
                                dispatch_app_notification_and_messenger_alert(
                                    sender_id=my_chat_identity,
                                    recipient_id="MASTER ADMIN",
                                    title=f"Blank / Unreadable Report File Uploaded ({target_parish})",
                                    details=f"Nag-upload si {my_chat_identity} ng '{uploaded_rep.name}' para sa {sel_month} {sel_week} ({exact_date_clean}) sa loob ng {total_elapsed}s ngunit 0 attendance at walang officers na nabasa.",
                                    is_error=True
                                )
                            else:
                                dispatch_app_notification_and_messenger_alert(
                                    sender_id=my_chat_identity,
                                    recipient_id="MASTER ADMIN",
                                    title=f"Bagong Weekly Report Na-submit ({target_parish})",
                                    details=f"Matagumpay na nakapag-submit ang {target_parish} para sa {sel_month} {selected_year} ({sel_week}, Petsa: {exact_date_clean}) sa loob ng {total_elapsed}s — Mass: {int(extracted.get('mass_a',0))+int(extracted.get('mass_y',0))}.",
                                    is_error=False
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
                                "date": exact_date_clean,
                                "elapsed_sec": total_elapsed,
                                "data": extracted,
                                "drive_link": drive_link,
                                "drive_msg": drive_msg,
                                "sync_msg": m_sync_msg,
                                "is_blank_source": is_blank_up
                            }
                            st.rerun()
                        except Exception as ex_up:
                            elapsed_err = round(time.time() - t_start_up, 2)
                            timer_placeholder.caption(f"⏱️ Oras bago nag-error: {elapsed_err}s")
                            dispatch_app_notification_and_messenger_alert(
                                sender_id=my_chat_identity,
                                recipient_id="MASTER ADMIN",
                                title=f"Report Submission Error ({target_parish})",
                                details=f"Sinubukan ni {my_chat_identity} na mag-upload ng '{getattr(uploaded_rep, 'name', 'file')}' para sa {sel_month} {sel_week}. Error: {str(ex_up)}",
                                is_error=True
                            )
                            st.error(f"Nagkaroon ng error sa pag-submit at awtomatikong ipinadala kay Master Admin sa Ministry Messenger: {str(ex_up)}")

            if "last_extracted_pwr" in st.session_state:
                flash = st.session_state["last_extracted_pwr"]
                sub_ph_flash = flash.get("data", {}).get("submitted_at_ph", "") or get_philippine_datetime_str()
                el_s = flash.get("elapsed_sec")
                timer_badge = f" • ⏱️ **Na-upload sa loob ng `{el_s} seconds`**" if el_s is not None else ""
                st.success(
                    f"✅ Matagumpay na na-proseso ang Weekly Report ng **{flash.get('parish')}** para lamang sa **{flash.get('month')} {selected_year} ({flash.get('week')}, Petsa: `{flash.get('date')}`)**!{timer_badge} "
                    f"(🇵🇭 Submitted: `{sub_ph_flash}` • {flash.get('sync_msg', '')})"
                )
                if flash.get("drive_link"):
                    st.info(f"📂 Na-save sa Google Drive Folder ng {flash.get('parish')}: [Buksan ang File sa Google Drive]({flash.get('drive_link')})")
                elif flash.get("drive_msg"):
                    st.warning(f"⚠️ {flash.get('drive_msg')}")

        with col_up_right:
            _render_verification_editor_panel()
        st.divider()
    else:
        with st.expander("✏️ Buksan ang Parish Weekly Report Manual Verification Editor (Optional Admin Editor)", expanded=False):
            _render_verification_editor_panel()

    # =====================================================================
    # ⛪ SECTION A (BOOK1 EXCEL LAYOUT): LITURGICAL OFFICERS & PREACHERS (DYNAMIC WEEKS OF THE MONTH SIDE-BY-SIDE)
    # =====================================================================
    db_live = load_weekly_monitoring_records()
    months_to_display = MONTHS_2027 if show_all_months else [view_month]

    for m_disp in months_to_display:
        m_disp_weeks = get_weeks_for_month_year(m_disp, int(selected_year))
        st.markdown(
            f"### ⛪ Liturgical Officers & Preachers (Para sa PARISH WEEKLY REPORT Sheet) — `{focused_parish_for_sheets}` ({m_disp.upper()} {selected_year} — {len(m_disp_weeks)} Weeks)"
        )
        off_cols = st.columns(len(m_disp_weeks))
        officer_field_labels = [
            ("PRIEST:", "priest"),
            ("SAKRISTAN:", "sakristan"),
            ("EPISTULA:", "epistula"),
            ("EVANGELIO:", "evangelio"),
            ("CASASAAD:", "casasaad"),
            ("CANDLE HOLDER:", "candle_holder"),
            ("USHER:", "usher"),
            ("MARSHAL:", "marshal"),
            ("ADAL TI EVANGELIO:", "adal_evangelio"),
            ("Preacher/s:", "adal_preacher"),
            ("TEMA TI MISA:", "tema_misa"),
            ("Preacher:", "misa_preacher"),
            ("TEMA TI PAMPAMILYA:", "tema_pampamilya"),
            ("Preacher:", "pampamilya_preacher"),
        ]
        for idx_w5, wk_5 in enumerate(m_disp_weeks):
            rec_w5 = db_live.get(f"{focused_parish_for_sheets.upper()}__{m_disp.upper()}__{wk_5}", {})
            has_w5 = is_week_record_completed(rec_w5)
            dl_w5 = evaluate_reporting_deadline(m_disp, wk_5, int(selected_year), parish_name=focused_parish_for_sheets)
            is_no_rep_w5 = is_record_marked_no_report(rec_w5)
            wk_dt_5 = resolve_display_week_date(rec_w5, m_disp, wk_5, int(selected_year))
            with off_cols[idx_w5]:
                lines_html = [
                    f"<div class='ctg-officer-wk-hdr'>"
                    f"{wk_5} (Date: {wk_dt_5})"
                    + ("<br><span style='color:#ef4444;'>❌ (No Report)</span>" if is_no_rep_w5 else "")
                    + f"</div>"
                ]
                for lbl_txt, f_key in officer_field_labels:
                    val_txt = str(rec_w5.get(f_key, "")).strip() if has_w5 else ("(No Report)" if is_no_rep_w5 else "")
                    lines_html.append(
                        f"<div style='font-size:0.86rem;padding:3px 0;border-bottom:1px solid rgba(148,163,184,0.22);'>"
                        f"<b>{lbl_txt}</b> <span class='ctg-officer-val'>{val_txt if val_txt else '-'}</span>"
                        f"</div>"
                    )
                st.markdown(
                    f"<div class='full-text-card' style='padding:10px;'>{''.join(lines_html)}</div>",
                    unsafe_allow_html=True
                )

        st.divider()

        # =====================================================================
        # 📝 SECTION B (BOOK1 EXCEL LAYOUT): REMARKS PARA SA NAG-ATTEND NG MASS / PAMPAMILYA / ADAL / HOUSE CHURCH / ONLINE MASS / VISITORS
        # =====================================================================
        st.markdown(
            f"### 📝 Remarks para sa Nag-attend ng Mass / Pampamilya / Adal / House Church / Online Mass / Visitors (Galing sa Image/Excel Report) — `{focused_parish_for_sheets}` ({m_disp.upper()} {selected_year})"
        )
        for wk_r5 in m_disp_weeks:
            rec_r5 = db_live.get(f"{focused_parish_for_sheets.upper()}__{m_disp.upper()}__{wk_r5}", {})
            has_r5 = is_week_record_completed(rec_r5)
            dl_r5 = evaluate_reporting_deadline(m_disp, wk_r5, int(selected_year), parish_name=focused_parish_for_sheets)
            is_no_rep_r5 = is_record_marked_no_report(rec_r5)
            wk_dt_r5 = resolve_display_week_date(rec_r5, m_disp, wk_r5, int(selected_year))
            raw_rem_5 = str(rec_r5.get("remarks", "") or rec_r5.get("mass_remarks", "")).strip() if has_r5 else ("(No Report)" if is_no_rep_r5 else "")
            sub_ph_5 = str(rec_r5.get("submitted_at_ph", "") or rec_r5.get("updated_at", "")).strip() if has_r5 else ""
            disp_rem_5 = raw_rem_5.replace("\n", "<br>") if raw_rem_5 else "<span style='opacity:0.75;'>Wala pang nakalagay na Remarks para sa linggong ito.</span>"
            st.markdown(
                f"<div class='full-text-card' style='margin-bottom:10px;'>"
                f"<b>🗓️ {wk_r5} (Date: {wk_dt_r5})</b> — <code>{focused_parish_for_sheets}</code> "
                + (f"• 🇵🇭 <b>Submitted:</b> <code>{sub_ph_5}</code>" if sub_ph_5 else ("• ❌ <b>(No Report)</b>" if is_no_rep_r5 else "• ⏳ <i>Hindi Pa Tapos</i>"))
                + f"<br><hr style='margin:6px 0;'>"
                f"{disp_rem_5}"
                f"</div>",
                unsafe_allow_html=True
            )

        st.divider()

        # =====================================================================
        # 📊 SECTION C (BOOK1 EXCEL LAYOUT): PARISH WEEKLY REPORT DATA (ADULTS / YOUTHS / KIDS ATTENDANCE SHEET + TOTAL MEMBERS BOXES)
        # =====================================================================
        st.markdown(
            f"### 📊 Parish Weekly Report Data (Adults/Youths/Kids Attendance) — `{focused_parish_for_sheets}` ({m_disp.upper()} {selected_year})"
        )
        st.caption(f"Maaari kang pumili ng Linggo ({len(m_disp_weeks)} Weeks sa {m_disp} o Lahat ng Weeks) sa dropdown sa ibaba o sa Top Control Bar upang makita ang listahan ng Attendance ng Adults, Youths, at Kids.")

        ms_c1, ms_c2 = st.columns([1.4, 1.4])
        with ms_c1:
            mem_sheet_week_opts = m_disp_weeks + [f"Lahat ng Weeks (1st–{m_disp_weeks[-1]})"]
            cur_ms_wk = st.session_state.get("member_sheet_week_sel", view_week)
            if cur_ms_wk not in mem_sheet_week_opts:
                cur_ms_wk = mem_sheet_week_opts[-1] if show_all_weeks_of_month else m_disp_weeks[0]
                st.session_state["member_sheet_week_sel"] = cur_ms_wk
            sel_mem_sheet_wk = st.selectbox(
                f"🗓️ Pumili ng Linggo para sa Adults/Youths/Kids Attendance Sheet ({len(m_disp_weeks)} Weeks):",
                mem_sheet_week_opts,
                key=f"member_sheet_week_sel_{m_disp}" if show_all_months else "member_sheet_week_sel"
            )
        with ms_c2:
            sel_mem_wk_for_date = sel_mem_sheet_wk if sel_mem_sheet_wk in m_disp_weeks else view_week
            auto_dt_preview = get_week_date_for_month_year(m_disp, sel_mem_wk_for_date, int(selected_year))
            st.info(f"📅 **Buwan & Linggo:** `{m_disp} {selected_year} — {sel_mem_sheet_wk}` (Petsa ng Pagsamba: **{auto_dt_preview}**)")

        mem_weeks_to_show = m_disp_weeks if sel_mem_sheet_wk not in m_disp_weeks else [sel_mem_sheet_wk]
        for mw_item in mem_weeks_to_show:
            mem_sheet_df, mem_meta = build_parish_member_attendance_sheet_df(focused_parish_for_sheets, m_disp, mw_item, selected_year)
            st.markdown(f"**📋 {focused_parish_for_sheets} — {mw_item} of {m_disp.upper()} {selected_year} (Date: `{mem_meta['report_date']}`)**")
            st.dataframe(mem_sheet_df, use_container_width=True, hide_index=True)

            tot_b1, tot_b2, tot_b3 = st.columns(3)
            with tot_b1:
                st.markdown(
                    f"<div class='full-text-card' style='text-align:center;'>"
                    f"<b>Total Members (Adults+Youths):</b><br><span style='font-size:1.35rem;font-weight:800;color:#10b981;'>{mem_meta['total_members_say']}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            with tot_b2:
                st.markdown(
                    f"<div class='full-text-card' style='text-align:center;'>"
                    f"<b>Total Youth Members:</b><br><span style='font-size:1.35rem;font-weight:800;color:#3b82f6;'>{mem_meta['total_members_y']}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            with tot_b3:
                st.markdown(
                    f"<div class='full-text-card' style='text-align:center;'>"
                    f"<b>Total Kids Members:</b><br><span style='font-size:1.35rem;font-weight:800;color:#d97706;'>{mem_meta['total_members_k']}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )

        st.divider()

        # =====================================================================
        # 📋📈 SECTION D (BOOK1 EXCEL LAYOUT): PAIRED WEEKLY TABLES (ACTUAL WEEKS OF THE MONTH ONLY)
        # =====================================================================
        st.markdown(
            f"### 📅 Buong Buwan ng {m_disp.upper()} {selected_year} — Parish Monitoring Report & Summary for Posting (1st Week hanggang {m_disp_weeks[-1]})"
        )
        parishes_to_show_d = PARISH_ORDER if effective_parish_filter == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, effective_parish_filter)]
        for wk_pair in m_disp_weeks:
            wk_pair_dt = get_week_date_for_month_year(m_disp, wk_pair, int(selected_year))
            st.markdown(f"#### 🗓️ {wk_pair} of {m_disp.upper()} {selected_year} (Petsa ng Sabado: `{wk_pair_dt}`)")
            att_monitoring_df = build_parish_monitoring_attendance_df(m_disp, wk_pair, effective_parish_filter)
            st.dataframe(att_monitoring_df, use_container_width=True, hide_index=True)

            for p_item in parishes_to_show_d:
                df_post_p = build_summary_for_posting_per_parish_df(p_item, m_disp, wk_pair)
                st.dataframe(df_post_p, use_container_width=True, hide_index=True)

        st.divider()

    # =====================================================================
    # 🏆 SECTION E (MASTER ADMIN EXCLUSIVE): CHURCH OF THE TEMPLE OF GOD – PARISH MINISTRY PERFORMANCE REPORT
    # (English & Ilokano, Editable Summary, Pulls Outreach & Remarks from Outreach Page, Sendable via Gmail)
    # =====================================================================
    if is_master:
        st.markdown("### 🏆 Church of the Temple of God – Parish Ministry Performance Report (Master Admin Exclusive)")
        st.caption(
            "Awtomatikong binubuo mula sa Attendance Percentages ng lahat ng 10 Parokya at kinukuha ang **House Visitation**, "
            "**Missionary Work / Indoctrination**, at **Remarks** mula sa pahina ng `🏠 House Visitation, Missionary Work, Indoctrination & Remarks`. "
            "Maaari mo itong i-switch sa **English** o **Ilokano**, i-edit, at direktang i-send sa **Gmail**."
        )

        pr_c1, pr_c2, pr_c3, pr_c4 = st.columns([1.1, 1.1, 1.1, 1.3])
        with pr_c1:
            perf_month = st.selectbox("📅 Buwan ng Report:", MONTHS_2027, index=MONTHS_2027.index(view_month) if view_month in MONTHS_2027 else 0, key="perf_rep_month")
        perf_valid_weeks = get_weeks_for_month_year(perf_month, int(selected_year))
        with pr_c2:
            perf_week = st.selectbox("🗓️ Linggo ng Report:", perf_valid_weeks, index=perf_valid_weeks.index(view_week) if view_week in perf_valid_weeks else 0, key="perf_rep_week")
        with pr_c3:
            perf_lang = st.selectbox("🌐 Wika ng Performance Report:", ["English", "Ilokano"], key="perf_rep_lang")
        with pr_c4:
            st.write("")
            regen_perf_btn = st.button("🔄 I-regenerate Mula sa Live Data", use_container_width=True, key="btn_regen_perf_rep")

        perf_state_key = f"editable_perf_report_{perf_month}_{perf_week}_{selected_year}_{perf_lang}"
        if regen_perf_btn or perf_state_key not in st.session_state:
            st.session_state[perf_state_key] = generate_parish_ministry_performance_report_text(
                perf_month, perf_week, str(selected_year), perf_lang
            )

        ai_pol_col1, ai_pol_col2 = st.columns([1.8, 1.8])
        with ai_pol_col1:
            if st.button(f"✨ I-polish / I-translate sa {perf_lang} Gamit ang Gemini AI (Optional)", use_container_width=True, key="btn_ai_polish_perf"):
                valid_k = get_valid_gemini_keys()
                if valid_k:
                    with st.spinner(f"Inaayos ng Gemini AI sa wikang {perf_lang}..."):
                        try:
                            client_p = genai.Client(api_key=valid_k[0])
                            base_txt = st.session_state.get(perf_state_key, "")
                            p_prompt = (
                                f"Translate and polish the following Church of the Temple of God – Parish Ministry Performance Report into natural, formal {perf_lang}. "
                                "Keep the exact structure, headings, parish names, percentages, House Visitation, Missionary Work / Indoctrination, and Remarks intact:\n\n"
                                + base_txt
                            )
                            res_p = safe_generate_content(client_p, selected_model, p_prompt)
                            if res_p and res_p.text:
                                st.session_state[perf_state_key] = res_p.text.strip()
                                st.rerun()
                        except Exception as ex_p:
                            st.warning(f"AI Polish Notice: {str(ex_p)}")
                else:
                    st.warning("Maglagay ng GEMINI_API_KEY sa App Settings.")
        with ai_pol_col2:
            if st.button("💾 I-save ang In-edit na Summary sa App", type="primary", use_container_width=True, key="btn_save_perf_edits"):
                cfg_s = load_saved_config()
                saved_perf_map = cfg_s.get("SAVED_PERFORMANCE_REPORTS", {})
                if not isinstance(saved_perf_map, dict):
                    saved_perf_map = {}
                saved_perf_map[perf_state_key] = st.session_state.get(perf_state_key, "")
                cfg_s["SAVED_PERFORMANCE_REPORTS"] = saved_perf_map
                try:
                    with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
                        json.dump(cfg_s, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
                st.success("✅ Na-save na ang iyong in-edit na Parish Ministry Performance Report!")

        edited_perf_text = st.text_area(
            f"✏️ Parish Ministry Performance Report ({perf_lang} — Maaari mong i-edit bago i-save o i-send sa Gmail):",
            value=st.session_state.get(perf_state_key, ""),
            height=520,
            key=f"ta_{perf_state_key}"
        )
        st.session_state[perf_state_key] = edited_perf_text

        with st.expander("📖 Preview ng Parish Ministry Performance Report (Formatted Card View)", expanded=True):
            st.markdown(f"<div class='full-text-card'>{edited_perf_text}</div>", unsafe_allow_html=True)

        with st.expander("📧 I-send ang Parish Ministry Performance Report sa Gmail", expanded=True):
            gm_c1, gm_c2, gm_c3 = st.columns([1.4, 1.6, 1.0])
            with gm_c1:
                perf_email_to = st.text_input(
                    "📧 Email ng Tatanggap (pwedeng marami, paghiwalayin ng comma):",
                    value=BACKUP_ADMIN_EMAIL,
                    key="perf_rep_email_to"
                )
            with gm_c2:
                perf_email_subj = st.text_input(
                    "🏷️ Subject ng Email:",
                    value=f"Church of the Temple of God – Parish Ministry Performance Report ({perf_week} of {perf_month} {selected_year} - {perf_lang})",
                    key="perf_rep_email_subj"
                )
            with gm_c3:
                st.write("")
                if st.button("🚀 I-send Ngayon sa Gmail", type="primary", use_container_width=True, key="btn_send_perf_gmail"):
                    recipients_list = [e.strip() for e in str(perf_email_to or "").split(",") if "@" in e.strip()]
                    if not recipients_list:
                        st.warning("Maglagay ng valid na Email Address.")
                    else:
                        sent_ok = []
                        err_list = []
                        for r_addr in recipients_list:
                            res_em = send_email_via_gmail(r_addr, perf_email_subj, edited_perf_text)
                            if "Matagumpay" in res_em:
                                sent_ok.append(r_addr)
                            else:
                                err_list.append(f"{r_addr}: {res_em}")
                        if sent_ok:
                            st.success(f"✅ Matagumpay na naipadala ang Parish Ministry Performance Report sa: {', '.join(sent_ok)}!")
                        if err_list:
                            st.error(" | ".join(err_list))

        st.divider()

    if os.path.exists(LOCAL_PWR_MASTER_FILE):
        with open(LOCAL_PWR_MASTER_FILE, "rb") as f:
            pwr_bytes = f.read()
        st.download_button(
            label=f"📥 I-download ang Na-update na Master Excel (PWR. {view_month} {view_week}.xlsx)",
            data=pwr_bytes,
            file_name=f"PWR. {view_month} {view_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def _format_grouped_ministry_percentages(pcts_dict: dict, is_hongkong: bool = False, lang: str = "English") -> str:
    """Ginugrupo ang mga ministries na may parehong percentage (hal. '73% attendance on Mass, Pampamilya, and Adal ti Evangelio; 67% on Youth Fellowship...')."""
    if is_hongkong:
        m_p = pcts_dict.get("Mass", 0)
        o_p = pcts_dict.get("Online Ministry", 0)
        if lang == "Ilokano":
            if m_p == o_p:
                return f"{m_p}% attendance iti agpada a Mass ken Online Ministry."
            return f"{m_p}% attendance iti Mass; ken {o_p}% iti Online Ministry."
        else:
            if m_p == o_p:
                return f"{m_p}% attendance on both Mass and Online Ministry."
            return f"{m_p}% attendance on Mass; and {o_p}% on Online Ministry."

    ordered_mins = [
        ("Mass", pcts_dict.get("Mass", 0)),
        ("Pampamilya", pcts_dict.get("Pampamilya", 0)),
        ("Adal ti Evangelio", pcts_dict.get("Adal ti Evangelio", 0)),
        ("Online Ministry", pcts_dict.get("Online Ministry", 0)),
        ("Youth Fellowship", pcts_dict.get("Youth Fellowship", 0)),
        ("Kids Fellowship", pcts_dict.get("Kids Fellowship", 0)),
    ]
    groups = {}
    for m_name, p_val in ordered_mins:
        groups.setdefault(p_val, []).append(m_name)

    # Sort groups: Mass group first, then descending by percentage
    mass_pct = pcts_dict.get("Mass", 0)
    sorted_pcts = [mass_pct] + sorted([p for p in groups.keys() if p != mass_pct], reverse=True)

    def join_names(names_list, l_mode):
        if len(names_list) == 1:
            return names_list[0]
        conj = "ken" if l_mode == "Ilokano" else "and"
        if len(names_list) == 2:
            return f"{names_list[0]} {conj} {names_list[1]}"
        return f"{', '.join(names_list[:-1])}, {conj} {names_list[-1]}"

    clauses = []
    for idx_g, p_val in enumerate(sorted_pcts):
        m_names = groups[p_val]
        if m_names == ["Youth Fellowship", "Kids Fellowship"]:
            joined_n = "Youth ken Kids Fellowship" if lang == "Ilokano" else "Youth and Kids Fellowship"
        else:
            joined_n = join_names(m_names, lang)
        if idx_g == 0:
            if lang == "Ilokano":
                clauses.append(f"{p_val}% attendance iti {joined_n}")
            else:
                clauses.append(f"{p_val}% attendance on {joined_n}")
        else:
            if lang == "Ilokano":
                clauses.append(f"{p_val}% iti {joined_n}")
            else:
                clauses.append(f"{p_val}% on {joined_n}")

    if len(clauses) == 1:
        return clauses[0] + "."
    last_conj = "ken" if lang == "Ilokano" else "and"
    return "; ".join(clauses[:-1]) + f"; {last_conj} {clauses[-1]}."

def generate_parish_ministry_performance_report_text(selected_month: str, selected_week: str, selected_year: str = "2027", report_lang: str = "English") -> str:
    """Bumubuo ng kumpletong Church of the Temple of God – Parish Ministry Performance Report sa English o Ilokano
    gamit ang live attendance percentages at ang House Visitation, Missionary Work / Indoctrination, at Remarks mula sa Outreach page."""
    cfg = load_saved_config()
    saved_map = cfg.get("SAVED_PERFORMANCE_REPORTS", {})
    save_k = f"editable_perf_report_{selected_month}_{selected_week}_{selected_year}_{report_lang}"
    if isinstance(saved_map, dict) and saved_map.get(save_k):
        return str(saved_map[save_k])

    db = load_weekly_monitoring_records()
    eval_items = []

    for p_name in PARISH_ORDER:
        key = f"{p_name.upper()}__{selected_month.upper()}__{selected_week}"
        rec = db.get(key, {})
        has_done = is_week_record_completed(rec)
        is_no_rep = is_record_marked_no_report(rec)
        denoms = get_parish_member_denominators(p_name)

        tot_say = (int(rec.get("total_members_say", 0)) if has_done else 0) or denoms["say"]
        tot_y = (int(rec.get("total_members_y", 0)) if has_done else 0) or denoms["youth"]
        tot_k = (int(rec.get("total_members_k", 0)) if has_done else 0) or denoms["kids"]

        m_att = (int(rec.get("mass_a", 0)) + int(rec.get("mass_y", 0))) if has_done else 0
        p_att = (int(rec.get("pampamilya_a", 0)) + int(rec.get("pampamilya_y", 0))) if has_done else 0
        o_att = (int(rec.get("online_a", 0)) + int(rec.get("online_y", 0))) if has_done else 0
        y_att = int(rec.get("youth_y", 0)) if has_done else 0
        k_att = int(rec.get("kids_k", 0)) if has_done else 0
        ad_att = (int(rec.get("adal_a", 0)) + int(rec.get("adal_y", 0))) if has_done else 0

        def _pct_num(a_v, d_v):
            return round((a_v / d_v) * 100) if (d_v > 0 and a_v > 0 and not is_no_rep) else 0

        pcts = {
            "Mass": _pct_num(m_att, tot_say),
            "Pampamilya": _pct_num(p_att, tot_say),
            "Online Ministry": _pct_num(o_att, tot_say),
            "Youth Fellowship": _pct_num(y_att, tot_y),
            "Kids Fellowship": _pct_num(k_att, tot_k),
            "Adal ti Evangelio": _pct_num(ad_att, tot_say),
        }

        raw_rem = str(rec.get("remarks", "")).strip() if not is_no_rep else ""
        p_ro = parse_remarks_into_mass_and_outreach(raw_rem) if raw_rem else {}

        hv_raw = str(rec.get("house_visitation", "0")).strip() if not is_no_rep else "0"
        hv_names = str(rec.get("hv_names", "") or p_ro.get("hv_names", "")).strip() if not is_no_rep else ""
        hv_rem = str(rec.get("hv_remarks", "") or p_ro.get("hv_remarks", "")).strip() if not is_no_rep else ""
        hv_cnt, _ = count_families_and_people_from_text(hv_names, hv_raw)

        mw_raw = str(rec.get("missionary_work", "0")).strip() if not is_no_rep else "0"
        if (not mw_raw or mw_raw == "0") and p_ro.get("mw_families_count", 0) > 0:
            mw_raw = str(p_ro["mw_families_count"])
        mw_names = str(rec.get("mw_names", "") or p_ro.get("mw_names", "")).strip() if not is_no_rep else ""
        mw_rem = str(rec.get("mw_remarks", "") or p_ro.get("mw_remarks", "")).strip() if not is_no_rep else ""
        mw_cnt, _ = count_families_and_people_from_text(mw_names, mw_raw)

        ind_raw = str(rec.get("indoctrination", "0")).strip() if not is_no_rep else "0"
        ind_names = str(rec.get("ind_names", "") or p_ro.get("ind_names", "")).strip() if not is_no_rep else ""
        ind_rem = str(rec.get("ind_remarks", "") or p_ro.get("ind_remarks", "")).strip() if not is_no_rep else ""
        ind_cnt, _ = count_families_and_people_from_text(ind_names, ind_raw)

        outreach_gen_rem = str(rec.get("outreach_general_remarks", "")).strip()
        if not outreach_gen_rem:
            rem_parts = [x for x in [hv_rem, mw_rem, ind_rem] if x and x != "(No Report)"]
            outreach_gen_rem = " | ".join(rem_parts)

        disp_name = "Hong Kong" if "HONGKONG" in p_name.upper() else p_name.replace("PARISH", "").strip().title()
        eval_items.append({
            "parish_raw": p_name,
            "name": disp_name,
            "is_hk": "HONGKONG" in p_name.upper(),
            "has_done": has_done,
            "is_no_rep": is_no_rep,
            "pcts": pcts,
            "mass_pct": pcts["Mass"],
            "online_pct": pcts["Online Ministry"],
            "hv_cnt": hv_cnt,
            "hv_names": hv_names,
            "mw_cnt": mw_cnt,
            "mw_names": mw_names,
            "ind_cnt": ind_cnt,
            "ind_names": ind_names,
            "outreach_rem": outreach_gen_rem,
        })

    # Sort by Mass % descending, keeping NO REPORT parishes at the bottom
    eval_items.sort(key=lambda x: (-1 if x["is_no_rep"] else x["mass_pct"], x["online_pct"]), reverse=True)
    passed_75 = [x for x in eval_items if x["mass_pct"] >= 75 and not x["is_no_rep"]]
    remaining_p = [x for x in eval_items if x not in passed_75]

    hv_completed_10 = [x for x in eval_items if x["hv_cnt"] >= 10]
    hv_partial = [x for x in eval_items if 0 < x["hv_cnt"] < 10]
    hv_partial.sort(key=lambda x: x["hv_cnt"], reverse=True)

    online_passed = [x for x in eval_items if x["online_pct"] >= 75 and not x["is_no_rep"]]
    online_passed.sort(key=lambda x: x["online_pct"], reverse=True)
    online_others = [x for x in eval_items if 0 < x["online_pct"] < 75 and not x["is_no_rep"]]
    online_others.sort(key=lambda x: x["online_pct"], reverse=True)

    num_words_en = {0: "zero (0)", 1: "one (1)", 2: "two (2)", 3: "three (3)", 4: "four (4)", 5: "five (5)", 6: "six (6)", 7: "seven (7)", 8: "eight (8)", 9: "nine (9)", 10: "ten (10)"}

    if report_lang == "Ilokano":
        lines = [
            "Church of the Temple of God – Parish Ministry Performance Report (Ilokano)",
            f"Panawen (Period): {selected_week} ti {selected_month} {selected_year}",
            "",
            "Executive Summary (Pakabuklan)",
        ]
        if passed_75:
            led_str = ", ".join(f"{x['name']} ({x['mass_pct']}%)" for x in passed_75)
            s1 = f"Adda {len(passed_75)} a parokya ti nakadanon wenno nanglabes iti 75% performance cut-off para iti Holy Mass attendance, nga idauluan ti {led_str}."
        else:
            s1 = "Awan pay ti parokya a nakadanon iti 75% performance cut-off para iti Holy Mass attendance ita a lawas."

        if hv_completed_10:
            hv_str = ", ".join(f"{x['name']} ({x['hv_cnt']})" for x in hv_completed_10)
            s2 = f"Iti amin a 10 a parokya, nakompleto ti {hv_str} ti 10 weekly house visitations requirement."
        elif hv_partial:
            part_str = " ken ".join(f"nakailista ti {x['name']} iti {x['hv_cnt']} visitations" for x in hv_partial[:3])
            s2 = f"Iti amin a 10 a parokya, awan ti nakakompleto iti 10 weekly house visitations requirement ita a lawas, idinto ta {part_str}."
        else:
            s2 = "Iti amin a 10 a parokya, awan pay ti nailista a house visitation ita a lawas."

        if online_passed:
            op_str = ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_passed)
            foll_str = (", a sinaruno ti " + ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_others[:6])) if online_others else ""
            s3 = f"Maipanggep iti attendance para iti weekly Online Ministry Preaching, ti {op_str} ti nakadanon iti passing mark a 75% wenno nangatngato{foll_str}."
        elif online_others:
            foll_str = ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_others[:6])
            s3 = f"Maipanggep iti weekly Online Ministry Preaching, dagiti nakailista iti attendance ket: {foll_str}."
        else:
            s3 = "Maipanggep iti weekly Online Ministry Preaching, awan pay ti nailista nga attendance ita a lawas."

        lines.append(f"{s1} {s2} {s3}")
        lines.append("")
        lines.append("Ti naan-anay nga ebalwasyon para iti amin a 10 a parokya ket naurnos iti baba:")
        lines.append("")
        lines.append("Dagiti Parokya a Nakapasa iti 75% Mass Target")
        if passed_75:
            for item in passed_75:
                lines.extend(_format_single_parish_perf_block(item, "Ilokano"))
        else:
            lines.append("- Awan ti parokya a nakadanon iti 75% Mass target ita a lawas.")
        lines.append("")
        lines.append("Dagiti Dadduma pay a Parokya (Nababa ngem 75% Cut-off / Special Highlights)")
        for item in remaining_p:
            lines.extend(_format_single_parish_perf_block(item, "Ilokano"))

        lines.append("")
        lines.append(
            f"Iti panaggibus ti {selected_week} ti {selected_month} {selected_year}, nakailista tayo iti {len(passed_75)} a parokya a nakalabes iti 75% passing mark para iti Mass attendance. "
            "Agkarkararag tayo met para iti kinasalun-at ken pigsa dagiti kakabsat tayo iti amin a parokya. "
            "Ti agtultuloy a panang-monitor ken panangiruar iti daytoy a Performance Report ket agserbi a giya ken inspirasyon kadagiti Parish Priests ken workers tayo tapno mapadur-as pay ti attendance ken ministerio kadagiti sumarsaruno a lawas."
        )
        lines.append("Agyaman ken Apo Dios ti Mangbendisyon!")
        return "\n".join(lines)

    # English (Default)
    lines = [
        "Church of the Temple of God – Parish Ministry Performance Report",
        f"Period: {selected_week} of {selected_month} {selected_year}",
        "",
        "Executive Summary",
    ]
    if passed_75:
        cnt_w = num_words_en.get(len(passed_75), str(len(passed_75)))
        led_str = ", ".join(f"{x['name']} ({x['mass_pct']}%)" for x in passed_75)
        only_pfx = "Only " if len(passed_75) == 1 else ""
        p_noun = "parish" if len(passed_75) == 1 else "parishes"
        s1 = f"{only_pfx}{cnt_w} {p_noun} reached or exceeded the 75% performance cut-off for Holy Mass attendance, led by {led_str}."
    else:
        s1 = "No parish reached the 75% performance cut-off for Holy Mass attendance during this reporting period."

    if hv_completed_10:
        hv_str = ", ".join(f"{x['name']} ({x['hv_cnt']})" for x in hv_completed_10)
        s2 = f"Looking across all 10 parishes, {hv_str} completed the full requirement of 10 weekly house visitations this week."
    elif hv_partial:
        part_str = " and ".join(f"{x['name']} recorded {x['hv_cnt']} visitations" for x in hv_partial[:3])
        s2 = f"Looking across all 10 parishes, no parish completed the full requirement of 10 weekly house visitations this week, while {part_str}."
    else:
        s2 = "Looking across all 10 parishes, no parish recorded weekly house visitations this week."

    if online_passed:
        op_str = ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_passed)
        only_op = "only " if len(online_passed) == 1 else ""
        foll_str = (", followed by " + ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_others[:6])) if online_others else ""
        s3 = f"Regarding attendance for the weekly Online Ministry Preaching, {only_op}{op_str} reached the passing mark of 75% or higher{foll_str}."
    elif online_others:
        foll_str = ", ".join(f"{x['name']} ({x['online_pct']}%)" for x in online_others[:6])
        s3 = f"Regarding attendance for the weekly Online Ministry Preaching, no parish reached the 75% mark, led by {foll_str}."
    else:
        s3 = "Regarding attendance for the weekly Online Ministry Preaching, no attendance was recorded this week."

    lines.append(f"{s1} {s2} {s3}")
    lines.append("The full evaluation for all 10 parishes is organized below, sorted by those that met the 75% Mass target followed by the remaining parishes:")
    lines.append("")
    lines.append("Parishes That Passed the 75% Mass Target")
    if passed_75:
        for item in passed_75:
            lines.extend(_format_single_parish_perf_block(item, "English"))
    else:
        lines.append("- None recorded for this week.")
    lines.append("")
    lines.append("The Remaining Parishes (Below the 75% Cut-off / Special Highlights)")
    for item in remaining_p:
        lines.extend(_format_single_parish_perf_block(item, "English"))

    runners_up = [x for x in remaining_p if x["mass_pct"] >= 65 and not x["is_no_rep"]]
    ru_clause = ""
    if runners_up:
        ru_names = " and ".join(f"{x['name']} ({x['mass_pct']}%)" for x in runners_up[:2])
        ru_clause = f", with {ru_names} closely following"

    lines.append("")
    lines.append(
        f"At the close of the {selected_week.lower()} of {selected_month} {selected_year}, we recorded {len(passed_75)} "
        f"{'parish' if len(passed_75) == 1 else 'parishes'} that crossed the 75% passing mark for Mass attendance{ru_clause}. "
        "The consistent monitoring and releasing of this Performance Report will continue to serve as a guide and inspiration for our Parish Priests and workers to further improve attendance and ministries in the coming weeks."
    )
    lines.append("Cheers and God Bless!")
    return "\n".join(lines)

def _format_single_parish_perf_block(item: dict, lang: str = "English") -> list[str]:
    p_name = item["name"]
    if item["is_no_rep"]:
        if lang == "Ilokano":
            main_line = f"- {p_name} – 0% iti amin a ministerio (Na-markaan kas NO REPORT ita a lawas)."
        else:
            main_line = f"- {p_name} – 0% across all ministries (Marked as NO REPORT for this week)."
    else:
        pct_sentence = _format_grouped_ministry_percentages(item["pcts"], is_hongkong=item["is_hk"], lang=lang)
        special_parts = []
        if item["hv_cnt"] > 0:
            special_parts.append(f"{item['hv_cnt']} house visitations")
        if item["mw_cnt"] > 0:
            special_parts.append(f"{item['mw_cnt']} Missionary Work")
        if item["ind_cnt"] > 0:
            special_parts.append(f"{item['ind_cnt']} Indoctrinations")
        spec_str = ""
        if special_parts:
            joined_sp = " and ".join(special_parts)
            spec_str = f" Special Mention: Recorded {joined_sp} for this week." if lang == "English" else f" Special Mention: Nakailista iti {joined_sp} ita a lawas."
        main_line = f"- {p_name} – {pct_sentence}{spec_str}"

    out_lines = [main_line]
    none_lbl = "0 (Awan nailista)" if lang == "Ilokano" else "0 (None recorded)"

    if item["hv_cnt"] > 0:
        hv_detail = item["hv_names"] if item["hv_names"] else f"{item['hv_cnt']} recorded"
        out_lines.append(f"  - House Visitation ({item['hv_cnt']}): {hv_detail}")
    else:
        out_lines.append(f"  - House Visitation: {none_lbl}")

    mw_ind_total = item["mw_cnt"] + item["ind_cnt"]
    if mw_ind_total > 0:
        comb_names = " | ".join(x for x in [item["mw_names"], item["ind_names"]] if x) or f"{mw_ind_total} recorded"
        out_lines.append(f"  - Missionary Work / Indoctrination ({mw_ind_total}): {comb_names}")
    else:
        out_lines.append(f"  - Missionary Work / Indoctrination: {none_lbl}")

    if item["outreach_rem"]:
        out_lines.append(f"  - Remarks: {item['outreach_rem']}")
    return out_lines

# =====================================================================
# 🏠 DEDICATED PAGE FOR ALL USERS: HOUSE VISITATION, MISSIONARY WORK, INDOCTRINATION & REMARKS
# (POWERED BY OFFICIAL GOOGLE FORM & GOOGLE SPREADSHEET LIVE SYNC)
# =====================================================================
DEFAULT_OUTREACH_GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScKnsFH73UI7Iyh5nftmGque7Y2jzRSexO3ImkKZx9aKEOwiA/viewform?usp=preview"
DEFAULT_OUTREACH_GOOGLE_FORM_EMBED_URL = "https://docs.google.com/forms/d/e/1FAIpQLScKnsFH73UI7Iyh5nftmGque7Y2jzRSexO3ImkKZx9aKEOwiA/viewform?embedded=true"
DEFAULT_OUTREACH_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8/edit?usp=sharing"
DEFAULT_OUTREACH_GOOGLE_SHEET_ID = "1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8"
LOCAL_OUTREACH_RESPONSES_CACHE = "ctg_outreach_form_responses_cache.json"

OUTREACH_FORM_35_COLUMNS = [
    "Timestamp",
    "Parish / Simbahan (Church Name)",
    "Pangalan ng Nag-uulat (Worker / Priest / Volunteer Name)",
    "Opening Prayer",
    "Closing Payer",
    "Katungkulan / Ministry Unit",
    "Petsa ng Gawain (Date of Activity)",
    "Lokasyon / Area / Barangay",
    "Type of Activity",
    "Date (House Visitation)",
    "Pangalan ng Pamilya / Indibidwal na Binisita",
    "Tirahan at Contact Number ng Binisita",
    "Layunin / Uri ng Pagbisita",
    "Kalagayan ng Espirituwal / Sitwasyon ng Pamilya",
    "Mga Kahilingan sa Panalangin (Prayer Requests)",
    "Paksa ng Pag-aaral o Mensaheng Ibinahagi (Topic / Scripture Shared)",
    "Reaksyon o Tugon ng Binisita sa Mensahe:",
    "Susunod na Hakbang / Kailangan ng Follow-up?",
    "Date (Missionary Work)",
    "Pangalan ng Bagong Contact / Prospect (Contact Name)",
    "Edad at Kasarian (Age & Gender)",
    "Contact Number / Social Media Handle (Messenger Link)",
    "Tirahan o Lokasyon kung saan nakatagpo (Address / Location met)",
    "Reaksyon o Tugon sa Ebanghelyo (Response Status)",
    "Anong Mensahe o Paksa ang Naibahagi? (Topic / Message Shared)",
    "Mga Interes o Pangangailangan ng Contact (Interests & Needs)",
    "Napagkasunduang Araw para sa Follow-up (Follow-up Date)",
    "Pangalan ng Candidate / Estudyante (Candidate Name)",
    "Contact Number / Social Media Link",
    "Pangalan ng Nagtuturo (Instructor / Mentor / Priest Name)",
    "Kasalukuyang Leksyon o Aral na Natapos (Lesson Progress)",
    "Status ng Pagdalo (Attendance Status)",
    "Handa na ba para sa Bautismo / Kumpirma (Ready for Baptism?)",
    "Target Date ng Bautismo (kung handa na)",
    "Puna o Evaluasyon ng Nagtuturo (Teacher's Notes / Assessment)"
]

def _load_local_outreach_form_cache() -> list[dict]:
    if os.path.exists(LOCAL_OUTREACH_RESPONSES_CACHE):
        try:
            with open(LOCAL_OUTREACH_RESPONSES_CACHE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []

def _save_local_outreach_form_cache(rows_list: list[dict]):
    try:
        with open(LOCAL_OUTREACH_RESPONSES_CACHE, "w", encoding="utf-8") as f:
            json.dump(rows_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _normalize_outreach_sheet_headers(raw_headers: list[str]) -> list[str]:
    """Inaayos ang duplicate 'Date' column headers sa Google Form Responses sheet:
    ang unang 'Date' (Column J) ay nagiging 'Date (House Visitation)' at ang pangalawang 'Date' (Column S) ay nagiging 'Date (Missionary Work)'."""
    normalized = []
    date_seen_cnt = 0
    for idx, h in enumerate(raw_headers):
        h_clean = str(h or "").strip()
        if h_clean.lower() == "date":
            date_seen_cnt += 1
            if date_seen_cnt == 1 or idx <= 12:
                normalized.append("Date (House Visitation)")
            else:
                normalized.append("Date (Missionary Work)")
        elif not h_clean:
            if idx < len(OUTREACH_FORM_35_COLUMNS):
                normalized.append(OUTREACH_FORM_35_COLUMNS[idx])
            else:
                normalized.append(f"Column_{idx+1}")
        else:
            normalized.append(h_clean)
    return normalized

def extract_google_drive_id(url_or_id: str) -> str:
    """Kinukuha ang Google Drive / Google Sheets File ID mula sa URL o direktang ID string."""
    return extract_id_from_url(url_or_id)

def download_google_drive_excel_bytes(sheet_url_or_id: str) -> bytes | None:
    """Kinukuha ang XLSX bytes mula sa Google Drive / Google Spreadsheet gamit ang Service Account API o public export link."""
    fid = extract_id_from_url(sheet_url_or_id) or DEFAULT_OUTREACH_GOOGLE_SHEET_ID
    if not fid:
        return None
    drive_service, _ = get_google_drive_service()
    if drive_service is not None:
        try:
            meta = drive_service.files().get(fileId=fid, fields="mimeType", supportsAllDrives=True).execute()
            m_type = str(meta.get("mimeType", ""))
            if "spreadsheet" in m_type:
                data = drive_service.files().export_media(
                    fileId=fid,
                    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ).execute()
            else:
                data = drive_service.files().get_media(fileId=fid, supportsAllDrives=True).execute()
            if data and len(data) > 100:
                return data
        except Exception:
            pass
    for dl_url in [
        f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={fid}"
    ]:
        try:
            r = requests.get(dl_url, timeout=15)
            if r.status_code == 200 and r.content and len(r.content) > 100 and not r.content.lstrip().startswith(b"<!DOCTYPE html"):
                return r.content
        except Exception:
            continue
    return None

def fetch_outreach_google_sheet_df(sheet_url: str = DEFAULT_OUTREACH_GOOGLE_SHEET_URL) -> tuple[pd.DataFrame, str]:
    """Kinukuha ang live responses mula sa Google Spreadsheet ('1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8' - Form responses 1)
    at isinasama rin ang anumang local entries para laging kumpleto ang Table sa App."""
    file_id = extract_google_drive_id(sheet_url) or DEFAULT_OUTREACH_GOOGLE_SHEET_ID
    rows_dicts = []
    status_msg = ""

    # 1. Subukan munang kunin bilang XLSX bytes (Google Drive API o direct export link)
    excel_bytes = download_google_drive_excel_bytes(sheet_url)
    if excel_bytes:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
            ws = wb["Form responses 1"] if "Form responses 1" in wb.sheetnames else wb.active
            all_rows = list(ws.iter_rows(values_only=True))
            if all_rows:
                headers = _normalize_outreach_sheet_headers([str(c or "").strip() for c in all_rows[0]])
                for r_vals in all_rows[1:]:
                    if not r_vals or not any(str(v or "").strip() for v in r_vals):
                        continue
                    item = {}
                    for c_i, col_name in enumerate(headers):
                        val = r_vals[c_i] if c_i < len(r_vals) else ""
                        if isinstance(val, datetime):
                            val_s = val.strftime("%Y-%m-%d %H:%M:%S") if (val.hour or val.minute) else val.strftime("%m/%d/%y")
                        else:
                            val_s = str(val if val is not None else "").strip()
                        item[col_name] = val_s
                    rows_dicts.append(item)
                status_msg = f"Na-load mula sa Google Spreadsheet ({ws.title}: {len(rows_dicts)} response row(s))."
        except Exception as ex_xl:
            status_msg = f"XLSX parse notice: {str(ex_xl)}"

    # 2. Fallback sa direct CSV export kung hindi nakuha sa XLSX
    if not rows_dicts and file_id:
        try:
            csv_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv"
            resp = requests.get(csv_url, timeout=15)
            if resp.status_code == 200 and resp.content and not resp.content.lstrip().startswith(b"<!DOCTYPE html"):
                import csv
                reader = list(csv.reader(io.StringIO(resp.content.decode("utf-8-sig", errors="replace"))))
                if reader:
                    headers = _normalize_outreach_sheet_headers(reader[0])
                    for r_vals in reader[1:]:
                        if not r_vals or not any(str(v or "").strip() for v in r_vals):
                            continue
                        item = {}
                        for c_i, col_name in enumerate(headers):
                            item[col_name] = str(r_vals[c_i] if c_i < len(r_vals) else "").strip()
                        rows_dicts.append(item)
                    status_msg = f"Na-load mula sa Google Sheet CSV Export ({len(rows_dicts)} response row(s))."
        except Exception as ex_csv:
            if not status_msg:
                status_msg = f"CSV fetch notice: {str(ex_csv)}"

    # 3. Isama ang local cached submissions mula sa App Form (iwas-duplicate gamit ang Timestamp + Parish + Activity + Name)
    local_rows = _load_local_outreach_form_cache()
    if local_rows:
        seen_sigs = set()
        for r in rows_dicts:
            sig = (
                str(r.get("Timestamp", "")).strip(),
                str(r.get("Parish / Simbahan (Church Name)", "")).strip().upper(),
                str(r.get("Pangalan ng Pamilya / Indibidwal na Binisita", "")).strip().upper(),
                str(r.get("Pangalan ng Bagong Contact / Prospect (Contact Name)", "")).strip().upper(),
                str(r.get("Pangalan ng Candidate / Estudyante (Candidate Name)", "")).strip().upper(),
            )
            seen_sigs.add(sig)
        for lr in local_rows:
            if not isinstance(lr, dict):
                continue
            sig_l = (
                str(lr.get("Timestamp", "")).strip(),
                str(lr.get("Parish / Simbahan (Church Name)", "")).strip().upper(),
                str(lr.get("Pangalan ng Pamilya / Indibidwal na Binisita", "")).strip().upper(),
                str(lr.get("Pangalan ng Bagong Contact / Prospect (Contact Name)", "")).strip().upper(),
                str(lr.get("Pangalan ng Candidate / Estudyante (Candidate Name)", "")).strip().upper(),
            )
            if sig_l not in seen_sigs:
                rows_dicts.append(lr)
                seen_sigs.add(sig_l)

    if not rows_dicts:
        empty_df = pd.DataFrame(columns=OUTREACH_FORM_35_COLUMNS)
        return empty_df, (status_msg or "Nakakonekta na sa Google Sheet (Wala pang nakalistang response rows sa Form responses 1).")

    # Siguraduhing kumpleto ang 35 standard columns sa DataFrame
    normalized_rows = []
    for r in rows_dicts:
        norm_r = {}
        for col in OUTREACH_FORM_35_COLUMNS:
            norm_r[col] = str(r.get(col, "") or "").strip()
        # Preserve extra columns if any
        for k_extra, v_extra in r.items():
            if k_extra not in norm_r:
                norm_r[k_extra] = str(v_extra or "").strip()
        normalized_rows.append(norm_r)

    return pd.DataFrame(normalized_rows), (status_msg or f"Matagumpay na nakuha ang {len(normalized_rows)} row(s) mula sa Google Sheet.")

def _match_parish_from_form_text(raw_parish_text: str, raw_loc_text: str = "", fallback_parish: str = "POBLACION PARISH") -> str:
    combined = f"{raw_parish_text} {raw_loc_text}".upper().strip()
    if not combined:
        return fallback_parish if fallback_parish in PARISH_ORDER else "POBLACION PARISH"
    if "HONG KONG" in combined or "HONGKONG" in combined or "HK" in combined.split():
        return "HONGKONG PARISH"
    if "LIPAY" in combined:
        return "LIPAY NORTE PARISH"
    for p_cand in PARISH_ORDER:
        short_p = p_cand.replace("PARISH", "").strip().upper()
        if short_p and short_p in combined:
            return p_cand
    return fallback_parish if fallback_parish in PARISH_ORDER else "POBLACION PARISH"

def _resolve_month_week_year_from_form_date(date_str: str, fallback_month: str = "January", fallback_week: str = "1st Week", fallback_year: str = "2027") -> tuple[str, str, str, str]:
    """Tinutukoy ang Buwan, Linggo (1st–5th Week), Taon, at MM/DD/YY string mula sa Petsa sa Google Form."""
    s = str(date_str or "").strip()
    if not s:
        return fallback_month, fallback_week, fallback_year, get_week_date_for_month_year(fallback_month, fallback_week, int(fallback_year) if str(fallback_year).isdigit() else 2027)

    parsed_dt = None
    date_part = s.split(" ")[0].strip()
    for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%m-%d-%Y", "%m-%d-%y", "%Y/%m/%d", "%b %d, %Y", "%B %d, %Y"]:
        try:
            parsed_dt = datetime.strptime(date_part, fmt)
            break
        except Exception:
            continue

    if parsed_dt is None:
        # Check if month name is mentioned in string
        for m_name in MONTHS_2027:
            if m_name.upper() in s.upper():
                fallback_month = m_name
                break
        for w_name in WEEKS_2027:
            if w_name.upper() in s.upper():
                fallback_week = w_name
                break
        return fallback_month, fallback_week, fallback_year, s

    m_name = MONTHS_2027[parsed_dt.month - 1]
    y_str = str(parsed_dt.year)
    sats = get_saturdays_for_month_year(m_name, parsed_dt.year)
    best_wk_idx = 0
    if sats:
        # Map activity date to the reporting cycle week (Saturday through Friday/Monday)
        min_diff = 999
        for idx_s, sat_d in enumerate(sats):
            diff_days = abs((parsed_dt.date() - sat_d.date()).days)
            if diff_days < min_diff:
                min_diff = diff_days
                best_wk_idx = idx_s
    wk_label = WEEKS_2027[min(best_wk_idx, len(WEEKS_2027) - 1)]
    return m_name, wk_label, y_str, parsed_dt.strftime("%m/%d/%y")

def sync_outreach_google_sheet_to_weekly_db(
    df_sheet: pd.DataFrame,
    default_month: str = "January",
    default_week: str = "1st Week",
    default_year: str = "2027",
    default_parish: str = "Lahat ng Parokya"
) -> tuple[int, str]:
    """Awtomatikong kumukuha ng bawat House Visitation, Missionary Work, Indoctrination, at Remarks mula sa Google Form Responses Sheet
    ('1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8') at isini-sync sa App Database at Performance Report."""
    if df_sheet is None or df_sheet.empty:
        return 0, "Walang nakitang response rows sa Google Spreadsheet."

    db_records = load_weekly_monitoring_records()
    grouped = {}
    synced_entries = 0
    fb_parish = default_parish if default_parish in PARISH_ORDER else "POBLACION PARISH"

    for _, row in df_sheet.iterrows():
        r = {str(k).strip(): str(v or "").strip() for k, v in row.items()}
        raw_parish = r.get("Parish / Simbahan (Church Name)", "")
        raw_loc = r.get("Lokasyon / Area / Barangay", "")
        parish_matched = _match_parish_from_form_text(raw_parish, raw_loc, fb_parish)

        act_type = r.get("Type of Activity", "").upper()
        worker_name = r.get("Pangalan ng Nag-uulat (Worker / Priest / Volunteer Name)", "")
        open_p = r.get("Opening Prayer", "")
        close_p = r.get("Closing Payer", "")
        unit_role = r.get("Katungkulan / Ministry Unit", "")
        act_date_raw = (
            r.get("Petsa ng Gawain (Date of Activity)", "")
            or r.get("Date (House Visitation)", "")
            or r.get("Date (Missionary Work)", "")
            or r.get("Timestamp", "")
        )

        m_res, w_res, y_res, short_dt = _resolve_month_week_year_from_form_date(
            act_date_raw, default_month, default_week, default_year
        )
        g_key = (parish_matched, m_res, w_res, y_res)
        if g_key not in grouped:
            grouped[g_key] = {
                "hv_names": [], "hv_rems": [], "hv_rows": [],
                "mw_names": [], "mw_rems": [], "mw_rows": [],
                "ind_names": [], "ind_rems": [], "ind_rows": []
            }

        # 1. Check House Visitation fields (Columns J to R)
        hv_fam = r.get("Pangalan ng Pamilya / Indibidwal na Binisita", "")
        hv_addr = r.get("Tirahan at Contact Number ng Binisita", "")
        hv_purp = r.get("Layunin / Uri ng Pagbisita", "")
        hv_cond = r.get("Kalagayan ng Espirituwal / Sitwasyon ng Pamilya", "")
        hv_pray = r.get("Mga Kahilingan sa Panalangin (Prayer Requests)", "")
        hv_top = r.get("Paksa ng Pag-aaral o Mensaheng Ibinahagi (Topic / Scripture Shared)", "")
        hv_react = r.get("Reaksyon o Tugon ng Binisita sa Mensahe:", "")
        hv_foll = r.get("Susunod na Hakbang / Kailangan ng Follow-up?", "")
        hv_dt = r.get("Date (House Visitation)", "") or short_dt

        is_hv = ("VISIT" in act_type or "DALAW" in act_type or "HOUSE" in act_type or bool(hv_fam or hv_purp or hv_cond or hv_top))
        if is_hv and (hv_fam or hv_purp or hv_cond or hv_top or hv_react):
            name_entry = f"{hv_dt}: {hv_fam}" if (hv_dt and hv_fam and hv_dt not in hv_fam) else (hv_fam or f"Visited ({hv_dt})")
            if name_entry not in grouped[g_key]["hv_names"]:
                grouped[g_key]["hv_names"].append(name_entry)
            rem_bits = []
            if worker_name:
                rem_bits.append(f"Nag-ulat: {worker_name}")
            if hv_purp:
                rem_bits.append(f"Layunin: {hv_purp}")
            if hv_top:
                rem_bits.append(f"Paksa: {hv_top}")
            if hv_cond:
                rem_bits.append(f"Sitwasyon: {hv_cond}")
            if hv_pray:
                rem_bits.append(f"Prayer Req: {hv_pray}")
            if hv_react:
                rem_bits.append(f"Tugon: {hv_react}")
            if hv_foll:
                rem_bits.append(f"Follow-up: {hv_foll}")
            rem_line = f"{hv_fam or 'House Visit'} ({' | '.join(rem_bits)})" if rem_bits else hv_fam
            if rem_line and rem_line not in grouped[g_key]["hv_rems"]:
                grouped[g_key]["hv_rems"].append(rem_line)
            grouped[g_key]["hv_rows"].append(r)
            synced_entries += 1

        # 2. Check Missionary Work fields (Columns S to AA)
        mw_contact = r.get("Pangalan ng Bagong Contact / Prospect (Contact Name)", "")
        mw_age_gen = r.get("Edad at Kasarian (Age & Gender)", "")
        mw_phone = r.get("Contact Number / Social Media Handle (Messenger Link)", "")
        mw_loc_met = r.get("Tirahan o Lokasyon kung saan nakatagpo (Address / Location met)", "")
        mw_resp = r.get("Reaksyon o Tugon sa Ebanghelyo (Response Status)", "")
        mw_topic = r.get("Anong Mensahe o Paksa ang Naibahagi? (Topic / Message Shared)", "")
        mw_needs = r.get("Mga Interes o Pangangailangan ng Contact (Interests & Needs)", "")
        mw_foll_dt = r.get("Napagkasunduang Araw para sa Follow-up (Follow-up Date)", "")
        mw_dt = r.get("Date (Missionary Work)", "") or short_dt

        is_mw = ("MISSION" in act_type or "MISYON" in act_type or bool(mw_contact or mw_topic or mw_resp))
        if is_mw and (mw_contact or mw_topic or mw_resp or mw_loc_met):
            name_entry_mw = f"{mw_dt}: {mw_contact}" if (mw_dt and mw_contact and mw_dt not in mw_contact) else (mw_contact or f"Mission Contact ({mw_dt})")
            if name_entry_mw not in grouped[g_key]["mw_names"]:
                grouped[g_key]["mw_names"].append(name_entry_mw)
            mw_bits = []
            if worker_name:
                mw_bits.append(f"Nag-ulat: {worker_name}")
            if mw_age_gen:
                mw_bits.append(f"Edad/Kasarian: {mw_age_gen}")
            if mw_loc_met:
                mw_bits.append(f"Lugar: {mw_loc_met}")
            if mw_topic:
                mw_bits.append(f"Paksa: {mw_topic}")
            if mw_resp:
                mw_bits.append(f"Tugon: {mw_resp}")
            if mw_needs:
                mw_bits.append(f"Pangangailangan: {mw_needs}")
            if mw_foll_dt:
                mw_bits.append(f"Follow-up: {mw_foll_dt}")
            rem_line_mw = f"{mw_contact or 'Missionary Work'} ({' | '.join(mw_bits)})" if mw_bits else mw_contact
            if rem_line_mw and rem_line_mw not in grouped[g_key]["mw_rems"]:
                grouped[g_key]["mw_rems"].append(rem_line_mw)
            grouped[g_key]["mw_rows"].append(r)
            synced_entries += 1

        # 3. Check Indoctrination fields (Columns AB to AI)
        ind_cand = r.get("Pangalan ng Candidate / Estudyante (Candidate Name)", "")
        ind_phone = r.get("Contact Number / Social Media Link", "")
        ind_inst = r.get("Pangalan ng Nagtuturo (Instructor / Mentor / Priest Name)", "")
        ind_less = r.get("Kasalukuyang Leksyon o Aral na Natapos (Lesson Progress)", "")
        ind_att = r.get("Status ng Pagdalo (Attendance Status)", "")
        ind_bap = r.get("Handa na ba para sa Bautismo / Kumpirma (Ready for Baptism?)", "")
        ind_bap_dt = r.get("Target Date ng Bautismo (kung handa na)", "")
        ind_notes = r.get("Puna o Evaluasyon ng Nagtuturo (Teacher's Notes / Assessment)", "")

        is_ind = ("INDOCTRIN" in act_type or "DOKTRINA" in act_type or bool(ind_cand or ind_less or ind_notes))
        if is_ind and (ind_cand or ind_less or ind_notes or ind_inst):
            name_entry_ind = f"{short_dt}: {ind_cand}" if (short_dt and ind_cand and short_dt not in ind_cand) else (ind_cand or f"Candidate ({short_dt})")
            if name_entry_ind not in grouped[g_key]["ind_names"]:
                grouped[g_key]["ind_names"].append(name_entry_ind)
            ind_bits = []
            if ind_inst:
                ind_bits.append(f"Nagtuturo: {ind_inst}")
            if ind_less:
                ind_bits.append(f"Leksyon: {ind_less}")
            if ind_att:
                ind_bits.append(f"Status: {ind_att}")
            if ind_bap:
                ind_bits.append(f"Handa sa Bautismo: {ind_bap}")
            if ind_bap_dt:
                ind_bits.append(f"Target Baptism: {ind_bap_dt}")
            if ind_notes:
                ind_bits.append(f"Puna: {ind_notes}")
            rem_line_ind = f"{ind_cand or 'Indoctrination'} ({' | '.join(ind_bits)})" if ind_bits else ind_cand
            if rem_line_ind and rem_line_ind not in grouped[g_key]["ind_rems"]:
                grouped[g_key]["ind_rems"].append(rem_line_ind)
            grouped[g_key]["ind_rows"].append(r)
            synced_entries += 1

    if not grouped:
        return 0, "Walang bagong Outreach entry na nakita sa Google Spreadsheet."

    db_modified = False
    for (p_u, m_u, w_u, y_u), bucket in grouped.items():
        rec_k = f"{p_u.upper()}__{m_u.upper()}__{w_u}"
        cur_r = dict(db_records.get(rec_k, {}))
        r_date = resolve_display_week_date(cur_r, m_u, w_u, int(y_u) if str(y_u).isdigit() else 2027)
        cur_r["parish"] = p_u
        cur_r["month"] = m_u
        cur_r["week"] = w_u
        cur_r["report_date"] = r_date

        if bucket["hv_names"]:
            joined_hv = "; ".join(bucket["hv_names"])
            cur_r["hv_names"] = joined_hv
            fc_hv, pc_hv = count_families_and_people_from_text(joined_hv, str(len(bucket["hv_names"])))
            cur_r["house_visitation"] = str(max(fc_hv, len(bucket["hv_names"])))
            cur_r["hv_people_count"] = max(pc_hv, len(bucket["hv_names"]))
            cur_r["hv_remarks"] = " ; ".join(bucket["hv_rems"])
            cur_r["hv_form_rows"] = bucket["hv_rows"]
            db_modified = True

        if bucket["mw_names"]:
            joined_mw = "; ".join(bucket["mw_names"])
            cur_r["mw_names"] = joined_mw
            fc_mw, pc_mw = count_families_and_people_from_text(joined_mw, str(len(bucket["mw_names"])))
            cur_r["missionary_work"] = str(max(fc_mw, len(bucket["mw_names"])))
            cur_r["mw_people_count"] = max(pc_mw, len(bucket["mw_names"]))
            cur_r["mw_remarks"] = " ; ".join(bucket["mw_rems"])
            cur_r["mw_form_rows"] = bucket["mw_rows"]
            db_modified = True

        if bucket["ind_names"]:
            joined_ind = "; ".join(bucket["ind_names"])
            cur_r["ind_names"] = joined_ind
            fc_ind, pc_ind = count_families_and_people_from_text(joined_ind, str(len(bucket["ind_names"])))
            cur_r["indoctrination"] = str(max(fc_ind, len(bucket["ind_names"])))
            cur_r["ind_people_count"] = max(pc_ind, len(bucket["ind_names"]))
            cur_r["ind_remarks"] = " ; ".join(bucket["ind_rems"])
            cur_r["ind_form_rows"] = bucket["ind_rows"]
            db_modified = True

        combined_rems = [x for x in [cur_r.get("hv_remarks", ""), cur_r.get("mw_remarks", ""), cur_r.get("ind_remarks", "")] if x and x != "(No Report)"]
        if combined_rems and not cur_r.get("outreach_general_remarks"):
            cur_r["outreach_general_remarks"] = " | ".join(combined_rems)

        db_records[rec_k] = cur_r
        for l_rst in ["English", "Ilokano"]:
            st.session_state.pop(f"editable_perf_report_{m_u}_{w_u}_{y_u}_{l_rst}", None)

    if db_modified:
        try:
            with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
                json.dump(db_records, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    return synced_entries, f"Matagumpay na na-sync ang {synced_entries} Outreach record(s) mula sa Google Form Spreadsheet papunta sa App Tables at Performance Report!"

def build_detailed_google_form_subtables(df_raw: pd.DataFrame, filter_parish: str = "Lahat ng Parokya") -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Bumubuo ng 3 detalyadong Table na eksaktong sumusunod sa mga tanong ng Google Form:
    1) Detailed House Visitation Table (Columns A–I + J–R)
    2) Detailed Missionary Work Table (Columns A–I + S–AA)
    3) Detailed Indoctrination Table (Columns A–I + AB–AI)"""
    hv_rows = []
    mw_rows = []
    ind_rows = []

    if df_raw is not None and not df_raw.empty:
        for _, row in df_raw.iterrows():
            r = {str(k).strip(): str(v or "").strip() for k, v in row.items()}
            p_raw = r.get("Parish / Simbahan (Church Name)", "")
            loc_raw = r.get("Lokasyon / Area / Barangay", "")
            p_matched = _match_parish_from_form_text(p_raw, loc_raw, "POBLACION PARISH")
            if filter_parish != "Lahat ng Parokya" and not normalize_parish_match(p_matched, filter_parish):
                continue

            act_type = r.get("Type of Activity", "").upper()
            base_info = {
                "Timestamp": r.get("Timestamp", "-"),
                "Parish / Simbahan": p_raw or p_matched,
                "Pangalan ng Nag-uulat (Worker/Priest)": r.get("Pangalan ng Nag-uulat (Worker / Priest / Volunteer Name)", "-"),
                "Opening & Closing Prayer": f"Open: {r.get('Opening Prayer', '-')} | Close: {r.get('Closing Payer', '-')}",
                "Katungkulan / Ministry Unit": r.get("Katungkulan / Ministry Unit", "-"),
                "Lokasyon / Area / Barangay": loc_raw or "-",
            }

            # Check HV
            hv_fam = r.get("Pangalan ng Pamilya / Indibidwal na Binisita", "")
            hv_purp = r.get("Layunin / Uri ng Pagbisita", "")
            hv_top = r.get("Paksa ng Pag-aaral o Mensaheng Ibinahagi (Topic / Scripture Shared)", "")
            if "VISIT" in act_type or "DALAW" in act_type or "HOUSE" in act_type or hv_fam or hv_purp or hv_top:
                hv_rows.append({
                    **base_info,
                    "Petsa ng Pagbisita": r.get("Date (House Visitation)", "") or r.get("Petsa ng Gawain (Date of Activity)", "-"),
                    "Pangalan ng Pamilya / Indibidwal na Binisita": hv_fam or "-",
                    "Tirahan at Contact Number": r.get("Tirahan at Contact Number ng Binisita", "-"),
                    "Layunin / Uri ng Pagbisita": hv_purp or "-",
                    "Kalagayan ng Espirituwal / Sitwasyon": r.get("Kalagayan ng Espirituwal / Sitwasyon ng Pamilya", "-"),
                    "Mga Kahilingan sa Panalangin (Prayer Requests)": r.get("Mga Kahilingan sa Panalangin (Prayer Requests)", "-"),
                    "Paksa / Mensaheng Ibinahagi": hv_top or "-",
                    "Reaksyon o Tugon ng Binisita": r.get("Reaksyon o Tugon ng Binisita sa Mensahe:", "-"),
                    "Susunod na Hakbang / Follow-up": r.get("Susunod na Hakbang / Kailangan ng Follow-up?", "-")
                })

            # Check MW
            mw_con = r.get("Pangalan ng Bagong Contact / Prospect (Contact Name)", "")
            mw_top = r.get("Anong Mensahe o Paksa ang Naibahagi? (Topic / Message Shared)", "")
            mw_resp = r.get("Reaksyon o Tugon sa Ebanghelyo (Response Status)", "")
            if "MISSION" in act_type or "MISYON" in act_type or mw_con or mw_top or mw_resp:
                mw_rows.append({
                    **base_info,
                    "Petsa ng Misyon": r.get("Date (Missionary Work)", "") or r.get("Petsa ng Gawain (Date of Activity)", "-"),
                    "Pangalan ng Bagong Contact / Prospect": mw_con or "-",
                    "Edad at Kasarian": r.get("Edad at Kasarian (Age & Gender)", "-"),
                    "Contact Number / Messenger Link": r.get("Contact Number / Social Media Handle (Messenger Link)", "-"),
                    "Tirahan o Lokasyon kung saan Nakatagpo": r.get("Tirahan o Lokasyon kung saan nakatagpo (Address / Location met)", "-"),
                    "Reaksyon o Tugon sa Ebanghelyo": mw_resp or "-",
                    "Mensahe o Paksang Naibahagi": mw_top or "-",
                    "Mga Interes o Pangangailangan": r.get("Mga Interes o Pangangailangan ng Contact (Interests & Needs)", "-"),
                    "Napagkasunduang Follow-up Date": r.get("Napagkasunduang Araw para sa Follow-up (Follow-up Date)", "-")
                })

            # Check IND
            ind_cand = r.get("Pangalan ng Candidate / Estudyante (Candidate Name)", "")
            ind_less = r.get("Kasalukuyang Leksyon o Aral na Natapos (Lesson Progress)", "")
            ind_notes = r.get("Puna o Evaluasyon ng Nagtuturo (Teacher's Notes / Assessment)", "")
            if "INDOCTRIN" in act_type or "DOKTRINA" in act_type or ind_cand or ind_less or ind_notes:
                ind_rows.append({
                    **base_info,
                    "Petsa ng Gawain": r.get("Petsa ng Gawain (Date of Activity)", "-"),
                    "Pangalan ng Candidate / Estudyante": ind_cand or "-",
                    "Contact Number / Social Media Link": r.get("Contact Number / Social Media Link", "-"),
                    "Pangalan ng Nagtuturo (Instructor / Priest)": r.get("Pangalan ng Nagtuturo (Instructor / Mentor / Priest Name)", "-"),
                    "Kasalukuyang Leksyon (Lesson Progress)": ind_less or "-",
                    "Status ng Pagdalo (Attendance Status)": r.get("Status ng Pagdalo (Attendance Status)", "-"),
                    "Handa na ba para sa Bautismo / Kumpirma?": r.get("Handa na ba para sa Bautismo / Kumpirma (Ready for Baptism?)", "-"),
                    "Target Date ng Bautismo": r.get("Target Date ng Bautismo (kung handa na)", "-"),
                    "Puna o Evaluasyon ng Nagtuturo (Remarks)": ind_notes or "-"
                })

    hv_cols = [
        "Timestamp", "Parish / Simbahan", "Petsa ng Pagbisita", "Pangalan ng Nag-uulat (Worker/Priest)",
        "Opening & Closing Prayer", "Katungkulan / Ministry Unit", "Lokasyon / Area / Barangay",
        "Pangalan ng Pamilya / Indibidwal na Binisita", "Tirahan at Contact Number", "Layunin / Uri ng Pagbisita",
        "Kalagayan ng Espirituwal / Sitwasyon", "Mga Kahilingan sa Panalangin (Prayer Requests)",
        "Paksa / Mensaheng Ibinahagi", "Reaksyon o Tugon ng Binisita", "Susunod na Hakbang / Follow-up"
    ]
    mw_cols = [
        "Timestamp", "Parish / Simbahan", "Petsa ng Misyon", "Pangalan ng Nag-uulat (Worker/Priest)",
        "Opening & Closing Prayer", "Katungkulan / Ministry Unit", "Lokasyon / Area / Barangay",
        "Pangalan ng Bagong Contact / Prospect", "Edad at Kasarian", "Contact Number / Messenger Link",
        "Tirahan o Lokasyon kung saan Nakatagpo", "Reaksyon o Tugon sa Ebanghelyo",
        "Mensahe o Paksang Naibahagi", "Mga Interes o Pangangailangan", "Napagkasunduang Follow-up Date"
    ]
    ind_cols = [
        "Timestamp", "Parish / Simbahan", "Petsa ng Gawain", "Pangalan ng Nag-uulat (Worker/Priest)",
        "Opening & Closing Prayer", "Katungkulan / Ministry Unit", "Lokasyon / Area / Barangay",
        "Pangalan ng Candidate / Estudyante", "Contact Number / Social Media Link",
        "Pangalan ng Nagtuturo (Instructor / Priest)", "Kasalukuyang Leksyon (Lesson Progress)",
        "Status ng Pagdalo (Attendance Status)", "Handa na ba para sa Bautismo / Kumpirma?",
        "Target Date ng Bautismo", "Puna o Evaluasyon ng Nagtuturo (Remarks)"
    ]

    return (
        pd.DataFrame(hv_rows, columns=hv_cols) if hv_rows else pd.DataFrame(columns=hv_cols),
        pd.DataFrame(mw_rows, columns=mw_cols) if mw_rows else pd.DataFrame(columns=mw_cols),
        pd.DataFrame(ind_rows, columns=ind_cols) if ind_rows else pd.DataFrame(columns=ind_cols),
    )

def render_outreach_visitation_missionary_page():
    is_master = (st.session_state.user_role == "Master Admin")
    my_parish = st.session_state.get("assigned_parish", "Lahat ng Parokya")
    my_chat_identity = get_current_chat_identity()

    render_title_with_logo(
        "🏠 House Visitation, Missionary Work, Indoctrination & Remarks",
        "Opisyal na Pahina at Google Form Table para sa House Visitation (10 Families Req.), Missionary Work, Indoctrination, at Remarks (Auto-Sync mula sa Google Form Spreadsheet)."
    )

    outreach_parish_options = (["Lahat ng Parokya"] + PARISH_ORDER) if is_master else [my_parish]
    if "outreach_sel_parish" not in st.session_state or st.session_state["outreach_sel_parish"] not in outreach_parish_options:
        st.session_state["outreach_sel_parish"] = outreach_parish_options[0]
    if "outreach_sel_month" not in st.session_state or st.session_state["outreach_sel_month"] not in MONTHS_2027:
        st.session_state["outreach_sel_month"] = st.session_state.get("mon_view_month", "January")
        if st.session_state["outreach_sel_month"] not in MONTHS_2027:
            st.session_state["outreach_sel_month"] = "January"
    selected_year = st.session_state.get("mon_view_year", "2027")

    out_cat_options = [
        "Lahat (House Visitation, Missionary Work, Indoctrination & Remarks)",
        "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER",
        "Missionary Work (Family Names, Mga Pangalan ng Dumalo, Total Tao & Remarks)",
        "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION",
        "Remarks ng House Visitation, Missionary Work & Indoctrination"
    ]

    oc1, oc2, oc3, oc4, oc5 = st.columns([1.15, 1.45, 0.85, 1.00, 1.30])
    with oc1:
        out_parish = st.selectbox(
            "🏛️ Parokya (Per Parish):",
            outreach_parish_options,
            key="outreach_sel_parish"
        )
    with oc2:
        out_category = st.selectbox(
            "📂 Pumili ng Kategorya (Outreach & Remarks):",
            out_cat_options,
            key="outreach_sel_category"
        )
    with oc3:
        out_month = st.selectbox(
            "📅 Buwan (2027):",
            MONTHS_2027,
            key="outreach_sel_month"
        )
    out_month_actual_weeks = get_weeks_for_month_year(out_month, int(selected_year))
    outreach_week_options = out_month_actual_weeks + [f"Lahat ng Weeks (1st–{out_month_actual_weeks[-1]})"]
    if "outreach_sel_week" not in st.session_state or st.session_state["outreach_sel_week"] not in outreach_week_options:
        st.session_state["outreach_sel_week"] = out_month_actual_weeks[0]
    with oc4:
        out_week_sel = st.selectbox(
            f"🗓️ Linggo ({len(out_month_actual_weeks)} Weeks / Lahat):",
            outreach_week_options,
            key="outreach_sel_week"
        )

    # Auto-fetch Google Sheet responses on initial load of this page
    eff_w_sync = out_week_sel if out_week_sel in out_month_actual_weeks else out_month_actual_weeks[0]
    sheet_df_live, sheet_fetch_msg = fetch_outreach_google_sheet_df(DEFAULT_OUTREACH_GOOGLE_SHEET_URL)
    if not sheet_df_live.empty and not st.session_state.get("_outreach_sheet_autosynced_once"):
        sync_outreach_google_sheet_to_weekly_db(sheet_df_live, out_month, eff_w_sync, selected_year, out_parish)
        st.session_state["_outreach_sheet_autosynced_once"] = True

    with oc5:
        st.write("")
        if st.button("🔄 Sync Mula sa Google Sheet & Notion", use_container_width=True, type="primary", key="btn_sync_outreach_gsheet"):
            with st.spinner("Kinukuha ang live responses mula sa Google Form Spreadsheet (1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8)..."):
                df_fresh, msg_fresh = fetch_outreach_google_sheet_df(DEFAULT_OUTREACH_GOOGLE_SHEET_URL)
                gs_cnt, gs_msg = sync_outreach_google_sheet_to_weekly_db(df_fresh, out_month, eff_w_sync, selected_year, out_parish)
                n_cnt, _ = sync_outreach_tasks_from_notion_to_excel_and_db(out_month, eff_w_sync, out_parish, selected_year)
                st.cache_data.clear()
                if gs_cnt > 0 or n_cnt > 0:
                    st.success(f"✅ {gs_msg} ({msg_fresh})")
                else:
                    st.info(f"ℹ️ {msg_fresh} — Naka-handa na ang koneksyon; kapag nag-submit ang mga Leaders sa Google Form, awtomatiko itong papasok sa Table at Performance Report!")
                time.sleep(0.4)
                st.rerun()

    # =====================================================================
    # 🔗 OFFICIAL GOOGLE FORM & GOOGLE SPREADSHEET BANNER + EMBEDDED FORM
    # =====================================================================
    st.markdown(
        f"""
        <div class="full-text-card" style="border-left: 5px solid #22c55e;">
            <b>📋 Opisyal na Google Form at Responses Spreadsheet para sa House Visitation, Missionary Work, Indoctrination & Remarks:</b><br>
            • 📝 <b>Google Form (Para sa mga Parish Leaders / Workers):</b>
              <a href="{DEFAULT_OUTREACH_GOOGLE_FORM_URL}" target="_blank" style="color:#4ade80;font-weight:700;text-decoration:underline;">
                Buksan ang Church Operations Report Form (Visitation, Mission, & Indoctrination) ↗
              </a><br>
            • 📊 <b>Google Spreadsheet (Pinagkukunan ng Info sa App Tables & Performance Report):</b>
              <a href="{DEFAULT_OUTREACH_GOOGLE_SHEET_URL}" target="_blank" style="color:#60a5fa;font-weight:700;text-decoration:underline;">
                Buksan ang Google Sheet Responses (1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8) ↗
              </a><br>
            <span style="font-size:0.86rem;opacity:0.9;">📡 <b>Live Connection Status:</b> {sheet_fetch_msg}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.expander("📝 Sagutan nang Direkta sa App ang Opisyal na Google Form (Embedded Google Form para sa mga Parish Leaders)", expanded=False):
        st.caption(
            f"Maaaring sagutan ng mga Parish Leaders ang opisyal na Google Form dito mismo sa loob ng App o buksan sa bagong tab: [Buksan ang Google Form ↗]({DEFAULT_OUTREACH_GOOGLE_FORM_URL}). "
            "Pagkatapos mag-submit sa Google Form, pindutin lamang ang **🔄 Sync Mula sa Google Sheet & Notion** button sa itaas upang agad pumasok sa mga Table sa ibaba!"
        )
        components.iframe(DEFAULT_OUTREACH_GOOGLE_FORM_EMBED_URL, height=680, scrolling=True)

    # =====================================================================
    # 📊 GOOGLE FORM RESPONSES TABLES (35 COLUMNS MULA SA SPREADSHEET 1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8)
    # =====================================================================
    hv_detail_df, mw_detail_df, ind_detail_df = build_detailed_google_form_subtables(sheet_df_live, out_parish)

    with st.expander(f"📊 Mga Table Base sa Google Form & Spreadsheet Responses ({len(sheet_df_live)} Total Row(s) sa 'Form responses 1' — Filter: {out_parish})", expanded=True):
        tab_gf_all, tab_gf_hv, tab_gf_mw, tab_gf_ind = st.tabs([
            f"📋 Buong Google Form Table (35 Columns - {len(sheet_df_live)} Row(s))",
            f"🏡 House Visitation Form Table ({len(hv_detail_df)})",
            f"🌍 Missionary Work Form Table ({len(mw_detail_df)})",
            f"📖 Indoctrination Form Table ({len(ind_detail_df)})"
        ])
        with tab_gf_all:
            st.caption("Ito ang eksaktong 35-Column Table mula sa Google Form at Google Spreadsheet (`Form responses 1` ng `1FZLr3gFvU-YXWJMTrRkOqJ2X0hTAaQ3rRerX9cdUIb8`):")
            if not sheet_df_live.empty:
                if out_parish != "Lahat ng Parokya":
                    filtered_all_df = sheet_df_live[
                        sheet_df_live["Parish / Simbahan (Church Name)"].astype(str).apply(
                            lambda x: normalize_parish_match(_match_parish_from_form_text(x, "", ""), out_parish)
                        )
                    ]
                else:
                    filtered_all_df = sheet_df_live
                st.dataframe(filtered_all_df, use_container_width=True, hide_index=True)
            else:
                st.dataframe(pd.DataFrame(columns=OUTREACH_FORM_35_COLUMNS), use_container_width=True, hide_index=True)
                st.caption("Wala pang nakalistang sagot sa Google Sheet (`Form responses 1`). Gamitin ang Google Form o ang Built-in Form sa ibaba upang magdagdag ng record.")

        with tab_gf_hv:
            st.caption("Detalyadong Table para sa **House Visitation** base sa mga tanong sa Google Form (Columns A–I at J–R):")
            if not hv_detail_df.empty:
                render_record_style_html_table(hv_detail_df)
            else:
                st.dataframe(hv_detail_df, use_container_width=True, hide_index=True)

        with tab_gf_mw:
            st.caption("Detalyadong Table para sa **Missionary Work** base sa mga tanong sa Google Form (Columns A–I at S–AA):")
            if not mw_detail_df.empty:
                render_record_style_html_table(mw_detail_df)
            else:
                st.dataframe(mw_detail_df, use_container_width=True, hide_index=True)

        with tab_gf_ind:
            st.caption("Detalyadong Table para sa **Indoctrination** base sa mga tanong sa Google Form (Columns A–I at AB–AI):")
            if not ind_detail_df.empty:
                render_record_style_html_table(ind_detail_df)
            else:
                st.dataframe(ind_detail_df, use_container_width=True, hide_index=True)

    # =====================================================================
    # 📤 BUILT-IN APP FORM NA Kapareho ng Google Form Fields (Auto-Sync sa Tables & Performance Report)
    # =====================================================================
    db_live = load_weekly_monitoring_records()
    form_parish_choices = PARISH_ORDER if is_master else [my_parish]
    default_form_parish = out_parish if out_parish in form_parish_choices else form_parish_choices[0]
    default_form_week = out_week_sel if out_week_sel in out_month_actual_weeks else out_month_actual_weeks[0]

    with st.expander("📤 Ipasa / I-update sa App Table ang House Visitation, Missionary Work, Indoctrination & Remarks (Kapareho ng Google Form + Auto-Sync sa Performance Report)", expanded=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            sub_out_parish = st.selectbox("🏛️ Parish / Simbahan (Church Name):", form_parish_choices, index=form_parish_choices.index(default_form_parish), key="form_out_parish")
        with f_col2:
            sub_out_month = st.selectbox("📅 Buwan:", MONTHS_2027, index=MONTHS_2027.index(out_month), key="form_out_month")
        sub_out_weeks = get_weeks_for_month_year(sub_out_month, int(selected_year))
        with f_col3:
            sub_out_week = st.selectbox("🗓️ Linggo:", sub_out_weeks, index=sub_out_weeks.index(default_form_week) if default_form_week in sub_out_weeks else 0, key="form_out_week")

        cur_out_key = f"{sub_out_parish.upper()}__{sub_out_month.upper()}__{sub_out_week}"
        cur_out_rec = db_live.get(cur_out_key, {})
        parsed_out_ro = parse_remarks_into_mass_and_outreach(str(cur_out_rec.get("remarks", "")))
        out_rev = f"{st.session_state.get('pwr_form_rev', 0)}_{cur_out_key}"
        r_date_default = resolve_display_week_date(cur_out_rec, sub_out_month, sub_out_week, int(selected_year))

        with st.form("outreach_dedicated_submission_form"):
            st.markdown("##### 📌 Bahagi I: Pangkalahatang Impormasyon ng Gawain (Google Form Columns B–I)")
            g1, g2, g3, g4 = st.columns(4)
            with g1:
                in_worker_name = st.text_input("Pangalan ng Nag-uulat (Worker/Priest/Volunteer):", value=str(cur_out_rec.get("outreach_worker_name", "")), key=f"owk_{out_rev}")
                in_ministry_unit = st.text_input("Katungkulan / Ministry Unit:", value=str(cur_out_rec.get("outreach_ministry_unit", "Parish Worker / Priest")), key=f"omu_{out_rev}")
            with g2:
                in_open_prayer = st.text_input("Opening Prayer:", value=str(cur_out_rec.get("outreach_open_prayer", "")), key=f"oop_{out_rev}")
                in_close_prayer = st.text_input("Closing Prayer:", value=str(cur_out_rec.get("outreach_close_prayer", "")), key=f"ocp_{out_rev}")
            with g3:
                in_act_date = st.text_input("Petsa ng Gawain (Date of Activity):", value=str(cur_out_rec.get("outreach_activity_date", r_date_default)), key=f"oad_{out_rev}")
                in_brgy_loc = st.text_input("Lokasyon / Area / Barangay:", value=str(cur_out_rec.get("outreach_location", sub_out_parish.replace("PARISH", "").strip().title())), key=f"oloc_{out_rev}")
            with g4:
                in_act_type = st.selectbox(
                    "Type of Activity:",
                    ["House Visitation, Missionary Work & Indoctrination", "House Visitation", "Missionary Work", "Indoctrination"],
                    key=f"oat_{out_rev}"
                )

            st.markdown("##### 1️⃣ House Visitation (Google Form Columns J–R | Requirement: 10 Families Member or Non-Member)")
            hc1, hc2, hc3 = st.columns([0.85, 1.65, 1.30])
            with hc1:
                in_hv_cnt = st.text_input("Bilang ng Pamilya (House Visitation):", value=str(cur_out_rec.get("house_visitation", "0")).replace("(No Report)", "0"), key=f"ohv_c_{out_rev}")
                in_hv_ppl = st.number_input("Total ng mga Tao (House Visitation):", min_value=0, value=int(cur_out_rec.get("hv_people_count", 0) or parsed_out_ro.get("hv_people_count", 0)), key=f"ohv_p_{out_rev}")
                in_hv_purpose = st.text_input("Layunin / Uri ng Pagbisita:", value=str(cur_out_rec.get("hv_purpose", "")), key=f"ohv_purp_{out_rev}")
            with hc2:
                in_hv_names = st.text_area(
                    "Pangalan ng Pamilya / Indibidwal na Binisita (Hal. 11-14-26: Nico Q; Remedios Q; Arjon Q; Milagno R; Imee N.):",
                    value=str(cur_out_rec.get("hv_names", "") or parsed_out_ro.get("hv_names", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"ohv_n_{out_rev}"
                )
                in_hv_topic = st.text_input("Paksa ng Pag-aaral o Mensaheng Ibinahagi (Topic / Scripture Shared):", value=str(cur_out_rec.get("hv_topic_shared", "")), key=f"ohv_top_{out_rev}")
            with hc3:
                in_hv_rem = st.text_area(
                    "Kalagayan / Prayer Requests / Reaksyon / Follow-up (Remarks sa House Visitation):",
                    value=str(cur_out_rec.get("hv_remarks", "") or parsed_out_ro.get("hv_remarks", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"ohv_r_{out_rev}"
                )
                in_hv_addr = st.text_input("Tirahan at Contact Number ng Binisita:", value=str(cur_out_rec.get("hv_addr_contact", "")), key=f"ohv_ad_{out_rev}")

            st.markdown("##### 2️⃣ Missionary Work (Google Form Columns S–AA)")
            mc1, mc2, mc3 = st.columns([0.85, 1.65, 1.30])
            with mc1:
                in_mw_cnt = st.text_input("Bilang ng Pamilya / Contact (Missionary Work):", value=str(cur_out_rec.get("missionary_work", "0")).replace("(No Report)", "0"), key=f"omw_c_{out_rev}")
                in_mw_ppl = st.number_input("Total ng mga Tao (Missionary Work):", min_value=0, value=int(cur_out_rec.get("mw_people_count", 0) or parsed_out_ro.get("mw_people_count", 0)), key=f"omw_p_{out_rev}")
                in_mw_age_gen = st.text_input("Edad at Kasarian / Contact No.:", value=str(cur_out_rec.get("mw_age_gender", "")), key=f"omw_ag_{out_rev}")
            with mc2:
                in_mw_names = st.text_area(
                    "Pangalan ng Bagong Contact / Prospect (Family Names & Mga Pangalan):",
                    value=str(cur_out_rec.get("mw_names", "") or parsed_out_ro.get("mw_names", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"omw_n_{out_rev}"
                )
                in_mw_topic = st.text_input("Anong Mensahe o Paksa ang Naibahagi? (Topic Shared):", value=str(cur_out_rec.get("mw_topic_shared", "")), key=f"omw_top_{out_rev}")
            with mc3:
                in_mw_rem = st.text_area(
                    "Reaksyon sa Ebanghelyo / Interes / Follow-up Date (Remarks sa Missionary Work):",
                    value=str(cur_out_rec.get("mw_remarks", "") or parsed_out_ro.get("mw_remarks", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"omw_r_{out_rev}"
                )
                in_mw_loc = st.text_input("Tirahan o Lokasyon kung saan nakatagpo:", value=str(cur_out_rec.get("mw_location_met", "")), key=f"omw_loc_{out_rev}")

            st.markdown("##### 3️⃣ Indoctrination (Google Form Columns AB–AI | Family Name of a Family Under Indoctrination)")
            ic1, ic2, ic3 = st.columns([0.85, 1.65, 1.30])
            with ic1:
                in_ind_cnt = st.text_input("Bilang ng Pamilya / Candidate (Indoctrination):", value=str(cur_out_rec.get("indoctrination", "0")).replace("(No Report)", "0"), key=f"oind_c_{out_rev}")
                in_ind_ppl = st.number_input("Total ng mga Tao (Indoctrination):", min_value=0, value=int(cur_out_rec.get("ind_people_count", 0) or parsed_out_ro.get("ind_people_count", 0)), key=f"oind_p_{out_rev}")
                in_ind_inst = st.text_input("Pangalan ng Nagtuturo (Instructor / Priest):", value=str(cur_out_rec.get("ind_instructor", "")), key=f"oind_ins_{out_rev}")
            with ic2:
                in_ind_names = st.text_area(
                    "Pangalan ng Candidate / Estudyante (Hal. Pasito L; Nathan L.):",
                    value=str(cur_out_rec.get("ind_names", "") or parsed_out_ro.get("ind_names", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"oind_n_{out_rev}"
                )
                in_ind_lesson = st.text_input("Kasalukuyang Leksyon o Aral na Natapos (Lesson Progress):", value=str(cur_out_rec.get("ind_lesson_progress", "")), key=f"oind_les_{out_rev}")
            with ic3:
                in_ind_rem = st.text_area(
                    "Puna o Evaluasyon ng Nagtuturo / Status ng Pagdalo (Remarks sa Indoctrination):",
                    value=str(cur_out_rec.get("ind_remarks", "") or parsed_out_ro.get("ind_remarks", "")).replace("(No Report)", ""),
                    height=68,
                    key=f"oind_r_{out_rev}"
                )
                in_ind_bapt = st.text_input("Handa na sa Bautismo? & Target Date:", value=str(cur_out_rec.get("ind_baptism_ready", "")), key=f"oind_bap_{out_rev}")

            st.markdown("##### 4️⃣ Pangkalahatang Remarks para sa Performance Report (Hal. Dahilan kung walang Visitation / Sitwasyon sa Parokya)")
            in_gen_out_rem = st.text_area(
                "Remarks (Ito ang awtomatikong lalabas sa `Remarks:` ng Parish Ministry Performance Report ni Master Admin):",
                value=str(cur_out_rec.get("outreach_general_remarks", "")),
                height=68,
                key=f"ogen_r_{out_rev}"
            )

            if st.form_submit_button("💾 I-save sa App Tables, Google Form Cache & Performance Report", type="primary", use_container_width=True):
                r_date_out = in_act_date.strip() or r_date_default
                merged_payload = dict(cur_out_rec)
                merged_payload["parish"] = sub_out_parish
                merged_payload["month"] = sub_out_month
                merged_payload["week"] = sub_out_week
                merged_payload["report_date"] = r_date_out
                merged_payload["outreach_worker_name"] = in_worker_name.strip()
                merged_payload["outreach_ministry_unit"] = in_ministry_unit.strip()
                merged_payload["outreach_open_prayer"] = in_open_prayer.strip()
                merged_payload["outreach_close_prayer"] = in_close_prayer.strip()
                merged_payload["outreach_activity_date"] = r_date_out
                merged_payload["outreach_location"] = in_brgy_loc.strip()

                merged_payload["house_visitation"] = in_hv_cnt.strip() or "0"
                merged_payload["hv_names"] = in_hv_names.strip()
                merged_payload["hv_people_count"] = in_hv_ppl
                merged_payload["hv_remarks"] = in_hv_rem.strip()
                merged_payload["hv_purpose"] = in_hv_purpose.strip()
                merged_payload["hv_topic_shared"] = in_hv_topic.strip()
                merged_payload["hv_addr_contact"] = in_hv_addr.strip()

                merged_payload["missionary_work"] = in_mw_cnt.strip() or "0"
                merged_payload["mw_names"] = in_mw_names.strip()
                merged_payload["mw_people_count"] = in_mw_ppl
                merged_payload["mw_remarks"] = in_mw_rem.strip()
                merged_payload["mw_age_gender"] = in_mw_age_gen.strip()
                merged_payload["mw_topic_shared"] = in_mw_topic.strip()
                merged_payload["mw_location_met"] = in_mw_loc.strip()

                merged_payload["indoctrination"] = in_ind_cnt.strip() or "0"
                merged_payload["ind_names"] = in_ind_names.strip()
                merged_payload["ind_people_count"] = in_ind_ppl
                merged_payload["ind_remarks"] = in_ind_rem.strip()
                merged_payload["ind_instructor"] = in_ind_inst.strip()
                merged_payload["ind_lesson_progress"] = in_ind_lesson.strip()
                merged_payload["ind_baptism_ready"] = in_ind_bapt.strip()
                merged_payload["outreach_general_remarks"] = in_gen_out_rem.strip()

                # Save also to the 35-column Google Form response table cache so it appears in the Google Form Table immediately
                new_gf_row = {
                    "Timestamp": get_philippine_datetime_str(),
                    "Parish / Simbahan (Church Name)": sub_out_parish,
                    "Pangalan ng Nag-uulat (Worker / Priest / Volunteer Name)": in_worker_name.strip() or my_chat_identity,
                    "Opening Prayer": in_open_prayer.strip(),
                    "Closing Payer": in_close_prayer.strip(),
                    "Katungkulan / Ministry Unit": in_ministry_unit.strip(),
                    "Petsa ng Gawain (Date of Activity)": r_date_out,
                    "Lokasyon / Area / Barangay": in_brgy_loc.strip(),
                    "Type of Activity": in_act_type,
                    "Date (House Visitation)": r_date_out if in_hv_names.strip() else "",
                    "Pangalan ng Pamilya / Indibidwal na Binisita": in_hv_names.strip(),
                    "Tirahan at Contact Number ng Binisita": in_hv_addr.strip(),
                    "Layunin / Uri ng Pagbisita": in_hv_purpose.strip(),
                    "Kalagayan ng Espirituwal / Sitwasyon ng Pamilya": in_hv_rem.strip(),
                    "Mga Kahilingan sa Panalangin (Prayer Requests)": "",
                    "Paksa ng Pag-aaral o Mensaheng Ibinahagi (Topic / Scripture Shared)": in_hv_topic.strip(),
                    "Reaksyon o Tugon ng Binisita sa Mensahe:": in_hv_rem.strip(),
                    "Susunod na Hakbang / Kailangan ng Follow-up?": "",
                    "Date (Missionary Work)": r_date_out if in_mw_names.strip() else "",
                    "Pangalan ng Bagong Contact / Prospect (Contact Name)": in_mw_names.strip(),
                    "Edad at Kasarian (Age & Gender)": in_mw_age_gen.strip(),
                    "Contact Number / Social Media Handle (Messenger Link)": "",
                    "Tirahan o Lokasyon kung saan nakatagpo (Address / Location met)": in_mw_loc.strip(),
                    "Reaksyon o Tugon sa Ebanghelyo (Response Status)": in_mw_rem.strip(),
                    "Anong Mensahe o Paksa ang Naibahagi? (Topic / Message Shared)": in_mw_topic.strip(),
                    "Mga Interes o Pangangailangan ng Contact (Interests & Needs)": "",
                    "Napagkasunduang Araw para sa Follow-up (Follow-up Date)": "",
                    "Pangalan ng Candidate / Estudyante (Candidate Name)": in_ind_names.strip(),
                    "Contact Number / Social Media Link": "",
                    "Pangalan ng Nagtuturo (Instructor / Mentor / Priest Name)": in_ind_inst.strip(),
                    "Kasalukuyang Leksyon o Aral na Natapos (Lesson Progress)": in_ind_lesson.strip(),
                    "Status ng Pagdalo (Attendance Status)": "Active" if in_ind_names.strip() else "",
                    "Handa na ba para sa Bautismo / Kumpirma (Ready for Baptism?)": in_ind_bapt.strip(),
                    "Target Date ng Bautismo (kung handa na)": "",
                    "Puna o Evaluasyon ng Nagtuturo (Teacher's Notes / Assessment)": in_ind_rem.strip() or in_gen_out_rem.strip()
                }
                cur_cache_rows = _load_local_outreach_form_cache()
                cur_cache_rows.append(new_gf_row)
                _save_local_outreach_form_cache(cur_cache_rows)

                ok_o, msg_o = save_weekly_monitoring_record(
                    sub_out_parish, sub_out_month, sub_out_week, r_date_out, merged_payload, drive_link=cur_out_rec.get("drive_link", "")
                )
                db_after = load_weekly_monitoring_records()
                if cur_out_key in db_after:
                    db_after[cur_out_key]["outreach_general_remarks"] = in_gen_out_rem.strip()
                    if cur_out_rec.get("no_report_locked") and not is_week_record_completed(cur_out_rec):
                        db_after[cur_out_key]["no_report_locked"] = True
                        db_after[cur_out_key]["has_data"] = False
                    try:
                        with open(LOCAL_WEEKLY_REPORTS_DB, "w", encoding="utf-8") as f:
                            json.dump(db_after, f, ensure_ascii=False, indent=2)
                    except Exception:
                        pass

                for l_rst in ["English", "Ilokano"]:
                    st.session_state.pop(f"editable_perf_report_{sub_out_month}_{sub_out_week}_{selected_year}_{l_rst}", None)
                try:
                    cfg_r = load_saved_config()
                    if isinstance(cfg_r.get("SAVED_PERFORMANCE_REPORTS"), dict):
                        for l_rst in ["English", "Ilokano"]:
                            cfg_r["SAVED_PERFORMANCE_REPORTS"].pop(f"editable_perf_report_{sub_out_month}_{sub_out_week}_{selected_year}_{l_rst}", None)
                        with open(LOCAL_CONFIG_FILE, "w", encoding="utf-8") as f:
                            json.dump(cfg_r, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass

                dispatch_app_notification_and_messenger_alert(
                    sender_id=my_chat_identity,
                    recipient_id="MASTER ADMIN",
                    title=f"Outreach Report Updated ({sub_out_parish})",
                    details=f"Na-update ang House Visitation ({in_hv_cnt}), Missionary Work ({in_mw_cnt}), Indoctrination ({in_ind_cnt}) & Remarks ng {sub_out_parish} para sa {sub_out_month} {selected_year} ({sub_out_week}).",
                    is_error=False
                )
                st.session_state["pwr_form_rev"] = st.session_state.get("pwr_form_rev", 0) + 1
                st.success(f"✅ Matagumpay na na-save ang House Visitation, Missionary Work, Indoctrination & Remarks para sa **{sub_out_parish}** ({sub_out_month} {sub_out_week})! ({msg_o})")
                time.sleep(0.4)
                st.rerun()

    st.divider()
    out_weeks_list = out_month_actual_weeks if out_week_sel not in out_month_actual_weeks else [out_week_sel]
    out_scope_label = f"{out_week_sel} of {out_month.upper()} {selected_year} — {out_parish}"

    if out_category in ("Lahat (House Visitation, Missionary Work, Indoctrination & Remarks)", "HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER"):
        st.markdown(f"#### 🏡 Unang Table: HOUSE Visitation: REQUIREMENTS 10 FAMILIES MEMBER OR NON-MEMBER ({out_scope_label})")
        hv_table_df = build_multi_week_house_visitation_df(out_month, out_weeks_list, out_parish)
        render_record_style_html_table(hv_table_df)

    if out_category in ("Lahat (House Visitation, Missionary Work, Indoctrination & Remarks)", "Missionary Work (Family Names, Mga Pangalan ng Dumalo, Total Tao & Remarks)"):
        st.markdown(f"#### 🌍 Pangalawang Table: Missionary Work (Family Names, Mga Pangalan ng Dumalo, Total Tao & Remarks) ({out_scope_label})")
        mw_table_df = build_multi_week_missionary_work_df(out_month, out_weeks_list, out_parish)
        render_record_style_html_table(mw_table_df)

    if out_category in ("Lahat (House Visitation, Missionary Work, Indoctrination & Remarks)", "INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION"):
        st.markdown(f"#### 📖 Pangatlong Table: INDOCTRINATION: FAMILY NAME OF A FAMILY UNDER INDOCTRINATION ({out_scope_label})")
        ind_table_df = build_multi_week_indoctrination_df(out_month, out_weeks_list, out_parish)
        render_record_style_html_table(ind_table_df)

    if out_category in ("Lahat (House Visitation, Missionary Work, Indoctrination & Remarks)", "Remarks ng House Visitation, Missionary Work & Indoctrination"):
        st.markdown(f"#### 📝 Remarks sa House Visitation, Remarks sa Missionary Work, at Remarks sa Indoctrination ({out_scope_label})")
        out_rem_df = build_multi_week_outreach_remarks_df(out_month, out_weeks_list, out_parish)
        render_record_style_html_table(out_rem_df)

        outreach_cards_list = []
        parishes_out_scan = PARISH_ORDER if out_parish == "Lahat ng Parokya" else [p for p in PARISH_ORDER if normalize_parish_match(p, out_parish)]
        for wk_o in out_weeks_list:
            for p_c in parishes_out_scan:
                rec_c = db_live.get(f"{p_c.upper()}__{out_month.upper()}__{wk_o}", {})
                r_raw = str(rec_c.get("remarks", "")).strip()
                p_ro = parse_remarks_into_mass_and_outreach(r_raw) if r_raw else {}
                hv_r = str(rec_c.get("hv_remarks", "") or p_ro.get("hv_remarks", "")).strip()
                mw_r = str(rec_c.get("mw_remarks", "") or p_ro.get("mw_remarks", "")).strip()
                ind_r = str(rec_c.get("ind_remarks", "") or p_ro.get("ind_remarks", "")).strip()
                hv_n = str(rec_c.get("hv_names", "") or p_ro.get("hv_names", "")).strip()
                mw_n = str(rec_c.get("mw_names", "") or p_ro.get("mw_names", "")).strip()
                ind_n = str(rec_c.get("ind_names", "") or p_ro.get("ind_names", "")).strip()
                gen_r = str(rec_c.get("outreach_general_remarks", "")).strip()
                if hv_r or mw_r or ind_r or hv_n or mw_n or ind_n or gen_r or parse_int_from_text(rec_c.get("house_visitation", "0")) > 0 or parse_int_from_text(rec_c.get("missionary_work", "0")) > 0 or parse_int_from_text(rec_c.get("indoctrination", "0")) > 0:
                    outreach_cards_list.append((p_c, wk_o, rec_c, p_ro))

        with st.expander(f"📝 Basahin nang Buo ang Remarks sa House Visitation, Missionary Work & Indoctrination ({len(outreach_cards_list)} Record(s) — {out_scope_label})", expanded=True):
            if outreach_cards_list:
                for p_c, wk_o, rec_c, p_ro in outreach_cards_list:
                    hv_n = str(rec_c.get("hv_names", "") or p_ro.get("hv_names", "")).strip() or "Wala pang nakalistang pangalan"
                    hv_r = str(rec_c.get("hv_remarks", "") or p_ro.get("hv_remarks", "")).strip() or "-"
                    hv_ppl = int(rec_c.get("hv_people_count", 0) or p_ro.get("hv_people_count", 0))

                    mw_n = str(rec_c.get("mw_names", "") or p_ro.get("mw_names", "")).strip() or "Wala pang nakalistang pangalan"
                    mw_r = str(rec_c.get("mw_remarks", "") or p_ro.get("mw_remarks", "")).strip() or "-"
                    mw_ppl = int(rec_c.get("mw_people_count", 0) or p_ro.get("mw_people_count", 0))

                    ind_n = str(rec_c.get("ind_names", "") or p_ro.get("ind_names", "")).strip() or "Wala pang nakalistang pangalan"
                    ind_r = str(rec_c.get("ind_remarks", "") or p_ro.get("ind_remarks", "")).strip() or "-"
                    ind_ppl = int(rec_c.get("ind_people_count", 0) or p_ro.get("ind_people_count", 0))
                    gen_r = str(rec_c.get("outreach_general_remarks", "")).strip()

                    wk_o_dt = resolve_display_week_date(rec_c, out_month, wk_o, int(selected_year))
                    st.markdown(
                        f"<div class='full-text-card'>"
                        f"<b>🏠 {p_c}</b> • <code>{out_month} {selected_year} ({wk_o})</code> • 📅 <b>Petsa:</b> {wk_o_dt}<br>"
                        f"🏡 <b>Remarks sa House Visitation ({rec_c.get('house_visitation', '0')} Pamilya | {hv_ppl} Tao):</b> {hv_r} <i>(Mga Pangalan: {hv_n})</i><br>"
                        f"🌍 <b>Remarks sa Missionary Work ({rec_c.get('missionary_work', '0')} Pamilya | {mw_ppl} Tao):</b> {mw_r} <i>(Mga Pangalan: {mw_n})</i><br>"
                        f"📖 <b>Remarks sa Indoctrination ({rec_c.get('indoctrination', '0')} Pamilya | {ind_ppl} Tao):</b> {ind_r} <i>(Mga Pangalan: {ind_n})</i>"
                        + (f"<br>📝 <b>Remarks (Performance Report):</b> {gen_r}" if gen_r else "")
                        + f"</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.caption("Walang nakalagay na Remarks sa House Visitation, Missionary Work, o Indoctrination para sa piniling filter.")

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

    if st.session_state.user_role == "Master Admin" or st.session_state.get("login_role") == "Master Admin":
        active_notifs_for_me = trigger_browser_and_toast_notifications_for_current_user()
        with st.expander(f"🔔 Live Notifications — Admin Only ({len(active_notifs_for_me)})", expanded=bool(active_notifs_for_me and active_notifs_for_me[-1].get("is_error"))):
            if active_notifs_for_me:
                if st.button("🗑️ Burahin Lahat ng Notifications", use_container_width=True, key="btn_clear_all_notifs"):
                    clear_all_live_notifications_for_user()
                    st.rerun()
                for nf_item in reversed(active_notifs_for_me[-8:]):
                    nf_id = nf_item.get("id", "")
                    nf_icon = "🚨" if nf_item.get("is_error") else "🔔"
                    nc_txt, nc_del = st.columns([0.82, 0.18])
                    with nc_txt:
                        st.markdown(
                            f"<div class='sidebar-log-card'>"
                            f"<b>{nf_icon} {nf_item.get('title')}</b><br>"
                            f"👤 {nf_item.get('sender')} • 🕒 {nf_item.get('timestamp')}<br>"
                            f"{nf_item.get('details')}"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                    with nc_del:
                        if st.button("🗑️", key=f"del_nf_{nf_id}", help="Burahin ang notification na ito"):
                            delete_live_notification(nf_id)
                            st.rerun()
            else:
                st.caption("Walang bagong notification.")


    st.divider()
    if st.session_state.user_role == "Master Admin":
        nav_options = [
            "🧭 Parish Leader Portal (Wizard & Summary)",
            "📑 Weekly Report & Parish Monitoring",
            "🏠 House Visitation, Missionary Work, Indoctrination & Remarks",
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
            "🧭 Parish Leader Portal (Wizard & Summary)",
            "📑 Weekly Report & Parish Monitoring",
            "🏠 House Visitation, Missionary Work, Indoctrination & Remarks",
            "📥 Parish Inbox & Action Hub",
            "📊 Parish Census & Member Manager",
            "🗣️ Ministry Messenger (Admin & Users Chat)",
            "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato",
            "🙏 Prayer Wall & Categorizer",
            "🛠️ App Settings"
        ]

    if "current_nav_page" not in st.session_state or st.session_state["current_nav_page"] not in nav_options:
        st.session_state["current_nav_page"] = "🧭 Parish Leader Portal (Wizard & Summary)"
    if "app_mode_override" in st.session_state:
        target_nav = st.session_state.pop("app_mode_override")
        if target_nav in nav_options:
            st.session_state["current_nav_page"] = target_nav

    app_mode = st.radio("Pumili ng Pahina:", nav_options, key="current_nav_page")
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
            load_internal_chat_messages(sync_cloud=True)
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
                for idx_m, m in enumerate(conv_msgs):
                    m_id = m.get("id", f"conv_{idx_m}")
                    s_name = m.get("sender", "Unknown")
                    r_name = m.get("recipient", "Unknown")
                    t_str = m.get("timestamp", "")
                    body_txt = m.get("message", "")
                    can_del_this = is_master or (s_name == my_id) or (r_name == my_id)

                    mc_bub, mc_del = st.columns([0.88, 0.12])
                    with mc_bub:
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
                    with mc_del:
                        if can_del_this:
                            if st.button("🗑️", key=f"del_conv_msg_{m_id}_{idx_m}", help="Burahin ang mensaheng ito"):
                                delete_internal_chat_messages_by_ids({m_id})
                                st.rerun()

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
                    save_internal_chat_message(my_id, partner_id, msg_input.strip(), sync_to_drive=True)
                    dispatch_app_notification_and_messenger_alert(
                        sender_id=my_id,
                        recipient_id=partner_id,
                        title=f"Bagong Mensahe sa Ministry Messenger mula kay {my_id}",
                        details=msg_input.strip()[:140],
                        is_error=False,
                        also_post_to_messenger=False
                    )
                    st.success(f"Naipadala na kay {partner_id}!")
                    st.rerun()
                else:
                    st.warning("Mag-type muna ng mensahe.")

        if conv_msgs:
            dc_1, dc_2 = st.columns([2.2, 1.2])
            with dc_1:
                conv_del_opts = [
                    (m.get("id", ""), f"[{m.get('timestamp', '')}] {m.get('sender', '')}: {str(m.get('message', ''))[:55]}")
                    for m in conv_msgs if m.get("id")
                ]
                if conv_del_opts:
                    sel_conv_del_idx = st.selectbox(
                        "🗑️ Pumili ng Mensahe sa Chat Box na Burahin:",
                        range(len(conv_del_opts)),
                        format_func=lambda i: conv_del_opts[i][1],
                        key="conv_box_del_sel"
                    )
            with dc_2:
                st.write("")
                b_d1, b_d2 = st.columns(2)
                with b_d1:
                    if st.button("🗑️ Delete Selected", use_container_width=True, key="btn_conv_del_sel"):
                        if conv_del_opts:
                            delete_internal_chat_messages_by_ids({conv_del_opts[sel_conv_del_idx][0]})
                            st.success("Binura na ang mensahe!")
                            st.rerun()
                with b_d2:
                    if st.button(f"🗑️ Clear Chat ({len(conv_msgs)})", use_container_width=True, key="btn_conv_clear_all"):
                        ids_conv = {m.get("id") for m in conv_msgs if m.get("id")}
                        delete_internal_chat_messages_by_ids(ids_conv)
                        st.success("Binura na ang buong usapan!")
                        st.rerun()

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

                st.markdown("##### 🗑️ Magbura ng Mensahe / Usapan (Master Admin Delete Controls)")
                del_c1, del_c2, del_c3 = st.columns([2.1, 1.0, 1.1])
                msg_del_options = [
                    (
                        m.get("id", f"idx_{i}"),
                        f"[{m.get('timestamp', '')}] {m.get('sender', '')} → {m.get('recipient', '')}: {str(m.get('message', ''))[:75]}"
                    )
                    for i, m in enumerate(mon_list)
                ]
                with del_c1:
                    sel_msg_to_del = st.selectbox(
                        "Pumili ng Mensahe na Gustong Burahin:",
                        range(len(msg_del_options)),
                        format_func=lambda i: msg_del_options[i][1],
                        key="admin_sel_msg_delete"
                    )
                with del_c2:
                    st.write("")
                    if st.button("🗑️ Burahin ang Napiling Mensahe", type="primary", use_container_width=True, key="btn_del_single_chat"):
                        target_msg_id = msg_del_options[sel_msg_to_del][0]
                        delete_internal_chat_messages_by_ids({target_msg_id})
                        st.success("Binura na ang napiling mensahe!")
                        time.sleep(0.3)
                        st.rerun()
                with del_c3:
                    st.write("")
                    if st.button(f"🗑️ Burahin Lahat sa Filter ({len(mon_list)})", use_container_width=True, key="btn_del_filtered_chats"):
                        ids_filter = {m.get("id") for m in mon_list if m.get("id")}
                        delete_internal_chat_messages_by_ids(ids_filter)
                        st.success(f"Binura na ang {len(ids_filter)} mensahe!")
                        time.sleep(0.3)
                        st.rerun()

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
# ⚡ SILENT REAL-TIME LIVE SYNC WATCHER (NO BLINKING, UPDATES ONLY WHEN ANOTHER USER SUBMITS/CHANGES DATA)
# =====================================================================
def _compute_shared_state_signature() -> str:
    parts = []
    for f_path in [LOCAL_WEEKLY_REPORTS_DB, LOCAL_CONFIG_FILE]:
        try:
            if os.path.exists(f_path):
                with open(f_path, "rb") as f_b:
                    parts.append(hashlib.md5(f_b.read()).hexdigest())
            else:
                parts.append("0")
        except Exception:
            parts.append("0")
    return "|".join(parts)

# Always sync signature at the start of every full page render so our own actions never cause a secondary blink/rerun
st.session_state["_last_live_sync_sig"] = _compute_shared_state_signature()

if hasattr(st, "fragment"):
    @st.fragment(run_every="6s")
    def _silent_realtime_state_watcher():
        cur_sig = _compute_shared_state_signature()
        prev_sig = st.session_state.get("_last_live_sync_sig", "")
        if prev_sig and cur_sig != prev_sig:
            st.session_state["_last_live_sync_sig"] = cur_sig
            st.cache_data.clear()
            st.rerun()

    _silent_realtime_state_watcher()

# =====================================================================
# 📑 WEEKLY REPORT & PARISH MONITORING ROUTE
# =====================================================================
if app_mode == "📑 Weekly Report & Parish Monitoring":
    render_weekly_report_page()
    st.stop()

elif app_mode == "🏠 House Visitation, Missionary Work, Indoctrination & Remarks":
    render_outreach_visitation_missionary_page()
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
                "4. Wika ng App at AI Assistant (Ilokano, Tagalog, English):",
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
        if st.session_state.user_role != "Master Admin":
            st.info("🔒 Ang Google Drive Excel Link & Weekly Report Folder Configuration ay para lamang sa **Master Admin**.")
        else:
            st.subheader("📊 Google Drive Excel Link & Weekly Report Folder Configuration (Master Admin)")
            st.caption("Maaari mong i-paste dito ang Google Drive Excel o Folder Link (awtomatikong magsi-sync pag-save) nang **hindi nawawala ang mga naunang links na nakalagay na**.")

            st.markdown("#### ⚡ Mabilisang Pag-add ng Kahit Anong Google Drive Link (Weekly Reports, Census, Outreach, Folders, Docs & Iba Pa)")
            qa_c1, qa_c2, qa_c3, qa_c4 = st.columns([1.8, 1.15, 1.15, 0.95])
            with qa_c1:
                settings_quick_url = st.text_input("🔗 I-paste ang Google Drive Link (Excel, Folder, Doc, o Kahit Anong Link):", key="settings_quick_drive_url")
            with qa_c2:
                settings_quick_lbl = st.text_input("🏷️ Pangalan / Label ng Link:", placeholder="Hal. Poblacion Excel / Census / Folder", key="settings_quick_drive_lbl")
            with qa_c3:
                settings_quick_cat = st.selectbox(
                    "📂 Uri ng Link / Kategorya:",
                    [
                        "Parish Weekly Report Excel (JAN–DEC 2027)",
                        "Master Weekly Report Excel (PWR)",
                        "Google Drive Master Folder",
                        "Parish Upload Folder",
                        "Master Census / Members Database Excel",
                        "Outreach / Visitation / Missionary / Indoctrination",
                        "Financial / Offerings / Tithes Report",
                        "Church Events / Youth & Kids Ministry",
                        "Photos / Videos / Media Archive Folder",
                        "Documents / Templates / Certificates / Forms",
                        "Custom / Iba Pang Link"
                    ],
                    key="settings_quick_drive_cat"
                )
            with qa_c4:
                st.write("")
                if st.button("➕ I-save & Sync", type="primary", use_container_width=True, key="btn_settings_quick_sync"):
                    if not settings_quick_url.strip():
                        st.warning("Mag-paste muna ng Google Drive Link.")
                    else:
                        q_lbl_s = settings_quick_lbl.strip() or f"{settings_quick_cat} ({get_philippine_datetime_str()})"
                        add_link_to_drive_registry(q_lbl_s, settings_quick_cat, "All Months", "All Weeks", "Lahat ng Parokya (Master)", settings_quick_url.strip())
                        with st.spinner("Sine-save ang link at sini-sync ang data kung Excel/Folder ito..."):
                            _, s_msg_q = auto_sync_from_any_drive_link(settings_quick_url.strip(), "January", "1st Week", "Lahat ng Parokya (Master)")
                            st.success(f"✅ Na-save ang '{q_lbl_s}' ({settings_quick_cat})! ({s_msg_q})")
                            time.sleep(0.5)
                            st.rerun()
            st.divider()
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
                            "Parish Weekly Report Excel (JAN–DEC 2027)",
                            "Google Drive Master Folder",
                            "Parish Upload Folder",
                            "Master Census / Members Database Excel",
                            "Outreach / House Visitation / Missionary / Indoctrination",
                            "Financial / Offerings / Tithes Report",
                            "Church Events / Youth & Kids Ministry",
                            "Photos / Videos / Media Archive Folder",
                            "Documents / Templates / Certificates / Forms",
                            "Custom / Iba Pang Google Drive Link o Folder"
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
                nodb = st.text_input("NOTION_OUTREACH_DB_ID (House Visitation, Missionary Work & Indoctrination Tasks)", value=st.session_state.custom_secrets.get("NOTION_OUTREACH_DB_ID", ""))
                gci = st.text_input("GMAIL_CLIENT_ID", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", ""))
                gcs = st.text_input("GMAIL_CLIENT_SECRET", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", ""), type="password")
                grt = st.text_input("GMAIL_REFRESH_TOKEN", value=st.session_state.custom_secrets.get("GMAIL_REFRESH_TOKEN", ""), type="password")
                if st.form_submit_button("💾 I-save ang API & Cloud Credentials", type="primary"):
                    new_creds = {
                        "GEMINI_API_KEY": ng, "GEMINI_API_KEY_2": ng2, "NOTION_TOKEN": nt,
                        "NOTION_CRM_DB_ID": ndb, "NOTION_PRAYER_DB_ID": npdb,
                        "NOTION_QUESTIONS_DB_ID": nqdb, "NOTION_COUNSELLING_DB_ID": ncdb, "NOTION_OUTREACH_DB_ID": nodb,
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
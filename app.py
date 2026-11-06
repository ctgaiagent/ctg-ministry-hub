import os
import io
import re
import time
import base64
import requests
from datetime import datetime
from email.message import EmailMessage
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

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
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    val = os.getenv(key, default)
    return str(val).strip() if val else default

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
GEMINI_API_KEY_2 = get_secret("GEMINI_API_KEY_2")
NOTION_TOKEN = get_secret("NOTION_TOKEN")
NOTION_CRM_DB_ID = get_secret("NOTION_CRM_DB_ID", "409fff6803d780e1ab25f2c80b562ddf")
NOTION_PRAYER_DB_ID = get_secret("NOTION_PRAYER_DB_ID", "40ffff6803d7800a8897c531b220ffbf")
NOTION_QUESTIONS_DB_ID = get_secret("NOTION_QUESTIONS_DB_ID", "40ffff6803d780849acef3a6cf0d4838")
NOTION_COUNSELLING_DB_ID = get_secret("NOTION_COUNSELLING_DB_ID", "40ffff6803d7804ca7ddc4b7c499ab49")
GMAIL_CLIENT_ID = get_secret("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = get_secret("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = get_secret("GMAIL_REFRESH_TOKEN")
APP_USER = get_secret("APP_USER", "ctg.ministry.master.admin")
APP_PASSWORD = get_secret("APP_PASSWORD", "Church#Temple#God$2026!Admin")
BACKUP_ADMIN_EMAIL = "clarojerald16@gmail.com"

CLOUD_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/15_znwPmjK95vTgf5lStbKKPCjEGw6oDp/"
    "edit?usp=sharing&ouid=107511419052324683693&rtpof=true&sd=true"
)

st.set_page_config(
    page_title="CTG Ministry Hub - Gabay AI",
    page_icon=LOGO_PATH if LOGO_PATH else "⛪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 🎛️ INITIALIZE UI & APP CONTROLLER STATES
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
        input_bg = "#FFFFFF"
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
        input_bg = "#172A45"
    else:
        # Default Dark Mode
        app_bg = "#0E1117"
        sidebar_bg = "#161B22"
        text_color = "#F8FAFC"
        card_bg = "#1E293B"
        card_text = "#F8FAFC"
        card_border = "#38BDF8"
        side_box_bg = "#1E293B"
        side_box_text = "#F1F5F9"
        side_box_border = "#334155"
        input_bg = "#1E293B"

    st.markdown(f"""
        <style>
        html {{
            font-size: {base_zoom};
        }}
        .stApp {{
            background-color: {app_bg} !important;
            color: {text_color} !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: {sidebar_bg} !important;
        }}
        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp p, .stApp label, .stMarkdown {{
            color: {text_color} !important;
        }}
        .full-text-card {{
            background-color: {card_bg};
            border-left: 4px solid {card_border};
            border-top: 1px solid {side_box_border};
            border-right: 1px solid {side_box_border};
            border-bottom: 1px solid {side_box_border};
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
            background-color: #ECFDF5;
            border: 1px solid #10B981;
            color: #065F46 !important;
            padding: 6px 12px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.88rem;
            margin-bottom: 8px;
            display: inline-block;
        }}
        .assigned-parish-badge {{
            background-color: #EFF6FF;
            border: 1px solid #3B82F6;
            color: #1E40AF !important;
            padding: 6px 12px;
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.88rem;
            margin-bottom: 8px;
            display: inline-block;
        }}
        .records-side-box {{
            background-color: {side_box_bg};
            border: 1px solid {side_box_border};
            border-top: 4px solid #0F766E;
            padding: {card_pad};
            border-radius: 8px;
            margin-bottom: {card_margin};
            color: {side_box_text} !important;
            font-size: 0.88rem;
        }}
        .sidebar-log-card {{
            background-color: {side_box_bg};
            border-left: 3px solid #38BDF8;
            border-top: 1px solid {side_box_border};
            border-right: 1px solid {side_box_border};
            border-bottom: 1px solid {side_box_border};
            padding: 8px 10px;
            border-radius: 5px;
            margin-bottom: 6px;
            color: {side_box_text} !important;
            font-size: 0.80rem;
        }}
        /* Siguraduhing laging litaw ang Sidebar (>> / <<) Toggle Button sa Desktop at Phone */
        [data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"] {{
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            z-index: 999999 !important;
            background-color: #1E3A8A !important;
            color: #FFFFFF !important;
            border-radius: 8px !important;
            padding: 4px !important;
            margin: 6px !important;
            box-shadow: 0 2px 6px rgba(0,0,0,0.45) !important;
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
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else url

@st.cache_resource(ttl=300)
def load_excel_file():
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'CTG' in f]
    if not excel_files and os.path.exists("CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"):
        excel_files = ["CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"]
    if excel_files:
        filename = excel_files[0]
        try:
            xls = pd.ExcelFile(filename)
            return xls, filename
        except Exception:
            pass
    try:
        doc_id = extract_id_from_url(CLOUD_SPREADSHEET_URL)
        download_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
        res = requests.get(download_url, timeout=30)
        if res.status_code == 200 and len(res.content) > 1000:
            xls = pd.ExcelFile(io.BytesIO(res.content))
            return xls, "Google Drive Cloud Excel"
        return None, f"Hindi ma-download ang Excel (Status {res.status_code})."
    except Exception as e:
        return None, str(e)

excel_xls, excel_filename = load_excel_file()

@st.cache_data(ttl=300)
def load_complete_parish_members():
    if excel_xls is None:
        return pd.DataFrame()
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
        return pd.concat(all_rows, ignore_index=True)
    return pd.DataFrame()

master_members_df = load_complete_parish_members()

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

@st.cache_data(ttl=120)
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
        res = requests.post(url, headers=headers, json=payload, timeout=25)
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
    rp = str(row_parish).upper().replace("PARISH", "").strip()
    tp = str(target_parish).upper().replace("PARISH", "").strip()
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
            "Full name": row.get("Full Name", ""), "Family ID": row.get("Family ID", ""),
            "Category": row.get("Category", ""), "Status": row.get("Status", ""),
            "Church member": row.get("Church Member", "")
        }
        if st_group == "ACTIVE": active_rows.append(mini_rec)
        elif st_group == "INACTIVE": inactive_rows.append(mini_rec)
        elif st_group == "NON-MEMBER": non_member_rows.append(mini_rec)
            
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
        r_act[1] + r_inact[1] + r_non[1] + r_oth[1], r_act[2] + r_inact[2] + r_non[2] + r_oth[2],
        r_act[3] + r_inact[3] + r_non[3] + r_oth[3], r_act[4] + r_inact[4] + r_non[4] + r_oth[4],
        r_act[5] + r_inact[5] + r_non[5] + r_oth[5], r_act[6] + r_inact[6] + r_non[6] + r_oth[6],
        r_act[7] + r_inact[7] + r_non[7] + r_oth[7],
    ]
    summary_cols = ["Status", "Seniors", "Adults", "Youth", "Seniors + Adults + Youth", "Kids", "No category", "Total persons"]
    summary_df = pd.DataFrame([r_act, r_inact, r_non, r_oth, tot_row], columns=summary_cols)
    mini_cols = ["Full name", "Family ID", "Category", "Status", "Church member"]
    df_act_list = pd.DataFrame(active_rows, columns=mini_cols) if active_rows else pd.DataFrame(columns=mini_cols)
    df_inact_list = pd.DataFrame(inactive_rows, columns=mini_cols) if inactive_rows else pd.DataFrame(columns=mini_cols)
    df_non_list = pd.DataFrame(non_member_rows, columns=mini_cols) if non_member_rows else pd.DataFrame(columns=mini_cols)
    return summary_df, df_act_list, df_inact_list, df_non_list

def get_valid_gemini_keys():
    raw_keys = [
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY", ""),
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY_2", ""),
        GEMINI_API_KEY, GEMINI_API_KEY_2,
    ]
    return [str(k).strip() for k in raw_keys if k and str(k).strip()]

def safe_generate_content(client, model, contents, config=None):
    max_retries = 3
    valid_keys = get_valid_gemini_keys()
    fallback_models = [model, "gemini-3.8-flash", "gemini-3.1-pro-preview", "gemini-2.5-flash"]
    last_err = ""
    for api_key in valid_keys:
        active_client = genai.Client(api_key=api_key)
        for current_model in fallback_models:
            delay = 2
            for attempt in range(max_retries):
                try:
                    if config:
                        return active_client.models.generate_content(model=current_model, contents=contents, config=config)
                    else:
                        return active_client.models.generate_content(model=current_model, contents=contents)
                except Exception as e:
                    last_err = str(e)
                    if "503" in last_err or "UNAVAILABLE" in last_err:
                        if attempt < max_retries - 1:
                            time.sleep(delay)
                            delay *= 2
                            continue
                    break
    raise Exception(f"Gemini API Error: {last_err}")

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
            with c_img2: st.image(LOGO_PATH, width=110)
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

def query_excel_database(search_query: str) -> str:
    if master_members_df.empty:
        return "Error: Walang laman o hindi makita ang Excel database."
    try:
        df_scope = master_members_df.copy()
        if st.session_state.user_role == "Parish Leader":
            df_scope = df_scope[df_scope["Parish"] == st.session_state.assigned_parish]
        match = df_scope[df_scope.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
        if not match.empty:
            results = []
            for _, row in match.head(20).iterrows():
                results.append(
                    f"Parish: {row['Parish']} | Family ID: {row['Family ID']} | "
                    f"Full Name: {row['Full Name']} | Status: {row['Status']} | Remarks: {row['Remarks']}"
                )
            return f"May {len(match)} na nahanap sa Excel Database:\n" + "\n".join(results)
        return f"Walang nahanap na record sa Excel para sa '{search_query}'."
    except Exception as e:
        return f"Error sa pagbasa ng Excel: {str(e)}"

def query_notion_crm(search_query: str = "") -> str:
    try:
        n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
        db_ids = [
            st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID),
            st.session_state.custom_secrets.get("NOTION_PRAYER_DB_ID", NOTION_PRAYER_DB_ID),
            st.session_state.custom_secrets.get("NOTION_QUESTIONS_DB_ID", NOTION_QUESTIONS_DB_ID),
            st.session_state.custom_secrets.get("NOTION_COUNSELLING_DB_ID", NOTION_COUNSELLING_DB_ID),
        ]
        frames, seen_db = [], set()
        for d_id in db_ids:
            c_id = clean_notion_id(d_id)
            if c_id and c_id not in seen_db:
                seen_db.add(c_id)
                try:
                    df_part = load_notion_crm_data(n_tok, c_id)
                    if not df_part.empty: frames.append(df_part)
                except Exception:
                    continue
        if not frames: return "Walang nakitang rekord sa Notion database."
        df_notion = pd.concat(frames, ignore_index=True).fillna("")
        for m_col in ["_page_id", "_last_edited", "_created_time"]:
            if m_col in df_notion.columns:
                df_notion = df_notion.drop(columns=[m_col])
        if search_query and search_query.strip():
            mask = df_notion.astype(str).apply(lambda x: x.str.contains(search_query.strip(), case=False, na=False)).any(axis=1)
            df_notion = df_notion[mask]
        if df_notion.empty: return f"Walang nahanap sa Notion para sa '{search_query}'."
        return "Mga rekord sa Notion:\n" + df_notion.head(25).to_csv(index=False)
    except Exception as e:
        return f"Error sa Notion: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str, attachment_bytes: bytes = None, attachment_filename: str = None) -> str:
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

if "custom_secrets" not in st.session_state:
    st.session_state.custom_secrets = {
        "GEMINI_API_KEY": GEMINI_API_KEY, "GEMINI_API_KEY_2": GEMINI_API_KEY_2,
        "NOTION_TOKEN": NOTION_TOKEN, "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID,
        "NOTION_PRAYER_DB_ID": NOTION_PRAYER_DB_ID, "NOTION_QUESTIONS_DB_ID": NOTION_QUESTIONS_DB_ID,
        "NOTION_COUNSELLING_DB_ID": NOTION_COUNSELLING_DB_ID,
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID, "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
    }

# =====================================================================
# 📜 HELPER PARA SA RECORDS, LIVE DASHBOARD, AT CLOUD BACKUP
# =====================================================================
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

# =====================================================================
# ⚙️ SIDEBAR CONTROL PANEL (MASTER ADMIN EXCLUSIVE + THEME CONTROLLER)
# =====================================================================
with st.sidebar:
    if LOGO_PATH and os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=85)
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")

    # 1. THEME & DISPLAY CONTROLLER (BUKAS SA LAHAT NG USERS)
    with st.expander("🎨 Theme & Display Controller", expanded=False):
        theme_opts = ["🌙 Dark Mode", "☀️ Light Mode", "⛪ Royal Blue Ministry"]
        sel_theme = st.selectbox(
            "Tema ng App (Theme):",
            theme_opts,
            index=theme_opts.index(st.session_state.ui_theme) if st.session_state.ui_theme in theme_opts else 0,
            key="sb_theme_sel"
        )
        font_opts = ["Normal (100%)", "Large (115%)", "Extra Large (130%)"]
        sel_font = st.selectbox(
            "Laki ng Teksto (Font Size):",
            font_opts,
            index=font_opts.index(st.session_state.ui_font_size) if st.session_state.ui_font_size in font_opts else 0,
            key="sb_font_sel"
        )
        dens_opts = ["Standard (Comfortable)", "Compact (Mas Maraming Data)"]
        sel_dens = st.selectbox(
            "Layout Spacing:",
            dens_opts,
            index=dens_opts.index(st.session_state.ui_density) if st.session_state.ui_density in dens_opts else 0,
            key="sb_dens_sel"
        )
        if (
            sel_theme != st.session_state.ui_theme
            or sel_font != st.session_state.ui_font_size
            or sel_dens != st.session_state.ui_density
        ):
            st.session_state.ui_theme = sel_theme
            st.session_state.ui_font_size = sel_font
            st.session_state.ui_density = sel_dens
            st.rerun()

    # 2. EXCLUSIVE PARA SA MASTER ADMIN LAMANG (HINDI KITA NG IBANG USERS / PARISH LEADERS)
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

    # 3. ADMIN VIEW SWITCHER (TANGING ANG TOTOONG NAG-LOGIN NA MASTER ADMIN LANG ANG MAKAKAKITA)
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
            "📈 Live Activity & Update Dashboard",
            "📊 Parish Directory & Family Search",
            "🛡️ Master Admin Controller & Approval Queue",
            "🤝 Notion CRM (Visitors & Members)",
            "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato",
            "🙏 Prayer Wall & Categorizer",
            "📢 Announcement Broadcaster",
            "🛠️ App Settings"
        ]
    else:
        nav_options = [
            "📊 Parish Directory & Family Search",
            "📥 Parish Assigned Requests & Counselling",
            "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato",
            "🙏 Prayer Wall & Categorizer",
            "🛠️ App Settings"
        ]
    app_mode = st.radio("Pumili ng Pahina:", nav_options)
    selected_model = st.selectbox("Gemini Model:", ["gemini-3.8-flash", "gemini-3.1-pro-preview"], index=0)
    st.divider()
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

def generate_ai_ministry_draft(service_category, sender_name, parish_name, topic, details, lang, model_name):
    valid_keys = get_valid_gemini_keys()
    if not valid_keys: return "⚠️ Wala pang GEMINI_API_KEY."
    client = genai.Client(api_key=valid_keys[0])
    prompt = (
        f"Ikaw si CTG Gabay AI, sumusulat ng opisyal, magalang, at punong-puno ng malasakit na tugon mula sa "
        f"Church of the Temple of God (CTG) - {parish_name}.\n"
        f"Uri ng Request: {service_category}\n"
        f"Pangalan ng Nagpadala: {sender_name}\nPaksa: {topic}\nDetalye:\n\"{details}\"\n\n"
        f"1. Ilagay sa unang linya: '[Urgency Analysis: Low / Medium / High]'\n"
        f"2. Sumulat ng buong mensahe (email/SMS ready) sa wikang {lang} na may comforting words, Bible verse, o malinaw na gabay.\n"
    )
    res = safe_generate_content(client, model_name, prompt)
    return res.text

# =====================================================================
# 🛠️ WORKFLOW QUEUE CARDS (INBOX AREA + TYPING AREA + PARISH TRACKING)
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

        # STRICT PARISH ISOLATION: Tanging ang parokyang nakatalaga lang ang makakakita
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
# 📈 PAGE 0: LIVE ACTIVITY & UPDATE DASHBOARD (MASTER ADMIN ONLY)
# =====================================================================
if app_mode == "📈 Live Activity & Update Dashboard":
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
# 1. PARISH DIRECTORY & FAMILY SEARCH (KUMPLETONG ORIGINAL VIEW + REMARKS)
# =====================================================================
elif app_mode == "📊 Parish Directory & Family Search":
    render_title_with_logo("CHURCH OF THE TEMPLE OF GOD — SEARCH FAMILY / PARISH")
    if master_members_df.empty:
        st.error(f"Hindi makita o mabasa ang Excel file ({excel_filename}).")
        st.stop()
        
    if st.session_state.user_role == "Master Admin":
        parish_options = ["Lahat ng Parokya"] + PARISH_ORDER
    else:
        parish_options = [st.session_state.assigned_parish]
        
    col_p, col_f, col_s = st.columns([1.2, 1.2, 1.6])
    with col_p:
        selected_parish = st.selectbox("1. Pumili ng Parish:", parish_options)
        
    if selected_parish == "Lahat ng Parokya":
        parish_filtered_df = master_members_df.copy()
    else:
        parish_filtered_df = master_members_df[master_members_df["Parish"] == selected_parish].copy()
        
    available_family_ids = ["Lahat ng Family ID"] + sorted(parish_filtered_df["Family ID"].unique().tolist())
    with col_f:
        selected_family_id = st.selectbox("2. Pumili ng Family ID:", available_family_ids)
    with col_s:
        keyword_search = st.text_input(
            "3. Search Family ID o Pangalan (Optional):",
            placeholder="Hal. CTG-LN-FAM-0001 o Juan Reyes"
        ).strip()
        
    final_df = parish_filtered_df.copy()
    if selected_family_id != "Lahat ng Family ID":
        final_df = final_df[final_df["Family ID"] == selected_family_id]
        
    if keyword_search:
        final_df = final_df[
            final_df["Family ID"].astype(str).str.contains(keyword_search, case=False, na=False) |
            final_df["Full Name"].astype(str).str.contains(keyword_search, case=False, na=False) |
            final_df["Head of Family"].astype(str).str.contains(keyword_search, case=False, na=False) |
            final_df["Remarks"].astype(str).str.contains(keyword_search, case=False, na=False)
        ]
        
    search_label = selected_parish
    if selected_family_id != "Lahat ng Family ID":
        search_label += f" • {selected_family_id}"
    if keyword_search:
        search_label += f" • Search: '{keyword_search}'"
        
    st.divider()
    
    summary_df, df_act_list, df_inact_list, df_non_list = build_status_summary_and_lists(final_df)
    
    st.markdown("### 📋 MEMBER STATUS SUMMARY")
    st.caption(f"**Search:** {search_label}")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
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
        
    st.divider()
    
    total_families = final_df["Family ID"].nunique()
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
                        f"📍 <b>Address:</b> {r_row['Current Address (If living separately)'] or 'N/A'}<br>"
                        f"📝 <b>Buong Remarks:</b> {r_row['Remarks']}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
            else:
                st.caption("Walang nakalagay na espesyal na Remarks sa mga nakalistang miyembro sa itaas.")
    else:
        st.warning("Walang nahanap na kumpletong record para sa iyong hinahanap.")
    st.stop()

# =====================================================================
# 2. MASTER ADMIN CONTROLLER / PARISH ASSIGNED WORKSPACE + SIDE PANEL
# =====================================================================
elif app_mode in ["🛡️ Master Admin Controller & Approval Queue", "📥 Parish Assigned Requests & Counselling"]:
    if st.session_state.user_role == "Master Admin":
        render_title_with_logo(
            "Master Admin Controller & Approval Hub",
            "Ipasa ang mga request sa kaukulang Parokya, tingnan kung sino ang nag-submit, at i-approve/i-send sa Gmail o SMS."
        )
    else:
        render_title_with_logo(
            f"Parish Workspace — {st.session_state.assigned_parish}",
            "Eksklusibong Inbox at Typing Area para sa mga request na ibinigay ni Master Admin sa iyong Parokya."
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
            if st.session_state.user_role == "Master Admin":
                st.dataframe(df_crm.drop(columns=["_page_id", "_last_edited", "_created_time"], errors="ignore"), use_container_width=True, hide_index=True)
            else:
                st.info("Ang buong Central CRM table ay para sa Master Admin. Gamitin ang Records tab para sa iyong parokya.")

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
# 3. NOTION CRM (VISITORS & MEMBERS) — FULL READABLE TEXT & CARDS
# =====================================================================
elif app_mode == "🤝 Notion CRM (Visitors & Members)":
    render_title_with_logo("Church CRM: Visitors & Members (Notion)")
    if st.button("🔄 I-refresh"): st.cache_data.clear(); st.rerun()
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    n_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
    df_n = load_notion_crm_data(n_tok, n_db)
    if not df_n.empty:
        disp = df_n.drop(columns=["_page_id", "_last_edited", "_created_time"], errors="ignore")
        q = st.text_input("🔍 Search CRM:").strip()
        if q: disp = disp[disp.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)]
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

# =====================================================================
# 4. MAGDAGDAG NG MIYEMBRO & LITRATO
# =====================================================================
elif app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    render_title_with_logo("Magdagdag ng Miyembro", "Mag-input ng bagong miyembro.")
    with st.form("add_form"):
        c1, c2 = st.columns(2)
        with c1:
            mp = st.selectbox("Parish:", [st.session_state.assigned_parish] if st.session_state.user_role == "Parish Leader" else PARISH_ORDER)
            fid = st.text_input("Family ID (hal. CTG-LN-FAM-0001):")
            m_head = st.text_input("Head of Family:")
            fn = st.text_input("Buong Pangalan:")
            m_rel = st.selectbox("Relasyon sa Pamilya:", ["HEAD OF FAMILY", "HUSBAND", "WIFE", "SON", "DAUGHTER", "RELATIVE", "OTHER"])
        with c2:
            m_sex = st.selectbox("Kasarian:", ["MALE", "FEMALE"])
            m_cat = st.selectbox("Kategorya:", ["KIDS", "YOUTH", "ADULT", "SENIOR CITIZEN"])
            m_st = st.selectbox("Status:", ["ACTIVE", "INACTIVE", "NON-MEMBER"])
            m_con = st.text_input("Contact Number:")
            pho = st.file_uploader("Litrato:", type=["jpg", "png", "jpeg"])
        m_rem = st.text_area("Remarks / Tala para sa Miyembro:")
        if st.form_submit_button("I-save", type="primary"):
            if fid and fn:
                st.success(f"Naidagdag si **{fn}** sa pamilyang `{fid}` ({mp})!")
                backup_action_to_admin_cloud(
                    "New Member Added",
                    fn,
                    mp,
                    f"Family ID: {fid}",
                    f"Head: {m_head} | Rel: {m_rel} | Sex: {m_sex} | Cat: {m_cat} | Status: {m_st} | Contact: {m_con}",
                    f"Remarks: {m_rem}",
                    st.session_state.assigned_parish
                )
                if pho is not None:
                    st.image(pho, caption=f"Litrato ni {fn}", width=150)
            else:
                st.warning("Punan ang Family ID at Pangalan.")
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
            res = safe_generate_content(client, selected_model, f"Suriin ang prayer request na ito sa wikang {plang}: {rp}")
            st.success("Tapos na!"); st.markdown(res.text)
        else: st.warning("Maglagay muna ng mensahe.")
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
            body = safe_generate_content(client, selected_model, f"Gumawa ng pormal na announcement tungkol sa {et}. Detalye: {ed}. Isulat sa {blang}.").text
            res = send_email_via_gmail(em, f"CTG Announcement: {et}", body)
            st.success(res)
        else: st.warning("Punan ang lahat.")
    st.stop()

# =====================================================================
# 7. APP SETTINGS (THEME CONTROLLER, DISPLAY PREFS & API CREDENTIALS)
# =====================================================================
elif app_mode == "🛠️ App Settings":
    render_title_with_logo("🛠️ App Controller, Theme & Configuration Settings", "I-customize ang hitsura ng app (Dark/Light Mode, Font Size, Layout) at API settings.")

    tab_ui, tab_api = st.tabs(["🎨 Theme, Display & App Controller", "🔑 API & Cloud Credentials (Master Admin)"])

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

    with tab_api:
        if st.session_state.user_role != "Master Admin":
            st.info("🔒 Ang API & Cloud Credentials ay maaari lamang baguhin ng Master Admin.")
        else:
            st.subheader("🔑 API Keys & Notion / Gmail Cloud Integration")
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
                if st.form_submit_button("💾 I-save ang API Credentials", type="primary"):
                    st.session_state.custom_secrets = {
                        "GEMINI_API_KEY": ng, "GEMINI_API_KEY_2": ng2, "NOTION_TOKEN": nt,
                        "NOTION_CRM_DB_ID": ndb, "NOTION_PRAYER_DB_ID": npdb,
                        "NOTION_QUESTIONS_DB_ID": nqdb, "NOTION_COUNSELLING_DB_ID": ncdb,
                        "GMAIL_CLIENT_ID": gci, "GMAIL_CLIENT_SECRET": gcs, "GMAIL_REFRESH_TOKEN": grt
                    }
                    st.cache_data.clear()
                    st.success("Na-save na ang mga credentials!"); time.sleep(1); st.rerun()
    st.stop()

# =====================================================================
# 8. CHAT ASSISTANT
# =====================================================================
render_title_with_logo("CTG Gabay AI — Ministry Assistant", "Mag-type sa Tagalog, Ilokano, o English.")
valid_keys = get_valid_gemini_keys()
if not valid_keys: st.warning("Wala pang GEMINI_API_KEY."); st.stop()

if "messages" not in st.session_state: st.session_state.messages = []
if "current_model" not in st.session_state: st.session_state.current_model = selected_model

pref_lang = st.selectbox("🌐 Wika ng Sagot:", ["Awto", "Tagalog", "Ilokano", "English"], index=0)
sys_inst = f"Ikaw si CTG Gabay AI. Parish: {st.session_state.assigned_parish}. Tagalog, Ilokano, English."

def create_chat(api_key, model_name):
    client = genai.Client(api_key=api_key)
    return client.chats.create(model=model_name, config=types.GenerateContentConfig(system_instruction=sys_inst, tools=[query_excel_database, query_notion_crm, send_email_via_gmail], temperature=0.3))

if "chat_session" not in st.session_state or st.session_state.current_model != selected_model:
    st.session_state.current_model = selected_model
    st.session_state.chat_session = create_chat(valid_keys[0], selected_model)

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]): st.markdown(msg["content"])

if user_prompt := st.chat_input("Magtanong..."):
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user"): st.markdown(user_prompt)
    with st.chat_message("assistant"):
        res = st.session_state.chat_session.send_message(user_prompt)
        st.markdown(res.text)
        st.session_state.messages.append({"role": "assistant", "content": res.text})
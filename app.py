import os
import io
import re
import time
import base64
import requests
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

CLOUD_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/15_znwPmjK95vTgf5lStbKKPCjEGw6oDp/"
    "edit?usp=sharing&ouid=107511419052324683693&rtpof=true&sd=true"
)

st.set_page_config(
    page_title="CTG Ministry Hub - Gabay AI",
    page_icon=LOGO_PATH if LOGO_PATH else "⛪",
    layout="wide"
)

st.markdown("""
    <style>
    .full-text-card {
        background-color: #F8FAFC;
        border-left: 4px solid #1E3A8A;
        padding: 12px 16px;
        margin-bottom: 10px;
        border-radius: 6px;
        color: #1F2937;
        white-space: pre-wrap;
        word-wrap: break-word;
        font-size: 0.95rem;
    }
    .urgent-card {
        background-color: #FEF2F2;
        border-left: 4px solid #DC2626;
        padding: 12px 16px;
        margin-bottom: 10px;
        border-radius: 6px;
        color: #7F1D1D;
        white-space: pre-wrap;
        word-wrap: break-word;
        font-size: 0.95rem;
    }
    [data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
    [data-testid="stAppDeployButton"] { display: none !important; }
    [data-testid="stHeaderActionElements"] { display: none !important; }
    #MainMenu { visibility: hidden !important; display: none !important; }
    footer { visibility: hidden !important; display: none !important; }
    .viewerBadge_container__1QSob, .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137, .viewerBadge_text__1JaDK { display: none !important; }
    </style>
""", unsafe_allow_html=True)

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
            row_data = {"_page_id": page.get("id", "")}
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
                other_cols = [c for c in df.columns if c not in [main_col, "_page_id"]]
                df = df[[main_col] + other_cols + ["_page_id"]]
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
            set_prop_value(["Approved By", "Reviewed By", "Approver", "Handled By"], approved_by)
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
        if "_page_id" in df_notion.columns: df_notion = df_notion.drop(columns=["_page_id"])
        if search_query and search_query.strip():
            mask = df_notion.astype(str).apply(lambda x: x.str.contains(search_query.strip(), case=False, na=False)).any(axis=1)
            df_notion = df_notion[mask]
        if df_notion.empty: return f"Walang nahanap sa Notion para sa '{search_query}'."
        return "Mga rekord sa Notion:\n" + df_notion.head(25).to_csv(index=False)
    except Exception as e:
        return f"Error sa Notion: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
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
        encoded_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": encoded_msg}).execute()
        return f"Matagumpay na naipadala ang email kay {recipient_email}!"
    except Exception as e:
        return f"Error sa Gmail: {str(e)}"

if "custom_secrets" not in st.session_state:
    st.session_state.custom_secrets = {
        "GEMINI_API_KEY": GEMINI_API_KEY, "GEMINI_API_KEY_2": GEMINI_API_KEY_2,
        "NOTION_TOKEN": NOTION_TOKEN, "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID,
        "NOTION_PRAYER_DB_ID": NOTION_PRAYER_DB_ID, "NOTION_QUESTIONS_DB_ID": NOTION_QUESTIONS_DB_ID,
        "NOTION_COUNSELLING_DB_ID": NOTION_COUNSELLING_DB_ID,
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID, "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
    }

with st.sidebar:
    if LOGO_PATH and os.path.exists(LOGO_PATH): st.image(LOGO_PATH, width=85)
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")
    st.divider()
    if st.session_state.user_role == "Master Admin":
        nav_options = [
            "📊 Parish Directory & Family Search", "🛡️ Master Admin Controller & Approval Queue",
            "🤝 Notion CRM (Visitors & Members)", "💬 AI Chat Assistant",
            "➕ Magdagdag ng Miyembro & Litrato", "🙏 Prayer Wall & Categorizer",
            "📢 Announcement Broadcaster", "🛠️ App Settings"
        ]
    else:
        nav_options = [
            "📊 Parish Directory & Family Search", "📥 Parish Assigned Requests & Counselling",
            "💬 AI Chat Assistant", "➕ Magdagdag ng Miyembro & Litrato", "🙏 Prayer Wall & Categorizer"
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
        f"Ikaw si CTG Gabay AI, sumusulat ng opisyal at magalang na tugon mula sa Church of the Temple of God (CTG) - {parish_name}.\n"
        f"Uri ng Request: {service_category}\n"
        f"Pangalan ng Nagpadala: {sender_name}\nPaksa: {topic}\nDetalye:\n\"{details}\"\n\n"
        f"1. Ilagay sa unang linya: '[Urgency Analysis: Low / Medium / High]'\n"
        f"2. Sumulat ng email sa wikang {lang}.\n"
    )
    res = safe_generate_content(client, model_name, prompt)
    return res.text

def render_workflow_queue_section(df_source: pd.DataFrame, category_label: str, key_prefix: str, default_lang: str):
    if df_source.empty:
        st.info(f"Walang nakitang rekord para sa **{category_label}**.")
        return
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    is_master = (st.session_state.user_role == "Master Admin")
    col_f1, col_f2 = st.columns([1.4, 2])
    with col_f1:
        status_filter = st.selectbox("Salain ayon sa Status:", ["Lahat ng Hindi Pa Naipapadala (Active Queue)", "Lahat ng Rekord", "Pending Review", "Assigned to Parish Leader", "Approved & Sent", "Rejected / Archived"], key=f"flt_st_{key_prefix}")
    with col_f2:
        search_txt = st.text_input("🔍 Hanapin:", key=f"flt_q_{key_prefix}").strip()
    
    for idx, row in df_source.iterrows():
        page_id = row.get("_page_id", f"{key_prefix}_{idx}")
        sender_name = get_first_non_empty(row, ["Full Name", "Name", "Pangalan"], f"Record #{idx+1}")
        email = get_first_non_empty(row, ["Email", "Email Address"], "")
        contact_no = get_first_non_empty(row, ["Contact No.", "Phone"], "N/A")
        parish = get_first_non_empty(row, ["Parish", "Parokya"], "Unassigned Parish")
        topic = get_first_non_empty(row, ["Prayer Title", "Question Topic", "Counselling Type", "Service Type", "Title"], category_label)
        details = get_first_non_empty(row, ["Prayer Details", "Question Details", "Situation Details", "Request Details", "Details", "Message"], "Walang detalye.")
        admin_status = get_first_non_empty(row, ["Admin Status", "Approval Status"], "Pending Review")
        
        if not is_master and not normalize_parish_match(parish, st.session_state.assigned_parish): continue
        if status_filter == "Lahat ng Hindi Pa Naipapadala (Active Queue)" and admin_status.lower() in ["approved & sent", "rejected / archived"]: continue
        elif status_filter != "Lahat ng Rekord" and admin_status.lower() != status_filter.lower(): continue
        if search_txt and search_txt.lower() not in f"{sender_name} {parish} {topic} {details}".lower(): continue

        draft_state_key = f"draft_val_{page_id}"
        if draft_state_key not in st.session_state: st.session_state[draft_state_key] = get_first_non_empty(row, ["AI Draft Reply", "AI Reply"], "")

        with st.expander(f"📌 {sender_name} | 📍 {parish} | 🏷️ {topic} | 📊 Status: {admin_status}", expanded=(admin_status != "Approved & Sent")):
            st.markdown(f"**📧 Email:** {email} | **📞 Contact:** {contact_no}")
            st.markdown(f"<div class='full-text-card'><b>📝 Paksa:</b> {topic}<br><b>💬 Mensahe:</b><br>{details}</div>", unsafe_allow_html=True)
            
            c_ai_btn, c_lang = st.columns([1.6, 2])
            with c_lang:
                item_lang = st.selectbox("Wika ng AI Draft:", ["Tagalog", "Ilokano", "English"], index=["Tagalog", "Ilokano", "English"].index(default_lang), key=f"lang_{page_id}")
            with c_ai_btn:
                st.write("")
                if st.button("🤖 Bumuo ng AI Draft", key=f"gen_ai_{page_id}", use_container_width=True):
                    with st.spinner("Ginagawa ng AI..."):
                        try:
                            generated = generate_ai_ministry_draft(category_label, sender_name, parish, topic, details, item_lang, selected_model)
                            st.session_state[draft_state_key] = generated
                            update_notion_page_record(n_tok, page_id, ai_draft=generated, admin_status="Pending Review")
                            st.cache_data.clear()
                            st.success("Na-save ang AI Draft!")
                        except Exception as ex: st.error(str(ex))

            edited_draft = st.text_area("✏️ Tugon para sa miyembro:", value=st.session_state[draft_state_key], height=180, key=f"editor_{page_id}")
            st.session_state[draft_state_key] = edited_draft

            if is_master:
                col_a1, col_a2, col_a3, col_a4 = st.columns([1.5, 1.5, 1.3, 1.1])
                with col_a1:
                    target_parish = st.selectbox("Ipasa sa Parokya:", PARISH_ORDER, key=f"ass_sel_{page_id}", label_visibility="collapsed")
                    if st.button("📤 Ipasa sa Parish", key=f"ass_btn_{page_id}", use_container_width=True):
                        update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="Assigned to Parish Leader", assigned_parish=target_parish)
                        st.cache_data.clear()
                        st.info(f"Naipasa sa {target_parish}!")
                with col_a2:
                    if st.button("💾 I-save", key=f"save_btn_{page_id}", use_container_width=True):
                        update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="Pending Review")
                        st.cache_data.clear()
                        st.success("Na-save!")
                with col_a3:
                    if st.button("🚀 Approve & Send", key=f"app_send_{page_id}", type="primary", use_container_width=True):
                        if email and "@" in email:
                            res_mail = send_email_via_gmail(email, f"CTG Update: {topic}", edited_draft)
                            if "Matagumpay" in res_mail:
                                update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="Approved & Sent", approved_by="Master Admin")
                                st.cache_data.clear()
                                st.success(res_mail)
                            else: st.error(res_mail)
                        else: st.error("Walang valid email.")
                with col_a4:
                    if st.button("❌ Reject", key=f"rej_btn_{page_id}", use_container_width=True):
                        update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="Rejected / Archived")
                        st.cache_data.clear()
                        st.warning("Na-reject.")
            else:
                p1, p2 = st.columns(2)
                with p1:
                    if st.button("💾 I-save", key=f"p_save_{page_id}", use_container_width=True):
                        update_notion_page_record(n_tok, page_id, ai_draft=edited_draft)
                        st.cache_data.clear()
                        st.success("Na-save.")
                with p2:
                    if st.button("📥 Ibalik kay Master Admin", key=f"p_ret_{page_id}", type="primary", use_container_width=True):
                        update_notion_page_record(n_tok, page_id, ai_draft=edited_draft, admin_status="For Master Admin Approval", approved_by=st.session_state.assigned_parish)
                        st.cache_data.clear()
                        st.success("Naibalik na sa Master Admin!")

if app_mode == "📊 Parish Directory & Family Search":
    render_title_with_logo("CHURCH OF THE TEMPLE OF GOD — SEARCH FAMILY / PARISH")
    if master_members_df.empty: st.error(f"Hindi mabasa ang Excel ({excel_filename})."); st.stop()
    parish_options = ["Lahat ng Parokya"] + PARISH_ORDER if st.session_state.user_role == "Master Admin" else [st.session_state.assigned_parish]
    col_p, col_f, col_s = st.columns([1.2, 1.2, 1.6])
    with col_p: sel_parish = st.selectbox("1. Parish:", parish_options)
    f_df = master_members_df.copy() if sel_parish == "Lahat ng Parokya" else master_members_df[master_members_df["Parish"] == sel_parish].copy()
    fam_ids = ["Lahat ng Family ID"] + sorted(f_df["Family ID"].unique().tolist())
    with col_f: sel_fam = st.selectbox("2. Family ID:", fam_ids)
    with col_s: kw = st.text_input("3. Search:", placeholder="Pangalan o ID").strip()
    if sel_fam != "Lahat ng Family ID": f_df = f_df[f_df["Family ID"] == sel_fam]
    if kw: f_df = f_df[f_df.astype(str).apply(lambda x: x.str.contains(kw, case=False, na=False)).any(axis=1)]
    st.divider()
    summary_df, act_l, inact_l, non_l = build_status_summary_and_lists(f_df)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    c1, c2, c3 = st.columns(3)
    with c1: st.success(f"**ACTIVE ({len(act_l)})**"); st.dataframe(act_l, use_container_width=True, hide_index=True)
    with c2: st.warning(f"**INACTIVE ({len(inact_l)})**"); st.dataframe(inact_l, use_container_width=True, hide_index=True)
    with c3: st.info(f"**NON-MEMBERS ({len(non_l)})**"); st.dataframe(non_l, use_container_width=True, hide_index=True)
    st.stop()

elif app_mode in ["🛡️ Master Admin Controller & Approval Queue", "📥 Parish Assigned Requests & Counselling"]:
    if st.session_state.user_role == "Master Admin":
        render_title_with_logo("Master Admin Controller & Approval Hub", "Sentralisadong Dashboard para sa Prayer, Questions, Counselling.")
    else:
        render_title_with_logo(f"Parish Assigned Requests — {st.session_state.assigned_parish}", "Mga Request na nakatalaga sa iyong Parokya.")
    
    top1, top2 = st.columns([2.5, 1])
    with top1: def_lang = st.selectbox("🌐 Wika para sa AI Draft:", ["Tagalog", "Ilokano", "English"], index=0)
    with top2: 
        st.write("")
        if st.button("🔄 I-refresh", use_container_width=True): st.cache_data.clear(); st.rerun()

    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    crm_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
    p_db = st.session_state.custom_secrets.get("NOTION_PRAYER_DB_ID", NOTION_PRAYER_DB_ID) or crm_db
    q_db = st.session_state.custom_secrets.get("NOTION_QUESTIONS_DB_ID", NOTION_QUESTIONS_DB_ID) or crm_db
    c_db = st.session_state.custom_secrets.get("NOTION_COUNSELLING_DB_ID", NOTION_COUNSELLING_DB_ID) or crm_db

    df_crm = load_notion_crm_data(n_tok, crm_db)
    df_p = load_notion_crm_data(n_tok, p_db) if clean_notion_id(p_db) != clean_notion_id(crm_db) else df_crm.copy()
    df_q = load_notion_crm_data(n_tok, q_db) if clean_notion_id(q_db) != clean_notion_id(crm_db) else df_crm.copy()
    df_co = load_notion_crm_data(n_tok, c_db) if clean_notion_id(c_db) != clean_notion_id(crm_db) else df_crm.copy()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🙏 Prayer Requests", len(df_p))
    m2.metric("📖 Church Questions", len(df_q))
    m3.metric("🤝 Counselling Requests", len(df_co))
    m4.metric("👥 Total CRM", len(df_crm))
    st.divider()

    tp, tq, tc, tcrm = st.tabs(["🙏 Prayer Requests", "📖 Church Questions", "🤝 Counselling", "📋 All CRM"])
    with tp: render_workflow_queue_section(df_p, "Prayer Request", "pray", def_lang)
    with tq: render_workflow_queue_section(df_q, "Church Question", "qna", def_lang)
    with tc: render_workflow_queue_section(df_co, "Counselling Request", "counsel", def_lang)
    with tcrm: st.dataframe(df_crm.drop(columns=["_page_id"], errors="ignore"), use_container_width=True, hide_index=True)
    st.stop()

elif app_mode == "🤝 Notion CRM (Visitors & Members)":
    render_title_with_logo("Church CRM: Visitors & Members (Notion)")
    if st.button("🔄 I-refresh"): st.cache_data.clear(); st.rerun()
    n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
    n_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
    df_n = load_notion_crm_data(n_tok, n_db)
    if not df_n.empty:
        disp = df_n.drop(columns=["_page_id"], errors="ignore")
        q = st.text_input("🔍 Search CRM:").strip()
        if q: disp = disp[disp.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)]
        st.dataframe(disp, use_container_width=True, hide_index=True)
    st.stop()

elif app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    render_title_with_logo("Magdagdag ng Miyembro", "Mag-input ng bagong miyembro.")
    with st.form("add_form"):
        c1, c2 = st.columns(2)
        with c1:
            mp = st.selectbox("Parish:", [st.session_state.assigned_parish] if st.session_state.user_role == "Parish Leader" else PARISH_ORDER)
            fid = st.text_input("Family ID (hal. CTG-LN-FAM-0001):")
            fn = st.text_input("Buong Pangalan:")
        with c2:
            st.text_input("Contact Number:")
            pho = st.file_uploader("Litrato:", type=["jpg", "png"])
        if st.form_submit_button("I-save"):
            if fid and fn: st.success(f"Naidagdag si {fn}!")
            else: st.warning("Punan ang Family ID at Pangalan.")
    st.stop()

elif app_mode == "🙏 Prayer Wall & Categorizer":
    render_title_with_logo("AI Prayer Request Categorizer")
    plang = st.selectbox("Wika:", ["Tagalog", "Ilokano", "English"])
    rp = st.text_area("Nilalaman ng Panalangin:")
    if st.button("Suriin sa AI"):
        if rp:
            client = genai.Client(api_key=get_valid_gemini_keys()[0])
            res = safe_generate_content(client, selected_model, f"Suriin ang prayer request na ito sa wikang {plang}: {rp}")
            st.success("Tapos na!"); st.markdown(res.text)
        else: st.warning("Maglagay muna ng mensahe.")
    st.stop()

elif app_mode == "📢 Announcement Broadcaster":
    render_title_with_logo("AI Announcement Broadcaster")
    blang = st.selectbox("Wika:", ["Tagalog", "Ilokano", "English"])
    et = st.text_input("Pamagat:")
    ed = st.text_area("Detalye:")
    em = st.text_input("Email ng Tatanggap:")
    if st.button("Magpadala"):
        if et and ed and em:
            client = genai.Client(api_key=get_valid_gemini_keys()[0])
            body = safe_generate_content(client, selected_model, f"Gumawa ng pormal na announcement tungkol sa {et}. Detalye: {ed}. Isulat sa {blang}.").text
            res = send_email_via_gmail(em, f"CTG Announcement: {et}", body)
            st.success(res)
        else: st.warning("Punan ang lahat.")
    st.stop()

elif app_mode == "🛠️ App Settings":
    render_title_with_logo("App Configuration & Settings")
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
        if st.form_submit_button("I-save"):
            st.session_state.custom_secrets = {
                "GEMINI_API_KEY": ng, "GEMINI_API_KEY_2": ng2, "NOTION_TOKEN": nt,
                "NOTION_CRM_DB_ID": ndb, "NOTION_PRAYER_DB_ID": npdb,
                "NOTION_QUESTIONS_DB_ID": nqdb, "NOTION_COUNSELLING_DB_ID": ncdb,
                "GMAIL_CLIENT_ID": gci, "GMAIL_CLIENT_SECRET": gcs, "GMAIL_REFRESH_TOKEN": grt
            }
            st.success("Na-save na!"); time.sleep(1); st.rerun()
    st.stop()

# =====================================================================
# CHAT ASSISTANT
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
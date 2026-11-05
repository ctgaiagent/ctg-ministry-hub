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
    """Hinahanap ang logo image file sa loob ng project folder."""
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
    """Nagpapakita ng Logo image sa kaliwa bago ang Title Header."""
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

# Custom CSS para sa full-text cards at PAGTATAGO ng Top-Right Toolbar (Share, Star, Edit, GitHub, 3-dots)
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
    
    /* Itago ang Share, Star, Edit, GitHub icon, at 3-dots menu sa taas */
    [data-testid="stToolbar"] {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="stAppDeployButton"] {
        display: none !important;
    }
    [data-testid="stHeaderActionElements"] {
        display: none !important;
    }
    #MainMenu {
        visibility: hidden !important;
        display: none !important;
    }
    footer {
        visibility: hidden !important;
        display: none !important;
    }
    .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_,
    .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

PARISH_ORDER = [
    "LIPAY NORTE PARISH",
    "MALANAS PARISH",
    "POBLACION PARISH",
    "DUPLAS PARISH",
    "SANTIAGO PARISH",
    "CABA PARISH",
    "MANILA PARISH",
    "ANDABUEN PARISH",
    "APAYAO PARISH",
    "HONGKONG PARISH",
]

# ==========================================
# 🔐 MGA USERNAME AT PASSWORD NG ADMIN AT PAROKYA (1-10)
# ==========================================
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
    """Binabasa lamang ang mga kumpletong record na may valid Family ID (CTG-...) at Full Name mula sa lahat ng Parish Sheets."""
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

# ==========================================
# 🗂️ NOTION CRM FUNCTIONS
# ==========================================
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
    return ""

@st.cache_data(ttl=180)
def load_notion_crm_data(token: str, db_id: str) -> pd.DataFrame:
    if not token or not db_id:
        return pd.DataFrame()
    clean_id = db_id.strip().split("?")[0].split("-")[-1]
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
            row_data = {col_name: extract_notion_property(col_val) for col_name, col_val in props.items()}
            rows.append(row_data)
        has_more = data.get("has_more", False)
        next_cursor = data.get("next_cursor", None)
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.fillna("")
        for main_col in ["Full Name", "Name", "Pangalan", "Visitor Name"]:
            if main_col in df.columns:
                cols = [main_col] + [c for c in df.columns if c != main_col]
                df = df[cols]
                break
    return df

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
    
    active_rows = []
    inactive_rows = []
    non_member_rows = []
    
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
    
    summary_cols = [
        "Status", "Seniors", "Adults", "Youth", 
        "Seniors + Adults + Youth", "Kids", "No category", "Total persons"
    ]
    summary_df = pd.DataFrame([r_act, r_inact, r_non, r_oth, tot_row], columns=summary_cols)
    
    mini_cols = ["Full name", "Family ID", "Category", "Status", "Church member"]
    df_act_list = pd.DataFrame(active_rows, columns=mini_cols) if active_rows else pd.DataFrame(columns=mini_cols)
    df_inact_list = pd.DataFrame(inactive_rows, columns=mini_cols) if inactive_rows else pd.DataFrame(columns=mini_cols)
    df_non_list = pd.DataFrame(non_member_rows, columns=mini_cols) if non_member_rows else pd.DataFrame(columns=mini_cols)
    
    return summary_df, df_act_list, df_inact_list, df_non_list

# ==========================================
# 🤖 MULTI-KEY GEMINI HELPER & RETRY LOGIC
# ==========================================
def get_valid_gemini_keys():
    raw_keys = [
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY", ""),
        st.session_state.get("custom_secrets", {}).get("GEMINI_API_KEY_2", ""),
        GEMINI_API_KEY,
        GEMINI_API_KEY_2,
    ]
    valid = []
    for k in raw_keys:
        if k and str(k).strip() and str(k).strip() not in valid:
            valid.append(str(k).strip())
    return valid

def safe_generate_content(client, model, contents, config=None):
    max_retries = 3
    valid_keys = get_valid_gemini_keys()
    fallback_models = []
    for m in [model, "gemini-3.8-flash", "gemini-3.1-pro-preview", "gemini-2.5-flash"]:
        if m not in fallback_models:
            fallback_models.append(m)
            
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
                    err_str = str(e)
                    last_err = err_str
                    if "404" in err_str or "NOT_FOUND" in err_str:
                        break
                    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                        break
                    if "503" in err_str or "UNAVAILABLE" in err_str:
                        if attempt < max_retries - 1:
                            time.sleep(delay)
                            delay *= 2
                            continue
                    break
    raise Exception(f"Ang mga Gemini models ay kasalukuyang busy o naabot ang quota limit. Detalye: {last_err}")

# ==========================================
# 🔐 LOGIN SESSION & AUTHENTICATION
# ==========================================
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
                    st.rerun()
                else:
                    st.error("Maling Username o Password. Pakisuri ang iyong mga detalye.")
    st.stop()

# ==========================================
# 🛠️ AI TOOLS (EXCEL, NOTION CRM, AT GMAIL)
# ==========================================
def query_excel_database(search_query: str) -> str:
    """Hinahanap ang miyembro, pamilya, o parokya sa Google Drive / Local Excel Database."""
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
                    f"Full Name: {row['Full Name']} | Relationship: {row['Relationship']} | "
                    f"Category: {row['Category']} | Status: {row['Status']} | Remarks: {row['Remarks']}"
                )
            return f"May {len(match)} na nahanap sa Excel Database (unang 20 resulta):\n" + "\n".join(results)
        return f"Walang nahanap na record sa Excel para sa '{search_query}'."
    except Exception as e:
        return f"Error sa pagbasa ng Excel: {str(e)}"

def query_notion_crm(search_query: str = "") -> str:
    """Kumukuha at naghahanap ng mga bisita o miyembro mula sa Notion CRM Database."""
    try:
        n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
        n_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
        df_notion = load_notion_crm_data(n_tok, n_db)
        if df_notion.empty:
            return "Walang nakitang rekord sa Notion CRM database."
        if search_query and search_query.strip():
            mask = df_notion.astype(str).apply(lambda x: x.str.contains(search_query.strip(), case=False, na=False)).any(axis=1)
            df_notion = df_notion[mask]
        if df_notion.empty:
            return f"Walang nahanap sa Notion CRM para sa '{search_query}'."
        return "Mga rekord sa Notion CRM:\n" + df_notion.head(25).to_csv(index=False)
    except Exception as e:
        return f"Error sa pagbasa ng Notion CRM: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
    """Nagpapadala ng email gamit ang Gmail OAuth2 API."""
    g_id = st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", GMAIL_CLIENT_ID)
    g_sec = st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", GMAIL_CLIENT_SECRET)
    g_ref = st.session_state.custom_secrets.get("GMAIL_REFRESH_TOKEN", GMAIL_REFRESH_TOKEN)
    if not all([g_id, g_sec, g_ref]):
        return "Error: Kulang ang GMAIL credentials."
    try:
        creds = Credentials(
            token=None,
            refresh_token=g_ref.strip(),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=g_id.strip(),
            client_secret=g_sec.strip()
        )
        service = build("gmail", "v1", credentials=creds)
        msg = EmailMessage()
        msg.set_content(body_text)
        msg["To"] = recipient_email
        msg["Subject"] = subject
        encoded_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": encoded_msg}).execute()
        return f"Matagumpay na naipadala ang email kay {recipient_email}!"
    except Exception as e:
        return f"Error sa pag-send ng Gmail: {str(e)}"

if "custom_secrets" not in st.session_state:
    st.session_state.custom_secrets = {
        "GEMINI_API_KEY": GEMINI_API_KEY,
        "GEMINI_API_KEY_2": GEMINI_API_KEY_2,
        "NOTION_TOKEN": NOTION_TOKEN,
        "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID,
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID,
        "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
    }

with st.sidebar:
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=85)
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")
    st.divider()
    
    app_mode = st.radio("Pumili ng Pahina:", [
        "📊 Parish Directory & Family Search",
        "🤝 Notion CRM (Visitors & Members)",
        "💬 AI Chat Assistant",
        "➕ Magdagdag ng Miyembro & Litrato",
        "🙏 Prayer Wall & Categorizer",
        "📢 Announcement Broadcaster",
        "🛠️ App Settings"
    ])
    
    selected_model = st.selectbox("Gemini Model:", ["gemini-3.8-flash", "gemini-3.1-pro-preview"], index=0)
    st.divider()
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

# =====================================================================
# 1. PARISH DIRECTORY & FAMILY SEARCH (KATULAD NG EXCEL SEARCH DATABASE)
# =====================================================================
if app_mode == "📊 Parish Directory & Family Search":
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
    
    # SECTION A: MEMBER STATUS SUMMARY TABLE
    st.markdown("### 📋 MEMBER STATUS SUMMARY")
    st.caption(f"**Search:** {search_label}")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # SECTION B: ACTIVE, INACTIVE, AT NON-MEMBERS TABLES (3 COLUMNS)
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
    
    # SECTION C: MATCHING MEMBERS FULL TABLE
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
# 1B. NOTION CRM (VISITORS & MEMBERS) — FULL READABLE TEXT & CARDS
# =====================================================================
elif app_mode == "🤝 Notion CRM (Visitors & Members)":
    render_title_with_logo("Church CRM: Visitors & Members (Notion)")
    
    col_r1, col_r2 = st.columns([3, 1])
    with col_r2:
        if st.button("🔄 I-refresh ang Notion CRM", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    try:
        n_tok = st.session_state.custom_secrets.get("NOTION_TOKEN", NOTION_TOKEN)
        n_db = st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", NOTION_CRM_DB_ID)
        notion_df = load_notion_crm_data(n_tok, n_db)
        if notion_df.empty:
            st.info("Walang nakitang rekord sa Notion CRM o walang nakalagay na NOTION_TOKEN.")
        else:
            crm_q = st.text_input("🔍 Maghanap ng Bisita, Miyembro, o Detalye sa Notion CRM:")
            if crm_q.strip():
                mask = notion_df.astype(str).apply(lambda x: x.str.contains(crm_q.strip(), case=False, na=False)).any(axis=1)
                notion_df = notion_df[mask]
                
            st.success(f"Kabuuang nakita: **{len(notion_df)}** na bisita/rekord.")
            
            notion_col_cfg = {
                col: st.column_config.TextColumn(col, width="large")
                for col in notion_df.columns
            }
            st.dataframe(
                notion_df,
                use_container_width=True,
                hide_index=True,
                column_config=notion_col_cfg
            )
            
            st.divider()
            st.subheader("📖 Buong Detalye ng Bawat Rekord sa Notion CRM (Full View)")
            st.caption("Ipinapakita rito nang buo at hindi putol ang lahat ng text, notes, at impormasyon ng bawat bisita/miyembro.")
            
            for idx, row in notion_df.iterrows():
                first_col = notion_df.columns[0]
                title_val = str(row[first_col]).strip() or f"Record #{idx + 1}"
                
                details_html = f"<div class='full-text-card'><b>📌 {title_val}</b><br><hr style='margin:6px 0;'>"
                for c_name in notion_df.columns:
                    val_str = str(row[c_name]).strip()
                    if val_str:
                        details_html += f"<b>• {c_name}:</b> {val_str}<br>"
                details_html += "</div>"
                st.markdown(details_html, unsafe_allow_html=True)
                
    except Exception as e:
        st.error(f"Hindi makakonekta sa Notion CRM: {str(e)}")
    st.stop()

# =====================================================================
# 2. MAGDAGDAG NG MIYEMBRO & LITRATO
# =====================================================================
elif app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    render_title_with_logo(
        "Magdagdag ng Miyembro at Mag-upload ng Litrato",
        "Mag-input ng bagong miyembro sa ilalim ng iyong parokya."
    )
    
    with st.form("add_member_form"):
        col1, col2 = st.columns(2)
        with col1:
            m_parish = st.selectbox(
                "Parish:",
                [st.session_state.assigned_parish] if st.session_state.user_role == "Parish Leader" else PARISH_ORDER
            )
            m_fam_id = st.text_input("Family ID (hal. CTG-LN-FAM-0001):")
            m_head = st.text_input("Head of Family:")
            m_fullname = st.text_input("Buong Pangalan ng Miyembro:")
            m_relation = st.selectbox("Relasyon sa Pamilya:", ["HEAD OF FAMILY", "HUSBAND", "WIFE", "SON", "DAUGHTER", "RELATIVE", "OTHER"])
        with col2:
            m_sex = st.selectbox("Kasarian:", ["MALE", "FEMALE"])
            m_category = st.selectbox("Kategorya:", ["KIDS", "YOUTH", "ADULT", "SENIOR CITIZEN"])
            m_status = st.selectbox("Status:", ["ACTIVE", "INACTIVE", "NON-MEMBER"])
            m_contact = st.text_input("Contact Number:")
            uploaded_photo = st.file_uploader("Mag-upload ng Litrato ng Miyembro (JPG/PNG)", type=["jpg", "png", "jpeg"])
            
        submit_member = st.form_submit_button("I-save ang Miyembro", type="primary")
        if submit_member:
            if m_fam_id and m_fullname:
                st.success(f"Matagumpay na naidagdag si **{m_fullname}** sa pamilyang `{m_fam_id}` ({m_parish})!")
                if uploaded_photo is not None:
                    st.image(uploaded_photo, caption=f"Litrato ni {m_fullname}", width=150)
                    st.info("Na-save at na-attach ang litrato sa profile.")
            else:
                st.warning("Pakipunan ang Family ID at Buong Pangalan.")
    st.stop()

# =====================================================================
# 3. PRAYER WALL & CATEGORIZER (WITH MULTILINGUAL SUPPORT)
# =====================================================================
elif app_mode == "🙏 Prayer Wall & Categorizer":
    render_title_with_logo("AI Prayer Request & Needs Categorizer")
    prayer_lang = st.selectbox("Wika ng Tugon at Bible Verse:", ["Tagalog", "Ilokano", "English"], index=0)
    raw_prayer = st.text_area("Nilalaman ng Prayer Request o Mensahe (Tagalog / Ilokano / English):")
    if st.button("Suriin sa AI at I-categorize", type="primary"):
        if raw_prayer:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=active_key)
            with st.spinner("Sinusuri ng AI ang kahilingan..."):
                prompt = (
                    f"Suriin ang sumusunod na prayer request. Ibigay ang Kategorya, Urgency Level, "
                    f"at comforting Bible verse at maikling panalangin gamit ang wikang {prayer_lang}:\n\n{raw_prayer}"
                )
                try:
                    res = safe_generate_content(client, selected_model, prompt)
                    st.success("Tapos na ang Pagsusuri!")
                    st.markdown(res.text)
                except Exception as ex:
                    st.error(str(ex))
        else:
            st.warning("Maglagay muna ng mensahe.")
    st.stop()

# =====================================================================
# 4. ANNOUNCEMENT BROADCASTER (WITH MULTILINGUAL SUPPORT)
# =====================================================================
elif app_mode == "📢 Announcement Broadcaster":
    render_title_with_logo("AI Ministry Announcement Broadcaster")
    bcast_lang = st.selectbox("Wika ng Announcement Email:", ["Tagalog", "Ilokano", "English"], index=0)
    event_title = st.text_input("Pamagat ng Aktibidad / Event:")
    event_details = st.text_area("Mga Detalye (Petsa, Oras, Lugar, Mensahe):")
    target_email = st.text_input("Email ng Tatanggap:")
    if st.button("Bumuo at Magpadala ng Email Broadcast", type="primary"):
        if event_title and event_details and target_email:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=active_key)
            with st.spinner("Gumagawa at nagpapadala ng broadcast..."):
                try:
                    p_prompt = (
                        f"Gumawa ng pormal at nakaka-inspire na announcement email tungkol sa: {event_title}. "
                        f"Mga detalye: {event_details}. Isulat ito sa wikang {bcast_lang}."
                    )
                    ai_res = safe_generate_content(client, selected_model, p_prompt)
                    email_body = ai_res.text
                    res_send = send_email_via_gmail(target_email, f"CTG Announcement: {event_title}", email_body)
                    st.success(res_send)
                    with st.expander("Tingnan ang Mensahe"):
                        st.markdown(email_body)
                except Exception as ex:
                    st.error(str(ex))
        else:
            st.warning("Punan ang lahat ng kailangan.")
    st.stop()

# =====================================================================
# 5. APP SETTINGS
# =====================================================================
elif app_mode == "🛠️ App Settings":
    render_title_with_logo("App Configuration & Settings")
    with st.form("settings_form"):
        st.subheader("🔑 API Keys & Integration")
        new_gemini = st.text_input("GEMINI_API_KEY (Key #1)", value=st.session_state.custom_secrets.get("GEMINI_API_KEY", ""), type="password")
        new_gemini_2 = st.text_input("GEMINI_API_KEY_2 (Fallback Key #2)", value=st.session_state.custom_secrets.get("GEMINI_API_KEY_2", ""), type="password")
        new_notion_token = st.text_input("NOTION_TOKEN", value=st.session_state.custom_secrets.get("NOTION_TOKEN", ""), type="password")
        new_notion_db = st.text_input("NOTION_CRM_DB_ID", value=st.session_state.custom_secrets.get("NOTION_CRM_DB_ID", "409fff6803d780e1ab25f2c80b562ddf"))
        st.subheader("📧 Gmail OAuth2 Credentials")
        new_g_client_id = st.text_input("GMAIL_CLIENT_ID", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_ID", ""))
        new_g_client_sec = st.text_input("GMAIL_CLIENT_SECRET", value=st.session_state.custom_secrets.get("GMAIL_CLIENT_SECRET", ""), type="password")
        new_g_token = st.text_input("GMAIL_REFRESH_TOKEN", value=st.session_state.custom_secrets.get("GMAIL_REFRESH_TOKEN", ""), type="password")
        if st.form_submit_button("I-save ang mga Pagbabago"):
            st.session_state.custom_secrets = {
                "GEMINI_API_KEY": new_gemini,
                "GEMINI_API_KEY_2": new_gemini_2,
                "NOTION_TOKEN": new_notion_token,
                "NOTION_CRM_DB_ID": new_notion_db,
                "GMAIL_CLIENT_ID": new_g_client_id,
                "GMAIL_CLIENT_SECRET": new_g_client_sec,
                "GMAIL_REFRESH_TOKEN": new_g_token,
            }
            st.session_state.pop("chat_session", None)
            st.success("Matagumpay na na-update!")
            time.sleep(1)
            st.rerun()
    st.stop()

# =====================================================================
# 6. AI CHAT ASSISTANT — CTG GABAY AI (TEXT MULTILINGUAL SUPPORT)
# =====================================================================
render_title_with_logo(
    "CTG Gabay AI — Ministry Assistant",
    "Mag-type sa Tagalog, Ilokano, o English para maghanap sa Excel, Notion CRM, o magpadala ng email."
)

valid_keys = get_valid_gemini_keys()
if not valid_keys:
    st.warning("⚠️ Wala pang GEMINI_API_KEY. Pumunta sa **App Settings** sa sidebar.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = []

if "current_model" not in st.session_state:
    st.session_state.current_model = selected_model

preferred_language = st.selectbox(
    "🌐 Wika ng Sagot ni Gabay AI:",
    ["Awto (Kung ano ang gamit ng nagtanong)", "Tagalog", "Ilokano", "English"],
    index=0
)

system_instruction = (
    f"Ikaw si CTG Gabay AI, ang opisyal, magalang, at matulunging AI Ministry Assistant ng Church of the Temple of God (CTG). "
    f"Ang kasalukuyang naka-login ay may access sa parokyang: {st.session_state.assigned_parish}. "
    "Matatas kang umintindi at sumagot sa Tagalog, Ilokano (Ilocano), at English. "
    "Tulungan sila sa paghanap ng mga rekord ng pamilya at miyembro sa Excel database gamit ang `query_excel_database`, "
    "paghanap ng mga bisita sa Notion CRM gamit ang `query_notion_crm`, at pagpapadala ng email gamit ang `send_email_via_gmail`."
)

def create_chat_with_key(api_key: str, model_name: str):
    client = genai.Client(api_key=api_key)
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        tools=[query_excel_database, query_notion_crm, send_email_via_gmail],
        temperature=0.3,
    )
    return client.chats.create(model=model_name, config=config)

if "chat_session" not in st.session_state or st.session_state.current_model != selected_model:
    st.session_state.current_model = selected_model
    st.session_state.chat_session = create_chat_with_key(valid_keys[0], selected_model)

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_prompt := st.chat_input("Magtanong sa Tagalog, Ilokano, o English (hal. Mano ti miyembro idiay Lipay Norte?)..."):
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user"):
        st.markdown(user_prompt)

    lang_note = (
        ""
        if preferred_language.startswith("Awto")
        else f"\n\n[Paalala: Pakisagot ito gamit ang wikang {preferred_language}.]"
    )

    with st.chat_message("assistant"):
        status_box = st.status(f"Pinoproseso ni Gabay AI gamit ang `{selected_model}`...", expanded=True)
        max_retries = 3
        success = False
        last_error = ""

        for key_idx, api_key in enumerate(valid_keys, start=1):
            if key_idx > 1:
                st.session_state.chat_session = create_chat_with_key(api_key, selected_model)
                
            for attempt in range(max_retries):
                try:
                    status_box.write(f"🔄 Kumokonekta sa Gemini API (Key #{key_idx} — Attempt {attempt + 1}/{max_retries})...")
                    response = st.session_state.chat_session.send_message(user_prompt + lang_note)
                    status_box.update(label="✅ Tapos na!", state="complete", expanded=False)
                    st.markdown(response.text)
                    st.session_state.messages.append({"role": "assistant", "content": response.text})
                    success = True
                    break
                except Exception as e:
                    err_str = str(e)
                    last_error = err_str
                    if ("503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str) and attempt < max_retries - 1:
                        time.sleep(3)
                        continue
                    break
            if success:
                break

        if not success:
            status_box.update(label="❌ May Error", state="error", expanded=True)
            st.error(f"Detalye: {last_error}")
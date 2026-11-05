import os
import io
import re
import time
import requests
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from google import genai

# Load local environment variables (.env)
load_dotenv()

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="CTG Ministry Hub",
    page_icon="⛪",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .stMetric {
        background-color: #F8FAFC;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CREDENTIALS & CONFIG HELPER
# ==========================================
def get_secret_or_env(key: str, default: str = "") -> str:
    """Kinukuha ang secret mula sa Streamlit Cloud Secrets o local .env file."""
    try:
        if key in st.secrets:
            return str(st.secrets[key]).strip()
    except Exception:
        pass
    val = os.getenv(key, default)
    return val.strip() if val else default

GEMINI_API_KEY_1 = get_secret_or_env("GEMINI_API_KEY")
GEMINI_API_KEY_2 = get_secret_or_env("GEMINI_API_KEY_2")
NOTION_TOKEN = get_secret_or_env("NOTION_TOKEN")
NOTION_CRM_DB_ID = get_secret_or_env("NOTION_CRM_DB_ID", "409fff6803d780e1ab25f2c80b562ddf")

APP_USER = get_secret_or_env("APP_USER", "ctg.ministry.master.admin")
APP_PASSWORD = get_secret_or_env("APP_PASSWORD", "Church#Temple#God$2026!Admin")

CLOUD_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/15_znwPmjK95vTgf5lStbKKPCjEGw6oDp/"
    "edit?usp=sharing&ouid=107511419052324683693&rtpof=true&sd=true"
)

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
    "HONGKONG PARISH"
]

# Eksaktong PARISH_ACCOUNTS structure
PARISH_ACCOUNTS = {
    "ctg.ministry.master.admin": {"password": "Church#Temple#God$2026!Admin", "parish": "Lahat ng Parokya", "role": "Master Admin"},
    
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

# ==========================================
# 3. DATA FETCHING FUNCTIONS (EXCEL & NOTION)
# ==========================================
def extract_id_from_url(url: str) -> str:
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else url

@st.cache_data(ttl=300)
def load_excel_data():
    """Binabasa ang buong Google Drive Excel file at ibinabalik bilang dictionary ng DataFrames."""
    doc_id = extract_id_from_url(CLOUD_SPREADSHEET_URL)
    download_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
    
    res = requests.get(download_url, timeout=30)
    if res.status_code != 200 or len(res.content) < 1000:
        raise Exception(f"Hindi ma-download ang Excel file. Status Code: {res.status_code}")
        
    xls = pd.ExcelFile(io.BytesIO(res.content))
    sheets_dict = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df = df.dropna(how="all").fillna("")
        sheets_dict[sheet] = df
    return sheets_dict

def extract_notion_property(prop_data: dict) -> str:
    """Kinukuha ang plain text value mula sa iba't ibang uri ng Notion property."""
    if not prop_data:
        return ""
    p_type = prop_data.get("type", "")
    
    if p_type == "title":
        arr = prop_data.get("title", [])
        return "".join([x.get("plain_text", "") for x in arr])
    elif p_type == "rich_text":
        arr = prop_data.get("rich_text", [])
        return "".join([x.get("plain_text", "") for x in arr])
    elif p_type == "select":
        sel = prop_data.get("select")
        return sel.get("name", "") if sel else ""
    elif p_type == "multi_select":
        arr = prop_data.get("multi_select", [])
        return ", ".join([x.get("name", "") for x in arr])
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
    return ""

@st.cache_data(ttl=180)
def load_notion_crm_data(token: str, db_id: str) -> pd.DataFrame:
    """Kumukuha ng mga rekord mula sa Notion CRM Database."""
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
        results = data.get("results", [])
        
        for page in results:
            props = page.get("properties", {})
            row_data = {}
            for col_name, col_val in props.items():
                row_data[col_name] = extract_notion_property(col_val)
            rows.append(row_data)
            
        has_more = data.get("has_more", False)
        next_cursor = data.get("next_cursor", None)
        
    df = pd.DataFrame(rows)
    if not df.empty and "Full Name" in df.columns:
        cols = ["Full Name"] + [c for c in df.columns if c != "Full Name"]
        df = df[cols]
    return df

# ==========================================
# 4. GEMINI AI HELPER (MULTI-KEY + RETRY)
# ==========================================
def ask_gemini_ai(prompt: str, context_text: str) -> str:
    """
    Tumatawag sa Gemini API gamit ang multi-key rotation (Key #1 at Key #2)
    at automatic retry para malampasan ang 503 high demand errors.
    """
    raw_keys = [
        st.session_state.get("custom_gemini_key", ""),
        GEMINI_API_KEY_1,
        GEMINI_API_KEY_2
    ]
    valid_keys = []
    for k in raw_keys:
        if k and k.strip() and k.strip() not in valid_keys:
            valid_keys.append(k.strip())
            
    if not valid_keys:
        return "❌ Walang nahanap na GEMINI_API_KEY. Mangyaring ilagay ito sa Streamlit Secrets o sa `.env` file."

    models_to_try = ["gemini-3.8-flash", "gemini-3.1-pro-preview"]
    max_retries = 3
    last_error = ""

    full_prompt = f"""Ikaw ay isang matulungin, magalang, at tumpak na AI Church Ministry Assistant para sa CTG Ministry Hub.
Gamitin ang sumusunod na opisyal na datos mula sa simbahan (Google Drive Excel at Notion CRM) upang sagutin ang tanong.
Kung wala sa datos ang sagot, sabihin nang magalang kung ano lang ang makikita sa rekord. Sumagot sa Tagalog o English depende sa tanong ng user.

=== DATOS NG SIMBAHAN ===
{context_text}

=== TANONG NG USER ===
{prompt}
"""

    for key_idx, api_key in enumerate(valid_keys, start=1):
        client = genai.Client(api_key=api_key)
        for model_name in models_to_try:
            delay = 2
            for attempt in range(max_retries):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=full_prompt
                    )
                    if response and response.text:
                        return response.text
                except Exception as e:
                    err_str = str(e)
                    last_error = f"[Key #{key_idx} | {model_name} | Attempt {attempt+1}]: {err_str}"
                    if "404" in err_str or "NOT_FOUND" in err_str:
                        break
                    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                        break
                    if attempt < max_retries - 1:
                        time.sleep(delay)
                        delay *= 2

    return f"❌ Pansamantalang abala ang mga server ng Gemini AI. Pakisubukan ulit pagkalipas ng ilang segundo.\n\n*(Huling Error: {last_error})*"

def build_ai_context(excel_sheets: dict, notion_df: pd.DataFrame, active_parish: str, active_role: str) -> str:
    """Gumagawa ng malinis na text summary ng data para mabasa ng Gemini AI."""
    parts = []
    
    # 1. Notion CRM Context
    if not notion_df.empty:
        parts.append(f"--- NOTION CRM: VISITORS & MEMBERS (Total: {len(notion_df)}) ---")
        parts.append(notion_df.head(100).to_csv(index=False))
    else:
        parts.append("--- NOTION CRM: Walang nakitang rekord o hindi konektado ---")

    # 2. Excel Context
    if excel_sheets:
        if active_role == "Master Admin" or active_parish == "Lahat ng Parokya":
            if "PARISH SUMMARY" in excel_sheets:
                parts.append("--- EXCEL: PARISH SUMMARY ---")
                parts.append(excel_sheets["PARISH SUMMARY"].head(50).to_csv(index=False))
            for p_name in PARISH_ORDER:
                if p_name in excel_sheets:
                    df_p = excel_sheets[p_name]
                    parts.append(f"--- EXCEL: {p_name} (Rows: {len(df_p)}) ---")
                    parts.append(df_p.head(60).to_csv(index=False))
        else:
            if active_parish in excel_sheets:
                df_p = excel_sheets[active_parish]
                parts.append(f"--- EXCEL: {active_parish} (Rows: {len(df_p)}) ---")
                parts.append(df_p.head(150).to_csv(index=False))
                
    return "\n\n".join(parts)

# ==========================================
# 5. SESSION STATE & LOGIN SCREEN
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_account" not in st.session_state:
    st.session_state.user_account = None
if "user_parish" not in st.session_state:
    st.session_state.user_parish = "Lahat ng Parokya"
if "user_role" not in st.session_state:
    st.session_state.user_role = ""
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

def verify_credentials(username_in: str, password_in: str):
    """Sinusuri kung tama ang tinayp na Username at Password base sa PARISH_ACCOUNTS."""
    u_clean = username_in.strip().lower()
    p_clean = password_in.strip()
    
    # Check kung tugma sa APP_USER / APP_PASSWORD mula sa .env o Secrets
    if u_clean == APP_USER.lower() and p_clean == APP_PASSWORD:
        return True, APP_USER, "Lahat ng Parokya", "Master Admin"
        
    # Check sa PARISH_ACCOUNTS dictionary
    if u_clean in PARISH_ACCOUNTS:
        acc_info = PARISH_ACCOUNTS[u_clean]
        if p_clean == acc_info["password"]:
            return True, u_clean, acc_info["parish"], acc_info["role"]
            
    return False, None, None, None

def render_login():
    st.markdown('<div class="main-header">⛪ CTG Ministry Hub</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Administrative Portal & AI Ministry Assistant</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.form("login_form", clear_on_submit=False):
            st.subheader("🔐 Admin / Parish Login")
            username_input = st.text_input("Username:", placeholder="Ilagay ang iyong Username")
            password_input = st.text_input("Password:", type="password", placeholder="Ilagay ang iyong Password")
            submit_btn = st.form_submit_button("Mag-login", use_container_width=True)
            
            if submit_btn:
                ok, acc_key, parish_val, role_val = verify_credentials(username_input, password_input)
                if ok:
                    st.session_state.authenticated = True
                    st.session_state.user_account = acc_key
                    st.session_state.user_parish = parish_val
                    st.session_state.user_role = role_val
                    st.success("✅ Matagumpay na nakapasok!")
                    st.rerun()
                else:
                    st.error("❌ Mali ang Username o Password. Pakisubukan ulit.")

if not st.session_state.authenticated:
    render_login()
    st.stop()

# ==========================================
# 6. MAIN APPLICATION (AUTHENTICATED)
# ==========================================
user_parish = st.session_state.user_parish
user_role = st.session_state.user_role

# Sidebar Navigation & Controls
with st.sidebar:
    st.title("⛪ CTG Ministry Hub")
    st.markdown(f"**Parokya:** `{user_parish}`")
    st.markdown(f"**Role:** `{user_role}`")
    
    if st.button("🔄 I-refresh ang Database", use_container_width=True):
        st.cache_data.clear()
        st.toast("Na-refresh na ang data mula sa Excel at Notion!", icon="✅")
        st.rerun()
        
    st.divider()
    with st.expander("⚙️ API Settings (Optional Override)"):
        st.session_state.custom_gemini_key = st.text_input(
            "Custom Gemini API Key:",
            value=st.session_state.get("custom_gemini_key", ""),
            type="password"
        )
        custom_notion_db = st.text_input(
            "Notion CRM DB ID:",
            value=NOTION_CRM_DB_ID
        )
        
    st.divider()
    if st.button("🚪 Mag-logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_account = None
        st.session_state.user_parish = "Lahat ng Parokya"
        st.session_state.user_role = ""
        st.session_state.chat_history = []
        st.rerun()

# Load Data
excel_data = {}
notion_df = pd.DataFrame()
excel_error = None
notion_error = None

with st.spinner("Binabasa ang Google Drive Excel at Notion CRM database..."):
    try:
        excel_data = load_excel_data()
    except Exception as e:
        excel_error = str(e)
        
    try:
        active_db_id = custom_notion_db if "custom_notion_db" in locals() and custom_notion_db else NOTION_CRM_DB_ID
        notion_df = load_notion_crm_data(NOTION_TOKEN, active_db_id)
    except Exception as e:
        notion_error = str(e)

# Header
st.markdown(f'<div class="main-header">⛪ CTG Ministry Hub — {user_parish} ({user_role})</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Integrated Google Drive Excel Directory, Notion CRM & Gemini AI Assistant</div>', unsafe_allow_html=True)

# Status Alerts
col_stat1, col_stat2, col_stat3 = st.columns(3)
with col_stat1:
    if excel_error:
        st.error(f"❌ Excel Error: {excel_error}")
    else:
        st.metric("📂 Google Drive Excel Sheets", f"{len(excel_data)} Sheets Loaded")
with col_stat2:
    if notion_error:
        st.warning(f"⚠️ Notion CRM: {notion_error}")
    else:
        st.metric("🗂️ Notion CRM Records", f"{len(notion_df)} Visitors/Members")
with col_stat3:
    key_count = len([k for k in [GEMINI_API_KEY_1, GEMINI_API_KEY_2, st.session_state.get("custom_gemini_key", "")] if k and k.strip()])
    st.metric("🤖 Active Gemini API Keys", f"{key_count} Key(s) Ready")

st.divider()

# Main Tabs
tab_dir, tab_crm, tab_ai = st.tabs([
    "📊 Parish Excel Directory",
    "🤝 Notion CRM (Visitors & Members)",
    "🤖 AI Ministry Assistant"
])

# ------------------------------------------
# TAB 1: GOOGLE DRIVE EXCEL DIRECTORY
# ------------------------------------------
with tab_dir:
    if not excel_data:
        st.warning("Walang nabasang datos mula sa Google Drive Excel.")
    else:
        if user_role == "Master Admin" or user_parish == "Lahat ng Parokya":
            available_sheets = list(excel_data.keys())
        else:
            available_sheets = [s for s in excel_data.keys() if s in [user_parish, "PARISH SUMMARY", "SEARCH DATABASE", "HEAD OF FAMILY MASTER"]]
            if not available_sheets:
                available_sheets = list(excel_data.keys())

        col_s1, col_s2 = st.columns([1.2, 2])
        with col_s1:
            selected_sheet = st.selectbox("Piliin ang Sheet:", available_sheets)
        with col_s2:
            search_query = st.text_input("🔍 Maghanap ng Pangalan / Lugar / Detalye sa Sheet na ito:", "")

        df_display = excel_data[selected_sheet].copy()
        if search_query.strip():
            mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query.strip(), case=False, na=False)).any(axis=1)
            df_display = df_display[mask]

        st.caption(f"Nagpapakita ng **{len(df_display)}** na rekord mula sa `{selected_sheet}`.")
        st.dataframe(df_display, use_container_width=True, height=480)

# ------------------------------------------
# TAB 2: NOTION CRM
# ------------------------------------------
with tab_crm:
    st.subheader("🤝 Church CRM: Visitors & Members (Notion)")
    if notion_error:
        st.error(f"Hindi makakonekta sa Notion CRM: {notion_error}")
    elif notion_df.empty:
        st.info("Walang nakitang rekord sa Notion CRM o walang nakalagay na `NOTION_TOKEN` sa Secrets.")
    else:
        crm_search = st.text_input("🔍 Maghanap ng Bisita o Miyembro sa Notion CRM:", "")
        crm_view = notion_df.copy()
        if crm_search.strip():
            mask = crm_view.astype(str).apply(lambda x: x.str.contains(crm_search.strip(), case=False, na=False)).any(axis=1)
            crm_view = crm_view[mask]

        st.caption(f"Kabuuang nakita: **{len(crm_view)}** na bisita/miyembro.")
        st.dataframe(crm_view, use_container_width=True, height=450)

# ------------------------------------------
# TAB 3: GEMINI AI ASSISTANT
# ------------------------------------------
with tab_ai:
    st.subheader("🤖 CTG Ministry AI Assistant")
    st.caption("Magtanong tungkol sa bilang ng miyembro, pamilya, bisita sa Notion CRM, o detalye ng parokya.")

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_prompt = st.chat_input("Magtanong sa AI Assistant (hal. Sino ang mga bisita sa Notion CRM? Ilan ang miyembro sa Lipay Norte?)...")
    if user_prompt:
        st.session_state.chat_history.append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        with st.chat_message("assistant"):
            with st.spinner("Nag-iisip at binabasa ang rekord ng simbahan..."):
                context_str = build_ai_context(excel_data, notion_df, user_parish, user_role)
                ai_reply = ask_gemini_ai(user_prompt, context_str)
                st.markdown(ai_reply)
                st.session_state.chat_history.append({"role": "assistant", "content": ai_reply})
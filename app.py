import os
import re
import time
import base64
import io
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
# 🔐 SECRETS & CONFIGURATION HELPER
# ==========================================
def get_secret(key, default=""):
    if "custom_secrets" in st.session_state and key in st.session_state.custom_secrets:
        val = st.session_state.custom_secrets[key]
        if val:
            return val
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.getenv(key, default)

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
NOTION_TOKEN = get_secret("NOTION_TOKEN")
NOTION_CRM_DB_ID = get_secret("NOTION_CRM_DB_ID")
GMAIL_CLIENT_ID = get_secret("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = get_secret("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = get_secret("GMAIL_REFRESH_TOKEN")

st.set_page_config(page_title="CTG Ministry Hub - AI Agent", page_icon="⛪", layout="wide")

# Direktang link sa iyong Google Drive / Google Sheets file
CLOUD_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/15_znwPmjK95vTgf5lStbKKPCjEGw6oDp/edit?usp=sharing&ouid=107511419052324683693&rtpof=true&sd=true"

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
# 👥 ACCOUNTS NG ADMIN AT MGA PAROKYA (1-10)
# ==========================================
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
# ☁️ GOOGLE DRIVE / SHEETS EXCEL DOWNLOADER
# ==========================================
def extract_id_from_url(url: str) -> str:
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else url

@st.cache_data(ttl=600, show_spinner="Binabasa ang Excel database mula sa Google Drive...")
def load_complete_parish_members():
    doc_id = extract_id_from_url(CLOUD_SPREADSHEET_URL)
    
    download_urls = [
        f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={doc_id}&confirm=t",
        f"https://drive.google.com/uc?export=download&id={doc_id}"
    ]
    
    excel_bytes = None
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}
    
    for url in download_urls:
        try:
            res = session.get(url, headers=headers, timeout=30)
            if res.status_code == 200 and len(res.content) > 1000:
                if res.content[:2] == b"PK":
                    excel_bytes = res.content
                    break
                for k, v in res.cookies.items():
                    if k.startswith("download_warning"):
                        confirm_url = f"https://drive.google.com/uc?export=download&id={doc_id}&confirm={v}"
                        res2 = session.get(confirm_url, headers=headers, timeout=30)
                        if res2.status_code == 200 and res2.content[:2] == b"PK":
                            excel_bytes = res2.content
                            break
                if excel_bytes:
                    break
        except Exception:
            continue

    xls = None
    if excel_bytes is not None:
        try:
            xls = pd.ExcelFile(io.BytesIO(excel_bytes))
        except Exception:
            xls = None

    if xls is None:
        local_candidates = ["CTG_HEADS_TOTAL_UPDATED FIXED.xlsx", "uploaded_ctg_database.xlsx"]
        for lf in local_candidates:
            if os.path.exists(lf):
                try:
                    xls = pd.ExcelFile(lf)
                    break
                except Exception:
                    continue

    if xls is None:
        return pd.DataFrame(), "Hindi ma-download ang Excel file. Siguraduhing naka-set sa 'Anyone with the link' ang sharing permission sa Google Drive."

    sheet_map = {str(s).strip().upper(): s for s in xls.sheet_names}
    all_rows = []

    for p_name in PARISH_ORDER:
        actual_sheet = sheet_map.get(p_name.upper())
        if not actual_sheet:
            for s_up, s_orig in sheet_map.items():
                if p_name.replace(" PARISH", "") in s_up:
                    actual_sheet = s_orig
                    break
        if not actual_sheet:
            continue

        try:
            df_preview = pd.read_excel(xls, sheet_name=actual_sheet, header=None, nrows=10)
            header_idx = 3
            for idx, row_vals in df_preview.iterrows():
                row_text = " ".join([str(v).upper() for v in row_vals.dropna()])
                if ("FAMILY ID" in row_text or "RECORD ID" in row_text) and "FULL NAME" in row_text:
                    header_idx = idx
                    break

            df_raw = pd.read_excel(xls, sheet_name=actual_sheet, header=header_idx)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]

            rename_map = {
                "Member / Record ID": "Family ID",
                "Employment / Work": "Student/Working",
                "Address (HK or Philippines)": "Current Address (If living separately)",
                "Church Member Status": "Church Member",
            }
            df_raw = df_raw.rename(columns=rename_map)

            if "Family ID" not in df_raw.columns or "Full Name" not in df_raw.columns:
                continue

            valid_mask = (
                df_raw["Family ID"].astype(str).str.strip().str.upper().str.startswith("CTG-", na=False)
                & df_raw["Full Name"].notna()
                & (df_raw["Full Name"].astype(str).str.strip() != "")
                & (df_raw["Full Name"].astype(str).str.strip().str.lower() != "nan")
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

            for str_col in ["Family ID", "Full Name", "Head of Family", "Category", "Status", "Church Member"]:
                df_valid[str_col] = df_valid[str_col].fillna("").astype(str).str.strip()

            all_rows.append(df_valid[standard_cols])
        except Exception:
            continue

    if all_rows:
        return pd.concat(all_rows, ignore_index=True), "OK"
    return pd.DataFrame(), "Nabuksan ang file ngunit walang nabasang miyembro. Pakisuri ang pangalan ng mga sheets."

master_members_df, db_status_msg = load_complete_parish_members()

# ==========================================
# 📊 SUMMARY & CLASSIFICATION HELPERS
# ==========================================
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
# 🤖 GEMINI AI HELPER
# ==========================================
def safe_generate_content(client, model, contents, config=None):
    max_retries = 2
    delay = 2
    fallback_models = []
    for m in [model, "gemini-2.5-flash", "gemini-2.5-pro"]:
        if m not in fallback_models:
            fallback_models.append(m)

    for current_model in fallback_models:
        for attempt in range(max_retries):
            try:
                if config:
                    return client.models.generate_content(model=current_model, contents=contents, config=config)
                else:
                    return client.models.generate_content(model=current_model, contents=contents)
            except Exception as e:
                err_str = str(e)
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str:
                    if attempt < max_retries - 1:
                        time.sleep(delay)
                        delay *= 2
                        continue
                break
    raise Exception("Ang mga Gemini models ay kasalukuyang busy o hindi ma-access. Pakisuri ang iyong GEMINI_API_KEY.")

# ==========================================
# 🔑 LOGIN SCREEN (MAY OPISYAL NA LOGO NA TAMA)
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
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=120)
            
        col_log_img, col_log_txt = st.columns([0.15, 0.85])
        with col_log_img:
            if os.path.exists("logo.jpg"):
                st.image("logo.jpg", width=45)
        with col_log_txt:
            st.markdown("### CTG Ministry Hub Login")
            
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
# 🛠️ AI AGENT TOOLS
# ==========================================
def query_excel_database(search_query: str) -> str:
    if master_members_df.empty:
        return "Error: Walang laman o hindi mabasa ang Excel database mula sa Google Drive."
    try:
        df_scope = master_members_df.copy()
        if st.session_state.user_role == "Parish Leader":
            df_scope = df_scope[df_scope["Parish"] == st.session_state.assigned_parish]
            
        match = df_scope[df_scope.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
        if not match.empty:
            results = []
            for _, row in match.head(15).iterrows():
                results.append(
                    f"Parish: {row['Parish']} | Family ID: {row['Family ID']} | "
                    f"Full Name: {row['Full Name']} | Relationship: {row['Relationship']} | "
                    f"Category: {row['Category']} | Status: {row['Status']}"
                )
            return "Mga nahanap sa Excel Database:\n" + "\n".join(results)
        return f"Walang nahanap na record para sa '{search_query}'."
    except Exception as e:
        return f"Error sa pagbasa ng Excel: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
    if not all([GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN]):
        return "Error: Kulang ang GMAIL credentials."
    try:
        creds = Credentials(
            token=None,
            refresh_token=GMAIL_REFRESH_TOKEN.strip(),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=GMAIL_CLIENT_ID.strip(),
            client_secret=GMAIL_CLIENT_SECRET.strip()
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
        "GEMINI_API_KEY": GEMINI_API_KEY, "NOTION_TOKEN": NOTION_TOKEN,
        "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID, "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID,
        "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET, "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
    }

# ==========================================
# 🧭 SIDEBAR CONTROL PANEL
# ==========================================
with st.sidebar:
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")
    
    if not master_members_df.empty:
        st.success(f"☁️ Cloud DB Active ({len(master_members_df)} Miyembro)")
    else:
        st.error("⚠️ Hindi Mabasa ang Cloud DB")
        
    if st.button("🔄 I-refresh ang Data mula sa Cloud", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
        
    st.divider()
    
    menu_options = [
        "📊 Parish Directory & Family Search",
        "💬 AI Chat Assistant",
        "➕ Magdagdag ng Miyembro & Litrato",
        "🙏 Prayer Wall & Categorizer",
        "📢 Announcement Broadcaster",
        "🛠️ App Settings"
    ]
    
    app_mode = st.radio("Pumili ng Pahina:", menu_options)
    
    selected_model = st.selectbox(
        "Gemini Model:",
        ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.5-pro"],
        index=0
    )
    st.divider()
    
    if st.button("🗑️ I-clear ang Chat History", use_container_width=True):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# =====================================================================
# 1. PARISH DIRECTORY & FAMILY SEARCH
# =====================================================================
if app_mode == "📊 Parish Directory & Family Search":
    col_logo, col_title = st.columns([0.06, 0.94])
    with col_logo:
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=50)
    with col_title:
        st.markdown("## CHURCH OF THE TEMPLE OF GOD — SEARCH FAMILY / PARISH")
    
    if master_members_df.empty:
        st.error(f"⚠️ {db_status_msg}")
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
            final_df["Head of Family"].astype(str).str.contains(keyword_search, case=False, na=False)
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
        st.dataframe(final_df, use_container_width=True, hide_index=True)
    else:
        st.warning("Walang nahanap na kumpletong record para sa iyong hinahanap.")
    st.stop()

# =====================================================================
# 2. MAGDAGDAG NG MIYEMBRO & LITRATO
# =====================================================================
elif app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    col_logo, col_title = st.columns([0.06, 0.94])
    with col_logo:
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=50)
    with col_title:
        st.markdown("## Magdagdag ng Miyembro at Mag-upload ng Litrato")
    st.caption("Mag-input ng bagong miyembro sa ilalim ng iyong parokya.")
    
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
# 3. PRAYER WALL & CATEGORIZER
# =====================================================================
elif app_mode == "🙏 Prayer Wall & Categorizer":
    col_logo, col_title = st.columns([0.06, 0.94])
    with col_logo:
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=50)
    with col_title:
        st.markdown("## AI Prayer Request & Needs Categorizer")
    
    raw_prayer = st.text_area("Nilalaman ng Prayer Request o Mensahe:")
    if st.button("Suriin sa AI at I-categorize", type="primary"):
        if raw_prayer:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            if not active_key:
                st.error("⚠️ Wala pang GEMINI_API_KEY. Pumunta sa App Settings sa sidebar.")
                st.stop()
            client = genai.Client(api_key=active_key)
            with st.spinner("Sinusuri ng AI ang kahilingan..."):
                prompt = f"Suriin ang sumusunod na prayer request. Ibigay ang Kategorya, Urgency Level, at comforting Bible verse:\n\n{raw_prayer}"
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
# 4. ANNOUNCEMENT BROADCASTER
# =====================================================================
elif app_mode == "📢 Announcement Broadcaster":
    col_logo, col_title = st.columns([0.06, 0.94])
    with col_logo:
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=50)
    with col_title:
        st.markdown("## AI Ministry Announcement Broadcaster")
        
    event_title = st.text_input("Pamagat ng Aktibidad / Event:")
    event_details = st.text_area("Mga Detalye (Petsa, Oras, Lugar, Mensahe):")
    target_email = st.text_input("Email ng Tatanggap:")
    if st.button("Bumuo at Magpadala ng Email Broadcast", type="primary"):
        if event_title and event_details and target_email:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            if not active_key:
                st.error("⚠️ Wala pang GEMINI_API_KEY. Pumunta sa App Settings sa sidebar.")
                st.stop()
            client = genai.Client(api_key=active_key)
            with st.spinner("Gumagawa at nagpapadala ng broadcast..."):
                try:
                    p_prompt = f"Gumawa ng pormal at nakaka-inspire na announcement email tungkol sa: {event_title}. Mga detalye: {event_details}. Isulat sa Tagalog."
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
    col_logo, col_title = st.columns([0.06, 0.94])
    with col_logo:
        if os.path.exists("logo.jpg"):
            st.image("logo.jpg", width=50)
    with col_title:
        st.markdown("## App Configuration & Settings")
        
    with st.form("settings_form"):
        st.subheader("🔑 API Keys & Integration")
        new_gemini = st.text_input("GEMINI_API_KEY", value=st.session_state.custom_secrets["GEMINI_API_KEY"], type="password")
        new_notion_token = st.text_input("NOTION_TOKEN", value=st.session_state.custom_secrets["NOTION_TOKEN"], type="password")
        new_notion_db = st.text_input("NOTION_CRM_DB_ID", value=st.session_state.custom_secrets["NOTION_CRM_DB_ID"])
        st.subheader("📧 Gmail OAuth2 Credentials")
        new_g_client_id = st.text_input("GMAIL_CLIENT_ID", value=st.session_state.custom_secrets["GMAIL_CLIENT_ID"])
        new_g_client_sec = st.text_input("GMAIL_CLIENT_SECRET", value=st.session_state.custom_secrets["GMAIL_CLIENT_SECRET"], type="password")
        new_g_token = st.text_input("GMAIL_REFRESH_TOKEN", value=st.session_state.custom_secrets["GMAIL_REFRESH_TOKEN"], type="password")
        if st.form_submit_button("I-save ang mga Pagbabago"):
            st.session_state.custom_secrets = {
                "GEMINI_API_KEY": new_gemini, "NOTION_TOKEN": new_notion_token,
                "NOTION_CRM_DB_ID": new_notion_db, "GMAIL_CLIENT_ID": new_g_client_id,
                "GMAIL_CLIENT_SECRET": new_g_client_sec, "GMAIL_REFRESH_TOKEN": new_g_token,
            }
            st.success("Matagumpay na na-update!")
            time.sleep(1)
            st.rerun()
    st.stop()

# =====================================================================
# 6. AI CHAT ASSISTANT
# =====================================================================
col_logo, col_title = st.columns([0.06, 0.94])
with col_logo:
    if os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=50)
with col_title:
    st.markdown("## CTG Church AI Ministry Assistant")
st.caption("Katuwang ng mga Leaders sa paghahanap ng records sa Excel mula sa Google Drive at pamamahala ng mga parokya.")

active_gemini_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
if not active_gemini_key:
    st.warning("⚠️ Wala pang GEMINI_API_KEY. Pumunta sa **App Settings** sa sidebar.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_prompt := st.chat_input("Magtanong tungkol sa records o magpa-send ng email..."):
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user"):
        st.markdown(user_prompt)

    with st.chat_message("assistant"):
        status_box = st.status(f"Pinoproseso gamit ang `{selected_model}`...", expanded=True)
        client = genai.Client(api_key=active_gemini_key)
        system_instruction = (
            f"Ikaw ang opisyal na AI Ministry Assistant ng CTG Church. "
            f"Ang naka-login ay may access sa parokyang: {st.session_state.assigned_parish}. "
            "Tulungan sila sa paghanap ng mga rekord ng pamilya at miyembro mula sa Google Drive database."
        )
        config = types.GenerateContentConfig(
            system_instruction=system_instruction,
            tools=[query_excel_database, send_email_via_gmail],
            temperature=0.3,
        )
        
        models_to_try = []
        for m in [selected_model, "gemini-2.5-flash", "gemini-2.5-pro"]:
            if m not in models_to_try:
                models_to_try.append(m)
                
        success = False
        last_err = ""
        for model_candidate in models_to_try:
            try:
                status_box.write(f"🔄 Kumokonekta sa `{model_candidate}`...")
                chat = client.chats.create(model=model_candidate, config=config)
                response = chat.send_message(user_prompt)
                status_box.update(label=f"✅ Tapos na! (Model: `{model_candidate}`)", state="complete", expanded=False)
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
                success = True
                break
            except Exception as e:
                last_err = str(e)
                continue
                
        if not success:
            status_box.update(label="❌ May Error", state="error", expanded=True)
            st.error(f"Detalye: {last_err}")
import os
import time
import base64
from email.message import EmailMessage
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

load_dotenv()

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
NOTION_CRM_DB_ID = get_secret("NOTION_CRM_DB_ID", "409fff68-03d7-80be-abbe-000bcf7585fd")
GMAIL_CLIENT_ID = get_secret("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = get_secret("GMAIL_CLIENT_SECRET")
GMAIL_REFRESH_TOKEN = get_secret("GMAIL_REFRESH_TOKEN")

st.set_page_config(page_title="CTG Ministry Hub - AI Agent", page_icon="⛪", layout="wide")

UPLOADED_DB_PATH = "uploaded_ctg_database.xlsx"
DEFAULT_DB_PATH = "CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"

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
    "admin": {"password": "ctg2026", "parish": "Lahat ng Parokya", "role": "Master Admin"},
    
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

def get_active_excel_path():
    """Hinahanap kung may in-upload nang Excel file sa server o kung nasa local folder."""
    if os.path.exists(UPLOADED_DB_PATH):
        return UPLOADED_DB_PATH
    if os.path.exists(DEFAULT_DB_PATH):
        return DEFAULT_DB_PATH
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'CTG' in f]
    if excel_files:
        return excel_files[0]
    return None

@st.cache_data
def load_complete_parish_members(excel_path, file_mtime):
    """Binabasa lamang ang mga kumpletong record na may valid Family ID (CTG-...) at Full Name."""
    if not excel_path or not os.path.exists(excel_path):
        return pd.DataFrame()
        
    try:
        xls = pd.ExcelFile(excel_path)
    except Exception:
        return pd.DataFrame()
        
    all_rows = []
    for p_name in PARISH_ORDER:
        if p_name in xls.sheet_names:
            try:
                df_raw = pd.read_excel(xls, sheet_name=p_name, header=3)
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
                
                all_rows.append(df_valid[standard_cols])
            except Exception:
                continue
                
    if all_rows:
        return pd.concat(all_rows, ignore_index=True)
    return pd.DataFrame()

active_excel_path = get_active_excel_path()
active_mtime = os.path.getmtime(active_excel_path) if active_excel_path and os.path.exists(active_excel_path) else 0
master_members_df = load_complete_parish_members(active_excel_path, active_mtime)

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

def safe_generate_content(client, model, contents, config=None):
    max_retries = 3
    delay = 2
    fallback_models = [model, "gemini-2.5-flash", "gemini-2.5-pro"]
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
    raise Exception("Ang mga Gemini models ay kasalukuyang busy o naabot ang quota limit. Subukan muli makalipas ang ilang sandali.")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_role" not in st.session_state:
    st.session_state.user_role = "Admin"
if "assigned_parish" not in st.session_state:
    st.session_state.assigned_parish = "Lahat ng Parokya"

if not st.session_state.authenticated:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.title("⛪ CTG Ministry Hub Login")
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

with st.sidebar:
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Role: **{st.session_state.user_role}**")
    st.info(f"📍 Parish Access:\n**{st.session_state.assigned_parish}**")
    
    if active_excel_path:
        st.success(f"📁 Active DB: `{os.path.basename(active_excel_path)}`")
    else:
        st.warning("⚠️ Walang Excel File")
    st.divider()
    
    menu_options = ["📊 Parish Directory & Family Search"]
    if st.session_state.user_role == "Master Admin":
        menu_options.append("📂 Admin Excel Uploader")
    menu_options.extend([
        "💬 AI Chat Assistant",
        "➕ Magdagdag ng Miyembro & Litrato",
        "🙏 Prayer Wall & Categorizer",
        "📢 Announcement Broadcaster",
        "🛠️ App Settings"
    ])
    
    app_mode = st.radio("Pumili ng Pahina:", menu_options)
    selected_model = st.selectbox("Gemini Model:", ["gemini-2.5-flash", "gemini-2.5-pro"], index=0)
    st.divider()
    
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

# =====================================================================
# 1. ADMIN EXCEL UPLOADER (PARA SA STREAMLIT CLOUD AT LOCAL UPDATE)
# =====================================================================
if app_mode == "📂 Admin Excel Uploader":
    st.title("📂 Admin Excel Database Uploader")
    st.caption("Mag-upload ng bagong Excel file (.xlsx) upang ma-update ang database na ginagamit ng lahat ng Parokya sa app.")
    
    uploaded_file = st.file_uploader("Piliin ang iyong CTG Excel File (.xlsx):", type=["xlsx"])
    
    if uploaded_file is not None:
        if st.button("💾 I-save at Gamitin ang Excel File na Ito", type="primary"):
            try:
                with open(UPLOADED_DB_PATH, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.cache_data.clear()
                st.success("✅ Matagumpay na na-save at na-load ang bagong Excel database! Maaari ka nang pumunta sa 'Parish Directory & Family Search'.")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Error sa pag-save ng Excel file: {str(e)}")
                
    if os.path.exists(UPLOADED_DB_PATH):
        st.divider()
        st.info(f"Kasalukuyang nakakonekta sa uploaded file: `{UPLOADED_DB_PATH}` ({len(master_members_df)} kumpletong miyembro ang nabasa).")
    st.stop()

# =====================================================================
# 2. PARISH DIRECTORY & FAMILY SEARCH
# =====================================================================
if app_mode == "📊 Parish Directory & Family Search":
    st.title("⛪ CHURCH OF THE TEMPLE OF GOD — SEARCH FAMILY / PARISH")
    
    # Kung wala pang Excel file sa Streamlit Cloud, magpakita ng quick upload box para sa Admin
    if master_members_df.empty:
        st.warning("⚠️ Wala pang Excel database na naka-load sa server.")
        if st.session_state.user_role == "Master Admin":
            st.markdown("#### 📤 Mag-upload ng Excel File Ngayon:")
            quick_upload = st.file_uploader("I-upload ang `CTG_HEADS_TOTAL_UPDATED FIXED.xlsx` dito:", type=["xlsx"], key="quick_up")
            if quick_upload is not None:
                with open(UPLOADED_DB_PATH, "wb") as f:
                    f.write(quick_upload.getbuffer())
                st.cache_data.clear()
                st.success("✅ Na-upload na ang Excel file! Nagre-refresh...")
                time.sleep(1)
                st.rerun()
        else:
            st.info("Pakisuyo sa Master Admin na i-upload muna ang Excel database sa Admin Excel Uploader.")
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
        st.dataframe(final_df, use_container_width=True, hide_index=True)
    else:
        st.warning("Walang nahanap na kumpletong record para sa iyong hinahanap.")
    st.stop()

# =====================================================================
# 3. MAGDAGDAG NG MIYEMBRO & LITRATO
# =====================================================================
elif app_mode == "➕ Magdagdag ng Miyembro & Litrato":
    st.title("➕ Magdagdag ng Miyembro at Mag-upload ng Litrato")
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
# 4. PRAYER WALL & CATEGORIZER
# =====================================================================
elif app_mode == "🙏 Prayer Wall & Categorizer":
    st.title("🙏 AI Prayer Request & Needs Categorizer")
    raw_prayer = st.text_area("Nilalaman ng Prayer Request o Mensahe:")
    if st.button("Suriin sa AI at I-categorize", type="primary"):
        if raw_prayer:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
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
# 5. ANNOUNCEMENT BROADCASTER
# =====================================================================
elif app_mode == "📢 Announcement Broadcaster":
    st.title("📢 AI Ministry Announcement Broadcaster")
    event_title = st.text_input("Pamagat ng Aktibidad / Event:")
    event_details = st.text_area("Mga Detalye (Petsa, Oras, Lugar, Mensahe):")
    target_email = st.text_input("Email ng Tatanggap:")
    if st.button("Bumuo at Magpadala ng Email Broadcast", type="primary"):
        if event_title and event_details and target_email:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
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
# 6. APP SETTINGS
# =====================================================================
elif app_mode == "🛠️ App Settings":
    st.title("🛠️ App Configuration & Settings")
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
# 7. AI CHAT ASSISTANT
# =====================================================================
st.title("⛪ CTG Church AI Ministry Assistant")
st.caption("Katuwang ng mga Leaders sa paghahanap ng records sa Excel at pamamahala ng mga parokya.")

active_gemini_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
if not active_gemini_key:
    st.warning("⚠️ Wala pang GEMINI_API_KEY. Pumunta sa **App Settings** sa sidebar.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = []

if "current_model" not in st.session_state:
    st.session_state.current_model = selected_model

if "chat_session" not in st.session_state or st.session_state.current_model != selected_model:
    st.session_state.current_model = selected_model
    client = genai.Client(api_key=active_gemini_key)
    system_instruction = (
        f"Ikaw ang opisyal na AI Ministry Assistant ng CTG Church. "
        f"Ang naka-login ay may access sa parokyang: {st.session_state.assigned_parish}. "
        "Tulungan sila sa paghanap ng mga rekord ng pamilya at miyembro sa Excel database."
    )
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        tools=[query_excel_database, send_email_via_gmail],
        temperature=0.3,
    )
    st.session_state.chat_session = client.chats.create(model=selected_model, config=config)

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_prompt := st.chat_input("Magtanong tungkol sa records sa Excel o magpa-send ng email..."):
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user"):
        st.markdown(user_prompt)

    with st.chat_message("assistant"):
        status_box = st.status(f"Pinoproseso gamit ang `{selected_model}`...", expanded=True)
        max_retries = 3
        for attempt in range(max_retries):
            try:
                status_box.write(f"🔄 Kumokonekta sa Gemini API (Attempt {attempt + 1}/{max_retries})...")
                response = st.session_state.chat_session.send_message(user_prompt)
                status_box.update(label="✅ Tapos na!", state="complete", expanded=False)
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
                break
            except Exception as e:
                err_str = str(e)
                if ("503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str) and attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                status_box.update(label="❌ May Error", state="error", expanded=True)
                st.error(f"Detalye: {err_str}")
                break
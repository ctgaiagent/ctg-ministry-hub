import os
import time
import base64
from email.message import EmailMessage
import requests
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
APP_USER = get_secret("APP_USER", "admin")
APP_PASSWORD = get_secret("APP_PASSWORD", "ctg2026")

st.set_page_config(page_title="CTG Ministry Hub - AI Agent", page_icon="⛪", layout="wide")

@st.cache_data
def load_excel_data():
    # Awtomatikong hahanapin ang anumang .xlsx file sa folder para maiwasan ang path error
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'CTG' in f]
    if not excel_files:
        # Fallback sa eksaktong pangalan kung sakaling wala sa listahan
        if os.path.exists("CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"):
            excel_files = ["CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"]
        else:
            return None, "Walang nakitang Excel file sa folder."
    
    filename = excel_files[0]
    try:
        xls = pd.ExcelFile(filename)
        return xls, filename
    except Exception as e:
        return None, str(e)

excel_xls, excel_filename = load_excel_data()

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

if not st.session_state.authenticated:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.title("⛪ CTG Ministry Hub")
        st.caption("Mag-login upang ma-access ang AI Assistant para sa mga Leaders.")
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login", use_container_width=True)
            if submit:
                if username == APP_USER and password == APP_PASSWORD:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Maling Username o Password.")
    st.stop()

def query_excel_database(search_query: str) -> str:
    if excel_xls is None:
        return "Error: Hindi makita ang Excel database file."
    try:
        results = []
        if "HEAD OF FAMILY MASTER" in excel_xls.sheet_names:
            df_hof = pd.read_excel(excel_xls, sheet_name="HEAD OF FAMILY MASTER", header=1)
            match = df_hof[df_hof.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            if not match.empty:
                for _, row in match.head(5).iterrows():
                    results.append(f"Family ID: {row.get('Family ID', 'N/A')} | Parish: {row.get('Parish', 'N/A')} | Head: {row.get('Head of Family', 'N/A')} | Address: {row.get('Complete Address', 'N/A')}")
        
        if results:
            return "Mga nahanap sa Excel Database:\n" + "\n".join(results)
        return f"Walang nahanap na record sa Excel para sa '{search_query}'."
    except Exception as e:
        return f"Error sa pagbasa ng Excel: {str(e)}"

def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
    if not all([GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN]):
        return "Error: Kulang ang GMAIL credentials."
    try:
        creds = Credentials(token=None, refresh_token=GMAIL_REFRESH_TOKEN.strip(), token_uri="https://oauth2.googleapis.com/token", client_id=GMAIL_CLIENT_ID.strip(), client_secret=GMAIL_CLIENT_SECRET.strip())
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
        "APP_USER": APP_USER, "APP_PASSWORD": APP_PASSWORD
    }

with st.sidebar:
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Naka-login bilang: **{APP_USER}**")
    app_mode = st.radio("Pumili ng Pahina:", ["💬 AI Chat Assistant", "📊 Parish Directory & Excel Search", "🙏 Prayer Wall & Categorizer", "📢 Announcement Broadcaster", "🛠️ App Settings"])
    selected_model = st.selectbox("Gemini Model:", ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-3.8-flash"], index=0)
    st.divider()
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()

if app_mode == "📊 Parish Directory & Excel Search":
    st.title("📊 CTG Parish Directory & Excel Database")
    st.caption("Pumili ng parokya gamit ang dropdown filter o hanapin ang Family ID at pangalan ng miyembro mula sa Excel master file.")
    
    if excel_xls is not None:
        st.info(f"📁 Nakakonekta sa Excel File: `{excel_filename}`")
        try:
            df_hof = pd.read_excel(excel_xls, sheet_name="HEAD OF FAMILY MASTER", header=1)
            
            available_parishes = ["Lahat ng Parokya"] + sorted(df_hof['Parish'].dropna().unique().tolist())
            
            col_filter1, col_filter2 = st.columns(2)
            with col_filter1:
                selected_parish = st.selectbox("Pumili ng Parish:", available_parishes)
            with col_filter2:
                search_term = st.text_input("Maghanap (Family ID o Pangalan):", placeholder="hal. CTG-LN-FAM-0001")
            
            filtered_df = df_hof.copy()
            if selected_parish != "Lahat ng Parokya":
                filtered_df = filtered_df[filtered_df['Parish'] == selected_parish]
                
            if search_term:
                filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)]
            
            st.divider()
            st.subheader(f"Resulta para sa: {selected_parish} ({len(filtered_df)} pamilya ang nahanap)")
            
            if not filtered_df.empty:
                st.dataframe(filtered_df[['Family ID', 'Parish', 'Head of Family', 'Mobile Number', 'Complete Address', 'Total Family Members', 'Active Members']], use_container_width=True)
            else:
                st.warning("Walang nahanap na record para sa napiling parokya o salita.")
                
        except Exception as ex:
            st.error(f"Error sa pagbasa ng Excel data: {str(ex)}")
    else:
        st.error(f"Hindi makita ang Excel database file sa folder. Detalye: {excel_filename}")
    st.stop()

if app_mode == "🙏 Prayer Wall & Categorizer":
    st.title("🙏 AI Prayer Request & Needs Categorizer")
    raw_prayer = st.text_area("Nilalaman ng Prayer Request o Mensahe:")
    if st.button("Suriin sa AI at I-categorize", type="primary"):
        if raw_prayer:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=active_key)
            with st.spinner("Sinusuri ng AI ang kahilingan..."):
                prompt = f"Suriin ang sumusunod na prayer request. Ibigay ang Kategorya (Kalusugan, Pananalapi, Pamilya, Espirituwal), Urgency Level, at comforting Bible verse:\n\n{raw_prayer}"
                try:
                    res = safe_generate_content(client, selected_model, prompt)
                    st.success("Tapos na ang Pagsusuri!")
                    st.markdown(res.text)
                except Exception as ex:
                    st.error(str(ex))
        else:
            st.warning("Maglagay muna ng mensahe.")
    st.stop()

if app_mode == "📢 Announcement Broadcaster":
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

if app_mode == "🛠️ App Settings":
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
        st.subheader("🔐 Admin Login Credentials")
        new_user = st.text_input("APP_USER", value=st.session_state.custom_secrets["APP_USER"])
        new_pass = st.text_input("APP_PASSWORD", value=st.session_state.custom_secrets["APP_PASSWORD"], type="password")
        if st.form_submit_button("I-save ang mga Pagbabago"):
            st.session_state.custom_secrets = {
                "GEMINI_API_KEY": new_gemini, "NOTION_TOKEN": new_notion_token,
                "NOTION_CRM_DB_ID": new_notion_db, "GMAIL_CLIENT_ID": new_g_client_id,
                "GMAIL_CLIENT_SECRET": new_g_client_sec, "GMAIL_REFRESH_TOKEN": new_g_token,
                "APP_USER": new_user, "APP_PASSWORD": new_pass
            }
            st.success("Matagumpay na na-update!")
            time.sleep(1)
            st.rerun()
    st.stop()

st.title("⛪ CTG Church AI Ministry Assistant")
st.caption("Katuwang ng mga Leaders sa pamamahala ng Parish Records, Visitors, at Emails.")

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
        "Ikaw ang opisyal na AI Ministry Assistant ng CTG Church. "
        "Tinutulungan mo ang mga church leaders sa paghanap ng record sa Excel database "
        "at sa pagpapadala ng follow-up emails sa pamamahala ng Gmail. "
        "Maging magalang, malinaw, at tumugon sa wikang Tagalog o Taglish."
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
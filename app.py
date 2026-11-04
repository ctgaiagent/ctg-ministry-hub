import os
import time
import base64
from email.message import EmailMessage
import requests
import streamlit as st
from dotenv import load_dotenv
from google import genai
from google.genai import types
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# 1. Load Environment Variables (Local .env)
load_dotenv()

def get_secret(key, default=""):
    """Kumukuha ng secret mula sa Session State (Settings), Streamlit Cloud Secrets, o Local .env"""
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

# 2. Page Configuration
st.set_page_config(
    page_title="CTG Ministry Hub - AI Agent",
    page_icon="⛪",
    layout="wide"
)

# 3. Authentication / Login Gate (Compact view)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.title("⛪ CTG Ministry Hub")
        st.caption("Mag-login upang ma-access ang AI Assistant para sa mga Leaders at Priests.")
        
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

# ==========================================
# 4. AI AGENT TOOLS (NOTION AT GMAIL)
# ==========================================

def get_notion_crm_data(search_name: str = "") -> str:
    """Kumukuha ng listahan ng mga visitors at members mula sa Notion CRM database."""
    if not NOTION_TOKEN or not NOTION_CRM_DB_ID:
        return "Error: Kulang ang NOTION_TOKEN o NOTION_CRM_DB_ID sa configuration."

    clean_db_id = NOTION_CRM_DB_ID.strip().replace("-", "")
    url = f"https://api.notion.com/v1/databases/{clean_db_id}/query"
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN.strip()}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }

    try:
        response = requests.post(url, headers=headers, json={"page_size": 50}, timeout=15)
        if response.status_code == 404:
            return "Error 404: Hindi makita ang Notion database. I-check ang Integration connection sa Notion page."
        if response.status_code != 200:
            return f"Notion API Error ({response.status_code}): {response.text}"

        results = response.json().get("results", [])
        if not results:
            return "Walang nakitang records sa Notion database."

        records = []
        for page in results:
            props = page.get("properties", {})
            row_info = []
            for prop_name, prop_val in props.items():
                ptype = prop_val.get("type")
                val_str = ""
                if ptype == "title":
                    val_str = "".join([t.get("plain_text", "") for t in prop_val.get("title", [])])
                elif ptype == "rich_text":
                    val_str = "".join([t.get("plain_text", "") for t in prop_val.get("rich_text", [])])
                elif ptype == "email":
                    val_str = prop_val.get("email") or ""
                elif ptype == "phone_number":
                    val_str = prop_val.get("phone_number") or ""
                elif ptype == "select":
                    sel = prop_val.get("select")
                    val_str = sel.get("name", "") if sel else ""
                elif ptype == "status":
                    stat = prop_val.get("status")
                    val_str = stat.get("name", "") if stat else ""

                if val_str:
                    row_info.append(f"{prop_name}: {val_str}")

            line = " | ".join(row_info)
            if line:
                if not search_name or search_name.lower() in line.lower():
                    records.append(f"- {line}")

        return "Mga Rekord mula sa Notion CRM:\n" + "\n".join(records) if records else f"Walang nahanap na record para sa '{search_name}'."
    except Exception as e:
        return f"System Error sa Notion: {str(e)}"


def add_visitor_to_notion(name: str, email: str = "", status: str = "New Visitor", notes: str = "") -> str:
    """Nagdadagdag ng bagong visitor o member record sa Notion CRM database."""
    if not NOTION_TOKEN or not NOTION_CRM_DB_ID:
        return "Error: Kulang ang NOTION credentials."

    clean_db_id = NOTION_CRM_DB_ID.strip().replace("-", "")
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN.strip()}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }

    db_resp = requests.get(f"https://api.notion.com/v1/databases/{clean_db_id}", headers=headers, timeout=15)
    if db_resp.status_code != 200:
        return f"Hindi mabasa ang Notion database schema ({db_resp.status_code})."

    db_props = db_resp.json().get("properties", {})
    new_props = {}

    for prop_name, prop_meta in db_props.items():
        ptype = prop_meta.get("type")
        if ptype == "title":
            new_props[prop_name] = {"title": [{"text": {"content": name}}]}
        elif ptype == "email" and email:
            new_props[prop_name] = {"email": email}
        elif ptype == "rich_text" and notes:
            new_props[prop_name] = {"rich_text": [{"text": {"content": notes}}]}

    payload = {"parent": {"database_id": clean_db_id}, "properties": new_props}
    create_resp = requests.post("https://api.notion.com/v1/pages", headers=headers, json=payload, timeout=15)
    
    if create_resp.status_code in [200, 201]:
        return f"Matagumpay na naidagdag si '{name}' sa Notion CRM!"
    return f"Nabigong idagdag sa Notion: {create_resp.text}"


def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
    """Nagpapadala ng email gamit ang Gmail REST API at OAuth2 Refresh Token."""
    if not all([GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN]):
        return "Error: Kulang ang GMAIL credentials sa configuration."

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
        send_result = service.users().messages().send(userId="me", body={"raw": encoded_msg}).execute()
        return f"Matagumpay na naipadala ang email kay {recipient_email}! (ID: {send_result.get('id')})"
    except Exception as e:
        return f"Error sa pag-send ng Gmail: {str(e)}"


# ==========================================
# 5. SIDEBAR & NAVIGATION (MENU)
# ==========================================
if "custom_secrets" not in st.session_state:
    st.session_state.custom_secrets = {
        "GEMINI_API_KEY": GEMINI_API_KEY,
        "NOTION_TOKEN": NOTION_TOKEN,
        "NOTION_CRM_DB_ID": NOTION_CRM_DB_ID,
        "GMAIL_CLIENT_ID": GMAIL_CLIENT_ID,
        "GMAIL_CLIENT_SECRET": GMAIL_CLIENT_SECRET,
        "GMAIL_REFRESH_TOKEN": GMAIL_REFRESH_TOKEN,
        "APP_USER": APP_USER,
        "APP_PASSWORD": APP_PASSWORD
    }

with st.sidebar:
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Naka-login bilang: **{APP_USER}**")
    
    app_mode = st.radio("Pumili ng Pahina:", ["💬 AI Chat Assistant", "🙏 Prayer Wall & Categorizer", "📢 Announcement Broadcaster", "🛠️ App Settings"])
    
    # Naidagdag na ang gemini-3.8-flash sa listahan
    selected_model = st.selectbox(
        "Gemini Model:",
        ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-2.5-pro"],
        index=0
    )
    
    st.divider()
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()

    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()


# ==========================================
# 6. FEATURE: PRAYER REQUEST CATEGORIZER
# ==========================================
if app_mode == "🙏 Prayer Wall & Categorizer":
    st.title("🙏 AI Prayer Request & Needs Categorizer")
    st.caption("I-paste dito ang mga prayer requests o pangangailangan para awtomatikong ma-classify ng AI.")
    
    raw_prayer = st.text_area("Nilalaman ng Prayer Request o Mensahe:")
    
    if st.button("Suriin at I-categorize", type="primary"):
        if raw_prayer:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=active_key)
            with st.spinner("Sinusuri ng AI ang kahilingan..."):
                prompt = f"Suriin ang sumusunod na prayer request. Ibigay ang Kategorya (Kalusugan, Pananalapi, Pamilya, Espirituwal, o Iba pa), Urgency Level, at comforting Bible verse:\n\n{raw_prayer}"
                res = client.models.generate_content(model=selected_model, contents=prompt)
                st.success("Tapos na ang Pagsusuri!")
                st.markdown(res.text)
        else:
            st.warning("Maglagay muna ng mensahe.")
    st.stop()


# ==========================================
# 7. FEATURE: ANNOUNCEMENT BROADCASTER
# ==========================================
if app_mode == "📢 Announcement Broadcaster":
    st.title("📢 AI Ministry Announcement Broadcaster")
    st.caption("Gumawa ng pormal na announcement para sa mga miyembro at ipadala via email.")
    
    event_title = st.text_input("Pamagat ng Aktibidad / Event:")
    event_details = st.text_area("Mga Detalye (Petsa, Oras, Lugar, Mensahe):")
    target_email = st.text_input("Email ng Tatanggap:")
    
    if st.button("Bumuo at Magpadala ng Email Broadcast", type="primary"):
        if event_title and event_details and target_email:
            active_key = st.session_state.custom_secrets["GEMINI_API_KEY"]
            client = genai.Client(api_key=active_key)
            with st.spinner("Gumagawa at nagpapadala ng broadcast..."):
                p_prompt = f"Gumawa ng pormal, nakaka-inspire, at malinaw na simbahan announcement email tungkol sa: {event_title}. Mga detalye: {event_details}. Isulat sa Tagalog."
                ai_res = client.models.generate_content(model=selected_model, contents=p_prompt)
                email_body = ai_res.text
                
                res_send = send_email_via_gmail(target_email, f"CTG Announcement: {event_title}", email_body)
                st.success(res_send)
                with st.expander("Tingnan ang Mensahe"):
                    st.markdown(email_body)
        else:
            st.warning("Punan ang lahat ng kailangan.")
    st.stop()


# ==========================================
# 8. APP SETTINGS PAGE
# ==========================================
if app_mode == "🛠️ App Settings":
    st.title("🛠️ App Configuration & Settings")
    st.markdown("Dito mo maaaring i-update ang iyong mga API keys at password.")
    
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
                "GEMINI_API_KEY": new_gemini,
                "NOTION_TOKEN": new_notion_token,
                "NOTION_CRM_DB_ID": new_notion_db,
                "GMAIL_CLIENT_ID": new_g_client_id,
                "GMAIL_CLIENT_SECRET": new_g_client_sec,
                "GMAIL_REFRESH_TOKEN": new_g_token,
                "APP_USER": new_user,
                "APP_PASSWORD": new_pass
            }
            st.success("Matagumpay na na-update!")
            time.sleep(1)
            st.rerun()
    st.stop()


# ==========================================
# 9. MAIN AI CHAT INTERFACE
# ==========================================
st.title("⛪ CTG Church AI Ministry Assistant")
st.caption("Katuwang ng mga Leaders at Priests sa pamamahala ng Visitors CRM, Follow-up Emails, at Iskedyul.")

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
        "Tinutulungan mo ang mga church leaders sa pag-check ng Notion CRM database "
        "at sa pagpapadala ng follow-up emails sa pamamagitan ng Gmail. "
        "Maging magalang, malinaw, at tumugon sa wikang Tagalog o Taglish."
    )
    
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        tools=[get_notion_crm_data, add_visitor_to_notion, send_email_via_gmail],
        temperature=0.3,
    )
    st.session_state.chat_session = client.chats.create(model=selected_model, config=config)

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_prompt := st.chat_input("Magtanong tungkol sa visitors sa Notion o magpa-send ng follow-up email..."):
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
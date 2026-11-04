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
    """Kumukuha ng secret mula sa Streamlit Cloud Secrets o sa Local .env"""
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

# 3. Authentication / Login Gate
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("⛪ CTG Ministry Hub Login")
    st.markdown("Maligayang pagdating! Mag-login upang ma-access ang AI Assistant para sa mga Leaders at Priests.")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")
        
        if submit:
            if username == APP_USER and password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Maling Username o Password. Pakisubukan ulit.")
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
            return (
                "Error 404 (object_not_found): Hindi makita ng Notion API ang database. "
                "SOLUSYON: Buksan ang Notion database page -> i-click ang '...' sa kanang itaas -> "
                "piliin ang 'Connections' -> i-add ang iyong integration (hal. CTG AI AGENT)."
            )
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
                elif ptype == "date":
                    dt = prop_val.get("date")
                    val_str = dt.get("start", "") if dt else ""

                if val_str:
                    row_info.append(f"{prop_name}: {val_str}")

            line = " | ".join(row_info)
            if line:
                if not search_name or search_name.lower() in line.lower():
                    records.append(f"- {line}")

        if not records:
            return f"Walang nahanap na record na tumutugma sa '{search_name}'."
        return "Mga Rekord mula sa Notion CRM:\n" + "\n".join(records)

    except Exception as e:
        return f"System Error sa pagkonekta sa Notion: {str(e)}"


def add_visitor_to_notion(name: str, email: str = "", status: str = "New Visitor", notes: str = "") -> str:
    """Nagdadagdag ng bagong visitor o member record sa Notion CRM database."""
    if not NOTION_TOKEN or not NOTION_CRM_DB_ID:
        return "Error: Kulang ang NOTION_TOKEN o NOTION_CRM_DB_ID."

    clean_db_id = NOTION_CRM_DB_ID.strip().replace("-", "")
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN.strip()}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }

    # Kunin muna ang schema ng database para malaman ang tamang pangalan ng columns
    db_resp = requests.get(f"https://api.notion.com/v1/databases/{clean_db_id}", headers=headers, timeout=15)
    if db_resp.status_code != 200:
        return f"Hindi mabasa ang Notion database schema ({db_resp.status_code}): {db_resp.text}"

    db_props = db_resp.json().get("properties", {})
    new_props = {}

    for prop_name, prop_meta in db_props.items():
        ptype = prop_meta.get("type")
        if ptype == "title":
            new_props[prop_name] = {"title": [{"text": {"content": name}}]}
        elif ptype == "email" and email:
            new_props[prop_name] = {"email": email}
        elif ptype == "rich_text" and notes and ("note" in prop_name.lower() or "remark" in prop_name.lower()):
            new_props[prop_name] = {"rich_text": [{"text": {"content": notes}}]}

    payload = {
        "parent": {"database_id": clean_db_id},
        "properties": new_props
    }

    create_resp = requests.post("https://api.notion.com/v1/pages", headers=headers, json=payload, timeout=15)
    if create_resp.status_code in [200, 201]:
        return f"Matagumpay na naidagdag si '{name}' sa Notion CRM!"
    return f"Nabigong idagdag sa Notion ({create_resp.status_code}): {create_resp.text}"


def send_email_via_gmail(recipient_email: str, subject: str, body_text: str) -> str:
    """Nagpapadala ng email gamit ang Gmail REST API at OAuth2 Refresh Token."""
    if not all([GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN]):
        return "Error: Kulang ang GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, o GMAIL_REFRESH_TOKEN sa configuration."

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
        return f"Matagumpay na naipadala ang email kay {recipient_email}! (Message ID: {send_result.get('id')})"

    except Exception as e:
        err_str = str(e)
        if "403" in err_str or "insufficientPermissions" in err_str or "Insufficient Permission" in err_str:
            return (
                "Error 403 (Insufficient Permission sa Gmail): Ang iyong GMAIL_REFRESH_TOKEN ay ginawa nang walang "
                "pahintulot na mag-send ng email. Kailangan mong gumawa ng bagong Refresh Token na may scope na: "
                "'https://www.googleapis.com/auth/gmail.send' o 'https://mail.google.com/'."
            )
        return f"Error sa pag-send ng Gmail: {err_str}"


# ==========================================
# 5. SIDEBAR CONTROLS & DIAGNOSTICS
# ==========================================
with st.sidebar:
    st.header("⚙️ Ministry Control Panel")
    st.write(f"👤 Naka-login bilang: **{APP_USER}**")
    
    selected_model = st.selectbox(
        "Piliin ang Gemini Model:",
        ["gemini-2.5-flash", "gemini-2.5-pro"],
        index=0
    )
    
    st.divider()
    st.subheader("🔍 Tool Diagnostics (Subukan Dito)")
    
    if st.button("1. I-test ang Notion Connection"):
        with st.spinner("Kinokonekta ang Notion..."):
            res = get_notion_crm_data()
            if "Error" in res:
                st.error(res)
            else:
                st.success("Gumagana ang Notion!")
                with st.expander("Tingnan ang Data"):
                    st.text(res)

    with st.expander("2. I-test ang Pag-send ng Gmail"):
        test_to = st.text_input("Padadalhan (Email):", placeholder="halimbawa@gmail.com")
        if st.button("Magpadala ng Test Email"):
            if test_to:
                with st.spinner("Nagpapadala ng email..."):
                    gres = send_email_via_gmail(
                        test_to,
                        "CTG Ministry Hub - Test Email",
                        "Pagbati! Ito ay test email mula sa iyong CTG Church AI App."
                    )
                    if "Error" in gres:
                        st.error(gres)
                    else:
                        st.success(gres)
            else:
                st.warning("Maglagay muna ng email address.")

    st.divider()
    if st.button("🗑️ I-clear ang Chat History"):
        st.session_state.messages = []
        st.session_state.pop("chat_session", None)
        st.rerun()

    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()


# ==========================================
# 6. MAIN AI CHAT INTERFACE
# ==========================================
st.title("⛪ CTG Church AI Ministry Assistant")
st.caption("Katuwang ng mga Leaders at Priests sa pamamahala ng Visitors CRM, Follow-up Emails, at Iskedyul.")

if not GEMINI_API_KEY:
    st.warning("⚠️ Wala pang GEMINI_API_KEY. Pakilagay ito sa `.env` file o sa Streamlit Cloud Secrets.")
    st.stop()

if "messages" not in st.session_state:
    st.session_state.messages = []

if "current_model" not in st.session_state:
    st.session_state.current_model = selected_model

# Initialize Gemini Client at Chat Session na may kasamang Tools
if "chat_session" not in st.session_state or st.session_state.current_model != selected_model:
    st.session_state.current_model = selected_model
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    system_instruction = (
        "Ikaw ang opisyal na AI Ministry Assistant ng CTG Church. "
        "Tinutulungan mo ang mga church leaders at mga pari sa pag-check ng Notion CRM database "
        "(gamit ang `get_notion_crm_data` at `add_visitor_to_notion`) at sa pagpapadala ng follow-up emails "
        "sa pamamagitan ng Gmail (gamit ang `send_email_via_gmail`). "
        "Maging magalang, malinaw, at tumugon sa wikang Tagalog o Taglish."
    )
    
    config = types.GenerateContentConfig(
        system_instruction=system_instruction,
        tools=[get_notion_crm_data, add_visitor_to_notion, send_email_via_gmail],
        temperature=0.3,
    )
    st.session_state.chat_session = client.chats.create(model=selected_model, config=config)

# Ipakita ang nakaraang pag-uusap
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Chat Input Box
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
                    wait_sec = 3
                    status_box.write(f"⏳ Busy ang server. Uulitin sa loob ng {wait_sec} segundo...")
                    time.sleep(wait_sec)
                    continue
                status_box.update(label="❌ May Error sa Pagkonekta", state="error", expanded=True)
                st.error(f"Detalye ng Error: {err_str}")
                break
import streamlit as st
import os, json, re, requests, time, tempfile
from datetime import datetime, timedelta, date
from lxml import etree
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from msal import ConfidentialClientApplication
from openai import AzureOpenAI
import PyPDF2
import docx
import pandas as pd
from PIL import Image

# =============== CONFIG ===============
def get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except:
        return default

CLIENT_ID = get_secret("CLIENT_ID")
CLIENT_SECRET = get_secret("CLIENT_SECRET")
TENANT_ID = get_secret("TENANT_ID")
REDIRECT_URI = get_secret("REDIRECT_URI", "http://localhost:8501")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else ""
SCOPE = ["https://graph.microsoft.com/User.Read"]

API = "https://api.ted.europa.eu/v3/notices/search"
COPILOT_STUDIO_ENDPOINT = get_secret("COPILOT_STUDIO_ENDPOINT", "")

JKM_LOGO_URL = "https://www.xing.com/imagecache/public/scaled_original_image/eyJ1dWlkIjoiMGE2MTk2MTYtODI4Zi00MWZlLWEzN2ItMjczZGM2ODc5MGJmIiwiYXBwX2NvbnRleHQiOiJlbnRpdHktcGFnZXMiLCJtYXhfd2lkdGgiOjMyMCwibWF4X2hlaWdodCI6MzIwfQ?signature=a21e5c1393125a94fc9765898c25d73a064665dc3aacf872667c902d7ed9c3f9"
BOT_AVATAR_URL = "https://raw.githubusercontent.com/PratikSondkarJKM/AkquiseWescraper/refs/heads/main/botavatar.svg"

# =============== AUTH ===============
def build_msal_app():
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        st.error("❌ Missing OAuth config")
        st.stop()
    return ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

def fetch_token(auth_code):
    msal_app = build_msal_app()
    return msal_app.acquire_token_by_authorization_code(auth_code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

def login_button():
    msal_app = build_msal_app()
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    st.markdown(f"""
    <style>
    .block-container {{ padding: 0 !important; max-width: 100vw !important; }}
    .center-root {{
        min-height: 100vh; width: 100vw;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        background: linear-gradient(120deg, #eaf6fb 0%, #f3e9f5 100%);
    }}
    .jkm-logo {{ height: 72px; margin-bottom: 14px; border-radius: 14px; }}
    .app-title {{ font-size: 2.3em; text-align: center; font-weight: 800; color: #283044; }}
    .login-card {{
        width: 375px; padding: 38px 34px; background: #fff; border-radius: 18px;
        box-shadow: 0 8px 32px rgba(50,72,140,.13);
    }}
    .login-button {{
        display: block; width: 100%; padding: 17px 0; background-color: #0078d7; 
        color: #fff !important; border: none; border-radius: 7px; cursor: pointer; 
        font-weight: 600; text-decoration: none;
    }}
    .login-button:hover {{ background-color: #005fa1; }}
    </style>
    <div class="center-root">
        <img src="{JKM_LOGO_URL}" class="jkm-logo"/>
        <div class="app-title">TED Scraper & AI Assistant</div>
        <div class="login-card">
            <h2 style="text-align: center;">Sign in with Microsoft</h2>
            <a href="{auth_url}" class="login-button">Sign in</a>
        </div>
    </div>
    """, unsafe_allow_html=True)

def auth_flow():
    params = st.query_params
    if "code" in params and "user_token" not in st.session_state:
        code = params["code"][0] if isinstance(params["code"], list) else params["code"]
        token_data = fetch_token(code)
        if "access_token" in token_data:
            st.session_state["user_token"] = token_data["access_token"]
            st.query_params.clear()
            st.rerun()
        else:
            st.error("Login failed")
            st.stop()
    if "user_token" not in st.session_state:
        login_button()
        st.stop()

class CopilotStudioClient:
    """Uses the user's Microsoft token to connect to Copilot Studio agent"""
    def __init__(self, endpoint_url, user_token):
        self.base_endpoint = endpoint_url
        self.user_token = user_token
        self.conversation_id = None
        self.watermark = None
        self.debug_log = []
        self.last_error = None
        
    def log(self, msg):
        self.debug_log.append(msg)
        print(f"[DEBUG] {msg}")  # Also print to console
        
    def get_debug(self):
        if not self.debug_log:
            return "\n\n**Debug:** No logs captured"
        return "\n\n**🔍 Debug Log:**\n``````"
        
    def start_conversation(self):
        try:
            self.debug_log = []  # Clear previous logs
            self.log("📌 Starting conversation...")
            
            headers = {
                "Authorization": f"Bearer {self.user_token}",
                "Content-Type": "application/json",
                "Accept": "application/json"
            }
            
            self.log(f"🔗 Endpoint: {self.base_endpoint[:80]}...")
            self.log(f"📋 Token length: {len(self.user_token)}")
            
            response = requests.post(
                self.base_endpoint, 
                headers=headers, 
                json={}, 
                timeout=30
            )
            
            self.log(f"📡 Response: {response.status_code}")
            
            if response.status_code in [200, 201]:
                data = response.json()
                self.conversation_id = data.get("id") or data.get("conversationId")
                self.log(f"✅ Conversation ID: {self.conversation_id[:20]}...")
                return True
            else:
                error_text = response.text[:200]
                self.log(f"❌ Error: {response.status_code}")
                self.log(f"Response: {error_text}")
                self.last_error = response.text
                return False
                
        except Exception as e:
            self.log(f"❌ Exception: {str(e)}")
            self.last_error = str(e)
            return False
    
    def send_message(self, message):
        try:
            if not self.conversation_id:
                if not self.start_conversation():
                    error_msg = f"❌ Konnte Konversation nicht starten.{self.get_debug()}"
                    return error_msg
            
            headers = {
                "Authorization": f"Bearer {self.user_token}",
                "Content-Type": "application/json",
                "Accept": "application/json"
            }
            
            user_id = f"user_{abs(hash(self.user_token)) % 100000}"
            payload = {
                "type": "message",
                "text": message,
                "from": {"id": user_id, "name": "User"},
                "locale": "de-DE"
            }
            
            # Build URL correctly
            if '?' in self.base_endpoint:
                base_url, query = self.base_endpoint.split('?', 1)
                url = f"{base_url}/{self.conversation_id}/activities?{query}"
            else:
                url = f"{self.base_endpoint}/{self.conversation_id}/activities"
            
            self.log(f"📤 Sending message...")
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            self.log(f"📡 Send status: {response.status_code}")
            
            if response.status_code in [200, 201, 202]:
                return self.get_response()
            else:
                error_msg = f"❌ Senden fehlgeschlagen: {response.status_code}{self.get_debug()}"
                return error_msg
                
        except Exception as e:
            error_msg = f"❌ Exception beim Senden: {str(e)}{self.get_debug()}"
            return error_msg
    
    def get_response(self, max_attempts=20, delay=1.5):
        try:
            headers = {
                "Authorization": f"Bearer {self.user_token}",
                "Accept": "application/json"
            }
            
            user_id = f"user_{abs(hash(self.user_token)) % 100000}"
            
            self.log(f"📥 Polling for response ({max_attempts} attempts)...")
            
            for attempt in range(max_attempts):
                time.sleep(delay)
                
                try:
                    # Build URL correctly
                    if '?' in self.base_endpoint:
                        base_url, query = self.base_endpoint.split('?', 1)
                        url = f"{base_url}/{self.conversation_id}/activities?{query}"
                    else:
                        url = f"{self.base_endpoint}/{self.conversation_id}/activities"
                    
                    if self.watermark:
                        url += f"&watermark={self.watermark}" if '?' in url else f"?watermark={self.watermark}"
                    
                    response = requests.get(url, headers=headers, timeout=30)
                    
                    if response.status_code == 200:
                        data = response.json()
                        activities = data.get("activities", [])
                        self.watermark = data.get("watermark")
                        
                        self.log(f"Poll {attempt+1}: {len(activities)} activities")
                        
                        # Look for bot response
                        for activity in reversed(activities):
                            if activity.get("type") == "message":
                                from_id = activity.get("from", {}).get("id", "")
                                from_name = activity.get("from", {}).get("name", "")
                                text = activity.get("text", "")
                                
                                # If it's not from the user and has text, it's a bot response
                                if from_id != user_id and text and text.strip():
                                    self.log(f"✅ Bot response from {from_name}")
                                    return text
                    else:
                        self.log(f"Poll {attempt+1}: Error {response.status_code}")
                        
                except Exception as e:
                    self.log(f"Poll {attempt+1}: Exception {str(e)[:50]}")
                    continue
            
            return f"⏱️ Bot did not respond in time{self.get_debug()}"
            
        except Exception as e:
            return f"❌ Error getting response: {str(e)}{self.get_debug()}"


# =============== TED SCRAPER (keeping your existing code) ===============
def fetch_all_notices_to_json(cpv_codes, date_start, date_end, buyer_country, json_file):
    query = f"(publication-date >={date_start}<={date_end}) AND (buyer-country IN ({buyer_country})) AND (classification-cpv IN ({cpv_codes})) AND (notice-type IN (pin-cfc-standard pin-cfc-social qu-sy cn-standard cn-social subco cn-desg))"
    payload = {
        "query": query, "fields": ["publication-number", "links"], "scope": "ACTIVE",
        "checkQuerySyntax": False, "paginationMode": "PAGE_NUMBER", "page": 1, "limit": 100
    }
    s = requests.Session()
    s.headers.update({"Accept": "application/json"})
    all_notices = []
    page = 1
    while True:
        body = dict(payload)
        body["page"] = page
        r = s.post(API, json=body, timeout=60)
        r.raise_for_status()
        data = r.json()
        notices = data.get("results") or data.get("items") or []
        if not notices:
            break
        all_notices.extend(notices)
        total = data.get("total") or data.get("totalCount")
        if not notices or (total and page * 100 >= total):
            break
        page += 1
        time.sleep(0.2)
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump({"notices": all_notices}, f, ensure_ascii=False, indent=2)

def _get_links_block(notice: dict) -> dict:
    links = notice.get("links") or {}
    if isinstance(links, dict) and "links" in links and isinstance(links["links"], dict):
        links = links["links"]
    return {(k.lower() if isinstance(k, str) else k): v for k, v in (links.items() if isinstance(links, dict) else [])}

def _extract_xml_urls_from_notice(notice: dict) -> list:
    block = _get_links_block(notice)
    xml_block = block.get("xml")
    urls = []
    if isinstance(xml_block, dict):
        for k, v in xml_block.items():
            if isinstance(k, str) and k.lower() == "mul" and v:
                urls.append(v)
        for k, v in xml_block.items():
            if isinstance(k, str) and k.lower() != "mul" and v:
                urls.append(v)
    elif isinstance(xml_block, str) and xml_block:
        urls.append(xml_block)
    return urls

def fetch_notice_xml(session, pubno: str, notice: dict) -> bytes:
    xml_headers = {"Accept": "application/xml", "User-Agent": "Mozilla/5.0"}
    for url in _extract_xml_urls_from_notice(notice):
        try:
            r = session.get(url, headers=xml_headers, timeout=60)
            if r.status_code == 200 and r.content.strip():
                return r.content
        except:
            pass
    for lang in ("en", "de", "fr"):
        url = f"https://ted.europa.eu/{lang}/notice/{pubno}/xml"
        try:
            r = session.get(url, headers=xml_headers, timeout=60)
            if r.status_code == 200 and r.content.strip():
                return r.content
        except:
            pass
    raise RuntimeError(f"No XML found for {pubno}")

def _first_text(nodes):
    for n in nodes or []:
        t = (n.text or "").strip()
        if t:
            return t
    return ""

def _norm_date(d: str) -> str:
    if not d:
        return ""
    d = d.rstrip("Zz")
    return d.split("T")[0].split("+")[0]

def _clean_title(raw: str) -> str:
    if not raw:
        return ""
    return re.sub(r"^\s*\d{4}[-_]\d{5,}[\s_\-–:]+", "", raw.strip())

def _parse_iso_date(d: str):
    try:
        return datetime.strptime(d, "%Y-%m-%d")
    except:
        return None

def _duration_to_days(val: str, unit: str):
    if not val:
        return None
    try:
        num = float(str(val).strip().replace(",", "."))
    except:
        return None
    u = (unit or "").upper()
    if u in ("DAY", "D", "DAYS"):
        return int(round(num))
    if u in ("MON", "M", "MONTH", "MONTHS"):
        return int(round(num * 30))
    if u in ("ANN", "Y", "YEAR", "YEARS"):
        return int(round(num * 365))
    return None

def parse_xml_fields(xml_bytes: bytes) -> dict:
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.parse(BytesIO(xml_bytes), parser)
    ns = {k: v for k, v in (root.getroot().nsmap or {}).items() if k}
    ns.setdefault("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")
    ns.setdefault("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")
    ns.setdefault("efac", "http://data.europa.eu/p27/eforms-ubl-extension-aggregate-components/1")
    ns.setdefault("efbc", "http://data.europa.eu/p27/eforms-ubl-extension-basic-components/1")

    out = {}
    out["Beschaffer"] = _first_text(root.xpath(".//cac:ContractingParty//cac:PartyName/cbc:Name", namespaces=ns) or root.xpath(".//efac:Organizations//efac:Company/cac:PartyName/cbc:Name", namespaces=ns))
    out["Projektbezeichnung"] = _clean_title(_first_text(root.xpath(".//cac:ProcurementProject/cbc:Name | .//cbc:Title | .//efbc:Title", namespaces=ns)))
    out["Ort/Region"] = _first_text(root.xpath("//cac:PostalAddress[1]/cbc:CityName", namespaces=ns))
    out["Vergabeplattform"] = _first_text(root.xpath(".//cbc:AccessToolsURI | .//cbc:WebsiteURI | .//cbc:URI | .//cbc:EndpointID", namespaces=ns))
    pub_id = _first_text(root.xpath(".//efbc:NoticePublicationID[@schemeName='ojs-notice-id']", namespaces=ns))
    out["Ted-Link"] = f"https://ted.europa.eu/en/notice/-/detail/{pub_id}" if pub_id else ""
    out["Projektstart"] = _norm_date(_first_text(root.xpath(".//cac:ProcurementProject/cac:PlannedPeriod/cbc:StartDate", namespaces=ns)))
    out["Projektende"] = _norm_date(_first_text(root.xpath(".//cac:ProcurementProject/cac:PlannedPeriod/cbc:EndDate", namespaces=ns)))
    out["Geforderte Unternehmensreferenzen"] = ""
    out["Geforderte Kriterien CVs"] = ""
    out["Projektvolumen"] = ""
    out["Frist Abgabedatum"] = _norm_date(_first_text(root.xpath(".//cac:TenderSubmissionDeadlinePeriod/cbc:EndDate", namespaces=ns)))
    out["Veröffentlichung Datum"] = _norm_date(_first_text(root.xpath(".//cbc:PublicationDate", namespaces=ns)))
    out["CPV Codes"] = ""
    out["Leistungen/Rollen"] = ""
    return out

def main_scraper(cpv_codes, date_start, date_end, buyer_country, output_excel):
    temp_json = tempfile.mktemp(suffix=".json")
    fetch_all_notices_to_json(cpv_codes, date_start, date_end, buyer_country, temp_json)
    with open(temp_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    notices = data.get("notices", [])
    s = requests.Session()
    rows = []
    for n in notices:
        pubno = n.get("publication-number")
        if not pubno:
            continue
        try:
            xml_bytes = fetch_notice_xml(s, pubno, n)
            fields = parse_xml_fields(xml_bytes)
            fields["publication-number"] = pubno
            rows.append(fields)
        except:
            pass
        time.sleep(0.25)
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["publication-number", "Beschaffer", "Projektbezeichnung", "Ort/Region", "Vergabeplattform", "Ted-Link", "Projektstart", "Projektende", "Geforderte Unternehmensreferenzen", "Geforderte Kriterien CVs", "Projektvolumen", "Frist Abgabedatum", "Veröffentlichung Datum", "CPV Codes", "Leistungen/Rollen"]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    last_row = len(rows) + 1
    last_col = len(headers)
    table_range = f"A1:{get_column_letter(last_col)}{last_row}"
    table = Table(displayName="Teddata", ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)
    wb.save(output_excel)
    os.remove(temp_json)

# =============== FILE PROCESSING ===============
def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num, page in enumerate(pdf_reader.pages):
            text += f"\n--- Page {page_num + 1} ---\n"
            text += page.extract_text()
        return text
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_txt(file):
    try:
        return file.read().decode('utf-8')
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_excel(file):
    try:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        return f"Rows: {len(df)}, Columns: {len(df.columns)}\n\n{df.head(50).to_string()}"
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_image(file):
    try:
        Image.open(file)
        return f"Image: {file.name}"
    except Exception as e:
        return f"Error: {str(e)}"

def process_uploaded_file(uploaded_file):
    ext = uploaded_file.name.split('.')[-1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(uploaded_file)
    elif ext == 'docx':
        return extract_text_from_docx(uploaded_file)
    elif ext == 'txt':
        return extract_text_from_txt(uploaded_file)
    elif ext in ['xlsx', 'xls', 'csv']:
        return extract_text_from_excel(uploaded_file)
    elif ext in ['png', 'jpg', 'jpeg']:
        return extract_text_from_image(uploaded_file)
    return f"Unsupported: {ext}"

# =============== MAIN ===============
def main():
    st.set_page_config(page_title="TED Scraper & AI Assistant", layout="wide", initial_sidebar_state="collapsed")
    
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] { background-color: #343541; }
    [data-testid="stHeader"] { background-color: #343541; }
    [data-testid="stSidebar"] { background-color: #202123; }
    .main .block-container { padding-top: 2rem !important; max-width: 48rem !important; margin: 0 auto !important; }
    .stChatMessage { background-color: transparent !important; padding: 1.5rem 0 !important; }
    [data-testid="stChatMessageContent"] { background-color: #444654 !important; border-radius: 0.5rem; padding: 1rem 1.5rem !important; color: #ececf1 !important; }
    [data-testid="stChatMessage"][data-testid*="user"] [data-testid="stChatMessageContent"] { background-color: #343541 !important; }
    .stMarkdown, .stText { color: #ececf1 !important; }
    h1, h2, h3, h4, h5, h6 { color: #ececf1 !important; }
    .stButton button { background-color: #10a37f !important; color: white !important; }
    .stButton button:hover { background-color: #1a7f64 !important; }
    </style>
    """, unsafe_allow_html=True)
    
    auth_flow()
    
    tab1, tab2 = st.tabs(["📄 TED Scraper", "💬 AI Assistant"])
    
    with tab1:
        st.header("📄 TED EU Notice Scraper")
        c1, c2 = st.columns(2)
        with c1:
            cpv = st.text_input("CPV Codes", "71541000")
        with c2:
            country = st.text_input("Country", "DEU")
        today = date.today()
        c1, c2 = st.columns(2)
        with c1:
            start = st.date_input("Start", today)
        with c2:
            end = st.date_input("End", today)
        filename = st.text_input("Filename", f"ted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        if st.button("Run Scraper"):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                try:
                    main_scraper(cpv, start.strftime("%Y%m%d"), end.strftime("%Y%m%d"), country, tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("Download", f.read(), filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    if os.path.exists(tmp.name):
                        os.remove(tmp.name)
    
    with tab2:
        if "document_store" not in st.session_state:
            st.session_state.document_store = {}
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = []
        
        with st.sidebar:
            st.markdown("## 🔑 Config")
            if COPILOT_STUDIO_ENDPOINT:
                st.success("✅ Copilot Studio")
            if st.button("🗑️ Clear", use_container_width=True):
                st.session_state.clear()
                st.rerun()
        
        if not COPILOT_STUDIO_ENDPOINT:
            st.error("No endpoint configured")
        else:
            if "copilot_client" not in st.session_state:
                st.session_state.copilot_client = CopilotStudioClient(COPILOT_STUDIO_ENDPOINT, st.session_state.user_token)
            
            if not st.session_state.chat_messages:
                with st.chat_message("assistant", avatar=JKM_LOGO_URL):
                    st.markdown("👋 Willkommen! Ich bin mit Ihrem SharePoint verbunden. Stellen Sie mir eine Frage!")
            
            for msg in st.session_state.chat_messages:
                with st.chat_message(msg["role"], avatar=JKM_LOGO_URL if msg["role"] == "assistant" else BOT_AVATAR_URL):
                    st.markdown(msg["content"])
            
            if prompt := st.chat_input("Message..."):
                st.session_state.chat_messages.append({"role": "user", "content": prompt})
                with st.chat_message("user", avatar=BOT_AVATAR_URL):
                    st.markdown(prompt)
                
                with st.chat_message("assistant", avatar=JKM_LOGO_URL):
                    placeholder = st.empty()
                    placeholder.markdown("💭 Thinking...")
                    response = st.session_state.copilot_client.send_message(prompt)
                    placeholder.markdown(response)
                    st.session_state.chat_messages.append({"role": "assistant", "content": response})

if __name__ == "__main__":
    main()


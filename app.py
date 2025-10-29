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

# ------------------- CONFIGURATION -------------------
def get_secret(key, default=""):
    """Safely get secrets with fallback"""
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

CLIENT_ID = get_secret("CLIENT_ID")
CLIENT_SECRET = get_secret("CLIENT_SECRET")
TENANT_ID = get_secret("TENANT_ID")
REDIRECT_URI = get_secret("REDIRECT_URI", "http://localhost:8501")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else ""
SCOPE = ["https://graph.microsoft.com/User.Read"]

API = "https://api.ted.europa.eu/v3/notices/search"

# JKM Logo URL
JKM_LOGO_URL = "https://www.xing.com/imagecache/public/scaled_original_image/eyJ1dWlkIjoiMGE2MTk2MTYtODI4Zi00MWZlLWEzN2ItMjczZGM2ODc5MGJmIiwiYXBwX2NvbnRleHQiOiJlbnRpdHktcGFnZXMiLCJtYXhfd2lkdGgiOjMyMCwibWF4X2hlaWdodCI6MzIwfQ?signature=a21e5c1393125a94fc9765898c25d73a064665dc3aacf872667c902d7ed9c3f9"

# ------------------- AUTHENTICATION -------------------
def build_msal_app():
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        st.error("❌ Microsoft OAuth credentials not configured!")
        st.info("""
        Please create `.streamlit/secrets.toml` in your project directory with:
        
        ```
        CLIENT_ID = "your-client-id"
        CLIENT_SECRET = "your-client-secret"
        TENANT_ID = "your-tenant-id"
        REDIRECT_URI = "http://localhost:8501"
        ```
        """)
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
    st.markdown("""
    <style>
    .block-container { padding: 0 !important; max-width: 100vw !important; }
    .center-root {
        min-height: 100vh; width: 100vw;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        background: linear-gradient(120deg, #eaf6fb 0%, #f3e9f5 100%);
    }
    .jkm-logo {
        height: 72px; margin-bottom: 14px; border-radius: 14px; box-shadow: 0 2px 14px rgba(70,80,120,0.08);
        background: #fff; display: block;
    }
    .app-title {
        font-family: 'Segoe UI', Arial,sans-serif;
        font-size: 2.3em; text-align: center; font-weight: 800; color: #283044; margin-bottom: 10px; margin-top: 4px;
    }
    .welcome-text {
        font-size: 1.07em; color: #505A69; margin-bottom: 22px; margin-top: 0; text-align: center;
    }
    .login-card {
        width: 375px; padding: 38px 34px 31px 34px; background: #fff; border-radius: 18px;
        box-shadow: 0 8px 32px rgba(50,72,140,.13); text-align: center; margin-top: 6px;
    }
    .microsoft-logo {
        height: 44px; margin-bottom: 16px; display:block; margin-left:auto; margin-right:auto;
    }
    .login-button {
        display: block; width: 100%; padding: 17px 0 14px 0; margin: 28px 0 18px 0; font-size: 17px;
        background-color: #0078d7; color: #fff !important; border: none; border-radius: 7px;
        cursor: pointer; text-decoration: none; font-weight: 600; transition: background 0.18s; outline: none;
    }
    .login-button:hover {
        background-color: #005fa1; color: #fff !important; text-decoration: none;
    }
    </style>
    """, unsafe_allow_html=True)
    st.markdown(f"""
    <div class="center-root">
        <img src="{JKM_LOGO_URL}" class="jkm-logo" alt="JKM Consult Logo"/>
        <div class="app-title">TED Scraper & AI Assistant</div>
        <div class="welcome-text">
            Welcome! Access project info securely.<br>
            Login with Microsoft to continue.
        </div>
        <div class="login-card">
            <img src="https://upload.wikimedia.org/wikipedia/commons/4/44/Microsoft_logo.svg" class="microsoft-logo" alt="Microsoft Logo"/>
            <h2 style="margin-bottom: 9px; font-size: 1.26em;">Sign in</h2>
            <p style="font-size: 1em; color: #232b39; margin-bottom: 9px;">
                to continue to <b>TED Scraper & AI Assistant</b>
            </p>
            <a href="{auth_url}" class="login-button">
                Sign in with Microsoft
            </a>
            <p style="margin-top: 32px; font-size: 0.98em; color: #888;">
                Your credentials are always handled by Microsoft.<br>
                We never see or store your password.
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

def auth_flow():
    params = st.query_params
    if "code" in params and "user_token" not in st.session_state:
        code = params["code"]
        if isinstance(code, list):
            code = code[0]
        token_data = fetch_token(code)
        if "access_token" in token_data:
            st.session_state["user_token"] = token_data["access_token"]
            st.query_params.clear()
            st.rerun()
        else:
            st.error("Microsoft login failed. Please try again.")
            st.stop()
    if "user_token" not in st.session_state:
        login_button()
        st.stop()
    return True

# ---------------- TED SCRAPER FUNCTIONS ----------------
def fetch_all_notices_to_json(cpv_codes, date_start, date_end, buyer_country, json_file):
    query = (
        f"(publication-date >={date_start}<={date_end}) AND (buyer-country IN ({buyer_country})) "
        f"AND (classification-cpv IN ({cpv_codes})) AND (notice-type IN (pin-cfc-standard pin-cfc-social qu-sy cn-standard cn-social subco cn-desg))"
    )
    payload = {
        "query": query,
        "fields": ["publication-number", "links"],
        "scope": "ACTIVE",
        "checkQuerySyntax": False,
        "paginationMode": "PAGE_NUMBER",
        "page": 1,
        "limit": 100
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
        notices = data.get("results") or data.get("items") or data.get("notices") or []
        if not notices:
            break
        all_notices.extend(notices)
        total = data.get("total") or data.get("totalCount")
        if not notices or (total and page * payload["limit"] >= total):
            break
        page += 1
        time.sleep(0.2)
    with open(json_file, "w", encoding="utf-8") as f:
        json.dump({"notices": all_notices}, f, ensure_ascii=False, indent=2)

def _get_links_block(notice: dict) -> dict:
    links = notice.get("links") or {}
    if isinstance(links, dict) and "links" in links and isinstance(links["links"], dict):
        links = links["links"]
    if isinstance(links, dict):
        return { (k.lower() if isinstance(k,str) else k): v for k,v in links.items() }
    return {}

def _extract_xml_urls_from_notice(notice: dict) -> list:
    block = _get_links_block(notice)
    xml_block = block.get("xml")
    urls = []
    if isinstance(xml_block, dict):
        for k,v in xml_block.items():
            if isinstance(k,str) and k.lower()=="mul" and v:
                urls.append(v)
        for k,v in xml_block.items():
            if isinstance(k,str) and k.lower()!="mul" and v:
                urls.append(v)
    elif isinstance(xml_block, str) and xml_block:
        urls.append(xml_block)
    return urls

def fetch_notice_xml(session: requests.Session, pubno: str, notice: dict) -> bytes:
    xml_headers = {"Accept":"application/xml","User-Agent":"Mozilla/5.0"}
    for url in _extract_xml_urls_from_notice(notice):
        try:
            r = session.get(url, headers=xml_headers, timeout=60)
            if r.status_code == 200 and r.content.strip():
                return r.content
        except requests.RequestException:
            pass
    for lang in ("en","de","fr"):
        url = f"https://ted.europa.eu/{lang}/notice/{pubno}/xml"
        try:
            r = session.get(url, headers=xml_headers, timeout=60)
            if r.status_code == 200 and r.content.strip():
                return r.content
        except requests.RequestException:
            pass
    detail_url = f"https://ted.europa.eu/en/notice/-/detail/{pubno}"
    try:
        html = session.get(detail_url, headers={"User-Agent":"Mozilla/5.0"}, timeout=60).text
        m = re.search(r'https://ted\.europa\.eu/(?:en|de|fr)/notice/' + re.escape(pubno) + r'/xml', html)
        if m:
            r = session.get(m.group(0), headers=xml_headers, timeout=60)
            if r.status_code == 200 and r.content.strip():
                return r.content
    except requests.RequestException:
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
    if not raw: return ""
    return re.sub(r"^\s*\d{4}[-_]\d{5,}[\s_\-–:]+", "", raw.strip())

def _parse_iso_date(d: str):
    try:
        return datetime.strptime(d, "%Y-%m-%d")
    except Exception:
        return None

def _duration_to_days(val: str, unit: str) -> int or None:
    if not val:
        return None
    try:
        num = float(str(val).strip().replace(",", "."))
    except Exception:
        return None
    u = (unit or "").upper()
    if u in ("DAY","D","DAYS"):
        return int(round(num))
    if u in ("MON","M","MONTH","MONTHS"):
        return int(round(num * 30))
    if u in ("ANN","Y","YEAR","YEARS"):
        return int(round(num * 365))
    return None

def parse_xml_fields(xml_bytes: bytes) -> dict:
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.parse(BytesIO(xml_bytes), parser)
    ns = {k: v for k, v in (root.getroot().nsmap or {}).items() if k}
    ns.setdefault("cbc","urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")
    ns.setdefault("cac","urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")
    ns.setdefault("efac","http://data.europa.eu/p27/eforms-ubl-extension-aggregate-components/1")
    ns.setdefault("efbc","http://data.europa.eu/p27/eforms-ubl-extension-basic-components/1")

    out = {}
    out["Beschaffer"] = _first_text(
        root.xpath(".//cac:ContractingParty//cac:PartyName/cbc:Name", namespaces=ns)
        or root.xpath(".//efac:Organizations//efac:Company/cac:PartyName/cbc:Name", namespaces=ns)
    )
    out["Projektbezeichnung"] = _clean_title(
        _first_text(root.xpath(".//cac:ProcurementProject/cbc:Name | .//cbc:Title | .//efbc:Title", namespaces=ns))
    )
    out["Ort/Region"] = _first_text(root.xpath("//cac:PostalAddress[1]/cbc:CityName", namespaces=ns))
    out["Vergabeplattform"] = _first_text(
        root.xpath(".//cbc:AccessToolsURI | .//cbc:WebsiteURI | .//cbc:URI | .//cbc:EndpointID", namespaces=ns)
    )
    pub_id = _first_text(root.xpath(".//efbc:NoticePublicationID[@schemeName='ojs-notice-id']", namespaces=ns))
    out["Ted-Link"] = f"https://ted.europa.eu/en/notice/-/detail/{pub_id}" if pub_id else ""

    start_nodes = root.xpath(
        ".//cac:ProcurementProject/cac:PlannedPeriod/cbc:StartDate "
        "| .//cac:ProcurementProjectLot//cac:ProcurementProject/cac:PlannedPeriod/cbc:StartDate",
        namespaces=ns
    )
    end_nodes = root.xpath(
        ".//cac:ProcurementProject/cac:PlannedPeriod/cbc:EndDate "
        "| .//cac:ProcurementProjectLot//cac:ProcurementProject/cac:PlannedPeriod/cbc:EndDate",
        namespaces=ns
    )
    start_norm = _norm_date(_first_text(start_nodes))
    end_norm = _norm_date(_first_text(end_nodes))

    if not start_norm and end_norm:
        dur_nodes = root.xpath(
            ".//cac:ProcurementProject/cac:PlannedPeriod/cbc:DurationMeasure "
            "| .//cac:ProcurementProjectLot//cac:ProcurementProject/cac:PlannedPeriod/cbc:DurationMeasure",
            namespaces=ns
        )
        dur_val, dur_unit = None, None
        for dn in dur_nodes:
            text_val = (dn.text or "").strip()
            unit = (dn.get("unitCode") or "").strip()
            if text_val:
                dur_val, dur_unit = text_val, unit
                break
        days = _duration_to_days(dur_val, dur_unit) if dur_val else None
        if days:
            end_dt = _parse_iso_date(end_norm)
            if end_dt:
                start_norm = (end_dt - timedelta(days=days)).strftime("%Y-%m-%d")

    out["Projektstart"] = start_norm
    out["Projektende"] = end_norm

    crit_nodes = root.xpath(
        ".//*[contains(local-name(),'SelectionCriteria') or contains(local-name(),'SelectionCriterion')]/cbc:Description",
        namespaces=ns
    )
    crit_text = " ".join((n.text or "").strip() for n in crit_nodes if (n.text or "").strip())
    crit_text = re.sub(r"\bslc-[a-z0-9\-]+\b", "", crit_text, flags=re.I).strip()
    out["Geforderte Unternehmensreferenzen"] = crit_text
    out["Geforderte Kriterien CVs"] = "CV" if re.search(
        r"\b(CV|Lebenslauf|Schlüsselpersonal|key staff|personaleinsatz)\b", crit_text, re.I
    ) else ""

    amount_nodes = root.xpath(
        ".//cbc:EstimatedOverallContractAmount | .//cbc:EstimatedOverallContractAmount/cbc:Value | .//efbc:EstimatedValue | .//cbc:PayableAmount",
        namespaces=ns
    )
    value_text = ""
    if amount_nodes:
        for node in amount_nodes:
            if node.text and node.text.strip():
                value_text = node.text.strip()
                parent = node.getparent()
                currency = node.get("currencyID") or (parent.get("currencyID") if parent is not None else None)
                if currency:
                    value_text += f" {currency}"
                break
    out["Projektvolumen"] = value_text or ""

    tender_deadline_date = _norm_date(
        _first_text(root.xpath(".//cac:TenderSubmissionDeadlinePeriod/cbc:EndDate", namespaces=ns))
    )
    if not tender_deadline_date:
        tender_deadline_date = _norm_date(
            _first_text(root.xpath(".//cac:TenderingTerms/cbc:SubmissionDeadlineDate", namespaces=ns))
        )
    if not tender_deadline_date:
        tender_deadline_date = _norm_date(
            _first_text(root.xpath(".//cac:InterestExpressionReceptionPeriod/cbc:EndDate", namespaces=ns))
        )
    if not tender_deadline_date:
        tender_deadline_date = _norm_date(
            _first_text(root.xpath(".//efac:InterestExpressionReceptionPeriod/cbc:EndDate", namespaces=ns))
        )
    participation_deadline_date = _norm_date(
        _first_text(root.xpath(".//cac:ParticipationRequestReceptionPeriod/cbc:EndDate", namespaces=ns))
    )
    if not participation_deadline_date:
        participation_deadline_date = _norm_date(
            _first_text(root.xpath(".//efac:ParticipationRequestReceptionPeriod/cbc:EndDate", namespaces=ns))
        )
    out["Frist Abgabedatum"] = tender_deadline_date or participation_deadline_date

    pub_date = _first_text(root.xpath(".//efbc:PublicationDate", namespaces=ns))
    if not pub_date:
        pub_date = _first_text(root.xpath(".//cbc:PublicationDate", namespaces=ns))
    out["Veröffentlichung Datum"] = _norm_date(pub_date)

    cpv_codes_set = set()
    main_cpv_nodes = root.xpath(".//cac:MainCommodityClassification/cbc:ItemClassificationCode", namespaces=ns)
    for node in main_cpv_nodes:
        if node.text:
            cpv_codes_set.add(node.text.strip())
    add_cpv_nodes = root.xpath(".//cac:AdditionalCommodityClassification/cbc:ItemClassificationCode", namespaces=ns)
    for node in add_cpv_nodes:
        if node.text:
            cpv_codes_set.add(node.text.strip())
    out["CPV Codes"] = ", ".join(sorted(cpv_codes_set))

    lots = root.xpath(".//cac:ProcurementProjectLot", namespaces=ns)
    lot_names = []
    for lot in lots:
        lot_name = lot.xpath(".//cac:ProcurementProject/cbc:Name", namespaces=ns)
        if lot_name and len(lot_name) > 0:
            text = lot_name[0].text.strip()
            if text:
                lot_names.append(text)
    out["Leistungen/Rollen"] = "; ".join(lot_names)

    return out

def main_scraper(cpv_codes, date_start, date_end, buyer_country, output_excel):
    temp_json = tempfile.mktemp(suffix=".json")
    fetch_all_notices_to_json(cpv_codes, date_start, date_end, buyer_country, temp_json)
    with open(temp_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    notices = data.get("results") or data.get("items") or data.get("notices") or []
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
            fields.setdefault("Ted-Link", f"https://ted.europa.eu/en/notice/-/detail/{pubno}")
            rows.append(fields)
        except Exception as e:
            print(f"ERR {pubno}: {e}")
        time.sleep(0.25)
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "publication-number","Beschaffer","Projektbezeichnung","Ort/Region",
        "Vergabeplattform","Ted-Link","Projektstart","Projektende",
        "Geforderte Unternehmensreferenzen","Geforderte Kriterien CVs",
        "Projektvolumen", "Frist Abgabedatum", "Veröffentlichung Datum", "CPV Codes", "Leistungen/Rollen"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    last_row = len(rows) + 1
    last_col = len(headers)
    table_range = f"A1:{get_column_letter(last_col)}{last_row}"
    table = Table(displayName="Teddata", ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    wb.save(output_excel)
    os.remove(temp_json)

# ---------------- CHATBOT FUNCTIONS ----------------
def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        return text
    except Exception as e:
        st.error(f"Error reading DOCX: {e}")
        return None

def extract_text_from_txt(file):
    try:
        return file.read().decode('utf-8')
    except Exception as e:
        st.error(f"Error reading TXT: {e}")
        return None

def process_uploaded_file(uploaded_file):
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension == 'pdf':
        return extract_text_from_pdf(uploaded_file)
    elif file_extension == 'docx':
        return extract_text_from_docx(uploaded_file)
    elif file_extension == 'txt':
        return extract_text_from_txt(uploaded_file)
    else:
        st.error(f"Unsupported file type: {file_extension}")
        return None

def get_azure_chatbot_response(messages, azure_endpoint, azure_key, deployment_name, api_version="2024-08-01-preview"):
    client = AzureOpenAI(
        azure_endpoint=azure_endpoint,
        api_key=azure_key,
        api_version=api_version
    )
    
    stream = client.chat.completions.create(
        model=deployment_name,
        messages=messages,
        stream=True,
        temperature=0.7,
    )
    return stream


# ------------------- MAIN APP -------------------
def main():
    st.set_page_config(page_title="TED Scraper & AI Assistant", layout="wide", initial_sidebar_state="collapsed")
    
    # ChatGPT-style Custom CSS with Fixed Bottom Input
    st.markdown("""
    <style>
    /* ChatGPT-style theme */
    [data-testid="stAppViewContainer"] {
        background-color: #343541;
    }
    
    [data-testid="stHeader"] {
        background-color: #343541;
    }
    
    [data-testid="stSidebar"] {
        background-color: #202123;
    }
    
    /* Main container with padding for fixed input */
    .main .block-container {
        padding-bottom: 180px !important;
        max-width: 48rem !important;
        margin: 0 auto !important;
    }
    
    /* Chat message styling */
    .stChatMessage {
        background-color: transparent !important;
        padding: 1.5rem 0 !important;
    }
    
    [data-testid="stChatMessageContent"] {
        background-color: #444654 !important;
        border-radius: 0.5rem;
        padding: 1rem 1.5rem !important;
        color: #ececf1 !important;
    }
    
    /* User message - darker background */
    [data-testid="stChatMessage"][data-testid*="user"] [data-testid="stChatMessageContent"] {
        background-color: #343541 !important;
    }
    
    /* Fixed input container at bottom */
    .stChatInputContainer {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        right: 0 !important;
        background-color: #343541 !important;
        padding: 1.5rem !important;
        border-top: 1px solid #565869 !important;
        z-index: 999 !important;
        max-width: 48rem !important;
        margin: 0 auto !important;
    }
    
    /* Larger vertical input field */
    [data-testid="stChatInput"] {
        background-color: #40414f !important;
        color: #ececf1 !important;
        border: 1px solid #565869 !important;
        border-radius: 0.75rem !important;
        padding: 1.25rem 3.5rem 1.25rem 3rem !important;
        font-size: 1rem !important;
        min-height: 70px !important;
        line-height: 1.5 !important;
    }
    
    [data-testid="stChatInput"]:focus {
        border-color: #10a37f !important;
        box-shadow: 0 0 0 1px #10a37f !important;
    }
    
    /* Attach button in input field */
    .attach-button-container {
        position: fixed;
        bottom: 2rem;
        left: calc(50% - 22rem);
        z-index: 1001;
    }
    
    .attach-button-container button {
        background-color: transparent !important;
        border: none !important;
        color: #8e8ea0 !important;
        font-size: 1.5rem !important;
        padding: 0.5rem !important;
        cursor: pointer !important;
        transition: color 0.2s !important;
    }
    
    .attach-button-container button:hover {
        color: #ececf1 !important;
        background-color: transparent !important;
    }
    
    /* Text and headers */
    .stMarkdown, .stText {
        color: #ececf1 !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        color: #ececf1 !important;
    }
    
    /* Buttons */
    .stButton button {
        background-color: #10a37f !important;
        color: white !important;
        border: none !important;
        border-radius: 0.375rem !important;
        padding: 0.5rem 1rem !important;
        font-weight: 500 !important;
    }
    
    .stButton button:hover {
        background-color: #1a7f64 !important;
    }
    
    /* Avatar styling - person icon for user */
    [data-testid="stChatMessage"] img {
        border-radius: 0.25rem !important;
        width: 32px !important;
        height: 32px !important;
    }
    
    /* Success/Info/Warning boxes */
    .stSuccess, .stInfo, .stWarning {
        background-color: #444654 !important;
        color: #ececf1 !important;
        border-radius: 0.5rem !important;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
        background-color: #343541;
        border-bottom: 1px solid #565869;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #ececf1 !important;
        background-color: transparent;
        border-bottom: 2px solid transparent;
        padding: 1rem 0;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        border-bottom-color: #10a37f !important;
        color: #10a37f !important;
    }
    
    /* Input fields */
    .stTextInput input, .stDateInput input, .stSelectbox select {
        background-color: #40414f !important;
        color: #ececf1 !important;
        border: 1px solid #565869 !important;
        border-radius: 0.375rem !important;
    }
    
    /* Labels */
    label {
        color: #ececf1 !important;
    }
    
    /* Download button */
    .stDownloadButton button {
        background-color: #10a37f !important;
        color: white !important;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #444654 !important;
        color: #ececf1 !important;
        border-radius: 0.5rem !important;
    }
    
    /* File uploader in sidebar */
    [data-testid="stFileUploader"] {
        background-color: #40414f !important;
        border: 1px dashed #565869 !important;
        border-radius: 0.5rem !important;
        padding: 1rem !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Authentication guard
    auth_flow()
    
    # Create tabs after authentication
    tab1, tab2 = st.tabs(["📄 TED Scraper", "💬 AI Assistant"])
    
    # ============= TAB 1: TED SCRAPER =============
    with tab1:
        st.header("📄 TED EU Notice Scraper")
        st.write("Download TED procurement notices to Excel (data is exported as a table for Power Automate).")
        
        with st.expander("ℹ️ How this works / Instructions", expanded=False):
            st.write("""
            1. Enter your filters (CPV, date range, country, filename).
            2. Click **Run Scraper**. The script downloads notices and attachments, saves an Excel file.
            3. Use the download button to save the Excel file wherever you want!
            4. The exported file now contains an Excel table named 'TEDData', ready for Power Automate!
            """)

        c1, c2 = st.columns(2)
        with c1:
            cpv_codes = st.text_input(
                "🔎 CPV Codes (space separated)",
                "71541000 71500000 71240000 79421000 71000000 71248000 71312000 71700000 71300000 71520000 71250000 90712000 71313000",
            )
        with c2:
            buyer_country = st.text_input("🌍 Buyer Country (ISO Alpha-3)", "DEU")

        today = date.today()
        date_col1, date_col2 = st.columns(2)
        with date_col1:
            start_date_obj = st.date_input("📆 Start Publication Date", value=today)
        with date_col2:
            end_date_obj = st.date_input("📆 End Publication Date", value=today)

        date_start = start_date_obj.strftime("%Y%m%d")
        date_end = end_date_obj.strftime("%Y%m%d")

        output_excel = st.text_input(
            "💾 Output Excel filename",
            f"ted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )

        if st.button("▶️ Run Scraper"):
            st.info("Scraping... Please wait (can take a few minutes).")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_excel:
                try:
                    main_scraper(cpv_codes, date_start, date_end, buyer_country, temp_excel.name)
                    st.success("✅ Done! Download your Excel file below.")
                    with open(temp_excel.name, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Excel",
                            data=f.read(),
                            file_name=output_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ Error during scraping: {e}")
                finally:
                    temp_excel.close()
                    if os.path.exists(temp_excel.name):
                        os.remove(temp_excel.name)
    
    # ============= TAB 2: CHATBOT =============
    with tab2:
        # Sidebar for document library
        with st.sidebar:
            # Get Azure credentials from secrets (auto-load)
            azure_endpoint = get_secret("AZURE_ENDPOINT", "")
            azure_key = get_secret("AZURE_API_KEY", "")
            deployment_name = get_secret("DEPLOYMENT_NAME", "gpt-4o-mini")
            api_version = "2024-08-01-preview"
            
            # Show configuration status
            st.markdown("## 🔑 Configuration")
            if azure_endpoint and azure_key:
                st.success("✅ Azure AI Connected")
                try:
                    masked_endpoint = azure_endpoint.replace("https://", "").split(".")[0]
                    st.info(f"🔗 {masked_endpoint}")
                except:
                    st.info("🔗 Endpoint configured")
                st.info(f"🤖 {deployment_name}")
            else:
                st.warning("⚠️ Azure credentials missing")
            
            st.markdown("---")
            st.markdown("## 📚 Document Library")
            
            # Initialize document store in session state
            if "document_store" not in st.session_state:
                st.session_state.document_store = {}
            
            # File uploader in sidebar under Document Library
            library_files = st.file_uploader(
                "Upload Documents", 
                type=['pdf', 'docx', 'txt'],
                accept_multiple_files=True,
                key="library_uploader",
                help="Upload PDFs, Word documents, or text files to library"
            )
            
            if library_files:
                for uploaded_file in library_files:
                    if uploaded_file.name not in st.session_state.document_store:
                        with st.spinner(f"Processing {uploaded_file.name}..."):
                            text = process_uploaded_file(uploaded_file)
                            if text:
                                st.session_state.document_store[uploaded_file.name] = text
                                st.success(f"✅ {uploaded_file.name}")
            
            if st.session_state.document_store:
                st.markdown(f"**📁 {len(st.session_state.document_store)} document(s)**")
                for doc_name in list(st.session_state.document_store.keys()):
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.caption(f"• {doc_name}")
                    with col2:
                        if st.button("🗑️", key=f"del_{doc_name}"):
                            del st.session_state.document_store[doc_name]
                            st.rerun()
            
            # Clear chat button in sidebar
            st.markdown("---")
            if st.button("🗑️ Clear Chat", use_container_width=True):
                st.session_state.chat_messages = []
                if "current_chat_doc" in st.session_state:
                    del st.session_state.current_chat_doc
                if "current_chat_doc_name" in st.session_state:
                    del st.session_state.current_chat_doc_name
                st.rerun()
        
        # Main chat interface
        if not azure_endpoint or not azure_key:
            st.error("❌ Azure AI Foundry credentials not configured!")
            st.info("""
            **Add to `.streamlit/secrets.toml`:**
            ```
            AZURE_ENDPOINT = "https://your-resource.openai.azure.com"
            AZURE_API_KEY = "your-api-key"
            DEPLOYMENT_NAME = "gpt-4o-mini"
            ```
            """)
        else:
            # Initialize chat history
            if "chat_messages" not in st.session_state:
                st.session_state.chat_messages = []
            
            # Display welcome message if no chat history
            if not st.session_state.chat_messages:
                with st.chat_message("assistant", avatar=JKM_LOGO_URL):
                    st.markdown("""
                    👋 **Willkommen beim JKM AI Assistant!**
                    
                    Ich kann Ihnen helfen, Fragen zu Ihren Dokumenten zu beantworten. 
                    
                    **So funktioniert's:**
                    - 📎 Laden Sie Dokumente in der Seitenleiste hoch (Document Library)
                    - 📄 Oder nutzen Sie die Datei-Schaltfläche unten im Chat
                    - ❓ Stellen Sie Fragen zu den Inhalten
                    - 🔍 Ich analysiere die Dokumente und antworte präzise
                    
                    Womit kann ich Ihnen helfen?
                    """)
            
            # Display chat history with person icon for user
            for message in st.session_state.chat_messages:
                avatar = JKM_LOGO_URL if message["role"] == "assistant" else "👨"
                with st.chat_message(message["role"], avatar=avatar):
                    st.markdown(message["content"])
            
            # Initialize file uploader state
            if "chat_file_uploader_key" not in st.session_state:
                st.session_state.chat_file_uploader_key = 0
            
            # Attach button container (positioned with CSS)
            st.markdown('<div class="attach-button-container">', unsafe_allow_html=True)
            
            # Toggle for file uploader
            if "show_file_uploader" not in st.session_state:
                st.session_state.show_file_uploader = False
            
            if st.button("📎", key="attach_btn", help="Attach a file"):
                st.session_state.show_file_uploader = not st.session_state.show_file_uploader
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # File uploader (conditionally shown)
            if st.session_state.show_file_uploader:
                chat_file = st.file_uploader(
                    "Choose a file", 
                    type=['pdf', 'docx', 'txt'],
                    key=f"chat_uploader_{st.session_state.chat_file_uploader_key}",
                    help="Upload a document for this conversation"
                )
                
                if chat_file:
                    if "current_chat_doc" not in st.session_state or st.session_state.get("current_chat_doc_name") != chat_file.name:
                        with st.spinner(f"Processing {chat_file.name}..."):
                            text = process_uploaded_file(chat_file)
                            if text:
                                st.session_state.current_chat_doc = text
                                st.session_state.current_chat_doc_name = chat_file.name
                                st.success(f"✅ {chat_file.name} loaded")
                                st.session_state.show_file_uploader = False
                                st.session_state.chat_file_uploader_key += 1
                                st.rerun()
            
            # Chat input (fixed at bottom via CSS)
            if prompt := st.chat_input("Message JKM AI Assistant..."):
                # Prepare context from library and current document
                context_parts = []
                
                if st.session_state.document_store:
                    library_context = "\n\n".join([
                        f"Document: {name}\n{content[:3000]}" 
                        for name, content in st.session_state.document_store.items()
                    ])
                    context_parts.append(f"=== DOCUMENT LIBRARY ===\n{library_context}")
                
                if "current_chat_doc" in st.session_state:
                    context_parts.append(f"=== CURRENT DOCUMENT ({st.session_state.current_chat_doc_name}) ===\n{st.session_state.current_chat_doc[:5000]}")
                
                if not context_parts:
                    st.warning("⚠️ Bitte laden Sie mindestens ein Dokument hoch.")
                else:
                    full_context = "\n\n".join(context_parts)
                    
                    # Add user message to chat
                    st.session_state.chat_messages.append({"role": "user", "content": prompt})
                    with st.chat_message("user", avatar="👨"):
                        st.markdown(prompt)
                    
                    # Prepare messages for API
                    system_message = {
                        "role": "system",
                        "content": f"""Du bist der JKM AI Assistant - ein hilfreicher KI-Assistent für Ausschreibungen und Projektdokumente.

Hier sind die verfügbaren Dokumente:

{full_context}

Anweisungen:
- Beantworte Fragen NUR basierend auf den bereitgestellten Dokumenten
- Falls die Antwort nicht in den Dokumenten steht, sage das klar
- Antworte immer auf Deutsch, wenn auf Deutsch gefragt wird
- Sei präzise und zitiere spezifische Teile der Dokumente
- Bei Ausschreibungsdokumenten: Hebe wichtige Termine, Anforderungen und Details hervor
- Sei professionell und hilfsbereit"""
                    }
                    
                    api_messages = [system_message] + [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.chat_messages
                    ]
                    
                    # Get and display response
                    with st.chat_message("assistant", avatar=JKM_LOGO_URL):
                        try:
                            stream = get_azure_chatbot_response(
                                api_messages, 
                                azure_endpoint, 
                                azure_key, 
                                deployment_name,
                                api_version
                            )
                            response = st.write_stream(stream)
                            st.session_state.chat_messages.append({"role": "assistant", "content": response})
                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")
                            st.info("Bitte überprüfen Sie Ihre Azure-Konfiguration in secrets.toml")

if __name__ == "__main__":
    main()

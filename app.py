import streamlit as st
import os, json, re, requests, time, tempfile
from datetime import datetime, timedelta, date
from lxml import etree
from io import BytesIO
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from msal import ConfidentialClientApplication

# --------------- CONFIGURATION ---------------
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TENANT_ID = st.secrets["TENANT_ID"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/User.Read"]
REDIRECT_URI = st.secrets["REDIRECT_URI"]

BASE_PATH = "./data"
os.makedirs(BASE_PATH, exist_ok=True)
API = "https://api.ted.europa.eu/v3/notices/search"

# --------------- AUTHENTICATION --------------
def build_msal_app():
    return ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

def fetch_token(auth_code):
    msal_app = build_msal_app()
    return msal_app.acquire_token_by_authorization_code(auth_code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

def login_button():
    jkm_logo_url = "https://www.xing.com/imagecache/public/scaled_original_image/eyJ1dWlkIjoiMGE2MTk2MTYtODI4Zi00MWZlLWEzN2ItMjczZGM2ODc5MGJmIiwiYXBwX2NvbnRleHQiOiJlbnRpdHktcGFnZXMiLCJtYXhfd2lkdGgiOjMyMCwibWF4X2hlaWdodCI6MzIwfQ?signature=a21e5c1393125a94fc9765898c25d73a064665dc3aacf872667c902d7ed9c3f9"
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
        <img src="{jkm_logo_url}" class="jkm-logo" alt="JKM Consult Logo"/>
        <div class="app-title">TED Scraper</div>
        <div class="welcome-text">
            Welcome! Access project info securely.<br>
            Login with Microsoft to continue.
        </div>
        <div class="login-card">
            <img src="https://upload.wikimedia.org/wikipedia/commons/4/44/Microsoft_logo.svg" class="microsoft-logo" alt="Microsoft Logo"/>
            <h2 style="margin-bottom: 9px; font-size: 1.26em;">Sign in</h2>
            <p style="font-size: 1em; color: #232b39; margin-bottom: 9px;">
                to continue to <b>TED Scraper</b>
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

# --------------- BUSINESS LOGIC ---------------

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

# All your XML/excel helpers here unchanged — use as in your latest version
# (Include _get_links_block, _extract_xml_urls_from_notice, fetch_notice_xml, all your field parsing, etc.)

# Place the main scraper and Streamlit UI logic here as in your latest main()
# Most of your earlier main(), main_scraper() code can be reused — the fix is only passing arguments above and in the UI handlers.


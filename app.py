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
import base64
from PIL import Image

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
POWER_PLATFORM_SCOPE = ["https://api.powerplatform.com/.default"]

API = "https://api.ted.europa.eu/v3/notices/search"

COPILOT_STUDIO_ENDPOINT = get_secret("COPILOT_STUDIO_ENDPOINT", "")

# Avatars
JKM_LOGO_URL = "https://www.xing.com/imagecache/public/scaled_original_image/eyJ1dWlkIjoiMGE2MTk2MTYtODI4Zi00MWZlLWEzN2ItMjczZGM2ODc5MGJmIiwiYXBwX2NvbnRleHQiOiJlbnRpdHktcGFnZXMiLCJtYXhfd2lkdGgiOjMyMCwibWF4X2hlaWdodCI6MzIwfQ?signature=a21e5c1393125a94fc9765898c25d73a064665dc3aacf872667c902d7ed9c3f9"
BOT_AVATAR_URL = "https://raw.githubusercontent.com/PratikSondkarJKM/AkquiseWescraper/refs/heads/main/botavatar.svg"

# ------------------- AUTHENTICATION -------------------
def build_msal_app():
    if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
        st.error("❌ Microsoft OAuth credentials not configured!")
        st.stop()
    
    return ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

def fetch_token(auth_code):
    msal_app = build_msal_app()
    return msal_app.acquire_token_by_authorization_code(auth_code, scopes=SCOPE, redirect_uri=REDIRECT_URI)

def get_power_platform_token():
    """Get token for Power Platform API"""
    msal_app = build_msal_app()
    result = msal_app.acquire_token_for_client(scopes=POWER_PLATFORM_SCOPE)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        print(f"Failed to get Power Platform token: {result.get('error_description')}")
        return None

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
            
            # Get Power Platform token
            pp_token = get_power_platform_token()
            if pp_token:
                st.session_state["power_platform_token"] = pp_token
            
            st.query_params.clear()
            st.rerun()
        else:
            st.error("Microsoft login failed. Please try again.")
            st.stop()
    if "user_token" not in st.session_state:
        login_button()
        st.stop()
    return True

# ------------------- COPILOT STUDIO CLIENT -------------------
class CopilotStudioM365Client:
    """Copilot Studio client using Power Platform token"""
    def __init__(self, endpoint_url, power_platform_token):
        self.base_endpoint = endpoint_url.split('?')[0]
        self.power_platform_token = power_platform_token
        self.conversation_id = None
        self.watermark = None
        
    def start_conversation(self):
        headers = {
            "Authorization": f"Bearer {self.power_platform_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        try:
            response = requests.post(
                self.base_endpoint,
                headers=headers,
                json={},
                timeout=30
            )
            
            if response.status_code in [200, 201]:
                data = response.json()
                self.conversation_id = data.get("id") or data.get("conversationId")
                return True
            else:
                print(f"Failed: {response.status_code} - {response.text}")
                return False
        except Exception as e:
            print(f"Exception: {e}")
            return False
    
    def send_message(self, message):
        if not self.conversation_id:
            if not self.start_conversation():
                return "❌ Konnte keine Verbindung herstellen."
        
        headers = {
            "Authorization": f"Bearer {self.power_platform_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        user_id = f"user_{abs(hash(self.power_platform_token)) % 100000}"
        
        activity = {
            "type": "message",
            "text": message,
            "from": {"id": user_id, "name": "User"},
            "channelId": "directline",
            "locale": "de-DE"
        }
        
        try:
            activity_url = f"{self.base_endpoint}/{self.conversation_id}/activities"
            response = requests.post(activity_url, headers=headers, json=activity, timeout=30)
            
            if response.status_code in [200, 201, 202]:
                return self.get_response()
            else:
                return f"❌ Fehler: {response.status_code}"
        except Exception as e:
            return f"❌ Fehler: {str(e)}"
    
    def get_response(self, max_attempts=25, delay=1.5):
        headers = {
            "Authorization": f"Bearer {self.power_platform_token}",
            "Accept": "application/json"
        }
        
        user_id = f"user_{abs(hash(self.power_platform_token)) % 100000}"
        
        for attempt in range(max_attempts):
            time.sleep(delay)
            
            try:
                activities_url = f"{self.base_endpoint}/{self.conversation_id}/activities"
                if self.watermark:
                    activities_url += f"?watermark={self.watermark}"
                
                response = requests.get(activities_url, headers=headers, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    activities = data.get("activities", [])
                    self.watermark = data.get("watermark")
                    
                    for activity in reversed(activities):
                        if activity.get("type") == "message":
                            from_id = activity.get("from", {}).get("id", "")
                            if from_id != user_id:
                                text = activity.get("text", "")
                                if text and text.strip():
                                    return text
            except Exception as e:
                continue
        
        return "⏱️ Antwort dauert länger. Bitte erneut versuchen."

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
    with open(json_file, "w", encoding="utf-8") as

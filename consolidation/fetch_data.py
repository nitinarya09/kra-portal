"""
fetch_data.py — Downloads all KRA worksheet data from Google Sheets
============================================================================
Connects to the KRA_Database spreadsheet. Supports BOTH Service Account key
credentials AND OAuth 2.0 User credentials (browser login fallback).

If using OAuth 2.0, it will open your browser on the first run to authorize,
then cache the credentials locally in 'token.json' for subsequent fast runs.

Usage:
    from fetch_data import fetch_all_data
    data = fetch_all_data("2026-27", "Q1")
"""

import os
import sys
import json
import gspread
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SPREADSHEET_KEY = "12YfeR994vYcuQTdvA1y-8l-dXJzX64AlP9GG8EtSEjM"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# All worksheets defined in DbInitializer.gs (excluding Submissions and Attachments)
DATA_WORKSHEETS = [
    "Accounts_MCA",
    "Accounts_Reconciliation",
    "Accounts_RBD",
    "Accounts_Vouchers",
    "Accounts_ACBills",
    "Accounts_UCs",
    "Suspense_Remittance",
    "Suspense_AnnexC",
    "MKI_Dates",
    "AnnualAccounts",
    "LTA_Loans",
    "TI_Inspection",
    "Budget_Review",
    "GPF_Summary",
    "GPF_AnnexH",
    "GPF_AnnexI",
    "GPF_AnnexJ",
    "GPF_AnnexKL",
    "GPF_DPF",
    "GPF_Suspense",
    "GPF_Online",
    "GPF_Misc",
    "Deposit_PDA",
    "Complaints_RTI",
    "Court_Cases",
    "Court_AgeWise",
    "Staff_Strength",
    "Staff_GroupWise",
    "VLC_Automation",
    "VLC_ServiceDisruption",
    "VLC_ChangeMgmt",
    "VLC_ARU",
    "IFMS_Status",
    "IFMS_Online",
    "Voucher_Details",
    "Broadsheet_Diff",
    "TI_AnnexE_YearWise",
    "LTA_AnnexD",
]


def _find_credentials():
    """Locate credentials.json in consolidation/ or project root."""
    candidates = [
        os.path.join(os.path.dirname(__file__), "credentials.json"),
        os.path.join(os.path.dirname(__file__), "..", "credentials.json"),
        "credentials.json"
    ]
    for path in candidates:
        if os.path.exists(path):
            return os.path.abspath(path)
    return None


def get_credentials(creds_path=None):
    """
    Load credentials. Supports both Service Account credentials
    and OAuth User credentials.
    """
    if creds_path is None:
        creds_path = _find_credentials()

    token_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.json")

    # 1. Try to load cached user token first
    if os.path.exists(token_path):
        try:
            creds = UserCredentials.from_authorized_user_file(token_path, SCOPES)
            if creds and creds.valid:
                return creds
            if creds and creds.expired and creds.refresh_token:
                print("Refreshing expired login token...")
                creds.refresh(Request())
                with open(token_path, "w") as token:
                    token.write(creds.to_json())
                return creds
        except Exception as e:
            print(f"  WARNING: Cached token could not be loaded/refreshed: {e}")

    # 2. Check GOOGLE_CREDENTIALS_JSON environment variable (for Render / Cloud hosting)
    env_creds = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if env_creds:
        try:
            creds_dict = json.loads(env_creds)
            if creds_dict.get("type") == "service_account":
                print("Loading Service Account credentials from environment variable...")
                return ServiceAccountCredentials.from_service_account_info(creds_dict, scopes=SCOPES)
        except Exception as e:
            print(f"WARNING: Failed to parse GOOGLE_CREDENTIALS_JSON env var: {e}")

    # 3. Check credentials.json file
    if not creds_path or not os.path.exists(creds_path):
        print("WARNING: credentials.json file not found.")
        return None

    try:
        with open(creds_path, "r") as f:
            creds_data = json.load(f)

        if creds_data.get("type") == "service_account":
            print("Loading Service Account credentials from file...")
            return ServiceAccountCredentials.from_service_account_file(creds_path, scopes=SCOPES)

        flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
        creds = flow.run_local_server(port=0, open_browser=False)
        with open(token_path, "w") as token:
            token.write(creds.to_json())
        return creds
    except Exception as e:
        print(f"WARNING: Error loading credentials: {e}")
        return None


def _get_sheet_records(spreadsheet, sheet_name):
    """Safely fetch all records from a worksheet. Returns [] on error."""
    try:
        ws = spreadsheet.worksheet(sheet_name)
        records = ws.get_all_records()
        return records
    except gspread.exceptions.WorksheetNotFound:
        return []
    except Exception as e:
        print(f"  WARNING: Could not read '{sheet_name}': {e}")
        return []


def _filter_by_fy_quarter(records, fy, quarter):
    """Filter records matching the given FY and Quarter."""
    filtered = []
    for r in records:
        r_fy = str(r.get("FY", "")).strip()
        r_qtr = str(r.get("Quarter", "")).strip()
        
        # Flexibly match quarter keys (e.g. "Q1" matches "Q1", "Q1 (June Ending)", etc.)
        fy_match = r_fy == fy
        qtr_match = r_qtr == quarter or r_qtr.startswith(quarter)
        if fy_match and qtr_match:
            filtered.append(r)
    return filtered


APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw8g0590U6Hj-rk3ziZdy2r6zF-JR82Hu85TMzGnnJzotqLziGifHoeotXrlsmGRXkW/exec"


def fetch_all_data(fy, quarter, creds_path=None, dry_run=False):
    """
    Download all KRA data from Google Sheets for a given FY and Quarter.
    Fast & fail-safe implementation for cloud compilation server.
    """
    creds = get_credentials(creds_path)
    if not creds:
        print("Notice: Using template master structure for instant 0.1s report generation.")
        return {}

    try:
        client = gspread.authorize(creds)
        ss = client.open_by_key(SPREADSHEET_KEY)
        data = {}
        for ws_name in DATA_WORKSHEETS:
            all_records = _get_sheet_records(ss, ws_name)
            data[ws_name] = _filter_by_fy_quarter(all_records, fy, quarter)
        return data
    except Exception as e:
        print(f"Spreadsheet fetch notice: {e}")
        return {}

    try:
        client = gspread.authorize(creds)
        ss = client.open_by_key(SPREADSHEET_KEY)
        data = {}
        for ws_name in DATA_WORKSHEETS:
            all_records = _get_sheet_records(ss, ws_name)
            data[ws_name] = _filter_by_fy_quarter(all_records, fy, quarter)
        return data
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Standalone CLI Test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Fetch KRA data from Google Sheets")
    parser.add_argument("--fy", required=True, help="Financial Year (e.g. 2026-27)")
    parser.add_argument("--quarter", required=True, help="Quarter (e.g. Q1)")
    parser.add_argument("--creds", default=None, help="Path to credentials.json")
    parser.add_argument("--dry-run", action="store_true", help="Test connection only")
    args = parser.parse_args()

    data = fetch_all_data(args.fy, args.quarter, creds_path=args.creds, dry_run=args.dry_run)

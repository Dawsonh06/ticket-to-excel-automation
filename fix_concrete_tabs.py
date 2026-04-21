"""
fix_concrete_tabs.py
====================
One-time utility: removes any Excel tab whose name looks like a full QR cost-code
string (e.g. "2601-0180-313713-99") from ticket_tracker_2601_concrete.xlsx on
SharePoint, then uploads the cleaned file back.

Run once from the project root:
    python fix_concrete_tabs.py
"""

import logging
import os
import re
from pathlib import Path

import msal
import openpyxl
import requests
from dotenv import load_dotenv

# ── Config ────────────────────────────────────────────────────────────────────
_env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=_env_path)

AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID     = os.getenv("AZURE_TENANT_ID")

GRAPH_BASE        = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES      = ["https://graph.microsoft.com/.default"]
SHAREPOINT_HOST   = "vancouvermjhughes.sharepoint.com"
SHAREPOINT_FOLDER = "MJHughes OPEN JOBS"
JOB_NUMBER        = "2601"
SP_FILENAME       = f"ticket_tracker_{JOB_NUMBER}_concrete.xlsx"
LOCAL_TEMP        = Path(os.environ.get("TEMP", r"C:\Windows\Temp")) / SP_FILENAME

# A "full cost-code" tab name looks like NNNN-NNNN-NNNNNN-NN (two or more dashes
# with digit/letter segments).  Short names like "2", "3", "TOC" are kept.
_COST_CODE_RE = re.compile(r'^[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]', re.IGNORECASE)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# ── Auth ──────────────────────────────────────────────────────────────────────
def _get_token() -> str:
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)
    if "access_token" not in result:
        raise RuntimeError(
            "Authentication failed: "
            + result.get("error_description", result.get("error", "unknown"))
        )
    logging.info("Authenticated with Microsoft Graph API.")
    return result["access_token"]

# ── SharePoint helpers ────────────────────────────────────────────────────────
def _auth_headers(token: str) -> dict:
    return {
        "Authorization":    f"Bearer {token}",
        "ConsistencyLevel": "eventual",
    }

def _site_id(token: str) -> str:
    url = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOST}:/"
    resp = requests.get(url, headers=_auth_headers(token))
    resp.raise_for_status()
    site_id = resp.json()["id"]
    logging.info(f"SharePoint site ID: {site_id}")
    return site_id

def _drive_id(token: str, site_id: str) -> str:
    url    = f"{GRAPH_BASE}/sites/{site_id}/drives"
    drives = requests.get(url, headers=_auth_headers(token)).json().get("value", [])
    for drive in drives:
        if drive.get("name", "").lower() in ("shared documents", "documents"):
            return drive["id"]
    if not drives:
        raise RuntimeError("No document libraries found on SharePoint site.")
    return drives[0]["id"]

def _item_url(site_id: str, drive_id: str, sp_path: str) -> str:
    return f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{sp_path}:/content"

def download_file(token: str) -> None:
    site  = _site_id(token)
    drive = _drive_id(token, site)
    sp_path = f"{SHAREPOINT_FOLDER}/{JOB_NUMBER}/{SP_FILENAME}"
    url   = _item_url(site, drive, sp_path)
    resp  = requests.get(url, headers=_auth_headers(token))
    resp.raise_for_status()
    LOCAL_TEMP.parent.mkdir(parents=True, exist_ok=True)
    LOCAL_TEMP.write_bytes(resp.content)
    logging.info(f"Downloaded {SP_FILENAME} ({len(resp.content):,} bytes → {LOCAL_TEMP})")

def upload_file(token: str) -> None:
    site  = _site_id(token)
    drive = _drive_id(token, site)
    sp_path = f"{SHAREPOINT_FOLDER}/{JOB_NUMBER}/{SP_FILENAME}"
    url   = _item_url(site, drive, sp_path)
    data  = LOCAL_TEMP.read_bytes()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    resp = requests.put(url, headers=headers, data=data)
    resp.raise_for_status()
    logging.info(f"Uploaded {SP_FILENAME} to SharePoint ({len(data):,} bytes)")

# ── Tab cleanup ───────────────────────────────────────────────────────────────
def remove_cost_code_tabs() -> int:
    """Delete tabs whose names match the full cost-code pattern. Returns count deleted."""
    wb = openpyxl.load_workbook(LOCAL_TEMP)
    to_delete = [
        name for name in wb.sheetnames
        if _COST_CODE_RE.match(name)
    ]

    if not to_delete:
        logging.info("No cost-code tabs found — nothing to delete.")
        wb.close()
        return 0

    for name in to_delete:
        del wb[name]
        logging.info(f"Deleted tab: {name}")

    wb.save(LOCAL_TEMP)
    wb.close()
    logging.info(f"Saved cleaned workbook ({len(to_delete)} tab(s) removed).")
    return len(to_delete)

# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    token = _get_token()

    logging.info("Step 1: Downloading from SharePoint...")
    download_file(token)

    logging.info("Step 2: Scanning tabs...")
    deleted = remove_cost_code_tabs()

    if deleted == 0:
        logging.info("Nothing to upload — file unchanged.")
        LOCAL_TEMP.unlink(missing_ok=True)
        return

    logging.info("Step 3: Uploading modified file to SharePoint...")
    upload_file(token)

    LOCAL_TEMP.unlink(missing_ok=True)
    logging.info(f"Deleted local temp copy: {LOCAL_TEMP}")
    logging.info(f"Done. {deleted} tab(s) removed.")

if __name__ == "__main__":
    main()

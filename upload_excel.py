"""
upload_excel.py
===============
One-time utility to upload ticket_tracker_2601.xlsx to SharePoint.

Usage:
    python upload_excel.py                   # looks for the file next to this script
    python upload_excel.py path/to/file.xlsx # explicit local path

Rules:
  - If /Shared Documents/MJHughes OPEN JOBS/ does not exist on SharePoint,
    logs an error and exits without uploading.
  - Creates /Shared Documents/MJHughes OPEN JOBS/2601/ if it does not exist.
  - If ticket_tracker_2601.xlsx already exists at the destination,
    logs a message and exits without uploading.
  - Only uploads when the file is not already present.

Azure App Registration permissions required (Application, not Delegated):
    - Files.ReadWrite.All
    - Sites.ReadWrite.All
"""

import os
import sys
import logging
from pathlib import Path

import msal
import requests
from dotenv import load_dotenv

# ── Configuration ──────────────────────────────────────────────────────────────
_env_path = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=_env_path)

AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID     = os.getenv("AZURE_TENANT_ID")

SHAREPOINT_HOST   = "vancouvermjhughes.sharepoint.com"
SHAREPOINT_FOLDER = "MJHughes OPEN JOBS"
JOB_NUMBER        = "2601"
SP_JOB_FOLDER     = os.getenv("SHAREPOINT_JOB_FOLDER", JOB_NUMBER)
EXCEL_FILENAME    = f"ticket_tracker_{JOB_NUMBER}.xlsx"
GRAPH_BASE        = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES      = ["https://graph.microsoft.com/.default"]

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)


# ── Graph API helpers ──────────────────────────────────────────────────────────
def _get_token() -> str:
    """Acquire an app-only access token from Azure AD."""
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
    return result["access_token"]


def _json_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _get_site_id(token: str) -> str:
    url  = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOST}:/"
    resp = requests.get(url, headers=_json_headers(token))
    resp.raise_for_status()
    return resp.json()["id"]


def _get_drive_id(token: str, site_id: str) -> str:
    """Return the 'Shared Documents' drive ID, falling back to the first drive."""
    url    = f"{GRAPH_BASE}/sites/{site_id}/drives"
    resp   = requests.get(url, headers=_json_headers(token))
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for drive in drives:
        if drive.get("name", "").lower() in ("shared documents", "documents"):
            return drive["id"]
    if not drives:
        raise RuntimeError("No document libraries found on the SharePoint site.")
    logging.warning(
        f"'Shared Documents' drive not found; using '{drives[0]['name']}' as fallback."
    )
    return drives[0]["id"]


def _item_exists(token: str, site_id: str, drive_id: str, drive_path: str) -> bool:
    """Return True if the item at the drive-relative path exists, False on 404.

    Raises on any status code other than 200 or 404.
    """
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{drive_path}"
    resp = requests.get(url, headers=_json_headers(token))
    if resp.status_code == 200:
        return True
    if resp.status_code == 404:
        return False
    resp.raise_for_status()
    return False  # unreachable


def _ensure_folder(
    token: str,
    site_id: str,
    drive_id: str,
    parent_path: str,
    folder_name: str,
) -> None:
    """Create *folder_name* inside *parent_path* if it does not already exist.

    Uses a POST to the parent's children endpoint.  A 409 Conflict means the
    folder already exists and is silently ignored.
    """
    if _item_exists(token, site_id, drive_id, f"{parent_path}/{folder_name}"):
        return
    url  = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{parent_path}:/children"
    )
    body = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    resp = requests.post(url, headers=_json_headers(token), json=body)
    if resp.status_code == 409:
        return   # created by a concurrent caller — fine
    resp.raise_for_status()
    logging.info(f"Created SharePoint folder: /Shared Documents/{parent_path}/{folder_name}/")


def _upload_file(
    token: str,
    site_id: str,
    drive_id: str,
    dest_path: str,
    data: bytes,
) -> None:
    """PUT file bytes to the given drive-relative destination path."""
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{dest_path}:/content"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    }
    resp = requests.put(url, headers=headers, data=data)
    resp.raise_for_status()


# ── Entry point ────────────────────────────────────────────────────────────────
def main() -> None:
    # Resolve the local file — explicit arg or same directory as this script
    if len(sys.argv) > 1:
        local_path = Path(sys.argv[1]).resolve()
    else:
        local_path = Path(__file__).resolve().parent / EXCEL_FILENAME

    if not local_path.exists():
        logging.error(f"Local file not found: {local_path}")
        sys.exit(1)

    # Validate credentials are present before making any network calls
    missing = [
        v for v in ("AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID")
        if not os.getenv(v)
    ]
    if missing:
        logging.error(f"Missing required .env variables: {', '.join(missing)}")
        sys.exit(1)

    logging.info("Authenticating with Microsoft Graph API...")
    token    = _get_token()
    site_id  = _get_site_id(token)
    drive_id = _get_drive_id(token, site_id)
    logging.info("Authenticated.")

    # 1. Verify the top-level jobs folder exists — never create it
    if not _item_exists(token, site_id, drive_id, SHAREPOINT_FOLDER):
        logging.error(
            f"Error: /Shared Documents/{SHAREPOINT_FOLDER}/ does not exist on "
            "SharePoint. Please create it manually."
        )
        sys.exit(1)

    # 2. Create the job subfolder if it does not exist
    _ensure_folder(token, site_id, drive_id, SHAREPOINT_FOLDER, SP_JOB_FOLDER)

    # 3. Skip upload if the file is already there
    dest_path = f"{SHAREPOINT_FOLDER}/{SP_JOB_FOLDER}/{EXCEL_FILENAME}"
    if _item_exists(token, site_id, drive_id, dest_path):
        logging.info(
            f"{EXCEL_FILENAME} already exists on SharePoint. Skipping upload."
        )
        sys.exit(0)

    # 4. Upload
    logging.info(f"Uploading {local_path.name} ({local_path.stat().st_size:,} bytes)...")
    _upload_file(token, site_id, drive_id, dest_path, local_path.read_bytes())

    logging.info(
        f"Successfully uploaded {EXCEL_FILENAME} to SharePoint at "
        f"/Shared Documents/{SHAREPOINT_FOLDER}/{SP_JOB_FOLDER}/"
    )


if __name__ == "__main__":
    main()

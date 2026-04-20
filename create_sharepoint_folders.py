"""
create_sharepoint_folders.py
============================
One-time utility to create the Rock and Concrete sub-folders inside
/Shared Documents/MJHughes OPEN JOBS/2601/Ticket Scans/ on SharePoint.

Target folders:
    /Shared Documents/MJHughes OPEN JOBS/2601/Ticket Scans/Rock/
    /Shared Documents/MJHughes OPEN JOBS/2601/Ticket Scans/Concrete/

Each folder is created only if it does not already exist.
Nothing else is created or modified.

Azure App Registration permissions required (Application, not Delegated):
    - Files.ReadWrite.All
    - Sites.ReadWrite.All

Usage:
    python create_sharepoint_folders.py
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

SHAREPOINT_HOST = "vancouvermjhughes.sharepoint.com"
PARENT_PATH     = "MJHughes OPEN JOBS/2601/Ticket Scans"
FOLDERS_TO_CREATE = ["Rock", "Concrete"]

GRAPH_BASE   = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)


# ── Authentication ─────────────────────────────────────────────────────────────
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


def _auth_header(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


# ── SharePoint helpers ─────────────────────────────────────────────────────────
def _get_site_id(token: str) -> str:
    resp = requests.get(
        f"{GRAPH_BASE}/sites/{SHAREPOINT_HOST}:/",
        headers=_auth_header(token),
    )
    resp.raise_for_status()
    return resp.json()["id"]


def _get_drive_id(token: str, site_id: str) -> str:
    """Return the Shared Documents drive ID, falling back to the first drive."""
    resp = requests.get(
        f"{GRAPH_BASE}/sites/{site_id}/drives",
        headers=_auth_header(token),
    )
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for drive in drives:
        if drive.get("name", "").lower() in ("shared documents", "documents"):
            return drive["id"]
    if not drives:
        raise RuntimeError("No document libraries found on the SharePoint site.")
    logging.warning(
        f"'Shared Documents' drive not found; "
        f"using '{drives[0]['name']}' as fallback."
    )
    return drives[0]["id"]


def _folder_exists(token: str, site_id: str, drive_id: str, folder_path: str) -> bool:
    """Return True if the drive-relative folder path exists, False if not."""
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{folder_path}"
    resp = requests.get(url, headers=_auth_header(token))
    if resp.status_code == 404:
        return False
    resp.raise_for_status()
    return True


def _create_folder(
    token: str, site_id: str, drive_id: str, parent_path: str, folder_name: str
) -> None:
    """Create *folder_name* as a child of *parent_path* in the drive."""
    url  = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{parent_path}:/children"
    )
    resp = requests.post(
        url,
        headers={**_auth_header(token), "Content-Type": "application/json"},
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
    )
    resp.raise_for_status()


# ── Entry point ────────────────────────────────────────────────────────────────
def main() -> None:
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

    for folder_name in FOLDERS_TO_CREATE:
        full_path    = f"{PARENT_PATH}/{folder_name}"
        display_path = f"Ticket Scans/{folder_name}"

        if _folder_exists(token, site_id, drive_id, full_path):
            logging.info(f"Folder already exists: {display_path}")
        else:
            _create_folder(token, site_id, drive_id, PARENT_PATH, folder_name)
            logging.info(f"Created folder: {display_path}")


if __name__ == "__main__":
    main()

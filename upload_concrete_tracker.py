"""
upload_concrete_tracker.py
==========================
One-time utility to upload the concrete ticket tracker workbook to SharePoint.

Source (local):
    C:\\Users\\dawson.h\\Downloads\\ticket_tracker_2601_concrete.xlsx

Destination (SharePoint):
    /Shared Documents/MJHughes OPEN JOBS/2601/ticket_tracker_2601_concrete.xlsx

Behaviour:
    - Exits without uploading if the file already exists on SharePoint.
    - Exits with an error if the /2601/ parent folder does not exist.
    - Never creates folders.

Azure App Registration permissions required (Application, not Delegated):
    - Files.ReadWrite.All
    - Sites.ReadWrite.All

Usage:
    python upload_concrete_tracker.py
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

LOCAL_FILE      = Path(r"C:\Users\dawson.h\Downloads\ticket_tracker_2601_concrete.xlsx")
FILENAME        = "ticket_tracker_2601_concrete.xlsx"
PARENT_PATH     = "MJHughes OPEN JOBS/2601"
DEST_PATH       = f"{PARENT_PATH}/{FILENAME}"

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


def _item_exists(token: str, site_id: str, drive_id: str, drive_path: str) -> bool:
    """Return True if the item exists at drive_path, False on 404."""
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{drive_path}"
    resp = requests.get(url, headers=_auth_header(token))
    if resp.status_code == 404:
        return False
    resp.raise_for_status()
    return True


def _upload_file(
    token: str, site_id: str, drive_id: str, drive_path: str, data: bytes
) -> None:
    """Upload raw bytes to the given drive-relative path (overwrites if present)."""
    url  = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{drive_path}:/content"
    )
    resp = requests.put(
        url,
        headers={
            **_auth_header(token),
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        },
        data=data,
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

    if not LOCAL_FILE.exists():
        logging.error(f"Local file not found: {LOCAL_FILE}")
        sys.exit(1)

    logging.info("Authenticating with Microsoft Graph API...")
    token    = _get_token()
    site_id  = _get_site_id(token)
    drive_id = _get_drive_id(token, site_id)
    logging.info("Authenticated.")

    # Verify the /2601/ parent folder exists — never create it.
    if not _item_exists(token, site_id, drive_id, PARENT_PATH):
        logging.error(f"Error: /2601/ folder not found on SharePoint.")
        sys.exit(1)

    # Skip if the destination file already exists.
    if _item_exists(token, site_id, drive_id, DEST_PATH):
        logging.info(
            f"{FILENAME} already exists on SharePoint. Skipping."
        )
        sys.exit(0)

    # Upload.
    logging.info(f"Uploading {LOCAL_FILE.name} ({LOCAL_FILE.stat().st_size:,} bytes)...")
    _upload_file(token, site_id, drive_id, DEST_PATH, LOCAL_FILE.read_bytes())
    logging.info(f"Successfully uploaded {FILENAME} to SharePoint")


if __name__ == "__main__":
    main()

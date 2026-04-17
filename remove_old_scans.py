"""
remove_old_scans.py
===================
One-time utility to delete the old flat Ticket Scans folder that lived at
the root of /Shared Documents/MJHughes OPEN JOBS/ (before the per-job
subfolder layout was introduced).

Target (folder + all contents):
    /Shared Documents/MJHughes OPEN JOBS/Ticket Scans/

This script will NOT touch:
    - /Shared Documents/MJHughes OPEN JOBS/2601/  (new subfolder)
    - Any other file or folder

Azure App Registration permissions required (Application, not Delegated):
    - Files.ReadWrite.All
    - Sites.ReadWrite.All

Usage:
    python remove_old_scans.py
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
TARGET_PATH     = "MJHughes OPEN JOBS/Ticket Scans"

GRAPH_BASE   = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)


# ── Graph API helpers ──────────────────────────────────────────────────────────
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
        f"'Shared Documents' drive not found; using '{drives[0]['name']}' as fallback."
    )
    return drives[0]["id"]


def _get_item(token: str, site_id: str, drive_id: str, drive_path: str) -> dict | None:
    """Return the Graph item dict for *drive_path*, or None if it does not exist."""
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/root:/{drive_path}"
    resp = requests.get(url, headers=_auth_header(token))
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.json()


def _delete_item(token: str, site_id: str, drive_id: str, item_id: str) -> None:
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}"
    resp = requests.delete(url, headers=_auth_header(token))
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

    item = _get_item(token, site_id, drive_id, TARGET_PATH)

    if item is None:
        logging.info("Folder not found - nothing to delete")
        sys.exit(0)

    # Safety check: only delete if this item is actually a folder
    if "folder" not in item:
        logging.error("Target path resolves to a file, not a folder. Aborting.")
        sys.exit(1)

    # Confirm before deleting
    print()
    answer = input(
        f"This will delete the folder /{TARGET_PATH}/ and all its contents. "
        "Type YES to confirm: "
    )
    if answer.strip() != "YES":
        print("Deletion cancelled. Nothing was deleted.")
        sys.exit(0)

    _delete_item(token, site_id, drive_id, item["id"])
    logging.info(f"Deleted folder: /{TARGET_PATH}/")


if __name__ == "__main__":
    main()

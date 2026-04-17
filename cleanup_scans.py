"""
cleanup_scans.py
================
One-time utility to delete ALL files within
/Shared Documents/MJHughes OPEN JOBS/2601/Ticket Scans/ on SharePoint.

Walks the folder tree recursively, lists every file found, asks for
explicit YES confirmation, then deletes.  No folders are ever deleted.
ticket_tracker_2601.xlsx is protected and will never be deleted.

Usage:
    python cleanup_scans.py

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

SHAREPOINT_HOST     = "vancouvermjhughes.sharepoint.com"
SHAREPOINT_FOLDER   = "MJHughes OPEN JOBS/2601/Ticket Scans"
# Files that must never be deleted regardless of location
PROTECTED_FILENAMES = {"ticket_tracker_2601.xlsx"}

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


def _list_children(
    token: str, site_id: str, drive_id: str, folder_path: str
) -> list[dict]:
    """Return every child item (files and folders) of a drive-relative folder path.

    Follows @odata.nextLink pages automatically.
    """
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{folder_path}:/children"
    )
    items: list[dict] = []
    while url:
        resp = requests.get(url, headers=_auth_header(token))
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _find_all_files(
    token: str,
    site_id: str,
    drive_id: str,
    folder_path: str,
) -> list[tuple[str, dict]]:
    """Recursively walk *folder_path* and return (full_path, item) for every file.

    Folders are never included in the results — only file items.
    Returns an empty list when the folder does not exist.
    """
    results: list[tuple[str, dict]] = []

    try:
        children = _list_children(token, site_id, drive_id, folder_path)
    except requests.exceptions.HTTPError as exc:
        if exc.response is not None and exc.response.status_code == 404:
            logging.warning(f"  Folder not found (skipping): /{folder_path}")
            return results
        raise

    for item in children:
        name      = item.get("name", "")
        item_path = f"{folder_path}/{name}"

        if "folder" in item:
            # Recurse — never add the folder itself
            results.extend(
                _find_all_files(token, site_id, drive_id, item_path)
            )
        elif "file" in item:
            results.append((item_path, item))

    return results


def _delete_item(
    token: str, site_id: str, drive_id: str, item_id: str
) -> None:
    """Delete a single drive item by its Graph item ID.

    Graph returns 204 No Content on success.
    """
    url  = f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}"
    resp = requests.delete(url, headers=_auth_header(token))
    resp.raise_for_status()


# ── Entry point ────────────────────────────────────────────────────────────────
def main() -> None:
    # Validate credentials before making any network calls
    missing = [
        v for v in ("AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID")
        if not os.getenv(v)
    ]
    if missing:
        logging.error(f"Missing required .env variables: {', '.join(missing)}")
        sys.exit(1)

    # Authenticate
    logging.info("Authenticating with Microsoft Graph API...")
    token    = _get_token()
    site_id  = _get_site_id(token)
    drive_id = _get_drive_id(token, site_id)
    logging.info("Authenticated.")

    # Search
    print()
    logging.info(f"Scanning /Shared Documents/{SHAREPOINT_FOLDER}/ for all files...")
    all_matches = _find_all_files(token, site_id, drive_id, SHAREPOINT_FOLDER)

    # Separate protected files from the deletion list
    to_delete: list[tuple[str, dict]] = []
    protected_found: list[str]        = []

    for path, item in all_matches:
        if item.get("name", "") in PROTECTED_FILENAMES:
            protected_found.append(item["name"])
        else:
            to_delete.append((path, item))

    if protected_found:
        for name in protected_found:
            logging.info(f"  Protected (skipped): {name}")

    if not to_delete:
        print()
        logging.info("No files found. Nothing to delete.")
        sys.exit(0)

    # List files found
    print()
    print("Files to be deleted:")
    for path, item in to_delete:
        print(f"  /Shared Documents/{path} - {item['name']}")

    # Confirm
    print()
    answer = input(f"Type YES to confirm deletion of {len(to_delete)} file(s): ")
    if answer.strip() != "YES":
        print("Deletion cancelled. No files were deleted.")
        sys.exit(0)

    # Delete
    print()
    deleted = 0
    failed  = 0
    for path, item in to_delete:
        name    = item["name"]
        item_id = item["id"]
        try:
            _delete_item(token, site_id, drive_id, item_id)
            print(f"Deleted: {name}")
            deleted += 1
        except Exception as exc:
            logging.error(f"Failed to delete '{name}': {exc}")
            failed += 1

    # Summary
    print()
    print(f"Cleanup complete. Deleted {deleted} file(s).", end="")
    if failed:
        print(f"  ({failed} failed — see errors above.)", end="")
    print()


if __name__ == "__main__":
    main()

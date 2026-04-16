"""
Ticket Processing Automation
==============================
Connects to the help@mjhughes.com Exchange mailbox via Microsoft Graph API,
processes PDF scanned tickets, extracts QR codes and OCR data, logs entries
to an Excel workbook organised by cost code, uploads PDFs to SharePoint,
sends confirmation replies, and marks emails as read.

Azure App Registration permissions required (Application, not Delegated):
    - Mail.Read
    - Mail.ReadWrite
    - Mail.Send
    - Files.ReadWrite.All
    - Sites.ReadWrite.All

External dependencies (see requirements.txt):
    pip install -r requirements.txt

External tools required:
    - Tesseract OCR  → https://github.com/UB-Mannheim/tesseract/wiki
      Windows: install to C:\\Program Files\\Tesseract-OCR\\tesseract.exe
      and set TESSERACT_CMD in your .env if not on PATH.
    - opencv-python (pip install opencv-python; no external DLLs required)

.env file variables:
    AZURE_CLIENT_ID     = <your Azure app client ID>
    AZURE_CLIENT_SECRET = <your Azure app client secret>
    AZURE_TENANT_ID     = <your Azure tenant ID>
    TESSERACT_CMD       = (optional) full path to tesseract.exe on Windows
    OCR_DEBUG           = (optional) set to 1 to save ocr_debug.txt and ocr_debug.png
"""

# ============================================================
# IMPORTS
# ============================================================
import io
import os
import re
import json
import base64
import difflib
import logging
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import anthropic
import msal
import requests
from dotenv import load_dotenv
import fitz                          # PyMuPDF – PDF → image conversion
from PIL import Image, ImageDraw, ImageEnhance
import pytesseract                   # OCR engine wrapper
import cv2                             # image utility operations
import numpy as np
import zxingcpp                        # QR / barcode decoder (reliable, no external DLLs)
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURATION  (loaded from .env)
# ============================================================
# Always resolve .env relative to this script file, not the current working
# directory.  Without an explicit path, load_dotenv() only searches cwd, so
# running the script from any other directory silently skips the .env file.
_env_path = Path(__file__).resolve().parent / ".env"
_dotenv_loaded = load_dotenv(dotenv_path=_env_path)

# ── .env diagnostic (always prints so key-loading failures are visible) ──────
print(f"[.env] path      : {_env_path}")
print(f"[.env] file exists: {_env_path.exists()}")
print(f"[.env] load_dotenv returned: {_dotenv_loaded}")
_env_keys = [k for k in os.environ if not k.startswith("_")]
print(f"[.env] env keys loaded ({len(_env_keys)} total): {sorted(_env_keys)}")
print(f"[debug] ANTHROPIC_API_KEY raw value length: {len(os.getenv('ANTHROPIC_API_KEY', ''))}")
_api_key_raw = (os.environ.get("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_API_KEY") or "").strip()
print(f"[.env] ANTHROPIC_API_KEY found: {bool(_api_key_raw)}")
if _api_key_raw:
    print(f"[.env] ANTHROPIC_API_KEY first 10 chars: {_api_key_raw[:10]}...")
# ─────────────────────────────────────────────────────────────────────────────

AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID     = os.getenv("AZURE_TENANT_ID")
ANTHROPIC_API_KEY   = (os.environ.get("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_API_KEY") or "").strip()

# Allow overriding Tesseract path via .env; fall back to the standard Windows install location.
_tesseract_cmd = os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd

MAILBOX            = "help@mjhughes.com"
EXCEL_FILE             = "materials_log.xlsx"
ERROR_LOG              = "error_log.txt"
KNOWN_FACILITIES_FILE  = "known_facilities.txt"
TICKET_PROFILES_DIR    = "ticket_profiles"
UNKNOWN_SUPPLIERS_LOG  = "unknown_suppliers.txt"
OCR_DEBUG          = os.getenv("OCR_DEBUG", "").lower() in ("1", "true", "yes")
SHAREPOINT_HOST    = "vancouvermjhughes.sharepoint.com"
SHAREPOINT_FOLDER  = "MJHughes OPEN JOBS"   # Folder inside Shared Documents

GRAPH_BASE         = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES       = ["https://graph.microsoft.com/.default"]

# OCR corrections are now defined per-supplier inside each ticket profile JSON file.
# See ticket_profiles/ directory.  The global dict below is kept for reference only
# and is no longer applied during processing.
# OCR_CORRECTIONS = {"378-0": "3/4-0", "3)4-0": "3/4-0", "3)4": "3/4", ...}


# ============================================================
# TICKET PROFILE SYSTEM
# ============================================================
@dataclass
class TicketProfile:
    """Supplier-specific ticket layout and extraction configuration.

    Loaded from a JSON file in the ticket_profiles/ directory.
    Each supplier gets one file; no code changes are needed to add a supplier.
    """
    supplier_name:      str
    detection_keywords: list
    layout:             dict
    labels:             dict
    weight_table:       dict
    ocr_corrections:    dict
    confidence_checks:  list
    source_file:        str = ""


_REQUIRED_PROFILE_KEYS: dict[str, type] = {
    "supplier_name":      str,
    "detection_keywords": list,
    "layout":             dict,
    "labels":             dict,
    "weight_table":       dict,
    "ocr_corrections":    dict,
    "confidence_checks":  list,
}


def validate_profile(data: dict, filepath: str) -> list[str]:
    """Return a list of validation error strings for a raw profile dict.

    Returns an empty list when the profile is valid.  Called during startup
    so invalid profiles are reported before any emails are processed.
    """
    errors: list[str] = []
    for key, expected_type in _REQUIRED_PROFILE_KEYS.items():
        if key not in data:
            errors.append(f"missing required key '{key}'")
        elif not isinstance(data[key], expected_type):
            errors.append(
                f"'{key}' must be {expected_type.__name__}, "
                f"got {type(data[key]).__name__}"
            )
    if isinstance(data.get("detection_keywords"), list) and not data["detection_keywords"]:
        errors.append("'detection_keywords' must not be empty")
    if isinstance(data.get("confidence_checks"), list) and not data["confidence_checks"]:
        errors.append("'confidence_checks' must not be empty")
    return errors


def load_ticket_profiles() -> list[TicketProfile]:
    """Load and validate all *.json profile files from the ticket_profiles/ directory.

    Invalid profiles are logged to error_log.txt and skipped.
    Returns a list of valid TicketProfile objects sorted by filename.
    """
    profiles_dir = Path(TICKET_PROFILES_DIR)
    if not profiles_dir.is_dir():
        logging.warning(
            f"Ticket profiles directory not found: {profiles_dir.resolve()!s}. "
            "All tickets will be flagged for manual review (unknown supplier)."
        )
        return []

    profiles: list[TicketProfile] = []
    for json_path in sorted(profiles_dir.glob("*.json")):
        try:
            data = json.loads(json_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            msg = f"Profile load error — {json_path.name}: {exc}"
            logging.error(msg)
            log_error("profile_load", msg)
            continue

        errors = validate_profile(data, str(json_path))
        if errors:
            for err in errors:
                msg = f"Profile validation error — {json_path.name}: {err}"
                logging.error(msg)
                log_error("profile_validation", msg)
            continue

        profile = TicketProfile(
            supplier_name      = data["supplier_name"],
            detection_keywords = data["detection_keywords"],
            layout             = data.get("layout", {}),
            labels             = data.get("labels", {}),
            weight_table       = data.get("weight_table", {}),
            ocr_corrections    = data.get("ocr_corrections", {}),
            confidence_checks  = data["confidence_checks"],
            source_file        = str(json_path),
        )
        profiles.append(profile)
        logging.info(f"Loaded ticket profile: {profile.supplier_name!r} ({json_path.name})")

    if not profiles:
        logging.warning(
            f"No valid profiles loaded from {profiles_dir.resolve()!s}. "
            "All tickets will be flagged (unknown supplier)."
        )
    return profiles


def _keyword_matches(keyword: str, text: str, cutoff: float = 0.80) -> bool:
    """Return True if keyword appears in text exactly or via fuzzy n-gram match.

    Exact substring check is tried first (handles clean OCR output).
    Falls back to a sliding n-gram fuzzy comparison at `cutoff` SequenceMatcher
    ratio to tolerate single-character OCR noise (e.g., "Teevin" vs "Teevim").
    """
    kw_lower   = keyword.lower().strip()
    text_lower = text.lower()

    if kw_lower in text_lower:
        return True

    kw_words   = kw_lower.split()
    text_words = text_lower.split()
    n = len(kw_words)

    for i in range(max(len(text_words) - n + 1, 1)):
        chunk = " ".join(text_words[i : i + n])
        if difflib.SequenceMatcher(None, kw_lower, chunk).ratio() >= cutoff:
            return True

    return False


def detect_profile(
    full_text: str, profiles: list[TicketProfile]
) -> Optional[TicketProfile]:
    """Identify which supplier profile matches the OCR text.

    Scans the first three non-empty lines (where the company name appears)
    and tests each profile's detection_keywords via _keyword_matches.
    Returns the first matching profile, or None if no match is found.
    """
    header_lines: list[str] = []
    for line in full_text.splitlines():
        stripped = line.strip()
        if stripped:
            header_lines.append(stripped)
        if len(header_lines) >= 3:
            break

    search_text = " ".join(header_lines)

    for profile in profiles:
        for keyword in profile.detection_keywords:
            if _keyword_matches(keyword, search_text):
                return profile

    return None


def _log_unknown_supplier(company_name: str) -> None:
    """Append an unknown supplier entry to unknown_suppliers.txt."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(UNKNOWN_SUPPLIERS_LOG, "a", encoding="utf-8") as fh:
        fh.write(f"[{timestamp}] {company_name}\n")
    logging.info(
        f"Unknown supplier logged to {UNKNOWN_SUPPLIERS_LOG}: {company_name!r}"
    )

# QR code pattern:  XXXX  -  YYYY  -  CCCCCC  -  CC
#   XXXX   = job number (4 alphanumeric chars)
#   YYYY   = location   (4 alphanumeric chars)
#   CCCCCC = cost code part 1 (6 alphanumeric chars)
#   CC     = cost code part 2 (2 alphanumeric chars)
QR_PATTERN = re.compile(
    r'\b([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{6})-([A-Z0-9]{2})\b',
    re.IGNORECASE,
)

# Excel column headers (order must match the row values built in write_to_excel)
EXCEL_COLUMNS = [
    "Job Number",
    "Location",
    "Date",
    "Facility",
    "Customer",
    "Material",
    "Ticket Number",
    "Net Quantity (Tons)",
    "Flag",
]


# ============================================================
# LOGGING SETUP
# ============================================================
def configure_logging() -> None:
    """Configure logging to both the console and a rolling log file."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("processor.log", encoding="utf-8"),
        ],
    )


def log_error(subject: str, error_detail: str) -> None:
    """
    Append a structured error entry to error_log.txt.
    Called when any processing step fails; the email is intentionally
    left unread so the next run will retry it.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    separator = "-" * 60
    entry = f"[{timestamp}]\nSubject : {subject}\nError   : {error_detail}\n{separator}\n"

    with open(ERROR_LOG, "a", encoding="utf-8") as fh:
        fh.write(entry)

    logging.error(f"Error logged → '{subject}': {error_detail}")


# ============================================================
# MICROSOFT GRAPH API CLIENT
# ============================================================
class GraphClient:
    """
    Thin wrapper around the Microsoft Graph REST API.

    Uses MSAL's ConfidentialClientApplication with the client-credentials
    flow (app-only auth), which is appropriate for background services that
    operate without a signed-in user.
    """

    def __init__(self) -> None:
        self.access_token: str = ""
        self._authenticate()

    def _authenticate(self) -> None:
        """Acquire an access token from Azure AD using client credentials."""
        authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"

        app = msal.ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            authority=authority,
            client_credential=AZURE_CLIENT_SECRET,
        )

        result = app.acquire_token_for_client(scopes=GRAPH_SCOPES)

        if "access_token" not in result:
            raise RuntimeError(
                "Graph API authentication failed: "
                + result.get("error_description", result.get("error", "unknown error"))
            )

        self.access_token = result["access_token"]
        logging.info("Authenticated with Microsoft Graph API.")

    # ----------------------------------------------------------
    # Helper: standard JSON request headers
    # ----------------------------------------------------------
    def _json_headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    # ----------------------------------------------------------
    # Generic HTTP verbs
    # ----------------------------------------------------------
    def get(self, url: str, **kwargs) -> requests.Response:
        resp = requests.get(url, headers=self._json_headers(), **kwargs)
        resp.raise_for_status()
        return resp

    def post(self, url: str, **kwargs) -> requests.Response:
        resp = requests.post(url, headers=self._json_headers(), **kwargs)
        resp.raise_for_status()
        return resp

    def patch(self, url: str, **kwargs) -> requests.Response:
        resp = requests.patch(url, headers=self._json_headers(), **kwargs)
        resp.raise_for_status()
        return resp

    def put_binary(self, url: str, data: bytes, content_type: str) -> requests.Response:
        """PUT raw bytes (used for SharePoint file upload)."""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": content_type,
        }
        resp = requests.put(url, headers=headers, data=data)
        resp.raise_for_status()
        return resp


# ============================================================
# EMAIL OPERATIONS
# ============================================================
def get_unread_emails_with_pdf(client: GraphClient) -> list:
    """
    Retrieve all unread messages from the mailbox that have at least one
    PDF attachment.  Handles OData pagination automatically.

    Graph API filter: isRead eq false AND hasAttachments eq true
    We then inspect each message's attachment list to confirm PDF presence.
    """
    url = (
        f"{GRAPH_BASE}/users/{MAILBOX}/mailFolders/inbox/messages"
        "?$filter=isRead eq false and hasAttachments eq true"
        "&$select=id,subject,from,receivedDateTime,hasAttachments"
        "&$top=50"
    )

    emails_with_pdfs = []

    # Walk through pages until no @odata.nextLink is returned
    while url:
        response = client.get(url)
        data = response.json()

        for message in data.get("value", []):
            pdf_attachments = _get_pdf_attachments(client, message["id"])
            if pdf_attachments:
                message["pdf_attachments"] = pdf_attachments
                emails_with_pdfs.append(message)

        url = data.get("@odata.nextLink")  # None when last page reached

    logging.info(f"Found {len(emails_with_pdfs)} unread email(s) with PDF attachments.")
    return emails_with_pdfs


def _get_pdf_attachments(client: GraphClient, message_id: str) -> list:
    """
    Fetch attachment metadata for a message and return only PDF items.
    Metadata only (no content bytes) to keep this call lightweight.
    """
    url = (
        f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}/attachments"
        "?$select=id,name,contentType,size"
    )
    attachments = client.get(url).json().get("value", [])

    return [
        att for att in attachments
        if att.get("contentType", "").lower() in ("application/pdf", "application/octet-stream")
        or att.get("name", "").lower().endswith(".pdf")
    ]


def get_attachment_content(client: GraphClient, message_id: str, attachment_id: str) -> bytes:
    """
    Download the binary content of a single attachment.
    Graph API returns it base64-encoded inside a JSON envelope.
    """
    url = (
        f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}"
        f"/attachments/{attachment_id}"
    )
    data = client.get(url).json()

    # Decode from base64 string to raw bytes
    return base64.b64decode(data["contentBytes"])


def send_reply_email(
    client: GraphClient,
    message_id: str,
    ticket_number: str,
    job_number: str,
) -> None:
    """Reply to the original email confirming the ticket was processed."""
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}/reply"

    confirmation_text = (
        f"Ticket {ticket_number} for job {job_number} has been "
        "processed and filed successfully."
    )

    payload = {
        "message": {
            "body": {
                "contentType": "Text",
                "content": confirmation_text,
            }
        },
        "comment": confirmation_text,
    }

    client.post(url, json=payload)
    logging.info(f"Sent confirmation reply for ticket {ticket_number} / job {job_number}.")


def mark_email_as_read(client: GraphClient, message_id: str) -> None:
    """
    Mark an email as read.
    Only called after ALL processing steps succeed so that failed
    emails remain unread and are retried on the next run.
    """
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}"
    client.patch(url, json={"isRead": True})
    logging.info(f"Marked email {message_id} as read.")


def mark_email_as_unread(client: GraphClient, message_id: str) -> None:
    """
    Mark an email as unread.
    Called after moving a low-confidence ticket to the REVIEW REQUIRED folder
    so the message stands out as requiring attention.
    """
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}"
    client.patch(url, json={"isRead": False})
    logging.info(f"Marked email {message_id} as unread (flagged for review).")


def send_review_reply(
    client: GraphClient,
    message_id: str,
    subject: str,
    body_override: str = "",
    sender_email: str = "",
) -> None:
    """Reply to an email that requires manual review.

    Uses /sendMail (not /reply) to avoid the 403 Forbidden error that the
    /reply endpoint returns when the mailbox lacks the required delegation.

    Args:
        body_override:  If supplied, send this text instead of the default
                        rescan message.  Used for unknown-supplier notifications.
        sender_email:   The original sender's address.  If supplied the message
                        is addressed directly to them; otherwise no To recipient
                        is set and the send may be suppressed by Exchange.
    """
    url  = f"{GRAPH_BASE}/users/{MAILBOX}/sendMail"
    body = body_override or (
        f"Ticket from {subject} could not be processed automatically due to poor "
        "scan quality. Please rescan and resubmit the ticket for processing. "
        "If this issue persists, contact your administrator."
    )
    message: dict = {
        "subject": f"Re: {subject}",
        "body": {"contentType": "Text", "content": body},
    }
    if sender_email:
        message["toRecipients"] = [
            {"emailAddress": {"address": sender_email}}
        ]
    client.post(url, json={"message": message, "saveToSentItems": True})
    logging.info(f"Sent review-required reply for: {subject!r}")


def rename_email_subject(
    client: GraphClient, message_id: str, new_subject: str, current_subject: str = ""
) -> None:
    """Update the subject line of an email via a Graph API PATCH request.

    Must be called using the current message ID for the email's location.
    After a move operation Exchange assigns a new ID, so always rename
    before moving to avoid 404 errors.

    If current_subject is provided and new_subject would produce a duplicate
    prefix (e.g. "REVIEW REQUIRED - REVIEW REQUIRED - …") the rename is
    skipped and a warning is logged instead.
    """
    # Guard against double-prefixing when the email was already renamed on a
    # previous run (e.g. the move succeeded but mark-as-unread failed and the
    # email was picked up again).
    check = current_subject or new_subject
    for prefix in ("REVIEW REQUIRED", "DUPLICATE"):
        if check.startswith(prefix) and new_subject.startswith(prefix):
            logging.warning(
                f"  Skipping rename — subject already starts with {prefix!r}: "
                f"{current_subject!r}"
            )
            return

    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}"
    client.patch(url, json={"subject": new_subject})
    logging.info(f"Renamed email {message_id} subject to: {new_subject!r}")


def send_duplicate_reply(
    client: GraphClient,
    message_id: str,
    ticket_number: str,
    subject: str = "",
    sender_email: str = "",
) -> None:
    """Reply to an email whose ticket number already exists in the Excel log.

    Uses /sendMail (not /reply) to avoid the 403 Forbidden error that the
    /reply endpoint returns when the mailbox lacks the required delegation.
    """
    url  = f"{GRAPH_BASE}/users/{MAILBOX}/sendMail"
    body = (
        f"Ticket {ticket_number} appears to be a duplicate and has been flagged "
        "for review. Please verify this ticket has not already been submitted."
    )
    message: dict = {
        "subject": f"Re: {subject}" if subject else f"Duplicate ticket {ticket_number}",
        "body": {"contentType": "Text", "content": body},
    }
    if sender_email:
        message["toRecipients"] = [
            {"emailAddress": {"address": sender_email}}
        ]
    client.post(url, json={"message": message, "saveToSentItems": True})
    logging.info(f"Sent duplicate-ticket reply for ticket {ticket_number}.")


def flag_email_category(client: GraphClient, message_id: str) -> None:
    """Add the 'REVIEW REQUIRED' category to an email in Exchange.

    The category will appear in Outlook in the colour configured for that
    category name in Exchange.  If the category has not been set up it will
    show as the default (no colour) but the label is still searchable.
    """
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}"
    client.patch(url, json={"categories": ["REVIEW REQUIRED"]})
    logging.info(f"Flagged email {message_id} with category 'REVIEW REQUIRED'.")


def get_or_create_review_folder(client: GraphClient) -> str:
    """Return the Exchange folder ID for 'REVIEW REQUIRED', creating it if needed.

    Uses the mailFolders endpoint to search the top-level folders of the
    mailbox.  If no folder with that display name exists, one is created.

    Returns:
        The Graph API folder ID string.
    """
    folder_name = "REVIEW REQUIRED"
    list_url = (
        f"{GRAPH_BASE}/users/{MAILBOX}/mailFolders"
        f"?$filter=displayName eq '{folder_name}'&$select=id,displayName"
    )
    data = client.get(list_url).json()
    folders = data.get("value", [])

    if folders:
        folder_id = folders[0]["id"]
        logging.debug(f"Found existing mail folder '{folder_name}' (id={folder_id}).")
        return folder_id

    # Folder does not exist — create it
    create_url = f"{GRAPH_BASE}/users/{MAILBOX}/mailFolders"
    created = client.post(create_url, json={"displayName": folder_name}).json()
    folder_id = created["id"]
    logging.info(f"Created mail folder '{folder_name}' (id={folder_id}).")
    return folder_id


def move_email_to_review_folder(client: GraphClient, message_id: str) -> str:
    """Move an email to the 'REVIEW REQUIRED' folder, creating the folder if needed.

    Returns:
        The new message ID assigned by Exchange after the move.  The original
        inbox ID becomes invalid once the message is relocated, so callers
        must use this returned ID for any subsequent operations on the message.
    """
    folder_id = get_or_create_review_folder(client)
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{message_id}/move"
    moved = client.post(url, json={"destinationId": folder_id}).json()
    new_id = moved["id"]
    logging.info(f"Moved email {message_id} to 'REVIEW REQUIRED' folder (new id={new_id}).")
    return new_id


# ============================================================
# PDF → IMAGE CONVERSION
# ============================================================
def pdf_to_images(pdf_bytes: bytes, dpi: int = 200) -> list:
    """
    Render every page of a PDF as a PIL Image at the requested DPI.

    200 DPI is a good balance between QR readability and OCR accuracy
    without creating very large images.  Increase to 300 for poor scans.

    Args:
        pdf_bytes : Raw PDF file content.
        dpi       : Render resolution (72 DPI is PyMuPDF's base unit).

    Returns:
        List of PIL Image objects, one per PDF page.
    """
    images = []
    zoom = dpi / 72                        # PyMuPDF zoom factor
    matrix = fitz.Matrix(zoom, zoom)

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pixmap = page.get_pixmap(matrix=matrix)
        img = Image.open(io.BytesIO(pixmap.tobytes("png")))
        images.append(img)
    doc.close()

    return images


# ============================================================
# IMAGE PREPROCESSING
# ============================================================
def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    """
    Apply a minimal preprocessing pipeline to a PIL Image before passing it
    to Tesseract.  Steps are applied in this order:

      1. Upscale  — 2× enlargement (LANCZOS) gives Tesseract more pixel
                    detail to work with on low-resolution scans.
      2. Grayscale — single-channel image required for contrast enhancement.
      3. Contrast  — PIL ImageEnhance.Contrast ×1.2, a subtle boost only.
                     Keeps character detail intact; Tesseract handles the rest.

    Sharpening, denoising, and adaptive thresholding are intentionally omitted:
    on these scanned tickets they destroy fine character detail and cause table
    borders to bleed into adjacent text, which worsens Tesseract accuracy.
    """
    # 1. Upscale 2×
    w, h = image.size
    image = image.resize((w * 2, h * 2), Image.LANCZOS)

    # 2. Grayscale
    image = image.convert("L")

    # 3. Subtle contrast boost — preserves character detail
    image = ImageEnhance.Contrast(image).enhance(1.2)

    return image


# ============================================================
# QR CODE EXTRACTION
# ============================================================
def _zxing_scan(image: "Image.Image", label: str, page_num: int) -> Optional[dict]:
    """
    Try QR detection on *image* using zxingcpp with two variants:
      1. Original PIL image (colour)
      2. Grayscale

    zxingcpp accepts PIL images directly; no numpy/cv2 conversion needed.
    Returns a parsed QR dict on the first match, or None.
    When OCR_DEBUG=1 every attempt is logged.
    """
    variants = [
        ("original",   image),
        ("grayscale",  image.convert("L")),
    ]

    for var_name, img in variants:
        results = zxingcpp.read_barcodes(img)
        for r in results:
            if not r.valid:
                continue
            text = r.text.strip()
            if OCR_DEBUG:
                logging.info(
                    f"  QR p{page_num} | {label} [{var_name}]: found {text!r}"
                )
            if QR_PATTERN.search(text):
                if OCR_DEBUG:
                    logging.info(
                        f"  QR p{page_num} | {label} [{var_name}]: MATCH {text!r}"
                    )
                return _parse_qr_string(text)
            if OCR_DEBUG:
                logging.info(
                    f"  QR p{page_num} | {label} [{var_name}]: "
                    f"barcode found but pattern not matched"
                )
        if OCR_DEBUG and not any(r.valid for r in results):
            logging.info(
                f"  QR p{page_num} | {label} [{var_name}]: no barcode detected"
            )
    return None


def extract_qr_code(images: list) -> Optional[dict]:
    """
    Scan each page image for a QR code matching the ticket pattern.

    Detection strategy per page (stops at first success):
      1. Full page at original resolution
      2. Full page at 2× resolution
      3. Right 25% strip (full height)
      4. Upper-right 25%×25% crop
      5. OCR fallback — pytesseract --psm 6 regex search on full page

    Each image region is tried with both colour and grayscale variants via
    zxingcpp.  When OCR_DEBUG=1 every attempt and its outcome is logged, and
    the first page is saved to qr_debug.png for inspection.

    Returns:
        Parsed QR data dict, or None if no valid code is found.
    """
    if OCR_DEBUG and images:
        images[0].save("qr_debug.png")
        logging.info(
            f"QR debug: first page saved to qr_debug.png "
            f"(size {images[0].size[0]}×{images[0].size[1]})"
        )

    for page_num, image in enumerate(images, start=1):
        w, h = image.size

        # ── Step 1: full page, original resolution ────────────────────────
        if OCR_DEBUG:
            logging.info(f"  QR p{page_num} | step 1: full page ({w}×{h})")
        result = _zxing_scan(image, "full-page", page_num)
        if result:
            return result

        # ── Step 2: full page, 2× upscale ────────────────────────────────
        img_2x = image.resize((w * 2, h * 2), Image.LANCZOS)
        if OCR_DEBUG:
            logging.info(f"  QR p{page_num} | step 2: full page 2× ({w*2}×{h*2})")
        result = _zxing_scan(img_2x, "full-page-2x", page_num)
        if result:
            return result

        # ── Step 3: right 25% strip (full height) ────────────────────────
        rs_box = (int(w * 0.75), 0, w, h)
        rs_crop = image.crop(rs_box)
        if OCR_DEBUG:
            logging.info(
                f"  QR p{page_num} | step 3: right-strip "
                f"({rs_crop.size[0]}×{rs_crop.size[1]}, box={rs_box})"
            )
        result = _zxing_scan(rs_crop, "right-strip", page_num)
        if result:
            return result

        # ── Step 4: upper-right 25%×25% crop ─────────────────────────────
        ur_box = (int(w * 0.75), 0, w, int(h * 0.25))
        ur_crop = image.crop(ur_box)
        if OCR_DEBUG:
            logging.info(
                f"  QR p{page_num} | step 4: upper-right "
                f"({ur_crop.size[0]}×{ur_crop.size[1]}, box={ur_box})"
            )
        result = _zxing_scan(ur_crop, "upper-right", page_num)
        if result:
            return result

        # ── Step 5: OCR fallback ──────────────────────────────────────────
        if OCR_DEBUG:
            logging.info(f"  QR p{page_num} | step 5: OCR fallback")
        ocr_text = pytesseract.image_to_string(image, config="--psm 6 --oem 3")
        match = QR_PATTERN.search(ocr_text)
        if match:
            if OCR_DEBUG:
                logging.info(
                    f"  QR p{page_num} | OCR fallback: MATCH {match.group(0)!r}"
                )
            return _parse_qr_string(match.group(0))
        else:
            if OCR_DEBUG:
                logging.info(
                    f"  QR p{page_num} | OCR fallback: pattern not found in OCR text"
                )

    return None   # Pattern not found on any page


def _parse_qr_string(raw: str) -> dict:
    """
    Parse a raw QR string into component fields.

    Input format : XXXX-YYYY-CCCCCC-CC
    Returns a dict with keys: job_number, location, cost_code_1,
    cost_code_2, cost_code, raw.
    """
    match = QR_PATTERN.search(raw)
    if not match:
        raise ValueError(f"String does not match QR pattern: {raw!r}")

    job_number   = match.group(1).upper()
    location     = match.group(2).upper()
    cost_code_1  = match.group(3).upper()
    cost_code_2  = match.group(4).upper()

    return {
        "job_number":  job_number,
        "location":    location,
        "cost_code_1": cost_code_1,
        "cost_code_2": cost_code_2,
        "cost_code":   f"{cost_code_1}-{cost_code_2}",
        "raw":         raw,
    }


# ============================================================
# OCR CONFIDENCE CHECK
# ============================================================
def _run_confidence_check(
    check_name: str, full_text: str, profile: TicketProfile,
    images: list = None,
) -> Optional[str]:
    """Run one named confidence check from the profile's confidence_checks list.

    Returns None when the check passes, or a human-readable failure description
    when it fails.  Unknown check names are logged and treated as passing so
    adding new names to a profile never silently breaks old installations.
    """
    if check_name == "valid_date_on_line_2":
        date_str, _ = _parse_ticket_header(full_text)
        return None if date_str else "date not found on header line"

    if check_name == "five_digit_ticket_on_line_2":
        # Primary: look for the number on the header line of the full-page text.
        _, ticket = _parse_ticket_header(full_text)
        if ticket:
            return None
        # Fallback: if the profile uses region OCR (bold corner number), try
        # cropping that region directly — many tickets place the number there
        # and it never appears in the full-page text stream.
        region_config = profile.layout.get("ticket_number_extraction", {})
        if region_config.get("method") == "region_ocr" and images:
            region_ticket = _ocr_ticket_number_from_region(images[0], region_config)
            if region_ticket:
                return None   # found via region OCR — check passes
        return "5-digit ticket number not found on header line or upper-right region"

    if check_name == "product_label_found":
        labels = profile.labels.get("material", ["Product"])
        if isinstance(labels, str):
            labels = [labels]
        found = any(
            re.search(re.escape(lbl.rstrip(":. ")), full_text, re.IGNORECASE)
            for lbl in labels
        )
        if found:
            return None
        # Print the full OCR text so we can see what Tesseract actually read
        logging.warning("  Full OCR text (product label not found):")
        for line in full_text.splitlines():
            logging.warning(f"    | {line}")
        return f"material/product label not found (expected one of: {labels})"

    if check_name == "net_row_found":
        net_label = profile.labels.get("net_row", "Net")
        found = bool(re.search(rf'\b{re.escape(net_label)}\b', full_text, re.IGNORECASE))
        return None if found else f"'{net_label}' row not found — weight table may not have been read"

    if check_name == "facility_label_found":
        labels = profile.labels.get("facility", ["Location"])
        if isinstance(labels, str):
            labels = [labels]
        found = any(
            re.search(re.escape(lbl.rstrip(":. ")), full_text, re.IGNORECASE)
            for lbl in labels
        )
        if not found:
            return f"facility label not found (expected one of: {labels})"
        facility_val = _ocr_facility(full_text, profile)
        if facility_val and re.search(r'[&$@#%]', facility_val):
            return f"facility name contains garble characters: {facility_val!r}"
        return None

    logging.warning(
        f"Unknown confidence check {check_name!r} in profile "
        f"{profile.supplier_name!r} — skipping"
    )
    return None


def _check_ocr_confidence(
    full_text: str, profile: TicketProfile, images: list = None,
) -> tuple[bool, list[str]]:
    """Run all confidence checks defined in the profile.

    Returns:
        (all_passed, failed_descriptions) — all_passed is True only when
        the failed list is empty.
    """
    failed: list[str] = []
    for check_name in profile.confidence_checks:
        result = _run_confidence_check(check_name, full_text, profile, images=images)
        if result is not None:
            failed.append(result)
    return len(failed) == 0, failed


# ============================================================
# OCR TEXT EXTRACTION
# ============================================================
def extract_ticket_data_ocr(images: list, profiles: list[TicketProfile], subject: str = "") -> dict:
    """Run OCR, detect supplier profile, apply corrections, check confidence, extract fields.

    Returns a dict.  Special keys (all start with "_"):
        _ocr_confidence_passed : bool — False when extraction should not be trusted.
        _no_profile_match      : bool — True when no supplier profile was identified.
        _company_name          : str  — first OCR line when no profile matched.
        _failed_checks         : list — check failure descriptions on confidence failure.
    """
    # 1. Run OCR on all pages (no corrections yet — need raw text for profile detection).
    raw_text = "\n".join(
        pytesseract.image_to_string(preprocess_for_ocr(img), config="--oem 3 --psm 6")
        for img in images
    )

    # 2. Detect supplier profile from raw OCR text.
    profile = detect_profile(raw_text, profiles)
    if not profile:
        company_name = next(
            (line.strip() for line in raw_text.splitlines() if line.strip()), "unknown"
        )
        logging.warning(
            f"  No matching profile found. First OCR line: {company_name!r}"
        )
        _log_unknown_supplier(company_name)
        return {
            "_ocr_confidence_passed": False,
            "_no_profile_match": True,
            "_company_name": company_name,
        }

    logging.info(f"  Detected profile: {profile.supplier_name!r}")

    # 3. Apply profile-specific OCR corrections.
    full_text = _apply_ocr_corrections(raw_text, profile)
    logging.debug(f"OCR full text (after corrections):\n{full_text}")

    if OCR_DEBUG:
        Path("ocr_debug.txt").write_text(full_text, encoding="utf-8")
        logging.info("OCR debug: corrected text saved to ocr_debug.txt")
        if images:
            debug_img = preprocess_for_ocr(images[0])

            # If the profile uses region OCR for ticket number, overlay a red
            # rectangle on the debug image showing the exact crop area, and save
            # the preprocessed crop separately so the region can be verified.
            region_config = profile.layout.get("ticket_number_extraction", {})
            if region_config.get("method") == "region_ocr":
                orig_w, orig_h = images[0].size
                rgn   = region_config.get("region", {})
                x1    = int(orig_w * rgn.get("x_start_pct", 0.70))
                y1    = int(orig_h * rgn.get("y_start_pct", 0.00))
                x2    = int(orig_w * rgn.get("x_end_pct",   1.00))
                y2    = int(orig_h * rgn.get("y_end_pct",   0.15))

                # Save the preprocessed crop (what Tesseract actually sees)
                crop_debug = preprocess_for_ocr(images[0].crop((x1, y1, x2, y2)))
                crop_debug.save("ocr_debug_topright.png")
                logging.info(
                    "OCR debug: ticket-number crop saved to ocr_debug_topright.png "
                    f"(original coords x={x1}-{x2}, y={y1}-{y2})"
                )

                # Draw a red rectangle on the full-page debug image.
                # The preprocessed image is 2× the original, so scale coords.
                scale = 2
                draw  = ImageDraw.Draw(debug_img)
                draw.rectangle(
                    [x1 * scale, y1 * scale, x2 * scale - 1, y2 * scale - 1],
                    outline="red",
                    width=4,
                )

            debug_img.save("ocr_debug.png")
            logging.info("OCR debug: preprocessed first page saved to ocr_debug.png")

    # 4. Run profile confidence checks.
    confidence_passed, failed_checks = _check_ocr_confidence(full_text, profile, images=images)
    if not confidence_passed:
        for check in failed_checks:
            logging.warning(f"  OCR confidence FAILED: {check}")
        logging.warning(
            f"OCR confidence check failed ({len(failed_checks)} check(s) failed). "
            "Ticket will be flagged for manual review."
        )
        return {"_ocr_confidence_passed": False, "_failed_checks": failed_checks}

    # 5. Extract all fields — try Claude AI first, fall back to regex.
    ticket_data: dict = {"_ocr_confidence_passed": True}
    ai_used = False
    _label = subject or "(unknown email)"

    if ANTHROPIC_API_KEY:
        logging.info(f"  Using AI extraction for: {_label}")
        try:
            ai_fields = extract_fields_with_ai(full_text, profile.supplier_name)
            ticket_data.update(ai_fields)
            ai_used = True
        except Exception as ai_exc:
            logging.warning(f"  AI extraction failed, falling back to regex: {ai_exc}")
    else:
        logging.warning("  ANTHROPIC_API_KEY not set — using regex extraction.")

    if not ai_used:
        logging.info(f"  Using regex extraction for: {_label}")
        ticket_data.update({
            "date":          _ocr_date(full_text),
            "facility":      _ocr_facility(full_text, profile),
            "customer":      _ocr_customer(full_text, profile),
            "material":      _ocr_material(full_text, profile),
            "ticket_number": _ocr_ticket_number(full_text, images, profile),
            "net_tons":      _ocr_net_tons(full_text, profile),
        })

    # 6. Post-extraction check: warn on any required field that is still empty.
    required_fields = ["date", "facility", "customer", "material", "ticket_number", "net_tons"]
    for field_name in required_fields:
        if not ticket_data.get(field_name):
            src = "AI" if ai_used else "regex"
            logging.warning(f"  {src} extraction: could not extract '{field_name}'.")

    return ticket_data


REVIEW_REQUIRED = "REVIEW REQUIRED"

# ============================================================
# AI FIELD EXTRACTION
# ============================================================
_AI_MODEL  = "claude-sonnet-4-6"
_AI_PROMPT = """\
You are extracting structured data from a scanned material delivery ticket.
The text below was produced by Tesseract OCR and may contain character-level
errors caused by scan quality (e.g. "Bradiey" = "Bradley", "3)4-0" = "3/4-0").

Extract exactly these six fields:

  ticket_number  — 5-digit bold number printed near the top-right corner
  date           — delivery date, normalised to MM/DD/YYYY
  facility       — quarry / pit / source location printed after the
                   "Location:" label.  NEVER a customer or contractor name.
                   Must not contain "MJ Hughes" or "Hughes Construction".
  customer       — company receiving the material, printed after "Customer:"
  material       — description of what was delivered.
                   The material/product field may be labeled as any of these
                   on the ticket: Product, Product., Material, Material:,
                   Item, Description, Mat, Prod. They all refer to the same
                   thing — extract the text after whichever label appears.
                   Include sizing notation (e.g. "3/4-0\"").
                   Exclude leading numeric product codes and quantity values.
  net_tons       — decimal number from the Tons column on the Net weight row.
                   Return the number only, no units.

Rules:
  • Correct obvious OCR errors based on context.
  • If a field genuinely cannot be determined, return null for that key.
  • Return ONLY a JSON object — no markdown fences, no explanation.

Supplier: {supplier_name}

OCR TEXT:
{ocr_text}"""


def extract_fields_with_ai(ocr_text: str, supplier_name: str) -> dict:
    """Send OCR text to Claude and return extracted ticket fields.

    Returns a dict with string values (or None) for the keys:
        ticket_number, date, facility, customer, material, net_tons

    Raises on API error or if the response is not valid JSON — the caller
    must catch and fall back to regex extraction.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    prompt = _AI_PROMPT.format(
        supplier_name=supplier_name,
        ocr_text=ocr_text,
    )
    response = client.messages.create(
        model=_AI_MODEL,
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}],
    )
    raw_json = response.content[0].text.strip()

    # Strip accidental markdown fences if the model adds them despite instructions
    if raw_json.startswith("```"):
        raw_json = re.sub(r"^```[a-z]*\n?", "", raw_json)
        raw_json = re.sub(r"\n?```$", "", raw_json)

    logging.info(f"  AI extraction result: {raw_json}")

    data = json.loads(raw_json)

    # Normalise: convert null → "" and coerce all values to str.
    # Log any field that was null before normalisation so we have a record.
    required_keys = {"ticket_number", "date", "facility", "customer", "material", "net_tons"}
    for key in required_keys:
        val = data.get(key)
        if val is None:
            logging.warning(f"  AI returned null for field: {key}")
        data[key] = str(val).strip() if val is not None else ""

    return data


# Matches a complete date with two separators: M/D/YYYY, MM/DD/YY, M-D-YYYY, etc.
# Intentionally uses only / and - so time strings like "10:18:01AM" are never matched.
_HEADER_DATE_RE   = re.compile(r'\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\b')
# Matches the first standalone 5-digit ticket number (10000-99999).
_HEADER_TICKET_RE = re.compile(r'\b([1-9]\d{4})\b')
# Lines containing weight-table keywords are skipped when scanning for the header.
_HEADER_SKIP_RE   = re.compile(
    r'\b(?:gross|tare|net|lbs?|pounds?|tons?|metric)\b', re.IGNORECASE
)


def _parse_ticket_header(text: str) -> tuple[str, str]:
    """Locate the header line and return (date_str, ticket_number).

    The Teevin Bros ticket layout is fixed:
      Line 1 — company name  (no date pattern → naturally skipped)
      Line 2 — date and ticket number, e.g. "8/6/2025 16800"
      Line 3 — time, e.g. "10:18:01AM"  (colon-separated → never matched as date)

    Strategy: scan non-empty lines from the top, skip weight-table rows, and
    return the first line that contains a complete date (two separators).  The
    ticket number is the first 5-digit number on that same line.

    Returns:
        (date_str, ticket_number) — either value is "" when not found on the
        header line.  Both being "" means the header is unreadable.
    """
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or _HEADER_SKIP_RE.search(stripped):
            continue
        date_m = _HEADER_DATE_RE.search(stripped)
        if not date_m:
            continue
        # Found the header line — extract ticket number from the same line.
        ticket_m = _HEADER_TICKET_RE.search(stripped)
        return date_m.group(1), (ticket_m.group(1) if ticket_m else "")

    return "", ""


def _ocr_date(text: str) -> str:
    """Return the date extracted from the ticket header line, or REVIEW_REQUIRED."""
    date_str, _ = _parse_ticket_header(text)
    return date_str if date_str else REVIEW_REQUIRED


def _ocr_ticket_number_from_region(
    image: Image.Image, region_config: dict
) -> str:
    """Crop the image to the configured region and extract the ticket number.

    The crop uses percentage coordinates from the profile.  Full-text OCR
    (--psm 6, no character whitelist) is run on the preprocessed crop.  The
    first *standalone* 5-digit number is returned — "standalone" meaning it
    is not surrounded by other digits or letters (i.e. not embedded in a
    longer number or word).  Any match whose position in the crop falls in
    the rightmost 15 % of the *full* page is skipped to avoid confusing a
    QR-code digit sequence with the ticket number.

    When OCR_DEBUG=1:
      - Logs the full page dimensions and crop box.
      - Logs the raw OCR text from the crop.
      - Saves the preprocessed crop to ocr_debug_topright.png and the raw
        OCR text to ocr_debug_topright.txt.

    Returns the first qualifying 5-digit number, or "" if none found.
    """
    # Standalone 5-digit number: not preceded/followed by a letter or digit.
    _TICKET_RE = re.compile(r'(?<![a-zA-Z\d])([1-9]\d{4})(?![a-zA-Z\d])')

    w, h = image.size

    if OCR_DEBUG:
        logging.info(f"  Page dimensions: {w}×{h} px")

    region = region_config.get("region", {})
    x1 = int(w * region.get("x_start_pct", 0.53))
    y1 = int(h * region.get("y_start_pct", 0.00))
    x2 = int(w * region.get("x_end_pct",   1.00))
    y2 = int(h * region.get("y_end_pct",   0.24))

    crop   = image.crop((x1, y1, x2, y2))
    prep   = preprocess_for_ocr(crop)
    config = region_config.get("tesseract_config", "--oem 3 --psm 6")
    raw    = pytesseract.image_to_string(prep, config=config).strip()

    if OCR_DEBUG:
        logging.info(
            f"  Region crop ({x1},{y1})–({x2},{y2})  raw OCR: {raw!r}"
        )
        try:
            prep.save("ocr_debug_topright.png")
            Path("ocr_debug_topright.txt").write_text(raw, encoding="utf-8")
            logging.info("  Saved ocr_debug_topright.png and ocr_debug_topright.txt")
        except Exception as dbg_exc:
            logging.warning(f"  Could not save region debug files: {dbg_exc}")
    else:
        logging.debug(f"  Region OCR raw: {raw!r}")

    # QR code sits in the rightmost 15 % of the full page.  The crop starts
    # at x1 pixels from the left edge of the full page, so the QR exclusion
    # zone begins at (0.85 * w - x1) pixels from the left edge of the crop.
    qr_zone_start_in_crop = max(0, int(w * 0.85) - x1)

    # Use pytesseract data to get per-word bounding boxes when available;
    # fall back to a plain regex scan if not.
    try:
        import pytesseract as _tess
        data = _tess.image_to_data(
            prep, config=config, output_type=_tess.Output.DICT
        )
        for i, word in enumerate(data["text"]):
            m = _TICKET_RE.fullmatch(word.strip()) if word.strip() else None
            if not m:
                continue
            word_x = data["left"][i]
            if word_x >= qr_zone_start_in_crop:
                if OCR_DEBUG:
                    logging.info(
                        f"  Skipping {word!r} at crop-x={word_x} "
                        f"(QR zone starts at {qr_zone_start_in_crop})"
                    )
                continue
            if OCR_DEBUG:
                logging.info(f"  Ticket number from region (bbox): {word.strip()!r}")
            return word.strip()
    except Exception:
        pass  # fall through to regex scan

    # Regex fallback — no positional info available
    for m in _TICKET_RE.finditer(raw):
        if OCR_DEBUG:
            logging.info(f"  Ticket number from region (regex): {m.group(1)!r}")
        return m.group(1)

    return ""


def _ocr_ticket_number(
    text: str, images: list, profile: TicketProfile
) -> str:
    """Extract the ticket number, trying region OCR first then full-page fallback.

    If profile.layout["ticket_number_extraction"]["method"] == "region_ocr",
    the configured region of the first page image is OCR'd with a digit
    whitelist (ideal for bold standalone numbers in a corner).  Falls back to
    scanning the date-line of the full-page text if the region yields nothing.
    """
    region_config = profile.layout.get("ticket_number_extraction", {})
    if region_config.get("method") == "region_ocr" and images:
        result = _ocr_ticket_number_from_region(images[0], region_config)
        if result:
            logging.info(f"  Ticket number from region OCR: {result!r}")
            return result
        logging.debug(
            "  Region OCR returned no 5-digit number — falling back to header line."
        )

    # Fallback: parse the header line of the full-page text
    _, ticket_number = _parse_ticket_header(text)
    return ticket_number


def _load_known_facilities() -> list:
    """Return the list of confirmed facility names from known_facilities.txt."""
    path = Path(KNOWN_FACILITIES_FILE)
    if not path.exists():
        return []
    return [l.strip() for l in path.read_text(encoding="utf-8").splitlines() if l.strip()]


def _save_known_facility(name: str) -> None:
    """Append a new facility name to known_facilities.txt if not already present."""
    if name in _load_known_facilities():
        return
    with open(KNOWN_FACILITIES_FILE, "a", encoding="utf-8") as fh:
        fh.write(name + "\n")
    logging.info(f"New facility saved to {KNOWN_FACILITIES_FILE}: {name!r}")


def _strip_facility_noise(candidate: str) -> str:
    """Filter OCR garbage tokens from a facility name candidate.

    Each token is tested against a set of noise rules.  Tokens that pass all
    rules are kept; tokens that fail any rule are dropped.  The leading numeric
    site code (e.g. "800") is always preserved.

    A token is kept when ALL of the following are true:
      1. It contains at least one alphabetic character  (drops "|", "?", "=")
      2. It has no characters other than letters, digits, hyphens, and
         apostrophes  (drops tokens with |, ?, @, #, etc.)
      3. It is not a single character  (drops stray "A", "I", etc.)
      4. It is not all-uppercase AND shorter than 4 letters
         (drops "OR", "Oo", "AR" but keeps "ROCK", "SAND", "MINE")
      5. It does not contain an uppercase letter after its first character
         when the token is lower-cased start  (drops "eRe", "oRe", "bRo")

    The leading numeric token (digits only, e.g. "800") is always kept
    regardless of the above rules.
    """
    tokens = candidate.split()
    clean: list[str] = []
    for i, tok in enumerate(tokens):
        # Always keep a leading numeric site code
        if i == 0 and tok.isdigit():
            clean.append(tok)
            continue
        # Rule 1: must contain at least one letter
        if not any(c.isalpha() for c in tok):
            continue
        # Rule 2: only letters, digits, hyphens, apostrophes allowed
        if re.search(r"[^a-zA-Z0-9\-']", tok):
            continue
        # Rule 3: must be longer than 1 character
        if len(tok) <= 1:
            continue
        # Rule 4: all-caps tokens must be 4+ letters  (drops "OR", "AR")
        if tok.isupper() and len(tok) < 4:
            continue
        # Rule 5: any 2-character token with an uppercase letter is noise
        #         (drops "Ar", "My", "Oo", "Or" — too short to be a real word
        #          in a facility name context)
        if len(tok) == 2 and any(c.isupper() for c in tok):
            continue
        # Rule 6: lowercase-start token with mid-uppercase (e.g. "eRe")
        if tok[0].islower() and any(c.isupper() for c in tok[1:]):
            continue
        clean.append(tok)
    return " ".join(clean)


def _fuzzy_correct_facility(raw: str) -> str:
    """Compare a raw facility name against known_facilities.txt and correct OCR errors.

    Strategy:
    - Split both the extracted value and each known facility into a leading
      numeric prefix ("800") and a text portion ("Bradley Quarry").
    - Fuzzy-match the text portions only so that different site numbers don't
      cause a false positive between two otherwise identical quarry names.
    - If the best text match scores >= 80%, return the numeric prefix from the
      extracted value combined with the corrected text from the known list.
    - If no match scores >= 80%, treat it as a new facility, save it, and
      return the raw value unchanged.
    """
    known = _load_known_facilities()

    def _split(s: str):
        """Split "800 Bradley Quarry" → ("800", "Bradley Quarry")."""
        m = re.match(r'^(\d[\d\s]*?)\s+([A-Za-z].*)$', s.strip())
        return (m.group(1).strip(), m.group(2).strip()) if m else ("", s.strip())

    raw_prefix, raw_text = _split(raw)

    # Build a map of text-portion → full known name
    known_text_map: dict[str, str] = {}
    for name in known:
        _, ktext = _split(name)
        known_text_map[ktext] = name

    matches = difflib.get_close_matches(
        raw_text, known_text_map.keys(), n=1, cutoff=0.80
    )

    if matches:
        best_text = matches[0]
        _, corrected_text = _split(known_text_map[best_text])
        corrected = f"{raw_prefix} {corrected_text}".strip() if raw_prefix else corrected_text
        if corrected != raw:
            logging.info(f"Facility fuzzy-corrected: {raw!r} → {corrected!r}")
        return corrected

    # New facility — strip trailing OCR bleed before persisting
    cleaned = _strip_facility_noise(raw)
    if cleaned != raw:
        logging.info(f"Facility trailing noise stripped: {raw!r} → {cleaned!r}")
    _save_known_facility(cleaned)
    return cleaned


def _ocr_facility(text: str, profile: TicketProfile) -> str:
    """Extract the source facility (quarry, mine, pit, plant) from OCR text.

    Primary label candidates come from profile.labels["facility"].
    Generic keyword fallback is retained for robustness.
    """
    company_name = re.compile(r'mj\s*hughes|hughes\s*construction', re.IGNORECASE)

    _STOP_RE = re.compile(
        r'\b(?:pays|customer|order|p\.o|product|carrier|vehicle|'
        r'weighmaster|gross|tare|net)\b',
        re.IGNORECASE,
    )

    def _clean(raw: str) -> str:
        cut = _STOP_RE.search(raw)
        return raw[:cut.start()].strip() if cut else raw.strip()

    # Build primary label patterns from the profile
    facility_labels = profile.labels.get("facility", ["Location"])
    if isinstance(facility_labels, str):
        facility_labels = [facility_labels]

    label_patterns = [
        rf'{re.escape(lbl.rstrip(":. "))}[.:\s]+([^\n]{{3,80}})'
        for lbl in facility_labels
    ]
    # Generic fallback labels (not supplier-specific)
    label_patterns.append(
        r'(?:shipped\s*from|from|source|origin|facility|pit|quarry|mine|'
        r'plant|producer|supplier|sold\s*by|vendor)[.:\s]+([^\n]{3,80})'
    )

    for pattern in label_patterns:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            candidate = _clean(m.group(1))
            if not candidate or company_name.search(candidate):
                continue
            if len(candidate) >= 3:
                return _fuzzy_correct_facility(_apply_ocr_corrections(candidate, profile))

    # Fallback: line containing a facility keyword
    facility_keywords = re.compile(
        r'\b(?:quarry|mine|pit|plant|aggregate|gravel|sand|rock|stone)\b',
        re.IGNORECASE,
    )
    for line in text.splitlines():
        line = line.strip()
        if not line or company_name.search(line):
            continue
        if facility_keywords.search(line):
            return _fuzzy_correct_facility(_apply_ocr_corrections(_clean(line), profile))

    return ""


def _apply_ocr_corrections(value: str, profile: TicketProfile) -> str:
    """Apply OCR misread corrections defined in the supplier profile."""
    for wrong, right in profile.ocr_corrections.items():
        value = value.replace(wrong, right)
    return value


def _ocr_material(text: str, profile: TicketProfile) -> str:
    """Extract the material / product description from OCR text.

    Primary label candidates come from profile.labels["material"].
    Generic fallback labels are appended after profile labels.
    """
    material_labels = profile.labels.get("material", ["Product"])
    if isinstance(material_labels, str):
        material_labels = [material_labels]

    patterns = [
        rf'{re.escape(lbl.rstrip(":. "))}[:\s]+([^\n]{{3,80}})'
        for lbl in material_labels
    ]
    patterns += [
        r'(?:material(?:\s+type)?|description|item|commodity|aggregate)[:\s]+([^\n]{3,80})',
        r'(?:type|grade)[:\s]+([^\n]{3,80})',
    ]

    def _clean_material(raw: str) -> str:
        raw = _apply_ocr_corrections(raw.strip(), profile)
        # Strip a leading product code prefix in two forms:
        #   "—-801 3/4-0"" — special chars before the code (original pattern)
        #   "801 3/4-0""   — plain code with no leading special chars,
        #                    but only when the code is followed by a non-letter
        #                    so "808 Bradley Rock Rip Rap" is left untouched.
        raw = re.sub(r'^[-—–=!\s]*\d{2,4}\s+(?=[^a-zA-Z])', '', raw)
        # Truncate at the first digit that follows a non-digit word character
        cut = re.search(r'(?<=\D)\s+\d', raw)
        if cut:
            raw = raw[:cut.start()]
        # Strip trailing noise symbols and unit words.
        # NOTE: double-quote (") is intentionally excluded from the character
        # class — it is part of aggregate sizing notation (e.g. '3/4-0"'
        # meaning 0-inch minus) and must be preserved.
        raw = re.sub(
            r"[\s=*/@'!|]+(?:tons?|lbs?|pounds?|kg)?\s*$",
            '',
            raw,
            flags=re.IGNORECASE,
        ).strip()
        return raw

    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result = _clean_material(m.group(1))
            if result:
                return result

    # Last-resort fallback: look for aggregate-style sizing patterns (e.g. "3/4-0"")
    # in the bottom third of the OCR text where the product line usually appears.
    lines = text.splitlines()
    bottom_lines = lines[len(lines) * 2 // 3:]
    for line in bottom_lines:
        m = re.search(r'(\d+/\d+[-\d"]*(?:\s+\w+)*)', line)
        if m:
            candidate = _clean_material(m.group(1))
            if candidate:
                logging.debug(f"  Material from bottom-third fallback: {candidate!r}")
                return candidate

    return ""


def _ocr_net_tons(text: str, profile: TicketProfile) -> str:
    """Extract net quantity in tons from OCR text.

    Uses profile.labels["net_row"] as the row identifier and
    profile.weight_table["tons_column_index"] (0-based) to select the
    correct number from the weight row (default: index 1 = second number).
    """
    net_label  = profile.labels.get("net_row", "Net")
    tons_col   = profile.weight_table.get("tons_column_index", 1)  # 0-based
    net_escape = re.escape(net_label)

    # Primary: Net row with two numbers — pick column by tons_column_index
    net_re = re.compile(
        rf'\b{net_escape}\b\s+(\d[\d,]*\.?\d*)\s+(\d[\d,]*\.?\d*)',
        re.IGNORECASE,
    )
    for line in text.splitlines():
        m = net_re.match(line.strip())
        if m:
            # group(1) = index 0 (pounds), group(2) = index 1 (tons)
            group = tons_col + 1
            try:
                return m.group(group).replace(",", "").strip()
            except IndexError:
                return m.group(2).replace(",", "").strip()

    # Fallback: explicit tons label
    fallback_patterns = [
        rf'{net_escape}\s*tons?[:\s]+(\d[\d,]*\.?\d*)',
        r'\btons?[:\s]+(\d[\d,]*\.?\d*)',
        r'(\d[\d,]*\.\d{1,4})\s*tons?\b',
    ]
    lbs_label = re.compile(r'\blbs?\b|\bpounds?\b', re.IGNORECASE)
    for pattern in fallback_patterns:
        for line in text.splitlines():
            if lbs_label.search(line):
                continue
            m = re.search(pattern, line, re.IGNORECASE)
            if m:
                return m.group(1).replace(",", "").strip()

    return ""


def _ocr_customer(text: str, profile: TicketProfile) -> str:
    """Extract the customer / company name from OCR text.

    Label candidates come from profile.labels["customer"].
    The value ends before "Pounds" (weight column header) or end of line.
    """
    customer_labels = profile.labels.get("customer", ["Customer"])
    if isinstance(customer_labels, str):
        customer_labels = [customer_labels]

    for lbl in customer_labels:
        clean_lbl = re.escape(lbl.rstrip(":. "))
        m = re.search(
            rf'{clean_lbl}[.:\s]+(.+?)(?=\s*pounds|\n|$)',
            text,
            re.IGNORECASE,
        )
        if m:
            return m.group(1).strip()
    return ""


# ============================================================
# EXCEL OPERATIONS
# ============================================================
def _find_duplicate_ticket(ticket_number: str) -> Optional[str]:
    """Search all tabs in materials_log.xlsx for an existing ticket number.

    Scans every sheet (skipping the header row) and looks at the
    "Ticket Number" column.  Returns the name of the first sheet where
    the ticket number is found, or None if no match exists.

    Args:
        ticket_number: The extracted ticket number string to look up.

    Returns:
        Sheet name string if a duplicate is found, else None.
    """
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists() or not ticket_number:
        return None

    ticket_col_idx = EXCEL_COLUMNS.index("Ticket Number")  # 0-based

    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > ticket_col_idx:
                    cell_val = row[ticket_col_idx]
                    if cell_val and str(cell_val).strip() == str(ticket_number).strip():
                        return sheet_name
    finally:
        workbook.close()

    return None


def write_to_excel(qr_data: dict, ticket_data: dict, flag: str = "") -> None:
    """
    Write one row of ticket data into the Excel workbook.

    - Opens materials_log.xlsx if it exists, or creates it.
    - Looks for a sheet named after the full cost code (e.g. "123456-78").
    - Creates the sheet with styled headers if it doesn't exist.
    - Appends the new row and saves.

    Args:
        flag: Optional flag value written to the last "Flag" column.
              Pass "DUPLICATE TICKET" to highlight the entire row orange.
              Leave empty for normal rows (REVIEW REQUIRED cells get yellow).
    """
    excel_path = Path(EXCEL_FILE)
    cost_code  = qr_data["cost_code"]

    # Load or create the workbook
    if excel_path.exists():
        workbook = openpyxl.load_workbook(excel_path)
    else:
        workbook = Workbook()
        # Remove the auto-created blank "Sheet" that Workbook() adds
        if "Sheet" in workbook.sheetnames:
            del workbook["Sheet"]
        logging.info(f"Created new workbook: {EXCEL_FILE}")

    # Get the correct cost-code tab, creating it if needed
    if cost_code in workbook.sheetnames:
        sheet = workbook[cost_code]
    else:
        sheet = workbook.create_sheet(title=cost_code)
        _add_sheet_headers(sheet)
        logging.info(f"Created new tab for cost code: {cost_code}")

    # Build the row in the same order as EXCEL_COLUMNS
    row = [
        qr_data["job_number"],
        qr_data["location"],
        ticket_data.get("date", ""),
        ticket_data.get("facility", ""),
        ticket_data.get("customer", ""),
        ticket_data.get("material", ""),
        ticket_data.get("ticket_number", ""),
        ticket_data.get("net_tons", ""),
        flag,   # Flag column — "DUPLICATE TICKET" or blank
    ]

    sheet.append(row)
    last_row = sheet.max_row

    if flag == "DUPLICATE TICKET":
        # Highlight entire row orange so duplicates are visually distinct
        # from both normal rows and low-confidence rows (red).
        _orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        for col_idx in range(1, len(EXCEL_COLUMNS) + 1):
            sheet.cell(row=last_row, column=col_idx).fill = _orange
    else:
        # Single pass: yellow for REVIEW REQUIRED values AND empty data cells.
        # The Flag column is skipped — it is intentionally blank on normal rows.
        _yellow       = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ticket_number = ticket_data.get("ticket_number", "N/A")

        for col_idx, (col_name, value) in enumerate(zip(EXCEL_COLUMNS, row), start=1):
            if value == REVIEW_REQUIRED:
                sheet.cell(row=last_row, column=col_idx).fill = _yellow
            elif col_name != "Flag" and (value is None or value == ""):
                sheet.cell(row=last_row, column=col_idx).fill = _yellow
                msg = f"Empty field '{col_name}' for ticket {ticket_number}"
                logging.warning(f"  WARNING: {msg}")
                log_error(f"ticket {ticket_number}", f"Empty field: {col_name}")

    workbook.save(excel_path)
    logging.info(
        f"Excel: wrote ticket {ticket_data.get('ticket_number', 'N/A')} "
        f"to tab '{cost_code}' in {EXCEL_FILE}."
    )


def write_flagged_row_to_excel(qr_data: Optional[dict]) -> None:
    """Write a REVIEW REQUIRED row with a red background for a low-confidence scan.

    If the QR code was readable the row is written to the cost-code tab and
    Job Number / Location are filled from QR data.  If the QR code also
    failed every column gets REVIEW_REQUIRED and the row goes to a dedicated
    "REVIEW REQUIRED" tab.
    """
    excel_path = Path(EXCEL_FILE)

    if excel_path.exists():
        workbook = openpyxl.load_workbook(excel_path)
    else:
        workbook = Workbook()
        if "Sheet" in workbook.sheetnames:
            del workbook["Sheet"]

    sheet_name = qr_data["cost_code"] if qr_data else "REVIEW REQUIRED"

    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.create_sheet(title=sheet_name)
        _add_sheet_headers(sheet)

    if qr_data:
        # QR was readable: populate Job Number and Location from QR data.
        # Date gets REVIEW REQUIRED so the row is easy to find.
        # All other data fields are left blank — do not write partial OCR data.
        row = [
            qr_data["job_number"],  # Job Number
            qr_data["location"],    # Location
            REVIEW_REQUIRED,        # Date
            "",                     # Facility
            "",                     # Customer
            "",                     # Material
            "",                     # Ticket Number
            "",                     # Net Quantity (Tons)
            "",                     # Flag — blank; red fill identifies these rows
        ]
    else:
        # QR also failed — no reliable data at all.
        # Flag column intentionally left blank; red fill is the visual indicator.
        row = [REVIEW_REQUIRED] * (len(EXCEL_COLUMNS) - 1) + [""]

    sheet.append(row)

    # Highlight entire row in red
    _red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    last_row = sheet.max_row
    for col_idx in range(1, len(EXCEL_COLUMNS) + 1):
        sheet.cell(row=last_row, column=col_idx).fill = _red

    workbook.save(excel_path)
    logging.info(
        f"Excel: low-confidence review row written to tab '{sheet_name}' in {EXCEL_FILE}."
    )


def _add_sheet_headers(sheet) -> None:
    """Write styled column headers to a freshly created cost-code sheet."""
    sheet.append(EXCEL_COLUMNS)

    # Dark-blue header row with white bold text
    fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)

    for cell in sheet[1]:
        cell.fill = cell.fill if False else fill   # apply fill to every header cell
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")

    # Set sensible column widths (one entry per column in EXCEL_COLUMNS order)
    widths = [14, 12, 14, 35, 35, 35, 16, 22, 20]
    for col_idx, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


# ============================================================
# SHAREPOINT UPLOAD
# ============================================================
def upload_to_sharepoint(
    client: GraphClient,
    pdf_bytes: bytes,
    job_number: str,
    ticket_number: str,
    date_str: str,
) -> str:
    """
    Upload the PDF to the configured SharePoint folder.

    File naming convention: [JobNumber]-[TicketNumber]-[Date].pdf
    Date is sanitised (slashes/spaces removed) to produce a safe filename.

    Returns:
        The SharePoint web URL of the uploaded file.
    """
    # Build a safe date token for the filename
    if date_str:
        safe_date = re.sub(r'[/\\\s\.\-]', '', date_str)
    else:
        safe_date = datetime.now().strftime("%Y%m%d")

    filename = f"{job_number}-{ticket_number}-{safe_date}.pdf"

    # Resolve the SharePoint site and document library
    site_id  = _sharepoint_site_id(client)
    drive_id = _sharepoint_drive_id(client, site_id)

    # Graph API path: PUT /sites/{site}/drives/{drive}/root:/{folder}/{file}:/content
    upload_url = (
        f"{GRAPH_BASE}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{SHAREPOINT_FOLDER}/{filename}:/content"
    )

    response  = client.put_binary(upload_url, pdf_bytes, content_type="application/pdf")
    web_url   = response.json().get("webUrl", "(unknown)")

    logging.info(f"SharePoint: uploaded '{filename}' → {web_url}")
    return web_url


def _sharepoint_site_id(client: GraphClient) -> str:
    """Retrieve the Graph site ID for the configured SharePoint hostname."""
    url  = f"{GRAPH_BASE}/sites/{SHAREPOINT_HOST}:/"
    data = client.get(url).json()
    return data["id"]


def _sharepoint_drive_id(client: GraphClient, site_id: str) -> str:
    """
    Find the 'Shared Documents' document library drive on the site.
    Falls back to the first available drive if the expected name is absent.
    """
    url    = f"{GRAPH_BASE}/sites/{site_id}/drives"
    drives = client.get(url).json().get("value", [])

    for drive in drives:
        if drive.get("name", "").lower() in ("shared documents", "documents"):
            return drive["id"]

    if not drives:
        raise RuntimeError("No document libraries found on SharePoint site.")

    logging.warning(
        f"'Shared Documents' drive not found; using '{drives[0]['name']}' instead."
    )
    return drives[0]["id"]


# ============================================================
# PER-EMAIL PROCESSING PIPELINE
# ============================================================
def process_email(client: GraphClient, email: dict, profiles: list[TicketProfile]) -> bool:
    """
    Process one email end-to-end:
        1. Download PDF attachment
        2. Convert PDF to images
        3. Extract QR code → job/location/cost-code
        4. Extract ticket data via OCR
        5. Write row to Excel
        6. Upload PDF to SharePoint
        7. Send confirmation reply
        8. Mark email as read

    Returns True on full success, False if any step fails.
    The email is intentionally NOT marked as read on failure so it
    will be retried the next time this script runs.
    """
    subject      = email.get("subject", "(no subject)")
    email_id     = email["id"]
    sender_email = (
        email.get("from", {}).get("emailAddress", {}).get("address", "")
    )

    logging.info(f"{'='*55}")
    logging.info(f"Processing: '{subject}'")

    # Iterate over every PDF attachment in the email
    for attachment in email["pdf_attachments"]:
        att_name = attachment.get("name", "attachment.pdf")

        try:
            # ---- 1. Download PDF ------------------------------------------------
            logging.info(f"  Downloading '{att_name}'...")
            pdf_bytes = get_attachment_content(client, email_id, attachment["id"])

            # ---- 2. Convert PDF to images ---------------------------------------
            logging.info("  Rendering PDF pages...")
            images = pdf_to_images(pdf_bytes, dpi=200)
            if not images:
                raise ValueError("PDF rendered 0 pages.")

            # ---- 3. QR code extraction ------------------------------------------
            logging.info("  Scanning for QR code...")
            qr_data = extract_qr_code(images)
            if not qr_data:
                logging.warning(
                    f"  QR DETECTION FAILED: {subject!r} — "
                    f"OpenCV and OCR fallback both returned no match across {len(images)} page(s)"
                )
                raise ValueError(
                    "No QR code matching XXXX-YYYY-CCCCCC-CC found in PDF."
                )
            logging.info(
                f"  QR → Job: {qr_data['job_number']}  "
                f"Location: {qr_data['location']}  "
                f"Cost Code: {qr_data['cost_code']}"
            )

            # ---- 4. OCR extraction and profile detection ------------------------
            logging.info("  Running OCR...")
            ticket_data = extract_ticket_data_ocr(images, profiles, subject=subject)

            # Handle unknown supplier — no matching profile found
            if ticket_data.get("_no_profile_match"):
                company_name = ticket_data.get("_company_name", "unknown")
                error_msg = (
                    f"UNKNOWN TICKET FORMAT: {company_name}. "
                    "No profile found. Please create a profile for this supplier."
                )
                logging.warning(f"  {error_msg}")
                log_error(subject, error_msg)
                reply_body = (
                    f"A ticket was received from an unrecognised supplier "
                    f"({company_name}). No processing profile exists for this "
                    "format. Please contact your administrator to add a profile "
                    "for this supplier."
                )
                try:
                    logging.info("  Sending unknown-supplier reply...")
                    send_review_reply(client, email_id, subject, body_override=reply_body, sender_email=sender_email)
                except Exception as reply_exc:
                    logging.warning(f"  Reply not sent (will continue): {reply_exc}")
                try:
                    logging.info("  Flagging email with 'REVIEW REQUIRED' category...")
                    flag_email_category(client, email_id)
                except Exception as cat_exc:
                    logging.warning(f"  Category flag not applied (will continue): {cat_exc}")
                try:
                    logging.info("  Renaming email subject...")
                    rename_email_subject(
                        client, email_id, f"REVIEW REQUIRED - {subject}",
                        current_subject=subject,
                    )
                except Exception as rename_exc:
                    logging.warning(f"  Subject rename not applied (will continue): {rename_exc}")
                moved_email_id = email_id
                try:
                    logging.info("  Moving email to 'REVIEW REQUIRED' folder...")
                    moved_email_id = move_email_to_review_folder(client, email_id)
                except Exception as move_exc:
                    logging.warning(f"  Email move not completed (will continue): {move_exc}")
                logging.info("  Marking email as unread in 'REVIEW REQUIRED' folder...")
                mark_email_as_unread(client, moved_email_id)
                return True   # handled; left unread for human review

            if not ticket_data.get("_ocr_confidence_passed", True):
                # OCR quality too low — flag the entire ticket for review
                failed_checks = ticket_data.get("_failed_checks", [])
                failed_detail = "; ".join(failed_checks) if failed_checks else "unknown"
                timestamp    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                error_msg    = (
                    f"REVIEW REQUIRED: {subject} - {timestamp} - "
                    f"Failed checks: {failed_detail}"
                )
                logging.warning(f"  TICKET FLAGGED FOR REVIEW: {subject!r}")
                for check in failed_checks:
                    logging.warning(f"    - {check}")
                log_error(subject, error_msg)

                # Send a rescan-request reply (best-effort — don't let a send
                # failure prevent the remaining steps from completing)
                try:
                    logging.info("  Sending review-required reply...")
                    send_review_reply(client, email_id, subject, sender_email=sender_email)
                except Exception as reply_exc:
                    logging.warning(f"  Review reply not sent (will continue): {reply_exc}")

                # Add "REVIEW REQUIRED" category in Exchange (best-effort)
                try:
                    logging.info("  Flagging email with 'REVIEW REQUIRED' category...")
                    flag_email_category(client, email_id)
                except Exception as cat_exc:
                    logging.warning(f"  Category flag not applied (will continue): {cat_exc}")

                # Rename subject before moving so the folder shows the new name.
                try:
                    logging.info("  Renaming email subject...")
                    rename_email_subject(
                        client, email_id, f"REVIEW REQUIRED - {subject}",
                        current_subject=subject,
                    )
                except Exception as rename_exc:
                    logging.warning(f"  Subject rename not applied (will continue): {rename_exc}")

                # Move email to the "REVIEW REQUIRED" folder (best-effort).
                # Capture the new ID — Exchange assigns a new message ID after
                # a move, so the original inbox ID is no longer valid.
                moved_email_id = email_id   # fallback: use original if move fails
                try:
                    logging.info("  Moving email to 'REVIEW REQUIRED' folder...")
                    moved_email_id = move_email_to_review_folder(client, email_id)
                except Exception as move_exc:
                    logging.warning(f"  Email move not completed (will continue): {move_exc}")

                # Mark as unread using the post-move ID so the message stands out
                # in the REVIEW REQUIRED folder.  If the move failed we still have
                # the original inbox ID as a best-effort fallback.
                logging.info("  Marking email as unread in 'REVIEW REQUIRED' folder...")
                mark_email_as_unread(client, moved_email_id)
                return True   # processed; left unread in review folder for human attention

            ticket_number = ticket_data.get("ticket_number", "")
            logging.info(
                f"  OCR → Ticket: {ticket_number or 'N/A'}  "
                f"Date: {ticket_data.get('date', 'N/A')}  "
                f"Facility: {ticket_data.get('facility', 'N/A')}  "
                f"Customer: {ticket_data.get('customer', 'N/A')}  "
                f"Net Tons: {ticket_data.get('net_tons', 'N/A')}"
            )

            # ---- 5. Duplicate ticket check --------------------------------------
            duplicate_tab = _find_duplicate_ticket(ticket_number)
            if duplicate_tab:
                logging.warning(
                    f"  DUPLICATE TICKET detected: {ticket_number!r} "
                    f"already exists in tab '{duplicate_tab}'"
                )
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_error(
                    subject,
                    f"DUPLICATE TICKET: {ticket_number} - {subject} - {timestamp}",
                )

                # Write full row with orange highlight and Flag = "DUPLICATE TICKET"
                logging.info("  Writing duplicate row to Excel...")
                write_to_excel(qr_data, ticket_data, flag="DUPLICATE TICKET")

                # Send duplicate-notice reply (best-effort)
                try:
                    logging.info("  Sending duplicate-ticket reply...")
                    send_duplicate_reply(client, email_id, ticket_number, subject=subject, sender_email=sender_email)
                except Exception as reply_exc:
                    logging.warning(f"  Duplicate reply not sent (will continue): {reply_exc}")

                # Rename subject before moving so the folder shows the new name.
                try:
                    logging.info("  Renaming email subject...")
                    _material = ticket_data.get("material", "")
                    _date     = ticket_data.get("date", "")
                    rename_email_subject(
                        client,
                        email_id,
                        f"DUPLICATE - Ticket {ticket_number} - {_material} - {_date}",
                    )
                except Exception as rename_exc:
                    logging.warning(f"  Subject rename not applied (will continue): {rename_exc}")

                # Move to REVIEW REQUIRED folder and mark unread (best-effort)
                moved_email_id = email_id
                try:
                    logging.info("  Moving duplicate email to 'REVIEW REQUIRED' folder...")
                    moved_email_id = move_email_to_review_folder(client, email_id)
                except Exception as move_exc:
                    logging.warning(f"  Email move not completed (will continue): {move_exc}")

                logging.info("  Marking duplicate email as unread in 'REVIEW REQUIRED' folder...")
                mark_email_as_unread(client, moved_email_id)
                return True   # processed; left unread in review folder for human attention

            # ---- 6. Write to Excel ----------------------------------------------
            logging.info("  Writing to Excel...")
            write_to_excel(qr_data, ticket_data)

            # ---- 7. Upload PDF to SharePoint ------------------------------------
            # SHAREPOINT UPLOAD DISABLED - RESUME LATER
            # logging.info("  Uploading to SharePoint...")
            # upload_to_sharepoint(
            #     client,
            #     pdf_bytes,
            #     qr_data["job_number"],
            #     ticket_data.get("ticket_number", "UNKNOWN"),
            #     ticket_data.get("date", ""),
            # )

            # ---- 8. Rename email subject ----------------------------------------
            try:
                logging.info("  Renaming email subject...")
                _material = ticket_data.get("material", "")
                _date     = ticket_data.get("date", "")
                rename_email_subject(
                    client,
                    email_id,
                    f"Ticket {ticket_number} - {_material} - {_date}",
                )
            except Exception as rename_exc:
                logging.warning(f"  Subject rename not applied (will continue): {rename_exc}")

            # ---- 9. Send confirmation reply (best-effort) -----------------------
            try:
                logging.info("  Sending confirmation email...")
                send_reply_email(
                    client,
                    email_id,
                    ticket_number or "UNKNOWN",
                    qr_data["job_number"],
                )
            except Exception as reply_exc:
                logging.warning(
                    f"  Confirmation email not sent (will continue): {reply_exc}"
                )

            # ---- 10. Mark as read -----------------------------------------------
            logging.info("  Marking email as read...")
            mark_email_as_read(client, email_id)

            logging.info(f"  Done: '{subject}'")
            return True

        except Exception as exc:
            detail = f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"
            logging.error(f"  FAILED on '{att_name}': {exc}")
            log_error(subject, f"Attachment '{att_name}': {detail}")
            return False   # Email stays unread → will retry next run

    # No PDF attachments were processed (shouldn't happen given the filter above)
    return False


# ============================================================
# ENTRY POINT
# ============================================================
def main() -> None:
    """
    Main driver:
        - Validates .env credentials
        - Authenticates with Graph API
        - Fetches unread emails with PDF attachments
        - Processes each one
        - Reports summary counts
    """
    configure_logging()

    logging.info("=" * 55)
    logging.info("Ticket Processing Automation — starting")
    logging.info(f"Run time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("=" * 55)

    # Fail fast if any required credential is missing
    missing_vars = [
        v for v in ("AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID")
        if not os.getenv(v)
    ]
    if missing_vars:
        raise EnvironmentError(
            f"Missing required .env variables: {', '.join(missing_vars)}\n"
            f"  .env loaded from: {_env_path}\n"
            f"  Make sure the file exists at that path and contains the required keys."
        )

    # Confirm credentials are being read from the right place
    tenant_preview = os.getenv("AZURE_TENANT_ID", "")[:4]
    logging.info(f"Credentials loaded from: {_env_path}")
    logging.info(f"AZURE_TENANT_ID starts with: {tenant_preview!r}")

    # Load and validate supplier ticket profiles
    profiles = load_ticket_profiles()
    logging.info(f"Loaded {len(profiles)} ticket profile(s).")

    # Authenticate
    client = GraphClient()

    # Fetch target emails
    emails = get_unread_emails_with_pdf(client)

    if not emails:
        logging.info("No unread emails with PDF attachments. Nothing to do.")
        return

    # Process each email, track outcomes
    successes, failures = 0, 0
    for email in emails:
        if process_email(client, email, profiles):
            successes += 1
        else:
            failures += 1

    logging.info("=" * 55)
    logging.info(
        f"Complete — success: {successes}  failed: {failures}  "
        f"(failed emails remain unread for retry)"
    )
    logging.info("=" * 55)


if __name__ == "__main__":
    main()

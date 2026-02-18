"""
Outlook Mail Scraper - FastAPI Backend (IMAP)
Connects to Outlook via IMAP to fetch emails, attachments, and metadata.
"""

import os
import io
import csv
import json
import re
import base64
import logging
import time
import mimetypes
import zipfile
import imaplib
import email
import email.header
import email.utils
import email.policy
import asyncio
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from typing import Optional, List
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, field_validator

# ─── Configuration ───────────────────────────────────────────────────────────

FRONTEND_URL = os.getenv("FRONTEND_URL", "http://localhost:5173")
IMAP_HOST = os.getenv("IMAP_SERVER", os.getenv("IMAP_HOST", "imap.gmail.com"))
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))

ATTACHMENTS_DIR = Path("./attachments")
ATTACHMENTS_DIR.mkdir(exist_ok=True)
_METADATA_FILE = ATTACHMENTS_DIR / "_metadata.json"

IMAGE_TYPES = {"image/jpeg", "image/png", "image/gif", "image/webp", "image/svg+xml", "image/bmp"}
PDF_TYPES = {"application/pdf"}
DOC_TYPES = {
    "application/msword", "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-powerpoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "text/plain", "text/csv",
}


def _read_attachment_metadata() -> dict:
    if _METADATA_FILE.exists():
        try:
            return json.loads(_METADATA_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _write_attachment_metadata(metadata: dict) -> None:
    _METADATA_FILE.write_text(json.dumps(metadata, indent=2, default=str), encoding="utf-8")


def _save_attachment_with_metadata(
    content: bytes, uid: int, index: int, original_name: str, content_type: str,
    email_subject: str = "", email_sender: str = "", email_date: str = "",
) -> str:
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", original_name)
    filename = f"{uid}_{index}_{safe_name}"
    filepath = ATTACHMENTS_DIR / filename
    filepath.write_bytes(content)
    metadata = _read_attachment_metadata()
    metadata[filename] = {
        "original_name": original_name,
        "content_type": content_type,
        "size": len(content),
        "email_subject": email_subject,
        "email_sender": email_sender,
        "email_date": email_date,
        "saved_at": datetime.now(timezone.utc).isoformat(),
    }
    _write_attachment_metadata(metadata)
    return filename


def _classify_file_type(content_type: str) -> str:
    if content_type in IMAGE_TYPES:
        return "image"
    if content_type in PDF_TYPES:
        return "pdf"
    if content_type in DOC_TYPES:
        return "document"
    return "other"


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── In-memory credential & connection store ─────────────────────────────────

_credentials: dict = {}
_imap_connection: imaplib.IMAP4_SSL | None = None
_imap_lock = asyncio.Lock()


# ─── IMAP Connection Management ─────────────────────────────────────────────

def _imap_connect() -> imaplib.IMAP4_SSL:
    """Create and authenticate a fresh IMAP connection."""
    if not _credentials.get("email") or not _credentials.get("password"):
        raise HTTPException(status_code=401, detail="Not authenticated. Please login first.")
    try:
        conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        conn.login(_credentials["email"], _credentials["password"])
        return conn
    except imaplib.IMAP4.error as e:
        raise HTTPException(status_code=401, detail=f"IMAP authentication failed: {e}")


def _get_imap() -> imaplib.IMAP4_SSL:
    """Return a working IMAP connection, reconnecting if stale."""
    global _imap_connection
    if _imap_connection is not None:
        try:
            _imap_connection.noop()
            return _imap_connection
        except Exception:
            try:
                _imap_connection.logout()
            except Exception:
                pass
            _imap_connection = None
    _imap_connection = _imap_connect()
    return _imap_connection


async def _imap_op(func, *args, **kwargs):
    """Run a synchronous IMAP function in a thread, with reconnect-on-failure."""
    async with _imap_lock:
        def _execute():
            global _imap_connection
            for attempt in range(2):
                try:
                    conn = _get_imap()
                    return func(conn, *args, **kwargs)
                except (imaplib.IMAP4.abort, OSError, ConnectionResetError):
                    _imap_connection = None
                    if attempt == 1:
                        raise
            raise HTTPException(status_code=500, detail="IMAP connection failed after retry")
        return await asyncio.to_thread(_execute)


# ─── Email ID encoding (folder + UID → opaque ID) ───────────────────────────

def _encode_email_id(folder: str, uid: str) -> str:
    raw = f"{folder}\x00{uid}"
    return base64.urlsafe_b64encode(raw.encode()).decode().rstrip("=")


def _decode_email_id(email_id: str) -> tuple[str, str]:
    padding = 4 - len(email_id) % 4
    if padding != 4:
        email_id += "=" * padding
    raw = base64.urlsafe_b64decode(email_id).decode()
    folder, uid = raw.split("\x00", 1)
    return folder, uid


# ─── Email parsing helpers ───────────────────────────────────────────────────

def _decode_header_value(raw) -> str:
    if raw is None:
        return ""
    parts = email.header.decode_header(raw)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return " ".join(decoded)


def _parse_address(addr_str: str) -> tuple[str, str]:
    if not addr_str:
        return ("Unknown", "")
    decoded = _decode_header_value(addr_str)
    name, address = email.utils.parseaddr(decoded)
    return (name or address.split("@")[0] if address else "Unknown", address)


def _parse_address_list(header_value: str) -> list[tuple[str, str]]:
    if not header_value:
        return []
    decoded = _decode_header_value(header_value)
    addresses = email.utils.getaddresses([decoded])
    return [(name or addr.split("@")[0], addr) for name, addr in addresses if addr]


def _get_body(msg: email.message.Message) -> tuple[str, str]:
    """Extract (text_body, html_body) from a parsed email message."""
    text_body = ""
    html_body = ""

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if "attachment" in disp:
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            content = payload.decode(charset, errors="replace")
            if ctype == "text/plain" and not text_body:
                text_body = content
            elif ctype == "text/html" and not html_body:
                html_body = content
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            content = payload.decode(charset, errors="replace")
            if msg.get_content_type() == "text/html":
                html_body = content
            else:
                text_body = content

    return text_body, html_body


def _get_attachments(msg: email.message.Message) -> list[dict]:
    """Extract attachment metadata from a parsed email message."""
    attachments = []
    idx = 0
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        filename = part.get_filename()
        if filename:
            filename = _decode_header_value(filename)

        is_attachment = "attachment" in disp or (
            filename and part.get_content_maintype() not in ("multipart",)
        )
        if not is_attachment:
            continue

        payload = part.get_payload(decode=True)
        attachments.append({
            "index": idx,
            "name": filename or f"attachment_{idx}",
            "content_type": part.get_content_type(),
            "size": len(payload) if payload else 0,
            "is_inline": "inline" in disp,
        })
        idx += 1
    return attachments


def _get_attachment_by_index(msg: email.message.Message, target_idx: int):
    """Get attachment content by index. Returns (filename, content_type, bytes)."""
    idx = 0
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        filename = part.get_filename()
        if filename:
            filename = _decode_header_value(filename)
        is_attachment = "attachment" in disp or (
            filename and part.get_content_maintype() not in ("multipart",)
        )
        if not is_attachment:
            continue
        if idx == target_idx:
            payload = part.get_payload(decode=True) or b""
            return (
                filename or f"attachment_{idx}",
                part.get_content_type() or "application/octet-stream",
                payload,
            )
        idx += 1
    return None


def _has_attachments(msg: email.message.Message) -> bool:
    """Check if message has attachments by examining Content-Disposition headers."""
    for part in msg.walk():
        disp = str(part.get("Content-Disposition", ""))
        if "attachment" in disp:
            return True
        filename = part.get_filename()
        if filename and part.get_content_maintype() not in ("multipart",):
            return True
    return False


def _parse_importance(msg: email.message.Message) -> str:
    imp = (msg.get("Importance") or "").lower()
    if imp in ("high", "low"):
        return imp
    xpri = msg.get("X-Priority", "")
    if xpri.startswith("1") or xpri.startswith("2"):
        return "high"
    if xpri.startswith("4") or xpri.startswith("5"):
        return "low"
    return "normal"


def _parse_date(msg: email.message.Message) -> str:
    raw = msg.get("Date", "")
    try:
        dt = email.utils.parsedate_to_datetime(raw)
        return dt.isoformat()
    except Exception:
        return raw


def _to_imap_date(date_str: str) -> str:
    """Convert YYYY-MM-DD to DD-Mon-YYYY for IMAP SEARCH."""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.strftime("%d-%b-%Y")


# ─── IMAP folder parsing ────────────────────────────────────────────────────

def _decode_imap_utf7(data: bytes) -> str:
    """Decode IMAP modified UTF-7 folder names (RFC 3501 section 5.1.3).

    IMAP modified UTF-7 differs from standard UTF-7:
    - '&' is the shift character (not '+')
    - '&-' encodes a literal '&'
    - Non-ASCII runs use modified base64 with ',' instead of '/'
    - Encoded as UTF-16BE inside the base64 sections
    """
    result = []
    i = 0
    while i < len(data):
        if data[i:i + 1] == b'&':
            dash = data.find(b'-', i + 1)
            if dash == -1:
                result.append(data[i:].decode("ascii", errors="replace"))
                break
            if dash == i + 1:
                result.append('&')
            else:
                encoded = data[i + 1:dash].replace(b',', b'/')
                pad = 4 - len(encoded) % 4
                if pad != 4:
                    encoded += b'=' * pad
                decoded = base64.b64decode(encoded).decode('utf-16-be')
                result.append(decoded)
            i = dash + 1
        else:
            result.append(chr(data[i]))
            i += 1
    return ''.join(result)


_FOLDER_RE = re.compile(rb'\(([^)]*)\)\s+"([^"]*)"\s+(.+)')


def _parse_folder_line(line: bytes) -> dict | None:
    match = _FOLDER_RE.match(line)
    if not match:
        return None
    flags, delimiter, name = match.groups()
    name = name.strip()
    if name.startswith(b'"') and name.endswith(b'"'):
        name = name[1:-1]
    try:
        folder_name = _decode_imap_utf7(name).rstrip()
    except Exception:
        folder_name = name.decode("ascii", errors="replace").rstrip()
    flags_str = flags.decode("ascii", errors="replace")
    return {"name": folder_name, "flags": flags_str}


# ─── IMAP FETCH response parsing ────────────────────────────────────────────

_UID_RE = re.compile(rb'UID (\d+)')
_FLAGS_RE = re.compile(rb'FLAGS \(([^)]*)\)')


def _parse_fetch_response(data: list) -> list[tuple[int, set, bytes]]:
    """Parse FETCH response into list of (uid, flags_set, raw_email_bytes)."""
    results = []
    for item in data:
        if not isinstance(item, tuple):
            continue
        metadata, raw_bytes = item
        uid_m = _UID_RE.search(metadata)
        flags_m = _FLAGS_RE.search(metadata)
        if not uid_m:
            continue
        uid = int(uid_m.group(1))
        flags = set()
        if flags_m:
            flags = {f.strip() for f in flags_m.group(1).decode("ascii", errors="replace").split() if f.strip()}
        results.append((uid, flags, raw_bytes))
    return results


# ─── Rate Limiter ────────────────────────────────────────────────────────────

RATE_LIMIT_MAX = 100
RATE_LIMIT_WINDOW = 60
_rate_limit_store: dict[str, list[float]] = defaultdict(list)


def _check_rate_limit(client_ip: str) -> None:
    now = time.monotonic()
    window_start = now - RATE_LIMIT_WINDOW
    timestamps = _rate_limit_store[client_ip]
    _rate_limit_store[client_ip] = [t for t in timestamps if t > window_start]
    if len(_rate_limit_store[client_ip]) >= RATE_LIMIT_MAX:
        raise HTTPException(status_code=429, detail="Rate limit exceeded. Try again later.")
    _rate_limit_store[client_ip].append(now)


# ─── App Setup ───────────────────────────────────────────────────────────────

app = FastAPI(
    title="Outlook Mail Scraper",
    description="Scrape and browse Outlook emails via IMAP",
    version="2.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[FRONTEND_URL, "http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def rate_limit_middleware(request: Request, call_next):
    client_ip = request.client.host if request.client else "unknown"
    _check_rate_limit(client_ip)
    return await call_next(request)


# ─── Pydantic Models ────────────────────────────────────────────────────────

class LoginRequest(BaseModel):
    email: str
    password: str


class SenderInfo(BaseModel):
    name: str = "Unknown"
    email: str = ""


class RecipientInfo(BaseModel):
    name: str = ""
    email: str = ""


class AttachmentInfo(BaseModel):
    id: str
    name: str = "unnamed"
    content_type: str = ""
    size: int = 0
    is_inline: bool = False


class HeaderInfo(BaseModel):
    name: str
    value: str


class FlagInfo(BaseModel):
    flag_status: str = "notFlagged"


class UserInfo(BaseModel):
    name: Optional[str] = None
    email: Optional[str] = None
    job_title: Optional[str] = None


class AuthStatus(BaseModel):
    authenticated: bool
    user: Optional[UserInfo] = None


class LogoutResponse(BaseModel):
    message: str


class EmailSummary(BaseModel):
    id: str
    subject: str
    sender: str
    sender_email: str
    received: str
    preview: str
    is_read: bool
    has_attachments: bool
    importance: str
    folder: Optional[str] = None
    categories: List[str] = []


class EmailListResponse(BaseModel):
    emails: List[EmailSummary]
    total: int
    next_link: Optional[str] = None
    skip: int
    top: int


class EmailDetail(BaseModel):
    id: str
    subject: str
    sender: SenderInfo
    to_recipients: List[RecipientInfo]
    cc_recipients: List[RecipientInfo]
    bcc_recipients: List[RecipientInfo]
    received: str
    sent: str
    body_html: str
    body_text: str
    is_read: bool
    has_attachments: bool
    importance: str
    internet_message_id: str
    conversation_id: str
    categories: List[str]
    flag: FlagInfo
    attachments: List[AttachmentInfo]
    headers: List[HeaderInfo]


class FolderInfo(BaseModel):
    id: str
    name: str
    total_count: int
    unread_count: int
    child_folder_count: int


class ScrapedAttachmentInfo(BaseModel):
    name: str = ""
    content_type: str = ""
    size: int = 0
    is_inline: bool = False
    saved_path: Optional[str] = None
    filename: Optional[str] = None
    download_url: Optional[str] = None
    preview_url: Optional[str] = None


class StoredAttachmentInfo(BaseModel):
    filename: str
    original_name: str = ""
    content_type: str = ""
    file_type: str = "other"
    size: int = 0
    email_subject: str = ""
    email_sender: str = ""
    email_date: str = ""
    saved_at: str = ""
    download_url: str = ""
    preview_url: str = ""


class ScrapedEmailData(BaseModel):
    id: str
    subject: str = ""
    sender_name: str = ""
    sender_email: str = ""
    to: List[RecipientInfo] = []
    cc: List[RecipientInfo] = []
    received: str = ""
    sent: str = ""
    body_type: str = ""
    body: str = ""
    is_read: bool = False
    has_attachments: bool = False
    importance: str = ""
    internet_message_id: str = ""
    conversation_id: str = ""
    categories: List[str] = []
    attachments: List[ScrapedAttachmentInfo] = []


class ScrapeRequest(BaseModel):
    folder_id: Optional[str] = None
    from_date: Optional[str] = None
    to_date: Optional[str] = None
    sender_filter: Optional[str] = None
    subject_filter: Optional[str] = None
    search: Optional[str] = None
    max_results: int = Field(default=50, ge=1, le=500)
    include_attachments: bool = True

    @field_validator("from_date", "to_date")
    @classmethod
    def validate_date_format(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        try:
            datetime.strptime(v, "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"Invalid date format: '{v}'. Expected YYYY-MM-DD.")
        return v


class ScrapeResult(BaseModel):
    total_scraped: int
    emails: List[ScrapedEmailData]
    exported_at: str


class FolderStatInfo(BaseModel):
    name: str
    total: int
    unread: int


class TopSenderInfo(BaseModel):
    name: str
    email: str
    count: int


class MailboxStats(BaseModel):
    total_emails: int
    total_unread: int
    emails_last_7_days: int
    folder_stats: List[FolderStatInfo]
    top_senders: List[TopSenderInfo]


class HealthResponse(BaseModel):
    status: str
    configured: bool
    timestamp: str


# ─── Auth Endpoints ─────────────────────────────────────────────────────────

@app.post("/auth/login")
async def login(request: LoginRequest):
    """Test IMAP credentials and store them if valid."""
    try:
        conn = await asyncio.to_thread(
            lambda: imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        )
        await asyncio.to_thread(conn.login, request.email, request.password)
        await asyncio.to_thread(conn.logout)
    except imaplib.IMAP4.error as e:
        raise HTTPException(status_code=401, detail=f"Authentication failed: {e}")

    _credentials["email"] = request.email
    _credentials["password"] = request.password
    return {"message": "Login successful", "email": request.email}


@app.get("/auth/status", response_model=AuthStatus)
async def auth_status():
    if _credentials.get("email"):
        addr = _credentials["email"]
        return AuthStatus(
            authenticated=True,
            user=UserInfo(name=addr.split("@")[0], email=addr),
        )
    return AuthStatus(authenticated=False)


@app.post("/auth/logout", response_model=LogoutResponse)
async def logout():
    global _imap_connection
    if _imap_connection:
        try:
            _imap_connection.logout()
        except Exception:
            pass
        _imap_connection = None
    _credentials.clear()
    return LogoutResponse(message="Logged out successfully")


# ─── IMAP operation implementations ─────────────────────────────────────────

def _list_folders_impl(conn: imaplib.IMAP4_SSL) -> List[FolderInfo]:
    status, folder_list = conn.list()
    if status != "OK" or not folder_list:
        return []

    folders = []
    for line in folder_list:
        if not line:
            continue
        parsed = _parse_folder_line(line)
        if not parsed or "\\Noselect" in parsed["flags"]:
            continue

        fname = parsed["name"]
        try:
            st, data = conn.status(f'"{fname}"', "(MESSAGES UNSEEN)")
            if st == "OK" and data[0]:
                info = data[0].decode("ascii", errors="replace")
                msgs = int(re.search(r"MESSAGES (\d+)", info).group(1))
                unseen = int(re.search(r"UNSEEN (\d+)", info).group(1))
            else:
                msgs, unseen = 0, 0
        except Exception:
            msgs, unseen = 0, 0

        folders.append(FolderInfo(
            id=fname, name=fname,
            total_count=msgs, unread_count=unseen, child_folder_count=0,
        ))
    return folders


def _build_search_criteria(
    search: str = None, from_date: str = None, to_date: str = None,
    sender: str = None, has_attachments: bool = None, is_read: bool = None,
    subject_filter: str = None,
) -> str:
    parts = []
    if from_date:
        parts.append(f'SINCE {_to_imap_date(from_date)}')
    if to_date:
        to_dt = datetime.strptime(to_date, "%Y-%m-%d") + timedelta(days=1)
        parts.append(f'BEFORE {to_dt.strftime("%d-%b-%Y")}')
    if sender:
        parts.append(f'FROM "{sender}"')
    if search:
        parts.append(f'TEXT "{search}"')
    if subject_filter:
        parts.append(f'SUBJECT "{subject_filter}"')
    if is_read is True:
        parts.append("SEEN")
    elif is_read is False:
        parts.append("UNSEEN")
    if not parts:
        return "ALL"
    return "(" + " ".join(parts) + ")"


def _list_emails_impl(
    conn: imaplib.IMAP4_SSL,
    folder_id: str = None, search: str = None,
    from_date: str = None, to_date: str = None,
    sender: str = None, importance: str = None,
    has_attachments: bool = None, is_read: bool = None,
    skip: int = 0, top: int = 25,
) -> tuple[List[EmailSummary], int]:
    folder = folder_id or "INBOX"
    conn.select(f'"{folder}"', readonly=True)

    criteria = _build_search_criteria(
        search=search, from_date=from_date, to_date=to_date,
        sender=sender, has_attachments=None, is_read=is_read,
    )

    status, data = conn.uid("SEARCH", None, criteria)
    if status != "OK" or not data[0]:
        return [], 0

    all_uids = data[0].split()
    all_uids.reverse()  # newest first
    total = len(all_uids)

    page_uids = all_uids[skip:skip + top]
    if not page_uids:
        return [], total

    uid_str = b",".join(page_uids)
    status, fetch_data = conn.uid("FETCH", uid_str, "(UID FLAGS BODY.PEEK[]<0.4096>)")

    emails = []
    for uid_int, flags, raw_bytes in _parse_fetch_response(fetch_data):
        uid = str(uid_int)
        is_seen = "\\Seen" in flags

        try:
            msg = email.message_from_bytes(raw_bytes)
            subject = _decode_header_value(msg.get("Subject", ""))
            from_name, from_email_addr = _parse_address(msg.get("From", ""))
            date_str = _parse_date(msg)
            imp = _parse_importance(msg)
            has_att = _has_attachments(msg)
            text_body, _ = _get_body(msg)
            preview_text = " ".join(text_body.split())[:200] if text_body else ""
        except Exception:
            subject, from_name, from_email_addr = "(Parse error)", "Unknown", ""
            date_str, imp, has_att = "", "normal", False
            preview_text = ""

        eid = _encode_email_id(folder, uid)
        emails.append(EmailSummary(
            id=eid, subject=subject or "(No Subject)",
            sender=from_name, sender_email=from_email_addr,
            received=date_str, preview=preview_text,
            is_read=is_seen, has_attachments=has_att,
            importance=imp, folder=folder, categories=[],
        ))

    # Filter has_attachments post-fetch if requested
    if has_attachments is not None:
        emails = [e for e in emails if e.has_attachments == has_attachments]
    if importance:
        emails = [e for e in emails if e.importance == importance]

    return emails, total


def _get_email_impl(conn: imaplib.IMAP4_SSL, folder: str, uid: str) -> EmailDetail:
    conn.select(f'"{folder}"', readonly=True)

    status, data = conn.uid("FETCH", uid, "(UID FLAGS BODY.PEEK[])")
    if status != "OK" or not data or not isinstance(data[0], tuple):
        raise HTTPException(status_code=404, detail="Email not found")

    parsed = _parse_fetch_response(data)
    if not parsed:
        raise HTTPException(status_code=404, detail="Email not found")

    uid_int, flags, raw_bytes = parsed[0]
    msg = email.message_from_bytes(raw_bytes)

    subject = _decode_header_value(msg.get("Subject", ""))
    from_name, from_addr = _parse_address(msg.get("From", ""))
    to_list = _parse_address_list(msg.get("To", ""))
    cc_list = _parse_address_list(msg.get("Cc", ""))
    bcc_list = _parse_address_list(msg.get("Bcc", ""))
    date_str = _parse_date(msg)
    text_body, html_body = _get_body(msg)
    attachments = _get_attachments(msg)
    imp = _parse_importance(msg)
    is_seen = "\\Seen" in flags
    is_flagged = "\\Flagged" in flags

    headers = [HeaderInfo(name=k, value=_decode_header_value(v)) for k, v in msg.items()]

    eid = _encode_email_id(folder, str(uid_int))
    att_infos = [
        AttachmentInfo(
            id=str(a["index"]), name=a["name"],
            content_type=a["content_type"], size=a["size"],
            is_inline=a["is_inline"],
        )
        for a in attachments
    ]

    return EmailDetail(
        id=eid, subject=subject or "(No Subject)",
        sender=SenderInfo(name=from_name, email=from_addr),
        to_recipients=[RecipientInfo(name=n, email=e) for n, e in to_list],
        cc_recipients=[RecipientInfo(name=n, email=e) for n, e in cc_list],
        bcc_recipients=[RecipientInfo(name=n, email=e) for n, e in bcc_list],
        received=date_str, sent=date_str,
        body_html=html_body, body_text=text_body,
        is_read=is_seen,
        has_attachments=len(attachments) > 0,
        importance=imp,
        internet_message_id=msg.get("Message-ID", ""),
        conversation_id=msg.get("In-Reply-To", ""),
        categories=[], flag=FlagInfo(flag_status="flagged" if is_flagged else "notFlagged"),
        attachments=att_infos, headers=headers,
    )


def _download_attachment_impl(conn: imaplib.IMAP4_SSL, folder: str, uid: str, att_idx: int):
    conn.select(f'"{folder}"', readonly=True)
    status, data = conn.uid("FETCH", uid, "(BODY.PEEK[])")
    if status != "OK" or not data or not isinstance(data[0], tuple):
        raise HTTPException(status_code=404, detail="Email not found")

    parsed = _parse_fetch_response(data)
    if not parsed:
        raise HTTPException(status_code=404, detail="Email not found")

    _, _, raw_bytes = parsed[0]
    msg = email.message_from_bytes(raw_bytes)
    result = _get_attachment_by_index(msg, att_idx)
    if not result:
        raise HTTPException(status_code=404, detail="Attachment not found")
    return result


# ─── API Endpoints ──────────────────────────────────────────────────────────

@app.get("/api/folders", response_model=List[FolderInfo])
async def get_folders():
    return await _imap_op(_list_folders_impl)


@app.get("/api/emails", response_model=EmailListResponse)
async def list_emails(
    folder_id: Optional[str] = None,
    search: Optional[str] = None,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    sender: Optional[str] = None,
    importance: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    is_read: Optional[bool] = None,
    skip: int = 0,
    top: int = 25,
    order_by: str = "receivedDateTime desc",
):
    emails, total = await _imap_op(
        _list_emails_impl,
        folder_id=folder_id, search=search,
        from_date=from_date, to_date=to_date,
        sender=sender, importance=importance,
        has_attachments=has_attachments, is_read=is_read,
        skip=skip, top=top,
    )
    return EmailListResponse(
        emails=emails, total=total, next_link=None, skip=skip, top=top,
    )


@app.get("/api/emails/{email_id}", response_model=EmailDetail)
async def get_email(email_id: str):
    folder, uid = _decode_email_id(email_id)
    return await _imap_op(_get_email_impl, folder, uid)


@app.get("/api/emails/{email_id}/attachments/{attachment_id}")
async def download_attachment(email_id: str, attachment_id: str):
    folder, uid = _decode_email_id(email_id)
    try:
        att_idx = int(attachment_id)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid attachment index")

    filename, content_type, content = await _imap_op(
        _download_attachment_impl, folder, uid, att_idx,
    )
    return Response(
        content=content, media_type=content_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ─── Stored Attachment Endpoints ─────────────────────────────────────────────

def _validate_attachment_filename(filename: str) -> Path:
    """Validate filename to prevent path traversal, return resolved path."""
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="Invalid filename")
    filepath = ATTACHMENTS_DIR / filename
    if not filepath.exists() or not filepath.is_file():
        raise HTTPException(status_code=404, detail="Attachment not found")
    return filepath


@app.get("/api/attachments", response_model=List[StoredAttachmentInfo])
async def list_stored_attachments(file_type: Optional[str] = None):
    """List all stored attachments with metadata."""
    metadata = _read_attachment_metadata()
    results = []
    for filename, meta in metadata.items():
        filepath = ATTACHMENTS_DIR / filename
        if not filepath.exists():
            continue
        ct = meta.get("content_type", "")
        ft = _classify_file_type(ct)
        if file_type and ft != file_type:
            continue
        results.append(StoredAttachmentInfo(
            filename=filename,
            original_name=meta.get("original_name", filename),
            content_type=ct,
            file_type=ft,
            size=meta.get("size", 0),
            email_subject=meta.get("email_subject", ""),
            email_sender=meta.get("email_sender", ""),
            email_date=meta.get("email_date", ""),
            saved_at=meta.get("saved_at", ""),
            download_url=f"/api/attachments/{filename}",
            preview_url=f"/api/attachments/{filename}/preview",
        ))
    results.sort(key=lambda a: a.saved_at, reverse=True)
    return results


@app.get("/api/attachments/{filename}")
async def serve_attachment(filename: str):
    """Download a stored attachment."""
    filepath = _validate_attachment_filename(filename)
    content = filepath.read_bytes()
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    metadata = _read_attachment_metadata()
    original_name = metadata.get(filename, {}).get("original_name", filename)
    return Response(
        content=content, media_type=content_type,
        headers={"Content-Disposition": f'attachment; filename="{original_name}"'},
    )


@app.get("/api/attachments/{filename}/preview")
async def preview_attachment(filename: str):
    """Preview a stored attachment. Images/PDFs served inline; others return metadata."""
    filepath = _validate_attachment_filename(filename)
    metadata = _read_attachment_metadata()
    meta = metadata.get(filename, {})
    content_type = meta.get("content_type", mimetypes.guess_type(filename)[0] or "application/octet-stream")

    if content_type in IMAGE_TYPES or content_type in PDF_TYPES:
        content = filepath.read_bytes()
        return Response(
            content=content, media_type=content_type,
            headers={"Content-Disposition": f'inline; filename="{meta.get("original_name", filename)}"'},
        )

    return {
        "filename": filename,
        "original_name": meta.get("original_name", filename),
        "content_type": content_type,
        "file_type": _classify_file_type(content_type),
        "size": meta.get("size", 0),
        "email_subject": meta.get("email_subject", ""),
        "email_sender": meta.get("email_sender", ""),
        "email_date": meta.get("email_date", ""),
        "download_url": f"/api/attachments/{filename}",
        "preview_available": False,
    }


class ZipDownloadRequest(BaseModel):
    filenames: Optional[List[str]] = None  # None = all attachments


@app.post("/api/attachments/download-zip")
async def download_attachments_zip(request: ZipDownloadRequest):
    """Download multiple attachments as a ZIP file."""
    metadata = _read_attachment_metadata()
    if request.filenames:
        filenames = [f for f in request.filenames if f in metadata]
    else:
        filenames = list(metadata.keys())

    if not filenames:
        raise HTTPException(status_code=404, detail="No attachments found")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in filenames:
            filepath = ATTACHMENTS_DIR / fname
            if not filepath.exists():
                continue
            original_name = metadata.get(fname, {}).get("original_name", fname)
            # Avoid duplicate names in ZIP by prefixing with UID part
            arc_name = original_name
            if arc_name in [info.filename for info in zf.infolist()]:
                arc_name = f"{fname.split('_')[0]}_{original_name}"
            zf.write(filepath, arc_name)

    buf.seek(0)
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    return Response(
        content=buf.getvalue(), media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="attachments_{ts}.zip"'},
    )


# ─── Scrape & Export ────────────────────────────────────────────────────────

def _scrape_impl(
    conn: imaplib.IMAP4_SSL,
    folder_id: str = None, from_date: str = None, to_date: str = None,
    sender_filter: str = None, subject_filter: str = None,
    search: str = None,
    max_results: int = 50, include_attachments: bool = True,
) -> List[ScrapedEmailData]:
    folder = folder_id or "INBOX"
    conn.select(f'"{folder}"', readonly=True)

    criteria = _build_search_criteria(
        search=search, from_date=from_date, to_date=to_date,
        sender=sender_filter, subject_filter=subject_filter,
    )

    status, data = conn.uid("SEARCH", None, criteria)
    if status != "OK" or not data[0]:
        return []

    all_uids = data[0].split()
    all_uids.reverse()
    uids_to_fetch = all_uids[:max_results]

    if not uids_to_fetch:
        return []

    # Fetch in batches of 25
    all_emails = []
    for i in range(0, len(uids_to_fetch), 25):
        batch = uids_to_fetch[i:i + 25]
        uid_str = b",".join(batch)
        st, fetch_data = conn.uid("FETCH", uid_str, "(UID FLAGS BODY.PEEK[])")
        if st != "OK":
            continue

        for parsed_uid, flags, raw_bytes in _parse_fetch_response(fetch_data):
            msg = email.message_from_bytes(raw_bytes)
            subject = _decode_header_value(msg.get("Subject", ""))
            from_name, from_addr = _parse_address(msg.get("From", ""))
            to_list = _parse_address_list(msg.get("To", ""))
            cc_list = _parse_address_list(msg.get("Cc", ""))
            date_str = _parse_date(msg)
            text_body, html_body = _get_body(msg)
            atts = _get_attachments(msg)
            is_seen = "\\Seen" in flags

            scraped_atts = []
            if include_attachments:
                for a in atts:
                    att_info = ScrapedAttachmentInfo(
                        name=a["name"], content_type=a["content_type"],
                        size=a["size"], is_inline=a["is_inline"],
                    )
                    result = _get_attachment_by_index(msg, a["index"])
                    if result:
                        saved_filename = _save_attachment_with_metadata(
                            content=result[2], uid=parsed_uid, index=a["index"],
                            original_name=a["name"], content_type=a["content_type"],
                            email_subject=subject, email_sender=from_addr,
                            email_date=date_str,
                        )
                        att_info.saved_path = str(ATTACHMENTS_DIR / saved_filename)
                        att_info.filename = saved_filename
                        att_info.download_url = f"/api/attachments/{saved_filename}"
                        att_info.preview_url = f"/api/attachments/{saved_filename}/preview"
                    scraped_atts.append(att_info)
            else:
                scraped_atts = [
                    ScrapedAttachmentInfo(
                        name=a["name"], content_type=a["content_type"],
                        size=a["size"], is_inline=a["is_inline"],
                    )
                    for a in atts
                ]

            eid = _encode_email_id(folder, str(parsed_uid))
            all_emails.append(ScrapedEmailData(
                id=eid, subject=subject, sender_name=from_name,
                sender_email=from_addr,
                to=[RecipientInfo(name=n, email=e) for n, e in to_list],
                cc=[RecipientInfo(name=n, email=e) for n, e in cc_list],
                received=date_str, sent=date_str,
                body_type="html" if html_body else "text",
                body=html_body or text_body,
                is_read=is_seen, has_attachments=len(atts) > 0,
                importance=_parse_importance(msg),
                internet_message_id=msg.get("Message-ID", ""),
                conversation_id=msg.get("In-Reply-To", ""),
                categories=[], attachments=scraped_atts,
            ))

    return all_emails


@app.post("/api/scrape", response_model=ScrapeResult)
async def scrape_emails(request: ScrapeRequest):
    emails = await _imap_op(
        _scrape_impl,
        folder_id=request.folder_id, from_date=request.from_date,
        to_date=request.to_date, sender_filter=request.sender_filter,
        subject_filter=request.subject_filter, search=request.search,
        max_results=request.max_results,
        include_attachments=request.include_attachments,
    )
    return ScrapeResult(
        total_scraped=len(emails), emails=emails,
        exported_at=datetime.now(timezone.utc).isoformat(),
    )


@app.post("/api/export/json")
async def export_json(request: ScrapeRequest):
    result = await scrape_emails(request)
    content = json.dumps(result.model_dump(), indent=2, default=str)
    return Response(
        content=content, media_type="application/json",
        headers={
            "Content-Disposition": f'attachment; filename="outlook_export_{datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")}.json"'
        },
    )


@app.post("/api/export/csv")
async def export_csv(request: ScrapeRequest):
    result = await scrape_emails(request)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Subject", "Sender Name", "Sender Email", "To", "CC",
        "Received", "Sent", "Is Read", "Has Attachments",
        "Importance", "Categories", "Attachment Names",
        "Attachment Filenames", "Attachment Paths",
        "Internet Message ID", "Conversation ID",
    ])
    for em in result.emails:
        to_str = "; ".join([f"{r.name} <{r.email}>" for r in em.to])
        cc_str = "; ".join([f"{r.name} <{r.email}>" for r in em.cc])
        att_names = "; ".join([a.name for a in em.attachments])
        att_filenames = "; ".join([a.filename for a in em.attachments if a.filename])
        att_paths = "; ".join([a.saved_path for a in em.attachments if a.saved_path])
        writer.writerow([
            em.subject, em.sender_name, em.sender_email, to_str, cc_str,
            em.received, em.sent, em.is_read, em.has_attachments,
            em.importance, "; ".join(em.categories), att_names,
            att_filenames, att_paths,
            em.internet_message_id, em.conversation_id,
        ])
    return Response(
        content=output.getvalue(), media_type="text/csv",
        headers={
            "Content-Disposition": f'attachment; filename="outlook_export_{datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")}.csv"'
        },
    )


# ─── Stats ──────────────────────────────────────────────────────────────────

def _get_stats_impl(conn: imaplib.IMAP4_SSL) -> MailboxStats:
    folders = _list_folders_impl(conn)
    total_emails = sum(f.total_count for f in folders)
    total_unread = sum(f.unread_count for f in folders)
    folder_stats = [FolderStatInfo(name=f.name, total=f.total_count, unread=f.unread_count) for f in folders]

    # Emails in last 7 days
    conn.select('"INBOX"', readonly=True)
    week_ago = (datetime.now(timezone.utc) - timedelta(days=7)).strftime("%d-%b-%Y")
    st, data = conn.uid("SEARCH", None, f"SINCE {week_ago}")
    recent_count = len(data[0].split()) if st == "OK" and data[0] else 0

    # Top senders from last 50 emails
    st, data = conn.uid("SEARCH", None, "ALL")
    sender_counts: dict[str, TopSenderInfo] = {}
    if st == "OK" and data[0]:
        all_uids = data[0].split()
        all_uids.reverse()
        recent_uids = all_uids[:50]
        if recent_uids:
            uid_str = b",".join(recent_uids)
            st2, fetch_data = conn.uid("FETCH", uid_str, "(BODY.PEEK[HEADER.FIELDS (FROM)])")
            if st2 == "OK":
                for item in fetch_data:
                    if not isinstance(item, tuple):
                        continue
                    try:
                        header_bytes = item[1] if len(item) > 1 else item[0]
                        msg = email.message_from_bytes(header_bytes)
                        from_name, from_addr = _parse_address(msg.get("From", ""))
                        if from_addr and from_addr not in sender_counts:
                            sender_counts[from_addr] = TopSenderInfo(name=from_name, email=from_addr, count=0)
                        if from_addr:
                            sender_counts[from_addr].count += 1
                    except Exception:
                        continue

    top_senders = sorted(sender_counts.values(), key=lambda x: x.count, reverse=True)[:10]

    return MailboxStats(
        total_emails=total_emails, total_unread=total_unread,
        emails_last_7_days=recent_count,
        folder_stats=folder_stats, top_senders=top_senders,
    )


@app.get("/api/stats", response_model=MailboxStats)
async def get_stats():
    return await _imap_op(_get_stats_impl)


# ─── Health Check ───────────────────────────────────────────────────────────

@app.get("/health", response_model=HealthResponse)
async def health_check():
    return HealthResponse(
        status="healthy",
        configured=bool(_credentials.get("email")),
        timestamp=datetime.now(timezone.utc).isoformat(),
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

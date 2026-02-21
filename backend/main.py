import sys
if sys.version_info >= (3, 12):
    print("FATAL: Python 3.12+ not supported. Use Python 3.11.")
    sys.exit(1)

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

from pathlib import Path
ENV_FILE = Path(__file__).resolve().parent / ".env"
TEMPLATE_FILE = Path(__file__).resolve().parent / ".env.template"
if not ENV_FILE.exists() and TEMPLATE_FILE.exists():
    import shutil
    shutil.copy(TEMPLATE_FILE, ENV_FILE)
    print("INFO: Created .env from template — please edit with your credentials")

load_dotenv()

from fastapi import FastAPI, HTTPException, Request, Response, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, field_validator

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.interval import IntervalTrigger

from database import (
    create_tables as _create_db_tables, DB_PATH,
    SessionLocal, Candidate, JobRequisition, MatchResult, ScrapedEmail, Attachment,
    SchedulerConfig, Notification,
)
from parsers import extract_text_from_pdf, extract_text_from_docx, extract_entities
from pipeline import process_attachment_into_candidate
from matcher import run_match, _candidate_to_dict, _job_to_dict
import graph_client
import requests as http_requests

# ─── Configuration ───────────────────────────────────────────────────────────

FRONTEND_URL = os.getenv("FRONTEND_URL", "http://localhost:5173")
IMAP_HOST = os.getenv("IMAP_SERVER", os.getenv("IMAP_HOST", "imap.gmail.com"))
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))

ATTACHMENTS_DIR = Path(__file__).resolve().parent / "attachments"
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

    now = datetime.now(timezone.utc)

    # Write to _metadata.json (backward compat)
    metadata = _read_attachment_metadata()
    metadata[filename] = {
        "original_name": original_name,
        "content_type": content_type,
        "size": len(content),
        "email_subject": email_subject,
        "email_sender": email_sender,
        "email_date": email_date,
        "saved_at": now.isoformat(),
    }
    _write_attachment_metadata(metadata)

    # Write to SQLite (primary store)
    db = SessionLocal()
    try:
        existing = db.query(Attachment).filter(Attachment.filename == filename).first()
        if existing:
            existing.original_name = original_name
            existing.content_type = content_type
            existing.size = len(content)
            existing.email_uid = str(uid)
            existing.email_subject = email_subject
            existing.email_sender = email_sender
            existing.email_date = email_date
            existing.saved_at = now
        else:
            db.add(Attachment(
                filename=filename, original_name=original_name,
                content_type=content_type, size=len(content),
                email_uid=str(uid), email_subject=email_subject,
                email_sender=email_sender, email_date=email_date,
                saved_at=now,
            ))
        db.commit()
    except Exception:
        db.rollback()
        logger.exception("Failed to persist attachment metadata to SQLite")
    finally:
        db.close()

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

# ─── Session persistence ─────────────────────────────────────────────────────
# TODO: Encrypt credentials at rest (e.g. via OS keyring or Fernet).

SESSION_FILE = Path(__file__).resolve().parent / ".session.json"

_credentials: dict = {}
_imap_connection: imaplib.IMAP4_SSL | None = None
_imap_lock = asyncio.Lock()


def _save_session():
    """Persist current IMAP credentials to .session.json, preserving other keys."""
    try:
        data = {}
        if SESSION_FILE.exists():
            try:
                data = json.loads(SESSION_FILE.read_text())
            except (json.JSONDecodeError, OSError):
                data = {}
        data["email"] = _credentials.get("email", "")
        data["password"] = _credentials.get("password", "")
        SESSION_FILE.write_text(json.dumps(data))
    except Exception:
        logger.exception("Failed to save session file")


def _clear_session():
    """Remove .session.json."""
    try:
        SESSION_FILE.unlink(missing_ok=True)
    except Exception:
        logger.exception("Failed to remove session file")


def _restore_session():
    """Auto-restore credentials from .session.json and reconnect silently."""
    global _imap_connection
    if not SESSION_FILE.exists():
        return
    try:
        data = json.loads(SESSION_FILE.read_text())

        # Check for OAuth2 session first
        if data.get("microsoft_tokens"):
            token = graph_client.get_valid_token()
            if token:
                email_addr = data["microsoft_tokens"].get("email", "")
                _credentials["email"] = email_addr
                _credentials["auth_method"] = "oauth2"
                logger.info("OAuth2 session restored for %s", email_addr)
                return

        # Fall back to IMAP credentials
        em = data.get("email", "")
        pw = data.get("password", "")
        if not em or not pw:
            return
        # Test the credentials with a quick IMAP login
        conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        conn.login(em, pw)
        _credentials["email"] = em
        _credentials["password"] = pw
        _credentials["auth_method"] = "imap"
        _imap_connection = conn
        logger.info("IMAP session restored for %s", em)
    except Exception:
        logger.warning("Session restore failed — credentials may be stale")
        _clear_session()


def get_auth_method() -> str:
    """Return current auth method: 'imap', 'oauth2', or 'none'."""
    method = _credentials.get("auth_method", "")
    if method:
        return method
    # Infer from credentials state
    if _credentials.get("password"):
        return "imap"
    if graph_client.get_tokens():
        return "oauth2"
    return "none"


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

def decode_mime_header(raw) -> str:
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
    decoded = decode_mime_header(addr_str)
    name, address = email.utils.parseaddr(decoded)
    return (name or address.split("@")[0] if address else "Unknown", address)


def _parse_address_list(header_value: str) -> list[tuple[str, str]]:
    if not header_value:
        return []
    decoded = decode_mime_header(header_value)
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
            filename = decode_mime_header(filename)

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
            filename = decode_mime_header(filename)
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
    raw = decode_mime_header(msg.get("Date", ""))
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


# ─── Scheduler ────────────────────────────────────────────────────────────────

scheduler = AsyncIOScheduler()
_SCHEDULER_JOB_ID = "auto_scrape"


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


# ─── Database init (idempotent) ──────────────────────────────────────────────
_create_db_tables()


def _init_scheduler():
    """Start scheduler and restore job from DB config if enabled."""
    db = SessionLocal()
    try:
        cfg = db.query(SchedulerConfig).first()
        if cfg and cfg.enabled and cfg.interval_minutes and cfg.interval_minutes > 0:
            scheduler.add_job(
                run_scheduled_scrape,
                trigger=IntervalTrigger(minutes=cfg.interval_minutes),
                id=_SCHEDULER_JOB_ID,
                replace_existing=True,
            )
            logger.info("Scheduler restored: every %d min, folder=%s", cfg.interval_minutes, cfg.folder)
    except Exception:
        logger.exception("Failed to restore scheduler from DB")
    finally:
        db.close()
    scheduler.start()
    logger.info("APScheduler started")


@app.on_event("startup")
async def _on_startup():
    _init_scheduler()


@app.on_event("shutdown")
async def _on_shutdown():
    if scheduler.running:
        scheduler.shutdown(wait=False)


# ─── Migrate _metadata.json → SQLite attachments table (one-time) ────────────
def _migrate_attachment_metadata():
    """Seed SQLite attachments table from _metadata.json if it has entries not yet in DB."""
    metadata = _read_attachment_metadata()
    if not metadata:
        return
    db = SessionLocal()
    try:
        migrated = 0
        for filename, meta in metadata.items():
            if db.query(Attachment).filter(Attachment.filename == filename).first():
                continue
            # Extract email_uid from filename pattern "uid_index_name"
            uid_part = filename.split("_", 1)[0] if "_" in filename else ""
            saved_at_str = meta.get("saved_at", "")
            saved_at = None
            if saved_at_str:
                try:
                    saved_at = datetime.fromisoformat(saved_at_str)
                except (ValueError, TypeError):
                    pass
            db.add(Attachment(
                filename=filename,
                original_name=meta.get("original_name", filename),
                content_type=meta.get("content_type", ""),
                size=meta.get("size", 0),
                email_uid=uid_part,
                email_subject=meta.get("email_subject", ""),
                email_sender=meta.get("email_sender", ""),
                email_date=meta.get("email_date", ""),
                saved_at=saved_at or datetime.now(timezone.utc),
            ))
            migrated += 1
        if migrated:
            db.commit()
            logger.info("Migrated %d attachment records from _metadata.json to SQLite", migrated)
    except Exception:
        db.rollback()
        logger.exception("Attachment metadata migration failed")
    finally:
        db.close()

_migrate_attachment_metadata()

# ─── Auto-restore saved session ──────────────────────────────────────────────
_restore_session()


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
    _credentials["auth_method"] = "imap"
    _save_session()
    return {"message": "Login successful", "email": request.email}


@app.get("/auth/status", response_model=AuthStatus)
async def auth_status():
    if _credentials.get("email"):
        addr = _credentials["email"]
        return AuthStatus(
            authenticated=True,
            user=UserInfo(name=addr.split("@")[0], email=addr),
        )
    # Also check for OAuth2 tokens on disk even if not in memory yet
    tokens = graph_client.get_tokens()
    if tokens and tokens.get("email"):
        addr = tokens["email"]
        _credentials["email"] = addr
        _credentials["auth_method"] = "oauth2"
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
    _clear_session()
    return LogoutResponse(message="Logged out successfully")


# ─── Microsoft OAuth2 Endpoints ─────────────────────────────────────────────

@app.get("/auth/microsoft/url")
async def microsoft_auth_url():
    """Build Microsoft OAuth2 authorization URL."""
    client_id = os.getenv("MICROSOFT_CLIENT_ID", "")
    if not client_id:
        return {"error": "MICROSOFT_CLIENT_ID not configured in .env"}
    authority = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    redirect_uri = "http://localhost:8000/auth/microsoft/callback"
    scope = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access User.Read"
    params = (
        f"client_id={client_id}"
        f"&response_type=code"
        f"&redirect_uri={redirect_uri}"
        f"&scope={scope}"
        f"&response_mode=query"
    )
    return {"url": f"{authority}?{params}"}


@app.get("/auth/microsoft/callback")
async def microsoft_auth_callback(code: str = ""):
    """Exchange authorization code for tokens, then redirect to frontend."""
    from fastapi.responses import RedirectResponse

    if not code:
        raise HTTPException(status_code=400, detail="Missing authorization code")

    client_id = os.getenv("MICROSOFT_CLIENT_ID", "")
    if not client_id:
        raise HTTPException(status_code=500, detail="MICROSOFT_CLIENT_ID not configured")

    redirect_uri = "http://localhost:8000/auth/microsoft/callback"
    scope = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access User.Read"
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

    # Exchange code for tokens
    try:
        resp = http_requests.post(token_url, data={
            "client_id": client_id,
            "code": code,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
            "scope": scope,
        }, timeout=15)
        resp.raise_for_status()
        token_data = resp.json()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Token exchange failed: {e}")

    access_token = token_data.get("access_token", "")
    refresh_token = token_data.get("refresh_token", "")
    expires_in = token_data.get("expires_in", 3600)
    expires_at = (datetime.utcnow() + timedelta(seconds=expires_in)).isoformat()

    # Get user profile
    user_email = ""
    try:
        me_resp = http_requests.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=10,
        )
        if me_resp.ok:
            me_data = me_resp.json()
            user_email = me_data.get("mail") or me_data.get("userPrincipalName", "")
    except Exception:
        pass

    # Save tokens to .session.json
    graph_client.save_tokens({
        "access_token": access_token,
        "refresh_token": refresh_token,
        "access_token_expires_at": expires_at,
        "email": user_email,
    })

    # Update session file with auth_method
    try:
        data = json.loads(SESSION_FILE.read_text()) if SESSION_FILE.exists() else {}
        data["auth_method"] = "oauth2"
        SESSION_FILE.write_text(json.dumps(data))
    except Exception:
        pass

    # Set in-memory credentials
    _credentials["email"] = user_email
    _credentials["auth_method"] = "oauth2"

    return RedirectResponse(url=f"{FRONTEND_URL}?auth=success&email={user_email}")


@app.get("/auth/microsoft/status")
async def microsoft_auth_status():
    """Check if Microsoft OAuth2 session is active."""
    tokens = graph_client.get_tokens()
    if tokens and tokens.get("access_token"):
        return {
            "connected": True,
            "email": tokens.get("email", ""),
            "auth_method": "oauth2",
        }
    return {"connected": False, "email": None, "auth_method": None}


@app.post("/auth/microsoft/logout")
async def microsoft_logout():
    """Remove Microsoft OAuth2 tokens from session."""
    global _imap_connection
    try:
        if SESSION_FILE.exists():
            data = json.loads(SESSION_FILE.read_text())
            data.pop("microsoft_tokens", None)
            if data.get("auth_method") == "oauth2":
                data.pop("auth_method", None)
            SESSION_FILE.write_text(json.dumps(data))
    except Exception:
        pass
    # Clear in-memory state if it was OAuth2
    if _credentials.get("auth_method") == "oauth2":
        _credentials.clear()
    return {"logged_out": True}


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
    status, fetch_data = conn.uid("FETCH", uid_str, "(UID FLAGS BODY.PEEK[]<0.65536>)")

    emails = []
    for uid_int, flags, raw_bytes in _parse_fetch_response(fetch_data):
        uid = str(uid_int)
        is_seen = "\\Seen" in flags

        try:
            msg = email.message_from_bytes(raw_bytes)
            subject = decode_mime_header(msg.get("Subject", ""))
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

    subject = decode_mime_header(msg.get("Subject", ""))
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

    headers = [HeaderInfo(name=k, value=decode_mime_header(v)) for k, v in msg.items()]

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
    auth = get_auth_method()
    if auth == "oauth2":
        try:
            raw_folders = await asyncio.to_thread(graph_client.list_folders)
            return [
                FolderInfo(
                    id=f["id"], name=f["name"],
                    total_count=f.get("total_count", 0),
                    unread_count=f.get("unread_count", 0),
                    child_folder_count=0,
                )
                for f in raw_folders
            ]
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Graph API error: {e}")
    elif auth == "imap":
        return await _imap_op(_list_folders_impl)
    else:
        raise HTTPException(status_code=401, detail="Not authenticated")


def _load_emails_from_cache(
    folder_id: Optional[str] = None, search: Optional[str] = None,
    sender: Optional[str] = None, has_attachments: Optional[bool] = None,
    is_read: Optional[bool] = None, skip: int = 0, top: int = 25,
) -> tuple[List[EmailSummary], int]:
    """Load cached emails from SQLite."""
    db = SessionLocal()
    try:
        q = db.query(ScrapedEmail)
        if folder_id:
            q = q.filter(ScrapedEmail.folder == folder_id)
        if search:
            term = f"%{search}%"
            q = q.filter(
                (ScrapedEmail.subject.ilike(term))
                | (ScrapedEmail.sender.ilike(term))
                | (ScrapedEmail.sender_email.ilike(term))
            )
        if sender:
            q = q.filter(
                (ScrapedEmail.sender.ilike(f"%{sender}%"))
                | (ScrapedEmail.sender_email.ilike(f"%{sender}%"))
            )
        if has_attachments is not None:
            q = q.filter(ScrapedEmail.has_attachments == has_attachments)
        if is_read is not None:
            q = q.filter(ScrapedEmail.is_read == is_read)

        total = q.count()
        rows = q.order_by(ScrapedEmail.date.desc()).offset(skip).limit(top).all()
        emails = []
        for r in rows:
            body_preview = ""
            if r.body_text:
                body_preview = " ".join(r.body_text.split())[:200]
            elif r.body_html:
                import re as _re
                plain = _re.sub(r"<[^>]+>", " ", r.body_html)
                body_preview = " ".join(plain.split())[:200]
            emails.append(EmailSummary(
                id=r.uid, subject=r.subject or "(No Subject)",
                sender=r.sender or "", sender_email=r.sender_email or "",
                received=r.date or "", preview=body_preview,
                is_read=r.is_read or False, has_attachments=r.has_attachments or False,
                importance="normal", folder=r.folder or "INBOX", categories=[],
            ))
        return emails, total
    finally:
        db.close()


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
    source: Optional[str] = None,
):
    # ── Cache-only mode: return from SQLite without touching IMAP ──
    if source == "cache":
        emails, total = _load_emails_from_cache(
            folder_id=folder_id, search=search, sender=sender,
            has_attachments=has_attachments, is_read=is_read,
            skip=skip, top=top,
        )
        return EmailListResponse(
            emails=emails, total=total, next_link=None, skip=skip, top=top,
        )

    # ── Default: try IMAP first, fall back to cache on failure ──
    try:
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
    except Exception:
        logger.info("IMAP unavailable, falling back to cached emails")
        emails, total = _load_emails_from_cache(
            folder_id=folder_id, search=search, sender=sender,
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
    """List all stored attachments with metadata from SQLite."""
    db = SessionLocal()
    try:
        rows = db.query(Attachment).order_by(Attachment.saved_at.desc()).all()
        results = []
        for row in rows:
            filepath = ATTACHMENTS_DIR / row.filename
            if not filepath.exists():
                continue
            ct = row.content_type or ""
            ft = _classify_file_type(ct)
            if file_type and ft != file_type:
                continue
            results.append(StoredAttachmentInfo(
                filename=row.filename,
                original_name=row.original_name or row.filename,
                content_type=ct,
                file_type=ft,
                size=row.size or 0,
                email_subject=row.email_subject or "",
                email_sender=row.email_sender or "",
                email_date=row.email_date or "",
                saved_at=row.saved_at.isoformat() if row.saved_at else "",
                download_url=f"/api/attachments/{row.filename}",
                preview_url=f"/api/attachments/{row.filename}/preview",
            ))
        return results
    finally:
        db.close()


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
            subject = decode_mime_header(msg.get("Subject", ""))
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

                        # ── Candidate pipeline: process resumes/CVs ──
                        file_type = _classify_file_type(a["content_type"])
                        if file_type in ("document", "pdf"):
                            try:
                                process_attachment_into_candidate(
                                    attachment_filepath=str(ATTACHMENTS_DIR / saved_filename),
                                    email_uid=str(parsed_uid),
                                    email_body=text_body or "",
                                    email_sender=from_addr,
                                    email_subject=subject,
                                )
                            except Exception:
                                logger.exception(
                                    "Candidate pipeline failed for %s", saved_filename
                                )

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
                id=eid, subject=subject or "(No Subject)", sender_name=from_name or "Unknown",
                sender_email=from_addr or "",
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


def _persist_scraped_emails(emails: List[ScrapedEmailData], folder: str = "INBOX"):
    """Upsert scraped emails into SQLite for cache persistence."""
    db = SessionLocal()
    try:
        for em in emails:
            existing = db.query(ScrapedEmail).filter(ScrapedEmail.uid == em.id).first()
            body_text = em.body if em.body_type == "text" else ""
            body_html = em.body if em.body_type == "html" else ""
            att_count = len(em.attachments) if em.attachments else 0

            if existing:
                existing.folder = folder
                existing.subject = em.subject
                existing.sender = em.sender_name
                existing.sender_email = em.sender_email
                existing.date = em.received
                existing.body_text = body_text
                existing.body_html = body_html
                existing.has_attachments = em.has_attachments
                existing.attachment_count = att_count
                existing.is_read = em.is_read
                existing.scraped_at = datetime.now(timezone.utc)
            else:
                db.add(ScrapedEmail(
                    uid=em.id, folder=folder, subject=em.subject,
                    sender=em.sender_name, sender_email=em.sender_email,
                    date=em.received, body_text=body_text, body_html=body_html,
                    has_attachments=em.has_attachments, attachment_count=att_count,
                    is_read=em.is_read,
                ))
        db.commit()
        logger.info("Persisted %d scraped emails to SQLite", len(emails))
    except Exception:
        db.rollback()
        logger.exception("Failed to persist scraped emails")
    finally:
        db.close()


def _scrape_via_graph(
    folder_id: str = None, from_date: str = None, to_date: str = None,
    sender_filter: str = None, subject_filter: str = None, search: str = None,
    max_results: int = 50, include_attachments: bool = True,
) -> List[ScrapedEmailData]:
    """Scrape emails via Microsoft Graph API — mirrors _scrape_impl logic."""
    folder = folder_id
    if not folder:
        # Get the Inbox folder ID
        folders = graph_client.list_folders()
        inbox = next((f for f in folders if f["name"].lower() == "inbox"), None)
        if inbox:
            folder = inbox["id"]
        elif folders:
            folder = folders[0]["id"]
        else:
            return []

    result = graph_client.list_messages(folder, {
        "top": max_results,
        "from_date": from_date,
        "to_date": to_date,
        "sender_filter": sender_filter,
        "search": search or subject_filter,
    })

    all_emails = []
    for msg_summary in result.get("messages", []):
        msg_id = msg_summary["id"]
        full_msg = graph_client.get_message(msg_id)

        text_body = full_msg.get("body_text", "")
        html_body = full_msg.get("body_html", "")

        scraped_atts = []
        if include_attachments and full_msg.get("attachments"):
            for att in full_msg["attachments"]:
                att_info = ScrapedAttachmentInfo(
                    name=att.get("name", ""),
                    content_type=att.get("content_type", ""),
                    size=att.get("size", 0),
                    is_inline=att.get("is_inline", False),
                )

                content_bytes_b64 = att.get("content_bytes")
                if content_bytes_b64:
                    try:
                        content = base64.b64decode(content_bytes_b64)
                    except Exception:
                        content = None
                else:
                    try:
                        content = graph_client.download_attachment(msg_id, att["id"])
                    except Exception:
                        content = None

                if content:
                    saved_filename = _save_attachment_with_metadata(
                        content=content, uid=msg_id[:16], index=0,
                        original_name=att.get("name", "attachment"),
                        content_type=att.get("content_type", ""),
                        email_subject=full_msg.get("subject", ""),
                        email_sender=full_msg.get("sender_email", ""),
                        email_date=full_msg.get("date", ""),
                    )
                    att_info.saved_path = str(ATTACHMENTS_DIR / saved_filename)
                    att_info.filename = saved_filename
                    att_info.download_url = f"/api/attachments/{saved_filename}"
                    att_info.preview_url = f"/api/attachments/{saved_filename}/preview"

                    # Candidate pipeline: process resumes/CVs
                    file_type = _classify_file_type(att.get("content_type", ""))
                    if file_type in ("document", "pdf"):
                        try:
                            process_attachment_into_candidate(
                                attachment_filepath=str(ATTACHMENTS_DIR / saved_filename),
                                email_uid=msg_id,
                                email_body=text_body or "",
                                email_sender=full_msg.get("sender_email", ""),
                                email_subject=full_msg.get("subject", ""),
                            )
                        except Exception:
                            logger.exception(
                                "Candidate pipeline failed for %s", saved_filename
                            )

                scraped_atts.append(att_info)

        all_emails.append(ScrapedEmailData(
            id=msg_id,
            subject=full_msg.get("subject", ""),
            sender_name=full_msg.get("sender", ""),
            sender_email=full_msg.get("sender_email", ""),
            to=[RecipientInfo(name=r["name"], email=r["email"]) for r in full_msg.get("to", [])],
            cc=[RecipientInfo(name=r["name"], email=r["email"]) for r in full_msg.get("cc", [])],
            received=full_msg.get("date", ""),
            sent=full_msg.get("date", ""),
            body_type="html" if html_body else "text",
            body=html_body or text_body,
            is_read=full_msg.get("is_read", False),
            has_attachments=full_msg.get("has_attachments", False),
            importance=full_msg.get("importance", "normal"),
            internet_message_id=full_msg.get("internet_message_id", ""),
            conversation_id=full_msg.get("conversation_id", ""),
            categories=full_msg.get("categories", []),
            attachments=scraped_atts,
        ))

    return all_emails


@app.post("/api/scrape", response_model=ScrapeResult)
async def scrape_emails(request: ScrapeRequest):
    auth = get_auth_method()
    if auth == "oauth2":
        emails = await asyncio.to_thread(
            _scrape_via_graph,
            folder_id=request.folder_id, from_date=request.from_date,
            to_date=request.to_date, sender_filter=request.sender_filter,
            subject_filter=request.subject_filter, search=request.search,
            max_results=request.max_results,
            include_attachments=request.include_attachments,
        )
    elif auth == "imap":
        emails = await _imap_op(
            _scrape_impl,
            folder_id=request.folder_id, from_date=request.from_date,
            to_date=request.to_date, sender_filter=request.sender_filter,
            subject_filter=request.subject_filter, search=request.search,
            max_results=request.max_results,
            include_attachments=request.include_attachments,
        )
    else:
        raise HTTPException(status_code=401, detail="Not authenticated")

    # ── Persist scraped emails to SQLite (upsert) ──
    _persist_scraped_emails(emails, folder=request.folder_id or "INBOX")

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


# ─── Recruitment: Pydantic models ────────────────────────────────────────────

class JobCreateRequest(BaseModel):
    title: str
    jd_raw: Optional[str] = None
    required_skills: Optional[List[str]] = None
    min_exp: Optional[float] = None
    location: Optional[str] = None
    remote_ok: bool = False


class NotesUpdateRequest(BaseModel):
    notes: str


class TagsUpdateRequest(BaseModel):
    tags: List[str]


# ─── Recruitment: Candidate endpoints ───────────────────────────────────────

def _normalize_name(name: str) -> str:
    """Lowercase, strip whitespace and punctuation for duplicate comparison."""
    return re.sub(r'[^a-z ]', '', (name or "").lower()).strip()


def _email_domain(addr: str) -> str:
    """Extract the domain from an email address."""
    if not addr or "@" not in addr:
        return ""
    return addr.split("@")[1].lower().strip()


def _detect_duplicates(candidates: list) -> list:
    """
    Mark duplicate candidates.  Two candidates are duplicates if:
    - Their normalized names are identical, OR
    - They share the same email domain AND have similar names
      (one name is a subset of the other, first names match, or
      Jaccard similarity > 0.5 on words).

    Returns the input list with added 'duplicate_group_id' and 'is_duplicate' keys.
    The earliest created_at in each group is the "original" (is_duplicate=False).
    """
    # Sort by created_at ascending so earliest comes first
    sorted_cands = sorted(candidates, key=lambda c: c.get("created_at") or "")

    # Build lookup structures
    groups: dict[int, list[int]] = {}   # group_id -> [indices]
    cand_group: dict[int, int] = {}     # index -> group_id
    next_group = 1

    for i, ci in enumerate(sorted_cands):
        if i in cand_group:
            continue
        norm_i = _normalize_name(ci.get("name"))
        domain_i = _email_domain(ci.get("email") or "")
        words_i = set(norm_i.split()) if norm_i else set()

        for j in range(i + 1, len(sorted_cands)):
            if j in cand_group and cand_group.get(i) and cand_group[j] == cand_group[i]:
                continue  # already in same group
            cj = sorted_cands[j]
            norm_j = _normalize_name(cj.get("name"))
            domain_j = _email_domain(cj.get("email") or "")
            words_j = set(norm_j.split()) if norm_j else set()

            is_dup = False

            # Exact normalized name match
            if norm_i and norm_j and norm_i == norm_j:
                is_dup = True
            # Exact email match (different name spellings)
            elif ci.get("email") and cj.get("email") and ci["email"].lower() == cj["email"].lower():
                is_dup = True
            # Same email domain + similar name
            elif domain_i and domain_j and domain_i == domain_j and words_i and words_j:
                intersection = words_i & words_j
                union = words_i | words_j
                # Jaccard similarity > 0.5
                if union and len(intersection) / len(union) > 0.5:
                    is_dup = True
                # One name is a subset of the other (e.g. "Alice J" vs "Alice Johnson")
                elif words_i <= words_j or words_j <= words_i:
                    is_dup = True
                # First names match
                elif norm_i.split()[0] == norm_j.split()[0]:
                    is_dup = True

            if is_dup:
                gid_i = cand_group.get(i)
                gid_j = cand_group.get(j)
                if gid_i and not gid_j:
                    cand_group[j] = gid_i
                    groups[gid_i].append(j)
                elif gid_j and not gid_i:
                    cand_group[i] = gid_j
                    groups[gid_j].append(i)
                elif not gid_i and not gid_j:
                    gid = next_group
                    next_group += 1
                    cand_group[i] = gid
                    cand_group[j] = gid
                    groups[gid] = [i, j]
                # else both already assigned — may be different groups, merge:
                elif gid_i != gid_j:
                    # merge gid_j into gid_i
                    for idx in groups[gid_j]:
                        cand_group[idx] = gid_i
                    groups[gid_i].extend(groups[gid_j])
                    del groups[gid_j]

    # Apply flags
    for c in sorted_cands:
        c["duplicate_group_id"] = None
        c["is_duplicate"] = False

    for gid, indices in groups.items():
        for rank, idx in enumerate(indices):
            sorted_cands[idx]["duplicate_group_id"] = gid
            # First in the group (earliest created_at) is the original
            if rank > 0:
                sorted_cands[idx]["is_duplicate"] = True

    # Re-sort by created_at descending (newest first) to match original ordering
    sorted_cands.sort(key=lambda c: c.get("created_at") or "", reverse=True)
    return sorted_cands


@app.get("/api/candidates")
async def list_candidates(
    skill: Optional[str] = None,
    location: Optional[str] = None,
    name: Optional[str] = None,
    tag: Optional[str] = None,
):
    """List all candidates with optional filters and duplicate detection."""
    db = SessionLocal()
    try:
        q = db.query(Candidate)
        if name:
            q = q.filter(Candidate.name.ilike(f"%{name}%"))
        if location:
            q = q.filter(Candidate.location.ilike(f"%{location}%"))
        if skill:
            # JSON array stored as text — use LIKE for contains
            q = q.filter(Candidate.skills.ilike(f"%{skill}%"))
        if tag:
            q = q.filter(Candidate.tags.ilike(f"%{tag}%"))
        candidates = q.order_by(Candidate.created_at.desc()).all()
        result = [_candidate_to_dict(c) for c in candidates]
        return _detect_duplicates(result)
    finally:
        db.close()


def _read_resume_text(raw_resume_path: str | None) -> str:
    """Extract readable text from a resume file (PDF/DOCX/TXT)."""
    if not raw_resume_path:
        return ""
    try:
        p = Path(raw_resume_path)
        if not p.is_absolute():
            p = Path(__file__).resolve().parent / raw_resume_path
        if not p.exists() or not p.is_file():
            return ""
        suffix = p.suffix.lower()
        if suffix == ".pdf":
            return extract_text_from_pdf(p)
        elif suffix in (".docx", ".doc"):
            return extract_text_from_docx(p)
        else:
            return p.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""


@app.get("/api/candidates/{candidate_id}")
async def get_candidate(candidate_id: int):
    """Get a single candidate by ID with match history and source email."""
    db = SessionLocal()
    try:
        c = db.query(Candidate).filter(Candidate.id == candidate_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="Candidate not found")

        result = _candidate_to_dict(c)

        # Resume text
        result["resume_text"] = _read_resume_text(c.raw_resume_path)

        # Match history
        matches = (
            db.query(MatchResult, JobRequisition)
            .join(JobRequisition, MatchResult.job_id == JobRequisition.id)
            .filter(MatchResult.candidate_id == candidate_id)
            .order_by(MatchResult.score.desc())
            .all()
        )
        result["match_history"] = [
            {
                "job_id": mr.job_id,
                "job_title": job.title,
                "score": mr.score,
                "fit_level": mr.fit_level,
                "match_reasons": json.loads(mr.match_reasons) if mr.match_reasons else [],
                "matched_at": job.created_at.isoformat() if job.created_at else None,
            }
            for mr, job in matches
        ]

        # Source email
        source_email = None
        if c.source_email_uid:
            se = db.query(ScrapedEmail).filter(ScrapedEmail.uid == c.source_email_uid).first()
            if se:
                source_email = {
                    "subject": se.subject,
                    "sender": se.sender,
                    "sender_email": se.sender_email,
                    "date": se.date,
                    "body_text": se.body_text or "",
                }
        result["source_email"] = source_email

        return result
    finally:
        db.close()


@app.get("/api/candidates/{candidate_id}/resume-text")
async def get_candidate_resume_text(candidate_id: int):
    """Return raw resume text as plain text for the resume viewer."""
    db = SessionLocal()
    try:
        c = db.query(Candidate).filter(Candidate.id == candidate_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="Candidate not found")
        text = _read_resume_text(c.raw_resume_path)
        return Response(content=text, media_type="text/plain; charset=utf-8")
    finally:
        db.close()


@app.delete("/api/candidates/{candidate_id}")
async def delete_candidate(candidate_id: int):
    """Delete a candidate and their associated match results."""
    db = SessionLocal()
    try:
        c = db.query(Candidate).filter(Candidate.id == candidate_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="Candidate not found")
        db.query(MatchResult).filter(MatchResult.candidate_id == candidate_id).delete()
        db.delete(c)
        db.commit()
        return {"detail": f"Candidate {candidate_id} deleted"}
    finally:
        db.close()


@app.patch("/api/candidates/{candidate_id}/notes")
async def update_candidate_notes(candidate_id: int, body: NotesUpdateRequest):
    """Update the notes field for a candidate."""
    db = SessionLocal()
    try:
        c = db.query(Candidate).filter(Candidate.id == candidate_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="Candidate not found")
        c.notes = body.notes
        db.commit()
        return {"id": candidate_id, "notes": c.notes}
    finally:
        db.close()


@app.patch("/api/candidates/{candidate_id}/tags")
async def update_candidate_tags(candidate_id: int, body: TagsUpdateRequest):
    """Update the tags field for a candidate."""
    db = SessionLocal()
    try:
        c = db.query(Candidate).filter(Candidate.id == candidate_id).first()
        if not c:
            raise HTTPException(status_code=404, detail="Candidate not found")
        c.tags = json.dumps(body.tags)
        db.commit()
        return {"id": candidate_id, "tags": body.tags}
    finally:
        db.close()


# ─── Recruitment: Job endpoints ─────────────────────────────────────────────

@app.post("/api/jobs/upload-jd")
async def upload_jd(file: UploadFile = File(...)):
    """Parse a JD file (PDF/DOCX) and return extracted fields for preview."""
    import tempfile

    filename = (file.filename or "").lower()
    if not filename.endswith((".pdf", ".docx", ".doc")):
        raise HTTPException(status_code=400, detail="Only PDF and DOCX files are supported")

    # Save to temp file
    content = await file.read()
    suffix = Path(filename).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(content)
        tmp_path = Path(tmp.name)

    try:
        # Extract text
        if suffix == ".pdf":
            raw_text = extract_text_from_pdf(tmp_path)
        else:
            raw_text = extract_text_from_docx(tmp_path)

        if not raw_text.strip():
            raise HTTPException(status_code=422, detail="Could not extract text from file")

        # Run entity extraction
        entities = extract_entities(raw_text)

        # Check for remote indicators in the text
        remote_ok = bool(re.search(r'\b(remote|work from home|wfh|hybrid|telecommute)\b', raw_text, re.IGNORECASE))

        return {
            "raw_text": raw_text,
            "title": entities.get("titles", [None])[0] if entities.get("titles") else None,
            "required_skills": entities.get("skills", []),
            "min_exp": entities.get("years_exp"),
            "location": entities.get("locations", [None])[0] if entities.get("locations") else None,
            "remote_ok": remote_ok,
        }
    finally:
        tmp_path.unlink(missing_ok=True)


@app.post("/api/jobs")
async def create_job(request: JobCreateRequest):
    """Create a new job requisition."""
    db = SessionLocal()
    try:
        job = JobRequisition(
            title=request.title,
            jd_raw=request.jd_raw,
            required_skills=json.dumps(request.required_skills or []),
            min_exp=request.min_exp,
            location=request.location,
            remote_ok=request.remote_ok,
        )
        db.add(job)
        db.commit()
        db.refresh(job)
        return _job_to_dict(job)
    finally:
        db.close()


@app.get("/api/jobs")
async def list_jobs():
    """List all job requisitions."""
    db = SessionLocal()
    try:
        jobs = db.query(JobRequisition).order_by(JobRequisition.created_at.desc()).all()
        return [_job_to_dict(j) for j in jobs]
    finally:
        db.close()


# ─── Recruitment: Matching endpoints ────────────────────────────────────────

@app.post("/api/jobs/{job_id}/match")
async def run_matching(job_id: int):
    """Run the matching engine for a job against all candidates. Saves results."""
    try:
        results = run_match(job_id)
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))

    db = SessionLocal()
    try:
        job = db.query(JobRequisition).filter(JobRequisition.id == job_id).first()
        return {
            "job": _job_to_dict(job),
            "total_candidates": len(results),
            "results": results,
        }
    finally:
        db.close()


@app.get("/api/jobs/{job_id}/results")
async def get_match_results(job_id: int):
    """Fetch saved match results for a job, ranked by score."""
    db = SessionLocal()
    try:
        job = db.query(JobRequisition).filter(JobRequisition.id == job_id).first()
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")

        matches = (
            db.query(MatchResult)
            .filter(MatchResult.job_id == job_id)
            .order_by(MatchResult.score.desc())
            .all()
        )

        results = []
        for mr in matches:
            c = db.query(Candidate).filter(Candidate.id == mr.candidate_id).first()
            results.append({
                "match_id": mr.id,
                "candidate": _candidate_to_dict(c) if c else None,
                "score": mr.score,
                "match_reasons": json.loads(mr.match_reasons) if mr.match_reasons else [],
                "fit_level": mr.fit_level,
            })

        return {
            "job": _job_to_dict(job),
            "total_candidates": len(results),
            "results": results,
        }
    finally:
        db.close()


# ─── Recruitment: CSV export ────────────────────────────────────────────────

@app.get("/api/export/candidates-csv")
async def export_candidates_csv(job_id: Optional[int] = None):
    """Export match results (or all candidates) as CSV via pandas."""
    import pandas as pd

    db = SessionLocal()
    try:
        if job_id:
            job = db.query(JobRequisition).filter(JobRequisition.id == job_id).first()
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")

            matches = (
                db.query(MatchResult)
                .filter(MatchResult.job_id == job_id)
                .order_by(MatchResult.score.desc())
                .all()
            )
            if not matches:
                raise HTTPException(status_code=404, detail="No match results found. Run matching first.")

            rows = []
            for mr in matches:
                c = db.query(Candidate).filter(Candidate.id == mr.candidate_id).first()
                if not c:
                    continue
                rows.append({
                    "Rank": None,
                    "Name": c.name,
                    "Email": c.email or "",
                    "Phone": c.phone or "",
                    "Location": c.location or "",
                    "Titles": "; ".join(json.loads(c.titles) if c.titles else []),
                    "Skills": "; ".join(json.loads(c.skills) if c.skills else []),
                    "Years Experience": c.years_exp or "",
                    "Match Score": mr.score,
                    "Fit Level": mr.fit_level or "",
                    "Match Reasons": "; ".join(json.loads(mr.match_reasons) if mr.match_reasons else []),
                    "Job Title": job.title,
                })
            # Fill rank after sorting (already sorted by score desc)
            for i, row in enumerate(rows, 1):
                row["Rank"] = i

            filename = f"match_results_job{job_id}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.csv"
        else:
            candidates = db.query(Candidate).order_by(Candidate.created_at.desc()).all()
            if not candidates:
                raise HTTPException(status_code=404, detail="No candidates in database")

            rows = []
            for c in candidates:
                rows.append({
                    "Name": c.name,
                    "Email": c.email or "",
                    "Phone": c.phone or "",
                    "Location": c.location or "",
                    "Titles": "; ".join(json.loads(c.titles) if c.titles else []),
                    "Skills": "; ".join(json.loads(c.skills) if c.skills else []),
                    "Years Experience": c.years_exp or "",
                    "Source Email UID": c.source_email_uid or "",
                    "Created At": c.created_at.isoformat() if c.created_at else "",
                })

            filename = f"candidates_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.csv"

        df = pd.DataFrame(rows)
        output = io.StringIO()
        df.to_csv(output, index=False)

        return Response(
            content=output.getvalue(),
            media_type="text/csv",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    finally:
        db.close()


# ─── Health Check ───────────────────────────────────────────────────────────

def _human_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


@app.get("/api/storage/health")
async def storage_health():
    """Comprehensive storage and persistence health check."""
    from sqlalchemy import func

    db = SessionLocal()
    try:
        # ── 1. Database section ──
        db_size = DB_PATH.stat().st_size if DB_PATH.exists() else 0

        # ── 2. Record counts ──
        candidates_count = db.query(func.count(Candidate.id)).scalar() or 0
        jobs_count = db.query(func.count(JobRequisition.id)).scalar() or 0
        match_count = db.query(func.count(MatchResult.id)).scalar() or 0
        emails_count = db.query(func.count(ScrapedEmail.id)).scalar() or 0
        attachments_count = db.query(func.count(Attachment.id)).scalar() or 0
        notes_count = db.query(func.count(Candidate.id)).filter(
            Candidate.notes.isnot(None), Candidate.notes != "",
        ).scalar() or 0
        tagged_count = db.query(func.count(Candidate.id)).filter(
            Candidate.tags.isnot(None), Candidate.tags != "[]",
        ).scalar() or 0

        # ── 3. Sync timestamps ──
        last_scraped = db.query(func.max(ScrapedEmail.scraped_at)).scalar()
        last_candidate = db.query(func.max(Candidate.created_at)).scalar()
        last_match = db.query(func.max(MatchResult.id)).scalar()
        # For last match, get the actual match result's job run time
        last_match_time = None
        if last_match:
            mr = db.query(MatchResult).filter(MatchResult.id == last_match).first()
            if mr and mr.job:
                last_match_time = mr.job.created_at

        imap_connected = False
        if _imap_connection:
            try:
                _imap_connection.noop()
                imap_connected = True
            except Exception:
                pass

        # ── 4. Attachments on disk ──
        att_files = [f for f in ATTACHMENTS_DIR.iterdir()
                     if f.is_file() and f.name != "_metadata.json"]
        total_att_size = sum(f.stat().st_size for f in att_files)
        by_type = {}
        for f in att_files:
            ext = f.suffix.lower().lstrip(".")
            if ext in ("jpg", "jpeg", "png", "gif", "webp", "bmp", "svg"):
                key = "image"
            elif ext == "pdf":
                key = "pdf"
            elif ext in ("docx", "doc"):
                key = ext
            else:
                key = "other"
            by_type[key] = by_type.get(key, 0) + 1

        # ── 5. Health status ──
        if not DB_PATH.exists():
            health_status = "warning"
        elif candidates_count == 0 and emails_count == 0:
            health_status = "empty"
        elif not SESSION_FILE.exists():
            health_status = "warning"
        else:
            health_status = "healthy"

        return {
            "database": {
                "db_file_path": str(DB_PATH.resolve()),
                "db_size_bytes": db_size,
                "db_size_human": _human_size(db_size),
            },
            "record_counts": {
                "candidates": candidates_count,
                "jobs": jobs_count,
                "match_results": match_count,
                "scraped_emails": emails_count,
                "attachments": attachments_count,
                "notes_count": notes_count,
                "tagged_count": tagged_count,
            },
            "sync": {
                "last_scraped_at": last_scraped.isoformat() if last_scraped else None,
                "last_candidate_added": last_candidate.isoformat() if last_candidate else None,
                "last_match_run": last_match_time.isoformat() if last_match_time else None,
                "session_file_exists": SESSION_FILE.exists(),
                "imap_connected": imap_connected,
            },
            "attachments": {
                "total_files": len(att_files),
                "total_size_bytes": total_att_size,
                "total_size_human": _human_size(total_att_size),
                "by_type": by_type,
            },
            "health_status": health_status,
        }
    finally:
        db.close()


class ClearDataRequest(BaseModel):
    confirm: str


@app.delete("/api/storage/clear-data")
async def clear_data(body: ClearDataRequest):
    """Truncate all data tables. Attachment files on disk and .session.json are preserved."""
    if body.confirm != "CONFIRM":
        raise HTTPException(status_code=400, detail='Body must contain {"confirm": "CONFIRM"}')
    db = SessionLocal()
    try:
        db.query(MatchResult).delete()
        db.query(Candidate).delete()
        db.query(JobRequisition).delete()
        db.query(ScrapedEmail).delete()
        db.query(Attachment).delete()
        db.commit()
        return {"deleted": True, "message": "All records cleared. Files on disk were preserved."}
    except Exception:
        db.rollback()
        raise HTTPException(status_code=500, detail="Failed to clear data")
    finally:
        db.close()


@app.get("/api/storage/backup")
async def backup_database():
    """Download a copy of the SQLite database file."""
    import shutil
    import tempfile

    if not DB_PATH.exists():
        raise HTTPException(status_code=404, detail="Database file not found")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
    tmp.close()
    tmp_path = Path(tmp.name)
    try:
        shutil.copy2(DB_PATH, tmp_path)
        content = tmp_path.read_bytes()
        date_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        return Response(
            content=content,
            media_type="application/octet-stream",
            headers={"Content-Disposition": f'attachment; filename="mailscraper_backup_{date_str}.db"'},
        )
    finally:
        tmp_path.unlink(missing_ok=True)


# ─── Scheduled Scrape ───────────────────────────────────────────────────────

async def run_scheduled_scrape():
    """Execute a scrape using saved SchedulerConfig settings."""
    db = SessionLocal()
    try:
        cfg = db.query(SchedulerConfig).first()
        if not cfg:
            return
        if not cfg.enabled:
            return
    finally:
        db.close()

    # Check we have IMAP credentials
    if not _credentials.get("email") or not _credentials.get("password"):
        logger.warning("Scheduled scrape skipped — no IMAP credentials")
        return

    # Count candidates before scrape
    db = SessionLocal()
    try:
        cand_before = db.query(Candidate).count()
    finally:
        db.close()

    # Run the scrape via IMAP
    try:
        emails = await _imap_op(
            _scrape_impl,
            folder_id=cfg.folder or "INBOX",
            subject_filter=cfg.subject_filter,
            max_results=50,
            include_attachments=True,
        )
        _persist_scraped_emails(emails, folder=cfg.folder or "INBOX")
    except Exception:
        logger.exception("Scheduled scrape IMAP operation failed")
        return

    # Count candidates after scrape
    db = SessionLocal()
    try:
        cand_after = db.query(Candidate).count()
        new_candidates = max(0, cand_after - cand_before)

        now = datetime.now(timezone.utc)
        cfg = db.query(SchedulerConfig).first()
        if cfg:
            cfg.last_run_at = now
            cfg.emails_found_last_run = len(emails)
            cfg.candidates_added_last_run = new_candidates
            cfg.next_run_at = now + timedelta(minutes=cfg.interval_minutes or 30)
            db.commit()

        # Create scrape_complete notification
        try:
            db.add(Notification(
                type="scrape_complete",
                title="Scrape Complete",
                message=f"Found {len(emails)} emails, {new_candidates} new candidates",
            ))
            db.commit()
        except Exception:
            db.rollback()
            logger.exception("Failed to create scrape_complete notification")

        logger.info(
            "Scheduled scrape complete: %d emails, %d new candidates",
            len(emails), new_candidates,
        )
    except Exception:
        db.rollback()
        logger.exception("Failed to update scheduler stats")
    finally:
        db.close()


def _scheduler_config_to_dict(cfg: SchedulerConfig) -> dict:
    def _iso(dt):
        if dt is None:
            return None
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.isoformat()

    return {
        "id": cfg.id,
        "enabled": cfg.enabled,
        "interval_minutes": cfg.interval_minutes,
        "folder": cfg.folder,
        "subject_filter": cfg.subject_filter,
        "last_run_at": _iso(cfg.last_run_at),
        "next_run_at": _iso(cfg.next_run_at),
        "emails_found_last_run": cfg.emails_found_last_run or 0,
        "candidates_added_last_run": cfg.candidates_added_last_run or 0,
    }


@app.get("/api/scheduler/status")
async def scheduler_status():
    """Return current scheduler configuration and runtime status."""
    db = SessionLocal()
    try:
        cfg = db.query(SchedulerConfig).first()
        if not cfg:
            cfg = SchedulerConfig(enabled=False, interval_minutes=30, folder="INBOX")
            db.add(cfg)
            db.commit()
            db.refresh(cfg)

        result = _scheduler_config_to_dict(cfg)

        # Is the scheduler job currently registered?
        job = scheduler.get_job(_SCHEDULER_JOB_ID) if scheduler.running else None
        result["is_running"] = job is not None

        # Seconds until next run
        if cfg.next_run_at:
            next_at = cfg.next_run_at
            if next_at.tzinfo is None:
                next_at = next_at.replace(tzinfo=timezone.utc)
            delta = (next_at - datetime.now(timezone.utc)).total_seconds()
            result["time_until_next_run_seconds"] = max(0, round(delta))
        else:
            result["time_until_next_run_seconds"] = None

        return result
    finally:
        db.close()


class SchedulerConfigRequest(BaseModel):
    enabled: bool = False
    interval_minutes: int = Field(default=30, ge=1, le=1440)
    folder: str = "INBOX"
    subject_filter: Optional[str] = None


@app.post("/api/scheduler/config")
async def update_scheduler_config(body: SchedulerConfigRequest):
    """Update scheduler settings. Starts or stops the scheduled job accordingly."""
    db = SessionLocal()
    try:
        cfg = db.query(SchedulerConfig).first()
        if not cfg:
            cfg = SchedulerConfig()
            db.add(cfg)

        cfg.enabled = body.enabled
        cfg.interval_minutes = body.interval_minutes
        cfg.folder = body.folder
        cfg.subject_filter = body.subject_filter

        if body.enabled:
            cfg.next_run_at = datetime.now(timezone.utc) + timedelta(minutes=body.interval_minutes)

        db.commit()
        db.refresh(cfg)

        # Update scheduler job
        if body.enabled and body.interval_minutes > 0:
            if scheduler.get_job(_SCHEDULER_JOB_ID):
                scheduler.remove_job(_SCHEDULER_JOB_ID)
            scheduler.add_job(
                run_scheduled_scrape,
                trigger=IntervalTrigger(minutes=body.interval_minutes),
                id=_SCHEDULER_JOB_ID,
                replace_existing=True,
            )
            logger.info("Scheduler enabled: every %d min, folder=%s", body.interval_minutes, body.folder)
        else:
            if scheduler.get_job(_SCHEDULER_JOB_ID):
                scheduler.remove_job(_SCHEDULER_JOB_ID)
            cfg.next_run_at = None
            db.commit()
            logger.info("Scheduler disabled")

        return _scheduler_config_to_dict(cfg)
    finally:
        db.close()


@app.post("/api/scheduler/run-now")
async def scheduler_run_now():
    """Trigger an immediate scrape in the background."""
    asyncio.create_task(run_scheduled_scrape())
    return {
        "message": "Scrape started",
        "started_at": datetime.now(timezone.utc).isoformat(),
    }


# ─── Notifications ─────────────────────────────────────────────────────────

def _notification_to_dict(n: Notification) -> dict:
    created = n.created_at
    if created and created.tzinfo is None:
        created = created.replace(tzinfo=timezone.utc)
    return {
        "id": n.id,
        "type": n.type,
        "title": n.title,
        "message": n.message,
        "job_id": n.job_id,
        "candidate_id": n.candidate_id,
        "is_read": n.is_read,
        "created_at": created.isoformat() if created else None,
    }


@app.get("/api/notifications")
async def get_notifications(limit: int = 50, offset: int = 0, unread_only: bool = False):
    """Return notifications ordered by newest first."""
    db = SessionLocal()
    try:
        q = db.query(Notification)
        if unread_only:
            q = q.filter(Notification.is_read == False)
        q = q.order_by(Notification.created_at.desc())
        total = q.count()
        items = q.offset(offset).limit(limit).all()
        return {
            "notifications": [_notification_to_dict(n) for n in items],
            "total": total,
        }
    finally:
        db.close()


@app.get("/api/notifications/count")
async def get_notification_count():
    """Return count of unread notifications."""
    db = SessionLocal()
    try:
        unread = db.query(Notification).filter(Notification.is_read == False).count()
        total = db.query(Notification).count()
        return {"unread": unread, "total": total}
    finally:
        db.close()


@app.patch("/api/notifications/{notification_id}/read")
async def mark_notification_read(notification_id: int):
    """Mark a single notification as read."""
    db = SessionLocal()
    try:
        n = db.query(Notification).filter(Notification.id == notification_id).first()
        if not n:
            raise HTTPException(status_code=404, detail="Notification not found")
        n.is_read = True
        db.commit()
        return _notification_to_dict(n)
    finally:
        db.close()


@app.patch("/api/notifications/read-all")
async def mark_all_notifications_read():
    """Mark all notifications as read."""
    db = SessionLocal()
    try:
        count = db.query(Notification).filter(Notification.is_read == False).update({"is_read": True})
        db.commit()
        return {"marked_read": count}
    finally:
        db.close()


@app.delete("/api/notifications/clear")
async def clear_notifications():
    """Delete all notifications."""
    db = SessionLocal()
    try:
        count = db.query(Notification).delete()
        db.commit()
        return {"deleted": count}
    finally:
        db.close()


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

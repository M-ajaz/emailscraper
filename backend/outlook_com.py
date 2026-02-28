"""
Outlook COM automation — read emails directly from the locally installed
Outlook desktop app via Windows COM (pywin32).

Requires:
    - Microsoft Outlook installed and configured with at least one account
    - pywin32 (pip install pywin32==306)
    - Windows OS
"""

import logging
from datetime import datetime, timezone
from pathlib import Path

import pythoncom
import win32com.client
import pywintypes

pythoncom.CoInitialize()

logger = logging.getLogger(__name__)

# Outlook folder type constants (OlDefaultFolders enumeration)
_OL_FOLDER_INBOX = 6

_MAX_SUBFOLDER_DEPTH = 3


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _com_date_to_iso(com_dt) -> str:
    """Convert a COM datetime (pywintypes.datetime) to an ISO 8601 string."""
    try:
        if hasattr(com_dt, "isoformat"):
            return com_dt.isoformat()
        # pywintypes.datetime is a subclass of datetime on some builds
        dt = datetime(
            com_dt.year, com_dt.month, com_dt.day,
            com_dt.hour, com_dt.minute, com_dt.second,
            tzinfo=timezone.utc,
        )
        return dt.isoformat()
    except Exception:
        return str(com_dt)


def _safe_str(value, default: str = "") -> str:
    """Return str(value) or *default* if the COM property is None / errors."""
    try:
        return str(value) if value is not None else default
    except Exception:
        return default


def _collect_folders(folder, current_path: str = "", depth: int = 0) -> list:
    """Recursively collect folder metadata up to *_MAX_SUBFOLDER_DEPTH* levels."""
    try:
        name = _safe_str(folder.Name)
        path = f"{current_path}/{name}" if current_path else name
        entry = {
            "id": folder.EntryID,
            "name": name,
            "path": path,
            "total_count": folder.Items.Count,
            "unread_count": folder.UnReadItemCount,
            "subfolders": [],
        }
    except Exception as exc:
        logger.debug("Skipping inaccessible folder: %s", exc)
        return []

    if depth < _MAX_SUBFOLDER_DEPTH:
        try:
            for i in range(1, folder.Folders.Count + 1):
                sub = folder.Folders.Item(i)
                entry["subfolders"].extend(
                    _collect_folders(sub, path, depth + 1)
                )
        except Exception:
            pass  # some special folders block enumeration

    return [entry]


# ─── 1. is_outlook_running ───────────────────────────────────────────────────

def is_outlook_running() -> bool:
    """Return True if the Outlook desktop app is currently running and accessible."""
    try:
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


# ─── 2. get_outlook_app ─────────────────────────────────────────────────────

def get_outlook_app():
    """
    Return a COM reference to the Outlook.Application object.

    Tries to attach to a running instance first; falls back to launching
    Outlook if it isn't open yet.

    Raises
    ------
    OSError  if Outlook is not installed or cannot be started.
    """
    # Try connecting to an already-running instance
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        pass

    # Fall back to launching / dispatching
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except pywintypes.com_error as exc:
        raise OSError(
            "Microsoft Outlook is not installed or could not be started. "
            f"COM error: {exc}"
        ) from exc
    except Exception as exc:
        raise OSError(
            f"Failed to connect to Outlook: {exc}"
        ) from exc


# ─── 3. get_folders ─────────────────────────────────────────────────────────

def get_folders() -> list:
    """
    List every mail folder in the default Outlook account.

    Returns a list of dicts, each containing:
        id, name, path, total_count, unread_count, subfolders
    Subfolders are collected recursively up to 3 levels deep.
    """
    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")

        results = []
        # Iterate over top-level folders of the default store
        inbox = namespace.GetDefaultFolder(_OL_FOLDER_INBOX)
        root = inbox.Parent  # root folder of the default store
        for i in range(1, root.Folders.Count + 1):
            folder = root.Folders.Item(i)
            results.extend(_collect_folders(folder))

        return results
    except OSError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Failed to list Outlook folders: {exc}") from exc


# ─── 4. get_messages ────────────────────────────────────────────────────────

def get_messages(folder_id: str, filters: dict | None = None) -> dict:
    """
    Retrieve messages from a folder identified by its EntryID.

    Parameters
    ----------
    folder_id : str
        The Outlook EntryID of the target folder.
    filters : dict, optional
        Supported keys:
            date_from      (str|datetime) – ReceivedTime >= value
            date_to        (str|datetime) – ReceivedTime <= value
            sender         (str)          – SenderEmailAddress contains value
            subject        (str)          – Subject contains value
            keyword        (str)          – Body or Subject contains value
            has_attachments (bool)        – Attachments.Count > 0
            limit          (int)          – max results (default 50)

    Returns
    -------
    dict  {messages: [...], total: int}
    """
    if filters is None:
        filters = {}
    limit = int(filters.get("limit", 50))

    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")
        folder = namespace.GetItemFromID(folder_id)
    except pywintypes.com_error as exc:
        raise ValueError(f"Folder not found for EntryID: {exc}") from exc
    except Exception as exc:
        raise RuntimeError(f"Failed to open folder: {exc}") from exc

    # ── Build a DASL / Restrict filter string ────────────────────────────
    restrict_parts: list[str] = []

    date_from = filters.get("date_from")
    if date_from:
        if isinstance(date_from, str):
            date_from = datetime.fromisoformat(date_from)
        restrict_parts.append(
            f"[ReceivedTime] >= '{date_from.strftime('%m/%d/%Y %H:%M %p')}'"
        )

    date_to = filters.get("date_to")
    if date_to:
        if isinstance(date_to, str):
            date_to = datetime.fromisoformat(date_to)
        restrict_parts.append(
            f"[ReceivedTime] <= '{date_to.strftime('%m/%d/%Y %H:%M %p')}'"
        )

    # ── Apply Restrict (date filters only — text filters done post-fetch)
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", Descending=True)

        if restrict_parts:
            restriction = " AND ".join(restrict_parts)
            items = items.Restrict(restriction)
    except Exception as exc:
        raise RuntimeError(f"Failed to query folder items: {exc}") from exc

    # ── Post-fetch filters & collect results ─────────────────────────────
    sender_filter = (filters.get("sender") or "").lower()
    subject_filter = (filters.get("subject") or "").lower()
    keyword_filter = (filters.get("keyword") or "").lower()
    att_filter = filters.get("has_attachments")

    messages: list[dict] = []
    try:
        count = items.Count
    except Exception:
        count = 0

    for idx in range(1, count + 1):
        if len(messages) >= limit:
            break
        try:
            item = items.Item(idx)

            # Only process mail items (class 43 = olMail)
            if item.Class != 43:
                continue

            sender_addr = _safe_str(item.SenderEmailAddress).lower()
            sender_name = _safe_str(item.SenderName)
            subject = _safe_str(item.Subject) or "(No Subject)"
            body = _safe_str(item.Body)
            att_count = item.Attachments.Count

            # Apply text-based filters
            if sender_filter and sender_filter not in sender_addr:
                continue
            if subject_filter and subject_filter not in subject.lower():
                continue
            if keyword_filter:
                if (keyword_filter not in subject.lower()
                        and keyword_filter not in body.lower()):
                    continue
            if att_filter is not None and (att_count > 0) != att_filter:
                continue

            messages.append({
                "id": item.EntryID,
                "uid": item.EntryID,
                "subject": subject,
                "sender": sender_name,
                "sender_email": _safe_str(item.SenderEmailAddress),
                "date": _com_date_to_iso(item.ReceivedTime),
                "body_text": body,
                "body_html": _safe_str(item.HTMLBody),
                "has_attachments": att_count > 0,
                "attachment_count": att_count,
                "is_read": not item.UnRead,
                "folder_name": _safe_str(item.Parent.Name),
            })
        except Exception as exc:
            logger.debug("Skipping unreadable item %d: %s", idx, exc)
            continue

    return {"messages": messages, "total": len(messages)}


# ─── 5. get_message ─────────────────────────────────────────────────────────

def get_message(message_id: str) -> dict:
    """
    Retrieve a single email by its EntryID.

    Returns the full message dict including body_text and body_html.
    """
    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.GetItemFromID(message_id)
    except pywintypes.com_error as exc:
        raise ValueError(f"Message not found for EntryID: {exc}") from exc
    except Exception as exc:
        raise RuntimeError(f"Failed to retrieve message: {exc}") from exc

    try:
        att_count = item.Attachments.Count
        attachments = []
        for i in range(1, att_count + 1):
            att = item.Attachments.Item(i)
            attachments.append({
                "index": i - 1,
                "name": _safe_str(att.FileName),
                "size": att.Size,
            })

        return {
            "id": item.EntryID,
            "uid": item.EntryID,
            "subject": _safe_str(item.Subject) or "(No Subject)",
            "sender": _safe_str(item.SenderName),
            "sender_email": _safe_str(item.SenderEmailAddress),
            "date": _com_date_to_iso(item.ReceivedTime),
            "body_text": _safe_str(item.Body),
            "body_html": _safe_str(item.HTMLBody),
            "has_attachments": att_count > 0,
            "attachment_count": att_count,
            "attachments": attachments,
            "is_read": not item.UnRead,
            "folder_name": _safe_str(item.Parent.Name),
        }
    except Exception as exc:
        raise RuntimeError(f"Failed to parse message: {exc}") from exc


# ─── 6. download_attachment ─────────────────────────────────────────────────

def download_attachment(message_id: str, attachment_index: int, save_path: str) -> str:
    """
    Save an attachment from a message to disk.

    Parameters
    ----------
    message_id : str
        EntryID of the parent message.
    attachment_index : int
        Zero-based index of the attachment.
    save_path : str
        Destination file path.

    Returns
    -------
    str  The resolved path where the file was saved.
    """
    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.GetItemFromID(message_id)
    except pywintypes.com_error as exc:
        raise ValueError(f"Message not found for EntryID: {exc}") from exc
    except Exception as exc:
        raise RuntimeError(f"Failed to retrieve message: {exc}") from exc

    att_count = item.Attachments.Count
    # COM attachments are 1-indexed
    com_index = attachment_index + 1
    if com_index < 1 or com_index > att_count:
        raise IndexError(
            f"Attachment index {attachment_index} out of range "
            f"(message has {att_count} attachment(s))"
        )

    try:
        dest = Path(save_path).resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        att = item.Attachments.Item(com_index)
        att.SaveAsFile(str(dest))
        return str(dest)
    except Exception as exc:
        raise RuntimeError(f"Failed to save attachment: {exc}") from exc


# ─── 7. get_account_info ────────────────────────────────────────────────────

def get_account_info() -> dict:
    """
    Return basic info about the default Outlook account.

    Returns
    -------
    dict  {name, email, exchange_server}
    """
    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")

        name = _safe_str(namespace.CurrentUser.Name)
        address = _safe_str(namespace.CurrentUser.Address)

        # Try to determine the Exchange server (only for Exchange accounts)
        exchange_server = ""
        try:
            account = namespace.Accounts.Item(1)
            exchange_server = _safe_str(account.ExchangeMailboxServerName)
        except Exception:
            pass

        return {
            "name": name,
            "email": address,
            "exchange_server": exchange_server,
        }
    except OSError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Failed to get account info: {exc}") from exc


# ─── 8. mark_as_read ────────────────────────────────────────────────────────

def mark_as_read(message_id: str) -> None:
    """Mark a message as read by its EntryID."""
    try:
        outlook = get_outlook_app()
        namespace = outlook.GetNamespace("MAPI")
        item = namespace.GetItemFromID(message_id)
        item.UnRead = False
        item.Save()
    except pywintypes.com_error as exc:
        raise ValueError(f"Message not found for EntryID: {exc}") from exc
    except Exception as exc:
        raise RuntimeError(f"Failed to mark message as read: {exc}") from exc

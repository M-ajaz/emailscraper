"""
Microsoft Graph API client — alternative to IMAP for accessing Outlook mail.
Handles OAuth2 token management, folder listing, message retrieval, and attachments.
"""

import requests
import json
import os
import base64
from datetime import datetime, timedelta
from pathlib import Path

SESSION_FILE = Path(__file__).resolve().parent / ".session.json"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"


def get_tokens() -> dict | None:
    """Read .session.json and return the microsoft_tokens dict if it exists."""
    if not SESSION_FILE.exists():
        return None
    try:
        data = json.loads(SESSION_FILE.read_text())
        return data.get("microsoft_tokens")
    except (json.JSONDecodeError, OSError):
        return None


def save_tokens(tokens: dict):
    """Read existing .session.json, add/update microsoft_tokens key, save back."""
    data = {}
    if SESSION_FILE.exists():
        try:
            data = json.loads(SESSION_FILE.read_text())
        except (json.JSONDecodeError, OSError):
            data = {}
    data["microsoft_tokens"] = tokens
    SESSION_FILE.write_text(json.dumps(data))


def refresh_access_token() -> str | None:
    """
    Check if access_token is expired; if so, use refresh_token to get a new one.
    Returns a valid access_token or None.
    """
    tokens = get_tokens()
    if not tokens:
        return None

    access_token = tokens.get("access_token")
    expires_at_str = tokens.get("access_token_expires_at")

    # Check if token is still valid
    if access_token and expires_at_str:
        try:
            expires_at = datetime.fromisoformat(expires_at_str)
            if datetime.utcnow() < expires_at - timedelta(minutes=2):
                return access_token
        except (ValueError, TypeError):
            pass

    # Token expired or missing — refresh it
    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        return None

    client_id = os.getenv("MICROSOFT_CLIENT_ID", "")
    if not client_id:
        return None

    try:
        resp = requests.post(TOKEN_URL, data={
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite offline_access User.Read",
        }, timeout=15)
        resp.raise_for_status()
        result = resp.json()

        new_access_token = result["access_token"]
        new_refresh_token = result.get("refresh_token", refresh_token)
        expires_in = result.get("expires_in", 3600)
        new_expires_at = (datetime.utcnow() + timedelta(seconds=expires_in)).isoformat()

        tokens["access_token"] = new_access_token
        tokens["refresh_token"] = new_refresh_token
        tokens["access_token_expires_at"] = new_expires_at
        save_tokens(tokens)

        return new_access_token
    except Exception:
        return None


def get_valid_token() -> str | None:
    """Return a valid access token, refreshing if necessary."""
    return refresh_access_token()


def graph_get(endpoint: str) -> dict:
    """
    GET request to Microsoft Graph API.
    Raises exception on 401 or other errors.
    """
    token = get_valid_token()
    if not token:
        raise Exception("No valid Microsoft Graph access token")

    url = f"{GRAPH_BASE}/{endpoint}"
    resp = requests.get(url, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }, timeout=30)

    if resp.status_code == 401:
        raise Exception("Microsoft Graph token expired or invalid (401)")
    if resp.status_code >= 400:
        detail = resp.text[:500]
        raise Exception(f"Microsoft Graph API error {resp.status_code}: {detail}")

    return resp.json()


def list_folders() -> list:
    """List mail folders via Graph API."""
    data = graph_get("me/mailFolders?$top=50")
    folders = []
    for f in data.get("value", []):
        folders.append({
            "id": f["id"],
            "name": f.get("displayName", ""),
            "total_count": f.get("totalItemCount", 0),
            "unread_count": f.get("unreadItemCount", 0),
        })
    return folders


def list_messages(folder_id: str, filters: dict = None) -> dict:
    """
    List messages in a folder with optional OData filters.

    filters can include:
      - top (int): max results, default 50
      - from_date (str): YYYY-MM-DD
      - to_date (str): YYYY-MM-DD
      - sender_filter (str): sender email/name
      - search (str): keyword search
      - subject_filter (str): subject keyword
    """
    if filters is None:
        filters = {}

    top = filters.get("top", 50)
    params = [f"$top={top}", "$orderby=receivedDateTime desc"]

    # Build $filter clauses
    filter_parts = []
    from_date = filters.get("from_date")
    if from_date:
        filter_parts.append(f"receivedDateTime ge {from_date}T00:00:00Z")
    to_date = filters.get("to_date")
    if to_date:
        # Add one day so "to_date" is inclusive
        try:
            to_dt = datetime.strptime(to_date, "%Y-%m-%d") + timedelta(days=1)
            filter_parts.append(f"receivedDateTime lt {to_dt.strftime('%Y-%m-%d')}T00:00:00Z")
        except ValueError:
            pass
    sender_filter = filters.get("sender_filter")
    if sender_filter:
        safe = sender_filter.replace("'", "''")
        filter_parts.append(f"from/emailAddress/address eq '{safe}'")

    if filter_parts:
        params.append("$filter=" + " and ".join(filter_parts))

    search = filters.get("search") or filters.get("subject_filter")
    if search:
        safe = search.replace("'", "''").replace('"', '\\"')
        params.append(f'$search="{safe}"')

    query = "&".join(params)
    endpoint = f"me/mailFolders/{folder_id}/messages?{query}"

    data = graph_get(endpoint)
    messages = []
    for m in data.get("value", []):
        sender_info = m.get("from", {}).get("emailAddress", {})
        messages.append({
            "id": m["id"],
            "subject": m.get("subject", ""),
            "sender": sender_info.get("name", ""),
            "sender_email": sender_info.get("address", ""),
            "date": m.get("receivedDateTime", ""),
            "body_preview": m.get("bodyPreview", ""),
            "has_attachments": m.get("hasAttachments", False),
            "is_read": m.get("isRead", False),
        })

    total = data.get("@odata.count", len(messages))
    return {"messages": messages, "total": total}


def get_message(message_id: str) -> dict:
    """Fetch a full message with attachments expanded."""
    data = graph_get(f"me/messages/{message_id}?$expand=attachments")

    body = data.get("body", {})
    sender_info = data.get("from", {}).get("emailAddress", {})

    attachments = []
    for att in data.get("attachments", []):
        attachments.append({
            "id": att.get("id", ""),
            "name": att.get("name", ""),
            "content_type": att.get("contentType", ""),
            "size": att.get("size", 0),
            "is_inline": att.get("isInline", False),
            "content_bytes": att.get("contentBytes"),
        })

    return {
        "id": data["id"],
        "subject": data.get("subject", ""),
        "sender": sender_info.get("name", ""),
        "sender_email": sender_info.get("address", ""),
        "date": data.get("receivedDateTime", ""),
        "body_html": body.get("content", "") if body.get("contentType") == "html" else "",
        "body_text": body.get("content", "") if body.get("contentType") == "text" else "",
        "has_attachments": data.get("hasAttachments", False),
        "is_read": data.get("isRead", False),
        "importance": data.get("importance", "normal"),
        "internet_message_id": data.get("internetMessageId", ""),
        "conversation_id": data.get("conversationId", ""),
        "categories": data.get("categories", []),
        "to": [
            {"name": r.get("emailAddress", {}).get("name", ""),
             "email": r.get("emailAddress", {}).get("address", "")}
            for r in data.get("toRecipients", [])
        ],
        "cc": [
            {"name": r.get("emailAddress", {}).get("name", ""),
             "email": r.get("emailAddress", {}).get("address", "")}
            for r in data.get("ccRecipients", [])
        ],
        "attachments": attachments,
    }


def download_attachment(message_id: str, attachment_id: str) -> bytes:
    """Download a specific attachment, returning the raw bytes."""
    data = graph_get(f"me/messages/{message_id}/attachments/{attachment_id}")
    content_b64 = data.get("contentBytes", "")
    return base64.b64decode(content_b64)

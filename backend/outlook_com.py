"""
Outlook COM automation — subprocess wrapper.

All COM work is delegated to outlook_helper.py running in its own process,
avoiding all COM apartment-threading issues with FastAPI's async event loop.
"""

import subprocess
import json
import sys
from pathlib import Path

# In PyInstaller bundles, data files land in sys._MEIPASS; otherwise use script dir.
_BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
HELPER = _BASE / "outlook_helper.py"
PYTHON = sys.executable  # Use same Python that's running FastAPI


def _run(cmd: str, args: dict = None) -> dict:
    """Run outlook_helper.py as subprocess and return parsed JSON."""
    if args is None:
        args = {}
    try:
        result = subprocess.run(
            [PYTHON, str(HELPER), cmd, json.dumps(args)],
            capture_output=True,
            text=True,
            timeout=15,
        )
        if result.stdout.strip():
            data = json.loads(result.stdout.strip())
            if "error" in data:
                raise Exception(data["error"])
            return data
        if result.stderr:
            raise Exception(result.stderr[:500])
        raise Exception("No output from Outlook helper")
    except subprocess.TimeoutExpired:
        raise Exception("Outlook connection timed out — make sure Outlook is open")
    except json.JSONDecodeError as e:
        raise Exception(f"Invalid response from Outlook helper: {e}")


def is_outlook_running() -> bool:
    try:
        _run("get_account_info")
        return True
    except Exception:
        return False


def get_account_info() -> dict:
    return _run("get_account_info")


def get_folders() -> list:
    return _run("get_folders")


def get_messages(folder_id: str, filters: dict = None) -> dict:
    if filters is None:
        filters = {}
    return _run("get_messages", {
        "folder_id": folder_id,
        "limit": filters.get("limit", 50),
        "subject_filter": filters.get("subject", ""),
        "sender_filter": filters.get("sender", ""),
        "keyword_filter": filters.get("keyword", ""),
    })


def get_message(message_id: str) -> dict:
    """Retrieve a single email by its EntryID with attachment metadata."""
    return _run("get_message", {"message_id": message_id})


def download_attachment(message_id: str, attachment_index: int, save_path: str) -> str:
    result = _run("download_attachment", {
        "message_id": message_id,
        "attachment_index": attachment_index,
        "save_path": save_path,
    })
    return result.get("saved", save_path)

"""
Standalone COM helper — spawned as a subprocess by outlook_com.py.

Runs Outlook COM operations in its own process with proper
CoInitialize/CoUninitialize, prints JSON to stdout.
"""

import sys
import json
import pythoncom
import win32com.client


def _get_outlook():
    """Connect to already-running Outlook instance"""
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        pass
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise Exception(f"Cannot connect to Outlook: {e}. Please make sure Outlook is open.")


def get_account_info():
    pythoncom.CoInitialize()
    try:
        outlook = _get_outlook()
        ns = outlook.GetNamespace("MAPI")
        return {"name": ns.CurrentUser.Name, "email": ns.CurrentUser.Address}
    finally:
        pythoncom.CoUninitialize()


def get_folders():
    pythoncom.CoInitialize()
    try:
        outlook = _get_outlook()
        ns = outlook.GetNamespace("MAPI")
        result = []
        for store in ns.Stores:
            try:
                root = store.GetRootFolder()
                result.extend(_get_subfolders(root, ""))
            except Exception:
                pass
        return result
    finally:
        pythoncom.CoUninitialize()


def _get_subfolders(folder, parent_path, depth=0):
    if depth > 3:
        return []
    path = f"{parent_path}/{folder.Name}" if parent_path else folder.Name
    try:
        count = folder.Items.Count
    except Exception:
        count = 0
    items = [{"id": folder.EntryID, "name": folder.Name, "path": path, "total_count": count, "unread_count": 0}]
    if depth < 3:
        try:
            for sub in folder.Folders:
                items.extend(_get_subfolders(sub, path, depth + 1))
        except Exception:
            pass
    return items


def get_messages(folder_id, limit=50, subject_filter="", sender_filter="", keyword_filter=""):
    pythoncom.CoInitialize()
    try:
        outlook = _get_outlook()
        ns = outlook.GetNamespace("MAPI")
        folder = ns.GetItemFromID(folder_id)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        results = []
        for item in items:
            if len(results) >= limit:
                break
            try:
                if not hasattr(item, "Subject"):
                    continue
                subj = item.Subject or "(No Subject)"
                sender = item.SenderName or "Unknown"
                sender_email = item.SenderEmailAddress or ""
                body = item.Body or ""
                if subject_filter and subject_filter.lower() not in subj.lower():
                    continue
                if sender_filter and sender_filter.lower() not in sender_email.lower():
                    continue
                if keyword_filter and keyword_filter.lower() not in body.lower() and keyword_filter.lower() not in subj.lower():
                    continue
                results.append({
                    "id": item.EntryID,
                    "uid": item.EntryID,
                    "subject": subj,
                    "sender": sender,
                    "sender_email": sender_email,
                    "date": item.ReceivedTime.isoformat(),
                    "body_text": body[:5000],
                    "body_html": item.HTMLBody or "",
                    "has_attachments": item.Attachments.Count > 0,
                    "attachment_count": item.Attachments.Count,
                    "is_read": not item.UnRead,
                    "folder_name": folder.Name,
                    "attachment_names": [item.Attachments[i + 1].FileName for i in range(item.Attachments.Count)]
                })
            except Exception:
                continue
        return {"messages": results, "total": len(results)}
    finally:
        pythoncom.CoUninitialize()


def get_message(message_id):
    pythoncom.CoInitialize()
    try:
        outlook = _get_outlook()
        ns = outlook.GetNamespace("MAPI")
        item = ns.GetItemFromID(message_id)
        att_count = item.Attachments.Count
        attachments = []
        for i in range(1, att_count + 1):
            att = item.Attachments.Item(i)
            attachments.append({
                "index": i - 1,
                "name": att.FileName or f"attachment_{i-1}",
                "size": att.Size,
            })
        return {
            "id": item.EntryID,
            "subject": item.Subject or "(No Subject)",
            "sender": item.SenderName or "",
            "sender_email": item.SenderEmailAddress or "",
            "date": item.ReceivedTime.isoformat(),
            "body_text": (item.Body or "")[:5000],
            "body_html": item.HTMLBody or "",
            "has_attachments": att_count > 0,
            "attachment_count": att_count,
            "attachments": attachments,
        }
    finally:
        pythoncom.CoUninitialize()


def download_attachment(message_id, attachment_index, save_path):
    pythoncom.CoInitialize()
    try:
        outlook = _get_outlook()
        ns = outlook.GetNamespace("MAPI")
        item = ns.GetItemFromID(message_id)
        att = item.Attachments[attachment_index + 1]
        att.SaveAsFile(save_path)
        return {"saved": save_path}
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    try:
        if cmd == "get_account_info":
            print(json.dumps(get_account_info()))
        elif cmd == "get_folders":
            print(json.dumps(get_folders()))
        elif cmd == "get_messages":
            print(json.dumps(get_messages(**args)))
        elif cmd == "get_message":
            print(json.dumps(get_message(**args)))
        elif cmd == "download_attachment":
            print(json.dumps(download_attachment(**args)))
        else:
            print(json.dumps({"error": f"Unknown command: {cmd}"}))
    except Exception as e:
        print(json.dumps({"error": str(e)}))

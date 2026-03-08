"""Outlook COM interface wrapper using win32com.client."""

import os
import tempfile
from datetime import datetime, timedelta

import pythoncom
import win32com.client

# OlDefaultFolders constants
OL_FOLDER_INBOX = 6
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10
OL_FOLDER_JUNK = 23

# OlItemType constants
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1

# OlFlagStatus
OL_NO_FLAG = 0
OL_FLAGGED = 2


def get_outlook():
    """Get or create Outlook Application COM object."""
    pythoncom.CoInitialize()
    return win32com.client.Dispatch("Outlook.Application")


def get_namespace(outlook=None):
    """Get MAPI namespace."""
    if outlook is None:
        outlook = get_outlook()
    ns = outlook.GetNamespace("MAPI")
    return ns


def _folder_to_path(folder, prefix=""):
    """Build folder path string."""
    return f"{prefix}/{folder.Name}" if prefix else folder.Name


def _collect_folders(folder, prefix=""):
    """Recursively collect folder info."""
    path = _folder_to_path(folder, prefix)
    try:
        total = folder.Items.Count
    except Exception:
        total = -1
    try:
        unread = folder.UnReadItemCount
    except Exception:
        unread = -1
    result = [{
        "name": folder.Name,
        "path": path,
        "totalCount": total,
        "unreadCount": unread,
    }]
    for i in range(1, folder.Folders.Count + 1):
        try:
            sub = folder.Folders.Item(i)
            result.extend(_collect_folders(sub, path))
        except Exception:
            pass
    return result


def _extract_sender(item):
    """Extract sender string from a MailItem COM object."""
    try:
        sender = item.SenderEmailAddress or ""
        sender_name = item.SenderName or ""
        if sender_name and sender_name != sender:
            return f"{sender_name} <{sender}>"
        return sender
    except Exception:
        return ""


def _mail_item_to_summary(item, folder_path=""):
    """Convert a MailItem COM object to a summary dict."""
    try:
        return {
            "messageId": item.EntryID,
            "folderPath": folder_path,
            "subject": item.Subject or "(no subject)",
            "sender": _extract_sender(item),
            "date": str(item.ReceivedTime),
            "read": not item.UnRead,
        }
    except Exception:
        return None


def _mail_item_to_full(item, folder_path="", save_attachments=False):
    """Convert a MailItem COM object to a full detail dict."""
    to_addrs = ""
    cc_addrs = ""
    try:
        to_addrs = item.To or ""
        cc_addrs = item.CC or ""
    except Exception:
        pass

    body = ""
    try:
        body = item.Body or ""
    except Exception:
        pass
    html_body = ""
    try:
        html_body = item.HTMLBody or ""
    except Exception:
        pass

    attachments = []
    if item.Attachments.Count > 0:
        for i in range(1, item.Attachments.Count + 1):
            att = item.Attachments.Item(i)
            att_info = {"filename": att.FileName, "size": att.Size}
            if save_attachments:
                tmp_dir = tempfile.mkdtemp(prefix="outlook_mcp_")
                save_path = os.path.join(tmp_dir, att.FileName)
                att.SaveAsFile(save_path)
                att_info["savedPath"] = save_path
            attachments.append(att_info)

    return {
        "messageId": item.EntryID,
        "folderPath": folder_path,
        "subject": item.Subject or "(no subject)",
        "sender": _extract_sender(item),
        "to": to_addrs,
        "cc": cc_addrs,
        "date": str(item.ReceivedTime),
        "body": body,
        "htmlBody": html_body,
        "attachments": attachments,
        "read": not item.UnRead,
    }


def _resolve_folder(ns, folder_path: str):
    """Resolve a folder by path string like 'AccountName/Inbox/Subfolder'.

    If folder_path is empty or None, returns the default Inbox.
    """
    if not folder_path:
        return ns.GetDefaultFolder(OL_FOLDER_INBOX)

    parts = folder_path.strip("/").split("/")
    # Start from top-level folders (accounts)
    folder = None
    for i in range(1, ns.Folders.Count + 1):
        if ns.Folders.Item(i).Name == parts[0]:
            folder = ns.Folders.Item(i)
            break

    if folder is None:
        raise ValueError(f"Top-level folder '{parts[0]}' not found")

    for part in parts[1:]:
        found = False
        for i in range(1, folder.Folders.Count + 1):
            if folder.Folders.Item(i).Name == part:
                folder = folder.Folders.Item(i)
                found = True
                break
        if not found:
            raise ValueError(f"Subfolder '{part}' not found in '{folder.Name}'")

    return folder


def _get_item_by_entry_id(ns, entry_id: str, folder_path: str):
    """Get a mail item by EntryID, searching in the specified folder."""
    # Prefer GetItemFromID (O(1) lookup by EntryID)
    try:
        return ns.GetItemFromID(entry_id)
    except Exception:
        pass
    # Fallback: iterate folder items
    folder = _resolve_folder(ns, folder_path)
    items = folder.Items
    for i in range(1, items.Count + 1):
        item = items.Item(i)
        if item.EntryID == entry_id:
            return item
    raise ValueError(f"Message with ID '{entry_id}' not found")


# ─── Public API ───────────────────────────────────────────────

def list_accounts():
    """List all configured email accounts."""
    ns = get_namespace()
    accounts = []
    for i in range(1, ns.Accounts.Count + 1):
        acc = ns.Accounts.Item(i)
        accounts.append({
            "name": acc.DisplayName,
            "email": acc.SmtpAddress,
            "accountType": str(acc.AccountType),
        })
    return accounts


def list_folders(account_id: str | None = None):
    """List all mail folders with message counts."""
    ns = get_namespace()
    results = []
    for i in range(1, ns.Folders.Count + 1):
        top = ns.Folders.Item(i)
        if account_id and top.Name != account_id:
            continue
        results.extend(_collect_folders(top))
    return results


def create_folder(parent_folder_path: str, name: str):
    """Create a new subfolder."""
    ns = get_namespace()
    parent = _resolve_folder(ns, parent_folder_path)
    new_folder = parent.Folders.Add(name)
    return {"name": new_folder.Name, "path": f"{parent_folder_path}/{name}"}


def search_messages(
    query: str,
    start_date: str | None = None,
    end_date: str | None = None,
    max_results: int = 50,
    sort_order: str = "desc",
):
    """Search messages by keyword, with optional date range."""
    max_results = min(max_results, 200)
    ns = get_namespace()

    # Build DASL filter (all clauses must use DASL syntax — cannot mix with Jet)
    q = query.replace("'", "''")
    keyword_filter = (
        "(\"urn:schemas:httpmail:subject\" LIKE '%{q}%'"
        " OR \"urn:schemas:httpmail:fromemail\" LIKE '%{q}%'"
        " OR \"urn:schemas:httpmail:displayto\" LIKE '%{q}%')"
    ).format(q=q)

    date_filters = []
    if start_date:
        date_filters.append(f"\"urn:schemas:httpmail:datereceived\" >= '{start_date}'")
    if end_date:
        date_filters.append(f"\"urn:schemas:httpmail:datereceived\" <= '{end_date}'")

    all_clauses = [keyword_filter] + date_filters
    restrict_str = "@SQL=" + " AND ".join(all_clauses)

    results = []

    def _search_folder(folder, folder_path):
        if len(results) >= max_results:
            return
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", sort_order == "desc")
            restricted = items.Restrict(restrict_str)

            for i in range(1, restricted.Count + 1):
                if len(results) >= max_results:
                    return
                item = restricted.Item(i)
                summary = _mail_item_to_summary(item, folder_path)
                if summary:
                    results.append(summary)
        except Exception:
            pass

        # Recurse into subfolders
        for i in range(1, folder.Folders.Count + 1):
            sub = folder.Folders.Item(i)
            _search_folder(sub, f"{folder_path}/{sub.Name}")

    for i in range(1, ns.Folders.Count + 1):
        top = ns.Folders.Item(i)
        _search_folder(top, top.Name)

    return results


def get_recent_messages(
    folder_path: str | None = None,
    days_back: int = 7,
    max_results: int = 50,
    unread_only: bool = False,
):
    """Get recent messages from a folder."""
    max_results = min(max_results, 200)
    ns = get_namespace()

    cutoff = datetime.now() - timedelta(days=days_back)
    cutoff_str = cutoff.strftime("%m/%d/%Y %H:%M %p")
    restrict_str = f"[ReceivedTime] >= '{cutoff_str}'"
    if unread_only:
        restrict_str += " AND [UnRead] = True"

    remaining = max_results

    def _collect_from_folder(folder, fp, limit):
        results = []
        if limit <= 0:
            return results
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            restricted = items.Restrict(restrict_str)
            for i in range(1, min(restricted.Count + 1, limit + 1)):
                summary = _mail_item_to_summary(restricted.Item(i), fp)
                if summary:
                    results.append(summary)
        except Exception:
            pass
        return results

    if folder_path:
        folder = _resolve_folder(ns, folder_path)
        return _collect_from_folder(folder, folder_path, max_results)

    # Default: collect from all Inboxes, respecting total max_results
    results = []
    for i in range(1, ns.Folders.Count + 1):
        if remaining <= 0:
            break
        top = ns.Folders.Item(i)
        inbox = None
        try:
            inbox = top.Folders.Item("Inbox")
        except Exception:
            try:
                inbox = top.Folders.Item("收件箱")
            except Exception:
                pass
        if inbox:
            fp = f"{top.Name}/{inbox.Name}"
            batch = _collect_from_folder(inbox, fp, remaining)
            results.extend(batch)
            remaining -= len(batch)

    # Sort by date descending (items from different accounts may interleave)
    results.sort(key=lambda x: x["date"], reverse=True)
    return results


def get_message(message_id: str, folder_path: str, save_attachments: bool = False):
    """Get full message content by EntryID."""
    ns = get_namespace()
    item = _get_item_by_entry_id(ns, message_id, folder_path)
    return _mail_item_to_full(item, folder_path, save_attachments)


def send_mail(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    bcc: str | None = None,
    is_html: bool = False,
    from_account: str | None = None,
    attachments: list[str] | None = None,
):
    """Create and display a new mail for user confirmation."""
    outlook = get_outlook()
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject

    if is_html:
        mail.HTMLBody = body
    else:
        mail.Body = body

    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc

    if from_account:
        ns = get_namespace(outlook)
        for i in range(1, ns.Accounts.Count + 1):
            acc = ns.Accounts.Item(i)
            if acc.SmtpAddress == from_account or acc.DisplayName == from_account:
                mail.SendUsingAccount = acc
                break

    if attachments:
        for path in attachments:
            if os.path.isfile(path):
                mail.Attachments.Add(path)

    # Display for user confirmation instead of auto-sending
    mail.Display()
    return {"status": "displayed", "message": "Mail compose window opened for user confirmation."}


def reply_to_message(
    message_id: str,
    folder_path: str,
    body: str,
    reply_all: bool = False,
    is_html: bool = False,
    to: str | None = None,
    cc: str | None = None,
    bcc: str | None = None,
    from_account: str | None = None,
    attachments: list[str] | None = None,
):
    """Reply to a message."""
    outlook = get_outlook()
    ns = get_namespace(outlook)
    item = _get_item_by_entry_id(ns, message_id, folder_path)

    reply = item.ReplyAll() if reply_all else item.Reply()

    if is_html:
        reply.HTMLBody = body + reply.HTMLBody
    else:
        reply.Body = body + "\n" + reply.Body

    if to:
        reply.To = to
    if cc:
        reply.CC = cc
    if bcc:
        reply.BCC = bcc

    if from_account:
        for i in range(1, ns.Accounts.Count + 1):
            acc = ns.Accounts.Item(i)
            if acc.SmtpAddress == from_account or acc.DisplayName == from_account:
                reply.SendUsingAccount = acc
                break

    if attachments:
        for path in attachments:
            if os.path.isfile(path):
                reply.Attachments.Add(path)

    reply.Display()
    return {"status": "displayed", "message": "Reply compose window opened for user confirmation."}


def forward_message(
    message_id: str,
    folder_path: str,
    to: str,
    body: str | None = None,
    is_html: bool = False,
    cc: str | None = None,
    bcc: str | None = None,
    from_account: str | None = None,
    attachments: list[str] | None = None,
):
    """Forward a message."""
    outlook = get_outlook()
    ns = get_namespace(outlook)
    item = _get_item_by_entry_id(ns, message_id, folder_path)

    fwd = item.Forward()
    fwd.To = to

    if body:
        if is_html:
            fwd.HTMLBody = body + fwd.HTMLBody
        else:
            fwd.Body = body + "\n" + fwd.Body

    if cc:
        fwd.CC = cc
    if bcc:
        fwd.BCC = bcc

    if from_account:
        for i in range(1, ns.Accounts.Count + 1):
            acc = ns.Accounts.Item(i)
            if acc.SmtpAddress == from_account or acc.DisplayName == from_account:
                fwd.SendUsingAccount = acc
                break

    if attachments:
        for path in attachments:
            if os.path.isfile(path):
                fwd.Attachments.Add(path)

    fwd.Display()
    return {"status": "displayed", "message": "Forward compose window opened for user confirmation."}


def update_message(
    message_id: str,
    folder_path: str,
    read: bool | None = None,
    flagged: bool | None = None,
    move_to: str | None = None,
    trash: bool = False,
):
    """Update message status: read/unread, flagged, move, or trash."""
    if move_to and trash:
        raise ValueError("Cannot use both 'moveTo' and 'trash' at the same time")

    ns = get_namespace()
    item = _get_item_by_entry_id(ns, message_id, folder_path)

    if read is not None:
        item.UnRead = not read

    if flagged is not None:
        item.FlagStatus = OL_FLAGGED if flagged else OL_NO_FLAG

    if read is not None or flagged is not None:
        item.Save()

    if trash:
        item.Delete()
        return {"status": "ok", "action": "trashed"}

    if move_to:
        target = _resolve_folder(ns, move_to)
        item.Move(target)
        return {"status": "ok", "action": f"moved to {move_to}"}

    return {"status": "ok", "action": "updated"}


def delete_messages(message_ids: list[str], folder_path: str):
    """Batch delete messages."""
    ns = get_namespace()
    deleted = 0
    errors = []
    for mid in message_ids:
        try:
            item = _get_item_by_entry_id(ns, mid, folder_path)
            item.Delete()
            deleted += 1
        except Exception as e:
            errors.append({"messageId": mid, "error": str(e)})
    return {"deleted": deleted, "errors": errors}


def search_contacts(query: str):
    """Search contacts by name or email."""
    ns = get_namespace()
    contacts_folder = ns.GetDefaultFolder(OL_FOLDER_CONTACTS)
    items = contacts_folder.Items

    q = query.replace("'", "''")
    restrict_str = (
        f"@SQL=(\"urn:schemas:contacts:cn\" LIKE '%{q}%'"
        f" OR \"urn:schemas:contacts:email1\" LIKE '%{q}%')"
    )

    results = []
    try:
        restricted = items.Restrict(restrict_str)
        for i in range(1, restricted.Count + 1):
            contact = restricted.Item(i)
            results.append({
                "name": contact.FullName or "",
                "email": contact.Email1Address or "",
                "phone": contact.BusinessTelephoneNumber or contact.MobileTelephoneNumber or "",
                "company": contact.CompanyName or "",
            })
    except Exception:
        # Fallback: simple iteration
        for i in range(1, min(items.Count + 1, 500)):
            try:
                contact = items.Item(i)
                name = contact.FullName or ""
                email = contact.Email1Address or ""
                if query.lower() in name.lower() or query.lower() in email.lower():
                    results.append({
                        "name": name,
                        "email": email,
                        "phone": contact.BusinessTelephoneNumber or contact.MobileTelephoneNumber or "",
                        "company": contact.CompanyName or "",
                    })
            except Exception:
                pass

    return results


def list_calendars():
    """List all calendars."""
    ns = get_namespace()
    calendars = []

    def _find_calendars(folder, prefix=""):
        path = _folder_to_path(folder, prefix)
        # Folder type 9 = olFolderCalendar
        if folder.DefaultItemType == OL_APPOINTMENT_ITEM:
            # Check if calendar is writable (ReadOnly property may not exist on all folders)
            writable = True
            try:
                # MAPI property PR_ACCESS (0x0FF40003), bit 0x2 = MAPI_ACCESS_MODIFY
                access = folder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0FF40003")
                writable = bool(access & 0x2)
            except Exception:
                pass
            calendars.append({
                "name": folder.Name,
                "path": path,
                "itemCount": folder.Items.Count,
                "writable": writable,
            })
        for i in range(1, folder.Folders.Count + 1):
            try:
                _find_calendars(folder.Folders.Item(i), path)
            except Exception:
                pass

    for i in range(1, ns.Folders.Count + 1):
        _find_calendars(ns.Folders.Item(i))

    return calendars


def create_event(
    title: str,
    start_date: str,
    end_date: str | None = None,
    location: str | None = None,
    description: str | None = None,
    calendar_id: str | None = None,
    all_day: bool = False,
):
    """Create a calendar event and display for user confirmation."""
    outlook = get_outlook()
    ns = get_namespace(outlook)
    appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)

    # If a specific calendar is requested, move the item there after setting properties
    target_calendar = None
    if calendar_id:
        target_calendar = _resolve_folder(ns, calendar_id)

    appt.Subject = title
    # Outlook COM requires datetime objects, not ISO strings
    start_dt = datetime.fromisoformat(start_date)
    appt.Start = start_dt

    if end_date:
        appt.End = datetime.fromisoformat(end_date)
    else:
        appt.End = start_dt + timedelta(hours=1)

    if location:
        appt.Location = location
    if description:
        appt.Body = description
    if all_day:
        appt.AllDayEvent = True

    if target_calendar:
        appt.Save()
        moved = appt.Move(target_calendar)
        moved.Display()
    else:
        appt.Display()

    return {"status": "displayed", "message": "Calendar event opened for user confirmation."}

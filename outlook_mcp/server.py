"""Outlook MCP Server — exposes Outlook operations as MCP tools via stdio."""

import json
import sys
from typing import Optional

from mcp.server.fastmcp import FastMCP

from . import outlook

mcp = FastMCP("outlook-mcp-server")


def _json(data) -> str:
    """Format response as indented JSON string."""
    return json.dumps(data, ensure_ascii=False, indent=2, default=str)


# ─── 1. Accounts & Folders ─────────────────────────────────────


@mcp.tool()
def listAccounts() -> str:
    """List all configured email accounts in Outlook (name, email, type)."""
    return _json(outlook.list_accounts())


@mcp.tool()
def listFolders(accountId: Optional[str] = None) -> str:
    """List all mail folders with message counts.

    Args:
        accountId: Optional account name to filter by.
    """
    return _json(outlook.list_folders(accountId))


@mcp.tool()
def createFolder(parentFolderPath: str, name: str) -> str:
    """Create a new subfolder under the specified parent folder.

    Args:
        parentFolderPath: Path of the parent folder (e.g. "AccountName/Inbox").
        name: Name for the new folder.
    """
    return _json(outlook.create_folder(parentFolderPath, name))


# ─── 2. Search & Read ──────────────────────────────────────────


@mcp.tool()
def searchMessages(
    query: str,
    startDate: Optional[str] = None,
    endDate: Optional[str] = None,
    maxResults: int = 50,
    sortOrder: str = "desc",
) -> str:
    """Search messages by keyword across all folders.

    Args:
        query: Search keyword (matches subject, sender, recipient).
        startDate: Start date filter (ISO 8601).
        endDate: End date filter (ISO 8601).
        maxResults: Maximum results to return (default 50, max 200).
        sortOrder: "desc" (newest first, default) or "asc".
    """
    return _json(outlook.search_messages(query, startDate, endDate, maxResults, sortOrder))


@mcp.tool()
def getRecentMessages(
    folderPath: Optional[str] = None,
    daysBack: int = 7,
    maxResults: int = 50,
    unreadOnly: bool = False,
) -> str:
    """Get recent messages, optionally filtered by folder and read status.

    Args:
        folderPath: Folder path (default: all inboxes).
        daysBack: Number of days to look back (default 7).
        maxResults: Maximum results (default 50, max 200).
        unreadOnly: Only return unread messages (default false).
    """
    return _json(outlook.get_recent_messages(folderPath, daysBack, maxResults, unreadOnly))


@mcp.tool()
def getMessage(
    messageId: str,
    folderPath: str,
    saveAttachments: bool = False,
) -> str:
    """Read full message content including body and attachments.

    Args:
        messageId: The message EntryID.
        folderPath: Folder path where the message resides.
        saveAttachments: Whether to save attachments to a temp directory (default false).
    """
    return _json(outlook.get_message(messageId, folderPath, saveAttachments))


# ─── 3. Mail Operations ────────────────────────────────────────


@mcp.tool()
def sendMail(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    isHtml: bool = False,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Compose a new email and open it in Outlook for user confirmation.

    Args:
        to: Recipient email address(es).
        subject: Email subject.
        body: Email body text.
        cc: CC recipients.
        bcc: BCC recipients.
        isHtml: Whether body is HTML (default false).
        fromAccount: Send-from account (email or display name).
        attachments: List of file paths to attach.
    """
    return _json(outlook.send_mail(to, subject, body, cc, bcc, isHtml, fromAccount, attachments))


@mcp.tool()
def replyToMessage(
    messageId: str,
    folderPath: str,
    body: str,
    replyAll: bool = False,
    isHtml: bool = False,
    to: Optional[str] = None,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Reply to a message (opens compose window for user confirmation).

    Args:
        messageId: Original message EntryID.
        folderPath: Folder path of the original message.
        body: Reply body text.
        replyAll: Reply to all recipients (default false).
        isHtml: Whether body is HTML.
        to: Override recipient.
        cc: CC recipients.
        bcc: BCC recipients.
        fromAccount: Send-from account.
        attachments: File paths to attach.
    """
    return _json(outlook.reply_to_message(
        messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, fromAccount, attachments
    ))


@mcp.tool()
def forwardMessage(
    messageId: str,
    folderPath: str,
    to: str,
    body: Optional[str] = None,
    isHtml: bool = False,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Forward a message (opens compose window for user confirmation).

    Args:
        messageId: Original message EntryID.
        folderPath: Folder path of the original message.
        to: Forward recipient.
        body: Additional body text prepended to the original.
        isHtml: Whether body is HTML.
        cc: CC recipients.
        bcc: BCC recipients.
        fromAccount: Send-from account.
        attachments: Additional file paths to attach.
    """
    return _json(outlook.forward_message(
        messageId, folderPath, to, body, isHtml, cc, bcc, fromAccount, attachments
    ))


@mcp.tool()
def updateMessage(
    messageId: str,
    folderPath: str,
    read: Optional[bool] = None,
    flagged: Optional[bool] = None,
    moveTo: Optional[str] = None,
    trash: bool = False,
) -> str:
    """Update message status: mark read/unread, flag/unflag, move, or trash.

    Args:
        messageId: Message EntryID.
        folderPath: Current folder path.
        read: Set read status (true=read, false=unread).
        flagged: Set flag status (true=flagged, false=unflagged).
        moveTo: Target folder path to move message to (cannot use with trash).
        trash: Move to deleted items (cannot use with moveTo).
    """
    return _json(outlook.update_message(messageId, folderPath, read, flagged, moveTo, trash))


@mcp.tool()
def deleteMessages(
    messageIds: list[str],
    folderPath: str,
) -> str:
    """Batch delete messages.

    Args:
        messageIds: List of message EntryIDs to delete.
        folderPath: Folder path where the messages reside.
    """
    return _json(outlook.delete_messages(messageIds, folderPath))


# ─── 4. Contacts ───────────────────────────────────────────────


@mcp.tool()
def searchContacts(query: str) -> str:
    """Search contacts by name or email.

    Args:
        query: Search keyword (matches name or email).
    """
    return _json(outlook.search_contacts(query))


# ─── 5. Calendar ───────────────────────────────────────────────


@mcp.tool()
def listCalendars() -> str:
    """List all calendars in Outlook."""
    return _json(outlook.list_calendars())


@mcp.tool()
def createEvent(
    title: str,
    startDate: str,
    endDate: Optional[str] = None,
    location: Optional[str] = None,
    description: Optional[str] = None,
    calendarId: Optional[str] = None,
    allDay: bool = False,
) -> str:
    """Create a calendar event (opens in Outlook for user confirmation).

    Args:
        title: Event title.
        startDate: Start time (ISO 8601).
        endDate: End time (default: start + 1 hour).
        location: Event location.
        description: Event description.
        calendarId: Target calendar ID/path.
        allDay: Whether this is an all-day event (default false).
    """
    return _json(outlook.create_event(
        title, startDate, endDate, location, description, calendarId, allDay
    ))


# ─── Entry point ────────────────────────────────────────────────


def main():
    """Run the MCP server with stdio transport."""
    if "--version" in sys.argv or "-V" in sys.argv:
        try:
            from importlib.metadata import version
            ver = version("outlook-mcp-server-server")
        except Exception:
            ver = "dev"
        print(f"outlook-mcp-server {ver}")
        return
    if "--help" in sys.argv or "-h" in sys.argv:
        print("Usage: outlook-mcp-server")
        print("  MCP server for Outlook (stdio transport)")
        print()
        print("Options:")
        print("  -V, --version  Show version and exit")
        print("  -h, --help     Show this help and exit")
        return
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

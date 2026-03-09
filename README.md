# outlook-mcp

[![PyPI version](https://img.shields.io/pypi/v/outlook-mcp-server)](https://pypi.org/project/outlook-mcp-server/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-mcp-server)](https://pypi.org/project/outlook-mcp-server/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

MCP server for Outlook — enables AI assistants to access email, contacts, and calendars via Windows COM interface.

Built on the [Model Context Protocol](https://modelcontextprotocol.io/) (MCP), works with [Claude Code](https://claude.ai/code) and other MCP-compatible clients.

## Features

- **Email**: Search, read, send, reply, forward, flag, move, delete
- **Contacts**: Search by name or email
- **Calendar**: List calendars, create events
- **Multi-account**: Supports all accounts configured in Outlook
- **Safe by design**: Send/reply/forward opens Outlook compose window for user confirmation — never auto-sends
- **Local only**: All operations via local COM interface, no cloud API or OAuth needed

## Requirements

- Windows with Outlook desktop client installed **and running**
- Python >= 3.10

## Install

```bash
pip install outlook-mcp-server
```

Verify installation:

```bash
outlook-mcp --version
```

## Usage with Claude Code

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

Then restart Claude Code. The 14 Outlook tools will be available to Claude automatically.

### Alternative: run from source

```bash
git clone https://github.com/lihaokun/outlook-mcp.git
cd outlook-mcp
pip install -e .
```

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": [
        "--directory", "C:\\path\\to\\outlook-mcp",
        "run", "outlook-mcp"
      ]
    }
  }
}
```

## Available Tools

### Accounts & Folders

| Tool | Description |
|------|-------------|
| `listAccounts` | List all configured email accounts (name, email, type) |
| `listFolders` | List all folders with message counts and unread counts |
| `createFolder` | Create a new subfolder under a specified parent |

### Email Search & Read

| Tool | Description |
|------|-------------|
| `searchMessages` | Search by keyword (subject/sender/recipient), with date range and sort |
| `getRecentMessages` | Get recent messages by days, folder, and unread filter |
| `getMessage` | Read full message content (body, HTML, attachments) |

### Email Operations

| Tool | Description |
|------|-------------|
| `sendMail` | Compose new email (opens Outlook compose window) |
| `replyToMessage` | Reply or reply-all (opens compose window) |
| `forwardMessage` | Forward with original attachments (opens compose window) |
| `updateMessage` | Mark read/unread, flag/unflag, move to folder, or trash |
| `deleteMessages` | Batch delete messages |

### Contacts & Calendar

| Tool | Description |
|------|-------------|
| `searchContacts` | Search contacts by name or email |
| `listCalendars` | List all calendars with item count and writable status |
| `createEvent` | Create calendar event (opens Outlook event window) |

## Safety

This server is designed with safety as a priority:

- **No auto-send**: `sendMail`, `replyToMessage`, `forwardMessage`, and `createEvent` all call `.Display()` instead of `.Send()` / `.Save()`. This opens the Outlook compose window so the user can review and confirm before sending.
- **Outlook Object Model Guard**: Outlook may show additional security prompts for send operations. This is expected behavior.
- **Local only**: All data stays on your machine. No external API calls, no cloud services, no OAuth tokens.

## Technical Details

- **COM interface**: Uses `pywin32` (`win32com.client`) to control Outlook
- **MCP transport**: stdio (standard input/output)
- **Message ID**: Uses Outlook's `EntryID` as unique identifier
- **Search**: Uses `Items.Restrict()` with DASL filters for efficient server-side filtering
- **Internationalization**: Automatically handles both "Inbox" and "收件箱" folder names

## Related Projects

- [mail-assistant](https://github.com/lihaokun/mail-assistant) — AI-powered mail assistant using MCP, works with Thunderbird and Outlook
- [thunderbird-mcp](https://github.com/TKasperczyk/thunderbird-mcp) — MCP server for Thunderbird

## License

MIT

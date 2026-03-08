# outlook-mcp

MCP server for Outlook — enables AI assistants to access email, contacts, and calendars via Windows COM interface.

## Requirements

- Windows with Outlook desktop client installed and running
- Python >= 3.10
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

## Install

```bash
git clone https://github.com/lihaokun/outlook-mcp.git
cd outlook-mcp
uv venv && uv pip install -e .
```

Or with pip:

```bash
pip install -e .
```

## Usage with Claude Code

Add to your project's `.mcp.json`:

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

Or if installed via pip:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `listAccounts` | List configured email accounts |
| `listFolders` | List folders with message counts |
| `createFolder` | Create a new subfolder |
| `searchMessages` | Search messages by keyword and date |
| `getRecentMessages` | Get recent messages |
| `getMessage` | Read full message content |
| `sendMail` | Compose new email |
| `replyToMessage` | Reply to a message |
| `forwardMessage` | Forward a message |
| `updateMessage` | Mark read/unread, flag, move, trash |
| `deleteMessages` | Batch delete messages |
| `searchContacts` | Search contacts |
| `listCalendars` | List calendars |
| `createEvent` | Create calendar event |

## Notes

- Outlook must be running for COM interface to work
- Send/reply/forward opens a compose window for user confirmation — does not auto-send
- Uses `EntryID` as unique message identifier
- Outlook Object Model Guard may show security prompts when sending mail

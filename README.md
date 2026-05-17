# outlook-mcp

A local MCP server that connects Claude to your classic Outlook desktop app — mail, calendar, contacts, tasks, and more — via the Outlook COM API on Windows. No Azure app registration or Microsoft Graph API required. Authentication piggybacks on whatever account Outlook is already signed into.

## Requirements

- Windows 10 or 11
- **Outlook Classic** (`OUTLOOK.EXE`) — the version shipped with Microsoft 365 / Office. The new Outlook (`olk.exe`) is **not** supported as it has no COM interface.
- Python 3.10+

## Installation

**Step 1 — Install the package**

```
pip install C:\path\to\outlook_mcp
```

Replace the path with wherever the source folder is located.

**Step 2 — Register with your MCP client**

For **Claude Code:**

```
claude mcp add outlook-mcp python -m outlook_mcp
```

For **Claude Desktop**, add this to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "python",
      "args": ["-m", "outlook_mcp"]
    }
  }
}
```

**Step 3 — Restart your MCP client**

Make sure Outlook is open, then restart Claude Code / Claude Desktop. On first use there may be a 5–10 second pause while the Outlook COM connection initialises — this is normal.

## Tools

| Category      | Tools |
|---------------|-------|
| Mail          | `outlook_list_mails`, `outlook_search_mails`, `outlook_remote_search_mails`, `outlook_get_mail`, `outlook_send_mail`, `outlook_reply_mail`, `outlook_forward_mail`, `outlook_move_mail`, `outlook_delete_mail`, `outlook_mark_mail`, `outlook_save_attachments` |
| Folders       | `outlook_list_folders`, `outlook_create_folder` |
| Calendar      | `outlook_list_events`, `outlook_get_event`, `outlook_create_event`, `outlook_update_event`, `outlook_delete_event`, `outlook_respond_event` |
| Contacts      | `outlook_list_contacts`, `outlook_search_contacts`, `outlook_get_contact` |
| Tasks         | `outlook_list_tasks`, `outlook_create_task`, `outlook_complete_task` |
| Categories    | `outlook_list_categories`, `outlook_set_category` |
| Rules         | `outlook_list_rules`, `outlook_toggle_rule` |
| Out-of-Office | `outlook_get_out_of_office` |
| Account       | `outlook_whoami` |

### Search tools

**`outlook_search_mails`** — fast search within a single folder using an indexed DASL subject filter. Best for targeted lookups when you know which folder the email is in.

**`outlook_remote_search_mails`** — full-text search across your entire mailbox using the Windows Search index (the same engine as Outlook's search bar). Searches subject and body in one pass, spans all folders, and can find content beyond the local MAPI cache window.

## Troubleshooting

**Outlook didn't connect** — make sure Outlook Classic is open and signed in before starting the MCP client.

**Search returns no results** — `outlook_remote_search_mails` relies on the Windows Search index. The index builds in the background; newly received emails may take a few minutes to appear.

**Send / reply is blocked** — some corporate IT policies restrict programmatic access to Outlook. Go to Outlook → File → Options → Trust Center → Programmatic Access and check the setting there.

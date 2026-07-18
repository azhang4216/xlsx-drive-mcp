# xlsx-drive-mcp

An MCP server that lets AI agents (Claude and others) read and write `.xlsx` files
stored on Google Drive. Fills the gap between Excel MCP servers (local files only)
and Google Drive MCP servers (native Google Sheets only).

## Features

- Read/write cell values, formulas, formatting, charts
- Full sheet management (create, delete, rename, copy)
- Google Drive authentication with zero extra setup for existing workspace-mcp users
- Conflict detection on concurrent edits

## Installation

```bash
pip install xlsx-drive-mcp
```

Requires Python 3.10+.

## Authentication

### Option A: Existing workspace-mcp users (zero extra setup)

If you already use [workspace-mcp](https://workspacemcp.com), credentials are reused
automatically. Skip to Claude Config below.

### Option B: Standalone setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/apis/credentials)
2. Create a project (or use an existing one)
3. Enable the **Google Drive API**
4. Create credentials: **OAuth client ID** → **Desktop app**
5. Download the client ID and secret, then:

```bash
export GOOGLE_CLIENT_ID=your_client_id
export GOOGLE_CLIENT_SECRET=your_client_secret
xlsx-drive-mcp auth --user you@example.com
```

This opens a browser, completes the OAuth flow, and stores credentials at
`~/.xlsx-drive-mcp/credentials/you@example.com.json`.

**Add to your global `.gitignore`:**
```
~/.xlsx-drive-mcp/
```

## Claude Config

```json
{
  "mcpServers": {
    "xlsx-drive": {
      "command": "xlsx-drive-mcp",
      "args": ["--user", "your-actual-email@example.com"]
    }
  }
}
```

Replace `your-actual-email@example.com` with the email you authenticated with.
If you have only one credential file, `--user` can be omitted.

## Scope note

- **workspace-mcp users** inherit workspace-mcp's scope, which is typically full Drive access.
- **Standalone users** get `drive.file` scope -- access only to files this app created or opened.
  With `drive.file` scope, `list_xlsx_files` only returns files you've previously opened with
  this tool. To search all Drive files, use full `drive` scope or the workspace-mcp path.

## Finding a folder_id

Open any Drive folder in your browser. The URL looks like:
`https://drive.google.com/drive/folders/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74`
The `folder_id` is the last segment: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74`.

## Important warnings

- **`write_range` over formula cells** replaces formulas with static values.
- **`read_range` with `data_only=true`** (default) returns cached values, not live formula results.
  Use `data_only=false` to get formula strings instead.
- **`copy_sheet`** silently drops charts and images from the copied sheet (openpyxl limitation).
- **`delete_xlsx`** moves the file to trash. The user has 30 days to recover it from Drive UI.

## Tool reference

| Tool | Description |
|---|---|
| `list_xlsx_files` | Search Drive for xlsx files |
| `get_xlsx_info` | Sheet names, data dimensions, named ranges |
| `create_xlsx` | Create a new blank xlsx on Drive |
| `delete_xlsx` | Move an xlsx to trash |
| `list_sheets` | List all sheets |
| `create_sheet` | Add a new sheet |
| `delete_sheet` | Remove a sheet |
| `rename_sheet` | Rename a sheet |
| `copy_sheet` | Duplicate a sheet (drops charts/images) |
| `read_range` | Read cell values |
| `write_range` | Write cell values |
| `append_rows` | Append rows below last data |
| `clear_range` | Clear values, preserve formatting |
| `read_format` | Read cell formatting (colors, bold, borders) |
| `format_range` | Apply formatting to a range |
| `set_column_width` | Set column widths |
| `set_row_height` | Set row heights |
| `merge_cells` | Merge a cell range |
| `unmerge_cells` | Unmerge a cell range |
| `create_chart` | Create a bar, line, pie, or scatter chart |
| `write_formula` | Write a formula to a single cell |

## Troubleshooting

**MCP server won't start / tool not found:**
Claude Desktop launches MCP servers as background processes with no visible output.
To see server logs, run the server manually:
```bash
xlsx-drive-mcp --user your@email.com
```
Any errors will appear in the terminal.

**`list_xlsx_files` returns no results:**
With `drive.file` scope, only files this app created or previously opened are visible.
Use `get_xlsx_info(file_id)` directly with a known `file_id` from the Drive URL instead.

**Shared Drive / Team Drive:**
Not officially tested in v1. May work; not supported. Open an issue if you test it.

**Service accounts:**
Not supported in v1. Tracked as a roadmap item.

## License

MIT

# xlsx-drive-mcp Design Spec

**Date:** 2026-07-18
**Status:** Approved

## Overview

An open-source MCP server that enables AI agents (Claude and others) to read and write `.xlsx` files stored on Google Drive. Fills a gap in the current MCP ecosystem: existing Excel MCP servers only operate on local files; existing Google Drive MCP servers only operate on native Google Sheets.

## Problem

No existing MCP tool bridges Google Drive authentication with xlsx file editing. Agents that need to work with `.xlsx` files on Drive must either convert to Google Sheets (lossy) or orchestrate a manual download-edit-upload cycle outside of MCP.

## Architecture

**Framework:** FastMCP (stdio transport), Python 3.10+

**Core pattern for every tool:**
1. Authenticate with Google Drive
2. Download xlsx to in-memory BytesIO
3. Open with openpyxl
4. Apply operation
5. Save to BytesIO
6. Upload back to Drive (update existing file)

**Module layout:**
```
src/xlsx_drive_mcp/
├── server.py       # FastMCP app, registers all tools
├── auth.py         # OAuth2 with workspace-mcp fallback
├── drive.py        # download/upload helpers (BytesIO)
└── tools/
    ├── files.py    # Drive-level file operations
    ├── sheets.py   # Worksheet management
    ├── data.py     # Read/write cell values
    ├── format.py   # Read/write cell formatting
    ├── charts.py   # Chart creation
    └── formulas.py # Formula writing
```

## Authentication

Two-tier auth, resolved at startup:

1. **workspace-mcp shortcut:** if `~/.google_workspace_mcp/credentials/{email}.json` exists, load it directly. Zero new setup for existing workspace-mcp users. Existing scopes include `drive` and `drive.file`.
2. **Standalone OAuth flow:** if no workspace-mcp credentials, run browser-based OAuth2 flow. Store credentials at `~/.xlsx-drive-mcp/credentials/{email}.json`. Bundled OAuth client ID used by default; users can override with `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET` env vars.

Required scope: `https://www.googleapis.com/auth/drive`

## Tools

### files.py
- `list_xlsx_files(folder_id?, query?)` -- search Drive for xlsx files; returns file_id, name, modified date
- `get_xlsx_info(file_id)` -- sheet names, dimensions (rows x cols per sheet), named ranges
- `create_xlsx(name, folder_id?)` -- create new blank xlsx on Drive; returns file_id
- `delete_xlsx(file_id)` -- move to trash

### sheets.py
- `list_sheets(file_id)` -- sheet names and indices
- `create_sheet(file_id, name, position?)` -- add sheet
- `delete_sheet(file_id, sheet_name)` -- remove sheet
- `rename_sheet(file_id, sheet_name, new_name)` -- rename
- `copy_sheet(file_id, sheet_name, new_name)` -- duplicate within workbook

### data.py
- `read_range(file_id, sheet, range)` -- read 2D array of values (e.g. `A1:D10`)
- `write_range(file_id, sheet, range, values)` -- write 2D array of values
- `append_rows(file_id, sheet, rows)` -- append rows below last row of data
- `clear_range(file_id, sheet, range)` -- clear values, preserve formatting

### format.py
- `read_format(file_id, sheet, range)` -- returns per-cell: bg_color, font_color, bold, italic, font_size, number_format, border, merged status
- `format_range(file_id, sheet, range, bold?, italic?, font_size?, font_color?, bg_color?, number_format?, border?)` -- apply formatting
- `set_column_width(file_id, sheet, columns, width)` -- set column widths
- `set_row_height(file_id, sheet, rows, height)` -- set row heights
- `merge_cells(file_id, sheet, range)` -- merge cells
- `unmerge_cells(file_id, sheet, range)` -- unmerge cells

### charts.py
- `create_chart(file_id, sheet, chart_type, data_range, title?, position?)` -- chart types: bar, line, pie, scatter

### formulas.py
- `write_formula(file_id, sheet, cell, formula)` -- write formula string (e.g. `=SUM(A1:A10)`)

## Project Structure

```
xlsx-drive-mcp/
├── pyproject.toml
├── README.md
├── LICENSE                        # MIT
├── .github/
│   └── workflows/
│       └── publish.yml            # auto-publish to PyPI on version tag
├── src/
│   └── xlsx_drive_mcp/
│       ├── __init__.py
│       ├── server.py
│       ├── auth.py
│       ├── drive.py
│       └── tools/
│           ├── __init__.py
│           ├── files.py
│           ├── sheets.py
│           ├── data.py
│           ├── format.py
│           ├── charts.py
│           └── formulas.py
└── tests/
    └── test_tools.py
```

## Distribution

- **PyPI:** `pip install xlsx-drive-mcp` (primary)
- **GitHub:** `pip install git+https://github.com/angep/xlsx-drive-mcp` (source)
- **Entry point:** `xlsx-drive-mcp` CLI command → runs server over stdio

## Claude Config (shown in README)

```json
{
  "mcpServers": {
    "xlsx-drive": {
      "command": "xlsx-drive-mcp",
      "args": ["--user", "you@example.com"]
    }
  }
}
```

## CI/CD

GitHub Actions workflow (`publish.yml`):
- On push to `main`: run tests
- On version tag (`v*`): run tests, build, publish to PyPI via Trusted Publisher (OIDC, no API key needed)

## Testing

`tests/test_tools.py` uses a mock Google Drive (BytesIO in-memory) and real openpyxl operations to test each tool in isolation. No real Drive credentials needed for tests.

## Non-goals (v1)

- Streaming large files (>50MB)
- Conditional formatting rules
- Pivot tables
- Macros / VBA
- In-process workbook caching (optimization, can be added later)

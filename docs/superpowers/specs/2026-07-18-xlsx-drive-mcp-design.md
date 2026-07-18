# xlsx-drive-mcp Design Spec

**Date:** 2026-07-18
**Status:** Approved (rev 2 -- updated after engineering, security, and OSS review)

## Overview

An open-source MCP server that enables AI agents (Claude and others) to read and write `.xlsx` files stored on Google Drive. Fills a gap in the current MCP ecosystem: existing Excel MCP servers only operate on local files; existing Google Drive MCP servers only operate on native Google Sheets.

## Problem

No existing MCP tool bridges Google Drive authentication with xlsx file editing. Agents that need to work with `.xlsx` files on Drive must either convert to Google Sheets (lossy) or orchestrate a manual download-edit-upload cycle outside of MCP.

## Architecture

**Framework:** FastMCP (stdio transport), Python 3.10+

**Core pattern for every tool:**
1. Authenticate with Google Drive
2. Check file size (reject >50MB with a clear error)
3. Download xlsx to in-memory BytesIO; record Drive revision ID (`headRevisionId`)
4. Open with openpyxl
5. Apply operation
6. Save to BytesIO
7. Upload back to Drive with `If-Match: {revision_id}` header
8. On 412 Precondition Failed: return structured conflict error (do not retry automatically)

All logging goes to stderr only. stdout is reserved exclusively for the MCP JSON-RPC transport.

**Module layout:**
```
src/xlsx_drive_mcp/
├── server.py       # FastMCP app, registers all tools
├── auth.py         # OAuth2 with workspace-mcp fallback
├── drive.py        # download/upload helpers (BytesIO, ETag, size check)
└── tools/
    ├── files.py    # Drive-level file operations
    ├── sheets.py   # Worksheet management
    ├── data.py     # Read/write cell values
    ├── format.py   # Read/write cell formatting
    ├── charts.py   # Chart creation
    └── formulas.py # Formula writing (convenience wrapper over data.py)
```

## Authentication

**No OAuth client ID is bundled with the package.** Google caps unverified OAuth apps at 100 users; shipping a client ID in a public PyPI package would break the project at any meaningful adoption.

### Two-tier auth, resolved at startup:

1. **workspace-mcp shortcut (recommended):** if `~/.google_workspace_mcp/credentials/{email}.json` exists, load it directly. Zero new setup for existing workspace-mcp users. These credentials already hold `drive` scope.

2. **Standalone OAuth flow:** requires the user to create a Google Cloud Console OAuth client (Desktop app type) and provide credentials via env vars:
   ```
   GOOGLE_CLIENT_ID=...
   GOOGLE_CLIENT_SECRET=...
   ```
   Credentials stored at `~/.xlsx-drive-mcp/credentials/{email}.json` with `0600` permissions.

### Auth CLI subcommand (required before first MCP use for standalone users)

MCP servers run headlessly -- no browser access, no terminal. OAuth browser flow must run before the server is added to Claude's config:

```bash
xlsx-drive-mcp auth --user you@example.com
```

This opens the browser, completes the OAuth flow, and stores the credential file. The MCP server then starts without prompting.

### Auto-detect

If `--user` is omitted and exactly one credential file exists in either credential directory, it is used automatically. If multiple exist and `--user` is omitted, the server exits with a clear error listing the available emails.

### Required scope

`https://www.googleapis.com/auth/drive.file` for standalone OAuth (access only to files created or opened by this app).

workspace-mcp shortcut users inherit whatever scope workspace-mcp requested (typically full `drive`). README must document this distinction so users understand what they are authorizing.

### Token refresh

Tokens are refreshed transparently before each API call using the stored refresh token. If the refresh token has been revoked, the server exits with a clear error directing the user to re-run `xlsx-drive-mcp auth`.

### Credential file security

All credential files created by this tool are written with `0600` permissions (owner read/write only). The README must warn users not to commit credential files and to add `~/.xlsx-drive-mcp/` to their global `.gitignore`.

## Tools

All tools that accept a cell range use the parameter name `cell_range` (not `range`, which shadows a Python builtin). All tools that reference a sheet use the parameter name `sheet_name` consistently.

### Input validation (applied by all tools before any Drive call)

- `cell_range`: must match `^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$` or a named range. Maximum range area: 1,000,000 cells. Reject with clear error otherwise.
- `file_id`: non-empty string.
- `sheet_name`: must exist in the workbook (validated after download).
- File size: if Drive metadata shows the file is >50MB, return an error before downloading.

---

### files.py

**`list_xlsx_files(folder_id?, name_contains?)`**
- Search Drive for xlsx files. `name_contains` is a plain string used for filename substring search only -- it is never passed raw to the Drive API `q` parameter. The server constructs the `q` string server-side: `mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and name contains '{escaped}'`.
- Returns: `[{file_id, name, modified_time, size_bytes}]`. Includes a `next_page_token` field if results are truncated.
- Note: `drive.file` scope only returns files this app created or opened. Users with full `drive` scope see all matching files. README must document this.

**`get_xlsx_info(file_id)`**
- Returns sheet names, actual data dimensions per sheet (scanned, not `max_row`/`max_column` which are unreliable after clears), and named ranges.

**`create_xlsx(name, folder_id?)`**
- Creates a new blank xlsx on Drive. Returns `file_id`.

**`delete_xlsx(file_id)`**
- Moves file to trash. This is non-reversible from the agent's perspective (user has 30 days to recover from Drive UI). README must warn agents to use this tool cautiously. No confirmation step in v1 -- the MCP caller is responsible for confirming before invoking.

---

### sheets.py

**`list_sheets(file_id)`**
- Returns: `[{name, index}]`

**`create_sheet(file_id, name, position?)`**
- Position is 0-based integer index. Defaults to last.
- Returns: `{name, index}` of the created sheet.

**`delete_sheet(file_id, sheet_name)`**
- Removes sheet. Raises error if it is the only sheet (Excel/openpyxl requirement).

**`rename_sheet(file_id, sheet_name, new_name)`**

**`copy_sheet(file_id, sheet_name, new_name)`**
- Duplicates within the same workbook.
- **Known limitation:** openpyxl's `copy_worksheet` silently drops charts, images, and some merged cell configurations. Documented in tool description and README.
- Returns: `{name, index}` of the new sheet.

---

### data.py

**`read_range(file_id, sheet_name, cell_range, data_only?)`**
- Reads a 2D array of cell values (e.g. `A1:D10`).
- `data_only` (default `true`): when true, opens workbook with `data_only=True` -- returns cached computed values for formula cells (the value from the last time Excel/Sheets calculated the workbook), but formula strings are lost. When false, returns formula strings instead of computed values. **These are mutually exclusive in openpyxl; you cannot get both in a single call.** The tool description must state this clearly.
- Returns: `{values: [[...], ...], cell_range: "A1:D10"}` (actual range written, which may be smaller than requested if the sheet has fewer rows/cols).

**`write_range(file_id, sheet_name, cell_range, values)`**
- `cell_range` is the **anchor** (top-left cell, e.g. `A1` or `A1:D10`). The `values` 2D array determines the actual write extent. If a full range is provided (e.g. `A1:D10`), it is used only for reference -- the actual cells written are determined by the dimensions of `values`. Extra cells in the range beyond the values array are not cleared. This contract is explicit and documented.
- Formula strings (starting with `=`) are written as formulas, not as literal strings.

**`append_rows(file_id, sheet_name, rows)`**
- Appends rows below the last non-empty row. "Last non-empty row" is determined by scanning from the bottom of the used range, not using `max_row` (which is a high-water mark and unreliable after clears). Documented in tool description.

**`clear_range(file_id, sheet_name, cell_range)`**
- Clears cell values, preserves formatting.
- If the range intersects a merged region, raises a descriptive error instructing the caller to unmerge first.

---

### format.py

**`read_format(file_id, sheet_name, cell_range)`**
- Returns formatting for each cell in the range.
- Return structure:
  ```json
  {
    "cells": [
      {
        "address": "A1",
        "bold": true,
        "italic": false,
        "font_size": 12,
        "font_color": "FF000000",
        "bg_color": "FFFFFF00",
        "number_format": "General",
        "border": {
          "top": "thin",
          "bottom": null,
          "left": null,
          "right": "thin"
        },
        "merged": false,
        "merge_range": null
      }
    ]
  }
  ```
- Colors are hex strings in ARGB format (`"FFRRGGBB"`). Theme-indexed colors are resolved to their hex equivalent. Transparent/no-fill is returned as `null`.
- `merge_range` is the full merge range (e.g. `"A1:C3"`) for the anchor cell of a merged region, or `null` if not merged. Non-anchor cells within a merged region return `"merged": true` and `"merge_range": null`.

**`format_range(file_id, sheet_name, cell_range, bold?, italic?, font_size?, font_color?, bg_color?, number_format?, border?)`**
- All formatting params are optional; omitted params are left unchanged.
- `border` accepts either a string shorthand (`"thin"`, `"medium"`, `"thick"`, `"none"`) applying to all 4 sides, or an object `{"top": "thin", "bottom": null, ...}` for per-side control. `null` means leave existing border unchanged; `"none"` explicitly removes the border on that side.
- Colors are accepted as `"RRGGBB"` (6-digit hex) or `"FFRRGGBB"` (8-digit ARGB).

**`set_column_width(file_id, sheet_name, columns, width)`**
- `columns`: a single column letter (`"A"`), a list (`["A", "B", "C"]`), or a range string (`"A:C"`). All three formats are accepted.
- `width`: Excel character width units (float).

**`set_row_height(file_id, sheet_name, rows, height)`**
- `rows`: a single row number (int), a list of ints, or a range string (`"1:5"`).
- `height`: points (float).

**`merge_cells(file_id, sheet_name, cell_range)`**

**`unmerge_cells(file_id, sheet_name, cell_range)`**

---

### charts.py

**`create_chart(file_id, sheet_name, chart_type, chart_subtype, values_range, category_range?, series_names?, title?, anchor_cell?)`**

- `chart_type`: `"bar"`, `"line"`, `"pie"`, `"scatter"`
- `chart_subtype`:
  - bar: `"col"` (vertical, default) or `"bar"` (horizontal)
  - line: `"line"` (default) or `"smooth"`
  - pie: `"pie"` (only option)
  - scatter: `"marker"` (default), `"smooth_line"`, `"straight_line"`
- `values_range`: cell range containing the data series values (e.g. `"B2:D10"`). Columns = series.
- `category_range?`: cell range containing category/X-axis labels (e.g. `"A2:A10"`). If omitted, uses row numbers.
- `series_names?`: list of series label strings, or a cell range containing them (e.g. `"B1:D1"`). If omitted, uses default series names.
- `title?`: chart title string.
- `anchor_cell?`: top-left cell for chart placement (e.g. `"E2"`). Defaults to one column right of the data range.

Returns: `{sheet_name, anchor_cell}` of the placed chart.

---

### formulas.py

**`write_formula(file_id, sheet_name, cell, formula)`**
- Convenience wrapper: equivalent to `write_range` with a 1x1 values array containing `formula`.
- `formula` must start with `=`. The following function names are blocked and return an error: `WEBSERVICE`, `FILTERXML`, `IMPORTDATA`, `IMPORTFEED`, `IMPORTHTML`, `IMPORTRANGE`, `IMPORTXML` -- these make external network connections when the file is opened in Excel/Sheets and can be used to exfiltrate data.

---

## Project Structure

```
xlsx-drive-mcp/
├── pyproject.toml              # package metadata, keywords, entry points
├── README.md                   # setup guide with screenshots, troubleshooting
├── CHANGELOG.md                # Keep a Changelog format
├── CONTRIBUTING.md             # how to run tests, add tools, PR policy
├── LICENSE                     # MIT
├── .github/
│   └── workflows/
│       ├── test.yml            # run tests on every push / PR
│       └── publish.yml         # publish to PyPI on version tag
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
    ├── fixtures/               # real .xlsx files covering edge cases
    │   ├── merged_cells.xlsx
    │   ├── named_ranges.xlsx
    │   ├── formulas.xlsx
    │   ├── multi_sheet.xlsx
    │   └── charts.xlsx
    ├── unit/
    │   ├── test_files.py
    │   ├── test_sheets.py
    │   ├── test_data.py
    │   ├── test_format.py
    │   ├── test_charts.py
    │   └── test_formulas.py
    └── integration/            # require DRIVE_TEST_FILE_ID env var
        └── test_drive_roundtrip.py
```

## Distribution

- **PyPI:** `pip install xlsx-drive-mcp` (primary)
- **GitHub:** `pip install git+https://github.com/angep/xlsx-drive-mcp` (source)
- **Entry points:**
  - `xlsx-drive-mcp` -- start MCP server (stdio)
  - `xlsx-drive-mcp auth` -- pre-auth CLI (must run before first MCP use for standalone users)

### pyproject.toml keywords (for PyPI discoverability)

```
keywords = ["google-drive", "excel", "xlsx", "openpyxl", "mcp",
            "model-context-protocol", "claude", "ai-agent", "spreadsheet"]
```

## Versioning

Semantic versioning. MCP tool names and parameter names are public API -- renaming or removing them requires a major version bump and a deprecation cycle (old tool kept with a deprecation notice in its description for one major version). `CHANGELOG.md` updated with every release.

## CI/CD

- `test.yml`: runs on every push and PR to `main`. Runs unit tests (no credentials needed). Integration tests are skipped unless `DRIVE_TEST_FILE_ID` secret is present.
- `publish.yml`: runs on `v*` tag push. Runs tests, builds, publishes to PyPI via Trusted Publisher (OIDC). Requires one-time setup: configure Trusted Publisher in PyPI UI pointing to this repo + workflow file + `pypi` GitHub environment.

## Testing strategy

- **Unit tests** (no credentials, no network): mock the `drive.py` download/upload layer with BytesIO. Use real fixture `.xlsx` files for openpyxl operations. One test file per tool module.
- **Integration tests** (gated behind `DRIVE_TEST_FILE_ID` env var): run against a real Drive file. Tests the full download-modify-upload-redownload-verify cycle. Run in a separate CI job with real credentials stored as GitHub secrets.
- Key scenarios that must be covered: merged cell intersection with clear_range, append_rows after a clear, concurrent write conflict (412 response), credential file missing, file too large (>50MB), malformed cell_range.

## README must cover

- Step-by-step OAuth setup with Google Cloud Console screenshots (for standalone users)
- How to find a `folder_id` from a Drive URL
- Troubleshooting: how to view server logs when running inside Claude Desktop
- Shared Drive / Team Drive support (not tested in v1 -- documented as untested)
- Service account support (not in v1 -- tracked as a roadmap item)
- Python version requirement (3.10+)
- Warning: `write_range` over cells containing formulas overwrites them with static values
- Warning: `copy_sheet` silently drops charts and images
- Warning: `read_range` with `data_only=true` (default) returns cached values, not live formula results

## Non-goals (v1)

- Streaming large files (>50MB)
- Conditional formatting rules
- Pivot tables
- Macros / VBA
- In-process workbook caching (optimization, can be added later)
- Service account authentication
- Shared Drive / Team Drive (untested; may work, not supported)

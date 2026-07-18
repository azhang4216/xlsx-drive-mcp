# xlsx-drive-mcp Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a FastMCP server that lets AI agents read and write `.xlsx` files stored on Google Drive.

**Architecture:** Each MCP tool follows a download-modify-upload cycle: authenticate, fetch the xlsx from Drive into BytesIO recording the revision ID, open with openpyxl, apply the operation, save back to BytesIO, re-upload with conflict detection. Business logic lives in pure functions (prefixed `_`) that take openpyxl Workbook objects; MCP wrappers are thin and call those pure functions directly, making unit tests fast and dependency-free.

**Tech Stack:** Python 3.10+, FastMCP 3.x, openpyxl 3.1+, google-api-python-client, google-auth-oauthlib

---

## File Map

| File | Responsibility |
|---|---|
| `pyproject.toml` | Package metadata, deps, entry points, PyPI keywords |
| `src/xlsx_drive_mcp/__init__.py` | Version string |
| `src/xlsx_drive_mcp/server.py` | FastMCP app creation, tool registration, CLI entry point |
| `src/xlsx_drive_mcp/auth.py` | Credential loading (workspace-mcp fallback → standalone), OAuth flow, token refresh |
| `src/xlsx_drive_mcp/drive.py` | Drive service factory, download/upload with revision check, size validation, cell range validation |
| `src/xlsx_drive_mcp/tools/files.py` | `list_xlsx_files`, `get_xlsx_info`, `create_xlsx`, `delete_xlsx` |
| `src/xlsx_drive_mcp/tools/sheets.py` | `list_sheets`, `create_sheet`, `delete_sheet`, `rename_sheet`, `copy_sheet` |
| `src/xlsx_drive_mcp/tools/data.py` | `read_range`, `write_range`, `append_rows`, `clear_range` |
| `src/xlsx_drive_mcp/tools/format.py` | `read_format`, `format_range`, `set_column_width`, `set_row_height`, `merge_cells`, `unmerge_cells` |
| `src/xlsx_drive_mcp/tools/charts.py` | `create_chart` |
| `src/xlsx_drive_mcp/tools/formulas.py` | `write_formula` (with denylist) |
| `tests/conftest.py` | Shared fixtures: in-memory workbook bytes, mock Drive context |
| `tests/fixtures/create_fixtures.py` | Script to generate fixture xlsx files |
| `tests/fixtures/*.xlsx` | Real xlsx files for edge cases |
| `tests/unit/test_drive.py` | Validate cell range, size check, conflict error |
| `tests/unit/test_auth.py` | Credential loading, auto-detect, missing credential error |
| `tests/unit/test_files.py` | list/create/delete/info pure logic |
| `tests/unit/test_sheets.py` | Sheet management pure logic |
| `tests/unit/test_data.py` | read/write/append/clear pure logic |
| `tests/unit/test_format.py` | Format read/write pure logic |
| `tests/unit/test_charts.py` | Chart creation logic |
| `tests/unit/test_formulas.py` | Formula denylist, write formula |
| `tests/integration/test_drive_roundtrip.py` | Full Drive roundtrip (gated behind env var) |
| `.github/workflows/test.yml` | Run unit tests on every push/PR |
| `.github/workflows/publish.yml` | Publish to PyPI on version tag |
| `README.md` | Setup, auth walkthrough, tool reference, troubleshooting |
| `CHANGELOG.md` | Keep a Changelog format |
| `CONTRIBUTING.md` | Local dev setup, how to add a tool, PR policy |

---

## Task 1: Project Scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `LICENSE`
- Create: `.gitignore`
- Create: `src/xlsx_drive_mcp/__init__.py`
- Create: `src/xlsx_drive_mcp/tools/__init__.py`

- [ ] **Step 1: Create directory structure**

```bash
cd /Users/ange/repos/xlsx-drive-mcp
mkdir -p src/xlsx_drive_mcp/tools
mkdir -p tests/unit tests/integration tests/fixtures
mkdir -p .github/workflows
```

- [ ] **Step 2: Create `pyproject.toml`**

```toml
[project]
name = "xlsx-drive-mcp"
version = "0.1.0"
description = "MCP server for reading and writing xlsx files on Google Drive"
readme = "README.md"
license = {text = "MIT"}
requires-python = ">=3.10"
keywords = [
    "google-drive", "excel", "xlsx", "openpyxl", "mcp",
    "model-context-protocol", "claude", "ai-agent", "spreadsheet"
]
classifiers = [
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.10",
    "Programming Language :: Python :: 3.11",
    "Programming Language :: Python :: 3.12",
    "License :: OSI Approved :: MIT License",
    "Operating System :: OS Independent",
]
dependencies = [
    "fastmcp>=3.0.0",
    "openpyxl>=3.1.0",
    "google-api-python-client>=2.0.0",
    "google-auth-oauthlib>=1.0.0",
    "google-auth-httplib2>=0.2.0",
]

[project.scripts]
xlsx-drive-mcp = "xlsx_drive_mcp.server:main"

[project.urls]
Homepage = "https://github.com/YOUR_GITHUB_USERNAME/xlsx-drive-mcp"
Issues = "https://github.com/YOUR_GITHUB_USERNAME/xlsx-drive-mcp/issues"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.hatch.build.targets.wheel]
packages = ["src/xlsx_drive_mcp"]

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-v"

[project.optional-dependencies]
dev = ["pytest>=7.0", "pytest-mock"]
```

- [ ] **Step 3: Create `LICENSE` (MIT)**

```
MIT License

Copyright (c) 2026 xlsx-drive-mcp contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

- [ ] **Step 4: Create `.gitignore`**

```
__pycache__/
*.py[cod]
*.egg-info/
dist/
build/
.venv/
venv/
.env
*.env
~/.xlsx-drive-mcp/
.pytest_cache/
```

- [ ] **Step 5: Create `src/xlsx_drive_mcp/__init__.py`**

```python
__version__ = "0.1.0"
```

- [ ] **Step 6: Create `src/xlsx_drive_mcp/tools/__init__.py`**

```python
```

- [ ] **Step 7: Install in editable mode**

```bash
cd /Users/ange/repos/xlsx-drive-mcp
pip3 install -e ".[dev]"
```

Expected: installs without errors, `xlsx-drive-mcp` command appears in PATH.

- [ ] **Step 8: Commit**

```bash
git add pyproject.toml LICENSE .gitignore src/
git commit -m "chore: initial project scaffolding"
```

---

## Task 2: Drive Client (`drive.py`)

**Files:**
- Create: `src/xlsx_drive_mcp/drive.py`
- Create: `tests/unit/test_drive.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_drive.py
import pytest
from xlsx_drive_mcp.drive import validate_cell_range, ConflictError, FileTooLargeError

def test_validate_cell_range_single_cell():
    validate_cell_range("A1")  # should not raise

def test_validate_cell_range_range():
    validate_cell_range("A1:D10")  # should not raise

def test_validate_cell_range_invalid():
    with pytest.raises(ValueError, match="Invalid cell range"):
        validate_cell_range("not_a_range")

def test_validate_cell_range_lowercase():
    validate_cell_range("a1:d10")  # should not raise (case-insensitive)

def test_validate_cell_range_too_large():
    with pytest.raises(ValueError, match="1,000,000"):
        validate_cell_range("A1:ZZZ999999")

def test_conflict_error_is_exception():
    with pytest.raises(ConflictError):
        raise ConflictError("conflict")

def test_file_too_large_error_is_exception():
    with pytest.raises(FileTooLargeError):
        raise FileTooLargeError("too large")
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd /Users/ange/repos/xlsx-drive-mcp
python -m pytest tests/unit/test_drive.py -v
```

Expected: `ImportError` or `ModuleNotFoundError` -- `drive.py` does not exist yet.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/drive.py`**

```python
import io
import logging
import re
import sys

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from openpyxl.utils import range_boundaries

# All logging to stderr; stdout is reserved for MCP JSON-RPC
logging.basicConfig(stream=sys.stderr, level=logging.WARNING)
logger = logging.getLogger(__name__)

XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024  # 50 MB
MAX_CELL_RANGE_AREA = 1_000_000
_CELL_RANGE_RE = re.compile(r"^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$")


class ConflictError(Exception):
    """Raised when Drive returns a concurrent-edit conflict."""


class FileTooLargeError(Exception):
    """Raised when the file exceeds the 50 MB limit."""


def validate_cell_range(cell_range: str) -> None:
    """Raise ValueError if cell_range is not a valid A1-notation range."""
    if not _CELL_RANGE_RE.match(cell_range):
        raise ValueError(
            f"Invalid cell range: '{cell_range}'. Expected A1-notation such as 'A1' or 'A1:D10'."
        )
    if ":" in cell_range:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range.upper())
        area = (max_col - min_col + 1) * (max_row - min_row + 1)
        if area > MAX_CELL_RANGE_AREA:
            raise ValueError(
                f"Range '{cell_range}' covers {area:,} cells; maximum is {MAX_CELL_RANGE_AREA:,}."
            )


def get_drive_service(creds: Credentials):
    """Build and return an authenticated Drive v3 service."""
    return build("drive", "v3", credentials=creds)


def get_file_metadata(service, file_id: str) -> dict:
    """Return id, name, size, mimeType, headRevisionId for a Drive file."""
    return service.files().get(
        fileId=file_id,
        fields="id,name,size,mimeType,headRevisionId",
    ).execute()


def download_xlsx(service, file_id: str) -> tuple[bytes, str]:
    """
    Download an xlsx file from Drive.

    Returns (content_bytes, revision_id).
    Raises FileTooLargeError if the file exceeds 50 MB.
    """
    metadata = get_file_metadata(service, file_id)
    size = int(metadata.get("size", 0))
    if size > MAX_FILE_SIZE_BYTES:
        raise FileTooLargeError(
            f"File is {size / 1024 / 1024:.1f} MB; maximum supported size is 50 MB."
        )
    revision_id = metadata["headRevisionId"]
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue(), revision_id


def upload_xlsx(service, file_id: str, content: bytes, revision_id: str) -> None:
    """
    Upload modified xlsx content back to Drive.

    Checks headRevisionId before uploading to detect concurrent edits.
    Raises ConflictError if the file was modified since it was downloaded.
    """
    current_metadata = get_file_metadata(service, file_id)
    if current_metadata["headRevisionId"] != revision_id:
        raise ConflictError(
            "The file was modified by another process since you downloaded it. "
            "Re-read the file and retry your operation."
        )
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=XLSX_MIME_TYPE, resumable=False)
    service.files().update(fileId=file_id, media_body=media).execute()
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_drive.py -v
```

Expected: all 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/drive.py tests/unit/test_drive.py
git commit -m "feat: add drive client with download/upload and conflict detection"
```

---

## Task 3: Auth Module (`auth.py`)

**Files:**
- Create: `src/xlsx_drive_mcp/auth.py`
- Create: `tests/unit/test_auth.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_auth.py
import json
import os
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from xlsx_drive_mcp.auth import (
    load_credentials,
    _find_credential_files,
    CredentialNotFoundError,
    WORKSPACE_MCP_CREDS_DIR,
    XLSX_DRIVE_CREDS_DIR,
)


@pytest.fixture
def cred_data():
    return {
        "token": "access_token",
        "refresh_token": "refresh_token",
        "token_uri": "https://oauth2.googleapis.com/token",
        "client_id": "client_id",
        "client_secret": "client_secret",
        "scopes": ["https://www.googleapis.com/auth/drive"],
        "expiry": "2099-01-01T00:00:00Z",
    }


def test_load_from_workspace_mcp_path(tmp_path, cred_data):
    creds_dir = tmp_path / "workspace_mcp"
    creds_dir.mkdir()
    cred_file = creds_dir / "user@example.com.json"
    cred_file.write_text(json.dumps(cred_data))
    cred_file.chmod(0o600)

    with patch("xlsx_drive_mcp.auth.WORKSPACE_MCP_CREDS_DIR", creds_dir), \
         patch("xlsx_drive_mcp.auth.XLSX_DRIVE_CREDS_DIR", tmp_path / "xlsx_drive"):
        creds, email = load_credentials("user@example.com")

    assert email == "user@example.com"
    assert creds.refresh_token == "refresh_token"


def test_auto_detect_single_credential(tmp_path, cred_data):
    creds_dir = tmp_path / "xlsx_drive"
    creds_dir.mkdir()
    cred_file = creds_dir / "solo@example.com.json"
    cred_file.write_text(json.dumps(cred_data))

    with patch("xlsx_drive_mcp.auth.WORKSPACE_MCP_CREDS_DIR", tmp_path / "workspace"),\
         patch("xlsx_drive_mcp.auth.XLSX_DRIVE_CREDS_DIR", creds_dir):
        creds, email = load_credentials(None)

    assert email == "solo@example.com"


def test_missing_credential_raises(tmp_path):
    with patch("xlsx_drive_mcp.auth.WORKSPACE_MCP_CREDS_DIR", tmp_path / "w"), \
         patch("xlsx_drive_mcp.auth.XLSX_DRIVE_CREDS_DIR", tmp_path / "x"):
        with pytest.raises(CredentialNotFoundError):
            load_credentials("nobody@example.com")


def test_ambiguous_credential_raises(tmp_path, cred_data):
    creds_dir = tmp_path / "xlsx_drive"
    creds_dir.mkdir()
    for email in ["a@example.com", "b@example.com"]:
        (creds_dir / f"{email}.json").write_text(json.dumps(cred_data))

    with patch("xlsx_drive_mcp.auth.WORKSPACE_MCP_CREDS_DIR", tmp_path / "w"), \
         patch("xlsx_drive_mcp.auth.XLSX_DRIVE_CREDS_DIR", creds_dir):
        with pytest.raises(CredentialNotFoundError, match="multiple"):
            load_credentials(None)
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_auth.py -v
```

Expected: `ImportError` -- `auth.py` does not exist yet.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/auth.py`**

```python
import json
import os
import sys
from pathlib import Path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

WORKSPACE_MCP_CREDS_DIR = Path.home() / ".google_workspace_mcp" / "credentials"
XLSX_DRIVE_CREDS_DIR = Path.home() / ".xlsx-drive-mcp" / "credentials"

STANDALONE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]


class CredentialNotFoundError(Exception):
    """Raised when no valid credentials can be found."""


def _find_credential_files() -> list[tuple[Path, str]]:
    """Return list of (path, email) for all known credential files."""
    results = []
    for creds_dir in [WORKSPACE_MCP_CREDS_DIR, XLSX_DRIVE_CREDS_DIR]:
        if creds_dir.exists():
            for f in creds_dir.glob("*.json"):
                email = f.stem
                results.append((f, email))
    return results


def _load_from_file(path: Path) -> Credentials:
    """Load and optionally refresh credentials from a JSON file."""
    with open(path) as f:
        data = json.load(f)
    creds = Credentials(
        token=data.get("token"),
        refresh_token=data["refresh_token"],
        token_uri=data["token_uri"],
        client_id=data["client_id"],
        client_secret=data["client_secret"],
        scopes=data.get("scopes"),
    )
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            raise CredentialNotFoundError(
                f"Failed to refresh credentials from {path}. "
                f"Run 'xlsx-drive-mcp auth --user {path.stem}' to re-authenticate. "
                f"Error: {e}"
            ) from e
    return creds


def load_credentials(email: str | None) -> tuple[Credentials, str]:
    """
    Resolve credentials for the given email (or auto-detect if None).

    Resolution order:
    1. ~/.google_workspace_mcp/credentials/{email}.json
    2. ~/.xlsx-drive-mcp/credentials/{email}.json
    3. Auto-detect if exactly one credential file exists across both dirs.

    Returns (Credentials, resolved_email).
    Raises CredentialNotFoundError if credentials cannot be found.
    """
    if email is not None:
        for creds_dir in [WORKSPACE_MCP_CREDS_DIR, XLSX_DRIVE_CREDS_DIR]:
            path = creds_dir / f"{email}.json"
            if path.exists():
                return _load_from_file(path), email
        raise CredentialNotFoundError(
            f"No credentials found for '{email}'. "
            f"Checked: {WORKSPACE_MCP_CREDS_DIR / f'{email}.json'} and "
            f"{XLSX_DRIVE_CREDS_DIR / f'{email}.json'}.\n"
            f"Run: xlsx-drive-mcp auth --user {email}"
        )

    all_creds = _find_credential_files()
    if len(all_creds) == 0:
        raise CredentialNotFoundError(
            "No credentials found. Run: xlsx-drive-mcp auth --user your@email.com"
        )
    if len(all_creds) > 1:
        emails = [e for _, e in all_creds]
        raise CredentialNotFoundError(
            f"Found credentials for multiple accounts: {emails}. "
            f"Specify which to use with --user."
        )
    path, resolved_email = all_creds[0]
    return _load_from_file(path), resolved_email


def run_auth_flow(email: str) -> None:
    """
    Run interactive OAuth2 browser flow and store credentials.
    Requires GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET env vars.
    """
    from google_auth_oauthlib.flow import InstalledAppFlow

    client_id = os.environ.get("GOOGLE_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
    if not client_id or not client_secret:
        print(
            "Error: GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET environment variables are required.\n"
            "Create a Desktop OAuth client at https://console.cloud.google.com/apis/credentials\n"
            "then set those variables before running this command.",
            file=sys.stderr,
        )
        sys.exit(1)

    client_config = {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    flow = InstalledAppFlow.from_client_config(client_config, scopes=STANDALONE_SCOPES)
    creds = flow.run_local_server(port=0)

    XLSX_DRIVE_CREDS_DIR.mkdir(parents=True, exist_ok=True)
    cred_path = XLSX_DRIVE_CREDS_DIR / f"{email}.json"
    cred_data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes or STANDALONE_SCOPES),
        "expiry": creds.expiry.isoformat() if creds.expiry else None,
    }
    cred_path.write_text(json.dumps(cred_data, indent=2))
    cred_path.chmod(0o600)
    print(f"Credentials saved to {cred_path}")
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_auth.py -v
```

Expected: all 4 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/auth.py tests/unit/test_auth.py
git commit -m "feat: add auth module with workspace-mcp fallback and auto-detect"
```

---

## Task 4: Server Entry Point (`server.py`)

**Files:**
- Create: `src/xlsx_drive_mcp/server.py`

- [ ] **Step 1: Create `src/xlsx_drive_mcp/server.py`**

```python
import argparse
import sys
from fastmcp import FastMCP
from .auth import load_credentials, run_auth_flow, CredentialNotFoundError
from .drive import get_drive_service


def create_app(email: str | None = None) -> FastMCP:
    """Create and return the configured FastMCP app."""
    try:
        creds, resolved_email = load_credentials(email)
    except CredentialNotFoundError as e:
        print(f"Authentication error: {e}", file=sys.stderr)
        sys.exit(1)

    service = get_drive_service(creds)
    mcp = FastMCP(
        "xlsx-drive-mcp",
        instructions=(
            "MCP server for reading and writing .xlsx files stored on Google Drive. "
            f"Authenticated as: {resolved_email}"
        ),
    )

    # Import here to avoid circular imports; importing registers tools via @mcp.tool()
    from .tools.files import register_tools as reg_files
    from .tools.sheets import register_tools as reg_sheets
    from .tools.data import register_tools as reg_data
    from .tools.format import register_tools as reg_format
    from .tools.charts import register_tools as reg_charts
    from .tools.formulas import register_tools as reg_formulas

    reg_files(mcp, service)
    reg_sheets(mcp, service)
    reg_data(mcp, service)
    reg_format(mcp, service)
    reg_charts(mcp, service)
    reg_formulas(mcp, service)

    return mcp


def main() -> None:
    parser = argparse.ArgumentParser(prog="xlsx-drive-mcp")
    parser.add_argument("--user", help="Google account email to use")

    subparsers = parser.add_subparsers(dest="command")
    auth_parser = subparsers.add_parser("auth", help="Authenticate and store credentials")
    auth_parser.add_argument("--user", required=True, help="Google account email")

    args = parser.parse_args()

    if args.command == "auth":
        run_auth_flow(args.user)
        return

    app = create_app(getattr(args, "user", None))
    app.run()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify CLI is wired up**

```bash
xlsx-drive-mcp --help
```

Expected: shows `--user` option and `auth` subcommand.

- [ ] **Step 3: Commit**

```bash
git add src/xlsx_drive_mcp/server.py
git commit -m "feat: add server entry point and CLI with auth subcommand"
```

---

## Task 5: Test Infrastructure

**Files:**
- Create: `tests/conftest.py`
- Create: `tests/fixtures/create_fixtures.py`
- Create: `tests/fixtures/*.xlsx`

- [ ] **Step 1: Create `tests/conftest.py`**

```python
import io
import pytest
import openpyxl
from unittest.mock import MagicMock


@pytest.fixture
def simple_wb_bytes() -> bytes:
    """Minimal xlsx: Sheet1 with A1=hello, B1=world, A2=42, B2=3.14."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello"
    ws["B1"] = "world"
    ws["A2"] = 42
    ws["B2"] = 3.14
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def mock_service():
    """Minimal Drive service mock. Tests configure return values as needed."""
    return MagicMock()


@pytest.fixture
def drive_store(simple_wb_bytes):
    """
    In-memory dict simulating Drive: {file_id: (content_bytes, revision_id)}.
    Returns the store dict; tests can add more files.
    """
    return {"file1": (simple_wb_bytes, "rev1")}


def make_download_fn(store: dict):
    from xlsx_drive_mcp.drive import FileTooLargeError, MAX_FILE_SIZE_BYTES

    def download(service, file_id: str) -> tuple[bytes, str]:
        if file_id not in store:
            raise KeyError(f"File {file_id!r} not in mock store")
        content, rev = store[file_id]
        if len(content) > MAX_FILE_SIZE_BYTES:
            raise FileTooLargeError("File too large")
        return content, rev

    return download


def make_upload_fn(store: dict):
    from xlsx_drive_mcp.drive import ConflictError

    def upload(service, file_id: str, content: bytes, revision_id: str) -> None:
        if file_id not in store:
            raise KeyError(f"File {file_id!r} not in mock store")
        _, current_rev = store[file_id]
        if current_rev != revision_id:
            raise ConflictError("Conflict")
        new_rev = f"rev{len(store[file_id][0]) + len(content)}"
        store[file_id] = (content, new_rev)

    return upload
```

- [ ] **Step 2: Create fixture xlsx files**

```python
# tests/fixtures/create_fixtures.py
"""Run once to generate fixture xlsx files: python tests/fixtures/create_fixtures.py"""
import io
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side

FIXTURES = Path(__file__).parent


def make_merged_cells():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Merged Header"
    ws.merge_cells("A1:C1")
    ws["A2"] = "Under merge"
    wb.save(FIXTURES / "merged_cells.xlsx")


def make_named_ranges():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for i in range(1, 6):
        ws[f"A{i}"] = i * 10
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["MyRange"] = DefinedName("MyRange", attr_text="Sheet1!$A$1:$A$5")
    wb.save(FIXTURES / "named_ranges.xlsx")


def make_formulas():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = "=SUM(A1:A2)"
    wb.save(FIXTURES / "formulas.xlsx")


def make_multi_sheet():
    wb = openpyxl.Workbook()
    wb.active.title = "Alpha"
    wb.active["A1"] = "alpha data"
    wb.create_sheet("Beta")
    wb["Beta"]["A1"] = "beta data"
    wb.create_sheet("Gamma")
    wb["Gamma"]["A1"] = "gamma data"
    wb.save(FIXTURES / "multi_sheet.xlsx")


def make_formatting():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Bold Red"
    ws["A1"].font = Font(bold=True, color="FFFF0000")
    ws["B1"] = "Yellow BG"
    ws["B1"].fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    ws["C1"] = "Bordered"
    thin = Side(style="thin")
    ws["C1"].border = Border(top=thin, bottom=thin, left=thin, right=thin)
    wb.save(FIXTURES / "formatting.xlsx")


if __name__ == "__main__":
    make_merged_cells()
    make_named_ranges()
    make_formulas()
    make_multi_sheet()
    make_formatting()
    print("Fixtures created.")
```

- [ ] **Step 3: Run the fixture creation script**

```bash
cd /Users/ange/repos/xlsx-drive-mcp
python tests/fixtures/create_fixtures.py
```

Expected: `Fixtures created.` and 5 xlsx files appear in `tests/fixtures/`.

- [ ] **Step 4: Commit**

```bash
git add tests/conftest.py tests/fixtures/
git commit -m "test: add test infrastructure and fixture xlsx files"
```

---

## Task 6: `tools/files.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/files.py`
- Create: `tests/unit/test_files.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_files.py
import io
import openpyxl
import pytest
from unittest.mock import MagicMock, patch
from xlsx_drive_mcp.tools.files import (
    _build_list_query,
    _get_xlsx_info_from_wb,
    _scan_sheet_dimensions,
)


def test_build_list_query_no_filters():
    q = _build_list_query(None, None)
    assert "spreadsheetml" in q
    assert "name contains" not in q


def test_build_list_query_with_name():
    q = _build_list_query(None, "budget")
    assert "name contains 'budget'" in q


def test_build_list_query_escapes_single_quote():
    q = _build_list_query(None, "Q1'report")
    assert "Q1\\'report" in q


def test_build_list_query_with_folder():
    q = _build_list_query("folder123", None)
    assert "'folder123' in parents" in q


def test_get_xlsx_info_from_wb():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "x"
    ws["C3"] = "y"
    result = _get_xlsx_info_from_wb(wb)
    assert result["sheets"][0]["name"] == "Data"
    assert result["sheets"][0]["rows"] == 3
    assert result["sheets"][0]["cols"] == 3


def test_scan_sheet_dimensions_empty():
    wb = openpyxl.Workbook()
    rows, cols = _scan_sheet_dimensions(wb.active)
    assert rows == 0
    assert cols == 0


def test_scan_sheet_dimensions_after_clear():
    """max_row is unreliable after clear; scan must find actual data."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A5"] = "data"
    ws["A5"].value = None  # simulate clear
    rows, cols = _scan_sheet_dimensions(ws)
    assert rows == 0
    assert cols == 0
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_files.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/files.py`**

```python
import io
import re
from typing import Optional

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource

from ..drive import (
    XLSX_MIME_TYPE,
    download_xlsx,
    get_file_metadata,
    upload_xlsx,
    validate_cell_range,
)

_SINGLE_QUOTE_RE = re.compile(r"'")


def _build_list_query(folder_id: Optional[str], name_contains: Optional[str]) -> str:
    parts = [f"mimeType='{XLSX_MIME_TYPE}'", "trashed=false"]
    if folder_id:
        parts.append(f"'{folder_id}' in parents")
    if name_contains:
        escaped = _SINGLE_QUOTE_RE.sub("\\'", name_contains)
        parts.append(f"name contains '{escaped}'")
    return " and ".join(parts)


def _scan_sheet_dimensions(ws) -> tuple[int, int]:
    """Return (last_data_row, last_data_col) by scanning actual cell values."""
    max_row = 0
    max_col = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
    return max_row, max_col


def _get_xlsx_info_from_wb(wb: openpyxl.Workbook) -> dict:
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        rows, cols = _scan_sheet_dimensions(ws)
        sheets.append({"name": name, "rows": rows, "cols": cols})
    named_ranges = list(wb.defined_names.keys())
    return {"sheets": sheets, "named_ranges": named_ranges}


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def list_xlsx_files(
        folder_id: Optional[str] = None,
        name_contains: Optional[str] = None,
        page_token: Optional[str] = None,
    ) -> dict:
        """List xlsx files on Google Drive.

        Args:
            folder_id: Restrict search to this Drive folder ID.
            name_contains: Filter by filename substring (safe, not raw query).
            page_token: Token from a previous response to fetch the next page.

        Returns dict with 'files' list and optional 'next_page_token'.
        """
        q = _build_list_query(folder_id, name_contains)
        params = dict(
            q=q,
            fields="nextPageToken,files(id,name,modifiedTime,size)",
            pageSize=100,
        )
        if page_token:
            params["pageToken"] = page_token
        result = service.files().list(**params).execute()
        files = [
            {
                "file_id": f["id"],
                "name": f["name"],
                "modified_time": f.get("modifiedTime"),
                "size_bytes": int(f.get("size", 0)),
            }
            for f in result.get("files", [])
        ]
        out = {"files": files}
        if "nextPageToken" in result:
            out["next_page_token"] = result["nextPageToken"]
        return out

    @mcp.tool()
    def get_xlsx_info(file_id: str) -> dict:
        """Get sheet names, data dimensions, and named ranges for an xlsx file.

        Dimensions are determined by scanning actual data, not openpyxl's
        max_row/max_column (which are unreliable after cells are cleared).
        """
        content, _ = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        return _get_xlsx_info_from_wb(wb)

    @mcp.tool()
    def create_xlsx(name: str, folder_id: Optional[str] = None) -> dict:
        """Create a new blank xlsx file on Google Drive. Returns the file_id."""
        wb = openpyxl.Workbook()
        buf = io.BytesIO()
        wb.save(buf)
        metadata = {"name": name, "mimeType": XLSX_MIME_TYPE}
        if folder_id:
            metadata["parents"] = [folder_id]
        from googleapiclient.http import MediaIoBaseUpload
        media = MediaIoBaseUpload(io.BytesIO(buf.getvalue()), mimetype=XLSX_MIME_TYPE)
        result = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name",
        ).execute()
        return {"file_id": result["id"], "name": result["name"]}

    @mcp.tool()
    def delete_xlsx(file_id: str) -> dict:
        """Move an xlsx file to Drive trash.

        WARNING: This is non-reversible from the agent's perspective.
        The user has 30 days to recover the file from the Drive trash UI.
        Use with caution.
        """
        service.files().trash(fileId=file_id).execute()
        return {"deleted": True, "file_id": file_id}
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_files.py -v
```

Expected: all 7 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/tools/files.py tests/unit/test_files.py
git commit -m "feat: add files tools (list, info, create, delete)"
```

---

## Task 7: `tools/sheets.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/sheets.py`
- Create: `tests/unit/test_sheets.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_sheets.py
import io
import openpyxl
import pytest
from xlsx_drive_mcp.tools.sheets import (
    _list_sheets,
    _create_sheet,
    _delete_sheet,
    _rename_sheet,
    _copy_sheet,
)


@pytest.fixture
def multi_wb():
    wb = openpyxl.Workbook()
    wb.active.title = "Alpha"
    wb.create_sheet("Beta")
    return wb


def test_list_sheets(multi_wb):
    result = _list_sheets(multi_wb)
    assert result == [{"name": "Alpha", "index": 0}, {"name": "Beta", "index": 1}]


def test_create_sheet_at_end(multi_wb):
    result = _create_sheet(multi_wb, "Gamma", None)
    assert result["name"] == "Gamma"
    assert result["index"] == 2
    assert "Gamma" in multi_wb.sheetnames


def test_create_sheet_at_position(multi_wb):
    _create_sheet(multi_wb, "First", 0)
    assert multi_wb.sheetnames[0] == "First"


def test_delete_sheet(multi_wb):
    _delete_sheet(multi_wb, "Beta")
    assert "Beta" not in multi_wb.sheetnames


def test_delete_last_sheet_raises(multi_wb):
    _delete_sheet(multi_wb, "Beta")
    with pytest.raises(ValueError, match="only sheet"):
        _delete_sheet(multi_wb, "Alpha")


def test_rename_sheet(multi_wb):
    _rename_sheet(multi_wb, "Beta", "Delta")
    assert "Delta" in multi_wb.sheetnames
    assert "Beta" not in multi_wb.sheetnames


def test_copy_sheet(multi_wb):
    result = _copy_sheet(multi_wb, "Alpha", "AlphaCopy")
    assert result["name"] == "AlphaCopy"
    assert "AlphaCopy" in multi_wb.sheetnames


def test_copy_sheet_nonexistent_raises(multi_wb):
    with pytest.raises(KeyError):
        _copy_sheet(multi_wb, "NoSuchSheet", "Copy")
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_sheets.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/sheets.py`**

```python
import io
from typing import Optional

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource

from ..drive import download_xlsx, upload_xlsx


def _list_sheets(wb: openpyxl.Workbook) -> list[dict]:
    return [{"name": name, "index": i} for i, name in enumerate(wb.sheetnames)]


def _create_sheet(wb: openpyxl.Workbook, name: str, position: Optional[int]) -> dict:
    ws = wb.create_sheet(title=name, index=position)
    return {"name": ws.title, "index": wb.sheetnames.index(ws.title)}


def _delete_sheet(wb: openpyxl.Workbook, sheet_name: str) -> None:
    if len(wb.sheetnames) == 1:
        raise ValueError(f"Cannot delete the only sheet in the workbook.")
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    del wb[sheet_name]


def _rename_sheet(wb: openpyxl.Workbook, sheet_name: str, new_name: str) -> None:
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    wb[sheet_name].title = new_name


def _copy_sheet(wb: openpyxl.Workbook, sheet_name: str, new_name: str) -> dict:
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    new_ws = wb.copy_worksheet(wb[sheet_name])
    new_ws.title = new_name
    return {"name": new_ws.title, "index": wb.sheetnames.index(new_ws.title)}


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def list_sheets(file_id: str) -> list[dict]:
        """List all sheets in an xlsx file. Returns [{name, index}]."""
        content, _ = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        return _list_sheets(wb)

    @mcp.tool()
    def create_sheet(file_id: str, name: str, position: Optional[int] = None) -> dict:
        """Add a new sheet to an xlsx file. Position is 0-based; defaults to last."""
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        result = _create_sheet(wb, name, position)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return result

    @mcp.tool()
    def delete_sheet(file_id: str, sheet_name: str) -> dict:
        """Remove a sheet from an xlsx file. Raises if it is the only sheet."""
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        _delete_sheet(wb, sheet_name)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"deleted": sheet_name}

    @mcp.tool()
    def rename_sheet(file_id: str, sheet_name: str, new_name: str) -> dict:
        """Rename a sheet in an xlsx file."""
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        _rename_sheet(wb, sheet_name, new_name)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"old_name": sheet_name, "new_name": new_name}

    @mcp.tool()
    def copy_sheet(file_id: str, sheet_name: str, new_name: str) -> dict:
        """Duplicate a sheet within the same workbook.

        WARNING: openpyxl's copy_worksheet silently drops charts, images,
        and some merged cell configurations from the copied sheet.
        """
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        result = _copy_sheet(wb, sheet_name, new_name)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return result
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_sheets.py -v
```

Expected: all 9 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/tools/sheets.py tests/unit/test_sheets.py
git commit -m "feat: add sheet management tools (list, create, delete, rename, copy)"
```

---

## Task 8: `tools/data.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/data.py`
- Create: `tests/unit/test_data.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_data.py
import io
import openpyxl
import pytest
from xlsx_drive_mcp.tools.data import (
    _read_range,
    _write_range,
    _append_rows,
    _clear_range,
    _find_last_data_row,
)


@pytest.fixture
def wb():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello"
    ws["B1"] = "world"
    ws["A2"] = 42
    ws["B2"] = 3.14
    return wb


def test_read_range_basic(wb):
    result = _read_range(wb, "Sheet1", "A1:B2")
    assert result["values"] == [["hello", "world"], [42, 3.14]]
    assert result["cell_range"] == "A1:B2"


def test_read_range_single_cell(wb):
    result = _read_range(wb, "Sheet1", "A1")
    assert result["values"] == [["hello"]]


def test_read_range_unknown_sheet_raises(wb):
    with pytest.raises(KeyError):
        _read_range(wb, "NoSheet", "A1")


def test_write_range_basic(wb):
    _write_range(wb, "Sheet1", "A1", [["new_a", "new_b"], [99, 100]])
    assert wb["Sheet1"]["A1"].value == "new_a"
    assert wb["Sheet1"]["B2"].value == 100


def test_write_range_formula(wb):
    _write_range(wb, "Sheet1", "C1", [["=SUM(A2:B2)"]])
    assert wb["Sheet1"]["C1"].value == "=SUM(A2:B2)"


def test_write_range_does_not_clear_extra_cells(wb):
    """Cells beyond the values array are untouched."""
    _write_range(wb, "Sheet1", "A1", [["only_a1"]])
    assert wb["Sheet1"]["B1"].value == "world"  # unchanged


def test_find_last_data_row_basic(wb):
    assert _find_last_data_row(wb["Sheet1"]) == 2


def test_find_last_data_row_after_clear(wb):
    wb["Sheet1"]["A2"].value = None
    wb["Sheet1"]["B2"].value = None
    assert _find_last_data_row(wb["Sheet1"]) == 1


def test_find_last_data_row_empty():
    wb2 = openpyxl.Workbook()
    assert _find_last_data_row(wb2.active) == 0


def test_append_rows(wb):
    _append_rows(wb, "Sheet1", [["new1", "new2"]])
    assert wb["Sheet1"]["A3"].value == "new1"
    assert wb["Sheet1"]["B3"].value == "new2"


def test_clear_range_preserves_formatting(wb):
    from openpyxl.styles import Font
    wb["Sheet1"]["A1"].font = Font(bold=True)
    _clear_range(wb, "Sheet1", "A1:B1")
    assert wb["Sheet1"]["A1"].value is None
    assert wb["Sheet1"]["A1"].font.bold is True


def test_clear_range_over_merged_raises():
    wb2 = openpyxl.Workbook()
    ws = wb2.active
    ws.title = "Sheet1"
    ws.merge_cells("A1:C1")
    with pytest.raises(ValueError, match="merged"):
        _clear_range(wb2, "Sheet1", "A1:C3")
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_data.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/data.py`**

```python
import io
from typing import Optional

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource
from openpyxl.utils import range_boundaries, get_column_letter

from ..drive import download_xlsx, upload_xlsx, validate_cell_range


def _read_range(wb: openpyxl.Workbook, sheet_name: str, cell_range: str) -> dict:
    ws = wb[sheet_name]  # raises KeyError if not found
    if ":" in cell_range:
        rows = list(ws[cell_range.upper()])
        values = [[cell.value for cell in row] for row in rows]
    else:
        cell = ws[cell_range.upper()]
        values = [[cell.value]]
    return {"values": values, "cell_range": cell_range}


def _write_range(wb: openpyxl.Workbook, sheet_name: str, cell_range: str, values: list) -> None:
    ws = wb[sheet_name]
    # Determine anchor from top-left of cell_range
    anchor = cell_range.upper().split(":")[0]
    min_col, min_row, _, _ = range_boundaries(anchor + ":" + anchor)
    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            ws.cell(row=min_row + r_offset, column=min_col + c_offset, value=value)


def _find_last_data_row(ws) -> int:
    """Scan from bottom to find the last row that has at least one non-None value."""
    last_row = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                last_row = max(last_row, cell.row)
    return last_row


def _append_rows(wb: openpyxl.Workbook, sheet_name: str, rows: list) -> None:
    ws = wb[sheet_name]
    last_row = _find_last_data_row(ws)
    for r_offset, row_values in enumerate(rows):
        for c_offset, value in enumerate(row_values):
            ws.cell(row=last_row + 1 + r_offset, column=1 + c_offset, value=value)


def _clear_range(wb: openpyxl.Workbook, sheet_name: str, cell_range: str) -> None:
    ws = wb[sheet_name]
    merged = {str(r) for r in ws.merged_cells.ranges}
    min_col, min_row, max_col, max_row = range_boundaries(cell_range.upper())
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            addr = f"{get_column_letter(c)}{r}"
            for merge_range in ws.merged_cells.ranges:
                if ws[addr].coordinate in [
                    ws.cell(row=mr, column=mc).coordinate
                    for mr in range(merge_range.min_row, merge_range.max_row + 1)
                    for mc in range(merge_range.min_col, merge_range.max_col + 1)
                ]:
                    raise ValueError(
                        f"Cell {addr} is part of a merged region {merge_range}. "
                        f"Call unmerge_cells first."
                    )
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def read_range(
        file_id: str,
        sheet_name: str,
        cell_range: str,
        data_only: bool = True,
    ) -> dict:
        """Read cell values from a range in an xlsx file on Google Drive.

        When data_only=True (default), formula cells return their last cached computed
        value from when Excel/Sheets last saved the file. When data_only=False, formula
        strings are returned instead. These modes are mutually exclusive in openpyxl.

        Args:
            cell_range: A1-notation range such as 'A1' or 'A1:D10'.
        """
        validate_cell_range(cell_range)
        content, _ = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=data_only)
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        return _read_range(wb, sheet_name, cell_range)

    @mcp.tool()
    def write_range(
        file_id: str,
        sheet_name: str,
        cell_range: str,
        values: list,
    ) -> dict:
        """Write a 2D array of values to an xlsx file on Google Drive.

        cell_range sets the top-left anchor. The values array determines the write
        extent. Cells beyond the values array are not cleared. Formula strings
        (starting with '=') are written as formulas.

        WARNING: Writing over cells that contain formulas replaces them with
        the static values provided.
        """
        validate_cell_range(cell_range)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        _write_range(wb, sheet_name, cell_range, values)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"written": True, "cell_range": cell_range, "rows": len(values)}

    @mcp.tool()
    def append_rows(file_id: str, sheet_name: str, rows: list) -> dict:
        """Append rows below the last non-empty row in a sheet.

        Uses an actual data scan to find the last row -- not openpyxl's max_row
        property, which is unreliable after cells have been cleared.
        """
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        _append_rows(wb, sheet_name, rows)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"appended_rows": len(rows)}

    @mcp.tool()
    def clear_range(file_id: str, sheet_name: str, cell_range: str) -> dict:
        """Clear cell values in a range, preserving formatting.

        Raises an error if the range intersects a merged region (call
        unmerge_cells first).
        """
        validate_cell_range(cell_range)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        _clear_range(wb, sheet_name, cell_range)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"cleared": True, "cell_range": cell_range}
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_data.py -v
```

Expected: all 12 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/tools/data.py tests/unit/test_data.py
git commit -m "feat: add data tools (read_range, write_range, append_rows, clear_range)"
```

---

## Task 9: `tools/format.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/format.py`
- Create: `tests/unit/test_format.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_format.py
import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill, Border, Side
from xlsx_drive_mcp.tools.format import (
    _read_cell_format,
    _apply_format_to_range,
    _parse_color_input,
    _parse_border_input,
    _expand_columns,
    _expand_rows,
)


@pytest.fixture
def wb():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].font = Font(bold=True, color="FFFF0000", size=14)
    ws["B1"].fill = PatternFill(fill_type="solid", fgColor="FF00FF00")
    thin = Side(style="thin")
    ws["C1"].border = Border(top=thin, right=thin)
    return wb


def test_read_cell_format_bold(wb):
    result = _read_cell_format(wb["Sheet1"]["A1"], wb["Sheet1"])
    assert result["bold"] is True
    assert result["font_color"] == "FFFF0000"
    assert result["font_size"] == 14
    assert result["merged"] is False


def test_read_cell_format_bg_color(wb):
    result = _read_cell_format(wb["Sheet1"]["B1"], wb["Sheet1"])
    assert result["bg_color"] == "FF00FF00"


def test_read_cell_format_border(wb):
    result = _read_cell_format(wb["Sheet1"]["C1"], wb["Sheet1"])
    assert result["border"]["top"] == "thin"
    assert result["border"]["right"] == "thin"
    assert result["border"]["left"] is None


def test_parse_color_six_digit():
    assert _parse_color_input("FF0000") == "FFFF0000"


def test_parse_color_eight_digit():
    assert _parse_color_input("FFFF0000") == "FFFF0000"


def test_parse_color_invalid():
    with pytest.raises(ValueError):
        _parse_color_input("ZZZZZZ")


def test_parse_border_string_shorthand():
    result = _parse_border_input("thin")
    assert result == {"top": "thin", "bottom": "thin", "left": "thin", "right": "thin"}


def test_parse_border_dict():
    result = _parse_border_input({"top": "medium", "bottom": "none"})
    assert result["top"] == "medium"
    assert result["bottom"] == "none"
    assert result.get("left") is None


def test_expand_columns_letter():
    assert _expand_columns("A") == ["A"]


def test_expand_columns_list():
    assert _expand_columns(["A", "C"]) == ["A", "C"]


def test_expand_columns_range_string():
    assert _expand_columns("A:C") == ["A", "B", "C"]


def test_expand_rows_single():
    assert _expand_rows(3) == [3]


def test_expand_rows_list():
    assert _expand_rows([1, 3, 5]) == [1, 3, 5]


def test_expand_rows_range_string():
    assert _expand_rows("2:4") == [2, 3, 4]


def test_apply_format_bold(wb):
    _apply_format_to_range(wb["Sheet1"], "A2:B2", bold=True)
    assert wb["Sheet1"]["A2"].font.bold is True
    assert wb["Sheet1"]["B2"].font.bold is True
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_format.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/format.py`**

```python
import io
import re
from copy import copy
from typing import Optional, Union

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries

from ..drive import download_xlsx, upload_xlsx, validate_cell_range

_HEX6_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_HEX8_RE = re.compile(r"^[0-9A-Fa-f]{8}$")


def _parse_color_input(color: str) -> str:
    """Normalize color to 8-digit ARGB hex string."""
    if _HEX8_RE.match(color):
        return color.upper()
    if _HEX6_RE.match(color):
        return "FF" + color.upper()
    raise ValueError(f"Invalid color '{color}'. Use RRGGBB or FFRRGGBB hex format.")


def _parse_border_input(border) -> dict:
    """Normalize border param to {top, bottom, left, right} dict."""
    if isinstance(border, str):
        return {"top": border, "bottom": border, "left": border, "right": border}
    return border  # already a dict


def _expand_columns(columns) -> list[str]:
    """Expand 'A', ['A','B'], or 'A:C' to list of column letters."""
    if isinstance(columns, str) and ":" in columns:
        start, end = columns.split(":")
        s = column_index_from_string(start.upper())
        e = column_index_from_string(end.upper())
        return [get_column_letter(i) for i in range(s, e + 1)]
    if isinstance(columns, str):
        return [columns.upper()]
    return [c.upper() for c in columns]


def _expand_rows(rows) -> list[int]:
    """Expand 3, [1,2,3], or '2:4' to list of row ints."""
    if isinstance(rows, int):
        return [rows]
    if isinstance(rows, str) and ":" in rows:
        start, end = rows.split(":")
        return list(range(int(start), int(end) + 1))
    return list(rows)


def _get_color_str(color_obj) -> Optional[str]:
    """Extract ARGB hex string from an openpyxl Color object, or None."""
    if color_obj is None:
        return None
    if hasattr(color_obj, "type"):
        if color_obj.type == "rgb" and color_obj.rgb not in ("00000000", "FFFFFFFF", "FF000000") or color_obj.rgb:
            return color_obj.rgb if color_obj.rgb != "00000000" else None
    return None


def _read_cell_format(cell, ws) -> dict:
    """Extract formatting dict from a single openpyxl cell."""
    font = cell.font or Font()
    fill = cell.fill
    border = cell.border or Border()

    bg_color = None
    if fill and fill.fill_type == "solid" and fill.fgColor:
        raw = fill.fgColor.rgb if hasattr(fill.fgColor, "rgb") else None
        if raw and raw not in ("00000000",):
            bg_color = raw

    font_color = None
    if font.color and hasattr(font.color, "rgb"):
        raw = font.color.rgb
        if raw and raw not in ("FF000000", "00000000"):
            font_color = raw
        elif raw == "FF000000":
            font_color = "FF000000"

    def side_style(side):
        return side.border_style if side and side.border_style else None

    merge_range = None
    merged = False
    for merge in ws.merged_cells.ranges:
        if (merge.min_row <= cell.row <= merge.max_row and
                merge.min_col <= cell.column <= merge.max_col):
            merged = True
            if cell.row == merge.min_row and cell.column == merge.min_col:
                merge_range = str(merge)
            break

    return {
        "address": cell.coordinate,
        "bold": font.bold or False,
        "italic": font.italic or False,
        "font_size": font.size,
        "font_color": font_color,
        "bg_color": bg_color,
        "number_format": cell.number_format,
        "border": {
            "top": side_style(border.top),
            "bottom": side_style(border.bottom),
            "left": side_style(border.left),
            "right": side_style(border.right),
        },
        "merged": merged,
        "merge_range": merge_range,
    }


def _apply_format_to_range(ws, cell_range: str, **kwargs) -> None:
    """Apply formatting kwargs to all cells in cell_range."""
    min_col, min_row, max_col, max_row = range_boundaries(cell_range.upper())
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if "bold" in kwargs or "italic" in kwargs or "font_size" in kwargs or "font_color" in kwargs:
                existing = copy(cell.font) if cell.font else Font()
                cell.font = Font(
                    bold=kwargs.get("bold", existing.bold),
                    italic=kwargs.get("italic", existing.italic),
                    size=kwargs.get("font_size", existing.size),
                    color=_parse_color_input(kwargs["font_color"]) if kwargs.get("font_color") else existing.color,
                )
            if "bg_color" in kwargs:
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor=_parse_color_input(kwargs["bg_color"]),
                )
            if "number_format" in kwargs:
                cell.number_format = kwargs["number_format"]
            if "border" in kwargs:
                parsed = _parse_border_input(kwargs["border"])

                def make_side(style_val):
                    if style_val == "none":
                        return Side(style=None)
                    if style_val is None:
                        return copy(getattr(cell.border, "top", Side()))  # unchanged handled per-side
                    return Side(style=style_val)

                existing_border = copy(cell.border) if cell.border else Border()
                cell.border = Border(
                    top=make_side(parsed.get("top")) if "top" in parsed else existing_border.top,
                    bottom=make_side(parsed.get("bottom")) if "bottom" in parsed else existing_border.bottom,
                    left=make_side(parsed.get("left")) if "left" in parsed else existing_border.left,
                    right=make_side(parsed.get("right")) if "right" in parsed else existing_border.right,
                )


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def read_format(file_id: str, sheet_name: str, cell_range: str) -> dict:
        """Read formatting (bold, colors, borders, merge status) for a cell range.

        Colors are returned as 8-digit ARGB hex strings (e.g. 'FFFF0000' = opaque red).
        Transparent/unset colors are returned as null.
        """
        validate_cell_range(cell_range)
        content, _ = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        ws = wb[sheet_name]
        cells_fmt = []
        if ":" in cell_range:
            for row in ws[cell_range.upper()]:
                for cell in row:
                    cells_fmt.append(_read_cell_format(cell, ws))
        else:
            cells_fmt.append(_read_cell_format(ws[cell_range.upper()], ws))
        return {"cells": cells_fmt}

    @mcp.tool()
    def format_range(
        file_id: str,
        sheet_name: str,
        cell_range: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        font_size: Optional[float] = None,
        font_color: Optional[str] = None,
        bg_color: Optional[str] = None,
        number_format: Optional[str] = None,
        border: Optional[Union[str, dict]] = None,
    ) -> dict:
        """Apply formatting to a cell range. Omitted params are left unchanged.

        border: 'thin'|'medium'|'thick'|'none' (all sides), or
                {"top": "thin", "bottom": "none", ...} for per-side control.
                null means leave unchanged; 'none' removes the border.
        Colors: 'RRGGBB' or 'FFRRGGBB' hex format.
        """
        validate_cell_range(cell_range)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        kwargs = {k: v for k, v in {
            "bold": bold, "italic": italic, "font_size": font_size,
            "font_color": font_color, "bg_color": bg_color,
            "number_format": number_format, "border": border,
        }.items() if v is not None}
        _apply_format_to_range(wb[sheet_name], cell_range, **kwargs)
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"formatted": True, "cell_range": cell_range}

    @mcp.tool()
    def set_column_width(
        file_id: str, sheet_name: str, columns: Union[str, list], width: float
    ) -> dict:
        """Set column width. columns: 'A', ['A','B'], or 'A:C'."""
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        ws = wb[sheet_name]
        for col in _expand_columns(columns):
            ws.column_dimensions[col].width = width
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"set_columns": _expand_columns(columns), "width": width}

    @mcp.tool()
    def set_row_height(
        file_id: str, sheet_name: str, rows: Union[int, str, list], height: float
    ) -> dict:
        """Set row height. rows: 3, [1,2,3], or '2:4'."""
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        ws = wb[sheet_name]
        for row in _expand_rows(rows):
            ws.row_dimensions[row].height = height
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"set_rows": _expand_rows(rows), "height": height}

    @mcp.tool()
    def merge_cells(file_id: str, sheet_name: str, cell_range: str) -> dict:
        """Merge a cell range."""
        validate_cell_range(cell_range)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        wb[sheet_name].merge_cells(cell_range.upper())
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"merged": cell_range}

    @mcp.tool()
    def unmerge_cells(file_id: str, sheet_name: str, cell_range: str) -> dict:
        """Unmerge a previously merged cell range."""
        validate_cell_range(cell_range)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        wb[sheet_name].unmerge_cells(cell_range.upper())
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"unmerged": cell_range}
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_format.py -v
```

Expected: all 15 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/tools/format.py tests/unit/test_format.py
git commit -m "feat: add format tools (read_format, format_range, widths, heights, merge)"
```

---

## Task 10: `tools/charts.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/charts.py`
- Create: `tests/unit/test_charts.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_charts.py
import openpyxl
import pytest
from xlsx_drive_mcp.tools.charts import _build_chart, _place_chart, CHART_TYPES


@pytest.fixture
def data_wb():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Month"
    ws["B1"] = "Sales"
    ws["C1"] = "Costs"
    for i, (month, sales, costs) in enumerate([
        ("Jan", 100, 80), ("Feb", 120, 90), ("Mar", 110, 85)
    ], start=2):
        ws[f"A{i}"] = month
        ws[f"B{i}"] = sales
        ws[f"C{i}"] = costs
    return wb


def test_build_bar_col_chart(data_wb):
    chart = _build_chart(data_wb["Sheet1"], "bar", "col", "B2:C4", "A2:A4", ["Sales", "Costs"], "Revenue")
    assert chart is not None
    assert chart.title == "Revenue"


def test_build_line_chart(data_wb):
    chart = _build_chart(data_wb["Sheet1"], "line", "line", "B2:B4", None, None, None)
    assert chart is not None


def test_build_pie_chart(data_wb):
    chart = _build_chart(data_wb["Sheet1"], "pie", "pie", "B2:B4", "A2:A4", None, "Pie")
    assert chart is not None


def test_build_scatter_chart(data_wb):
    chart = _build_chart(data_wb["Sheet1"], "scatter", "marker", "B2:C4", None, None, None)
    assert chart is not None


def test_invalid_chart_type_raises(data_wb):
    with pytest.raises(ValueError, match="chart_type"):
        _build_chart(data_wb["Sheet1"], "funnel", "col", "B2:C4", None, None, None)


def test_place_chart(data_wb):
    chart = _build_chart(data_wb["Sheet1"], "bar", "col", "B2:C4", "A2:A4", None, None)
    _place_chart(data_wb["Sheet1"], chart, "E2")
    assert len(data_wb["Sheet1"]._charts) == 1
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_charts.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/charts.py`**

```python
import io
from typing import Optional

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference, Series
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

from ..drive import download_xlsx, upload_xlsx, validate_cell_range

CHART_TYPES = {"bar", "line", "pie", "scatter"}

_CHART_SUBTYPES = {
    "bar": {"col", "bar"},
    "line": {"line", "smooth"},
    "pie": {"pie"},
    "scatter": {"marker", "smooth_line", "straight_line"},
}


def _build_chart(ws, chart_type: str, chart_subtype: str, values_range: str,
                 category_range: Optional[str], series_names: Optional[list],
                 title: Optional[str]):
    if chart_type not in CHART_TYPES:
        raise ValueError(f"Invalid chart_type '{chart_type}'. Must be one of: {sorted(CHART_TYPES)}")

    min_col, min_row, max_col, max_row = range_boundaries(values_range.upper())

    if chart_type == "bar":
        chart = BarChart()
        chart.type = chart_subtype if chart_subtype in {"col", "bar"} else "col"
    elif chart_type == "line":
        chart = LineChart()
    elif chart_type == "pie":
        chart = PieChart()
    elif chart_type == "scatter":
        chart = ScatterChart()

    if title:
        chart.title = title

    # Add data series -- each column in values_range is one series
    for col_idx in range(min_col, max_col + 1):
        values_ref = Reference(ws, min_col=col_idx, min_row=min_row, max_row=max_row)
        series = Series(values_ref)
        if series_names and isinstance(series_names, list):
            s_idx = col_idx - min_col
            if s_idx < len(series_names):
                series.title = series_names[s_idx]
        elif series_names and isinstance(series_names, str):
            # series_names is a cell range like "B1:D1"
            sc, sr, ec, er = range_boundaries(series_names.upper())
            name_ref = Reference(ws, min_col=sc + (col_idx - min_col), min_row=sr)
            series.title = name_ref
        chart.series.append(series)

    if category_range:
        cat_min_col, cat_min_row, _, cat_max_row = range_boundaries(category_range.upper())
        categories = Reference(ws, min_col=cat_min_col, min_row=cat_min_row, max_row=cat_max_row)
        chart.set_categories(categories)

    return chart


def _place_chart(ws, chart, anchor_cell: str) -> None:
    ws.add_chart(chart, anchor_cell.upper())


def _default_anchor(values_range: str) -> str:
    """Place chart one column to the right of the data range."""
    _, _, max_col, min_row = range_boundaries(values_range.upper())
    return f"{get_column_letter(max_col + 2)}{min_row}"


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def create_chart(
        file_id: str,
        sheet_name: str,
        chart_type: str,
        values_range: str,
        chart_subtype: Optional[str] = None,
        category_range: Optional[str] = None,
        series_names: Optional[list] = None,
        title: Optional[str] = None,
        anchor_cell: Optional[str] = None,
    ) -> dict:
        """Create a chart in an xlsx file on Google Drive.

        chart_type: 'bar', 'line', 'pie', 'scatter'
        chart_subtype: bar='col'(default)/'bar'; line='line'(default)/'smooth';
                       pie='pie'; scatter='marker'(default)/'smooth_line'/'straight_line'
        values_range: cell range of data values, e.g. 'B2:D10'. Columns = series.
        category_range: cell range of X-axis/category labels, e.g. 'A2:A10'.
        series_names: list of series label strings, or a cell range like 'B1:D1'.
        anchor_cell: top-left cell for chart placement, e.g. 'E2'.
        """
        validate_cell_range(values_range)
        if category_range:
            validate_cell_range(category_range)

        subtypes = _CHART_SUBTYPES.get(chart_type, set())
        if chart_subtype is None:
            chart_subtype = next(iter(subtypes))  # default = first
        elif chart_subtype not in subtypes:
            raise ValueError(
                f"Invalid chart_subtype '{chart_subtype}' for chart_type '{chart_type}'. "
                f"Valid options: {sorted(subtypes)}"
            )

        if anchor_cell is None:
            anchor_cell = _default_anchor(values_range)

        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

        ws = wb[sheet_name]
        chart = _build_chart(ws, chart_type, chart_subtype, values_range,
                              category_range, series_names, title)
        _place_chart(ws, chart, anchor_cell)

        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"sheet_name": sheet_name, "anchor_cell": anchor_cell, "chart_type": chart_type}
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_charts.py -v
```

Expected: all 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add src/xlsx_drive_mcp/tools/charts.py tests/unit/test_charts.py
git commit -m "feat: add create_chart tool"
```

---

## Task 11: `tools/formulas.py`

**Files:**
- Create: `src/xlsx_drive_mcp/tools/formulas.py`
- Create: `tests/unit/test_formulas.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/unit/test_formulas.py
import pytest
from xlsx_drive_mcp.tools.formulas import _validate_formula, BLOCKED_FUNCTIONS


def test_valid_formula_passes():
    _validate_formula("=SUM(A1:A10)")  # should not raise


def test_formula_must_start_with_equals():
    with pytest.raises(ValueError, match="must start with"):
        _validate_formula("SUM(A1:A10)")


def test_webservice_blocked():
    with pytest.raises(ValueError, match="WEBSERVICE"):
        _validate_formula("=WEBSERVICE(\"http://example.com\")")


def test_importrange_blocked():
    with pytest.raises(ValueError, match="IMPORTRANGE"):
        _validate_formula("=IMPORTRANGE(\"sheet\", \"A1\")")


def test_blocked_functions_case_insensitive():
    with pytest.raises(ValueError):
        _validate_formula("=webservice(\"http://example.com\")")


def test_blocked_functions_list_nonempty():
    assert len(BLOCKED_FUNCTIONS) >= 6
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
python -m pytest tests/unit/test_formulas.py -v
```

Expected: `ImportError`.

- [ ] **Step 3: Implement `src/xlsx_drive_mcp/tools/formulas.py`**

```python
import io
import re

import openpyxl
from fastmcp import FastMCP
from googleapiclient.discovery import Resource

from ..drive import download_xlsx, upload_xlsx
from .data import _write_range

BLOCKED_FUNCTIONS = {
    "WEBSERVICE", "FILTERXML", "IMPORTDATA", "IMPORTFEED",
    "IMPORTHTML", "IMPORTRANGE", "IMPORTXML",
}

_BLOCKED_RE = re.compile(
    r"\b(" + "|".join(BLOCKED_FUNCTIONS) + r")\s*\(",
    re.IGNORECASE,
)


def _validate_formula(formula: str) -> None:
    """Raise ValueError if formula is invalid or uses a blocked function."""
    if not formula.startswith("="):
        raise ValueError(
            f"Formula must start with '='. Got: '{formula[:30]}'"
        )
    match = _BLOCKED_RE.search(formula)
    if match:
        func_name = match.group(1).upper()
        raise ValueError(
            f"Formula uses blocked function '{func_name}', which makes external network "
            f"connections when the file is opened in Excel or Google Sheets. "
            f"Blocked functions: {sorted(BLOCKED_FUNCTIONS)}"
        )


def register_tools(mcp: FastMCP, service: Resource) -> None:

    @mcp.tool()
    def write_formula(file_id: str, sheet_name: str, cell: str, formula: str) -> dict:
        """Write a formula to a single cell in an xlsx file on Google Drive.

        Equivalent to write_range with a 1x1 values array. The formula must
        start with '='. Functions that make external network connections are
        blocked: WEBSERVICE, FILTERXML, IMPORTDATA, IMPORTFEED, IMPORTHTML,
        IMPORTRANGE, IMPORTXML.
        """
        _validate_formula(formula)
        content, rev_id = download_xlsx(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
        _write_range(wb, sheet_name, cell, [[formula]])
        buf = io.BytesIO()
        wb.save(buf)
        upload_xlsx(service, file_id, buf.getvalue(), rev_id)
        return {"written": True, "cell": cell.upper(), "formula": formula}
```

- [ ] **Step 4: Run tests**

```bash
python -m pytest tests/unit/test_formulas.py -v
```

Expected: all 6 tests pass.

- [ ] **Step 5: Run full unit test suite**

```bash
python -m pytest tests/unit/ -v
```

Expected: all unit tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/xlsx_drive_mcp/tools/formulas.py tests/unit/test_formulas.py
git commit -m "feat: add write_formula tool with external-connection function denylist"
```

---

## Task 12: GitHub Actions CI/CD

**Files:**
- Create: `.github/workflows/test.yml`
- Create: `.github/workflows/publish.yml`

- [ ] **Step 1: Create `.github/workflows/test.yml`**

```yaml
name: Tests

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        python-version: ["3.10", "3.11", "3.12"]

    steps:
      - uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.python-version }}
        uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}

      - name: Install dependencies
        run: |
          pip install -e ".[dev]"

      - name: Run unit tests
        run: python -m pytest tests/unit/ -v

      - name: Run integration tests (if credentials available)
        if: ${{ secrets.DRIVE_TEST_FILE_ID != '' }}
        env:
          DRIVE_TEST_FILE_ID: ${{ secrets.DRIVE_TEST_FILE_ID }}
          GOOGLE_TEST_CREDENTIALS: ${{ secrets.GOOGLE_TEST_CREDENTIALS }}
        run: python -m pytest tests/integration/ -v
```

- [ ] **Step 2: Create `.github/workflows/publish.yml`**

```yaml
name: Publish to PyPI

on:
  push:
    tags:
      - "v*"

jobs:
  test:
    uses: ./.github/workflows/test.yml

  publish:
    needs: test
    runs-on: ubuntu-latest
    environment: pypi
    permissions:
      id-token: write  # required for OIDC Trusted Publisher

    steps:
      - uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install build tools
        run: pip install build

      - name: Build package
        run: python -m build

      - name: Publish to PyPI
        uses: pypa/gh-action-pypi-publish@release/v1
        # No API token needed -- uses OIDC Trusted Publisher
        # One-time setup: configure Trusted Publisher at https://pypi.org/manage/account/publishing/
        # Settings: owner=YOUR_GITHUB_USERNAME, repo=xlsx-drive-mcp,
        #            workflow=publish.yml, environment=pypi
```

- [ ] **Step 3: Commit**

```bash
git add .github/
git commit -m "ci: add test and publish GitHub Actions workflows"
```

---

## Task 13: Docs (README, CONTRIBUTING, CHANGELOG)

**Files:**
- Create: `README.md`
- Create: `CONTRIBUTING.md`
- Create: `CHANGELOG.md`

- [ ] **Step 1: Create `CHANGELOG.md`**

```markdown
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

**Note:** MCP tool names and parameter names are public API. Renaming or removing
them requires a major version bump and one-version deprecation cycle.

## [Unreleased]

### Added
- Initial release with full xlsx editing capabilities on Google Drive
```

- [ ] **Step 2: Create `CONTRIBUTING.md`**

```markdown
# Contributing to xlsx-drive-mcp

## Local dev setup

```bash
git clone https://github.com/YOUR_USERNAME/xlsx-drive-mcp
cd xlsx-drive-mcp
pip install -e ".[dev]"
python tests/fixtures/create_fixtures.py  # generate test xlsx files
python -m pytest tests/unit/ -v           # run unit tests (no credentials needed)
```

## How to add a new tool

1. Choose the appropriate module in `src/xlsx_drive_mcp/tools/`.
2. Add a pure function prefixed with `_` that takes an `openpyxl.Workbook` and returns a result.
3. Add a unit test for the pure function in `tests/unit/test_<module>.py` using a real Workbook.
4. Register the MCP tool in the module's `register_tools(mcp, service)` function.
5. The MCP wrapper should: validate inputs, download, open workbook, call the pure function, save, upload.
6. Add the tool to the README tool reference table.

## Running integration tests

Integration tests run against a real Drive file. You need:
- A credential file at `~/.google_workspace_mcp/credentials/your@email.com.json` or configured via env vars
- `export DRIVE_TEST_FILE_ID=<a real xlsx file_id on your Drive>`

```bash
python -m pytest tests/integration/ -v
```

## PR policy

- Tests must pass on Python 3.10, 3.11, 3.12.
- One logical change per PR.
- Update CHANGELOG.md under `[Unreleased]`.
- Do not rename or remove existing tool names or parameters without a major version discussion.
```

- [ ] **Step 3: Create `README.md`**

```markdown
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
```

- [ ] **Step 4: Commit**

```bash
git add README.md CONTRIBUTING.md CHANGELOG.md
git commit -m "docs: add README, CONTRIBUTING, and CHANGELOG"
```

---

## Task 14: Integration Tests

**Files:**
- Create: `tests/integration/test_drive_roundtrip.py`

- [ ] **Step 1: Create integration test**

```python
# tests/integration/test_drive_roundtrip.py
"""
Integration tests against a real Google Drive file.
Requires:
  - DRIVE_TEST_FILE_ID env var: a real xlsx file_id on Drive
  - Credentials at ~/.google_workspace_mcp/credentials/{email}.json
    OR GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET with stored credentials

Run: python -m pytest tests/integration/ -v
"""
import io
import os
import openpyxl
import pytest

DRIVE_TEST_FILE_ID = os.environ.get("DRIVE_TEST_FILE_ID")

pytestmark = pytest.mark.skipif(
    not DRIVE_TEST_FILE_ID,
    reason="DRIVE_TEST_FILE_ID env var not set",
)


@pytest.fixture(scope="module")
def service():
    from xlsx_drive_mcp.auth import load_credentials
    from xlsx_drive_mcp.drive import get_drive_service
    creds, _ = load_credentials(None)
    return get_drive_service(creds)


def test_download_upload_roundtrip(service):
    """Download, write a value, upload, re-download, verify the value persisted."""
    from xlsx_drive_mcp.drive import download_xlsx, upload_xlsx

    content, rev_id = download_xlsx(service, DRIVE_TEST_FILE_ID)
    wb = openpyxl.load_workbook(io.BytesIO(content))
    ws = wb.active
    original_value = ws["Z99"].value
    ws["Z99"] = "__integration_test_marker__"
    buf = io.BytesIO()
    wb.save(buf)
    upload_xlsx(service, DRIVE_TEST_FILE_ID, buf.getvalue(), rev_id)

    # Re-download and verify
    content2, _ = download_xlsx(service, DRIVE_TEST_FILE_ID)
    wb2 = openpyxl.load_workbook(io.BytesIO(content2), data_only=True)
    assert wb2.active["Z99"].value == "__integration_test_marker__"

    # Restore original value
    content3, rev_id3 = download_xlsx(service, DRIVE_TEST_FILE_ID)
    wb3 = openpyxl.load_workbook(io.BytesIO(content3))
    wb3.active["Z99"] = original_value
    buf3 = io.BytesIO()
    wb3.save(buf3)
    upload_xlsx(service, DRIVE_TEST_FILE_ID, buf3.getvalue(), rev_id3)


def test_conflict_detection(service):
    """Simulate a concurrent edit: second upload must raise ConflictError."""
    from xlsx_drive_mcp.drive import download_xlsx, upload_xlsx, ConflictError

    content, rev_id = download_xlsx(service, DRIVE_TEST_FILE_ID)

    # First upload succeeds and bumps the revision
    wb = openpyxl.load_workbook(io.BytesIO(content))
    wb.active["Z98"] = "first_write"
    buf = io.BytesIO()
    wb.save(buf)
    upload_xlsx(service, DRIVE_TEST_FILE_ID, buf.getvalue(), rev_id)

    # Second upload with the old rev_id must fail
    wb2 = openpyxl.load_workbook(io.BytesIO(content))
    wb2.active["Z98"] = "second_write"
    buf2 = io.BytesIO()
    wb2.save(buf2)
    with pytest.raises(ConflictError):
        upload_xlsx(service, DRIVE_TEST_FILE_ID, buf2.getvalue(), rev_id)
```

- [ ] **Step 2: Verify unit tests still all pass**

```bash
python -m pytest tests/unit/ -v
```

Expected: all pass.

- [ ] **Step 3: Commit**

```bash
git add tests/integration/
git commit -m "test: add integration tests for Drive roundtrip and conflict detection"
```

---

## Task 15: Final Wiring and Smoke Test

- [ ] **Step 1: Install and verify CLI**

```bash
pip install -e .
xlsx-drive-mcp --help
```

Expected output includes `--user` option and `auth` subcommand.

- [ ] **Step 2: Run full test suite**

```bash
python -m pytest tests/unit/ -v --tb=short
```

Expected: all unit tests pass.

- [ ] **Step 3: Verify package builds cleanly**

```bash
pip install build
python -m build
ls dist/
```

Expected: `xlsx_drive_mcp-0.1.0-py3-none-any.whl` and `.tar.gz` appear in `dist/`.

- [ ] **Step 4: Tag and push**

```bash
git tag v0.1.0
git push origin main --tags
```

This triggers the `publish.yml` workflow (once Trusted Publisher is configured on PyPI).

---

## Self-Review: Spec Coverage Check

| Spec requirement | Task |
|---|---|
| No bundled OAuth client ID | Task 3 (`auth.py` requires env vars) |
| workspace-mcp credential fallback | Task 3 |
| `auth` CLI subcommand | Task 4 (`server.py`) |
| Auto-detect single credential | Task 3 |
| 50MB file size limit | Task 2 (`drive.py`) |
| ETag/revision conflict detection | Task 2 |
| `cell_range` not `range` | All tool tasks |
| `sheet_name` consistent naming | All tool tasks |
| `validate_cell_range` (1M cell cap) | Task 2 |
| All logging to stderr | Task 2 |
| `list_xlsx_files` sanitized query | Task 6 |
| `get_xlsx_info` scanned dimensions | Task 6 |
| `append_rows` scans actual last row | Task 8 |
| `clear_range` raises on merged cells | Task 8 |
| `read_format` defined JSON structure | Task 9 |
| `format_range` border string/dict | Task 9 |
| `set_column_width` letter/list/range | Task 9 |
| `set_row_height` int/list/range | Task 9 |
| `create_chart` full spec | Task 10 |
| `write_formula` denylist | Task 11 |
| `copy_sheet` limitation documented | Task 7 |
| `read_range` data_only documented | Task 8 |
| Credential files `0600` perms | Task 3 |
| Token refresh on expiry | Task 3 |
| GitHub Actions test + publish | Task 12 |
| PyPI keywords | Task 1 |
| Semantic versioning + CHANGELOG | Task 13 |
| README auth walkthrough | Task 13 |
| README folder_id note | Task 13 |
| README troubleshooting | Task 13 |
| README formula/copy_sheet warnings | Task 13 |
| CONTRIBUTING.md | Task 13 |
| Fixture xlsx files | Task 5 |
| Integration tests (gated) | Task 14 |

# Contributing to xlsx-drive-mcp

## Local dev setup

```bash
git clone https://github.com/angep/xlsx-drive-mcp
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

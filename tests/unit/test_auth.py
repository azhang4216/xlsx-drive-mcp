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

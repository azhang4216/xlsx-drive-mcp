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

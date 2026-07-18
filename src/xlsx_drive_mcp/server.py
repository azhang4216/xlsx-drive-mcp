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

import io
import logging
import re
import sys

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError  # noqa: F401 — re-exported for callers
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
    return (
        service.files()
        .get(
            fileId=file_id,
            fields="id,name,size,mimeType,headRevisionId",
        )
        .execute()
    )


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
    media = MediaIoBaseUpload(
        io.BytesIO(content), mimetype=XLSX_MIME_TYPE, resumable=False
    )
    service.files().update(fileId=file_id, media_body=media).execute()

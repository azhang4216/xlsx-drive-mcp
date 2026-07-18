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

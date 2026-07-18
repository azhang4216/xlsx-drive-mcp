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

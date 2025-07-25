import pytest
from setup import *


def test_is_sheet_empty(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    assert excel_helper.is_sheet_empty(worksheet)


def test_is_sheet_empty_with_non_empty_sheet(add_content_to_sheet, excel_helper):
    worksheet = add_content_to_sheet
    assert not excel_helper.is_sheet_empty(worksheet)


def test_row_start_with_empty_sheet(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    assert excel_helper.is_sheet_empty(worksheet)
    row_start = excel_helper.get_row_start(worksheet)
    assert row_start == 1


def test_row_start_with_non_empty_sheet(add_content_to_sheet, config, excel_helper):
    worksheet = add_content_to_sheet
    current_max_row = worksheet.max_row
    assert (
        excel_helper.get_row_start(worksheet) == current_max_row + config.SEPARATOR + 1
    )


import pytest


@pytest.mark.parametrize(
    "col, expected",
    [
        (None, "A"),
        ("A", "A"),
        ("Z", "Z"),
        ("AA", "AA"),
        (1, "A"),
        (26, "Z"),
        (27, "AA"),
        (16384, "XFD"),
        # NOTE a decision was made to prevent raising exceptions resulted from excel-related
        # code. a UserWarning is giving instead and a default config/value is used.
        # ("XFE", ValueError),
        # (16385, ValueError),
        # (-1, ValueError),
        # ("$", ValueError),
        # ("1", ValueError),
    ],
)
def test_get_column_letter(excel_helper, col, expected):
    if isinstance(expected, type) and issubclass(expected, Exception):
        with pytest.raises(expected):
            excel_helper.get_column_letter(col)
    else:
        result = excel_helper.get_column_letter(col)
        assert result == expected

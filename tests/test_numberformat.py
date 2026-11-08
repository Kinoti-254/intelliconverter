"""Tests for intelliconverter.numberformat: showing values the way Excel does.

Covers the basic formats agreed for Phase 2b. Anything unrecognised falls
back to Excel's General format rather than failing.
"""
import datetime as dt

import pytest

from intelliconverter.numberformat import format_value


@pytest.mark.parametrize(
    "value, expected",
    [
        (None, ""),
        ("Nairobi", "Nairobi"),
        (True, "TRUE"),
        (False, "FALSE"),
        (23.0, "23"),
        (23, "23"),
        (-7, "-7"),
        (0.1 + 0.2, "0.3"),
        (1234567.89, "1234567.89"),
        (2.5, "2.5"),
        (1e20, "1E+20"),
    ],
)
def test_general(value, expected):
    assert format_value(value, "General") == expected
    assert format_value(value, None) == expected


@pytest.mark.parametrize(
    "value, number_format, expected",
    [
        (1234567.891, "#,##0.00", "1,234,567.89"),
        (1234567.891, "#,##0", "1,234,568"),
        (5, "0.00", "5.00"),
        (2.675, "0.00", "2.68"),            # rounds half up, like Excel
        (-2.675, "0.00", "-2.68"),
        (7.5, "0", "8"),
        (42, "000", "042"),                 # leading zeros, e.g. codes
        (0.5, "#.00", ".50"),
        (1.5, "0.##", "1.5"),               # optional decimals
        (2.0, "0.##", "2."),
        (0.256, "0.0%", "25.6%"),
        (0.5, "0%", "50%"),
        (42, '"KES "#,##0', "KES 42"),
        (1250, "$#,##0.00", "$1,250.00"),
        (1250, "#,##0.00 [$KES]", "1,250.00 KES"),
        (12345.678, "0.00E+00", "1.23E+04"),
        (-1234.5, "#,##0.00_);(#,##0.00)", "(1,234.50)"),
        (1234.5, "#,##0.00_);(#,##0.00)", "1,234.50"),
        (-5, "#,##0;[Red]-#,##0", "-5"),
        (0, '#,##0;-#,##0;"-"', "-"),
        (3, "@", "3"),
        (3, "# ?/?", "3"),                  # fractions not supported: General
    ],
)
def test_number_formats(value, number_format, expected):
    assert format_value(value, number_format) == expected


@pytest.mark.parametrize(
    "value, number_format, expected",
    [
        (dt.datetime(2026, 11, 8), "dd/mm/yyyy", "08/11/2026"),
        (dt.datetime(2026, 11, 8), "d/m/yy", "8/11/26"),
        (dt.datetime(2026, 11, 8), "yyyy-mm-dd", "2026-11-08"),
        (dt.datetime(2026, 11, 8), "mm-dd-yy", "11-08-26"),
        (dt.datetime(2026, 11, 8), "d-mmm-yy", "8-Nov-26"),
        (dt.datetime(2026, 11, 8), "dddd, mmmm d, yyyy", "Sunday, November 8, 2026"),
        (dt.datetime(2026, 11, 8, 14, 5), "d-mmm-yy h:mm AM/PM", "8-Nov-26 2:05 PM"),
        (dt.datetime(2026, 11, 8, 9, 7, 3), "hh:mm:ss", "09:07:03"),
        (dt.datetime(2026, 11, 8, 0, 30), "h:mm AM/PM", "12:30 AM"),
        (dt.time(14, 5), "h:mm", "14:05"),
        (dt.date(2026, 1, 2), "dd mmm yyyy", "02 Jan 2026"),
        (dt.datetime(2026, 11, 8), "General", "2026-11-08"),
        (dt.datetime(2026, 11, 8, 14, 5), "General", "2026-11-08 14:05:00"),
    ],
)
def test_dates_and_times(value, number_format, expected):
    assert format_value(value, number_format) == expected


def test_minutes_and_months_are_told_apart():
    when = dt.datetime(2026, 3, 8, 10, 45)
    assert format_value(when, "m/d h:mm") == "3/8 10:45"
    assert format_value(when, "mm:ss") == "45:00"

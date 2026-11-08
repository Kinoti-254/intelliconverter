"""Characterisation tests for intelliconverter.measure.

These pin the current behaviour so the Phase 0 restructure can be shown
not to change anything. Behaviour that is a known defect is NOT pinned
here; see test_known_defects.py.
"""
import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.reader import load_active_sheet
from tests.conftest import SAMPLE_XLSX


class TestColWidthToInches:
    def test_zero_width_is_padding_only(self):
        assert col_width_to_inches(0) == 5 / 96

    def test_excel_default_width(self):
        # 8.43 characters is Excel's default column width
        assert col_width_to_inches(8.43) == pytest.approx(64.01 / 96)

    def test_formula_assumes_7px_per_character_at_96_dpi(self):
        assert col_width_to_inches(10) == 75 / 96


class TestGetTotalTableWidth:
    def test_sample_workbook(self):
        # test.xlsx: column B is 25.63 wide; A, C, D use the sheet default 12.63
        ws = load_active_sheet(SAMPLE_XLSX)
        expected = 3 * col_width_to_inches(12.63) + col_width_to_inches(25.63)
        assert get_total_table_width(ws) == pytest.approx(expected)
        assert f"{get_total_table_width(ws):.2f}" == "4.84"

    def test_falls_back_to_843_when_sheet_has_no_default(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws.sheet_format.defaultColWidth is None
        for col in range(1, 4):
            ws.cell(row=1, column=col, value="x")
        assert get_total_table_width(ws) == pytest.approx(3 * col_width_to_inches(8.43))

    def test_uses_sheet_default_width_when_set(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.sheet_format.defaultColWidth = 20
        ws.cell(row=1, column=2, value="x")
        assert get_total_table_width(ws) == pytest.approx(2 * col_width_to_inches(20))

    def test_custom_column_widths_are_used(self, reload_sheet):
        wb = openpyxl.Workbook()
        ws = wb.active
        widths = [5, 15, 30]
        for col, width in enumerate(widths, start=1):
            ws.cell(row=1, column=col, value="x")
            ws.column_dimensions[get_column_letter(col)].width = width
        ws = reload_sheet(wb)
        expected = sum(col_width_to_inches(w) for w in widths)
        assert get_total_table_width(ws) == pytest.approx(expected)

    def test_empty_sheet_counts_one_column(self):
        # openpyxl reports max_column = 1 for an empty sheet
        ws = openpyxl.Workbook().active
        assert get_total_table_width(ws) == pytest.approx(col_width_to_inches(8.43))

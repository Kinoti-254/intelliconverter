"""Tests for intelliconverter.measure.

The unit conversion is unchanged since Phase 0. Since Phase 1 the width is
measured from a SheetModel, which supplies the correct columns (visible
only, ranges expanded, used range only).
"""
import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.reader import load_sheet
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
        # test.xlsx: column B is 25.63 wide; A, C and D use the sheet default 12.63
        model = load_sheet(SAMPLE_XLSX)
        expected = 3 * col_width_to_inches(12.63) + col_width_to_inches(25.63)
        assert get_total_table_width(model) == pytest.approx(expected)
        assert f"{get_total_table_width(model):.2f}" == "4.84"

    def test_falls_back_to_843_when_sheet_has_no_default(self, save_workbook):
        wb = openpyxl.Workbook()
        wb.active.append(["x", "x", "x"])
        model = load_sheet(save_workbook(wb))
        assert get_total_table_width(model) == pytest.approx(3 * col_width_to_inches(8.43))

    def test_uses_sheet_default_width_when_set(self, save_workbook):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.sheet_format.defaultColWidth = 20
        ws.cell(row=1, column=2, value="x")
        model = load_sheet(save_workbook(wb))
        assert get_total_table_width(model) == pytest.approx(2 * col_width_to_inches(20))

    def test_custom_column_widths_are_used(self, save_workbook):
        wb = openpyxl.Workbook()
        ws = wb.active
        widths = [5, 15, 30]
        for col, width in enumerate(widths, start=1):
            ws.cell(row=1, column=col, value="x")
            ws.column_dimensions[get_column_letter(col)].width = width
        model = load_sheet(save_workbook(wb))
        expected = sum(col_width_to_inches(w) for w in widths)
        assert get_total_table_width(model) == pytest.approx(expected)

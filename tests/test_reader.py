"""Tests for intelliconverter.reader: building a SheetModel and validating input.

Includes the regression tests for defects 2, 3, 7 and 8, fixed in Phase 1.
"""
import re
import zipfile

import openpyxl
import pytest
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

from intelliconverter import reader
from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.reader import InputError, load_sheet
from tests.conftest import SAMPLE_XLSX


def _edit_sheet_xml(path, pattern, replacement):
    """Apply one regex substitution to the first sheet's XML inside an .xlsx.

    openpyxl writes slightly different XML depending on whether the optional
    lxml library is installed (`<v></v>` or `<v />`), so patterns must allow
    both. The edit must match exactly once.
    """
    sheet_xml = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as src:
        parts = {name: src.read(name) for name in src.namelist()}
    parts[sheet_xml], count = re.subn(pattern, replacement, parts[sheet_xml])
    assert count == 1, f"pattern {pattern!r} matched {count} times"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name, data in parts.items():
            dst.writestr(name, data)


def _sheet_with_widths(widths):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col, width in enumerate(widths, start=1):
        ws.cell(row=1, column=col, value="x")
        ws.column_dimensions[get_column_letter(col)].width = width
    return wb


# --- What the model contains ------------------------------------------------

class TestSampleModel:
    @pytest.fixture
    def model(self):
        return load_sheet(SAMPLE_XLSX)

    def test_title_and_size(self, model):
        assert model.title == "Sheet1"
        assert (model.row_count, model.column_count) == (3, 4)

    def test_column_widths(self, model):
        # Column B has a custom width; A, C and D use the sheet default.
        assert [c.width for c in model.columns] == [12.63, 25.63, 12.63, 12.63]
        assert [c.letter for c in model.columns] == ["A", "B", "C", "D"]
        assert not any(c.hidden for c in model.columns)

    def test_values(self, model):
        assert model.visible_rows() == [
            ["Name", "Age", "City", "Notes"],
            ["Sam", 23.0, "Nairobi", "Likes coffee"],
            ["Alex", 31.0, "Mombasa", "Plays guitar"],
        ]

    def test_no_hidden_content_or_formulas(self, model):
        assert model.hidden_rows == frozenset()
        assert model.uncached_formula_count == 0


def test_default_width_is_843_when_sheet_has_none(save_workbook):
    wb = openpyxl.Workbook()
    wb.active.append(["a", "b"])
    model = load_sheet(save_workbook(wb))
    assert [c.width for c in model.columns] == [8.43, 8.43]


def test_selects_a_sheet_by_name(save_workbook):
    wb = openpyxl.Workbook()
    wb.active.title = "First"
    wb.active["A1"] = "first"
    wb.create_sheet("Second")["A1"] = "second"
    path = save_workbook(wb)
    assert load_sheet(path).title == "First"
    assert load_sheet(path, sheet="Second").visible_rows() == [["second"]]


def test_hidden_rows_are_excluded_from_visible_rows(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    for value in ["keep 1", "drop", "keep 2"]:
        ws.append([value])
    ws.row_dimensions[2].hidden = True
    model = load_sheet(save_workbook(wb))
    assert model.hidden_rows == frozenset({2})
    assert model.visible_rows() == [["keep 1"], ["keep 2"]]


def test_empty_cells_inside_the_used_range_are_kept(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "top left"
    ws["C3"] = "bottom right"
    model = load_sheet(save_workbook(wb))
    assert model.visible_rows() == [
        ["top left", None, None],
        [None, None, None],
        [None, None, "bottom right"],
    ]


def test_wrong_size_recorded_in_the_file_does_not_hide_cells(save_workbook):
    # Some programs write a wrong <dimension> into the sheet. openpyxl's
    # read-only mode trusts it and, without care, silently returns only A1.
    wb = openpyxl.Workbook()
    wb.active["A1"] = "a"
    wb.active["C5"] = "z"
    path = save_workbook(wb, "dimension.xlsx")
    _edit_sheet_xml(path, rb'<dimension ref="A1:C5"\s*/>', b'<dimension ref="A1"/>')
    model = load_sheet(path)
    assert (model.row_count, model.column_count) == (5, 3)
    assert model.visible_rows()[4] == [None, None, "z"]


def test_uncached_formulas_are_counted(save_workbook):
    # openpyxl saves formulas without results, like a file never opened in Excel.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    ws["A3"] = "=A1*4"
    model = load_sheet(save_workbook(wb))
    assert model.uncached_formula_count == 2
    assert model.visible_rows() == [[2], [None], [None]]


# --- Regression tests for defects fixed in Phase 1 --------------------------

def test_defect_2_hidden_columns_are_excluded(save_workbook):
    wb = _sheet_with_widths([20] * 10)
    for letter in "CDEFGH":
        wb.active.column_dimensions[letter].hidden = True
    model = load_sheet(save_workbook(wb))
    # Only A, B, I and J are visible.
    assert [c.letter for c in model.visible_columns] == ["A", "B", "I", "J"]
    assert get_total_table_width(model) == pytest.approx(4 * col_width_to_inches(20))


def test_defect_3_grouped_column_range_counts_every_column(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col in range(1, 6):
        ws.cell(row=1, column=col, value="x")
    ws.column_dimensions.group("B", "E", hidden=False)
    ws.column_dimensions["B"].width = 30
    model = load_sheet(save_workbook(wb))
    # openpyxl stores B:E as one entry keyed "B"; C, D and E are also width 30.
    assert [c.width for c in model.columns] == [8.43, 30, 30, 30, 30]
    expected = col_width_to_inches(8.43) + 4 * col_width_to_inches(30)
    assert get_total_table_width(model) == pytest.approx(expected)


def test_defect_7_formatted_empty_cells_do_not_add_width(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "data"
    ws["Z1"].number_format = "0.00"  # formatted, but empty
    ws["A40"].font = openpyxl.styles.Font(bold=True)  # formatted, but empty
    model = load_sheet(save_workbook(wb))
    assert (model.row_count, model.column_count) == (1, 1)
    assert get_total_table_width(model) == pytest.approx(col_width_to_inches(8.43))


def test_defect_8_formula_cells_read_as_cached_values(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    path = save_workbook(wb, "formula.xlsx")
    # Write the result Excel would cache when saving (openpyxl writes none).
    _edit_sheet_xml(path, rb"<f>A1\*3</f>(<v></v>|<v\s*/>)", b"<f>A1*3</f><v>6</v>")
    model = load_sheet(path)
    assert model.visible_rows() == [[2], [6]]
    assert model.uncached_formula_count == 0


# --- Input validation -------------------------------------------------------

def _message(excinfo):
    return str(excinfo.value)


def test_missing_file(tmp_path):
    with pytest.raises(InputError) as excinfo:
        load_sheet(tmp_path / "nope.xlsx")
    assert "File not found" in _message(excinfo)


def test_folder_instead_of_file(tmp_path):
    with pytest.raises(InputError) as excinfo:
        load_sheet(tmp_path)
    assert "not a file" in _message(excinfo)


@pytest.mark.parametrize(
    "name, expected",
    [
        ("data.csv", "CSV files aren't supported yet"),
        ("old.xls", "save it as .xlsx"),
        ("notes.txt", "Unsupported file type"),
    ],
)
def test_unsupported_extensions(tmp_path, name, expected):
    path = tmp_path / name
    path.write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(InputError) as excinfo:
        load_sheet(path)
    assert expected in _message(excinfo)


def test_text_file_renamed_to_xlsx(tmp_path):
    path = tmp_path / "fake.xlsx"
    path.write_text("this is not a spreadsheet", encoding="utf-8")
    with pytest.raises(InputError) as excinfo:
        load_sheet(path)
    assert "isn't a valid Excel file" in _message(excinfo)


def test_password_protected_or_legacy_file(tmp_path):
    # Encrypted .xlsx files and old .xls files share this header.
    path = tmp_path / "locked.xlsx"
    path.write_bytes(bytes.fromhex("d0cf11e0a1b11ae1") + b"\x00" * 504)
    with pytest.raises(InputError) as excinfo:
        load_sheet(path)
    assert "password-protected" in _message(excinfo)


def test_zip_that_is_not_a_workbook(tmp_path):
    path = tmp_path / "archive.xlsx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("readme.txt", "hello")
    with pytest.raises(InputError) as excinfo:
        load_sheet(path)
    assert "isn't a valid Excel file" in _message(excinfo)


def test_unknown_sheet_lists_the_available_ones(save_workbook):
    wb = openpyxl.Workbook()
    wb.active.title = "Sales"
    wb.create_sheet("Costs")
    with pytest.raises(InputError) as excinfo:
        load_sheet(save_workbook(wb), sheet="Profit")
    assert "No sheet named 'Profit'" in _message(excinfo)
    assert "Sales, Costs" in _message(excinfo)


def test_chart_sheet_is_rejected(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([1])
    ws.append([2])
    chart_sheet = wb.create_chartsheet("Chart")
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=1, min_row=1, max_row=2))
    chart_sheet.add_chart(chart)
    wb.active = 1
    with pytest.raises(InputError) as excinfo:
        load_sheet(save_workbook(wb))
    assert "chart sheet" in _message(excinfo)


def test_empty_sheet(save_workbook):
    wb = openpyxl.Workbook()
    wb.active["C3"].number_format = "0.00"  # formatting only
    with pytest.raises(InputError) as excinfo:
        load_sheet(save_workbook(wb))
    assert "is empty" in _message(excinfo)


def test_sheet_with_everything_hidden(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "secret"
    ws.column_dimensions["A"].hidden = True
    with pytest.raises(InputError) as excinfo:
        load_sheet(save_workbook(wb))
    assert "no visible data" in _message(excinfo)


class TestLimits:
    def test_file_size(self, monkeypatch):
        monkeypatch.setattr(reader, "MAX_FILE_BYTES", 1000)
        with pytest.raises(InputError) as excinfo:
            load_sheet(SAMPLE_XLSX)
        assert "the limit is" in _message(excinfo) and "MB" in _message(excinfo)

    def test_rows(self, monkeypatch):
        monkeypatch.setattr(reader, "MAX_ROWS", 2)
        with pytest.raises(InputError) as excinfo:
            load_sheet(SAMPLE_XLSX)
        assert "3 rows; the limit is 2" in _message(excinfo)

    def test_columns(self, monkeypatch):
        monkeypatch.setattr(reader, "MAX_COLUMNS", 3)
        with pytest.raises(InputError) as excinfo:
            load_sheet(SAMPLE_XLSX)
        assert "4 columns; the limit is 3" in _message(excinfo)

    def test_default_limits(self):
        assert reader.MAX_ROWS == 50_000
        assert reader.MAX_COLUMNS == 100
        assert reader.MAX_FILE_BYTES == 20 * 1024 * 1024


def test_load_active_sheet_still_returns_a_worksheet(save_workbook):
    # Kept from Phase 0 for existing callers.
    wb = openpyxl.Workbook()
    wb.active.title = "First"
    wb.create_sheet("Second")
    wb.active = 1
    assert reader.load_active_sheet(save_workbook(wb)).title == "Second"

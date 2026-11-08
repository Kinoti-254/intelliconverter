"""Tests for PDF generation (intelliconverter.render).

These pin what the PDF code does: page sizes, page count, table contents,
column widths and the confirmation output. Since Phase 1 the PDF is built
from a SheetModel, so hidden rows and columns are left out and formula
cells show their saved values. Behaviour that is a known defect is tested
in test_known_defects.py instead.
"""
import time

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A3, A4, LEGAL, landscape, portrait

import excel_to_pdf
from intelliconverter import render
from intelliconverter.layout import LAYOUTS
from intelliconverter.measure import get_total_table_width
from intelliconverter.reader import ColumnInfo, SheetModel, load_sheet
from intelliconverter.report import analyze
from tests.conftest import SAMPLE_XLSX, read_pdf

INCH = 72  # points

# Sample file: best layout is A4 Portrait (7.7 in) at the fill shown in the report.
SAMPLE_SCALE = get_total_table_width(load_sheet(SAMPLE_XLSX)) / 7.7 * 100


def _sheet_with_widths(widths):
    wb = openpyxl.Workbook()
    ws = wb.active
    for col, width in enumerate(widths, start=1):
        ws.cell(row=1, column=col, value=f"H{col}")
        ws.cell(row=2, column=col, value="x")
        ws.column_dimensions[get_column_letter(col)].width = width
    return wb


class TestPageSizes:
    @pytest.mark.parametrize(
        "layout_name, expected",
        [
            ("A4 Portrait", portrait(A4)),
            ("A4 Landscape", landscape(A4)),
            ("A3 Portrait", portrait(A3)),
            ("A3 Landscape", landscape(A3)),
        ],
    )
    def test_layout_names_map_to_reportlab_sizes(self, layout_name, expected):
        assert render.get_page_size(layout_name) == expected

    def test_unknown_layout_falls_back_to_a4_landscape(self):
        assert render.get_page_size("Something else") == landscape(A4)
        assert render.get_page_size(None) == landscape(A4)

    def test_defect_6b_legal_landscape_uses_legal_paper(self):
        # Regression test: before Phase 2a it used Letter paper (11 in, not 14 in).
        assert render.get_page_size("Legal Landscape") == landscape(LEGAL)

    def test_every_scored_layout_has_a_page_size(self):
        assert {name for name, _ in LAYOUTS} == set(render.PAGE_SIZES)


class TestExtractSheetData:
    def test_values_become_strings_and_empty_cells_blank(self, reload_sheet):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Name", 5, None])
        ws.append(["Sam", 2.5, "ok"])
        assert render.extract_sheet_data(reload_sheet(wb)) == [
            ["Name", "5", ""],
            ["Sam", "2.5", "ok"],
        ]


class TestBuildPdfOnSample:
    """build_pdf on test.xlsx with the layout the report recommends."""

    @pytest.fixture
    def built(self, tmp_path, captured_tables, capsys):
        out = tmp_path / "out.pdf"
        render.build_pdf(SAMPLE_XLSX, str(out), "A4 Portrait", SAMPLE_SCALE)
        return {
            "path": str(out),
            "pdf": read_pdf(out),
            "table": captured_tables[0],
            "console": capsys.readouterr().out,
        }

    def test_one_a4_portrait_page(self, built):
        assert built["pdf"]["pages"] == 1
        width, height = built["pdf"]["sizes"][0]
        assert (width, height) == pytest.approx(portrait(A4))

    def test_title_and_subtitle(self, built):
        text = built["pdf"]["text"]
        assert "Sheet1" in text
        assert "Layout: A4 Portrait | Scale: 63% | Rows: 3 | Columns: 4" in text

    def test_every_cell_value_appears(self, built):
        for value in ["Name", "Age", "City", "Notes", "Sam", "Nairobi", "Likes coffee",
                      "Alex", "Mombasa", "Plays guitar"]:
            assert value in built["pdf"]["text"]

    def test_table_holds_every_row_and_text_cell(self, built):
        # Number formatting (23.0) is defect 6e, so numeric cells are not pinned.
        data = built["table"]["data"]
        assert len(data) == 3 and all(len(row) == 4 for row in data)
        assert data[0] == ["Name", "Age", "City", "Notes"]
        assert [data[1][i] for i in (0, 2, 3)] == ["Sam", "Nairobi", "Likes coffee"]
        assert [data[2][i] for i in (0, 2, 3)] == ["Alex", "Mombasa", "Plays guitar"]

    def test_table_width_is_usable_page_width_times_scale(self, built):
        # How the width is split between columns is defect 6d, so only the
        # total is pinned.
        widths = built["table"]["col_widths"]
        usable_width = portrait(A4)[0] - INCH
        assert len(widths) == 4
        assert sum(widths) == pytest.approx(usable_width * SAMPLE_SCALE / 100)

    def test_header_row_repeats_on_each_page(self, built):
        assert built["table"]["repeat_rows"] == 1

    def test_prints_a_confirmation(self, built):
        assert built["console"] == (
            f"\n✅ PDF saved to: {built['path']}\n"
            "   Layout  : A4 Portrait\n"
            "   Scale   : 63%\n"
            "   Rows    : 3\n"
            "   Columns : 4\n\n"
        )


def test_long_tables_paginate(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Row", "Value"])
    for row in range(1, 301):
        ws.append([row, "row"])
    src = tmp_path / "tall.xlsx"
    wb.save(src)
    out = tmp_path / "tall.pdf"
    render.build_pdf(src, str(out), "A4 Portrait", 50)
    pdf = read_pdf(out)
    assert pdf["pages"] > 1
    assert "300" in pdf["text"]


def test_no_data_prints_a_warning_and_writes_nothing(tmp_path, capsys):
    # load_sheet() rejects sheets with no visible data, so this branch is
    # reached with a hand-built model whose only row is hidden.
    model = SheetModel(
        title="Sheet1",
        columns=[ColumnInfo(index=1, letter="A", width=8.43, hidden=False)],
        hidden_rows=frozenset({1}),
        cells=[["x"]],
        uncached_formula_count=0,
    )
    out = tmp_path / "out.pdf"
    render.render_pdf(model, str(out), "A4 Portrait", 50)
    assert capsys.readouterr().out == "⚠️  No data found in sheet.\n"
    assert not out.exists()


class TestWhatThePdfContains:
    def test_hidden_rows_and_columns_are_left_out(self, tmp_path, captured_tables, capsys):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Name", "Secret", "City"])
        ws.append(["Sam", "s1", "Nairobi"])
        ws.append(["Hidden row", "s2", "Nowhere"])
        ws.column_dimensions["B"].hidden = True
        ws.row_dimensions[3].hidden = True
        src = tmp_path / "hidden.xlsx"
        wb.save(src)
        out = tmp_path / "out.pdf"
        render.build_pdf(src, str(out), "A4 Portrait", 50)
        assert captured_tables[0]["data"] == [["Name", "City"], ["Sam", "Nairobi"]]
        text = read_pdf(out)["text"]
        assert "Secret" not in text and "Hidden row" not in text
        assert "Rows: 2 | Columns: 2" in text

    def test_formula_cells_show_saved_values_not_formulas(self, tmp_path, captured_tables, capsys):
        wb = openpyxl.Workbook()
        wb.active["A1"] = 2
        wb.active["A2"] = "=A1*3"
        src = tmp_path / "formula.xlsx"
        wb.save(src)  # no saved result, so the cell is blank
        render.build_pdf(src, str(tmp_path / "out.pdf"), "A4 Portrait", 50)
        assert captured_tables[0]["data"] == [["2"], [""]]

    def test_a_named_sheet_can_be_rendered(self, tmp_path, captured_tables, capsys):
        wb = openpyxl.Workbook()
        wb.active["A1"] = "first"
        wb.create_sheet("Second")["A1"] = "second"
        src = tmp_path / "two.xlsx"
        wb.save(src)
        render.build_pdf(src, str(tmp_path / "out.pdf"), "A4 Portrait", 50, sheet="Second")
        assert captured_tables[0]["data"] == [["second"]]

    def test_render_pdf_and_build_pdf_give_the_same_file(self, tmp_path, deterministic_pdfs, capsys):
        from_file, from_model = tmp_path / "file.pdf", tmp_path / "model.pdf"
        render.build_pdf(SAMPLE_XLSX, str(from_file), "A4 Portrait", SAMPLE_SCALE)
        render.render_pdf(load_sheet(SAMPLE_XLSX), str(from_model), "A4 Portrait", SAMPLE_SCALE)
        assert from_file.read_bytes() == from_model.read_bytes()


@pytest.mark.parametrize(
    "widths, layout_name, page_size",
    [
        ([11] * 10, "A4 Landscape", landscape(A4)),
        ([10] * 14, "A3 Portrait", portrait(A3)),
        ([20] * 10, "A3 Landscape", landscape(A3)),
    ],
)
def test_analyze_exports_pdf_on_the_recommended_page(workdir, capsys, widths, layout_name, page_size):
    _sheet_with_widths(widths).save(workdir / "sheet.xlsx")
    analyze("sheet.xlsx")
    assert f"Best layout : {layout_name}" in capsys.readouterr().out
    pdf = read_pdf(workdir / "output.pdf")
    assert pdf["sizes"][0] == pytest.approx(page_size)


def test_pdf_is_byte_identical_across_runs_in_invariant_mode(tmp_path, deterministic_pdfs):
    first, second = tmp_path / "a.pdf", tmp_path / "b.pdf"
    render.build_pdf(SAMPLE_XLSX, str(first), "A4 Portrait", SAMPLE_SCALE)
    # Without invariant mode the PDF embeds a timestamp to the second; the
    # pause makes sure this test would catch that.
    time.sleep(1.1)
    render.build_pdf(SAMPLE_XLSX, str(second), "A4 Portrait", SAMPLE_SCALE)
    assert first.read_bytes() == second.read_bytes()


class TestExcelToPdfWrapper:
    def test_reexports_the_same_objects(self):
        assert excel_to_pdf.build_pdf is render.build_pdf
        assert excel_to_pdf.get_page_size is render.get_page_size
        assert excel_to_pdf.extract_sheet_data is render.extract_sheet_data
        assert excel_to_pdf.PAGE_SIZES is render.PAGE_SIZES

    def test_old_import_produces_the_same_pdf(self, tmp_path, deterministic_pdfs):
        from excel_to_pdf import build_pdf

        via_wrapper, via_package = tmp_path / "old.pdf", tmp_path / "new.pdf"
        build_pdf(SAMPLE_XLSX, str(via_wrapper), "A4 Portrait", SAMPLE_SCALE)
        render.build_pdf(SAMPLE_XLSX, str(via_package), "A4 Portrait", SAMPLE_SCALE)
        assert via_wrapper.read_bytes() == via_package.read_bytes()

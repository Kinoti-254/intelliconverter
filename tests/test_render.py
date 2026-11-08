"""Tests for PDF generation: the PDF must follow the LayoutPlan exactly.

For each sheet the plan is made first, then the PDF is drawn and read back,
and every promise the plan makes is checked: page count, page size, column
widths, row heights, wrapped lines, page breaks, the repeated header, the
footer, and that every visible cell appears once.
"""
import time
import warnings

import openpyxl
import pytest

import excel_to_pdf
from intelliconverter import planner, render
from intelliconverter.planner import PAD_X, plan_layout
from intelliconverter.reader import load_sheet
from intelliconverter.textmetrics import BODY_FONT, HEADER_FONT, text_width, wrap_text
from tests.conftest import SAMPLE_XLSX, read_pdf

LOREM = "Delivery moved to the afternoon because the loading bay is closed for maintenance"


def _book(tmp_path, name, rows, setup=None):
    wb = openpyxl.Workbook()
    for row in rows:
        wb.active.append(row)
    if setup:
        setup(wb.active)
    path = tmp_path / f"{name}.xlsx"
    wb.save(path)
    return load_sheet(path)


def sample(tmp_path):
    return load_sheet(SAMPLE_XLSX)


def tall(tmp_path):
    return _book(tmp_path, "tall", [["Row", "Value"]] + [[i, f"value {i}"] for i in range(1, 301)])


def wide(tmp_path):
    rows = [["ID"] + [f"Metric {c:02d}" for c in range(1, 30)]]
    rows += [[f"ROW-{r:03d}"] + [f"val{r:02d}-{c:02d}" for c in range(1, 30)] for r in range(60)]
    return _book(tmp_path, "wide", rows)


def notes(tmp_path):
    return _book(tmp_path, "notes", [["ID", "Notes"]] + [[i, LOREM] for i in range(40)])


def header_only(tmp_path):
    return _book(tmp_path, "header", [["Only", "A", "Header"]])


SHEETS = [sample, tall, wide, notes, header_only]


@pytest.fixture(params=SHEETS, ids=[s.__name__ for s in SHEETS])
def rendered(request, tmp_path, captured_tables):
    model = request.param(tmp_path)
    result = plan_layout(model)
    out = tmp_path / "out.pdf"
    render.render_pdf(model, result, str(out))
    return {"model": model, "result": result, "plan": result.plan,
            "pdf": read_pdf(out), "tables": list(captured_tables)}


# --- The PDF follows the plan ---------------------------------------------------

def test_page_count_matches_the_plan(rendered):
    assert rendered["pdf"]["pages"] == rendered["plan"].pages == len(rendered["tables"])


def test_every_page_has_the_planned_size(rendered):
    page = rendered["plan"].page
    assert all(size == pytest.approx((page.width, page.height)) for size in rendered["pdf"]["sizes"])


def _pages_by_band(rendered):
    """Pair each drawn table with its band and the body rows it should hold."""
    tables = iter(rendered["tables"])
    for band in rendered["plan"].bands:
        bounds = list(band.page_starts) + [len(band.row_heights)]
        for start, end in zip(bounds, bounds[1:]):
            yield band, start, end, next(tables)


def test_column_widths_and_row_heights_are_the_planned_ones(rendered):
    for band, start, end, table in _pages_by_band(rendered):
        assert list(table["col_widths"]) == pytest.approx(list(band.widths))
        expected = [band.header_height] + list(band.row_heights[start:end])
        assert list(table["row_heights"]) == pytest.approx(expected)


def test_cells_hold_the_planned_wrapped_lines(rendered):
    plan, texts = rendered["plan"], rendered["result"].measured
    size = plan.font_size
    for band, start, end, table in _pages_by_band(rendered):
        header, *body = table["data"]
        for position, column in enumerate(band.columns):
            content = band.widths[position] - 2 * PAD_X
            expected = wrap_text(texts.header[column], content, HEADER_FONT, size)
            assert header[position] == "\n".join(expected)
            assert len(expected) == band.header_lines[position]
            for offset, row in enumerate(range(start, end)):
                lines = wrap_text(texts.body[row][column], content, BODY_FONT, size)
                assert body[offset][position] == "\n".join(lines)
                assert len(lines) == band.body_lines[row][position]


def test_no_printed_line_is_wider_than_its_column(rendered):
    size = rendered["plan"].font_size
    for band, _, _, table in _pages_by_band(rendered):
        for row_index, row in enumerate(table["data"]):
            font = HEADER_FONT if row_index == 0 else BODY_FONT
            for cell, width in zip(row, band.widths):
                for line in cell.split("\n"):
                    assert text_width(line, font, size) <= width - 2 * PAD_X + 1e-6


def test_every_body_row_appears_once_per_band_in_order(rendered):
    count = len(rendered["result"].measured.body)
    for band in rendered["plan"].bands:
        pages = [(s, e) for b, s, e, _ in _pages_by_band(rendered) if b is band]
        rows = [r for s, e in pages for r in range(s, e)]
        assert rows == list(range(count))


def test_every_page_repeats_the_header_and_has_a_page_number(rendered):
    plan, texts = rendered["plan"], rendered["result"].measured
    page_texts = rendered["pdf"]["page_texts"]
    label = texts.header[0].split()[0]  # first word of the row-label header
    for number, text in enumerate(page_texts, start=1):
        assert label in text
        assert f"Page {number} of {plan.pages}" in text


def test_title_and_summary_are_on_the_first_page_only(rendered):
    plan, model = rendered["plan"], rendered["model"]
    page_texts = rendered["pdf"]["page_texts"]
    assert model.title in page_texts[0]
    assert plan.page.name in page_texts[0]
    assert all(plan.page.name not in text for text in page_texts[1:])


# --- Specific behaviour -----------------------------------------------------------

def test_wide_table_prints_its_bands_as_labelled_parts(tmp_path, captured_tables):
    model = wide(tmp_path)
    result = plan_layout(model)
    assert len(result.plan.bands) >= 2
    out = tmp_path / "wide.pdf"
    render.render_pdf(model, result, str(out))
    page_texts = read_pdf(out)["page_texts"]
    parts = len(result.plan.bands)
    assert f"Part 1 of {parts}" in page_texts[0]
    assert f"Part {parts} of {parts}" in page_texts[-1]
    # The row label column (values ROW-...) appears on every page, in every part.
    assert all("ROW-" in text for text in page_texts)


def test_numbers_print_as_excel_shows_them(tmp_path):
    def formats(ws):
        ws["B2"].number_format = "#,##0.00"
        ws["C2"].number_format = "0%"

    model = _book(tmp_path, "numbers", [["Count", "Amount", "Share"], [23.0, 1234.5, 0.25]], formats)
    out = tmp_path / "numbers.pdf"
    render.render_pdf(model, plan_layout(model), str(out))
    text = read_pdf(out)["text"]
    assert "1,234.50" in text and "25%" in text


def test_defect_6e_whole_numbers_print_without_decimals(tmp_path):
    # Regression test: before Phase 2b, 23 printed as "23.0".
    out = tmp_path / "out.pdf"
    model = load_sheet(SAMPLE_XLSX)
    render.render_pdf(model, plan_layout(model), str(out))
    text = read_pdf(out)["text"]
    assert "23" in text and "23.0" not in text


def test_defect_6d_column_widths_follow_the_content(tmp_path):
    # Regression test: before Phase 2b every column got the same width.
    plan = plan_layout(load_sheet(SAMPLE_XLSX)).plan
    widths = plan.bands[0].widths
    assert len(set(widths)) > 1


def test_numeric_columns_are_right_aligned():
    model = load_sheet(SAMPLE_XLSX)
    # Columns: Name, Age, City, Notes - only Age holds numbers.
    assert render.numeric_columns(model) == [False, True, False, False]


def test_hidden_rows_and_columns_are_left_out(tmp_path, captured_tables):
    def hide(ws):
        ws.column_dimensions["B"].hidden = True
        ws.row_dimensions[3].hidden = True

    model = _book(tmp_path, "hidden", [["Name", "Secret", "City"], ["Sam", "s1", "Nairobi"],
                                       ["Hidden row", "s2", "Nowhere"]], hide)
    out = tmp_path / "out.pdf"
    render.render_pdf(model, plan_layout(model), str(out))
    assert captured_tables[0]["data"] == [["Name", "City"], ["Sam", "Nairobi"]]
    text = read_pdf(out)["text"]
    assert "Secret" not in text and "Hidden row" not in text


def test_pdf_is_byte_identical_across_runs_in_invariant_mode(tmp_path, deterministic_pdfs):
    model = load_sheet(SAMPLE_XLSX)
    first, second = tmp_path / "a.pdf", tmp_path / "b.pdf"
    render.render_pdf(model, plan_layout(model), str(first))
    # Without invariant mode the PDF embeds a timestamp to the second; the
    # pause makes sure this test would catch that.
    time.sleep(1.1)
    render.render_pdf(model, plan_layout(model), str(second))
    assert first.read_bytes() == second.read_bytes()


def test_build_pdf_plans_and_renders_from_a_file(tmp_path):
    result = render.build_pdf(SAMPLE_XLSX, str(tmp_path / "out.pdf"), paper="A3")
    assert result.plan.page.paper == "A3"
    assert read_pdf(tmp_path / "out.pdf")["sizes"][0] == pytest.approx(
        (result.plan.page.width, result.plan.page.height))


# --- Backwards compatibility --------------------------------------------------------

class TestExcelToPdfWrapper:
    def test_old_call_still_makes_a_pdf_and_warns_that_layout_is_automatic(self, tmp_path):
        out = tmp_path / "old.pdf"
        with pytest.warns(DeprecationWarning, match="chooses the layout"):
            excel_to_pdf.build_pdf(SAMPLE_XLSX, str(out), "A4 Portrait", 63)
        assert read_pdf(out)["pages"] == 1

    def test_call_without_layout_arguments_does_not_warn(self, tmp_path):
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            excel_to_pdf.build_pdf(SAMPLE_XLSX, str(tmp_path / "new.pdf"))

    def test_old_helpers_are_still_importable(self):
        from excel_to_pdf import PAGE_SIZES, extract_sheet_data, get_page_size

        assert get_page_size("A4 Portrait") == PAGE_SIZES["A4 Portrait"]
        assert callable(extract_sheet_data)


def test_defect_6c_a_table_too_wide_for_the_page_prints_usable_columns(tmp_path, captured_tables):
    # Regression test: before Phase 2b it printed "Layout: None" with zero-width columns.
    model = wide(tmp_path)
    out = tmp_path / "wide.pdf"
    render.render_pdf(model, plan_layout(model), str(out))
    assert "None" not in read_pdf(out)["text"]
    assert all(width > 0 for table in captured_tables for width in table["col_widths"])


# --- Several sheets in one PDF ------------------------------------------------------

from pypdf import PdfReader  # noqa: E402

from intelliconverter.reader import load_workbook_sheets  # noqa: E402


@pytest.fixture
def mixed_document(tmp_path):
    """A wide sheet (landscape), a tall narrow sheet (portrait) and a tiny one."""
    wb = openpyxl.Workbook()
    wide_ws = wb.active
    wide_ws.title = "Wide"
    wide_ws.append(["Account"] + [f"Period {p:02d}" for p in range(1, 12)])
    for a in range(40):
        wide_ws.append([f"ACC-{a:04d}"] + [1234567.89 + a * p for p in range(1, 12)])
    tall_ws = wb.create_sheet("Tall")
    tall_ws.append(["Item"])
    for i in range(300):
        tall_ws.append([f"Inventory item {i}"])
    tiny = wb.create_sheet("Tiny")
    tiny.append(["Note"])
    tiny.append(["Done"])
    path = tmp_path / "mixed.xlsx"
    wb.save(path)

    sheets = load_workbook_sheets(path)
    sections = [(model, plan_layout(model)) for model in sheets.models]
    out = tmp_path / "mixed.pdf"
    render.render_document(sections, str(out))
    return {"sections": sections, "pdf": read_pdf(out), "path": out}


def test_each_sheet_is_planned_independently(mixed_document):
    pages = [result.plan.page for _, result in mixed_document["sections"]]
    assert pages[0].orientation == "Landscape"
    assert pages[1].orientation == "Portrait"


def test_each_page_has_its_own_sheets_paper_size(mixed_document):
    expected = []
    for _, result in mixed_document["sections"]:
        page = result.plan.page
        expected += [(page.width, page.height)] * result.plan.pages
    sizes = mixed_document["pdf"]["sizes"]
    assert len(sizes) == len(expected)
    assert all(a == pytest.approx(b) for a, b in zip(sizes, expected))


def test_each_sheet_starts_a_new_page_with_its_own_title(mixed_document):
    page_texts = mixed_document["pdf"]["page_texts"]
    first = 0
    for model, result in mixed_document["sections"]:
        assert model.title in page_texts[first]
        assert result.plan.page.name in page_texts[first]
        first += result.plan.pages


def test_footer_numbers_pages_per_sheet_and_for_the_document(mixed_document):
    page_texts = mixed_document["pdf"]["page_texts"]
    total = len(page_texts)
    document_page = 0
    for model, result in mixed_document["sections"]:
        for sheet_page in range(1, result.plan.pages + 1):
            text = page_texts[document_page]
            document_page += 1
            assert f"{model.title}  |  Page {sheet_page} of {result.plan.pages}" in text
            assert f"Document page {document_page} of {total}" in text


def test_bookmarks_point_to_each_sheets_first_page(mixed_document):
    reader = PdfReader(str(mixed_document["path"]))
    marks = [(entry.title, reader.get_destination_page_number(entry) + 1) for entry in reader.outline]
    expected, first = [], 1
    for model, result in mixed_document["sections"]:
        expected.append((model.title, first))
        first += result.plan.pages
    assert marks == expected


def test_a_single_sheet_footer_has_no_document_page_count(tmp_path):
    model = load_sheet(SAMPLE_XLSX)
    out = tmp_path / "one.pdf"
    render.render_document([(model, plan_layout(model))], str(out))
    text = read_pdf(out)["text"]
    assert "Page 1 of 1" in text and "Document page" not in text

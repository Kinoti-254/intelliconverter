"""Reference sheets with the layout a person would expect.

These expectations were written BEFORE the layout engine existed and
before its cost weights were tuned, so the weights were tuned to meet
them, not the other way round. Each expectation states only what matters
for readability (paper, bands, minimum font), not every detail.

If a weight change breaks one of these, the change is wrong or the
expectation needs an explicit, reviewed update.
"""
import openpyxl
import pytest

from intelliconverter.planner import LayoutError, plan_layout
from intelliconverter.reader import load_sheet
from tests.conftest import SAMPLE_XLSX

LOREM = (
    "Customer asked for the delivery to be moved to the afternoon because the "
    "loading bay is closed in the morning for maintenance work this month"
)


def _save(tmp_path, name, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = tmp_path / f"{name}.xlsx"
    wb.save(path)
    return load_sheet(path)


def _plan(model):
    return plan_layout(model).plan


# Each builder returns a SheetModel. Values are deterministic.

def sample(tmp_path):
    return load_sheet(SAMPLE_XLSX)


def contacts_tall(tmp_path):
    rows = [["Name", "Phone", "City"]]
    rows += [[f"Person {i:04d} Kamau", f"+254 7{i:08d}", "Nairobi"] for i in range(300)]
    return _save(tmp_path, "contacts", rows)


def sales_10_columns(tmp_path):
    header = ["Region", "Product", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total", "Owner"]
    rows = [header]
    rows += [[f"Region {i}", f"Item {i:03d}", *[1000 + i * m for m in range(1, 7)], 9999 + i, f"Agent {i}"]
             for i in range(25)]
    return _save(tmp_path, "sales", rows)


def wide_30_columns(tmp_path):
    rows = [["ID"] + [f"Metric {c:02d}" for c in range(1, 30)]]
    rows += [[f"ROW-{r:03d}"] + [f"val{r:02d}-{c:02d}" for c in range(1, 30)] for r in range(20)]
    return _save(tmp_path, "wide", rows)


def notes_long_text(tmp_path):
    rows = [["ID", "Customer", "Date", "Notes"]]
    rows += [[i, f"Customer {i}", "2026-11-08", LOREM] for i in range(30)]
    return _save(tmp_path, "notes", rows)


def single_column_list(tmp_path):
    rows = [["Item"]] + [[f"Inventory item number {i}"] for i in range(1000)]
    return _save(tmp_path, "list", rows)


def financial_12_columns(tmp_path):
    rows = [["Account"] + [f"Period {p:02d}" for p in range(1, 12)]]
    rows += [[f"ACC-{a:04d}"] + [1234567.89 + a * p for p in range(1, 12)] for a in range(40)]
    return _save(tmp_path, "financial", rows)


def narrow_tall_list(tmp_path):
    # Added in Phase 2b, written before running: should favour portrait on its own.
    rows = [["Code", "Description"]] + [[f"C{i:05d}", f"Stock line {i}"] for i in range(2000)]
    return _save(tmp_path, "narrow", rows)


def wide_20_columns(tmp_path):
    # Added in Phase 2b, written before running: chose A3 in Phase 2a; A4 with
    # column bands must win now.
    rows = [[f"Header {c}" for c in range(20)]]
    rows += [[r * 20 + c if c % 2 else f"text {r}-{c}" for c in range(20)] for r in range(400)]
    return _save(tmp_path, "wide20", rows)


def long_unbreakable_words(tmp_path):
    url = "https://example.com/" + "segment/" * 12 + "end"
    rows = [["Name", "Link"]] + [[f"Page {i}", url] for i in range(10)]
    return _save(tmp_path, "urls", rows)


# (builder, expected paper names, expected orientation or None, min font, bands rule)
REFERENCES = [
    (sample, {"A4"}, "Portrait", 10.0, "single"),
    (contacts_tall, {"A4"}, "Portrait", 9.0, "single"),
    (sales_10_columns, {"A4"}, None, 9.0, "single"),
    (wide_30_columns, {"A4"}, "Landscape", 9.0, "multiple"),
    (notes_long_text, {"A4"}, None, 9.0, "single"),
    (single_column_list, {"A4"}, "Portrait", 10.0, "single"),
    (financial_12_columns, {"A4"}, "Landscape", 9.0, "single"),
    (long_unbreakable_words, {"A4"}, None, 9.0, "single"),
    (narrow_tall_list, {"A4"}, "Portrait", 9.0, "single"),
    (wide_20_columns, {"A4"}, "Landscape", 9.0, "multiple"),
]


@pytest.mark.parametrize(
    "builder, papers, orientation, min_font, bands_rule",
    REFERENCES,
    ids=[ref[0].__name__ for ref in REFERENCES],
)
def test_reference_layout(tmp_path, builder, papers, orientation, min_font, bands_rule):
    plan = _plan(builder(tmp_path))
    assert plan.page.paper in papers
    if orientation:
        assert plan.page.orientation == orientation
    assert plan.font_size >= min_font
    if bands_rule == "single":
        assert len(plan.bands) == 1
    else:
        assert len(plan.bands) >= 2


def test_wide_table_repeats_the_first_column_in_every_band(tmp_path):
    plan = _plan(wide_30_columns(tmp_path))
    assert all(band.columns[0] == 0 for band in plan.bands)
    # Every other column appears in exactly one band.
    shown = sorted(c for band in plan.bands for c in band.columns[1:])
    assert shown == list(range(1, 30))


def test_long_notes_get_the_widest_column_and_wrap(tmp_path):
    plan = _plan(notes_long_text(tmp_path))
    widths = plan.bands[0].widths
    assert widths[3] == max(widths)
    assert plan.bands[0].body_lines[0][3] > 1


def test_tall_list_spans_several_pages(tmp_path):
    plan = _plan(single_column_list(tmp_path))
    assert plan.pages > 1


def test_a_cell_too_long_for_any_page_is_a_friendly_error(tmp_path):
    model = _save(tmp_path, "huge", [["Text"], ["word " * 6000]])
    with pytest.raises(LayoutError) as excinfo:
        plan_layout(model)
    assert "too long to fit on one page" in str(excinfo.value)

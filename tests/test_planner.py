"""Unit tests for intelliconverter.planner: pages, column widths, bands, pagination, cost."""
import openpyxl
import pytest
from reportlab.lib.pagesizes import A3, A4, LEGAL, LETTER

from intelliconverter import planner
from intelliconverter.planner import (
    allocate_widths,
    candidate_pages,
    paginate,
    plan_layout,
    split_into_bands,
)
from intelliconverter.reader import load_sheet
from tests.conftest import SAMPLE_XLSX


def _model(save_workbook, rows):
    wb = openpyxl.Workbook()
    for row in rows:
        wb.active.append(row)
    return load_sheet(save_workbook(wb))


class TestCandidatePages:
    def test_four_papers_in_both_orientations(self):
        pages = candidate_pages()
        assert len(pages) == 8
        sizes = {(p.paper, p.orientation): (p.width, p.height) for p in pages}
        assert sizes[("A4", "Portrait")] == pytest.approx(A4)
        assert sizes[("A4", "Landscape")] == pytest.approx((A4[1], A4[0]))
        assert sizes[("Letter", "Portrait")] == pytest.approx(LETTER)
        assert sizes[("Legal", "Landscape")] == pytest.approx((LEGAL[1], LEGAL[0]))
        assert sizes[("A3", "Portrait")] == pytest.approx(A3)

    def test_names(self):
        assert candidate_pages()[0].name == "A4 Portrait"

    def test_font_sizes_run_from_11_down_to_the_7pt_floor(self):
        assert planner.FONT_SIZES[0] == 11.0
        assert planner.FONT_SIZES[-1] == planner.MIN_FONT == 7.0
        assert all(a - b == 0.5 for a, b in zip(planner.FONT_SIZES, planner.FONT_SIZES[1:]))


class TestAllocateWidths:
    def test_natural_widths_are_kept_when_they_fit(self):
        assert allocate_widths([100, 50], [30, 30], usable=200) == [100, 50]

    def test_widest_columns_shrink_first_when_space_is_short(self):
        widths = allocate_widths([300, 60, 50], [40, 40, 40], usable=250)
        assert widths[1:] == [60, 50]
        assert sum(widths) == pytest.approx(250)

    def test_no_column_goes_below_its_minimum(self):
        widths = allocate_widths([300, 300], [100, 40], usable=180)
        assert widths[0] >= 100 and widths[1] >= 40
        assert sum(widths) == pytest.approx(180)

    def test_returns_none_when_even_minimums_do_not_fit(self):
        assert allocate_widths([300, 300], [100, 100], usable=150) is None


class TestSplitIntoBands:
    def test_one_band_when_everything_fits(self):
        bands = split_into_bands([50, 50, 50], [30, 30, 30], usable=200)
        assert [b.columns for b in bands] == [(0, 1, 2)]

    def test_first_column_repeats_in_every_band(self):
        bands = split_into_bands([60] * 7, [50] * 7, usable=200)
        assert all(band.columns[0] == 0 for band in bands)
        assert sorted(c for band in bands for c in band.columns[1:]) == list(range(1, 7))
        assert all(sum(band.widths) <= 200 + 1e-6 for band in bands)

    def test_columns_are_balanced_across_the_fewest_bands(self):
        # Filling band 1 to the brim would give 4 + 2 columns; balanced is 3 + 3.
        bands = split_into_bands([50] * 7, [40] * 7, usable=220)
        assert [band.columns for band in bands] == [(0, 1, 2, 3), (0, 4, 5, 6)]

    def test_balancing_never_adds_a_band(self):
        bands = split_into_bands([60] * 7, [50] * 7, usable=200)
        assert len(bands) == 2

    def test_returns_none_when_one_column_plus_the_label_cannot_fit(self):
        assert split_into_bands([60, 300], [60, 250], usable=200) is None


class TestPaginate:
    def test_header_repeats_and_first_page_holds_fewer_rows(self):
        # Body area 100 pt, title takes 30 pt on page 1, header 10 pt, rows 20 pt.
        starts = paginate([20] * 10, header_height=10, body_height=100, first_page_offset=30)
        # Page 1: (100-30-10)/20 = 3 rows; then (100-10)/20 = 4 rows per page.
        assert starts == [0, 3, 7]

    def test_returns_none_when_a_row_is_taller_than_a_page(self):
        assert paginate([20, 500], header_height=10, body_height=100, first_page_offset=0) is None


def test_plan_for_the_sample(save_workbook):
    result = plan_layout(load_sheet(SAMPLE_XLSX))
    plan = result.plan
    assert plan.page.name == "A4 Portrait"
    assert plan.pages == 1
    assert len(plan.bands) == 1
    band = plan.bands[0]
    assert band.columns == (0, 1, 2, 3)
    assert sum(band.widths) <= plan.page.width - planner.MARGIN_LEFT - planner.MARGIN_RIGHT


def test_candidates_are_ranked_and_the_plan_is_the_cheapest_feasible(save_workbook):
    result = plan_layout(load_sheet(SAMPLE_XLSX))
    feasible = [c for c in result.candidates if c.feasible]
    totals = [c.total for c in feasible]
    assert totals == sorted(totals)
    best = feasible[0]
    assert (best.page, best.font_size) == (result.plan.page, result.plan.font_size)
    assert set(best.cost) == set(planner.WEIGHTS)


def test_planning_is_deterministic(save_workbook):
    model = _model(save_workbook, [["A", "B"]] + [[f"row {i}", "x" * (i % 40)] for i in range(200)])
    first, second = plan_layout(model), plan_layout(model)
    assert first.plan == second.plan
    assert [c.total for c in first.candidates] == [c.total for c in second.candidates]


def test_plan_wraps_with_the_widths_it_chose(save_workbook):
    from intelliconverter.textmetrics import BODY_FONT, count_lines

    text = "word " * 60
    model = _model(save_workbook, [["Notes"], [text.strip()]])
    plan = plan_layout(model).plan
    band = plan.bands[0]
    content_width = band.widths[0] - 2 * planner.PAD_X
    assert band.body_lines[0][0] == count_lines(text.strip(), content_width, BODY_FONT, plan.font_size)


def test_only_visible_columns_are_planned(save_workbook):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "Hidden", "C"])
    ws.append(["a", "h", "c"])
    ws.column_dimensions["B"].hidden = True
    plan = plan_layout(load_sheet(save_workbook(wb))).plan
    assert plan.column_letters == ("A", "C")


def test_a_header_only_sheet_can_be_planned(save_workbook):
    plan = plan_layout(_model(save_workbook, [["Only", "Header"]])).plan
    assert plan.pages == 1


class TestPhase2bRules:
    def test_no_cost_part_depends_on_orientation(self):
        assert "landscape" not in planner.WEIGHTS
        assert "orientation" not in " ".join(planner.WEIGHTS)

    def test_a4_and_letter_are_practical_papers_legal_and_a3_are_last_resorts(self):
        cost = planner.PAPER_COST
        assert cost["A4"] == 0 and cost["Letter"] <= 1
        # More than an extra column band, so A4 with bands beats A3 or Legal.
        assert min(cost["Legal"], cost["A3"]) * planner.WEIGHTS["paper"] > planner.WEIGHTS["column_bands"]

    def test_a_page_smaller_than_a4_still_counts_as_one_sheet(self):
        a4 = planner.Candidate(candidate_pages()[0], 10, True, "", (), 3, {}, 0)
        letter = planner.Candidate(candidate_pages()[2], 10, True, "", (), 3, {}, 0)
        a3 = planner.Candidate(candidate_pages()[6], 10, True, "", (), 3, {}, 0)
        assert planner._a4_sheets(a4) == planner._a4_sheets(letter) == 3
        assert planner._a4_sheets(a3) == pytest.approx(6)

    def test_ties_go_to_fewer_pages_then_less_wrapping_then_less_empty_space(self):
        def candidate(pages, wrapping, empty):
            cost = dict.fromkeys(planner.WEIGHTS, 0.0)
            cost.update(wrapping=wrapping, empty_width=empty)
            return planner.Candidate(candidate_pages()[0], 10, True, "", (), pages, cost, 1.0)

        assert planner._rank_key(candidate(2, 0, 0), 9) < planner._rank_key(candidate(3, 0, 0), 0)
        assert planner._rank_key(candidate(2, 0.1, 0), 9) < planner._rank_key(candidate(2, 0.2, 0), 0)
        assert planner._rank_key(candidate(2, 0, 0.1), 9) < planner._rank_key(candidate(2, 0, 0.2), 0)
        # Only when every measured factor is equal does candidate order decide.
        assert planner._rank_key(candidate(2, 0, 0), 0) < planner._rank_key(candidate(2, 0, 0), 1)


class TestPaperChoice:
    def test_a_chosen_paper_limits_the_candidates(self):
        result = plan_layout(load_sheet(SAMPLE_XLSX), paper="A3")
        assert result.plan.page.paper == "A3"
        assert {c.page.paper for c in result.candidates} == {"A3"}

    def test_both_orientations_are_still_compared(self):
        result = plan_layout(load_sheet(SAMPLE_XLSX), paper="Letter")
        assert {c.page.orientation for c in result.candidates} == {"Portrait", "Landscape"}

    def test_unknown_paper_is_rejected(self):
        with pytest.raises(ValueError):
            plan_layout(load_sheet(SAMPLE_XLSX), paper="B5")


def test_defect_4_long_text_gets_a_wider_column_than_short_text(save_workbook):
    # Regression test: before Phase 2b cell content was ignored when sizing columns.
    model = _model(save_workbook, [["Code", "Description"],
                                   ["A1", "Nairobi Regional Sales Office, second floor"]])
    widths = plan_layout(model).plan.bands[0].widths
    assert widths[1] > 2 * widths[0]

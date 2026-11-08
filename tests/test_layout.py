"""Characterisation tests for intelliconverter.layout.

The scoring logic is pinned exactly as it is today, including behaviour the
assessment flagged for replacement in Phase 2 (marked "current behaviour").
When the new cost model lands, these tests are expected to be rewritten.
"""
import pytest

from intelliconverter.layout import LAYOUTS, score_layout


def test_layout_candidates_are_unchanged():
    assert LAYOUTS == [
        ("A4 Portrait", 7.7),
        ("A4 Landscape", 10.7),
        ("A3 Portrait", 11.7),
        ("Legal Landscape", 13.2),
        ("A3 Landscape", 16.5),
    ]


class TestScoreLayout:
    def test_wider_than_page_is_disqualified(self):
        assert score_layout(10.8, 10.7) == -1

    def test_exactly_page_width_fits_with_top_score(self):
        assert score_layout(10.7, 10.7) == 100.0

    @pytest.mark.parametrize(
        "fill, expected",
        [
            (0.10, 6.0),
            (0.50, 30.0),
            (0.749, 44.9),
        ],
    )
    def test_below_75_percent_fill_scores_fill_times_60(self, fill, expected):
        assert score_layout(fill, 1.0) == expected

    @pytest.mark.parametrize(
        "fill, expected",
        [
            (0.75, 87.5),
            (0.90, 95.0),
            (1.00, 100.0),
        ],
    )
    def test_from_75_percent_fill_scores_50_plus_fill_times_50(self, fill, expected):
        assert score_layout(fill, 1.0) == expected

    def test_current_behaviour_jumps_43_points_at_75_percent(self):
        # Current behaviour, to be replaced in Phase 2: a 0.1% change in
        # width moves the score from 44.9 to 87.5.
        assert score_layout(0.75, 1.0) - score_layout(0.749, 1.0) == pytest.approx(42.6)

    def test_scores_are_rounded_to_one_decimal(self):
        assert score_layout(1, 3) == 20.0
        assert score_layout(0.8333, 1.0) == 91.7

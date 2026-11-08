"""Tests for intelliconverter.textmetrics: measuring and wrapping text in the PDF font."""
import random
import string

import pytest
from reportlab.pdfbase.pdfmetrics import stringWidth

from intelliconverter.textmetrics import (
    BODY_FONT,
    HEADER_FONT,
    cell_text,
    count_lines,
    longest_word_width,
    text_width,
    wrap_text,
)


def test_cell_text_matches_what_the_pdf_prints():
    assert cell_text(None) == ""
    assert cell_text(23.0) == "23.0"
    assert cell_text("Nairobi") == "Nairobi"


def test_text_width_matches_reportlab():
    for font in (BODY_FONT, HEADER_FONT):
        assert text_width("Likes coffee", font, 9) == pytest.approx(stringWidth("Likes coffee", font, 9))


def test_short_text_stays_on_one_line():
    assert wrap_text("Likes coffee", 200, BODY_FONT, 10) == ["Likes coffee"]


def test_wraps_at_spaces_and_keeps_every_word_in_order():
    text = "the quick brown fox jumps over the lazy dog"
    lines = wrap_text(text, 60, BODY_FONT, 10)
    assert len(lines) > 1
    assert " ".join(lines) == text
    assert all(text_width(line, BODY_FONT, 10) <= 60 for line in lines)


def test_a_word_wider_than_the_column_is_broken_into_pieces():
    url = "https://example.com/a/very/long/path"
    lines = wrap_text(url, 40, BODY_FONT, 10)
    assert "".join(lines) == url
    assert all(text_width(line, BODY_FONT, 10) <= 40 for line in lines)


def test_line_breaks_in_the_cell_are_kept():
    assert wrap_text("first\nsecond", 500, BODY_FONT, 10) == ["first", "second"]


def test_empty_text_is_one_empty_line():
    assert wrap_text("", 50, BODY_FONT, 10) == [""]
    assert count_lines("", 50, BODY_FONT, 10) == 1


def test_a_column_narrower_than_one_character_still_progresses():
    assert wrap_text("abc", 0.5, BODY_FONT, 10) == ["a", "b", "c"]


def test_count_lines_always_agrees_with_wrap_text():
    random.seed(7)
    alphabet = string.ascii_letters + string.digits + "  .,-"
    for _ in range(500):
        text = "".join(random.choice(alphabet) for _ in range(random.randint(0, 120)))
        width = random.uniform(5, 200)
        size = random.choice([7, 8.5, 10, 11])
        font = random.choice([BODY_FONT, HEADER_FONT])
        assert count_lines(text, width, font, size) == len(wrap_text(text, width, font, size))


def test_longest_word_width():
    assert longest_word_width("a bb cccc", BODY_FONT, 10) == pytest.approx(
        text_width("cccc", BODY_FONT, 10)
    )
    assert longest_word_width("", BODY_FONT, 10) == 0

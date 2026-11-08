"""Measuring and wrapping text in the PDF font.

The PDF uses ReportLab's built-in Helvetica, so text is measured in exactly
that font. Helvetica widths add up (no kerning) and scale linearly with the
font size, so each word is measured once at 1 pt and reused for every
candidate size. wrap_text() is also what the PDF will print, so the plan's
line counts match the PDF exactly.
"""
from functools import lru_cache

from reportlab.pdfbase.pdfmetrics import stringWidth

BODY_FONT = "Helvetica"
HEADER_FONT = "Helvetica-Bold"


def cell_text(value):
    """The text the PDF prints for a cell value (None becomes empty)."""
    return str(value) if value is not None else ""


@lru_cache(maxsize=500_000)
def _width_at_1pt(text, font):
    return stringWidth(text, font, 1)


def text_width(text, font, size):
    """Width of `text` in points."""
    return _width_at_1pt(text, font) * size


@lru_cache(maxsize=200_000)
def _paragraph_words(text, font):
    """Split text into paragraphs (at line breaks) of (word, width at 1 pt)."""
    return tuple(
        tuple((word, _width_at_1pt(word, font)) for word in paragraph.split())
        for paragraph in text.split("\n")
    )


def longest_word_width(text, font, size):
    """Width of the widest single word in `text`, in points."""
    widths = [width for paragraph in _paragraph_words(text, font) for _, width in paragraph]
    return max(widths, default=0) * size


def wrap_text(text, max_width, font, size):
    """Break text into lines no wider than max_width points.

    Lines break at spaces. A word wider than the column is broken between
    characters. Line breaks inside the cell are kept. Empty text is one line.
    """
    lines = []
    _wrap(text, max_width, font, size, lines.append)
    return lines


def count_lines(text, max_width, font, size):
    """Number of lines wrap_text() would produce, without building them."""
    counter = [0]

    def add(_line):
        counter[0] += 1

    if "\n" not in text and _width_at_1pt(text, font) * size <= max_width:
        return 1  # fast path: most cells fit on one line
    _wrap(text, max_width, font, size, add)
    return counter[0]


def _wrap(text, max_width, font, size, emit):
    limit = max_width / size  # compare widths at 1 pt
    space = _width_at_1pt(" ", font)
    for paragraph in _paragraph_words(text, font):
        line_words = []
        line_width = 0.0
        for word, width in paragraph:
            needed = width if not line_words else line_width + space + width
            if needed <= limit:
                line_words.append(word)
                line_width = needed
                continue
            if line_words:
                emit(" ".join(line_words))
            if width <= limit:
                line_words, line_width = [word], width
            else:
                pieces = _break_word(word, limit, font)
                for piece in pieces[:-1]:
                    emit(piece)
                last = pieces[-1]
                line_words, line_width = [last], _width_at_1pt(last, font)
        emit(" ".join(line_words))


def _break_word(word, limit, font):
    """Split a word that is wider than the column into pieces that fit.

    Each piece holds at least one character, so narrow columns still progress.
    """
    pieces = []
    current = ""
    current_width = 0.0
    for char in word:
        char_width = _width_at_1pt(char, font)
        if current and current_width + char_width > limit:
            pieces.append(current)
            current, current_width = char, char_width
        else:
            current += char
            current_width += char_width
    pieces.append(current)
    return pieces

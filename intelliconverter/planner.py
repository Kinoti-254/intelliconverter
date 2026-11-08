"""The layout engine: chooses paper, orientation, font size and column widths.

How it works:
1. Measure every visible cell once, in the PDF font at 1 pt (textmetrics).
2. For each candidate (4 papers x 2 orientations x 9 font sizes), size the
   columns from their content, split them into bands if they can't fit
   across the page, and simulate wrapping and page breaks.
3. Reject candidates that break a hard limit; give the rest a cost made of
   named parts (small font, extra pages, bands, wrapping...). The cheapest
   wins; ties go to the earlier candidate (A4 first, portrait first,
   larger font first).
4. Work out the winner exactly (every row, every page break) as a LayoutPlan.

Large sheets: candidates are compared using an evenly spaced sample of up
to SAMPLE_ROWS rows; only the winner is computed for every row.
"""
import math
from dataclasses import dataclass

from reportlab.lib.pagesizes import A3, A4, LEGAL, LETTER

from intelliconverter.errors import LayoutError
from intelliconverter.textmetrics import (
    BODY_FONT,
    HEADER_FONT,
    cell_text,
    count_lines,
    longest_word_width,
    text_width,
)

__all__ = ["LayoutError", "LayoutPlan", "plan_layout"]

# Candidates, in tie-break order
PAPERS = (("A4", A4), ("Letter", LETTER), ("Legal", LEGAL), ("A3", A3))
ORIENTATIONS = ("Portrait", "Landscape")
FONT_SIZES = tuple(11.0 - 0.5 * step for step in range(9))  # 11.0 down to 7.0
MIN_FONT = 7.0          # hard floor
COMFORT_FONT = 9.0      # smaller than this costs extra
PREFERRED_FONT = 10.0

# Page geometry in points (72 per inch), matching the PDF renderer
MARGIN_LEFT = MARGIN_RIGHT = 36.0
MARGIN_TOP = 43.2
MARGIN_BOTTOM = 36.0
FOOTER_HEIGHT = 14.0        # room for page numbers
TITLE_BLOCK_HEIGHT = 56.0   # title and summary line on the first page

# Cells
PAD_X = 6.0                 # left and right padding in each cell
PAD_Y = 4.0                 # top and bottom padding
LEADING = 1.2               # line height as a multiple of the font size
MIN_COLUMN_EM = 2.0         # narrowest column content, in font sizes (ems)
MAX_COLUMN_EM = 20.0        # widest column content; longer text wraps
WIDTH_PERCENTILE = 0.9      # one unusually long cell doesn't widen a column

SAMPLE_ROWS = 500

PAPER_COST = {"A4": 0.0, "Letter": 1.0, "Legal": 3.0, "A3": 4.0}

# Cost of each named part. Tuned against tests/test_planner_reference.py.
WEIGHTS = {
    "small_font": 8.0,       # per point below COMFORT_FONT
    "font_distance": 1.0,    # per point away from PREFERRED_FONT
    "extra_paper": 4.0,      # times (paper used / least paper of any candidate - 1),
                             # paper measured in A4 sheets so 1 A3 page counts as 2
    "column_bands": 10.0,    # per band beyond the first
    "wrapping": 6.0,         # times the share of cells that wrap
    "empty_width": 2.0,      # times the share of the page width left empty
    "paper": 1.0,            # times PAPER_COST (A4 is standard here)
    "landscape": 0.5,        # portrait is preferred when all else is equal
    "short_last_page": 1.0,  # a last page holding only one or two rows
}


@dataclass(frozen=True)
class Page:
    paper: str
    orientation: str
    width: float
    height: float

    @property
    def name(self):
        return f"{self.paper} {self.orientation}"

    @property
    def usable_width(self):
        return self.width - MARGIN_LEFT - MARGIN_RIGHT

    @property
    def body_height(self):
        return self.height - MARGIN_TOP - MARGIN_BOTTOM - FOOTER_HEIGHT


@dataclass(frozen=True)
class Band:
    columns: tuple   # positions among the visible columns; 0 is the label column
    widths: tuple    # column widths in points, padding included


@dataclass(frozen=True)
class BandPlan:
    columns: tuple
    widths: tuple
    header_lines: tuple   # line count per header cell
    header_height: float
    body_lines: tuple     # line count per body cell: body_lines[row][position in band]
    row_heights: tuple
    page_starts: tuple    # body row index where each page of this band starts


@dataclass(frozen=True)
class LayoutPlan:
    page: Page
    font_size: float
    column_letters: tuple
    bands: tuple
    pages: int
    cost: dict
    total_cost: float


@dataclass(frozen=True)
class Candidate:
    page: Page
    font_size: float
    feasible: bool
    reason: str           # why it was rejected, when not feasible
    bands: tuple
    pages: int            # estimated when the sheet is sampled
    cost: dict
    total: float


@dataclass(frozen=True)
class PlanResult:
    plan: LayoutPlan
    candidates: tuple     # feasible ones first, cheapest first
    sampled: bool         # True when candidates' page counts are estimates


@dataclass(frozen=True)
class MeasuredSheet:
    letters: tuple
    header: tuple         # header row texts
    body: tuple           # body rows as texts
    natural_em: tuple     # preferred content width per column, at 1 pt
    min_em: tuple         # narrowest acceptable content width per column, at 1 pt


# --- Building blocks ---------------------------------------------------------

def candidate_pages():
    return [
        Page(paper, orientation, *(size if orientation == "Portrait" else (size[1], size[0])))
        for paper, size in PAPERS
        for orientation in ORIENTATIONS
    ]


def allocate_widths(natural, minimum, usable):
    """Fit columns into `usable` points.

    Natural widths are kept when they fit. Otherwise the widest columns are
    narrowed first (to a common cap), never below their minimum. Returns
    None when even the minimum widths don't fit.
    """
    if sum(natural) <= usable:
        return list(natural)
    if sum(minimum) > usable + 1e-9:
        return None

    def total(cap):
        return sum(max(low, min(high, cap)) for high, low in zip(natural, minimum))

    low, high = 0.0, max(natural)
    for _ in range(60):
        cap = (low + high) / 2
        if total(cap) > usable:
            high = cap
        else:
            low = cap
    return [max(lo, min(hi, low)) for hi, lo in zip(natural, minimum)]


def split_into_bands(natural, minimum, usable):
    """Split columns into bands that each fit across the page.

    Column 0 (the row label) is repeated at the start of every band. Returns
    None when the label plus any single column can't fit.
    """
    every = list(range(len(natural)))
    widths = allocate_widths(natural, minimum, usable)
    if widths is not None:
        return [Band(tuple(every), tuple(widths))]
    if len(natural) == 1:
        return None

    groups = []
    current = [0]
    for column in every[1:]:
        trial = current + [column]
        if sum(minimum[c] for c in trial) <= usable:
            current = trial
            continue
        if current == [0]:
            return None
        groups.append(current)
        current = [0, column]
        if sum(minimum[c] for c in current) > usable:
            return None
    groups.append(current)

    return [
        Band(tuple(group), tuple(allocate_widths(
            [natural[c] for c in group], [minimum[c] for c in group], usable)))
        for group in groups
    ]


def paginate(row_heights, header_height, body_height, first_page_offset):
    """Page start indices for rows that may not be split across pages.

    The header repeats on every page; the first page loses
    first_page_offset points (the title). Returns None if a row can't fit.
    """
    starts = [0]
    capacity = body_height - first_page_offset - header_height
    used = 0.0
    for index, height in enumerate(row_heights):
        if used and used + height > capacity:
            starts.append(index)
            used = 0.0
            capacity = body_height - header_height
        if height > capacity:
            return None
        used += height
    return starts


# --- Measuring ---------------------------------------------------------------

def measure_sheet(model):
    rows = [[cell_text(value) for value in row] for row in model.visible_rows()]
    header, body = rows[0], rows[1:]
    natural, minimum = [], []
    for position, header_text in enumerate(header):
        column = [row[position] for row in body]
        body_width = _percentile([text_width(text, BODY_FONT, 1) for text in column])
        body_word = _percentile([longest_word_width(text, BODY_FONT, 1) for text in column])
        header_word = longest_word_width(header_text, HEADER_FONT, 1)
        header_full = text_width(header_text, HEADER_FONT, 1)

        low = _clamp(max(header_word, body_word), MIN_COLUMN_EM, MAX_COLUMN_EM)
        high = _clamp(max(body_width, header_full), low, MAX_COLUMN_EM)
        minimum.append(low)
        natural.append(high)

    letters = tuple(column.letter for column in model.visible_columns)
    return MeasuredSheet(letters, tuple(header), tuple(tuple(r) for r in body),
                         tuple(natural), tuple(minimum))


def _percentile(values):
    if not values:
        return 0.0
    ordered = sorted(values)
    return ordered[min(len(ordered) - 1, math.ceil(WIDTH_PERCENTILE * len(ordered)) - 1)]


def _clamp(value, low, high):
    return max(low, min(high, value))


# --- Simulating one candidate -------------------------------------------------

def _line_counts(texts, widths, font, size):
    return tuple(count_lines(text, width - 2 * PAD_X, font, size) for text, width in zip(texts, widths))


def _height(lines, size):
    return max(lines, default=1) * LEADING * size + 2 * PAD_Y


def _bands_for(measured, page, size):
    natural = [em * size + 2 * PAD_X for em in measured.natural_em]
    minimum = [em * size + 2 * PAD_X for em in measured.min_em]
    return split_into_bands(natural, minimum, page.usable_width)


def _simulate_band(measured, band, page, size, rows, first):
    """Header height, line counts and heights for `rows` (body row indices)."""
    header_lines = _line_counts([measured.header[c] for c in band.columns], band.widths, HEADER_FONT, size)
    body_lines = [
        _line_counts([measured.body[r][c] for c in band.columns], band.widths, BODY_FONT, size)
        for r in rows
    ]
    heights = [_height(lines, size) for lines in body_lines]
    offset = TITLE_BLOCK_HEIGHT if first else 0.0
    return header_lines, _height(header_lines, size), body_lines, heights, offset


def _sample_rows(count):
    if count <= SAMPLE_ROWS:
        return list(range(count))
    step = (count - 1) / (SAMPLE_ROWS - 1)
    return sorted({round(i * step) for i in range(SAMPLE_ROWS)})


def _evaluate(measured, page, size, rows):
    bands = _bands_for(measured, page, size)
    if bands is None:
        return _rejected(page, size, "columns don't fit even when split into bands")

    exact = len(rows) == len(measured.body)
    pages = 0
    wrapped = cells = 0
    short_last = False
    for index, band in enumerate(bands):
        _, header_height, body_lines, heights, offset = _simulate_band(
            measured, band, page, size, rows, first=index == 0)
        wrapped += sum(1 for lines in body_lines for n in lines if n > 1)
        cells += sum(len(lines) for lines in body_lines)
        if exact:
            starts = paginate(heights, header_height, page.body_height, offset)
            if starts is None:
                return _rejected(page, size, "a row is taller than the page")
            pages += len(starts)
            if len(starts) > 1 and len(heights) - starts[-1] <= 2:
                short_last = True
        else:
            band_pages = _estimate_pages(heights, len(measured.body), header_height, page, offset)
            if band_pages is None:
                return _rejected(page, size, "a row is taller than the page")
            pages += band_pages

    cost = {
        "small_font": WEIGHTS["small_font"] * max(0.0, COMFORT_FONT - size),
        "font_distance": WEIGHTS["font_distance"] * abs(size - PREFERRED_FONT),
        "extra_paper": 0.0,  # filled in once every candidate is known
        "column_bands": WEIGHTS["column_bands"] * (len(bands) - 1),
        "wrapping": WEIGHTS["wrapping"] * (wrapped / cells if cells else 0.0),
        "empty_width": WEIGHTS["empty_width"] * (1 - sum(bands[0].widths) / page.usable_width),
        "paper": WEIGHTS["paper"] * PAPER_COST[page.paper],
        "landscape": WEIGHTS["landscape"] * (page.orientation == "Landscape"),
        "short_last_page": WEIGHTS["short_last_page"] * short_last,
    }
    return Candidate(page, size, True, "", tuple(bands), pages, cost, sum(cost.values()))


def _estimate_pages(sample_heights, row_count, header_height, page, offset):
    first = page.body_height - offset - header_height
    rest = page.body_height - header_height
    if max(sample_heights, default=0) > min(first, rest):
        return None
    total = sum(sample_heights) / len(sample_heights) * row_count if sample_heights else 0.0
    if total <= first:
        return 1
    return 1 + math.ceil((total - first) / rest)


def _rejected(page, size, reason):
    return Candidate(page, size, False, reason, (), 0, {}, math.inf)


# --- The engine ----------------------------------------------------------------

def plan_layout(model):
    """Choose the most readable layout for a SheetModel. Returns a PlanResult."""
    measured = measure_sheet(model)
    rows = _sample_rows(len(measured.body))

    candidates = [
        _evaluate(measured, page, size, rows)
        for page in candidate_pages()
        for size in FONT_SIZES
    ]
    feasible = [c for c in candidates if c.feasible]
    if not feasible:
        raise LayoutError(_no_layout_message(candidates))

    least = min(_a4_sheets(c) for c in feasible)
    priced = []
    for order, candidate in enumerate(candidates):
        if candidate.feasible:
            cost = dict(candidate.cost)
            cost["extra_paper"] = WEIGHTS["extra_paper"] * (_a4_sheets(candidate) / least - 1)
            candidate = Candidate(candidate.page, candidate.font_size, True, "",
                                  candidate.bands, candidate.pages, cost, sum(cost.values()))
        priced.append((candidate.total, order, candidate))
    ranked = tuple(candidate for _, _, candidate in sorted(priced, key=lambda item: item[:2]))

    for candidate in ranked:
        if not candidate.feasible:
            break
        plan = _exact_plan(measured, candidate)
        if plan is not None:
            return PlanResult(plan, ranked, sampled=len(rows) < len(measured.body))
    raise LayoutError(_no_layout_message(candidates))


def _a4_sheets(candidate):
    """Paper used, in A4 sheets: an A3 page counts as two."""
    area = candidate.page.width * candidate.page.height
    return candidate.pages * area / (A4[0] * A4[1])


def _exact_plan(measured, candidate):
    """Every row and page break for the chosen candidate, or None if a row can't fit."""
    rows = range(len(measured.body))
    bands = []
    for index, band in enumerate(candidate.bands):
        header_lines, header_height, body_lines, heights, offset = _simulate_band(
            measured, band, candidate.page, candidate.font_size, rows, first=index == 0)
        starts = paginate(heights, header_height, candidate.page.body_height, offset)
        if starts is None:
            return None
        bands.append(BandPlan(band.columns, band.widths, header_lines, header_height,
                              tuple(body_lines), tuple(heights), tuple(starts)))
    return LayoutPlan(
        page=candidate.page,
        font_size=candidate.font_size,
        column_letters=measured.letters,
        bands=tuple(bands),
        pages=sum(len(band.page_starts) for band in bands),
        cost=candidate.cost,
        total_cost=candidate.total,
    )


def _no_layout_message(candidates):
    reasons = {c.reason for c in candidates if not c.feasible}
    if "a row is taller than the page" in reasons:
        return (
            "A cell is too long to fit on one page, even with the smallest font on "
            "A3 paper. Split very long text across several cells or rows."
        )
    return "The columns are too wide to fit on any supported page, even when split into bands."

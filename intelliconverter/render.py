"""Drawing the PDF from a LayoutPlan with ReportLab.

The renderer makes no layout decisions. It draws exactly what the plan says:
paper, font size, column widths, wrapped lines, row heights and page
breaks. Lines are wrapped with the same function the planner used, so the
PDF always matches the plan the report describes.

Each page is one table (header row plus that page's rows). A table too
wide for the page is printed as consecutive parts (bands), each starting on
a new page and repeating the first column. The first page carries a title
block; every page has a footer with the page number.
"""
import datetime as dt

from reportlab.lib import colors
from reportlab.platypus import BaseDocTemplate, Flowable, Frame, PageBreak, PageTemplate, Table, TableStyle

from intelliconverter.errors import OutputError
from intelliconverter.planner import (
    FOOTER_HEIGHT,
    LEADING,
    MARGIN_BOTTOM,
    MARGIN_LEFT,
    MARGIN_RIGHT,
    MARGIN_TOP,
    PAD_X,
    PAD_Y,
    TITLE_BLOCK_HEIGHT,
    plan_layout,
)
from intelliconverter.reader import load_sheet
from intelliconverter.textmetrics import BODY_FONT, HEADER_FONT, text_width, wrap_text

HEADER_BACKGROUND = colors.HexColor("#2C3E50")
STRIPE = colors.HexColor("#F2F3F4")
GRID = colors.HexColor("#BDC3C7")
MUTED = colors.HexColor("#555555")
TITLE_SIZE = 14
SUBTITLE_SIZE = 8.5
FOOTER_SIZE = 7.5


def build_pdf(filepath, output_path, sheet=None, paper=None):
    """Read a sheet, plan its layout and write the PDF. Returns the PlanResult."""
    model = load_sheet(filepath, sheet)
    result = plan_layout(model, paper=paper)
    render_pdf(model, result, output_path)
    return result


def render_pdf(model, result, output_path):
    """Write the PDF for `model` following result.plan exactly."""
    plan, texts = result.plan, result.measured
    page = plan.page
    doc = BaseDocTemplate(
        output_path,
        pagesize=(page.width, page.height),
        leftMargin=MARGIN_LEFT,
        rightMargin=MARGIN_RIGHT,
        topMargin=MARGIN_TOP,
        bottomMargin=MARGIN_BOTTOM + FOOTER_HEIGHT,  # the footer sits in this strip
        title=model.title,
        author="IntelliConverter",
    )
    # No frame padding: the planner's page geometry is the frame, exactly.
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="body")

    right_aligned = numeric_columns(model)
    story = [TitleBlock(model.title, _subtitle(model, plan), doc.width)]
    page_bands = []  # which band each page shows, for the footer
    for band_index, band in enumerate(plan.bands):
        bounds = list(band.page_starts) + [len(band.row_heights)]
        for start, end in zip(bounds, bounds[1:]):
            if page_bands:
                story.append(PageBreak())
            story.append(_page_table(band, start, end, texts, right_aligned, plan.font_size))
            page_bands.append(band_index)

    doc.addPageTemplates([PageTemplate(id="pages", frames=[frame], onPage=_footer(model, plan, page_bands))])
    try:
        doc.build(story)
    except OSError as error:
        reason = f" ({error.strerror})" if error.strerror else ""
        raise OutputError(
            f"Could not write {output_path}. If it is open in another program, "
            f"close it and try again.{reason}"
        ) from error


def numeric_columns(model):
    """For each visible column: True when most of its body cells are numbers or dates.

    Those columns are right-aligned, as Excel does.
    """
    rows = model.visible_rows()[1:]
    flags = []
    for position in range(len(model.visible_columns)):
        values = [row[position] for row in rows if row[position] not in (None, "")]
        numbers = sum(1 for value in values if _is_numeric(value))
        flags.append(bool(values) and numbers * 2 > len(values))
    return flags


def _is_numeric(value):
    if isinstance(value, bool):
        return False
    return isinstance(value, (int, float, dt.datetime, dt.date, dt.time))


def _page_table(band, start, end, texts, right_aligned, size):
    widths = list(band.widths)

    def cell(text, width, font):
        return "\n".join(wrap_text(text, width - 2 * PAD_X, font, size))

    header = [cell(texts.header[c], w, HEADER_FONT) for c, w in zip(band.columns, widths)]
    body = [
        [cell(texts.body[row][c], w, BODY_FONT) for c, w in zip(band.columns, widths)]
        for row in range(start, end)
    ]
    heights = [band.header_height] + list(band.row_heights[start:end])
    table = Table([header] + body, colWidths=widths, rowHeights=heights)

    style = [
        ("FONT", (0, 0), (-1, 0), HEADER_FONT, size, LEADING * size),
        ("FONT", (0, 1), (-1, -1), BODY_FONT, size, LEADING * size),
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BACKGROUND),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, STRIPE]),
        ("GRID", (0, 0), (-1, -1), 0.4, GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), PAD_X),
        ("RIGHTPADDING", (0, 0), (-1, -1), PAD_X),
        ("TOPPADDING", (0, 0), (-1, -1), PAD_Y),
        ("BOTTOMPADDING", (0, 0), (-1, -1), PAD_Y),
    ]
    for position, column in enumerate(band.columns):
        if right_aligned[column]:
            style.append(("ALIGN", (position, 1), (position, -1), "RIGHT"))
    table.setStyle(TableStyle(style))
    return table


class TitleBlock(Flowable):
    """Sheet title and a one-line summary, exactly TITLE_BLOCK_HEIGHT tall.

    Drawn as plain strings (no markup), so any sheet title prints safely.
    """

    def __init__(self, title, subtitle, width):
        super().__init__()
        self.title, self.subtitle, self.block_width = title, subtitle, width

    def wrap(self, available_width, available_height):
        return self.block_width, TITLE_BLOCK_HEIGHT

    def draw(self):
        canvas = self.canv
        canvas.setFont(HEADER_FONT, TITLE_SIZE)
        canvas.drawString(0, TITLE_BLOCK_HEIGHT - TITLE_SIZE - 4,
                          _fit(self.title, HEADER_FONT, TITLE_SIZE, self.block_width))
        canvas.setFont(BODY_FONT, SUBTITLE_SIZE)
        canvas.setFillColor(MUTED)
        canvas.drawString(0, TITLE_BLOCK_HEIGHT - TITLE_SIZE - 22,
                          _fit(self.subtitle, BODY_FONT, SUBTITLE_SIZE, self.block_width))


def _subtitle(model, plan):
    parts = [
        plan.page.name,
        f"{plan.font_size:g} pt",
        f"{model.visible_row_count:,} rows x {len(model.visible_columns)} columns",
    ]
    if len(plan.bands) > 1:
        parts.append(f"printed in {len(plan.bands)} parts to fit the page width")
    return "  |  ".join(parts)


def _footer(model, plan, page_bands):
    letters = plan.column_letters

    def draw(canvas, doc):
        band_index = page_bands[doc.page - 1] if doc.page <= len(page_bands) else 0
        parts = [model.title, f"Page {doc.page} of {plan.pages}"]
        if len(plan.bands) > 1:
            columns = plan.bands[band_index].columns[1:]
            parts.append(
                f"Part {band_index + 1} of {len(plan.bands)}: "
                f"columns {letters[columns[0]]} to {letters[columns[-1]]}"
            )
        canvas.saveState()
        canvas.setFont(BODY_FONT, FOOTER_SIZE)
        canvas.setFillColor(MUTED)
        text = _fit("  |  ".join(parts), BODY_FONT, FOOTER_SIZE, doc.width)
        canvas.drawCentredString(doc.pagesize[0] / 2, MARGIN_BOTTOM + 3, text)
        canvas.restoreState()

    return draw


def _fit(text, font, size, width):
    """Shorten text with '...' so it fits `width` points."""
    if text_width(text, font, size) <= width:
        return text
    while text and text_width(text + "...", font, size) > width:
        text = text[:-1]
    return text + "..."

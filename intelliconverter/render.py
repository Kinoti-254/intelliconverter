"""Drawing the PDF from a LayoutPlan with ReportLab.

The renderer makes no layout decisions. It draws exactly what the plan says:
paper, font size, column widths, wrapped lines, row heights and page
breaks. Lines are wrapped with the same function the planner used, so the
PDF always matches the plan the report describes.

Each page is one table (header row plus that page's rows). A table too
wide for the page is printed as consecutive parts (bands), each starting on
a new page and repeating the first column.

Several sheets can go into one PDF (render_document). Each sheet is a
section with its own plan, so pages can change size and orientation between
sheets. Each section starts on a new page with a title block and a bookmark;
every page has a footer with the sheet's page number (and, with several
sheets, the document page number).
"""
import datetime as dt

from reportlab.lib import colors
from reportlab.platypus import (
    BaseDocTemplate, Flowable, Frame, NextPageTemplate, PageBreak, PageTemplate, Table, TableStyle,
)

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
    """Write the PDF for one sheet, following result.plan exactly."""
    render_document([(model, result)], output_path)


def render_document(sections, output_path):
    """Write one PDF from several sheets: a list of (SheetModel, PlanResult).

    Each sheet keeps its own plan: paper, orientation, font and parts.
    """
    first_page = sections[0][1].plan.page
    doc = BaseDocTemplate(
        output_path,
        pagesize=(first_page.width, first_page.height),
        leftMargin=MARGIN_LEFT,
        rightMargin=MARGIN_RIGHT,
        topMargin=MARGIN_TOP,
        bottomMargin=MARGIN_BOTTOM + FOOTER_HEIGHT,  # the footer sits in this strip
        title=", ".join(model.title for model, _ in sections),
        author="IntelliConverter",
    )

    story, templates, pages = [], [], []
    for index, (model, result) in enumerate(sections):
        template_id = f"sheet{index}"
        templates.append(_page_template(template_id, result.plan.page))
        if index:
            story += [NextPageTemplate(template_id), PageBreak()]
        story += _section_story(index, model, result, pages)

    footer = _footer(sections, pages)
    for template in templates:
        template.onPage = footer
    doc.addPageTemplates(templates)
    try:
        doc.build(story)
    except OSError as error:
        reason = f" ({error.strerror})" if error.strerror else ""
        raise OutputError(
            f"Could not write {output_path}. If it is open in another program, "
            f"close it and try again.{reason}"
        ) from error


def _page_template(template_id, page):
    width = page.width - MARGIN_LEFT - MARGIN_RIGHT
    height = page.height - MARGIN_TOP - MARGIN_BOTTOM - FOOTER_HEIGHT
    # No frame padding: the planner's page geometry is the frame, exactly.
    frame = Frame(MARGIN_LEFT, MARGIN_BOTTOM + FOOTER_HEIGHT, width, height,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id=template_id)
    return PageTemplate(id=template_id, frames=[frame], pagesize=(page.width, page.height))


def _section_story(index, model, result, pages):
    """Flowables for one sheet. Appends (section, band, sheet page) to `pages` per page."""
    plan, texts = result.plan, result.measured
    width = plan.page.width - MARGIN_LEFT - MARGIN_RIGHT
    right_aligned = numeric_columns(model)
    story = [
        SectionStart(f"sheet{index}", model.title),
        TitleBlock(model.title, _subtitle(model, plan), width),
    ]
    sheet_page = 0
    for band_index, band in enumerate(plan.bands):
        bounds = list(band.page_starts) + [len(band.row_heights)]
        for start, end in zip(bounds, bounds[1:]):
            if sheet_page:
                story.append(PageBreak())
            story.append(_page_table(band, start, end, texts, right_aligned, plan.font_size))
            sheet_page += 1
            pages.append((index, band_index, sheet_page))
    return story


class SectionStart(Flowable):
    """Takes no space; marks where a sheet starts with a PDF bookmark."""

    def __init__(self, key, title):
        super().__init__()
        self.key, self.title = key, title

    def wrap(self, available_width, available_height):
        return 0, 0

    def draw(self):
        self.canv.bookmarkPage(self.key)
        self.canv.addOutlineEntry(self.title, self.key, level=0)


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


def _footer(sections, pages):
    """Footer for every page: sheet name, sheet page, part, and document page."""
    total = len(pages)

    def draw(canvas, doc):
        if doc.page > total:
            return
        section, band_index, sheet_page = pages[doc.page - 1]
        model, result = sections[section]
        plan = result.plan
        parts = [model.title, f"Page {sheet_page} of {plan.pages}"]
        if len(plan.bands) > 1:
            letters = plan.column_letters
            columns = plan.bands[band_index].columns[1:]
            parts.append(
                f"Part {band_index + 1} of {len(plan.bands)}: "
                f"columns {letters[columns[0]]} to {letters[columns[-1]]}"
            )
        if len(sections) > 1:
            parts.append(f"Document page {doc.page} of {total}")
        width = plan.page.width - MARGIN_LEFT - MARGIN_RIGHT
        canvas.saveState()
        canvas.setFont(BODY_FONT, FOOTER_SIZE)
        canvas.setFillColor(MUTED)
        canvas.drawCentredString(plan.page.width / 2, MARGIN_BOTTOM + 3,
                                 _fit("  |  ".join(parts), BODY_FONT, FOOTER_SIZE, width))
        canvas.restoreState()

    return draw


def _fit(text, font, size, width):
    """Shorten text with '...' so it fits `width` points."""
    if text_width(text, font, size) <= width:
        return text
    while text and text_width(text + "...", font, size) > width:
        text = text[:-1]
    return text + "..."

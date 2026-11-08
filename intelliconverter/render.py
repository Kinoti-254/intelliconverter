"""Building the PDF with ReportLab (moved from excel_to_pdf.py in Phase 0).

Since Phase 1 the PDF is drawn from a SheetModel: hidden rows and columns
are left out, and formula cells show the value Excel last saved.
"""
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A3, LEGAL, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

from intelliconverter.errors import OutputError
from intelliconverter.reader import load_sheet


# Map layout names to reportlab page sizes
PAGE_SIZES = {
    "A4 Portrait":      portrait(A4),
    "A4 Landscape":     landscape(A4),
    "A3 Portrait":      portrait(A3),
    "A3 Landscape":     landscape(A3),
    "Legal Landscape":  landscape(LEGAL),
}


def get_page_size(layout_name):
    """Return the reportlab page size tuple for a given layout name."""
    return PAGE_SIZES.get(layout_name, landscape(A4))


def extract_sheet_data(ws):
    """Extract all cell values from the worksheet as a 2D list."""
    data = []
    for row in ws.iter_rows(values_only=True):
        # Convert None to empty string so the table renders cleanly
        data.append([str(cell) if cell is not None else "" for cell in row])
    return data


def build_pdf(filepath, output_path, layout_name, scale_percent, sheet=None):
    """
    Generate a PDF from an Excel file using the specified layout.

    Args:
        filepath: Path to the .xlsx file
        output_path: Where to save the PDF
        layout_name: e.g. "A4 Landscape"
        scale_percent: How much to scale the table (e.g. 86.0)
        sheet: Sheet name; the active sheet when omitted
    """
    render_pdf(load_sheet(filepath, sheet), output_path, layout_name, scale_percent)


def render_pdf(sheet, output_path, layout_name, scale_percent):
    """Generate a PDF from a SheetModel. Arguments as for build_pdf()."""
    page_size = get_page_size(layout_name)
    page_width, page_height = page_size

    # Set up the document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=page_size,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.6 * inch,
        bottomMargin=0.5 * inch,
    )

    # Visible cells as text; None becomes an empty string so the table renders cleanly
    data = [
        [str(cell) if cell is not None else "" for cell in row]
        for row in sheet.visible_rows()
    ]

    if not data:
        print("⚠️  No data found in sheet.")
        return

    # Calculate column widths scaled to fit the page
    usable_width = page_width - inch  # subtract margins
    num_cols = len(data[0])
    col_width = (usable_width / num_cols) * (scale_percent / 100)

    col_widths = [col_width] * num_cols

    # Build the table
    table = Table(data, colWidths=col_widths, repeatRows=1)

    # Style the table
    table.setStyle(TableStyle([
        # Header row
        ("BACKGROUND",  (0, 0), (-1, 0),  colors.HexColor("#2C3E50")),
        ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0),  9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),

        # Data rows
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F3F4")]),

        # Grid
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
    ]))

    # Add a title above the table
    styles = getSampleStyleSheet()
    rows = sheet.visible_row_count
    columns = len(sheet.visible_columns)
    title = Paragraph(f"<b>{sheet.title}</b>", styles["Title"])
    subtitle = Paragraph(
        f"Layout: {layout_name}  |  Scale: {scale_percent:.0f}%  |  "
        f"Rows: {rows}  |  Columns: {columns}",
        styles["Normal"]
    )

    elements = [title, Spacer(1, 0.2 * inch), subtitle, Spacer(1, 0.2 * inch), table]

    try:
        doc.build(elements)
    except OSError as error:
        reason = f" ({error.strerror})" if error.strerror else ""
        raise OutputError(
            f"Could not write {output_path}. If it is open in another program, "
            f"close it and try again.{reason}"
        ) from error
    print(f"\n✅ PDF saved to: {output_path}")
    print(f"   Layout  : {layout_name}")
    print(f"   Scale   : {scale_percent:.0f}%")
    print(f"   Rows    : {rows}")
    print(f"   Columns : {columns}\n")

"""The layout analysis report printed to the console, followed by PDF export."""
from intelliconverter.layout import LAYOUTS, score_layout
from intelliconverter.measure import get_total_table_width
from intelliconverter.reader import load_sheet
from intelliconverter.render import render_pdf

OUTPUT_PDF = "output.pdf"


def analyze(filepath, sheet=None):
    """Load an Excel sheet, print a full layout analysis report and export the PDF.

    Raises InputError or OutputError (both ConversionError) for problems the
    user can fix.
    """
    model = load_sheet(filepath, sheet)

    total_width = get_total_table_width(model)

    print(f"\n{'='*48}")
    print(f"  IntelliConverter — Layout Analysis Report")
    print(f"{'='*48}")
    print(f"  Sheet   : {model.title}")
    print(f"  Columns : {len(model.visible_columns)}")
    print(f"  Rows    : {model.visible_row_count}")
    _print_hidden_summary(model)
    print(f"  Width   : {total_width:.2f} inches")
    _print_formula_warning(model)
    print(f"{'='*48}")
    print(f"  {'Layout':<20} {'Fits':<6} {'Scale':>6}  {'Score':>6}")
    print(f"  {'-'*42}")

    best_name = None
    best_score = -1
    best_scale = 0

    for name, page_width in LAYOUTS:
        fits = total_width <= page_width
        scale = (total_width / page_width) * 100
        score = score_layout(total_width, page_width)

        fits_icon = "✅" if fits else "❌"
        score_display = f"{score:.0f}" if score >= 0 else "—"

        print(f"  {name:<20} {fits_icon:<6} {scale:>5.0f}%  {score_display:>6}")

        if score > best_score:
            best_score = score
            best_name = name
            best_scale = scale

    print(f"  {'-'*42}")
    print(f"\n  ✨ Best layout : {best_name}")
    print(f"  📐 Scale       : {best_scale:.0f}%")
    print(f"  🏆 Score       : {best_score:.0f}/100")
    print(f"\n{'='*48}\n")

    # Generate the PDF of the analysed sheet using the best layout
    render_pdf(model, OUTPUT_PDF, best_name, best_scale)


def _plural(count, word):
    return f"{count} {word}" if count == 1 else f"{count} {word}s"


def _print_hidden_summary(model):
    hidden_columns = model.column_count - len(model.visible_columns)
    hidden_rows = len(model.hidden_rows)
    if hidden_columns or hidden_rows:
        print(
            f"  Hidden  : {_plural(hidden_columns, 'column')}, "
            f"{_plural(hidden_rows, 'row')} (not included)"
        )


def _print_formula_warning(model):
    count = model.uncached_formula_count
    if count:
        verb = "has" if count == 1 else "have"
        print(
            f"  Warning : {_plural(count, 'formula cell')} {verb} no saved value "
            "and will show blank. Open and save the file in Excel to fix this."
        )

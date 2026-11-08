"""Console reports.

analyze()    - the current layout report, followed by PDF export
print_plan() - a preview of the new layout engine's plan (no PDF)
"""
from intelliconverter.layout import LAYOUTS, score_layout
from intelliconverter.measure import get_total_table_width
from intelliconverter.planner import plan_layout
from intelliconverter.reader import load_sheet
from intelliconverter.render import render_pdf

OUTPUT_PDF = "output.pdf"
POINTS_PER_INCH = 72
TOP_CANDIDATES = 5
WIDTHS_PER_LINE = 6


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


def print_plan(filepath, sheet=None):
    """Print the new layout engine's plan for a sheet. Writes no PDF.

    Raises InputError or LayoutError (both ConversionError) for problems the
    user can fix.
    """
    model = load_sheet(filepath, sheet)
    result = plan_layout(model)
    plan = result.plan

    print(f"\n{'='*48}")
    print(f"  IntelliConverter — Layout Plan (preview)")
    print(f"{'='*48}")
    print(f"  Sheet   : {model.title}")
    print(f"  Columns : {len(model.visible_columns)}")
    print(f"  Rows    : {model.visible_row_count}")
    _print_hidden_summary(model)
    print(f"{'='*48}")
    print(f"  Paper   : {plan.page.name}")
    print(f"  Font    : {plan.font_size:g} pt")
    print(f"  Bands   : {len(plan.bands)}")
    print(f"  Pages   : {plan.pages:,}")
    for number, band in enumerate(plan.bands, start=1):
        _print_band_widths(number, band, plan.column_letters)
    print(f"  {'-'*46}")
    _print_candidates(result)
    print(f"  {'-'*46}")
    print("  Cost of the chosen layout:")
    for part, value in plan.cost.items():
        if value:
            print(f"    {part:<16} {value:6.2f}")
    print(f"    {'total':<16} {plan.total_cost:6.2f}")
    print(f"{'='*48}")
    print(f"  Preview only: {OUTPUT_PDF} still uses the current layout.\n")


def _print_band_widths(number, band, letters):
    entries = [
        f"{letters[column]} {width / POINTS_PER_INCH:.2f}"
        for column, width in zip(band.columns, band.widths)
    ]
    for start in range(0, len(entries), WIDTHS_PER_LINE):
        label = f"Band {number}" if start == 0 else ""
        chunk = ", ".join(entries[start:start + WIDTHS_PER_LINE])
        end = " (inches)" if start + WIDTHS_PER_LINE >= len(entries) else ","
        print(f"  {label:<8}: {chunk}{end}")


def _print_candidates(result):
    pages_label = "Pages*" if result.sampled else "Pages"
    print("  Best candidates (lower cost is better):")
    print(f"  {'Candidate':<18} {'Font':>5} {'Bands':>6} {pages_label:>7} {'Cost':>7}")
    feasible = [c for c in result.candidates if c.feasible]
    for index, candidate in enumerate(feasible[:TOP_CANDIDATES]):
        marker = "  <- chosen" if index == 0 else ""
        print(
            f"  {candidate.page.name:<18} {candidate.font_size:>5g} {len(candidate.bands):>6} "
            f"{candidate.pages:>7,} {candidate.total:>7.2f}{marker}"
        )
    if result.sampled:
        print("  * estimated from a sample of rows; the chosen plan is exact")

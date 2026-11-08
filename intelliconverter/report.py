"""The console report: what the layout engine chose, then the PDF export."""
from intelliconverter.planner import plan_layout
from intelliconverter.reader import load_sheet
from intelliconverter.render import render_pdf

OUTPUT_PDF = "output.pdf"
POINTS_PER_INCH = 72
TOP_CANDIDATES = 5
WIDTHS_PER_LINE = 6


def analyze(filepath, sheet=None, paper=None, explain=False):
    """Plan the layout of an Excel sheet, print a report and write output.pdf.

    paper:   "A4", "Letter", "Legal" or "A3" to fix the paper; None lets the
             engine choose (both orientations are always compared).
    explain: also print the best alternatives and the cost breakdown.

    Raises InputError, LayoutError or OutputError (all ConversionError) for
    problems the user can fix.
    """
    model = load_sheet(filepath, sheet)
    result = plan_layout(model, paper=paper)
    plan = result.plan

    print(f"\n{'='*48}")
    print(f"  IntelliConverter — Layout Report")
    print(f"{'='*48}")
    print(f"  Sheet   : {model.title}")
    print(f"  Columns : {len(model.visible_columns)}")
    print(f"  Rows    : {model.visible_row_count}")
    _print_hidden_summary(model)
    _print_formula_warning(model)
    print(f"{'='*48}")
    print(f"  Paper   : {plan.page.name}")
    print(f"  Font    : {plan.font_size:g} pt")
    print(f"  Parts   : {len(plan.bands)}{_parts_note(plan)}")
    print(f"  Pages   : {plan.pages:,}")
    for number, band in enumerate(plan.bands, start=1):
        _print_band_widths(number, band, plan.column_letters)
    if explain:
        print(f"  {'-'*46}")
        _print_candidates(result)
        print(f"  {'-'*46}")
        _print_cost(plan)
    print(f"{'='*48}\n")

    render_pdf(model, result, OUTPUT_PDF)
    print(f"✅ PDF saved to: {OUTPUT_PDF}\n")


def _plural(count, word):
    return f"{count} {word}" if count == 1 else f"{count} {word}s"


def _parts_note(plan):
    return " (columns split to fit the page width)" if len(plan.bands) > 1 else ""


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


def _print_band_widths(number, band, letters):
    entries = [
        f"{letters[column]} {width / POINTS_PER_INCH:.2f}"
        for column, width in zip(band.columns, band.widths)
    ]
    for start in range(0, len(entries), WIDTHS_PER_LINE):
        label = f"Part {number}" if start == 0 else ""
        chunk = ", ".join(entries[start:start + WIDTHS_PER_LINE])
        end = " (inches)" if start + WIDTHS_PER_LINE >= len(entries) else ","
        print(f"  {label:<8}: {chunk}{end}")


def _print_candidates(result):
    pages_label = "Pages*" if result.sampled else "Pages"
    print("  Best candidates (lower cost is better):")
    print(f"  {'Candidate':<18} {'Font':>5} {'Parts':>6} {pages_label:>7} {'Cost':>7}")
    feasible = [c for c in result.candidates if c.feasible]
    for index, candidate in enumerate(feasible[:TOP_CANDIDATES]):
        marker = "  <- chosen" if index == 0 else ""
        print(
            f"  {candidate.page.name:<18} {candidate.font_size:>5g} {len(candidate.bands):>6} "
            f"{candidate.pages:>7,} {candidate.total:>7.2f}{marker}"
        )
    if result.sampled:
        print("  * estimated from a sample of rows; the chosen plan is exact")


def _print_cost(plan):
    print("  Cost of the chosen layout:")
    for part, value in plan.cost.items():
        if value:
            print(f"    {part:<16} {value:6.2f}")
    print(f"    {'total':<16} {plan.total_cost:6.2f}")

"""The layout analysis report printed to the console, followed by PDF export."""
from intelliconverter.layout import LAYOUTS, score_layout
from intelliconverter.measure import get_total_table_width
from intelliconverter.reader import load_active_sheet


def analyze(filepath):
    """Load an Excel file and print a full layout analysis report."""
    ws = load_active_sheet(filepath)

    total_width = get_total_table_width(ws)

    print(f"\n{'='*48}")
    print(f"  IntelliConverter — Layout Analysis Report")
    print(f"{'='*48}")
    print(f"  Sheet   : {ws.title}")
    print(f"  Columns : {ws.max_column}")
    print(f"  Rows    : {ws.max_row}")
    print(f"  Width   : {total_width:.2f} inches")
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

# Generate the PDF using the best layout
    from intelliconverter.render import build_pdf
    build_pdf("test.xlsx", "output.pdf", best_name, best_scale)

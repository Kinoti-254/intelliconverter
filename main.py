"""Command-line entry point: python main.py [file] [--sheet NAME].

With no arguments it analyses test.xlsx, as before. The code lives in the
intelliconverter package; the names below are re-exported so existing
`import main` usage keeps working.
"""
from intelliconverter.cli import main as run_cli
from intelliconverter.layout import LAYOUTS, score_layout
from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.report import analyze

__all__ = [
    "LAYOUTS",
    "analyze",
    "col_width_to_inches",
    "get_total_table_width",
    "score_layout",
]

if __name__ == "__main__":
    raise SystemExit(run_cli())

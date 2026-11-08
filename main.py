"""Command-line entry point.

    python main.py [file] [--sheet NAME ...] [--paper PAPER] [--explain] [-o PDF]

With no arguments it converts every visible sheet of test.xlsx in the
current folder. The code lives in the intelliconverter package; the names
below are re-exported so existing `import main` usage keeps working.
"""
from intelliconverter.cli import main as run_cli
from intelliconverter.measure import col_width_to_inches, get_total_table_width
from intelliconverter.report import analyze

__all__ = ["analyze", "col_width_to_inches", "get_total_table_width"]

if __name__ == "__main__":
    raise SystemExit(run_cli())

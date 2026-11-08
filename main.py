"""Command-line entry point, kept so `python main.py` works as before.

The code now lives in the intelliconverter package. The names below are
re-exported so existing `import main` usage keeps working.
"""
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
    analyze("test.xlsx")

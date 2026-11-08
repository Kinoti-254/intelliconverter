"""Command line: python main.py [file] [--sheet NAME] [--paper PAPER] [--explain]."""
import argparse
import sys

from intelliconverter.errors import ConversionError
from intelliconverter.planner import PAPERS
from intelliconverter.report import OUTPUT_PDF, analyze

DEFAULT_FILE = "test.xlsx"
AUTO = "auto"


def build_parser():
    parser = argparse.ArgumentParser(
        prog="python main.py",
        description=(
            "Choose the most readable page layout for an Excel sheet and export "
            f"{OUTPUT_PDF} to the current folder."
        ),
    )
    parser.add_argument(
        "file",
        nargs="?",
        default=DEFAULT_FILE,
        help=f"the .xlsx file to convert (default: {DEFAULT_FILE})",
    )
    parser.add_argument("--sheet", help="sheet name (default: the active sheet)")
    parser.add_argument(
        "--paper",
        choices=[AUTO] + [name for name, _ in PAPERS],
        default=AUTO,
        help="paper size (default: auto, which prefers A4); both orientations are compared",
    )
    parser.add_argument(
        "--explain",
        action="store_true",
        help="also show the best alternatives and why this layout was chosen",
    )
    return parser


def main(argv=None):
    """Run the command line. Returns the exit code: 0 on success, 1 on a user error."""
    args = build_parser().parse_args(argv)
    _avoid_encoding_crashes()
    paper = None if args.paper == AUTO else args.paper
    try:
        analyze(args.file, sheet=args.sheet, paper=paper, explain=args.explain)
    except ConversionError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


def _avoid_encoding_crashes():
    """Show '?' for characters the console can't display instead of crashing.

    For example, a Windows console writing to a file uses cp1252, which has
    no emoji. Consoles that can show them are unaffected.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

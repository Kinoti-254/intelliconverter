"""Command line: python main.py [file] [--sheet NAME ...] [--paper PAPER] [--explain] [-o PDF]."""
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
            "Choose the most readable page layout for each sheet of an Excel "
            "workbook and export them to one PDF."
        ),
    )
    parser.add_argument(
        "file",
        nargs="?",
        default=DEFAULT_FILE,
        help=f"the .xlsx file to convert (default: {DEFAULT_FILE})",
    )
    parser.add_argument(
        "--sheet",
        action="append",
        metavar="NAME",
        help="a sheet to convert; repeat for several (default: every visible sheet)",
    )
    parser.add_argument(
        "--paper",
        choices=[AUTO] + [name for name, _ in PAPERS],
        default=AUTO,
        help="paper size (default: auto, which prefers A4); both orientations are compared",
    )
    parser.add_argument(
        "-o", "--output",
        default=OUTPUT_PDF,
        metavar="PDF",
        help=f"the PDF to write (default: {OUTPUT_PDF} in the current folder)",
    )
    parser.add_argument(
        "--explain",
        action="store_true",
        help="also show each sheet's best alternatives and why its layout was chosen",
    )
    return parser


def main(argv=None):
    """Run the command line. Returns the exit code: 0 on success, 1 on a user error."""
    args = build_parser().parse_args(argv)
    _avoid_encoding_crashes()
    paper = None if args.paper == AUTO else args.paper
    try:
        analyze(args.file, sheets=args.sheet, paper=paper, explain=args.explain, output=args.output)
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

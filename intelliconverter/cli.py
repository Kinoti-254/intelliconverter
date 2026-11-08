"""Command line: python main.py [file] [--sheet NAME]."""
import argparse
import sys

from intelliconverter.errors import ConversionError
from intelliconverter.report import OUTPUT_PDF, analyze

DEFAULT_FILE = "test.xlsx"


def build_parser():
    parser = argparse.ArgumentParser(
        prog="python main.py",
        description=(
            "Analyse an Excel sheet, recommend a page layout and export "
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
    return parser


def main(argv=None):
    """Run the command line. Returns the exit code: 0 on success, 1 on a user error."""
    args = build_parser().parse_args(argv)
    _avoid_encoding_crashes()
    try:
        analyze(args.file, sheet=args.sheet)
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

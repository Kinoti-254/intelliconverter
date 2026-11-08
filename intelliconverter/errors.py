"""Errors the user can fix. Their messages are written to be shown as they are."""


class ConversionError(Exception):
    """Base class for problems reported to the user without a traceback."""


class InputError(ConversionError):
    """The input file or sheet can't be used (missing, wrong type, too big, empty...)."""


class OutputError(ConversionError):
    """The output file can't be written."""

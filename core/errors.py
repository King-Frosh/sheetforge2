"""User-facing exceptions for the FROSHMerge API."""


class ExcelToolError(Exception):
    """An expected error that should be shown to the user as a friendly
    message instead of a 500 crash (e.g. unsupported format, corrupted
    file, limits exceeded)."""

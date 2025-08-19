from datetime import datetime
import pandas as pd
import re


def standardize_date(
    date_input,
    out_format: str | None = None,
    in_format: str | list[str] | None = None,
):
    """Parse various date formats and return a date object or a formatted string.

    - Supports common numeric formats and month-abbreviation formats like
      '2025-Feb-23 00:00:00'.
    - Returns a `datetime.date` by default.
    - If `out_format` is provided (e.g.,"%d/% m/%Y"), returns a string.
    - Optionally provide `in_format` (str or list[str]) to specify known input format(s)
      that will be tried first using datetime.strptime.

    Returns None when parsing fails.
    """
    if date_input is None:
      	return None

    # Handle NaN from pandas
    try:
        if pd.isna(date_input):
            return None
    except Exception:
        pass

    # If already a datetime or date-like
    if hasattr(date_input, "date") and callable(getattr(date_input, "date")):
        date_obj = date_input.date()
        return date_obj.strftime(out_format) if out_format else date_obj

    text = str(date_input).strip()

    # If caller provided explicit input format(s), try those first
    if in_format:
        formats_to_try = [in_format] if isinstance(in_format, str) else list(in_format)
        for fmt in formats_to_try:
            try:
                date_obj = datetime.strptime(text, fmt).date()
                return date_obj.strftime(out_format) if out_format else date_obj
            except Exception:
                continue

    # Try explicit strptime patterns first (fast path)
    candidate_formats = (
        # ISO-like and dashes
        "%Y-%m-%d",
        "%d-%m-%Y",
        # Slashes (EU and US) with 4-digit year
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M:%S",
        # Slashes with 2-digit year
        "%d/%m/%y",
        "%m/%d/%y",
        "%m/%d/%y %H:%M:%S",
        # Year-first with slashes
        "%Y/%m/%d",
        # Month name variants
        "%Y-%b-%d %H:%M:%S",  # e.g., 2025-Feb-23 00:00:00
        "%Y-%b-%d",
        "%d-%b-%Y",
        "%d-%b-%Y %H:%M:%S",
        "%Y-%B-%d %H:%M:%S",
        "%Y-%B-%d",
        "%d-%B-%Y",
        "%d-%B-%Y %H:%M:%S",
    )
    for fmt in candidate_formats:
        try:
            date_obj = datetime.strptime(text, fmt).date()
            return date_obj.strftime(out_format) if out_format else date_obj
        except Exception:
            continue

    # Fallback to pandas parser (handles Excel serials and many strings)
    try:
        date_obj = pd.to_datetime(date_input, errors="raise").date()
        return date_obj.strftime(out_format) if out_format else date_obj
    except Exception:
        return None


def format_date_mmddyyyy(date_input, in_format: str | list[str] | None = None) -> str | None:
    """Return the date formatted as MM/DD/YYYY, or None if parsing fails.

    Accepts optional `in_format` to guide parsing.
    """
    return standardize_date(date_input, out_format="%m/%d/%Y", in_format=in_format)


def has_year_component(value) -> bool:
    """Return True if the input string/value appears to include an explicit year.

    - Datetime-like objects always return True
    - Strings are matched against common patterns with year present
    """
    if value is None:
        return False
    # Datetime-like objects
    if hasattr(value, "year"):
        return True

    text = str(value).strip()
    if text == "":
        return False

    # 4-digit year anywhere
    if re.search(r"\b\d{4}\b", text):
        return True
    # dd/mm/yy or mm/dd/yy
    if re.match(r"^\s*\d{1,2}[/-]\d{1,2}[/-]\d{2}\s*$", text):
        return True
    # Month name followed by 4 or 2 digit year somewhere
    month_name = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*"
    if re.search(month_name + r".*\b(\d{4}|\d{2})\b", text, re.IGNORECASE):
        return True

    return False



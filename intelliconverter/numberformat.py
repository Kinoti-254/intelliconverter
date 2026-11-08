"""Turning cell values into the text Excel would show (basic number formats).

Supported: General; fixed and optional decimals ("0.00", "0.##"); thousands
separators ("#,##0"); leading zeros ("000"); percentages; scientific
("0.00E+00"); literal text and currency symbols ('"KES "#,##0', "$#,##0",
"[$KES]"); separate formats for negative numbers and zero ("pos;neg;zero");
and date/time formats (d, dd, ddd, dddd, m, mm, mmm, mmmm, yy, yyyy, h, hh,
mm, ss, AM/PM).

Anything else (fractions, scaling by commas, elapsed-time arithmetic...)
falls back to General rather than failing. Rounding is half-up, like Excel.
"""
import datetime as dt
import re
from decimal import ROUND_HALF_UP, Decimal

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# A number placeholder run such as "#,##0.00" or "0.00E+00"
_NUMBER = re.compile(r"[#0?,]*\.?[#0?]*(?:[Ee][+-][#0?]+)?")
_DATE_TOKENS = re.compile(
    r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p",
    re.IGNORECASE,
)


def format_value(value, number_format="General"):
    """The text Excel shows for `value` under `number_format`."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return _format_date(value, number_format or "General")
    if isinstance(value, (int, float)):
        return _format_number(value, number_format or "General")
    return str(value)


# --- Numbers ------------------------------------------------------------------

def _general(value):
    if isinstance(value, int) or (value.is_integer() and abs(value) < 1e11):
        return str(int(value))
    text = f"{value:.10g}"  # Excel's General shows up to about 11 significant digits
    if "e" in text:
        mantissa, exponent = text.split("e")
        text = f"{mantissa}E{int(exponent):+03d}"
    return text


def _format_number(value, number_format):
    sections = _split_sections(number_format)
    if value < 0 and len(sections) >= 2:
        section, sign, value = sections[1], "", -value  # the section shows any sign itself
    elif value == 0 and len(sections) >= 3:
        section, sign = sections[2], ""
    else:
        section, sign = sections[0], "-" if value < 0 else ""
        value = abs(value)

    literal = _clean_section(section)
    match = next((m for m in _NUMBER.finditer(literal) if re.search(r"[#0?]", m.group())), None)
    if match is None:
        if literal.strip().lower() in ("general", "@", ""):
            return sign + _general(value)
        return literal  # a section that is only text, such as "-" for zero

    token = match.group()
    if "?" in token or "/" in literal:
        return sign + _general(value)  # fractions and aligned digits: not supported

    if "%" in literal:
        value *= 100
    body = _format_scientific(value, token) if "E" in token.upper() else _format_fixed(value, token)
    return sign + literal[:match.start()] + body + literal[match.end():]


def _split_sections(number_format):
    sections, current, quoted = [], "", False
    for char in number_format:
        if char == '"':
            quoted = not quoted
        if char == ";" and not quoted:
            sections.append(current)
            current = ""
        else:
            current += char
    sections.append(current)
    return sections


def _clean_section(section):
    """Resolve quotes, escapes, padding and bracket codes into plain literal text."""
    out = []
    i = 0
    while i < len(section):
        char = section[i]
        if char == '"':
            end = section.find('"', i + 1)
            end = len(section) if end == -1 else end
            out.append(section[i + 1:end])
            i = end + 1
        elif char == "\\" and i + 1 < len(section):
            out.append(section[i + 1])
            i += 2
        elif char in "_*" and i + 1 < len(section):
            i += 2  # padding or fill character: no visible text
        elif char == "[":
            end = section.find("]", i)
            end = len(section) if end == -1 else end
            code = section[i + 1:end]
            if code.startswith("$"):
                out.append(code[1:].split("-")[0])  # currency symbol, e.g. [$KES-441]
            i = end + 1  # colours and conditions are dropped
        else:
            out.append(char)
            i += 1
    return "".join(out)


def _round(value, decimals):
    quantum = Decimal(1).scaleb(-decimals)
    return Decimal(repr(value)).quantize(quantum, rounding=ROUND_HALF_UP)


def _format_fixed(value, token):
    integer_part, _, decimal_part = token.partition(".")
    max_decimals = sum(1 for c in decimal_part if c in "#0?")
    min_decimals = decimal_part.count("0")
    min_integer_digits = integer_part.count("0")

    rounded = _round(value, max_decimals)
    whole, _, fraction = f"{rounded:f}".partition(".")
    fraction = fraction.rstrip("0").ljust(min_decimals, "0") if max_decimals else ""

    if whole == "0" and min_integer_digits == 0:
        whole = ""
    whole = whole.rjust(min_integer_digits, "0")
    if "," in integer_part and whole:
        whole = f"{int(whole):,}"
    return whole + ("." + fraction if "." in token else "")


def _format_scientific(value, token):
    mantissa = token.upper().split("E")[0]
    decimals = len(mantissa.partition(".")[2])
    return f"{value:.{decimals}E}"


# --- Dates and times ------------------------------------------------------------

def _format_date(value, number_format):
    section = _clean_section(_split_sections(number_format)[0])
    if not _DATE_TOKENS.search(section) or section.strip().lower() == "general":
        return _iso(value)

    tokens = _tokenize_date(section)
    twelve_hour = any(t.lower() in ("am/pm", "a/p") for t, _ in tokens)
    parts = []
    for index, (token, is_code) in enumerate(tokens):
        parts.append(_date_part(value, token, tokens, index, twelve_hour) if is_code else token)
    return "".join(parts)


def _iso(value):
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d") if value.time() == dt.time(0) else value.strftime("%Y-%m-%d %H:%M:%S")
    return value.isoformat()


def _tokenize_date(section):
    tokens, position = [], 0
    for match in _DATE_TOKENS.finditer(section):
        if match.start() > position:
            tokens.append((section[position:match.start()], False))
        tokens.append((match.group(), True))
        position = match.end()
    if position < len(section):
        tokens.append((section[position:], False))
    return tokens


def _is_minutes(tokens, index):
    """m or mm means minutes right after an hour code or right before seconds."""
    codes = [(i, t.lower()) for i, (t, is_code) in enumerate(tokens) if is_code]
    position = next(p for p, (i, _) in enumerate(codes) if i == index)
    before = codes[position - 1][1] if position > 0 else ""
    after = codes[position + 1][1] if position + 1 < len(codes) else ""
    return before.startswith("h") or after.startswith("s")


def _date_part(value, token, tokens, index, twelve_hour):
    code = token.lower()
    date = value if isinstance(value, (dt.datetime, dt.date)) else None
    time = value if isinstance(value, (dt.datetime, dt.time)) else None
    if code in ("m", "mm") and _is_minutes(tokens, index):
        minute = time.minute if time else 0
        return f"{minute:02d}" if code == "mm" else str(minute)
    if code.startswith("y"):
        year = date.year if date else 1900
        return str(year) if code == "yyyy" else f"{year % 100:02d}"
    if code.startswith("m"):
        month = date.month if date else 1
        return {"m": str(month), "mm": f"{month:02d}",
                "mmm": MONTHS[month - 1][:3]}.get(code, MONTHS[month - 1])
    if code.startswith("d"):
        if code in ("ddd", "dddd"):
            name = DAYS[date.weekday()] if date else DAYS[0]
            return name[:3] if code == "ddd" else name
        day = date.day if date else 0
        return f"{day:02d}" if code == "dd" else str(day)
    if code.startswith("h"):
        hour = time.hour if time else 0
        if twelve_hour:
            hour = hour % 12 or 12
        return f"{hour:02d}" if code == "hh" else str(hour)
    if code.startswith("s"):
        second = time.second if time else 0
        return f"{second:02d}" if code == "ss" else str(second)
    hour = time.hour if time else 0
    if code == "a/p":
        return "A" if hour < 12 else "P"
    return "AM" if hour < 12 else "PM"

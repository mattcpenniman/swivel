"""Cell values as a pivot cache sees them.

Source cells are normalised to a small set of Python types:

* ``None``                 - a blank cell
* ``bool``                 - TRUE / FALSE
* ``int`` / ``float``      - numbers
* ``datetime.datetime``    - dates (and date-times)
* ``str``                  - text
* :class:`XLError`         - an error value such as ``#N/A``

Everything here mirrors how Excel itself treats those values inside a
PivotTable: items are grouped case-insensitively, sorted numbers-first, and
aggregated with Excel's rules for blanks, text and errors.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal

EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
_LEAP_BUG_CUTOFF = _dt.datetime(1900, 3, 1)


@dataclass(frozen=True)
class XLError:
    """An Excel error value (``#N/A``, ``#DIV/0!`` ...)."""

    code: str

    def __str__(self) -> str:  # pragma: no cover - trivial
        return self.code


DIV0 = XLError("#DIV/0!")

ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
     "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!",
     "#BUSY!", "#UNKNOWN!", "#PYTHON!"}
)


def normalize(value, data_type: str | None = None):
    """Convert a value read by openpyxl into one of the canonical types."""
    if value is None:
        return None
    if data_type == "e" or isinstance(value, XLError):
        return value if isinstance(value, XLError) else XLError(str(value))
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return XLError("#NUM!")
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, _dt.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    if isinstance(value, _dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second
                + value.microsecond / 1e6) / 86400.0
    if isinstance(value, _dt.timedelta):
        return value.total_seconds() / 86400.0
    if isinstance(value, str):
        if value == "":
            return None
        return value
    return str(value)


def to_serial(d: _dt.datetime) -> float:
    """Excel serial number of a datetime (1900 date system)."""
    delta = d - EXCEL_EPOCH
    serial = delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0
    if d < _LEAP_BUG_CUTOFF:
        serial -= 1  # Excel pretends 1900-02-29 existed
    return serial


def is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def numeric(v):
    """The value as a float when Excel would aggregate it as a number."""
    if is_number(v):
        return float(v)
    if isinstance(v, _dt.datetime):
        return to_serial(v)
    return None


def item_key(v):
    """Grouping key: Excel merges text items case-insensitively."""
    if v is None:
        return ("m",)
    if isinstance(v, bool):
        return ("b", v)
    if is_number(v):
        return ("n", float(v))
    if isinstance(v, _dt.datetime):
        return ("d", v)
    if isinstance(v, XLError):
        return ("e", v.code)
    return ("s", v.lower())


_IGNORED_IN_SORT = re.compile(r"['\-‐‑‒–—­]")


def sort_key(v):
    """Excel's ascending item order: numbers, text, logicals, errors, dates, blanks.

    (Dates only interleave with nothing: in a field mixing types Excel lists
    them after the errors.)
    """
    if v is None:
        return (5, 0, "", "", "")
    if isinstance(v, bool):
        return (2, int(v), "", "", "")
    if is_number(v):
        return (0, float(v), "", "", "")
    if isinstance(v, _dt.datetime):
        return (4, to_serial(v), "", "", "")
    if isinstance(v, XLError):
        return (3, 0, v.code, "", "")
    low = v.lower()
    # Excel sorts text like Windows' CompareString "word sort": case, accents,
    # hyphens and apostrophes are ignored on the first pass; on a tie the
    # plain spelling comes first, then unaccented before accented.
    base = "".join(ch for ch in unicodedata.normalize("NFD", low)
                   if not unicodedata.combining(ch))
    return (1, 0, _IGNORED_IN_SORT.sub("", base),
            _IGNORED_IN_SORT.sub("￿", base), low)


def format_general(x: float) -> str:
    """Render a number the way Excel's General format writes it into text."""
    if isinstance(x, int) or (isinstance(x, float) and x.is_integer() and abs(x) < 1e15):
        return str(int(x))
    s = format(x, ".15g")
    if "e" in s:
        mant, exp = s.split("e")
        sign = exp[0]
        exp = exp[1:].lstrip("0").rjust(2, "0")
        return "%sE%s%s" % (mant, "+" if sign == "+" else "-", exp)
    return s


# --------------------------------------------------------------------------
# Aggregation
# --------------------------------------------------------------------------

#: user-facing function names -> (OOXML ``subtotal`` value, caption prefix)
FUNCTIONS = {
    "sum": ("sum", "Sum"),
    "count": ("count", "Count"),
    "average": ("average", "Average"),
    "max": ("max", "Max"),
    "min": ("min", "Min"),
    "product": ("product", "Product"),
    "count_nums": ("countNums", "Count"),
    "stdev": ("stdDev", "StdDev"),
    "stdevp": ("stdDevp", "StdDevp"),
    "var": ("var", "Var"),
    "varp": ("varp", "Varp"),
}

_ALIASES = {
    "avg": "average", "mean": "average", "countnums": "count_nums",
    "count_numbers": "count_nums", "stddev": "stdev", "stddevp": "stdevp",
    "std": "stdev", "variance": "var", "maximum": "max", "minimum": "min",
    "total": "sum", "counta": "count",
}


def canonical_function(name: str) -> str:
    key = name.strip().lower().replace(" ", "_")
    key = _ALIASES.get(key, key)
    if key not in FUNCTIONS:
        raise ValueError(
            "unknown summary function %r; expected one of %s"
            % (name, ", ".join(sorted(FUNCTIONS)))
        )
    return key


class Accumulator:
    """Streaming accumulator for one value field in one pivot cell."""

    __slots__ = ("nonblank", "nums", "total", "prod", "lo", "hi", "mean", "m2", "error",
                 "error_at", "errors")

    def __init__(self):
        self.nonblank = 0
        self.nums = 0
        self.total = 0.0
        self.prod = 1.0
        self.lo = math.inf
        self.hi = -math.inf
        self.mean = 0.0
        self.m2 = 0.0
        self.error = None
        self.error_at = None
        self.errors = 0

    def add(self, v, order=None) -> None:
        """Add a value. ``order`` ranks error values: Excel builds totals from
        their child items in display order, so the error shown is the one from
        the earliest (row path, column path, record)."""
        if v is None:
            return
        self.nonblank += 1
        if isinstance(v, XLError):
            self.errors += 1
            if self.error is None or (order is not None and order < self.error_at):
                self.error = v
                self.error_at = order
            return
        x = numeric(v)
        if x is None:
            return
        self.nums += 1
        self.total += x
        self.prod *= x
        if x < self.lo:
            self.lo = x
        if x > self.hi:
            self.hi = x
        d = x - self.mean
        self.mean += d / self.nums
        self.m2 += d * (x - self.mean)

    def result(self, func: str, numeric_only: bool = False):
        """Final value, ``None`` for an empty cell, or an :class:`XLError`.

        ``numeric_only`` ignores text and logicals entirely, as Excel does in
        some total cells (see model.py).
        """
        nonblank = self.nums + self.errors if numeric_only else self.nonblank
        if nonblank == 0:
            return None
        if func == "count":
            return nonblank
        if func == "count_nums":
            return self.nums
        if self.error is not None:
            return self.error
        n = self.nums
        if func == "sum":
            return self.total
        if func == "average":
            return DIV0 if n == 0 else self.total / n
        if func == "max":
            return 0 if n == 0 else self.hi
        if func == "min":
            return 0 if n == 0 else self.lo
        if func == "product":
            return 0 if n == 0 else self.prod
        if func in ("stdev", "var"):
            if n < 2:
                return DIV0
            var = max(self.m2, 0.0) / (n - 1)
            return math.sqrt(var) if func == "stdev" else var
        if func in ("stdevp", "varp"):
            if n == 0:
                return DIV0
            var = max(self.m2, 0.0) / n
            return math.sqrt(var) if func == "stdevp" else var
        raise ValueError(func)  # pragma: no cover - guarded by canonical_function


# --------------------------------------------------------------------------
# Excel date format codes -> text (for labels such as "2024-02-29 Total")
# --------------------------------------------------------------------------

_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_DATE_TOKEN = re.compile(
    r'"[^"]*"|\\.|\[[^\]]*\]|AM/PM|am/pm|A/P|a/p|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d'
    r'|hh|h|ss|s|\.0+|.', re.I)


def format_serial(x: float, code: str):
    """Render a number as Excel shows it under a date format, or None if Excel
    could not (negative or beyond 9999-12-31)."""
    x = float(x)
    if x < 0 or x >= 2958466:
        return None
    whole = int(x)
    frac = _dt.timedelta(days=x - whole)
    if whole == 0:  # Excel's "January 0, 1900"
        return format_date(_dt.datetime(1900, 1, 1) + frac, code, zero_day=True)
    if whole < 61:  # before Excel's fictitious 1900-02-29
        base = _dt.datetime(1899, 12, 31) + _dt.timedelta(days=min(whole, 59))
    else:
        base = EXCEL_EPOCH + _dt.timedelta(days=whole)
    return format_date(base + frac, code)


def format_date(d: _dt.datetime, code: str, zero_day: bool = False) -> str:
    """Render ``d`` with an Excel number format code (date/time subset)."""
    code = code.split(";")[0] or "m/d/yyyy"
    tokens = _DATE_TOKEN.findall(code)
    # Excel rounds to the precision the format shows (whole seconds by default)
    digits = max((len(t) - 1 for t in tokens if t.startswith(".0")), default=0)
    step = 10 ** (6 - digits)
    us = int(round(d.microsecond / step)) * step
    d = d.replace(microsecond=0) + _dt.timedelta(microseconds=us)
    ampm = any(t.lower() in ("am/pm", "a/p") for t in tokens)
    kinds = []
    for t in tokens:
        low = t.lower()
        if low.startswith("h"):
            kinds.append("h")
        elif low.startswith("s"):
            kinds.append("s")
        elif low in ("m", "mm"):
            kinds.append("m?")
        else:
            kinds.append(None)
    out = []
    for i, t in enumerate(tokens):
        low = t.lower()
        if t.startswith('"'):
            out.append(t[1:-1])
        elif t.startswith("\\"):
            out.append(t[1:])
        elif t.startswith("["):
            continue
        elif low == "yyyy":
            out.append("%04d" % d.year)
        elif low == "yy":
            out.append("%02d" % (d.year % 100))
        elif kinds[i] == "m?":
            prev = next((k for k in reversed(kinds[:i]) if k in ("h", "s", "m?", None)
                         and k is not None), None)
            nxt = next((k for k in kinds[i + 1:] if k in ("h", "s", "m?")), None)
            minute = prev == "h" or nxt == "s"
            if minute:
                out.append("%02d" % d.minute if low == "mm" else str(d.minute))
            else:
                out.append("%02d" % d.month if low == "mm" else str(d.month))
        elif low == "mmmmm":
            out.append(_MONTHS[d.month - 1][0])
        elif low == "mmmm":
            out.append(_MONTHS[d.month - 1])
        elif low == "mmm":
            out.append(_MONTHS[d.month - 1][:3])
        elif low == "dddd":
            out.append(_DAYS[d.weekday()])
        elif low == "ddd":
            out.append(_DAYS[d.weekday()][:3])
        elif low == "dd":
            out.append("%02d" % (0 if zero_day else d.day))
        elif low == "d":
            out.append(str(0 if zero_day else d.day))
        elif low in ("hh", "h"):
            h = d.hour % 12 or 12 if ampm else d.hour
            out.append("%02d" % h if low == "hh" else str(h))
        elif low in ("ss", "s"):
            out.append("%02d" % d.second if low == "ss" else str(d.second))
        elif low.startswith(".0"):
            frac = d.microsecond / 1e6
            out.append(("%.*f" % (len(t) - 1, frac))[1:])
        elif low == "am/pm":
            out.append("AM" if d.hour < 12 else "PM")
        elif low == "a/p":
            out.append("A" if d.hour < 12 else "P")
        else:
            out.append(t)
    return "".join(out)

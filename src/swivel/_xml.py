"""Small XML helpers shared by the part writers."""
from __future__ import annotations

import re

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

RT_WORKSHEET = REL_NS + "/worksheet"
RT_PIVOT_TABLE = REL_NS + "/pivotTable"
RT_PIVOT_CACHE = REL_NS + "/pivotCacheDefinition"
RT_PIVOT_RECORDS = REL_NS + "/pivotCacheRecords"
RT_STYLES = REL_NS + "/styles"
RT_CALC_CHAIN = REL_NS + "/calcChain"
RT_OFFICE_DOC = REL_NS + "/officeDocument"
RT_TABLE = REL_NS + "/table"
RT_PRINTER = REL_NS + "/printerSettings"

CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_PIVOT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
CT_PIVOT_CACHE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
CT_PIVOT_RECORDS = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# Characters XML 1.0 cannot carry. OOXML encodes them as _xHHHH_.
_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_LOOKS_ESCAPED = re.compile(r"_(x[0-9A-Fa-f]{4}_)")


def ooxml_encode(s: str) -> str:
    """Apply the ST_Xstring escaping Excel expects for awkward characters."""
    s = _LOOKS_ESCAPED.sub(r"_x005F_\1", s)
    return _ILLEGAL.sub(lambda m: "_x%04X_" % ord(m.group()), s)


def esc(s: str) -> str:
    """Escape text content."""
    s = ooxml_encode(s)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def attr(s) -> str:
    """Escape an attribute value (without the surrounding quotes)."""
    s = ooxml_encode(str(s))
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;").replace("\n", "&#10;").replace("\r", "&#13;")
            .replace("\t", "&#9;"))


def num(x) -> str:
    """Serialise a number in xsd:double form, integers without a decimal point."""
    if isinstance(x, bool):
        return "1" if x else "0"
    if isinstance(x, int):
        return str(x)
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


def col_letter(idx: int) -> str:
    """0-based column index -> letters."""
    if idx < 0:
        raise ValueError(idx)
    out = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


def col_index(letters: str) -> int:
    """Letters -> 0-based column index."""
    n = 0
    for ch in letters.upper():
        if not "A" <= ch <= "Z":
            raise ValueError("bad column %r" % letters)
        n = n * 26 + ord(ch) - 64
    return n - 1


_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def parse_cell(ref: str) -> tuple[int, int]:
    """'B3' -> (row0, col0) = (2, 1)."""
    m = _CELL.match(ref.strip())
    if not m:
        raise ValueError("not a cell reference: %r" % ref)
    row = int(m.group(2)) - 1
    col = col_index(m.group(1))
    if row < 0 or row >= 1048576 or col >= 16384:
        raise ValueError("cell reference out of range: %r" % ref)
    return row, col


def cell_ref(row: int, col: int) -> str:
    return "%s%d" % (col_letter(col), row + 1)


def parse_range(ref: str) -> tuple[int, int, int, int]:
    """'A1:C9' -> (r0, c0, r1, c1), inclusive and normalised."""
    parts = ref.strip().split(":")
    if len(parts) == 1:
        r, c = parse_cell(parts[0])
        return r, c, r, c
    if len(parts) != 2:
        raise ValueError("not a range: %r" % ref)
    r0, c0 = parse_cell(parts[0])
    r1, c1 = parse_cell(parts[1])
    return min(r0, r1), min(c0, c1), max(r0, r1), max(c0, c1)


def range_ref(r0: int, c0: int, r1: int, c1: int) -> str:
    a, b = cell_ref(r0, c0), cell_ref(r1, c1)
    return a if a == b else "%s:%s" % (a, b)

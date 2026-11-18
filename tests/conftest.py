import datetime as dt
import math
import os
import posixpath
import zipfile

import openpyxl
import pytest
from lxml import etree

from swivel._values import XLError, to_serial

GOLDEN = os.path.join(os.path.dirname(__file__), "golden")


def decode(x):
    if isinstance(x, dict) and "$date" in x:
        return dt.datetime.fromisoformat(x["$date"])
    if isinstance(x, dict) and "$error" in x:
        return x["$error"]
    return x


def write_book(path, rows, sheet="Data", extra=None):
    """Write ``rows`` (header first) to a new workbook and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append([decode(c) for c in r])
    if extra:
        extra(wb)
    wb.save(path)
    return str(path)


@pytest.fixture
def book(tmp_path):
    def make(rows, name="book.xlsx", **kw):
        return write_book(tmp_path / name, rows, **kw)
    return make


SALES = [
    ["Region", "Product", "Year", "Amount", "Qty"],
    ["East", "Apple", 2023, 10.5, 1],
    ["West", "Pear", 2024, 3, 2],
    ["East", "pear", 2024, 7.25, 3],
    ["North", "Apple", 2023, None, 4],
    [None, "Fig", 2024, 1, 5],
    ["West", "Apple", 2023, 2, None],
]


@pytest.fixture
def sales(book):
    return book(SALES)


def comparable(v):
    """Cell value as Excel's Range.Value2 reports it."""
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, dt.datetime):
        return to_serial(v)
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return float(v)
    return v


def assert_cells_equal(ours: dict, theirs: dict, context=""):
    keys = sorted(set(ours) | set(theirs), key=lambda k: (int("".join(filter(str.isdigit, k))), len(k), k))
    diffs = []
    for k in keys:
        a = comparable(ours.get(k))
        b = comparable(theirs.get(k))
        if isinstance(a, float) and isinstance(b, float):
            if math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9):
                continue
        elif a == b:
            continue
        diffs.append((k, ours.get(k), theirs.get(k)))
    assert not diffs, "%s cell differences (cell, swivel, excel):\n%s" % (
        context, "\n".join(map(repr, diffs[:40])))


# --------------------------------------------------------------------------
# Package integrity checks shared by many tests
# --------------------------------------------------------------------------

def check_package(path):
    """Structural checks an OPC consumer (Excel) relies on. Returns the zip names."""
    with zipfile.ZipFile(path) as z:
        return _check_package(z)


def _check_package(z):
    names = set(z.namelist())
    lower = {n.lower(): n for n in names}
    assert len(lower) == len(names), "duplicate part names differing in case"
    ct = etree.fromstring(z.read("[Content_Types].xml"))
    ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
    overrides = {o.get("PartName").lstrip("/").lower(): o.get("ContentType")
                 for o in ct.findall("ct:Override", ns)}
    defaults = {d.get("Extension").lower() for d in ct.findall("ct:Default", ns)}
    for o in overrides:
        assert o in lower, "content type override for missing part %s" % o
    for n in names:
        if n.endswith("/"):
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        assert n.lower() in overrides or ext in defaults, "no content type for %s" % n
    # every part parses and every relationship resolves
    for n in names:
        if n.endswith(".xml") or n.endswith(".rels"):
            etree.fromstring(z.read(n))
        if n.endswith(".rels"):
            d = posixpath.dirname(posixpath.dirname(n))
            for r in etree.fromstring(z.read(n)):
                if r.get("TargetMode") == "External":
                    continue
                t = r.get("Target")
                p = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
                assert p.lower() in lower, "%s points at missing part %s" % (n, p)
    # pivot caches: unique ids, each referenced by exactly the tables that use it
    M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    ids = [int(pc.get("cacheId")) for pc in wb.iter(M + "pivotCache")]
    assert len(ids) == len(set(ids)), "duplicate cacheId"
    used = set()
    for n in names:
        if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml"):
            t = etree.fromstring(z.read(n))
            assert int(t.get("cacheId")) in ids, "%s uses an undeclared cache" % n
            used.add(int(t.get("cacheId")))
            _check_pivot_counts(t)
    assert used == set(ids), "orphaned pivot caches: %s" % (set(ids) - used)
    sheets = [s.get("name").lower() for s in wb.iter(M + "sheet")]
    assert len(sheets) == len(set(sheets)), "duplicate sheet names"
    return names


def _check_pivot_counts(t):
    for el in t.iter():
        c = el.get("count")
        if c is not None and isinstance(el.tag, str) and not el.tag.endswith("}sharedItems"):
            assert int(c) == len(el), "%s count=%s but has %d children" % (el.tag, c, len(el))


# --------------------------------------------------------------------------
# Excel (optional oracle)
# --------------------------------------------------------------------------

def _excel_available():
    if os.environ.get("SWIVEL_NO_EXCEL"):
        return False
    try:
        from swivel.excel import available, excel_app
        if not available():
            return False
        with excel_app():
            return True
    except Exception:
        return False


_EXCEL = None


def excel_ok():
    global _EXCEL
    if _EXCEL is None:
        _EXCEL = _excel_available()
    return _EXCEL


@pytest.fixture(scope="session")
def excel():
    if not excel_ok():
        pytest.skip("Microsoft Excel (COM) not available")
    from swivel.excel import excel_app
    with excel_app() as app:
        yield app


def pytest_collection_modifyitems(config, items):
    for item in items:
        if "excel" in getattr(item, "fixturenames", ()):
            item.add_marker(pytest.mark.excel)

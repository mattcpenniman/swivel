"""Property-based tests against an independent brute-force oracle (no Excel)."""
import datetime as dt
import math

import openpyxl
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

import swivel
from conftest import check_package, comparable, decode, write_book
from swivel._values import Accumulator, XLError, item_key, to_serial

_val = st.one_of(
    st.none(),
    st.sampled_from(["a", "A", "b", "é", "e", "x y", "Zed"]),
    st.integers(-3, 3),
    st.floats(-100, 100, allow_nan=False).map(lambda x: round(x, 3)),
    st.sampled_from([{"$date": "2024-01-01T00:00:00"}, {"$date": "2023-06-30T00:00:00"}]),
    st.sampled_from([{"$error": "#N/A"}]),
)
_key = st.one_of(st.none(), st.sampled_from(["p", "P", "q", "r s"]), st.integers(0, 3), st.booleans())
FUNCS = ["sum", "count", "average", "max", "min", "count_nums", "stdevp", "varp"]


@st.composite
def crosstab(draw):
    n = draw(st.integers(1, 30))
    data = [["R", "C", "V"]] + [[draw(_key), draw(_key), draw(_val)] for _ in range(n)]
    return data, draw(st.sampled_from(FUNCS))


def _oracle(c):
    if isinstance(c, dict) and "$error" in c:
        return XLError(c["$error"])
    return decode(c)


def _label_key(label):
    """Map a rendered row/column label back to the grouping key of its item."""
    if label == "(blank)":
        return ("m",)
    if label in ("TRUE", "FALSE"):
        return ("b", label == "TRUE")
    return item_key(label)


def _close(a, b):
    a, b = comparable(a), comparable(b)
    if isinstance(a, float) and isinstance(b, float):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    return a == b


@settings(max_examples=150, deadline=None, suppress_health_check=list(HealthCheck))
@given(case=crosstab())
def test_crosstab_matches_brute_force(case, tmp_path_factory):
    data, func = case
    path = write_book(tmp_path_factory.mktemp("p") / "p.xlsx", data)
    res = swivel.add_pivot(path, source="Data", target="P", rows=["R"], columns=["C"],
                           values=["V:" + func])
    check_package(path)
    records = [[_oracle(c) for c in r] for r in data[1:]]
    # the auto-detected range ends at the last row holding anything (min. one row)
    while len(records) > 1 and all(v is None for v in records[-1]):
        records.pop()
    ws = openpyxl.load_workbook(path)["P"]
    top = 3  # anchor A3: caption row, column-label row, then data
    ncols = res.ref.split(":")[1]
    width = openpyxl.utils.column_index_from_string("".join(filter(str.isalpha, ncols)))
    height = int("".join(filter(str.isdigit, ncols)))
    col_keys = {}
    for c in range(2, width + 1):
        lab = ws.cell(top + 1, c).value
        col_keys[c] = None if lab == "Grand Total" else _label_key(lab)
    blank_cross = func in ("max", "min", "product")
    for r in range(top + 2, height + 1):
        lab = ws.cell(r, 1).value
        rk = None if lab == "Grand Total" else _label_key(lab)
        for c in range(2, width + 1):
            ck = col_keys[c]
            sel = [rec for rec in records
                   if (rk is None or item_key(rec[0]) == rk) and (ck is None or item_key(rec[1]) == ck)]
            acc = Accumulator()
            for rec in sel:
                acc.add(rec[2])
            numeric_only = blank_cross and rk is None and ck is None
            expected = acc.result(func, numeric_only=numeric_only) if sel else None
            got = res.cells.get(ws.cell(r, c).coordinate)
            if acc.error is not None and isinstance(expected, XLError) and isinstance(got, XLError):
                # which error a total shows depends on item order (see model.py)
                assert got in [rec[2] for rec in sel], (got, sel)
                continue
            assert _close(got, expected), (ws.cell(r, c).coordinate, lab, got, expected)
    # every item appears exactly once, case-insensitively merged
    assert len({_label_key(ws.cell(r, 1).value) for r in range(top + 2, height)}) == \
        len({item_key(rec[0]) for rec in records})


@st.composite
def anything(draw):
    ncols = draw(st.integers(2, 6))
    n = draw(st.integers(1, 20))
    headers = ["H%d" % i for i in range(ncols)]
    data = [headers] + [[draw(st.one_of(_val, _key)) for _ in range(ncols)] for _ in range(n)]
    perm = draw(st.permutations(headers))
    nr = draw(st.integers(0, min(3, ncols - 1)))
    nc = draw(st.integers(0, min(3, ncols - 1 - nr)))
    np_ = draw(st.integers(0, min(2, ncols - nr - nc)))
    values = draw(st.lists(st.tuples(st.sampled_from(headers), st.sampled_from(FUNCS)),
                           min_size=1, max_size=4))
    return data, dict(rows=list(perm[:nr]), columns=list(perm[nr:nr + nc]),
                      filters=list(perm[nr + nc:nr + nc + np_]),
                      values=[{"field": f, "func": fn, "name": "v%d" % i}
                              for i, (f, fn) in enumerate(values)],
                      row_grand_totals=draw(st.booleans()),
                      column_grand_totals=draw(st.booleans()))


@settings(max_examples=200, deadline=None, suppress_health_check=list(HealthCheck))
@given(case=anything())
def test_any_spec_produces_a_consistent_package(case, tmp_path_factory):
    data, spec = case
    path = write_book(tmp_path_factory.mktemp("a") / "a.xlsx", data)
    res = swivel.add_pivot(path, source="Data", target="P", **spec)
    check_package(path)
    wb = openpyxl.load_workbook(path)        # openpyxl parses every pivot part
    (pt,) = wb["P"]._pivots
    assert pt.location.ref == res.ref
    from swivel._xml import parse_cell, parse_range
    t0, tc0, t1, tc1 = parse_range(res.ref)
    # the declared items account for exactly the table's rows and columns
    assert len(pt.rowItems) == (t1 - t0 + 1) - pt.location.firstDataRow
    assert len(pt.colItems) == (tc1 - tc0 + 1) - pt.location.firstDataCol
    # every rendered cell lies inside the reported footprint
    r0, c0, r1, c1 = parse_range(res.full_ref)
    for ref in res.cells:
        r, c = parse_cell(ref)
        assert r0 <= r <= r1 and c0 <= c <= c1, ref
    # a second run is a no-op structurally
    again = swivel.add_pivot(path, source="Data", target="P", **spec)
    assert again.cells == res.cells or _same(again.cells, res.cells)
    assert len(swivel.list_pivots(path)) == 1


def _same(a, b):
    return a.keys() == b.keys() and all(_close(a[k], b[k]) for k in a)

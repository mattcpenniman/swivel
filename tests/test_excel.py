"""Differential tests: real Excel opens swivel's output, refreshes it, and must
show exactly the cells swivel pre-rendered.

Skipped automatically when Excel/pywin32 are unavailable (CI, macOS, Linux).
Set SWIVEL_NO_EXCEL=1 to skip them on a Windows box too.
"""
import datetime as dt
import glob
import itertools
import json
import os
import random

import openpyxl
import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

import swivel
from conftest import (GOLDEN, SALES, assert_cells_equal, check_package, decode,
                      write_book)
from swivel.excel import snapshot

pytestmark = pytest.mark.excel


def excel_matches(app, path, results, *, expect_total=None):
    """Excel refreshes every pivot; each must equal what swivel rendered."""
    # 1) as stored: Excel accepts the definition and agrees on its footprint
    stored = snapshot(path, app=app, refresh_first=False)
    # 2) after an explicit refresh from the source data
    fresh = snapshot(path, app=app, refresh_first=True)
    n = expect_total if expect_total is not None else len(results)
    assert len(fresh) == n, "Excel sees %d pivot tables, expected %d" % (len(fresh), n)
    by_stored = {(s.sheet, s.name): s for s in stored}
    by_fresh = {(s.sheet, s.name): s for s in fresh}
    for r in results:
        key = (r.sheet, r.name)
        assert key in by_fresh, "Excel dropped pivot %s (repaired file?)" % (key,)
        s0, s1 = by_stored[key], by_fresh[key]
        assert s0.table_range1 == r.ref == s1.table_range1, (s0.table_range1, r.ref, s1.table_range1)
        assert s1.table_range2 == r.full_ref
        assert_cells_equal(r.cells, s0.cells, "%s (as stored)" % (key,))
        assert_cells_equal(r.cells, s1.cells, "%s (after refresh)" % (key,))


# --------------------------------------------------------------------------
# Every golden scenario, but now Excel renders *swivel's* file
# --------------------------------------------------------------------------

CASES = sorted(os.path.basename(p)[:-5] for p in glob.glob(os.path.join(GOLDEN, "*.json")))


@pytest.mark.parametrize("case", CASES)
def test_golden_scenarios_in_excel(case, tmp_path, excel):
    with open(os.path.join(GOLDEN, case + ".json"), encoding="utf-8") as fh:
        g = json.load(fh)
    path = write_book(tmp_path / "g.xlsx", g["data"])
    spec = dict(g["spec"])
    spec["filters"] = [tuple(f) for f in spec.get("filters", [])]
    res = swivel.add_pivot(path, source="Data", target="Pivot", **spec)
    excel_matches(excel, path, [res])


# --------------------------------------------------------------------------
# A grid of layouts over a dataset with every value type
# --------------------------------------------------------------------------

MIXED = [
    ["Region", "Rep", "Year", "Quarter", "Amount", "Units", "Note", "Flag", "Day"],
]
_rng = random.Random(7)
_regions = ["East", "west", "North", None, "Süd", "co-op", "coop"]
_reps = ["Ann", "bob", "Bob", "Cy", "Dee", "  Eve", "O'Neil", "Zed"]
for i in range(60):
    MIXED.append([
        _rng.choice(_regions), _rng.choice(_reps), _rng.choice([2022, 2023, 2024, 2024.5]),
        _rng.choice(["Q1", "Q2", "Q3", "Q4", None, 3]),
        _rng.choice([None, 0, -1.5, 2.25, 10, 1e6, 0.1, "n/a", {"$error": "#N/A"}]),
        _rng.choice([None, 1, 2, 3, 4, True]),
        _rng.choice(["x", None, "y" * 300, "multi\nline", "a&b<c>\"d'"]),
        _rng.choice([True, False, None]),
        {"$date": (dt.datetime(2024, 1, 1) + dt.timedelta(days=_rng.randrange(0, 5))).isoformat()},
    ])

GRID = list(itertools.product(
    [[], ["Region"], ["Region", "Rep"], ["Year", "Quarter", "Flag"]],
    [[], ["Quarter"], ["Flag", "Year"]],
    [["Amount"], ["Amount:average", "Units:count"], ["Units:max", "Amount:sum", "Amount:min"]],
))


def _grid_id(p):
    r, c, v = p
    return "R%s-C%s-V%s" % ("+".join(r) or "0", "+".join(c) or "0", "+".join(v))


@pytest.mark.parametrize("rows,cols,values", [g for g in GRID if not set(g[0]) & set(g[1])],
                         ids=[_grid_id(g) for g in GRID if not set(g[0]) & set(g[1])])
def test_layout_grid_in_excel(rows, cols, values, tmp_path, excel):
    path = write_book(tmp_path / "grid.xlsx", MIXED)
    res = swivel.add_pivot(path, source="Data", target="P", rows=rows, columns=cols,
                           values=values)
    check_package(path)
    excel_matches(excel, path, [res])


@pytest.mark.parametrize("filters", [
    [("Region", None)],
    [("Region", "East")],
    [("Region", ["East", "west"]), ("Flag", True)],
    [("Day", dt.datetime(2024, 1, 2)), ("Year", 2024)],
    [("Region", [None, "co-op"])],
])
def test_filters_in_excel(filters, tmp_path, excel):
    path = write_book(tmp_path / "f.xlsx", MIXED)
    res = swivel.add_pivot(path, source="Data", target="P", rows=["Rep"], columns=["Quarter"],
                           values=["Amount", "Units:count"], filters=filters)
    excel_matches(excel, path, [res])


# --------------------------------------------------------------------------
# Random data and random specs (hypothesis), checked by Excel
# --------------------------------------------------------------------------

_cell = st.one_of(
    st.none(),
    st.sampled_from(["a", "A", "b", "B ", "c-d", "cd", "é", "e", "zz", "10"]),
    st.integers(-5, 5),
    st.sampled_from([0.5, -2.25, 1e-3, 123456.789]),
    st.booleans(),
    st.sampled_from([{"$error": "#N/A"}, {"$error": "#DIV/0!"}]),
    st.sampled_from([{"$date": "2024-02-29T00:00:00"}, {"$date": "2023-12-31T00:00:00"},
                     {"$date": "2024-02-29T13:30:00"}]),
)


@st.composite
def pivot_case(draw):
    ncols = draw(st.integers(3, 6))
    nrows = draw(st.integers(1, 25))
    headers = ["F%d" % i for i in range(ncols)]
    data = [headers] + [[draw(_cell) for _ in range(ncols)] for _ in range(nrows)]
    fields = list(headers)
    draw(st.randoms()).shuffle(fields)
    nr = draw(st.integers(0, min(2, ncols - 1)))
    nc = draw(st.integers(0, min(2, ncols - 1 - nr)))
    npg = draw(st.integers(0, min(1, ncols - nr - nc)))
    rows, cols, pages = fields[:nr], fields[nr:nr + nc], fields[nr + nc:nr + nc + npg]
    funcs = ["sum", "count", "average", "max", "min", "product", "count_nums",
             "stdev", "stdevp", "var", "varp"]
    nv = draw(st.integers(1, 3))
    values = [{"field": draw(st.sampled_from(headers)), "func": draw(st.sampled_from(funcs)),
               "name": "V%d" % i} for i in range(nv)]
    filters = []
    for p in pages:
        col = [decode(r[headers.index(p)]) for r in data[1:]]
        choice = draw(st.sampled_from(["all", "one"]))
        if choice == "one":
            filters.append((p, draw(st.sampled_from(col))))
        else:
            filters.append((p, None))
    return data, dict(rows=rows, columns=cols, values=values, filters=filters)


@settings(max_examples=int(os.environ.get("SWIVEL_EXCEL_EXAMPLES", "60")), deadline=None,
          suppress_health_check=list(HealthCheck), derandomize=True)
@given(case=pivot_case())
def test_random_pivots_in_excel(case, tmp_path_factory, excel):
    data, spec = case
    path = write_book(tmp_path_factory.mktemp("h") / "h.xlsx", data)
    res = swivel.add_pivot(path, source="Data", target="P", **spec)
    check_package(path)
    excel_matches(excel, path, [res])


# --------------------------------------------------------------------------
# Lifecycle: re-runs, Excel-saved files, tables, existing content
# --------------------------------------------------------------------------

def test_rerun_after_data_change(tmp_path, excel):
    path = write_book(tmp_path / "r.xlsx", SALES)
    spec = dict(source="Data", target="Summary", rows=["Region"], columns=["Year"],
                values=["Amount"])
    swivel.add_pivots(path, [spec])
    wb = openpyxl.load_workbook(path)
    wb["Data"].append(["South", "Kiwi", 2025, 99, 9])
    wb["Data"]["D2"] = 1000
    wb.save(path)  # openpyxl keeps pivot parts; swivel then rebuilds the sheet
    res = swivel.add_pivots(path, [spec])
    check_package(path)
    assert "South" in res[0].cells.values()
    assert 1000 in res[0].cells.values()  # East / 2023 after the edit
    excel_matches(excel, path, res)


def test_rerun_on_file_saved_by_excel(tmp_path, excel):
    path = write_book(tmp_path / "x.xlsx", SALES)
    spec = dict(source="Data", target="Summary", rows=["Region", "Product"], values=["Amount", "Qty"])
    swivel.add_pivots(path, [spec])
    resaved = str(tmp_path / "resaved.xlsx")
    swivel.excel.refresh(path, output=resaved, app=excel)   # Excel rewrites every part
    res = swivel.add_pivots(resaved, [spec, dict(spec, target="Other", columns=["Year"],
                                                 rows=["Region"])])
    check_package(resaved)
    excel_matches(excel, resaved, res)


def test_table_source_in_excel(tmp_path, excel):
    from openpyxl.worksheet.table import Table

    def add_table(wb):
        wb["Data"].add_table(Table(displayName="Sales", ref="A1:E7"))
    path = write_book(tmp_path / "t.xlsx", SALES, extra=add_table)
    res = swivel.add_pivot(path, source="Sales", target="P", rows=["Product"], values=["Qty:sum"])
    assert res.source == "Sales"
    excel_matches(excel, path, [res])


def test_existing_excel_pivot_and_chart_survive(tmp_path, excel):
    """Excel builds a pivot + chart; swivel adds its own; all must survive."""
    path = write_book(tmp_path / "e.xlsx", SALES)
    import os as _os
    full = _os.path.abspath(path)
    wb = excel.Workbooks.Open(full)
    try:
        src = wb.Worksheets("Data").Range("A1:E7")
        pc = wb.PivotCaches().Create(1, src)
        sh = wb.Worksheets.Add()
        sh.Name = "Theirs"
        pt = pc.CreatePivotTable(sh.Range("A3"), "TheirPivot")
        pt.PivotFields("Region").Orientation = 1
        pt.AddDataField(pt.PivotFields("Qty"), "Total Qty", -4157)
        ch = wb.Worksheets("Data").ChartObjects().Add(300, 10, 300, 200)
        ch.Chart.SetSourceData(wb.Worksheets("Data").Range("D1:D7"))
        wb.Save()
    finally:
        wb.Close(False)
    res = swivel.add_pivot(path, source="Data", target="Mine", rows=["Product"], values=["Amount"])
    check_package(path)
    excel_matches(excel, path, [res], expect_total=2)
    wb = excel.Workbooks.Open(full, 0, True)
    try:
        assert wb.Worksheets("Data").ChartObjects().Count == 1
        assert wb.Worksheets("Theirs").PivotTables(1).Name == "TheirPivot"
    finally:
        wb.Close(False)


def test_many_pivots_one_sheet_and_number_formats(tmp_path, excel):
    path = write_book(tmp_path / "m.xlsx", MIXED)
    specs = [
        dict(source="Data", target="Dash", rows=["Region"], values=[
            {"field": "Amount", "func": "sum", "number_format": "#,##0.00"}]),
        dict(source="Data", target="Dash", rows=["Day"], columns=["Flag"],
             values=["Units:count"], name="ByDay"),
        dict(source="Data", target="Dash", rows=["Rep"], values=["Units:average"],
             filters=["Year"], anchor="A40", name="Below"),
    ]
    res = swivel.add_pivots(path, specs)
    check_package(path)
    excel_matches(excel, path, res)


def test_large_source_in_excel(tmp_path, excel):
    rnd = random.Random(1)
    rows = [["K1", "K2", "V"]] + [[rnd.choice("ABCDEFGHIJ"), rnd.randrange(40), rnd.random()]
                                  for _ in range(20000)]
    path = write_book(tmp_path / "big.xlsx", rows)
    res = swivel.add_pivot(path, source="Data", target="P", rows=["K1"], columns=["K2"],
                           values=["V", "V:max"])
    excel_matches(excel, path, [res])


def test_formatted_numbers_and_dates_as_column_groups(tmp_path, excel):
    def fmt(wb):
        ws = wb["Data"]
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 1).number_format = "#,##0.00"
            ws.cell(r, 2).number_format = "mmm d, yyyy"
    rows = [["Num", "When", "K", "V"]]
    for i in range(12):
        rows.append([1000 * (i % 3), {"$date": "2024-0%d-1%dT00:00:00" % (1 + i % 2, i % 3)},
                     "k%d" % (i % 2), i])
    path = write_book(tmp_path / "fmt.xlsx", rows, extra=fmt)
    res = swivel.add_pivots(path, [
        dict(source="Data", target="A", rows=["K"], columns=["Num", "When"], values=["V"]),
        dict(source="Data", target="B", rows=["K"], columns=["When", "Num"], values=["V"]),
    ])
    excel_matches(excel, path, res)


# Cases hypothesis found against Excel, pinned as plain regressions.
REGRESSIONS = {
    # text + logical items are not "mixed"; flagging them so made Excel refuse the file
    "text_and_bool_items": ([["F0", "F1"], [None, "a"], [None, False]],
                            dict(columns=["F0", "F1"], values=["F0"])),
    # a number below 0 in a date field once produced a minDate Excel rejects
    "negative_number_in_date_field": ([["F0", "F1"], [None, {"$date": "2024-02-29T00:00:00"}],
                                       [None, -1]], dict(values=["F0"])),
    "huge_number_in_date_field": ([["F0", "F1"], [None, {"$date": "2024-02-29T00:00:00"}],
                                   [None, 1e9]], dict(values=["F0"])),
    # dates sort after text in a mixed field; subtotal labels use the date format
    "text_before_dates": ([["F0", "F1"], [None, "a"], [None, {"$date": "2024-02-29T00:00:00"}]],
                          dict(columns=["F0", "F1"], values=["F0"])),
    # numbers in a date field get date-formatted subtotal labels ("1900-01-00 0:00:00 Total")
    "number_in_date_field_label": ([["F0", "F1", "F5"], [None, None, {"$date": "2024-02-29T00:00:00"}],
                                    [None, None, 0], [None, None, 45351.25], [None, None, -1],
                                    [None, None, 123456.789]],
                                   dict(columns=["F5", "F1"], values=["F0"])),
    # count shows the error in totals once another value field is not a count
    "count_error_in_totals": ([["K", "J", "V", "W"], ["a", "x", {"$error": "#N/A"}, 5],
                               ["a", "y", 1, 6], ["b", "x", 2, 7]],
                              dict(rows=["K", "J"], values=["V:count", "V:count_nums", "W:sum"])),
    "count_error_in_totals_alone": ([["K", "J", "V"], ["a", "x", {"$error": "#N/A"}],
                                     ["a", "y", 1]], dict(rows=["K", "J"], values=["V:count"])),
    # with a max field present, both-way totals ignore text (count 1, not 2)
    "numeric_only_cross_totals": ([["F0", "F1", "F2"], [None, None, 0], [None, "a", "a"],
                                   [None, None, None]],
                                  dict(rows=["F0", "F1"], columns=["F2"],
                                       values=["F0:max", "F2:count", "F2:sum", "F2:average"])),
    # a total shows the error of its first child item (rows before columns)
    "error_precedence": ([["R", "C", "V"], ["r2", "c1", {"$error": "#DIV/0!"}],
                          ["r1", "c2", {"$error": "#N/A"}], ["r1", "c1", 1],
                          ["r1", "c2", {"$error": "#REF!"}]],
                         dict(rows=["R"], columns=["C"], values=["V", "V:max"])),
    # an error in a both-ways total is shown, not blanked
    "error_in_cross_total": ([["F0", "F1", "F3"], [None, None, {"$error": "#N/A"}]],
                             dict(rows=["F0"], columns=["F1"], values=["F3:max"])),
}


@pytest.mark.parametrize("name", sorted(REGRESSIONS))
def test_excel_regressions(name, tmp_path, excel):
    data, spec = REGRESSIONS[name]
    path = write_book(tmp_path / "reg.xlsx", data)
    res = swivel.add_pivot(path, source="Data", target="P", **spec)
    excel_matches(excel, path, [res])

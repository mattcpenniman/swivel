"""Regenerate the golden files by having *Excel itself* build each pivot table.

Run on Windows with Excel installed:

    python tests/golden/regenerate.py

For every scenario below, a workbook is written with openpyxl, Excel creates
the pivot table through COM (no swivel involved), and the cells Excel renders
are stored in ``<scenario>.json``. ``tests/test_golden.py`` then checks that
swivel renders exactly the same cells - so the suite runs anywhere while the
expectations remain Excel's own output.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", ".."))

XL_FUNC = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139,
           "product": -4149, "count_nums": -4113, "stdev": -4155, "stdevp": -4156,
           "var": -4164, "varp": -4165}

SALES = [
    ["Region", "Product", "Year", "Amount", "Qty", "When"],
    ["East", "Apple", 2023, 10.5, 1, {"$date": "2024-01-05T00:00:00"}],
    ["West", "Pear", 2024, 3, 2, {"$date": "2024-02-05T00:00:00"}],
    ["East", "pear", 2024, 7.25, 3, {"$date": "2024-01-05T00:00:00"}],
    ["North", "Apple", 2023, None, 4, {"$date": "2024-03-05T00:00:00"}],
    [None, "Fig", 2024, 1, 5, {"$date": "2024-03-05T00:00:00"}],
    ["West", "Apple", 2023, 2, None, {"$date": "2024-03-09T00:00:00"}],
]

TYPES = [
    ["A", "B", "C", "N", "T", "Flag", "Err", "Mix"],
    ["a1", "b1", "c1", 1, "x", True, 1, 1],
    ["a1", "b1", "c2", 2, "y", False, {"$error": "#N/A"}, "s"],
    ["a1", "b2", "c1", 3, None, True, 3, None],
    ["a2", "b1", "c1", None, "z", None, 4, True],
    ["a2", "b2", "c2", 5, "w", False, 5, 2.5],
]

WORDS = [
    ["Name", "Cat", "V"],
    ["co-op", "k", 1], ["coop", "k", 2], ["Zeta", "k", 3], ["alpha", "j", 4],
    ["ALPHA", "j", 5], ["Beta", "j", 6], ["10", "j", 7], [10, "k", 8], [2.5, "k", 9],
    ["b'c", "k", 10], ["bc", "j", 11], ["Ébène", "k", 12], ["eagle", "j", 13],
    [True, "j", 14], [None, "k", 15], ["", "k", 16], ["  pad", "j", 17],
]


def v(name, func="sum", caption=None):
    d = {"field": name, "func": func}
    if caption:
        d["name"] = caption
    return d


SCENARIOS = {
    "r1_d1": (SALES, dict(rows=["Region"], values=[v("Amount")])),
    "r2_d1": (SALES, dict(rows=["Region", "Product"], values=[v("Amount")])),
    "r1_c1_d1": (SALES, dict(rows=["Region"], columns=["Year"], values=[v("Amount")])),
    "r1_d2": (SALES, dict(rows=["Region"], values=[v("Amount"), v("Qty", "count", "Count of Qty")])),
    "r1_c1_d2": (SALES, dict(rows=["Region"], columns=["Year"],
                            values=[v("Amount"), v("Qty", "average", "Average of Qty")])),
    "r0_c1_d1": (SALES, dict(columns=["Year"], values=[v("Amount")])),
    "r0_c0_d1": (SALES, dict(values=[v("Amount")])),
    "r1_d1_p1": (SALES, dict(rows=["Product"], values=[v("Amount")], filters=[["Region", None]])),
    "rdate_d1": (SALES, dict(rows=["When"], values=[v("Qty")])),
    "r2_c2_d1": (SALES, dict(rows=["Region", "Product"], columns=["Year", "Qty"], values=[v("Amount")])),
    "r0_d2": (TYPES, dict(values=[v("N"), v("T", "count", "Count of T")])),
    "r0_c1_d2": (TYPES, dict(columns=["A"], values=[v("N"), v("T", "count", "Count of T")])),
    "r1_c2_d2": (TYPES, dict(rows=["C"], columns=["A", "B"],
                            values=[v("N", "sum", "Sum of N"), v("N", "count", "Count of N2")])),
    "c3_d1": (TYPES, dict(rows=["T"], columns=["A", "B", "C"], values=[v("N")])),
    "r3_d1": (TYPES, dict(rows=["A", "B", "C"], values=[v("N")])),
    "types": (TYPES, dict(rows=["Flag"], columns=["Mix"], values=[
        v("T", "sum", "Sum of T"), v("T", "count", "Count of T"), v("T", "average", "Average of T2"),
        v("N", "max", "Max of N"), v("N", "min", "Min of N"), v("Err", "sum", "Sum of Err"),
        v("N", "count", "Count of N"), v("N", "average", "Average of N")])),
    "funcs": (TYPES, dict(rows=["A"], values=[
        v("N", "product", "Product of N"), v("N", "count_nums", "Count of N"),
        v("N", "stdev", "StdDev of N"), v("N", "stdevp", "StdDevp of N"),
        v("N", "var", "Var of N"), v("N", "varp", "Varp of N")])),
    "page_sel": (TYPES, dict(rows=["B"], values=[v("N")], filters=[["A", ["a2"]]])),
    "page_multi": (TYPES, dict(rows=["B"], values=[v("N")],
                               filters=[["C", ["c1"]], ["A", ["a1", "a2"]]])),
    "page_hidden": (TYPES, dict(rows=["A"], values=[v("N")], filters=[["B", ["b1"]], ["C", ["c1", "c2"]], ["T", ["x", "z"]]])),
    "word_sort": (WORDS, dict(rows=["Name"], values=[v("V")])),
    "word_sort_cols": (WORDS, dict(rows=["Cat"], columns=["Name"], values=[v("V", "max", "Max of V")])),
    "no_grand": (SALES, dict(rows=["Region"], columns=["Year"], values=[v("Amount")],
                             row_grand_totals=False, column_grand_totals=False)),
}


def decode(x):
    if isinstance(x, dict) and "$date" in x:
        return dt.datetime.fromisoformat(x["$date"])
    if isinstance(x, dict) and "$error" in x:
        return x["$error"]  # openpyxl stores error codes as error cells
    return x


def encode(x):
    if isinstance(x, dt.datetime):
        return {"$date": x.isoformat()}
    return x


def anchor_row(spec):
    return max(3, len(spec.get("filters", [])) + 2)


def build_with_excel(xl, name, data, spec, folder):
    path = os.path.join(folder, name + ".xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in data:
        ws.append([decode(c) for c in row])
    wb.save(path)
    book = xl.Workbooks.Open(path)
    try:
        src = book.Worksheets("Data").Range("A1").CurrentRegion
        src = book.Worksheets("Data").Range(
            "A1", book.Worksheets("Data").Cells(len(data), len(data[0])))
        pc = book.PivotCaches().Create(1, src)
        tgt = book.Worksheets.Add()
        tgt.Name = "Pivot"
        top = anchor_row(spec)
        pt = pc.CreatePivotTable(tgt.Range("A%d" % top), "PT1")
        for i, f in enumerate(spec.get("rows", [])):
            pt.PivotFields(f).Orientation = 1
            pt.PivotFields(f).Position = i + 1
        for i, f in enumerate(spec.get("columns", [])):
            pt.PivotFields(f).Orientation = 2
            pt.PivotFields(f).Position = i + 1
        # Excel stacks new report filters on top, so add them bottom-up
        for f, _ in reversed(spec.get("filters", [])):
            pt.PivotFields(f).Orientation = 3
        for val in spec["values"]:
            df = pt.AddDataField(pt.PivotFields(val["field"]), None, XL_FUNC[val["func"]])
            if val.get("name"):
                df.Caption = val["name"]
        for f, sel in spec.get("filters", []):
            if sel is None:
                continue
            pf = pt.PivotFields(f)
            if len(sel) == 1:
                pf.CurrentPage = sel[0]
            else:
                pf.EnableMultiplePageItems = True
                items = [pf.PivotItems(i) for i in range(1, pf.PivotItems().Count + 1)]
                for it in items:  # show the selection first so one stays visible
                    if it.Name in sel:
                        it.Visible = True
                for it in items:
                    if it.Name not in sel:
                        it.Visible = False
        if spec.get("row_grand_totals") is False:
            pt.RowGrand = False
        if spec.get("column_grand_totals") is False:
            pt.ColumnGrand = False
        rng = pt.TableRange2
        cells = {}
        vals = rng.Value2
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        from swivel.excel import _norm
        from swivel._xml import cell_ref
        for i, row in enumerate(vals):
            for j, val in enumerate(row):
                if val is not None and val != "":
                    cells[cell_ref(rng.Row - 1 + i, rng.Column - 1 + j)] = _norm(val)
        return {
            "data": data, "spec": spec, "anchor": "A%d" % top,
            "table_range1": pt.TableRange1.GetAddress(False, False),
            "table_range2": rng.GetAddress(False, False),
            "cells": cells,
            "excel_version": xl.Version,
        }
    finally:
        book.Close(False)


def main(only=None):
    from swivel.excel import excel_app
    with tempfile.TemporaryDirectory() as tmp, excel_app() as xl:
        for name, (data, spec) in SCENARIOS.items():
            if only and name not in only:
                continue
            g = build_with_excel(xl, name, data, spec, tmp)
            with open(os.path.join(HERE, name + ".json"), "w", encoding="utf-8") as fh:
                json.dump(g, fh, indent=1, ensure_ascii=False, sort_keys=True)
            print("wrote", name)


if __name__ == "__main__":
    main(sys.argv[1:])

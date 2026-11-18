"""Spec handling, source detection and every refusal path."""
import datetime as dt

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

import swivel
from conftest import SALES, check_package, write_book
from swivel import PivotSpec, SwivelError, Value


def test_value_coercion():
    assert Value.coerce("Amount") == Value("Amount", "sum")
    assert Value.coerce("Amount:avg") == Value("Amount", "average")
    assert Value.coerce("Ratio: 1:2") == Value("Ratio: 1:2", "sum")  # not a function suffix
    assert Value.coerce(("Qty", "count")) == Value("Qty", "count")
    assert Value.coerce({"field": "Qty", "func": "max", "name": "Peak"}).name == "Peak"
    with pytest.raises(SwivelError):
        Value("Qty", "median")
    with pytest.raises(SwivelError):
        Value.coerce(42)


def test_spec_validation():
    with pytest.raises(SwivelError, match="at least one value"):
        PivotSpec(source="Data", target="P", values=[])
    with pytest.raises(SwivelError, match="only one of"):
        PivotSpec(source="Data", target="P", values="A", rows=["X"], columns=["x"])
    with pytest.raises(SwivelError, match="31"):
        PivotSpec(source="Data", target="P" * 32, values="A")
    with pytest.raises(SwivelError, match="characters"):
        PivotSpec(source="Data", target="a/b", values="A")
    with pytest.raises(SwivelError, match="unknown pivot option"):
        PivotSpec.from_dict({"source": "Data", "target": "P", "values": "A", "colour": "red"})
    s = PivotSpec(source="Data", target="P", values="Amount", rows="Region", filters="Year")
    assert s.rows == ["Region"] and s.filters == [("Year", None)]


def test_source_forms(tmp_path):
    def extras(wb):
        ws = wb.create_sheet("Offset")
        ws["C5"], ws["D5"] = "K", "V"
        for i, (k, v) in enumerate([("a", 1), ("b", 2), ("a", 3)], start=6):
            ws.cell(i, 3, k)
            ws.cell(i, 4, v)
        ws["F5"] = "not part of the table (gap)"
        wb["Data"].add_table(Table(displayName="SalesTbl", ref="A1:E7", totalsRowCount=0))
    path = write_book(tmp_path / "s.xlsx", SALES, extra=extras)
    r = swivel.add_pivot(path, source="Data", target="P1", rows=["Region"], values=["Qty"])
    assert r.source == "Data!A1:E7" and r.records == 6
    r = swivel.add_pivot(path, source="Data!A1:E4", target="P2", rows=["Region"], values=["Qty"])
    assert r.records == 3
    r = swivel.add_pivot(path, source="salestbl", target="P3", rows=["Region"], values=["Qty"])
    assert r.source == "SalesTbl"
    r = swivel.add_pivot(path, source="Offset", target="P4", rows=["K"], values=["V"],
                         header_row=5)
    assert r.source == "Offset!C5:D8" and r.cells["B4"] == 4
    check_package(path)


def test_table_totals_row_is_excluded(tmp_path):
    def extras(wb):
        ws = wb["Data"]
        ws.append(["Total", None, None, 23.75, 15])
        wb["Data"].add_table(Table(displayName="T", ref="A1:E8", totalsRowCount=1))
    path = write_book(tmp_path / "t.xlsx", SALES, extra=extras)
    r = swivel.add_pivot(path, source="T", target="P", values=["Qty"])
    assert r.records == 6 and r.cells["A4"] == 15


def test_trailing_blank_rows_and_ragged_data(book):
    rows = SALES + [[None] * 5, [None] * 5, ["Late", None, None, 5, None]]
    path = book(rows)
    r = swivel.add_pivot(path, source="Data", target="P", rows=["Region"], values=["Amount:count"])
    assert r.source == "Data!A1:E10"
    assert "Late" in r.cells.values()


def test_empty_source_gives_an_empty_but_valid_pivot(book):
    path = book([["A", "B"]])
    r = swivel.add_pivot(path, source="Data", target="P", rows=["A"], values=["B"])
    assert r.records == 1
    assert r.cells["A4"] == "(blank)"
    check_package(path)
    with pytest.raises(SwivelError, match="no data rows"):
        swivel.add_pivot(path, source="Data!A1:B1", target="P", rows=["A"], values=["B"])


def test_headers(book):
    path = book([["Amt", "amt", 2024, None], [1, 2, 3, 4]])
    with pytest.raises(SwivelError, match="no header"):
        swivel.add_pivot(path, source="Data!A1:D2", target="P", values=["Amt"])
    r = swivel.add_pivot(path, source="Data", target="P", values=["amt2", "2024"])
    assert [c for c in r.cells.values() if isinstance(c, str)] == ["Sum of amt2", "Sum of 2024"]


@pytest.mark.parametrize("kwargs,match", [
    (dict(source="Nope"), "not a sheet"),
    (dict(source="Nope!A1:B2"), "not found"),
    (dict(source="Data!A1:ZZZZ9"), "column|cell"),
    (dict(rows=["Missing"]), "not a column"),
    (dict(values=["Missing"]), "not a column"),
    (dict(filters={"Region": "Atlantis"}), "not found in field"),
    (dict(anchor="A2", filters=["Region", "Year"]), "no room"),
    (dict(anchor="ZZ"), "cell"),
    (dict(values=[{"field": "Amount", "name": "Qty"}]), "name of a source column"),
])
def test_errors(kwargs, match, sales):
    base = dict(source="Data", target="P", values=["Amount"])
    base.update(kwargs)
    with pytest.raises(SwivelError, match=match):
        swivel.add_pivots(sales, [base])


def test_duplicate_default_value_names_are_numbered(sales):
    r = swivel.add_pivot(sales, source="Data", target="P", values=["Amount", "Amount"])
    assert r.cells["A3"] == "Sum of Amount" and r.cells["B3"] == "Sum of Amount2"
    with pytest.raises(SwivelError, match="both named"):
        swivel.add_pivot(sales, source="Data", target="P",
                         values=[Value("Amount", name="X"), Value("Qty", name="x")])


def test_several_pivots_on_one_sheet_are_laid_out_side_by_side(sales):
    res = swivel.add_pivots(sales, [
        dict(source="Data", target="P", rows=["Region"], values=["Amount"]),
        dict(source="Data", target="P", rows=["Product"], columns=["Year"], values=["Qty"],
             filters=["Region"]),
        dict(source="Data", target="P", values=["Qty"], name="Tiny"),
    ])
    assert [r.full_ref for r in res] == ["A3:B8", "D1:G8", "I3:I4"]
    assert [r.name for r in res] == ["PivotTable1", "PivotTable2", "Tiny"]
    check_package(sales)
    with pytest.raises(SwivelError, match="overlap"):
        swivel.add_pivots(sales, [
            dict(source="Data", target="P", rows=["Region"], values=["Amount"]),
            dict(source="Data", target="P", values=["Qty"], anchor="B5")])
    with pytest.raises(SwivelError, match="named"):
        swivel.add_pivots(sales, [
            dict(source="Data", target="P", values=["Qty"], name="Same"),
            dict(source="Data", target="P", values=["Qty"], name="same")])


def test_filter_selection_forms(sales):
    one = swivel.add_pivot(sales, source="Data", target="P", rows=["Product"], values=["Qty"],
                           filters={"Region": "east"})          # case-insensitive
    assert one.cells["B1"] == "East"
    multi = swivel.add_pivot(sales, source="Data", target="P", rows=["Product"], values=["Qty"],
                             filters={"Region": ["East", None]})
    assert multi.cells["B1"] == "(Multiple Items)"
    num = swivel.add_pivot(sales, source="Data", target="P", rows=["Product"], values=["Qty"],
                           filters={"Year": "2024"})             # "2024" finds 2024
    assert num.cells["B1"] == 2024


def test_not_an_xlsx(tmp_path):
    p = tmp_path / "x.xlsx"
    p.write_text("hello")
    with pytest.raises(SwivelError, match="not an .xlsx"):
        swivel.add_pivot(p, source="Data", target="P", values=["A"])
    with pytest.raises(FileNotFoundError):
        swivel.add_pivot(tmp_path / "missing.xlsx", source="Data", target="P", values=["A"])


def test_formula_sources_use_cached_values(tmp_path):
    """Sources are read as Excel last calculated them, like Excel's own pivot cache."""
    path = write_book(tmp_path / "f.xlsx", [["K", "V"], ["a", 2], ["b", 3]])
    wb = openpyxl.load_workbook(path)
    wb["Data"]["C1"] = "W"
    wb["Data"]["C2"] = "=B2*10"   # openpyxl stores no cached value
    wb["Data"]["C3"] = "=B3*10"
    wb.save(path)
    r = swivel.add_pivot(path, source="Data", target="P", rows=["K"], values=["W:count"])
    # no cached values -> blanks now; Excel recalculates and refreshes on open
    assert r.cells.get("B4") is None


def test_xlsm_keeps_its_content_type(tmp_path, sales):
    import zipfile
    with zipfile.ZipFile(sales) as z:
        data = {n: z.read(n) for n in z.namelist()}
    data["[Content_Types].xml"] = data["[Content_Types].xml"].replace(
        b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        b"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    out = tmp_path / "m.xlsm"
    with zipfile.ZipFile(out, "w") as w:
        for n, d in data.items():
            w.writestr(n, d)
    swivel.add_pivot(out, source="Data", target="P", values=["Qty"])
    with zipfile.ZipFile(out) as z:
        assert b"application/vnd.ms-excel.sheet.macroEnabled.main+xml" in z.read("[Content_Types].xml")


def test_list_pivots(sales):
    swivel.add_pivots(sales, [dict(source="Data", target="P", rows=["Region"], values=["Qty"],
                                   name="ByRegion")])
    (p,) = swivel.list_pivots(sales)
    assert (p.sheet, p.name, p.ref, p.source, p.refresh_on_load) == \
        ("P", "ByRegion", "A3:B8", "Data!A1:E7", True)

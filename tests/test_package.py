from pathlib import Path
"""The written file: structure, preservation of everything else, idempotency, safety."""
import datetime as dt
import io
import os
import zipfile

import openpyxl
import pytest
from lxml import etree
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.image import Image
from openpyxl.worksheet.table import Table

import swivel
from conftest import SALES, check_package, write_book
from swivel import SwivelError

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SPEC = dict(source="Data", target="Summary", rows=["Region"], columns=["Year"], values=["Amount"])


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def pivot_parts(path):
    return sorted(n for n in parts(path) if "pivot" in n.lower())


def test_structure_and_flags(sales):
    swivel.add_pivots(sales, [SPEC])
    names = check_package(sales)
    assert "xl/pivotTables/pivotTable1.xml" in names
    p = parts(sales)
    cache = etree.fromstring(p["xl/pivotCache/pivotCacheDefinition1.xml"])
    assert cache.get("refreshOnLoad") == "1"
    assert cache.get("recordCount") == "6"
    assert cache.find(M + "cacheSource/" + M + "worksheetSource").get("ref") == "A1:E7"
    recs = etree.fromstring(p["xl/pivotCache/pivotCacheRecords1.xml"])
    assert recs.get("count") == "6" and len(recs) == 6
    assert all(len(r) == 5 for r in recs)


def test_openpyxl_can_read_the_result(sales):
    swivel.add_pivots(sales, [SPEC])
    wb = openpyxl.load_workbook(sales)
    ws = wb["Summary"]
    assert len(ws._pivots) == 1
    pt = ws._pivots[0]
    assert pt.location.ref == "A3:D9"
    assert pt.cache.refreshOnLoad
    assert ws["A5"].value == "East" and ws["D5"].value == 17.75
    wb.save(io.BytesIO())  # and can write it back


def test_untouched_parts_are_byte_identical(tmp_path):
    def extras(wb):
        ws = wb["Data"]
        chart = BarChart()
        chart.add_data(Reference(ws, min_col=4, min_row=1, max_row=7), titles_from_data=True)
        ws.add_chart(chart, "H2")
        from PIL import Image as PILImage
        buf = io.BytesIO()
        PILImage.new("RGB", (4, 4), "red").save(buf, "PNG")
        buf.seek(0)
        ws.add_image(Image(buf), "H20")
        ws.add_table(Table(displayName="Sales", ref="A1:E7"))
        wb.create_sheet("Notes")["A1"] = "keep me"
    pytest.importorskip("PIL")
    path = write_book(tmp_path / "rich.xlsx", SALES, extra=extras)
    before = parts(path)
    swivel.add_pivots(path, [SPEC])
    after = parts(path)
    owned = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
             "docProps/app.xml", "xl/styles.xml"}
    for n, data in before.items():
        if n in owned:
            continue
        assert n in after, "lost part %s" % n
        assert after[n] == data, "part %s was modified" % n
    check_package(path)
    wb = openpyxl.load_workbook(path)
    assert wb["Notes"]["A1"].value == "keep me"
    assert wb.sheetnames == ["Data", "Notes", "Summary"]


def test_vba_and_custom_parts_survive(tmp_path, sales):
    # graft a fake vbaProject + customXml part into the package
    data = parts(sales)
    data["xl/vbaProject.bin"] = b"\x00\x01binary\xff"
    data["customXml/item1.xml"] = b"<root>mine</root>"
    out = tmp_path / "graft.xlsx"
    with zipfile.ZipFile(out, "w") as w:
        for n, d in data.items():
            if n == "[Content_Types].xml":
                d = d.replace(b"</Types>", b'<Default Extension="bin" ContentType="application/'
                              b'vnd.ms-office.vbaProject"/></Types>')
            w.writestr(n, d)
    swivel.add_pivots(out, [SPEC])
    after = parts(out)
    assert after["xl/vbaProject.bin"] == b"\x00\x01binary\xff"
    assert after["customXml/item1.xml"] == b"<root>mine</root>"


def test_rerun_is_idempotent(sales):
    swivel.add_pivots(sales, [SPEC])
    first = parts(sales)
    for _ in range(3):
        swivel.add_pivots(sales, [SPEC])
    again = parts(sales)
    assert pivot_parts(sales) == ["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                                  "xl/pivotCache/pivotCacheDefinition1.xml",
                                  "xl/pivotCache/pivotCacheRecords1.xml",
                                  "xl/pivotTables/_rels/pivotTable1.xml.rels",
                                  "xl/pivotTables/pivotTable1.xml"]
    check_package(sales)
    assert openpyxl.load_workbook(sales).sheetnames == ["Data", "Summary"]
    # identical except the refresh timestamp
    strip = lambda b: b.split(b'refreshedDate="')[0]
    assert strip(first["xl/pivotCache/pivotCacheDefinition1.xml"]) == \
        strip(again["xl/pivotCache/pivotCacheDefinition1.xml"])
    assert first["xl/pivotTables/pivotTable1.xml"] == again["xl/pivotTables/pivotTable1.xml"]
    assert first["xl/worksheets/sheet2.xml"] == again["xl/worksheets/sheet2.xml"]


def test_rerun_with_fewer_pivots_removes_orphans(sales):
    swivel.add_pivots(sales, [SPEC, dict(SPEC, rows=["Product"], name="Second")])
    assert len(swivel.list_pivots(sales)) == 2
    swivel.add_pivots(sales, [SPEC])
    assert len(swivel.list_pivots(sales)) == 1
    names = check_package(sales)
    assert not any("pivotCacheDefinition2" in n or "pivotTable2" in n for n in names)


def test_replacing_keeps_sheet_position_and_other_pivots(sales):
    swivel.add_pivots(sales, [dict(SPEC, target="One"), dict(SPEC, target="Two", rows=["Product"])])
    wb = openpyxl.load_workbook(sales)
    wb.move_sheet("Two", offset=-2)
    wb.save(sales)
    swivel.add_pivots(sales, [dict(SPEC, target="One", values=["Qty:count"])])
    assert openpyxl.load_workbook(sales).sheetnames == ["Two", "Data", "One"]
    info = {p.sheet: p for p in swivel.list_pivots(sales)}
    assert set(info) == {"One", "Two"}
    check_package(sales)


def test_refuses_to_wipe_a_data_sheet(tmp_path):
    def extras(wb):
        ws = wb.create_sheet("Budget")
        ws.append(["Line", "Plan"])
        ws.append(["Rent", 1200])
    path = write_book(tmp_path / "b.xlsx", SALES, extra=extras)
    with pytest.raises(SwivelError, match="outside its pivot tables"):
        swivel.add_pivots(path, [dict(SPEC, target="budget")])  # typo'd target
    assert openpyxl.load_workbook(path)["Budget"]["B2"].value == 1200


def test_refuses_target_with_foreign_content(tmp_path):
    def extras(wb):
        swivel_sheet = wb.create_sheet("Summary")
        swivel_sheet["Z99"] = "hand typed"
    path = write_book(tmp_path / "f.xlsx", SALES, extra=extras)
    with pytest.raises(SwivelError, match="outside its pivot tables"):
        swivel.add_pivots(path, [SPEC])
    res = swivel.add_pivots(path, [SPEC], if_exists="overwrite")
    assert "Z99" not in res[0].cells
    assert b"hand typed" not in parts(path)["xl/worksheets/sheet2.xml"]
    with pytest.raises(SwivelError, match="already exists"):
        swivel.add_pivots(path, [SPEC], if_exists="error")


def test_values_inside_old_pivot_area_are_fine(sales):
    """A sheet that only holds swivel's own (bigger) pivot is replaceable."""
    swivel.add_pivots(sales, [dict(SPEC, rows=["Region", "Product"], filters=["Qty"])])
    swivel.add_pivots(sales, [SPEC])
    check_package(sales)


def test_refusal_leaves_file_untouched(tmp_path):
    def extras(wb):
        wb.create_sheet("Summary")["A1"] = "precious"
    path = write_book(tmp_path / "f.xlsx", SALES, extra=extras)
    before = Path(path).read_bytes()
    with pytest.raises(SwivelError):
        swivel.add_pivots(path, [SPEC])
    assert Path(path).read_bytes() == before


def test_refuses_sheet_with_drawings(tmp_path):
    def extras(wb):
        ws = wb.create_sheet("Summary")
        chart = BarChart()
        chart.add_data(Reference(wb["Data"], min_col=4, min_row=1, max_row=7))
        ws.add_chart(chart, "A1")
    path = write_book(tmp_path / "d.xlsx", SALES, extra=extras)
    with pytest.raises(SwivelError, match="drawing"):
        swivel.add_pivots(path, [SPEC], if_exists="overwrite")


def test_source_and_target_must_differ(sales):
    with pytest.raises(SwivelError, match="both a pivot source and a pivot target"):
        swivel.add_pivots(sales, [dict(SPEC, target="data")])


def test_output_path_leaves_input_alone(sales, tmp_path):
    before = Path(sales).read_bytes()
    out = tmp_path / "out.xlsx"
    swivel.add_pivots(sales, [SPEC], output=out)
    assert Path(sales).read_bytes() == before
    check_package(out)


def test_no_records_mode(sales):
    swivel.add_pivots(sales, [dict(SPEC, save_records=False)])
    names = check_package(sales)
    assert not any("pivotCacheRecords" in n for n in names)
    cache = etree.fromstring(parts(sales)["xl/pivotCache/pivotCacheDefinition1.xml"])
    assert cache.get("saveData") == "0" and cache.get("refreshOnLoad") == "1"


def test_refresh_on_open_toggle(sales):
    swivel.add_pivots(sales, [dict(SPEC, refresh_on_load=False)])
    assert [p.refresh_on_load for p in swivel.list_pivots(sales)] == [False]
    assert swivel.set_refresh_on_open(sales) == 1
    assert [p.refresh_on_load for p in swivel.list_pivots(sales)] == [True]
    assert swivel.set_refresh_on_open(sales) == 0
    assert swivel.set_refresh_on_open(sales, False) == 1
    check_package(sales)


def test_number_formats_registered_once(sales):
    v = {"field": "Amount", "func": "sum", "number_format": "#,##0.000;[Red]-#,##0.000"}
    swivel.add_pivots(sales, [dict(SPEC, values=[v, dict(v, func="max")])])
    swivel.add_pivots(sales, [dict(SPEC, values=[v])])
    styles = etree.fromstring(parts(sales)["xl/styles.xml"])
    fmts = [f.get("formatCode") for f in styles.iter(M + "numFmt")]
    assert fmts.count("#,##0.000;[Red]-#,##0.000") == 1
    wb = openpyxl.load_workbook(sales)
    assert wb["Summary"]["B5"].number_format == "#,##0.000;[Red]-#,##0.000"
    builtin = dict(v, number_format="0.00")
    swivel.add_pivots(sales, [dict(SPEC, values=[builtin])])
    assert openpyxl.load_workbook(sales)["Summary"]["B5"].number_format == "0.00"


def test_dates_get_a_date_format(book):
    path = book([["D", "V"], [dt.datetime(2024, 5, 6), 1], [dt.datetime(2024, 5, 7, 8, 30), 2]])
    swivel.add_pivot(path, source="Data", target="P", rows=["D"], values=["V"])
    ws = openpyxl.load_workbook(path)["P"]
    assert ws["A4"].value == dt.datetime(2024, 5, 6)
    assert ws["A4"].is_date


def test_awkward_text_round_trips(book):
    nasty = ["tab\there", "nl\nhere", "_x0041_ literal", "<&>\"'",
             "  lead", "trail  ", "emoji 🚀", "x" * 400, "日本語"]
    rows = [["K", "V"]] + [[s, i] for i, s in enumerate(nasty)]
    path = book(rows)
    swivel.add_pivot(path, source="Data", target="P", rows=["K"], values=["V"])
    check_package(path)
    wb = openpyxl.load_workbook(path)
    labels = {wb["P"].cell(r, 1).value for r in range(4, 4 + len(nasty))}
    assert {"  lead", "trail  ", "emoji 🚀", "x" * 400, "日本語", "<&>\"'"} <= labels
    cache = etree.fromstring(parts(path)["xl/pivotCache/pivotCacheDefinition1.xml"])
    k = cache.find(M + "cacheFields")[0].find(M + "sharedItems")
    assert k.get("longText") == "1"


def test_unicode_and_quoted_sheet_names(tmp_path):
    path = write_book(tmp_path / "u.xlsx", SALES, sheet="Données 'Q1'")
    res = swivel.add_pivot(path, source="'Données ''Q1'''!A1:E7", target="Résumé & Co",
                           rows=["Region"], values=["Amount"])
    assert res.source == "Données 'Q1'!A1:E7"
    check_package(path)
    assert openpyxl.load_workbook(path).sheetnames[-1] == "Résumé & Co"
    res = swivel.add_pivot(path, source="Données 'Q1'", target="Résumé & Co",
                           rows=["Region"], values=["Amount"])
    assert res.records == 6


EXCEL_APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="4" baseType="variant">'
    '<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>'
    '<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>'
    '</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr">'
    '<vt:lpstr>Data</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts>'
    '</Properties>')


def test_app_xml_lists_new_sheet(sales, tmp_path):
    data = parts(sales)
    data["docProps/app.xml"] = EXCEL_APP_XML.encode()
    path = tmp_path / "app.xlsx"
    with zipfile.ZipFile(path, "w") as w:
        for n, d in data.items():
            w.writestr(n, d)
    swivel.add_pivots(path, [SPEC])
    app = etree.fromstring(parts(path)["docProps/app.xml"])
    vt = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
    titles = [e.text for e in app.iter(vt + "lpstr")]
    assert titles == ["Worksheets", "Named Ranges", "Data", "Summary", "Data!Print_Area"]
    assert [e.text for e in app.iter(vt + "i4")] == ["2", "1"]
    vec = app.find("{*}TitlesOfParts/" + vt + "vector")
    assert vec.get("size") == "3"
    swivel.add_pivots(path, [SPEC])  # re-run does not add it twice
    app = etree.fromstring(parts(path)["docProps/app.xml"])
    assert [e.text for e in app.iter(vt + "lpstr")].count("Summary") == 1


def test_large_dataset_is_fast(tmp_path):
    import random
    import time
    rnd = random.Random(0)
    rows = [["A", "B", "C", "V", "W"]]
    rows += [[rnd.choice("PQRSTUVWXY"), rnd.randrange(100), rnd.choice(["x", "y", None]),
              rnd.random() * 100, rnd.randrange(10)] for _ in range(100_000)]
    path = write_book(tmp_path / "big.xlsx", rows)
    t = time.perf_counter()
    res = swivel.add_pivot(path, source="Data", target="P", rows=["A", "C"], columns=["W"],
                           values=["V", "V:average", "B:max"])
    elapsed = time.perf_counter() - t
    assert res.records == 100_000
    check_package(path)
    assert elapsed < 60, elapsed

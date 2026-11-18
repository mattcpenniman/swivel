"""Serialise a :class:`PivotModel` into the three OOXML pivot parts and a sheet."""
from __future__ import annotations

import datetime as _dt
from typing import Callable, Iterable

from ._values import EXCEL_EPOCH, FUNCTIONS, XLError, is_number, to_serial
from ._xml import (MAIN_NS, REL_NS, XML_DECL, attr, cell_ref, col_letter, esc, num,
                   range_ref)
from .model import DATA, CacheField, PivotModel


def _iso(d: _dt.datetime) -> str:
    return d.strftime("%Y-%m-%dT%H:%M:%S")


def _item(v) -> str:
    """One shared item / record value."""
    if v is None:
        return "<m/>"
    if isinstance(v, bool):
        return '<b v="%d"/>' % v
    if is_number(v):
        return '<n v="%s"/>' % num(v)
    if isinstance(v, _dt.datetime):
        return '<d v="%s"/>' % _iso(v)
    if isinstance(v, XLError):
        return '<e v="%s"/>' % attr(v.code)
    return '<s v="%s"/>' % attr(v)


def _shared_items_attrs(values: Iterable) -> list[tuple[str, str]]:
    """The type flags Excel writes on <sharedItems> (rules taken from Excel's own files).

    Text, logicals and errors count as one "string" kind; a field is "mixed"
    when it combines more than one of string / number / date. Excel rejects a
    file whose flags contradict the items, so these must be exact.
    """
    has = {"s": False, "n": False, "d": False, "b": False, "e": False, "m": False}
    integer = True
    lo = hi = None
    dlo = dhi = None
    long_text = False
    for v in values:
        if v is None:
            has["m"] = True
        elif isinstance(v, bool):
            has["b"] = True
        elif is_number(v):
            has["n"] = True
            if not float(v).is_integer():
                integer = False
            lo = v if lo is None or v < lo else lo
            hi = v if hi is None or v > hi else hi
        elif isinstance(v, _dt.datetime):
            has["d"] = True
            dlo = v if dlo is None or v < dlo else dlo
            dhi = v if dhi is None or v > dhi else dhi
        elif isinstance(v, XLError):
            has["e"] = True
        else:
            has["s"] = True
            if len(v) > 255:
                long_text = True
    stringish = has["s"] or has["b"] or has["e"]
    kinds = int(stringish) + int(has["n"]) + int(has["d"])
    out = []
    if not (stringish or has["m"]):
        out.append(("containsSemiMixedTypes", "0"))
    if not (stringish or has["n"]):
        out.append(("containsNonDate", "0"))
    if has["d"]:
        out.append(("containsDate", "1"))
    if not stringish:
        out.append(("containsString", "0"))
    if has["m"]:
        out.append(("containsBlank", "1"))
    if kinds > 1:
        out.append(("containsMixedTypes", "1"))
    if has["n"] and not has["d"]:
        out.append(("containsNumber", "1"))
        if integer:
            out.append(("containsInteger", "1"))
        out.append(("minValue", num(lo)))
        out.append(("maxValue", num(hi)))
    if has["d"]:
        # numbers in a date field are reported as dates too; Excel stores the
        # day after the latest date as maxDate
        cands_lo = [dlo] + ([_from_serial(lo)] if has["n"] else [])
        cands_hi = [dhi] + ([_from_serial(hi)] if has["n"] else [])
        top = max(cands_hi)
        top = _dt.datetime(top.year, top.month, top.day)
        top = top + _dt.timedelta(days=1) if top < _MAX_DATE else _MAX_DATE
        out.append(("minDate", _iso(max(min(cands_lo), EXCEL_EPOCH))))
        out.append(("maxDate", _iso(top)))
    if long_text:
        out.append(("longText", "1"))
    return out


_MAX_DATE = _dt.datetime(9999, 12, 31)


def _from_serial(x: float) -> _dt.datetime:
    """Serial -> datetime, clamped to the range Excel accepts in min/maxDate
    (an earlier minDate makes Excel refuse to open the file)."""
    x = float(x)
    if x <= 0:
        return EXCEL_EPOCH
    try:
        return min(EXCEL_EPOCH + _dt.timedelta(days=x), _MAX_DATE)
    except OverflowError:
        return _MAX_DATE


def cache_definition_xml(model: PivotModel, *, records_rid: str | None,
                         numfmt_id: Callable[[str | None], int],
                         refreshed: _dt.datetime) -> bytes:
    src = model.source
    spec = model.spec
    head = ['<pivotCacheDefinition xmlns="%s" xmlns:r="%s"' % (MAIN_NS, REL_NS)]
    if records_rid:
        head.append(' r:id="%s"' % records_rid)
    else:
        head.append(' saveData="0"')
    if spec.refresh_on_load:
        head.append(' refreshOnLoad="1"')
    head.append(' refreshedBy="swivel" refreshedDate="%s" createdVersion="6"'
                ' refreshedVersion="6" minRefreshableVersion="3" recordCount="%d">'
                % (num(to_serial(refreshed)), src.record_count))
    out = ["".join(head)]
    if src.table:
        out.append('<cacheSource type="worksheet"><worksheetSource name="%s"/></cacheSource>'
                   % attr(src.table))
    else:
        out.append('<cacheSource type="worksheet"><worksheetSource ref="%s" sheet="%s"/>'
                   '</cacheSource>' % (src.ref, attr(src.sheet)))
    out.append('<cacheFields count="%d">' % len(model.fields))
    for cf in model.fields:
        nf = numfmt_id(cf.date_format) if cf.date_format else 0
        out.append('<cacheField name="%s" numFmtId="%d">' % (attr(cf.name), nf))
        values = cf.shared if cf.shared is not None else cf.values
        attrs = "".join(' %s="%s"' % kv for kv in _shared_items_attrs(values))
        if cf.shared is None:
            out.append("<sharedItems%s/>" % attrs)
        else:
            out.append('<sharedItems%s count="%d">' % (attrs, len(cf.shared)))
            out.extend(_item(v) for v in cf.shared)
            out.append("</sharedItems>")
        out.append("</cacheField>")
    out.append("</cacheFields></pivotCacheDefinition>")
    return (XML_DECL + "".join(out)).encode("utf-8")


def cache_records_chunks(model: PivotModel) -> Iterable[str]:
    """Yield pivotCacheRecords XML in pieces (it can be large)."""
    n = model.source.record_count
    yield XML_DECL
    yield '<pivotCacheRecords xmlns="%s" xmlns:r="%s" count="%d">' % (MAIN_NS, REL_NS, n)
    fields: list[CacheField] = model.fields
    cols = []
    for cf in fields:
        if cf.shared is not None:
            cols.append(("x", cf.record_index))
        else:
            cols.append(("v", cf.values))
    buf = []
    for r in range(n):
        parts = ["<r>"]
        for kind, data in cols:
            if kind == "x":
                parts.append('<x v="%d"/>' % data[r])
            else:
                parts.append(_item(data[r]))
        parts.append("</r>")
        buf.append("".join(parts))
        if len(buf) >= 2000:
            yield "".join(buf)
            buf = []
    if buf:
        yield "".join(buf)
    yield "</pivotCacheRecords>"


def _x(v: int) -> str:
    return "<x/>" if v == 0 else '<x v="%d"/>' % v


def location(model: PivotModel, top: int, left: int) -> str:
    return range_ref(top, left, top + model.height - 1, left + model.width - 1)


def pivot_table_xml(model: PivotModel, *, name: str, cache_id: int, top: int, left: int,
                    numfmt_id: Callable[[str | None], int]) -> bytes:
    spec = model.spec
    o = ['<pivotTableDefinition xmlns="%s" name="%s" cacheId="%d"'
         ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"'
         ' applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"'
         ' dataCaption="Values" updatedVersion="6" minRefreshableVersion="3"'
         ' useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0"'
         % (MAIN_NS, attr(name), cache_id)]
    if not spec.row_grand_totals:
        o.append(' rowGrandTotals="0"')
    if not spec.column_grand_totals:
        o.append(' colGrandTotals="0"')
    o.append(' multipleFieldFilters="0">')

    loc = ('<location ref="%s" firstHeaderRow="%d" firstDataRow="%d" firstDataCol="%d"'
           % (location(model, top, left), model.first_header_row, model.first_data_row,
              model.first_data_col))
    if model.page_fields:
        loc += ' rowPageCount="%d" colPageCount="1"' % len(model.page_fields)
    o.append(loc + "/>")

    axis = {}
    for f in model.row_fields:
        axis[f] = "axisRow"
    for f in model.col_fields:
        axis[f] = "axisCol"
    hidden = {}
    for f, sel in model.page_fields:
        axis[f] = "axisPage"
        if sel is not None and len(sel) > 1:
            hidden[f] = set(range(len(model.fields[f].shared))) - set(sel)
    data_flds = {d.field for d in model.data_fields}

    o.append('<pivotFields count="%d">' % len(model.fields))
    for i, cf in enumerate(model.fields):
        a = ["<pivotField"]
        if i in axis:
            a.append(' axis="%s"' % axis[i])
        if i in data_flds:
            a.append(' dataField="1"')
        if cf.date_format:
            a.append(' numFmtId="%d"' % numfmt_id(cf.date_format))
        if i in hidden:
            a.append(' multipleItemSelectionAllowed="1"')
        a.append(' showAll="0"')
        if i in axis and axis[i] != "axisPage":
            a.append(' sortType="ascending"')
        if i not in axis:
            o.append("".join(a) + "/>")
            continue
        o.append("".join(a) + ">")
        o.append('<items count="%d">' % (len(cf.order) + 1))
        hid = hidden.get(i, ())
        for ci in cf.order:
            o.append('<item x="%d"%s/>' % (ci, ' h="1"' if ci in hid else ""))
        o.append('<item t="default"/></items></pivotField>')
    o.append("</pivotFields>")

    if model.row_fields:
        o.append('<rowFields count="%d">' % len(model.row_fields))
        o.extend('<field x="%d"/>' % f for f in model.row_fields)
        o.append("</rowFields>")
    o.append('<rowItems count="%d">' % len(model.row_entries))
    for e in model.row_entries:
        if e.type == "item":
            lvl = len(e.path) - 1
            o.append("<i%s>%s</i>" % (' r="%d"' % lvl if lvl else "", _x(e.path[-1])))
        elif e.type == "grand":
            o.append('<i t="grand"><x/></i>')
        else:
            o.append("<i/>")
    o.append("</rowItems>")

    col_axis = model.col_axis
    if col_axis:
        o.append('<colFields count="%d">' % len(col_axis))
        o.extend('<field x="%d"/>' % f for f in col_axis)
        o.append("</colFields>")
    o.append('<colItems count="%d">' % len(model.col_entries))
    prev = None
    multi = model.multi_data
    for e in model.col_entries:
        iattr = ' i="%d"' % e.data if e.data else ""
        if e.type == "data":
            xs = list(e.path) + ([e.data] if multi else [])
            r = 0
            if prev:
                for a_, b_ in zip(prev, xs):
                    if a_ != b_:
                        break
                    r += 1
            ra = ' r="%d"' % r if r else ""
            o.append("<i%s%s>%s</i>" % (ra, iattr, "".join(_x(v) for v in xs[r:])))
            prev = xs
        elif e.type == "default":
            lvl = len(e.path) - 1
            ra = ' r="%d"' % lvl if lvl else ""
            o.append('<i t="default"%s%s>%s</i>' % (ra, iattr, _x(e.path[-1])))
            prev = list(e.path)
        elif e.type == "grand":
            o.append('<i t="grand"%s><x/></i>' % iattr)
            prev = None
        else:
            o.append("<i/>")
    o.append("</colItems>")

    if model.page_fields:
        o.append('<pageFields count="%d">' % len(model.page_fields))
        for f, sel in model.page_fields:
            if sel is not None and len(sel) == 1:
                o.append('<pageField fld="%d" item="%d" hier="-1"/>'
                         % (f, model.fields[f].pos[sel[0]]))
            else:
                o.append('<pageField fld="%d" hier="-1"/>' % f)
        o.append("</pageFields>")

    o.append('<dataFields count="%d">' % len(model.data_fields))
    for d in model.data_fields:
        sub = FUNCTIONS[d.value.func][0]
        a = '<dataField name="%s" fld="%d"' % (attr(d.name), d.field)
        if sub != "sum":
            a += ' subtotal="%s"' % sub
        a += ' baseField="0" baseItem="0"'
        if d.value.number_format:
            a += ' numFmtId="%d"' % numfmt_id(d.value.number_format)
        o.append(a + "/>")
    o.append("</dataFields>")
    o.append('<pivotTableStyleInfo name="%s" showRowHeaders="1" showColHeaders="1"'
             ' showRowStripes="0" showColStripes="0" showLastColumn="1"/>' % attr(spec.style))
    # Excel 2010+ hides the "Values" caption row by default; without this the
    # table grows by one row after a refresh when there are no column fields.
    o.append('<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"'
             ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
             '<x14:pivotTableDefinition hideValuesRow="1"'
             ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst>')
    o.append("</pivotTableDefinition>")
    return (XML_DECL + "".join(o)).encode("utf-8")


# --------------------------------------------------------------------------
# Worksheet
# --------------------------------------------------------------------------

def _cell_xml(r: int, c: int, value, style: int) -> str:
    ref = cell_ref(r, c)
    s = ' s="%d"' % style if style else ""
    if isinstance(value, XLError):
        return '<c r="%s"%s t="e"><v>%s</v></c>' % (ref, s, esc(value.code))
    if isinstance(value, bool):
        return '<c r="%s"%s t="b"><v>%d</v></c>' % (ref, s, value)
    if is_number(value):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, num(value))
    if isinstance(value, _dt.datetime):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, num(to_serial(value)))
    text = str(value)
    space = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return '<c r="%s"%s t="inlineStr"><is><t%s>%s</t></is></c>' % (ref, s, space, esc(text))


def _display_len(value) -> int:
    if isinstance(value, _dt.datetime):
        return 10
    if is_number(value):
        return min(len(num(round(value, 2) if isinstance(value, float) else value)), 15)
    if isinstance(value, XLError):
        return len(value.code)
    return max((len(line) for line in str(value).split("\n")), default=0)


def worksheet_xml(cells: dict, style_id: Callable[[str | None], int]) -> bytes:
    """A worksheet holding exactly ``cells`` ({(row0, col0): Cell})."""
    rows: dict = {}
    widths: dict = {}
    for (r, c), cell in cells.items():
        rows.setdefault(r, []).append((c, cell))
        w = _display_len(cell.value)
        widths[c] = max(widths.get(c, 0), w)
    out = [XML_DECL, '<worksheet xmlns="%s" xmlns:r="%s">' % (MAIN_NS, REL_NS)]
    if cells:
        rs = [k[0] for k in cells]
        cs = [k[1] for k in cells]
        out.append('<dimension ref="%s"/>' % range_ref(min(rs), min(cs), max(rs), max(cs)))
    else:
        out.append('<dimension ref="A1"/>')
    out.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="15"/>')
    if widths:
        out.append("<cols>")
        for c in sorted(widths):
            w = min(max(widths[c] + 2, 9), 60)
            out.append('<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (c + 1, c + 1, w))
        out.append("</cols>")
    out.append("<sheetData>")
    for r in sorted(rows):
        out.append('<row r="%d">' % (r + 1))
        for c, cell in sorted(rows[r], key=lambda t: t[0]):
            fmt = cell.number_format
            out.append(_cell_xml(r, c, cell.value, style_id(fmt) if fmt else 0))
        out.append("</row>")
    out.append("</sheetData>")
    out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3"'
               ' footer="0.3"/>')
    out.append("</worksheet>")
    return "".join(out).encode("utf-8")


__all__ = ["cache_definition_xml", "cache_records_chunks", "pivot_table_xml",
           "worksheet_xml", "location", "col_letter", "DATA"]

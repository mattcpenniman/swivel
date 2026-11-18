"""Minimal, lossless editing of an .xlsx package (an OPC zip).

swivel never round-trips a workbook through a spreadsheet library: every
part it does not own (charts, images, VBA, slicers, custom XML ...) is copied
byte-for-byte. Only ``[Content_Types].xml``, ``workbook.xml``, the workbook
relationships, ``styles.xml`` and the parts swivel itself writes are edited.
"""
from __future__ import annotations

import io
import os
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Iterable

from lxml import etree

from ._xml import (CT_NS, CT_PIVOT_CACHE, CT_PIVOT_RECORDS, CT_PIVOT_TABLE, CT_WORKSHEET,
                   MAIN_NS, PKG_REL_NS, REL_NS, RT_CALC_CHAIN, RT_OFFICE_DOC,
                   RT_PIVOT_CACHE, RT_PIVOT_RECORDS, RT_PIVOT_TABLE, RT_STYLES, RT_TABLE,
                   RT_WORKSHEET, parse_range)
from .spec import SwivelError

_M = "{%s}" % MAIN_NS
_R = "{%s}" % REL_NS

_WORKBOOK_ORDER = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews",
    "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr",
    "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes",
    "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
]
_STYLES_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
                 "cellStyles", "dxfs", "tableStyles", "colors", "extLst"]

# Relationship types a swivel-owned sheet may carry and that we know how to drop.
_DROPPABLE = {RT_PIVOT_TABLE, REL_NS + "/printerSettings"}

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False,
                          huge_tree=True, no_network=True)


def _parse(data: bytes):
    return etree.fromstring(data, _PARSER)


def _serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


@dataclass
class Rel:
    id: str
    type: str
    target: str
    external: bool = False


@dataclass
class SheetInfo:
    name: str
    sheet_id: int
    rid: str
    part: str
    rel_type: str


@dataclass
class PivotInfo:
    sheet: str
    part: str
    name: str
    cache_id: int
    ref: str
    page_rows: int
    page_cols: int
    cache_part: str | None
    refresh_on_load: bool = False
    source: str = ""


@dataclass
class TableInfo:
    sheet: str
    name: str
    display_name: str
    ref: str
    header_rows: int
    totals_rows: int


class Package:
    def __init__(self, data: bytes):
        try:
            zf = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            raise SwivelError("not an .xlsx file (it is not a zip package)") from None
        self._order: list[str] = []
        self.parts: dict[str, bytes] = {}
        for info in zf.infolist():
            if info.is_dir():
                continue
            self._order.append(info.filename)
            self.parts[info.filename] = zf.read(info.filename)
        self._lower = {n.lower(): n for n in self.parts}
        if "[content_types].xml" not in self._lower:
            raise SwivelError("not an .xlsx file ([Content_Types].xml is missing)")
        self._xml_cache: dict[str, object] = {}
        self._ro_cache: dict[str, object] = {}
        self.workbook_part = self._find_workbook()
        if self.workbook_part.lower().endswith(".bin"):
            raise SwivelError("binary workbooks (.xlsb) are not supported")
        if "macroEnabled" in self.content_type(self.workbook_part) or "template" in \
                self.content_type(self.workbook_part):
            pass  # .xlsm / .xltx are fine: the workbook part is still XML

    # ---- raw parts -------------------------------------------------------
    @classmethod
    def from_file(cls, path) -> "Package":
        with open(path, "rb") as fh:
            return cls(fh.read())

    def name(self, part: str) -> str | None:
        return self._lower.get(part.lstrip("/").lower())

    def has(self, part: str) -> bool:
        return self.name(part) is not None

    def read(self, part: str) -> bytes:
        n = self.name(part)
        if n is None:
            raise KeyError(part)
        if n in self._xml_cache:
            return _serialize(self._xml_cache[n])
        return self.parts[n]

    def write(self, part: str, data: bytes) -> None:
        part = part.lstrip("/")
        n = self.name(part)
        if n is None:
            n = part
            self._order.append(n)
            self._lower[n.lower()] = n
        self._xml_cache.pop(n, None)
        self._ro_cache.pop(n, None)
        self.parts[n] = data

    def delete(self, part: str) -> None:
        n = self.name(part)
        if n is None:
            return
        self._xml_cache.pop(n, None)
        self._ro_cache.pop(n, None)
        del self.parts[n]
        del self._lower[n.lower()]
        self._order.remove(n)
        self._remove_override(n)
        rp = self.rels_path(n)
        if self.has(rp):
            self.delete(rp)

    def xml(self, part: str):
        """Parsed, cached, mutable tree of an XML part (written back on save)."""
        n = self.name(part)
        if n is None:
            raise KeyError(part)
        if n not in self._xml_cache:
            self._xml_cache[n] = self._ro_cache.pop(n, None)
            if self._xml_cache[n] is None:
                self._xml_cache[n] = _parse(self.parts[n])
        return self._xml_cache[n]

    def peek(self, part: str):
        """Parsed tree for reading only; the part's bytes are left untouched."""
        n = self.name(part)
        if n is None:
            raise KeyError(part)
        if n in self._xml_cache:
            return self._xml_cache[n]
        if n not in self._ro_cache:
            self._ro_cache[n] = _parse(self.parts[n])
        return self._ro_cache[n]

    def to_bytes(self) -> bytes:
        for n, root in self._xml_cache.items():
            self.parts[n] = _serialize(root)
        self._xml_cache.clear()
        buf = io.BytesIO()
        order = sorted(self._order, key=lambda n: 0 if n == self.name("[Content_Types].xml") else 1)
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for n in order:
                zf.writestr(n, self.parts[n])
        return buf.getvalue()

    def save(self, path) -> None:
        data = self.to_bytes()
        path = os.fspath(path)
        d = os.path.dirname(os.path.abspath(path))
        fd, tmp = tempfile.mkstemp(prefix=".swivel-", suffix=".xlsx", dir=d)
        try:
            with os.fdopen(fd, "wb") as fh:
                fh.write(data)
            os.replace(tmp, path)
        except BaseException:
            try:
                os.unlink(tmp)
            except OSError:
                pass
            raise

    # ---- content types ---------------------------------------------------
    def _ct(self):
        return self.xml("[Content_Types].xml")

    def content_type(self, part: str) -> str:
        pn = "/" + part.lstrip("/")
        root = self._ct()
        for o in root.iter("{%s}Override" % CT_NS):
            if o.get("PartName", "").lower() == pn.lower():
                return o.get("ContentType", "")
        ext = pn.rsplit(".", 1)[-1].lower()
        for d in root.iter("{%s}Default" % CT_NS):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType", "")
        return ""

    def add_override(self, part: str, ctype: str) -> None:
        self._remove_override(part)
        el = etree.SubElement(self._ct(), "{%s}Override" % CT_NS)
        el.set("PartName", "/" + part.lstrip("/"))
        el.set("ContentType", ctype)

    def _remove_override(self, part: str) -> None:
        pn = "/" + part.lstrip("/")
        root = self._ct()
        for o in list(root.iter("{%s}Override" % CT_NS)):
            if o.get("PartName", "").lower() == pn.lower():
                root.remove(o)

    # ---- relationships ---------------------------------------------------
    @staticmethod
    def rels_path(part: str) -> str:
        d, f = posixpath.split(part.lstrip("/"))
        return posixpath.join(d, "_rels", f + ".rels")

    def resolve(self, source_part: str, target: str) -> str:
        if target.startswith("/"):
            return posixpath.normpath(target.lstrip("/"))
        base = posixpath.dirname(source_part.lstrip("/"))
        return posixpath.normpath(posixpath.join(base, target))

    @staticmethod
    def relative(source_part: str, target_part: str) -> str:
        base = posixpath.dirname(source_part.lstrip("/")) or "."
        return posixpath.relpath(target_part.lstrip("/"), base)

    def rels(self, part: str) -> list[Rel]:
        rp = self.rels_path(part)
        if not self.has(rp):
            return []
        out = []
        for r in self.peek(rp).iter("{%s}Relationship" % PKG_REL_NS):
            out.append(Rel(r.get("Id"), r.get("Type"), r.get("Target"),
                           r.get("TargetMode") == "External"))
        return out

    def rel_targets(self, part: str, rtype: str) -> list[tuple[str, str]]:
        """(rId, resolved part) for every internal relationship of ``rtype``."""
        return [(r.id, self.resolve(part, r.target)) for r in self.rels(part)
                if r.type == rtype and not r.external]

    def _rels_root(self, part: str):
        rp = self.rels_path(part)
        if not self.has(rp):
            self.write(rp, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                           b'<Relationships xmlns="%s"/>' % PKG_REL_NS.encode())
        return self.xml(rp)

    def add_rel(self, part: str, rtype: str, target_part: str) -> str:
        root = self._rels_root(part)
        ids = {r.get("Id") for r in root}
        n = 1
        while "rId%d" % n in ids:
            n += 1
        rid = "rId%d" % n
        el = etree.SubElement(root, "{%s}Relationship" % PKG_REL_NS)
        el.set("Id", rid)
        el.set("Type", rtype)
        el.set("Target", self.relative(part, target_part))
        return rid

    def remove_rel(self, part: str, rid: str) -> None:
        rp = self.rels_path(part)
        if not self.has(rp):
            return
        root = self.xml(rp)
        for r in list(root):
            if r.get("Id") == rid:
                root.remove(r)

    def next_part(self, pattern: str) -> str:
        """First free name for e.g. 'xl/pivotTables/pivotTable{}.xml'."""
        n = 1
        while self.has(pattern.format(n)):
            n += 1
        return pattern.format(n)

    # ---- workbook --------------------------------------------------------
    def _find_workbook(self) -> str:
        if not self.has("_rels/.rels"):
            raise SwivelError("not an .xlsx file (_rels/.rels is missing)")
        for r in self.rels(""):
            if r.type == RT_OFFICE_DOC:
                p = self.resolve("", r.target)
                if self.has(p):
                    return self.name(p)
        raise SwivelError("the package has no workbook part")

    @property
    def workbook(self):
        return self.xml(self.workbook_part)

    @property
    def workbook_ro(self):
        return self.peek(self.workbook_part)

    def sheets(self) -> list[SheetInfo]:
        rels = {r.id: r for r in self.rels(self.workbook_part)}
        out = []
        sheets = self.workbook_ro.find(_M + "sheets")
        if sheets is None:
            return out
        for s in sheets.findall(_M + "sheet"):
            rid = s.get(_R + "id")
            rel = rels.get(rid)
            part = self.resolve(self.workbook_part, rel.target) if rel else ""
            out.append(SheetInfo(s.get("name"), int(s.get("sheetId")), rid,
                                 self.name(part) or part, rel.type if rel else ""))
        return out

    def sheet(self, name: str) -> SheetInfo | None:
        for s in self.sheets():
            if s.name.lower() == name.lower():
                return s
        return None

    def _insert_ordered(self, parent, tag: str, order: list[str]):
        idx = order.index(tag)
        el = etree.Element(_M + tag)
        for i, child in enumerate(parent):
            if not isinstance(child.tag, str) or not child.tag.startswith(_M):
                continue
            local = child.tag[len(_M):]
            if local in order and order.index(local) > idx:
                parent.insert(i, el)
                return el
        parent.append(el)
        return el

    def _child(self, parent, tag: str, order: list[str]):
        el = parent.find(_M + tag)
        if el is None:
            el = self._insert_ordered(parent, tag, order)
        return el

    def add_sheet(self, name: str) -> SheetInfo:
        part = self.next_part("xl/worksheets/sheet{}.xml")
        self.write(part, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                         b'<worksheet xmlns="%s"><sheetData/></worksheet>' % MAIN_NS.encode())
        self.add_override(part, CT_WORKSHEET)
        rid = self.add_rel(self.workbook_part, RT_WORKSHEET, part)
        sheets = self._child(self.workbook, "sheets", _WORKBOOK_ORDER)
        ids = [int(s.get("sheetId")) for s in sheets.findall(_M + "sheet")]
        sid = max(ids, default=0) + 1
        el = etree.SubElement(sheets, _M + "sheet")
        el.set("name", name)
        el.set("sheetId", str(sid))
        el.set(_R + "id", rid)
        self._update_app_titles()
        return SheetInfo(name, sid, rid, part, RT_WORKSHEET)

    def _update_app_titles(self) -> None:
        """Keep docProps/app.xml's sheet list in step (cosmetic, best effort)."""
        app = None
        for r in self.rels(""):
            if r.type.endswith("/extended-properties"):
                app = self.resolve("", r.target)
        if not app or not self.has(app):
            return
        try:
            root = self.xml(app)
        except etree.XMLSyntaxError:
            return
        ep = root.nsmap.get(None)
        vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        titles = root.find("{%s}TitlesOfParts" % ep) if ep else None
        pairs = root.find("{%s}HeadingPairs" % ep) if ep else None
        if titles is None or pairs is None:
            return
        vec = titles.find("{%s}vector" % vt)
        pvec = pairs.find("{%s}vector" % vt)
        if vec is None or pvec is None:
            return
        variants = list(pvec)
        # HeadingPairs is (name, count) pairs; only the simple "Worksheets" +
        # optional "Named Ranges" layout is maintained.
        names = [s.name for s in self.sheets() if s.rel_type == RT_WORKSHEET]
        counts = []
        for i in range(0, len(variants) - 1, 2):
            label = variants[i].findtext("{%s}lpstr" % vt)
            cnt = variants[i + 1].find("{%s}i4" % vt)
            counts.append((label, cnt))
        if not counts or counts[0][0] != "Worksheets" or counts[0][1] is None:
            return
        old = int(counts[0][1].text)
        lpstrs = vec.findall("{%s}lpstr" % vt)
        rest = lpstrs[old:]
        for el in lpstrs:
            vec.remove(el)
        for n in names:
            e = etree.SubElement(vec, "{%s}lpstr" % vt)
            e.text = n
        for el in rest:
            vec.append(el)
        counts[0][1].text = str(len(names))
        vec.set("size", str(len(names) + len(rest)))

    # ---- styles ----------------------------------------------------------
    def _styles_part(self) -> str | None:
        for _, p in self.rel_targets(self.workbook_part, RT_STYLES):
            if self.has(p):
                return self.name(p)
        return None

    def numfmt_id(self, code: str | None) -> int:
        """numFmtId for a format code, registering a custom format if needed."""
        if not code or code == "General":
            return 0
        from openpyxl.styles.numbers import BUILTIN_FORMATS_REVERSE
        if code in BUILTIN_FORMATS_REVERSE:
            return BUILTIN_FORMATS_REVERSE[code]
        part = self._ensure_styles()
        root = self.xml(part)
        fmts = self._child(root, "numFmts", _STYLES_ORDER)
        used = []
        for f in fmts.findall(_M + "numFmt"):
            if f.get("formatCode") == code:
                return int(f.get("numFmtId"))
            used.append(int(f.get("numFmtId")))
        nid = max(used + [163]) + 1
        el = etree.SubElement(fmts, _M + "numFmt")
        el.set("numFmtId", str(nid))
        el.set("formatCode", code)
        fmts.set("count", str(len(fmts)))
        return nid

    def style_id(self, code: str | None) -> int:
        """cellXfs index of a plain style with this number format."""
        nid = self.numfmt_id(code)
        if nid == 0:
            return 0
        root = self.xml(self._ensure_styles())
        xfs = self._child(root, "cellXfs", _STYLES_ORDER)
        items = xfs.findall(_M + "xf")
        for i, xf in enumerate(items):
            if (xf.get("numFmtId") == str(nid) and xf.get("fontId", "0") == "0"
                    and xf.get("fillId", "0") == "0" and xf.get("borderId", "0") == "0"
                    and len(xf) == 0):
                return i
        if not items:  # cellXfs must start with the default style
            d = etree.SubElement(xfs, _M + "xf")
            for k in ("numFmtId", "fontId", "fillId", "borderId", "xfId"):
                d.set(k, "0")
            items = [d]
        el = etree.SubElement(xfs, _M + "xf")
        el.set("numFmtId", str(nid))
        el.set("fontId", "0")
        el.set("fillId", "0")
        el.set("borderId", "0")
        el.set("xfId", "0")
        el.set("applyNumberFormat", "1")
        xfs.set("count", str(len(items) + 1))
        return len(items)

    def _ensure_styles(self) -> str:
        part = self._styles_part()
        if part:
            return part
        from openpyxl.styles.stylesheet import Stylesheet  # minimal default sheet
        from openpyxl.xml.functions import tostring
        part = "xl/styles.xml"
        self.write(part, tostring(Stylesheet().to_tree()))
        self.add_override(part, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        self.add_rel(self.workbook_part, RT_STYLES, part)
        return part

    # ---- tables ----------------------------------------------------------
    def tables(self) -> list[TableInfo]:
        out = []
        for s in self.sheets():
            if s.rel_type != RT_WORKSHEET or not self.has(s.part):
                continue
            for _, tp in self.rel_targets(s.part, RT_TABLE):
                if not self.has(tp):
                    continue
                t = self.peek(tp)
                out.append(TableInfo(s.name, t.get("name") or "", t.get("displayName") or "",
                                     t.get("ref"), int(t.get("headerRowCount", "1")),
                                     int(t.get("totalsRowCount", "0"))))
        return out

    # ---- pivots ----------------------------------------------------------
    def pivot_caches(self) -> dict[int, str]:
        rels = {r.id: r for r in self.rels(self.workbook_part)}
        out = {}
        pcs = self.workbook_ro.find(_M + "pivotCaches")
        if pcs is None:
            return out
        for pc in pcs.findall(_M + "pivotCache"):
            rel = rels.get(pc.get(_R + "id"))
            if rel is not None:
                out[int(pc.get("cacheId"))] = self.name(self.resolve(self.workbook_part, rel.target)) or ""
        return out

    def pivots(self, sheet: SheetInfo | None = None) -> list[PivotInfo]:
        sheets = [sheet] if sheet else [s for s in self.sheets() if s.rel_type == RT_WORKSHEET]
        caches = self.pivot_caches()
        out = []
        for s in sheets:
            if not self.has(s.part):
                continue
            for _, pp in self.rel_targets(s.part, RT_PIVOT_TABLE):
                if not self.has(pp):
                    continue
                t = self.peek(pp)
                loc = t.find(_M + "location")
                cid = int(t.get("cacheId"))
                cpart = caches.get(cid)
                info = PivotInfo(
                    s.name, self.name(pp), t.get("name"), cid,
                    loc.get("ref") if loc is not None else "",
                    int(loc.get("rowPageCount", "0")) if loc is not None else 0,
                    int(loc.get("colPageCount", "0")) if loc is not None else 0,
                    cpart)
                if cpart and self.has(cpart):
                    c = self.peek(cpart)
                    info.refresh_on_load = c.get("refreshOnLoad") in ("1", "true")
                    ws = c.find("%scacheSource/%sworksheetSource" % (_M, _M))
                    if ws is not None:
                        info.source = ws.get("name") or "%s!%s" % (ws.get("sheet"), ws.get("ref"))
                out.append(info)
        return out

    def pivot_footprints(self, sheet: SheetInfo) -> list[tuple[int, int, int, int]]:
        """Cell areas (r0, c0, r1, c1) the pivots on a sheet occupy, filters included."""
        out = []
        for p in self.pivots(sheet):
            if not p.ref:
                continue
            r0, c0, r1, c1 = parse_range(p.ref)
            if p.page_rows:
                r0 = max(0, r0 - p.page_rows - 1)
                c1 = max(c1, c0 + 3 * max(p.page_cols, 1) - 1)
            out.append((r0, c0, r1, c1))
        return out

    def sheet_value_cells(self, sheet: SheetInfo) -> Iterable[tuple[int, int]]:
        """Coordinates of every cell on a sheet that holds a value or formula."""
        from ._xml import parse_cell
        data = self.read(sheet.part)
        for _, el in etree.iterparse(io.BytesIO(data), tag=_M + "c", huge_tree=True):
            if el.find(_M + "v") is not None or el.find(_M + "is") is not None \
                    or el.find(_M + "f") is not None:
                ref = el.get("r")
                if ref:
                    yield parse_cell(ref)
            el.clear()

    def clear_sheet(self, sheet: SheetInfo, *, force: bool) -> None:
        """Remove a sheet's pivot tables and content, keeping its tab and position."""
        rels = self.rels(sheet.part)
        blockers = sorted({r.type.rsplit("/", 1)[-1] for r in rels
                           if r.type not in _DROPPABLE and not r.external})
        if blockers:
            raise SwivelError(
                "sheet %r holds content swivel will not delete (%s); choose another "
                "target sheet or remove that content first" % (sheet.name, ", ".join(blockers)))
        if not force:
            areas = self.pivot_footprints(sheet)
            for r, c in self.sheet_value_cells(sheet):
                if not any(r0 <= r <= r1 and c0 <= c <= c1 for r0, c0, r1, c1 in areas):
                    from ._xml import cell_ref
                    raise SwivelError(
                        "sheet %r has data outside its pivot tables (e.g. %s); pass "
                        "if_exists='overwrite' to replace the sheet anyway"
                        % (sheet.name, cell_ref(r, c)))
        self._check_slicers(sheet)
        for r in rels:
            if r.type == RT_PIVOT_TABLE:
                self.delete(self.resolve(sheet.part, r.target))
            if not r.external:
                if r.type != RT_PIVOT_TABLE:
                    self.delete(self.resolve(sheet.part, r.target))
            self.remove_rel(sheet.part, r.id)
        self._drop_calc_chain()
        self.gc_pivot_caches()

    def _check_slicers(self, sheet: SheetInfo) -> None:
        names = {p.name for p in self.pivots(sheet)}
        if not names:
            return
        for n in list(self.parts):
            low = n.lower()
            if low.startswith(("xl/slicercaches/", "xl/timelinecaches/")) and low.endswith(".xml"):
                root = self.peek(n)
                for el in root.iter():
                    if (isinstance(el.tag, str) and el.tag.endswith("}pivotTable")
                            and el.get("name") in names
                            and (el.get("tabId") is None or el.get("tabId") == str(sheet.sheet_id))):
                        raise SwivelError(
                            "a slicer or timeline is connected to a pivot table on sheet %r; "
                            "remove it before letting swivel replace the sheet" % sheet.name)

    def _drop_calc_chain(self) -> None:
        for rid, p in self.rel_targets(self.workbook_part, RT_CALC_CHAIN):
            self.delete(p)
            self.remove_rel(self.workbook_part, rid)

    def gc_pivot_caches(self) -> None:
        """Delete pivot caches no pivot table refers to any more."""
        referenced = set()
        for p in self.pivots():
            referenced.add(p.cache_id)
        pcs = self.workbook.find(_M + "pivotCaches")
        if pcs is None:
            return
        # caches referenced from slicers / data model (x14/x15 extLst) are kept
        keep_text = b"".join(self.parts[n] for n in self.parts
                             if n.lower().startswith(("xl/slicercaches/", "xl/timelinecaches/")))
        for pc in list(pcs.findall(_M + "pivotCache")):
            cid = int(pc.get("cacheId"))
            if cid in referenced:
                continue
            if keep_text and re.search(rb'pivotCacheId="%d"' % cid, keep_text):
                continue
            rid = pc.get(_R + "id")
            for r in self.rels(self.workbook_part):
                if r.id == rid:
                    cpart = self.resolve(self.workbook_part, r.target)
                    for _, rec in self.rel_targets(cpart, RT_PIVOT_RECORDS):
                        self.delete(rec)
                    self.delete(cpart)
                    self.remove_rel(self.workbook_part, rid)
            pcs.remove(pc)
        if len(pcs) == 0:
            pcs.getparent().remove(pcs)

    def add_pivot(self, sheet: SheetInfo, *, cache_xml, records_chunks, table_xml) -> tuple[int, str]:
        """Write a cache (+records) and a pivot table; returns (cacheId, table part).

        ``cache_xml(records_rid)`` and ``table_xml(cache_id)`` build the XML once
        ids are known. ``records_chunks`` is an iterable of str, or None.
        """
        caches = self.pivot_caches()
        cache_id = max(list(caches) + [0]) + 1
        cpart = self.next_part("xl/pivotCache/pivotCacheDefinition{}.xml")
        records_rid = None
        if records_chunks is not None:
            rpart = self.next_part("xl/pivotCache/pivotCacheRecords{}.xml")
            self.write(rpart, "".join(records_chunks).encode("utf-8"))
            self.add_override(rpart, CT_PIVOT_RECORDS)
            self.write(cpart, b"")  # reserve the name so add_rel can target it
            records_rid = self.add_rel(cpart, RT_PIVOT_RECORDS, rpart)
        self.write(cpart, cache_xml(records_rid))
        self.add_override(cpart, CT_PIVOT_CACHE)
        wb_rid = self.add_rel(self.workbook_part, RT_PIVOT_CACHE, cpart)
        pcs = self._child(self.workbook, "pivotCaches", _WORKBOOK_ORDER)
        el = etree.SubElement(pcs, _M + "pivotCache")
        el.set("cacheId", str(cache_id))
        el.set(_R + "id", wb_rid)

        tpart = self.next_part("xl/pivotTables/pivotTable{}.xml")
        self.write(tpart, table_xml(cache_id))
        self.add_override(tpart, CT_PIVOT_TABLE)
        self.add_rel(tpart, RT_PIVOT_CACHE, cpart)
        self.add_rel(sheet.part, RT_PIVOT_TABLE, tpart)
        return cache_id, tpart

    def set_refresh_on_load(self, on: bool = True) -> int:
        """Flag every pivot cache in the workbook to refresh when opened."""
        n = 0
        for _, part in self.pivot_caches().items():
            if not part or not self.has(part):
                continue
            root = self.xml(part)
            if on:
                if root.get("refreshOnLoad") not in ("1", "true"):
                    root.set("refreshOnLoad", "1")
                    n += 1
            elif root.get("refreshOnLoad") is not None:
                del root.attrib["refreshOnLoad"]
                n += 1
        return n

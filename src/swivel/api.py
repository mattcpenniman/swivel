"""Public entry points."""
from __future__ import annotations

import datetime as _dt
import os
from dataclasses import dataclass
from typing import Any, Iterable, Sequence

from ._xml import cell_ref, parse_cell, range_ref
from .model import PivotModel, build
from .package import Package, PivotInfo
from .source import SourceReader
from .spec import PivotSpec, SwivelError, coerce_specs
from .writer import (cache_definition_xml, cache_records_chunks, location,
                     pivot_table_xml, worksheet_xml)

IF_EXISTS = ("replace", "overwrite", "error")


@dataclass
class PivotResult:
    """What was written for one pivot table."""

    sheet: str
    name: str
    ref: str               # the table body, e.g. "A3:D9" (what Excel calls TableRange1)
    full_ref: str          # including report filters above the table (TableRange2)
    cache_id: int
    source: str
    records: int
    cells: dict            # {"A3": value, ...} - the cells swivel pre-rendered

    def __str__(self) -> str:
        return "%s!%s  %s  (%d records from %s)" % (self.sheet, self.ref, self.name,
                                                    self.records, self.source)


def add_pivot(path, source: str, target: str, values, rows: Sequence[str] = (),
              columns: Sequence[str] = (), filters: Any = (), *, output=None,
              if_exists: str = "replace", **options) -> PivotResult:
    """Add one refreshed pivot table to an .xlsx file. See :class:`PivotSpec`.

    >>> add_pivot("sales.xlsx", source="Data", target="Summary",
    ...           rows=["Region"], columns=["Year"], values=["Amount"])
    """
    spec = PivotSpec(source=source, target=target, values=values, rows=rows,
                     columns=columns, filters=filters, **options)
    return add_pivots(path, [spec], output=output, if_exists=if_exists)[0]


def add_pivots(path, specs: Iterable[PivotSpec | dict], *, output=None,
               if_exists: str = "replace") -> list[PivotResult]:
    """Add several pivot tables in one pass.

    Every target sheet named in ``specs`` is rebuilt from scratch, so running
    the same call again (for example after the source data changed) leaves
    exactly one up-to-date copy of each pivot table - never duplicates.

    ``if_exists`` decides what happens when a target sheet already exists:

    ``"replace"`` (default)
        rebuild it, but refuse if it holds anything other than pivot tables
        (so a typo in ``target`` can never wipe a data sheet);
    ``"overwrite"``
        rebuild it whatever it contains;
    ``"error"``
        refuse.

    The file is written atomically to ``output`` (default: in place).
    """
    if if_exists not in IF_EXISTS:
        raise SwivelError("if_exists must be one of %s" % ", ".join(IF_EXISTS))
    specs = coerce_specs(specs)
    if not specs:
        raise SwivelError("no pivot tables to add")
    path = os.fspath(path)
    with open(path, "rb") as fh:
        data = fh.read()
    pkg = Package(data)
    reader = SourceReader(data, pkg)
    try:
        results = _apply(pkg, reader, specs, if_exists)
    finally:
        reader.close()
    pkg.save(output if output is not None else path)
    return results


def _apply(pkg: Package, reader: SourceReader, specs: list[PivotSpec],
           if_exists: str) -> list[PivotResult]:
    groups: dict[str, list[PivotSpec]] = {}
    display: dict[str, str] = {}
    for s in specs:
        key = s.target.lower()
        groups.setdefault(key, []).append(s)
        display.setdefault(key, s.target)

    # read every source before anything is modified
    sources = {}
    models: dict[int, PivotModel] = {}
    for s in specs:
        k = (s.source.strip().lower(), s.header_row)
        if k not in sources:
            sources[k] = reader.resolve(s.source, s.header_row)
        src = sources[k]
        if src.sheet.lower() in groups:
            raise SwivelError(
                "sheet %r is both a pivot source and a pivot target; swivel rebuilds "
                "target sheets, so they must be separate" % src.sheet)
        models[id(s)] = build(s, src)

    now = _dt.datetime.now().replace(microsecond=0)
    results = []
    for key, group in groups.items():
        name = display[key]
        sheet = pkg.sheet(name)
        if sheet is not None:
            if not sheet.rel_type.endswith("/worksheet"):
                raise SwivelError("%r is a chart or dialog sheet, not a worksheet" % sheet.name)
            if if_exists == "error":
                raise SwivelError("sheet %r already exists" % sheet.name)
            pkg.clear_sheet(sheet, force=(if_exists == "overwrite"))
        else:
            sheet = pkg.add_sheet(name)

        cells: dict = {}
        footprints: list = []
        names: set = set()
        next_left = 0
        for n, spec in enumerate(group, 1):
            model = models[id(spec)]
            P = len(model.page_fields)
            if spec.anchor:
                try:
                    top, left = parse_cell(spec.anchor)
                except ValueError as exc:
                    raise SwivelError(str(exc)) from None
                if top - P - 1 < 0:
                    raise SwivelError(
                        "anchor %s leaves no room for %d report filter(s) above the table; "
                        "use row %d or lower" % (spec.anchor, P, P + 2))
            else:
                top, left = max(2, P + 1), next_left
            bottom = top + model.height - 1
            right = left + model.width - 1
            fp_top = top - P - 1 if P else top
            fp_right = max(right, left + 1) if P else right
            if bottom >= 1048576 or right >= 16384:
                raise SwivelError("pivot table %s does not fit on a worksheet" % (spec.name or n))
            for (a, b, c, d) in footprints:
                if not (bottom < a or fp_top > c or fp_right < b or left > d):
                    raise SwivelError("pivot tables on sheet %r overlap; give them "
                                      "separate anchors" % sheet.name)
            footprints.append((fp_top, left, bottom, fp_right))
            next_left = max(next_left, fp_right + 2)

            pname = spec.name or "PivotTable%d" % n
            if pname.lower() in names:
                raise SwivelError("two pivot tables on sheet %r are named %r" % (sheet.name, pname))
            names.add(pname.lower())

            cache_id, _ = pkg.add_pivot(
                sheet,
                cache_xml=lambda rid, m=model: cache_definition_xml(
                    m, records_rid=rid, numfmt_id=pkg.numfmt_id, refreshed=now),
                records_chunks=cache_records_chunks(model) if spec.save_records else None,
                table_xml=lambda cid, m=model, nm=pname, t=top, l=left: pivot_table_xml(
                    m, name=nm, cache_id=cid, top=t, left=l, numfmt_id=pkg.numfmt_id),
            )
            rendered = {}
            for (r, c), cell in model.cells.items():
                cells[(top + r, left + c)] = cell
                rendered[cell_ref(top + r, left + c)] = cell.value
            src = model.source
            results.append(PivotResult(
                sheet=sheet.name, name=pname, ref=location(model, top, left),
                full_ref=range_ref(fp_top, left, bottom, fp_right), cache_id=cache_id,
                source=src.table or "%s!%s" % (src.sheet, src.ref),
                records=src.record_count, cells=rendered))
        pkg.write(sheet.part, worksheet_xml(cells, pkg.style_id))
    return results


def list_pivots(path) -> list[PivotInfo]:
    """Every pivot table in a workbook, with its location, source and refresh flag."""
    return Package.from_file(path).pivots()


def set_refresh_on_open(path, on: bool = True, *, output=None) -> int:
    """Make every pivot table in a workbook refresh when Excel opens it.

    Works on pivot tables swivel did not create. Returns the number of pivot
    caches changed.
    """
    pkg = Package.from_file(path)
    n = pkg.set_refresh_on_load(on)
    if n or output is not None:
        pkg.save(output if output is not None else path)
    return n

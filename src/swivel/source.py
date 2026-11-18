"""Read pivot source data out of a workbook."""
from __future__ import annotations

import datetime as _dt
import io
import warnings

from ._values import format_general, normalize
from ._xml import parse_range, range_ref
from .model import Source
from .package import Package
from .spec import SwivelError


class SourceReader:
    """Reads ranges with openpyxl in read-only mode (cached values, not formulas)."""

    def __init__(self, data: bytes, pkg: Package):
        import openpyxl
        self._pkg = pkg
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self._wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)

    def close(self) -> None:
        self._wb.close()

    def _worksheet(self, name: str):
        for ws_name in self._wb.sheetnames:
            if ws_name.lower() == name.lower():
                ws = self._wb[ws_name]
                if not hasattr(ws, "iter_rows"):
                    raise SwivelError("source sheet %r is not a worksheet" % ws_name)
                return ws_name, ws
        raise SwivelError("source sheet %r not found; sheets are: %s"
                          % (name, ", ".join(repr(s) for s in self._wb.sheetnames)))

    def resolve(self, source: str, header_row: int = 1) -> Source:
        text = source.strip()
        if "!" in text:
            sheet, rng = text.rsplit("!", 1)
            sheet = sheet.strip()
            if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
                sheet = sheet[1:-1].replace("''", "'")
            try:
                bounds = parse_range(rng)
            except ValueError as exc:
                raise SwivelError(str(exc)) from None
            name, ws = self._worksheet(sheet)
            return self._read(name, ws, bounds, None)
        sheet_names = {n.lower(): n for n in self._wb.sheetnames}
        if text.lower() in sheet_names:
            name, ws = self._worksheet(text)
            return self._read(name, ws, self._detect(ws, header_row), None)
        for t in self._pkg.tables():
            if text.lower() in (t.display_name.lower(), t.name.lower()):
                if t.header_rows < 1:
                    raise SwivelError("table %r has no header row" % t.display_name)
                r0, c0, r1, c1 = parse_range(t.ref)
                r1 -= t.totals_rows
                name, ws = self._worksheet(t.sheet)
                return self._read(name, ws, (r0, c0, r1, c1), t.display_name or t.name)
        tables = [t.display_name for t in self._pkg.tables()]
        raise SwivelError(
            "source %r is not a sheet, a Sheet!A1:D9 range or a table; sheets: %s%s"
            % (source, ", ".join(repr(s) for s in self._wb.sheetnames),
               ("; tables: " + ", ".join(repr(t) for t in tables)) if tables else ""))

    def _detect(self, ws, header_row: int):
        """Header row -> contiguous header cells -> last row holding data."""
        r0 = header_row - 1
        ws.reset_dimensions()  # never trust a stale <dimension> element
        first = last_col = None
        for row in ws.iter_rows(min_row=header_row, max_row=header_row):
            for cell in row:
                if getattr(cell, "value", None) not in (None, ""):
                    c = cell.column - 1
                    if first is None:
                        first = last_col = c
                    elif c == last_col + 1:
                        last_col = c
                    else:
                        break
        if first is None:
            raise SwivelError("sheet %r has no header row at row %d" % (ws.title, header_row))
        last_row = r0
        for row in ws.iter_rows(min_row=header_row + 1, min_col=first + 1,
                                max_col=last_col + 1):
            for cell in row:
                if getattr(cell, "value", None) not in (None, ""):
                    last_row = cell.row - 1
                    break
        # Excel needs two rows of source; an empty table gets one blank record
        # (exactly what Excel shows for a header plus an empty row), so a
        # scheduled run on an empty extract still produces a valid workbook.
        return r0, first, max(last_row, r0 + 1), last_col

    def _read(self, sheet: str, ws, bounds, table: str | None) -> Source:
        r0, c0, r1, c1 = bounds
        if r1 <= r0:
            raise SwivelError(
                "source %s!%s has a header row but no data rows; Excel cannot build a "
                "pivot table from it" % (sheet, range_ref(r0, c0, r1, c1)))
        width = c1 - c0 + 1
        columns = [[] for _ in range(width)]
        date_formats: list = [None] * width
        headers_raw = None
        for row in ws.iter_rows(min_row=r0 + 1, max_row=r1 + 1, min_col=c0 + 1, max_col=c1 + 1):
            cells = list(row)
            if len(cells) < width:  # read-only rows can be ragged
                cells += [None] * (width - len(cells))
            if headers_raw is None:
                headers_raw = [getattr(c, "value", None) for c in cells]
                continue
            for i, c in enumerate(cells):
                if c is None:
                    v = None
                else:
                    v = normalize(getattr(c, "value", None), getattr(c, "data_type", None))
                    if isinstance(v, _dt.datetime) and date_formats[i] is None:
                        fmt = getattr(c, "number_format", None)
                        date_formats[i] = fmt if fmt and fmt != "General" else "m/d/yyyy"
                columns[i].append(v)
        # iter_rows stops early when trailing rows are absent from the file
        n = r1 - r0
        for col in columns:
            col.extend([None] * (n - len(col)))
        if headers_raw is None:
            headers_raw = [None] * width
        headers = _headers(headers_raw, sheet, r0, c0)
        return Source(sheet=sheet, ref=range_ref(r0, c0, r1, c1), table=table,
                      headers=headers, columns=columns, date_formats=date_formats)


def _headers(raw, sheet, r0, c0) -> list[str]:
    from ._xml import cell_ref
    out: list[str] = []
    seen: set[str] = set()
    for i, h in enumerate(raw):
        if h is None or (isinstance(h, str) and h.strip() == ""):
            raise SwivelError(
                "source column at %s!%s has no header; every column of a pivot source "
                "needs a name" % (sheet, cell_ref(r0, c0 + i)))
        if isinstance(h, bool):
            text = "TRUE" if h else "FALSE"
        elif isinstance(h, (int, float)):
            text = format_general(h)
        elif isinstance(h, _dt.datetime):
            text = h.strftime("%m/%d/%Y").lstrip("0").replace("/0", "/")
        else:
            text = str(h)
        base, n = text, 1
        while text.lower() in seen:  # Excel renames duplicates to Name2, Name3 ...
            n += 1
            text = "%s%d" % (base, n)
        seen.add(text.lower())
        out.append(text)
    return out

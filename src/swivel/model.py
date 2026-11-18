"""Turn source rows + a :class:`PivotSpec` into a pivot cache and a laid-out table.

The layout reproduces Excel's default ("compact") PivotTable form exactly,
so the cells swivel writes are the cells Excel itself would show after a
refresh:

* all row fields share one label column; outer items carry their subtotal on
  their own row (subtotals at top);
* column fields are nested left to right with ``<item> Total`` columns after
  every group of an outer field;
* with several value fields a ``Values`` pseudo-field is the innermost column
  field.
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Any

from ._values import (Accumulator, XLError, format_date, format_general, format_serial,
                      is_number,
                      item_key, sort_key)
from .spec import ALL, PivotSpec, SwivelError, Value

DATA = -2  # the Values pseudo-field, as OOXML numbers it


@dataclass
class Source:
    """Data read from the workbook for one pivot cache."""

    sheet: str
    ref: str                     # e.g. "A1:F7" (header row included)
    table: str | None            # Excel Table name when the source is a table
    headers: list[str]
    columns: list[list[Any]]     # columns[field][record]
    date_formats: list[str | None]

    @property
    def record_count(self) -> int:
        return len(self.columns[0]) if self.columns else 0


@dataclass
class CacheField:
    name: str
    values: list
    date_format: str | None = None
    shared: list | None = None           # labels in order of first appearance
    record_index: list | None = None     # per record: index into shared
    order: list | None = None            # shared indices in display order
    pos: dict | None = None              # shared index -> display position

    def share(self) -> None:
        index: dict = {}
        shared: list = []
        rec = []
        for v in self.values:
            k = item_key(v)
            i = index.get(k)
            if i is None:
                i = index[k] = len(shared)
                shared.append(v)
            rec.append(i)
        self.shared, self.record_index = shared, rec
        self.order = sorted(range(len(shared)), key=lambda i: sort_key(shared[i]))
        self.pos = {ci: p for p, ci in enumerate(self.order)}

    def lookup(self, wanted) -> int:
        """Shared index of a user-supplied item (text matched case-insensitively)."""
        from ._values import normalize
        k = item_key(normalize(wanted))
        for i, v in enumerate(self.shared):
            if item_key(v) == k:
                return i
        # be forgiving about "2023" vs 2023 and similar
        text = label_text(normalize(wanted)).lower()
        hits = [i for i, v in enumerate(self.shared) if label_text(v).lower() == text]
        if len(hits) == 1:
            return hits[0]
        raise SwivelError(
            "filter item %r not found in field %r (items: %s)"
            % (wanted, self.name, ", ".join(repr(label_text(v)) for v in self.shared[:20])))


def label_text(v) -> str:
    """Text Excel shows for an item when it has to be rendered as a string."""
    if v is None:
        return "(blank)"
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if is_number(v):
        return format_general(v)
    if isinstance(v, _dt.datetime):
        if v.hour or v.minute or v.second or v.microsecond:
            return "%d/%d/%d %d:%02d" % (v.month, v.day, v.year, v.hour, v.minute)
        return "%d/%d/%d" % (v.month, v.day, v.year)
    if isinstance(v, XLError):
        return v.code
    return v


@dataclass
class Cell:
    value: Any
    number_format: str | None = None
    kind: str = "data"   # "data", "label", "header"


@dataclass
class ColEntry:
    type: str            # "data" (leaf), "default" (subtotal), "grand", "all"
    path: tuple          # display positions of real column-field items
    data: int            # value-field index


@dataclass
class RowEntry:
    type: str            # "item", "grand", "all"
    path: tuple


@dataclass
class DataFieldModel:
    field: int
    value: Value
    name: str


@dataclass
class PivotModel:
    spec: PivotSpec
    source: Source
    fields: list[CacheField]
    row_fields: list[int]
    col_fields: list[int]                 # real fields only
    page_fields: list[tuple[int, Any]]    # (field, selection: None | [shared idx])
    data_fields: list[DataFieldModel]
    row_entries: list[RowEntry]
    col_entries: list[ColEntry]
    cells: dict                           # (row_off, col_off) -> Cell, relative to top-left
    first_header_row: int
    first_data_row: int
    first_data_col: int
    width: int
    height: int

    @property
    def multi_data(self) -> bool:
        return len(self.data_fields) > 1

    @property
    def col_axis(self) -> list[int]:
        return self.col_fields + ([DATA] if self.multi_data else [])


def _resolve_field(name: str, headers: list[str], role: str) -> int:
    if name in headers:
        return headers.index(name)
    low = [h.lower() for h in headers]
    if low.count(name.lower()) == 1:
        return low.index(name.lower())
    raise SwivelError(
        "%s field %r is not a column of the source; columns are: %s"
        % (role, name, ", ".join(repr(h) for h in headers)))


def build(spec: PivotSpec, source: Source) -> PivotModel:
    headers = source.headers
    fields = [CacheField(h, col, fmt) for h, col, fmt in
              zip(headers, source.columns, source.date_formats)]

    rows = [_resolve_field(n, headers, "row") for n in spec.rows]
    cols = [_resolve_field(n, headers, "column") for n in spec.columns]
    pages_raw = [(_resolve_field(n, headers, "filter"), sel) for n, sel in spec.filters]
    used = rows + cols + [f for f, _ in pages_raw]
    if len(set(used)) != len(used):
        raise SwivelError("a field can be used only once across rows, columns and filters")
    for f in used:
        fields[f].share()

    pages = []
    for f, sel in pages_raw:
        if sel is ALL:
            pages.append((f, None))
        else:
            sels = sel if isinstance(sel, list) else [sel]
            idx = []
            for s in sels:
                i = fields[f].lookup(s)
                if i not in idx:
                    idx.append(i)
            if len(idx) == len(fields[f].shared) and len(idx) > 1:
                idx = None  # everything selected is the same as (All)
            pages.append((f, idx))

    data_fields = []
    taken = {h.lower() for h in headers}
    for v in spec.values:
        f = _resolve_field(v.field, headers, "value")
        name = v.name or v.default_caption
        if v.name is not None and v.name.lower() in {h.lower() for h in headers}:
            raise SwivelError(
                "value name %r is the name of a source column; Excel rejects that "
                "(add a trailing space, e.g. %r)" % (v.name, v.name + " "))
        base, n = name, 1
        while name.lower() in taken:
            if v.name is not None:
                raise SwivelError("two value fields are both named %r" % name)
            n += 1
            name = "%s%d" % (base, n)
        taken.add(name.lower())
        data_fields.append(DataFieldModel(f, v, name))

    n = source.record_count
    keep = list(range(n))
    for f, sel in pages:
        if sel is not None:
            s = set(sel)
            ri = fields[f].record_index
            keep = [r for r in keep if ri[r] in s]

    def path_of(r, flds):
        return tuple(fields[f].pos[fields[f].record_index[r]] for f in flds)

    # ---- accumulate every (row prefix, col prefix) cell -------------------
    D = len(data_fields)
    acc: dict = {}
    row_paths = set()
    col_paths = set()
    data_cols = [fields[d.field].values for d in data_fields]
    for r in keep:
        rp = path_of(r, rows)
        cp = path_of(r, cols)
        row_paths.add(rp)
        col_paths.add(cp)
        vals = [c[r] for c in data_cols]
        for i in range(len(rp) + 1):
            rk = rp[:i]
            for j in range(len(cp) + 1):
                key = (rk, cp[:j])
                a = acc.get(key)
                if a is None:
                    a = acc[key] = [Accumulator() for _ in range(D)]
                for k in range(D):
                    a[k].add(vals[k], (rp, cp, r))

    # ---- row entries (compact form, subtotals on the group row) ----------
    row_entries: list[RowEntry] = []
    if rows:
        emitted = set()
        for p in sorted(row_paths):
            for i in range(1, len(p) + 1):
                pre = p[:i]
                if pre not in emitted:
                    emitted.add(pre)
                    row_entries.append(RowEntry("item", pre))
        if spec.row_grand_totals:
            row_entries.append(RowEntry("grand", ()))
    else:
        row_entries.append(RowEntry("all", ()))

    # ---- column entries --------------------------------------------------
    col_entries: list[ColEntry] = []
    C = len(cols)
    multi = D > 1
    dloop = range(D) if multi else [0]
    if C:
        children: dict = {}
        for p in col_paths:
            for i in range(C):
                children.setdefault(p[:i], set()).add(p[i])

        def walk(prefix):
            for ch in sorted(children.get(prefix, ())):
                p = prefix + (ch,)
                if len(p) == C:
                    for d in dloop:
                        col_entries.append(ColEntry("data", p, d))
                else:
                    walk(p)
                    for d in dloop:
                        col_entries.append(ColEntry("default", p, d))

        walk(())
        if spec.column_grand_totals:
            for d in dloop:
                col_entries.append(ColEntry("grand", (), d))
    elif multi:
        for d in range(D):
            col_entries.append(ColEntry("data", (), d))
    else:
        col_entries.append(ColEntry("all", (), 0))

    # ---- geometry --------------------------------------------------------
    K = C + (1 if multi else 0)
    R = len(rows)
    label_col = 1 if (R > 0 or (C > 0 and not multi)) else 0
    if K == 0:
        first_header_row, first_data_row = 1, 1
    elif C == 0:
        first_header_row, first_data_row = 0, 1
    else:
        first_header_row, first_data_row = 1, 1 + K
    H = first_data_row
    height = H + len(row_entries)
    width = label_col + len(col_entries)

    cells: dict = {}

    def put(r, c, value, fmt=None, kind="header"):
        if value is not None:
            cells[(r, c)] = Cell(value, fmt, kind)

    def item_label(f, p):
        cf = fields[f]
        v = cf.shared[cf.order[p]]
        if v is None:
            return "(blank)", None
        if isinstance(v, bool):
            return ("TRUE" if v else "FALSE"), None
        if isinstance(v, _dt.datetime):
            return v, cf.date_format or "m/d/yyyy"
        return v, None

    caption = [d.name for d in data_fields]

    # header block
    if C > 0:
        put(0, label_col, "Column Labels")
        if not multi and R > 0:
            put(0, 0, caption[0])
        prev = None
        for j, e in enumerate(col_entries):
            c = label_col + j
            if e.type == "data":
                xs = list(e.path) + ([e.data] if multi else [])
                r = _common_prefix(prev, xs)
                for lvl in range(r, K):
                    if lvl < C:
                        v, fmt = item_label(cols[lvl], xs[lvl])
                        put(1 + lvl, c, v, fmt)
                    else:
                        put(1 + lvl, c, caption[xs[lvl]])
                prev = xs
            elif e.type == "default":
                lvl = len(e.path) - 1
                v, fmt = item_label(cols[lvl], e.path[-1])
                date_fmt = fields[cols[lvl]].date_format
                if isinstance(v, _dt.datetime):
                    text = format_date(v, fmt or "m/d/yyyy")
                elif date_fmt and is_number(v):
                    # a field holding dates formats its numbers as dates too;
                    # numbers no date can show (negative, too large) become ""
                    text = format_serial(v, date_fmt) or ""
                else:
                    text = label_text(v) if not isinstance(v, str) else v
                put(1 + lvl, c, "%s %s" % (text, caption[e.data] if multi else "Total"))
                prev = list(e.path)
            else:  # grand
                put(1, c, ("Total %s" % caption[e.data]) if multi else "Grand Total")
                prev = None
        if R > 0:
            put(H - 1, 0, "Row Labels")
    elif multi:
        if R > 0:
            put(0, 0, "Row Labels")
        for j, e in enumerate(col_entries):
            put(0, label_col + j, caption[e.data])
    else:
        if R > 0:
            put(0, 0, "Row Labels")
            put(0, 1, caption[0])
        else:
            put(0, 0, caption[0])

    # row labels
    for i, e in enumerate(row_entries):
        r = H + i
        if e.type == "item":
            v, fmt = item_label(rows[len(e.path) - 1], e.path[-1])
            put(r, 0, v, fmt, "label")
        elif e.type == "grand":
            put(r, 0, "Grand Total", None, "label")
        elif label_col:
            put(r, 0, caption[0], None, "label")

    # values
    # Excel quirk (verified against Excel 16): when the table has both row and
    # column fields and any value field uses max, min or product, cells that
    # are totals in *both* directions (row group/grand total crossed with a
    # column subtotal/grand total) only see numbers and errors: text and
    # logicals are ignored, so such a cell can come out empty or with a
    # smaller count.
    numeric_cross = bool(R and C) and any(
        d.value.func in ("max", "min", "product") for d in data_fields)
    errors_in_count_totals = any(d.value.func not in ("count", "count_nums")
                                 for d in data_fields)
    for i, re_ in enumerate(row_entries):
        row_total = re_.type == "grand" or len(re_.path) < R
        for j, ce in enumerate(col_entries):
            a = acc.get((re_.path, ce.path))
            if a is None:
                continue
            df = data_fields[ce.data]
            col_total = ce.type in ("grand", "default")
            value = a[ce.data].result(df.value.func,
                                      numeric_only=numeric_cross and row_total and col_total)
            # Excel quirk: once any value field uses a non-count function,
            # count / count_nums return the error in *total* cells holding one
            # (leaf cells still count it).
            if errors_in_count_totals and df.value.func in ("count", "count_nums")                     and (row_total or col_total) and a[ce.data].error is not None:
                value = a[ce.data].error
            put(H + i, label_col + j, value, df.value.number_format, "data")

    # report filters above the table
    P = len(pages)
    for i, (f, sel) in enumerate(pages):
        r = -(P + 1) + i
        put(r, 0, fields[f].name)
        if sel is None:
            put(r, 1, "(All)")
        elif len(sel) == 1:
            cf = fields[f]
            v, fmt = item_label(f, cf.pos[sel[0]])
            put(r, 1, v, fmt)
        else:
            put(r, 1, "(Multiple Items)")

    return PivotModel(
        spec=spec, source=source, fields=fields, row_fields=rows, col_fields=cols,
        page_fields=pages, data_fields=data_fields, row_entries=row_entries,
        col_entries=col_entries, cells=cells, first_header_row=first_header_row,
        first_data_row=first_data_row, first_data_col=label_col, width=width,
        height=height)


def _common_prefix(a, b) -> int:
    if not a:
        return 0
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n

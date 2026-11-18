"""Declarative description of a pivot table."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping, Sequence

from ._values import FUNCTIONS, canonical_function


class SwivelError(Exception):
    """Raised for anything swivel refuses to do (bad spec, unsafe overwrite...)."""


ALL = None  # a filter showing every item


@dataclass
class Value:
    """One field in the Values area.

    ``func`` is one of ``sum, count, average, max, min, product, count_nums,
    stdev, stdevp, var, varp``. ``name`` defaults to Excel's caption
    (``"Sum of Amount"``). ``number_format`` is an Excel format code such as
    ``"#,##0.00"`` applied to the value cells.
    """

    field: str
    func: str = "sum"
    name: str | None = None
    number_format: str | None = None

    def __post_init__(self):
        if not isinstance(self.field, str) or not self.field:
            raise SwivelError("a value needs a field name")
        try:
            self.func = canonical_function(self.func)
        except ValueError as exc:
            raise SwivelError(str(exc)) from None

    @property
    def default_caption(self) -> str:
        return "%s of %s" % (FUNCTIONS[self.func][1], self.field)

    @classmethod
    def coerce(cls, v) -> "Value":
        """Accept ``Value``, ``"Field"``, ``"Field:func"``, a tuple or a dict."""
        if isinstance(v, Value):
            return v
        if isinstance(v, str):
            if ":" in v:
                fld, func = v.rsplit(":", 1)
                if func.strip().lower().replace(" ", "_") in _known_funcs():
                    return cls(fld, func.strip())
            return cls(v)
        if isinstance(v, Mapping):
            return cls(**v)
        if isinstance(v, (tuple, list)):
            return cls(*v)
        raise SwivelError("cannot interpret value spec %r" % (v,))


def _known_funcs():
    from ._values import _ALIASES
    return set(FUNCTIONS) | set(_ALIASES)


@dataclass
class PivotSpec:
    """Everything needed to build one pivot table.

    ``source``
        Where the data lives: a sheet name (``"Data"`` - the used range from
        the header row down is detected), a sheet-qualified range
        (``"Data!A1:F500"``), or the name of an Excel Table (``"Sales"``).
        Tables are the most robust choice: when rows are added, a refresh
        picks them up automatically.
    ``target``
        Sheet the pivot table is placed on. Created if missing.
    ``rows`` / ``columns``
        Field names for the row and column areas, outermost first.
    ``values``
        One or more :class:`Value` (or ``"Field"`` / ``"Field:func"``).
    ``filters``
        Report filters. A list of field names shows ``(All)``; a mapping
        ``{field: selection}`` selects one item, or several when the selection
        is a list.
    ``anchor``
        Top-left cell of the table body. Defaults to ``A3`` (lower when
        filters need the room); further pivots on the same sheet are placed
        to the right of the previous one.
    """

    source: str
    target: str
    values: Sequence[Any]
    rows: Sequence[str] = ()
    columns: Sequence[str] = ()
    filters: Any = ()
    name: str | None = None
    anchor: str | None = None
    row_grand_totals: bool = True
    column_grand_totals: bool = True
    style: str = "PivotStyleLight16"
    refresh_on_load: bool = True
    save_records: bool = True
    header_row: int = 1

    def __post_init__(self):
        if isinstance(self.rows, str):
            self.rows = [self.rows]
        if isinstance(self.columns, str):
            self.columns = [self.columns]
        if isinstance(self.values, (str, Value, Mapping)):
            self.values = [self.values]
        self.rows = list(self.rows or ())
        self.columns = list(self.columns or ())
        self.values = [Value.coerce(v) for v in (self.values or ())]
        self.filters = _coerce_filters(self.filters)
        if not self.values:
            raise SwivelError("a pivot table needs at least one value field")
        if not isinstance(self.source, str) or not self.source.strip():
            raise SwivelError("source must be a sheet name, range or table name")
        if not isinstance(self.target, str) or not self.target.strip():
            raise SwivelError("target must be a sheet name")
        _check_sheet_name(self.target)
        if self.header_row < 1:
            raise SwivelError("header_row is 1-based")
        seen = {}
        for area, names in (("rows", self.rows), ("columns", self.columns),
                            ("filters", [f for f, _ in self.filters])):
            for n in names:
                if not isinstance(n, str) or not n:
                    raise SwivelError("field names must be non-empty strings")
                k = n.lower()
                if k in seen:
                    raise SwivelError(
                        "field %r is used in both %s and %s; Excel allows a field "
                        "in only one of rows, columns and filters" % (n, seen[k], area))
                seen[k] = area

    @classmethod
    def from_dict(cls, d: Mapping[str, Any]) -> "PivotSpec":
        known = set(cls.__dataclass_fields__)
        extra = set(d) - known
        if extra:
            raise SwivelError("unknown pivot option(s): %s" % ", ".join(sorted(extra)))
        return cls(**d)


def _coerce_filters(filters) -> list[tuple[str, Any]]:
    if not filters:
        return []
    if isinstance(filters, str):
        return [(filters, ALL)]
    if isinstance(filters, Mapping):
        items = list(filters.items())
    else:
        items = []
        for f in filters:
            if isinstance(f, str):
                items.append((f, ALL))
            elif isinstance(f, (tuple, list)) and len(f) == 2:
                items.append((f[0], f[1]))
            else:
                raise SwivelError("cannot interpret filter %r" % (f,))
    out = []
    for name, sel in items:
        if isinstance(sel, (list, tuple, set, frozenset)):
            sel = list(sel)
            if not sel:
                raise SwivelError("filter %r selects nothing" % name)
        out.append((name, sel))
    return out


_BAD_SHEET_CHARS = set("[]:*?/\\")


def _check_sheet_name(name: str) -> None:
    if len(name) > 31:
        raise SwivelError("sheet name %r is longer than Excel's 31 characters" % name)
    if any(c in _BAD_SHEET_CHARS for c in name) or name.startswith("'") or name.endswith("'"):
        raise SwivelError("sheet name %r contains characters Excel does not allow" % name)
    if name.lower() == "history":
        raise SwivelError("'History' is a reserved sheet name in Excel")


def coerce_specs(specs: Iterable[Any]) -> list[PivotSpec]:
    out = []
    for s in specs:
        if isinstance(s, PivotSpec):
            out.append(s)
        elif isinstance(s, Mapping):
            out.append(PivotSpec.from_dict(s))
        else:
            raise SwivelError("expected a PivotSpec or a dict, got %r" % (s,))
    return out

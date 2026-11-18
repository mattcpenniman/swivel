"""``swivel`` command line."""
from __future__ import annotations

import argparse
import json
import sys

from . import __version__
from .api import add_pivots, list_pivots, set_refresh_on_open
from .spec import PivotSpec, SwivelError


def _split(s: str | None) -> list[str]:
    return [p.strip() for p in s.split(",") if p.strip()] if s else []


def _filters(items: list[str]) -> list:
    out = []
    for it in items or []:
        if "=" in it:
            name, sel = it.split("=", 1)
            parts = [p for p in sel.split("|")]
            out.append((name.strip(), parts if len(parts) > 1 else parts[0]))
        else:
            out.append((it.strip(), None))
    return out


def _load_config(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    if path.lower().endswith((".yaml", ".yml")):
        try:
            import yaml
        except ImportError:
            raise SwivelError("YAML configs need PyYAML (pip install pyyaml)") from None
        cfg = yaml.safe_load(text)
    else:
        cfg = json.loads(text)
    if isinstance(cfg, dict):
        cfg = cfg.get("pivots", [cfg])
    if not isinstance(cfg, list):
        raise SwivelError("config must be a pivot spec, a list of them, or {'pivots': [...]}")
    return cfg


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="swivel", description="Add refreshed pivot tables to .xlsx files.")
    p.add_argument("--version", action="version", version="swivel " + __version__)
    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("add", help="add one pivot table")
    a.add_argument("workbook")
    a.add_argument("--source", required=True, help="sheet, Sheet!A1:F99 or table name")
    a.add_argument("--target", required=True, help="sheet to put the pivot table on")
    a.add_argument("--rows", help="comma-separated row fields")
    a.add_argument("--columns", help="comma-separated column fields")
    a.add_argument("--values", required=True,
                   help="comma-separated Field or Field:func (sum, count, average, max, min, ...)")
    a.add_argument("--filter", action="append", metavar="FIELD[=ITEM[|ITEM...]]",
                   help="report filter; repeatable")
    a.add_argument("--name", help="pivot table name")
    a.add_argument("--anchor", help="top-left cell of the table, e.g. A3")
    a.add_argument("--number-format", help="number format for all value fields, e.g. '#,##0.00'")
    a.add_argument("--no-records", action="store_true", help="do not store cache records")
    a.add_argument("-o", "--output", help="write here instead of in place")
    a.add_argument("--if-exists", choices=["replace", "overwrite", "error"], default="replace")

    ap = sub.add_parser("apply", help="add pivot tables described in a JSON/YAML file")
    ap.add_argument("workbook")
    ap.add_argument("config")
    ap.add_argument("-o", "--output")
    ap.add_argument("--if-exists", choices=["replace", "overwrite", "error"], default="replace")

    ls = sub.add_parser("list", help="list the pivot tables in a workbook")
    ls.add_argument("workbook")
    ls.add_argument("--json", action="store_true")

    r = sub.add_parser("refresh-on-open", help="make every pivot table refresh when opened")
    r.add_argument("workbook")
    r.add_argument("--off", action="store_true", help="turn it off instead")
    r.add_argument("-o", "--output")

    ex = sub.add_parser("excel-refresh", help="refresh all pivot tables using installed Excel")
    ex.add_argument("workbook")
    ex.add_argument("-o", "--output")

    args = p.parse_args(argv)
    try:
        if args.cmd == "add":
            values = []
            for v in _split(args.values):
                values.append(v)
            spec = PivotSpec(source=args.source, target=args.target, values=values,
                             rows=_split(args.rows), columns=_split(args.columns),
                             filters=_filters(args.filter), name=args.name,
                             anchor=args.anchor, save_records=not args.no_records)
            if args.number_format:
                for v in spec.values:
                    v.number_format = args.number_format
            for res in add_pivots(args.workbook, [spec], output=args.output,
                                  if_exists=args.if_exists):
                print(res)
        elif args.cmd == "apply":
            for res in add_pivots(args.workbook, _load_config(args.config),
                                  output=args.output, if_exists=args.if_exists):
                print(res)
        elif args.cmd == "list":
            pivots = list_pivots(args.workbook)
            if args.json:
                print(json.dumps([{"sheet": x.sheet, "name": x.name, "ref": x.ref,
                                   "cache_id": x.cache_id, "source": x.source,
                                   "refresh_on_load": x.refresh_on_load} for x in pivots],
                                 indent=2))
            else:
                for x in pivots:
                    print("%s!%s  %s  source=%s  refresh_on_open=%s"
                          % (x.sheet, x.ref, x.name, x.source, "yes" if x.refresh_on_load else "no"))
        elif args.cmd == "refresh-on-open":
            n = set_refresh_on_open(args.workbook, not args.off, output=args.output)
            print("%d pivot cache(s) changed" % n)
        elif args.cmd == "excel-refresh":
            from .excel import available, refresh
            if not available():
                raise SwivelError("excel-refresh needs Windows, Excel and pywin32 "
                                  "(pip install swivel[excel])")
            print("%d pivot table(s) refreshed" % refresh(args.workbook, output=args.output))
    except SwivelError as exc:
        print("swivel: error: %s" % exc, file=sys.stderr)
        return 2
    except FileNotFoundError as exc:
        print("swivel: error: %s" % exc, file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())

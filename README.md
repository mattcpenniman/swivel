# swivel

Add **refreshed** pivot tables to Excel `.xlsx` files from Python — no Excel
required, nothing else in the workbook touched.

```python
import swivel

swivel.add_pivot(
    "sales.xlsx",
    source="Data",                      # a sheet, "Data!A1:F500", or an Excel Table name
    target="Summary",                   # created, or rebuilt if it already holds pivots
    rows=["Region", "Product"],
    columns=["Year"],
    values=["Amount", "Qty:count"],     # or swivel.Value("Amount", "average", number_format="#,##0.00")
    filters={"Channel": "Retail"},      # report filters: None = (All), a value, or a list
)
```

## Install

Each tagged release (`vX.Y.Z`) on GitHub carries a pure-Python wheel
(`py3-none-any`). Pin it by URL, which also works for builds that only accept
pre-built wheels, such as a Lambda build using
`pip install --platform manylinux2014_x86_64 --only-binary=:all:`:

```
swivel @ https://github.com/mattcpenniman/swivel/releases/download/v0.1.0/swivel-0.1.0-py3-none-any.whl
```

Its dependencies (`openpyxl`, `lxml`) ship manylinux wheels, so they install
under the same flags; CI checks this on every change.

To cut a release: bump `__version__` in `src/swivel/__init__.py`, merge, then
push the tag `v<version>`. The *Release* workflow tests, builds, checks the tag
matches the version, smoke-tests a Lambda-style install, and publishes.

## Why it is "consistently refreshed"

Every pivot table swivel writes is refreshed twice over:

1. **At write time.** swivel builds the pivot cache, its records, and the
   rendered cells itself, laid out exactly the way Excel lays out its default
   (compact) PivotTable. The file is correct the moment it is written — in
   Excel's Protected View, in a previewer, in openpyxl or pandas, or in
   anything else that reads cells without refreshing.
2. **At open time.** The cache is flagged `refreshOnLoad`, so Excel rebuilds
   it from the source every time the file is opened.

The two agree: the test-suite has real Excel open swivel's files, refresh
them, and compare cell-for-cell (see *Testing*).

Re-running the same call after the data changes rebuilds the target sheet in
place — same tab, same position, one pivot table, no orphaned caches — so it
is safe to put in a scheduled job.

## Lossless

swivel never round-trips the workbook through a spreadsheet library. It edits
the `.xlsx` package directly: only `[Content_Types].xml`, `workbook.xml`, its
relationships, `styles.xml`, `docProps/app.xml` and the parts swivel owns are
rewritten. Charts, images, VBA (`.xlsm`), slicers, tables, defined names,
comments and existing pivot tables are copied byte-for-byte.

## Safety

* A target sheet is only rebuilt if it holds nothing but pivot tables. A sheet
  with any other data (say a typo'd `target="Budget"`) is refused unless you
  pass `if_exists="overwrite"`; `if_exists="error"` refuses any existing sheet.
* A sheet that is both a source and a target is refused.
* Sheets with drawings, tables or comments are never cleared.
* Output is written to a temp file and atomically moved into place; on any
  error the original file is untouched.

## Several pivot tables at once

```python
swivel.add_pivots("ops.xlsx", [
    {"source": "Jobs", "target": "Dashboard", "rows": ["Tech"], "values": ["Hours"]},
    {"source": "Jobs", "target": "Dashboard", "rows": ["Site"], "columns": ["Month"],
     "values": ["Cost:sum"], "name": "CostBySite"},          # placed to the right
    {"source": "Parts", "target": "Parts Pivot", "rows": ["Bin"], "values": ["Qty:max"]},
])
```

## Command line

```
swivel add book.xlsx --source Data --target Summary --rows Region --columns Year \
                     --values Amount:sum,Qty:count --filter "Region=East|West"
swivel apply book.xlsx pivots.json          # or .yaml (pip install swivel[yaml])
swivel list book.xlsx [--json]
swivel refresh-on-open book.xlsx            # flag *existing* pivots to refresh when opened
swivel excel-refresh book.xlsx              # Windows + Excel: have Excel refresh & save
```

`pivots.json` is a spec, a list of specs, or `{"pivots": [...]}` with the same
keys as `PivotSpec`.

## Supported

| | |
|---|---|
| Layout | Excel's default compact form, any number of row and column fields, subtotals, grand totals (each can be turned off) |
| Values | `sum count average max min product count_nums stdev stdevp var varp`, several at once, custom names, number formats |
| Filters | report filters showing (All), one item, or several items |
| Sources | sheet (range auto-detected from a header row), explicit range, Excel Table (refresh picks up new rows) |
| Data | text (grouped case-insensitively, sorted like Excel), numbers, dates, booleans, errors, blanks, long text, any Unicode |
| Files | `.xlsx`, `.xlsm`, `.xltx`; not `.xlsb` / `.xls` |

Not (yet) supported: tabular/outline layouts, calculated fields, grouping
(dates by month etc.), "show values as", value filters, OLAP/data-model
sources, pivot charts.

Pivot sources are read the way Excel's pivot cache reads them: formula cells
contribute their last *calculated* value. A file produced by a script that
never ran through Excel has no calculated values; swivel's pre-rendered cells
then show blanks for those columns, and Excel fills them in on open.

## Testing

```
pip install -e .[test,excel]
pytest                    # ~170 tests; the Excel ones auto-skip without Excel
```

* `tests/golden/` — expectations produced by **Excel itself**:
  `regenerate.py` has Excel build each pivot through COM and records the cells
  it renders; `test_golden.py` checks swivel renders the same cells (runs
  anywhere).
* `test_excel.py` — differential tests: real Excel opens *swivel's* file,
  refreshes every pivot, and must show exactly what swivel wrote. Covers a
  layout grid over mixed-type data, filters, hypothesis-generated random
  data/specs, re-runs, files re-saved by Excel, tables, coexistence with
  Excel-authored pivots and charts, and 20k-row sources.
* `test_properties.py` — hypothesis against a brute-force oracle, plus
  structural invariants for arbitrary specs.
* `test_package.py` — OPC integrity, byte-for-byte preservation of untouched
  parts, idempotency, every refusal path.

Excel behaviours swivel reproduces that were discovered this way include:

* case-insensitive grouping that keeps the first-seen spelling;
* word-sort ordering (`coop` < `co-op`, accents ignored), and dates listed
  after text, logicals and errors in mixed fields;
* blank vs `0` vs `#DIV/0!` for text-only or blank values, per function;
* the exact `<sharedItems>` type flags (a wrong flag makes Excel refuse to
  open the file, e.g. text+logical is *not* "mixed"; dates before 1899-12-30);
* totals take the error of their first child item, rows before columns;
* once a value field is not a count, Count shows errors in total cells;
* with a max/min/product value field, cells that are totals in both
  directions ignore text and logicals;
* date-formatted subtotal labels, including numbers in date fields
  (`1900-01-00 0:00:00 Total`).

## License

MIT

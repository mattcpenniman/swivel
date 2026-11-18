"""Smoke-test a swivel install the way a Lambda build produces it.

Run with the install target on sys.path and nothing else, e.g.:

    pip install --platform manylinux2014_x86_64 --python-version 3.13 \
        --implementation cp --only-binary=:all: --target lambda dist/*.whl
    PYTHONPATH=lambda python -s scripts/lambda_smoke.py

It writes a workbook with openpyxl, adds a pivot with swivel, and checks the
rendered cells, so it exercises lxml's compiled extension as well as swivel.
"""
import os
import sys
import tempfile

import openpyxl

import swivel


def main() -> int:
    print("swivel", swivel.__version__, "from", os.path.dirname(swivel.__file__))
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "smoke.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Region", "Amount"])
        for region, amount in [("East", 10), ("West", 5), ("East", 7)]:
            ws.append([region, amount])
        wb.save(path)

        swivel.add_pivot(path, source="Data", target="PIVOT", rows=["Region"],
                         values=["Amount:sum"], anchor="A4")

        cells = {(r[0], r[1]) for r in openpyxl.load_workbook(path)["PIVOT"]
                 .iter_rows(values_only=True) if r[0] is not None}
        expected = {("East", 17), ("West", 5), ("Grand Total", 22)}
        missing = expected - cells
        if missing:
            print("FAIL: missing", missing, "in", cells)
            return 1
        print("pivot cells OK:", sorted(expected))
    return 0


if __name__ == "__main__":
    sys.exit(main())

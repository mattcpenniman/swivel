"""swivel must render exactly what Excel rendered for the same pivot table.

The expectations in ``golden/*.json`` were produced by Excel itself (see
``golden/regenerate.py``); no Excel is needed to run these tests.
"""
import glob
import json
import os

import pytest

import swivel
from conftest import GOLDEN, assert_cells_equal, check_package, write_book

CASES = sorted(os.path.basename(p)[:-5] for p in glob.glob(os.path.join(GOLDEN, "*.json")))


def test_goldens_exist():
    assert len(CASES) >= 20


@pytest.mark.parametrize("case", CASES)
def test_matches_excel(case, tmp_path):
    with open(os.path.join(GOLDEN, case + ".json"), encoding="utf-8") as fh:
        g = json.load(fh)
    path = write_book(tmp_path / "g.xlsx", g["data"])
    spec = dict(g["spec"])
    spec["filters"] = [tuple(f) for f in spec.get("filters", [])]
    res = swivel.add_pivot(path, source="Data", target="Pivot", anchor=g["anchor"], **spec)
    assert res.ref == g["table_range1"]
    assert res.full_ref == g["table_range2"]
    assert_cells_equal(res.cells, g["cells"], case)
    check_package(path)

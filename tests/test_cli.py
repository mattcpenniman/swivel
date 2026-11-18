import json

import openpyxl

from swivel import list_pivots
from swivel.cli import main


def test_add(sales, capsys):
    rc = main(["add", sales, "--source", "Data", "--target", "Sum", "--rows", "Product",
               "--columns", "Year", "--values", "Amount:sum,Qty:count", "--filter", "Region=East|West",
               "--number-format", "0.0", "--name", "CLI"])
    assert rc == 0
    assert "Sum!A3" in capsys.readouterr().out
    (p,) = list_pivots(sales)
    assert p.name == "CLI"
    assert openpyxl.load_workbook(sales)["Sum"]["B1"].value == "(Multiple Items)"


def test_apply_json_and_list(sales, tmp_path, capsys):
    cfg = tmp_path / "pivots.json"
    cfg.write_text(json.dumps({"pivots": [
        {"source": "Data", "target": "A", "rows": ["Region"], "values": ["Amount"]},
        {"source": "Data", "target": "B", "columns": ["Year"],
         "values": [{"field": "Qty", "func": "average", "number_format": "0.00"}],
         "filters": {"Region": "East"}},
    ]}))
    out = tmp_path / "out.xlsx"
    assert main(["apply", sales, str(cfg), "-o", str(out)]) == 0
    capsys.readouterr()
    assert main(["list", str(out), "--json"]) == 0
    listed = json.loads(capsys.readouterr().out)
    assert [x["sheet"] for x in listed] == ["A", "B"]


def test_apply_yaml(sales, tmp_path):
    import pytest
    pytest.importorskip("yaml")
    cfg = tmp_path / "p.yaml"
    cfg.write_text("- source: Data\n  target: Y\n  rows: [Region]\n  values: [Qty:max]\n")
    assert main(["apply", sales, str(cfg)]) == 0
    assert [p.sheet for p in list_pivots(sales)] == ["Y"]


def test_refresh_on_open_command(sales, capsys):
    main(["add", sales, "--source", "Data", "--target", "P", "--values", "Qty"])
    assert main(["refresh-on-open", sales, "--off"]) == 0
    assert "1 pivot cache(s) changed" in capsys.readouterr().out
    assert main(["refresh-on-open", sales]) == 0
    assert list_pivots(sales)[0].refresh_on_load


def test_errors_exit_2(sales, capsys, tmp_path):
    assert main(["add", sales, "--source", "Nope", "--target", "P", "--values", "Qty"]) == 2
    assert "swivel: error:" in capsys.readouterr().err
    assert main(["list", str(tmp_path / "missing.xlsx")]) == 2

import datetime as dt
import math
import statistics
from decimal import Decimal

import pytest
from hypothesis import given
from hypothesis import strategies as st
from openpyxl.utils.datetime import to_excel

from swivel._values import (DIV0, Accumulator, XLError, canonical_function, format_general,
                            item_key, normalize, sort_key, to_serial)
from swivel._xml import (attr, cell_ref, col_index, col_letter, esc, num, ooxml_encode,
                         parse_cell, parse_range, range_ref)


@pytest.mark.parametrize("raw,expected", [
    (None, None), ("", None), ("x", "x"), (True, True), (0, 0), (1.5, 1.5),
    (Decimal("2.5"), 2.5), (dt.date(2024, 1, 2), dt.datetime(2024, 1, 2)),
    (dt.time(12, 0), 0.5), (dt.timedelta(hours=6), 0.25), (float("nan"), XLError("#NUM!")),
])
def test_normalize(raw, expected):
    assert normalize(raw) == expected


def test_normalize_errors_and_timezones():
    assert normalize("#N/A", "e") == XLError("#N/A")
    aware = dt.datetime(2024, 1, 1, tzinfo=dt.timezone.utc)
    assert normalize(aware).tzinfo is None


@given(st.datetimes(min_value=dt.datetime(1900, 3, 1), max_value=dt.datetime(9999, 12, 31)))
def test_serial_matches_openpyxl(d):
    assert math.isclose(to_serial(d), to_excel(d), abs_tol=1e-6)


def test_serial_leap_year_bug():
    assert to_serial(dt.datetime(1900, 1, 1)) == 1
    assert to_serial(dt.datetime(1900, 2, 28)) == 59
    assert to_serial(dt.datetime(1900, 3, 1)) == 61


def test_item_key_groups_like_excel():
    assert item_key("Pear") == item_key("pear") == item_key("PEAR")
    assert item_key("pear") != item_key("pear ")
    assert item_key(1) == item_key(1.0)
    assert item_key(1) != item_key("1")
    assert item_key(True) != item_key(1)
    assert item_key(dt.datetime(2024, 1, 1)) != item_key(to_serial(dt.datetime(2024, 1, 1)))


def test_sort_order_like_excel():
    values = ["b", None, True, XLError("#N/A"), 3, "A", -1, False, "co-op", "coop",
              dt.datetime(1900, 1, 5), "Ébène", "eagle", "zeta"]
    got = sorted(values, key=sort_key)
    assert got == [-1, 3, "A", "b", "coop", "co-op", "eagle", "Ébène",
                   "zeta", False, True, XLError("#N/A"), dt.datetime(1900, 1, 5), None]


@pytest.mark.parametrize("x,s", [(1, "1"), (2.5, "2.5"), (2023.0, "2023"), (0.1 + 0.2, "0.3"),
                                 (1e20, "1E+20"), (1.5e-7, "1.5E-07"), (-3.25, "-3.25")])
def test_format_general(x, s):
    assert format_general(x) == s


def _acc(values):
    a = Accumulator()
    for v in values:
        a.add(v)
    return a


@given(st.lists(st.floats(-1e6, 1e6, allow_nan=False), min_size=2, max_size=40))
def test_accumulator_matches_statistics(xs):
    a = _acc(xs)
    assert math.isclose(a.result("sum"), math.fsum(xs), rel_tol=1e-9, abs_tol=1e-6)
    assert a.result("count") == len(xs) == a.result("count_nums")
    assert math.isclose(a.result("average"), statistics.fmean(xs), rel_tol=1e-9, abs_tol=1e-6)
    assert a.result("max") == max(xs) and a.result("min") == min(xs)
    assert math.isclose(a.result("var"), statistics.variance(xs), rel_tol=1e-6, abs_tol=1e-4)
    assert math.isclose(a.result("varp"), statistics.pvariance(xs), rel_tol=1e-6, abs_tol=1e-4)
    assert math.isclose(a.result("stdev"), statistics.stdev(xs), rel_tol=1e-6, abs_tol=1e-4)
    assert math.isclose(a.result("stdevp"), statistics.pstdev(xs), rel_tol=1e-6, abs_tol=1e-4)


def test_accumulator_blank_text_error_rules():
    # only blanks -> empty cell, whatever the function
    for f in ("sum", "count", "average", "max", "product", "stdev"):
        assert _acc([None, None]).result(f) is None
    text = _acc(["x", True, None])
    assert text.result("sum") == 0
    assert text.result("count") == 2
    assert text.result("count_nums") == 0
    assert text.result("average") == DIV0
    assert text.result("max") == 0 and text.result("min") == 0 and text.result("product") == 0
    assert text.result("stdev") == DIV0 and text.result("varp") == DIV0
    err = _acc([1, XLError("#N/A"), XLError("#REF!")])
    assert err.result("sum") == XLError("#N/A")
    assert err.result("count") == 3
    one = _acc([5])
    assert one.result("stdev") == DIV0 and one.result("stdevp") == 0
    dates = _acc([dt.datetime(2024, 1, 1), 1])
    assert dates.result("sum") == to_serial(dt.datetime(2024, 1, 1)) + 1


def test_function_names():
    assert canonical_function("Average") == "average"
    assert canonical_function("avg") == "average"
    assert canonical_function("Count Numbers") == "count_nums"
    with pytest.raises(ValueError):
        canonical_function("median")


def test_cell_refs():
    for i in (0, 25, 26, 701, 702, 16383):
        assert col_index(col_letter(i)) == i
    assert col_letter(16383) == "XFD"
    assert parse_cell("$B$3") == (2, 1)
    assert cell_ref(2, 1) == "B3"
    assert parse_range("C9:A1") == (0, 0, 8, 2)
    assert range_ref(0, 0, 0, 0) == "A1"
    for bad in ("", "A0", "1A", "XFE1", "A1048577"):
        with pytest.raises(ValueError):
            parse_cell(bad)


def test_xml_escaping():
    assert esc("a<b>&c") == "a&lt;b&gt;&amp;c"
    assert attr('x"y\nz') == "x&quot;y&#10;z"
    assert ooxml_encode("a\x01b") == "a_x0001_b"
    assert ooxml_encode("_x0041_") == "_x005F_x0041_"
    assert num(2.0) == "2" and num(0.1) == "0.1" and num(True) == "1" and num(1e300) == "1e+300"


def test_cell_writer_escapes_ooxml_sequences():
    from lxml import etree
    from swivel.writer import _cell_xml
    xml = _cell_xml(0, 0, "_x0041_ and  and  space ", 0)
    assert "_x005F_x0041_" in xml and "_x0001_" in xml and 'xml:space="preserve"' in xml
    etree.fromstring(xml.replace("<c ", '<c xmlns="urn:x" ', 1))


@pytest.mark.parametrize("values,expected", [
    # straight from <sharedItems> in files Excel 16 saved
    (["a", False], {}),
    (["a", False, None], {"containsBlank": "1"}),
    ([1, True], {"containsMixedTypes": "1", "containsNumber": "1", "containsInteger": "1",
                 "minValue": "1", "maxValue": "1"}),
    ([True, XLError("#N/A")], {}),
    ([dt.datetime(2024, 2, 29), None], {"containsNonDate": "0", "containsDate": "1",
                                        "containsString": "0", "containsBlank": "1",
                                        "minDate": "2024-02-29T00:00:00",
                                        "maxDate": "2024-03-01T00:00:00"}),
    ([None], {"containsNonDate": "0", "containsString": "0", "containsBlank": "1"}),
    ([dt.datetime(2024, 2, 29), "a"], {"containsDate": "1", "containsMixedTypes": "1",
                                       "minDate": "2024-02-29T00:00:00",
                                       "maxDate": "2024-03-01T00:00:00"}),
    ([2023, 2024], {"containsSemiMixedTypes": "0", "containsString": "0", "containsNumber": "1",
                    "containsInteger": "1", "minValue": "2023", "maxValue": "2024"}),
    ([10.5, None, 1], {"containsString": "0", "containsBlank": "1", "containsNumber": "1",
                       "minValue": "1", "maxValue": "10.5"}),
    (["x", None], {"containsBlank": "1"}),
    ([dt.datetime(2024, 2, 29), -1], {"containsSemiMixedTypes": "0", "containsDate": "1",
                                      "containsString": "0", "containsMixedTypes": "1",
                                      "minDate": "1899-12-30T00:00:00",
                                      "maxDate": "2024-03-01T00:00:00"}),
    ([dt.datetime(2024, 2, 29), 1e9], {"containsSemiMixedTypes": "0", "containsDate": "1",
                                       "containsString": "0", "containsMixedTypes": "1",
                                       "minDate": "2024-02-29T00:00:00",
                                       "maxDate": "9999-12-31T00:00:00"}),
])
def test_shared_items_flags_match_excel(values, expected):
    from swivel.writer import _shared_items_attrs
    assert dict(_shared_items_attrs(values)) == expected


@pytest.mark.parametrize("x,code,text", [
    (0, r"yyyy\-mm\-dd\ h:mm:ss", "1900-01-00 0:00:00"),   # seen in Excel 16
    (1, "m/d/yyyy", "1/1/1900"), (59, "m/d/yyyy", "2/28/1900"), (61, "m/d/yyyy", "3/1/1900"),
    (45351.5, "yyyy-mm-dd hh:mm", "2024-02-29 12:00"), (-1, "m/d/yyyy", None),
    (123456.789, r"yyyy\-mm\-dd\ h:mm:ss", "2238-01-03 18:56:10"),   # rounds, like Excel
    (123456.789, "h:mm:ss.00", "18:56:09.60"),
])
def test_format_serial(x, code, text):
    from swivel._values import format_serial
    assert format_serial(x, code) == text

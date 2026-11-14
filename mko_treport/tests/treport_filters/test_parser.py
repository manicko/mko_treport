"""Parser tests for filters DSL (D5/D6)."""

from __future__ import annotations

import pytest

from mko_treport.src.scripts.treport_filters import FilterResolutionError, parse_filters
from mko_treport.src.scripts.treport_filters.models import ListExpr, RangeExpr, ScalarExpr


def test_parse_clauses_trim_and_first_equals_and_spaces_in_field_names():
    parsed = parse_filters(" Week Number = (1-12) | Years = max ", delimeter="&", row_identity=1, data_file="report_w")
    assert [c.field for c in parsed.clauses] == ["Week Number", "Years"]
    assert isinstance(parsed.clauses[0].expression, RangeExpr)
    assert isinstance(parsed.clauses[1].expression, ScalarExpr)
    assert parsed.clauses[1].expression.value == "max"


def test_parse_list_range_scalar_disambiguation_with_delimeter_and_bounds():
    p_list = parse_filters("Years=(2024&2025)", delimeter="&")
    assert isinstance(p_list.clauses[0].expression, ListExpr)
    assert p_list.clauses[0].expression.values == ("2024", "2025")

    p_range = parse_filters("Month Number=(1-max)", delimeter="&")
    r = p_range.clauses[0].expression
    assert isinstance(r, RangeExpr)
    assert r.start.kind == "literal" and r.start.value == "1"
    assert r.end.kind == "max"

    p_range2 = parse_filters("Week Number=(min-53)", delimeter="&")
    r2 = p_range2.clauses[0].expression
    assert isinstance(r2, RangeExpr)
    assert r2.start.kind == "min"
    assert r2.end.kind == "literal" and r2.end.value == "53"

    p_scalar = parse_filters("Years=max", delimeter="&")
    assert p_scalar.clauses[0].expression == ScalarExpr(value="max")


def test_parse_preserves_exact_after_trim_tokens():
    p = parse_filters("Status=( Open &Closed )|Years=max", delimeter="&")
    lst = p.clauses[0].expression
    assert isinstance(lst, ListExpr)
    assert lst.values == ("Open", "Closed")  # trimmed tokens


def test_d6_ignores_exact_duplicate_clause_and_assignment():
    p = parse_filters("Years=max | Years=max", delimeter="&")
    assert len(p.clauses) == 1
    p2 = parse_filters("Years=max|  Years  =  max  ", delimeter="&")
    assert len(p2.clauses) == 1


def test_d6_conflicting_duplicate_field_fails_fast_with_context():
    with pytest.raises(FilterResolutionError) as ei:
        parse_filters("Years=max|Years=(1-2)", delimeter="&", row_identity=7, data_file="df")
    err = ei.value
    assert err.code == "conflicting_duplicate_field"
    assert err.field == "Years"
    assert err.row_identity == 7
    assert err.data_file == "df"
    assert err.table == "data_files"


@pytest.mark.parametrize(
    "raw,code",
    [
        ("", "empty_filters"),
        ("   ", "empty_filters"),
        ("Years", "invalid_clause"),
        ("=max", "empty_field_name"),
        ("Years=", "empty_expression"),
        ("Years=max||Month Number=1", "empty_clause"),
        ("Month Number=(max)", "invalid_range_bound"),
        ("Month Number=(1-)", "invalid_range_bound"),
        ("Month Number=(a-b)", "invalid_range_bound"),
    ],
)
def test_parser_fail_fast_malformed_forms(raw, code):
    with pytest.raises(FilterResolutionError) as ei:
        parse_filters(raw, delimeter="&", row_identity="row-1", data_file="f.xlsx")
    assert ei.value.code == code
    assert ei.value.operation == "filter_parse"
    assert ei.value.table == "data_files"
    assert ei.value.row_identity == "row-1"
    assert ei.value.data_file == "f.xlsx"
    assert ei.value.to_dict()["code"] == code

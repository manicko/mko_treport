"""Placeholder derivation and report_sheet substitution tests (D11)."""

from __future__ import annotations

import pytest

from mko_treport.src.scripts.treport_filters import (
    FilterResolutionError,
    derive_max_placeholders,
    extract_report_sheet_tokens,
    substitute_report_sheet,
)
from mko_treport.src.scripts.treport_filters.models import PlaceholderDerivation, ResolvedField, ScalarExpr


def test_derive_max_placeholders_in_same_resolved_context():
    fields = [
        ResolvedField(field="Years", values=("2025",), expression=ScalarExpr("max"), max_value="2025"),
        ResolvedField(field="Week Number", values=("1", "2", "52"), expression=ScalarExpr("x"), max_value="52"),
        ResolvedField(field="Month Number", values=("1", "12"), expression=ScalarExpr("x"), max_value="12"),
    ]
    d = derive_max_placeholders(fields, row_identity=1, data_file="df")
    assert d.max_year == "2025"
    assert d.max_weeknum == "52"
    assert d.max_monthnum == "12"


def test_substitute_report_sheet_only_with_tokens_no_zero_padding_and_extract():
    d = PlaceholderDerivation(max_year="2025", max_weeknum="7", max_monthnum="3")
    sheet = "W<max_weeknum> Y<max_year> M<max_monthnum>"
    assert extract_report_sheet_tokens(sheet) == ["max_weeknum", "max_year", "max_monthnum"]
    assert substitute_report_sheet(sheet, placeholders=d) == "W7 Y2025 M3"


def test_fail_fast_referenced_but_underivable_placeholder():
    d = PlaceholderDerivation(max_year="2025", max_weeknum=None, max_monthnum=None)
    with pytest.raises(FilterResolutionError) as ei:
        substitute_report_sheet("Отчет неделя <max_weeknum>", placeholders=d, row_identity=9, data_file="report_w")
    err = ei.value
    assert err.code == "placeholder_underivable"
    assert err.operation == "placeholder_resolution"
    assert err.field == "report_sheet"
    assert err.value == "max_weeknum"
    assert err.row_identity == 9
    assert err.data_file == "report_w"


def test_error_to_dict_contains_context_fields():
    d = PlaceholderDerivation(max_year=None, max_weeknum=None, max_monthnum=None)
    with pytest.raises(FilterResolutionError) as ei:
        substitute_report_sheet("<max_year>", placeholders=d, row_identity="r1", data_file="df1")
    data = ei.value.to_dict()
    assert data["operation"] == "placeholder_resolution"
    assert data["table"] == "data_files"
    assert data["row_identity"] == "r1"
    assert data["data_file"] == "df1"
    assert data["field"] == "report_sheet"
    assert data["value"] == "max_year"
    assert data["code"] == "placeholder_underivable"

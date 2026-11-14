"""Resolver tests for dependency-aware D10 behavior."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Tuple

import pytest

from mko_treport.src.scripts.treport_filters import (
    FieldDomainResult,
    FilterResolutionError,
    YearsContext,
    parse_filters,
    resolve_data_file_filters,
    resolve_parsed_filters,
)


@dataclass
class FakeProvider:
    domains: Dict[Tuple[str, str | None, Tuple[str, ...] | None], FieldDomainResult]
    calls: list = field(default_factory=list)

    @staticmethod
    def key(field: str, years_context: YearsContext | None):
        yc = tuple(years_context.values) if years_context else None
        return (field, years_context.field if years_context else None, yc)

    def get_field_domain(self, *, data_file, data_sheet, field, years_context=None, row_identity=None):
        self.calls.append((field, years_context.values if years_context else None))
        k = self.key(field, years_context)
        if k in self.domains:
            return self.domains[k]
        # fallback by field name only
        k2 = (field, None, None) if years_context is None else None
        if k2 and k2 in self.domains:
            return self.domains[k2]
        return FieldDomainResult(status="missing_field", reason=f"no domain for {field}")


def test_resolves_years_first_then_dependents_within_years_context():
    parsed = parse_filters("Week Number=(1-max)|Years=max", delimeter="&", row_identity=1, data_file="df")
    provider = FakeProvider(
        domains={
            ("Years", None, None): FieldDomainResult(status="ok", values=("2024", "2025")),
            ("Week Number", "Years", ("2025",)): FieldDomainResult(status="ok", values=("1", "2", "3", "4")),
        }
    )
    resolved, years_ctx = resolve_parsed_filters(
        parsed, provider=provider, data_file="df", data_sheet="S", row_identity=1
    )
    assert [c[0] for c in provider.calls[:2]] == ["Years", "Week Number"]
    assert provider.calls[1] == ("Week Number", ("2025",))
    assert years_ctx and years_ctx.values == ("2025",)
    by = {r.field: r for r in resolved}
    assert by["Years"].values == ("2025",)
    assert by["Week Number"].values == ("1", "2", "3", "4")


def test_numeric_inclusive_ranges_and_min_max_in_context():
    parsed = parse_filters("Years=max|Month Number=(min-3)", delimeter="&")
    provider = FakeProvider(
        domains={
            ("Years", None, None): FieldDomainResult(status="ok", values=("2023", "2024")),
            ("Month Number", "Years", ("2024",)): FieldDomainResult(status="ok", values=("2", "3", "4", "5")),
        }
    )
    resolved, _ = resolve_parsed_filters(parsed, provider=provider, data_file="df", row_identity=1)
    by = {r.field: r for r in resolved}
    assert by["Years"].values == ("2024",)
    assert by["Month Number"].values == ("2", "3")
    assert by["Month Number"].max_value == "3"


def test_scalar_and_list_resolution_exact_after_trim_and_min_max_tokens():
    parsed = parse_filters("Week Number=(1&max&3)", delimeter="&")
    provider = FakeProvider(
        domains={("Week Number", None, None): FieldDomainResult(status="ok", values=("1", "2", "3", "4"))}
    )
    resolved, _ = resolve_parsed_filters(parsed, provider=provider, data_file="df")
    assert resolved[0].values == ("1", "4", "3")


@pytest.mark.parametrize(
    "domain_result,code",
    [
        (FieldDomainResult(status="missing_field"), "missing_field"),
        (FieldDomainResult(status="empty_domain"), "empty_domain"),
        (FieldDomainResult(status="unavailable"), "value_unavailable"),
    ],
)
def test_fail_fast_on_provider_status(domain_result, code):
    parsed = parse_filters("Years=max", delimeter="&")
    provider = FakeProvider(domains={("Years", None, None): domain_result})
    with pytest.raises(FilterResolutionError) as ei:
        resolve_parsed_filters(parsed, provider=provider, data_file="df", row_identity=2)
    assert ei.value.code == code
    assert ei.value.field == "Years"
    assert ei.value.operation == "filter_resolution"


def test_fail_fast_absent_scalar_value_and_invalid_range_bounds():
    parsed = parse_filters("Years=2022", delimeter="&")
    provider = FakeProvider(domains={("Years", None, None): FieldDomainResult(status="ok", values=("2023", "2024"))})
    with pytest.raises(FilterResolutionError) as ei:
        resolve_parsed_filters(parsed, provider=provider, data_file="df")
    assert ei.value.code == "value_unavailable"

    parsed_r = parse_filters("Month Number=(5-2)", delimeter="&")
    provider2 = FakeProvider(domains={("Month Number", None, None): FieldDomainResult(status="ok", values=("1", "2", "3", "4", "5"))})
    with pytest.raises(FilterResolutionError) as ei2:
        resolve_parsed_filters(parsed_r, provider=provider2, data_file="df")
    assert ei2.value.code == "invalid_range_bounds"


def test_end_to_end_row_resolution_derives_placeholders_and_substitutes_report_sheet():
    parsed = parse_filters("Month Number=(1-max)|Years=max", delimeter="&")
    provider = FakeProvider(
        domains={
            ("Years", None, None): FieldDomainResult(status="ok", values=("2024", "2025")),
            ("Month Number", "Years", ("2025",)): FieldDomainResult(status="ok", values=("1", "2", "10", "12")),
        }
    )
    result = resolve_data_file_filters(
        parsed=parsed,
        provider=provider,
        data_file="report_m",
        data_sheet="Data",
        report_sheet="Отчет <max_year> мес <max_monthnum>",
        delimeter="&",
        row_identity=3,
    )
    assert result.placeholders.max_year == "2025"
    assert result.placeholders.max_monthnum == "12"
    assert result.resolved_report_sheet == "Отчет 2025 мес 12"
    assert {r.field for r in result.resolved_fields} == {"Years", "Month Number"}

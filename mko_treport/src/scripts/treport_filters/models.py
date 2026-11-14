"""Typed models for parsed filters and resolution results (pure, immutable)."""

from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType
from typing import Any, Mapping, Tuple


@dataclass(frozen=True)
class ScalarExpr:
    value: str


@dataclass(frozen=True)
class BoundToken:
    kind: str  # "literal" | "min" | "max"
    value: str | None = None  # numeric/text token for literal bounds


@dataclass(frozen=True)
class RangeExpr:
    start: BoundToken
    end: BoundToken
    raw_inner: str


@dataclass(frozen=True)
class ListExpr:
    values: Tuple[str, ...]


FilterExpr = ScalarExpr | RangeExpr | ListExpr


@dataclass(frozen=True)
class FilterClause:
    field: str
    expression: FilterExpr
    raw_clause: str
    raw_expression: str


@dataclass(frozen=True)
class ParsedFilters:
    clauses: Tuple[FilterClause, ...]

    def by_field(self) -> Mapping[str, FilterClause]:
        return MappingProxyType({c.field: c for c in self.clauses})


@dataclass(frozen=True)
class ResolvedField:
    field: str
    values: Tuple[str, ...]
    expression: FilterExpr
    max_value: str | None = None


@dataclass(frozen=True)
class PlaceholderDerivation:
    max_year: str | None
    max_weeknum: str | None
    max_monthnum: str | None

    def as_mapping(self) -> Mapping[str, str]:
        data: dict[str, str] = {}
        if self.max_year is not None:
            data["max_year"] = self.max_year
        if self.max_weeknum is not None:
            data["max_weeknum"] = self.max_weeknum
        if self.max_monthnum is not None:
            data["max_monthnum"] = self.max_monthnum
        return MappingProxyType(data)


@dataclass(frozen=True)
class ResolvedRowResult:
    data_file: Any
    parsed: ParsedFilters
    resolved_fields: Tuple[ResolvedField, ...]
    placeholders: PlaceholderDerivation
    resolved_report_sheet: str

    @property
    def resolved_by_field(self) -> Mapping[str, ResolvedField]:
        return MappingProxyType({r.field: r for r in self.resolved_fields})

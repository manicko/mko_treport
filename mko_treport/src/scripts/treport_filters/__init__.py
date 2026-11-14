"""B5 filters DSL parsing and dependency-aware resolution core (pure)."""

from .errors import FilterResolutionError
from .models import (
    BoundToken,
    FilterClause,
    ListExpr,
    ParsedFilters,
    PlaceholderDerivation,
    RangeExpr,
    ResolvedField,
    ResolvedRowResult,
    ScalarExpr,
)
from .parser import parse_filters
from .placeholders import derive_max_placeholders, extract_report_sheet_tokens, substitute_report_sheet
from .resolver import resolve_data_file_filters, resolve_parsed_filters
from .values_provider import AvailableValuesProvider, FieldDomainResult, YearsContext

__all__ = [
    "AvailableValuesProvider",
    "BoundToken",
    "FieldDomainResult",
    "FilterClause",
    "FilterResolutionError",
    "ListExpr",
    "ParsedFilters",
    "PlaceholderDerivation",
    "RangeExpr",
    "ResolvedField",
    "ResolvedRowResult",
    "ScalarExpr",
    "YearsContext",
    "derive_max_placeholders",
    "extract_report_sheet_tokens",
    "parse_filters",
    "resolve_data_file_filters",
    "resolve_parsed_filters",
    "substitute_report_sheet",
]

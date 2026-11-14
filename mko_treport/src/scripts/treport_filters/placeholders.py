"""D11 placeholder derivation and report_sheet substitution."""

from __future__ import annotations

import re
from typing import Any, Iterable, Mapping

from .errors import FilterResolutionError
from .models import PlaceholderDerivation, ResolvedField

OP = "placeholder_resolution"
_TOKEN_RE = re.compile(r"<([^<>]+)>")


def _err(message: str, *, code: str, row_identity=None, data_file=None, field=None, value=None, clause=None, expression=None):
    return FilterResolutionError(
        message=message,
        operation=OP,
        table="data_files",
        row_identity=row_identity,
        field=field,
        value=value,
        code=code,
        data_file=data_file,
        clause=clause,
        expression=expression,
    )


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def extract_report_sheet_tokens(report_sheet: Any) -> list[str]:
    text = _as_text(report_sheet)
    return [m.group(1) for m in _TOKEN_RE.finditer(text)]


def derive_max_placeholders(
    resolved_fields: Iterable[ResolvedField],
    *,
    row_identity: str | int | None = None,
    data_file: Any = None,
) -> PlaceholderDerivation:
    by_field = {rf.field: rf for rf in resolved_fields}

    def max_for(field_name: str) -> str | None:
        rf = by_field.get(field_name)
        if rf is None:
            return None
        if rf.max_value is None:
            raise _err(
                f"Cannot derive max placeholder for field '{field_name}': max value underivable",
                code="placeholder_underivable",
                row_identity=row_identity,
                data_file=data_file,
                field=field_name,
            )
        return rf.max_value

    return PlaceholderDerivation(
        max_year=max_for("Years"),
        max_weeknum=max_for("Week Number"),
        max_monthnum=max_for("Month Number"),
    )


def substitute_report_sheet(
    report_sheet: Any,
    *,
    placeholders: PlaceholderDerivation | Mapping[str, str],
    row_identity: str | int | None = None,
    data_file: Any = None,
) -> str:
    text = _as_text(report_sheet)
    mapping = placeholders.as_mapping() if isinstance(placeholders, PlaceholderDerivation) else dict(placeholders)
    tokens = extract_report_sheet_tokens(text)
    # referenced-but-underivable => fail-fast
    for token in tokens:
        if token not in mapping:
            raise _err(
                f"Referenced placeholder '<{token}>' in report_sheet is underivable in this context",
                code="placeholder_underivable",
                row_identity=row_identity,
                data_file=data_file,
                field="report_sheet",
                value=token,
            )

    def repl(match: re.Match[str]) -> str:
        token = match.group(1)
        if token not in mapping:
            raise _err(
                f"Referenced placeholder '<{token}>' in report_sheet is underivable in this context",
                code="placeholder_underivable",
                row_identity=row_identity,
                data_file=data_file,
                field="report_sheet",
                value=token,
            )
        return str(mapping[token])

    return _TOKEN_RE.sub(repl, text)

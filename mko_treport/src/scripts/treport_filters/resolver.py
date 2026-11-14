"""Dependency-aware filter resolver (D10) using AvailableValuesProvider."""

from __future__ import annotations

from typing import Any

from .errors import FilterResolutionError
from .models import BoundToken, FilterClause, ListExpr, ParsedFilters, RangeExpr, ResolvedField, ResolvedRowResult, ScalarExpr
from .parser import parse_filters
from .placeholders import derive_max_placeholders, substitute_report_sheet
from .values_provider import AvailableValuesProvider, YearsContext

OP = "filter_resolution"


def _err(message: str, *, code: str, row_identity=None, data_file=None, field=None, clause=None, expression=None, value=None):
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


def _is_numeric_text(token: str) -> bool:
    t = token.strip()
    if t.startswith(("+", "-")):
        t = t[1:]
    return t.isdigit() and t != ""


def _to_int(token: str, *, row_identity, data_file, field, clause, expression, value) -> int:
    t = token.strip()
    if not _is_numeric_text(t):
        raise _err(
            f"Value '{t}' is not numeric-comparable as required",
            code="non_numeric_value",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
            value=value if value is not None else t,
        )
    return int(t)


def _domain_for(
    provider: AvailableValuesProvider,
    *,
    data_file: Any,
    data_sheet: Any,
    field: str,
    years_context: YearsContext | None,
    row_identity,
    clause: str | None = None,
    expression: str | None = None,
) -> list[str]:
    result = provider.get_field_domain(
        data_file=data_file,
        data_sheet=data_sheet,
        field=field,
        years_context=years_context,
        row_identity=row_identity,
    )
    status = getattr(result, "status", None)
    values = list(getattr(result, "values", ()))
    reason = getattr(result, "reason", None)
    if status != "ok":
        if status == "missing_field":
            raise _err(
                reason or f"Field '{field}' is missing or unresolvable in available values",
                code="missing_field",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expression,
            )
        if status == "empty_domain":
            raise _err(
                reason or f"Available domain for field '{field}' is empty",
                code="empty_domain",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expression,
            )
        if status == "unavailable":
            raise _err(
                reason or f"Value/domain for field '{field}' is unavailable",
                code="value_unavailable",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expression,
            )
        raise _err(
            reason or f"Provider returned unknown status '{status}' for field '{field}'",
            code="provider_error",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
            value=status,
        )
    if not values:
        raise _err(
            reason or f"Available domain for field '{field}' is empty",
            code="empty_domain",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
        )
    # keep exact-after-trim string tokens
    return [str(v).strip() for v in values]


def _domain_min_max(domain: list[str], *, row_identity, data_file, field, clause, expression) -> tuple[str, str]:
    ints = [_to_int(v, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression, value=v) for v in domain]
    min_i = min(ints)
    max_i = max(ints)
    # return tokens as they appear in domain (exact-after-trim) if possible
    min_s = domain[ints.index(min_i)]
    max_s = domain[ints.index(max_i)]
    return min_s, max_s


def _resolve_bound_value(bound: BoundToken, *, domain: list[str], row_identity, data_file, field, clause, expression) -> str:
    dmin, dmax = _domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression)
    if bound.kind == "min":
        return dmin
    if bound.kind == "max":
        return dmax
    # literal: must be present in domain? presence checked by caller for scalars/lists/range inclusion semantics;
    # but D10 says value absent => fail-fast. For range bounds, literal bound must be comparable; membership is implied by inclusive filter over domain.
    if bound.value is None:
        raise _err(
            "Literal bound missing value",
            code="invalid_range_bound",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
        )
    # ensure numeric comparable
    _to_int(bound.value, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression, value=bound.value)
    return bound.value.strip()


def _select_in_range(domain: list[str], start_val: str, end_val: str, *, row_identity, data_file, field, clause, expression) -> list[str]:
    start_i = _to_int(start_val, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression, value=start_val)
    end_i = _to_int(end_val, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression, value=end_val)
    if start_i > end_i:
        raise _err(
            f"Range start {start_val} is greater than end {end_val}",
            code="invalid_range_bounds",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
            value=f"{start_val}-{end_val}",
        )
    selected: list[str] = []
    for v in domain:
        vi = _to_int(v, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expression, value=v)
        if start_i <= vi <= end_i:
            selected.append(v)
    if not selected:
        raise _err(
            f"No available values in range ({start_val}-{end_val}) for field '{field}'",
            code="value_unavailable",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
            value=f"{start_val}-{end_val}",
        )
    return selected


def _resolve_clause(
    clause: FilterClause,
    *,
    provider: AvailableValuesProvider,
    data_file: Any,
    data_sheet: Any,
    years_context: YearsContext | None,
    row_identity,
) -> ResolvedField:
    domain = _domain_for(
        provider,
        data_file=data_file,
        data_sheet=data_sheet,
        field=clause.field,
        years_context=years_context,
        row_identity=row_identity,
        clause=clause.raw_clause,
        expression=clause.raw_expression,
    )
    domain_set = set(domain)
    expr = clause.expression
    if isinstance(expr, ScalarExpr):
        val = expr.value.strip()
        # min/max scalars
        if val in {"min", "max"}:
            dmin, dmax = _domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)
            chosen = dmin if val == "min" else dmax
            return ResolvedField(field=clause.field, values=(chosen,), expression=expr, max_value=_domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)[1])
        if val not in domain_set:
            raise _err(
                f"Value '{val}' is not available for field '{clause.field}'",
                code="value_unavailable",
                row_identity=row_identity,
                data_file=data_file,
                field=clause.field,
                clause=clause.raw_clause,
                expression=clause.raw_expression,
                value=val,
            )
        max_v = _domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)[1]
        return ResolvedField(field=clause.field, values=(val,), expression=expr, max_value=max_v)
    if isinstance(expr, ListExpr):
        selected: list[str] = []
        for v in expr.values:
            vv = v.strip()
            if vv in {"min", "max"}:
                dmin, dmax = _domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)
                selected.append(dmin if vv == "min" else dmax)
                continue
            if vv not in domain_set:
                raise _err(
                    f"Value '{vv}' is not available for field '{clause.field}'",
                    code="value_unavailable",
                    row_identity=row_identity,
                    data_file=data_file,
                    field=clause.field,
                    clause=clause.raw_clause,
                    expression=clause.raw_expression,
                    value=vv,
                )
            selected.append(vv)
        max_v = _domain_min_max(domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)[1]
        return ResolvedField(field=clause.field, values=tuple(selected), expression=expr, max_value=max_v)
    if isinstance(expr, RangeExpr):
        start_v = _resolve_bound_value(expr.start, domain=domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)
        end_v = _resolve_bound_value(expr.end, domain=domain, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)
        selected = _select_in_range(domain, start_v, end_v, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression)
        # for range, max of resolved selection
        ints = [_to_int(v, row_identity=row_identity, data_file=data_file, field=clause.field, clause=clause.raw_clause, expression=clause.raw_expression, value=v) for v in selected]
        max_sel = selected[ints.index(max(ints))]
        return ResolvedField(field=clause.field, values=tuple(selected), expression=expr, max_value=max_sel)
    raise _err(
        "Unsupported filter expression type",
        code="unsupported_expression",
        row_identity=row_identity,
        data_file=data_file,
        field=clause.field,
        clause=clause.raw_clause,
        expression=clause.raw_expression,
    )


def resolve_parsed_filters(
    parsed: ParsedFilters,
    *,
    provider: AvailableValuesProvider,
    data_file: Any,
    data_sheet: Any = None,
    row_identity: str | int | None = None,
) -> tuple[list[ResolvedField], YearsContext | None]:
    by_field = {c.field: c for c in parsed.clauses}
    resolved_map: dict[str, ResolvedField] = {}
    years_context: YearsContext | None = None

    # D10: resolve Years first if present (configured field name as written)
    if "Years" in by_field:
        years_clause = by_field["Years"]
        years_resolved = _resolve_clause(
            years_clause,
            provider=provider,
            data_file=data_file,
            data_sheet=data_sheet,
            years_context=None,
            row_identity=row_identity,
        )
        resolved_map["Years"] = years_resolved
        years_context = YearsContext(field="Years", values=years_resolved.values)

    # then dependents/others, skipping Years already resolved
    for clause in parsed.clauses:
        if clause.field == "Years":
            continue
        rf = _resolve_clause(
            clause,
            provider=provider,
            data_file=data_file,
            data_sheet=data_sheet,
            years_context=years_context,
            row_identity=row_identity,
        )
        resolved_map[clause.field] = rf

    # preserve deterministic parsed order in results
    resolved = [resolved_map[c.field] for c in parsed.clauses]
    return resolved, years_context


def resolve_data_file_filters(
    *,
    raw_filters: Any = None,
    parsed: ParsedFilters | None = None,
    provider: AvailableValuesProvider,
    data_file: Any = None,
    data_sheet: Any = None,
    report_sheet: Any = None,
    delimeter: Any = "&",
    row_identity: str | int | None = None,
) -> ResolvedRowResult:
    if parsed is None:
        parsed = parse_filters(
            raw_filters,
            delimeter=delimeter,
            table="data_files",
            row_identity=row_identity,
            data_file=data_file,
        )
    resolved_fields, _years_ctx = resolve_parsed_filters(
        parsed,
        provider=provider,
        data_file=data_file,
        data_sheet=data_sheet,
        row_identity=row_identity,
    )
    placeholders = derive_max_placeholders(resolved_fields, row_identity=row_identity, data_file=data_file)
    resolved_sheet = substitute_report_sheet(
        report_sheet,
        placeholders=placeholders,
        row_identity=row_identity,
        data_file=data_file,
    )
    return ResolvedRowResult(
        data_file=data_file,
        parsed=parsed,
        resolved_fields=tuple(resolved_fields),
        placeholders=placeholders,
        resolved_report_sheet=resolved_sheet,
    )

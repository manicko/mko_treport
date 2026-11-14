"""Pure filters DSL parser (D5, D6)."""

from __future__ import annotations

from typing import Any

from .errors import FilterResolutionError
from .models import BoundToken, FilterClause, ListExpr, ParsedFilters, RangeExpr, ScalarExpr

OP = "filter_parse"


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


def _parse_bound(token: str, *, row_identity, data_file, field, clause, expression) -> BoundToken:
    t = token.strip()
    if t == "":
        raise _err(
            "Empty range bound is not allowed",
            code="invalid_range_bound",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expression,
            value=token,
        )
    if t == "min":
        return BoundToken(kind="min", value=None)
    if t == "max":
        return BoundToken(kind="max", value=None)
    if _is_numeric_text(t):
        return BoundToken(kind="literal", value=t)
    raise _err(
        f"Invalid range bound '{t}'; expected numeric literal, 'min', or 'max'",
        code="invalid_range_bound",
        row_identity=row_identity,
        data_file=data_file,
        field=field,
        clause=clause,
        expression=expression,
        value=t,
    )


def _split_range_bounds(inner: str) -> tuple[str, str] | None:
    # Find separator '-' between two non-empty bound tokens. Support negative numeric literals.
    # We try each '-' position and accept if both sides are valid bound tokens.
    candidates: list[tuple[str, str]] = []
    for i, ch in enumerate(inner):
        if ch != "-":
            continue
        left = inner[:i]
        right = inner[i + 1 :]
        if left.strip() == "" or right.strip() == "":
            continue
        # validate tokens
        def ok(tok: str) -> bool:
            t = tok.strip()
            return t in {"min", "max"} or _is_numeric_text(t)
        if ok(left) and ok(right):
            candidates.append((left, right))
    if len(candidates) == 1:
        return candidates[0]
    return None


def _disambiguate_expression(expr: str, *, delimeter: str, row_identity, data_file, field, clause):
    trimmed = expr.strip()
    if trimmed == "":
        raise _err(
            "Empty filter expression is not allowed",
            code="empty_expression",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expr,
            value=expr,
        )
    # 1) list if contains general.delimeter
    if delimeter != "" and delimeter in trimmed:
        if not (trimmed.startswith("(") and trimmed.endswith(")")):
            raise _err(
                "List expression containing delimeter must be parenthesized",
                code="invalid_list_expression",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expr,
                value=expr,
            )
        inner = trimmed[1:-1]
        if inner.strip() == "":
            raise _err(
                "List expression must contain at least one value",
                code="invalid_list_expression",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expr,
                value=expr,
            )
        parts = [p.strip() for p in inner.split(delimeter)]
        if any(p == "" for p in parts):
            raise _err(
                "List expression contains an empty value after trim",
                code="invalid_list_expression",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expr,
                value=expr,
            )
        return ListExpr(values=tuple(parts))
    # 2) range if '(start-end)' with two bound tokens
    if trimmed.startswith("(") and trimmed.endswith(")"):
        inner = trimmed[1:-1]
        bounds = _split_range_bounds(inner)
        if bounds is not None:
            start_t, end_t = bounds
            start = _parse_bound(start_t, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expr)
            end = _parse_bound(end_t, row_identity=row_identity, data_file=data_file, field=field, clause=clause, expression=expr)
            return RangeExpr(start=start, end=end, raw_inner=inner)
        # parenthesized non-list/non-valid-range forms are malformed/ambiguous for this DSL.
        if "-" in inner:
            # likely malformed range attempt (e.g. (1-), (a-b))
            # choose bound-level code when obvious, else invalid range expression.
            if inner.strip().startswith("-") or inner.strip().endswith("-"):
                raise _err(
                    "Invalid range bounds in parenthesized expression",
                    code="invalid_range_bound",
                    row_identity=row_identity,
                    data_file=data_file,
                    field=field,
                    clause=clause,
                    expression=expr,
                    value=expr,
                )
            raise _err(
                "Invalid range bounds in parenthesized expression",
                code="invalid_range_bound",
                row_identity=row_identity,
                data_file=data_file,
                field=field,
                clause=clause,
                expression=expr,
                value=expr,
            )
        raise _err(
            "Invalid parenthesized filter expression; expected list or range",
            code="invalid_range_bound",
            row_identity=row_identity,
            data_file=data_file,
            field=field,
            clause=clause,
            expression=expr,
            value=expr,
        )
    # 3) scalar
    return ScalarExpr(value=trimmed)


def _canonical_assignment_key(field: str, expr_obj) -> tuple:
    if isinstance(expr_obj, ScalarExpr):
        return (field, "scalar", expr_obj.value)
    if isinstance(expr_obj, ListExpr):
        return (field, "list", expr_obj.values)
    if isinstance(expr_obj, RangeExpr):
        return (field, "range", expr_obj.start.kind, expr_obj.start.value, expr_obj.end.kind, expr_obj.end.value)
    return (field, "unknown", repr(expr_obj))


def parse_filters(
    raw_filters: Any,
    *,
    delimeter: Any,
    table: str = "data_files",
    row_identity: str | int | None = None,
    data_file: Any = None,
) -> ParsedFilters:
    if table != "data_files":
        raise _err(f"Unsupported table '{table}' for filter parsing", code="unsupported_table", row_identity=row_identity, data_file=data_file)
    if raw_filters is None:
        raise _err("filters value is missing", code="missing_filters", row_identity=row_identity, data_file=data_file, field="filters")
    text = str(raw_filters)
    if text.strip() == "":
        raise _err("filters value is empty", code="empty_filters", row_identity=row_identity, data_file=data_file, field="filters", value=raw_filters)
    delim_text = "" if delimeter is None else str(delimeter)
    if delim_text.strip() == "":
        # FR-CFG-2/A-B4-2 requires delimeter; fail-fast rather than guessing.
        raise _err(
            "general.delimeter is required for filter parsing",
            code="missing_delimeter",
            row_identity=row_identity,
            data_file=data_file,
            field="delimeter",
            value=delimeter,
        )

    # Clause split by '|', trim around clauses.
    raw_clauses = text.split("|")
    seen_exact: set[str] = set()
    by_field: dict[str, FilterClause] = {}
    field_key: dict[str, tuple] = {}
    result: list[FilterClause] = []

    for raw_clause in raw_clauses:
        clause = raw_clause.strip()
        if clause == "":
            raise _err(
                "Empty filter clause is not allowed",
                code="empty_clause",
                row_identity=row_identity,
                data_file=data_file,
                clause=raw_clause,
                value=raw_filters,
            )
        # D6 exact duplicate clause (identical text after trim) -> ignore
        if clause in seen_exact:
            continue
        seen_exact.add(clause)

        # Split field/expression on first '=' and trim.
        if "=" not in clause:
            raise _err(
                "Filter clause must contain '=' separating field and expression",
                code="invalid_clause",
                row_identity=row_identity,
                data_file=data_file,
                clause=clause,
                value=raw_filters,
            )
        field_part, expr_part = clause.split("=", 1)
        field_name = field_part.strip()
        expr_text = expr_part.strip()
        if field_name == "":
            raise _err(
                "Filter field name is empty",
                code="empty_field_name",
                row_identity=row_identity,
                data_file=data_file,
                clause=clause,
                expression=expr_part,
                value=raw_filters,
            )
        expr_obj = _disambiguate_expression(
            expr_text,
            delimeter=delim_text,
            row_identity=row_identity,
            data_file=data_file,
            field=field_name,
            clause=clause,
        )
        key = _canonical_assignment_key(field_name, expr_obj)
        if field_name in by_field:
            if field_key[field_name] == key:
                # D6: ignore exact duplicate assignment (same normalized meaning after trim)
                continue
            raise _err(
                f"Conflicting duplicate filter assignments for field '{field_name}'",
                code="conflicting_duplicate_field",
                row_identity=row_identity,
                data_file=data_file,
                field=field_name,
                clause=clause,
                expression=expr_text,
                value=raw_filters,
            )
        fc = FilterClause(field=field_name, expression=expr_obj, raw_clause=clause, raw_expression=expr_text)
        by_field[field_name] = fc
        field_key[field_name] = key
        result.append(fc)

    if not result:
        raise _err(
            "No filter clauses parsed",
            code="empty_filters",
            row_identity=row_identity,
            data_file=data_file,
            field="filters",
            value=raw_filters,
        )
    return ParsedFilters(clauses=tuple(result))

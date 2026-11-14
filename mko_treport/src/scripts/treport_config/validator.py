"""Deterministic validation for treport settings config (FR-CFG, D12/D13/D17)."""

from __future__ import annotations

from typing import Any, Mapping, Sequence

from .errors import ConfigValidationError
from .models import DataFileRow, GeneralSettings, PlaceholderRow, ReportRow, TreportConfig
from .reader import SettingsSource
from .schema import (
    DATA_FILES_ALLOWED_FIELDS,
    DATA_FILES_REQUIRED_FIELDS,
    DATA_FILES_TABLE,
    DROPPED_PARAMETERS,
    EXPECTED_PLACEHOLDERS,
    GENERAL_ALLOWED_FIELDS,
    GENERAL_REQUIRED_FIELDS,
    GENERAL_TABLE,
    PLACEHOLDERS_ALLOWED_FIELDS,
    PLACEHOLDERS_REQUIRED_FIELDS,
    PLACEHOLDERS_TABLE,
    REPORTS_ALLOWED_FIELDS,
    REPORTS_REQUIRED_FIELDS,
    REPORTS_TABLE,
    REQUIRED_TABLES,
)

OP = "config_validation"


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _require_table(rows_by_table: Mapping[str, Sequence[Mapping[str, Any]]], table: str) -> Sequence[Mapping[str, Any]]:
    if table not in rows_by_table:
        raise ConfigValidationError(
            message=f"Missing required settings table '{table}'",
            table=table,
            code="missing_table",
        )
    return rows_by_table[table]


def _reject_dropped_params(raw: Mapping[str, Any], table: str, row_identity: str | int | None) -> None:
    present = sorted(k for k in raw.keys() if k in DROPPED_PARAMETERS)
    if present:
        raise ConfigValidationError(
            message=f"Dropped parameters must not be provided/processed: {', '.join(present)} (D17)",
            table=table,
            row_identity=row_identity,
            field=present[0],
            value=raw.get(present[0]),
            code="dropped_parameter_present",
        )


def _validate_allowed_fields(raw: Mapping[str, Any], allowed: Sequence[str], table: str, row_identity: str | int | None) -> None:
    allowed_set = set(allowed)
    # Dropped params handled separately for D17-specific messaging.
    unknown = [k for k in raw.keys() if k not in allowed_set and k not in DROPPED_PARAMETERS]
    if unknown:
        field = sorted(unknown)[0]
        raise ConfigValidationError(
            message=f"Unknown/unexpected field '{field}' in table '{table}'",
            table=table,
            row_identity=row_identity,
            field=field,
            value=raw.get(field),
            code="unknown_field",
        )


def _validate_required_non_blank(raw: Mapping[str, Any], required: Sequence[str], table: str, row_identity: str | int | None) -> None:
    for field in required:
        if field not in raw or _is_blank(raw.get(field)):
            raise ConfigValidationError(
                message=f"Missing or empty required field '{field}'",
                table=table,
                row_identity=row_identity,
                field=field,
                value=raw.get(field),
                code="required_missing_or_empty",
            )


def _validate_general(rows: Sequence[Mapping[str, Any]]) -> GeneralSettings:
    table = GENERAL_TABLE
    if len(rows) != 1:
        raise ConfigValidationError(
            message=f"Table '{table}' must contain exactly one row",
            table=table,
            row_identity=None,
            field=None,
            value=len(rows),
            code="invalid_row_count",
        )
    raw = rows[0]
    row_id = "general"
    _reject_dropped_params(raw, table, row_id)
    _validate_allowed_fields(raw, GENERAL_ALLOWED_FIELDS, table, row_id)
    _validate_required_non_blank(raw, GENERAL_REQUIRED_FIELDS, table, row_id)
    # FR-CFG-2: delimeter mandatory when multi-value splitting is required by config.
    # A-B4-2 (least-permissive, deterministic): require non-empty delimeter so config can
    # express list splitting per D5/D9 without runtime guesswork.
    if "delimeter" not in raw or _is_blank(raw.get("delimeter")):
        raise ConfigValidationError(
            message="Missing or empty required field 'delimeter' in general settings (FR-CFG-2/A-B4-2)",
            table=table,
            row_identity=row_id,
            field="delimeter",
            value=raw.get("delimeter"),
            code="required_missing_or_empty",
        )
    return GeneralSettings(
        raw=dict(raw),
        input_path=raw.get("input_path"),
        output_path=raw.get("output_path"),
        delimeter=raw.get("delimeter"),
    )


def _validate_reports(rows: Sequence[Mapping[str, Any]]) -> tuple[ReportRow, ...]:
    table = REPORTS_TABLE
    if len(rows) == 0:
        raise ConfigValidationError(
            message=f"Table '{table}' must contain at least one row",
            table=table,
            code="invalid_row_count",
            value=0,
        )
    result: list[ReportRow] = []
    seen_names: dict[str, tuple[int, Any]] = {}
    for idx, raw in enumerate(rows, start=1):
        name_val = raw.get("name")
        row_id: str | int = _as_text(name_val).strip() if not _is_blank(name_val) else idx
        _reject_dropped_params(raw, table, row_id)
        _validate_allowed_fields(raw, REPORTS_ALLOWED_FIELDS, table, row_id)
        _validate_required_non_blank(raw, REPORTS_REQUIRED_FIELDS, table, row_id)
        # Config-level collision/risk checks (A-B4-3): exact and exact-after-trim duplicates.
        norm_name = _as_text(name_val).strip()
        if norm_name in seen_names:
            prev_idx, prev_val = seen_names[norm_name]
            raise ConfigValidationError(
                message=(
                    f"Duplicate report name after trim '{norm_name}' in table '{table}' "
                    f"(A-B4-3 config-level collision risk; first at row {prev_idx})"
                ),
                table=table,
                row_identity=row_id,
                field="name",
                value=name_val,
                code="report_name_collision",
            )
        seen_names[norm_name] = (idx, name_val)
        result.append(
            ReportRow(
                index=idx,
                raw=dict(raw),
                articleLevel4Name_clean=raw.get("articleLevel4Name_clean"),
                targetAudience=raw.get("targetAudience"),
                name=name_val,
            )
        )
    return tuple(result)


def _parse_group_rows(value: Any) -> bool | None:
    if _is_blank(value):
        return None  # no-op
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and value in (0, 1):
        return bool(int(value))
    text = _as_text(value).strip().lower()
    if text in {"true", "1", "yes"}:
        return True
    if text in {"false", "0", "no"}:
        return False
    return "invalid"  # type: ignore[return-value]


def _validate_data_files(rows: Sequence[Mapping[str, Any]]) -> tuple[DataFileRow, ...]:
    table = DATA_FILES_TABLE
    if len(rows) == 0:
        raise ConfigValidationError(
            message=f"Table '{table}' must contain at least one row",
            table=table,
            code="invalid_row_count",
            value=0,
        )
    result: list[DataFileRow] = []
    for idx, raw in enumerate(rows, start=1):
        df_val = raw.get("data_file")
        row_id: str | int = _as_text(df_val).strip() if not _is_blank(df_val) else idx
        _reject_dropped_params(raw, table, row_id)
        _validate_allowed_fields(raw, DATA_FILES_ALLOWED_FIELDS, table, row_id)
        _validate_required_non_blank(raw, DATA_FILES_REQUIRED_FIELDS, table, row_id)
        # D13: group_rows blank/false => no-op valid; TRUE => fail-fast (R4).
        group_rows = raw.get("group_rows")
        parsed = _parse_group_rows(group_rows)
        if parsed is True:
            raise ConfigValidationError(
                message=(
                    "group_rows=TRUE is not allowed until an explicit deterministic rule is approved (D13/R4)"
                ),
                table=table,
                row_identity=row_id,
                field="group_rows",
                value=group_rows,
                code="group_rows_true_not_allowed",
            )
        if parsed == "invalid":
            raise ConfigValidationError(
                message="Invalid value for 'group_rows'; expected blank/false/true (D13)",
                table=table,
                row_identity=row_id,
                field="group_rows",
                value=group_rows,
                code="invalid_boolean_field",
            )
        # Optional noop fields: blank is valid (D13). Non-blank values are preserved;
        # runtime type/semantic application is out of B4 scope.
        result.append(
            DataFileRow(
                index=idx,
                raw=dict(raw),
                data_file=raw.get("data_file"),
                data_sheet=raw.get("data_sheet"),
                report_sheet=raw.get("report_sheet"),
                filters=raw.get("filters"),
                paste_row=raw.get("paste_row"),
                paste_col=raw.get("paste_col"),
                freeze_row=raw.get("freeze_row"),
                freeze_col=raw.get("freeze_col"),
                zoom=raw.get("zoom"),
                group_rows=group_rows,
            )
        )
    return tuple(result)


def _validate_placeholders(rows: Sequence[Mapping[str, Any]]) -> tuple[PlaceholderRow, ...]:
    table = PLACEHOLDERS_TABLE
    if len(rows) == 0:
        raise ConfigValidationError(
            message=f"Table '{table}' must contain at least one row",
            table=table,
            code="invalid_row_count",
            value=0,
        )
    result: list[PlaceholderRow] = []
    seen: set[str] = set()
    for idx, raw in enumerate(rows, start=1):
        ph_val = raw.get("placeholder")
        row_id: str | int = _as_text(ph_val).strip() if not _is_blank(ph_val) else idx
        _reject_dropped_params(raw, table, row_id)
        _validate_allowed_fields(raw, PLACEHOLDERS_ALLOWED_FIELDS, table, row_id)
        _validate_required_non_blank(raw, PLACEHOLDERS_REQUIRED_FIELDS, table, row_id)
        name = _as_text(ph_val).strip()
        if name in seen:
            raise ConfigValidationError(
                message=f"Duplicate placeholder '{name}' in table '{table}'",
                table=table,
                row_identity=row_id,
                field="placeholder",
                value=ph_val,
                code="duplicate_placeholder",
            )
        seen.add(name)
        result.append(
            PlaceholderRow(
                index=idx,
                raw=dict(raw),
                placeholder=ph_val,
                description=raw.get("description"),
            )
        )
    # FR-CFG-5: expected placeholders present (contract/shape/presence; derivation is B5).
    missing = [p for p in EXPECTED_PLACEHOLDERS if p not in seen]
    if missing:
        raise ConfigValidationError(
            message=f"Missing expected placeholder definitions: {', '.join(missing)} (FR-CFG-5)",
            table=table,
            row_identity=None,
            field="placeholder",
            value=missing,
            code="missing_expected_placeholders",
        )
    return tuple(result)


def validate_config_from_tables(rows_by_table: Mapping[str, Sequence[Mapping[str, Any]]]) -> TreportConfig:
    # Deterministic fail-fast in table order.
    for table in REQUIRED_TABLES:
        _require_table(rows_by_table, table)
    general = _validate_general(rows_by_table[GENERAL_TABLE])
    reports = _validate_reports(rows_by_table[REPORTS_TABLE])
    data_files = _validate_data_files(rows_by_table[DATA_FILES_TABLE])
    placeholders = _validate_placeholders(rows_by_table[PLACEHOLDERS_TABLE])
    return TreportConfig(
        general=general,
        reports=reports,
        data_files=data_files,
        placeholders=placeholders,
    )


def validate_config(source: SettingsSource) -> TreportConfig:
    rows_by_table = {name: list(source.read_table(name)) for name in REQUIRED_TABLES}
    return validate_config_from_tables(rows_by_table)

"""Unit tests for treport config validation outcomes and D12 context (B4)."""

from __future__ import annotations

import pytest

from mko_treport.src.scripts.treport_config.errors import ConfigValidationError
from mko_treport.src.scripts.treport_config.validator import validate_config_from_tables


def base_tables():
    return {
        "general": [{"input_path": "C:/in", "output_path": "C:/out", "delimeter": ";"}],
        "reports": [
            {"articleLevel4Name_clean": "A", "targetAudience": "TA", "name": "R1"}
        ],
        "data_files": [
            {
                "data_file": "df1",
                "data_sheet": "S",
                "report_sheet": "RS",
                "filters": "Years=max",
                "paste_row": "",
                "paste_col": "",
                "freeze_row": "",
                "freeze_col": "",
                "zoom": "",
                "group_rows": "",
            }
        ],
        "placeholders": [
            {"placeholder": "max_weeknum", "description": "w"},
            {"placeholder": "max_monthnum", "description": "m"},
            {"placeholder": "max_year", "description": "y"},
        ],
    }


@pytest.mark.parametrize("missing", ["general", "reports", "data_files", "placeholders"])
def test_missing_required_table_fails_fast(missing):
    tables = base_tables()
    del tables[missing]
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == missing
    assert ei.value.code == "missing_table"
    assert ei.value.operation == "config_validation"


def test_invalid_general_row_count():
    tables = base_tables()
    tables["general"].append({"input_path": "x", "output_path": "y", "delimeter": ","})
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "general"
    assert ei.value.code == "invalid_row_count"


def test_report_name_collision_after_trim_a_b4_3():
    tables = base_tables()
    tables["reports"] = [
        {"articleLevel4Name_clean": "A", "targetAudience": "TA", "name": "DOLGIT "},
        {"articleLevel4Name_clean": "B", "targetAudience": "TB", "name": "DOLGIT"},
    ]
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    err = ei.value
    assert err.table == "reports"
    assert err.field == "name"
    assert err.code == "report_name_collision"
    assert "DOLGIT" in err.message


def test_duplicate_placeholder_fails_fast():
    tables = base_tables()
    tables["placeholders"].append({"placeholder": "max_year", "description": "dup"})
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "placeholders"
    assert ei.value.field == "placeholder"
    assert ei.value.code == "duplicate_placeholder"


def test_unknown_field_fails_fast_with_context():
    tables = base_tables()
    tables["reports"][0]["unexpected"] = "x"
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "reports"
    assert ei.value.field == "unexpected"
    assert ei.value.code == "unknown_field"


def test_invalid_group_rows_value_fails_fast():
    tables = base_tables()
    tables["data_files"][0]["group_rows"] = "maybe"
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.field == "group_rows"
    assert ei.value.code == "invalid_boolean_field"


def test_error_string_contains_context():
    tables = base_tables()
    tables["data_files"][0]["filters"] = ""
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    text = str(ei.value)
    assert "operation=config_validation" in text
    assert "table=data_files" in text
    assert "field=filters" in text

"""Unit tests for treport config schema/requiredness/default-noop/D17/D13 (B4)."""

from __future__ import annotations

import pytest

from mko_treport.src.scripts.treport_config.errors import ConfigValidationError
from mko_treport.src.scripts.treport_config.validator import validate_config_from_tables


def base_tables():
    return {
        "general": [{"input_path": "C:/in", "output_path": "C:/out", "delimeter": ";"}],
        "reports": [
            {
                "articleLevel4Name_clean": "A",
                "targetAudience": "TA",
                "name": "Report A",
            }
        ],
        "data_files": [
            {
                "data_file": "file1",
                "data_sheet": "Sheet1",
                "report_sheet": "Sheet <max_year>",
                "filters": "Years = max",
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


def test_valid_minimal_config_preserves_order_and_raw():
    tables = base_tables()
    tables["data_files"].append(
        {
            "data_file": "file2",
            "data_sheet": "Sheet2",
            "report_sheet": "S2",
            "filters": "Years = max",
            "paste_row": None,
            "paste_col": None,
            "freeze_row": None,
            "freeze_col": None,
            "zoom": None,
            "group_rows": None,
        }
    )
    cfg = validate_config_from_tables(tables)
    assert cfg.general.input_path == "C:/in"
    assert [r.data_file for r in cfg.data_files] == ["file1", "file2"]
    assert cfg.data_files[0].raw["report_sheet"] == "Sheet <max_year>"
    assert cfg.raw["reports"][0]["name"] == "Report A"


@pytest.mark.parametrize(
    "field", ["input_path", "output_path", "delimeter"]
)
def test_general_required_fields_non_empty(field):
    tables = base_tables()
    tables["general"][0][field] = "  "
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    err = ei.value
    assert err.table == "general"
    assert err.field == field
    assert err.code == "required_missing_or_empty"


def test_reports_required_fields_and_non_empty():
    tables = base_tables()
    tables["reports"][0]["name"] = ""
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "reports"
    assert ei.value.field == "name"


def test_data_files_required_fields_non_empty():
    tables = base_tables()
    tables["data_files"][0]["data_sheet"] = None
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "data_files"
    assert ei.value.field == "data_sheet"


def test_placeholders_contract_requires_expected_set():
    tables = base_tables()
    tables["placeholders"] = [{"placeholder": "max_year", "description": "y"}]
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.table == "placeholders"
    assert ei.value.code == "missing_expected_placeholders"


def test_d17_dropped_params_not_required_and_rejected_if_present():
    tables = base_tables()
    # Absence is compliant (implicit). Presence must fail fast and not be required.
    tables["data_files"][0]["keep_col_hidden"] = False
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.code == "dropped_parameter_present"
    assert ei.value.field == "keep_col_hidden"
    assert ei.value.table == "data_files"

    tables = base_tables()
    tables["general"][0]["new_file_delimeter"] = ","
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    assert ei.value.code == "dropped_parameter_present"


@pytest.mark.parametrize("value", [None, "", "  ", False, "false", "0"])
def test_d13_group_rows_blank_or_false_is_noop(value):
    tables = base_tables()
    tables["data_files"][0]["group_rows"] = value
    cfg = validate_config_from_tables(tables)
    assert cfg.data_files[0].group_rows == value


@pytest.mark.parametrize("value", [True, "TRUE", "true", 1, "1"])
def test_d13_group_rows_true_fails_fast(value):
    tables = base_tables()
    tables["data_files"][0]["group_rows"] = value
    tables["data_files"][0]["data_file"] = "report_m"
    with pytest.raises(ConfigValidationError) as ei:
        validate_config_from_tables(tables)
    err = ei.value
    assert err.table == "data_files"
    assert err.row_identity == "report_m"
    assert err.field == "group_rows"
    assert err.value == value
    assert err.code == "group_rows_true_not_allowed"

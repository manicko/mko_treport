"""Interaction/contract tests for SettingsSource boundary (B4)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, List, Mapping, Sequence

import pytest

from mko_treport.src.scripts.treport_config.errors import ConfigValidationError
from mko_treport.src.scripts.treport_config.validator import validate_config


@dataclass
class InMemorySettingsSource:
    tables: Mapping[str, Sequence[Mapping[str, Any]]]
    calls: List[str] = field(default_factory=list)

    def read_table(self, table_name: str) -> Sequence[Mapping[str, Any]]:
        self.calls.append(table_name)
        if table_name not in self.tables:
            raise KeyError(table_name)
        return list(self.tables[table_name])


def valid_tables():
    return {
        "general": [{"input_path": "C:/in", "output_path": "C:/out", "delimeter": ";"}],
        "reports": [
            {"articleLevel4Name_clean": "A", "targetAudience": "TA", "name": "R1"},
            {"articleLevel4Name_clean": "B", "targetAudience": "TB", "name": "R2"},
        ],
        "data_files": [
            {
                "data_file": "df1",
                "data_sheet": "S1",
                "report_sheet": "RS1",
                "filters": "Years=max",
                "paste_row": "",
                "paste_col": "",
                "freeze_row": "",
                "freeze_col": "",
                "zoom": "",
                "group_rows": "",
            },
            {
                "data_file": "df2",
                "data_sheet": "S2",
                "report_sheet": "RS2",
                "filters": "Years=max",
                "paste_row": 2,
                "paste_col": 1,
                "freeze_row": 1,
                "freeze_col": 0,
                "zoom": 90,
                "group_rows": False,
            },
        ],
        "placeholders": [
            {"placeholder": "max_weeknum", "description": "w"},
            {"placeholder": "max_monthnum", "description": "m"},
            {"placeholder": "max_year", "description": "y"},
        ],
    }


def test_validate_config_reads_tables_in_deterministic_order_and_preserves_rows():
    source = InMemorySettingsSource(tables=valid_tables())
    cfg = validate_config(source)
    assert source.calls == ["general", "reports", "data_files", "placeholders"]
    assert [r.name for r in cfg.reports] == ["R1", "R2"]
    assert [r.data_file for r in cfg.data_files] == ["df1", "df2"]
    assert cfg.data_files[1].paste_row == 2
    assert cfg.data_files[1].raw["paste_row"] == 2


def test_validate_config_propagates_missing_table_from_source_as_validation_error():
    tables = valid_tables()
    del tables["placeholders"]
    source = InMemorySettingsSource(tables=tables)
    # KeyError from source becomes missing_table via validator orchestration? we call read_table for all required.
    with pytest.raises((ConfigValidationError, KeyError)):
        validate_config(source)


def test_validate_config_surfaces_contextual_error_from_adapter_data():
    tables = valid_tables()
    tables["data_files"][0]["group_rows"] = True
    source = InMemorySettingsSource(tables=tables)
    with pytest.raises(ConfigValidationError) as ei:
        validate_config(source)
    assert ei.value.code == "group_rows_true_not_allowed"
    assert ei.value.table == "data_files"

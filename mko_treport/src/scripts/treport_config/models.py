"""Typed config models preserving source order and raw values."""

from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType
from typing import Any, Mapping, Tuple


def _freeze_row(row: Mapping[str, Any]) -> Mapping[str, Any]:
    return MappingProxyType(dict(row))


@dataclass(frozen=True)
class GeneralSettings:
    raw: Mapping[str, Any]
    input_path: Any
    output_path: Any
    delimeter: Any


@dataclass(frozen=True)
class ReportRow:
    index: int
    raw: Mapping[str, Any]
    articleLevel4Name_clean: Any
    targetAudience: Any
    name: Any


@dataclass(frozen=True)
class DataFileRow:
    index: int
    raw: Mapping[str, Any]
    data_file: Any
    data_sheet: Any
    report_sheet: Any
    filters: Any
    paste_row: Any
    paste_col: Any
    freeze_row: Any
    freeze_col: Any
    zoom: Any
    group_rows: Any


@dataclass(frozen=True)
class PlaceholderRow:
    index: int
    raw: Mapping[str, Any]
    placeholder: Any
    description: Any


@dataclass(frozen=True)
class TreportConfig:
    general: GeneralSettings
    reports: Tuple[ReportRow, ...]
    data_files: Tuple[DataFileRow, ...]
    placeholders: Tuple[PlaceholderRow, ...]

    @property
    def raw(self) -> Mapping[str, Any]:
        return MappingProxyType(
            {
                "general": self.general.raw,
                "reports": tuple(r.raw for r in self.reports),
                "data_files": tuple(r.raw for r in self.data_files),
                "placeholders": tuple(r.raw for r in self.placeholders),
            }
        )

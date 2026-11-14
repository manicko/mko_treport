"""Typed models for file resolution, readiness and session policy (B6)."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import Tuple


RESOLUTION_PRECEDENCE: Tuple[str, ...] = (".xlsx", ".xlsm", ".xlsb")


class OperationMode(StrEnum):
    CREATE_REPORT = "CreateReport"
    REFRESH_DATA = "RefreshData"


class AccessMode(StrEnum):
    READ_ONLY = "read_only"
    WRITABLE = "writable"


@dataclass(frozen=True)
class ResolvedDataFile:
    configured_data_file: str
    base_name: str
    input_path: Path
    selected_path: Path
    selected_name: str
    selected_extension: str


@dataclass(frozen=True)
class SessionPolicy:
    mode: OperationMode
    access_mode: AccessMode
    allow_save: bool
    always_close: bool = True

"""Contextual file resolution/readiness/session errors (FR-FILE, D1/D2/D12)."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence


def _fmt_context(items: list[tuple[str, Any]]) -> str:
    parts = [f"{key}={value!r}" for key, value in items if value is not None]
    if not parts:
        return ""
    return " (" + ", ".join(parts) + ")"


@dataclass(frozen=True)
class FileResolutionError(ValueError):
    """Fail-fast D2 resolution error with deterministic context."""

    message: str
    operation: str
    configured_data_file: str
    searched_input_path: Path | str
    precedence: Sequence[str]
    cause: str
    code: str
    candidates: Sequence[str] = ()

    def __str__(self) -> str:
        return self.message + _fmt_context(
            [
                ("operation", self.operation),
                ("configured_data_file", self.configured_data_file),
                ("searched_input_path", str(self.searched_input_path)),
                ("precedence", tuple(self.precedence)),
                ("candidates", tuple(self.candidates)),
                ("cause", self.cause),
                ("code", self.code),
            ]
        )

    def to_dict(self) -> Mapping[str, Any]:
        return {
            "message": self.message,
            "operation": self.operation,
            "configured_data_file": self.configured_data_file,
            "searched_input_path": str(self.searched_input_path),
            "precedence": tuple(self.precedence),
            "candidates": tuple(self.candidates),
            "cause": self.cause,
            "code": self.code,
        }


@dataclass(frozen=True)
class OutputReadinessError(ValueError):
    """Fail-fast FR-FILE-5 output readiness error with deterministic context."""

    message: str
    operation: str
    output_path: Path | str
    stage: str
    cause: str
    code: str

    def __str__(self) -> str:
        return self.message + _fmt_context(
            [
                ("operation", self.operation),
                ("output_path", str(self.output_path)),
                ("stage", self.stage),
                ("cause", self.cause),
                ("code", self.code),
            ]
        )

    def to_dict(self) -> Mapping[str, Any]:
        return {
            "message": self.message,
            "operation": self.operation,
            "output_path": str(self.output_path),
            "stage": self.stage,
            "cause": self.cause,
            "code": self.code,
        }


@dataclass(frozen=True)
class FileAccessSessionError(ValueError):
    """Fail-fast D1/session lifecycle error with deterministic context."""

    message: str
    operation: str
    mode: str
    effective_path: Path | str | None
    stage: str
    policy_access_mode: str
    cause: str
    code: str

    def __str__(self) -> str:
        return self.message + _fmt_context(
            [
                ("operation", self.operation),
                ("mode", self.mode),
                ("effective_path", None if self.effective_path is None else str(self.effective_path)),
                ("stage", self.stage),
                ("policy_access_mode", self.policy_access_mode),
                ("cause", self.cause),
                ("code", self.code),
            ]
        )

    def to_dict(self) -> Mapping[str, Any]:
        return {
            "message": self.message,
            "operation": self.operation,
            "mode": self.mode,
            "effective_path": None if self.effective_path is None else str(self.effective_path),
            "stage": self.stage,
            "policy_access_mode": self.policy_access_mode,
            "cause": self.cause,
            "code": self.code,
        }

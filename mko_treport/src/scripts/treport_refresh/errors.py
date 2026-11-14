"""Contextual refresh orchestration errors (D12-compatible)."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from .models import RefreshExecutionOutcome, RefreshStage


def _fmt_context(items: list[tuple[str, Any]]) -> str:
    parts = [f"{key}={value!r}" for key, value in items if value is not None]
    if not parts:
        return ""
    return " (" + ", ".join(parts) + ")"


@dataclass
class RefreshOrchestrationError(ValueError):
    """Fail-fast refresh workflow error with deterministic context."""

    message: str
    operation: str
    stage: RefreshStage | str
    configured_data_file: str | None = None
    row_identity: str | int | None = None
    resolved_path: Path | str | None = None
    component: str | None = None
    completion_state: str | None = None
    allows_persist: bool | None = None
    evidence: Mapping[str, Any] | None = None
    diagnostics: Mapping[str, Any] | None = None
    cause: str | None = None
    code: str | None = None

    def __str__(self) -> str:
        return self.message + _fmt_context(
            [
                ("operation", self.operation),
                ("configured_data_file", self.configured_data_file),
                ("row_identity", self.row_identity),
                ("resolved_path", None if self.resolved_path is None else str(self.resolved_path)),
                ("stage", self.stage.value if isinstance(self.stage, RefreshStage) else self.stage),
                ("component", self.component),
                ("completion_state", self.completion_state),
                ("allows_persist", self.allows_persist),
                ("cause", self.cause),
                ("code", self.code),
            ]
        )

    def to_dict(self) -> Mapping[str, Any]:
        return {
            "message": self.message,
            "operation": self.operation,
            "configured_data_file": self.configured_data_file,
            "row_identity": self.row_identity,
            "resolved_path": None if self.resolved_path is None else str(self.resolved_path),
            "stage": self.stage.value if isinstance(self.stage, RefreshStage) else self.stage,
            "component": self.component,
            "completion_state": self.completion_state,
            "allows_persist": self.allows_persist,
            "evidence": None if self.evidence is None else dict(self.evidence),
            "diagnostics": None if self.diagnostics is None else dict(self.diagnostics),
            "cause": self.cause,
            "code": self.code,
        }

    @classmethod
    def from_outcome(
        cls,
        *,
        message: str,
        operation: str,
        stage: RefreshStage | str,
        configured_data_file: str | None,
        row_identity: str | int | None,
        resolved_path: Path | str | None,
        outcome: RefreshExecutionOutcome | None,
        component: str | None = None,
        code: str | None = None,
        cause: str | None = None,
    ) -> "RefreshOrchestrationError":
        evidence = None
        diagnostics = None
        completion_state = None
        allows_persist = None
        if outcome is not None:
            completion_state = outcome.completion_state.value
            allows_persist = outcome.allows_persist
            diagnostics = dict(outcome.diagnostics)
            evidence = outcome.evidence.to_dict() if outcome.evidence is not None else None
        return cls(
            message=message,
            operation=operation,
            stage=stage,
            configured_data_file=configured_data_file,
            row_identity=row_identity,
            resolved_path=resolved_path,
            component=component,
            completion_state=completion_state,
            allows_persist=allows_persist,
            evidence=evidence,
            diagnostics=diagnostics,
            cause=cause,
            code=code,
        )

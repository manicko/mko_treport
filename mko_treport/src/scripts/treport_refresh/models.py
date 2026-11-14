"""Typed contracts for refresh orchestration (B7, OQ-T2 seam)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any, Mapping, Optional, Tuple

from mko_treport.src.scripts.treport_files.models import OperationMode, ResolvedDataFile


class RefreshCompletionState(StrEnum):
    COMPLETED = "completed"
    NOT_COMPLETED = "not_completed"
    TIMED_OUT = "timed_out"
    FAILED = "failed"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class RefreshRequest:
    """Semantic refresh request; no mechanism-specific fields."""

    operation: OperationMode
    configured_data_file: str
    row_identity: str | int | None
    resolved_path: Path
    session_handle: Any = None
    data_sheet: Any = None
    context: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_resolved(
        cls,
        *,
        operation: OperationMode,
        resolved: ResolvedDataFile,
        row_identity: str | int | None = None,
        session_handle: Any = None,
        data_sheet: Any = None,
        context: Mapping[str, Any] | None = None,
    ) -> "RefreshRequest":
        return cls(
            operation=operation,
            configured_data_file=resolved.configured_data_file,
            row_identity=row_identity,
            resolved_path=resolved.selected_path,
            session_handle=session_handle,
            data_sheet=data_sheet,
            context=dict(context or {}),
        )


@dataclass(frozen=True)
class RefreshEvidence:
    """Extensible nullable evidence hooks; absence/ambiguity means not persistable."""

    details: Mapping[str, Any] = field(default_factory=dict)
    diagnostics: Mapping[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Mapping[str, Any]:
        return {"details": dict(self.details), "diagnostics": dict(self.diagnostics)}


@dataclass(frozen=True)
class RefreshExecutionOutcome:
    completion_state: RefreshCompletionState
    allows_persist: bool
    evidence: Optional[RefreshEvidence] = None
    diagnostics: Mapping[str, Any] = field(default_factory=dict)
    started_at: Optional[datetime] = None
    finished_at: Optional[datetime] = None
    duration_ms: Optional[float] = None
    error: Optional[BaseException] = None

    @property
    def is_proven_complete(self) -> bool:
        return (
            self.completion_state is RefreshCompletionState.COMPLETED
            and self.allows_persist is True
        )


class RefreshStage(StrEnum):
    RESOLVE_INPUT = "resolve_input"
    PRECONDITION = "precondition"
    SESSION_OPEN = "session_open"
    REFRESH_TRIGGER = "refresh_trigger"
    AWAIT_COMPLETION = "await_completion_or_refresh_completion"
    REFRESH_EXECUTION = "refresh_execution"
    PRE_SAVE_GATE = "pre_save_gate"
    SAVE = "save"
    CLOSE_CLEANUP = "close_cleanup"
    STATE_INTEGRITY = "state_integrity"


@dataclass(frozen=True)
class ProgressEvent:
    stage: RefreshStage
    operation: str
    configured_data_file: str | None = None
    row_identity: str | int | None = None
    resolved_path: Path | str | None = None
    message: str | None = None
    timestamp: Optional[datetime] = None
    details: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class RefreshBatchItem:
    """Semantic ordered batch item for one data file."""

    configured_data_file: str
    row_identity: str | int | None = None
    data_sheet: Any = None
    context: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class RefreshFileResult:
    configured_data_file: str
    row_identity: str | int | None
    resolved_path: Path | None
    operation: str
    outcome: Optional[RefreshExecutionOutcome]
    saved: bool
    closed: bool
    progress_events: Tuple[ProgressEvent, ...] = ()

    def to_dict(self) -> Mapping[str, Any]:
        out = self.outcome
        return {
            "configured_data_file": self.configured_data_file,
            "row_identity": self.row_identity,
            "resolved_path": None if self.resolved_path is None else str(self.resolved_path),
            "operation": self.operation,
            "saved": self.saved,
            "closed": self.closed,
            "completion_state": None if out is None else out.completion_state.value,
            "allows_persist": None if out is None else out.allows_persist,
            "duration_ms": None if out is None else out.duration_ms,
            "diagnostics": {} if out is None else dict(out.diagnostics),
            "evidence": None if out is None or out.evidence is None else out.evidence.to_dict(),
        }

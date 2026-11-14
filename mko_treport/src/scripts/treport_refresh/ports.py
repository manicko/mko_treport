"""Abstract ports for refresh execution, progress and time (B7)."""

from __future__ import annotations

from datetime import datetime
from typing import Optional, Protocol

from .models import ProgressEvent, RefreshExecutionOutcome, RefreshRequest


class RefreshExecutorPort(Protocol):
    """Completion-aware refresh executor (mechanism OQ-T2 gated)."""

    def execute(self, request: RefreshRequest) -> RefreshExecutionOutcome: ...


class ProgressSink(Protocol):
    def emit(self, event: ProgressEvent) -> None: ...


class Clock(Protocol):
    def now(self) -> datetime: ...


class Deadline(Protocol):
    @property
    def deadline_at(self) -> Optional[datetime]: ...

    def is_expired(self, now: Optional[datetime] = None) -> bool: ...

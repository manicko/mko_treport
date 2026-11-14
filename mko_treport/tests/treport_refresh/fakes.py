"""Test doubles for B7 treport_refresh."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, List

from mko_treport.src.scripts.treport_files.models import AccessMode
from mko_treport.src.scripts.treport_refresh.models import ProgressEvent, RefreshExecutionOutcome, RefreshRequest


@dataclass
class RecordingProgressSink:
    events: List[ProgressEvent] = field(default_factory=list)

    def emit(self, event: ProgressEvent) -> None:
        self.events.append(event)


@dataclass
class FixedClock:
    current: datetime

    def now(self) -> datetime:
        return self.current


@dataclass
class FakeRefreshExecutor:
    outcome: RefreshExecutionOutcome | Any = None
    exception: BaseException | None = None
    calls: List[RefreshRequest] = field(default_factory=list)

    def execute(self, request: RefreshRequest) -> RefreshExecutionOutcome:
        self.calls.append(request)
        if self.exception is not None:
            raise self.exception
        return self.outcome


# Reuse-compatible minimal fakes (aligned with treport_files tests patterns)
@dataclass
class OpenCall:
    path: Path
    access_mode: AccessMode


@dataclass
class FakeWorkbookSessionPort:
    fail_on_open: Exception | None = None
    fail_on_save: Exception | None = None
    fail_on_close: Exception | None = None
    opened: List[OpenCall] = field(default_factory=list)
    saved: List[Any] = field(default_factory=list)
    closed: List[Any] = field(default_factory=list)
    _counter: int = 0

    def open(self, path: Path, access_mode: AccessMode) -> Any:
        if self.fail_on_open:
            raise self.fail_on_open
        self._counter += 1
        handle = {"id": self._counter, "path": path, "mode": access_mode}
        self.opened.append(OpenCall(path=Path(path), access_mode=access_mode))
        return handle

    def save(self, handle: Any) -> None:
        if self.fail_on_save:
            raise self.fail_on_save
        self.saved.append(handle)

    def close(self, handle: Any) -> None:
        if self.fail_on_close:
            raise self.fail_on_close
        self.closed.append(handle)


@dataclass
class FakeFileSystemPort:
    dirs: set[Path] = field(default_factory=set)
    files: set[Path] = field(default_factory=set)
    children: dict[Path, list[Path]] = field(default_factory=dict)
    writable_dirs: set[Path] | None = None

    def _norm(self, path: Path) -> Path:
        return Path(path)

    def is_dir(self, path: Path) -> bool:
        return self._norm(path) in self.dirs

    def exists(self, path: Path) -> bool:
        p = self._norm(path)
        return p in self.dirs or p in self.files

    def iter_direct_children(self, path: Path) -> list[Path]:
        return list(self.children.get(self._norm(path), []))

    def mkdir(self, path: Path, exist_ok: bool = False) -> None:
        p = self._norm(path)
        if p in self.dirs and not exist_ok:
            raise FileExistsError(str(p))
        self.dirs.add(p)

    def is_writable_dir(self, path: Path) -> bool:
        p = self._norm(path)
        if self.writable_dirs is None:
            return p in self.dirs
        return p in self.writable_dirs

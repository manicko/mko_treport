"""Test doubles for B6 file/session ports."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Sequence

from mko_treport.src.scripts.treport_files.models import AccessMode


@dataclass
class FakeFileSystemPort:
    dirs: set[Path] = field(default_factory=set)
    files: set[Path] = field(default_factory=set)
    children: Dict[Path, List[Path]] = field(default_factory=dict)
    writable_dirs: set[Path] | None = None
    mkdir_failures: Dict[Path, Exception] = field(default_factory=dict)
    created_dirs: List[Path] = field(default_factory=list)

    def _norm(self, path: Path) -> Path:
        return Path(path)

    def is_dir(self, path: Path) -> bool:
        return self._norm(path) in self.dirs

    def exists(self, path: Path) -> bool:
        p = self._norm(path)
        return p in self.dirs or p in self.files

    def iter_direct_children(self, path: Path) -> Sequence[Path]:
        return list(self.children.get(self._norm(path), []))

    def mkdir(self, path: Path, exist_ok: bool = False) -> None:
        p = self._norm(path)
        if p in self.mkdir_failures:
            raise self.mkdir_failures[p]
        if p in self.dirs and not exist_ok:
            raise FileExistsError(str(p))
        self.dirs.add(p)
        self.created_dirs.append(p)

    def is_writable_dir(self, path: Path) -> bool:
        p = self._norm(path)
        if self.writable_dirs is None:
            return p in self.dirs
        return p in self.writable_dirs


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

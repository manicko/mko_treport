"""Abstract ports for filesystem and workbook session (B6 core)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol, Sequence

from .models import AccessMode


class FileSystemPort(Protocol):
    def is_dir(self, path: Path) -> bool: ...

    def exists(self, path: Path) -> bool: ...

    def iter_direct_children(self, path: Path) -> Sequence[Path]: ...

    def mkdir(self, path: Path, exist_ok: bool = False) -> None: ...

    def is_writable_dir(self, path: Path) -> bool: ...


class WorkbookSessionPort(Protocol):
    def open(self, path: Path, access_mode: AccessMode) -> Any: ...

    def save(self, handle: Any) -> None: ...

    def close(self, handle: Any) -> None: ...

"""Session lifecycle orchestration over abstract ports (D1, no real I/O)."""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from .errors import FileAccessSessionError
from .models import OperationMode, SessionPolicy
from .policy import policy_for_mode
from .ports import WorkbookSessionPort


@dataclass
class ManagedWorkbookSession:
    policy: SessionPolicy
    handle: Any
    effective_path: Path
    port: WorkbookSessionPort
    _closed: bool = False

    def save(self) -> None:
        if not self.policy.allow_save:
            raise FileAccessSessionError(
                message="Save is not allowed for this operation mode",
                operation=self.policy.mode.value,
                mode=self.policy.mode.value,
                effective_path=self.effective_path,
                stage="save",
                policy_access_mode=self.policy.access_mode.value,
                cause="save_not_allowed",
                code="save_not_allowed",
            )
        try:
            self.port.save(self.handle)
        except Exception as exc:  # noqa: BLE001
            raise FileAccessSessionError(
                message="Workbook save failed",
                operation=self.policy.mode.value,
                mode=self.policy.mode.value,
                effective_path=self.effective_path,
                stage="save",
                policy_access_mode=self.policy.access_mode.value,
                cause=str(exc),
                code="save_failed",
            ) from exc

    def close(self) -> None:
        if self._closed:
            return
        try:
            self.port.close(self.handle)
            self._closed = True
        except Exception as exc:  # noqa: BLE001
            raise FileAccessSessionError(
                message="Workbook close failed",
                operation=self.policy.mode.value,
                mode=self.policy.mode.value,
                effective_path=self.effective_path,
                stage="close",
                policy_access_mode=self.policy.access_mode.value,
                cause=str(exc),
                code="close_failed",
            ) from exc


@contextmanager
def open_workbook_session(
    mode: OperationMode,
    effective_path: Path | str,
    port: WorkbookSessionPort,
    *,
    operation: str | None = None,
) -> Iterator[ManagedWorkbookSession]:
    policy = policy_for_mode(mode)
    path = Path(effective_path)
    op = operation or policy.mode.value

    try:
        handle = port.open(path, policy.access_mode)
    except Exception as exc:  # noqa: BLE001
        raise FileAccessSessionError(
            message="Workbook open failed",
            operation=op,
            mode=policy.mode.value,
            effective_path=path,
            stage="open",
            policy_access_mode=policy.access_mode.value,
            cause=str(exc),
            code="open_failed",
        ) from exc

    session = ManagedWorkbookSession(
        policy=policy,
        handle=handle,
        effective_path=path,
        port=port,
    )
    primary: BaseException | None = None
    cleanup: FileAccessSessionError | None = None
    try:
        yield session
    except BaseException as exc:  # preserve primary (incl. KeyboardInterrupt/SystemExit unlikely here)
        primary = exc
        raise
    finally:
        try:
            session.close()
        except FileAccessSessionError as exc:
            cleanup = exc
            if primary is None:
                raise
            # Primary preserved; cleanup failure available as context for diagnostics if needed.
            primary.__context__ = exc
        except Exception as exc:  # defensive
            cleanup = FileAccessSessionError(
                message="Workbook close failed",
                operation=op,
                mode=policy.mode.value,
                effective_path=path,
                stage="cleanup",
                policy_access_mode=policy.access_mode.value,
                cause=str(exc),
                code="close_failed",
            )
            if primary is None:
                raise cleanup from exc
            primary.__context__ = cleanup
    _ = cleanup  # keep intent explicit

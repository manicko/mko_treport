"""Deterministic sequential refresh workflow over B6 ports/session (OQ-T2 gated)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, List, Sequence

from mko_treport.src.scripts.treport_files.errors import (
    FileAccessSessionError,
    FileResolutionError,
)
from mko_treport.src.scripts.treport_files.models import AccessMode, OperationMode, ResolvedDataFile
from mko_treport.src.scripts.treport_files.policy import policy_for_mode
from mko_treport.src.scripts.treport_files.ports import FileSystemPort, WorkbookSessionPort
from mko_treport.src.scripts.treport_files.resolver import resolve_data_file

from .errors import RefreshOrchestrationError
from .models import (
    ProgressEvent,
    RefreshBatchItem,
    RefreshCompletionState,
    RefreshExecutionOutcome,
    RefreshFileResult,
    RefreshRequest,
    RefreshStage,
)
from .policy import evaluate_pre_save_gate
from .ports import Clock, ProgressSink, RefreshExecutorPort


def _now(clock: Clock | None) -> datetime | None:
    return clock.now() if clock is not None else None


def _emit(
    sink: ProgressSink | None,
    events: List[ProgressEvent],
    *,
    stage: RefreshStage,
    operation: str,
    configured_data_file: str | None = None,
    row_identity: str | int | None = None,
    resolved_path: Path | str | None = None,
    message: str | None = None,
    clock: Clock | None = None,
    details: dict[str, Any] | None = None,
) -> None:
    event = ProgressEvent(
        stage=stage,
        operation=operation,
        configured_data_file=configured_data_file,
        row_identity=row_identity,
        resolved_path=resolved_path,
        message=message,
        timestamp=_now(clock),
        details=details or {},
    )
    events.append(event)
    if sink is not None:
        sink.emit(event)


def _coerce_completion_state(value: Any) -> RefreshCompletionState:
    if isinstance(value, RefreshCompletionState):
        return value
    if isinstance(value, str):
        try:
            return RefreshCompletionState(value)
        except ValueError:
            return RefreshCompletionState.UNKNOWN
    return RefreshCompletionState.UNKNOWN


def _ensure_outcome(
    outcome: Any,
    *,
    started_at: datetime | None,
    finished_at: datetime | None,
) -> RefreshExecutionOutcome:
    if isinstance(outcome, RefreshExecutionOutcome):
        state = _coerce_completion_state(outcome.completion_state)
        if (
            state is not outcome.completion_state
            or outcome.started_at is None
            or outcome.finished_at is None
        ):
            duration_ms = outcome.duration_ms
            if duration_ms is None and (outcome.started_at or started_at) and (outcome.finished_at or finished_at):
                s = outcome.started_at or started_at
                f = outcome.finished_at or finished_at
                if s and f:
                    duration_ms = (f - s).total_seconds() * 1000.0
            return RefreshExecutionOutcome(
                completion_state=state,
                allows_persist=bool(outcome.allows_persist),
                evidence=outcome.evidence,
                diagnostics=outcome.diagnostics,
                started_at=outcome.started_at or started_at,
                finished_at=outcome.finished_at or finished_at,
                duration_ms=duration_ms,
                error=outcome.error,
            )
        return outcome
    duration_ms = None
    if started_at and finished_at:
        duration_ms = (finished_at - started_at).total_seconds() * 1000.0
    # Invalid/missing-shaped outcome is treated as non-persistable failure evidence.
    return RefreshExecutionOutcome(
        completion_state=RefreshCompletionState.UNKNOWN,
        allows_persist=False,
        diagnostics={"invalid_outcome_type": type(outcome).__name__},
        started_at=started_at,
        finished_at=finished_at,
        duration_ms=duration_ms,
    )


def _close_handle(
    *,
    workbook_port: WorkbookSessionPort,
    handle: Any,
    operation: str,
    mode: OperationMode,
    resolved_path: Path,
    policy_access_mode: str,
) -> FileAccessSessionError | None:
    try:
        workbook_port.close(handle)
        return None
    except Exception as exc:  # noqa: BLE001
        return FileAccessSessionError(
            message="Workbook close failed",
            operation=operation,
            mode=mode.value,
            effective_path=resolved_path,
            stage="close",
            policy_access_mode=policy_access_mode,
            cause=str(exc),
            code="close_failed",
        )


def refresh_data_file(
    *,
    item: RefreshBatchItem,
    input_path: Path | str,
    fs: FileSystemPort,
    workbook_port: WorkbookSessionPort,
    executor: RefreshExecutorPort,
    operation: str = OperationMode.REFRESH_DATA.value,
    progress_sink: ProgressSink | None = None,
    clock: Clock | None = None,
) -> RefreshFileResult:
    events: List[ProgressEvent] = []
    configured = item.configured_data_file
    row_id = item.row_identity
    resolved_path: Path | None = None
    outcome: RefreshExecutionOutcome | None = None
    saved = False
    closed = False

    _emit(
        progress_sink,
        events,
        stage=RefreshStage.RESOLVE_INPUT,
        operation=operation,
        configured_data_file=configured,
        row_identity=row_id,
        message="resolving input",
        clock=clock,
    )
    try:
        resolved: ResolvedDataFile = resolve_data_file(
            configured,
            input_path,
            fs,
            operation=operation,
        )
    except FileResolutionError as exc:
        raise RefreshOrchestrationError(
            message=str(exc),
            operation=operation,
            stage=RefreshStage.RESOLVE_INPUT,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=None,
            component="file_resolver",
            cause=exc.code,
            code=exc.code,
            diagnostics=exc.to_dict(),
        ) from exc
    resolved_path = resolved.selected_path

    policy = policy_for_mode(OperationMode.REFRESH_DATA)
    op = operation or policy.mode.value

    _emit(
        progress_sink,
        events,
        stage=RefreshStage.SESSION_OPEN,
        operation=operation,
        configured_data_file=configured,
        row_identity=row_id,
        resolved_path=resolved_path,
        message="opening workbook session",
        clock=clock,
    )
    try:
        handle = workbook_port.open(resolved_path, AccessMode.WRITABLE)
    except Exception as exc:  # noqa: BLE001
        raise FileAccessSessionError(
            message="Workbook open failed",
            operation=op,
            mode=policy.mode.value,
            effective_path=resolved_path,
            stage="open",
            policy_access_mode=policy.access_mode.value,
            cause=str(exc),
            code="open_failed",
        ) from exc

    primary: BaseException | None = None
    try:
        _emit(
            progress_sink,
            events,
            stage=RefreshStage.REFRESH_TRIGGER,
            operation=operation,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=resolved_path,
            message="invoking refresh executor",
            clock=clock,
        )
        request = RefreshRequest.from_resolved(
            operation=OperationMode.REFRESH_DATA,
            resolved=resolved,
            row_identity=row_id,
            session_handle=handle,
            data_sheet=item.data_sheet,
            context=item.context,
        )
        started = _now(clock)
        try:
            raw_outcome = executor.execute(request)
        except BaseException as exc:  # executor failure -> no save
            finished = _now(clock)
            duration_ms = None
            if started and finished:
                duration_ms = (finished - started).total_seconds() * 1000.0
            outcome = RefreshExecutionOutcome(
                completion_state=RefreshCompletionState.FAILED,
                allows_persist=False,
                diagnostics={"exception_type": type(exc).__name__},
                started_at=started,
                finished_at=finished,
                duration_ms=duration_ms,
                error=exc if isinstance(exc, Exception) else None,
            )
            raise RefreshOrchestrationError.from_outcome(
                message="Refresh executor failed; save forbidden",
                operation=operation,
                stage=RefreshStage.REFRESH_EXECUTION,
                configured_data_file=configured,
                row_identity=row_id,
                resolved_path=resolved_path,
                outcome=outcome,
                component="refresh_executor",
                code="executor_failed",
                cause=str(exc),
            ) from exc

        finished = _now(clock)
        outcome = _ensure_outcome(raw_outcome, started_at=started, finished_at=finished)
        # ensure timing present when clock provided
        if outcome.duration_ms is None and outcome.started_at and outcome.finished_at:
            outcome = RefreshExecutionOutcome(
                completion_state=_coerce_completion_state(outcome.completion_state),
                allows_persist=outcome.allows_persist,
                evidence=outcome.evidence,
                diagnostics=outcome.diagnostics,
                started_at=outcome.started_at,
                finished_at=outcome.finished_at,
                duration_ms=(outcome.finished_at - outcome.started_at).total_seconds() * 1000.0,
                error=outcome.error,
            )

        _emit(
            progress_sink,
            events,
            stage=RefreshStage.AWAIT_COMPLETION,
            operation=operation,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=resolved_path,
            message="refresh outcome received",
            clock=clock,
            details={"completion_state": _coerce_completion_state(outcome.completion_state).value},
        )
        _emit(
            progress_sink,
            events,
            stage=RefreshStage.PRE_SAVE_GATE,
            operation=operation,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=resolved_path,
            message="evaluating pre-save gate",
            clock=clock,
        )
        evaluate_pre_save_gate(
            outcome,
            operation=operation,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=str(resolved_path),
            component="pre_save_gate",
        )
        _emit(
            progress_sink,
            events,
            stage=RefreshStage.SAVE,
            operation=operation,
            configured_data_file=configured,
            row_identity=row_id,
            resolved_path=resolved_path,
            message="saving workbook",
            clock=clock,
        )
        try:
            workbook_port.save(handle)
        except Exception as exc:  # noqa: BLE001
            raise FileAccessSessionError(
                message="Workbook save failed",
                operation=op,
                mode=policy.mode.value,
                effective_path=resolved_path,
                stage="save",
                policy_access_mode=policy.access_mode.value,
                cause=str(exc),
                code="save_failed",
            ) from exc
        saved = True
    except BaseException as exc:
        primary = exc
        raise
    finally:
        cleanup = _close_handle(
            workbook_port=workbook_port,
            handle=handle,
            operation=op,
            mode=policy.mode,
            resolved_path=resolved_path,
            policy_access_mode=policy.access_mode.value,
        )
        if cleanup is None:
            closed = True
        elif primary is None:
            raise cleanup
        else:
            try:
                primary.__context__ = cleanup  # type: ignore[attr-defined]
            except Exception:
                pass

    _emit(
        progress_sink,
        events,
        stage=RefreshStage.CLOSE_CLEANUP,
        operation=operation,
        configured_data_file=configured,
        row_identity=row_id,
        resolved_path=resolved_path,
        message="session closed",
        clock=clock,
    )
    return RefreshFileResult(
        configured_data_file=configured,
        row_identity=row_id,
        resolved_path=resolved_path,
        operation=operation,
        outcome=outcome,
        saved=saved,
        closed=closed,
        progress_events=tuple(events),
    )


def run_refresh_batch(
    *,
    items: Sequence[RefreshBatchItem],
    input_path: Path | str,
    fs: FileSystemPort,
    workbook_port: WorkbookSessionPort,
    executor: RefreshExecutorPort,
    operation: str = OperationMode.REFRESH_DATA.value,
    progress_sink: ProgressSink | None = None,
    clock: Clock | None = None,
) -> List[RefreshFileResult]:
    """Deterministic sequential fail-fast batch. Stops on first failure."""
    results: List[RefreshFileResult] = []
    for item in items:
        result = refresh_data_file(
            item=item,
            input_path=input_path,
            fs=fs,
            workbook_port=workbook_port,
            executor=executor,
            operation=operation,
            progress_sink=progress_sink,
            clock=clock,
        )
        results.append(result)
    return results


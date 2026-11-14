"""Tests for B7 refresh orchestration core."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from mko_treport.src.scripts.treport_files import AccessMode, FileAccessSessionError
from mko_treport.src.scripts.treport_refresh import (
    RefreshBatchItem,
    RefreshCompletionState,
    RefreshEvidence,
    RefreshExecutionOutcome,
    RefreshOrchestrationError,
    RefreshStage,
    evaluate_pre_save_gate,
    refresh_data_file,
    run_refresh_batch,
)

from .fakes import (
    FakeFileSystemPort,
    FakeRefreshExecutor,
    FakeWorkbookSessionPort,
    FixedClock,
    RecordingProgressSink,
)


def _fs_with_file(input_path: Path, filename: str) -> FakeFileSystemPort:
    f = input_path / filename
    return FakeFileSystemPort(dirs={input_path}, files={f}, children={input_path: [f]})


def _completed_outcome() -> RefreshExecutionOutcome:
    return RefreshExecutionOutcome(
        completion_state=RefreshCompletionState.COMPLETED,
        allows_persist=True,
        evidence=RefreshEvidence(details={"proof": "ok"}),
        diagnostics={"step": "done"},
    )


@pytest.mark.parametrize(
    "outcome",
    [
        None,
        RefreshExecutionOutcome(RefreshCompletionState.UNKNOWN, True),
        RefreshExecutionOutcome(RefreshCompletionState.UNKNOWN, False),
        RefreshExecutionOutcome(RefreshCompletionState.NOT_COMPLETED, False),
        RefreshExecutionOutcome(RefreshCompletionState.TIMED_OUT, False),
        RefreshExecutionOutcome(RefreshCompletionState.FAILED, False),
        RefreshExecutionOutcome(RefreshCompletionState.COMPLETED, False),
    ],
)
def test_gate_forbids_all_non_proven_outcomes(outcome):
    with pytest.raises(RefreshOrchestrationError) as ei:
        evaluate_pre_save_gate(
            outcome,
            operation="RefreshData",
            configured_data_file="data",
            row_identity=1,
            resolved_path="C:/in/data.xlsx",
            component="pre_save_gate",
        )
    err = ei.value
    assert err.stage == RefreshStage.PRE_SAVE_GATE
    assert err.code == "persist_blocked"
    assert err.operation == "RefreshData"
    assert err.configured_data_file == "data"


def test_gate_allows_only_completed_with_persist():
    outcome = _completed_outcome()
    evaluate_pre_save_gate(
        outcome,
        operation="RefreshData",
        configured_data_file="data",
        row_identity=None,
        resolved_path=None,
    )


def test_happy_path_sequences_and_captures_hooks():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort()
    clock = FixedClock(datetime(2026, 1, 1, 12, 0, 0))
    sink = RecordingProgressSink()
    executor = FakeRefreshExecutor(outcome=_completed_outcome())
    item = RefreshBatchItem(configured_data_file="data", row_identity=7, data_sheet="S1")

    result = refresh_data_file(
        item=item,
        input_path=in_path,
        fs=fs,
        workbook_port=wb,
        executor=executor,
        progress_sink=sink,
        clock=clock,
    )

    assert result.saved is True
    assert result.closed is True
    assert result.resolved_path == in_path / "data.xlsx"
    assert result.outcome and result.outcome.is_proven_complete
    assert result.outcome.evidence and result.outcome.evidence.details["proof"] == "ok"
    assert len(wb.opened) == 1
    assert wb.opened[0].access_mode is AccessMode.WRITABLE
    assert len(wb.saved) == 1
    assert len(wb.closed) == 1
    assert len(executor.calls) == 1
    req = executor.calls[0]
    assert req.configured_data_file == "data"
    assert req.row_identity == 7
    assert req.data_sheet == "S1"
    assert req.session_handle == wb.saved[0]
    stages = [e.stage for e in result.progress_events]
    assert RefreshStage.RESOLVE_INPUT in stages
    assert RefreshStage.SESSION_OPEN in stages
    assert RefreshStage.REFRESH_TRIGGER in stages
    assert RefreshStage.AWAIT_COMPLETION in stages
    assert RefreshStage.PRE_SAVE_GATE in stages
    assert RefreshStage.SAVE in stages
    assert RefreshStage.CLOSE_CLEANUP in stages
    assert sink.events == list(result.progress_events)


@pytest.mark.parametrize("state", [
    RefreshCompletionState.UNKNOWN,
    RefreshCompletionState.NOT_COMPLETED,
    RefreshCompletionState.TIMED_OUT,
    RefreshCompletionState.FAILED,
])
def test_blocked_states_fail_fast_no_save_but_close(state):
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort()
    executor = FakeRefreshExecutor(
        outcome=RefreshExecutionOutcome(state, False, diagnostics={"s": state.value})
    )

    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data", row_identity=2),
            input_path=in_path,
            fs=fs,
            workbook_port=wb,
            executor=executor,
        )
    err = ei.value
    assert err.stage == RefreshStage.PRE_SAVE_GATE
    assert err.code == "persist_blocked"
    assert err.completion_state == state.value
    assert err.configured_data_file == "data"
    assert err.row_identity == 2
    assert err.resolved_path == str(in_path / "data.xlsx")
    assert wb.saved == []
    assert len(wb.closed) == 1  # close attempted by session context


def test_ambiguous_allows_persist_false_completed_blocks_save():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort()
    executor = FakeRefreshExecutor(
        outcome=RefreshExecutionOutcome(RefreshCompletionState.COMPLETED, False)
    )
    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    assert ei.value.code == "persist_blocked"
    assert ei.value.completion_state == "completed"
    assert ei.value.allows_persist is False
    assert wb.saved == []
    assert len(wb.closed) == 1


def test_executor_exception_treated_as_failure_no_save_primary_preserved():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort()
    executor = FakeRefreshExecutor(exception=RuntimeError("boom-exec"))

    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data", row_identity="r1"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    err = ei.value
    assert err.stage == RefreshStage.REFRESH_EXECUTION
    assert err.code == "executor_failed"
    assert err.completion_state == "failed"
    assert "boom-exec" in (err.cause or "")
    assert isinstance(err.__cause__, RuntimeError)
    assert wb.saved == []
    assert len(wb.closed) == 1
    assert len(executor.calls) == 1


def test_session_open_failure_surfaced_with_stage_context():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort(fail_on_open=OSError("locked"))
    executor = FakeRefreshExecutor(outcome=_completed_outcome())

    with pytest.raises(FileAccessSessionError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    err = ei.value
    assert err.code == "open_failed"
    assert err.stage == "open"
    assert err.mode == "RefreshData"
    assert executor.calls == []
    assert wb.saved == []


def test_save_failure_surfaced_with_stage_context_and_close_attempted():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort(fail_on_save=OSError("disk full"))
    executor = FakeRefreshExecutor(outcome=_completed_outcome())

    with pytest.raises(FileAccessSessionError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    err = ei.value
    assert err.code == "save_failed"
    assert err.stage == "save"
    assert err.mode == "RefreshData"
    # save recorded attempt failed in fake before append; close still attempted by context
    assert wb.saved == []
    assert len(wb.closed) == 1


def test_primary_error_preserved_when_close_fails_after_gate_block():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort(fail_on_close=RuntimeError("close broke"))
    executor = FakeRefreshExecutor(
        outcome=RefreshExecutionOutcome(RefreshCompletionState.FAILED, False)
    )

    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    assert ei.value.stage == RefreshStage.PRE_SAVE_GATE
    assert ei.value.code == "persist_blocked"
    # B6 open_workbook_session attaches cleanup close failure as __context__ when primary exists
    ctx = ei.value.__context__
    assert isinstance(ctx, FileAccessSessionError)
    assert ctx.code == "close_failed"
    assert ctx.stage == "close"
    assert wb.saved == []


def test_resolve_failure_mapped_to_refresh_stage_error():
    in_path = Path("C:/in")
    fs = FakeFileSystemPort(dirs={in_path}, files=set(), children={in_path: []})
    wb = FakeWorkbookSessionPort()
    executor = FakeRefreshExecutor(outcome=_completed_outcome())

    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="missing", row_identity=9),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    err = ei.value
    assert err.stage == RefreshStage.RESOLVE_INPUT
    assert err.code == "not_found"
    assert err.configured_data_file == "missing"
    assert err.row_identity == 9
    assert wb.opened == []
    assert executor.calls == []


def test_batch_deterministic_fail_fast_stops_before_third_item():
    in_path = Path("C:/in")
    files = [in_path / "a.xlsx", in_path / "b.xlsx", in_path / "c.xlsx"]
    fs = FakeFileSystemPort(dirs={in_path}, files=set(files), children={in_path: files})
    wb = FakeWorkbookSessionPort()

    outcomes = {
        "a": _completed_outcome(),
        "b": RefreshExecutionOutcome(RefreshCompletionState.UNKNOWN, False),
        "c": _completed_outcome(),
    }

    class SeqExecutor:
        def __init__(self):
            self.calls = []

        def execute(self, request):
            self.calls.append(request.configured_data_file)
            return outcomes[request.configured_data_file]

    executor = SeqExecutor()
    items = [
        RefreshBatchItem(configured_data_file="a"),
        RefreshBatchItem(configured_data_file="b"),
        RefreshBatchItem(configured_data_file="c"),
    ]

    with pytest.raises(RefreshOrchestrationError):
        run_refresh_batch(
            items=items, input_path=in_path, fs=fs, workbook_port=wb, executor=executor
        )

    assert executor.calls == ["a", "b"]
    # a: open, save, close; b: open, close only (blocked) -> total 2 opens, 1 save, 2 closes
    assert [c.path.name for c in wb.opened] == ["a.xlsx", "b.xlsx"]
    assert len(wb.saved) == 1
    assert len(wb.closed) == 2


def test_invalid_outcome_object_treated_as_non_persistable():
    in_path = Path("C:/in")
    fs = _fs_with_file(in_path, "data.xlsx")
    wb = FakeWorkbookSessionPort()
    executor = FakeRefreshExecutor(outcome={"state": "completed"})  # wrong shape

    with pytest.raises(RefreshOrchestrationError) as ei:
        refresh_data_file(
            item=RefreshBatchItem(configured_data_file="data"),
            input_path=in_path, fs=fs, workbook_port=wb, executor=executor,
        )
    assert ei.value.stage == RefreshStage.PRE_SAVE_GATE
    assert ei.value.completion_state in (None, "unknown")
    assert wb.saved == []
    assert len(wb.closed) == 1


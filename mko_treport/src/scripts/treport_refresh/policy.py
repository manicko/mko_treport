"""Conservative refresh completion/pre-save gate policy."""

from __future__ import annotations

from .errors import RefreshOrchestrationError
from .models import RefreshCompletionState, RefreshExecutionOutcome, RefreshStage


PERSISTABLE_CODE = "persist_allowed"
BLOCKED_CODE = "persist_blocked"


def can_persist(outcome: RefreshExecutionOutcome | None) -> bool:
    if outcome is None:
        return False
    return outcome.is_proven_complete


def evaluate_pre_save_gate(
    outcome: RefreshExecutionOutcome | None,
    *,
    operation: str,
    configured_data_file: str | None,
    row_identity: str | int | None,
    resolved_path: str | None,
    component: str | None = None,
) -> None:
    """Fail-fast unless completion is explicitly proven persistable.

    Unknown/ambiguous/missing/failed/timed_out/not_completed outcomes forbid save.
    """
    if can_persist(outcome):
        return

    if outcome is None:
        raise RefreshOrchestrationError(
            message="Refresh outcome missing; save is forbidden without proven completion",
            operation=operation,
            stage=RefreshStage.PRE_SAVE_GATE,
            configured_data_file=configured_data_file,
            row_identity=row_identity,
            resolved_path=resolved_path,
            component=component,
            completion_state=None,
            allows_persist=None,
            cause="missing_outcome",
            code=BLOCKED_CODE,
        )

    state = outcome.completion_state
    if state is RefreshCompletionState.COMPLETED and outcome.allows_persist is not True:
        cause = "completed_without_persist_permission"
    elif state in {
        RefreshCompletionState.NOT_COMPLETED,
        RefreshCompletionState.TIMED_OUT,
        RefreshCompletionState.FAILED,
        RefreshCompletionState.UNKNOWN,
    }:
        cause = state.value
    else:
        cause = "ambiguous_completion"

    raise RefreshOrchestrationError.from_outcome(
        message="Refresh completion not proven; stale save forbidden",
        operation=operation,
        stage=RefreshStage.PRE_SAVE_GATE,
        configured_data_file=configured_data_file,
        row_identity=row_identity,
        resolved_path=resolved_path,
        component=component,
        outcome=outcome,
        cause=cause,
        code=BLOCKED_CODE,
    )

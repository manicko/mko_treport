"""FR-FILE-5 output path readiness."""

from __future__ import annotations

from pathlib import Path

from .errors import OutputReadinessError
from .ports import FileSystemPort


def ensure_output_ready(
    output_path: Path | str,
    fs: FileSystemPort,
    *,
    operation: str = "ensure_output_ready",
) -> Path:
    out = Path(output_path)

    if fs.exists(out) and not fs.is_dir(out):
        raise OutputReadinessError(
            message="output_path exists but is not a directory",
            operation=operation,
            output_path=out,
            stage="validate_existing",
            cause="output_path_not_dir",
            code="output_path_not_dir",
        )

    if not fs.exists(out):
        try:
            fs.mkdir(out, exist_ok=True)
        except Exception as exc:  # noqa: BLE001 - port boundary mapped to D12
            raise OutputReadinessError(
                message="Failed to create missing output_path",
                operation=operation,
                output_path=out,
                stage="create_missing",
                cause=str(exc),
                code="output_path_create_failed",
            ) from exc

    # Re-check dir state after create (idempotent / race-safe intent).
    if not fs.is_dir(out):
        raise OutputReadinessError(
            message="output_path is not a directory after ensure step",
            operation=operation,
            output_path=out,
            stage="post_create_validate",
            cause="output_path_not_dir",
            code="output_path_not_dir",
        )

    if not fs.is_writable_dir(out):
        raise OutputReadinessError(
            message="output_path is not writable",
            operation=operation,
            output_path=out,
            stage="writability_check",
            cause="output_path_not_writable",
            code="output_path_not_writable",
        )

    return out

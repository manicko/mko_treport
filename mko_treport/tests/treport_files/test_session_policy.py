"""D1 access/session policy tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from mko_treport.src.scripts.treport_files import (
    AccessMode,
    FileAccessSessionError,
    OperationMode,
    open_workbook_session,
    policy_for_mode,
)

from .fakes import FakeWorkbookSessionPort


def test_policy_modes():
    cr = policy_for_mode(OperationMode.CREATE_REPORT)
    assert cr.access_mode is AccessMode.READ_ONLY
    assert cr.allow_save is False
    assert cr.always_close is True
    rd = policy_for_mode(OperationMode.REFRESH_DATA)
    assert rd.access_mode is AccessMode.WRITABLE
    assert rd.allow_save is True
    assert rd.always_close is True


def test_create_report_open_read_only_no_save():
    port = FakeWorkbookSessionPort()
    path = Path("C:/in/report.xlsx")
    with open_workbook_session(OperationMode.CREATE_REPORT, path, port) as session:
        assert port.opened[0].access_mode is AccessMode.READ_ONLY
        assert port.opened[0].path == path
        with pytest.raises(FileAccessSessionError) as ei:
            session.save()
        err = ei.value
        assert err.code == "save_not_allowed"
        assert err.stage == "save"
        assert err.policy_access_mode == "read_only"
        assert err.mode == "CreateReport"
        assert err.effective_path == path
    assert len(port.closed) == 1
    assert port.saved == []


def test_refresh_data_open_writable_save_and_close():
    port = FakeWorkbookSessionPort()
    path = Path("C:/in/data.xlsx")
    with open_workbook_session(OperationMode.REFRESH_DATA, path, port) as session:
        assert port.opened[0].access_mode is AccessMode.WRITABLE
        session.save()
        assert len(port.saved) == 1
    assert len(port.closed) == 1


def test_close_always_attempted_on_body_error():
    port = FakeWorkbookSessionPort()
    path = Path("C:/in/data.xlsx")
    with pytest.raises(RuntimeError, match="boom"):
        with open_workbook_session(OperationMode.REFRESH_DATA, path, port):
            raise RuntimeError("boom")
    assert len(port.closed) == 1


def test_primary_error_preserved_over_cleanup_failure():
    port = FakeWorkbookSessionPort(fail_on_close=RuntimeError("close broke"))
    path = Path("C:/in/data.xlsx")
    with pytest.raises(ValueError, match="primary") as ei:
        with open_workbook_session(OperationMode.REFRESH_DATA, path, port):
            raise ValueError("primary")
    # primary raised; cleanup attached as context
    assert ei.value.__context__ is not None
    ctx = ei.value.__context__
    assert isinstance(ctx, FileAccessSessionError)
    assert ctx.code == "close_failed"
    assert ctx.stage == "close"
    assert len(port.closed) == 0  # fake raises before recording


def test_close_failure_without_primary_is_mapped():
    port = FakeWorkbookSessionPort(fail_on_close=RuntimeError("close broke"))
    path = Path("C:/in/data.xlsx")
    with pytest.raises(FileAccessSessionError) as ei:
        with open_workbook_session(OperationMode.REFRESH_DATA, path, port):
            pass
    assert ei.value.code == "close_failed"
    assert ei.value.stage == "close"
    assert ei.value.policy_access_mode == "writable"
    assert ei.value.effective_path == path
    assert ei.value.mode == "RefreshData"


def test_port_failures_mapped_with_mode_path_stage_open():
    port = FakeWorkbookSessionPort(fail_on_open=OSError("locked"))
    path = Path("C:/in/data.xlsx")
    with pytest.raises(FileAccessSessionError) as ei:
        with open_workbook_session(OperationMode.CREATE_REPORT, path, port):
            pass
    err = ei.value
    assert err.code == "open_failed"
    assert err.stage == "open"
    assert err.policy_access_mode == "read_only"
    assert err.mode == "CreateReport"
    assert err.effective_path == path
    assert "locked" in err.cause
    assert err.to_dict()["operation"] == "CreateReport"


def test_save_failure_mapped():
    port = FakeWorkbookSessionPort(fail_on_save=OSError("disk"))
    path = Path("C:/in/data.xlsx")
    with open_workbook_session(OperationMode.REFRESH_DATA, path, port) as session:
        with pytest.raises(FileAccessSessionError) as ei:
            session.save()
    err = ei.value
    assert err.code == "save_failed"
    assert err.stage == "save"
    assert err.policy_access_mode == "writable"
    assert err.effective_path == path

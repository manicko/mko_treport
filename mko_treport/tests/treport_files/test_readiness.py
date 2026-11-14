"""FR-FILE-5 output readiness tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from mko_treport.src.scripts.treport_files import OutputReadinessError, ensure_output_ready

from .fakes import FakeFileSystemPort


def test_existing_dir_ready():
    out = Path("C:/out")
    fs = FakeFileSystemPort(dirs={out})
    assert ensure_output_ready(out, fs) == out


def test_missing_dir_created():
    out = Path("C:/out/new")
    fs = FakeFileSystemPort(dirs=set())
    assert ensure_output_ready(out, fs) == out
    assert out in fs.created_dirs
    assert out in fs.dirs


def test_uncreatable_fails_with_context():
    out = Path("C:/out/new")
    fs = FakeFileSystemPort(dirs=set(), mkdir_failures={out: PermissionError("denied")})
    with pytest.raises(OutputReadinessError) as ei:
        ensure_output_ready(out, fs, operation="RefreshData")
    err = ei.value
    assert err.code == "output_path_create_failed"
    assert err.stage == "create_missing"
    assert err.operation == "RefreshData"
    assert str(err.output_path) == str(out)
    assert "denied" in err.cause
    assert err.to_dict()["stage"] == "create_missing"


def test_not_dir_fails_with_context():
    out = Path("C:/out/file.txt")
    fs = FakeFileSystemPort(files={out})
    with pytest.raises(OutputReadinessError) as ei:
        ensure_output_ready(out, fs)
    assert ei.value.code == "output_path_not_dir"
    assert ei.value.stage == "validate_existing"


def test_unwritable_fails_with_context():
    out = Path("C:/out")
    fs = FakeFileSystemPort(dirs={out}, writable_dirs=set())
    with pytest.raises(OutputReadinessError) as ei:
        ensure_output_ready(out, fs)
    assert ei.value.code == "output_path_not_writable"
    assert ei.value.stage == "writability_check"


def test_idempotent_existing_dir():
    out = Path("C:/out")
    fs = FakeFileSystemPort(dirs={out})
    ensure_output_ready(out, fs)
    ensure_output_ready(out, fs)
    assert fs.created_dirs == []

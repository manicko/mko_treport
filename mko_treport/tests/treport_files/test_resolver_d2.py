"""D2 resolver tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from mko_treport.src.scripts.treport_files import (
    FileResolutionError,
    RESOLUTION_PRECEDENCE,
    resolve_data_file,
)

from .fakes import FakeFileSystemPort


def _fs_with_children(input_path: Path, names: list[str]) -> FakeFileSystemPort:
    return FakeFileSystemPort(
        dirs={input_path},
        children={input_path: [input_path / n for n in names]},
    )


def test_unique_by_precedence_xlsx_over_xlsb():
    inp = Path("C:/in")
    fs = FakeFileSystemPort(
        dirs={inp},
        children={inp: [inp / "report.xlsx"]},
    )
    res = resolve_data_file("report", inp, fs)
    assert res.selected_name == "report.xlsx"
    assert res.selected_extension == ".xlsx"
    assert res.selected_path == inp / "report.xlsx"
    assert res.base_name == "report"
    assert res.configured_data_file == "report"


def test_not_found_lists_context():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["other.xlsx"])
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("report_w", inp, fs, operation="CreateReport")
    err = ei.value
    assert err.code == "not_found"
    assert err.cause == "not_found"
    assert err.configured_data_file == "report_w"
    assert str(err.searched_input_path) == str(inp)
    assert tuple(err.precedence) == RESOLUTION_PRECEDENCE
    assert err.candidates == ()
    assert err.operation == "CreateReport"
    d = err.to_dict()
    assert d["code"] == "not_found"
    assert d["searched_input_path"] == str(inp)


def test_ambiguous_in_tier_two_xlsx():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["b.xlsx", "a.xlsx"])
    # same base only? names are b.xlsx/a.xlsx (different stems) - not same base.
    # need same base
    fs = _fs_with_children(inp, ["report.xlsx", "report(1).xlsx"])  # stems differ -> no
    fs = _fs_with_children(inp, ["report.xlsx", "Report_copy.xlsx"])
    # better explicit: same stem case-insensitively
    fs = _fs_with_children(inp, ["report.xlsx", "REPORT.xlsx"])
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("report", inp, fs)
    err = ei.value
    assert err.code == "ambiguous_in_tier"
    assert sorted(err.candidates) == ["REPORT.xlsx", "report.xlsx"]


def test_ambiguous_across_tiers_lists_all_candidates():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["report.xlsx", "report.xlsb", "report.xlsm"])
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("report", inp, fs)
    err = ei.value
    assert err.code == "ambiguous_across_tiers"
    assert err.cause == "ambiguous_across_tiers"
    assert err.candidates == ("report.xlsx", "report.xlsm", "report.xlsb")


def test_no_recursion_into_subfolders():
    inp = Path("C:/in")
    sub = inp / "base"
    # direct children contain subdir only; matching file lives "inside" subdir but
    # port returns direct children only so resolver must not find it.
    fs = FakeFileSystemPort(dirs={inp, sub}, children={inp: [sub], sub: [sub / "report.xlsx"]})
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("report", inp, fs)
    assert ei.value.code == "not_found"


def test_case_insensitive_discovery_preserves_actual_name():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["RePoRt_W.XLSX"])
    res = resolve_data_file("report_w", inp, fs)
    assert res.selected_name == "RePoRt_W.XLSX"
    assert res.selected_extension == ".xlsx"
    assert res.base_name == "report_w"


def test_configured_with_extension_strips_for_base_match():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["report_w.xlsm"])
    res = resolve_data_file("report_w.xlsb", inp, fs)
    assert res.selected_name == "report_w.xlsm"
    assert res.selected_extension == ".xlsm"


def test_missing_or_non_dir_input_path_fails_fast():
    inp = Path("C:/missing")
    fs = FakeFileSystemPort(dirs=set())
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("report", inp, fs)
    assert ei.value.code == "input_path_missing_or_not_dir"
    assert ei.value.cause == "input_path_missing_or_not_dir"

    file_path = Path("C:/in.txt")
    fs2 = FakeFileSystemPort(dirs=set(), files={file_path})
    with pytest.raises(FileResolutionError) as ei2:
        resolve_data_file("report", file_path, fs2)
    assert ei2.value.code == "input_path_missing_or_not_dir"


def test_case_variant_duplicates_same_tier_fail_fast():
    inp = Path("C:/in")
    fs = _fs_with_children(inp, ["Data.XLSX", "data.xlsx"])
    with pytest.raises(FileResolutionError) as ei:
        resolve_data_file("data", inp, fs)
    assert ei.value.code == "ambiguous_in_tier"

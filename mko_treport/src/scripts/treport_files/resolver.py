"""Deterministic D2 file resolution in input_path direct children only."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from .errors import FileResolutionError
from .models import RESOLUTION_PRECEDENCE, ResolvedDataFile
from .ports import FileSystemPort


@dataclass(frozen=True)
class _Candidate:
    path: Path
    name: str
    extension: str


def _base_from_configured(configured_data_file: str) -> str:
    p = Path(configured_data_file)
    if p.suffix.lower() in {ext.lower() for ext in RESOLUTION_PRECEDENCE}:
        return p.stem
    return p.name


def _raise(
    *,
    message: str,
    operation: str,
    configured_data_file: str,
    searched_input_path: Path,
    precedence: Sequence[str],
    cause: str,
    code: str,
    candidates: Sequence[str] = (),
) -> None:
    raise FileResolutionError(
        message=message,
        operation=operation,
        configured_data_file=configured_data_file,
        searched_input_path=searched_input_path,
        precedence=tuple(precedence),
        cause=cause,
        code=code,
        candidates=tuple(candidates),
    )


def resolve_data_file(
    configured_data_file: str,
    input_path: Path | str,
    fs: FileSystemPort,
    *,
    operation: str = "resolve_data_file",
    precedence: Sequence[str] = RESOLUTION_PRECEDENCE,
) -> ResolvedDataFile:
    in_path = Path(input_path)
    base = _base_from_configured(str(configured_data_file))
    norm_precedence = tuple(ext.lower() for ext in precedence)

    if not fs.is_dir(in_path):
        _raise(
            message="input_path must exist and be a directory",
            operation=operation,
            configured_data_file=str(configured_data_file),
            searched_input_path=in_path,
            precedence=norm_precedence,
            cause="input_path_missing_or_not_dir",
            code="input_path_missing_or_not_dir",
        )

    # Direct children only, non-recursive.
    children = [Path(p) for p in fs.iter_direct_children(in_path)]
    base_lower = base.lower()
    # Group by lowercased extension among precedence.
    by_ext: dict[str, list[_Candidate]] = {ext: [] for ext in norm_precedence}
    for child in children:
        ext = child.suffix.lower()
        if ext not in by_ext:
            continue
        if child.stem.lower() != base_lower:
            continue
        by_ext[ext].append(_Candidate(path=child, name=child.name, extension=ext))

    tiers_with_matches = [ext for ext in norm_precedence if by_ext[ext]]

    if not tiers_with_matches:
        _raise(
            message="No matching data file found for configured data_file",
            operation=operation,
            configured_data_file=str(configured_data_file),
            searched_input_path=in_path,
            precedence=norm_precedence,
            cause="not_found",
            code="not_found",
        )

    highest = tiers_with_matches[0]
    highest_matches = sorted(by_ext[highest], key=lambda c: c.name)

    # Ambiguity in highest-precedence tier.
    if len(highest_matches) > 1:
        _raise(
            message="Multiple matching files found in highest-precedence tier",
            operation=operation,
            configured_data_file=str(configured_data_file),
            searched_input_path=in_path,
            precedence=norm_precedence,
            cause="ambiguous_in_tier",
            code="ambiguous_in_tier",
            candidates=[c.name for c in highest_matches],
        )

    # Unique highest, but lower-precedence same-base exists -> fail fast across tiers.
    if len(tiers_with_matches) > 1:
        all_candidates: list[str] = []
        for ext in tiers_with_matches:
            for c in sorted(by_ext[ext], key=lambda m: m.name):
                all_candidates.append(c.name)
        _raise(
            message="Lower-precedence same-base file exists alongside highest-precedence match",
            operation=operation,
            configured_data_file=str(configured_data_file),
            searched_input_path=in_path,
            precedence=norm_precedence,
            cause="ambiguous_across_tiers",
            code="ambiguous_across_tiers",
            candidates=all_candidates,
        )

    selected = highest_matches[0]
    return ResolvedDataFile(
        configured_data_file=str(configured_data_file),
        base_name=base,
        input_path=in_path,
        selected_path=selected.path,
        selected_name=selected.name,
        selected_extension=selected.extension,
    )

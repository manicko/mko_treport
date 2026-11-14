"""Settings source boundary for treport config (no Excel runtime coupling)."""

from __future__ import annotations

from typing import Any, Mapping, Protocol, Sequence


RawRow = Mapping[str, Any]
RawTable = Sequence[RawRow]


class SettingsSource(Protocol):
    """Semantic adapter providing raw ListObject/table contents."""

    def read_table(self, table_name: str) -> RawTable:
        """Return rows for a named ListObject in source order."""


def read_all_tables(source: SettingsSource, table_names: Sequence[str]) -> Mapping[str, RawTable]:
    return {name: source.read_table(name) for name in table_names}

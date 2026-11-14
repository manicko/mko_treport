"""Provider abstraction for available filter values (OQ-T1 gate)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Protocol, Tuple


@dataclass(frozen=True)
class YearsContext:
    """Resolved Years scope for dependent field queries."""

    field: str
    values: Tuple[str, ...]


@dataclass(frozen=True)
class FieldDomainResult:
    """Domain/status outcome for a field query in a given context."""

    status: str  # "ok" | "missing_field" | "empty_domain" | "unavailable"
    values: Tuple[str, ...] = field(default_factory=tuple)
    reason: str | None = None

    @property
    def ok(self) -> bool:
        return self.status == "ok"


class AvailableValuesProvider(Protocol):
    """Abstraction over value domains for DSL resolution (B5 pure boundary)."""

    def get_field_domain(
        self,
        *,
        data_file: Any,
        data_sheet: Any,
        field: str,
        years_context: YearsContext | None = None,
        row_identity: str | int | None = None,
    ) -> FieldDomainResult:
        """Return domain/status for field, optionally scoped to years_context."""

"""Contextual config validation errors (D12, config scope)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping


@dataclass(frozen=True)
class ConfigValidationError(ValueError):
    """Fail-fast config validation error with deterministic context."""

    message: str
    operation: str = "config_validation"
    table: str | None = None
    row_identity: str | int | None = None
    field: str | None = None
    value: Any = None
    code: str | None = None

    def __str__(self) -> str:
        parts: list[str] = [self.message]
        context: list[str] = []
        if self.operation:
            context.append(f"operation={self.operation}")
        if self.table is not None:
            context.append(f"table={self.table}")
        if self.row_identity is not None:
            context.append(f"row={self.row_identity!r}")
        if self.field is not None:
            context.append(f"field={self.field}")
        if self.value is not None:
            context.append(f"value={self.value!r}")
        if self.code is not None:
            context.append(f"code={self.code}")
        if context:
            parts.append("(" + ", ".join(context) + ")")
        return " ".join(parts)

    def to_dict(self) -> Mapping[str, Any]:
        return {
            "message": self.message,
            "operation": self.operation,
            "table": self.table,
            "row_identity": self.row_identity,
            "field": self.field,
            "value": self.value,
            "code": self.code,
        }

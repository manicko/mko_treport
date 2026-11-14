"""B5 filter parsing/resolution errors (D12-compatible, no config mutation)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping


@dataclass(frozen=True)
class FilterResolutionError(ValueError):
    """Fail-fast parse/resolve/placeholder error with deterministic context."""

    message: str
    operation: str = "filter_resolution"
    table: str | None = "data_files"
    row_identity: str | int | None = None
    field: str | None = None
    value: Any = None
    code: str | None = None
    data_file: Any = None
    clause: str | None = None
    expression: str | None = None

    def __str__(self) -> str:
        parts: list[str] = [self.message]
        context: list[str] = []
        if self.operation:
            context.append(f"operation={self.operation}")
        if self.table is not None:
            context.append(f"table={self.table}")
        if self.row_identity is not None:
            context.append(f"row={self.row_identity!r}")
        if self.data_file is not None:
            context.append(f"data_file={self.data_file!r}")
        if self.field is not None:
            context.append(f"field={self.field}")
        if self.clause is not None:
            context.append(f"clause={self.clause!r}")
        if self.expression is not None:
            context.append(f"expression={self.expression!r}")
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
            "data_file": self.data_file,
            "clause": self.clause,
            "expression": self.expression,
        }

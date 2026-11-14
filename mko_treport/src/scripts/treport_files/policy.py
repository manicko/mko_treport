"""D1 operation access policy."""

from __future__ import annotations

from .models import AccessMode, OperationMode, SessionPolicy


def policy_for_mode(mode: OperationMode) -> SessionPolicy:
    if mode is OperationMode.CREATE_REPORT:
        return SessionPolicy(
            mode=mode,
            access_mode=AccessMode.READ_ONLY,
            allow_save=False,
            always_close=True,
        )
    if mode is OperationMode.REFRESH_DATA:
        return SessionPolicy(
            mode=mode,
            access_mode=AccessMode.WRITABLE,
            allow_save=True,
            always_close=True,
        )
    raise ValueError(f"Unsupported operation mode: {mode!r}")

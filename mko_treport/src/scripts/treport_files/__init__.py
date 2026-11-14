"""B6 file resolution, output readiness and access/session policy core."""

from .errors import FileAccessSessionError, FileResolutionError, OutputReadinessError
from .models import AccessMode, OperationMode, RESOLUTION_PRECEDENCE, ResolvedDataFile, SessionPolicy
from .policy import policy_for_mode
from .ports import FileSystemPort, WorkbookSessionPort
from .readiness import ensure_output_ready
from .resolver import resolve_data_file
from .session import ManagedWorkbookSession, open_workbook_session

__all__ = [
    "AccessMode",
    "FileAccessSessionError",
    "FileResolutionError",
    "FileSystemPort",
    "ManagedWorkbookSession",
    "OperationMode",
    "OutputReadinessError",
    "RESOLUTION_PRECEDENCE",
    "ResolvedDataFile",
    "SessionPolicy",
    "WorkbookSessionPort",
    "ensure_output_ready",
    "open_workbook_session",
    "policy_for_mode",
    "resolve_data_file",
]

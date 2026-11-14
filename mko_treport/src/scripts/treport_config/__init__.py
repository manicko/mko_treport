"""Treport settings config contract and validation (FR-CFG, D12/D13/D17)."""

from .errors import ConfigValidationError
from .models import DataFileRow, GeneralSettings, PlaceholderRow, ReportRow, TreportConfig
from .reader import SettingsSource
from .validator import validate_config

__all__ = [
    "ConfigValidationError",
    "DataFileRow",
    "GeneralSettings",
    "PlaceholderRow",
    "ReportRow",
    "SettingsSource",
    "TreportConfig",
    "validate_config",
]

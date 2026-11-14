"""Schema contracts for treport settings ListObjects (FR-CFG, D13, D17)."""

from __future__ import annotations

GENERAL_TABLE = "general"
REPORTS_TABLE = "reports"
DATA_FILES_TABLE = "data_files"
PLACEHOLDERS_TABLE = "placeholders"

REQUIRED_TABLES = (GENERAL_TABLE, REPORTS_TABLE, DATA_FILES_TABLE, PLACEHOLDERS_TABLE)

GENERAL_REQUIRED_FIELDS = ("input_path", "output_path")
GENERAL_OPTIONAL_FIELDS = ("delimeter",)
GENERAL_ALLOWED_FIELDS = GENERAL_REQUIRED_FIELDS + GENERAL_OPTIONAL_FIELDS

REPORTS_REQUIRED_FIELDS = ("articleLevel4Name_clean", "targetAudience", "name")
REPORTS_ALLOWED_FIELDS = REPORTS_REQUIRED_FIELDS

DATA_FILES_REQUIRED_FIELDS = ("data_file", "data_sheet", "report_sheet", "filters")
DATA_FILES_OPTIONAL_NOOP_FIELDS = (
    "paste_row",
    "paste_col",
    "freeze_row",
    "freeze_col",
    "zoom",
    "group_rows",
)
DATA_FILES_ALLOWED_FIELDS = DATA_FILES_REQUIRED_FIELDS + DATA_FILES_OPTIONAL_NOOP_FIELDS

PLACEHOLDERS_REQUIRED_FIELDS = ("placeholder", "description")
PLACEHOLDERS_ALLOWED_FIELDS = PLACEHOLDERS_REQUIRED_FIELDS
EXPECTED_PLACEHOLDERS = ("max_weeknum", "max_monthnum", "max_year")

DROPPED_PARAMETERS = frozenset({"new_file_delimeter", "keep_col_hidden"})

# Values that deterministically require multi-value splitting by general.delimeter.
# Config-only check: presence of these characters in configured multi-value fields.
LIST_DELIMETER_TRIGGER_CHARS = (",", ";")

---
id: treport-configuration-model
domain: vba-reporting
tags:
  - vba
  - excel
  - configuration
  - schema
  - settings
related:
  - vba-reporting-overview
  - treport-filters-dsl
  - treport-pipeline-dataflow
---

## Purpose

Defines the configuration model that drives treport behavior: the `settings` workbook structures, key fields, normalization/validation expectations, naming/placeholder rules, and deferred options. This is the SSOT for configuration/schema concepts in governed docs.

## Main Concepts

- Configuration is read from the `settings` sheet/ListObjects by `modTreportConfig` (`ReadConfig`).
- Behavior is data-driven: source resolution, report generation, filtering, layout, and output naming depend on configured rows and values.
- Values are normalized/trimmed (`Norm`) and validated with contextual fail-fast errors.
- `data_files` rows are processed in configured order; this order also drives output sheet assembly.

## Tables and Fields

### `general`

| Field | Purpose (high-level) | Notes |
|---|---|---|
| `input_path` | Base folder for resolving source data workbooks. | Required, non-empty; normalized/validated. |
| `output_path` | Folder for generated report workbooks. | Required, non-empty; readiness/creation checks apply. |
| `delimeter` | Delimiter used by filter list expressions. | Required, non-empty (spelling follows implementation/config contract). |

### `reports`

| Field | Purpose (high-level) | Notes |
|---|---|---|
| `articleLevel4Name_clean` | Report-level category filter value(s) applied per report. | Required, non-empty, trimmed. |
| `targetAudience` | Report-level target audience filter value(s) applied per report. | Required, non-empty, trimmed. |
| `name` | Output report workbook/file naming basis. | Required, non-empty, trimmed; subject to sanitization/collision rules. |

### `data_files`

Rows define source data and sheet/output processing in order.

| Field | Purpose (high-level) | Notes |
|---|---|---|
| `data_file` | Source workbook identifier/base name to resolve. | Required, non-empty. |
| `data_sheet` | Sheet containing the source pivot/data context. | Required, non-empty. |
| `report_sheet` | Output sheet name template for this data file. | Required, non-empty; supports placeholder substitution. |
| `filters` | Filter DSL applied at data-file scope. | Required, non-empty (may contain clauses per DSL). |
| `paste_row` | Target row for snapshot paste/layout. | Layout input; defaults may apply in implementation. |
| `paste_col` | Target column for snapshot paste/layout. | Layout input; defaults may apply in implementation. |
| `freeze_row` | Freeze panes row setting. | Optional/layout control. |
| `freeze_col` | Freeze panes column setting. | Optional/layout control. |
| `zoom` | Sheet zoom setting. | Optional/layout control; normalized/clamped by implementation. |
| `group_rows` | Row grouping/banding option. | Optional column if present; **deferred/no-op** (see below). |

### `placeholders`

| Field | Purpose (high-level) | Notes |
|---|---|---|
| `max_weeknum` | Placeholder value for maximum week number. | Derived in runtime resolution context; referenced in `report_sheet` substitution. |
| `max_monthnum` | Placeholder value for maximum month number. | Derived in runtime resolution context; referenced in `report_sheet` substitution. |
| `max_year` | Placeholder value for maximum year. | Derived in runtime resolution context; referenced in `report_sheet` substitution. |

## Validation and Normalization

Implemented/expected behavior at a high level (see `modTreportConfig`, `modTreportRun`, `modTreportOutput`):

- **Required structures**: `settings` sheet and required ListObjects/columns must exist; missing/invalid structures fail fast with contextual errors.
- **Required values**: key fields in `general`, `reports`, and `data_files` must be non-empty after normalization.
- **Normalization**: `Norm` trims and normalizes null/empty/error-like inputs.
- **Path handling**: input/output paths are normalized/validated; output readiness checks are enforced.
- **Naming rules**: report/file and sheet names are sanitized and constrained (including Excel sheet name length considerations) with collision checks.
- **Fail-fast collisions**: naming/sheet collisions are treated as errors (scope and exact behavior described in [Pipeline and data flow](./pipeline-and-dataflow.md)); successful-run evidence may be blocked by dataset/config collisions.

## Placeholders and Substitution

- `report_sheet` templates support placeholder substitution (implemented in output/naming flow; see `modTreportOutput.SubstitutePlaceholders`).
- Supported derived placeholders: `<max_year>`, `<max_monthnum>`, `<max_weeknum>`.
- Placeholder values are resolved from dependency-aware filter/period resolution during processing, not by copying raw `placeholders` table rows as literals.
- Referencing a placeholder that cannot be derived is expected to fail fast.

## Deferred Behavior

- **`group_rows`**: present in `data_files` schema but intentionally deferred; treated as a no-op in current processing (non-blocking, not an acceptance criterion).

## Runtime Verification Status

- Configuration contract and validation logic are implemented in exported VBA modules, but live workbook configuration contents and runtime validation outcomes (including successful positive-run evidence on specific datasets/configs) are not fully verified in this audit.
- Any discrepancy between documented contract and observed runtime behavior must be treated as unverified until evidenced and, if material, tracked outside governed docs.

## Related Docs

- [VBA reporting overview](./overview.md)
- [Filters DSL](./filters-dsl.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)

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
- `report_sheet` placeholder derivation/substitution is token-driven at runtime from `report_sheet` templates (resolved in code, e.g. `modTreportOutput` / `modTreportRun.ResolveReportSheetPlaceholderMaxValues`). Presence of a `placeholders` ListObject does not imply table-driven runtime placeholder resolution.

## Tables and Fields

### `general`

| Field | Purpose (high-level) | Notes |
|---|---|---|
| `input_path` | Base folder for resolving source data workbooks. | Required, non-empty; normalized/validated. |
| `output_path` | Folder for generated report workbooks. | Required, non-empty; readiness/creation checks apply. |
| `delimeter` | Delimiter used by filter list expressions. | Required, non-empty (spelling follows implementation/config contract). |
| `rows_collapsed` | Default collapsed state for Outline row groups created from `data_files.group_rows`. | Optional. Absent, blank, or `False`/`0`/`No` => expanded (default). `True`/`1`/`Yes` => collapsed readiness/behavior as implemented. Any other present non-blank value fails config validation with context. |
| `report_rows_limit_general` | Global per-subreport row limit applied during combine validation. | **Required (P01).** Explicitly configured positive integer (`>= 1`), Long-compatible, no fractional/Boolean/error/non-numeric/blank/duplicate values and no fallback/default/override. Applied independently to every prepared subreport worksheet in a combine group (including the base/first subreport) against that worksheet's `UsedRange.Rows.Count` before append; it is **not** applied to the total combined report. Projected final combined worksheet dimensions are validated separately against Excel limits before each append paste. Missing/invalid values fail fast with contextual diagnostics. |

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
| `report_sheet` | Output sheet name template for this data file (effective base name before combine suffixing). | Required, non-empty; supports placeholder substitution. After placeholder resolution and sanitization, rows that produce the same effective base worksheet name (case-insensitive) within the **same output workbook** form a combine group in original `data_files` emission order (P01). |
| `filters` | Filter DSL applied at data-file scope. | Required, non-empty (may contain clauses per DSL). |
| `gap_rows` | Spacing rows inserted before this subreport when appended into a combined worksheet. | **Optional (P01).** If the `gap_rows` column is absent, default is `0` for all rows. If present: blank/empty/null => `0`; accepted values are integers `>= 0` (Long-compatible, non-fractional/non-Boolean/non-error/non-negative). Invalid values fail fast with contextual diagnostics. The first/base subreport in a combine group starts the sheet, so its `gap_rows` is ignored during combine. |
| `paste_row` | Target row for snapshot paste/layout. | Layout input; defaults may apply in implementation. For `insert_rows=FALSE` and legacy/non-template `insert_rows=TRUE` (P05) this is the final first row of inserted data. For template-scoped P09 `insert_rows=TRUE` (template set), actual snapshot top becomes `paste_row-1` (requires `paste_row >= 2`) and final first row is expected `paste_row-1` (A1 L3-verified). |
| `paste_col` | Target column for snapshot paste/layout. | Layout input; defaults may apply in implementation. |
| `freeze_row` | Freeze panes row setting. | Optional/layout control. If explicitly set by report/data-file settings, it overrides template display state for freeze panes. |
| `freeze_col` | Freeze panes column setting. | Optional/layout control. If explicitly set, overrides template display state for freeze panes. |
| `zoom` | Sheet zoom setting. | Optional/layout control; normalized/clamped by implementation. If explicitly set, overrides template zoom. |
| `group_rows` | 1-based column index (relative to the copied/pasted snapshot range) used for Excel Outline row grouping. | Optional. If missing/blank/`0`, non-positive, non-integer, out of copied-range width, or otherwise invalid at runtime, Outline grouping is skipped for that subreport with an explicit diagnostic (no fail-fast). |
| `template` | Worksheet template name in macro workbook (`ThisWorkbook`, `treport.xlsm`) used as target-sheet source. | Optional (P04). Absent column or blank/empty => ordinary plain sheet creation (current default behavior). If set: Worksheet only (not chart), case-insensitive lookup in `ThisWorkbook` (visible/hidden/very hidden allowed), fresh copy per subreport use into output workbook, source template never modified, copied sheet renamed to final `data_sheet` and made visible. Missing/unusable template, copy/rename failure, or name conflict fails fast with contextual diagnostics. |
| `only_values` | Controls snapshot transfer mode. | Optional boolean (P04). Default `False` (absent column/blank/`FALSE`/`0`/`NO`) = existing standard transfer (values/number formats, formats, column widths, row heights). `True` (`TRUE`/`1`/`YES`, trimmed case-insensitive) = strict values only into target range (no source formatting/number formats/widths/heights); target/template/prepared-row formatting is preserved. Non-empty unknown value fails config validation. |
| `insert_rows` | Controls row-preparation/space creation for insertion using target sample row. | Optional boolean (P04, behavior revised P05/P09). Default `False` = no row add/delete, insert at `paste_row`. **Template-scoped P09 (2026-11-17, OP-1..OP-4):** when `insert_rows=TRUE` **and** `template` is set (`Len(Norm(template))>0`), uses effective snapshot height `H` (after `skip_header` slicing when applicable) and configured `paste_row` `P` (sample/template row, entire row): requires `P >= 2` (fail-fast with actionable diagnostic before structural mutation/paste if `P < 2`, so snapshot can be placed at `P-1`); if `H >= 3` copy row `P` and insert `H-2` copies starting at `P`, if `H == 2` or `H == 1` insert `0` copies (H==1 allowed in-scope); paste snapshot at `(P-1, paste_col)` over `H x snapshot.Columns.Count`. Final/returned range follows actual placement, expected `P-1 .. P-1+H-1` over `paste_col .. paste_col+W-1` (subject to L3 Excel-shift verification, A1). **Legacy/non-template insert_rows=TRUE (P05, unchanged):** if template is not set, if `H > 1` copy row `P` and insert `H-1` copies starting at `P` (`H==1` => 0 insertions), prepared space `P..P+H-1`, paste at `P,paste_col`, final/returned range `P..P+H-1`; P09 `P>=2`/`P-1` anchor does not apply. In all `insert_rows=TRUE` cases: original `P`/rows below may shift down structurally when rows are inserted; `only_values=TRUE` preserves prepared/target formatting, `FALSE` applies existing snapshot transfer semantics (values/number formats, formats, column widths, row heights). No helper rows/deletion. Final range used for formatting/grouping. Row copy/insert/prep/anchor errors fail fast with contextual diagnostics. `H <= 0` remains fail-fast. |
| `skip_header` | Omits pivot column-header band from copied snapshot. | Optional boolean (P07). Default `False` (absent column/blank/`FALSE`/`0`/`NO`). `True` (`TRUE`/`1`/`YES`, trimmed case-insensitive) = remove top rows above the data/value body within the presented range (column-header/field-caption band). **Preserves** row item labels, data values, and subtotals/grand totals as shown in remaining rows of the sliced presented range; page fields/report filters remain governed by presented-range selection. Basis: `GetPresentedRange` result (`TableRange1`, fallback `TableRange2`). Header offset computed from `DataBodyRange.Row - presentedRange.Row`: `offset==0` => no-op; `offset>0 and offset < presented rows` => slice; `DataBodyRange` missing/unusable, `offset<0`, or `offset >= presented rows` => fail-fast with contextual diagnostics (no silent fallback). Sliced dimensions drive snapshot transfer, `insert_rows` prep/final range, formatting, and bounds-scoped grouping. Initial support/verification scope: layouts used by current `data/theiss` sources (see L3 matrix when executed). Non-empty unknown/error value fails config validation. |

### `placeholders`

`settings` contains a required `placeholders` ListObject (read by configuration loading; missing structure fails validation). In the current workbook snapshot it defines informational rows with columns `placeholder` and `description` (for example `max_Week Number`, `max_Month Number`, `max_Year`). Current runtime logic does not consume `placeholders` table row values for substitution or behavior.

#### `report_sheet` max-placeholder tokens (derived at runtime)

`report_sheet` templates support runtime-derived `max` placeholders resolved pre-filter and independently of `data_files.filters` (see [Pipeline and data flow](./pipeline-and-dataflow.md)):

| Placeholder form | Purpose (high-level) | Notes |
|---|---|---|
| `<max_<PivotTable field name>>` | Maximum numeric value of the named PivotTable field in the relevant pivot/data context. | Supported in `report_sheet` only. Field name is matched exactly as specified inside the token. |
| `<max-N_<PivotTable field name>>` | Offset from maximum: `max_value - N`. | `N` is a base-10 non-negative integer (`max`, `max-N`, `max - N` accepted; case-insensitive `max`). Rules: `0 <= N < max_value`, `N == 0` same as `max`. Computed value must exist among field numeric domain items (exact domain membership). Fails fast with context (field, resolved max, N, computed value/token/template as applicable). Formatting/coercion follows existing max-placeholder behavior (`CLng` when integer-valued, else numeric string). |

## Validation and Normalization

Implemented/expected behavior at a high level (see `modTreportConfig`, `modTreportRun`, `modTreportOutput`):

- **Required structures**: `settings` sheet and required ListObjects/columns must exist; missing/invalid structures fail fast with contextual errors.
- **Required values**: key fields in `general`, `reports`, and `data_files` must be non-empty after normalization. `general.report_rows_limit_general` is required and must be explicitly provided as a positive integer (no default/fallback/override).
- **Normalization**: `Norm` trims and normalizes null/empty/error-like inputs.
- **Optional P04/P07 boolean/string fields**: `template`, `only_values`, `insert_rows`, `skip_header` are optional in `data_files` (backward compatible when columns are absent). Absent column or blank/empty/null values use defaults (`template` empty, booleans `False`). Accepted boolean tokens (trimmed, case-insensitive): `TRUE` = `TRUE`/`1`/`YES` or Excel Boolean `True`; `FALSE` = `FALSE`/`0`/`NO` or Excel Boolean `False`. Non-empty unknown/error values fail config validation with contextual diagnostics (row/data_file/data_sheet/parameter/raw value/expectation). No `height`/`width` config columns are required (snapshot size is runtime-derived).
- **Optional P01 combine spacing field**: `data_files.gap_rows` is optional. If the column is absent, every row defaults to `0`. If present, blank/empty/null => `0`; otherwise only integer values `>= 0` are accepted (stored as Long). Invalid values fail fast with contextual diagnostics.
- **Path handling**: input/output paths are normalized/validated; output readiness checks are enforced.
- **Naming/planning rules (P01)**: `report_sheet` names are resolved (placeholders/substitutions) then sanitized to an effective base worksheet name. Members with the same case-insensitive effective base name within the same output workbook form one combine group in original `data_files`/report emission order. Planned worksheet names are `base` for the first member and `base1..base(N-1)` for subsequent members in order (actual suffix length used). Planned-name length is validated against Excel's 31-character limit before affected worksheet creation; length violations are reported with suffix-aware diagnostics (effective base name without suffix and maximum permitted base length `31 - suffix_length`).
- **Fail-fast collisions (P01)**: before creating affected report worksheets, planned names are preflighted against existing output-workbook worksheet names and all names planned for the current run (case-insensitive), including default sheets. Conflicts fail fast with contextual diagnostics and **no** overwrite/delete/rename resolution. Unused default worksheets may still be removed later only when they do not conflict and are not produced finals (see [Pipeline and data flow](./pipeline-and-dataflow.md)).

## Placeholders and Substitution

- `report_sheet` templates support placeholder substitution (implemented in output/naming flow; `modTreportOutput` / `modTreportRun.ResolveReportSheetPlaceholderMaxValues`).
- Supported derived placeholders: `<max_<field>>` and `<max-N_<field>>` for named PivotTable fields in `report_sheet`.
- Placeholder values are resolved pre-filter from the relevant pivot/data context for the processed `data_files` row and are independent of `data_files.filters` application/selection/derivation.
- Offset validity/domain rules apply (`0 <= N < max_value`, `N==0` same as `max`, computed value must exist in field numeric domain items); invalid/unresolvable offsets or missing domain context fail fast with clear context.
- Existing naming/sanitization/length/collision handling and legacy/malformed/unclosed/unresolved/remnant placeholder guards remain in effect. Referencing a placeholder that cannot be resolved fails fast.

## Row grouping behavior (implemented contract)

- **`group_rows`** (per `data_files` row): after snapshot copy/paste and formatting, Excel Outline row grouping may be applied within the inserted/pasted range bounds only. The grouping column is `group_rows` (1-based, relative to copied range width). Contiguous row runs are grouped when every cell in that column inside bounds has no fill/static fill: pattern none (`xlNone`), solid white `RGB(255,255,255)`, or a solid/system-automatic background treated as white/no-fill for grouping compatibility (e.g. system/indexed background). Conditional formatting is ignored. Grouping is per-subreport and does not cross bounds.
- **Invalid/missing/out-of-width `group_rows`**: deterministic skip for that subreport with explicit diagnostic (no fail-fast, no COM grouping attempt on parameter-invalid cases).
- **`general.rows_collapsed`**: controls collapsed/expanded state for created groups (absent/blank/False => expanded/default; True => collapsed). Invalid present values fail config validation.
- **Runtime verification note:** end-to-end Outline/collapse behavior and edge-case rendering are not fully runtime-verified in this audit.

## Runtime Verification Status

- Configuration contract and validation logic are implemented in exported VBA modules, but live workbook configuration contents and runtime validation outcomes (including successful positive-run evidence on specific datasets/configs) are not fully verified in this audit.
- Any discrepancy between documented contract and observed runtime behavior must be treated as unverified until evidenced and, if material, tracked outside governed docs.

## Related Docs

- [VBA reporting overview](./overview.md)
- [Filters DSL](./filters-dsl.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)

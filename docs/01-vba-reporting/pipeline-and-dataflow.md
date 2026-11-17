---
id: treport-pipeline-dataflow
domain: vba-reporting
tags:
  - vba
  - excel
  - pipeline
  - pivots
  - refresh
  - output
related:
  - vba-reporting-overview
  - treport-configuration-model
  - treport-filters-dsl
---

## Purpose

Documents the high-level execution pipeline and data flow for treport: refresh and report workflows, path/session handling, filter and pivot processing, snapshot extraction, output assembly, and validation/state-safety behavior. This is the SSOT for runtime flow concepts in governed docs.

## Main Concepts

- Two primary workflows: `Refresh_data` (source refresh/persist) and `Create_report` (static report generation).
- Processing is configuration-driven and ordered by `data_files` rows.
- `Create_report` uses read-only source access and never saves source workbooks; `Refresh_data` opens writable and saves refreshed sources.
- Validation is deterministic and fail-fast; application state is captured/restored around runs.

## Refresh_data Flow

High-level implemented sequence (`modTreportRun.Refresh_data`):

1. Capture and apply application runtime state (`modTreportState`).
2. Read and validate configuration (`modTreportConfig.ReadConfig`).
3. Normalize/validate input and output paths/readiness.
4. For each `data_files` row in order:
   - Resolve source workbook path (`modTreportFiles.ResolveDataFilePath`).
   - Open workbook writable (`OpenWorkbookWritable`).
   - Refresh workbook data (`modTreportRefresh.RefreshWorkbookData`) using best-effort synchronous behavior (connections/queries/pivots) before save.
   - Save and close on success; close without save on error paths.
5. Restore application state and report explicit outcome/status.

**Verification note:** Power Query/pivot completion and persisted-state guarantees are environment-dependent and not runtime-proven in this audit.

## Create_report Flow

High-level implemented sequence (`modTreportRun.Create_report` / `ProcessDataFile`):

1. Capture/apply application state and read/validate configuration.
2. Validate paths/readiness and precompute per-report sanitized output names; fail fast on report-filename collisions.
3. Create one output workbook per report and maintain per-output-workbook sheet-name tracking.
4. For each `data_files` row in order:
   - Resolve source workbook and open read-only.
   - Parse/classify `data_files.filters` (including supported `max-N` offset bounds/values), resolve min/max/year/period context against pre-resolved per-field numeric domain, and apply data-file filters once per source.
   - Resolve `report_sheet` `max` placeholders for the row pre-filter and independently of `data_files.filters` (including `<max-N_<field>>` with D1/D2 rules).
   - For each report row: apply report-level filters, verify data presence, derive/substitute `report_sheet` name, enforce sheet-name collision rules, **prepare target sheet** (ordinary plain sheet when `template` is blank; or fresh copy of Worksheet template from `ThisWorkbook` when `template` is set, renamed to final `data_sheet`, made visible, source template unchanged), **insert snapshot with unified P04 pipeline** (`only_values` / `insert_rows`), determine **final inserted range**, apply sheet formatting with template/report precedence, and apply bounds-scoped Outline row grouping on final range when configured/valid (`group_rows`, `general.rows_collapsed`).
   - Close source workbook without save.
5. Remove Excel-default empty sheets from each produced output workbook only on normal successful completion paths (guarded to never delete produced report sheets), enforce produced-sheet invariants, then save output workbooks (overwrite-on-success intent) and close them.
6. Restore application state and report explicit outcome/status.

**P04/P05 unified insertion/formatting sequence (per subreport):** prepare target sheet -> determine `P=paste_row`, `H=snapshot.Rows.Count`, `W=snapshot.Columns.Count` -> if `insert_rows=TRUE` and `H>1`, copy entire target row `P` (sample/template row) and insert `H-1` copies starting at `P` (original `P`/rows below shift down; prepared space `P..P+H-1` inherits sample-row formatting/properties). `H=1` with `insert_rows=TRUE` performs no row insertions. -> insert snapshot at `P,paste_col` over `H x W` (`only_values=TRUE`: strict values only, preserve prepared/target formatting; `FALSE`: existing values/number formats, formats, column widths, row heights transfer semantics) -> final inserted range is `P..P+H-1` (or destination-sized range when `insert_rows=FALSE`, i.e. `P..P+H-1` at `P`) -> apply sheet formatting (template-preserving by default; explicit report zoom/freeze settings override) -> apply Outline grouping on final range/bounds. Snapshot size is derived from presented range (`TableRange1/2`) at runtime (no `height/width` config). **P05 supersedes P04 GA-3 helper-row insertion/delete model for `insert_rows=TRUE`** (no helper rows/deletion). Fail-fast on template-not-found/unusable, copy/rename/target-conflict, invalid params, and row copy/insert/prep errors with contextual diagnostics.

## Path Resolution and Output Readiness

- Source workbooks are resolved from `general.input_path` and `data_files.data_file`.
- **Current implementation behavior (exported code):** resolution searches recursively under `input_path` and requires exactly one total matching workbook across allowed extensions/locations; otherwise it fails fast.
- **Alignment note:** this is stricter on ambiguity and broader in search scope than a simple "pick by precedence" description. Treat external precedence wording as not fully aligned until reconciled; this document records as-implemented behavior.
- Output path readiness/creation and write-precondition checks are enforced before report generation proceeds.

## Pivot Filter Application (P01 Semantics)

High-level semantics as specified/implemented in exported modules (`modTreportPivot`, `modTreportFilters`, `modTreportRun`), **runtime-unverified**:

- **Availability basis:** field handling is determined by PivotTable available fields (not solely by current Filters placement).
- **Placement behavior:**
  - Available-but-inactive fields: move to Filters area, then filter/sort as required.
  - Active fields (in Filters/Columns/Rows): filter in place and do not move fields out of Columns/Rows; preserve layout.
  - Unavailable required fields: fail fast (subject to the uniform selection rule below).
- **Staging:** `data_files.filters` are applied once per source; report-level `articleLevel4Name_clean` and `targetAudience` filters are reapplied per report row without leaking across reports.
- **Matching:** exact-after-trim/value-normalized matching semantics are used.
- **Uniform selection rule (D9):** if a required field is missing/unavailable in the applicable context, or none of the configured values/tokens are matchable, processing fails fast with contextual diagnostics. If at least one configured value is matchable, the matched subset is applied and processing proceeds. This applies uniformly to data-file and report-level filters (including multi-value `articleLevel4Name_clean` split by `general.delimeter` and `targetAudience`); no warn-and-skip path is used for no-match cases.

**Verification note:** P01 runtime behavior, staging edge cases, layout preservation, and dataset-specific pivot behavior are not evidenced in this audit.

## Extraction and Static Assembly

- Presented pivot data range is determined after filtering (`modTreportExtract`, preferring `TableRange1` with fallback behavior as implemented). Optional `skip_header` (P07, per `data_files` row): when `True`, compute header offset from the presented range using `DataBodyRange.Row - presentedRange.Row` (basis = `GetPresentedRange` result). `offset == 0` is a no-op; `offset > 0` and `< presented rows` removes the top column-header/field-caption band above the data/value body (preserving row item labels, data values, and subtotals/grand totals as shown in remaining sliced rows). `DataBodyRange` missing/unusable, negative offset, or offset `>= presented rows` fails fast with contextual diagnostics (`parameter='skip_header'`, no silent fallback). When `False`, behavior unchanged.
- Snapshot transfer supports two modes (P04, `only_values`): **standard** copies values and number formats, cell formats, column widths, and row heights (as before); **values-only** assigns snapshot values to the target range without transferring source formats/number formats/widths/heights, preserving existing target/template/prepared-row formatting (number/date display on blank ordinary targets follows Excel default unless provided by target/template).
- `insert_rows=TRUE` (P05, supersedes P04 GA-3 helper-row model): uses target-sheet row `P=paste_row` as sample/template row (entire row). If `H=snapshot.Rows.Count > 1`, copies row `P` and inserts `H-1` copies starting at `P` to create prepared space `P..P+H-1` inheriting sample-row formatting/properties (original `P` and rows below shift down structurally). If `H=1`, no row insertions are performed. Snapshot is written at `P,paste_col` over `H x W`; final inserted/pasted range is `P..P+H-1`. No helper rows or helper-row deletion. When `FALSE`, insertion starts at `P` with no row add/delete. **`H/W` come from the effective (post-`skip_header` slice when applicable) snapshot range** and are used consistently for transfer, row prep, final bounds, formatting, and grouping.
- Sheet formatting/layout steps are applied after final-range determination. Ordinary (non-template) sheets keep current behavior; template-based sheets preserve template zoom/freeze/gridlines/display settings by default, while explicitly set report `zoom`/`freeze_row`/`freeze_col` values override corresponding template settings.
- **Row grouping (Outline):** when applicable, grouping is applied after paste/formatting/structural finalization on the **final** inserted/pasted range bounds only, using `data_files.group_rows` (1-based relative column) and static-fill predicate (no fill/pattern none, solid white `RGB(255,255,255)`, or solid/system-automatic background treated as white/no-fill for compatibility; conditional formatting ignored). Invalid/missing/out-of-width parameters are skipped for that subreport with explicit diagnostics (no fail-fast). Collapse state follows `general.rows_collapsed` (default expanded).
- **Verification note:** exact presented-range boundaries and visual/data fidelity (including exclusion of pivot/page-filter/PQ artifacts in all cases), `skip_header` slicing correctness across pivot layouts, template/sample-row copy fidelity/side effects (merges, formulas, conditional formatting), and Outline grouping/collapse behavior are not fully runtime-verified in this environment. P07 initial verification scope is layouts used by current `data/theiss` sources (L3 matrix required). L2 manual VBE compile and L3 functional checks on working copies with real data/templates are required to claim compile/behavior.

## Output Workbooks and Sheets

- One output workbook is created per `reports` row.
- Sheets are produced in `data_files` order with `report_sheet` names after placeholder substitution and sanitization. `report_sheet` supports `<max_<field>>` and `<max-N_<field>>` placeholders resolved pre-filter and independently of `data_files.filters`.
- Target sheets are prepared per subreport: ordinary plain worksheet when `template` is blank/default, or a fresh copy of a Worksheet template from `ThisWorkbook` when `template` is set (template reusable, source immutable, copied sheet made visible). Target name conflicts in the output workbook are treated as fail-fast errors (no overwrite/delete).
- Default Excel-empty sheets are removed from normally produced output workbooks on successful completion paths with guards that preserve produced report sheets and enforce invariants before save (fail-fast on invalid produced-sheet state as implemented).
- Successful generation overwrites existing output artifacts as applicable (save semantics implemented in `modTreportOutput`).
- **Sheet-name collisions (current exports):** collision tracking/checking is scoped per output workbook during `Create_report` processing (not a single global namespace across all output workbooks in the same run). Same final sheet name within one output workbook fails fast with contextual diagnostics.
- `group_rows` and `general.rows_collapsed` control bounds-scoped Outline grouping behavior as implemented on the final inserted range (see [Configuration model](./configuration-model.md)); invalid grouping parameters are skipped per subreport with diagnostics.

## Validation, Diagnostics, and State Safety

- **Deterministic validation and fail-fast:** contextual errors are raised for missing/invalid configuration, paths/files/sheets/pivots/fields, malformed/unresolvable filters/placeholders, collisions, and no-data conditions (with defined exceptions where applicable).
- **Explicit outcomes:** workflows report success/failure via user-facing messages/status updates in the implementation.
- **State safety:** application state is captured, adjusted for execution, and restored; cleanup paths attempt to close unsaved sources/outputs and restore UI/calculation state on failure.

## Known Limitations and Unverified Areas

- **Workbook/VBA parity:** workbook binaries and embedded VBA/control bindings are not inspectable here; deployed workbook state may differ from exported sources.
- **Runtime acceptance:** end-to-end `Refresh_data`/`Create_report` runs, PQ/pivot persistence, P01 filter behavior, snapshot/output fidelity, and long-run UX are not runtime-proven.
- **Dataset-specific behavior:** `report_wclip` `Year` vs grouped `Years` semantics and other pivot-member-dependent resolutions are unverified.
- **Build/deployment outcome:** auxiliary build script exists but successful build/compile/sync acceptance is not evidenced (see [VBA build automation](../02-build-and-deployment/vba-build.md)).
- **Evidence gaps:** repository tests currently provide no evidentiary `.py` sources for VBA behavior; discrepancies and verification tracking belong in `.ai/audit/problems` and `.ai/reports`, not as asserted facts in this document.

## Related Docs

- [VBA reporting overview](./overview.md)
- [Configuration model](./configuration-model.md)
- [Filters DSL](./filters-dsl.md)
- [VBA build automation](../02-build-and-deployment/vba-build.md)

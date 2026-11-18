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

High-level implemented sequence (`modTreportRun.Create_report` / per-data-file execution helpers):

1. Capture/apply application state and read/validate configuration.
2. Validate paths/readiness and precompute per-report sanitized output names; fail fast on report-filename collisions.
3. Create one output workbook per report (for inspection/planning context) and initialize per-output registries/state: final/base-produced names, planned members/groups/names, and runtime scaffolding. No planned report worksheets are created before the full per-run/per-output plan is resolved, length-validated, and collision-preflighted.
4. **Plan/resolve/preflight pass (P01, before Stage 1 worksheet creation):** for each `data_files` row in original order, open the source read-only, resolve placeholder max values in the pre-data-file-filter context, and collect resolved `(data_file, report)` emissions (effective base `report_sheet` name after placeholder substitution and sanitization, `gap_rows`, contexts). Build per-output groups by case-insensitive effective base name in original emission order, assign planned worksheet names (`base, base1..base(N-1)` with actual suffix length), validate suffix-aware 31-character limits, and preflight planned names against existing output-workbook worksheet names and all names planned for the run (case-insensitive). Any name-length or collision issue fails fast before affected worksheet creation.
5. **Stage 1 — prepare all subreports (per output, all members/groups):** execute in original `data_files`/report emission order. For each data file reopen source read-only, obtain sheet/pivot, apply data-file filters once and report filters per emission, enforce no-data checks, prepare the target sheet using the B2-assigned planned name (base or temporary), run the existing snapshot transfer/formatting/outline pipeline, and record prepared worksheet/range/runtime state. Only base/final members are registered in the final-produced registry consumed by save/cleanup/verify guards; temporary members remain runtime-only and unregistered.
6. **Stage 2 — combine/finalize (per output, deterministic group order, P01):** for single-member groups no combine is required. For combine groups, append prepared temporary subreports into the base/final worksheet in original order after successful per-group validations (valid content range, per-subreport `report_rows_limit` for base and temps, and projected Excel row/column bounds before each paste). Combine transfer copies values (not formulas), cell/number formats, row heights, hidden row states, and row outline/grouping levels at new positions; it excludes column widths, hidden column states, zoom, freeze panes/splits, and does not reapply `paste_row`/`paste_col`. The first/base subreport's `gap_rows` is ignored. After successful combine, temporary worksheets are deleted and final/base registry/runtime state is reconciled; on failure, processing aborts before save (no false-success cleanup/save).
7. **Save/finalize outputs (successful completion path only):** for each output workbook, require a non-empty final-produced registry, remove unproduced default sheets only when non-conflicting, verify all final-produced sheets are present, then save and close. Source/output workbooks are closed without save on error paths.
8. Restore application state and report explicit outcome/status.

**P04/P05/P09 unified insertion/formatting sequence (Stage 1, per subreport):** prepare target sheet -> determine `P=paste_row`, `H=snapshot.Rows.Count`, `W=snapshot.Columns.Count` (effective after `skip_header` slicing when applicable). **Scope (P09, 2026-11-17, OP-1=A):** `inScopeP09 = (insert_rows=TRUE AND template is set)`.
- **In-scope template P09 (OP-2=B, OP-3=A, OP-4):** require `P >= 2` before structural mutation/paste (fail-fast with actionable diagnostic if `P < 2`). If `H >= 3` copy entire target row `P` (sample/template row) and insert `H-2` copies starting at `P`; if `H == 2` or `H == 1` insert `0` copies (H==1 allowed). No deletions, no snapshot shrink. Write snapshot at `(P-1, paste_col)` over `H x W`. Returned/final range follows actual placement, expected `P-1 .. P-1+H-1` (A1: must be proven/verified against Excel shift behavior in L3, especially H==1/H==2 where no inserts occur).
- **Legacy out-of-scope (P05, unchanged):** `insert_rows=FALSE` inserts at `P` with no row add/delete (final range `P..P+H-1` at `P`). Non-template `insert_rows=TRUE` preserves P05 behavior: if `H>1` copy entire row `P` and insert `H-1` copies starting at `P`, if `H==1` no insertions, prepared/written at `P` with final/returned range `P..P+H-1`. P09 `P>=2` fail-fast and `P-1` anchor do **not** apply to these out-of-scope cases.
- `only_values=TRUE`: strict values only, preserve prepared/target formatting; `FALSE`: existing values/number formats, formats, column widths, row heights transfer semantics. Apply sheet formatting (template-preserving by default; explicit report zoom/freeze settings override) -> apply Outline grouping on final/returned range/bounds. Snapshot size is derived from presented range (`TableRange1/2`) at runtime (no `height/width` config). **P05 supersedes P04 GA-3 helper-row insertion/delete model** for legacy insert path; **P09 reworks only template-scoped insert_rows=TRUE** and preserves legacy/non-template and `insert_rows=FALSE` behavior (Q4=B). Fail-fast on template-not-found/unusable, copy/rename/target-conflict, invalid params, anchor validation (in-scope), and row copy/insert/prep errors with contextual diagnostics.

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

- **Stage 1 (per subreport):** presented pivot data range is determined after filtering (`modTreportExtract`, preferring `TableRange1` with fallback behavior as implemented). Optional `skip_header` (P07, per `data_files` row): when `True`, compute header offset from the presented range using `DataBodyRange.Row - presentedRange.Row` (basis = `GetPresentedRange` result). `offset == 0` is a no-op; `offset > 0` and `< presented rows` removes the top column-header/field-caption band above the data/value body (preserving row item labels, data values, and subtotals/grand totals as shown in remaining sliced rows). `DataBodyRange` missing/unusable, negative offset, or offset `>= presented rows` fails fast with contextual diagnostics (`parameter='skip_header'`, no silent fallback). When `False`, behavior unchanged.
- Snapshot transfer supports two modes (P04, `only_values`): **standard** copies values and number formats, cell formats, column widths, and row heights (as before); **values-only** assigns snapshot values to the target range without transferring source formats/number formats/widths/heights, preserving existing target/template/prepared-row formatting (number/date display on blank ordinary targets follows Excel default unless provided by target/template).
- `insert_rows` placement/row-preparation behavior is defined by the unified P04/P05/P09 sequence above (Stage 1): `insert_rows=FALSE` performs no row insert/delete and places the snapshot at `P,paste_col`; non-template `insert_rows=TRUE` (legacy/P05, template not set) inserts `H-1` sample-row copies at `P` when `H>1` (else none), pastes at `P`, and returns/finalizes `P..P+H-1`; template-scoped P09 (`insert_rows=TRUE` and template set) requires `P>=2`, inserts `H-2` copies at `P` when `H>=3` (else 0 for `H==1|2`), pastes at `P-1`, and returns/finalizes `P-1..P-1+H-1`. In all cases `H/W` come from the effective (post-`skip_header` slice when applicable) presented snapshot range and drive transfer, row preparation, final bounds, formatting, and grouping; `H<=0` is invalid. No helper rows or deletions are used.
- Sheet formatting/layout steps are applied after final-range determination. Ordinary (non-template) sheets keep current behavior; template-based sheets preserve template zoom/freeze/gridlines/display settings by default, while explicitly set report `zoom`/`freeze_row`/`freeze_col` values override corresponding template settings.
- **Row grouping (Outline):** when applicable, grouping is applied after paste/formatting/structural finalization on the **final** inserted/pasted range bounds only, using `data_files.group_rows` (1-based relative column) and static-fill predicate (no fill/pattern none, solid white `RGB(255,255,255)`, or solid/system-automatic background treated as white/no-fill for compatibility; conditional formatting ignored). Invalid/missing/out-of-width parameters are skipped for that subreport with explicit diagnostics (no fail-fast). Collapse state follows `general.rows_collapsed` (default expanded).
- **Stage 2 combine (P01, per combine group):** prepared member sheets are append-combined into the base/final worksheet in original emission order. Before each append, the implementation validates that the subreport has valid content, enforces `report_rows_limit` per prepared subreport (base and temps), and checks projected destination bounds against Excel worksheet limits. Append placement is computed from the base sheet used area and configured `gap_rows` (first/base gap ignored). The transfer profile is values-only (no formulas) with cell/number formats, row heights, hidden row states, and row outline levels restored at destination positions; it must not transfer column widths, hidden column states, zoom, freeze panes/splits, and must preserve the base sheet layout/display and source starting column. Temporary member sheets are removed only after successful combine/finalization for the group.
- **Verification note:** exact presented-range/`UsedRange` boundaries and visual/data fidelity (including exclusion of pivot/page-filter/PQ artifacts in all cases), `skip_header` slicing correctness across pivot layouts, template/sample-row copy fidelity/side effects, Stage 1/Stage 2 Outline grouping/hidden-row restoration at new positions, and large/composite combine scenarios are not fully runtime-verified in this environment. P07 initial verification scope is layouts used by current `data/theiss` sources (L3 matrix required). L2 manual VBE compile and L3 functional checks on working copies with real data/templates are required to claim compile/behavior.

## Output Workbooks and Sheets

- One output workbook is created per `reports` row.
- **Worksheet planning/naming (P01):** for each resolved `(data_file, report)` emission, the effective base worksheet name is derived by placeholder substitution followed by sanitization. Emissions with the same case-insensitive effective base name within the same output workbook form a combine group in original `data_files` emission order. Planned names are `base` for the first member and `base1..base(N-1)` for subsequent members (actual suffix length). Planning validates 31-character limits with suffix-aware diagnostics and preflights all planned names against existing workbook worksheets and all planned run names before affected worksheet creation; conflicts fail fast with no overwrite/delete/rename.
- **Stage 1 preparation:** each planned member (single or combine group) is prepared on its own worksheet using the existing preparation pipeline (`PrepareTargetSheet`, snapshot copy, formatting, outline grouping) in deterministic order. Base/final sheets are registered as produced finals for save guards; temporary sheets are runtime-only until successful Stage 2 finalization.
- **Stage 2 combine/finalization (P01):** for multi-member groups, prepared temps are appended to the base sheet in original order with configured `gap_rows` spacing (first/base gap ignored), required content/limit/bounds validations, and the specified transfer profile (values, formats/number formats, row heights, hidden rows, row outline levels at new positions; base layout/display preserved). Temporary worksheets are deleted only after successful transfer/finalization. Single-member groups do not create temps.
- Default Excel-empty/unproduced sheets are removed from normally produced output workbooks only on successful completion paths, with guards that preserve final-produced report sheets and enforce invariants before save (fail-fast on invalid produced-sheet state as implemented). Non-conflicting unused defaults may be removed; conflicting defaults are rejected earlier by naming preflight.
- Successful generation overwrites existing output artifacts as applicable (save semantics implemented in `modTreportOutput`).
- **Sheet-name collisions:** planning/preflight and produced-sheet tracking are scoped per output workbook (not a single global namespace across all output workbooks in the same run). Planned/existing and planned-vs-planned name collisions, and planned-name length violations, fail fast with contextual diagnostics before affected worksheet creation.
- `group_rows` and `general.rows_collapsed` control bounds-scoped Outline grouping behavior as implemented on the prepared/final inserted range (see [Configuration model](./configuration-model.md)); invalid grouping parameters are skipped per subreport with diagnostics.

## Validation, Diagnostics, and State Safety

- **Deterministic validation and fail-fast:** contextual errors are raised for missing/invalid configuration, paths/files/sheets/pivots/fields, malformed/unresolvable filters/placeholders, no-data conditions, invalid/oversized subreport content or projected Excel bounds during combine, and naming/length/collision issues (with defined exceptions where applicable).
- **Explicit outcomes:** workflows report success/failure via user-facing messages/status updates in the implementation.
- **State safety:** application state is captured, adjusted for execution, and restored; cleanup paths attempt to close unsaved sources/outputs and restore UI/calculation state on failure. On combine/finalization failures, outputs are not saved as successful artifacts.

## Known Limitations and Unverified Areas

- **Workbook/VBA parity:** workbook binaries and embedded VBA/control bindings are not inspectable here; deployed workbook state may differ from exported sources.
- **Runtime acceptance:** end-to-end `Refresh_data`/`Create_report` runs, PQ/pivot persistence, P01 filter behavior (including multi-subreport combine, `gap_rows`, `report_rows_limit`, suffix-aware naming/collisions, projected Excel-bounds checks, transfer fidelity, and temp cleanup), snapshot/output fidelity, and long-run UX are not runtime-proven.
- **Dataset-specific behavior:** `report_wclip` `Year` vs grouped `Years` semantics and other pivot-member-dependent resolutions are unverified.
- **Build/deployment outcome:** auxiliary build script exists but successful build/compile/sync acceptance is not evidenced (see [VBA build automation](../02-build-and-deployment/vba-build.md)).
- **Evidence gaps:** repository tests currently provide no evidentiary `.py` sources for VBA behavior; discrepancies and verification tracking belong in `.ai/audit/problems` and `.ai/reports`, not as asserted facts in this document.

## Related Docs

- [VBA reporting overview](./overview.md)
- [Configuration model](./configuration-model.md)
- [Filters DSL](./filters-dsl.md)
- [VBA build automation](../02-build-and-deployment/vba-build.md)

# Analytical specification: P05 — `insert_rows=TRUE` uses `paste_row` as template/sample row

- **Spec ID:** 005_p05_insert_rows_template_row_spec
- **Status:** Ready for implementation planning (PO decisions confirmed 2026-11-17: all recommended options accepted)
- **Date:** 2026-11-17
- **Source request:** `.ai/problems/P05.md` (2026-11-17)
- **Primary artifacts in scope:** VBA export sources `mko_treport/src/vbs/export/modTreportExtract.bas`, `modTreportRun.bas`, `modTreportConfig.bas`, `modTreportOutput.bas` (authoritative edit targets), docs `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`
- **Baseline context:** `.ai/problems/004_p04_template_insert_spec.md` (GA-1..GA-7, GA-3 helper-row insertion model — **superseded for `insert_rows=TRUE` by this spec per PO decision GA-P05-1**), current implementation in `modTreportExtract.CopyPivotSnapshot` / `PasteSnapshotData` / `DeleteInsertRowsHelpers`, `modTreportRun.ProcessDataFile`

> This document specifies **what** must be achieved, not low-level implementation code. Final end-user runtime remains **VBA-only** in Excel; auxiliary development/deployment/inspection tools are permitted off-runtime only. VBA changes must follow project rules: edit exports under `mko_treport/src/vbs/export/`, backup/working copy, L1 `build_treport_vba.ps1` + report review, **L2 manual VBE `Debug -> Compile VBAProject`** required to claim compilation, L3 on copies when behavior changes. No golden-output workbook required.

## 1. Problem statement

P05 changes the existing `insert_rows=TRUE` algorithm so that the target-sheet row at `paste_row` (`P`) is used as a **formatting/sample (template) row** for the space created under the inserted snapshot.

Required sequence (per P05, confirmed by PO):
1. **Copy sample row:** copy target-sheet row `P` (entire row; formatting and properties needed for new rows).
2. **Create space:** insert the copied row `H - 1` times starting at `P`, where `H = snapshot.Rows.Count` (actual snapshot height; minus 1 accounts for the existing sample row). New rows are placed starting at `P`; original row `P` and rows below shift down and are preserved (shifted structurally).
3. **Insert data:** insert `snapshot` into the created space at row `P`, column `paste_col`, height `H`, width `snapshot.Columns.Count`.
   - `only_values = TRUE`: insert **only values**, preserving formatting of the created rows.
   - `only_values = FALSE`: insert **values and formatting** of the source range according to the existing mechanism.

This **supersedes** the current GA-3/P04 implementation for `insert_rows=TRUE`, which inserts `H` blank rows at `P+1`, performs a temporary insert at `P+1..P+H`, then deletes helper rows at `P+H+1` and `P` via `DeleteInsertRowsHelpers` to yield final data at `P..P+H-1` (`modTreportExtract.bas:38-53,97-116`, `004_p04_template_insert_spec.md` GA-3/CR-4). Under P05 there are **no helper rows** and **no helper-row deletion** in this path.

## 2. Business goal and scope

### 2.1 Business goals
- Ensure rows created for `insert_rows=TRUE` inherit the formatting/properties of the configured sample row at `paste_row` (target-sheet/template-driven), not generic inserted-blank-row defaults.
- Preserve the ability to paste snapshot values while keeping created-row formatting (`only_values=TRUE`) or to transfer snapshot values/formatting via the existing mechanism (`only_values=FALSE`).
- Keep insertion behavior deterministic and compatible with the unified subreport pipeline (ordinary and template-prepared target sheets).

### 2.2 In scope
- Change `insert_rows=TRUE` row-preparation/insertion algorithm in `modTreportExtract.CopyPivotSnapshot` (and related helpers) to follow the P05 sample-row model per PO decisions in §4.
- Define behavior for `H = snapshot.Rows.Count` including edge case `H = 1` (`H-1 = 0` copied rows).
- Preserve `paste_row` / `paste_col` as the final snapshot placement anchor.
- Keep `insert_rows=FALSE` and `only_values` config parsing/defaults unchanged except where required by P05 interaction.
- Maintain GA-4 stage ordering intent: structural row changes before final-range determination, sheet formatting, and `ApplyOutlineGrouping` on final range (`modTreportRun.bas:207-211`).
- Remove helper-deletion call from `insert_rows=TRUE` path and reconcile `DeleteInsertRowsHelpers` API per PO decision (GA-P05-7).
- Update `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`, and record P04 GA-3 supersession/conflict resolution as needed.
- L1/L2/L3 verification on working copies per project workflow (L3 after implementation, on copies only).

### 2.3 Out of scope
- Changes to snapshot source selection (`TableRange1/2` / presented range), `data_files` schema/parsing for `template|only_values|insert_rows`, placeholder/filter logic, or pivot refresh semantics except necessary integration.
- New UI or non-VBA runtime dependencies.
- Redesign of template copy/naming/immutability rules (P04 GA-7) except interaction required by row-copy insertion.
- Guaranteed fidelity for arbitrary unsupported Excel structures (e.g. illegal merged-cell boundaries that Excel cannot insert/copy) beyond fail-fast diagnostics.
- Golden-output acceptance artifacts.

## 3. Confirmed requirements

### 3.1 From P05 request
- **CR-P05-1 (sample row):** For `insert_rows=TRUE`, row `P = paste_row` on the target sheet is the sample row whose formatting/properties must be transferred to created rows.
- **CR-P05-2 (copy + insert count):** Copy entire row `P`; insert copied row `H - 1` times starting at `P`, where `H = snapshot.Rows.Count`.
- **CR-P05-3 (shift/preserve):** Original row `P` and rows below shift down and are preserved structurally during space creation. Sample row is used as formatting/properties source for prepared space; final cell contents in `P..P+H-1` within snapshot bounds are determined by snapshot insertion per `only_values` (overwritten where snapshot is written).
- **CR-P05-4 (snapshot placement):** Insert `snapshot` at row `P`, column `paste_col`, height `H`, width `snapshot.Columns.Count`.
- **CR-P05-5 (only_values=TRUE):** Insert only values, preserving formatting of the created/prepared rows.
- **CR-P05-6 (only_values=FALSE):** Insert values and formatting of source range according to the existing mechanism (same snapshot transfer semantics as today: values+number formats, formats, column widths, row heights from snapshot), applied at final anchor `P`.
- **CR-P05-7 (target context):** Applies to `insert_rows=TRUE` insertion flow on the target sheet per the unified subreport pipeline (ordinary or template-prepared sheet).

### 3.2 Product Owner decisions (confirmed 2026-11-17, all recommended)

| ID | Decision (Option A) | Confirmed choice |
|---|---|---|
| GA-P05-1 | Replace P04 GA-3 helper-row deletion model | **Supersede GA-3 for `insert_rows=TRUE`** with P05 algorithm: no helper rows, no `P+H+1`/`P` deletion; final data at `P..P+H-1` via sample-row duplication + snapshot write at `P`. Remove helper-deletion call from this path; update P04/docs/contracts accordingly. |
| GA-P05-2 | Final position/content of original sample row | Sample row is formatting/properties source for prepared space; final contents in `P..P+H-1` within snapshot bounds determined by snapshot per `only_values`; "preserved (shifted)" is structural except where overwritten by snapshot. |
| GA-P05-3 | Scope of copy/insert | **Entire row `P`** copied/duplicated as sample `H-1` times (whole-row semantics); snapshot written at `P,paste_col` over `H x W`. |
| GA-P05-4 | `only_values=FALSE` mechanism | Keep **same snapshot transfer semantics as today** (values+number formats, formats, column widths, row heights from snapshot) applied at final anchor `P`. |
| GA-P05-5 | `H = 1` | Valid: perform **0** sample-row insertions and write snapshot at `P` per `only_values` (final footprint `P..P`). |
| GA-P05-6 | Row-copy/insert failures | **Fail-fast** with contextual diagnostics (operation, `data_sheet`, `paste_row`, `H/W`, `insert_rows/only_values/template` as applicable), consistent with existing behavior. |
| GA-P05-7 | `DeleteInsertRowsHelpers` API | Remove helper-deletion call from `insert_rows=TRUE` path; **remove unused `DeleteInsertRowsHelpers` logic from exports** if no external callers (preferred, keep module clean) and document change in spec/docs/code comments as appropriate. No in-repo external callers identified by research (usage centered in `modTreportExtract`). |

## 4. Conceptual development tasks

| ID | Task | Purpose | Expected outcome | Dependencies |
|---|---|---|---|---|
| T1 | PO decisions (GA-P05-1..7) | Resolve conflicts with GA-3 and acceptance details | **Done (2026-11-17):** all recommended options A confirmed, including GA-P05-7 disposition = remove unused helper logic. | — |
| T2 | Research closure (as needed during implementation) | Validate Excel row-copy/insert behavior for merged/template/formula/CF/grouping edge cases on real working copies | Evidence-backed notes if issues arise; informs L3 scenarios | T1 |
| T3 | Update insertion algorithm (`modTreportExtract`) | Implement P05 `insert_rows=TRUE` sequence; adjust anchoring/branching; remove helper-deletion call and reconcile/remove `DeleteInsertRowsHelpers` | Correct final range `P..P+H-1`, sample-formatted created rows, `only_values` semantics per PO decisions | T1 |
| T4 | Integrate call/ordering invariants (`modTreportRun`, formatting/grouping) | Ensure GA-4-compatible order and final-range usage remain correct | No regressions in `ApplySheetFormatting`/`ApplyOutlineGrouping` bounds; `insert_rows=FALSE` unchanged | T3 |
| T5 | Diagnostics, config/compat review | Preserve fail-fast errors and P04 parameter behavior; align error codes/messages for new prep failures | Actionable failures with context; no unintended config regressions | T3,T4 |
| T6 | Documentation and spec lineage updates | Align `configuration-model.md`, `pipeline-and-dataflow.md`, and P04/P05 lineage with P05 contract | Docs reflect implemented behavior and GA-3 supersession | T1,T3–T5 |
| T7 | Verification (project workflow) | L1 build+report on working copy, **L2 manual VBE compile** (required to claim compilation), L3 on copies for relevant scenarios | Recorded evidence per project rules (no false compile claims) | T3–T6 |

## 5. Research summary

### 5.1 Current architecture (evidence-backed)
- **Orchestration:** `modTreportRun.ProcessDataFile` resolves filters/placeholders/sheet name, normalizes `paste_row/paste_col` to >=1, prepares target via `modTreportOutput.PrepareTargetSheet` (ordinary or `ThisWorkbook` worksheet template copy, name conflict fail-fast, template immutable), calls `modTreportExtract.CopyPivotSnapshot(..., onlyValues, insertRows)`, applies `ApplySheetFormatting` (template-preserving when template set), then `ApplyOutlineGrouping` on returned range (`modTreportRun.bas:152-215`, `modTreportOutput.bas:356-416`).
- **Current GA-3 insert:** `CopyPivotSnapshot` computes `H/W` from presented pivot range; for `insertRows` inserts `H` blank rows at `P+1`, pastes snapshot at `P+1` via `PasteSnapshotData`, calls `DeleteInsertRowsHelpers(dstWs,P,H,ctx)` deleting `P+H+1` then `P`, returns `P..P+H-1` (`modTreportExtract.bas:16-116`).
- **only_values today:** `TRUE` => `dstRg.Value = srcRg.Value` (preserves target formatting); `FALSE` => `xlPasteValuesAndNumberFormats`, `xlPasteFormats`, `xlPasteColumnWidths`, per-row `RowHeight` (`modTreportExtract.bas:73-95`).
- **Config:** P04 fields parsed with GA-6 defaults/validation (`modTreportConfig.bas:211-336`).

### 5.2 Research findings (completed; actionable for implementation)
- **P05 vs GA-3 conflict:** sample-row-at-`P`, `H-1` inserts starting at `P`, snapshot at `P`, and no helper deletion conflict with current helper-row model and `DeleteInsertRowsHelpers` contract/usage — resolved by **GA-P05-1** (supersede GA-3).
- **Preferred technical approach (confirmed):** duplicate **entire row `P`** as sample `H-1` times starting at `P` (whole-row copy+insert/block) so original shifts down and prepared area carries sample formatting/properties; write snapshot at `P` over `H x W`. For `only_values=TRUE` use direct `Value/Value2` assignment only to preserve created-row formatting; for `only_values=FALSE` keep existing snapshot transfer semantics (GA-P05-4). Handle `H=1` as 0 inserts (GA-P05-5). Fail-fast on row-prep/insert COM errors with context (GA-P05-6).
- **Risks/edge cases to verify in L3:** `paste_row=1`, `insert_rows=FALSE` regression, merged-cell boundaries, formulas/conditional formatting in sample/shifted rows, outline/grouping shifts, column-width/row-height effects under `only_values=FALSE`, template vs ordinary sheets.

## 6. Assumptions

- **A-1:** `paste_row`/`paste_col` remain the configured anchor for snapshot placement; final inserted data starts at `P` with size `H x W`.
- **A-2:** `H = snapshot.Rows.Count >= 1` for valid insertion; `H=1` means 0 extra row copies and snapshot written at `P` (GA-P05-5).
- **A-3:** "Formatting and other necessary properties" of the sample row are those propagated by Excel whole-row copy/insert in the target workbook context (entire-row semantics, GA-P05-3); exact fidelity for merges/formulas/CF is subject to Excel behavior and L3 validation.
- **A-4:** `insert_rows=FALSE` behavior remains unchanged.
- **A-5:** Excel 2021+/Windows, same application instance; operations run on output-workbook copies/targets only (templates in `ThisWorkbook` remain immutable).
- **A-6:** No external callers depend on `DeleteInsertRowsHelpers` outside in-repo usage; removing it per GA-P05-7 is acceptable (research found no external in-repo callers).

## 7. Constraints

- VBA-only runtime; modular exports in `mko_treport/src/vbs/export/*`.
- Follow VBA workflow rules (backup/working copy, L1 report, **L2 manual VBE compile required to claim compilation**, L3 on copies for behavior changes).
- Do not modify canonical/original workbooks or source data during verification; use working copies.
- Deterministic, fail-fast behavior consistent with existing `Create_report` architecture.

## 8. Risks

| Risk | Impact | Mitigation |
|---|---|---|
| Contract break vs P04 GA-3 (`DeleteInsertRowsHelpers`, helper rows, docs) | Inconsistent behavior/docs if not fully reconciled | GA-P05-1/7 confirmed; update spec/docs/code consistently and record supersession |
| Merged cells or protected structures at insertion boundary | COM/runtime failure during row copy/insert | GA-P05-6 fail-fast diagnostics; L3 on real templates; document preconditions |
| Sample-row formulas/CF/objects copied into `H-1` rows or shifted | Unintended duplicated formulas/formatting | GA-P05-2/3 decisions; L3 validation on representative templates |
| `only_values=FALSE` overwrites sample widths/heights/number formats per snapshot semantics | Visual/layout differences vs sample row | GA-P05-4 confirmed; regression checks for ordinary and template targets |
| Grouping/freeze applied to wrong range if final range/anchor miscomputed | Broken outlines/layout | Preserve GA-4 order and return correct final range; L3 checks |
| `H=1` or `paste_row=1` edge cases | Off-by-one, no-op errors | GA-P05-5 confirmed; boundary tests in L3 |

## 9. Open questions

- **Blocking:** none (GA-P05-1..GA-P05-7 resolved 2026-11-17).
- **OQ-P05-V1 (non-blocking):** Exact L3 scenario matrix (combinations of `insert_rows`/`only_values`/template, `H=1`, `paste_row=1`, merges, formulas, grouping) to be selected/executed on available working copies after implementation.
- **OQ-P05-V2 (non-blocking):** If any external/manual caller of `DeleteInsertRowsHelpers` is discovered during implementation, confirm replacement/compatibility note; in-repo research found no external callers.

## 10. Out of scope (explicit)

- Reworking P04 GA-7 template discovery/copy rules beyond required interaction.
- Changes to filter/placeholder/pivot algorithms.
- Partial-run/skip-on-error redesign.
- Golden-output acceptance artifacts.

## 11. Definition of Ready

**Ready for implementation planning.** Implementation remains gated by project verification workflow (L1/L2/L3) and must not claim compilation without L2.

- [x] PO decisions recorded for GA-P05-1..GA-P05-7 (all recommended A, 2026-11-17).
- [x] Contract reconciled in spec with P04 GA-3 (supersession for `insert_rows=TRUE`) and accepted by PO.
- [x] Problem statement, current-state evidence, conceptual tasks, risks, assumptions, constraints, and research summary captured.
- [ ] Implementation plan/tasks sequenced with L1/L2/L3 verification obligations (next step: planning/implementation task breakdown).

## 12. Research traceability

- **Repo/code research:** `.ai/problems/P05.md`; `.ai/problems/004_p04_template_insert_spec.md`; `mko_treport/src/vbs/export/modTreportExtract.bas`, `modTreportRun.bas`, `modTreportConfig.bas`, `modTreportOutput.bas`; `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`.
- **Internal researcher tasks completed:** (1) current architecture/invariants/call sites and P05-vs-GA-3 conflict analysis; (2) VBA/Excel row-duplication and `only_values`/PasteSpecial interaction best practices with recommended approach.
- **Follow-up:** targeted spike on real template workbooks only if implementation/L3 reveals unresolved Excel-specific edge cases (merges/formulas/CF/grouping).

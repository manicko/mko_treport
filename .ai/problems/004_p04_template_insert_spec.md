# Analytical specification: P04 — template-based subreport sheets and unified insertion parameters (`template`, `only_values`, `insert_rows`)

- **Spec ID:** 004_p04_template_insert_spec
- **Status:** Ready for implementation planning (analytical, PO decisions GA-1..GA-7 recorded)
- **Date:** 2026-11-17 (updated with PO decisions)
- **Source request:** `.ai/problems/P04.md` (2026-11-17)
- **PO decisions:** user-provided GA-1..GA-7 (2026-11-17, recorded in §4)
- **Primary artifacts in scope:** VBA export sources `mko_treport/src/vbs/export/*.bas,*.cls` (authoritative edit targets), macro workbook `data/theiss/treport.xlsm` (`settings` / `data_files` ListObject), output workbooks produced by `Create_report`
- **Baseline context:** `.ai/problems/001_treport_spec.md` (rev. 7 lineage), `.ai/problems/003_p01_max_offset_default_sheet_grouping_spec.md`, current implementation in `modTreportRun.bas`, `modTreportExtract.bas`, `modTreportOutput.bas`, `modTreportConfig.bas`, governed docs `docs/01-vba-reporting/configuration-model.md`, `pipeline-and-dataflow.md`

> This document specifies **what** must be achieved. It does not prescribe low-level implementation code. Final end-user runtime remains **VBA-only** in Excel; auxiliary development/deployment/inspection tools are permitted off-runtime only. VBA changes must follow project rules: edit exports, backup/working copy, L1 build script + report, **L2 manual VBE compile** required to claim compilation, L3 on copies when behavior changes. No golden-output workbook required; acceptance is by functional verification against actual source data on working copies. A detailed L3 acceptance matrix is **not mandatory before implementation** — different scenarios can be verified after implementation.

## 1. Problem statement

The current treport subreport pipeline creates a plain output worksheet per report/subreport iteration and copies a static pivot snapshot (values/number formats, formats, column widths, row heights) into it at `paste_row`/`paste_col`. There is no way to (1) control whether source formatting/number formats/widths/heights are transferred versus preserving target formatting, (2) insert/prep rows in the target sheet before pasting and remove helper rows after while keeping `paste_row` as the final data start, or (3) create the target sheet as a copy of a predefined worksheet template stored in the macro workbook.

P04 adds two **independent** capabilities in one unified subreport flow:
1. **Insertion control:** `only_values`, `insert_rows` — apply on **any target sheet** (ordinary newly created sheet or template-based copy).
2. **Target sheet preparation from template:** `template` — controls only how the target sheet is created/prepared; insertion behavior remains governed by the unified mechanism.

All three parameters must be combinable in all valid combinations and must preserve backward compatibility for existing `data_files` configurations.

## 2. Business goal and scope

### 2.1 Business goals
- Enable reports to use pre-designed Excel worksheet templates (layout, merges, formulas, formatting, structure) from `treport.xlsm` without altering the source template.
- Enable pasting subreport data as values-only while preserving target/template formatting/layout, or as standard copy using existing behavior.
- Enable `insert_rows` so pasted blocks expand within target areas via helper-row prep/cleanup with `paste_row` as the stable final start address of inserted data.
- Keep one maintainable, unified insertion pipeline for ordinary and template targets; no duplicated normal-vs-template insertion logic.

### 2.2 In scope
- `settings.data_files` handling for `template`, `only_values`, `insert_rows`: reading, defaults, boolean/value parsing, validation, and backward-compatible behavior when columns/cells are absent or blank (GA-6). **No new `height`/`width` columns** (GA-1).
- Effective inserted-range size derived from actual snapshot dimensions: `snapshot.Rows.Count` (height `H`) and `snapshot.Columns.Count` (width) from the existing snapshot mechanism (`TableRange1/2` / presented range) (GA-1).
- Target-sheet preparation: existing ordinary-sheet path when `template` is blank/default; whole-sheet copy of a Worksheet template from `ThisWorkbook` into output workbook per use, renamed to `data_sheet`, made visible, source template unchanged (GA-7, CR-2).
- Unified snapshot insertion pipeline on the prepared target sheet honoring `only_values` and `insert_rows` (GA-2, GA-3).
- Separate, independently callable helper-row deletion function for `insert_rows=TRUE` (GA-3).
- Stage ordering with structural changes completed before final-range determination, formatting, and grouping; template settings preserved unless explicitly overridden by report settings (GA-4).
- Fail-fast error handling consistent with existing `Create_report` architecture, with contextual diagnostics (GA-5).
- Documentation impact for `docs/01-vba-reporting/configuration-model.md` and `pipeline-and-dataflow.md`.

### 2.3 Out of scope
- Redesigning report architecture beyond agreed VBA/Excel macro model and P04 requirements.
- Changing source pivot/data files, Power Query logic, or filter/placeholder semantics except necessary integration without regression.
- New UI controls or non-VBA runtime dependencies.
- Future toggle UX/config for helper-row deletion beyond requiring a separate function/call point now.
- Partial-report continuation/skip-on-error redesign (GA-5 keeps fail-fast).
- Full runtime acceptance proof in this analysis (L1/L2/L3 on working copies with actual data/templates can be performed after implementation; a predefined L3 acceptance matrix is not a mandatory precondition for execution).

## 3. Confirmed requirements

### CR-1: Independent parameters and combinations
- `template`, `only_values`, `insert_rows` are independent; `only_values` and `insert_rows` apply regardless of ordinary vs template target sheet.
- All 8 combinations must work: (`template` empty/filled) × (`only_values` FALSE/TRUE) × (`insert_rows` FALSE/TRUE).
- Unified insertion/prep/cleanup logic operates on the provided target sheet — no separate normal-vs-template insertion implementations.

### CR-2: Target sheet preparation (`template`) (GA-7, GA-5)
- **`template` empty/not set (default):** use existing mechanism — create new sheet in output workbook, name per `data_sheet`/existing rules, prepare for insertion (backward-compatible behavior).
- **`template` filled:** find template only in `ThisWorkbook` (macro workbook `treport.xlsm`), as a **Worksheet** (not Chart Sheet). Use Excel-standard case-insensitive sheet-name lookup. Visible, hidden, and VeryHidden sheets are allowed as templates.
- For **each use** create a separate fresh copy of the template into the output workbook; rename copied sheet per `data_sheet` using existing validation/normalization rules; set copied/final sheet visible (`xlSheetVisible`). The same template may be reused by multiple `data_files` rows.
- **Source template must never be modified.** All subsequent operations apply only to the copy.
- **Do not overwrite or delete existing sheets** on name conflict where existing logic does not provide safe handling — treat as error (GA-5).
- **Errors (fail-fast, contextual):** template not found or not usable, copy failure, invalid final sheet name, or name conflict per above. Error message must include `data_sheet` and problematic parameter/template name.

### CR-3: `only_values` (GA-2)
- **`TRUE`:** copy **strictly values only** from source snapshot range.
  - Do not transfer source formatting, source number formats, column widths, or row heights.
  - Do not alter existing formatting of the target range/sheet (template or preformatted target).
  - Prefer direct value assignment (`Value` or `Value2`) rather than current multi-step `PasteSpecial` transfer.
  - Number formats remain as already defined on the target sheet. For an ordinary new blank sheet, Excel default (General) display applies; no additional formatting is performed in this mode. If specific date/number display is required, it must be provided by the target sheet/template in advance.
- **`FALSE` (default):** preserve existing standard behavior (values and formatting using current mechanism).
- Behavior is identical for ordinary and template targets.

### CR-4: `insert_rows` coordinate model and helper-row deletion (GA-1, GA-3)
- Applies to the **target sheet** (ordinary or template).
- **`FALSE` (default):** normal insertion — determine position by `paste_row`,`paste_col`; insert starting at `paste_row`; do not add or delete target-sheet rows.
- **`TRUE`:** `paste_row` (`P`) is the **final first row of inserted data after all operations complete** (logical address unchanged).
  - Let `H = snapshot.Rows.Count` (actual snapshot height, GA-1); require `H > 0`.
  - **Step 1 (row prep):** add `H` rows starting at row `P + 1`.
  - **Step 2 (temporary insert):** insert snapshot data into rows `P + 1 .. P + H`, starting at column `paste_col`, honoring `only_values`.
  - **Step 3 (cleanup, separate function):** delete helper row directly below range — row `P + H + 1` — **then** delete helper row directly above range — row `P` (delete bottom before top to avoid extra coordinate shift issues). Pass target sheet and inserted-range parameters to the function.
  - **Result:** final inserted data occupies rows `P .. P + H - 1`. After all operations, re-determine the actual inserted data range.
- **Boundary/assumptions:** `paste_row = 1` is allowed (row above is first sheet row and may be deleted after insert as helper). Rows above and below deleted as helper rows are removed as whole Excel rows and must be provisioned as helper rows not containing data that must be preserved.
- **Errors (fail-fast):** invalid prep/insert/delete operations, invalid snapshot height, or inability to complete required structural steps (GA-5).

### CR-5: Stage ordering and post-processing integration (GA-4)
1. Create/prepare target sheet (ordinary or template copy).
2. Add rows if `insert_rows=TRUE`.
3. Insert data.
4. Delete helper rows if `insert_rows=TRUE`.
5. Determine final coordinates of inserted data range.
6. Apply sheet formatting considering report settings and template.
7. Apply outline row grouping (`ApplyOutlineGrouping`) using **final** range coordinates.

- Grouping must use final coordinates after helper-row deletion and remain within inserted/subreport bounds (consistent with existing grouping constraints).
- **Template settings:** when a template is used, preserve its formatting/display settings by default. If report settings explicitly set zoom, freeze panes, or gridlines, report settings take precedence over template. If report settings are not set, preserve corresponding template settings (do not forcibly reset). Do not unconditionally clear FreezePanes or force Gridlines off for template sheets unless explicitly set by report settings.
- **Ordinary sheets without template:** preserve current `ApplySheetFormatting` behavior.

### CR-6: Boolean parsing and defaults (GA-6)
- **TRUE:** Excel Boolean TRUE, number `1`, text `TRUE`, text `YES` (trimmed, case-insensitive).
- **FALSE:** Excel Boolean FALSE, number `0`, text `FALSE`, text `NO` (trimmed, case-insensitive).
- **Defaults:** `template` = empty/absent value; `only_values` = `FALSE`; `insert_rows` = `FALSE`.
- If column is absent in older config or cell value is empty => use default.
- If a non-empty cell contains an unknown value => configuration error (fail-fast) with clear message.

### CR-7: Fail-fast error model and diagnostics (GA-5)
- Preserve existing `Create_report` fail-fast architecture (stop formation on critical error).
- Error cases include at minimum: template not found/unusable, template copy failure, invalid final sheet name, name conflict not safely handled by existing logic, invalid parameter values, row add/delete errors, data insertion errors.
- Error messages must be actionable and include `data_sheet` and the problematic parameter or template name (and other relevant context where available).
- No partial-formation/skip-subreport redesign required for this task.

### CR-8: Architecture and separation
- Separate independent operations/functions: (1) target-sheet preparation (ordinary vs template copy), (2) row preparation (add rows when `insert_rows=TRUE`), (3) data insertion (values-only vs existing standard), (4) helper-row deletion (above/below inserted range).
- Main subreport algorithm determines parameters and calls functions sequentially; no hardcoded row-deletion logic in main insertion path.
- `only_values` and `insert_rows` work identically regardless of `template`.

## 4. Product Owner decisions (GA-1..GA-7)

| ID | Decision | Summary |
|---|---|---|
| GA-1 | **A** | No `height`/`width` columns in `data_files`. Use actual snapshot size: `H = snapshot.Rows.Count`, width = `snapshot.Columns.Count` (from existing `TableRange1/2` snapshot mechanism). `insert_rows=TRUE` adds `H` rows; all range operations use actual snapshot size. |
| GA-2 | **A (target-preserving)** | `only_values=TRUE` = strict values only (prefer direct `Value/Value2` assignment): no source formatting/number formats/column widths/row heights; preserve existing target formatting. Target number formats remain as on target sheet/template; ordinary blank sheet displays Excel default (General). `only_values=FALSE` keeps existing behavior. |
| GA-3 | **Confirmed model** | `paste_row` (`P`) is final first data row after all operations. For `insert_rows=TRUE`, `H=snapshot.Rows.Count`: add `H` rows at `P+1`; temp insert to `P+1..P+H`; delete helper row `P+H+1` (below) then `P` (above) via separate function; final data `P..P+H-1`; re-determine actual range. `paste_row=1` allowed; helper rows must be expendable whole rows. `insert_rows=FALSE` => no row add/delete, insert at `P`. |
| GA-4 | **Confirmed** | Complete structural row changes before final-range determination, formatting, grouping. Order: prepare sheet -> add rows -> insert -> delete helpers -> final range -> sheet formatting (respect template/report precedence) -> grouping on final coords. Template settings preserved by default; explicit report zoom/freeze/gridlines override. Do not unconditionally reset FreezePanes/Gridlines for templates unless explicitly set. Ordinary sheets keep current `ApplySheetFormatting` behavior. |
| GA-5 | **C (fail-fast)** | Preserve existing `Create_report` fail-fast model. Template-not-found/unusable/copy/name/conflict/invalid params/row ops/insert errors stop formation with clear message including `data_sheet` and problematic parameter/template. No partial-formation redesign. |
| GA-6 | **A** | Accept TRUE: Excel TRUE/1/`TRUE`/`YES`; FALSE: Excel FALSE/0/`FALSE`/`NO` (trim, case-insensitive). Defaults: `template` empty, `only_values` FALSE, `insert_rows` FALSE. Missing column or empty cell => default. Non-empty unknown value => config error. |
| GA-7 | **Confirmed** | Template lookup only in `ThisWorkbook`, Worksheet only (not Chart), Excel case-insensitive name match; visible/hidden/VeryHidden allowed. Fresh copy per use into output workbook; template reusable across `data_files`; template unchanged. Final copied sheet visible (`xlSheetVisible`). Final name from `data_sheet` via existing rules. No overwrite/delete of existing sheets on conflict (error). Template-not-found/unusable => fail-fast diagnostic. |

## 5. Conceptual development tasks

| ID | Task | Purpose | Expected outcome | Dependencies |
|---|---|---|---|---|
| T1 | Config contract extension | Support `template,only_values,insert_rows` read/default/parse/validation per GA-6 with backward-compatible absent/blank handling; no `height/width` columns (GA-1) | Correct `TreportDataFile`/`ReadConfig` behavior and contextual config errors | — |
| T2 | Target-sheet preparation abstraction | Unified ordinary creation and template copy/rename/visibility from `ThisWorkbook` per GA-7; fail-fast diagnostics per GA-5 | Single prepared-target entry point returning usable worksheet copy for any subreport | T1 |
| T3 | Unified insertion/row-prep pipeline | Stage sequence: optional row insert at `P+1` (`H` rows), temp insert at `P+1..P+H` honoring `only_values` (GA-2), using actual snapshot size (GA-1) | One deterministic insertion mechanism for ordinary/template targets | T1, T2 |
| T4 | Helper-row deletion function | Separate callable shift-safe function: delete `P+H+1` then `P` (GA-3); boundary handling; re-determine actual inserted range | Correct final range `P..P+H-1` and stable `paste_row` semantics | T3 |
| T5 | Post-processing integration | Apply GA-4 ordering: final range -> formatting with template/report precedence -> `ApplyOutlineGrouping` on final coords; preserve template settings unless explicitly overridden | No regressions in grouping/freeze/zoom/gridlines/layout; correct template behavior | T3, T4 |
| T6 | Diagnostics, robustness, compatibility | Enforce GA-5 fail-fast cases/messages (`data_sheet` + problematic parameter/template); preserve existing invariants and no template/source mutation on failure paths | Predictable, actionable failures; backward compatibility for existing configs | T1–T5 |
| T7 | Documentation and verification plan | Update `docs/01-vba-reporting/configuration-model.md`, `pipeline-and-dataflow.md`; identify L1/L2 requirements and recommended post-implementation L3 scenarios (all 8 combinations and edge cases: GA decisions, template missing, `paste_row=1`, invalid params, number/date display on blank vs template targets, grouping/freeze precedence) that can be checked after implementation | Documentation impact and traceability; post-implementation verification scenarios identified (L3 matrix not mandatory before execution) | T1–T6 |

## 6. Research summary

### 6.1 Current architecture (evidence-backed)
- `ProcessDataFile` (`modTreportRun.bas`): per `data_files` in order -> open source read-only, resolve `report_sheet` max placeholders pre-filter, parse/apply data-file filters once, per report apply report filters/no-data fail, substitute/sanitize sheet name + per-output collision check, create target via `AddReportSheet`, copy snapshot via `CopyPivotSnapshot`, apply `ApplySheetFormatting`, apply `ApplyOutlineGrouping` on inserted range.
- `CopyPivotSnapshot` (`modTreportExtract.bas`): presented range from `TableRange1/2`; static copy via `xlPasteValuesAndNumberFormats`, `xlPasteFormats`, `xlPasteColumnWidths` + per-row row heights; returns destination range sized to snapshot. Paste anchor defaults/clamps to 1.
- `AddReportSheet` (`modTreportOutput.bas`): plain `Worksheets.Add` at end (after deleting same-named sheet in target workbook); no template support.
- `TreportDataFile`/`ReadConfig` (`modTreportConfig.bas`): no `template,only_values,insert_rows,height,width`; `group_rows` read only if column exists (runtime-tolerant precedent); `rows_collapsed` has strict boolean parse/default/fail-fast precedent.
- `ApplyOutlineGrouping` is bounds-scoped to `insertedRng` and fill-based; `ApplySheetFormatting` activates sheet, clears freeze panes, hides gridlines, applies zoom/freeze via `ActiveWindow` (GA-4 requires conditional/template-preserving behavior).
- Run-level behavior is predominantly fail-fast with contextual `Fail`; `Create_report` aborts and closes without save on raised errors (GA-5 preserves model).

### 6.2 Best-practice findings (VBA/Excel 2021, Windows)
- **Template copy:** `ThisWorkbook.Worksheets(template).Copy After:=outWb.Sheets(outWb.Sheets.Count)` in same Excel instance, then rename/make visible in output workbook; preserves full sheet fidelity; never modify template.
- **Values-only, preserve target:** direct `Range.Value/Value2` assignment avoids transferring source formats/number formats/widths/heights (GA-2).
- **Row prep:** block `Rows(...).Insert Shift:=xlDown` at `P+1` for `H` rows; coordinate shifts are deterministic. GA-3 fixes the final-position contract (`P..P+H-1`) and delete order (`P+H+1` then `P`) via separate function.
- **Formatting/grouping:** apply grouping and bounds-dependent operations only on final post-delete range; preserve template display settings unless explicitly overridden.

### 6.3 Preferred implementation direction (research-backed, aligned to GA-1..GA-7)
- Unified stage pipeline: **prepare target (ordinary/template copy)** -> **optional row insert** -> **insert data (`Value/Value2` when `only_values`, existing mechanism when not)** -> **optional helper-row delete (separate fn)** -> **final range** -> **conditional sheet formatting** -> **grouping on final range**.
- No `height/width` config fields; use `snapshot.Rows.Count/Columns.Count`.
- Strict GA-6 boolean parsing/defaults; GA-5 fail-fast diagnostics including `data_sheet` and problematic parameter/template.
- Template lookup in `ThisWorkbook`, Worksheet-only, case-insensitive, visible/hidden/VeryHidden allowed, fresh visible copy per use, immutable source.

## 7. Assumptions

- **A-1:** Macro workbook containing templates is `ThisWorkbook` (`treport.xlsm`) runtime context.
- **A-2:** One prepared target worksheet per subreport instance in current model; unified mechanism operates on that target sheet.
- **A-3:** Snapshot size and presented source range are obtained via existing snapshot mechanism (`TableRange1/2`) consistently with current behavior; `H = snapshot.Rows.Count > 0` when inserting.
- **A-4:** Helper rows deleted at `P` and `P+H+1` (per GA-3 sequence) are expendable whole rows with no required preserved data.
- **A-5:** `data_sheet` naming/sanitization/validation follows existing project rules except where conflict handling is explicitly an error per GA-5/GA-7.
- **A-6:** Excel 2021+/Windows, same Excel application instance for template copy and output operations.
- **A-7:** For `only_values=TRUE` on ordinary blank sheets, General/default Excel display is acceptable unless target/template provides formats (GA-2).

## 8. Constraints

- **VBA-only runtime** for end-user operations; no new external runtime dependencies.
- **Modular export architecture** preserved; edit `mko_treport/src/vbs/export/*` and follow build/sync/verification workflow (backup/working copy, L1 report, **L2 manual VBE compile** required to claim compilation, L3 on copies for behavior changes).
- **Template immutability:** template sheets in `ThisWorkbook` must never be modified.
- **Determinism:** configured-order processing and repeatable outcomes for same inputs/config/data state (subject to Excel/pivot data state).
- **No invention:** implementation must follow GA-1..GA-7; any newly discovered edge case outside these decisions must be escalated before guessing user-visible behavior.

## 9. Risks

| Risk | Impact | Mitigation |
|---|---|---|
| Coordinate/shift errors in GA-3 sequence | Misplaced data, wrong deletions, broken grouping/freeze | Implement exact GA-3 algorithm (delete `P+H+1` then `P`), recompute final range `P..P+H-1`, boundary tests incl. `paste_row=1`, L3 on working copies |
| `only_values` (Value/Value2) loses date/number display on unformatted targets | Misleading presentation on blank ordinary sheets | Document/rely on target/template-provided formats per GA-2; L3 cases for blank vs preformatted/template targets |
| Template-preserving formatting integration (GA-4) vs current `ApplySheetFormatting` behavior | Regressions (forced gridlines off/freeze cleared) or unintended template overrides | Conditional logic: preserve template settings by default; apply explicit report settings as overrides; regression tests for ordinary and template sheets |
| Grouping/freeze using pre-delete coordinates | Outline/freeze attached to wrong rows | Enforce GA-4 stage order and pass final range to `ApplyOutlineGrouping`/layout steps |
| Name conflicts and template usability checks (GA-7/GA-5) | Fail-fast stops expected by design; possible surprise for existing conflicting configs | Clear diagnostics with `data_sheet` and parameter/template; document behavior |
| Schema/backward-compat in live workbooks | Older configs fail or behave unexpectedly | GA-6 absent/blank defaults, no new required `height/width` columns, optional-column-safe reads |
| Template copy side effects (formulas/links/objects in different output context) | Output differs from template expectations | Whole-sheet copy in same instance, immutable source, L3 validation on real templates/data |

## 10. Open questions

No blocking PO decisions remain (GA-1..GA-7 resolved). Implementation-time clarifications to capture during/after implementation (non-blocking if handled by stated rules/evidence):
- **OQ-V1 (post-implementation verification, optional):** A detailed L3 acceptance matrix is **not mandatory for execution**. After implementation, different scenarios can be checked on working copies as needed (e.g. all 8 combinations, `paste_row=1`, invalid/unknown param values, template-not-found/name-conflict, date/number display on blank vs template targets, freeze/gridlines/zoom/grouping precedence), using available real template sheets and data.
- **OQ-V2 (integration detail):** Exact interaction of existing per-output sheet-name collision registry/`AddReportSheet` behavior with GA-7 "no overwrite/delete existing sheets on conflict" must be implemented consistently with current invariants and GA-5 diagnostics (implementer to verify against existing flow and document).
- **OQ-V3 (template precedence mapping):** Which existing `data_files`/formatting fields count as "explicitly set" report overrides for zoom/freeze/gridlines per GA-4 (implementer to map to current config semantics and verify in post-implementation checks as needed).

## 11. Out of scope (explicit)

- Implementation/code changes as part of this analysis deliverable.
- Adding `height`/`width` to `data_files` (GA-1).
- Partial report formation/skip-on-error behavior (GA-5).
- Changes to pivot refresh/filter/placeholder algorithms except necessary integration for P04 parameters.
- Golden-output baseline; acceptance by functional verification on working copies.
- Automated VBA compile proof beyond L1/L2/L3 evidence.

## 12. Definition of Ready

**Ready for implementation planning and execution.**

- [x] PO decisions recorded for GA-1..GA-7.
- [x] Core behavioral contracts defined: snapshot-sized range (no `height/width` columns), strict values-only semantics, `paste_row` final-position/row-delete algorithm, stage ordering with template/report formatting precedence, fail-fast model, boolean parsing/defaults, template lookup/copy/naming rules.
- [x] Conceptual tasks, risks, assumptions, constraints, and verification obligations identified.
- [ ] L3 acceptance matrix defined/selected before implementation — **not required** (optional). Different L3 scenarios can be checked after implementation on working copies as needed (T7/OQ-V1).
- [ ] Doc updates drafted/applied as part of implementation (T7).

Implementation may proceed to planning/decomposition and execution. L1/L2 follow project VBA workflow for changed exports; L3 verification scenarios may be performed after implementation as needed. OQ-V2/OQ-V3 integration details should be resolved against current code behavior during implementation and documented.

## 13. Research traceability

- **Repo/architecture research (completed):** subreport flow, config/schema, snapshot/format/grouping behavior, insertion/size constraints, and integration risks analyzed from `modTreportConfig.bas`, `modTreportRun.bas`, `modTreportExtract.bas`, `modTreportOutput.bas`, `modTreportFiles.bas`, `modTreportPivot.bas`, and `docs/01-vba-reporting/*`.
- **VBA/Excel technique research (completed):** whole-sheet template copy from `ThisWorkbook`, block row insert, `PasteSpecial`/value-assignment modes for strict values-only while preserving target, shift-safe helper-row deletion sequencing, template-not-found handling, Excel 2021/Windows target.
- **Spec gap/risk audit (completed):** P04 mapped against current implementation; gray areas GA-1..GA-7 identified and resolved by PO decisions recorded in §4.
- **Follow-up research:** only implementation-time verification spikes on real templates/workbooks if OQ-V2/OQ-V3 reveal behavior not covered by GA decisions.

# Specification: Combine Multiple Subreports on One Worksheet and Worksheet Name Length Validation (P01)

- **Source problem:** `.ai/problems/P01.md` (titled in-file as "P06. Combine Multiple Subreports on One Worksheet and Validate Worksheet Names")
- **Status:** Analytical specification — **ready for implementation planning** (all Product Owner decisions D-01–D-06 and OQ-7 resolved; preferred approach selected)
- **Prepared by:** Analyst (specification only; no architecture/code changes)

## 1. Problem statement

The report-generation process must support combining multiple subreports that target the same final report worksheet within the same output workbook, and it must enforce Excel worksheet name length constraints for all generated worksheet names (including temporary names). Today each `(data_files row, report)` produces a separate final worksheet name in the output workbook and duplicate final names fail as collisions; there is no two-stage prepare-then-combine flow, no `gap_rows` handling, no per-subreport `report_rows_limit_general` check, no projected final Excel-dimension validation for a combined sheet, and worksheet name validation is not suffix-aware for generated temporary names (`.ai/problems/P01.md`; current behavior evidenced in `mko_treport/src/vbs/export/modTreportRun.bas`, `modTreportOutput.bas`, `modTreportExtract.bas`, `modTreportConfig.bas` and `docs/01-vba-reporting/{pipeline-and-dataflow,configuration-model,overview}.md`).

**Business goal (desired outcome):** allow users to define multiple `data_files` records that contribute to the same `report_sheet` in the same output workbook and receive one correctly assembled final worksheet (in original `data_files` order, with required content/formatting/grouping/row-state preservation and required spacing/limits), while preventing invalid or conflicting worksheet names with clear, actionable errors.

## 2. Confirmed requirements

### 2.1 Grouping and ordering
- **CR-01 (P01 §1.1, D-03):** In `data_files`, `report_sheet` specifies the final report worksheet name. When multiple `data_files` records within the **same output workbook** have the same `report_sheet` value, their subreports must be combined into one final worksheet. Grouping identity is the **final resolved and sanitized worksheet name** within the same output workbook.
  - Processing order for grouping/name resolution: (1) resolve dynamic placeholders and substitutions; (2) sanitize the resulting worksheet name using existing naming rules; (3) group subreports by the resulting name within the same output workbook.
  - Worksheet name comparisons must be **case-insensitive**, consistent with Excel naming rules.
- **CR-02 (P01 §1.1, §1.3, §3):** Subreports must appear vertically, top-to-bottom, in the **same order as their records in `data_files`**. Processing/combining must be performed separately for each output workbook (no mixing across output workbooks).
- **CR-03 (P01 §3):** Preserve the original `data_files` processing order.

### 2.2 Stage 1 — prepare individual subreports
- **CR-04 (P01 §1.2, §3):** Preserve the existing subreport preparation process as much as possible. Keep initial report generation independent from subsequent combining. Each subreport must first be **fully generated on a separate worksheet**, regardless of whether a template is used.
- **CR-05 (P01 §1.2, OQ-7):** Worksheet naming for preparation follows sequential suffixes `1..N-1` in original `data_files` order for a group containing N subreports:
  - The first subreport uses the **base** name without a suffix.
  - Subsequent subreports receive sequential numeric suffixes starting at `1`: `<base>1`, `<base>2`, etc.
  - Example (three subreports): `Sales`, `Sales1`, `Sales2`.
  - Use the **actual suffix length** when validating the maximum permitted base-name length.
- **CR-06 (P01 §1.2, D-06):** If a generated/planned name already exists in the output workbook or conflicts with another **planned** worksheet name, **stop with an error** identifying the conflicting worksheet name and ask the user to change the conflicting name.
  - **Never** automatically delete, overwrite, or rename an existing worksheet to resolve a naming conflict.
  - Check both existing worksheet names and all names planned for creation during the current run.
  - Preserve existing cleanup behavior for unused default worksheets **only when they do not conflict** with planned report names.
- **CR-07 (P01 §1.2):** For each output workbook, maintain a grouping structure containing: base `report_sheet` name, an ordered list of temporary worksheets belonging to that report, and the `gap_rows` value associated with each temporary worksheet. Keep isolated per output workbook.
- **CR-08 (P01 §1.3, D-01):** `gap_rows` is an **optional column in `data_files`**.
  - Accept non-negative integers (`>= 0`).
  - Blank or missing values default to `0`.
  - `gap_rows` specifies the number of empty rows before the corresponding subreport when combining worksheets.
  - The first subreport's `gap_rows` is ignored because it starts the final worksheet.

### 2.3 Stage 2 — combine prepared subreports
- **CR-09 (P01 §1.3):** After all individual subreports have been prepared, perform a **second pass** through the output workbooks. For each report group: base worksheet already contains the first subreport.
- **CR-10 (P01 §1.3):** Compute insertion using base worksheet used area:
  - `lastRow = UsedRange.Row + UsedRange.Rows.Count - 1`
  - `nextRow = lastRow + 1`
  - process temporary worksheets sequentially in original `data_files` order.
- **CR-11 (P01 §1.3):** For each temporary worksheet: get `UsedRange`; validate size using per-worksheet `report_rows_limit_general` (see CR-15); calculate `destinationRow = nextRow + gap_rows`; copy source `UsedRange` to base starting at `destinationRow`; preserve source range's starting column; **do not apply `paste_row` or `paste_col` again** (already applied during initial preparation); set `nextRow = destinationRow + sourceUsedRange.Rows.Count`.
- **CR-12 (P01 §1.4):** Transfer from temp to base: **cell values (not formulas)**, cell formatting including number formats, row grouping and outline levels, hidden row states, row heights.
- **CR-13 (P01 §1.4):** Do **not** transfer from temps: column widths, hidden column states, zoom, FreezePanes, SplitRow and SplitColumn. Base worksheet must retain its existing column layout and display settings.
- **CR-14 (P01 §1.4):** Grouping and hidden rows must be restored on the final worksheet using the **new row positions** of each subreport.

### 2.4 Limits and validation
- **CR-15 (P01 §1.5A, §3, D-02):** `report_rows_limit_general` is a **global parameter in the `general` configuration** (individual subreport limit).
  - It must be an **explicitly configured positive integer**.
  - No per-`data_files` overrides and no hardcoded fallback.
  - Apply independently to every prepared subreport worksheet, including the first (base) worksheet.
  - Before combining, validate each worksheet's own `UsedRange.Rows.Count` against `report_rows_limit_general`. If exceeded, stop with an error identifying: worksheet name, associated `report_sheet`, actual row count, configured `report_rows_limit_general` value.
  - Do **not** apply this limit to the total combined report.
  - Validate the final combined worksheet separately against Excel's hard row and column limits (CR-16).
- **CR-16 (P01 §1.5B, §3):** Independently validate the **projected dimensions of the final combined worksheet before each paste** against Excel hard limits (rows **1,048,576**, columns **16,384**). Calculate destination range boundaries using actual source `UsedRange` dimensions and destination coordinates. If exceeded, stop with a clear error containing: final `report_sheet` name, projected row or column boundary, applicable Excel limit. Do not attempt the paste when limits would be exceeded.
- **CR-17v (D-05, complements CR-10/CR-11/CR-15):** If an expected subreport has **no valid content range**, stop report generation with a contextual error identifying the source subreport, worksheet, and associated `report_sheet`. Do **not** silently skip empty subreports.
  - **Important:** a valid one-row report must **not** be treated as empty solely because `UsedRange.Rows.Count = 1`. Check whether the range actually contains data or formulas.

### 2.5 Cleanup
- **CR-18c (P01 §1.6, §3):** After all subreports belonging to a group have been successfully combined, delete their temporary worksheets. Final output workbook must contain base worksheets with combined reports, not temporary subreport worksheets. Do not delete temporary worksheets until their content has been successfully transferred.

### 2.6 Worksheet name length validation
- **CR-19n (P01 §2):** Excel worksheet names cannot exceed **31** characters. Validate all worksheet names generated during report preparation, including temporary names with numeric suffixes.
- **CR-20n (P01 §2):** Validate the original/base `report_sheet` name and every generated temporary worksheet name. If a name exceeds the limit, **stop before creating the affected worksheets**.
- **CR-21n (P01 §2, D-04):** On length violation (including when a generated suffix causes the worksheet name to exceed 31 characters):
  - Display the **effective base name after substitution and sanitization** without the numeric suffix.
  - Calculate the maximum permitted base-name length as `31 - suffix_length`.
  - Explicitly state the maximum permitted base-name length in the error message.
  - The raw template name may be included as additional diagnostic context when useful.

### 2.7 General/non-regression constraints
- **CR-22g (P01 §3):** Do not apply `paste_row`/`paste_col` offsets a second time during combine. Do not transfer worksheet-level display settings from temporary worksheets. Do not delete or overwrite existing worksheets when name conflicts occur. Avoid unnecessary changes to existing report-generation logic.
- **CR-23g (existing, must be preserved):** Per-output-workbook isolation and current output/save/cleanup safety model for produced artifacts; `data_files` order preservation; template copy-from-`ThisWorkbook` behavior and template immutability per use; existing `insert_rows`/`skip_header`/`only_values`/formatting/grouping/placeholder behaviors during Stage 1 preparation where applicable; fail-fast contextual diagnostics patterns (evidenced in `modTreportRun.bas`, `modTreportOutput.bas`, `modTreportExtract.bas`, docs).

## 3. Conceptual development tasks

Each task is independent at the conceptual/spec level but has implementation dependencies.

| ID | Purpose | Expected outcome | Dependencies | Decision alignment |
|---|---|---|---|---|
| T1 | **Configuration model extension** | Add and validate `gap_rows` in `data_files` (optional, integer >= 0, blank/missing => 0, first subreport ignored at combine) and `report_rows_limit_general` in `general` (explicitly configured positive integer, no overrides/fallback). Update config read/validation and type structures with contextual errors consistent with existing patterns. | — | D-01, D-02 |
| T2 | **Sheet naming, planning, and suffix-aware validation** | Implement deterministic grouping key derivation (resolve placeholders -> sanitize -> case-insensitive group within output workbook), base/temp name planning with suffixes `1..N-1` in original `data_files` order, suffix-aware 31-char validation (max base = `31 - suffix_length`, effective base name displayed, raw template as optional context), and collision checks against existing and all planned names before creating affected worksheets. | T1 (for config/context as needed) | D-03, D-04, D-06, OQ-7 |
| T3 | **Grouping state and two-stage orchestration** | Build per-output-workbook grouping structure (base name, ordered temp worksheets/members, `gap_rows` per temp), run Stage 1 to fully prepare each subreport on separate worksheets preserving existing prep, then Stage 2 combine pass in preserved order with per-output isolation. | T1, T2 | D-01, D-03, D-06, OQ-7, CR-04/CR-09 |
| T4 | **Combine/copy semantics and bounds/limit enforcement** | Compute `lastRow/nextRow/destinationRow` per P01 (first `gap_rows` ignored), validate valid content range (fail fast per D-05, do not reject valid one-row report solely by count), enforce per-subreport `report_rows_limit_general` incl. base, validate projected final Excel limits before each paste, and copy `UsedRange` blocks with required transfer profile and restored grouping/hidden rows at new positions without reapplying `paste_row/paste_col`. | T1, T3 | D-02, D-05, CR-10–CR-17v |
| T5 | **Cleanup, produced-registry, and lifecycle integration** | Delete temporary worksheets only after successful transfer per group; ensure final output contains only base/combined sheets (plus allowed non-conflicting defaults behavior as applicable); align produced-sheet registry and cleanup/verify/save guards with P01 lifecycle and D-06. | T3, T4 | D-06, CR-18c |
| T6 | **Documentation, verification support, and acceptance alignment** | Update governed docs for config/schema/behavior (`docs/01-vba-reporting/configuration-model.md`, `pipeline-and-dataflow.md`, and related docs if impacted), define acceptance scenarios from confirmed requirements/decisions, and complete executor-side checks (L1 build and other non-user checks as feasible) while recording user-performed final verification (L2 compile and run from workbook). | T1–T5 | All, execution directives |

## 4. Product Owner decisions (approved and resolved)

The following decisions are **approved** and are part of the confirmed specification baseline.

### D-01. `gap_rows` location and scope — **Approved: Option A**
- `gap_rows` is an optional column in `data_files`.
- Accept non-negative integers (`>= 0`).
- Blank or missing values default to `0`.
- `gap_rows` specifies the number of empty rows before the corresponding subreport when combining worksheets.
- The first subreport's `gap_rows` is ignored because it starts the final worksheet.

### D-02. `report_rows_limit_general` location and mandatoriness — **Approved: Option A**
- Keep `report_rows_limit_general` in the `general` configuration.
- It must be an explicitly configured positive integer.
- No per-`data_files` overrides and no hardcoded fallback.
- Apply the limit independently to every prepared subreport worksheet, including the first (base) worksheet.
- Do not apply this limit to the total combined report.
- Validate the final combined worksheet separately against Excel's hard row and column limits.

### D-03. Grouping identity for dynamic `report_sheet` — **Approved: Option A**
- Group by the final resolved and sanitized worksheet name.
- Processing order: (1) resolve dynamic placeholders and substitutions; (2) sanitize the resulting worksheet name using existing naming rules; (3) group subreports by the resulting name within the same output workbook.
- Worksheet name comparisons must be case-insensitive, consistent with Excel naming rules.

### D-04. Name displayed in worksheet length errors — **Approved: Option A**
- Display the effective base name after substitution and sanitization (without the numeric suffix).
- Calculate the maximum permitted base-name length as `31 - suffix_length`.
- Explicitly state the maximum permitted base-name length in the error message.
- The raw template name may be included as additional diagnostic context when useful.

### D-05. Empty or undefined `UsedRange` — **Approved: Option A**
- If an expected subreport has no valid content range, stop report generation with a contextual error identifying the source subreport, worksheet, and associated `report_sheet`.
- Do not silently skip empty subreports.
- Important: a valid one-row report must not be treated as empty solely because `UsedRange.Rows.Count = 1`. Check whether the range actually contains data or formulas.

### D-06. Conflicts with pre-existing/default worksheets — **Approved: Option A**
- Never overwrite, delete, or rename an existing worksheet to resolve a naming conflict.
- Check both existing worksheet names and all names planned for creation during the current run.
- Preserve existing cleanup behavior for unused default worksheets only when they do not conflict with planned report names.
- On conflict, stop with an error identifying the conflicting worksheet name.

### OQ-7. Temporary suffix sequence — **Approved/Confirmed**
- Use suffixes `1..N-1` for a group containing N subreports.
- The first subreport uses the base name without a suffix.
- Subsequent subreports receive sequential suffixes starting at `1`, assigned in original `data_files` order.
- Example for three subreports: `Sales`, `Sales1`, `Sales2`.
- Use the actual suffix length when validating the maximum permitted base-name length.

## 5. Research summary

### 5.1 Current architecture (evidence-based)
- **Orchestration and flow:** `modTreportRun.Create_report` creates one output workbook per report, maintains per-output `outUsedSheetNames(r)`, iterates `cfg.DataFiles` in order, opens sources read-only, and calls `ProcessDataFile` per data file. `ProcessDataFile` resolves per-row `report_sheet` max placeholders, parses/applies data-file filters once, then for each report applies report filters, fails on no-data (2410), substitutes+sanitizes sheet name, enforces per-output collision (2612) and registers name, prepares target sheet, copies snapshot, applies formatting and outline grouping, and returns. Save path (success only): guard no-produced-sheets (2630), `CleanupDefaultSheets`, `EnsureProducedSheetsPresent`, `SaveOutputWorkbook` (`modTreportRun.bas`; `docs/01-vba-reporting/pipeline-and-dataflow.md`).
- **Naming/collision/preparation:** `modTreportOutput.SanitizeSheetName` replaces `:\/?*[]` and `'`, fails on empty (2601) or length > 31 (2602). `PrepareTargetSheet` fails if target sheet exists (2634, no overwrite/delete), creates plain sheet or copies Worksheet template from `ThisWorkbook` (fresh per use), renames and makes visible; template source not modified (`modTreportOutput.bas`; `docs/01-vba-reporting/configuration-model.md`).
- **Snapshot/transfer:** `modTreportExtract.CopyPivotSnapshot` determines presented range (`TableRange1` fallback `TableRange2`), handles `skip_header`, computes placement/final range based on `insert_rows` and template scope (P09 vs legacy), and transfers via `PasteSnapshotData`: standard mode = values+number formats, formats, column widths, per-row row heights; `onlyValues` mode = values only to target preserving target formatting. Returns final inserted range used for formatting/grouping (`modTreportExtract.bas`; docs pipeline/config).
- **Config/schema:** `data_files` fields in code/docs include `data_file`, `data_sheet`, `report_sheet`, `filters`, `paste_row`, `paste_col`, `freeze_row`, `freeze_col`, `zoom`, `group_rows`, `template`, `only_values`, `insert_rows`, `skip_header`; `general` includes `input_path`, `output_path`, `delimeter`, `rows_collapsed`. No `gap_rows`, no `report_rows_limit_general`, no `UsedRange`-based combine logic, no Stage 2 (`modTreportConfig.bas`; `docs/01-vba-reporting/configuration-model.md`).

### 5.2 Gaps vs P01 (facts, now bounded by approved decisions)
- No config/validation/runtime handling for `gap_rows` (D-01) or `report_rows_limit_general` (D-02).
- No per-output-workbook grouping by final resolved+sanitized (case-insensitive) `report_sheet` for combining (D-03); duplicate final names currently fail (2612) instead of combining.
- No temp naming scheme with suffixes `1..N-1` (OQ-7), no planned-name set including all generated/existing names checked before creation (D-06), no suffix-aware 31-char validation/messages per D-04.
- No Stage 2 combine pass, no `UsedRange`-based append (`lastRow/nextRow/destinationRow`, first `gap_rows` ignored), no "valid content range" check per D-05, no per-subreport `report_rows_limit_general` enforcement, no Excel hard-limit projection before combine paste, no worksheet-to-worksheet block copy with P01 transfer profile (values-not-formulas, no column widths/hidden columns/display settings transfer, restore grouping/hidden rows at new positions).
- No temp lifecycle as specified (delete temps only after successful transfer; final output = combined base sheets only) integrated with current produced-sheet registry/cleanup guards.

### 5.3 Feasible implementation approaches and selected direction
- **Approach 1 (Selected, recommended): Two-stage orchestration in `modTreportRun` + dedicated combine helper in `modTreportExtract` + suffix-aware naming/planning helpers in `modTreportOutput` (with `modTreportConfig` extensions).**
  - **Stage 1 (planning + preparation):** For each output workbook, build a deterministic plan in `data_files` order: resolve group key per D-03 (resolve placeholders -> sanitize -> case-insensitive), compute base and sequential temp names per OQ-7, perform suffix-aware length validation per D-04 and planned/existing collision checks per D-06 before creating affected sheets, maintain per-output grouping state `(baseName, ordered temp entries, gap_rows per temp)` per D-01. Generate each subreport on its own sheet using existing preparation pipeline (`PrepareTargetSheet`, `CopyPivotSnapshot`, `ApplySheetFormatting`, `ApplyOutlineGrouping`) so templates/`insert_rows`/`skip_header`/`only_values`/placeholders behave as today. Enforce valid-content-range and per-subreport limit checks per D-05/D-02 as specified.
  - **Stage 2 (combine):** Per output workbook/group in preserved order, append each temp `UsedRange` to base using P01 formulas and `destinationRow = nextRow + gap_rows` (ignore first subreport `gap_rows`); before each paste validate projected Excel row/column bounds; copy with P01 transfer profile (values not formulas, formats incl number formats, row heights, hidden row states, row grouping/outline levels restored at new positions; no column widths/hidden columns/zoom/freeze/split transfer; preserve base layout/display; no `paste_row/paste_col` reapplication). Delete temp sheets only after successful transfer of the group; update produced/final registry to base sheets for existing cleanup/verify/save guards.
  - **Rationale:** Preserves existing Stage 1 prep behavior and module responsibilities, isolates combine as a new concern, and aligns with P01 and approved decisions D-01–D-06/OQ-7.
- **Approach 2 (Rejected):** In-place direct generation into base with offset math (no real temp worksheets). Conflicts with explicit "fully generated on a separate worksheet first" and independent-combine intent; high regression risk.
- **Approach 3 (Rejected as primary):** Broad sheet/range copy then post-process. High risk of violating exact transfer inclusions/exclusions and disturbing base layout.

### 5.4 Relevant Excel/VBA best-practice notes (for implementation planning)
- Avoid `Select`/`Activate` in combine loops; operate on explicit `Workbook/Worksheet/Range` objects.
- Prefer block/value-array assignment for values (`dest.Value = src.Value`) and minimal `Copy/PasteSpecial` operations for formats; explicitly avoid `xlPasteColumnWidths` for combine.
- Preserve performance/stability with application-state controls and deterministic cleanup of `CutCopyMode`.
- Treat `UsedRange` as specified by P01 but implement D-05 content-validity checks (do not reject valid one-row reports solely by `UsedRange.Rows.Count = 1`; check for actual data/formulas). Validate Excel row/column boundaries **before** paste.
- Recreating outline/grouping and hidden rows at offset positions requires careful row/range-mapped handling (structural outline levels/groups); complex nested outlines/merged cells are higher-risk cases requiring L3 verification.
- Perform all suffix/planned-name/existence checks **before** creating affected worksheets; use per-output-workbook reserved-name sets with case-insensitive comparison.

## 6. Assumptions

- **A-01:** P01 (`.ai/problems/P01.md`) is the authoritative business/request specification for this change despite the in-file heading text; references to sections map to that file.
- **A-02 (resolved by D-03):** Grouping key is the final resolved and sanitized worksheet name within the same output workbook, compared case-insensitively.
- **A-03 (resolved by D-01/D-02):** `gap_rows` is added as an optional `data_files` column (integer >= 0, blank/missing => 0, first ignored at combine). `report_rows_limit_general` is added as an explicitly configured positive-integer parameter in `general` (no overrides/fallback), applied per prepared subreport incl. base.
- **A-04:** `UsedRange`-based placement follows P01 formulas; D-05 defines "valid content range" (not a blind `Rows.Count = 1` rejection) and requires fail-fast on no valid content range.
- **A-05:** "Stop before creating the affected worksheets" means no base/temp worksheet for the offending planned creation is created; planning prevents partial conflicting state where feasible.
- **A-06:** Existing Stage 1 preparation semantics (templates, `only_values`, `insert_rows` including template/legacy modes, `skip_header`, formatting precedence, config-driven outline grouping, placeholder resolution rules) remain the source of truth unless explicitly overridden by P01/D-01–D-06/OQ-7.
- **A-07:** VBA source-of-truth remains `mko_treport/src/vbs/export/*`; executor-side and user-side verification responsibilities follow the approved execution directives in §7.
- **A-08 (environment):** For development, it is **not required** to have an already prepared template with the needed parameters, although such a template/workbook already exists and required parameters are entered. Development should proceed against code/config/spec and existing repository artifacts without being blocked by template-preparation prerequisites.

## 7. Constraints and execution/verification directives (approved)

### 7.1 Product/technical constraints
- **C-01:** Excel worksheet name max length is 31 characters; Excel worksheet hard limits are 1,048,576 rows and 16,384 columns (P01 §1.5B, §2).
- **C-02:** No auto delete/overwrite/rename of existing worksheets on name conflicts (D-06). Do not reapply `paste_row/paste_col` in combine. Do not transfer forbidden temp worksheet properties/settings. Do not delete temps before successful transfer (P01).
- **C-03:** `report_rows_limit_general` applies to each individual prepared subreport only (not combined/sum); projected combined worksheet must be validated before each paste (D-02, P01 §1.5).
- **C-04:** Grouping/combining must be isolated per output workbook and preserve original `data_files` order (P01 §1.1–§1.3, §3, D-03, OQ-7).
- **C-05:** Current architecture constraints: `Create_report` creates per-report output workbooks, preparation uses `PrepareTargetSheet`/`CopyPivotSnapshot`/formatting/grouping pipeline, and end-of-run cleanup/save is guarded by produced-sheet registries (`modTreportRun.bas`, `modTreportOutput.bas`, docs).

### 7.2 Execution and verification directives (separately stated)
- **EVD-01 (template prerequisite):** For development it is **not required** to have an already prepared template with the needed parameters (although such a template already exists and everything is entered). Lack of a separately prepared "ready template" must not block implementation.
- **EVD-02 (verification ownership):** **Final verification — compilation and run exactly from the workbook — will be done by the user.** The executor performs the remaining checks (feasible automated/scripted/static checks, L1 build/report review, and other non-user verifications as applicable) and must clearly report what was and was not verified.
- **EVD-03 (continuity):** **Development must not stop until the code is implemented.** Implementation work should proceed to completion of the required code changes according to this specification and approved decisions, subject to the verification split in EVD-02 (i.e., user-performed workbook compile/run is the final gate but not a reason to halt code implementation).

## 8. Risks

- **R-01 (High): Outline/grouping transfer fidelity.** P01 requires transferring/restoring row grouping and outline levels at new positions; Excel outlines (especially nested/multi-level or pre-existing complex structures) are non-trivial to reconstruct exactly in VBA. **Mitigation:** targeted block-copy logic with explicit row/level mapping; L3 scenarios by user (and executor-prepared test cases/instructions where feasible).
- **R-02 (Medium): `UsedRange` accuracy/inflation and edge cases.** `UsedRange` may include stale/formatted cells or merged-cell complexities; D-05 requires distinguishing "no valid content range" from a valid one-row report (check actual data/formulas). **Mitigation:** follow P01 formulas, implement D-05 validity checks, and produce contextual diagnostics.
- **R-03 (Medium): Naming/grouping correctness under substitution/sanitization.** Grouping and length-error messaging depend on final resolved+sanitized names (D-03/D-04); case-insensitive comparison is required. **Mitigation:** deterministic resolve->sanitize->compare pipeline and explicit message content including effective base name and max base length.
- **R-04 (Medium): Interaction with existing prep modes.** `skip_header`, `insert_rows` (legacy vs template P09), `only_values`, templates, and config-driven `ApplyOutlineGrouping` affect prepared content/layout; combine must copy final prepared `UsedRange` semantics correctly without breaking base layout/display.
- **R-05 (Medium): Large data/performance.** Many/larger subreports increase copy time and memory pressure. **Mitigation:** block/value transfers, minimal paste operations, existing fast app-state controls.
- **R-06 (Low-Medium): Lifecycle/cleanup regressions.** Temps and deferred deletion must integrate with produced-sheet registry/cleanup guards (D-06, CR-18c). **Mitigation:** clear separation of planned/temp/final-base registries and delete temps only after successful transfer.
- **R-07 (Verification risk):** Runtime behavior (Excel COM, outline/hidden-row fidelity, exact UsedRange/content validity results) cannot be fully proven without user-performed workbook compile/run (EVD-02). **Mitigation:** executor performs L1 and other feasible checks; user performs final L2 compile and run from workbook.

## 9. Open questions

No open Product Owner questions remain. **D-01, D-02, D-03, D-04, D-05, D-06, and OQ-7 are approved and resolved** (see §4).

## 10. Out of scope

- Changes to `Refresh_data` workflow or source-data refresh behavior.
- Changes to existing filter semantics, placeholder resolution rules, `insert_rows`/`skip_header`/`only_values`/template preparation semantics except as strictly required to support P01 combine/validation and confirmed requirements/decisions.
- Reapplying or redesigning `paste_row`/`paste_col`, worksheet display/layout settings transfer from temps, or automatic conflict resolution (rename/overwrite/delete) beyond what P01/D-06 require.
- Applying `report_rows_limit_general` to combined totals/sums.
- Per-`data_files` overrides or hardcoded fallback for `report_rows_limit_general` (D-02).
- Performance optimization beyond what is necessary for correct implementation.
- New report formats, new output file types, or unrelated configuration/schema changes.

## 11. Definition of Ready

The specification is **ready for implementation planning and implementation execution**.

- **DoR-01 (satisfied):** P01 confirmed requirements (CR-01..CR-23g, including D-01–D-06/OQ-7 refinements) are accepted as the scope baseline.
- **DoR-02 (satisfied):** Product Owner decisions D-01–D-06 and OQ-7 are resolved and recorded in §4.
- **DoR-03 (satisfied):** Configuration contract is finalized: `gap_rows` in `data_files` (optional, integer >= 0, blank/missing => 0, first ignored at combine); `report_rows_limit_general` in `general` (explicitly configured positive integer, no overrides/fallback, per prepared subreport incl. base); grouping key = final resolved+sanitized case-insensitive name (D-03); length-error display = effective base name + `31 - suffix_length` with raw template optional context (D-04).
- **DoR-04 (satisfied):** Implementation approach selected: **Approach 1** with module responsibilities — `modTreportRun` (orchestration/grouping state/two-stage flow/lifecycle), `modTreportOutput` (naming/planning/suffix-aware validation/collision helpers, worksheet prep integration), `modTreportExtract` (combine/copy helper with P01 transfer profile and bounds checks), `modTreportConfig` (schema/validation for new parameters) — and integration with existing prep/cleanup flows.
- **DoR-05 (defined):** Acceptance scenarios must cover, at minimum: multi-subreport same final `report_sheet` combine in order with `gap_rows` (first ignored); base `<name>` and temps `<name>1..N-1`; planned/existing name collisions (D-06) with no overwrite/delete/rename; suffix-aware 31-char errors showing effective base name and max `31 - suffix_length` (including multi-digit suffix case); per-subreport `report_rows_limit_general` pass/fail (base and temps) and D-05 valid-content-range handling (including valid one-row report not rejected solely by `Rows.Count = 1`); projected Excel row/column limit fail before paste; transfer inclusions/exclusions and restored grouping/hidden rows at new positions with base layout/display preserved; no temp sheets left after successful combine; no `paste_row/paste_col` reapplication in combine; preservation of existing single-subreport/non-combined behaviors.
- **DoR-06 (defined with ownership):** Executor performs remaining feasible checks (including L1 build/report review and other non-user verifications) and reports results/limitations. **User performs final verification — compilation and run exactly from the workbook** (L2/L3 in workbook context). Any `-SyncExports`/`-DeployToDataTheiss` actions remain explicit opt-in only per project rules.
- **DoR-07 (satisfied/managed):** Risks R-01–R-07 recorded with mitigations/verification notes; D-05 resolves the prior empty-`UsedRange` ambiguity for specification purposes (implementation must apply the stated validity check).
- **DoR-08 (identified):** Governed documentation update scope includes `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`, and related docs (e.g. `overview.md`) as impacted by new config/behavior.

## 12. Recommended next step

Produce an implementation-ready technical plan/decomposition and proceed with implementation based on **Approach 1** and confirmed decisions D-01–D-06/OQ-7. **Development must not stop until the code is implemented** (EVD-03). The executor should complete all non-user checks (EVD-02) and provide clear verification/reporting, while the user performs final compilation and run exactly from the workbook. No separately prepared template with needed parameters is required to start or continue development (EVD-01).

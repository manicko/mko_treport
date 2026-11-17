# 006 — P07: `data_files.skip_header` (copy pivot without header)

- **Spec ID:** 006_p07_skip_header_data_files
- **Status:** Draft — PO decisions confirmed (all recommended). Ready for implementation planning.
- **Date:** 2026-11-17
- **Source request:** `.ai/problems/P07.md`
- **Related artifacts:** `.ai/problems/001_treport_spec.md` (FR-RPT-3/D15, FR-CFG-4/6, D13/D14), `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`, `mko_treport/src/vbs/export/modTreportConfig.bas`, `mko_treport/src/vbs/export/modTreportExtract.bas`, `mko_treport/src/vbs/export/modTreportRun.bas`, `mko_treport/src/vbs/export/modTreportOutput.bas`

## 1. Problem statement

`treport.xlsm` copies a pivot presented range (currently via `GetPresentedRange` preferring `PivotTable.TableRange1`, fallback `TableRange2`) into report sheets (`FR-RPT-3`, D15). P07 requests: when `data_files.skip_header` is `TRUE`/`1`/`YES`, copy the pivot summary **without Header**.

P07 provides illustrative pseudo-code (not authoritative): `src = pt.TableRange1; body = pt.DataBodyRange; headerRows = body.Row - src.Row`, then copy without header rows. The implementation must fit current architecture and preserve determinism, fail-fast behavior, and existing contracts except where explicitly revised by PO decisions below.

## 2. Business goal

- Allow report generation to omit pivot header rows per `data_files` row so target sheets receive summary/body content in the required form without manual post-editing.
- Keep behavior configurable, deterministic, and backward-compatible for existing configurations.

## 3. Scope

### 3.1 In scope (conceptual)
- Add `data_files.skip_header` configuration support (schema, parsing, validation, defaults, storage).
- Implement deterministic "copy without header" semantics for the presented pivot snapshot per confirmed PO decisions.
- Preserve/fail-fast interactions with `paste_row/paste_col`, `template`, `only_values`, `insert_rows`, formatting, bounds-scoped row grouping, placeholders, and output naming/idempotency as applicable.
- Update governed/spec documentation (`001_treport_spec.md` deltas as needed, `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`) and produce verification evidence (L1/L2/L3 per project rules) when implemented.

### 3.2 Out of scope
- Changing source pivot layouts/data, Power Query, or business meaning of metrics/categories beyond this copy behavior.
- Redesigning VBA module architecture; runtime dependencies outside Excel VBA.
- Implementing `group_rows` feature scope beyond existing behavior (grouping remains bounds-scoped on final inserted range as today).
- Broadening support to all pivot layouts/options beyond the agreed initial verification scope (see D-P07-5).
- Unrelated report/filter/placeholder changes.

## 4. Constraints

- **Runtime:** final solution VBA-only in Excel; auxiliary build/test tools non-runtime (project/001 constraints).
- **Contract baseline:** FR-RPT-3/D15 define default presented-data copy; amended for `skip_header=TRUE` by D-P07-1/D-P07-2.
- **Pipeline invariants:** snapshot size drives `only_values` transfer, `insert_rows` prep/helper deletion/final range, formatting target, and Outline grouping bounds (`modTreportExtract.CopyPivotSnapshot`, `modTreportRun.ProcessDataFile`, pipeline docs). Any skip must keep these consistent using sliced dimensions.
- **Config patterns:** optional boolean-like columns in `data_files` follow tolerant defaults and strict validation for unknown non-empty values (aligned with `modTreportConfig.ReadP04BooleanValue` for `only_values/insert_rows`).
- **Determinism & fail-fast:** invalid/unusable geometry, invalid config values, or impossible skip conditions must fail fast with contextual diagnostics (operation, data_file, data_sheet, report context as applicable, parameter, raw/value details as applicable). No silent fallback.
- **VBA naming:** identifiers must not collide with VBA reserved words/case-insensitive keywords.

## 5. Confirmed facts (evidence)

| Area | Fact | Evidence |
|---|---|---|
| Request | `skip_header = TRUE or 1 or YES` in `data_files`; copy pivot summary without Header; pseudo-code uses `TableRange1`/`DataBodyRange` row offset. | `.ai/problems/P07.md` |
| Presented range | `modTreportExtract.GetPresentedRange` returns `pt.TableRange1`, fallback `pt.TableRange2`, fail **2701** if unavailable. | `modTreportExtract.bas:4-14` |
| Snapshot/copy | `CopyPivotSnapshot` derives `h,w` from presented range; supports `only_values` and `insert_rows` (`H` from snapshot); pastes/copies full `h x w`; deletes helpers; returns final/destination range sized from `h,w`. | `modTreportExtract.bas:16-127` |
| Downstream use | `ProcessDataFile` passes `onlyValues/insertRows`, uses returned `insertedRng` for formatting and `ApplyOutlineGrouping` on final bounds. | `modTreportRun.bas:201-214`, `modTreportExtract.bas:173-326` |
| Config model today | `TreportDataFile` has no `skipHeader`; `data_files` supports optional `template/only_values/insert_rows` with defaults and strict bool parsing; docs list no `skip_header`. | `modTreportConfig.bas:20-35,126-242,284-336`, `configuration-model.md:49-66` |
| Current contract | Copy presented data after filters incl. row/column labels and grand totals as shown; exclude pivot object/page-filter area and PQ connections. | `.ai/problems/001_treport_spec.md:104,143`, `pipeline-and-dataflow.md:86-91` |
| `skip_header` not implemented | No existing `skip_header` contract/usage in repo outside P07. | repo scan (research) |
| Pivot semantics (research) | `headerRows = DataBodyRange.Row - TableRange1.Row` measures rows above value body inside `TableRange1` (typically column/field header band) but is layout-dependent; `PageRange` is outside `TableRange1`; `DataBodyRange` alone drops labels; no universal Excel "header" range. | R2 findings (Microsoft VBA/Excel docs + analysis) |

## 6. Product Owner decisions (confirmed, 2026-11-17)

All recommended options accepted.

| ID | Decision | Confirmed choice |
|---|---|---|
| D-P07-1 (Q1) | Definition of "Header" when `skip_header=TRUE` | **Column-header band only.** Remove the top rows above the data/value body within the presented range (column field captions/header band). **Preserve:** row item labels, data values, and subtotals/grand totals as shown in the remaining rows of the sliced presented range. Page fields/report filters remain governed by presented-range selection (not inside `TableRange1`; see D-P07-3). Exact preserved/removed elements and supported cases must be documented in config/pipeline docs and validated per D-P07-5. |
| D-P07-2 (Q2) | FR-RPT-3/D15 amendment | **Explicit exception.** Default FR-RPT-3/D15 unchanged; when `skip_header=TRUE`, omission of the agreed header rows per D-P07-1 is allowed. The copy must still preserve row labels and data values and subtotals/grand totals as applicable/rendered in the resulting slice per D-P07-1. Spec/docs must state this exception precisely. |
| D-P07-3 (Q3) | Source range basis for slicing | **`GetPresentedRange` result.** Use the same range object selected today (`TableRange1`, fallback `TableRange2`) as the basis for header-offset computation and slicing, preserving current extraction/fidelity semantics. |
| D-P07-4 (Q4) | Edge-case / fail-fast semantics | **Strict fail-fast** when `skip_header=TRUE` and: presented range/body geometry is invalid or unavailable for deterministic slicing, computed header offset `< 0`, or computed header offset `>= snapshot row count` (would produce empty/invalid snapshot). **`offset == 0` => no-op** (copy as today). Errors must be contextual and deterministic; no silent fallback. |
| D-P07-5 (Q5) | Supported layouts/options (initial scope) | **Layouts used by current treport data sources only** for initial delivery. Behavior must be verified against actual `data/theiss` pivots/layouts in scope with a documented L3 matrix. Layouts/options not covered by this matrix are out of initial verification scope; if encountered and not deterministically supported by implementation guards, behavior must fail fast with clear context rather than produce ambiguous output. |
| D-P07-6 (Q6) | Configuration contract | **Optional `data_files` column `skip_header`**, default `False` when absent/blank/null. Accepted tokens (trimmed, case-insensitive): `TRUE`/`1`/`YES` => True, `FALSE`/`0`/`NO` => False (Excel Boolean supported). Any other present non-empty/error value => **fail-fast config validation** with contextual diagnostics (row, `data_file`, `data_sheet`, parameter, raw value, expectation), aligned with existing boolean parsing. |

## 7. Confirmed functional requirements (conceptual)

- **FR-CFG-P07-1:** Extend `data_files` schema/`TreportDataFile` with `skip_header` per D-P07-6 (backward compatible). Read/validate in config loading before run side effects.
- **FR-CFG-P07-2:** Update `001_treport_spec.md` (FR-CFG-4/FR-CFG-6 as applicable), `docs/01-vba-reporting/configuration-model.md`, and `docs/01-vba-reporting/pipeline-and-dataflow.md` to document column, defaults, tokens, validation, and `skip_header` semantics.
- **FR-RPT-P07-1:** When `skip_header=False` (default), snapshot extraction/copy behavior remains unchanged.
- **FR-RPT-P07-2:** When `skip_header=True`, compute header offset from the presented range (`GetPresentedRange` result, D-P07-3) and slice to remove the column-header band per D-P07-1/D-P07-2. Sliced source range/dimensions must be used consistently for transfer mode (`only_values`), row-height/width handling as applicable, `insert_rows` prep count, helper-row deletion, and returned final inserted range so downstream formatting and bounds-scoped grouping operate on correct bounds.
- **FR-RPT-P07-3:** Apply D-P07-4 fail-fast rules with contextual diagnostics; `offset == 0` is no-op.
- **FR-RPT-P07-4:** `skip_header` is independent per `data_files` row and compatible with existing per-row/report processing order, placeholders, filters, template/`only_values`/`insert_rows` options, and output naming/idempotency, subject to documented semantics and L3 verification matrix (D-P07-5).
- **FR-DEP/VFY-P07:** Implementation delivery must follow project VBA workflow: L1 build/import report on working copy, L2 manual VBE compile for any compile claim, L3 functional checks on copies for logic/behavior changes (no production/original data corruption).

## 8. Conceptual development tasks

| ID | Purpose | Expected outcome | Dependencies |
|---|---|---|---|
| T1 | Configuration contract & validation | `data_files.skip_header` parsed/stored/validated per D-P07-6; defaults/backward compatibility; contextual fail-fast errors; schema/docs deltas identified. | D-P07-6 |
| T2 | Snapshot slicing & extraction | Deterministic header-offset computation and slicing per D-P07-1/2/3/4; sliced `H/W` propagated through `CopyPivotSnapshot`/`PasteSnapshotData`/`DeleteInsertRowsHelpers` and final range; `offset==0` no-op; fail-fast paths covered. | T1, D-P07-1–4 |
| T3 | Pipeline integration & downstream invariants | `modTreportRun` passes flag; final range correctness for formatting and bounds-scoped `ApplyOutlineGrouping`; no regressions when `skip_header=False`; interactions with `template/only_values/insert_rows` preserved per invariants. | T2 |
| T4 | Documentation updates | `001_treport_spec.md` (FR-CFG-4/6, FR-RPT-3/D15 exception notes as needed), `configuration-model.md`, `pipeline-and-dataflow.md` updated to reflect D-P07-1–6 and verification scope. | T1–T3 (final wording aligned to implementation) |
| T5 | Verification & deployment evidence | L1 build/import on working copy with report; L2 manual VBE Compile recorded (path/time); L3 functional checks on copies covering current `data/theiss` layouts/cases, `skip_header` on/off, and key option combinations (`only_values`, `insert_rows`, template if used) per D-P07-5 matrix. | T1–T4 |

## 9. Research summary

- **R1 — Current architecture (completed, actionable):** `GetPresentedRange` selects `TableRange1`/`TableRange2` fallback (2701 on failure). `CopyPivotSnapshot` owns `h,w` and all dependent operations: paste/transfer, `insert_rows` insert at `P+1`/helper deletes/final range `P..P+H-1`, and returned range consumed by formatting and bounds-scoped grouping. Config uses optional-column + strict boolean parsing pattern (`ReadP04BooleanValue`, 2118-style context). **Preferred approach:** localized slice in snapshot-copy path using sliced dimensions (Approach A) to preserve invariants; run-side pass-through only.
- **R2 — Excel PivotTable semantics (completed, actionable):** `TableRange1` excludes page fields; `TableRange2` includes them; `DataBodyRange` is value grid (labels live in row/column/page areas). `headerRows = DataBodyRange.Row - TableRange1.Row` removes top band above value body (layout-dependent), not a universal "header"; may affect field-caption rows and does not capture page area. `DataBodyRange`-only drops labels. Deterministic VBA behavior requires explicit preserved/removed contract, range-basis choice, and `Nothing`/bounds fail-fast checks. PO decisions D-P07-1–5 resolve the required ambiguities for initial scope.

**Review:** R1/R2 are sufficient and support the confirmed decisions. No additional Researcher tasks required for implementation planning unless new data sources/layouts are introduced outside D-P07-5.

## 10. Assumptions

- **A1:** Error code(s) for new `skip_header` geometry/offset validation failures will follow existing extract/validation numeric conventions and include contextual diagnostics (implementation detail to align during T2).
- **A2:** For `offset == 0` (no-op), no warning is required unless implementation adds a diagnostic consistent with existing style (default: silent no-op, behavior unchanged).
- **A3:** "Subtotals/grand totals as shown in the remaining rows" (D-P07-1) means totals rendered within the sliced presented range after removing the column-header band; exact rendering is pivot/layout-dependent and validated per D-P07-5 matrix.

## 11. Constraints (summary)

- VBA-only runtime; no external runtime dependencies.
- Backward compatibility when `skip_header` absent/False.
- Deterministic processing and contextual fail-fast (no silent fallback/ambiguous output).
- L1/L2/L3 evidence required for implementation claims per project workflow.
- Initial layout support/verification limited to D-P07-5.

## 12. Risks

- **Layout sensitivity (medium):** pivot geometry varies by layout/options; initial scope limited to D-P07-5 and guarded by fail-fast on unsupported/invalid geometry to avoid ambiguous output.
- **Regression risk (medium):** incorrect propagation of sliced dimensions can break `insert_rows`, final bounds, or row grouping; mitigated by consistent single-source sizing in extract path and L3 matrix.
- **Template/format alignment (low–medium):** removing header rows can shift content relative to template/target formatting expectations; validate in L3 combinations.
- **Verification/environment risk (medium):** compile/behavior claims require L1/L2/L3 on working copies with representative `data/theiss` data; runtime not provable by static analysis alone.

## 13. Open questions

- **OQ-1 (implementation detail):** specific error code numbers/messages for `skip_header` invalid geometry/offset cases, to be aligned with existing patterns during T2/T4 (non-blocking for planning).
- **OQ-2 (operational):** whether any existing target workbook `data_files` configurations need migration notes once column is introduced (expected none if optional/default `False`; confirm during rollout).

## 14. Out of scope (confirmed)

- Physical merge of VBA modules or runtime external tooling.
- Changes to PQ/pivot source definitions, unrelated filter/placeholder semantics, or new report types.
- Golden-workbook acceptance baseline (functional verification against actual source data per 001/D18).
- Support/verification for pivot layouts/options outside D-P07-5 initial matrix.

## 15. Definition of Ready

**Ready for implementation planning.**

- [x] PO decisions Q1–Q6 confirmed (all recommended, D-P07-1–D-P07-6).
- [x] "Header" removal/preservation and FR-RPT-3/D15 exception defined (D-P07-1, D-P07-2).
- [x] Range basis, fail-fast/no-op rules, and config contract defined (D-P07-3, D-P07-4, D-P07-6).
- [x] Initial supported-layout scope and verification requirement defined (D-P07-5).
- [x] Conceptual tasks T1–T5 accepted as decomposition baseline.
- [x] Research (R1, R2) reviewed and sufficient; no additional Researcher tasks required at this stage.

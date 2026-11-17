# Execution Plan: P01 — max-offset, default sheet removal, row grouping

- **Plan file:** `.ai/plans/005-p01-max-offset-default-sheet-grouping.md`
- **Plan status:** Draft — decomposed for dependency-safe execution (B0 baseline/research collected 2026-11-16)
- **Date:** 2026-11-16
- **Scope mode:** Minimal, incremental, semantic (no speculative redesign)

## 1. Source references (authoritative)

| Ref | Path / artifact | Role |
|---|---|---|
| S1 | `.ai/problems/003_p01_max_offset_default_sheet_grouping_spec.md` | Requirement baseline and PO decisions (CR-1..CR-6, D1–D7, OQ-1–OQ-3) |
| S2 | `.ai/problems/P01.md` | Original request (as cited by S1) |
| S3 | Step-0 Auditor report (session task) | Current-state evidence for affected VBA exports, gaps, and constraints |
| S4 | B0 Auditor report (session task, 2026-11-16) | Baseline confirmation, call-path facts, decision-gate checklist (a)–(g) |
| S5 | B0 Researcher memo (session task, 2026-11-16) | Excel/VBA/COM behavior/options for uncertainties (no final choices) |
| S6 | `mko_treport/src/vbs/export/modTreportRun.bas` (`Create_report`, `ProcessDataFile`, `ResolveReportSheetPlaceholderMaxValues`) | Run orchestration, output workbook lifecycle, pre-filter placeholder resolution, paste/format sequence |
| S7 | `mko_treport/src/vbs/export/modTreportOutput.bas` (`AddReportSheet`, `SaveOutputWorkbook`, `GetOrCreateOutputWorkbook`, `CleanupDefaultSheets`, placeholder/name helpers) | Output workbook/sheet creation, naming/sanitization, placeholder parsing/substitution |
| S8 | `mko_treport/src/vbs/export/modTreportFilters.bas` (`ParseAndClassify`, `ClassifyExpr`, `IsBoundToken`, `CountValidRangeSplits`, `ResolveFilterContext`, `ResolveGlobalBound`, `BuildDesiredValues`, `Apply*`) | Filter DSL, bound/range parsing, min/max resolution, selection semantics |
| S9 | `mko_treport/src/vbs/export/modTreportPivot.bas` (`GetPivotFieldNumericMaxValue`, numeric parsing/domain helpers) | Named-field numeric domain/max extraction and domain helpers |
| S10 | `mko_treport/src/vbs/export/modTreportConfig.bas` (`ReadConfig`, `TreportGeneral`, `TreportDataFile`) | Config schema/types, `general` keys, `data_files.group_rows` storage |
| S11 | `mko_treport/src/vbs/export/modTreportExtract.bas` (`GetPresentedRange`, `CopyPivotSnapshot`, `ApplySheetFormatting`) | Presented/pasted range source and size derivation, freeze/zoom/gridlines |
| S12 | `mko_treport/src/vbs/export/modTreportState.bas` | State capture/restore and error cleanup context |
| S13 | `mko_treport/src/scripts/build_treport_vba.ps1`, `docs/02-build-and-deployment/vba-build.md` | L1 sync/build workflow and report artifacts |
| S14 | `docs/01-vba-reporting/configuration-model.md`, `filters-dsl.md`, `pipeline-and-dataflow.md`, `overview.md` | User/contract documentation impacted by P01 |

**Evidence convention:** code references use semantic units (modules/procedures/types/ListObjects/keys). Prior audit line references are context-only.

## 2. Overview

P01 introduces: (1) deterministic removal of Excel-default empty sheets from normally produced output workbooks without deleting produced report sheets (CR-1, D3, D7), (2) `max-N` offset semantics for `data_files.filters` bounds/values and `report_sheet` `<max-N_<field>>` placeholders with shared validity/domain rules and filter-independence for placeholders (CR-2, CR-3, D1, D2, OQ-1/OQ-2), and (3) per-subreport, bounds-scoped Excel Outline row grouping driven by `data_files.group_rows` and fill rules with silent skip on invalid params and `general.rows_collapsed` readiness/default-expanded behavior (CR-4, CR-5, D4, D5, D6).

**B0 baseline (confirmed):** Exports do **not** satisfy CR-1..CR-5. `CleanupDefaultSheets` is a stub and never called; workbooks created via `Application.Workbooks.Add(xlWBATWorksheet)` with report sheets added after. `modTreportFilters` has no offset grammar (`-` treated as range separator candidate in `CountValidRangeSplits`/bound-token validation). `modTreportOutput` only supports `<max_<field>>` with legacy/remnant rejection and row-scoped pre-filter resolution in `modTreportRun.ResolveReportSheetPlaceholderMaxValues` via `modTreportPivot.GetPivotFieldNumericMaxValue`. `TreportDataFile.GroupRows` is read but runtime no-op; `TreportGeneral` parses only `input_path/output_path/delimeter` and `rows_collapsed` is absent. Inserted bounds are available at runtime in `CopyPivotSnapshot` (`srcRg` rows/cols at paste anchor) but not returned to caller.

## 3. Assumptions and constraints carried over

- **Baseline readiness:** Spec analytically ready; no open business questions (S1 §4 D1–D7, §11). OQ-1..OQ-3 are technical/verification clarifications resolved per block with evidence.
- **Runtime/scope:** Final `Refresh_data`/`Create_report` behavior remains VBA-only in Excel. Applies generally to all applicable flows/cases (CR-6, D7).
- **Source discipline:** Edit only VBA export sources `mko_treport/src/vbs/export/*.bas,*.cls` as needed; do not modify canonical workbooks blindly; never claim compile without L2; use working copies/backups for verification.
- **Verification gates:** L1 build/sync report required when exports change; **L2 manual `Debug → Compile VBAProject` in Excel on a working copy is required to claim compilation**; L3 functional checks on copies only when behavior changes.
- **Determinism/fail-fast:** Invalid/unresolvable filter or placeholder offsets/values fail fast with clear context (field, resolved max, N, computed value where applicable) per D1/D2. Invalid/missing/out-of-width `group_rows` => **silent skip for that subreport only** (D6). Default-sheet edge cases align with existing fail-fast/no-data behavior (D3).
- **Independence/bounds:** Placeholder resolution remains independent of `data_files.filters` application/selection/derivation and pre-filter timing preserved (CR-3). Grouping is per-subreport, bounded to inserted/pasted range, contiguous groupable runs only (CR-4, D5).
- **Compatibility:** Existing valid behavior preserved except explicit P01 changes (CR-6). `general.rows_collapsed` absent/blank/`False` => expanded; `True` => collapsed (D4, A-3).

## 4. Decision-gate checklist (B0, must be resolved before/within block implementation)

| ID | Uncertainty | Blocks | Required resolution (no speculative final choice outside evidence) |
|---|---|---|---|
| DG-a | Default-sheet removal strategy and identity guards | B1, B5 | Define workbook-scoped cleanup with produced-sheet identity tracking/guards; never delete produced sheets; respect Excel minimum-sheet/count and D3; choose deterministic lifecycle point (preferred end-of-per-report assembly before save per S1) based on Auditor/Researcher/Validator evidence. |
| DG-b | `max-N` lexical/grammar disambiguation vs range `-` and canonical forms (OQ-1) | B2, B3 | Define unambiguous bound/token grammar and precedence preserving existing `(1-max)`, `(min-max)`, numeric ranges/lists/scalars; specify accepted spacing/canonicalization and shared D1/D2 evaluation contract for filters and placeholders. |
| DG-c | Runtime source/handoff of inserted width/height for grouping bounds | B4, B5 | Use authoritative presented/pasted range dimensions at runtime (A-4); define semantic handoff (e.g. from `CopyPivotSnapshot`/post-paste destination range) to grouping step without changing paste semantics. |
| DG-d | Fill detection for groupable cells (D5) | B4 | Define property-based predicate for “no fill/pattern-none OR explicit solid white RGB(255,255,255)”; ignore conditional formatting; validate on styled samples. |
| DG-e | `rows_collapsed` config contract and backward compatibility (CR-5, D4, A-3) | B4, B5 | Extend `TreportGeneral`/`ReadConfig` for optional `rows_collapsed` (case-insensitive boolean-like), default expanded when absent/blank/False; define mechanism/timing for collapsed readiness without breaking configs. |
| DG-f | Integration timing/order and error-path interactions | B5 | Fix exact per-subreport sequence (paste -> format -> grouping hook with bounds) and per-output cleanup-before-save point; preserve existing filter application/state/error cleanup semantics. |
| DG-g | L1/L2/L3 evidence scope on working copies (OQ-3) | B7 (+all) | Select representative copied artifacts/scenarios proving CR-1..CR-5 and regressions; capture build report path and explicit L2/L3 status. |

## 5. Execution blocks (ordered)

### Block B0 — Baseline confirmation and decision-gate framing (completed: collection)
- **Status:** Auditor + Researcher baseline notes collected (S3–S5). No code changes.
- **Objective:** Confirm exports baseline and freeze technical uncertainties (DG-a..g).
- **Required agents:** Auditor (done), Researcher (done), Planner (required for checklists/plan), Validator optional.
- **Exit:** Uncertainties mapped to blocks; dependency order feasible.

### Block B1 — Default-sheet removal contract and implementation (CR-1, D3, D7) — HIGH RISK
- **Objective:** No Excel-default empty sheet left in normal successful report workbooks; never delete produced report sheets; D3 edge behavior aligned.
- **Scope (semantic):** `modTreportOutput` (`CleanupDefaultSheets` or replacement API; `GetOrCreateOutputWorkbook`/`AddReportSheet`/`SaveOutputWorkbook` as needed for identity/guards), `modTreportRun.Create_report` (deterministic per-output finalization point before save), produced-sheet identity tracking.
- **Dependencies:** B0; integrate sequencing in B5.
- **Risks:** Wrong-sheet deletion, last-sheet invariant, alerts/state, localized names, D3 paths.
- **Alternatives (no final choice until resolved):** cleanup at per-report end before save vs creation-time handling with restored global state; identity via tracked produced sheets/objects vs heuristics.
- **Required agents (all):** Auditor, Researcher, Planner, Validator, then Implementor.
- **Acceptance:** Deterministic guarded cleanup defined and integrated; produced sheets preserved; D3 consistent.
- **Verification:** L1, **L2 mandatory to claim compile**, L3 on copies (single/multi data_files/reports, no leftover default sheet in success, no-data/fail-fast cases).

### Block B2 — `max-N` offset support in `data_files.filters` (CR-2, D1, D2, D7) — HIGH RISK
- **Objective:** Accept/resolve `max-N` offset bounds/values wherever `max` is valid with D1/D2 fail-fast; preserve existing expressions.
- **Scope (semantic):** `modTreportFilters` (bound/token recognition, `CountValidRangeSplits`/classification/parse, `ResolveGlobalBound`, range/desired-value building, clause types/diagnostics as needed), reuse per-field domain/max from `ResolveFilterContext`/`GetOrCreateFieldState` and `modTreportPivot` domain helpers.
- **Dependencies:** B0; align D1/D2 contract with B3.
- **Risks:** `-` ambiguity/regressions, domain/computed-value edge cases, unclear diagnostics.
- **Alternatives (no final choice until resolved):** structured bound grammar with optional offset vs constrained forms; canonical spacing/forms per OQ-1.
- **Required agents (all):** Auditor, Researcher, Planner, Validator, then Implementor.
- **Acceptance:** Unambiguous grammar/precedence; `0 <= N < max_value`, `N==0==max`, computed value must exist in domain; clear fail-fast contexts.
- **Verification:** L1, **L2**, L3 on copies (`max-N` scalar/range, `N==0`, failures, and regressions `(1-max)`, `(min-max)`, numeric ranges/lists).

### Block B3 — `max-N_<field>` placeholder support in `report_sheet` (CR-3, D1, D2, D7) — HIGH RISK
- **Objective:** Support offset placeholders resolved from named field max, independent of filters, with aligned D1/D2 and preserved naming/sanitization/collision behavior.
- **Scope (semantic):** `modTreportOutput` (discovery/validation/substitution/remnant guards), `modTreportRun.ResolveReportSheetPlaceholderMaxValues` and pre-filter call point, `modTreportPivot` max/domain resolution helpers as needed.
- **Dependencies:** B0, B2 contract alignment for D1/D2.
- **Risks:** Token/guard regressions, timing/independence break, value formatting/collision (OQ-2).
- **Alternatives (no final choice until resolved):** canonical token form/spelling/spacing (OQ-1); resolver extension shape while keeping contract fixed.
- **Required agents (all):** Auditor, Researcher, Planner, Validator, then Implementor.
- **Acceptance:** Pre-filter independence preserved; D1/D2 contexts defined; legacy/malformed/unresolved handling consistent.
- **Verification:** L1, **L2**, L3 on copies (existing `<max_<field>>`, offset valid/invalid/domain-missing, independence from filters, sanitization/length/collision).

### Block B4 — `group_rows` grouping engine and `rows_collapsed` config (CR-4, CR-5, D4, D5, D6, D7) — HIGH RISK
- **Objective:** Per-subreport bounds-scoped Outline grouping by fill rule with silent skip on invalid params; optional `rows_collapsed` in `general` (default expanded).
- **Scope (semantic):** `modTreportConfig` (`TreportGeneral` + `ReadConfig` optional `rows_collapsed`), grouping engine procedure(s) with semantic inputs (worksheet, anchor, inserted width/height, 1-based `group_rows`, effective collapsed flag), bounds handoff from `modTreportExtract`/`modTreportRun` as decided.
- **Dependencies:** B0; integrate in B5.
- **Risks:** Out-of-bounds/cross-subreport grouping, fill/theme/white detection variance, Outline/COM side effects, collapse mechanism.
- **Alternatives (no final choice until resolved):** bounds handoff mechanism; fill-property predicate; collapsed-state application method.
- **Required agents (all):** Auditor, Researcher, Planner, Validator, then Implementor.
- **Acceptance:** D5 predicate, contiguous/single-run/start-end rules, D6 silent-skip scope, D4/A-3 defaulting/readiness all specified and implemented.
- **Verification:** L1, **L2**, L3 on copies (styled boundaries/no-fill/white, single/edge runs, multi-subreport independence, invalid/out-of-width skip, `rows_collapsed` absent/False/True).

### Block B5 — Pipeline integration and sequencing (CR-1, CR-4, CR-6) — HIGH RISK
- **Objective:** Wire cleanup, offset handling, grouping, and config into correct `Create_report`/`ProcessDataFile` lifecycle points without leakage/regressions.
- **Scope (semantic):** `modTreportRun` (cleanup-before-save per output, per-subreport post-paste/format grouping hook), `modTreportExtract`/`modTreportOutput`/`modTreportConfig` interfaces as required by resolved contracts.
- **Dependencies:** B1–B4.
- **Risks:** Ordering regressions, state/error-path interactions, bounds misuse.
- **Required agents (all):** Auditor, Researcher, Planner, Validator, then Implementor.
- **Acceptance:** Explicit per-iteration and per-output order documented and implemented; invariants preserved.
- **Verification:** L1, **L2**, L3 combined scenarios (offset filters+placeholders+grouping+default-sheet removal across outputs).

### Block B6 — Documentation and configuration contract updates (CR-5, CR-6)
- **Objective:** Align docs with finalized P01 behavior (minimal, evidence-based).
- **Scope (semantic):** `docs/01-vba-reporting/configuration-model.md`, `filters-dsl.md`, `pipeline-and-dataflow.md`, `overview.md` as needed.
- **Dependencies:** B2–B5 finalized.
- **Risks:** Documenting undecided syntax/behavior.
- **Required agents:** Planner required, Validator recommended, then Doc-specialist/Implementor per workflow if code-adjacent; docs-only edits committed separately as appropriate.
- **Verification:** Review against implemented behavior/evidence.

### Block B7 — Verification execution, evidence capture, and rollout readiness (CR-6) — HIGH RISK
- **Objective:** Produce auditable L1/L2/L3 evidence for CR-1..CR-5 and regressions on working copies.
- **Scope (semantic):** `build_treport_vba.ps1` on timestamped working copy + `.ai/tmp` report; manual **L2 Debug → Compile VBAProject** on working copy; **L3 on copies only**; record outcomes/blockers; no originals corrupted.
- **Dependencies:** B1–B6 integrated.
- **Required agents (all):** Auditor, Researcher, Planner, Validator (and Implementor/support as needed for fixes).
- **Acceptance:** L1 path/result recorded; L2 explicitly passed with working-copy identity/time or explicitly blocked (no compile claim without pass); L3 evidence for each CR and key regressions recorded.

## 6. Dependency and execution order

| Order | Block | Depends on | Gate |
|---|---|---|---|
| 1 | B0 | — | Completed (collection) |
| 2 | B1, B2, B3, B4 | B0 (+B3 aligned with B2 contract) | DG-a..e resolved per block with required agents |
| 3 | B5 | B1–B4 | All upstream contracts implemented/interface-stable; DG-f resolved |
| 4 | B6 | B2–B5 | Final behavior/OQ resolutions known |
| 5 | B7 | B1–B6 | Working copies + L2-capable environment or documented blockers |

**Critical path:** B0 → (B2/B3 + B4 + B1) → B5 → B7. High-risk blocks require Auditor, Researcher, Planner, Validator participation and must not choose final approaches where uncertainty remains without evidence.

## 7. Cross-cutting verification matrix

| CR / area | Blocks | L1 | L2 (manual, required to claim compile) | L3 (copies) minimum intent |
|---|---|---|---|---|
| CR-1, D3 | B1,B5,B7 | Required | Required | No default empty sheet in success; produced sheets preserved; no-data/fail-fast alignment |
| CR-2, D1,D2 | B2,B5,B7 | Required | Required | Offset forms/`N==0`, boundary/domain failures with context, expression regressions |
| CR-3, D1,D2 | B3,B5,B7 | Required | Required | Pre-filter independence, token/guard behavior, naming/sanitization/collision |
| CR-4, D5,D6 | B4,B5,B7 | Required | Required | Bounds-scoped contiguous grouping, fill rules, invalid-param silent skip, no cross-subreport merge |
| CR-5, D4 | B4,B6,B7 | Required | Required | `rows_collapsed` absent/False/True behavior, backward compatibility |
| CR-6, D7 | B1–B7 | Required | Required | General applicability and non-regression |

## 8. Notes and non-goals

- No speculative redesign or scope beyond P01 (no generic expression engine, no new filter/placeholder semantics outside offsets, no non-fill grouping criteria).
- Semantic targeting only (modules/procedures/types/keys/ListObjects), not line numbers in implementation tasks.
- Do not commit `data/theiss` binaries/results; keep working copies/reports under `.ai/tmp` or ignored paths as appropriate.
- **Block execution status (live):** B0 collection complete; B1–B7 pending execution in this run.

## 9. Execution log

- **2026-11-17:** B0 completed (Auditor + Researcher). Execution plan saved. B1 Auditor/Researcher/Planner completed; DG-a/DG-f contract drafted (per-output cleanup before save, produced identity via outUsedSheetNames, D3 fail-fast). B1 Validator review pending before implementation.


### B1 — Validator outcome (2026-11-17)
- **Verdict:** APPROVED with required corrections (must-fix before implementation).
- **Must-fix (1):** Safe sheet-deletion iteration — do not delete from Worksheets via For Each while mutating; use deterministic safe approach (reverse index loop or collect-then-delete pass) with per-delete Count/produced-invariant rechecks.
- **Must-fix (2):** Guaranteed DisplayAlerts restoration on all paths including runtime/delete errors (error-handler/cleanup that restores then propagates/fails with clear context).
- **Must-fix (3):** Explicit normalization and collection choice — protection check producedSheetNames.Exists(LCase) against lowercased (trimmed) registry keys; operate on wb.Worksheets consistently.
- **Must-fix (4):** D3 + post-cleanup invariants enforced concretely — before save fail fast with clear context if producedSheetNames is Nothing or .Count = 0 (no SaveOutputWorkbook); after cleanup fail fast before save if intended produced sheet(s) not present/workbook invalid. Cleanup/delete only on successful normal-completion path.
- **Must-fix (5):** Validation task must require evidence (L1 report path recorded; L2 manual Debug->Compile VBAProject pass on identified working copy/time or explicitly not done/not claimed; L3 on copies covering multi-data_files same-output inventory and locale/name-independence of removal).
- **Notes:** CleanupDefaultSheets signature change safe in-repo (unused). GetOrCreateOutputWorkbook unused by run/outside new contract. SaveOutputWorkbook remains persistence-only. EH preserved.

- **2026-11-17 (B1 implementation):** Implementor completed B1 per DG-a/DG-f and validator must-fix corrections. Changed exports: mko_treport/src/vbs/export/modTreportRun.bas, mko_treport/src/vbs/export/modTreportOutput.bas (cleanup with produced-sheet registry guards, D3 pre/post invariants, safe deletion/alerts restore, Create_report save-loop integration). Commit: 9dfe3d ix(treport): remove default sheets from report outputs with produced-sheet guards. L1: SUCCESS on working copy .ai\tmp\treport_vba_check_20261117_092902.xlsm, report .ai\tmp\build_treport_vba_20261117_092902_workcopy.txt (compile not confirmed by script/VBOM). L2: not performed (interactive VBE required) - compilation not claimed. L3: not performed in environment; runtime evidence pending (B7). Proceeding to B2/B3/B4.

- **2026-11-17 (B2/B3 in progress):** B2 Auditor/Researcher/Planner completed (DG-b/OQ-1/OQ-2/D1/D2 contract drafted; shared notes for B3 produced). B3 Auditor completed. B2 Validator and B3 Researcher launched in parallel.

- **2026-11-17 (B2 validator):** APPROVED with required corrections (RC-1..RC-10 must-fix before implementation): D1/D2 error-context precision (computed value when derived/relevant), explicitly define max-N behavior in list elements (recommended out of scope/preserve existing), exact deterministic bound-token grammar/classification covering signed numeric literals and (max-1) reinterpretation, concrete D2 membership test consistent with existing numeric domain equality semantics, minimal bound representation carrying offset/kind, B2/B3 semantic equivalence for domain-membership source/meaning and placeholder token grammar parity/remnant integration, complete L2/L3 regression matrix. Proceeding to B3 planning and B4 audit in parallel.

- **2026-11-17 (B2 corrected, B3/B4 designs):** B2 planner produced corrected design/task incorporating validator RC-1..RC-10 (explicit list-element exclusion, deterministic bound grammar/decomposition, (max-1) scalar reinterpretation, concrete D2 membership vs pre-resolved numeric domain, minimal bound representation, precise D1/D2 error contexts, full regression matrix). B3 planner produced design/task aligned to B2 shared D1/D2 contract (grammar-aware <max[-N]_field> parsing with remnant/legacy guards, pre-filter PivotItems/TryParseNumericConvertible domain membership parity, OQ-2 formatting, independence). B4 planner produced design/task (DG-c bounds handoff via CopyPivotSnapshot inserted range, DG-d static Interior predicate/contiguous runs/silent skip, DG-e ows_collapsed config+bounded collapse, DG-f post-format hook). Proceeding with validators before implementation.

- **2026-11-17 (B3/B4 validators):** B3 design conditionally approved with required corrections - must ensure grammar-aware scanning integrated with all legacy/malformed/unresolved/remnant guards (RC-9), exact-token substitution/keying for repeated/multiple/offset spellings, D2 domain membership via pre-filter PivotItems+TryParseNumericConvertible parity (not FieldAvailableNumericValues/ResolvedContext), RC-1 error contexts, and OQ-2 formatting. B4 design approved with corrections (2 must-fix blockers): **B4-B1** preserve existing paste/copy fail-fast semantics and logging in CopyPivotSnapshot refactor (do not convert paste errors to silent skip/continued processing; Nothing only for defensive/no-bounds cases as defined) and derive bounds/grouping target from returned inserted Range's worksheet/address (no ActiveWindow dependence). **B4-B2** deterministic D6 group_rows parsing/validation (> insCols using runtime inserted width A-4, whole-number/Long-safe/Trim, reject fractional/non-numeric/overflow/<1) with no COM attempt on param skip; localized COM-error handling logged with context and distinct from param skip, without masking paste failures. Proceeding to sequential implementation B2->B3->B4 (designs treated as incorporating corrections).

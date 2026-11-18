# Execution Plan: Combine Multiple Subreports on One Worksheet and Worksheet Name Length Validation (P01)

- **Plan file:** `.ai/plans/01-combine-subreports-worksheet-name-validation.md`
- **Source of truth (spec):** `.ai/problems/02_combine_subreports_worksheet_name_validation_spec.md`
- **Source problem:** `.ai/problems/P01.md`
- **Baseline context:** Auditor `{code_context}` provided with this planning request (current implementation, gaps, architecture, risks) — treated as the as-is baseline for decomposition.
- **Status:** Planning updated (no VBA/code or audit artifacts modified). **B1 (Config model extension): Planned — Validator re-review required before implementation.** B2–B6 remain sequenced per this plan; implementation execution follows approved spec decisions and execution directives after required reviews/gates.
- **Date:** 2026-11-18 (updated with B1 detailed implementor task and Validator-correction alignment)
- **Scope type:** Dependency-aware, minimal rollout plan for P01 only. No speculative redesign.

## 1. Executive summary

This plan decomposes P01 into logical, dependency-safe execution blocks to implement: (1) combining multiple subreports targeting the same final `report_sheet` within the same output workbook into one assembled worksheet (original `data_files` order, required spacing/transfer/limits), and (2) suffix-aware worksheet name length validation and conflict preflight for all generated/planned names.

**Spec baseline is implementation-ready:** DoR satisfied, PO decisions **D-01..D-06** and **OQ-7** approved/resolved, and **Approach 1** selected in the spec (§5.3, §11, §12). This plan follows that baseline and preserves existing Stage 1 preparation semantics unless explicitly overridden by P01/decisions.

**Product Owner clarification (Step 2): skipped.** All PO/business decisions are resolved in the spec and no business uncertainty remains; no additional PO questions are required before implementation.

**Key implementation surfaces (VBA SSOT):** `mko_treport/src/vbs/export/modTreportConfig.bas`, `modTreportRun.bas`, `modTreportOutput.bas`, `modTreportExtract.bas` (primary), with governed documentation updates in `docs/01-vba-reporting/*` as impacted. **No changes** to `Refresh_data`, unrelated filter/placeholder/insert/skip/template semantics except as strictly required, and no implicit `-SyncExports` / `-DeployToDataTheiss`.

**Verification split (EVD-02):** executor performs feasible checks (static review, L1 build/report review and other non-user checks) and reports results/limitations; **user performs final verification — compilation and run exactly from the workbook** (L2/L3 in workbook context). **Development must not stop until code is implemented** (EVD-03), subject to that final user gate.

**B1 planning update (this revision):** §11 adds a complete, implementor-ready B1 task with semantic targets, exact config contract, presence/parse diagnostics, edge-case rules, and done/acceptance checklist. B1 addresses Validator notes (presence-tracked required `report_rows_limit_general`, fully specified `gap_rows` semantics, strictness/diagnostics/error-code alignment, explicit non-goals/regression guards). **B1 requires Validator re-review before implementation.**

## 2. Scope, constraints, and non-goals

### 2.1 In scope (confirmed requirements)
- `gap_rows` in `data_files` (optional, integer `>= 0`, blank/missing => `0`, first subreport ignored at combine) — **D-01 / CR-07, CR-08**.
- `report_rows_limit_general` in `general` (explicit positive integer, no overrides/fallback, per prepared subreport incl. base; not applied to combined total) — **D-02 / CR-15**.
- Grouping by final resolved + sanitized `report_sheet` name, case-insensitive, within same output workbook; resolve -> sanitize -> group — **D-03 / CR-01, CR-02, CR-03**.
- Temp/base naming `base, base1..base(N-1)` in original `data_files` order; suffix-aware 31-char validation with effective base name and max `31 - suffix_length`; conflict checks against existing and all planned names before creating affected worksheets; never overwrite/delete/rename — **D-04, D-06, OQ-7 / CR-05, CR-06, CR-19n–CR-21n**.
- Two-stage flow: Stage 1 fully prepare each subreport on a separate worksheet (preserve existing prep), Stage 2 combine in preserved order per output/group — **CR-04, CR-09**.
- Combine placement/limits/content: `lastRow/nextRow/destinationRow`, first `gap_rows` ignored, no `paste_row/paste_col` reapply, valid-content-range fail-fast (D-05, valid one-row not rejected solely by `Rows.Count=1`), per-subreport `report_rows_limit_general`, projected final Excel row/column validation before each paste — **D-05, CR-10, CR-11, CR-15, CR-16, CR-17v**.
- Combine transfer profile: values (not formulas), formats incl. number formats, row heights, hidden row states, row grouping/outline levels restored at new positions; do **not** transfer column widths, hidden column states, zoom, FreezePanes, SplitRow/SplitColumn; preserve base layout/display — **CR-12, CR-13, CR-14**.
- Temp lifecycle: delete temps only after successful transfer; final workbook keeps base combined sheets (no temps); align produced/final registry and cleanup/verify/save guards — **CR-18c, D-06**.
- Documentation updates for config/schema/behavior and acceptance/verification reporting per DoR/CR-23g/EVDs.

### 2.2 Non-goals / out of scope
- `Refresh_data` workflow or source-data refresh behavior.
- Changing existing filter semantics, placeholder resolution rules, `insert_rows`/`skip_header`/`only_values`/template preparation semantics except as strictly required for P01.
- Reapplying/redesigning `paste_row`/`paste_col` in combine; transferring forbidden worksheet-level display/layout settings from temps; automatic conflict resolution.
- Applying `report_rows_limit_general` to combined totals.
- Per-`data_files` overrides or hardcoded fallback for `report_rows_limit_general`.
- Performance optimization beyond correctness/safe implementation needs.
- New report/output formats or unrelated config/schema changes.

### 2.3 Mandatory constraints and execution directives
- **SSOT:** VBA source of truth is `mko_treport/src/vbs/export/*`. Do not modify tracked workbook/VBA artifacts as part of planning; any `-SyncExports` / `-DeployToDataTheiss` is explicit opt-in only and never implicit.
- **No audit/code modification in this phase:** this document is planning output only; existing audit/problem artifacts remain unchanged.
- **EVD-01:** no separately prepared "ready template" is required to start/continue development.
- **EVD-02:** executor performs remaining feasible checks (incl. L1 build/report review) and reports what was/was not verified; **user performs final compilation and run exactly from the workbook**.
- **EVD-03:** development must not stop until code is implemented.
- **Excel limits:** sheet name 31 chars; 1,048,576 rows; 16,384 columns.
- **Safety:** no overwrite/delete/rename on conflicts; no temp deletion before successful transfer; preserve per-output isolation, `data_files` order, and existing cleanup/save guard safety model.

## 3. Baseline and confirmed gaps (from spec + Auditor `{code_context}`)

| Area | Current baseline (as-is) | Required target (P01) | Gap |
|---|---|---|---|
| Config (`modTreportConfig`) | `TreportGeneral`: input/output/delimeter/rowsCollapsed; `TreportDataFile` has no `gap_rows`; no `report_rows_limit_general` parsing/validation | `gap_rows` (D-01) + `report_rows_limit_general` (D-02) with contextual validation | **Missing** |
| Orchestration (`modTreportRun`) | Single-pass per data file/report to final sheet; duplicate final name => **2612** fail; `outUsedSheetNames` acts as collision+produced registry | Two-stage prep-then-combine, per-output grouping state, base/temp planning, lifecycle-aware registries | **Missing/major refactor** |
| Naming/prep (`modTreportOutput`) | `SanitizeSheetName` rejects >31 with sanitized name only; `PrepareTargetSheet` fails if exists (2634); no planned-name preflight across run | Suffix-aware planning/validation (D-04/OQ-7), existing+planned preflight (D-06), temp/base creation strategy | **Partial/missing** |
| Extraction/combine (`modTreportExtract`) | Stage 1 snapshot/format/grouping primitives exist; no `UsedRange` block append, no D-05 combine gate, no projected Excel-bound checks, no combine transfer profile | Stage 2 append helper(s) with CR-10–CR-17v semantics and CR-12–CR-14 transfer profile | **Missing (high risk)** |
| Finalization | Cleanup/verify assume produced finals; no temps | Delete temps only after successful transfer; finals/bases only; guards preserved | **Missing integration** |
| Docs | Describe current behavior (2612 collision, no combine/new params) | Updated config/schema/pipeline/overview behavior and acceptance alignment | **Missing** |

## 4. Logical execution blocks

Agent requirement rule used below: **High-risk blocks require Auditor, Researcher, Planner, and Validator.** Lower-risk blocks list only the agents that add material value.

### B0 — Baseline lock and implementation readiness (planning/control)

- **Objective:** Lock implementation baseline to spec + `{code_context}`, confirm no blocking business ambiguity, and establish execution guardrails before code changes.
- **Primary artifacts (read-only in this phase):** `.ai/problems/02_combine_subreports_worksheet_name_validation_spec.md`, Auditor `{code_context}`, project rules, `docs/02-build-and-deployment/vba-build.md`.
- **Key tasks:**
  - Confirm D-01..D-06/OQ-7, Approach 1 baseline, EVD-01..03, and acceptance targets (DoR-05) are the working contract.
  - Record/track technical uncertainties that prevent final low-level transfer-technique choice (see §7).
  - Prepare an executor verification checklist template (L1/static feasible now; user L2/L3 from workbook later).
- **Inputs:** spec, `{code_context}`, rules.
- **Outputs:** ready-to-execute work breakdown; uncertainty/risk register carried into B3–B6.
- **Dependencies:** none (precedes all implementation blocks).
- **Acceptance criteria:** no open PO/business questions; scope/non-goals and verification ownership explicit.
- **Risk level:** Low.
- **Agents required:** **Planner** (structure/sequence). *Auditor/Researcher/Validator not required unless new codebase uncertainty is discovered.*
- **Rationale:** baseline and decisions are already approved; work is coordination/control, not technical investigation.

### B1 — Configuration model extension and validation

- **Status:** **Planned — Validator re-review required before implementation.** Detailed implementor task: **§11**.
- **Objective:** Extend config contract and fail-fast validation for `gap_rows` (D-01) and `report_rows_limit_general` (D-02) with contextual diagnostics consistent with existing patterns.
- **Primary SSOT targets:** `mko_treport/src/vbs/export/modTreportConfig.bas` (types, `ReadConfig`, column detection/parsers/defaulting/validation). Doc touchpoints prepared/aligned in B6: `docs/01-vba-reporting/configuration-model.md`.
- **Key tasks:**
  - Add `gap_rows` to `TreportDataFile` and parsing: optional column (use existing `HasColumn`-style detection consistent with module), accept integer `>= 0`, blank/missing/empty/null => `0`, row-contextual errors (see §11).
  - Add `report_rows_limit_general` to `TreportGeneral` and parsing: **required explicit positive integer in `general` with presence tracking**, no default/fallback/override, contextual errors (see §11).
  - Preserve existing config behavior/validation for all current fields and error-code/style conventions.
  - Keep `data_files` row index/order metadata (`RowIndex`) available for deterministic planning.
- **Inputs:** D-01, D-02, existing config patterns (`HasColumn`, `Norm`, `Fail`, strict boolean/template parsers, `DataFilesParamContext`, general-parameter loop diagnostics).
- **Outputs:** parsed/validated config values available to run/planning/combine logic.
- **Dependencies:** B0.
- **Acceptance criteria:** valid configs load with correct defaults/values; missing/invalid/non-positive/non-integer/error cases fail fast with clear context; existing configs/fields behave unchanged except the new required `report_rows_limit_general` contract per D-02 and the additive optional `gap_rows` column semantics per D-01.
- **Risk level:** Medium (contract change; can block runs if mis-specified) — **not high by itself**.
- **Agents required:** **Planner** (detailed field/validation design, compatibility with existing parser conventions), **Validator** (independent check that D-01/D-02 semantics, presence-tracking, strictness, diagnostics, and non-goals/regression guards are fully covered). **Validator re-review required before implementation.**
- **Rationale:** behavior is well-specified; primary risk is omission/inconsistency rather than technical unknowns.

### B2 — Name resolution, grouping key, suffix-aware planning, and collision preflight

- **Objective:** Implement deterministic naming/grouping planning per D-03/D-04/D-06/OQ-7 before creating affected worksheets.
- **Primary SSOT targets:** `mko_treport/src/vbs/export/modTreportOutput.bas` (helpers), integration contract consumed by `modTreportRun.bas`.
- **Key tasks:**
  - Define grouping key pipeline: resolve placeholders/substitutions -> sanitize -> case-insensitive key, scoped to same output workbook (D-03/CR-01).
  - Compute sequential planned names per group in original `data_files` order: `base`, `base1`..`base(N-1)` using actual suffix length (OQ-7/CR-05).
  - Suffix-aware length validation (CR-19n–CR-21n, D-04): reject when generated name exceeds 31; error must show effective base name (post-subst/sanitize, without numeric suffix), max permitted base length `31 - suffix_length`, and may include raw template as context.
  - Collision preflight (D-06/CR-06): check existing worksheet names and all names planned for creation in current run before creating affected worksheets; fail identifying conflicting name; never overwrite/delete/rename. Preserve default-sheet cleanup behavior only when non-conflicting.
  - Produce planning artifacts usable by orchestration (e.g. per-output group plan: base/final name, ordered members, temp names, `gap_rows`, source row/report context) — exact structure is an implementation detail but must be deterministic and isolated per output.
- **Inputs:** resolved/sanitized names from current pipeline, D-03/D-04/D-06/OQ-7, existing 2601/2602/2634 semantics.
- **Outputs:** validated, conflict-free per-output naming/group plan and clear error diagnostics.
- **Dependencies:** B0; uses config context from B1 where needed.
- **Acceptance criteria:** correct grouping for identical final names incl. dynamic resolution/case-insensitivity; correct base/temp sequence; accurate suffix-aware limits/messages; collisions (existing or planned) fail before creation with identifying name and no destructive resolution.
- **Risk level:** Medium (logic/edge-case risk) — **can escalate to High if coupled with lifecycle refactor**.
- **Agents required:** **Planner** (algorithm/data-structure design, integration seams), **Auditor** (targeted review of existing naming/collision/cleanup behavior and hidden coupling), **Validator** (independent review of D-03/D-04/D-06/OQ-7 edge cases incl. multi-digit suffixes, case variants, sanitization effects).
- **Rationale:** behavior is specified but correctness depends on subtle interactions with current 2612/2634/cleanup and per-output registries.

### B3 — Orchestration refactor, grouping state, and two-stage lifecycle (core flow)

- **Objective:** Refactor run flow to prepare-then-combine per output workbook while preserving Stage 1 semantics, per-output isolation, `data_files` order, and safety invariants.
- **Primary SSOT targets:** `mko_treport/src/vbs/export/modTreportRun.bas` (primary), `modTreportOutput.bas` (prep/lifecycle integration), coordination with `modTreportExtract.bas`.
- **Key tasks:**
  - Introduce per-output grouping/planning state: base report sheet name, ordered prepared members, temp worksheet identities/names, associated `gap_rows`, and source/report context.
  - Stage 1: for each data file/report emission, resolve/group/plan names and fully prepare each subreport on its **own** worksheet (base or temp per plan) using existing prep pipeline (`PrepareTargetSheet`, `CopyPivotSnapshot`, `ApplySheetFormatting`, `ApplyOutlineGrouping`) and existing `paste_row/paste_col`/template/insert/skip/only_values/placeholder behaviors.
  - Replace/reshape current duplicate-name handling (2612) with grouping behavior; separate concerns: planned/temp namespace vs final produced/base registry used for 2630/cleanup/verify/save.
  - Stage 2 orchestration hook per output/group in preserved order (actual append logic in B4): ensure base contains first subreport, sequence temps, enforce "first `gap_rows` ignored" at combine boundary.
  - Maintain per-output-workbook isolation and existing fail-fast/error-context patterns throughout.
- **Inputs:** B1 config, B2 naming/group plan, existing Stage 1 pipeline and invariants.
- **Outputs:** integrated two-stage execution flow with correct state handoff to combine/finalization.
- **Dependencies:** B1, B2. **Precedes/parallel design input to B4; must be integrated before B5 completion.**
- **Acceptance criteria:** single-subreport/non-combined behavior preserved; multi-subreport groups prepared as separate sheets then combined in order; no `paste_row/paste_col` reapplication in Stage 2; registries/guards remain coherent.
- **Risk level:** **High** (core control-flow/lifecycle change, regression-prone).
- **Agents required:** **Auditor, Researcher, Planner, Validator** (all).
- **Rationale:** High-risk refactor of orchestration and produced-registry semantics; requires deep code-path understanding, design scrutiny, and independent review.
- **B3 implementation context (2026-11-18, post-B1/B2 SSOT):** RC-5 scan/build/preflight and B2 per-output plan maps (`outPlanMembers`, `outPlanGroups`, `outPlanNames`) remain the pre-creation baseline; remove RC-1 gate (`modTreportRun.Create_report`, current 2699 path) only when replaced by Stage 1 preparation for all members and B4/B5 orchestration hooks. Keep `outUsedSheetNames(r)` as **final/base-produced registry only** (2630/cleanup/verify/save); track **planned/reserved** names via `outPlanNames(r)` and **temp/transient** prepared non-base worksheets via per-output runtime group state (worksheet refs, planned names, gapRows, contexts, creation/transfer status) — temps must not be registered as finals and must be deleted only after successful transfer (B5). Generalize single-member execution to df-ordered Stage 1 preparation on assigned planned sheets (base for first group member, temps `base1..N-1`) preserving placeholder-max-before-df-filters, df/report filter order, 2410, placement, `PrepareTargetSheet`/`CopyPivotSnapshot`/`ApplySheetFormatting`/`ApplyOutlineGrouping`. Hand off ordered groups (base sheet + temp members + gaps/contexts/limits) to B4 after Stage 1 completion per output; run B5 temp deletion after successful combine and before cleanup/verify/save. Uncertainties R-01/R-02/CR-13 remain B4-scoped transfer/validation choices.

### B4 — Stage 2 combine semantics, transfer fidelity, and bounds/limit enforcement

- **Objective:** Implement safe, spec-compliant append of prepared subreports into base worksheets with required validations and transfer profile.
- **Primary SSOT targets:** `mko_treport/src/vbs/export/modTreportExtract.bas` (combine helper(s) / reusable primitives), called from `modTreportRun.bas` (and/or `modTreportOutput.bas` as appropriate) with strict separation from Stage 1 prep.
- **Key tasks:**
  - Compute append geometry per P01: `lastRow = UsedRange.Row + UsedRange.Rows.Count - 1`, `nextRow = lastRow + 1`, `destinationRow = nextRow + gap_rows` (ignore first subreport `gap_rows`), sequential processing in original `data_files` order.
  - D-05 valid-content-range gate: fail fast with contextual error (source/subreport, worksheet, associated `report_sheet`) if expected subreport has no valid content range; do **not** reject a valid one-row report solely because `UsedRange.Rows.Count = 1` (check actual data/formulas).
  - Enforce `report_rows_limit_general` per prepared subreport (incl. base) before combine with required context (worksheet name, report_sheet, actual, configured).
  - Validate projected final combined dimensions before **each** paste against Excel limits (1,048,576 rows / 16,384 columns) with required error fields; do not attempt paste if exceeded.
  - Transfer profile (CR-12–CR-14): values not formulas; formats incl. number formats; row heights; hidden row states; row grouping/outline levels restored at **new positions**. Exclude column widths, hidden column states, zoom, FreezePanes, SplitRow/SplitColumn. Preserve base layout/display and source range starting column. No `paste_row/paste_col` reapplication.
  - Keep operations on explicit objects (no `Select`/`Activate` dependency in combine path), manage `CutCopyMode`/state deterministically, and preserve performance controls where safe.
- **Inputs:** prepared base/temp worksheets and planning metadata (B3), config limits (B1), D-02/D-05/CR-10–CR-17v.
- **Outputs:** correctly combined base worksheets ready for finalization; no temps modified into final artifacts except as sources prior to deletion.
- **Dependencies:** B1, B3 (and B2 planning outputs). **Highest technical uncertainty — see §7.**
- **Acceptance criteria:** all combine/transfer/limit/content/bounds acceptance cases in DoR-05 satisfied in feasible verification; no forbidden properties transferred; grouping/hidden rows fidelity addressed per chosen/validated technique and user L3 verification focus.
- **Risk level:** **High (critical fidelity and environment-dependent behavior).**
- **Agents required:** **Auditor, Researcher, Planner, Validator** (all).
- **Rationale:** Excel/VBA object-model behavior for outlines/hidden rows/UsedRange/paste semantics is technically uncertain; multiple viable low-level techniques exist and must be investigated/designed/reviewed before/alongside implementation.

### B5 — Cleanup, produced-registry, finalization, and end-to-end integration

- **Objective:** Ensure temp worksheets are removed only after successful transfer and final outputs/guards remain correct for save and workbook integrity.
- **Primary SSOT targets:** `modTreportRun.bas`, `modTreportOutput.bas` (`CleanupDefaultSheets`, `EnsureProducedSheetsPresent`, save flow integration).
- **Key tasks:**
  - Track final/base produced sheets separately from transient temp sheets so **2630**/cleanup/verify/save semantics remain valid and temps are not saved as finals.
  - Delete temporary worksheets only after successful transfer per group (CR-18c); on failure, preserve fail-fast behavior and avoid false-success cleanup claims.
  - Preserve existing cleanup behavior for unused default worksheets only when they do not conflict with planned report names (D-06).
  - Verify end-to-end per-output isolation and unchanged success/failure save semantics (no-save on error path as today, unless explicitly required otherwise by existing invariants).
- **Inputs:** B3/B4 integrated flow and registries.
- **Outputs:** final output workbooks containing base combined sheets only (plus allowed non-conflicting defaults as applicable).
- **Dependencies:** B3, B4.
- **Acceptance criteria:** no temp sheets remain after successful combine; produced-sheet guards (2630/2620–2629 family behavior as applicable) remain correct; no regression in save/cleanup safety model.
- **Risk level:** **High** (artifact lifecycle, save/cleanup invariants, rollback/error-path implications).
- **Agents required:** **Auditor, Researcher, Planner, Validator** (all).
- **Rationale:** subtle coupling between registries, cleanup, and Excel workbook state requires careful design and independent validation.

### B6 — Documentation, acceptance alignment, verification execution/support, and handoff

- **Objective:** Complete governed documentation updates and verification reporting consistent with EVD-01..03 and project rules.
- **Primary targets:** `docs/01-vba-reporting/configuration-model.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`, `docs/01-vba-reporting/overview.md` (and related governed docs only if impacted). Verification artifacts/reports as produced by executor runs (e.g. `.ai/tmp` build reports) referenced in handoff notes.
- **Key tasks:**
  - Update config/schema docs for `gap_rows` and `report_rows_limit_general` (location, type, defaults/mandatoriness, validation, semantics) and pipeline/overview for two-stage combine, naming/planning, transfer profile, limits, and lifecycle.
  - Document acceptance scenarios mapped to DoR-05 (multi-group combine/order/gaps, naming sequence, collisions, suffix-aware errors incl. multi-digit, per-subreport limits/base+temps, D-05 incl. valid one-row, projected Excel limits before paste, transfer inclusions/exclusions + restored grouping/hidden rows at new positions, base layout/display preserved, no temps left, no `paste_row/paste_col` reapply, single-subreport/non-combined regression cases).
  - Executor feasible checks: static review, **L1** build via `mko_treport/src/scripts/build_treport_vba.ps1` on a safe target (prefer copy in `.ai/tmp` for verification-only runs) with report capture; record exact commands, report paths, and any VBOM/compile-note limitations. Do **not** claim L2 compile success without manual VBE compile.
  - Prepare clear user handoff for **final L2 (Debug → Compile VBAProject) and L3 (functional run on copies)** from workbook (EVD-02), including focus areas tied to R-01/R-02/R-04 and output fidelity.
  - No implicit `-SyncExports` / `-DeployToDataTheiss`; document them only as explicit opt-in post-verification actions if needed.
- **Inputs:** implemented behavior from B1–B5, DoR-05/08, EVD-01..03, build script/docs.
- **Outputs:** updated governed docs, verification checklist/results summary, user L2/L3 handoff.
- **Dependencies:** B1–B5 (docs and final acceptance checks reflect final behavior).
- **Acceptance criteria:** docs match implemented contract/behavior; executor reports L1/static outcomes and limitations transparently; user-owned final workbook compile/run is explicitly requested/documented.
- **Risk level:** Medium (documentation/verification completeness; no runtime code risk by itself).
- **Agents required:** **Planner** (acceptance/verification/doc structure), **Validator** (independent completeness check against DoR-05/08 and spec decisions). *Auditor/Researcher not required unless implementation exposed unresolved technical behavior needing investigation.*
- **Rationale:** primary risk is omission/misalignment in reporting and docs, not new technical uncertainty.

## 5. Dependencies and execution order

### 5.1 Dependency graph (blocks)

| Block | Depends on | Enables | Notes |
|---|---|---|---|
| B0 Baseline/readiness | — | B1–B6 | Control gate; skip PO clarification (resolved) |
| B1 Config model/validation | B0 | B2, B3, B4 | Contract foundation for planning and limits; **§11 task, Validator re-review required before implementation** |
| B2 Naming/group planning/preflight | B0, B1 (context) | B3, B4 | Must run before affected worksheet creation in runtime design |
| B3 Orchestration/two-stage lifecycle | B1, B2 | B4, B5 | Core flow refactor and state handoff |
| B4 Combine transfer/bounds/limits | B1, B3 (B2 outputs) | B5, B6 | Highest uncertainty/fidelity block |
| B5 Cleanup/registry/finalization | B3, B4 | B6 | Lifecycle completion and save-safety alignment |
| B6 Docs/verification/handoff | B1–B5 | Release-readiness handoff | Final executor checks + user L2/L3 request |

### 5.2 Recommended execution order
1. **B0** — confirm baseline, uncertainties, and guardrails.
2. **B1** — implement/validate config extensions per **§11** after **Validator re-review approval** (smallest safe contract change).
3. **B2** — implement naming/group planning and preflight logic (pure/deterministic where possible, ready for orchestration integration).
4. **B3** — integrate two-stage orchestration and grouping state (preserve Stage 1).
5. **B4** — implement Stage 2 combine/transfer/validations (iterate with investigation/design as needed per uncertainties).
6. **B5** — integrate cleanup/registry/finalization and error-path safety.
7. **B6** — update docs, run executor-feasible verification (L1/static), and prepare user L2/L3 handoff.

**Parallelization guidance (limited):** B1 and parts of B2 design/prep can overlap once contracts are fixed; B4 design/spike investigation may begin during B3 if interfaces are stable, but integration must follow dependency order. Prefer sequential integration for B3→B4→B5 to control regression risk.

## 6. Risk assessment and mitigations

### 6.1 Risk register (implementation, rollout, regression, compatibility, fidelity)

| ID | Category | Risk | Likelihood | Impact | Mitigation / control | Primary blocks |
|---|---|---|---|---|---|---|
| R-01 | Implementation / Fidelity / Compatibility | **Row grouping/outline levels and hidden row states restored at new positions** may be incorrect for nested/complex outlines or differ by Excel behavior | Medium-High | High | Explicit range/row-mapped transfer strategy; avoid assuming simple block-copy preserves offsets; focused test matrix; Auditor/Researcher investigation; user L3 validation on representative templates/data | B4, B5 |
| R-02 | Implementation / Fidelity | **`UsedRange` inflation/staleness** and D-05 validity detection; must not reject valid one-row report solely by `Rows.Count=1` while failing true empty/no-content subreports | Medium | High | Implement D-05 data/formula-aware checks per spec intent; contextual diagnostics; validate before combine; targeted edge-case tests (one-row valid, empty, formatted-only ranges) | B4 |
| R-03 | Implementation / Regression | Naming/grouping errors (resolve->sanitize->case-insensitive, actual suffix length, multi-digit suffixes, planned/existing preflight) causing wrong combine, false collisions, or invalid names | Medium | High | Deterministic planner, pre-creation checks, unit-like scenario coverage in L3 plan, Validator review of D-03/D-04/D-06/OQ-7 | B2, B3 |
| R-04 | Regression | Interaction with existing Stage 1 prep modes (`skip_header`, `insert_rows` incl. template/legacy P09 behavior, `only_values`, templates, formatting, config-driven `ApplyOutlineGrouping`) and `paste_row/paste_col` non-reapplication in Stage 2 | Medium | High | Preserve current prep calls/semantics; isolate combine from Stage 1 placement logic; regression scenarios for single and multi-subreport cases | B3, B4, B5 |
| R-05 | Implementation / Performance | Large/many subreports increase copy time, clipboard contention, and memory pressure | Medium | Medium | Block/value transfers where compliant, minimal paste ops, deterministic `CutCopyMode` cleanup, retain existing fast app-state controls; avoid speculative optimization | B4 |
| R-06 | Regression / Rollout | Temp/final registry and cleanup/save guard coupling (2630, 2620–2629) causing premature temp deletion, temps saved in output, or broken invariants | Medium | High | Separate planned/temp/final registries; delete temps only after successful transfer; preserve per-output isolation and guard order; error-path review | B3, B5 |
| R-07 | Verification / Rollout | Runtime COM/Excel behavior cannot be fully proven without user-performed workbook compile/run; executor cannot claim L2/L3 completion | High (certainty of limitation) | Medium-High | EVD-02 split: executor L1/static/feasible checks + transparent report; user performs final L2 compile and run from workbook; provide focused L3 scenarios | B6 (all blocks) |
| R-08 | Compatibility (Excel version / Windows) | Differences in outline/grouping APIs, `UsedRange`, PasteSpecial behavior, window/state handling across Excel builds/Windows environments | Medium | Medium-High | Use explicit object-model operations, avoid reliance on UI selection, validate projected bounds, document tested/untested assumptions, prioritize user L3 in target environment | B4, B6 |
| R-09 | Compatibility (macro security / VBOM) | Macro settings, Trust Center/VBOM access, or policies may block build automation or L2 compile steps in some environments | Medium | Medium | Follow project VBA workflow; run L1 on copies when useful; report VBOM/compile limitations clearly; user L2 in trusted workbook context | B0, B6 |
| R-10 | Compatibility (Power Query refresh timing) | Power Query refresh timing/state could affect data freshness in broader workflows, but P01 targets `Create_report` combine/validation and `Refresh_data` is out of scope | Low (for P01 scope) | Low-Medium | No changes to refresh flow; keep P01 verification focused on `Create_report` outputs on copies; note dependency if user scenarios involve stale/external-query pivots | B6 |
| R-11 | Compatibility (pivot/object-model behavior) | Pivot presented-range updates, `DataBodyRange`/`TableRange` behavior, and COM timing can affect Stage 1 snapshot correctness feeding Stage 2 | Low-Medium | Medium | Preserve existing Stage 1 extraction/update logic; do not alter pivot/filter semantics; include regression checks where source pivots drive combined outputs | B3, B4 |
| R-12 | Compatibility / Rollout (workbook state/locking) | Workbook locks, open Excel sessions, or in-place build on tracked workbook can interfere with safe verification/deploy | Medium | Medium | Prefer verification builds on `.ai/tmp` copies; avoid implicit sync/deploy; do not run macros on production/original data; respect EVD-01/02 and project rules | B0, B6 |
| R-13 | Implementation | Projected-dimension validation and append math errors (off-by-one, starting column preservation, first-gap ignore) causing overflow or misplaced data | Medium | High | Implement formulas exactly per CR-10/CR-11/CR-16; pre-paste boundary checks; acceptance cases for boundary and multi-append sequences | B4 |

### 6.2 Rollout / compatibility notes summary
- **Excel version / Windows:** core row/column/name limits are stable, but COM behavior (outline/grouping, hidden rows, clipboard/PasteSpecial, `UsedRange`) may vary; treat fidelity as environment-sensitive until user L3 confirmation.
- **Macro security:** may prevent automated compile signals; does not block code implementation in exports but affects verification evidence.
- **Power Query refresh timing:** out of direct scope; no functional dependency introduced by P01.
- **Pivot/object-model behavior:** preserve existing Stage 1 behavior; combined output correctness depends on unchanged snapshot preparation.
- **Workbook state/locking:** use copy-based L1/L3 workflows and avoid in-place risky operations.
- **Output fidelity:** highest-priority acceptance dimension (values/formats/number formats/row heights/hidden rows/grouping positions/layout/display, no temps, correct order/spacing).

## 7. Implementation path alternatives and technical uncertainties

### 7.1 Alternatives considered (relevant)

| Area | Alternative | Pros | Cons / risks | Status vs uncertainty |
|---|---|---|---|---|
| Overall flow | **A1 (spec baseline): Two-stage orchestration in `modTreportRun` + combine helper in `modTreportExtract` + naming/planning in `modTreportOutput` + config in `modTreportConfig`** | Preserves Stage 1 semantics and module boundaries; aligns with D-01..D-06/OQ-7 and CRs | Requires careful lifecycle/registry refactor and robust Stage 2 transfer implementation | **Preferred baseline direction**; low-level Stage 2 transfer technique remains technically uncertain (see §7.2) |
| Overall flow | **A2:** Direct in-place generation into base with offsets (no real temp worksheets) | Potentially simpler runtime objects | Conflicts with "fully generated on separate worksheet first"/independent combine intent; high regression risk to Stage 1 semantics | **Rejected by spec intent; not recommended** |
| Overall flow | **A3:** Broad sheet/range copy then heavy post-processing | May simplify some copy operations | High risk of violating transfer inclusions/exclusions and altering base layout/display; hard to control CR-13 | **Rejected as primary by spec; not recommended** |
| Stage 2 transfer technique (within A1 sequencing) | **T-a:** Explicit row/range-mapped reconstruction (copy values/formats/row heights, then reapply hidden states and outline/group levels at destination offsets programmatically) | Best control over CR-12–CR-14 and offset correctness; avoids unintended worksheet-level transfers | Complex for nested outlines/group structures; Excel object-model behavior can be subtle | **Viable; not fully chosen** |
| Stage 2 transfer technique | **T-b:** Block `Copy`/`PasteSpecial` sequences limited to allowed attributes + targeted post-pass for hidden rows/grouping restoration at offsets | Potentially efficient for values/formats; fewer loops for some attributes | Clipboard/state side effects; risk of accidentally transferring excluded properties; grouping/hidden-row offset restoration still required | **Viable; not fully chosen** |
| Stage 2 transfer technique | **T-c:** Hybrid (value arrays/format-safe pastes for content + explicit per-row/outline metadata reconstruction) | Balances performance and control | More moving parts; requires careful ordering and state management | **Viable; not fully chosen** |
| D-05 content validity detection | **V-a:** `UsedRange` bounds + scan for non-empty constants/formulas in relevant area | Closer to "actual data or formulas" intent; avoids rejecting valid one-row reports by count alone | Must define efficient/reliable scan boundaries given `UsedRange` quirks | **Viable; detail not fully chosen** |
| D-05 content validity detection | **V-b:** Heuristic based on `UsedRange` size/special-cells with safeguards for one-row valid content | Simple in some cases | Risk of false positives/negatives with stale `UsedRange` or formatted-only cells | **Needs caution; not chosen** |
| Naming/collision planning | **P-a:** Single per-output preflight plan (all resolved/sanitized names, temp/final sets, reservations) before any affected worksheet creation | Directly satisfies D-06/OQ-7 and fail-before-create intent; deterministic | Requires clear integration with existing per-output registries | **Strongly indicated by spec; implementation details open** |
| Naming/collision planning | **P-b:** Incremental checks during creation | Simpler local checks | Higher risk of partial-state/conflict-order issues; harder to prove "all planned names checked before creating affected worksheets" | **Not recommended** |
| `gap_rows` config parsing | **G-a:** `HasColumn` optional detection (blank/missing => 0) consistent with existing optional-column patterns | Compatible with existing workbook schemas and D-01 | Must align with project conventions and tests | **Chosen baseline for B1 (consistent with existing optional-column handling); implementation details per §11** |
| `gap_rows` config parsing | **G-b:** Require physical column presence but allow blank values | Explicit schema | Slightly stricter than "optional column" wording if interpreted as absence-allowed | **Not selected** (B1 follows G-a to satisfy "column absent => 0") |

### 7.2 Technical uncertainties preventing a fully final low-level implementation choice

Despite approved decisions and selected high-level **Approach 1**, the **final low-level Stage 2 transfer/validation technique is not fully chosen** at planning time due to the following technical uncertainties:

1. **Outline/grouping and hidden-row fidelity (R-01, R-08):** exact, reliable method to transfer/restore **row grouping/outline levels and hidden row states at new row positions** (including nested/complex outlines) across Excel/Windows environments is not conclusively determined from static analysis/spec alone.
2. **D-05 "valid content range" detection (R-02):** robust, performant criterion that accepts a valid one-row report (including cases where `UsedRange.Rows.Count = 1`) but fails true empty/no-content subreports under `UsedRange` inflation/staleness requires targeted implementation validation.
3. **Clipboard/PasteSpecial side effects (R-05, R-08, CR-13):** whether any block copy/paste sequence can satisfy "values not formulas, formats incl. number formats, row heights" while guaranteeing exclusion of column widths/hidden columns/zoom/freeze/split in all target environments without unintended state leakage.
4. **Performance vs. correctness tradeoff (R-05):** optimal balance for large combines is environment/data-dependent and must not drive speculative redesign.

**Planning implication:** B4/B5 must proceed with Auditor/Researcher/Planner/Validator involvement, prefer correctness/fidelity over speed, and include targeted user L3 scenarios focused on these uncertainties. If investigation during implementation identifies a materially different safe technique within Approach 1, this plan permits selecting that technique provided it satisfies all confirmed CR/D/OQ requirements.

### 7.3 B3/B4 research notes (2026-11-18, post-B2 SSOT, implementation handoff)

- **B3 orchestration/runtime state (viable VBA patterns):** B2 already produces per-output `outPlanGroups(r)` (group-key -> group with ordered `Members` payload dicts carrying `ReportIndex/DataFileIndex/DataFileRowIndex/EmissionOrder/GapRows/EffectiveBaseName/GroupKey/PlannedName/SuffixText/SuffixLength/OutputName/ReportName/FileCtx/RepCtx/RawReportSheetTemplate`) and `outPlanNames(r)` / `outUsedSheetNames(r)`. For B3, extend existing `Scripting.Dictionary`-of-dictionaries/Variant structures (consistent with current module style) rather than introducing cross-module UDT collections: UDTs cannot be stored in `Collection`/`Dictionary` items and public exposure adds friction, while report/df/group counts are moderate. Recommended: enrich each planned member record with runtime fields (`IsBase`, `Sheet` As Worksheet, `PreparedRange`/inserted bounds, `SourceUsedRange`/valid-content range, `CombineStatus`, error context refs) and keep per-group ordered member access from B2 order; maintain `outPlanNames` as planned/reserved namespace, keep `outUsedSheetNames` as **final/base-produced only** at reconciliation point for 2630/cleanup/verify/save, and track temp worksheets separately until B5 deletes them after successful transfer. This aligns with RC-3 and §4 B3 lifecycle notes.
- **B3 Stage 1 generalization:** preserve RC-5 scan/plan/preflight semantics (placeholder max before df filters already resolved in scan; planning/2637/2638/2699 gate happen before Stage 1 creation). Stage 1 execution must still reopen each source workbook read-only per `data_files` row in original order, apply `data_files` filters once and per-report filters in the same sequence, enforce 2410 on empty pivot data, compute placement, and call `PrepareTargetSheet`/`CopyPivotSnapshot`/`ApplySheetFormatting`/`ApplyOutlineGrouping` on the **B2-assigned planned sheet** (base or temp). Member lookup must be emission/member-aware (by `EmissionOrder` or `DataFileIndex`+`ReportIndex` from plan) because a df row emits per report and groups may span rows; do not assume one sheet per df/report index outside the plan. Record prepared worksheet/range refs into runtime state for B4; do not register temps in final-produced registry.
- **D-05 valid-content detection (alternatives/tradeoffs, B4):** spec requires fail-fast when expected subreport has "no valid content range" and must not reject a valid one-row report solely because `UsedRange.Rows.Count = 1`; it also requires checking actual data/formulas and CR-11/CR-10 use `UsedRange` for append geometry/block copy. `Worksheet.UsedRange` is convenient but can be inflated/stale (formatting, cleared cells) and is not a proof of meaningful content. Viable checks: (a) `Application.WorksheetFunction.CountA(candidateRange) > 0` as quick non-empty signal (but formula returning `""` may be counted in some cases; verify), (b) `SpecialCells(xlCellTypeConstants)` / `SpecialCells(xlCellTypeFormulas)` on candidate range with "no cells found" error handling to prove data/formula presence, (c) `Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows/ByColumns, SearchDirection:=xlPrevious, ...)` for last-content detection/presence. **Recommendation:** treat D-05 as a presence gate (data or formula exists in the prepared range) rather than a row-count-only gate; for a one-row UsedRange, explicitly test constants/formulas before failing. Keep append/copy block aligned with spec's `UsedRange` semantics unless implementation proves a safer bounded range that still satisfies CR-10/11/12–14 and all acceptance cases; document the chosen criterion in code comments/diagnostics. Residual uncertainty remains for formatted-only/stale UsedRange cases (R-02).
- **CR-12–CR-14 Stage 2 transfer (alternatives/tradeoffs, B4):** required profile = values not formulas, formats incl. number formats, row heights, hidden row states, row grouping/outline levels restored at new row positions; forbidden = column widths, hidden column states, zoom, FreezePanes, SplitRow/SplitColumn; preserve base layout/display and source starting column; no `paste_row/paste_col` reapply. Viable techniques: (T-a) explicit reconstruction: `dest.Value = src.Value` (or equivalent value-only transfer) + `PasteSpecial xlPasteFormats` for formats/number formats + per-row `RowHeight` + per-row `Hidden` + per-row `OutlineLevel` mapped to destination offsets; avoids `xlPasteColumnWidths`/`xlPasteAll` and gives strong control. (T-b) clipboard block pastes (e.g. `xlPasteValuesAndNumberFormats` + `xlPasteFormats`) plus explicit per-row/outline/hidden restoration; fewer loops for some attributes but depends on clipboard/CutCopyMode discipline. (T-c) hybrid (value arrays/range assignment + formats paste + metadata reconstruction). **Recommendation (correctness-first):** prefer T-a/T-c style separation: transfer values without formulas, apply formats via `xlPasteFormats` only (not column widths), then explicitly restore row heights, hidden-row states, and `OutlineLevel` row-by-row at `destinationRow + offset`; never use `xlPasteColumnWidths`, `xlPasteAll`, or sheet/window copy operations. For outline fidelity, setting destination row `OutlineLevel` from corresponding source row is the most controllable offset-preserving method, but nested outline behavior and summary-level interactions must be validated in user L3 (R-01). Always clear `Application.CutCopyMode` deterministically (including error paths where practical) and avoid `Select`/`Activate`.
- **CR-10/CR-11/CR-16 geometry and limits (B4, high certainty from spec/math):** for each append, compute base `lastRow = wsBase.UsedRange.Row + wsBase.UsedRange.Rows.Count - 1`, `nextRow = lastRow + 1`, `destinationRow = nextRow + member.gapRows` except ignore first subreport gap in combine sequence; source block is the member's valid prepared range (`UsedRange` per spec, subject to D-05 gate) preserving `sourceRange.Column`. Before paste, validate projected bounds: `lastDestRow = destinationRow + src.Rows.Count - 1 <= 1048576` and `lastDestCol = src.Column + src.Columns.Count - 1 <= 16384`; fail with final `report_sheet`, projected boundary, and Excel limit, and do not paste. Advance `nextRow = destinationRow + src.Rows.Count` (or recompute from base UsedRange after successful append). Enforce `report_rows_limit_general` per prepared subreport (including base) before combine with required context (worksheet/planned name, report_sheet/base name, actual rows, configured limit).
- **B5 lifecycle/registry guidance (high certainty from invariants):** delete temp worksheets only after successful transfer of required content for that group/member sequence; never before successful combine. Keep temps out of `outUsedSheetNames` (final-produced registry consumed by 2630/cleanup/verify/save). After successful combine, ensure only base/final sheet remains represented as produced and all temps removed; preserve existing cleanup behavior for unproduced default sheets only when non-conflicting (D-06). On any failure, preserve fail-fast/no-save behavior and do not claim successful cleanup.
- **Selected path guidance (where certain):** stay within **Approach 1** and current module boundaries (`modTreportRun` orchestration/lifecycle/state, `modTreportOutput` naming/planning/preflight integration, `modTreportExtract` Stage 2 combine/validation/transfer helpers, `modTreportConfig` unchanged except B1-complete contract). Reuse B2 plan artifacts/maps and remove 2699 RC-1 gate only together with complete B3/B4/B5-safe flow. **Not fully chosen (technical uncertainty remains):** final low-level D-05 range-selection criterion under stale/formatted UsedRange, exact outline/hidden-row restoration mechanics for nested/complex cases across Excel versions, and clipboard/PasteSpecial edge behavior; these require implementation-time validation and focused user L3 confirmation.

## 8. Verification strategy and handoff

### 8.1 Executor-feasible checks (this implementation effort)
- **Static/semantic review:** confirm code paths align with D-01..D-06/OQ-7, CR-01..CR-23g, EVDs, and module responsibilities; verify no out-of-scope changes.
- **L1 build (feasible):** run `mko_treport/src/scripts/build_treport_vba.ps1` using a safe verification target when appropriate (copy tracked workbook to `.ai/tmp` and build the copy) and capture report path/summary. Do not claim compile success beyond what the report supports.
- **Limitations reporting:** explicitly state if L1 compile is best-effort/unavailable due to VBOM/macro policy and that L2 manual VBE compile was not performed by executor.
- **No implicit actions:** do not run `-SyncExports` or `-DeployToDataTheiss`; do not run macros on production/original data.

### 8.2 User-performed final verification (required gate)
- **L2:** open the built/checked workbook in Excel and run **Debug → Compile VBAProject**; record pass/fail and environment context.
- **L3:** run `Create_report` (and any relevant scenarios) on **copies** of data/workbooks only, validating DoR-05 acceptance scenarios with emphasis on R-01/R-02/R-04 fidelity and compatibility in the target Excel/Windows environment.

### 8.3 Suggested L3 focus matrix (minimum)
- Single-subreport / non-combined regression (no behavior drift).
- Multi-subreport combine in order with `gap_rows` (first ignored), including multi-group and per-output isolation.
- Naming: `base, base1..N-1`, case-insensitive grouping by final resolved/sanitized name, suffix-aware 31-char errors (effective base, `31 - suffix_length`, multi-digit suffix), existing/planned collisions with no destructive resolution.
- Limits/content: per-subreport `report_rows_limit_general` (base and temps), D-05 valid-content-range (including valid one-row not rejected solely by `UsedRange.Rows.Count = 1`), projected Excel row/column limit fail before paste.
- Transfer fidelity: values not formulas, formats/number formats, row heights, hidden rows and row grouping/outline levels restored at new positions; no column widths/hidden columns/zoom/freeze/split transferred; base layout/display preserved; no `paste_row/paste_col` reapply in combine.
- Lifecycle: no temp sheets remain after successful combine; final workbook/save/cleanup guards behave correctly.

## 9. Documentation impact (DoR-08)

| Document | Planned update |
|---|---|
| `docs/01-vba-reporting/configuration-model.md` | Add `gap_rows` (`data_files`) and `report_rows_limit_general` (`general`) schema/validation/default/mandatoriness semantics; clarify naming/collision/length-error behavior as specified |
| `docs/01-vba-reporting/pipeline-and-dataflow.md` | Document two-stage prepare-then-combine flow, grouping key/order, Stage 2 placement/limits/transfer profile, temp lifecycle and finalization guards |
| `docs/01-vba-reporting/overview.md` | Reflect user-visible behavior/capabilities and constraints impacted by P01 |
| Related governed docs (only if impacted) | Keep consistent with implemented behavior; avoid unrelated changes |

## 10. Deliverables and completion criteria for this planning phase

- **Delivered:** this execution plan at `.ai/plans/01-combine-subreports-worksheet-name-validation.md`, including **§11 B1 detailed implementor task**.
- **Planning completeness:** logical blocks, dependencies/order, risk assessment (incl. requested compatibility dimensions), alternatives, per-block agent requirements, verification/handoff strategy, and B1 implementor-ready specification defined.
- **B1 status:** **Planned — Validator re-review required before implementation** (§11). Do not begin B1 code changes until Validator re-review confirms the corrected design.
- **No code/audit modifications performed:** planning only.
- **Implementation may proceed** for subsequent blocks following B0–B6 without further PO clarification (business decisions resolved), subject to required reviews/gates and while explicitly addressing technical uncertainties in §7.2 during B4/B5 with required agent support and user L2/L3 final verification.

## 11. B1 — Detailed plan and implementor task (Config model extension)

- **Block:** B1 — Configuration model extension and validation
- **Plan status:** **Planned — Validator re-review required before implementation**
- **Review gate:** **Validator review required before implementation** (this §11 is the review artifact; implementation must not start until Validator approves the corrected design)
- **Spec/decision baseline:** D-01/CR-07/CR-08 (`gap_rows`), D-02/CR-15 (`report_rows_limit_general`), A-03, DoR-03
- **Scope:** `mko_treport/src/vbs/export/modTreportConfig.bas` only (config contract and validation). **No doc edits, no combine/naming/runtime/orchestration logic changes in B1.**

### 11.1 Implementor task (YAML-style, template-aligned)

```yaml
id: B1_config_model_extension_gap_rows_and_report_rows_limit_general
title: "B1 — Extend Treport config model/validation for gap_rows and report_rows_limit_general"
priority: high
depends_on:
  - B0
source_reference: .ai/plans/01-combine-subreports-worksheet-name-validation.md
source_section: "§11 B1 — Detailed plan and implementor task (Config model extension)"
review_gate:
  validator_review_required_before_implementation: true
  status: "Planned — Validator re-review required before implementation"

description: >
  Implement P01 config-model extensions in modTreportConfig.bas: add optional data_files.gap_rows
  parsing/validation/defaulting (D-01) and required general.report_rows_limit_general parsing/validation
  with explicit presence tracking and no fallback (D-02). Preserve existing behavior/diagnostics for all
  current fields and keep data_files RowIndex/order intact. Do not implement combine/runtime naming logic
  or edit docs in this task.

goals:
  - "Add TreportDataFile.gapRows (Long) and TreportGeneral.reportRowsLimitGeneral (Long) to config types."
  - "Parse/validate report_rows_limit_general as a required, presence-tracked, positive integer in general with contextual fail-fast diagnostics."
  - "Parse/validate gap_rows as an optional data_files column (absent => 0 for all rows; present => per-row blank/null/empty => 0, else integer >= 0) with contextual fail-fast diagnostics."
  - "Preserve all existing config fields, validation behavior, RowIndex/order, and existing error-code/message conventions (no renumbering/changes to existing codes)."

non_goals_regression_guards:
  - "No changes to group_rows semantics, storage, parsing behavior, diagnostics, or comments."
  - "No changes to existing required/optional fields, accepted values, defaults, or validations (except additive introduction of report_rows_limit_general as required by D-02)."
  - "No doc file edits in B1 (configuration-model.md/pipeline-and-dataflow.md/overview.md remain unchanged in this block)."
  - "No combine, naming, grouping, worksheet creation, transfer, lifecycle, or runtime logic changes."
  - "No implicit -SyncExports / -DeployToDataTheiss; no VBA artifact/workbook sync/deploy actions."

files:
  - path: mko_treport/src/vbs/export/modTreportConfig.bas
    semantic_targets:
      - type: udt
        name: TreportGeneral
        intent: "Add Long field reportRowsLimitGeneral for required per-subreport limit (D-02)."
      - type: udt
        name: TreportDataFile
        intent: "Add Long field gapRows for per-row combine spacing (D-01), stored as deterministic Long-compatible value."
      - type: function
        name: ReadConfig
        intent: >
          Extend general-parameter parsing/validation (presence-tracked required report_rows_limit_general,
          duplicate detection, contextual fail-fast) and data_files row parsing (optional gap_rows column
          detection, per-row default/validation, preserve RowIndex/order and all existing field handling).
      - type: helper_optional
        name: "private integer parsing/validation helper(s) (implementor may add if consistent with module style)"
        intent: >
          Centralize strict integer parsing/range checks for new parameters to apply uniform rules and
          diagnostics; use Fail/DataFilesParamContext conventions and do not alter existing helpers' behavior.

implementation_scope_and_sequence:
  - step: 1
    target: "TreportGeneral (UDT)"
    action: "Add member: reportRowsLimitGeneral As Long."
    notes: "No default/fallback semantics implied; value must come from validated required general parameter."
  - step: 2
    target: "TreportDataFile (UDT)"
    action: "Add member: gapRows As Long."
    notes: "Deterministic Long-compatible storage; defaulted/validated per D-01 and row parsing rules below."
  - step: 3
    target: "ReadConfig > general section (local state, loop, post-loop validation)"
    action: >
      Introduce explicit presence/validity tracking for parameter 'report_rows_limit_general' (e.g. seen flag
      and/or parsed-valid flag as implementation detail). In the general parameter Select Case, add case
      "report_rows_limit_general" that: detects duplicate provision, validates/parses the value per strict
      rules, stores Long value on cfg.General.reportRowsLimitGeneral, and marks provided/valid state. After
      the general loop (and after existing required-field checks as appropriate), fail if the parameter was
      not provided/valid. Keep rows_collapsed parsing/diagnostics and all existing general parameters unchanged.
  - step: 4
    target: "ReadConfig > data_files section (pre-loop column detection and per-row assignment)"
    action: >
      Detect optional 'gap_rows' column once using HasColumn(loD, \"gap_rows\") (consistent with existing
      optional-column handling). For every data_files row, set cfg.DataFiles(i).gapRows = 0 by default;
      if column is present, read the cell and parse/validate per-row: blank/Empty/Null => 0, else strict
      integer >= 0 => store Long, else fail with DataFilesParamContext-style diagnostic. Preserve
      cfg.DataFiles(i).RowIndex assignment, existing required/optional field reads, and group_rows handling exactly.
  - step: 5
    target: "self-check"
    action: >
      Verify no existing error codes/messages/branches changed, no line-number assumptions introduced, and all
      new failures use Fail with contextual strings aligned to existing conventions.

exact_contract:
  report_rows_limit_general:
    location: "settings/general parameter row: parameter='report_rows_limit_general', value=<configured>"
    target_field: "TreportGeneral.reportRowsLimitGeneral As Long"
    required: true
    presence_tracking: >
      Required. ReadConfig must detect that the parameter was provided and parsed as valid (presence/seen and
      validity tracked explicitly; implementation mechanism is local detail). A row that exists but has an
      invalid value does not satisfy the requirement.
    missing_parameter_row: "Fail (no fallback/default)."
    blank_empty_error_or_invalid_value: "Fail with clear context identifying parameter and row/section."
    accepted_values: "Explicit positive integer (> 0), Long-compatible, no fractional component."
    no_overrides_no_fallback: true
    duplicate_parameter_rows: >
      Fail-fast as invalid/ambiguous configuration (recommended deterministic behavior); include parameter and
      row/section context for the duplicate occurrence.

  gap_rows:
    location: "settings/data_files optional column 'gap_rows'"
    target_field: "TreportDataFile.gapRows As Long"
    column_absent: "Default gapRows = 0 for all data_files rows."
    column_present_per_row:
      default_to_zero: "Blank/empty/null/IsEmpty/IsNull cell => gapRows = 0."
      accepted_values: "Integer value >= 0 (0 allowed), Long-compatible, no fractional component."
      rejected_values: "Negative, non-integer/fractional, non-numeric text, error, overflow/out-of-range, Boolean-typed."
    storage: "Store deterministic Long value (CLng of validated whole number) per row."
    row_order_metadata: "Preserve existing data_files iteration order and cfg.DataFiles(i).RowIndex for every row."

integer_strictness_and_edge_cases:
  common_rules:
    - "Boolean-typed cells (VarType = vbBoolean) MUST be rejected for both parameters (do not coerce True/False)."
    - "Error values (IsError) MUST be rejected."
    - "Accept numeric/integer-valued inputs only when value is finite, within VBA Long range, and has no fractional component (e.g. 1 and 1.0 are acceptable as integer value 1 where representable); store as Long."
    - "Reject fractional values (e.g. 1.5), non-numeric text (e.g. \"abc\"), negative values where not allowed, and overflow/unrepresentable values."
    - "Apply rules uniformly; implementation may use CDbl/IsNumeric prechecks but must enforce whole-number, range, and type constraints explicitly."
  report_rows_limit_general:
    - "0 => reject (must be positive)."
    - "1 => accept."
    - "Large positive integer within Long range => accept."
    - "1.0 / integer-valued numeric cell => accept as 1 (if Long-compatible, no fractional component)."
    - "-1 => reject."
    - "1.5 => reject."
    - "\"abc\" => reject."
    - "#VALUE! (or any error cell/value) => reject."
    - "Boolean True/False => reject."
    - "Blank/empty/null/IsEmpty/IsNull when parameter row exists => reject (required value)."
  gap_rows:
    - "0 => accept."
    - "1 => accept."
    - "Large non-negative integer within Long range => accept."
    - "1.0 / integer-valued numeric cell => accept as 1 (if Long-compatible, no fractional component)."
    - "-1 => reject."
    - "1.5 => reject."
    - "\"abc\" => reject."
    - "#VALUE! (or any error cell/value) => reject."
    - "Boolean True/False => reject."
    - "Blank/empty/null/IsEmpty/IsNull when gap_rows column present => default to 0 (not an error)."
    - "gap_rows column absent => default to 0 for all rows."

diagnostics_and_error_codes:
  preserve_existing:
    - "Do not renumber or change behavior/messages/codes of existing validations (2100-2107, 2109-2123, 2118 helpers/usage, etc.)."
  report_rows_limit_general:
    code: "2108 (recommended, currently unused/conflict-free in modTreportConfig.bas and docs search baseline at planning time; use consistently for missing/invalid/duplicate/non-positive cases of this parameter)"
    context: "config/general"
    missing_row_message_style: "parameter 'report_rows_limit_general' is required and was not found."
    invalid_or_duplicate_message_style: >
      Include parameter name, raw value when applicable, general-row index (row=<i>, section='general')
      consistent with rows_collapsed diagnostics, and expectation (explicit positive integer / not duplicated).
    examples_intent:
      - "missing: Fail 2108, \"config/general\", \"parameter 'report_rows_limit_general' is required and was not found.\""
      - "invalid: Fail 2108, \"config/general\", \"parameter 'report_rows_limit_general' has invalid value '...' (row=<i>, section='general'); expected a positive integer (>= 1).\""
      - "duplicate: Fail 2108, \"config/general\", \"parameter 'report_rows_limit_general' is duplicated (row=<i>, section='general'); expected exactly one valid entry.\""
  gap_rows:
    code: "2118 (reuse/align with existing data_files parameter validation convention)"
    context_helper: "DataFilesParamContext(rowIdx/data row index, dataFileName, dataSheetName) — same style as template/only_values/insert_rows/skip_header"
    message_style: >
      Include parameter name 'gap_rows', raw/observed value, and expectation (blank or integer >= 0).
    example_intent:
      - "Fail 2118, DataFilesParamContext(...), \"parameter 'gap_rows' has invalid value '...' ; expected blank or an integer >= 0.\""

acceptance_criteria:
  - "TreportGeneral contains reportRowsLimitGeneral As Long; TreportDataFile contains gapRows As Long."
  - "report_rows_limit_general: missing parameter row => fail 2108 (config/general); present but blank/empty/non-integer/non-positive/fractional/non-numeric/error/Boolean/overflow/duplicate => fail 2108 with row/section context consistent with rows_collapsed style."
  - "report_rows_limit_general: valid positive integer (including integer-valued numeric cells like 1.0, Long-compatible) is stored as Long; no default/fallback."
  - "gap_rows column absent => every data_files row has gapRows = 0."
  - "gap_rows column present => blank/empty/null/IsEmpty/IsNull => 0; valid integer >= 0 (including 0 and integer-valued numeric cells) => stored as Long; invalid values fail 2118 with DataFilesParamContext."
  - "All existing data_files fields, required validations, RowIndex assignment, and row order remain unchanged."
  - "group_rows behavior/storage/parsing/diagnostics/comments remain unchanged."
  - "Existing general parameters (input_path, output_path, delimeter, rows_collapsed) and their diagnostics/error codes remain unchanged."
  - "No new/changed behavior outside modTreportConfig.bas contract extensions specified above."

verification_notes:
  - "Static review against this §11 contract and modTreportConfig.bas existing conventions."
  - "Executor may run L1 build on a safe copy if code is implemented later; L1 is best-effort and not proof of compile."
  - "User performs final L2 (Debug → Compile VBAProject) and L3 from workbook on copies (EVD-02)."
  - "No -SyncExports / -DeployToDataTheiss as part of this task."
```

### 11.2 Non-goals / regression guard (explicit, B1)

- **No changes** to `group_rows` semantics, comment, parsing, defaulting, diagnostics, or runtime interpretation hooks.
- **No changes** to existing required/optional `general`, `reports`, `data_files`, or `placeholders` fields/columns/validations except the additive required `report_rows_limit_general` parameter per D-02.
- **No doc file edits** in B1 (`docs/01-vba-reporting/configuration-model.md`, `pipeline-and-dataflow.md`, `overview.md` unchanged by this block).
- **No combine/naming/grouping/runtime/workbook lifecycle logic** changes in B1 (B2–B5 scope).
- **No implicit** export sync or data deploy actions.

### 11.3 Done / acceptance checklist (B1)

- [ ] §11 reviewed and approved by **Validator** before implementation.
- [ ] `TreportGeneral.reportRowsLimitGeneral As Long` added; `TreportDataFile.gapRows As Long` added.
- [ ] `ReadConfig` parses `report_rows_limit_general` with explicit presence/validity tracking; missing row fails; invalid/blank/duplicate/non-positive/non-integer/fractional/error/Boolean/overflow cases fail with **2108** and `config/general` + row/section context as specified.
- [ ] `ReadConfig` handles `gap_rows` column absent => all rows default `0`; column present => per-row blank/Empty/Null defaults `0`, valid `>= 0` integer stored as `Long`, invalid cases fail with **2118** and `DataFilesParamContext` as specified.
- [ ] `RowIndex` and existing `data_files` row processing/order preserved; `group_rows` and all existing fields/diagnostics unchanged.
- [ ] No existing error codes/messages renumbered or altered; no out-of-scope files/logic changed.
- [ ] Static/self-review complete; any L1/L2/L3 evidence recorded per project verification rules and EVD-02 ownership split.

## 12. B2 — Detailed plan and implementor task (Name resolution, grouping key, suffix-aware planning, and collision preflight)

- **Block:** B2 — Name resolution, grouping key, suffix-aware planning, and collision preflight
- **Plan status:** **Planned — Validator re-review required before implementation** (amended with mandatory Validator corrections **RC-1..RC-5**)
- **Review gate:** **Validator re-review required before implementation** (this §12 is the binding review artifact; implementation must not start until Validator approves the corrected RC-1..RC-5 design and B2/B3 boundary decisions)
- **Spec/decision baseline:** D-03/CR-01/CR-02/CR-03 (grouping key, per-output isolation, order), D-04/CR-19n–CR-21n (suffix-aware length validation/messages), D-06/CR-06/CR-22g (planned/existing collision preflight, no destructive resolution, default-sheet conflict rule), OQ-7/CR-05 (sequential suffixes `1..N-1` in original `data_files` order, actual suffix length), CR-07 context (per-output grouping structure), CR-08 (per-member `gap_rows`), CR-18c/DoR-04/05 relevant parts (lifecycle/registry integration contracts for B3–B5)
- **Dependencies:** B0, B1 (implemented in `modTreportConfig.bas`, commit `3f653e4`: `TreportDataFile.gapRows`, `TreportDataFile.RowIndex`, `TreportGeneral.reportRowsLimitGeneral`; template/order preserved)
- **Risk level:** Medium (logic/edge-case/coupling risk); can escalate if coupled with lifecycle refactor. Requires careful integration with 2612/2634/cleanup/registry semantics.

### 12.1 Objective and B2/B3 boundary (explicit)

**Objective:** Deliver deterministic, spec-compliant naming/group planning and preflight primitives plus per-output plan contracts so that, for every `(data_file row, report)` emission in the same output workbook: (1) the effective base worksheet name is derived via **resolve placeholders/substitutions -> sanitize** (D-03), (2) members are grouped by **case-insensitive final resolved+sanitized base name** within the same output workbook in original `data_files` emission order (D-03/CR-01–CR-03/OQ-7), (3) planned worksheet names are assigned as `base, base1..base(N-1)` using actual suffix length (OQ-7/CR-05), (4) every generated/planned name is validated for Excel's 31-char limit with suffix-aware D-04 diagnostics, and (5) collisions against **existing output-workbook worksheet names and all names planned for creation in the current run** are detected before affected worksheet creation with no overwrite/delete/rename (D-06/CR-06).

**B2 delivers (implementation scope):**
- `modTreportOutput.bas`: name-resolution/sanitization helpers for planning, grouping-key helper, suffix-aware planned-name generation/validation, collision preflight helpers, and documented plan/registry data contracts (dictionary schemas and/or Public types as needed for cross-module use).
- `modTreportRun.bas` (minimal integration scaffolding necessary for correctness): introduce per-output planning state and a **complete pre-creation naming/planning pass** that resolves base names for all relevant emissions, builds groups/planned names via B2 APIs, and runs B2 full preflight before **any** Stage 1 report worksheet is created; enforce the **RC-1 multi-member gate**; preserve existing non-combined single-member behavior and existing failure semantics except where P01 explicitly changes them (grouping/planning, D-04/D-06 diagnostics/codes, registry handling, 2612 evolution per RC-4). B2 must not implement Stage 2 combine transfer/lifecycle finalization (B4/B5).

**B3 consumes/extends:** validated B2 plans to perform Stage 1 preparation for multi-member/combine groups on assigned base/temp names, maintain separate planned/temp/final registries for 2630/cleanup/verify/save, execute Stage 2 combine hooks, and complete lifecycle/registry refactor. B3 owns end-to-end "prepare all, then combine" sequencing and temp/final lifecycle for combine groups.

**Boundary (binding, RC-1):** B2 must **not** create or prepare any member report worksheet for any group with `GroupSize > 1` / `IsCombineGroup = true`. If any such group exists after complete resolve/collect and successful plan/finalize/preflight for all affected outputs, B2 must fail-fast with a contextual diagnostic **before creating any Stage 1 report worksheet**. "Safe Stage 1-only preparation" for multi-member groups and mid-execution gating after creation are **not permitted** in B2 (B4/B5 lifecycle-safe completion is out of scope). Only when **all** groups across all affected outputs are single-member may B2 proceed to Stage 1 execution.

**Boundary/residual dependency (RC-5):** D-06 requires all planned names checked before creating affected worksheets. Because base names depend on per-row pivot-resolved placeholders, local per-emission checks are insufficient. **B2 implementation must use explicit sequencing:** create/make output workbook objects available for existing-name inspection as needed, but perform **no planned report worksheet creation** until the **complete per-run/per-output plan set is resolved, length-validated, and collision-preflighted** (recommended: full `data_files`/report scan -> build/preflight all output plans -> execute supported single-member-only path or fail per RC-1). Any inability to prove full preflight coverage for a specific edge path is a blocking uncertainty for B2 completion and must be escalated to Validator/Auditor (not silently deferred).

### 12.2 Current-state constraints (semantic baseline)

- `modTreportRun.ProcessDataFile` currently: resolves placeholder max values per data file, applies filters, then per `(dfIdx, report r)` substitutes placeholders -> `SanitizeSheetName` -> fails 2612 if `outUsedSheetNames(r)` already contains `LCase$(sheetName)` -> registers key -> computes placement -> `PrepareTargetSheet` -> `CopyPivotSnapshot` -> formatting/grouping. `Create_report` owns per-output `outBooks/outNames/outUsedSheetNames` and final cleanup/verify/save guards (2630/2620–2629).
- `modTreportOutput.SanitizeSheetName` normalizes chars `:\/?*[]` and `'` to `_`, fails 2601 (empty) and 2602 (`> 31` showing sanitized name only, not suffix-aware). `PrepareTargetSheet` fails 2634 if target exists (no overwrite/delete), copies template from `ThisWorkbook` (case-insensitive lookup) or creates plain sheet. `CleanupDefaultSheets` deletes unproduced worksheets by `LCase$(Trim$(ws.Name))` not in produced keys.
- `outUsedSheetNames(r)` is currently dual-purpose (collision namespace + produced/final registry). B2 must evolve collision/grouping semantics without breaking existing guards and must satisfy RC-3 final-produced registry coherence on the single-member execution path.

### 12.3 Detailed design

#### 12.3.1 Naming pipeline and grouping key (D-03, CR-01–CR-03)

- **Base-name derivation (per member/emission):** `effectiveBaseName = SanitizeForPlanning(SubstituteReportSheetPlaceholders(reportSheetTemplate, resolvedMaxValues, ctx), ctxBase)` where sanitization applies the same normalization/replacement rules as current `SanitizeSheetName` (trim/norm, replace `: \ / ? * [ ]` and `'` with `_`). Empty result => **2601** (preserve existing fail-fast semantics). Do **not** apply non-suffix-aware `>31` rejection as the final D-04 check; length is validated against the concrete planned name (base or suffixed) per §12.3.2 using **2637**.
- **Grouping key:** `groupKey = LCase$(effectiveBaseName)` (case-insensitive, Excel-consistent textual key) scoped to the **same output workbook/report output index** `r`. Identical final resolved+sanitized base names across different output workbooks remain isolated and allowed.
- **Member identity/order:** members are grouped in deterministic original emission order: for each `dfIdx` in `cfg.DataFiles` order, for each `r` in `cfg.Reports` order (same iteration order as current processing), the emission `(dfIdx, r)` belongs to group `groupKey` for output `r` if its effective base name maps to that key. Preserve existing per-output isolation and placeholder resolution timing basis (per data_files row/pivot context, pre-`data_files.filters` application, as today); any timing refactor must not change resolved values/semantics.

#### 12.3.2 Suffix assignment and suffix-aware length validation (D-04, CR-19n–CR-21n, OQ-7, CR-05) — RC-4

- For a group with `N` members in original emission order, assign planned worksheet names by zero-based member position `k = 0..N-1`:
  - `k == 0` => `plannedName = effectiveBaseName` (base, no suffix)
  - `k >= 1` => `suffix = CStr(k)`; `plannedName = effectiveBaseName & suffix` (suffixes `1..N-1`)
- **Actual suffix length** is `Len(suffix)` for that member (`0` for base). Multi-digit suffixes must be handled by actual length.
- **Validation (RC-4a):** all P01-planned worksheet-name length violations (base-only or suffixed, `suffix_length >= 0`) must emit D-04-compliant diagnostics via **2637** (confirm unused at implementation time; if conflict found, use an explicitly justified equivalent unused code without renumbering existing codes). **2601** remains for empty-after-normalization. Legacy **2602** must not be the path for planning-validated names when its message is non-compliant with D-04.
- **2637 required message fields (D-04):** identify/report the **effective base name after substitution and sanitization without the numeric suffix** (`effectiveBaseName`); compute and explicitly state **maximum permitted base-name length** as `31 - Len(suffix)` for the failing generated name; include `suffix_length`/`suffix` and concrete planned name; include output/report/data-file/row context; **raw template name** (`report_sheet` template before substitution/sanitization) may be included as additional diagnostic context when useful.
- A base-only name (`N==1` or `k==0`) that exceeds 31 is a length violation with `suffix_length = 0`, max permitted base length `31`.
- `SanitizeSheetName` may be retained as a compatibility wrapper only if it does not emit legacy-only 2602 diagnostics for planning-validated names; otherwise refactor callers to the planning-aware validation path.

#### 12.3.3 Collision preflight (D-06, CR-06, CR-22g)

- Preflight scope is **per output workbook** and must run **before creating any affected Stage 1 worksheet** for that output's validated plan (and globally before any execution per RC-5).
- **Inputs:** output workbook (`wb`), complete set/list of concrete planned worksheet names for the run in that output (including base and all suffixed names across all groups), and source/group/member context for diagnostics.
- **Checks (case-insensitive, Excel-consistent):**
  1. **Existing worksheet conflict:** for each planned name, if any `wb.Worksheets` has name equal to planned name by case-insensitive comparison (`StrComp(..., vbTextCompare)=0` or equivalent lookup over all worksheets), fail identifying the conflicting worksheet name. This includes default sheets (e.g. `Sheet1`) — conflicting defaults must **not** be auto-deleted/renamed/overwritten; preserve existing cleanup behavior for unused default worksheets **only when they do not conflict** with planned report names.
  2. **Planned-vs-planned conflict:** if any two planned concrete names in the same output workbook collide case-insensitively (cross-group or within-plan anomaly), fail identifying the conflicting worksheet name and involved planning contexts.
- **Behavior:** fail-fast, no worksheet creation/deletion/overwrite/rename as conflict resolution. Preflight helper must be pure validation over provided inputs/workbook state (no side effects).
- **Error code policy (RC-4):** use dedicated unused code **2638** for planned/existing worksheet name collision preflight failures (D-06), with message identifying conflicting worksheet name, conflict class (`existing` vs `planned`), output context, and source/group/member context(s). Confirm unused at implementation time.

#### 12.3.4 Plan artifacts and registry contracts (B2 output, B3/B4/B5 input) — RC-3

B2 must define and produce deterministic, per-output planning artifacts. Preferred representation: `Scripting.Dictionary` / `Collection` objects with documented keys; Public Type(s) may be added in `modTreportOutput.bas` only if they improve clarity without VBA cross-module friction.

**Per-output plan (`OutputPlan`) minimum contract:**
- `OutputIndex` / report identity context (`r`, `outNames(r)`, report row/name)
- `Members`: ordered collection/list of member records in original emission order `(dfIdx asc, report r fixed per output)`; each member record contains:
  - source identity: `DfIdx`, `DataFileRowIndex` (`RowIndex`), `DataFileName`, `DataSheet`, `ReportSheetTemplate` (raw), `ReportRowIndex`/report name/context as applicable, `SourceCtx`/`RepCtx`
  - planning inputs: `GapRows` (Long, from B1), `ResolvedMaxValues` reference/keying context if needed for traceability (not required to persist object if base already resolved), `EffectiveBaseName`, `GroupKey` (`LCase$(EffectiveBaseName)`)
  - assignment (filled by group planning): `PlannedSheetName`, `SuffixIndex` (0-based in group), `GroupSize`, `IsBaseMember`, `GroupKey`/group ref
- `Groups`: dictionary keyed by `GroupKey` -> group record:
  - `GroupKey`, `EffectiveBaseName`, `RawTemplateExample` (optional), `Members` (ordered references/indices in group order), `GroupSize`, `PlannedNames` (ordered), `IsCombineGroup` (`GroupSize > 1`)
- `PlannedNameRegistry`: case-insensitive dictionary keyed by `LCase$(PlannedSheetName)` -> conflict/info record (`PlannedSheetName`, `GroupKey`, member ref/context, class=`planned`) used for D-06 planned-vs-planned and preflight reporting
- `Validated` / status flags as needed for B3 handoff (implementation detail)

**Registry coherence (binding, RC-3):**
- B2 planning registry is not the same as the final produced/base registry consumed by existing `2630` / `CleanupDefaultSheets` / `EnsureProducedSheetsPresent` / save checks, but B2 must **not** defer coherence.
- On the **single-member execution path** (RC-1 allows execution only when all groups are single-member), B2 must register each created final/base report sheet in the registry/structure consumed by `2630`, `CleanupDefaultSheets`, `EnsureProducedSheetsPresent`, and save checks in the **equivalent final-produced role** as today (case-insensitive keying consistent with existing cleanup/verify behavior), so guards remain correct.
- Keep B3/B5 extension points explicit for separate planned/reserved, transient temp, and final/base produced registries needed for combine lifecycle; document which structure carries final-produced keys in B2 and how B3/B5 must extend/split it without regressing guards.

#### 12.3.5 2612 evolution and grouping semantics — RC-4

- Current unconditional 2612 "sheet name collision in output workbook..." on duplicate final name in the same output becomes **grouping-aware** (RC-4b):
  - Same `GroupKey` (same effective resolved+sanitized base name, case-insensitive) within same output => same combine group, append member in original order (**not** an error by itself).
  - **2612 must not trigger** for same effective-base grouping; the duplicate-final-name 2612 path must be removed or made unreachable for that case and must not be repurposed for grouping.
  - Concrete planned worksheet name collisions (base or suffixed) against existing workbook names or other planned names => **2638 preflight failure** per D-06 (not silent merge, not destructive resolution).
- Single-member groups (`N==1`) must preserve current observable behavior as closely as possible: planned/final name is `EffectiveBaseName`, no suffix, no combine. Any intentional deviation required by spec (e.g. D-04 diagnostics for base-too-long) must be explicit and validated.
- B2 does not create/maintain temp lifecycle; it only assigns deterministic names/metadata needed by B3.

#### 12.3.6 Pre-creation sequencing, RC-1 multi-member gate, and RC-2 execution-pass fidelity (binding)

To satisfy D-06 and RC-1/RC-2/RC-3/RC-5, B2 must separate name resolution/plan construction/preflight from Stage 1 worksheet creation with explicit sequencing:

1. **Output workbook availability (inspection only, no planned report sheets):** create or obtain per-output workbook objects as needed so existing worksheet names can be inspected for D-06 preflight. This may follow existing output-workbook creation semantics, but **no planned report worksheet may be created** at this stage.
2. **Resolve/collect phase (full run scan, no report-sheet creation):** iterate `cfg.DataFiles` in original order and `cfg.Reports` in original order consistent with current behavior; for each emission resolve placeholder max values and compute `EffectiveBaseName`/`GroupKey` using existing pivot/filter-resolution prerequisites required for correct placeholder substitution (placeholder max resolution basis must remain equivalent to current pre-`data_files.filters` behavior), collect member records into per-output plans including `GapRows` and contexts. Do not call `PrepareTargetSheet`, `CopyPivotSnapshot`, formatting/grouping, or mutate output workbooks by creating report sheets in this phase. Preserve read-only source handling discipline.
3. **Plan/finalize/preflight phase (complete set, no creation):** for **all** outputs, build groups in preserved member order, assign suffix-aware planned names, validate lengths (**2637**/2601 as applicable), build planned-name registries, run D-06 preflight (**2638**) against existing output workbook worksheets and all planned names. Any failure aborts before creating any affected report sheet.
4. **RC-1 multi-member gate (before any Stage 1 report-sheet creation):** after complete resolve/collect and successful plan/finalize/preflight for all affected outputs, if **any** group has `GroupSize > 1` / `IsCombineGroup = true`, fail-fast with contextual diagnostic including at minimum output/report name/index, effective base/group, member count, and source contexts as available, stating that B3/B4/B5 combine implementation is required — **before creating any Stage 1 report worksheet**. **Do not** create/prepare any member sheet for multi-member groups in B2. No mid-execution gating after creation is permitted.
5. **Execution phase (only if all groups are single-member, RC-2/RC-3):** proceed to Stage 1 execution for the supported single-member-only path. Execution must preserve existing Stage 1 semantics/order for single-member groups:
   - iterate `cfg.DataFiles` and `cfg.Reports` in original order consistent with current behavior;
   - for each supported emission reopen the source workbook read-only as needed, obtain sheet/pivot, apply data-file/report filters in the same order, enforce 2410 no-data checks, compute placement, prepare the target sheet using the **B2-assigned planned name**, call `CopyPivotSnapshot`, `ApplySheetFormatting`, `ApplyOutlineGrouping`, and close the source;
   - preserve existing context/error patterns except intentional P01 changes (grouping/planning, D-04/D-06 diagnostics/codes, registry handling);
   - placeholder max resolution basis must remain equivalent to current pre-`data_files.filters` behavior (resolved values must not change due to pass separation);
   - register each created final/base report sheet in the final-produced registry consumed by `2630`/`CleanupDefaultSheets`/`EnsureProducedSheetsPresent`/save checks (RC-3).
6. **Finalization:** existing cleanup/verify/save guards run against the B2 final-produced registry for the single-member-only path; do not claim B4/B5 completion.

**Preferred/required sequencing (RC-5):** full `data_files` scan -> build/preflight all output plans -> execute supported single-member-only path or fail per RC-1. "Stage 1-only preparation for multi-member groups" and "mid-execution gating after creation" are explicitly disallowed in B2.

### 12.4 Implementor task (YAML-style, template-aligned)

```yaml
id: B2_name_resolution_grouping_suffix_planning_collision_preflight
title: "B2 — Name resolution, grouping key, suffix-aware planning, and collision preflight"
priority: high
status: "Planned — Validator re-review required before implementation (RC-1..RC-5 amended)"

depends_on:
  - B0
  - B1  # implemented in modTreportConfig.bas (commit 3f653e4)

source_reference: .ai/plans/01-combine-subreports-worksheet-name-validation.md
source_section: "§12 B2 — Detailed plan and implementor task (Name resolution, grouping key, suffix-aware planning, and collision preflight)"
spec_reference: .ai/problems/02_combine_subreports_worksheet_name_validation_spec.md
spec_decisions: [D-03, D-04, D-06, OQ-7, CR-01, CR-02, CR-03, CR-05, CR-06, CR-07, CR-08, CR-19n, CR-20n, CR-21n, CR-22g]

review_gate:
  validator_re_review_required_before_implementation: true
  mandatory_corrections_incorporated: [RC-1, RC-2, RC-3, RC-4, RC-5]

description: >
  Implement P01 B2 naming/group planning and preflight in modTreportOutput.bas with minimal, necessary
  integration in modTreportRun.bas. Deliver deterministic base-name resolution (placeholder substitution +
  sanitization), case-insensitive per-output grouping in original emission order, suffix-aware planned-name
  assignment (base, base1..base(N-1), actual suffix length), D-04 length diagnostics (2637), D-06
  planned/existing collision preflight (2638) before affected worksheet creation, and per-output plan/registry
  contracts consumable by B3. Refactor run flow to explicit sequencing: output workbooks available for
  inspection -> full scan resolve/collect -> complete per-run/per-output plan/finalize/preflight -> RC-1
  multi-member gate before any Stage 1 report-sheet creation -> single-member-only execution with RC-2 semantic
  fidelity and RC-3 final-produced registry coherence, or fail per RC-1. Do not implement Stage 2 combine
  transfer/lifecycle finalization (B4/B5). No documentation edits.

goals:
  - "RC-5/RC-1: guarantee no planned report worksheet creation until complete per-run/per-output plans are resolved, length-validated, and preflighted; if any IsCombineGroup exists, fail-fast before any Stage 1 report-sheet creation."
  - "D-03/CR-01..03: derive effective base name via substitute -> sanitize and group by case-insensitive key scoped to same output workbook in original emission order."
  - "D-04/OQ-7/CR-05/CR-19n..21n: assign base, base1..base(N-1) with actual suffix length; validate all planned names (suffix_length >= 0) with D-04-compliant diagnostics via 2637 (2601 for empty)."
  - "D-06/CR-06/CR-22g: preflight existing and all planned names (case-insensitive, incl. defaults) before affected creation; fail 2638 identifying conflicting name/class/context with no overwrite/delete/rename."
  - "RC-4: 2612 must not trigger for same effective-base grouping (remove/make unreachable for that case); 2637/2638 confirmed unused at implementation time; existing codes/messages otherwise unchanged."
  - "RC-2: single-member execution pass preserves existing Stage 1 semantics/order (reopen source read-only as needed, sheet/pivot, filters in same order, 2410, placement, PrepareTargetSheet with B2-assigned name, CopyPivotSnapshot, ApplySheetFormatting, ApplyOutlineGrouping, close source) except intentional P01 changes; placeholder max resolution basis remains equivalent to current pre-data_files.filters behavior."
  - "RC-3: single-member execution registers each created final/base sheet in final-produced registry consumed by 2630/CleanupDefaultSheets/EnsureProducedSheetsPresent/save checks; B3/B5 planned/temp/final separation extension points explicit."
  - "Produce deterministic per-output plan/registry contracts (members, groups, planned names, gapRows, contexts) for B3/B4/B5."

non_goals_regression_guards:
  - "No Stage 2 combine append/transfer/Excel-bounds/content-limit enforcement (B4)."
  - "No temp deletion/final produced-registry cleanup-save lifecycle completion for combine groups (B5)."
  - "No creation/preparation of any member report worksheet for any multi-member/combine group in B2 (RC-1 binding)."
  - "No mid-execution gating after worksheet creation; no Stage 1-only preparation for multi-member groups unless B4/B5 lifecycle-safe completion is in scope (it is not)."
  - "No changes to Refresh_data, filter semantics, placeholder resolution rules, insert_rows/skip_header/only_values/template preparation semantics except strictly necessary integration seams that preserve RC-2 equivalence."
  - "No reapplication of paste_row/paste_col in combine context; no transfer of forbidden worksheet-level display/layout settings."
  - "No overwrite/delete/rename of existing worksheets for conflict resolution; no temp deletion before successful transfer."
  - "No per-data_files override or fallback for report_rows_limit_general; no application of report_rows_limit_general to combined totals."
  - "No implicit -SyncExports / -DeployToDataTheiss; no docs edits in B2."

files:
  - path: mko_treport/src/vbs/export/modTreportOutput.bas
    semantic_targets:
      - type: public_helper
        name: "effective base-name planning helper (e.g. BuildEffectiveReportSheetBaseName / ResolveAndSanitizeReportSheetBase)"
        intent: "Substitute placeholders -> sanitize using existing char rules; preserve 2601 on empty; expose raw template and effective base for diagnostics; no legacy 2602 path for planning-validated D-04 cases."
      - type: public_helper
        name: "group-key helper (e.g. MakeReportSheetGroupKey)"
        intent: "Case-insensitive key for effective base name scoped by caller to same output workbook."
      - type: public_helper_or_procedure
        name: "suffix-aware planned-name assignment/validation helpers (e.g. BuildPlannedSheetName, ValidatePlannedSheetNameLength, AssignGroupPlannedNames)"
        intent: "Generate base/base1..base(N-1) with actual suffix length; fail 2637 with D-04-compliant fields for all planned-name length violations (suffix_length >= 0)."
      - type: public_helper_or_procedure
        name: "collision preflight helper (e.g. PreflightPlannedSheetNames / ValidateNoPlannedNameCollisions)"
        intent: "Case-insensitive preflight of all planned names vs existing workbook sheets and all planned names before affected creation; fail 2638 identifying conflicting name/class/context; no side effects."
      - type: function
        name: "SanitizeSheetName (compatibility review/refactor as needed)"
        intent: "Preserve normalization/caller compatibility; ensure planning-validated names do not route to non-compliant legacy 2602 diagnostics."
      - type: contract
        name: "per-output plan/registry data contract (Public Type(s) and/or documented Dictionary schema)"
        intent: "Expose members/groups/planned-name registry and final-produced registry handoff for B2 single-member path plus B3/B5 extension points."

  - path: mko_treport/src/vbs/export/modTreportRun.bas
    semantic_targets:
      - type: procedure
        name: "Create_report"
        intent: "Own output workbook availability for inspection, per-output plan containers, full-scan planning/preflight sequencing, RC-1 gate, single-member execution/finalization integration, and final-produced registry coherence for 2630/cleanup/verify/save."
      - type: procedure_or_helper
        name: "ProcessDataFile and/or extracted resolve/collect helper(s)"
        intent: "Collect resolved per-emission members without creating report sheets; preserve placeholder max resolution basis (pre-data_files.filters), filter/source semantics inputs, and contexts."
      - type: helper_procedure
        name: "planning/finalize/preflight helper(s)"
        intent: "Build complete per-output plans, assign names, validate lengths, run D-06 preflight, and apply RC-1 gate before any Stage 1 report-sheet creation."
      - type: helper_procedure
        name: "single-member execution helper(s)"
        intent: "Execute RC-2-compliant Stage 1 path only when all groups are single-member; use B2-assigned planned names; register created sheets in final-produced registry (RC-3)."

implementation_scope_and_sequence:
  - step: 1
    target: "modTreportOutput.bas (primitives and codes)"
    action: "Implement effective base derivation, group-key, and diagnostic payload helpers; confirm 2637/2638 unused; preserve 2601 and prevent legacy 2602 path for planning-validated D-04 cases."
  - step: 2
    target: "modTreportOutput.bas (plan/registry contracts)"
    action: "Define per-output member/group/planned-name and final-produced registry contracts (GapRows, contexts, order, B3/B5 extension points)."
  - step: 3
    target: "modTreportOutput.bas (suffix-aware planning/validation)"
    action: "Implement deterministic group name assignment and 2637 D-04 validation for all planned-name length violations (suffix_length >= 0), including multi-digit suffixes."
  - step: 4
    target: "modTreportOutput.bas (collision preflight)"
    action: "Implement 2638 D-06 preflight over existing workbook worksheets and all planned run names (case-insensitive, incl. defaults), no side effects."
  - step: 5
    target: "modTreportRun.bas (RC-5 sequencing: inspect + full resolve/collect)"
    action: "Make output workbooks available for inspection; perform full data_files/report scan to collect resolved members without creating report sheets; preserve placeholder resolution basis and source handling."
  - step: 6
    target: "modTreportRun.bas (plan/finalize/preflight + RC-1 gate)"
    action: "Build/preflight all output plans; remove/make unreachable 2612 for same-base grouping; if any IsCombineGroup exists fail-fast with contextual diagnostic before any Stage 1 report-sheet creation."
  - step: 7
    target: "modTreportRun.bas (RC-2/RC-3 single-member execution only)"
    action: "If and only if all groups are single-member, execute preserved Stage 1 semantics/order using B2-assigned planned names and register created sheets in final-produced registry for 2630/cleanup/verify/save."
  - step: 8
    target: "self-check"
    action: "Verify RC-1..RC-5, D-03/D-04/D-06/OQ-7, no line numbers, no destructive conflicts, no multi-member sheet creation, and no unrelated regressions."

contract_and_diagnostics:
  sequencing:
    required_order: "output workbooks available for inspection -> full resolve/collect (no report sheets) -> complete per-run/per-output plan/finalize/preflight -> RC-1 gate -> single-member-only execution or fail"
    rc1_gate: "if any GroupSize > 1 / IsCombineGroup: fail-fast before any Stage 1 report-sheet creation with output/report name/index, effective base/group, member count, source contexts; do not create/prepare any member sheet"
  grouping_key:
    input: "effective base name after placeholder substitution and sanitization"
    output: "case-insensitive key (LCase$), scoped to same output workbook/report output"
  planned_names:
    single_member: "base"
    multi_member_order: "original data_files/report emission order -> base, base1, ..., base(N-1)"
    suffix_rule: "actual suffix length per member (multi-digit supported)"
  length_validation:
    limit: 31
    on_violation_code: 2637
    applies_to: "all P01-planned worksheet names, base-only or suffixed (suffix_length >= 0)"
    required_message_fields:
      - "effective base name after substitution and sanitization without numeric suffix"
      - "maximum permitted base length = 31 - suffix_length (explicit)"
      - "suffix_length and/or suffix and concrete planned name"
      - "output/report/data_file/row/context"
      - "raw report_sheet template as optional diagnostic context"
  collision_preflight:
    on_violation_code: 2638
    checks:
      - "existing output workbook worksheet names (all sheets, incl. defaults) vs all planned names"
      - "all planned names vs all planned names in same output run"
    required_message_fields:
      - "conflicting worksheet name"
      - "conflict class/source (existing worksheet or planned name)"
      - "output/report/group/member context"
    prohibited_resolution: "no overwrite/delete/rename"
  empty_base_name:
    code: 2601
    behavior: "preserve fail-fast empty-after-normalization semantics"
  legacy_codes:
    2602: "must not be the path for planning-validated D-04 name-length violations"
    2612: "must not trigger for same effective-base grouping; remove/make unreachable for that case; do not repurpose"
    2637_2638: "confirm unused at implementation time; all other existing codes/messages unchanged"

acceptance_criteria:
  - "RC-5: no planned report worksheets created until complete per-run/per-output plans are resolved, length-validated, and collision-preflighted."
  - "RC-1: after successful full planning/preflight, any IsCombineGroup causes fail-fast before any Stage 1 report-sheet creation with required contextual diagnostic; no member sheets created/prepared for multi-member groups in B2."
  - "RC-2: single-member-only execution preserves existing Stage 1 semantics/order (source reopen/read-only, sheet/pivot, filters same order, 2410, placement, PrepareTargetSheet with B2-assigned planned name, CopyPivotSnapshot, ApplySheetFormatting, ApplyOutlineGrouping, close source, context/error patterns) except intentional P01 changes; placeholder max resolution basis equivalent to current pre-data_files.filters behavior."
  - "RC-3: in single-member execution, each created final/base sheet is registered in final-produced registry consumed by 2630/CleanupDefaultSheets/EnsureProducedSheetsPresent/save checks; B3/B5 extension points explicit."
  - "D-03: grouping by resolved+sanitized effective base name, case-insensitive, scoped to same output, original emission order; per-output isolation preserved."
  - "D-04/OQ-7: planned names base/base1..base(N-1) with actual suffix length; all planned-name length violations (suffix_length >= 0) emit D-04-compliant 2637 diagnostics with effective base (no suffix), explicit 31 - suffix_length, and required context."
  - "D-06: preflight detects existing and planned-vs-planned collisions (incl. defaults) before affected creation and fails 2638 identifying conflicting name/class/context with no overwrite/delete/rename."
  - "RC-4: 2612 not triggered for same-base grouping; 2637/2638 verified unused at implementation time; 2601 preserved; 2602 not used for planning-validated D-04 cases; existing unrelated codes/messages unchanged."
  - "Per-output plan/registry contracts produced with ordered members, groups, planned names, gapRows, and source contexts sufficient for B3/B4/B5."

constraints:
  - "Never change VBA/code in this planning turn; this is a plan/task artifact only."
  - "No line numbers; use semantic anchors only."
  - "No docs edits in B2; no -SyncExports / -DeployToDataTheiss."
  - "No Stage 2 combine transfer/lifecycle implementation (B4/B5)."
  - "Never overwrite/delete/rename existing worksheets to resolve conflicts."
  - "Preserve per-output isolation, original data_files/report emission order, and existing Stage 1 preparation semantics where not explicitly changed by P01."

verification_required_from_implementor:
  executor_feasible:
    - "Static/semantic review: RC-1..RC-5 and D-03/D-04/D-06/OQ-7 compliance; no planned report-sheet creation before complete preflight; RC-1 gate before any Stage 1 creation when combine group exists; RC-2 execution fidelity; RC-3 registry coherence."
    - "L1 build on safe copy (prefer .ai/tmp copy of tracked workbook) via mko_treport/src/scripts/build_treport_vba.ps1; record exact command, report path, SUCCESS/ISSUES/ERROR and compile/VBOM limitation note (best-effort only)."
    - "Deterministic/manual verification design or execution on copies: single-member regression (naming/format/grouping/output/save guards), multi-member config fails before any report-sheet creation with required context, case-insensitive grouping/collision, dynamic placeholders, multi-digit suffix and suffix_length=0 length errors, existing default/non-default conflicts, planned-vs-planned cross-group collision, different-output same-name isolation."
    - "Workbook-state checks on copies: on preflight/RC-1 failure no planned report sheets created; final-produced registry drives cleanup/verify/save correctly in single-member path."
  user_performed_final:
    - "L2 (required for compile claim): Debug -> Compile VBAProject in Excel on built/checked workbook; record file/time/result."
    - "L3 (required final gate): run Create_report and relevant scenarios from workbook on copies only; validate B2-relevant behavior and regressions."
  reporting:
    - "Report changed export files, exact L1 command/report summary/path, L2 done/not done with explicit limitation if not done, L3 result or not performed, and confirmation 2637/2638 unused or equivalent justified."

review_gate:
  validator_re_review_required_before_implementation: true
  auditor_attention_recommended: true
  notes:
    - "RC-1 is binding: no multi-member sheet creation/preparation and no mid-execution gate in B2."
    - "RC-2/RC-3 prevent single-member regressions and save/cleanup guard breakage."
    - "RC-4/RC-5 remove diagnostic ambiguity and preflight sequencing gaps."
```

### 12.5 Implementation sequence (semantic units)

1. **`modTreportOutput` — planning primitives and code policy (RC-4):** implement effective base-name derivation (substitute placeholders + sanitize chars, 2601 on empty), raw-template/effective-base diagnostic payload, case-insensitive group-key helper; confirm 2637/2638 unused and prevent planning-validated D-04 cases from using legacy 2602.
2. **`modTreportOutput` — plan/registry contracts (RC-3):** define documented per-output member/group/planned-name registry schema (and Public Type(s) only if beneficial) including `GapRows`, source contexts, emission order, B3/B5 handoff notes, and explicit final-produced registry role for B2 single-member execution.
3. **`modTreportOutput` — suffix-aware name planner (D-04, RC-4):** deterministic assignment `base, base1..base(N-1)` in original order with actual suffix length; 2637 validation for all planned-name length violations (`suffix_length >= 0`) with required D-04 fields.
4. **`modTreportOutput` — collision preflight (D-06):** case-insensitive checks against existing workbook worksheets and all planned run names; 2638 fail-fast diagnostics; no side effects.
5. **`modTreportRun` — inspection + full resolve/collect pass (RC-5):** make output workbooks available for inspection; perform full `data_files`/report scan to collect resolved effective-base members without creating report sheets, preserving placeholder resolution basis (pre-`data_files.filters`) and read-only source handling.
6. **`modTreportRun` — finalize/preflight + RC-1 gate:** build groups, assign names, validate, preflight all outputs before any affected creation; remove/make unreachable 2612 for same-base grouping; if any `IsCombineGroup` exists, fail-fast before any Stage 1 report-sheet creation with required contextual diagnostic.
7. **`modTreportRun` — RC-2/RC-3 single-member-only execution:** if and only if all groups are single-member, execute preserved Stage 1 semantics/order using B2-assigned planned names and register created sheets in final-produced registry consumed by 2630/cleanup/verify/save; then finalize via existing guards.
8. **Self-review:** confirm RC-1..RC-5, D-03/D-04/D-06/OQ-7, no destructive behavior, no multi-member sheet creation, and regression guards.

### 12.6 Architectural and workbook-state constraints

- **SSOT/boundaries:** VBA source of truth remains `mko_treport/src/vbs/export/*`; only `modTreportOutput.bas` and `modTreportRun.bas` in scope for B2 code changes.
- **Semantic targeting only:** no line-number references in implementation or verification instructions.
- **RC-5 sequencing:** output workbook objects may exist for existing-name inspection, but **no planned report worksheets** may be created until the complete per-run/per-output plan set is resolved, length-validated, and collision-preflighted.
- **RC-1 gate:** if any group is multi-member/combine (`GroupSize > 1` / `IsCombineGroup = true`) after successful full planning/preflight, fail-fast **before any Stage 1 report-sheet creation** with required contextual diagnostic; **no member sheet creation/preparation** for multi-member groups in B2; no mid-execution gating after creation.
- **RC-2 execution fidelity:** single-member execution pass must preserve existing Stage 1 semantics/order and placeholder max resolution basis (pre-`data_files.filters`) except intentional P01 changes.
- **RC-3 registry coherence:** single-member execution must register created final/base sheets in the final-produced registry used by 2630/cleanup/verify/save; B3/B5 planned/temp/final separation documented but not deferred for correctness of B2 path.
- **RC-4 diagnostics/codes:** 2637 for all planned-name length violations (`suffix_length >= 0`) with D-04 fields; 2638 for D-06 preflight collisions; 2601 preserved; 2602 not used for planning-validated D-04 cases; 2612 not triggered for same-base grouping and not repurposed; 2637/2638 verified unused at implementation time; all other existing codes/messages unchanged.
- **No destructive conflict resolution:** never overwrite/delete/rename existing worksheets to resolve naming conflicts (D-06/CR-22g).
- **Isolation/order:** grouping/planning scoped per output workbook; preserve original `data_files`/report emission order (CR-02/CR-03/OQ-7).
- **No out-of-scope behavior changes:** `Refresh_data`, filter semantics, placeholder resolution rules, `insert_rows`/`skip_header`/`only_values`/template prep semantics unchanged except strictly necessary integration seams that preserve RC-2 equivalence.

### 12.7 Required validation

- **Logic and component interaction focus:** base-name pipeline (substitute -> sanitize -> group key), grouping accumulation with 2612 removed/unreachable for same-base grouping, suffix assignment/actual suffix length, 2637 D-04 diagnostic content (incl. `suffix_length >= 0`), 2638 D-06 preflight coverage (existing incl. defaults, planned-vs-planned, cross-group), per-output isolation, RC-5 sequencing, RC-1 gate timing, RC-2 execution fidelity, and RC-3 final-produced registry coherence.
- **Deterministic/manual verification scenarios (feasible to specify/execute on copies):**
  - single-member/non-combined regression with preserved Stage 1 outputs and successful cleanup/verify/save guards
  - multi-member/combine config: fail after successful full planning/preflight but **before any Stage 1 report-sheet creation** with required contextual diagnostic
  - identical resolved+sanitized base names in same output group in original order (grouping intent; B2 execution blocked by RC-1)
  - case-insensitive grouping/collision detection and cross-output same-name isolation
  - suffix-aware overflow with `suffix_length=0` and multi-digit suffix cases (effective base, explicit `31 - suffix_length`)
  - planned name conflicts with existing default/non-default sheets and planned-vs-planned cross-group collisions, no destructive resolution
  - dynamic placeholders affecting resolved base name
- **Build/compile checks:** executor may run L1 build on safe copy and capture report/limitations (best-effort, not proof of compile). **User performs final L2** (`Debug -> Compile VBAProject`) and **L3** run from workbook on copies (EVD-02). VBA compile checks in IDE are required for "compiles" claims but are user-performed in this workflow.

### 12.8 Non-goals and regression guards

- Do not implement B4 combine transfer/append/bounds/content-limit enforcement or B5 temp-deletion/final-registry completion.
- Do not create/prepare any multi-member group worksheet in B2; do not use mid-execution gating after creation.
- Do not change `group_rows`, existing config fields/diagnostics, or B1 contract behavior.
- Do not alter `Refresh_data`, source refresh, or unrelated report-generation logic.
- Do not add automatic conflict resolution, speculative refactors, or documentation edits in B2 implementation.
- Do not use line numbers; do not modify audit/problem artifacts except this planning update.

### 12.9 Done / acceptance checklist and gate

- [ ] **Validator re-review completed and approved for §12/B2 RC-1..RC-5 corrections and B2/B3 boundary before implementation.**
- [ ] `modTreportOutput.bas` provides deterministic base-name derivation, case-insensitive group-key, suffix-aware name planning (actual suffix length), 2637 D-04 length validation, and 2638 D-06 collision preflight APIs/helpers with documented contracts; 2637/2638 verified unused (or explicitly justified equivalent) at implementation time.
- [ ] Per-output plan/registry artifacts produced with member order, effective base/group key, planned names, `GapRows`, source contexts, and explicit final-produced registry/B3-B5 extension contract.
- [ ] `modTreportRun.bas` implements RC-5 sequencing: output workbooks available for inspection, full resolve/collect without report-sheet creation, complete per-run/per-output plan/finalize/preflight before any affected Stage 1 report-sheet creation.
- [ ] RC-1 enforced: if any `IsCombineGroup`/`GroupSize > 1` exists after successful full planning/preflight, fail-fast before any Stage 1 report-sheet creation with required contextual diagnostic; no multi-member member sheets created/prepared.
- [ ] RC-2 enforced for single-member-only execution: preserved Stage 1 semantics/order and placeholder max resolution basis except intentional P01 changes.
- [ ] RC-3 enforced: created final/base sheets registered in final-produced registry consumed by 2630/cleanup/verify/save checks.
- [ ] RC-4 enforced: 2612 not triggered for same-base grouping; 2637 covers all planned-name length violations (`suffix_length >= 0`) with D-04 fields; 2601 preserved; 2602 not used for planning-validated D-04 cases; 2638 used for D-06 preflight collisions; other existing codes/messages unchanged.
- [ ] Single-member/non-combined behavior preserved except spec-mandated D-04/D-06/grouping/registry-handling changes; no out-of-scope modules/behaviors changed.
- [ ] Static review complete; executor-feasible checks (incl. L1 on safe copy when applicable) and limitations recorded; user L2/L3 requested per EVD-02/project VBA workflow.

## 13. B3 — Detailed plan and implementor task (Orchestration refactor, grouping state, and two-stage lifecycle)

- **Block:** B3 — Orchestration refactor, grouping state, and two-stage lifecycle (core flow)
- **Plan status:** **Planned — Validator review required before implementation**
- **Review gate:** **Validator review required before implementation** (this §13 is the binding review artifact; implementation must not start until Validator approves the B3 orchestration/state/registry/hook design and B3/B4/B5 boundary/integration contract)
- **Spec/decision baseline:** D-01..D-06, OQ-7, CR-01..CR-18c, CR-22g/CR-23g, Approach 1, EVD-01..03, DoR-04/05; B2 contracts (RC-1..RC-5, 2637/2638, per-output `outPlanMembers/outPlanGroups/outPlanNames`, `B2SheetMemberPlan`)
- **Dependencies:** B0 baseline/guardrails; B1 config contract implemented (`TreportGeneral.reportRowsLimitGeneral`, `TreportDataFile.gapRows`, `RowIndex`); B2 naming/group planning/preflight implemented (commit `a096911`, `BuildOutputSheetPlan`/`AssignPlannedSheetNames`/`PreflightPlannedSheetNameCollisions`, RC-5 scan/build/preflight, `outPlanMembers/outPlanGroups/outPlanNames`, current RC-1 combine gate 2699)
- **Risk level:** **High** (core control-flow/lifecycle/registry refactor, regression-prone). **Agents required:** Auditor, Researcher, Planner, Validator (all) per §4 B3.
- **B3 implementation context (post-B1/B2 SSOT):** RC-5 scan/build/preflight and B2 per-output plan maps remain the pre-creation baseline. Remove/replace the RC-1 gate in `modTreportRun.Create_report` (current 2699 fail-fast when any combine group exists) only together with B3 Stage 1 preparation for all members and B4/B5 orchestration hooks. Keep `outUsedSheetNames(r)` as **final/base-produced registry only** for 2630/`CleanupDefaultSheets`/`EnsureProducedSheetsPresent`/save checks. Keep `outPlanNames(r)` as **planned/reserved namespace** (all planned base+temp names, D-06/RC-5). Track **transient temp prepared non-base worksheets** in per-output runtime state only — never in `outUsedSheetNames(r)` and never left after successful combine/finalization. R-01/R-02/CR-13 remain B4-scoped transfer/validation technique choices (§7.2, §7.3).

### 13.1 Objective and scope boundary

**Objective:** Refactor `Create_report`/run orchestration to a safe two-stage, dependency-aware flow that: (1) preserves RC-5 scan/build/preflight and Stage 1 preparation semantics/order/contexts for every emitted member, (2) prepares each planned member on its own assigned worksheet (base for group-first member, temp `base1..base(N-1)` for subsequent members) using B2-planned names, (3) establishes deterministic per-output runtime grouping state and registry separation for combine/finalization, (4) orchestrates Stage 2 combine and B5 lifecycle hooks in correct order per output/group, and (5) maintains fail-fast/no-save safety and existing guard invariants.

**In scope (B3):**
- `modTreportRun.bas`: remove/replace RC-1 2699 gate; generalize single-member Stage 1 execution to **all members** (single-member and combine groups) with emission/member-aware lookup; build/maintain per-output runtime group/member state from B2 plans; separate final-produced registry from planned/transient registries; add Stage 2 orchestration sequencing and B4/B5 hook calls/integration points; reconcile final/base registry before cleanup/verify/save.
- `modTreportOutput.bas` and/or `modTreportExtract.bas`: add **minimal hook/interface contracts** required by B3 orchestration (semantic procedure names/signatures and fail-fast stub behavior) only if necessary to define stable B3→B4/B5 handoff boundaries in the implementation task; no speculative B4 transfer/validation or B5 deletion logic design beyond required interface and safety contract.
- Integration with existing B1/B2/config/output/extract APIs without changing their validated contracts except where B3 must consume/extend them (e.g. runtime enrichment of plan member info dicts).

**B3 delivers:** complete, safe orchestration framework and state handoff; Stage 1 preparation for base+temp members; correct registry/lifecycle sequencing hooks. B3 may include minimal compile-safe fail-fast stubs for B4/B5 hook procedures if the implementation requires callable interfaces, but must not implement B4 combine semantics or B5 temp-deletion/final-registry completion logic.

**Not in scope (deferred to B4/B5 with their own high-risk cycles):** B4 D-05 valid-content gate, CR-15 per-subreport `report_rows_limit_general` enforcement in combine path, CR-16 projected Excel bounds before each paste, CR-10/CR-11 append geometry execution, CR-12–CR-14 transfer profile implementation (values/formats/number formats/row heights/hidden rows/outline restoration at offsets, exclusions), and B5 actual temp worksheet deletion/reconciliation completion after successful transfer. R-01/R-02/CR-13 low-level technique choices remain B4-scoped.

### 13.2 B3/B4/B5 boundaries and handoff hooks (binding)

| Boundary | B3 responsibility | B4 responsibility (hook consumer/implementor, separate block) | B5 responsibility (hook consumer/implementor, separate block) |
|---|---|---|---|
| RC-1/Stage 1 | Replace 2699 with Stage 1 preparation of **all** planned members (base+temps) across all outputs in preserved emission order; no combine-group blocking before preparation. | None (B4 runs only after Stage 1 complete). | None. |
| Runtime state | Build deterministic per-output runtime groups/members from B2 `outPlanGroups/outPlanMembers/outPlanNames`; enrich member records with runtime fields (IsBase, sheet ref, prepared/inserted range, status/contexts); maintain explicit group/member order without relying on Dictionary enumeration order. | Consume provided state/inputs; must not reorder members/groups contrary to preserved emission order. | Consume provided state/inputs for successful groups. |
| Registry separation | `outPlanNames(r)` = planned/reserved (unchanged). `outUsedSheetNames(r)` = **final/base-produced only**. Temps never added to final-produced. Register base/final sheets as produced when prepared (or at a deterministic point before cleanup/verify/save) so 2630/cleanup/verify/save remain valid. | Must not register temps as finals. | Must delete temps only after successful transfer and ensure final/base registry remains coherent before cleanup/verify/save. |
| Stage 2 orchestration | After Stage 1 completes for all outputs/members, iterate outputs/groups in deterministic order and call B4 combine hook for each `IsCombineGroup` group (`GroupSize > 1`); for single-member groups no combine call required. If B4 hook is not implemented, fail-fast with explicit diagnostic (recommended dedicated/contextual fail, e.g. "B4 combine hook not implemented") **before save/finalization**; EH must close outputs without save (no corrupted saved workbook). | Implement `CombinePreparedGroup` (semantic name) to append temp members into base per CR-10–CR-17v/D-05 and report success/failure via normal error handling/fail-fast contract. | Not invoked before successful combine. |
| B5 lifecycle | After successful combine of all required groups in an output (or per-group success boundary as defined by safe contract), call B5 finalization hook to delete temps and reconcile final state; only then proceed to `CleanupDefaultSheets`/`EnsureProducedSheetsPresent`/`SaveOutputWorkbook`. On any failure, do not delete temps as success cleanup, do not save, and preserve EH no-save behavior. | Must complete all required transfer/validation for group before signaling success to B5. | Implement temp deletion (`DeleteCombinedGroupTempSheets`/equivalent semantic hook) only after successful transfer; ensure temps absent and base finals still present/registered. |
| Error handling | Preserve existing fail-fast contextual diagnostics and EH close-no-save behavior for outputs/source workbooks/state restoration; no partial-success save claims. | Fail-fast with required P01/B4 context (D-05/CR-15/CR-16). | Fail-fast if temp deletion/final reconciliation cannot be completed safely; no false-success. |

**B3 hook contract (minimum, semantic; exact VBA signatures finalized in implementation but must satisfy these inputs/outputs):**

- **B4 hook — combine one prepared group (in `modTreportExtract.bas` or module chosen per B4, callable from `modTreportRun`):**
  - **Inputs:** output workbook context, base member runtime record (planned/base name, `Worksheet` ref, prepared inserted range/metadata as available, contexts), ordered temp member runtime records in emission order (planned name, `Worksheet` ref, `GapRows`, prepared range/metadata, contexts), group identity (`GroupKey`, `EffectiveBaseName`, output/report identifiers), `reportRowsLimitGeneral` and any required config/general context.
  - **Behavior contract:** perform/own B4 validations and combine semantics (D-05, CR-10–CR-17v, CR-12–CR-14) in B4 implementation. In B3-only state, a minimal stub may fail-fast with explicit "B4 combine hook not implemented" diagnostic (must not claim success, must not save).
  - **Output/side effect:** on success base sheet contains combined result and temps remain present for B5 deletion; on failure raise/fail-fast (no save).
- **B5 hook — finalize temps after successful combine (in `modTreportRun.bas`/`modTreportOutput.bas` per B5, callable from `modTreportRun`):**
  - **Inputs:** output workbook, group/runtime temp member refs (or group record), final-produced registry/ref as needed, context.
  - **Behavior contract:** delete temp worksheets only after successful transfer and reconcile final/base registry/state in B5 implementation. In B3-only state, a minimal stub may fail-fast with explicit "B5 finalization hook not implemented" diagnostic before save (must not claim success, must not save).
  - **Output/side effect:** on success no temp sheets remain for group and final/base produced registry coherent; on failure raise/fail-fast.

B3 must call hooks only at safe points: B4 after Stage 1 complete; B5 only after B4 success for the relevant group/output scope. B3 must not attempt to emulate B4/B5 logic.

### 13.3 Detailed design

#### 13.3.1 Runtime state model (per-output, deterministic)

B3 extends B2 plan artifacts in-memory (no cross-module UDT collection dependency). Recommended structures (all keyed/ordered deterministically):

- **Per-output runtime container** (parallel arrays/dictionaries indexed by output/report `r`, consistent with existing `outPlan*` arrays): `outRuntimeGroups(r)` or enriched `outPlanGroups(r)` + `outGroupOrder(r)`.
- **Group order (`outGroupOrder(r)`):** ordered sequence of `GroupKey` values in first-seen emission order for output `r` (derived by scanning `outPlanMembers(dfIdx)` in `dfIdx`/emission order for `ReportIndex=r`). Required because Dictionary key enumeration must not be relied upon for member/group sequence.
- **Group runtime record** (Dictionary, keyed case-insensitively where appropriate): existing B2 keys (`Key`, `BaseName`, `Members` collection) plus runtime fields:
  - `GroupKey`, `EffectiveBaseName`, `OutputIndex`/`OutputName` context
  - `IsCombineGroup` = (`Members.Count > 1`)
  - `CombineStatus` (`Pending`/`Prepared`/`Combined`/`Failed`/`Finalized` as needed)
  - `FirstEmissionOrder` (for diagnostics/order verification)
  - `Members` remains ordered collection keyed `"1"..CStr(N)` in preserved emission order (B2 pattern). Each member entry is the cloned/enriched member info Dictionary.
- **Member runtime record** (Dictionary; extends B2 `CloneB2MemberInfo` payload): B2 fields retained (`ReportIndex`, `DataFileIndex`, `DataFileRowIndex`, `EmissionOrder`, `GapRows`, `EffectiveBaseName`, `GroupKey`, `PlannedName`, `SuffixText`, `SuffixLength`, `OutputName`, `ReportName`, `FileCtx`, `RepCtx`, `RawReportSheetTemplate`) plus runtime fields:
  - `IsBase` (Boolean, `True` for first member in group order)
  - `IsTemp` (Boolean, `Not IsBase`)
  - `MemberIndexInGroup` (Long, 1-based)
  - `SheetNamePlanned` (= `PlannedName`)
  - `Sheet` (Object/`Worksheet`, set during Stage 1 preparation)
  - `InsertedRange` (Object/`Range` or Nothing, from `CopyPivotSnapshot`)
  - `Prepared` (Boolean)
  - `PreparedCtx` (String, concrete execution context)
  - optional metadata for B4 handoff: `PasteRow`, `PasteCol`, template/insert/skip flags as applicable from `cfg.DataFiles(DataFileIndex)` and report context (or resolvable via indices) — B3 must provide sufficient context/refs for B4 without hardcoding B4 technique.

**Member lookup:** do not scan by `ReportIndex` alone. Use `EmissionOrder` (globally unique in RC-5 scan) and/or composite key `(DataFileIndex, ReportIndex, EmissionOrder)` to locate the runtime member record for each execution emission. Recommended: build lookup map per output and/or per df keyed by `CStr(EmissionOrder)` or `CStr(DataFileIndex) & "|" & CStr(ReportIndex)` only if uniqueness is guaranteed; prefer `EmissionOrder` as primary key.

#### 13.3.2 Registry model and lifecycle rules (RC-3 aligned, B3/B5-ready)

| Registry / state | Contents | Populated by | Consumed by | Rules |
|---|---|---|---|---|
| `outPlanNames(r)` | All planned/reserved concrete sheet names for output `r` (base + temps), case-insensitive keys | B2 build/preflight (`BuildOutputSheetPlan`) | D-06 evidence/audit, B3/B4 diagnostics; not used as produced registry | Complete before any report-sheet creation (RC-5). Immutable for the run once built/preflighted. |
| `outUsedSheetNames(r)` | **Final/base-produced worksheet names only**, case-insensitive keys (`LCase$(plannedName)`), value `True` | B3 Stage 1 for base/final members (and reconciled by B5 after combine as needed) | 2630 no-produced guard, `CleanupDefaultSheets`, `EnsureProducedSheetsPresent`, save loop | Must contain every intended final/base report sheet and **must not** contain temp/transient sheets. For combine groups, base member (`IsBase=True`, planned base name) is the final sheet and must be registered. Temps (`base1..`) never registered. Non-empty only when at least one final/base sheet produced. |
| Runtime temp tracking | Temp worksheet objects/refs and status in group member runtime records (`IsTemp=True`, `Sheet`) | B3 Stage 1 preparation | B4 combine hook, B5 finalization hook, diagnostics/EH context | Temps exist on output workbook only between Stage 1 creation and successful B5 deletion (or EH no-save discard). Never treated as produced/final. Deleted only after successful transfer (B5). |
| `outBooks/outNames` | Output workbook objects and sanitized file names | existing `Create_report` | save/finalization/EH | Unchanged role. |

**Cleanup/save invariants:** `CleanupDefaultSheets` deletes worksheets not in `outUsedSheetNames(r)`; therefore temps must not be in `outUsedSheetNames(r)` and must be removed before cleanup runs on success path (B5 responsibility after successful combine). If B5 fails or is not implemented, B3 must fail before cleanup/save and EH must close outputs without save. `EnsureProducedSheetsPresent` requires all keys in `outUsedSheetNames(r)` to exist; base finals must remain after B5 deletes temps.

#### 13.3.3 Stage 1 generalization (all members, preserved semantics)

- **Prerequisites (unchanged):** RC-5 full scan (`CollectDataFilePlanMembers`) completed with placeholder max resolution per df row before df filters (equivalent basis), `BuildOutputPlan`/`PreflightPlannedSheetNameCollisions` completed for all outputs, `outPlanNames/outPlanGroups/outPlanMembers` populated.
- **Gate change:** remove RC-1 2699 combine-group block. Proceed to Stage 1 preparation for **all** emissions/members in all outputs.
- **Execution order:** iterate `cfg.DataFiles` in original order (`i = LBound..UBound`) and within each df emit/process members in original report order (`r = LBound..UBound(cfg.Reports)`) consistent with current behavior and `EmissionOrder`. For each emission, resolve runtime member record via emission/member-aware lookup and prepare that member's B2-assigned planned sheet.
- **Per-df processing (generalized from `ExecuteDataFileSingleMember`):** for each df row, reopen source workbook read-only (same as today), resolve `ws`/`pt` (`GetDataSheet`, `GetPrimaryPivot`), parse/classify and apply data-file filters once (`ParseAndClassify`/`ResolveFilterContext`/`ApplyDataFileFilters`), then for each report `r`: apply report filters (`ApplyReportFilters`), enforce 2410 if pivot data empty, resolve planned member runtime record and planned name, compute `pasteRow/pasteCol` from `cfg.DataFiles(i)` with same normalization/defaulting as current code, call `PrepareTargetSheet(outBooks(r), plannedName, cfg.DataFiles(i).template, repCtx/memberCtx)` with B2-assigned planned name, call `CopyPivotSnapshot(pt, outWs, pr, pc, ctx, cfg.DataFiles(i).onlyValues, cfg.DataFiles(i).insertRows, cfg.DataFiles(i).skipHeader, (Len(Norm(cfg.DataFiles(i).template))>0))`, call `ApplySheetFormatting(outWs, cfg.DataFiles(i).freezeRow, cfg.DataFiles(i).freezeCol, cfg.DataFiles(i).Zoom, repCtx, (Len(Norm(cfg.DataFiles(i).template))>0))`, call `ApplyOutlineGrouping(insertedRng, cfg.DataFiles(i).GroupRows, cfg.General.rowsCollapsed, ctx & ",output='" & outNames(r) & "',data_sheet='" & plannedName & "'")`, store `Sheet`/`InsertedRange`/`Prepared` in runtime member record, register **only base/final members** in `outUsedSheetNames(r)` (`LCase$(plannedName)`), close source workbook no-save. Preserve all existing error contexts/codes/behaviors (2410, 2601/2634/2635/2636/2637/2638 context patterns as applicable, extract/output errors) except intentional P01 changes (2699 removed, grouping/combine flow enabled).
- **Template/window side effects:** existing `PrepareTargetSheet` copies template from `ThisWorkbook` per use and `ApplySheetFormatting` may activate workbook/window for freeze/zoom behavior; B3 must preserve these Stage 1 behaviors and avoid additional side effects. Stage 2/B4/B5 must operate without reapplying `paste_row/paste_col` and without transferring forbidden worksheet-level display/layout settings from temps.
- **No reliance on Dictionary iteration for member emission order:** iterate by `cfg.DataFiles`/`cfg.Reports` order and/or `EmissionOrder` sequence; use group/member ordered collections (`"1"..N`) for per-group order.

#### 13.3.4 Stage 2 orchestration and finalization sequencing

After Stage 1 completes successfully for all df rows/outputs:
1. **Per-output Stage 2 combine pass:** for each output `r` in report/output order, iterate groups in deterministic first-seen emission order (`outGroupOrder(r)` / equivalent). For each group:
   - if `Not IsCombineGroup` (single member): no B4 combine call; base already prepared and registered.
   - if `IsCombineGroup`: resolve base member runtime record (`IsBase=True`) and ordered temp member records (`IsTemp=True`, `MemberIndexInGroup` 2..N) from runtime state; call **B4 combine hook** with required inputs (base sheet, ordered temps, gapRows per temp member record, contexts, `cfg.General.reportRowsLimitGeneral`/limits context). On B4 success, call **B5 finalization hook** for the group (or output-scope finalize after all group combines in output, per B5 contract) to delete temps and reconcile registry. Do not run cleanup/verify/save for output until all combines and B5 finalization succeed for that output's required groups.
2. **Save/finalization path (success only, per output):** preserve existing sequence/guards: require `outUsedSheetNames(r).Count > 0` else 2630, `CleanupDefaultSheets(outBooks(r), outUsedSheetNames(r), ctx & "/cleanup")`, `EnsureProducedSheetsPresent(outBooks(r), outUsedSheetNames(r), ctx & "/verify")`, `SaveOutputWorkbook(...)`. Temps must already be absent and not in `outUsedSheetNames(r)` before cleanup.
3. **Failure behavior:** any failure in Stage 1, B4, B5, or finalization follows existing EH path (close source/output workbooks without save, restore state, show error). No temp deletion as "success cleanup" on failure; no save.

**B3-only interim behavior (if B4/B5 not yet implemented):** B3 implementation must still replace 2699 and complete Stage 1 preparation for all members/groups. Upon reaching a combine group requiring B4 (or requiring B5 after hypothetical combine), call the defined hook; if hook is an explicit fail-fast stub, error must clearly identify missing B4/B5 implementation and context, and execution must abort before save via EH (unsaved output workbooks closed without save). This is intentional fail-safe behavior, not a success path.

### 13.4 Implementation sequence (semantic units)

1. **`modTreportRun.Create_report` — planning/state scaffolding (post-RC-5/preflight):** retain scan/build/preflight and `outPlanNames/outPlanGroups/outPlanMembers`; add per-output runtime state containers (`outGroupOrder`, runtime group/member enrichment structures) and build deterministic group order and member lookup maps from B2 plans without creating sheets.
2. **`modTreportRun.Create_report` — remove RC-1 gate:** delete/replace 2699 `BuildCombineGateMessage` fail-fast block so combine groups proceed to Stage 1 preparation; keep `BuildCombineGateMessage` only if reused for diagnostics, otherwise retire as dead code in B3 scope.
3. **`modTreportRun` — Stage 1 generalization:** refactor/replace `ExecuteDataFileSingleMember` (and related member lookup) with a generalized Stage 1 executor (e.g. `ExecuteDataFilePlannedMembers` / equivalent semantic unit) that prepares every planned member emission for the df/report using B2-assigned names, preserves existing filter/2410/placement/snapshot/format/grouping semantics, updates runtime member records (`Sheet`, `InsertedRange`, `Prepared`, `IsBase/IsTemp`), and registers only base/final members in `outUsedSheetNames(r)`.
4. **`modTreportRun` — Stage 2 orchestration integration:** after Stage 1 loop, add per-output/group orchestration in deterministic order calling B4 combine hook for `IsCombineGroup` groups and B5 finalization hook at correct success boundaries; integrate with existing EH/no-save semantics.
5. **B4/B5 hook interface definition/stubs (minimal, module-appropriate):** introduce callable hook procedures with documented semantic signatures/contracts (`modTreportExtract` combine hook for B4, `modTreportOutput` and/or `modTreportRun` finalization hook for B5 as appropriate) that fail-fast with explicit "not implemented / B4 required / B5 required" diagnostics if invoked before B4/B5 implementation. Do not implement combine/transfer/delete logic.
6. **Finalization/guard verification:** ensure save loop runs only after successful Stage 2/B5 completion, `outUsedSheetNames(r)` contains only final/base sheets, temps absent, and 2630/cleanup/verify/save behavior unchanged for single-member and correct for combined finals.
7. **Self-review:** confirm RC-5 preserved, 2699 removed safely, Stage 1 semantics/order preserved, member-aware lookup correct, registry separation correct, deterministic ordering, hook boundaries explicit, and no B4/B5 transfer logic implemented.

### 13.5 Architectural and workbook-state constraints

- **SSOT/boundaries:** VBA source of truth `mko_treport/src/vbs/export/*`. Primary B3 target `modTreportRun.bas`; hook/interface touchpoints may include `modTreportOutput.bas` and/or `modTreportExtract.bas` only for minimal B3→B4/B5 contracts/stubs. No docs edits in B3.
- **Semantic targeting only:** no line-number references.
- **Preserve B1/B2 contracts:** config fields/validation, 2637/2638/2601/2634 behavior, B2 plan semantics, RC-5 scan/build/preflight, planned-name assignment, and placeholder max resolution basis (pre-df-filters) unchanged.
- **Preserve Stage 1 semantics:** reopen source read-only per df, df filters once + per-report filters in same order, 2410, placement calculation, `PrepareTargetSheet` with B2-assigned planned name, `CopyPivotSnapshot`/`ApplySheetFormatting`/`ApplyOutlineGrouping` calls/parameters/contexts equivalent to current supported behavior.
- **Deterministic ordering:** do not rely on `Scripting.Dictionary` key/item enumeration order for emission/group/member sequence; use explicit ordered collections/indices (`EmissionOrder`, `"1"..N`, `outGroupOrder`).
- **Registry/workbook-state invariants:** `outUsedSheetNames` = finals only; `outPlanNames` = planned/reserved; temps transient runtime-only; no temp in final registry; no temp deletion before successful transfer; no cleanup/verify/save on failure; EH closes outputs without save.
- **No Stage 2 placement reapplication:** B3/B4 orchestration must not reapply `paste_row`/`paste_col` during combine (B4 appends per CR-10/11).
- **No destructive conflict resolution / no out-of-scope changes:** no `Refresh_data` changes, no filter/placeholder/insert/skip/only_values/template semantic redesign, no automatic naming conflict resolution, no `-SyncExports`/`-DeployToDataTheiss`.

### 13.6 Required validation

**Logic and component interaction focus (executor-feasible):**
- RC-5 scan/build/preflight still runs completely before any report-sheet creation; 2699 removed and replaced safely.
- Single-member/non-combined path: behavior/output/registry/save guards preserved (base planned name, one sheet per group, registered in `outUsedSheetNames`, cleanup/verify/save succeed).
- Multi-member/combine-group configs: Stage 1 prepares all base+temp worksheets with B2-assigned names in preserved order and correct filter/2410/placement/snapshot/format/grouping semantics; base registered as final-produced, temps not registered.
- B4/B5 hook invocation points correct (after Stage 1; B5 only after B4 success boundary); if stubs/not implemented, fail-fast is explicit and occurs before save with EH no-save (no corrupted saved output).
- Member-aware lookup correctness for same df emitting multiple reports and groups spanning df rows (no ReportIndex-only ambiguity).
- Registry coherence: 2630/cleanup/verify/save consume finals only; temps absent before cleanup on success path (or never saved on failure).
- Deterministic group/member ordering independent of Dictionary iteration.

**Build/compile checks:** executor may run L1 build on a safe `.ai/tmp` copy via `mko_treport/src/scripts/build_treport_vba.ps1` and capture exact command/report path/summary (best-effort, not proof of compile). **User performs final L2** (`Debug -> Compile VBAProject`) and **L3** run from workbook on copies (EVD-02). VBA compile checks in IDE are required for "compiles" claims but user-performed.

**B3-specific L3 focus (user, on copies):** single-member regression; multi-member scenario reaches Stage 1-prepared base+temps and fails safely at B4/B5 hook boundary without saved output (until B4/B5 implemented), with clear diagnostic; inspect unsaved/failure behavior only via controlled copies if needed and never claim success.

### 13.7 Non-goals and regression guards

- Do not implement B4 combine append/transfer/validation logic (D-05, CR-10–CR-16, CR-12–CR-14, R-01/R-02/CR-13 technique choices) or B5 temp-deletion/final reconciliation logic beyond minimal fail-fast hook stubs/interfaces required by B3.
- Do not change `Refresh_data`, unrelated filter/placeholder/`insert_rows`/`skip_header`/`only_values`/template semantics, or config/docs.
- Do not reintroduce 2612 same-base collision behavior; do not alter 2637/2638/2601 contracts.
- Do not register temps in `outUsedSheetNames`; do not delete temps before successful transfer; do not save on failure.
- Do not rely on Dictionary iteration order for emission/group/member sequencing.
- No line numbers; no implicit `-SyncExports`/`-DeployToDataTheiss`; no audit/problem artifact edits.

### 13.8 Done / acceptance checklist and gate

- [ ] **Validator review completed and approved for §13/B3 orchestration, runtime state, registry separation, Stage 1 generalization, hook contracts, and B3/B4/B5 boundary before implementation.**
- [ ] RC-5 scan/build/preflight flow preserved; RC-1 2699 combine-group gate removed/replaced so all planned members (including combine groups) proceed to B3 Stage 1 preparation.
- [ ] B3 builds deterministic per-output runtime group/member state from B2 plans with explicit group/member ordering and emission/member-aware lookup (no ReportIndex-only ambiguity; no reliance on Dictionary iteration order for sequence).
- [ ] Stage 1 generalized to prepare all members (base and temps) on B2-assigned planned sheets with preserved semantics/order/contexts (source reopen per df, df filters once, per-report filters, 2410, placement, `PrepareTargetSheet`, `CopyPivotSnapshot`, `ApplySheetFormatting`, `ApplyOutlineGrouping`).
- [ ] Registry separation enforced: `outPlanNames` planned/reserved; `outUsedSheetNames` finals/base-produced only; base/final members registered appropriately; temps never registered as finals and tracked only in runtime state.
- [ ] B4 combine hook and B5 finalization hook integration points/call sequencing implemented per contract (after Stage 1; B5 only after B4 success); minimal fail-fast stubs/interfaces present only if required, with explicit diagnostics and no save/no false-success behavior.
- [ ] Finalization path preserves 2630/cleanup/verify/save invariants; temps absent and finals coherent before cleanup/verify/save on success; EH no-save behavior preserved on failure.
- [ ] Single-member/non-combined behavior preserved; no B4 transfer/validation technique or B5 deletion logic implemented in B3 beyond defined hooks/stubs.
- [ ] Static review complete; executor-feasible checks (incl. L1 on safe copy when applicable) and limitations recorded; user L2/L3 requested per EVD-02/project VBA workflow.

### 13.9 Implementor task (YAML-style, template-aligned)

```yaml
id: B3_orchestration_refactor_grouping_state_two_stage_lifecycle
title: "B3 — Orchestration refactor, grouping state, and two-stage lifecycle"
priority: high
status: "Planned — Validator review required before implementation"

depends_on:
  - B0
  - B1  # implemented in modTreportConfig.bas (commit 3f653e4)
  - B2  # implemented in modTreportOutput.bas/modTreportRun.bas (commit a096911)

source_reference: .ai/plans/01-combine-subreports-worksheet-name-validation.md
source_section: "§13 B3 — Detailed plan and implementor task (Orchestration refactor, grouping state, and two-stage lifecycle)"
spec_reference: .ai/problems/02_combine_subreports_worksheet_name_validation_spec.md
spec_decisions: [D-01, D-02, D-03, D-04, D-05, D-06, OQ-7, CR-01, CR-02, CR-03, CR-04, CR-05, CR-06, CR-07, CR-08, CR-09, CR-10, CR-11, CR-12, CR-13, CR-14, CR-15, CR-16, CR-17v, CR-18c, CR-22g, CR-23g]

review_gate:
  validator_review_required_before_implementation: true
  required_agents_for_block: [Auditor, Researcher, Planner, Validator]
  mandatory_context:
    - "§4 B3 and B3 implementation context (post-B1/B2 SSOT)"
    - "§7.2, §7.3 uncertainties (R-01/R-02/CR-13 remain B4-scoped)"

description: >
  Refactor modTreportRun.Create_report orchestration from RC-1-blocked single-member execution to a
  two-stage flow: preserve RC-5 scan/build/preflight, remove/replace 2699 combine-group gate, build
  deterministic per-output runtime grouping/member state from B2 plans, generalize Stage 1 to prepare
  all planned members (base+temps) on B2-assigned sheets with preserved semantics/order/contexts,
  separate planned/reserved, transient temp, and final/base-produced registries, and orchestrate B4
  combine and B5 finalization hooks before cleanup/verify/save. Do not implement B4 combine
  transfer/validation logic or B5 temp-deletion/final-registry completion logic; provide only minimal
  explicit fail-fast hook interfaces/stubs if required, with no-save safety.

goals:
  - "Preserve RC-5 scan/build/preflight and B2 plan artifacts; remove/replace 2699 only together with safe Stage 1-all-members flow."
  - "Build deterministic per-output runtime state (group order, enriched member records: IsBase/IsTemp/Sheet/InsertedRange/Prepared/contexts/gaps) with emission/member-aware lookup and no Dictionary-order dependence."
  - "Generalize Stage 1 to prepare every member on its B2-planned sheet (base or temp) in preserved df/report emission order with existing filter/2410/placement/snapshot/format/grouping semantics."
  - "Enforce registry separation: outPlanNames planned/reserved; outUsedSheetNames final/base-produced only; temps transient runtime-only and never registered as finals."
  - "Orchestrate B4 combine hook per combine group after Stage 1 and B5 finalization hook only after B4 success, before cleanup/verify/save; preserve 2630/cleanup/verify/save and EH no-save invariants."
  - "Define stable B3->B4/B5 handoff inputs/hooks; if not implemented, fail-fast explicitly before save with no corrupted output."

non_goals_regression_guards:
  - "No B4 implementation of D-05 valid-content detection, CR-10/CR-11 append geometry execution, CR-12-CR-14 transfer profile, CR-15 per-subreport limit enforcement in combine, or CR-16 projected Excel bounds checks (R-01/R-02/CR-13 deferred)."
  - "No B5 implementation of temp worksheet deletion/final reconciliation logic beyond minimal fail-fast hook contract/stub if required."
  - "No changes to Refresh_data, filter semantics, placeholder resolution rules/timing basis, insert_rows/skip_header/only_values/template preparation semantics, or paste_row/paste_col behavior in Stage 1."
  - "No reapplication of paste_row/paste_col in Stage 2/combine; no transfer of forbidden worksheet-level display/layout settings."
  - "No overwrite/delete/rename on naming conflicts; no temp deletion before successful transfer; no save/cleanup-as-success on failure."
  - "No registration of temp sheets in outUsedSheetNames; no reliance on Dictionary iteration order for emission/group/member sequencing."
  - "No implicit -SyncExports / -DeployToDataTheiss; no docs edits in B3."

files:
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    semantic_targets:
      - type: procedure
        name: "Create_report"
        intent: "Remove/replace RC-1 2699 gate; add runtime state/group-order construction; run Stage 1 for all members; orchestrate Stage 2 B4/B5 hooks per output/group; preserve RC-5, finalization guards, and EH no-save."
      - type: procedure_or_helper
        name: "ExecuteDataFileSingleMember (refactor/replace) / generalized Stage 1 executor"
        intent: "Prepare all planned member emissions for a df/report using emission/member-aware lookup and B2-assigned names; preserve existing Stage 1 semantics; update runtime member state; register only base/final members in outUsedSheetNames."
      - type: helper_procedures
        name: "runtime state/build/lookup/order helpers (e.g. build output group order, locate member by emission, mark prepared)"
        intent: "Deterministic state model over B2 artifacts; no Dictionary-order dependence."
      - type: helper_procedure
        name: "Stage 2 orchestration helper(s)"
        intent: "Iterate outputs/groups in deterministic order, invoke B4 combine hook for combine groups and B5 finalization hook only after B4 success, before cleanup/verify/save."
      - type: procedure_or_hook
        name: "B5 finalization hook wrapper/interface (if placed in Run)"
        intent: "Minimal contract/stub only if required; fail-fast explicit before save when not implemented."
  - path: mko_treport/src/vbs/export/modTreportExtract.bas
    semantic_targets:
      - type: procedure
        name: "B4 combine hook (e.g. CombinePreparedGroup) — interface/stub only if required"
        intent: "Define callable semantic contract for B4; if unimplemented fail-fast with explicit diagnostic and no side effects implying success."
  - path: mko_treport/src/vbs/export/modTreportOutput.bas
    semantic_targets:
      - type: procedure
        name: "B5 temp-finalization hook/interface (if placed in Output) — stub only if required"
        intent: "Define callable semantic contract for B5; if unimplemented fail-fast with explicit diagnostic and no side effects implying success."

implementation_scope_and_sequence:
  - step: 1
    target: "modTreportRun.bas (post-preflight runtime scaffolding)"
    action: "Build per-output deterministic group order and enrich B2 group/member records with runtime fields and emission/member lookup structures; keep outPlanNames/outUsedSheetNames roles explicit."
  - step: 2
    target: "modTreportRun.bas (Create_report gate)"
    action: "Remove/replace 2699 RC-1 combine-group fail-fast so execution proceeds to Stage 1 for all members; preserve RC-5 scan/build/preflight behavior."
  - step: 3
    target: "modTreportRun.bas (Stage 1 generalization)"
    action: "Refactor Stage 1 execution to prepare every planned member (base+temp) using B2-assigned names and emission/member-aware lookup with preserved semantics/order/contexts; update runtime sheet/range/status; register only base/final members in outUsedSheetNames."
  - step: 4
    target: "modTreportRun/modTreportExtract/modTreportOutput (B4/B5 hook contracts)"
    action: "Define minimal callable hook interfaces/stubs only as needed for B3 orchestration; fail-fast explicitly before save if invoked unimplemented; do not add combine/delete logic."
  - step: 5
    target: "modTreportRun.bas (Stage 2 orchestration and finalization)"
    action: "After Stage 1, call B4 per combine group in deterministic order, then B5 only on B4 success, then existing cleanup/verify/save guards; maintain EH no-save on any failure."
  - step: 6
    target: "self-check"
    action: "Verify RC-5 preserved, 2699 replaced safely, Stage 1 parity, registry separation, deterministic ordering, hook boundaries, and no B4/B5 logic leakage."

contract_and_state_rules:
  registries:
    outPlanNames: "planned/reserved concrete names (base+temps), complete pre-creation; unchanged role."
    outUsedSheetNames: "final/base-produced names only; consumed by 2630/CleanupDefaultSheets/EnsureProducedSheetsPresent/save; temps never included."
    runtime_temp_tracking: "Worksheet refs/status in runtime member records only; deleted only after successful B4+B5 path; never saved on failure."
  member_identity:
    lookup: "emission/member-aware (prefer EmissionOrder); never ReportIndex-only."
    base_temp: "first group member IsBase=True/final sheet; subsequent members IsTemp=True with B2 suffix names base1..base(N-1)."
  ordering:
    deterministic: "explicit df/report emission order, group first-seen emission order, and member order 1..N; no Dictionary iteration dependence."
  stage1_preservation:
    - "reopen source read-only per df"
    - "resolve sheet/pivot, apply df filters once then per-report filters in same order"
    - "2410 on empty pivot data"
    - "placement/paste_row/paste_col/template/insert_rows/skip_header/only_values semantics preserved"
    - "PrepareTargetSheet with B2-assigned planned name, CopyPivotSnapshot, ApplySheetFormatting, ApplyOutlineGrouping"
  hooks:
    b4_combine: "inputs include base runtime member, ordered temp members with GapRows/contexts, group/output identity, reportRowsLimitGeneral/config context; success required before B5; stub must fail-fast explicit before save."
    b5_finalize: "inputs include output/group temp refs and registry/final context; runs only after B4 success; deletes temps and reconciles finals in B5 impl; stub must fail-fast explicit before save."

acceptance_criteria:
  - "2699 RC-1 gate removed/replaced and combine groups reach Stage 1 preparation; RC-5 scan/build/preflight behavior preserved."
  - "All planned members (single and multi-member groups) prepared on B2-assigned sheets in preserved order with Stage 1 semantic parity and correct contexts."
  - "Runtime state provides deterministic group/member ordering and emission/member-aware lookup; member records carry required B3/B4 handoff fields (planned name, IsBase/IsTemp, Sheet, prepared range/status, GapRows, contexts)."
  - "outUsedSheetNames contains only final/base-produced sheets; temps never registered; outPlanNames remains planned/reserved."
  - "B4/B5 hooks invoked at correct lifecycle points; unimplemented hooks fail-fast explicitly before cleanup/verify/save with EH no-save and no false-success."
  - "On success path (when B4/B5 implemented), temps removed before cleanup/verify/save and 2630/cleanup/verify/save invariants hold; single-member behavior unchanged."
  - "No B4 transfer/validation technique or B5 deletion logic implemented in B3 beyond minimal contracts/stubs."

constraints:
  - "Never change VBA/code in this planning turn; plan/task artifact only."
  - "No line numbers; use semantic anchors only."
  - "No docs edits in B3; no -SyncExports / -DeployToDataTheiss."
  - "No B4/B5 logic implementation beyond required hook interfaces/stubs."
  - "Preserve B1/B2 contracts and RC-5/Stage 1 semantics except intentional P01 orchestration/lifecycle changes."
  - "Never overwrite/delete/rename worksheets to resolve conflicts; never delete temps before successful transfer; never save on failure."

verification_required_from_implementor:
  executor_feasible:
    - "Static/semantic review: 2699 replacement safety, RC-5/Stage 1 parity, deterministic ordering, emission/member-aware lookup, registry separation, hook sequencing, and EH/no-save behavior."
    - "L1 build on safe copy (prefer .ai/tmp copy of tracked workbook) via mko_treport/src/scripts/build_treport_vba.ps1; record exact command, report path, SUCCESS/ISSUES/ERROR and compile/VBOM limitation note (best-effort only)."
    - "Deterministic/manual verification design or execution on copies: single-member regression unchanged; multi-member Stage 1 creates base+temp planned sheets with correct prepared content/order; base registered final, temps unregistered; B4/B5-unimplemented path fails explicitly before save with no saved corrupted output."
    - "Workbook-state checks on copies: no temps in final-produced registry; cleanup/verify/save guards correct on supported paths; no-save on failure."
  user_performed_final:
    - "L2 (required for compile claim): Debug -> Compile VBAProject in Excel on built/checked workbook; record file/time/result."
    - "L3 (required final gate): run Create_report and relevant scenarios from workbook on copies only; validate B3-relevant orchestration/state/registry behavior and regressions."
  reporting:
    - "Report changed export files, exact L1 command/report summary/path, L2 done/not done with explicit limitation if not done, L3 result or not performed, hook/stub status (B4/B5 implemented or explicit fail-fast stubs), and any residual uncertainties observed."

review_gate:
  validator_review_required_before_implementation: true
  auditor_researcher_validator_required: true
  notes:
    - "B3 must not hardcode B4 transfer/validation technique (R-01/R-02/CR-13 remain B4-scoped)."
    - "B3 may replace 2699 only with complete Stage 1-all-members flow plus correct hook/lifecycle sequencing."
    - "If B4/B5 are unavailable, safe fail-fast before save is required (no false success, no corrupted saved workbook)."
```

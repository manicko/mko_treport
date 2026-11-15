# Execution plan: treport (001_treport)

- **Plan ID:** 001-treport-execution-plan
- **Based on (current baseline):** `.ai/problems/001_treport_spec.md` — **rev. 5** (2026-11-14), incorporating PO decisions: final runtime VBA-only with auxiliary dev tools allowed, no golden-output workbook required, `group_rows` deferred/no-op, explicit `Refresh_data`/`Create_report` run controls, `Trim` whitespace policy, optional `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`, and evidence-driven V-007/V-018/V-008 handling.
- **Revision lineage:** rev. 5 supersedes rev. 4 as planning/implementation baseline. Rev. 4 finalized D1–D17 and the VBA-only runtime constraint; rev. 5 clarifies runtime-vs-auxiliary tooling and resolves/adjusts open items per PO direction.
- **Planning intent:** dependency-safe, minimal-scope, incremental decomposition into VBA implementation and verification blocks aligned to rev. 5. **No final implementation approach is chosen where runtime/Excel behavior remains technically uncertain** (exact PQ refresh completion mechanics, pivot object/selection/copy mechanics, and runtime-vs-export parity remain subject to OQ-T1–T4 validation in target Excel). No source code files are changed by this plan document; only `.ai/problems/001_treport_spec.md` and `.ai/plans/001-treport-execution-plan.md` are edited in this documentation update.
- **Execution posture:** proceed with **VBA-compliant implementation in `treport.xlsm`** and required local/config/control updates as evidence allows. Auxiliary development tools (Python/PowerShell/BAT/etc.) may be used for inspection, workbook/pivot metadata reading, VBA extraction/analysis, testing/checking, test-data preparation, and diagnostics, but **must not become runtime dependencies** of the final user-facing solution. Large Excel workbooks/generated reports remain local/gitignored (backup/working-copy workflow); do not commit them.

## Scope and guardrails

- **In scope:** implement and verify rev. 5 behavior in `treport.xlsm` (Excel VBA macros/modules) for `Refresh_data` and `Create_report` using existing `settings`/`run`/ListObjects (`general`, `reports`, `data_files`, `placeholders`): config read/validation, deterministic file resolution/access modes (D1,D2), `filters` DSL parse and dependency-aware `min/max`/`max_*` resolution (D5,D6,D10,D11), evidence-based per-file field identity/alignment (V-007/V-018), filter staging/placement/matching with optional-token exception (D4,D7–D9,D12), pivot snapshot extraction and static paste (D15), formatting/layout with `group_rows` deferred/no-op (D13), output assembly/naming/placeholders/collision/idempotency with `Trim` (D3,D11,D14), fail-fast contextual diagnostics (D12), and required run controls/entry points (FR-RUN-1).
- **Runtime boundary (mandatory):** end user runs final solution in Excel with VBA only (no Python/PowerShell/BAT/external runtime). Auxiliary dev/test/inspection scripts/tools are allowed but non-runtime.
- **Out of scope (current):** `group_rows` color-band grouping (deferred, no-op), golden expected-output workbook, non-Excel runtimes/schedulers/UI redesign, source data/upstream PQ/model changes, committing large `data/theiss` files/results, invented requirements/mappings.
- **Local artifact policy:** `data/theiss/*` and `treport.xlsm` are gitignored/local; perform changes on working copies with backups, keep originals for comparison/recovery, store evidence/results locally as appropriate.
- **Non-runtime artifacts:** `mko_treport/src/vbs/` exports/snapshots are maintainability aids (runtime-vs-export parity unverified, OQ-T4). `mko_treport/src/scripts/` currently has no treport runtime modules; prior Python runtime modules were removed (4e60165,148b668,3d7b4c2,93e1ae4 in e040c8c). Tracked `mko_treport/tests/treport_*` and any restored/created helpers may be used as auxiliary (non-runtime) aids.

## Resolved PO decisions and current evidence (rev. 5)

| Topic | Decision / evidence | Planning implication |
|---|---|---|
| **D18 / acceptance baseline** | Windows 10+ / Excel 2021+; `data/theiss` workbooks and `treport.xlsm` available. **No golden expected-output workbook required.** Validate against functional requirements and actual source data (correct data/sheets/values/formatting/names/structure). Record exact Excel build and macro/PQ prerequisites in test environment but do not block unrelated development. | Remove golden-workbook gate from planning/execution; verify via functional/source comparison and recorded environment evidence. |
| **Runtime vs auxiliary tools** | Final solution VBA-only/Excel-native for user. Python/PowerShell/BAT/etc. allowed for inspection, pivot/metadata reading, testing/checking, VBA extraction/analysis, test-data prep/validation, diagnostics. Helper scripts may be restored/recreated for auxiliary use only (not runtime dependencies). | Use auxiliary tools freely during development/validation; keep VBA as sole runtime path. |
| **Git/local Excel** | No requirement to manage local Excel workbooks via Git; `data/theiss` large and gitignored. Backup/working-copy workflow; evidence/results local. | Do not commit large Excel/generated reports; document local evidence paths/hashes as useful. |
| **D13 / R4 `group_rows`** | Excluded from current scope (deferred). Do not implement row grouping; ignore values as no-op; do not block or require for acceptance (incl. `report_m.group_rows=TRUE`). | Treat as no-op in VBA logic/validation; no R4 compliance blocker. |
| **FR-RUN-1 controls** | `run` sheet must have exactly two buttons: `Refresh_data` and `Create_report`; VBA procedures same names. Remove old `RunXlconcat` button and obsolete workflow. | Replace legacy control/binding and implement/rename procedures accordingly. |
| **V-007/V-018 period fields** (config + evidence) | `data_files.filters` now use `Year` for all rows (user-updated, verified in `treport.xlsm`): `Month Number=(1-max)|Year=max` (`report_m`), `Week Number=(1-max)|Year=max` (`report_w`, `report_wclip`). Inspected caches show base `Year` present in all three files; `report_wclip` also has a grouped `Years` field plus base `Year`. | **Alignment:** config is aligned to `Year` for all files. Implement/resolve/filter/derive `max_year` against verified/configured `Year` per file (no guessed mapping). Validate runtime behavior in target Excel for all files, including that `Year` filtering is correct for `report_wclip` despite the grouped `Years` field; escalate only a genuine residual business ambiguity with concrete alternatives if discovered. |
| **V-008 category token** (evidence + PO) | Config token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` paired with `ЗУБНАЯ ПАСТА`. Inspected `articleLevel4Name_clean` cache items include `ЗУБНАЯ ПАСТА` but not `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` as distinct item in `report_m/report_w/report_wclip`. PO: optional/conditional — select if present, skip if absent, absence must not fail; at least one valid value must remain selected. | Implement optional-skip logic in filter application/validation; not a mandatory match. |
| **V-020 whitespace** | Use `Trim` for configured names/relevant strings before matching/substitution/naming/sanitization/collision checks; preserve spaces inside names (e.g. `DOLGIT ` → `DOLGIT`). | Apply consistently in VBA validation/resolution/naming. |

## Artifact and evidence context (current)

- **`treport.xlsm` (local, gitignored):** `run`/`settings`, 4 ListObjects (`general`: `input_path,output_path,delimeter`; `reports`: 6 rows; `data_files`: 3 rows with `filters` and `group_rows` column present; `placeholders`: `max_weeknum,max_monthnum,max_year`). `run` has one legacy button “Button 1” bound to `[0]!RunXlconcat` (caption “Run Macro”) — to be removed/replaced per FR-RUN-1. `data_files.filters` (user-updated, verified): `report_m` `Month Number=(1-max)|Year=max`, `report_w`/`report_wclip` `Week Number=(1-max)|Year=max`. Trailing-space name `DOLGIT ` present in config (handled by `Trim`).
- **Data workbooks:** `data/theiss/monthly/report_m.xlsx`, `weekly/report_w.xlsx`, `base/report_wclip.xlsx` exist locally with `clean_data_full`, `PivotTable1`, external PQ/cache structures. Key field evidence summarized above (Year/Years/period/category presence).
- **Repo:** `mko_treport/src/vbs/vbaProject.bin` and exports in `mko_treport/src/vbs/export/` present (maintainability; parity unverified). `mko_treport/src/scripts/` no treport runtime modules. Tracked `mko_treport/tests/treport_*` dirs (`config,files,filters,filter_application,refresh`) present as auxiliary/non-runtime artifacts.

## Agent role key

- **Auditor:** evidence/architecture/artifact investigation in code/workbooks/VBA where complexity/uncertainty exists.
- **Researcher:** Excel/PQ/pivot/VBA behavior and viable techniques where platform behavior matters.
- **Planner:** sequencing, detailed design shaping, and verification structure.
- **Validator:** independent review/verification gate for correctness/risk.
- **High-risk rule:** blocks marked **High risk** require **Auditor, Researcher, Planner, Validator** involvement as appropriate.

## R1–R9 traceability (rev. 5)

| Risk | Summary | Blocks / handling |
|---|---|---|
| R1 | Pivot topology, presented ranges, year field runtime behavior (base `Year` in all files; `report_wclip` also has grouped `Years`) | B2,B5,B8,B9,B13 (evidence-based/config-aligned to `Year`; runtime validation) |
| R2 | PQ refresh completeness/timing and connection portability | B2,B7,B13 (OQ-T2, record env/prereqs) |
| R3 | `min`/`max` member semantics in resolved context (incl. grouped `Years`) | B2,B5,B13 (log/validate resolved values) |
| R4 | `group_rows` previously conflicting — now **deferred/no-op** | B4,B10,B12 (ignore/no-op; not blocking) |
| R5 | Copy fidelity vs actual presented source state (no golden) | B2,B9,B10,B13 (functional/source comparison) |
| R6 | Locked/open files, permissions, save/overwrite | B6,B7,B11,B12,B13 (D1/D12, fail-fast) |
| R7 | VBA runtime behavior/parity and legacy control removal | B1,B2,B12,B13 (FR-RUN-1, OQ-T4) |
| R8 | Performance on large pivots | B13 (benchmark if needed; preserve semantics) |
| R9 | Environment prerequisites (Win10+/Excel 2021+, macros, PQ) | B0,B7,B13 (record; non-blocking where inapplicable) |

## Logical execution blocks (B0–B14)

### B0. Environment, artifact baseline & evidence capture
- **Objective:** establish reproducible local baseline for development/verification: confirm local artifact inventory and gitignored status, capture available Excel build/macro-trust/PQ prerequisites in test environment (non-blocking where inapplicable), define evidence storage (paths/hashes/snapshots) using backup/working-copy workflow.
- **Status:** **Completed (baseline checklist/evidence context available; environment details recorded as available).** Prior audit captured artifact inventory/state; exact target Excel build/PQ prerequisites to be recorded during runtime validation as encountered.
- **Depends on:** —
- **Risks:** evidence drift if local artifacts change without record.
- **Agents:** Auditor, Planner.

### B1. Control/workbook preparation and runtime-boundary alignment
- **Objective:** on a backed-up working copy of `treport.xlsm`, implement FR-RUN-1: create exactly two `run` controls/buttons `Refresh_data` and `Create_report` bound to same-named VBA procedures and remove legacy `RunXlconcat` button/obsolete workflow; ensure VBA export/snapshots can be refreshed for maintainability (auxiliary). Preserve runtime boundary (no external runtime dependencies).
- **Status:** **Ready / in scope.**
- **Depends on:** B0.
- **Risks:** Impl low-medium; Regression low-medium (remove legacy entry point). Compatibility low.
- **Agents:** Planner, Auditor, Validator.

### B2. Technical investigation & config alignment (evidence-driven, Excel/runtime)
- **Objective:** confirm/complete OQ-T1–T4 in target Excel as needed: per-file pivot object/field identities and member behavior (base `Year` present in all files; `report_wclip` also has grouped `Years`), presented ranges/layout, PQ refresh behavior/dependencies, and VBA runtime parity. Confirm config alignment to `Year` for all `data_files` rows (already applied by user) is correct in runtime, including `report_wclip`; validate optional-token handling context. Escalate only genuine residual business ambiguity with concrete alternatives.
- **Status:** **Substantially evidenced from package/XML; config aligned to `Year` for all files (user-updated). Runtime confirmation in target Excel still pending (esp. `report_wclip` behavior with grouped `Years` also present).**
- **Depends on:** B0 (parallel with B1 where useful).
- **Risks:** Impl/Compatibility high (Excel/PQ/pivot behavior). **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B3. (Removed as mandatory gate) Golden baseline — not required
- **Objective:** N/A as prerequisite. Use functional verification against actual source data per D18 during B13.
- **Status:** **Not applicable by PO decision.**

### B4. VBA config contract & validation (FR-CFG, D2,D12–D14,D17)
- **Objective:** implement deterministic VBA config reading/validation for `general,reports,data_files,placeholders`: required/non-empty checks, D17 dropped params ignored, `Trim` on configured strings, `group_rows` treated as no-op/deferred, fail-fast contextual errors per D12 with optional-token exception contract (D9/FR-FLT-5) respected.
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B1, B2 (field/config alignment inputs as needed).
- **Risks:** Impl/Regression medium.
- **Agents:** Planner, Validator, Auditor.

### B5. VBA `filters` parsing & dependency-aware resolution (FR-FLT-1/2/3, D5,D6,D10,D11)
- **Objective:** strict DSL parser (D5,D6) with `Trim`; resolve `min/max` and derive `max_*` per D10/D11 using verified per-file year/period field identities and available members in resolved context (no guessed mapping). Fail fast on invalid/unparseable/unresolvable/underivable cases per D12.
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B4, B2.
- **Risks:** Impl/Regression high; Compatibility medium-high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B6. VBA file resolution & access/session policy (FR-FILE, D1,D2)
- **Objective:** deterministic base-name resolution in `input_path` with precedence `.xlsx,.xlsm,.xlsb` (fail if none/>1), mode-specific open/save/close (Create read-only/never save; Refresh writable refresh/save/close), output path readiness, locked/missing handling (D12), and application/workbook state cleanup.
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B4, B2.
- **Risks:** Impl/Regression high; Compatibility medium-high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B7. VBA refresh workflow (FR-REF, OQ-T2)
- **Objective:** implement `Refresh_data` to refresh PQ and dependent pivot data to persisted current state (synchronous completion before save/close) per target Excel behavior; contextual fail-fast on errors; no stale-save.
- **Status:** **Not started (rev. 5 VBA); runtime-validated in target Excel.**
- **Depends on:** B6, B2.
- **Risks:** Impl/Rollout/Compatibility high (PQ behavior). **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B8. VBA filter application, staging & placement (FR-FLT-4/5/6/7, D4,D7–D9,D12)
- **Objective:** apply resolved `data_files.filters` once per data file and keep; per report row apply/reapply `articleLevel4Name_clean`/`targetAudience` without leak (D8); placement in-place or add-to-Filters (D7); exact-after-trim matching (D9); implement optional `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` skip-if-absent (must not fail) while requiring ≥1 valid value selected; fail fast on missing required field/value/match or no-data per D12.
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B5, B6, B2, B7 context where relevant.
- **Risks:** Impl/Regression/Compatibility high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B9. VBA pivot snapshot extraction & static paste (FR-RPT-3/4, D15, OQ-T3)
- **Objective:** capture pivot presented data range after filtering (labels and grand totals as shown; exclude pivot object/page-filter area/PQ connections) and paste static values+formats (+ widths if needed) at configured position; no pivot/PQ artifacts in output.
- **Status:** **Not started (rev. 5 VBA); fidelity validated in Excel vs actual source state.**
- **Depends on:** B8, B2.
- **Risks:** Impl/Regression high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B10. VBA formatting & layout (FR-RPT-5/6, D13)
- **Objective:** apply freeze panes, zoom, hide gridlines, and required row/column sizing/state per config; blanks/no-op; `group_rows` ignored as no-op (deferred, non-blocking).
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B4, B9.
- **Risks:** Impl/Regression medium.
- **Agents:** Planner, Validator, Auditor.

### B11. VBA report assembly, naming, placeholders & idempotency (FR-RPT-1/2/7, FR-NAME, D3,D11,D14)
- **Objective:** create one output workbook per report (`reports.name` + `.xlsx`, overwrite on success) with sheets in `data_files` order; substitute placeholders in `report_sheet` only (no zero-padding) using resolved `max_*`; apply `Trim`, sanitize Excel-invalid chars, enforce 31-char limit, fail fast on collisions; idempotent recreate/replace; save/close per D1.
- **Status:** **Not started (rev. 5 VBA).**
- **Depends on:** B5, B6, B9, B10.
- **Risks:** Impl/Regression high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B12. VBA diagnostics, run control & state safety (FR-RUN, D12)
- **Objective:** unified contextual fail-fast error taxonomy/outcomes; ensure `run` exposes only `Refresh_data`/`Create_report` with correct bindings/procedures (legacy removed); restore application/workbook state on completion/failure; respect optional-token exception and deferred `group_rows` no-op.
- **Status:** **Not started (rev. 5 VBA); incremental with B4–B11.**
- **Depends on:** B1, B4–B11.
- **Risks:** Impl/Regression medium.
- **Agents:** Planner, Validator, Auditor.

### B13. Verification & acceptance (functional, target Excel)
- **Objective:** execute verification against functional requirements and actual source data in target Excel (Windows 10+/Excel 2021+, recorded build/macro/PQ prerequisites as available): config/Trim/naming/collision/idempotency, file resolution/access modes, refresh persistence (OQ-T2), per-file period field alignment and `min/max`/`max_*` resolution (V-007/V-018), filter staging/placement/matching/no-leak, optional `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` skip behavior and ≥1-valid requirement (V-008), no-data/fail-fast cases, presented copy/format/layout fidelity (OQ-T3), required run controls/procedures and legacy removal (FR-RUN-1), and `group_rows` no-op. Use auxiliary tools/scripts as needed for inspection/comparison (non-runtime).
- **Status:** **Scenario preparation in scope; runtime execution pending VBA implementation and Excel validation. No golden workbook required.**
- **Depends on:** B1–B12, B2 runtime confirmations.
- **Risks:** Compatibility/Regression high. **High risk**.
- **Agents:** Auditor, Researcher, Planner, Validator.

### B14. Independent review & readiness
- **Objective:** independent review that delivered VBA behavior satisfies rev. 5 (D1–D18, runtime boundary, controls, deferred `group_rows`, Trim, optional filter, evidence-based field alignment) with verification evidence recorded or clearly bounded by any remaining runtime-validated unknowns; confirm no external runtime dependency introduced.
- **Status:** **Gated on B13 evidence.**
- **Depends on:** B1–B13.
- **Agents:** Validator (lead), Planner/Auditor/Researcher as needed.

## Execution order and dependency DAG

1. **B0** (baseline/evidence) — completed/informational.
2. **B1** (run controls + legacy removal) — depends B0. Can start immediately on working copy.
3. **B2** (runtime investigation + per-file config/field alignment evidence) — depends B0; parallel with B1/B4 design.
4. **B4** (VBA config/validation, Trim, `group_rows` no-op) — depends B1,B2 inputs.
5. **B5** (VBA filters parse/resolve, `min/max`/`max_*`, verified fields) — depends B4,B2. *(high risk)*
6. **B6** (VBA file resolution/access/session) — depends B4,B2. *(high risk)*
7. **B7** (VBA refresh) — depends B6,B2. *(high risk)*
8. **B8** (VBA filter apply/staging/placement + optional token) — depends B5,B6,B2. *(high risk)*
9. **B9** (VBA pivot snapshot/static paste) — depends B8,B2. *(high risk)*
10. **B10** (VBA formatting/layout, `group_rows` no-op) — depends B4,B9.
11. **B11** (VBA assembly/naming/placeholders/idempotency) — depends B5,B6,B9,B10. *(high risk)*
12. **B12** (diagnostics/state safety, control reconciliation) — incremental B4–B11, finalize before B13.
13. **B13** (functional verification in target Excel, no golden) — depends B1–B12 and B2 runtime confirmations. *(high risk)*
14. **B14** (independent review/readiness) — depends B13.

**Parallel opportunities:** B1 ∥ B2; B5 ∥ B6 after B4; B10/B12 can progress incrementally as extraction/assembly contracts stabilize; verification scenario design can proceed in parallel with implementation.

## Verification approach (D18, no golden workbook)

- **Primary method:** compare generated outputs against actual source pivot presented state/data in target Excel for the same resolved configuration/filters (values, labels/totals, sheet/workbook names/structure, formatting/layout properties as specified).
- **Evidence to record locally:** Excel build/version, macro-trust/PQ prerequisites relevant to run, fixed config snapshot (incl. any per-file filter alignment applied), input data/workbook state identifiers (paths/timestamps/hashes where practical), output artifacts, and observed results/screenshots/notes.
- **Auxiliary tooling:** may use Python/PowerShell/BAT or other tools to inspect workbooks, read metadata, compare extracted values/structures, and automate checks; these are non-runtime aids only.
- **Environment note:** do not block unrelated development if a specific environment detail is not yet recorded, provided it does not affect the current task; capture it when performing runtime-dependent validation.

## Cross-cutting risks and containment

- **Runtime boundary breach:** keep all end-user behavior in VBA; treat helpers/tests/scripts as auxiliary only.
- **Incorrect period-field mapping:** use evidenced per-file identities (`Year` for m/w, grouped `Years` for wclip), align config per row, validate in Excel, no guessed mapping.
- **Optional token mishandling:** `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` must be skipped if absent and never fail alone; require ≥1 valid selected value.
- **Legacy control remnants:** remove/replace `RunXlconcat` completely per FR-RUN-1.
- **`group_rows` scope creep:** enforce no-op/deferred behavior only.
- **PQ refresh staleness/portability:** validate synchronous completion and persisted results in target Excel; document prerequisites/paths issues.
- **Local-only artifacts:** backup before edits, preserve originals, avoid committing large Excel/generated files.

## Definition of Ready / implementation start (rev. 5)

- **Ready to proceed:** Yes. PO decisions D1–D18 (rev. 5) are specified; key clarifications (controls, Trim, `group_rows` deferral, optional filter, no-golden validation, auxiliary-tools policy, local handling) are resolved; V-007/V-018/V-008 have evidence-based determinations and config is aligned to `Year` for all files (user-updated).
- **Not blocking start:** recording exact Excel build/PQ prerequisites where inapplicable to a current task, and Excel-runtime confirmation details (OQ-T2/T3/T4) not yet executed — complete them during relevant implementation/verification steps.
- **Must complete for sign-off:** VBA implementation per rev. 5, runtime-verified `Year`-based field/config alignment in target Excel for all files including `report_wclip` (or explicitly escalated genuine business ambiguity with concrete alternatives), successful functional verification against actual source data including required controls/refresh/filter/copy/format/naming/idempotency/error behaviors, and confirmation of no external runtime dependency.

## Expected deliverables

- Updated VBA implementation in `treport.xlsm` (working copy, with backups/evidence) providing `Refresh_data`/`Create_report` per rev. 5 and removing legacy `RunXlconcat` workflow.
- Per-file config/field alignment applied as evidenced (esp. year-field filters/`max_year` derivation) and validated.
- Refreshed VBA exports/snapshots in `mko_treport/src/vbs/` as maintainability aids (auxiliary, parity noted).
- Optional auxiliary scripts/tests/notes for inspection and verification (non-runtime), kept consistent with repo/workspace conventions as appropriate.
- Local verification evidence (environment notes, run outcomes, comparisons against actual source data, issue resolutions) sufficient for B14 review/sign-off.

## Implementation Tasks

### Task B1 — Control/workbook preparation and runtime-boundary alignment (Implementor)

- **Task ID:** `001-treport-b1`
- **Block:** B1 (depends on B0 baseline/evidence; can proceed in parallel with B2 design/investigation)
- **Status:** Ready for implementation
- **Objective:** On a backed-up working copy of `treport.xlsm`, prepare the runtime control surface for rev. 5 by providing exactly two `run`-sheet entry points `Refresh_data` and `Create_report` bound to same-named VBA procedures, removing the legacy `RunXlconcat` control/obsolete workflow, and refreshing maintainability VBA exports/snapshots as auxiliary artifacts — without introducing any external runtime dependency.
- **Primary references:** `.ai/problems/001_treport_spec.md` rev. 5 — FR-RUN-1/FR-RUN-2/FR-RUN-3, D1, D16, T9 context, R7, OQ-T4, §9.1, §13.3; `.ai/plans/001-treport-execution-plan.md` rev. 5 — B1/B12/B13 scope boundaries.

#### Exact implementation scope

- **In scope (semantic units / artifacts):**
  - **Workbook artifact (working copy, local/binary):** `treport.xlsm` control workbook used for treport execution (evidence shows local/gitignored `data/theiss/treport.xlsm` and untracked `mko_treport/src/xls/treport.xlsm`; implement against the canonical working copy selected per local backup workflow and record which copy/hash was modified — do not commit workbook binaries unless explicitly requested and policy-compliant).
  - **Worksheet/UI surface:** `run` sheet (current evidence: `Sheet2`, one Form Control button “Button 1” caption “Run Macro” bound to `[0]!RunXlconcat`). End state: exactly two controls/buttons on `run` named/labeled `Refresh_data` and `Create_report`, each bound to the VBA procedure of the same name.
  - **VBA procedures (macro entry points):** `Public Sub Refresh_data()` and `Public Sub Create_report()` in a standard VBA module dedicated to treport run entry points/workflow orchestration (create a clearly named module if none exists, e.g. `modTreportRun`/`modRun`; avoid embedding business logic in sheet class modules). Procedure names must match control bindings exactly.
  - **Legacy removal:** remove/replace the legacy `RunXlconcat` button/control and its binding from `run`; remove or neutralize the obsolete `xlconcat_macro`/`RunXlconcat` workflow as an invocable run entry point (do not leave a third runnable legacy control). Retain or archive legacy implementation details only if needed for historical reference, but ensure no active UI/control routes to it.
  - **VBA project maintainability exports (auxiliary, repo-tracked):** `mko_treport/src/vbs/vbaProject.bin` and `mko_treport/src/vbs/export/*` (current: `Sheet1.cls`, `Sheet2.cls`, `ThisWorkbook.cls`, `xlconcat_macro.bas`) refreshed to reflect the post-change VBA project/control-relevant state as far as export tooling allows; record parity limitations per OQ-T4 (exports are auxiliary, runtime source of truth is `treport.xlsm`).
  - **Local evidence (auxiliary, not committed as large binaries):** backup of the original working workbook(s), short change note (which workbook copy edited, pre/post control inventory, macro names/bindings, export steps/results) stored in an appropriate local/workspace evidence location (e.g. `.ai/tmp/...`) if useful.
- **Out of scope for B1 (defer to later blocks):**
  - Implementing refresh PQ/pivot persistence semantics, config parsing/validation, filters/`min`/`max`/`max_*`, file resolution/access modes, pivot extraction/paste, formatting/layout, output assembly/naming/idempotency, and full diagnostics/state restoration beyond minimal safe stubs needed so the two required entry points exist and do not invoke legacy behavior (B4–B12).
  - Resolving OQ-T1–T3 runtime behavior in data workbooks (B2/B7–B9/B13).
  - Adding Python/PowerShell/BAT runtime dependencies for end users (auxiliary dev/inspection tools only).
  - Committing `data/theiss/*`, `treport.xlsm` working binaries, or generated reports.

#### Implementation sequence (dependency-safe)

1. **Preflight & backup (local):** identify canonical `treport.xlsm` working copy to modify; create timestamped backup(s) in a local path outside Git-tracked source (preserve original for recovery/comparison). Record macro-trust/Excel availability notes only if already known (do not block on unavailable details).
2. **Inspect current control/VBA state (semantic):** in `treport.xlsm` VBA project, confirm `run` sheet control(s) and assignments (expect legacy `RunXlconcat`/`xlconcat_macro`), existing module/class structure, and any object names/captions that could conflict with required names.
3. **Establish entry-point module/procedures:** add or update a standard module to declare `Public Sub Refresh_data()` and `Public Sub Create_report()` as the sole rev. 5 run entry points. If full behavior is not yet implemented (expected at B1), provide minimal, explicit, non-destructive stubs that fail fast with a clear contextual message indicating the operation is not yet implemented for rev. 5 (no legacy concatenation, no silent success) while preserving application state (align with FR-RUN-2 intent as applicable).
4. **Replace run-sheet controls:** remove legacy button/control bound to `RunXlconcat`; create/add exactly two controls on `run` sheet assigned to `Refresh_data` and `Create_report` (names/captions/procedure bindings unambiguous). Verify no other shape/button/control on `run` invokes macros (especially legacy).
5. **Retire legacy workflow routing:** ensure `RunXlconcat` is not reachable from `run` UI; update/rename/archive VBA artifacts as needed so the active workflow is rev. 5 entry points only (document if `xlconcat_macro` module remains in-project for reference but is not bound/called by UI).
6. **Save & refresh maintainability exports (auxiliary):** save `treport.xlsm`; export/refresh `mko_treport/src/vbs/vbaProject.bin` and `mko_treport/src/vbs/export/*.bas|*.cls` from the modified workbook using available Excel/VBA export tooling (auxiliary). Note any elements not representable in text exports (forms, control properties/bindings) as parity gaps per OQ-T4.
7. **Verify B1 acceptance criteria (manual/Excel + repo checks):** perform required checks below on the working copy; capture local evidence.

#### Architectural and safety constraints

- **Runtime boundary (mandatory):** final user-facing behavior must run in Excel/VBA only; no Python/PowerShell/BAT/external runtime dependencies introduced by B1.
- **Minimal/atomic change:** limit modifications to control surface, entry-point procedures/stubs, legacy-run decommissioning, and auxiliary exports needed for B1; do not refactor or implement unrelated report logic.
- **Semantic targeting:** identify workbook objects by sheet name (`run`), control names/captions/OnAction bindings, VBA procedure/module names, and ListObjects only as needed — never by line numbers or positional cell assumptions except where the `run` sheet is the designated control surface.
- **Local artifact handling:** work on backups/working copies; do not commit large Excel workbooks or generated outputs; keep evidence local.
- **Compatibility:** preserve existing `settings` schema/data and non-`run` workbook behavior unless explicitly required by B1; avoid breaking macro security/signature state beyond what editing requires (record observations).
- **State safety:** stubs/entry points must not leave Excel in a degraded global state on exit/error (respect FR-RUN-2 direction where practical; full centralized handling is B12).

#### Required verification / tests (proportional to risk; verify logic/interaction)

- **Control/UI (Excel, manual on working copy):**
  - `run` sheet contains exactly two macro controls/buttons: `Refresh_data` and `Create_report`.
  - Each control is bound to the VBA procedure with the identical name (`Refresh_data` / `Create_report`).
  - Legacy “Button 1” / `RunXlconcat` control and any binding to `[0]!RunXlconcat` or `xlconcat_macro.RunXlconcat` are absent from `run`.
- **VBA entry points (Excel/VBA IDE):**
  - `Public Sub Refresh_data` and `Public Sub Create_report` exist, compile in the project context, and are invocable from their controls.
  - Invoking either entry point does not execute legacy concatenation workflow; if stubbed, it returns/fails with a clear non-destructive “not yet implemented (rev. 5/Bx)” style message and does not corrupt workbook/application state.
- **Maintainability exports (repo/auxiliary):**
  - `mko_treport/src/vbs/export/` and `mko_treport/src/vbs/vbaProject.bin` updated from the modified workbook; exported module/class set reflects added/renamed entry-point module(s) as applicable.
  - Parity note recorded: control shapes/bindings may not be fully captured in text exports (OQ-T4) — runtime workbook is authoritative.
- **Regression/sanity:**
  - `settings` sheet and ListObjects (`general`, `reports`, `data_files`, `placeholders`) remain present and unchanged by B1 edits unless a specific B1 necessity is documented.
  - No external runtime dependency introduced (no scripts/tools required for button execution).

#### Deliverables and evidence

- Modified `treport.xlsm` working copy with required controls/procedures and legacy run control removed (local, backed up).
- Refreshed auxiliary VBA exports/snapshots in `mko_treport/src/vbs/` (+ parity note).
- Local B1 verification note/evidence (pre/post control inventory, bindings, stub behavior observed, backup path/hash if recorded).

#### Risks and gates

- **Risks:** binary workbook edit/control-binding mistakes (medium), residual legacy invocation paths (low-medium), VBA export/runtime parity gaps (known, OQ-T4), environment/macro-trust differences when validating (low).
- **Gate to B4+ integration:** B1 complete when FR-RUN-1 control/procedure surface is satisfied on the working copy and legacy `RunXlconcat` UI entry is removed; full operational behavior remains gated on B4–B13.

---
*End of execution plan (rev. 5).*

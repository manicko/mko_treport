# Execution Plan: P01 Closure — max-offset, default sheet removal, row grouping

- **Plan file:** `.ai/plans/006-p01-max-offset-default-sheet-grouping-closure.md`
- **Plan status:** Draft — closure/decomposition for remaining evidence, verification, and traceability (no rewrite of plan 005)
- **Date:** 2026-11-17
- **Scope mode:** Minimal, closure-focused, dependency-safe, semantic (no speculative redesign)

## 1. Context and source references (authoritative)

| Ref | Path / artifact | Role |
|---|---|---|
| S1 | `.ai/problems/003_p01_max_offset_default_sheet_grouping_spec.md` | Requirement baseline (CR-1..CR-6, D1–D7, OQ-1–OQ-3); **currently untracked** in git |
| S2 | `.ai/plans/005-p01-max-offset-default-sheet-grouping.md` | Existing execution plan/log (B0–B6 implementation largely recorded in exports/docs, B7 evidence incomplete); **currently untracked** in git; **must not be replaced or rewritten** |
| S3 | Step-0 Auditor code context (task input, 2026-11-17) | Current-state assessment: core P01 mechanisms appear implemented in exports/docs; L1 observed, L2/L3 not evidenced; artifact/sync ambiguity noted |
| S4 | `mko_treport/src/vbs/export/modTreportRun.bas`, `modTreportOutput.bas`, `modTreportFilters.bas`, `modTreportConfig.bas`, `modTreportPivot.bas`, `modTreportExtract.bas` | VBA export sources (semantic implementation units for P01) |
| S5 | `docs/01-vba-reporting/configuration-model.md`, `filters-dsl.md`, `pipeline-and-dataflow.md` | Updated contract/pipeline documentation to be reconciled against implemented behavior |
| S6 | `mko_treport/src/scripts/build_treport_vba.ps1`, `docs/02-build-and-deployment/vba-build.md` | L1 sync/build workflow (auxiliary), report artifacts |
| S7 | `.ai/tmp/build_treport_vba_20261117_092902_workcopy.txt`, `.ai/tmp/build_treport_vba_b7_20261117_141522.txt` | Observed L1 SUCCESS reports on timestamped working copies; compile not programmatically confirmed; L2 manual compile recommended/not recorded |
| S8 | Working-tree state (2026-11-17): modified `mko_treport/src/xls/treport.xlsm`, `docs/01-vba-reporting/*.md`; untracked `.ai/plans/005-...md`, `.ai/problems/003_...spec.md`, `.ai/problems/P03.md`, `.ai/audit/problems/...`; present `mko_treport/src/xls/~$treport.xlsm` and `mko_treport/src/vbs/vbaProject.bin` | Traceability/sync/lock context (environment-sensitive) |

**Evidence convention:** semantic units only (modules/procedures/types/ListObjects/config keys/workbook artifacts). No line numbers. Runtime Excel/VBE/Pivot/Outline behavior is **not** verified by static audit alone.

## 2. Closure objective and scope

**Objective:** Complete P01 (CR-1..CR-6, D1–D7) to an auditable, dependency-safe closure state by (a) reconciling current evidence and identifying any concrete export/doc gaps that still require fixes, (b) executing/capturing B7-equivalent verification evidence on working copies with an explicit **interactive L2 manual compile gate** and L3 behavioral/regression scenarios, (c) resolving or explicitly documenting artifact traceability and export-vs-binary sync considerations without speculative decisions, and (d) running a final validation and issue-fix loop until acceptance criteria are satisfied or blockers are explicitly recorded.

**Assumed baseline (from S3, to be confirmed in C0):** B0–B6 implementation appears largely present in VBA exports/docs for CR-1..CR-5; B7 verification/acceptance evidence is incomplete. Do not reimplement completed behavior without a concrete, evidenced gap.

**Out of scope (closure):** New P01 requirements, speculative redesign, non-P01 refactors, changes to canonical/original data workbooks for verification, and any claim of compilation/functional correctness without required evidence.

## 3. Constraints and verification policy (carried forward)

- **Source discipline:** Edit only VBA export sources `mko_treport/src/vbs/export/*.bas,*.cls` if a concrete gap is found; do not modify canonical workbooks blindly. No VBA/code or existing audit phase files are changed by this planning task.
- **Backups/working copies:** All build, compile, and functional checks use timestamped working copies/backups (e.g. `.ai/tmp/treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm`) and copied data where needed. Never corrupt originals.
- **L1 (automated sync/import):** Run `mko_treport/src/scripts/build_treport_vba.ps1` on a working copy when exports/binary sync must be evidenced; record exact command, report path, and result summary (SUCCESS/ISSUES/ERROR, VBOM/compile notes).
- **L2 (manual VBE compile) — mandatory to claim compilation:** A human/operator must open the **identified working copy** in Excel and run **Debug → Compile VBAProject**. Evidence must record: working-copy path, date/time, operator/context if available, and explicit PASS/FAIL/not performed. **This is an interactive environment/user-executed gate; automation may prepare artifacts but cannot substitute for L2. Do not claim "compiled successfully" without explicit L2 pass evidence.**
- **L3 (functional checks on copies):** Required because behavior changed. Execute representative scenarios for CR-1..CR-5, D1–D7, and key regressions on **copies only**; capture expected vs observed outcomes and artifact identifiers. **Do not claim L3 done without recorded evidence.**
- **Environment prerequisites/risks:** Windows + Excel COM automation, VBOM trust ("Trust access to the VBA project object model"), macro/security policies, file locks, and interactive Excel session availability affect L1/L2/L3. `mko_treport/src/xls/~$treport.xlsm` indicates Excel may have the workbook open/locked — this must be checked before workbook-targeted operations.

## 4. Execution blocks (closure)

### Block C0 — Evidence reconciliation and concrete gap triage
- **Objective:** Establish a verified closure baseline: confirm which P01 requirements are satisfied by exports/docs, identify any concrete remaining implementation/documentation gaps, and freeze the minimum fix set (or explicitly record "no code/doc fix required") before further verification.
- **Scope (semantic):**
  - Read-only reconciliation of S1–S8 against VBA export modules/procedures/types (`modTreportRun.Create_report` / `ProcessDataFile` / placeholder resolution; `modTreportOutput` cleanup/placeholder/naming APIs; `modTreportFilters` offset/bound/domain logic; `modTreportConfig` `TreportGeneral.rowsCollapsed` / `TreportDataFile.group_rows`; `modTreportPivot` numeric domain/max helpers; `modTreportExtract` snapshot bounds/grouping integration).
  - Reconcile docs contracts in S5 with export behavior and S1 decisions (D1–D7, OQ-1/OQ-2 canonical forms as implemented).
  - Inventory L1/L2/L3 evidence artifacts and traceability state (untracked plan/spec, modified workbook/binary, lock).
- **Dependencies:** None (entry block).
- **Execution order:** 1.
- **Risks:** False confidence from static audit (runtime/COM/Pivot/Outline behavior unproven); missed edge-case gap in D1–D7; environment/lock ambiguity mistaken for code issue.
- **Alternatives (relevant; do not choose if uncertain):**
  - **Gap handling path A (preferred if evidence supports):** no export/doc code changes; proceed directly to C2/C3/C4 with verification and traceability tasks.
  - **Gap handling path B:** targeted, minimal fixes in specific export modules/docs only where a concrete S1/S3 mismatch is evidenced; require C1 before C2.
  - **Uncertainty path:** if runtime behavior cannot be inferred from exports/docs, record uncertainty and defer to C2 L3 scenarios rather than speculative redesign.
- **Required agents:**
  - **Auditor (required):** deep, dependency-aware reconciliation of VBA/workbook/doc artifacts and identification of concrete gaps/uncertainties.
  - **Researcher (optional, conditional):** required only if C0 reveals unresolved Excel/VBA/PQ/pivot/Outline/COM behavior uncertainty affecting acceptance interpretation.
  - **Planner (required):** maintain minimal closure sequencing and acceptance mapping.
  - **Validator (recommended):** independent check that claimed gaps/non-gaps are evidence-based and no speculative scope is introduced.
  - **Implementor/Doc-specialist:** only if C1 is triggered by evidenced gaps.
- **Acceptance / verification criteria:**
  - Written reconciliation outcome states, per CR-1..CR-6/D1–D7: satisfied-by-evidence, gap-found (with semantic unit + required fix), or unverified-pending-runtime.
  - Explicit list of concrete files/semantic units requiring change, or explicit "no implementation/doc changes required" decision with rationale.
  - L2/L3 status recorded as **not done/not evidenced** unless new evidence exists; no compile/behavior claims made.

### Block C1 — Targeted remediation for evidenced gaps only (conditional)
- **Objective:** Apply minimal, isolated fixes to any concrete gaps identified in C0 while preserving P01 contracts and existing valid behavior.
- **Scope (semantic):** Only the specific export modules/procedures/types and/or specific governed docs identified in C0 (e.g. `modTreportRun`, `modTreportOutput`, `modTreportFilters`, `modTreportConfig`, `modTreportPivot`, `modTreportExtract`, `docs/01-vba-reporting/configuration-model.md`, `filters-dsl.md`, `pipeline-and-dataflow.md`). No broad rewrites.
- **Dependencies:** C0 with evidenced gap(s). **Skipped if C0 records no required changes.**
- **Execution order:** 2 (before L1/L2/L3 evidence capture for the fixed state).
- **Risks:** Regression in existing valid filter/placeholder/naming/paste behavior; introducing coupling across CR areas; changing D1–D7 semantics unintentionally; export/workbook drift if not re-synced.
- **Alternatives (relevant; do not choose if technical uncertainty exists):**
  - Fix in parsing/resolution layer vs orchestration/integration layer depending on root cause (choose based on C0 evidence; if ambiguous, escalate to Auditor/Researcher/Validator before coding).
  - Doc-only correction vs export+doc correction when behavior is correct but contract text is inaccurate.
- **Required agents:**
  - **Auditor (required):** confirm root cause and affected semantic boundaries.
  - **Researcher (required if uncertainty):** validate Excel/VBA/COM/pivot behavior assumptions behind the fix options.
  - **Planner (required):** define atomic fix tasks and regression checks.
  - **Validator (required):** independent review of proposed fix/design before and/or after change due to regression/compatibility risk.
  - **Implementor/Doc-specialist (required):** perform minimal edits in allowed sources only.
- **Acceptance / verification criteria:**
  - Each fix traces to a C0 gap and S1 requirement/decision.
  - No unrelated files/behaviors modified; semantic targeting preserved.
  - Post-fix state ready for C2 L1 on a fresh timestamped working copy.

### Block C2 — B7-equivalent verification and evidence capture (L1/L2/L3 on working copies)
- **Objective:** Produce auditable L1/L2/L3 evidence for the closure candidate (post-C0 or post-C1 state) covering CR-1..CR-5, D1–D7, and key regressions without using original data/workbooks.
- **Scope (semantic):**
  - **L1:** `mko_treport/src/scripts/build_treport_vba.ps1` execution on timestamped working copy(ies); `.ai/tmp/build_treport_vba_<timestamp>.txt` report capture; component/import/legacy/duplicate/VBOM/compile-note summary.
  - **L2 (interactive gate):** manual Excel VBE **Debug → Compile VBAProject** on the identified working copy by user/operator.
  - **L3:** functional scenarios on copied data/workbooks exercising: CR-1/D3 (no default empty sheet in normal success; produced sheets preserved; no-data/fail-fast alignment), CR-2/D1/D2 (offset scalar/range, `N==0`, boundary/domain failures with context, existing expression regressions), CR-3/D1/D2 (pre-filter independence, `<max_field>` and `<max-N_field>` valid/invalid/domain-missing, formatting/sanitization/collision guards), CR-4/D5/D6 (bounds-scoped contiguous grouping, fill rules/no-fill/white boundaries, single/edge runs, multi-subreport independence, invalid/out-of-width silent skip), CR-5/D4 (`rows_collapsed` absent/False/True/default-expanded behavior), CR-6/D7 (general applicability/non-regression).
- **Dependencies:** C0 complete; C1 complete if triggered. For authoritative L1/L2/L3, the working copy must reflect the intended export state (coordination with C3 may be required if workbook sync is in question; see alternatives).
- **Execution order:** 3.
- **Risks (must assess):**
  - **Excel/Windows environment:** Office build/locale differences, Excel COM automation stability, interactive session requirements.
  - **Macro security/VBOM:** trust/policy blocks can prevent L1 import or mislead compile diagnostics.
  - **Workbook state/locking:** `mko_treport/src/xls/~$treport.xlsm` lock and open Excel instances can block save/copy/build/sync.
  - **Power Query refresh timing:** refresh completion/state may affect pivot availability and L3 reproducibility.
  - **Pivot/object-model behavior:** PivotItems/domain extraction, numeric coercion, outline/grouping COM side effects, and worksheet/sheet-count behavior vary by runtime state.
  - **Output fidelity:** sheet naming/sanitization/length/collision, paste bounds/formatting, grouping visibility/collapse state, and absence of default sheets must be validated on actual outputs.
  - **Data representativeness:** copied datasets may not cover all D1–D7 edge cases unless explicitly selected.
- **Alternatives (relevant; do not choose final path when technical uncertainty exists):**
  - **Working-copy source:** build from `mko_treport/src/xls/treport.xlsm` vs from another known-good workbook baseline; if parity is uncertain, record both options and require C3 traceability decision/evidence before claiming definitive L2/L3 for canonical sync.
  - **L3 data selection:** use existing copied/sanitized datasets under ignored/temp paths vs construct minimal synthetic copies; choose based on availability and coverage, but never originals.
  - **Compile confirmation:** programmatic compile attempt (best-effort) plus mandatory manual L2 vs manual L2 only; manual L2 remains the acceptance gate regardless.
  - **If interactive L2/L3 cannot be executed now:** produce L1 artifacts + explicit blocker record (environment, lock, VBOM, operator unavailable) and stop short of compile/behavior claims.
- **Required agents:**
  - **Auditor (required):** define/verify scenario coverage against CR/D matrix and inspect produced evidence artifacts.
  - **Researcher (required when uncertainty):** interpret Excel/Pivot/PQ/Outline/COM anomalies and distinguish environment issues from product defects.
  - **Planner (required):** sequence L1→L2→L3, select minimal sufficient scenario set, and define fix-loop triggers.
  - **Validator (required, high risk):** independent review of evidence sufficiency, regression coverage, and correctness of pass/fail conclusions.
  - **Implementor/support (conditional):** assist with script runs or targeted fixes if failures are evidenced (feeds C4).
- **Acceptance / verification criteria:**
  - **L1:** exact command, working-copy path, and `.ai/tmp` report path recorded with clear result summary; VBOM/compile notes captured (not treated as L2 proof).
  - **L2 (environment/user-executed gate):** explicit record of PASS with working-copy identity and date/time, or explicit NOT DONE/FAIL with blocker/error details. **No compile claim without PASS.**
  - **L3:** per-scenario matrix for CR-1..CR-5, D1–D7, and regressions with pass/fail, observed outputs, and copy identifiers; all runs on copies only. **No L3 claim without recorded evidence.**
  - Failures classified as implementation defect, documentation mismatch, environment/blocker, or insufficient evidence (drives C4).

### Block C3 — Artifact traceability and export-vs-binary sync reconciliation
- **Objective:** Make closure artifacts auditable and reduce export-vs-binary ambiguity by explicitly reconciling plan/spec/docs/workbook/binary states and recording required sync/commit actions as traceable tasks (without speculative implementation decisions).
- **Scope (semantic):**
  - Traceability of closure artifacts: `.ai/plans/006-p01-max-offset-default-sheet-grouping-closure.md` (this plan), existing `.ai/plans/005-...md` and `.ai/problems/003_...spec.md` (untracked), and any verification evidence records produced in C2.
  - Binary/sync considerations: `mko_treport/src/xls/treport.xlsm` (modified, possibly locked by `~$treport.xlsm`) and `mko_treport/src/vbs/vbaProject.bin` relative to `mko_treport/src/vbs/export/*.bas,*.cls` and L1/L2-verified working copies.
  - Documentation/workflow alignment: `docs/02-build-and-deployment/vba-build.md` and project VBA workflow constraints (L1/L2/L3, working copies, VBOM).
- **Dependencies:** C0 (inventory); best executed in parallel with C2 preparation but final sync/commit decisions depend on C2 outcomes and environment availability (especially workbook lock/VBOM/L2).
- **Execution order:** 3 (parallel-prep) / 4 (finalization actions) depending on evidence.
- **Risks:**
  - **Export-vs-binary ambiguity:** modified `treport.xlsm` and existing `vbaProject.bin` may not match current exports; cannot be fully proven from text-only audit.
  - **Workbook lock/open state:** `~$treport.xlsm` suggests Excel lock; sync/build/save operations may fail or capture wrong state.
  - **Untracked spec/plan docs:** implementation commits exist but planning/spec docs not committed, weakening audit trail.
  - **Overreach risk:** syncing canonical binaries or committing artifacts without explicit task/approval/evidence.
- **Alternatives (relevant; do not choose final approach when technical uncertainty exists):**
  - **Sync target:** sync exports to a fresh working copy only (safe, evidence-oriented) vs sync/update canonical `mko_treport/src/xls/treport.xlsm` and/or `vbaProject.bin` (requires explicit approval/context, unlocked workbook, and post-sync L1/L2 evidence).
  - **Doc tracking:** commit untracked `.ai/plans/005`, `.ai/problems/003` and this closure plan as audit/plan artifacts vs leave untracked per repository/workflow decision — record as explicit action item, not assumed.
  - **Lock handling:** close Excel/retry on copies vs defer workbook-touching steps; do not force/delete lock artifacts.
  - **Parity proof:** rely on L1/L2 on a working copy built from exports vs require additional binary comparison/manual workbook inspection — choose based on available environment and risk.
- **Required agents:**
  - **Auditor (required):** map artifact states and identify exact traceability gaps.
  - **Planner (required):** define safe, ordered sync/recording tasks and approval gates.
  - **Validator (required for high-risk sync decisions):** review any proposed canonical-binary sync/commit rationale.
  - **Researcher (optional):** only if build/VBOM/COM sync behavior is unclear.
  - **Implementor/Doc-specialist (conditional):** execute only explicitly approved, minimal, evidenced actions.
- **Acceptance / verification criteria:**
  - Traceability matrix recorded: each key artifact (S1–S8, L1 reports, L2 record, L3 matrix) marked as present/missing/untracked/modified/locked with implications.
  - Export-vs-binary state explicitly classified as **verified**, **unverified**, or **blocked** with reason and required next action.
  - Any canonical workbook/`vbaProject.bin` sync action is listed as gated (approval + unlocked environment + post-sync L1/L2 evidence) or explicitly out of scope for this run.
  - No speculative claim that embedded VBA matches exports without evidence.

### Block C4 — Final validation and issue-fix loop (closure gate)
- **Objective:** Independently validate closure readiness, drive remediation loops for any failures/gaps found in C1–C3, and produce the final P01 closure assessment.
- **Scope (semantic):** Cross-check C0–C3 outputs against S1 acceptance intent (CR-1..CR-6, D1–D7), L1/L2/L3 evidence quality, regression coverage, and artifact traceability. If issues remain, feed minimal fixes back through C1 (code/doc) and/or C2 (re-verification) and/or C3 (traceability/sync records) until resolved or explicitly blocked.
- **Dependencies:** C2 and C3 completed to the extent possible (including explicit blocker records where environment prevents completion).
- **Execution order:** 5 (iterative loop as needed).
- **Risks:**
  - **Residual regression/compatibility risk:** fixes for one CR break existing expressions, placeholder behavior, layout, or grouping bounds.
  - **Incomplete evidence accepted accidentally:** L1 notes mistaken for L2, or partial L3 mistaken for full D1–D7 coverage.
  - **Environment-driven non-determinism:** flaky Excel/COM/PQ behavior obscures true pass/fail.
  - **Traceability drift:** final report not tied to concrete artifact paths/timestamps.
- **Alternatives (relevant; do not choose final implementation approach when technical uncertainty exists):**
  - **On failure:** targeted semantic fix + focused re-test vs broader rollback/reassessment of affected CR block; choose minimal safe option based on evidenced root cause.
  - **On environment blocker:** record explicit blocker with reproducible context and stop vs attempt alternative working-copy/environment path if available and safe.
  - **On sync ambiguity:** accept export+working-copy evidence with documented binary ambiguity vs gate closure on canonical sync (decision depends on approval/environment; do not assume).
- **Required agents:**
  - **Validator (required):** independent closure review and pass/fail/blocker determination (high risk).
  - **Auditor (required):** confirm evidence artifacts, scenario coverage, and traceability.
  - **Planner (required):** control loop scope, stop conditions, and minimal remediation tasks.
  - **Researcher (required when failures/behavior are ambiguous):** distinguish implementation defects from Excel/PQ/pivot/COM/environment behavior and propose viable investigation paths without forcing a final uncertain fix.
  - **Implementor/Doc-specialist (conditional):** execute only evidenced minimal remediations.
- **Acceptance / verification criteria:**
  - Final closure record states, per CR-1..CR-6 and D1–D7: **PASS (evidenced)**, **FAIL (evidenced defect)**, or **BLOCKED (explicit reason/evidence path)**.
  - **L2:** PASS recorded for identified working copy/date-time **or** explicitly not done/blocked (never claimed done).
  - **L3:** scenario matrix recorded with outcomes on copies **or** explicitly not done/blocked.
  - L1 report path(s), working-copy identifiers, and any fix-loop iterations documented.
  - No unresolved concrete gap left undocumented; no speculative redesign introduced.

## 5. Dependency and execution order

| Order | Block | Depends on | Gate / notes |
|---|---|---|---|
| 1 | C0 Evidence reconciliation and concrete gap triage | — | Establish baseline, gap list/no-gap decision, evidence inventory |
| 2 | C1 Targeted remediation for evidenced gaps only (conditional) | C0 (gap evidenced) | Skip if no gaps; minimal semantic fixes only |
| 3 | C2 B7-equivalent verification and evidence capture (L1/L2/L3) | C0, and C1 if run | **L2 is interactive user/operator gate**; L3 on copies only; may be blocked by environment/lock |
| 3/4 | C3 Artifact traceability and export-vs-binary sync reconciliation | C0; final actions informed by C2 | Parallel prep allowed; canonical sync/commit actions gated by approval, unlock, and evidence |
| 5 | C4 Final validation and issue-fix loop | C2, C3 (as available) | Iterate C1/C2/C3 as needed; closure = evidenced pass/fail/blocked |

**Critical closure path:** C0 → (C1 if needed) → C2 → C4, with C3 providing traceability/sync gates throughout. If C0 finds no code/doc gaps, path becomes C0 → C2 → C3/C4 finalization.

## 6. Cross-cutting verification matrix (closure)

| Requirement / decision area | Primary blocks | L1 (working copy, scripted) | L2 (manual VBE, user/operator gate) | L3 (copies only) minimum closure intent |
|---|---|---|---|---|
| CR-1, D3 (default sheet removal) | C0,C1?,C2,C4 | Required when exports/state changed | Required to claim compile | No Excel-default empty sheet in normal success; produced sheets preserved; no-data/fail-fast alignment |
| CR-2, D1,D2 (`max-N` in filters) | C0,C1?,C2,C4 | Required | Required to claim compile | `max-N` scalar/range, `N==0`, D1/D2 failures with context, regressions for existing valid expressions |
| CR-3, D1,D2 (`max-N_<field>` placeholders) | C0,C1?,C2,C4 | Required | Required to claim compile | Pre-filter/filter-independence, token/remnant guards, valid/invalid/domain-missing offsets, naming/sanitization/collision |
| CR-4, D5,D6 (`group_rows` Outline grouping) | C0,C1?,C2,C4 | Required | Required to claim compile | Bounds-scoped contiguous grouping, no-fill/solid-white rules, single/edge runs, multi-subreport independence, invalid/out-of-width silent skip |
| CR-5, D4 (`rows_collapsed`) | C0,C1?,C2,C4 | Required | Required to claim compile | Absent/blank/False/0/No => expanded; True/1/Yes => collapsed readiness/behavior as implemented; invalid present value handling per config contract |
| CR-6, D7 (scope, determinism, non-regression) | C0–C4 | Required | Required to claim compile | General applicability across relevant flows and preservation of existing valid behavior except explicit P01 changes |
| Artifact traceability / export-vs-binary sync | C0,C3,C4 | Evidence/input | Post-sync confirmation if canonical artifacts changed | N/A (process evidence); record verified/unverified/blocked |

## 7. Agent requirement summary (by block)

| Block | Risk level | Auditor | Researcher | Planner | Validator | Implementor/Doc-specialist | Justification |
|---|---|---|---|---|---|---|---|
| C0 | High (evidence/uncertainty) | Required | Conditional | Required | Recommended | Conditional | Closure baseline, gap detection, and avoidance of speculative claims/changes |
| C1 | High (implementation/regression) | Required | Required if uncertainty | Required | Required | Required if run | Minimal fixes only for evidenced gaps with regression control |
| C2 | High (environment/runtime/verification) | Required | Required when uncertainty | Required | Required | Conditional | L1/L2/L3 evidence on copies; Excel/COM/PQ/pivot/lock/fidelity risks; L2 is user-executed gate |
| C3 | High (traceability/sync) | Required | Optional | Required | Required for canonical sync decisions | Conditional | Export-vs-binary ambiguity, modified/locked workbook, untracked plan/spec artifacts |
| C4 | High (closure decision) | Required | Required when ambiguous | Required | Required | Conditional | Independent final pass/fail/blocked determination and controlled fix loop |

**Note:** Because overall closure risk is high (unverified runtime behavior, environment dependencies, and artifact sync ambiguity), all listed core agents (Auditor, Researcher, Planner, Validator) are involved across the closure path, with Researcher/Implementor conditional on concrete uncertainty or evidenced work.

## 8. L2 interactive gate — execution and evidence protocol (explicit)

- **Gate owner:** user/operator with access to Excel UI on Windows (environment-dependent).
- **Preconditions:** identified timestamped working copy (`.ai\tmp/...xlsm`), Excel closed/unlocked as needed, macros/VBA project access permitted (VBOM trust) where required, no use of canonical/original data for destructive checks.
- **Action:** open working copy in Excel → VBA Editor → **Debug → Compile VBAProject**.
- **Evidence to record (mandatory fields):**
  - Working-copy full path and identity (timestamp/build context)
  - Date/time (and operator/context if available)
  - Result: **PASS** / **FAIL** / **NOT PERFORMED (BLOCKED)**
  - If FAIL/BLOCKED: exact error text/screenshot-equivalent notes if available, suspected cause (VBOM trust, lock, missing refs, code error), and next action
- **Reporting rule:** state "L2 passed (manual VBE Compile)" only with PASS evidence; otherwise state L2 not done/failed and that compilation is **not claimed**.

## 9. Notes and non-goals

- Do not replace or rewrite `.ai/plans/005-p01-max-offset-default-sheet-grouping.md`; this document is a closure companion.
- Do not modify VBA/code or existing audit phase files as part of producing this plan.
- Treat `mko_treport/src/xls/~$treport.xlsm` as a potential active-lock signal; verify before any workbook operation.
- Do not commit `data/theiss` binaries/results; keep working copies/reports in `.ai/tmp` or ignored paths as appropriate.
- Prefer smallest safe closure actions: verify and document first; change code/docs/binaries only when evidenced and gated.

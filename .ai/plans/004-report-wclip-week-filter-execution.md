# Execution plan: 004 — report_wclip week filter / rev.7 placeholders / D2,D9 alignment (decomposed)

- **Plan ID:** 004-report-wclip-week-filter-execution
- **Status:** Step 1 decomposition (draft for downstream design/validation/execution)
- **Date:** 2026-11-16
- **Step:** Decompose Plan (workflow Step 1)
- **Parent/inputs:**
  - `.ai/plans/003-report-wclip-week-filter.md` (audit-reviewed intent/tasks)
  - `.ai/problems/001_treport_spec.md` (**rev.7, normative**)
  - `{code_context}` from Auditor Step 0 (read-only, provided in-session)
- **Scope:** Execution decomposition, dependencies, risks, alternatives, and per-block agent assignments only. **No VBA/source-code changes, no workbook edits, no build runs in this step.**
- **Constraints:** Semantic anchors only (no line numbers). VBA runtime-only in Excel; auxiliary tools non-runtime. Working copies/backups mandatory for any future L1/L2/L3/build; never claim compile from L1 alone (L2 VBE required). Treat `mko_treport/src/vbs/vbaProject.bin` (modified/uncommitted) and possible export↔workbook drift as rollout risk. Minimal, evidence-backed scope; no speculative redesign.

## 1. Current-state baseline (from Step 0 / plan 003)

- **Provisional alignment in exports:** generic filter parsing/classification and per-field pre-apply numeric-domain resolution (D10 direction), uniform selection path with fail-on-missing-field and fail-on-none-matchable behavior in core selection flow (D9 direction), fail-fast on empty pivot data, sanitization/collision/static snapshot basics, `group_rows` no-op, run orchestration/state handling.
- **Critical/normative gaps:** rev.7 `report_sheet` placeholders `max_<PivotTable field name>` **not implemented** (legacy fixed tokens/args and filter-representative coupling remain; observed workbook strings may use rev.7-like tokens → migration/mismatch risk). **D2 regression in exports:** current behavior is non-recursive root-only search (commit `ea6409eb`), which conflicts with updated normative D2 requiring deterministic **recursive** search within `input_path` subtree, allowed extensions `.xlsx/.xlsm/.xlsb` as a fixed set without precedence tie-breaker, and fail-fast on 0 or >1 total candidates with diagnostics. **`report_wclip` (U2)** correctness/axis behavior unresolved and evidence-dependent. FR-RUN-1/FR-DEP/build and all runtime behavior **unverified** (no L1/L2/L3 evidence).
- **Implication:** plan 003 is directionally correct but execution must be gated by evidence, design clarification where uncertain, migration inventory, and disciplined verification.

## 2. Execution principles (apply to all blocks)

- **Dependency safety:** do not start implementation blocks that depend on unresolved normative/semantic uncertainty without a recorded decision/evidence gate.
- **Isolation:** prefer changes separable by concern (D2 resolution, placeholder resolution contract, D9 hardening, config/migration, U2 handling) with independent reviewability where feasible.
- **Risk-proportional process:** higher-risk blocks require Auditor/Researcher/Planner/Validator involvement as specified; low-risk follow-ups may be leaner but never bypass L1/L2/L3 gates for logic changes.
- **No invented requirements:** alternatives recorded; final approach **not chosen** when technical uncertainty remains (must be resolved in design/evidence blocks before or during gated implementation).
- **Verification discipline:** L1 build report on working copy; **L2 manual VBE compile** required for compile claims; L3 on copies only for logic/run behavior; workbook-level checks for FR-RUN-1/outputs; record commands, report paths, timestamps, before/after template strings.

## 3. Dependency graph and execution order

```text
B0 (Preconditions, drift & inventory, baseline evidence)
  ├─> B1 (Decision/design gates: placeholder semantics point, D9 edge cases, D2 bounds, U2 hypotheses) 
  │     ├─> B2 (D2 conformance implementation)
  │     ├─> B3 (Rev.7 placeholder contract implementation, critical)
  │     ├─> B4 (D9 alignment confirmation/hardening)
  │     └─> B6 (U2 report_wclip resolution; may be config-only or minimal code per evidence)
  │            └─> B5 (Template/config migration inventory & apply) [depends on B3 contract + B6 outcome]
  └─> B7 (L1/L2/L3 verification, regression matrix, FR-RUN-1/FR-DEP evidence) [after B2-B6 as applicable]
         └─> B8 (Controlled sync/rollout of canonical artifacts) [after B7 pass + explicit approval]
```

**Parallelization (safe, if resourced):** after B0+B1 decisions, B2 and B4 can proceed independently if they touch disjoint semantics; B3 is critical-path and may require coordination with B5/B6; B6 evidence work can overlap B2/B4 but its resolution may constrain B3/B5. No parallel block may bypass shared decision gates.

## 4. Block decomposition

### B0 — Preconditions, drift control, and evidence inventory
- **Objective:** establish a safe, auditable baseline before design/implementation: confirm artifact/workbook inventory, drift, and collect missing evidence needed for U2/migration/D2/FR-RUN-1.
- **In scope:** read-only inventory of `mko_treport/src/vbs/export/*`, `mko_treport/src/vbs/vbaProject.bin`, `mko_treport/src/xls/treport.xlsm`, `data/theiss/treport.xlsm` (and related config-bearing copies if present), `data/theiss/{base,monthly,weekly,output}` presence, build script/docs; identify modified/untracked state; extract/record all `data_files.report_sheet` templates and placeholder tokens, relevant `filters`, and run-control presence where observable without unsafe edits; capture U2 baseline symptoms/expected period/axis evidence on copies if available.
- **Out of scope:** code/workbook modification, claiming L2/L3, speculative fixes.
- **Impacted components (semantic):** deployment/build artifacts, workbook `settings`/`run` artifacts, source data context for evidence only.
- **Risks:** **High** (drift/export↔binary↔workbook mismatch, gitignored/local-only artifacts, incomplete visibility into workbook internals, environment-specific PQ/pivot behavior). **Medium** (misreading template inventory leading to incomplete migration).
- **Uncertainties / alternatives:** workbook internals may require Excel/manual inspection on working copies; inventory method alternatives: (a) export/table extraction via auxiliary script/tooling if available and non-runtime, (b) manual documented inspection in Excel, (c) hybrid. **Do not assume completeness from OOXML strings alone.**
- **Required agents:**
  - **Auditor:** **Required (lead)** — deep investigation of drift, artifact parity, template/config inventory, U2 symptoms, FR-RUN-1 evidence gaps.
  - **Researcher:** Optional — only if tool/Excel-object-model inspection approaches for evidence collection need support implications review.
  - **Planner:** Optional — to refine block sequencing if major new constraints found.
  - **Validator:** Optional — spot-check inventory completeness/risk log for high-impact gaps.
- **Entry criteria:** Step 0 `{code_context}` available (done); read access to repo paths.
- **Exit criteria:** documented inventory (workbooks/config locations, all in-scope `report_sheet` templates & tokens, relevant filters, control presence signals), drift/risk register, explicit U2 evidence gaps, and confirmed working-copy/backup locations for future steps.

### B1 — Decision and detailed design gates (resolve/contain uncertainty)
- **Objective:** produce implementation-ready semantic decisions for uncertain areas without choosing prematurely where evidence is insufficient; define acceptance of alternatives and escalation points.
- **In scope:** clarify (a) **rev.7 placeholder resolution contract** implementation semantics consistent with D11: token grammar `max_<PivotTable field name>` in `report_sheet`, field identity/lookup rules, numeric-convertible domain and "maximum" definition, deterministic resolution point/timing in Create flow **independent of `data_files.filters` derivation/representatives**, behavior when field missing/unusable/no numeric values, and legacy token removal/migration policy; (b) **D9** edge cases (trim/numeric equivalence, multi-token subsets, report vs data-file contexts, error context requirements) vs current exports; (c) **D2** normative behavior (recursive search within `input_path` subtree, allowed extensions `.xlsx,.xlsm,.xlsb` as fixed set without precedence tie-breaker, deterministic single-match or fail-fast on 0/>1 with diagnostics) and correction of the current non-recursive regression (`ea6409eb`); (d) **U2** hypotheses and minimal remediation options (filter-only/config vs axis/layout-aware minimal change) with evidence requirements; (e) module/API impact map (semantic) and task slicing for B2–B6.
- **Out of scope:** code changes; final selection among viable implementation approaches where uncertainty remains unresolved by evidence.
- **Impacted components (semantic):** `modTreportOutput` (substitution contract), `modTreportFilters`/`ResolvedContext` (decouple legacy aliases), `modTreportRun` (resolution sequencing/context), `modTreportPivot` (field/domain lookup support as needed), `modTreportFiles` (D2), workbook config/templates, build/verify process.
- **Risks:** **High** (wrong placeholder semantics/timing breaks rev.7 determinism; D2 behavior change can affect existing environments; U2 misdiagnosis). **Medium** (API ripple across modules).
- **Uncertainties / alternatives (must remain open until evidence/decision):**
  - **Placeholder resolution source:** alternatives include (a) derive max from pivot field items/numeric domain in current pivot state at a defined pre/post filter-application point consistent with "relevant pivot/data context for processed `data_files` row" and independent of filter representatives, (b) derive from pivot cache/available field data independent of current visible selection state, (c) hybrid with explicit precedence/fail-fast rules. **Choice deferred** pending spec-consistent interpretation and Excel behavior evidence.
  - **Token parsing/field naming:** support exact configured field names including spaces/special chars inside `<max_...>`; alternatives for delimiter/escaping if field names contain `>` — **evidence/design needed** (prefer minimal deterministic grammar aligned with spec examples).
  - **D2 location scope (DECIDED 2026-11-16, normative):** resolve within `general.input_path` **and all descendant subfolders** (recursive), allowed extensions `.xlsx,.xlsm,.xlsb` (fixed set/scan order, not a precedence tie-breaker), exactly one candidate total across full scope or fail-fast (0/>1) with diagnostics including candidates when ambiguous. Non-recursive root-only search is non-compliant/regressive (`ea6409eb`) and must not be retained.
  - **U2:** (a) `data_files.filters` on the correct period field constrains axis sufficiently, (b) requires minimal axis/layout/pivot presentation adjustment, (c) config/template correction only. **No implementation approach selected without evidence.**
- **Required agents:**
  - **Auditor:** **Required** — supply evidence from B0, validate assumptions against artifacts, identify contradictions.
  - **Researcher:** **Required** — Excel/VBA PivotTable field/item/domain behavior, filter/axis/layout interactions, numeric/locale item handling, macro security/build automation support implications, and viable implementation patterns with trade-offs.
  - **Planner:** **Required (lead for design artifact)** — produce detailed semantic design notes/options, dependency-safe task slicing, and rollback/migration considerations.
  - **Validator:** **Required** — independent review of design against spec rev.7 (D2/D9/D10/D11/D12/D14), risk coverage, and absence of speculative scope.
- **Entry criteria:** B0 inventory/risk register complete (or explicitly bounded gaps recorded).
- **Exit criteria:** approved/recorded design decisions for implementable parts, explicit open decisions list for unresolved uncertainties with required evidence, and per-block implementation readiness criteria.

### B2 — D2 data-file resolution conformance (normative)
- **Objective:** align `data_files.data_file` resolution with updated normative spec D2 (base name resolved within `input_path` subtree recursively, allowed extensions `.xlsx,.xlsm,.xlsb` as fixed set without precedence tie-breaker, deterministic; fail fast on no match or >1 matching candidate with diagnostics) and eliminate the non-recursive regression/ambiguous behavior.
- **In scope:** semantic changes in file-resolution module/entry points; clear contextual errors; update any dependent assumptions; verification cases for precedence and collision/no-match.
- **Out of scope:** unrelated I/O refactors, silent wrong-file selection on ambiguity.
- **Impacted components (semantic):** `modTreportFiles` (resolution/search/selection), `modTreportRun` call sites/context, config usage of `input_path`.
- **Risks:** **High** (implementation must be deterministic and must not pick among ambiguous candidates; regression fix changes behavior vs `ea6409eb`). **Medium** (case/extension matching on Windows, traversal determinism).
- **Decision (2026-11-16):** normative D2 requires recursive subtree search. No non-recursive fallback; ensure compatibility with expected `monthly/weekly/base`-under-root layouts and fail-fast on ambiguity.
- **Required agents:**
  - **Auditor:** **Required** — confirm real folder/layout cases and regression targets.
  - **Researcher:** Optional — Windows/FS edge cases if relevant.
  - **Planner:** **Required** — precise semantic change spec and migration/compatibility notes.
  - **Validator:** **Required** if rollout impact is high; otherwise recommended.
- **Entry criteria:** B1 D2 decision recorded (recursive, 2026-11-16); B0 layout evidence available.
- **Exit criteria:** implementation-ready task spec + (later) code changes passing L1/L2 and D2 regression checks on copies (incl. subfolder resolution and ambiguity fail-fast); plan/status no longer claims pre-fix compliance.

### B3 — Rev.7 `max_<PivotTable field name>` placeholder contract (critical)
- **Objective:** implement spec-compliant placeholder detection/resolution/substitution for `report_sheet` only, bound to named PivotTable field maxima in relevant pivot/data context for the processed `data_files` row, **independent of `data_files.filters`**.
- **In scope:** semantic redesign of placeholder pipeline: token discovery in templates, per-token field lookup, numeric-maximum computation/fail-fast, deterministic resolution point per B1, substitution integration with existing Trim/sanitize/length/collision rules; remove/refactor legacy fixed-token handling and filter-representative/alias coupling as primary mechanism; preserve unrelated output behaviors.
- **Out of scope:** deriving placeholders from filters, speculative caching beyond correctness needs, `group_rows` changes.
- **Impacted components (semantic):** `modTreportOutput` (substitution), `modTreportRun` (orchestration/context passing, sequencing), `modTreportFilters`/`ResolvedContext` (remove legacy alias dependency), possibly `modTreportPivot` (field numeric-domain/max helpers) — final API shape per B1 design.
- **Risks:** **Critical/High** (core correctness and rev.7 compliance; template breakage until migrated; subtle Excel pivot/item type issues; sequencing errors could reintroduce filter coupling).
- **Uncertainties / alternatives:** resolution timing/source and field-name parsing remain open per B1 until decided; implementation patterns (central resolver vs module-local helpers) are alternatives — **detailed design in Planner output required before coding; do not preselect if uncertainty persists**.
- **Required agents:**
  - **Auditor:** **Required** — validate against real templates/field names and ensure all call paths covered.
  - **Researcher:** **Required** — PivotField/PivotItem numeric domain behavior, captions vs source names, locale/format issues, performance/safety trade-offs.
  - **Planner:** **Required (lead)** — detailed pre-implementation design (data flow, contracts, fail-fast conditions, migration hooks).
  - **Validator:** **Required** — independent spec-compliance and regression-risk review (high risk).
- **Entry criteria:** B1 placeholder decision(s) sufficiently resolved for safe implementation; B0 template inventory available.
- **Exit criteria:** code/design implements D11 contract, legacy coupling removed per decision, unit/behavioral checks defined, and B5 migration can be executed deterministically.

### B4 — D9 uniform rule confirmation and hardening (all filters)
- **Objective:** ensure implementation and diagnostics fully satisfy D9 for every applied filter: missing field in PivotTable available fields → fail fast; none matchable → fail fast; ≥1 matchable → matched subset applied (no warn-and-skip). Preserve D10 field-only `min`/`max`.
- **In scope:** gap analysis against exports, targeted fixes only if discrepancies/edge cases found (data-file and report filters, context-rich errors, consistent behavior), regression cases.
- **Out of scope:** behavior changes beyond D9/D10 compliance, broad refactor.
- **Impacted components (semantic):** `modTreportPivot` (lookup/selection/placement), `modTreportFilters` (apply paths), `modTreportRun` (error surfacing/context as needed).
- **Risks:** **Medium/High** (user-visible failure-mode changes vs prior behavior; Excel object-model edge cases). **Regression** risk for existing valid configs.
- **Uncertainties / alternatives:** matching equivalence details (trim, numeric/string forms) must align with spec and evidence; if ambiguity exists, **Researcher/Auditor evidence required before changing semantics**.
- **Required agents:**
  - **Auditor:** **Required** — confirm current behavior against D9 cases using code evidence and, later, runtime evidence.
  - **Researcher:** **Required** if PivotItems/value matching or placement edge cases are unclear.
  - **Planner:** Recommended — minimal change plan if fixes needed.
  - **Validator:** **Required** when behavior changes; recommended otherwise.
- **Entry criteria:** B0/B1 context; can run in parallel with B2 after gates if isolated.
- **Exit criteria:** explicit conformance statement per D9 condition with evidence (code + L3 where applicable); D10 preserved.

### B5 — Template/config migration for rev.7 placeholders (mandatory)
- **Objective:** update every in-scope `data_files.report_sheet` template (and placeholder metadata where required by process) in affected workbooks/configs to use correct `max_<PivotTable field name>` tokens matching actual pivot fields; remove legacy tokens.
- **In scope:** per-row/template checklist, before/after recording, alignment with B3 implementation and actual field names/captions; coordination across canonical/local workbook copies as applicable under working-copy discipline.
- **Out of scope:** changing business meaning of reports, editing unrelated config, committing gitignored data artifacts.
- **Impacted components (semantic):** workbook `settings`/`data_files` templates (and possibly `placeholders` descriptive content), verification artifacts.
- **Risks:** **High** (missed template causes runtime fail-fast; wrong field name token; drift between workbook copies). **Compatibility:** outputs/sheet names will change by design where tokens change.
- **Uncertainties / alternatives:** authoritative source of truth among multiple workbook copies (`data/theiss/treport.xlsm`, `mko_treport/src/xls/treport.xlsm`, others) — **must be established in B0/B1**. Migration method alternatives: manual Excel edit on working copies vs controlled scripted transformation if safe tooling exists — **not chosen here**.
- **Required agents:**
  - **Auditor:** **Required** — complete inventory and post-migration verification of every template.
  - **Researcher:** Optional — tooling/options for safe workbook table edits if automation considered.
  - **Planner:** **Required** — migration checklist, sequencing with B3/B6, rollback/backup plan.
  - **Validator:** **Required** — verify completeness and spec alignment before rollout.
- **Entry criteria:** B3 contract implemented/ready; B6 outcome known for affected templates (e.g. `report_wclip`) where token/field choice depends on evidence.
- **Exit criteria:** 100% in-scope templates migrated and recorded (before/after), no legacy placeholder tokens remain in scope, and L3 confirms expected names/outputs on copies.

### B6 — `report_wclip` correctness (U2, evidence-driven)
- **Objective:** ensure `report_wclip` period/axis/column output matches intended scope (resolve unintended 54–350 expansion) with minimal, evidence-backed changes.
- **In scope:** collect/interpret evidence on pivot layout, axis-driving fields, and effect of configured filters; determine required remedy; implement minimal config and/or code adjustment only if justified; update templates/expectations accordingly.
- **Out of scope:** speculative pivot redesign, broad report-model changes.
- **Impacted components (semantic):** `data_files` config/filters for `report_wclip`, possibly `modTreportPivot`/`modTreportFilters`/`modTreportRun` only if evidence proves axis/layout handling gap, output naming context.
- **Risks:** **High** (incorrect diagnosis; environment-specific pivot behavior). **Regression** for `report_w`/`report_m`.
- **Uncertainties / alternatives (open):** (a) filter field/value constrains axis as intended once semantics correct, (b) requires minimal placement/axis/layout handling in code, (c) config correction only. **No final approach selected without evidence from B0/Auditor/Researcher and B1 decision gate.**
- **Required agents:**
  - **Auditor:** **Required (lead for evidence)** — reproduce/characterize issue on copies, capture before/after behavior.
  - **Researcher:** **Required** — PivotTable axis/filter/page/row/column semantics and Excel behavior across relevant versions/configurations.
  - **Planner:** **Required** — minimal intervention options and decision criteria.
  - **Validator:** **Required** — ensure fix is spec-consistent and regression-bounded.
- **Entry criteria:** B0 baseline evidence; B1 hypotheses prepared.
- **Exit criteria:** evidenced root cause/resolution path selected and implemented (or config-fixed), with B7 demonstrating intended `report_wclip` output structure on copies.

### B7 — Build and runtime verification (L1/L2/L3) + regression/acceptance evidence
- **Objective:** produce objective evidence for compile/build and functional correctness on working copies per project rules and plan 003 acceptance checks.
- **In scope:** L1 `build_treport_vba.ps1` runs with recorded command/report/summary/VBOM notes; **L2 manual VBE Compile** on working copy (required for compile claim); L3 `Refresh_data`/`Create_report` on copies only; regression matrix covering D2, D9, D10, D11/migration/independence, D12 fail-fast, D14 naming/collisions, `report_wclip` correctness, existing rows (`report_m`,`report_w`), FR-RUN-1 workbook controls, FR-DEP repeatability/idempotence where exercised.
- **Out of scope:** running macros on production/original data, claiming success without evidence.
- **Impacted components (semantic):** all changed modules, workbooks/configs on working copies, build script/reports.
- **Risks:** **High** (environment/macro security/VBOM/PQ refresh timing/pivot cache behavior/locking can block or distort results; Excel 2021+/Windows differences). **Medium** (false confidence if copies not representative).
- **Uncertainties / alternatives:** verification depth per environment may vary; if L2/L3 blocked, record precise blocker and do not claim pass. Alternative evidence collection methods (manual logs/screenshots/report files) acceptable if traceable.
- **Required agents:**
  - **Auditor:** **Required** — execute/collect evidence, compare against acceptance checks, log discrepancies.
  - **Researcher:** **Required** when environment-specific Excel/PQ/pivot behavior affects interpretation or remediation.
  - **Planner:** Recommended — adjust sequencing/remediation tasks from failures.
  - **Validator:** **Required** — independent review of evidence sufficiency vs spec/plan gates before rollout.
- **Entry criteria:** relevant implementation/migration blocks complete on working copies.
- **Exit criteria:** traceable evidence package (commands, reports, L2 record, L3 outcomes, before/after templates, FR-RUN-1/FR-DEP notes) with pass/fail per check and explicit limitations.

### B8 — Controlled sync and rollout of canonical artifacts
- **Objective:** propagate verified changes to intended canonical/export/binary/workbook artifacts only after gates pass and approval is explicit.
- **In scope:** sync exports/bin/workbooks per approved process on working copies then controlled replacement; reconcile `vbaProject.bin`/xlsm parity; document what was updated and why; ensure no gitignored data committed.
- **Out of scope:** rollout without B7 evidence/approval, destructive edits without backups.
- **Impacted components (semantic):** `mko_treport/src/vbs/export/*`, `mko_treport/src/vbs/vbaProject.bin`, `mko_treport/src/xls/treport.xlsm`, target runtime workbook(s) as authorized.
- **Risks:** **High** (artifact drift reintroduced, accidental overwrite of canonical state, dual-target build script side effects).
- **Uncertainties / alternatives:** which workbook(s) are canonical for which consumers — must follow project/task approval; single-target build/sync alternatives exist in tooling and should be preferred when safer (**operational choice recorded, not preselected here**).
- **Required agents:**
  - **Auditor:** **Required** — parity/diff verification post-sync.
  - **Researcher:** Optional.
  - **Planner:** Recommended — rollout checklist and rollback.
  - **Validator:** **Required** for high-risk release/canonical updates.
- **Entry criteria:** B7 pass (or explicitly accepted residual risks) and explicit authorization in context.
- **Exit criteria:** artifacts consistent, backups retained, changes documented, no unintended files modified.

## 5. Cross-cutting risk register

| Risk | Type | Severity | Notes / mitigation |
|---|---|---|---|
| Rev.7 placeholder semantics/timing wrong or coupled to filters | Implementation/spec | Critical | B1 design gate + B3/B5 separation + Validator review + targeted L3 checks for D11 independence. |
| Workbook/config template inventory incomplete or inconsistent across copies | Rollout/compatibility | High | B0 exhaustive inventory, B5 checklist with before/after evidence, B7 regression on all in-scope rows. |
| D2 implementation error (missed subfolders, non-determinism, or wrong selection on ambiguity) | Functional/compatibility | High | Recursive deterministic traversal/collection, fail-fast on 0/>1 with candidate diagnostics, verification on root and subfolder layouts (e.g. `monthly/weekly/base`). |
| `report_wclip` axis/layout misunderstood | Functional/regression | High | Evidence-first B6, Researcher input, minimal change, compare against intended period/columns on copies. |
| Export ↔ vbaProject.bin ↔ workbook drift (`bin` modified/uncommitted) | Build/rollout | High | Treat all artifacts as potentially out of sync; work from exports + controlled rebuild on copies; B8 parity checks. |
| L2/L3 blocked by VBOM trust, macro security, Excel/session/COM issues | Environment/verification | High | Record exact blockers; use working copies; do not claim compile/runtime success without evidence. |
| Power Query refresh timing/background queries/stale data | Runtime/compatibility | High | Verify completion semantics in L3; Researcher guidance if issues observed; environment-specific notes. |
| Pivot/object-model behavior differs (field names/captions, item types, placement, cache) across workbooks/Excel builds | Compatibility | High | Semantic field lookup rules validated in B1/B3/B4; L3 on actual target-like data/copies. |
| Workbook/file locking, read-only state, save failures | Runtime/operational | Medium | Preserve existing access-mode handling; test in B7; clear diagnostics. |
| Output fidelity (values/formats/layout/sheet names/31-char/collisions) regressions | Regression | High | Full acceptance matrix in B7 including existing reports and naming rules. |
| Build script dual-target default mutates multiple workbooks | Operational | Medium | Prefer explicit single-target/working-copy parameters in execution; document commands in evidence. |

## 6. Uncertainty and alternatives register (no final choice mandated here)

| ID | Uncertainty | Alternatives (viable candidates) | Required to resolve |
|---|---|---|---|
| U-A | **D11 placeholder resolution point/source** for `max_<field>` (must be independent of `data_files.filters`, deterministic, in relevant pivot/data context) | (A1) compute from field numeric domain in pivot state at defined point in per-`data_files` processing; (A2) compute from cache/available values independent of current filter visibility; (A3) hybrid with explicit rules. | B0 evidence + B1 Researcher/Planner/Validator decision. |
| U-B | **Token grammar/field identity** for `<max_<PivotTable field name>>` including spaces and exact names | (B1) direct substring/token parse supporting known field-name set; (B2) stricter grammar with validation against pivot field names. | B0 inventory of actual names/tokens + B1 design. |
| U-C | **D2 resolution scope** (normative: recursive within `input_path` subtree, no precedence tie-breaker, 0/>1 fail-fast) | **Resolved (2026-11-16):** recursive subtree search required. | B2 implementation + B7 verification. |
| U-D | **`report_wclip` (U2)** cause of excess columns/axis expansion | (D1) filter semantics/config/template fix sufficient; (D2) minimal axis/layout handling change; (D3) combination. | B0/B6 evidence + Researcher analysis. |
| U-E | **Canonical workbook/config source** for migration and rollout | (E1) `data/theiss/treport.xlsm`; (E2) `mko_treport/src/xls/treport.xlsm`; (E3) synchronized set per approved process. | B0 audit + task-owner approval. |
| U-F | **Migration/apply method** for workbook table/template updates | (F1) manual controlled edits on working copies; (F2) auxiliary scripted transformation with verification. | B1/B5 planning and risk acceptance. |

## 7. Verification and acceptance gates (execution-level)

- **Gate G0 (after B0/B1):** baseline inventory complete, open uncertainties logged, high-risk design reviewed.
- **Gate G1 (per implementation block B2–B4,B6):** semantic spec satisfied, no scope creep, L1 on working copy run where code changed.
- **Gate G2 (compile):** **L2 manual VBE `Debug → Compile VBAProject` passed** on working copy (record path/date/time) before any compile-success claim.
- **Gate G3 (functional, copies only):** L3 `Refresh_data`/`Create_report` evidence covers plan 003 §4 checks relevant to changed scope (incl. all-template placeholder migration, D11 independence, D9 uniformity, D10, D2, no-data fail-fast, naming/collisions, `report_wclip`, regression for existing reports, FR-RUN-1).
- **Gate G4 (deployment readiness):** FR-DEP repeatability/idempotence evidenced if exercised; Validator sign-off for high-risk rollout.
- **Gate G5 (rollout):** controlled sync completed with backups/parity checks and explicit approval.

## 8. Deliverables expected from subsequent steps

- B0/B1: evidence inventory, risk/uncertainty log, and detailed semantic design/decision record(s).
- B2–B4,B6: implementation task specs and minimal, reviewable code changes (by implementors) with traceability to spec decisions.
- B5: migration checklist with before/after template strings per in-scope row/workbook.
- B7: verification report package (L1 command+report, L2 record, L3 results, regression matrix outcomes, FR-RUN-1/FR-DEP notes, limitations).
- B8: rollout record and artifact parity confirmation (if authorized).

## 9. Notes on agent involvement (summary)

| Block | Auditor | Researcher | Planner | Validator | Rationale |
|---|---:|---:|---:|---:|---|
| B0 | Required (lead) | Optional | Optional | Optional | Evidence/drift/inventory uncertainty is high; correctness of downstream work depends on baseline. |
| B1 | Required | Required | Required (lead) | Required | High technical/spec risk and multiple viable alternatives; must not preselect uncertain approach. |
| B2 | Required | Optional | Required | Required/Recommended | Normative compliance with rollout/compatibility impact. |
| B3 | Required | Required | Required (lead) | Required | Critical rev.7 contract, cross-module coupling, high regression risk. |
| B4 | Required | Required (if unclear) | Recommended | Required if behavior changes | Uniform fail-fast semantics and Excel matching/placement edge cases. |
| B5 | Required | Optional | Required | Required | Migration completeness and workbook drift risk. |
| B6 | Required (lead evidence) | Required | Required | Required | U2 unresolved, pivot axis behavior uncertain, regression-sensitive. |
| B7 | Required | Required (env-dependent) | Recommended | Required | Verification integrity, environment risks, gatekeeping for rollout. |
| B8 | Required | Optional | Recommended | Required | Canonical artifact consistency and high operational risk. |

## 10. B1 Decision and design record (2026-11-16)

- **Status:** B1 draft (evidence-based). Inputs: B0 inventory, `{context_b1_auditor}`, `{context_b1_researcher}`, `{context_u2_pivot}`. **No VBA/code or existing audit files changed.** Spec 001 rev.7 is normative.
- **Purpose:** record implementable, spec-compliant decisions for B2–B6 where clear; keep alternatives/open items explicit where runtime evidence is still required (esp. U2 fix selection, D2 rollout layout alignment, placeholder numeric-domain edge cases).

### 10.1 Rev.7 placeholder contract — `max_<PivotTable field name>` in `report_sheet` (D11, FR-FLT-3, FR-NAME-1) — **DECIDED (contract) / sequencing IMPLEMENTATION-DEFINED (defensible choice)**

**Normative requirements (must satisfy):**
- Target: `report_sheet` only. Token form `<max_<PivotTable field name>>` (e.g. `<max_Year>`, `<max_Week Number>`), supporting field names containing spaces as configured.
- Value: **maximum numeric value** of that named PivotTable field in the **relevant pivot/data context for the processed `data_files` row**, at the time/place resolution is performed (implementation-defined), from **available numeric-convertible items of that field in that context**.
- **Independent of `data_files.filters` clause derivation/representatives** (not derived from `data_files.filters`).
- Fail fast with clear context if referenced but not deterministically derivable (field missing/unusable in relevant pivot context, no numeric-convertible values, or max indeterminable). Preserve D14/FR-NAME rules (Trim/sanitize/31-char/collision after substitution). No zero-padding assumed.

**Token parsing (DECIDED):**
- **P-Scan (deterministic):** scan `report_sheet` template for every occurrence of prefix `<max_` and the next `>`; extracted inner text after `max_` is the configured PivotTable field name (preserve inner spaces, trim outer whitespace per configured-string handling only if consistent with template intent; prefer exact inner content match to configured/actual field name). Support multiple distinct tokens and repeated tokens per template. Empty field name or malformed/unclosed token → fail fast (unresolvable placeholder) with template/row context.
- No regex/external scripting dependency required. If a field name can contain `>` (not evidenced), treat as evidence gap; current inventory shows no such case.

**Field identity/lookup (DECIDED):**
- Resolve each named field via the PivotTable’s **available fields** for the pivot used for the processed `data_files` row (consistent with D7/D9 availability model): field must be findable in `pt.PivotFields` by the configured name (exact/trimmed as appropriate). If not found/unavailable/unusable for domain read → fail fast (D11/D12 context: data_file, data_sheet, report_sheet/template, field name). Use the same pivot instance/context as `ProcessDataFile` operates on for that row.

**Numeric domain and maximum (DECIDED, with bounded edge-case evidence note):**
- “Numeric-convertible item”: an item/value from the field’s available items in that context convertible to a numeric scalar using deterministic VBA-safe conversion consistent with project numeric handling (locale-stable as feasible; no reliance on filter-representative formatting). Non-convertible items (blank/text-only where not numeric-convertible) are excluded from max computation.
- “Maximum numeric value”: greatest numeric value among convertible items. If zero convertible items exist → fail fast (cannot derive deterministic maximum). Substituted text is the resolved numeric value rendered without forced zero-padding (unless later specified).

**Resolution source/context and sequencing (DECIDED defensible choice; implementation-defined allowed by spec):**
- **Chosen point:** resolve `max_<field>` values **once per `data_files` row**, after obtaining the target `data_sheet` and primary `PivotTable` for that row and **before applying any `data_files.filters` clauses** (and before per-report filter application). Rationale: (1) explicitly independent of `data_files.filters` clause derivation/representatives and independent of applied filter selection effects; (2) uses the pivot/data context of the processed `data_files` row (relevant row context) at a deterministic point; (3) aligns with pre-capture-of-field-domain concept already present for D10 while being decoupled from clause representatives.
- **Source:** compute max from the named field’s available item/domain values in that pivot context at that point (field-level domain). Implementation may enumerate `PivotField` items/available values via Excel object model; must not read/derive from `ResolvedContext` legacy aliases, clause representatives, or selected-subset results produced by filter application. Caching per (row, fieldName) within row processing is allowed for determinism/performance if it uses the same captured domain.
- **Independence statement:** satisfies “independent of `data_files.filters` clause derivation/representatives”. It is intentionally independent of filter application effects by resolving pre-filter; this is the selected defensible interpretation of “relevant pivot/data context…” for this implementation (recorded). Any future change to sequencing requires spec-consistent justification and evidence.
- **Substitution point:** continue to produce final `sheetName` per report output as needed, but substitution must use the pre-resolved placeholder values for that `data_files` row (not legacy `rc.MaxYear/Month/Week`). If no placeholder tokens present, resolution may be skipped.
- **Open evidence note (not blocking contract):** runtime confirmation on actual pivots (regular/grouped/date-like fields, item visibility/retention behavior) that pre-apply item/domain enumeration yields the intended “available numeric-convertible items…”; if Excel behavior differs materially, adjust source extraction minimally while preserving independence from filter derivation/representatives and determinism (escalate to Planner/Validator).

**Legacy removal (DECIDED):**
- Remove legacy fixed-token-only substitution path as primary behavior for rev.7 templates: eliminate dependence on `<max_year>/<max_monthnum>/<max_weeknum>` and on `MaxYear/MaxMonth/MaxWeek` derived from filter representatives/`DeriveLegacyPlaceholderAliases` in the `report_sheet` placeholder flow. `ResolvedContext`/filter plumbing may retain only what is needed for D9/D10 filter semantics (not for placeholders) or be refactored accordingly (semantic/API change expected in B3). No silent fallback to legacy semantics.

**Current inventory impact (evidence):** in-scope templates in both inspected workbooks: `report_m` = `помесячное сравнение` (no tokens), `report_w` = `сравн конкур_<max_Week Number>w_<max_Year>_прямая` (rev.7-like tokens; fields `Week Number`,`Year` present in caches), `report_wclip` = `хроно по неделям 2025` (no tokens). B3+B5 must align implementation and templates (row 2 already rev.7-like but code incompatible today).

### 10.2 D2 data-file resolution — **DECIDED (normative, updated 2026-11-16)**

**Spec (normative, D2/FR-FILE-1/2):** resolve `data_files.data_file` (base name) within `general.input_path` (including all descendant subfolders, recursive) to `<data_file>.<ext>` considering allowed extensions **`.xlsx`, `.xlsm`, `.xlsb`**. The extension list is a fixed allowed set/scan order and **not** a precedence tie-breaker. **Fail fast** if no match or if more than one matching candidate exists across the full searched scope and allowed extensions. Resolution must be deterministic; errors must list configured name, searched path/scope (including subfolders), cause, and candidate paths when ambiguous.

**Compliant algorithm (DECIDED):**
- **Search scope:** `general.input_path` subtree — `input_path` directory and all descendant subfolders (recursive). This is required to support structured layouts (e.g. `monthly/`, `weekly/`, `base/` under a shared root) while keeping `input_path` as the configured root.
- **Candidate matching (deterministic):** for each file encountered in a stable traversal order, match base name (case-insensitively on Windows, consistent with existing behavior) and allowed extensions `.xlsx,.xlsm,.xlsb`. Collect all matching candidates across the entire searched scope.
  - If **exactly one** candidate exists in total → select it.
  - If **zero** candidates exist → fail fast (no match).
  - If **more than one** candidate exists in total (including same/different subfolders and/or different allowed extensions) → fail fast as **ambiguous/>1 match**; do not select by extension precedence, folder proximity, or traversal order.
- **Errors:** include operation/context, configured `data_file`, resolved `input_path`/searched scope (including subfolders), and cause (not found / ambiguous with candidate paths).

**Regression/compatibility note:** commit `ea6409eb` changed resolution to non-recursive root-only search, breaking subfolder-based layouts; B2 must restore D2-compliant recursive deterministic search. No silent fallback selection is permitted.
### 10.3 D9/D10 alignment — **DECIDED (confirm + minimal hardening only if evidenced)**

**D9 (uniform, all filters — `data_files.filters`, `reports.articleLevel4Name_clean`, `reports.targetAudience`):**
1. Configured field name **not found** among PivotTable **available fields** → **fail fast** with clear context (data_file/sheet/pivot, filter source/report row if applicable, field name, expression/value).
2. Field exists but **none** of specified configured values matchable/present → **fail fast** with clear context (no warn-and-skip).
3. **≥1** specified value matchable → apply **matched subset** (unmatched tokens ignored), proceed. Preserve D7/D8 placement/layout/staging rules.
- Exports appear directionally aligned (`GetPivotFieldRequired`→2403, `ApplyFieldSelection`→2407 on zero wanted, subset apply; report filters share path). **B4 scope:** confirm no divergent code paths/edge cases (including non-filterable placement handling, page/row/column in-place behavior, token trim/numeric equivalence per spec) and improve diagnostics only where needed for D12-quality context. **No behavior change beyond compliance.**

**D10:** `min`/`max` resolve against available numeric values of **the referenced field only** in relevant pivot context for the row; **no cross-field/year-first dependency**. Exports pre-capture per-field numeric domains before apply (field-only) — appears aligned; **B4 must confirm** no hidden cross-field coupling remains and that placeholder resolution is fully separated (B3).

### 10.4 U2 `report_wclip` correctness — **EVIDENCE-BASED HYPOTHESES, FINAL FIX NOT SELECTED**

**Observed config (B0/B1):** `data_files` row 3: `data_file=report_wclip`, `data_sheet=clean_data_full`, `report_sheet=хроно по неделям 2025` (no placeholders), `filters=Week of Year=(202501-202553)`, `paste 1,1`, `freeze 2,1`, `zoom 0.7`.

**Pivot evidence (`{context_u2_pivot}`, `data/theiss/base/report_wclip.xlsx`, `PivotTable1` `A10:BC273`):**
- Column axis fields: `Week Number` (idx 6, axisCol, 53 items, hidden=0) and `Start of Week` (idx 7, axisCol, 149 date items spanning 2024-01-01..2026-11-03, hidden=0). Data layout/values include measures with `baseField` related to `Week of Year`.
- Page/filter fields include `Week of Year` (idx 34, axisPage, 152 items, hidden=98 in saved state = selective) and others (`Years`, `Year`, `researchDate`, etc.). `Week of Year` cache values include 202401..202645; `Week Number` 1..53.
- No explicit `showItemsWithNoData`/retain flags found in scanned attrs (extLst not parsed). Weekly/monthly pivots keep time fields off column axis (columns = Values only) — wclip layout differs.

**Root-cause hypotheses (ranked by evidence):**
1. **H1 (most likely):** `Week of Year` filter targets a **page-axis** field in-place, but **column proliferation is driven by different column-axis fields** (`Start of Week`, and possibly `Week Number`) that remain broad/unhidden across multi-year cache; filtering `Week of Year` alone does not constrain these axis-driving fields. `Start of Week`’s 149 items are consistent with large column expansion magnitude.
2. **H2 (contributing/possible):** retained/broad item membership on column-axis time fields (and multi-year cache) keeps members available; interaction with saved visibility/layout may amplify columns.
3. **H3 (possible):** configured value/range semantics for `Week of Year` match intended 2025 weeks, but without constraining axis drivers the presented range remains wide (needs runtime confirmation).

**Viable minimal remedies (NOT selected; evidence/decision in B6):**
- **R1 (config, preferred if sufficient):** adjust `data_files.filters` for `report_wclip` to constrain the **actual axis-driving period field(s)** (e.g. `Start of Week` and/or `Week Number` and/or keep/supplement `Week of Year` as appropriate) to the intended 2025 period, using correct field names/item value formats as present in pivot. Must preserve D7/D8 and be verified on copies.
- **R2 (minimal code + config, only if evidenced):** if correct filtering cannot be expressed/achieved without minimal axis/layout-aware handling while preserving layout intent and spec compliance, implement the smallest evidenced change in pivot/filter handling; requires Auditor/Researcher/Planner/Validator agreement.
- **R3 (source pivot/layout artifact adjustment, if authorized):** modify pivot layout in source workbook/template so constrained field drives/does not over-expand columns — outside VBA logic but affects artifacts; only if explicitly in scope/approved and regression-safe.

**B6 required evidence/decision criteria (gate):**
- Reproduce on working copy: capture baseline column count/headers/period span and filtered result for `Week of Year=(202501-202553)` and for candidate axis-driver filters.
- Confirm which field(s) truly drive displayed columns after filtering and whether item values (e.g. `Start of Week` dates vs week numbers vs `Week of Year` codes) match intended 2025 scope.
- Select **minimal** remedy that yields intended period/columns with no regression to `report_m`/`report_w`; record before/after evidence. If config-only insufficient, escalate before code/layout change.

### 10.5 Module/API impact map (semantic)

| Component | Planned semantic impact | Blocks |
|---|---|---|
| `modTreportFiles` | D2-compliant resolution (recursive `input_path` subtree search, fixed allowed extensions without precedence tie-breaker, deterministic 0/1/>1 fail-fast, contextual errors incl. candidates). | B2 |
| `modTreportOutput` | Replace legacy fixed placeholder substitution with generic rev.7 token parsing/substitution over pre-resolved named-field maxima; keep D14 sanitization/length/collision integration. | B3 |
| `modTreportRun` | Per-`data_files` placeholder-max resolution step at defined point (after pivot acquired, before data-file filter apply) producing row-scoped resolved values independent of filter reps; pass resolved values to substitution; adjust sequencing/call signatures as needed. | B3, (B6 if sequencing/context touched) |
| `modTreportFilters` / `ResolvedContext` | Remove legacy placeholder alias/representative coupling from placeholder flow; retain/refactor only D9/D10 filter parse/resolve/apply responsibilities as needed. | B3, B4 (verify no regression) |
| `modTreportPivot` | Possible helper(s) for field numeric-domain/max extraction from pivot field context (reusable/decoupled) and/or diagnostics hardening; preserve D7/D8/D9 selection behavior. | B3, B4 |
| `modTreportConfig` | No contract change expected except any clarified validation/error text; only if required. | as needed |
| `modTreportExtract`, `modTreportRefresh`, `modTreportState` | No functional change expected (regression watch). | B7 |
| `settings`/workbooks (`data_files.report_sheet`, possibly `placeholders` descriptions, `report_wclip` filters per U2) | Template token migration (B5) and evidence-driven U2 config correction (B6) on working copies with before/after records. | B5, B6 |
| `build_treport_vba.ps1` / docs | No change required by design; use single-target working-copy flows for L1. | B7 |

### 10.6 Dependency-safe task slicing (implementation-ready)

**B2 — D2 conformance (after B1):** Planner-refined task → Implementor (single). Semantic scope: `modTreportFiles.ResolveDataFilePath` (+ helpers) implement deterministic recursive search within `input_path` subtree, allowed extensions `.xlsx/.xlsm/.xlsb` (no precedence tie-breaker), unambiguous single-candidate resolution, and fail-fast no-match/>1-match with contextual errors (searched scope incl. subfolders, candidates when ambiguous). Acceptance: spec-aligned behavior restoring subfolder layouts (e.g. `monthly/weekly/base`); L1 on working copy; L2 claim only after manual compile; L3 on copies validates resolution and run behavior. Agents in execution: Auditor evidence + Validator recommended/required given regression/compatibility impact.

**B3 — Rev.7 placeholders (critical, after B1):** Planner detailed task → Implementor. Semantic scope: add rev.7 token discovery/parsing and per-row pre-filter max resolution for named fields in pivot context; refactor `modTreportOutput` substitution signature/logic to accept resolved rev.7 values; remove legacy alias/rep coupling in `modTreportRun`/`modTreportFilters`/`ResolvedContext` for placeholder path; fail-fast per D11/D12. Acceptance: row 2 template resolves correctly on copy (`Week Number`/`Year` maxima in row context, independent of filter reps), no legacy token dependence, existing naming/sanitization/collision rules preserved. High risk → Auditor/Researcher support as needed, Validator required.

**B4 — D9/D10 hardening (parallelizable with B2 if disjoint; after B1):** Auditor-led confirmation → Planner micro-task only if gaps found → Implementor minimal fix. Acceptance: all D9 conditions enforced for all filters with clear context; D10 field-only preserved; no unrelated behavior change.

**B6 — U2 `report_wclip` (after B0/B1 evidence, before finalizing affected migration):** Auditor reproduce/measure on working copy → Researcher/Planner choose minimal remedy based on evidence → Implementor apply config-only or minimal code fix. Acceptance: intended period/column scope achieved and evidenced; no regression `report_m/report_w`; remedy recorded.

**B5 — Template/config migration (after B3 contract ready and B6 outcome known where relevant):** Auditor inventory + Planner checklist → controlled workbook edits on working copies (auxiliary/manual, non-runtime) or Implementor only if code-side config artifacts exist in repo (not the case for gitignored xlsm data) → Validator completeness check. Checklist (per workbook copy in scope): before/after `report_sheet` per `data_files` row, confirm tokens `max_<PivotTable field name>` only where needed, field names match pivot fields, `placeholders` table descriptions aligned if kept. Current known: row 1/3 no tokens; row 2 already rev.7-like; `report_wclip` name static (revisit only if B6 requires). Record explicit before/after strings.

**B7 — Verification (after B2–B6 applicable):** Auditor executes/collects evidence on working copies. L1: run `build_treport_vba.ps1` with single-target/working-copy parameters and backups (avoid default dual-target unless explicitly intended), capture command + report path + summary/VBOM notes. L2: manual VBE `Debug → Compile VBAProject` on working copy (path/date/time/pass-fail) — required for compile claim. L3: `Refresh_data`/`Create_report` on copies only; regression matrix per plan 003 §4 (D2/D9/D10/D11 independence+migration, D12, D14, U2, existing reports, FR-RUN-1, FR-DEP if exercised). Validator independent review.

**B8 — Controlled sync/rollout (after B7 pass + explicit approval):** sync exports/bin/workbooks from verified working copy per approved single-target flow; parity check; no gitignored data/output commits. Auditor + Validator.

### 10.7 Rollback / migration / safety notes

- **Backups & working copies:** mandatory for all build/import/macro runs and workbook config edits; preserve originals; record artifact paths/hashes/timestamps where practical.
- **Drift control:** `mko_treport/src/vbs/vbaProject.bin` is currently modified/uncommitted and exports/workbooks may be out of sync — treat rebuilt verified working copy as source of truth for sync, not blind overwrite.
- **D2 fix/rollout:** restore recursive subtree search per updated D2; verify against existing structured layouts (`monthly/weekly/base`) and ensure ambiguity is fail-fast with diagnostics. Do not reintroduce non-recursive root-only behavior or silent selection.
- **Template migration:** complete per-row/per-workbook checklist with before/after values; validate no legacy placeholder tokens remain in scope and no unintended name collisions/length violations.
- **U2:** prefer config-only remedy if evidence supports; any code change must be minimal and regression-tested.
- **Verification claims:** never claim compile success without L2; never claim runtime success without L3 evidence on copies; report blockers explicitly (VBOM/macro security/COM/PQ timing/locking).

### 10.8 Open uncertainties (must remain explicit)

| ID | Uncertainty | Status / how resolved |
|---|---|---|
| U-A/B (B1 partial) | Pivot field item/domain edge cases for pre-apply max (grouped/date/retained/visibility semantics) and exact field-name matching nuances in target Excel | Contract/point decided; **confirm in B3/B7 L3**; adjust extraction minimally if needed without coupling to filter reps. |
| U-C | D2 scope/behavior for structured layouts | **Resolved (2026-11-16):** recursive subtree search required; B2/B7 must verify subfolder resolution and ambiguity fail-fast. |
| U-D (critical) | U2 final remedy (which axis-driving field(s)/values or minimal code/layout change yields intended 54–350 fix) | **B6 evidence + decision gate** on working copy; not preselected. |
| U-E | Canonical workbook/config source(s) to migrate/sync (`data/theiss/treport.xlsm` vs `mko_treport/src/xls/treport.xlsm`) | **B5/B8 with explicit approval**; currently package-identical in inspected state but runtime/formula context differs. |
| U-F | Migration method for workbook table edits (manual Excel vs controlled tooling) | **B5**; choose safest traceable method on working copies. |

### 10.9 B1 exit assessment (gate G0 readiness)

- **Decisions recorded (spec-compliant):** D11 placeholder contract/parsing/lookup/independence + defensible pre-apply sequencing and legacy decoupling; D2 compliant algorithm and explicit rollout implication; D9/D10 confirmation/hardening scope; B2–B8 slicing and risks.
- **Not preselected (requires evidence):** U2 final fix (R1/R2/R3 choice), placeholder edge-case extraction adjustments if runtime differs, canonical migration/sync targets and method.
- **Next:** proceed to block execution per dependencies with **Validator (B1)** review of this design record, then B2/B3/B4/B6 as gated (sequential implementor, working copies, L1/L2/L3 discipline).

### 10.10 B1 Validator required corrections (RC1–RC4) — incorporated 2026-11-16

**Status:** RC1–RC4 accepted and incorporated into the authoritative B1 design guidance before B2–B6 implementation. §10.1–§10.4/§10.6/§10.7 remain in force except where clarified below.

#### RC1 — D2 uniqueness vs allowed extensions (D2, FR-FILE-1/2) — **LOCKED (updated 2026-11-16)**
- **Search scope:** resolve within `general.input_path` **and all descendant subfolders** (recursive), per updated normative D2.
- For the same `data_files.data_file` base name, existence of **more than one matching candidate across the allowed extensions** (`.xlsx`, `.xlsm`, `.xlsb`) anywhere in the full searched scope is a fail-fast **">1 match/ambiguous"** condition in **all** cases, including cross-extension coexistence and/or same-name files in different subfolders.
- The fixed extension list `.xlsx, .xlsm, .xlsb` defines the allowed candidate set (and may define scan/check order for efficiency) but **must not** be used as a precedence tie-breaker to select among coexisting candidates.
- Deterministic outcomes: **exactly one** candidate in scope → select it; **zero** → fail fast (no match); **>1** → fail fast (ambiguous/>1 match). Error context includes configured `data_file`, resolved `input_path`/searched scope (including subfolders), and cause (not found / ambiguous; listing candidate paths when useful and safe).
#### RC2 — D11 “available numeric-convertible items in that context” excludes filter-applied selection state (D11, FR-FLT-3, FR-NAME-1) — **LOCKED**
- `max_<PivotTable field name>` is computed from the named field’s **available item/domain values** in the relevant pivot/data context for the processed `data_files` row, **independent of any `data_files.filters` application/selection effects** and **not limited to items currently visible or selected due to filters** (and not derived from filter clause derivation/representatives or selected-subset results).
- The chosen pre-apply capture/resolution point (per §10.1: once per `data_files` row, after target `data_sheet`/primary `PivotTable` acquired and **before applying any `data_files.filters` clauses** and before per-report filters) must be implemented and documented accordingly.
- **B3 requirement:** specify the concrete Excel object-model enumeration source/approach used for the field item/domain (and how numeric-convertible values are collected deterministically) in implementation notes/task evidence.
- **B7 L3 requirement:** evidence that resolved placeholder values are independent of filter derivation/representatives and consistent with the intended row/pivot context; validate edge cases as applicable (grouped/date-like/retained/hidden item behaviors) and record any necessary minimal extraction adjustment that preserves independence and determinism (escalate to Planner/Validator if semantics diverge).

#### RC3 — Placeholder token handling scope and non-rev.7 `<...>` behavior (D11, D12, FR-NAME-1/2) — **LOCKED**
- **Recognized rev.7 placeholders:** only well-formed `<max_<PivotTable field name>>` tokens in `report_sheet` are resolved/substituted per D11.
- **Malformed/unclosed max tokens:** any malformed/unclosed `<max_...` construct (including empty field name) → **fail fast** with clear context (data_file row/sheet, template, token/field as applicable).
- **Legacy and other placeholder-like tokens:** any legacy fixed placeholder tokens (`<max_weeknum>`, `<max_monthnum>`, `<max_year>`) and any other unresolved placeholder-like tokens remaining in in-scope `report_sheet` templates must **fail fast** as unresolved/legacy-noncompliant placeholders under D11/D12 (not silent pass-through). This applies at substitution/validation time for templates in scope.
- **B5 requirement:** exhaustive per-workbook/per-`data_files`-row verification that no legacy placeholder tokens and no unintended/unresolved placeholder-like tokens remain in in-scope `report_sheet` templates; record before/after strings and token inventory.

#### RC4 — D2 operational readiness with recursive subtree search (D2, FR-FILE-1/2) — **LOCKED (updated 2026-11-16)**
- **B2** delivers **code/spec compliance** for D2 resolution algorithm and errors (recursive `input_path` subtree search, no precedence tie-breaker, deterministic 0/1/>1 outcomes, diagnostics).
- **L3/functional acceptance** on working copies must verify resolution against intended structured layouts (e.g. `data_file` sources located under `monthly/`, `weekly/`, `base/` beneath `input_path`) and confirm fail-fast ambiguity behavior where applicable.
- Readiness claims must distinguish **"D2 algorithm compliant (recursive)"** from broader runtime success; record evidence (searched scope, resolved path/candidates/errors) and any remaining blockers.
- No silent fallback selection or non-deterministic tie-breaking may be introduced.
#### Correction tracking
- **Incorporated into (updated 2026-11-16):** §10.1 (D11 sequencing/source/independence and token handling clarifications), §10.2 (D2 recursive subtree scope, uniqueness/no tie-breaker, diagnostics), §10.5–§10.7 (B2/module/rollout), §10.8 (uncertainties), RC1/RC4 (scope and readiness).
- **Residual risks/uncertainties:** updated to reflect recursive D2 (RC1/RC4 revised 2026-11-16) (see §10.8 and B1 Validator risks: item-domain edge cases, U2 remedy, canonical source/method, environment blockers for L2/L3).

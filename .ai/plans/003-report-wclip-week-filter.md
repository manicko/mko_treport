# Execution plan: 003 — report_wclip week filter / generalized `data_files.filters` model

- **Plan ID:** 003-report-wclip-week-filter
- **Status:** B1 (R-1 structural + live pivot-field check) completed; B2 PO decisions **recorded and confirmed** (2026-11-16); B3 R-3/R-4 options and draft decomposition prepared but **provisionally frozen for implementation start** (2026-11-16) with explicit technical contingencies (U1-U10, F1-F5 from Validator) to be verified/resolved during T6/T7/L3; **implementation in progress** (per PO direction to proceed). No VBA/code changed by planning steps.
- **Date:** 2026-11-16 (updated 2026-11-16 with PO decisions)
- **Source spec:** `.ai/problems/002_report_wclip_week_filter_spec.md` (Draft analytical specification, rev. 2, 2026-11-16, Spec ID `002_report_wclip_week_filter`)
- **Inputs:** Spec §1–§9, Appendix B; Auditor `{code_context}` (current implementation vs spec, read-only scan of `mko_treport/src/vbs/export/{modTreportConfig,modTreportExtract,modTreportFiles,modTreportFilters,modTreportOutput,modTreportPivot,modTreportRefresh,modTreportRun,modTreportState}.bas`, `{Sheet1,Sheet2,ThisWorkbook}.cls`; `.ai/problems/P02.md`, `001_treport_spec.md`)
- **Task authoring template (for future implementation `{task_description}` artifacts):** `.ai/tasks/templates/task_template.yaml`
- **Runtime/verification constraints (project):** VBA export sources in `mko_treport/src/vbs/export/*`; runtime remains **VBA-only in Excel**; build/sync via `mko_treport/src/scripts/build_treport_vba.ps1`; L1 automated script report, **L2 manual VBE `Debug → Compile VBAProject`** required to claim compile, **L3** functional checks on copies when logic/run behavior changes (`docs/02-build-and-deployment/vba-build.md`, repo command rules). Backups/working copies only for workbook verification.

## 1. Objective and readiness context

### Objective
Decompose spec 002 into dependency-safe, evidence-gated execution blocks that can restore correct `report_wclip` output (symptom: extra columns 54–350 in `output/DOLGIT W30-50.xlsx`, sheet `хроно по неделям 2025`) by adopting a **general, field-driven filter model** for `data_files.filters` (CR-G1..CR-G5, §4), without hardcoded week/month/year semantics in the filter engine.

### Current readiness state (updated)
- **B1/R-1 (structural, completed on working copies `.ai/tmp/r1_20261116_165838/`, originals untouched):** `treport.xlsm` `data_files` row `report_wclip` = `data_sheet=clean_data_full`, `report_sheet=Хроно по неделям 2025`, **`filters=Week of Year=(202501-202553)`**, `freeze 2,1, zoom 0.7`; `reports` includes `DOLGIT W30-50` / `W 30-50`. `base/report_wclip.xlsx` pivot/cache contains **`Week of Year` (numeric ~202401–202645, covers 202501–202553)** plus `Year/Month Number/Week Number/Start of Week`; column axis includes week/time fields (e.g. `Week Number`/`Start of Week`). `output/DOLGIT W30-50.xlsx` sheet `Хроно по неделям 2025` is static **A1:MM88 (351 cols)** with row 2 showing `Start of Week`-style serials across ~cols 2–350 (irregularity around 53→54) — consistent with over-expanded weekly/period members beyond intended 202501–202553. **Live Excel UI/runtime confirmation still pending** (field caption/name resolution as VBA sees it, actual placement at execution, definitive causal attribution: mis-application vs placement/axis/cache interaction).
- **B1/R-2 (completed, repo trace):** `Create_report` → `ProcessDataFile` parses/applies `data_files.filters` via hardcoded `ResolveMinMaxAndYear` (`Year/Month Number/Week Number`, Year-first, fixed `ResolvedContext.MaxYear/MaxMonth/MaxWeek`) → `ApplyDataFileFilters` → per-report `ApplyReportFilters` (hardcoded `articleLevel4Name_clean` skip-if-none 2411, `targetAudience` strict) → `SubstitutePlaceholders` only `<max_year,max_monthnum,max_weeknum>` → snapshot full presented range. Minimal refactor mapping identified (remove hardcoded period/Max*/dependency coupling, generalize per-field resolution/placeholders/call sites) with compatibility risks for `report_m`/`report_w`.
- **B2 (PO decisions recorded 2026-11-16):** see §2a. **Remaining before implementation (DoR):** R-1 live/runtime confirmation recommended where feasible (field caption/name resolution as VBA sees it, placement at execution, definitive causal attribution for 54–350), B3 design/options (R-3/R-4) with recommendation and frozen approach after uncertainties resolved, acceptance/regression criteria locked (incl. all `data_files` rows per Q2-G9), and L1/L2/L3 verification plan. No final implementation approach chosen yet.

### How to proceed when local artifacts are unavailable in-repo
- Treat R-1 as an **external evidence gate**: produce an evidence-collection checklist/runbook (paths, sheets, `data_files`/`reports` rows, pivot `PivotFields`/items/formats, column 54–350 axis mapping, comparison with `report_m`/`report_w`) to be executed on a machine with access to `data/theiss/*` on **working copies**; attach screenshots/tables/exported notes as evidence artifacts (store paths decided by orchestrator, e.g. under `.ai/tmp` or an agreed evidence location) without committing sensitive workbooks.
- Partition work: advance **repo-evidence and design-prep blocks** (R-2 code trace/refactor map, R-3 option framing constrained by spec, R-4 verification/runbook drafting, PO-question capture templates) in parallel where safe; **do not select/lock implementation behavior** that depends on unresolved §6.2 or R-1 facts.
- If access cannot be obtained, mark plan state **evidence-blocked**, record explicit unknowns (symptom causality, real field model for Q2-G8), and limit outputs to decision-ready options and risk-contained stubs — no speculative “fix” coding.

## 2a. PO decisions recorded (2026-11-16) — B2 progress

| ID | Decision (as provided by PO) | Notes / implications |
|---|---|---|
| **Q2-G8** | **Confirmed (per spec §4).** No hardcoded field semantics. For `report_wclip` (`DOLGIT W30-50`, sheet `Хроно по неделям 2025`, `data_file=report_wclip`, `data_sheet=clean_data_full`), intended `data_files.filters = Week of Year = (202501-202553)` → select **all available values** in interval 202501–202553 for field `Week of Year`. Core principle §4.1 holds: field must exist in PivotTable available fields or fail fast; user-controlled `<field>=<expr>`; clauses joined by `|`; `general.delimeter` splits multi-value lists inside expressions. | R-1 confirms `Week of Year` field and 202501–202553 items exist in `base/report_wclip.xlsx` cache and output currently expands to 351 cols with extra weekly/`Start of Week` members. **Acceptance:** intended 2025 week period columns only (53 periods implied by 202501–202553) + intended layout/totals, **no extra columns 54–350**. Runtime field-name/caption resolution and causal fix path require implementation-time verification (L3 on copies). |
| **Q2-G9** | **Filter logic must be fully reworked and applied equally to all reports.** | Broad-scope refactor (not limited to `report_wclip`). Regression risk for `report_m`/`report_w` and all `data_files`/`reports` rows; requires compatibility/migration analysis and full regression plan in B3/B4/B6. Uniform filter rules (Q2-G4/Q2-G5/Q2-G3/Q2-G2/Q2-G6/Q2-G7) apply across the reworked filter logic. |
| **Q2-G4 / Q2-G5 (uniform rule for all filters, incl. report-level)** | **For all filters** (field name vs value distinction is explicit): **(1)** if the field name/caption does not exist in the PivotTable’s available fields → **fail with a clear error**; **(2)** if for the specified field **none** of the specified values are present/available/matchable → **fail with a clear error**; **(3)** if for the specified field **at least one** of the specified values is present → **filter by the present/matching values and complete without error** (apply matched subset). No hardcoded field semantics. **Applies equally to `data_files.filters` and report-level filters.** (Re-confirmed by PO 2026-11-16.) | **Behavior change vs current implementation:** today `articleLevel4Name_clean` report filter uses “matched subset if any; if none match → warn and skip that report (2411), continue others” and `targetAudience`/data-file filters reject missing tokens. Under this decision, “none match” becomes an **error** (no skip-and-continue for that case), and “some match, some missing” becomes **non-error subset filtering** for all filters. This affects `data_files.filters` and report-level filters and must be reflected in B3/B4 design, error handling, and regression/acceptance criteria (including whether partial subset is the intended business outcome where configs list extra values). |
| **Q2-G1** | **`max` = maximum among all available numeric values** of that field. | Derive from all available numeric items of the referenced pivot field (independent/global for that field in current pivot context), not hardcoded Year→Month/Week dependency or special field names. Exact resolution timing/interaction with other clauses to be specified in B3 consistent with “no hardcode / apply equally”. |
| **Q2-G2** | **Numeric literals, and allow literals `min` and `max` as start and end bounds** (examples: `1-max`, and combinations such as `(min-max)`, `(min-50)`, `(10-max)` are allowed as bound forms). After resolving `min`/`max`, if `start > end` → **fail with a clear error**. | Range bounds are numeric/min/max tokens per clause parsing rules; resolved interval must be valid. |
| **Q2-G3** | **Range specified for non-numeric values → fail with a clear error.** Non-numeric values must be enumerated via `general.delimeter` (list form). **Numeric vs non-numeric definition:** anything that can be converted to a number is treated as numeric (e.g. `Week of Year = (202501-202553)` uses numeric values/range). | Range form `(start-end)` requires numeric-comparable values/bounds (convertible to numbers) per this rule; otherwise use list syntax. Exact conversion rules/culture-invariance to be specified in B3 (VBA `IsNumeric`/`CDbl` semantics, locale safety) without changing the stated principle. |
| **Q2-G6** | **Keep its placement and apply the filter there; if it’s available but not placed, move it to Filters and apply** (current flexible behavior), failing clearly only if impossible. | Preserve flexible, safe placement for generic user-chosen fields. |
| **Q2-G7** | **Same field repeated in `filters`: not allowed → fail with a clear error.** | Consistent with duplicate-clause rejection intent. |

## 2. Decomposition into logical execution blocks

| Block | Goal | In scope | Out of scope | Dependencies | Execution order / parallelism | Required agents |
|---|---|---|---|---|---|---|
| **B1 — Evidence & research (Excel + VBA trace)** | Establish root cause and refactor surface with evidence (R-1, R-2). | R-1 checklist/results on working copies (pivot fields/items, config rows, 54–350 mapping, compare `report_m`/`report_w`); R-2 trace `Create_report` path and catalog hardcoded assumptions/special cases and minimal refactor touch-points (`modTreportRun`, `modTreportFilters`, `modTreportPivot`, `modTreportOutput`, `modTreportConfig` as implicated). | Code/workbook edits; deciding PO policy. | Spec + code_context; B1-R1 needs local `data/theiss/*` access. | **Start early.** B1-R2 (repo/code) can run **in parallel** with B1-R1 prep; B1-R1 execution depends on environment access. | **Auditor** (deeper investigation/evidence), **Researcher** (R-1/R-2 methods, VBA/Pivot behavior facts). |
| **B2 — PO decision capture & scope lock** | Convert §6.2 ambiguities into explicit, recorded decisions and acceptance criteria (CT-5). | Capture answers to **Q2-G1..Q2-G9**; confirm intended `data_files.filters` string for `report_wclip` post-R-1 (Q2-G8); confirm fix scope/regression breadth (Q2-G9); resolve GA-G1..GA-G9 as applicable. | Implementing behavior before decisions; changing report-level semantics without decision (Q2-G5). | B1 evidence (esp. R-1 for Q2-G8) informs but some policy can be pre-discussed; **final lock requires R-1 where factual**. | After/alongside B1; **gate before B4/B5 design freeze & implementation**. | **Planner** (structure options/questions), **Researcher** (implications), **Validator** (independent check that decisions are complete/unambiguous) — optional but recommended given risk. |
| **B3 — Design options & target behavior specification** | Produce implementation-ready design choices aligned to CR-G1..CR-G4 and decided policies (R-3; pre-DoR drafting allowed, finalization gated). | Top 2–3 approaches for generic parsing/`max`/ranges/lists/matching/placement/error handling; grammar/disambiguation; `ResolvedContext`/placeholder evolution; interaction with report-level filters (per Q2-G5) and `reports`/config schema impacts; backward-compat/migration for `report_m`/`report_w`; recommendation + rationale; R-4 build/verify plan outline if VBA changes expected. **When uncertainty remains, enumerate paths and criteria — do not choose final approach.** | Committing to a single refactor strategy without R-1/PO resolution where required; runtime code changes. | B1 findings; B2 decisions for affected areas. | Iterative: draft in parallel with late B1/B2; **freeze only after B2 gate**. | **Researcher** (options/best practices), **Planner** (detailed pre-implementation design), **Validator** (independent plan/design review). |
| **B4 — Implementation decomposition (VBA/config semantics)** | Define atomic, dependency-safe change sets/tasks for the chosen design (post-gate). | Semantic task breakdown (engine generalization, removal of hardcoded `Year/Month Number/Week Number`/`Max*` coupling, placeholder strategy per decision, config/text migrations if any, report-level handling per Q2-G5/Q2-G6, `report_wclip` filter correction per Q2-G8/Q2-G9). Each future task to follow `.ai/tasks/templates/task_template.yaml` (id/title/priority/depends_on/source_reference/goals/files with semantic targets/anchors/changes/acceptance_criteria as applicable). | Actual VBA edits in this planning step; broad redesign beyond spec; non-Excel runtime dependencies. | **Hard gate:** B1 complete, B2 recorded, B3 frozen, DoR met. | Sequenced by dependency DAG (shared config/parsing primitives before call-site behavior; high-risk pivots isolated). | **Planner** (task specs), **Auditor** (targeted pre-change checks if scope shifts), **Validator** (plan/task review) — **all required for high-risk implementation blocks** (shared filter engine, pivot apply semantics, public/config contract). |
| **B5 — Documentation & knowledge transfer** | Capture only necessary durable knowledge (proportional). | Update/specify: filter grammar/examples/error semantics, migration notes for existing `data_files.filters` (if changed), PO decisions record, verification/runbook notes; reference build docs. Only where ambiguity/operational risk warrants. | Documentation noise for self-evident internal refactors. | B3/B4 outcomes. | After design freeze; finalize after implementation verification evidence. | **Planner** (scope of docs), **Validator** (accuracy vs decided behavior). |
| **B6 — Validation, build & rollout verification** | Prove correctness and containment on working copies (CT-4 execution when implementing). | L1 `build_treport_vba.ps1` report review (VBOM caveats noted); **L2 manual VBE Compile** on working copy; **L3** functional checks on copies for `Create_report`/`Refresh_data` as relevant (no corruption of originals; confirm source files unchanged per rules); regression smoke for `report_m`/`report_w` and other `data_files` rows per Q2-G9; compare outputs to expected column behavior for `report_wclip`. | Running macros on production/original data; claiming compile without L2 when asked for verification. | B4 implementation tasks done in a working copy context. | Final gate before any canonical sync consideration (sync not in scope of this plan). | **Auditor** (evidence of outcomes/regressions), **Validator** (independent acceptance against Q2-G8/criteria). |

## 3. Dependencies and execution order (summary DAG)

1. **B1-R2 (repo trace/refactor map)** — can start immediately (no local workbook access required for static analysis).
2. **B1-R1 (Excel evidence)** — starts when access to `data/theiss/*` working copies is available; **parallel** to B1-R2/B3-draft/B4-prep templates.
3. **B2 (PO decisions Q2-G1..Q2-G9)** — proceeds with known info; **Q2-G8 finalized after R-1**; B2 is a **hard gate** for B3 freeze and B4.
4. **B3 (design/options)** — draft early, **no final selection under technical/PO uncertainty**; freeze after B1+B2.
5. **B4 (implementation decomposition & execution)** — only after DoR; internal order: shared parsing/resolution primitives and contract decisions → pivot/filter application semantics → orchestration/placeholders/output touch-points → `report_wclip` config alignment per Q2-G8 → regression-safe cleanup.
6. **B5/B6** — verification planning in B3/R-4, execution in B6 post-implementation; docs updated with verified behavior.

**Critical path:** B1-R1 + B2(Q2-G8/Q2-G1) → B3 freeze → B4 → B6 (L2/L3 evidence).

## 4. Risk assessment

| Area | Risk | Severity | Notes / containment |
|---|---|---|---|
| Implementation | Generalizing filter engine touches shared parsing, context resolution, placeholder substitution, and report orchestration (`modTreportFilters`, `modTreportOutput`, `modTreportRun`, config expectations). | **HIGH** | Isolate primitives; semantic, minimal refactor per R-2; avoid broad rewrites; preserve compatibility where configs unchanged unless migrated. |
| Implementation | Removing hardcoded `Year/Month Number/Week Number`/`MaxYear|MaxMonth|MaxWeek` assumptions changes `ResolvedContext` and placeholder model. | **HIGH** | Explicit design for `max` (Q2-G1..Q2-G3) and placeholders; fail-fast errors with context; migration plan for affected configs. |
| Rollout / compatibility | Existing `report_m`/`report_w` and other `data_files` rows may rely on current semantics or special cases. | **HIGH/MED** | Q2-G9 scope control; regression smoke on copies for all impacted rows; backward-compat analysis in R-2/R-3. |
| Regression | Symptom cause (extra columns 54–350) may be pivot layout/field placement/cache state, not only filter text; fix could mask or shift issues. | **MED** | R-1 must map columns to axes/fields and confirm intended period constraint; validate expected layout, not just absence of extras. |
| Compatibility (Excel/VBA) | PivotTable/object-model behavior varies by Excel version/build, locale, field caption vs name, page-field placement APIs, and `PivotFields` availability semantics. | **MED** | Evidence-based field lookup rules (GA-G8); test on target Excel environment(s) during L3; avoid version-specific assumptions without guards. |
| Environment (Windows/Office) | Macro security/VBOM trust blocks L1 compile note and automation; policies differ by machine; “cannot compile” ≠ syntax error. | **MED** | Follow L1/L2 workflow; record VBOM issues; never claim compile from script alone. |
| Power Query / refresh | Async refresh timing, connection state, or stale caches can affect pivot items/max resolution and outputs. | **MED** | Honor existing refresh waits; run L3 with controlled refresh on copies; distinguish filter defects from data-refresh issues (GA-G9). |
| Workbook state/locking | Concurrent opens, protected structure, or locked workbooks can break automation/verification. | **LOW/MED** | Use dedicated working copies; document environment preconditions in R-4 runbook. |
| Output fidelity | Formatting, widths/heights, sheet naming/collision handling, and placeholder substitution must remain correct. | **MED** | Compare full presented ranges on copies; include fidelity checks in acceptance criteria. |
| Evidence gap | `data/theiss/*` gitignored/unavailable in-repo; root cause and Q2-G8 cannot be finalized here. | **HIGH (to readiness)** | B1-R1 external gate + checklist; block implementation freeze until evidence captured. |

## 5. Implementation paths / alternatives (do not finalize under uncertainty)

> **Uncertainty present (R-1 facts, Q2-G1..Q2-G9).** The following are viable paths to evaluate in **B3 (R-3)**; **no final implementation approach is chosen in this plan.**

| Option | Summary | Pros (maintainability/evolution/conventions) | Cons / risks | When preferable |
|---|---|---|---|---|
| **A — General filter-engine refactor (spec-aligned)** | Replace hardcoded period/`Max*` logic with a generic clause parser/resolver over arbitrary pivot fields; explicit `max`/range/list semantics per decisions; evolve `ResolvedContext`/placeholders generically or via explicit mapping. | Matches CR-G1..CR-G4; removes special-case debt; extensible for future fields; aligns with “no hardcoding” principle. | Highest change surface; compatibility/migration burden; requires strong regression plan. | Preferred direction **if** Q2-G9 allows broader alignment and R-2 shows feasible containment; final choice deferred to B3 freeze. |
| **B — Targeted fix for `report_wclip` + constrained generalization** | Correct the specific misconfiguration/handling causing 54–350 columns for the confirmed field model (Q2-G8) while limiting engine changes to the minimum needed now; defer full `Max*`/placeholder generalization. | Lower immediate blast radius; faster remediation for reported defect. | Risks preserving hardcoded assumptions (conflicts with CR-G3 long-term); may not satisfy “general model” target; technical debt. | Only if PO constrains scope tightly (Q2-G9 option 1) **and** evidence shows a safe minimal path — must still respect CR-G2/explicit failures; decision deferred. |
| **C — Hybrid: generic parsing/validation core + phased semantics rollout** | Introduce generic field-existence/parse/validate pipeline now (fail-fast, no hardcoded names), but phase `max`/context/placeholder generality behind clear rules or staged config migration. | Balances safety and evolution; enables incremental verification (L2/L3 per phase). | Requires careful sequencing and clear interim behavior; more planning overhead. | Useful if R-3 identifies dependency/compatibility tensions (e.g. placeholder contract) — evaluate in B3. |

**Evaluation criteria for B3:** conformance to §4/CR-G1..CR-G4 and PO decisions, compatibility with `report_m`/`report_w`, clarity of error handling, testability on working copies, minimal necessary coupling, and alignment with project conventions (modular VBA exports, semantic anchors, no speculative abstractions).

## 6. Scope control (minimal, non-speculative)

- **In plan scope:** evidence gathering, decision capture, design, implementation decomposition, proportional docs, and verification strategy for spec 002.
- **Out of plan scope:** any VBA/source/workbook modification now; redesign beyond the general filter model and defect fix; inventing Excel facts or PO answers; runtime dependencies outside VBA/Excel.
- **Change discipline:** prefer isolated, atomic tasks with explicit `blocked_by`/`depends_on` semantics in future task files; do not split work unless it improves isolation, risk containment, parallelism, or reviewability.

## 7. Block-level agent requirements (consolidated)

- **B1:** **Auditor** (required, deeper investigation), **Researcher** (required for R-1/R-2).
- **B2:** **Planner** (required, question/decision structuring), **Researcher** (recommended), **Validator** (recommended for completeness check).
- **B3:** **Researcher** (required, R-3/R-4 options), **Planner** (required, detailed pre-implementation design), **Validator** (required, independent review before freeze).
- **B4 (high-risk implementation blocks):** **Planner** (required, task specs per `.ai/tasks/templates/task_template.yaml`), **Auditor** (required for targeted checks when touching shared engine/config), **Validator** (required for plan/task acceptance) — i.e. **all three** engaged for high-risk blocks.
- **B5:** **Planner** (required), **Validator** (recommended).
- **B6:** **Auditor** (required, evidence), **Validator** (required, acceptance/regression sign-off).

## 8. Definition of Ready for implementation (aligned to spec §9.6)

Implementation may start only when:
- PO answers **Q2-G1..Q2-G9** are recorded (including **Q2-G8** exact target `data_files.filters` and **Q2-G9** scope).
- **R-1** and **R-2** evidence complete and root cause confirmed on working copies/repo trace.
- **R-3** (and **R-4** if VBA changes proposed) approved; **final approach selected only after uncertainties resolved** (B3 freeze).
- Acceptance criteria defined for agreed scope (including expected `report_wclip` column/layout behavior) and regression plan for other `data_files` rows included.
- Verification workflow planned/executable: backups/working copies, `build_treport_vba.ps1` (L1), **L2 manual VBE Compile**, **L3** on copies if logic/run behavior changes.

## 9. Artifacts and traceability

- **This plan:** `.ai/plans/003-report-wclip-week-filter.md` (updated 2026-11-16: B1 evidence, B2 PO decisions, B3 draft+freeze defaults, B4 tasking, implementation tracking)
- **Spec:** `.ai/problems/002_report_wclip_week_filter_spec.md`
- **Task template:** `.ai/tasks/templates/task_template.yaml`
- **R-1 working copies/evidence:** `.ai/tmp/r1_20261116_165838/` (read-only inspection artifacts; originals in `data/theiss/` untouched)
- **Implementation:** T1–T7 executed sequentially via Implementor agents (exports in sync, no gitignored `data/theiss/*` committed); L1/L2/L3 evidence recorded in plan/build reports as produced.

### 11.6 Provisional implementation freeze (2026-11-16, PO-directed)
PO directed move to implementation; business decisions Q2-G1..Q2-G9 are complete. B3 draft (§11.1–11.5, Option B-like direction) is **provisionally adopted** for implementation with the following explicit choices/contingencies (to be validated in L3/T6/T7 and adjusted only with evidence):

- **F1/U3/U4 (`max`/`min`):** per Q2-G1, resolve `min`/`max` from **all available numeric values of the referenced field** in the pivot (field-scoped, independent of other filter clauses; no hardcoded Year→Month/Week dependency or Year-first apply). Clause application remains deterministic AND in declared order in pivot state; `max`/`min` resolution for a field reads that field’s available numeric items independent of other clauses’ selections (global-per-field). *Compatibility note:* may differ from legacy Year-context maxima for `report_m/w` (`Year=max|Week Number=(1-max)` / `Month Number=(1-max)`); L3 must compare outputs/sheet names and adjust config order/values only if evidence requires (no engine hardcoding).
- **F3/U5 (placeholders):** legacy tokens `<max_year>`, `<max_monthnum>`, `<max_weeknum>` resolved via generic per-field results by field-name alias (`Year`, `Month Number`, `Week Number`) without filter-semantics hardcoding. Representative value: for `max`/range end `max` use resolved max; for `min`/range start `min` use resolved min; for scalar numeric use scalar; for numeric range use resolved end bound as representative; for list use numeric maximum among present/matched numeric values in that clause (deterministic) — finalize exact rule in implementation and verify against `report_m/w` sheet names in L3.
- **U2 (Q2-G8 causality, critical):** `Week of Year=(202501-202553)` must be applied per uniform/generic rules; acceptance (53 intended 2025 periods, no 54–350) depends on live pivot/layout behavior. L3 on working copy is the gate; if filtering `Week of Year` does not constrain column-axis members, record evidence and make only evidence-backed, minimal scoped adjustments (including config/field/placement choices by PO if needed) — no speculative hardcoding.
- **U1/U5-U10:** verify field identity/name resolution, numeric-convertible/locale-safe handling, `Week of Year` item typing, report schema touchpoints, uniform-rule behavior deltas (2411 removal, subset-on-partial-match), and placement feasibility during implementation/L3.

Implementation proceeds sequentially via Implementor on T1→T2→T3→T4→T5→T6→T7 (plan §11.3), with exports kept in sync and no gitignored `data/theiss/*` committed. L1 `build_treport_vba.ps1` (report reviewed), **L2 manual VBE Compile on working copy** (required to claim compile), L3 on copies for `Create_report`/`Refresh_data` and regression across all `data_files`/reports remain mandatory before any acceptance claim.

- **B2:** Q2-G1..Q2-G9 **recorded and confirmed** in §2a (2026-11-16, latest PO confirmation: uniform rule applies to **all filters including report-level**; Q2-G3/Q2-G2 clarified). Design must honor: field missing OR none of specified values present → fail with clear error; at least one present → filter by present/matching values (subset) without error. Ranges require numeric-convertible values/bounds (`min`/`max` allowed); non-numeric range → fail; non-numeric selections via `general.delimeter` list. No hardcoded field semantics; logic applied equally to all reports.
- **R-3/R-4 (B3, Researcher):** options A/B/C analyzed; provisional recommendation **Option B** (generic engine refactor + narrowly scoped legacy-placeholder compatibility layer, single uniform matching policy) with contingencies — **not final/locked** due to runtime uncertainties (U1 live field name/caption resolution, U2 whether filtering `Week of Year` constrains column axis `Week Number`/`Start of Week` for 54–350, U3/U4 `max` resolution timing/global-per-field vs legacy Year-context impact on `report_m/w`, U5 placeholder contract, U6 report schema generality, U7 numeric/locale, U8 `Week of Year` item typing, U9 config reliance on 2411-skip/strict-missing, U10 placement/multi-select feasibility). R-4 L1/L2/L3 checklist outlined.
- **R-1 live/runtime confirmation:** still recommended on target Excel/working copies where feasible (causal attribution and Q2-G8 acceptance).

## 11. B3 detailed design and B4 task decomposition (2026-11-16, draft pending freeze)

- **Status:** Draft B3 design + B4 decomposition prepared from B1 (R-1 structural, R-2 trace), B2 PO decisions (Q2-G1..Q2-G9, uniform rule for all filters incl. report-level), and R-3/R-4 inputs. **Not final/locked.** No VBA/code changes made by this planning step.
- **Authoring basis:** semantic units only (modules/procedures/types/fields/config keys/contracts). No line references.
- **Binding decisions honored:** no hardcoded field semantics (§4.1, Q2-G8); filter logic applied equally to all reports (Q2-G9); uniform rule for all filters incl. report-level (field missing or none of specified values present/matchable → fail with clear error; at least one present → filter by present/matching values and complete without error); `max` from all available numeric values of that field, global/independent per field (Q2-G1); numeric literals and `min`/`max` allowed as range bounds, `start > end` after resolution → fail (Q2-G2); range on non-numeric values → fail, non-numeric selections via `general.delimeter` list, numeric = convertible to number (Q2-G3); keep placement or move available-unplaced field to Filters and apply, fail only if impossible (Q2-G6); duplicate/repeated field in `filters` not allowed → fail (Q2-G7).

### 11.1 Provisional design choice (recommended, contingent — not final)

**Recommended direction (provisional, Option B-like):** a **generic filter-engine refactor** with a **narrowly scoped legacy-placeholder compatibility layer** and a **single uniform matching/application policy** across `data_files.filters` and report-level filters.

- **Rationale:** aligns with CR-G1..CR-G5 and Q2-G1..Q2-G9 by removing hardcoded `Year/Month Number/Week Number` semantics, fixed `ResolvedContext` maxima, dependency coupling, and special-case report-skip behavior; enables arbitrary user-chosen pivot fields/expressions with deterministic fail-fast validation; limits compatibility risk by preserving existing placeholder tokens and report/config column names where feasible.
- **Contingent / not final:** this is **not** a final design selection. Freeze requires verification/resolution of key uncertainties:
  - **U1 (field identity/caption resolution):** live VBA/Excel resolution of configured field names vs pivot field captions/names for `Week of Year` and other fields.
  - **U2 (critical, defect causality):** whether applying `Week of Year=(202501-202553)` constrains the actual column/period axis driving output expansion (e.g. fields such as `Week Number`/`Start of Week` observed in R-1) or whether layout/axis/cache state also requires adjustment; acceptance “no extra columns 54–350, 53 intended 2025 periods” depends on this.
  - **U3/U4 (`max` timing/context, legacy compatibility):** clause application order and `min`/`max` resolution timing/source (global-per-field per Q2-G1 in current pivot context) vs legacy Year-first/Year-scoped behavior and impact on `report_m` (`Month Number=(1-max)|Year=max`) / `report_w` (`Week Number=(1-max)|Year=max`), including whether current clause order in configs must change.
  - **U8 (`Week of Year` item typing/matching):** numeric/cache item representation and robust numeric/string equivalence for matching `202501`–`202553`.
- **Other open contingencies (inform B3 freeze/refinement):** U5 placeholder contract details, U6 report schema generality, U7 numeric/locale conversion details, U9 configs relying on 2411-skip or strict-missing-token behavior, U10 placement/multi-select feasibility for arbitrary fields/orientations.

### 11.2 Detailed semantic design (provisional)

#### 11.2.1 Parsing and validation (`modTreportFilters` — replace/extend `ParsedClause`/`ParsedFilters`, `ParseAndClassify`, `ClassifyExpr`, `IsBoundToken`, `StripOuterParens`)

**Grammar (normative intent, generic):**
- `filters` is one or more clauses joined by `|`: `<field> = <expr> [ | <field> = <expr> ... ]`.
- Each clause: trimmed `<field>` (may contain spaces), `=`, trimmed `<expr>`. Empty filter text, empty clause set, missing/empty field or expression → fail with clear contextual error.
- `general.delimeter` (from config) is required for parsing list expressions and is used **only** inside expressions to split multi-value lists (not as clause separator).

**Disambiguation (deterministic, per expression):**
1. **List** if the normalized expression contains `general.delimeter`: treat as list. Strip optional outer parentheses for token extraction when present; split by `general.delimeter`; trim each token; empty token → fail with clear error.
2. **Range** if not a list and expression matches outer-parenthesized `(left-right)` form with exactly one valid range split into two **bound tokens** (after trim): `left` and `right` each must be a bound token (`min`, `max`, or numeric-convertible literal). Recognize `min`/`max` case-insensitively. Examples allowed per Q2-G2: `(1-max)`, `(min-max)`, `(min-50)`, `(10-max)`.
3. **Scalar** otherwise: single value token (including `min`/`max` where semantically valid as a scalar resolution token for the field, or any literal).

**Normalization and identity:**
- **Whitespace:** trim around `|`, `=`, field names, expressions, list tokens, and bound tokens.
- **Duplicate/repeated field (Q2-G7):** not allowed → fail with clear error. Detect duplicates by normalized field identity (trimmed; compare case-insensitively for duplicate detection to prevent ambiguous same-field repeats differing only by case). Preserve original/normalized display form for error messages.
- **Field identity for Pivot lookup:** resolve referenced field against PivotTable available fields using deterministic lookup semantics (prefer case-insensitive match against available pivot field names/captions as exposed by VBA `PivotFields`, while reporting the configured field string in errors). **U1 contingency:** confirm whether `Week of Year` must be matched by `PivotField.Name`, `PivotField.Caption`/source name, or both in target Excel/runtime.
- **Duplicate detection vs Pivot lookup:** duplicate detection operates on parsed filter clause field identities; Pivot existence/availability is validated per resolved clause against the actual pivot.

**Numeric-convertible test and range/list validity (Q2-G2, Q2-G3):**
- **Numeric = convertible to a number** (principle). Use a deterministic numeric-conversion test consistent with VBA/Excel numeric values; implementation must define locale-safe handling (avoid locale-dependent decimal/thousands surprises) while accepting integer and decimal forms that appear in pivot items/config. For cache values like `202501`/`202501.0`, numeric equivalence must hold (see U8).
- **Range validity:** allowed only when both bounds are resolvable to numeric values for that field context — bound tokens may be numeric literals or `min`/`max` (resolved from available numeric values of the referenced field). If range form is specified but bounds/field context cannot be treated as numeric-comparable per Q2-G3 → **fail with clear error** (non-numeric range not allowed). After resolving `min`/`max`, if `start > end` → **fail with clear error**. Range is inclusive.
- **List/non-range selections:** non-numeric values must be selected via list syntax using `general.delimeter` (single scalar for one value). List tokens may be numeric or non-numeric literals matched to pivot items by normalized comparison/numeric equivalence rules.
- **Scalar `min`/`max`:** treated as bound-token resolution against all available numeric values of the referenced field (requires numeric-available set; if none/unsuitable → fail clearly).

#### 11.2.2 Generic per-field resolution and context model (`modTreportFilters`, `modTreportRun`, `modTreportOutput`, `modTreportPivot` helpers)

**Replace hardcoded context:**
- Remove/retire `ResolvedContext` fixed fields (`YearFieldName`, `MaxYear`, `MaxMonth`, `MaxWeek`, `HasYear/HasMonth/HasWeek`), `ResolveMinMaxAndYear`, and `FindClauseIndex`-based Year-first logic.
- Introduce a **generic resolved-filter context** (semantic model) that records, per data-file filter application, clause outcomes needed for execution and placeholders without field-name hardcoding in the engine:
  - Parsed/validated clauses in declared order.
  - Per-clause resolution details: field (configured/normalized), expression kind, resolved bound/value set inputs as applicable, and a **resolved representative numeric value** when deterministically derivable (e.g. scalar numeric or resolved `min`/`max` value; range resolved `max`/end bound or other explicitly defined representative — must be specified consistently for placeholder compatibility).
  - A lookup by normalized field identity to resolved clause/value (for placeholder alias resolution and diagnostics), with duplicate prevention already enforced.

**`min`/`max` resolution (Q2-G1):**
- `min` / `max` = minimum/maximum among **all available numeric values of that referenced field** (independent/global per field, no hardcoded Year→Month/Week dependency). “Available numeric values” are values from that pivot field’s available items convertible to numbers under the numeric test.
- If no numeric values are available to resolve `min`/`max` → fail with clear contextual error.

**Clause application order and `max` resolution timing (explicit recommendation, provisional):**
- **Recommended:** process `data_files.filters` clauses in **declared order** (left-to-right). For each clause: (1) validate clause/field existence and placement applicability, (2) resolve `min`/`max` bound tokens (if any) from **all available numeric values of that specific field in the pivot’s current state at resolution time**, (3) build desired selection set, (4) apply selection per uniform/placement rules, (5) refresh/update pivot state as needed before the next clause so subsequent resolutions reflect prior selections in this batch.
- **Rationale:** generic and field-name agnostic; respects Q2-G1 “independent/global per field” (no cross-field formula) while acknowledging `current pivot context`; avoids old Year-first special case.
- **U3/U4 compatibility implications (`report_m`/`report_w`):** legacy configs are `report_m: Month Number=(1-max)|Year=max` and `report_w: Week Number=(1-max)|Year=max` (period clause before `Year=max`). If `max` is resolved at application time against current context, applying `Year=max` after the period range may yield different results than historical Year-scoped maxima. To preserve likely intent under generic order-driven behavior, configs may require reordering to `Year=max|Month Number=(1-max)` / `Year=max|Week Number=(1-max)` or equivalent verified form. **B3 freeze must decide** whether implementation keeps declared order strictly (with config migration/regression verification) or introduces a non-hardcoded generic pre-resolution strategy; any chosen strategy must not reintroduce field-name semantics. This is a primary compatibility contingency.

#### 11.2.3 Uniform selection/matching and placement (`modTreportPivot.ApplyFieldSelection`, `modTreportFilters.ApplySingleClauseSelection`/`BuildDesiredValues`/`NormalizeToPivotItem`, `modTreportFilters.ApplyDataFileFilters`, `modTreportRun.ProcessDataFile`)

**Uniform rule (all filters, including report-level, Q2-G4/Q2-G5):**
1. **Field missing from PivotTable available fields** → **fail with clear error** (context: operation, data_file, data_sheet, report/row if applicable, field, expression). Preserve/align with existing unavailable-field failure semantics (currently `GetPivotFieldRequired` → 2403) but ensure message/context is explicit.
2. **For specified field, none of specified values present/available/matchable** → **fail with clear error** (no skip-and-continue). Must include field and requested values/expression in context.
3. **At least one specified value present/matchable** → **filter by present/matching values and complete without error** (apply matched subset; unmatched tokens do not cause error).

**Evolution of selection behavior:**
- `modTreportPivot.ApplyFieldSelection` must be refactored from flag-driven special cases (`skipMissingTokensIfAnyMatch`, `skipReportIfNoMatches`) to a **single uniform policy** (or equivalent API) implementing the above rule for all callers:
  - Build desired/matchable item set by comparing requested tokens to available pivot items with robust normalization (trim, exact string match, and numeric equivalence where both sides convertible) — extend/reuse `NormalizeToPivotItem`/matching helpers; **U1/U8** verification required for caption/name and `Week of Year` numeric item forms.
  - Missing/unmatchable individual tokens: record/ignore **only when** at least one token matches; do **not** fail per-token in that case.
  - Zero matches / no visible valid selection after application: **fail with clear error** (replaces 2411 report-skip condition). Do not emit “skip report” warnings as success path.
- `modTreportFilters.ApplyDataFileFilters` and `ApplyReportFilters` call the same uniform selection semantics (no strict-vs-skip split).
- `modTreportRun.ProcessDataFile`: **remove 2411 skip-and-continue control flow** for “no matching articleLevel4Name_clean values” (and any similar report-skip-on-no-match path). The orchestrator must propagate filter failures as run-failing errors with full context; keep other non-filter error handling intact. Update/remove `ERR_REPORT_SKIP_NO_MATCH`/`WarnSkippedReport` usage as appropriate to the new semantics.

**Placement (Q2-G6, generic):**
- Preserve current flexible, safe behavior for all user-chosen fields: if field is active/placed in Filters/Rows/Columns (or otherwise already applied in a supported orientation), keep placement and apply filter there; if field is available but not placed, move to Filters/page area and apply (including best-effort sort/enable-multi-select behavior where supported). Fail clearly only if placement/selection is impossible for the field/orientation/pivot type.
- Retain field-scoped filter operations (no `ClearAllFilters` on pivot) and layout preservation intent from current implementation.

**`Week of Year` / report_wclip application note (U2):**
- Applying the configured filter to field `Week of Year` is necessary but may be insufficient if the pivot’s presented column axis is driven by a different field (R-1 indicates `Week Number`/`Start of Week` involvement). Design must not guess; T6/T7 must verify causality and whether any pivot layout/field placement alignment is required within scope (if required, treat as a scoped follow-up contingent on evidence and B3 freeze, not as assumed behavior here).

#### 11.2.4 Placeholder and `report_sheet` strategy (`modTreportOutput.SubstitutePlaceholders`, `modTreportRun` call sites, placeholder context)

**Compatibility goal:** preserve existing templates for `report_m` / `report_w` without reintroducing hardcoded filter-field semantics in parsing/matching/resolution logic.

**Concrete minimal approach (proposed):**
- **Generic source:** derive placeholder values from the **generic resolved-filter context** produced for the applied `data_files.filters` clauses (not from hardcoded Year/Month/Week engine branches).
- **Legacy token support:** continue to recognize `<max_year>`, `<max_monthnum>`, `<max_weeknum>`.
  - Resolve `<max_year>` → resolved numeric value associated with clause field `Year` in the generic context (by normalized field-name lookup/alias).
  - Resolve `<max_monthnum>` → resolved numeric value associated with clause field `Month Number`.
  - Resolve `<max_weeknum>` → resolved numeric value associated with clause field `Week Number`.
- **Representative value rule (to finalize at implementation/freeze):** for each field, use a deterministic value derived from the clause’s resolved result (recommended: resolved `max`/upper bound for range/`max` expressions, resolved scalar numeric value for scalar numeric/`min`/`max` scalar forms) consistent with legacy expectations for `1-max`/`max` usage. **U3/U4/U5 contingency:** exact meaning under global-per-field/order-driven resolution must be verified against `report_m/w` outputs.
- **No hardcoded filter semantics:** engine does not special-case Year/Month/Week during filter evaluation; it only maps legacy placeholder tokens to generically resolved per-field results by field name when substitution is requested.
- **Failure behavior:** if a placeholder token referenced in `report_sheet` cannot be resolved from available generic context (missing clause/field result, non-numeric or undefined representative as required) → **fail with clear contextual error** (operation/data_file/report context, placeholder token, expected source field/value).
- **`report_wclip`:** `report_sheet` is literal `Хроно по неделям 2025` with no placeholders; unaffected by substitution except existing normalization/sanitization/collision checks.
- **Extensibility (non-goal now):** optional generic placeholder syntax (e.g. field-parameterized tokens) may be considered later but is out of B4 scope unless required by freeze.

#### 11.2.5 Report-level filters and config touchpoints (`modTreportFilters.ApplyReportFilters`, `modTreportConfig`)

**Uniform application (Q2-G9, Q2-G4/Q2-G5):**
- Apply report-level filters with the same generic, uniform rule as data-file filters: field must exist in pivot available fields; none of specified values match → fail; at least one matches → apply matched subset; no 2411 skip semantics.
- Affects both `articleLevel4Name_clean` and `targetAudience` selections (behavior change where configs relied on skip-on-none or strict-missing-token rejection).

**Minimal compatible handling (proposed, U6 noted):**
- **Config/schema:** retain existing `reports` table columns and required fields (`articleLevel4Name_clean`, `targetAudience`, `name`) and existing `TreportReport` structure/validation in `modTreportConfig` to minimize migration risk, unless B3 freeze identifies a required schema evolution.
- **Application:** stop hardcoding differential behavior by field; treat each configured report-level field as a generic selection on the named pivot field using uniform matching, current delimiter-based multi-value splitting (`SplitReportValues` semantics may be simplified/reused but must align with uniform rule), and contextual errors.
- **Schema generality (U6):** full generalization to arbitrary report-level filter columns/expressions is **not assumed** in this draft; if required, define a separate config evolution. Current proposal is “same columns/names, generic uniform semantics.”

#### 11.2.6 Error taxonomy and messages (cross-cutting)

All failures must be contextual and clear (fail-fast). Standard context elements to include where applicable:
- **Operation:** `Refresh_data` / `Create_report` phase (and sub-operation: config read, parse, resolve, apply filter, placeholder substitution, output/sheet naming).
- **Data context:** `data_file`, `data_sheet`, `row`/`data_files` row index, source path as available.
- **Report context:** `report` name, `reports` row index, output workbook/name as available.
- **Filter context:** configured field name, raw clause/expression or token/value(s), expression kind (scalar/range/list) when relevant.
- **Cause:** e.g. field not found in pivot available fields, none of specified values present/matchable, invalid/empty clause or token, non-numeric range, `min`/`max` unresolved (no numeric values), `start > end`, duplicate field, placeholder unresolved, placement/application impossible.

Preserve meaningful error codes where stable/useful, but prioritize correctness and clarity of the uniform rule over legacy skip/warn codes (2411 no longer indicates skip-and-continue).

### 11.3 B4 task decomposition (sequential, dependency-ordered; semantic)

**Global prerequisites/constraints (apply to all implementation tasks):**
- **B3 freeze gate:** do not implement/finalize behavior-dependent code paths until B3 design is frozen with accepted resolutions for U1-U4,U8 (and other Ux as they affect the task). Tasks below are implementation-ready specifications **pending freeze**.
- Edit VBA export sources under `mko_treport/src/vbs/export/*.bas|*.cls` only (keep exports in sync; do not treat workbooks as source of truth except controlled build/sync flow). Do **not** commit gitignored `data/theiss/*`.
- Follow L1/L2/L3 workflow: work on backups/working copies only; run `build_treport_vba.ps1` and review report (L1, VBOM caveats); **manual VBE `Debug → Compile VBAProject` on working copy required to claim compile (L2)**; functional checks on copies when logic/run behavior changes (L3). No macros on originals.

```yaml
id: 003-T1-generic-filter-parse-validate
title: Generic filter parsing, classification, and validation model
priority: high
depends_on:
  - B3-freeze (U1,U7,U8 affect field identity and numeric-convertible/matching details)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Replace hardcoded-period parsing assumptions with generic clause parsing/validation for arbitrary pivot fields.
  - Implement grammar/disambiguation: clause split by '|', '<field>=<expr>', list when delimiter present, range when '(bound-bound)' with two bound tokens, else scalar.
  - Enforce duplicate/repeated field detection (Q2-G7), whitespace trimming, and normalized field identity handling for duplicates vs pivot lookup.
  - Enforce numeric-convertible semantics and range validity (Q2-G2,Q2-G3): numeric literals/min/max bounds, min/max case-insensitive tokens, non-numeric range fail, start>end after resolution fail (resolution hook/interface prepared), locale-safe conversion considerations documented in code behavior per design.
  - Provide clear contextual parse/validation errors (operation/data_file/field/expression/kind).
files:
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
      - type: type
        name: ParsedClause
      - type: type
        name: ParsedFilters
      - type: procedure
        name: ParseAndClassify
      - type: procedure
        name: ClassifyExpr
      - type: procedure
        name: StripOuterParens
      - type: procedure
        name: IsBoundToken (or successor bound-token validation)
      - type: procedure
        name: SplitReportValues (reuse/adapt for delimiter-based tokenization where compatible)
    semantic_anchors:
      insert_after:
        type: declaration_region
        value: generic clause/expression metadata types and helpers replacing Year/Month/Week-specific assumptions
      insert_before:
        type: legacy_procedure
        value: ResolveMinMaxAndYear / hardcoded period-resolution entry points (to be removed/refactored in T2)
    changes:
      - action: replace_types_and_parsing
        description: >
          Redefine parsed clause structures to carry generic field identity, raw/normalized expression,
          kind (scalar/range/list), bound/value tokens, and validation metadata needed by generic
          resolution/selection. Implement deterministic parse/classify/validate pipeline per B3 design.
      - action: add_helpers
        description: >
          Add helpers for normalized field identity comparison (duplicate detection), numeric-convertible
          testing, bound-token recognition/resolution interface, and contextual error construction.
acceptance_criteria:
  - Accepts arbitrary field names and valid scalar/range/list expressions per Q2-G2/Q2-G3 and general.delimeter rules.
  - Rejects empty/invalid clauses, duplicate fields (Q2-G7), non-numeric ranges, invalid bounds, and malformed expressions with clear contextual errors.
  - No hardcoded Week/Month/Year/Week of Year semantics in parsing/validation logic.
  - Ready for consumption by generic per-field resolution and uniform selection tasks.
```

```yaml
id: 003-T2-generic-min-max-resolution
title: Generic per-field min/max resolution and context (remove hardcoded period context)
priority: high
depends_on:
  - 003-T1-generic-filter-parse-validate
  - B3-freeze (U3,U4,U5 clause order/timing and placeholder representative-value contract)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Remove Year-first/hardcoded period resolution (FindClauseIndex/ResolveMinMaxAndYear/ResolvedContext MaxYear|MaxMonth|MaxWeek coupling).
  - Implement field-agnostic min/max resolution from all available numeric values of the referenced field (Q2-G1).
  - Define and implement clause processing order and min/max resolution timing per frozen B3 decision, with explicit handling for current pivot context updates.
  - Produce generic resolved-filter context/results consumable by application and placeholder substitution (no field-name semantics in engine).
files:
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
      - type: type
        name: ResolvedContext (remove or replace)
      - type: procedure
        name: ResolveMinMaxAndYear (remove/replace)
      - type: procedure
        name: FindClauseIndex (remove/replace if period-specific)
      - type: procedure
        name: ResolveBound
      - type: procedure
        name: ResolveClauseMax
      - type: procedure
        name: MaxOfStringArray
      - type: procedure
        name: ApplyDataFileFilters (signature/flow adaptation)
    semantic_anchors:
      insert_after:
        type: parsing_api
        value: generic parsed-filters representation from T1
      insert_before:
        type: selection_api
        value: uniform clause selection/application entry points used by run orchestration
    changes:
      - action: replace_resolution_engine
        description: >
          Implement generic bound/value resolution per field and clause using pivot available numeric
          values at resolution time; eliminate hardcoded Year/Month/Week indices and dependency logic.
      - action: redefine_context
        description: >
          Introduce generic resolved context/map for per-clause resolved values and placeholder-compatible
          representative outputs, preserving enough information for diagnostics and substitution.
  - path: mko_treport/src/vbs/export/modTreportPivot.bas
    targets:
      - type: module
        name: modTreportPivot
      - type: procedure
        name: GetPivotFieldRequired
      - type: procedure
        name: FieldAvailableNumericValues
      - type: procedure
        name: FieldAvailableValues
      - type: procedure
        name: MinNumeric
      - type: procedure
        name: MaxNumeric
    semantic_anchors:
      insert_after:
        type: field_lookup_helpers
        value: available-field/value enumeration helpers
    changes:
      - action: extend_helpers_if_needed
        description: >
          Adapt or add helpers to support generic, deterministic numeric/value enumeration and
          case-insensitive/robust field resolution required by T1/T2 without introducing field semantics.
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: module
        name: modTreportRun
      - type: procedure
        name: ProcessDataFile
    semantic_anchors:
      insert_after:
        type: parse_step
        value: ParseAndClassify/validated filter parsing for data_files row
    changes:
      - action: update_orchestration_context
        description: >
          Replace ResolveMinMaxAndYear/ResolvedContext usage with generic resolution/application flow and
          pass generic resolved context to downstream steps (selection, placeholders) as defined by freeze.
acceptance_criteria:
  - min/max resolve from all available numeric values of the referenced field only (Q2-G1), no hardcoded field dependencies.
  - Clause order/timing and pivot refresh behavior follow frozen B3 decision; U3/U4 compatibility impacts for report_m/report_w identified/handled per freeze (including config ordering if required).
  - No remaining fixed MaxYear/MaxMonth/MaxWeek or Year-first resolution path in runtime logic.
```

```yaml
id: 003-T3-uniform-selection-placement-and-run-flow
title: Uniform matching/selection, placement preservation, and removal of 2411 skip semantics
priority: high
depends_on:
  - 003-T2-generic-min-max-resolution
  - B3-freeze (U1,U2,U8,U10 field resolution, axis/placement feasibility, item matching)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Implement uniform rule for all filters: field missing in available fields -> fail; none of specified values present/matchable -> fail; >=1 present -> apply matched subset without error.
  - Evolve selection API/behavior to remove strict-vs-skip flag special cases and 2411 report-skip semantics.
  - Preserve generic placement behavior (Q2-G6): keep active placement, move available-unplaced to Filters and apply, fail only if impossible.
  - Ensure pivot item matching is robust for numeric/string equivalence and configured field identity per freeze.
files:
  - path: mko_treport/src/vbs/export/modTreportPivot.bas
    targets:
      - type: module
        name: modTreportPivot
      - type: procedure
        name: ApplyFieldSelection
      - type: procedure
        name: GetPivotFieldRequired
      - type: procedure
        name: EnsureFilterPlacement
      - type: procedure
        name: FindMatchingAvailable
      - type: procedure
        name: IsWantedItem
      - type: procedure
        name: FieldAvailableValues
    semantic_anchors:
      insert_after:
        type: field_resolution
        value: available-field lookup/classification
      insert_before:
        type: error_raise
        value: selection/placement failure paths (must emit clear contextual errors)
    changes:
      - action: unify_selection_policy
        description: >
          Replace skipMissingTokensIfAnyMatch/skipReportIfNoMatches behavior with uniform subset-if-any,
          fail-if-none policy; remove 2411-skip signaling; keep field-scoped operations and layout safety.
      - action: preserve_generic_placement
        description: >
          Maintain/apply Q2-G6 placement rules for arbitrary fields with clear failure only when impossible.
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
      - type: procedure
        name: ApplySingleClauseSelection
      - type: procedure
        name: BuildDesiredValues
      - type: procedure
        name: NormalizeToPivotItem
      - type: procedure
        name: ApplyDataFileFilters
    semantic_anchors:
      insert_after:
        type: resolution_output
        value: generic resolved clause values/bounds
    changes:
      - action: align_clause_application
        description: >
          Build desired selections from resolved generic clauses and apply via uniform pivot selection;
          ensure unmatched tokens do not error when any match exists and none-match errors clearly.
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: module
        name: modTreportRun
      - type: procedure
        name: ProcessDataFile
      - type: constant
        name: ERR_REPORT_SKIP_NO_MATCH
      - type: procedure
        name: WarnSkippedReport
    semantic_anchors:
      insert_after:
        type: per_report_filter_apply
        value: ApplyReportFilters call site
    changes:
      - action: remove_skip_continue
        description: >
          Remove report-skip-on-no-match handling/warning flow; propagate uniform none-match/missing-field
          failures as contextual run errors. Keep report iteration structure otherwise intact unless freeze requires change.
acceptance_criteria:
  - Uniform rule holds for data-file and report-level filters (Q2-G4/Q2-G5,Q2-G9): missing field or zero matches -> fail; >=1 match -> subset applied without error.
  - No 2411 skip-and-continue behavior remains; no silent ignore of all-missing selections.
  - Placement behavior matches Q2-G6 for available/placed/unplaced fields; failures are clear and contextual.
  - Selection works for generic fields including numeric-like values (e.g. Week of Year items) subject to U1/U8 verification.
```

```yaml
id: 003-T4-placeholder-substitution-compatibility
title: Placeholder and report_sheet substitution compatibility via generic resolved values
priority: high
depends_on:
  - 003-T2-generic-min-max-resolution
  - B3-freeze (U3,U4,U5 placeholder contract and representative-value semantics)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Preserve existing <max_year>, <max_monthnum>, <max_weeknum> placeholders for report_m/report_w without hardcoded filter semantics in engine.
  - Resolve legacy placeholder tokens from generic resolved-filter context by field-name alias mapping to generically computed per-field results.
  - Fail clearly when a referenced placeholder cannot be resolved from available context.
  - Keep literal report_sheet names (e.g. report_wclip) behavior unchanged except existing normalization/collision handling.
files:
  - path: mko_treport/src/vbs/export/modTreportOutput.bas
    targets:
      - type: module
        name: modTreportOutput
      - type: procedure
        name: SubstitutePlaceholders
    semantic_anchors:
      insert_after:
        type: procedure_signature
        value: placeholder substitution entry point
    changes:
      - action: refactor_substitution
        description: >
          Replace fixed (yearVal,monthVal,weekVal) dependency with generic resolved-context/value lookup
          (or compatible adapter) supporting legacy tokens mapped to Year/Month Number/Week Number
          resolved values; enforce fail-fast unresolved-placeholder errors with context.
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: module
        name: modTreportRun
      - type: procedure
        name: ProcessDataFile
    semantic_anchors:
      insert_after:
        type: post_filter_apply
        value: point where report_sheet name is derived per report iteration
    changes:
      - action: update_call_site
        description: >
          Pass generic resolved context/lookup to SubstitutePlaceholders and handle failures contextually.
acceptance_criteria:
  - report_m/report_w sheet naming remains correct under frozen resolution/order semantics (U3/U4 verified).
  - Unresolvable referenced placeholders fail with clear contextual errors.
  - No reintroduction of hardcoded Year/Month/Week evaluation logic in filter engine; only alias-based lookup into generic results.
```

```yaml
id: 003-T5-report-level-filters-and-config-touchpoints
title: Unify report-level filter application and align config touchpoints
priority: high
depends_on:
  - 003-T3-uniform-selection-placement-and-run-flow
  - B3-freeze (U6 report schema generality, U9 behavior-change impact)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Apply uniform filter rule equally to report-level filters (articleLevel4Name_clean, targetAudience) with no field-specific skip/strict split.
  - Keep minimal compatible config/schema handling (existing reports columns/required fields and TreportReport shape) unless freeze requires change.
  - Ensure contextual errors for report-level field/value failures and preserve delimiter-based value parsing consistent with general.delimeter.
files:
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
      - type: procedure
        name: ApplyReportFilters
      - type: procedure
        name: SplitReportValues
    semantic_anchors:
      insert_after:
        type: uniform_selection_api
        value: shared field selection/matching entry point from T3
    changes:
      - action: unify_report_filters
        description: >
          Refactor ApplyReportFilters to apply configured report fields using uniform selection semantics
          (matched subset if any, fail if none/missing field), removing articleLevel4Name_clean-only
          skip behavior and strict-vs-subset divergence.
  - path: mko_treport/src/vbs/export/modTreportConfig.bas
    targets:
      - type: module
        name: modTreportConfig
      - type: type
        name: TreportReport
      - type: procedure
        name: ReadConfig
    semantic_anchors:
      insert_after:
        type: reports_table_validation
        value: reports ListObject/column reads and required-field checks
    changes:
      - action: adjust_only_if_required
        description: >
          Preserve existing reports schema/required columns unless B3 freeze mandates evolution (U6).
          Adjust validation/messages only as needed for compatibility with uniform runtime semantics and
          clear diagnostics; do not introduce speculative schema generalization.
acceptance_criteria:
  - Report-level filters follow the same uniform rule as data_files.filters (Q2-G4/Q2-G5,Q2-G9).
  - No report-skip-on-none-match path; subset applied when at least one configured value matches.
  - Existing reports table/columns remain supported per minimal-compat decision (or documented migration if freeze changes this).
```

```yaml
id: 003-T6-report-wclip-alignment-and-regression-sweep
title: report_wclip config alignment verification and regression sweep across all data_files/reports
priority: high
depends_on:
  - 003-T4-placeholder-substitution-compatibility
  - 003-T5-report-level-filters-and-config-touchpoints
  - B3-freeze (U1-U4,U8; acceptance contingent on U2)
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Verify report_wclip configuration intent (data_files.filters = Week of Year=(202501-202553), data_sheet/report_sheet as configured) is correctly applied by the generalized engine on working copies.
  - Assess U2 causality: confirm whether filtering Week of Year constrains the output column axis/members responsible for extra columns 54-350; record evidence and any scoped follow-up required (no speculative fix).
  - Execute regression verification for all data_files rows (report_m, report_w, report_wclip) and all reports/outputs per Q2-G9, including placeholder-derived sheet names and layout/formatting expectations on copies.
  - Validate behavior changes from uniform rule (no 2411 skip; subset-if-any) against existing configs and document observed impacts.
files:
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
    semantic_anchors:
      insert_after:
        type: uniform_engine
        value: finalized generic parse/resolve/select behavior from T1-T5
    changes:
      - action: verify_and_adjust_if_evidenced
        description: >
          Only minimal, evidence-backed adjustments required by U1-U4/U8/U2 findings and frozen design;
          no unrelated refactors.
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: module
        name: modTreportRun
      - type: procedure
        name: Create_report
      - type: procedure
        name: ProcessDataFile
    semantic_anchors:
      insert_after:
        type: orchestration_flow
        value: end-to-end per-data_file/per-report processing
    changes:
      - action: verify_regression
        description: >
          Confirm orchestration correctness for all configured data_files/reports under uniform semantics
          and generic placeholder/context behavior.
acceptance_criteria:
  - On working-copy evidence (L3 context), report_wclip targets intended Week of Year 202501-202553 selection; if U2 confirms axis coupling, output shows intended 2025 week period columns only (53 periods implied) with intended layout/totals and no extra columns 54-350. If U2 not satisfied by filter-only change, causality and gap are explicitly evidenced (acceptance contingent, not falsely claimed).
  - report_m and report_w produce expected outputs/sheet names on copies under frozen compatibility decisions (U3/U4/U5 resolved or explicitly mitigated).
  - All configured data_files/reports processed per Q2-G9 with uniform rule behavior; failures (if any) are clear, contextual, and intentional per spec/decisions.
```

```yaml
id: 003-T7-build-verification-and-evidence
title: L1/L2/L3 build verification and evidence capture
priority: high
depends_on:
  - 003-T6-report-wclip-alignment-and-regression-sweep
source_reference: .ai/plans/003-report-wclip-week-filter.md
goals:
  - Execute L1 build/sync on working copies and review generated reports (including VBOM/compile-note caveats).
  - Obtain L2 manual VBE Debug -> Compile VBAProject result on working copy (required to claim compile).
  - Execute L3 functional checks on copies (Create_report/Refresh_data as relevant) without modifying originals; capture regression and report_wclip evidence.
  - Record changed export files, exact script commands, report paths/summaries, L2 outcome, L3 outcome, and U1-U10 observations relevant to acceptance.
files:
  - path: mko_treport/src/vbs/export/modTreportFilters.bas
    targets:
      - type: module
        name: modTreportFilters
    semantic_anchors:
      insert_after:
        type: final_state
        value: implemented generic filter engine
    changes:
      - action: verify
        description: No source changes except evidence-driven fixes already scoped in prior tasks; confirm export consistency.
  - path: mko_treport/src/vbs/export/modTreportPivot.bas
    targets:
      - type: module
        name: modTreportPivot
    semantic_anchors:
      insert_after:
        type: final_state
        value: uniform selection/placement helpers
    changes:
      - action: verify
        description: Confirm build/compile/runtime behavior on copies.
  - path: mko_treport/src/vbs/export/modTreportRun.bas
    targets:
      - type: module
        name: modTreportRun
    semantic_anchors:
      insert_after:
        type: final_state
        value: orchestration without 2411 skip semantics
    changes:
      - action: verify
        description: Confirm end-to-end execution and error handling on copies.
  - path: mko_treport/src/vbs/export/modTreportOutput.bas
    targets:
      - type: module
        name: modTreportOutput
    semantic_anchors:
      insert_after:
        type: final_state
        value: generic placeholder substitution compatibility
    changes:
      - action: verify
        description: Confirm sheet naming/substitution outcomes on copies.
  - path: mko_treport/src/vbs/export/modTreportConfig.bas
    targets:
      - type: module
        name: modTreportConfig
    semantic_anchors:
      insert_after:
        type: final_state
        value: compatible config read/validation touchpoints
    changes:
      - action: verify
        description: Confirm config loading/validation behavior on copies if touched.
acceptance_criteria:
  - L1 report produced from working copy and reviewed; VBOM trust/automation limitations explicitly noted if present.
  - L2 manual VBE compile performed on working copy with result recorded (pass/fail with errors); no "compiled" claim without L2.
  - L3 run on copies only (no originals corrupted, source files unchanged per rules) with outcomes recorded for report_wclip and regression set.
  - Evidence package sufficient to judge Q2-G8 acceptance (contingent on U2) and uniform-rule/regression outcomes.
```

### 11.4 Block-level acceptance criteria (B3/B4 handoff)

- **Q2-G8 / defect target (contingent on U2):** for `report_wclip` (`DOLGIT W30-50`, sheet `Хроно по неделям 2025`, `data_file=report_wclip`, `data_sheet=clean_data_full`, intended `filters = Week of Year = (202501-202553)`), engine must select all available values in interval 202501–202553 for field `Week of Year` when field exists in pivot available fields. Intended outcome is 2025 week period columns only (53 periods implied) plus intended layout/totals, with **no extra columns 54–350**. This outcome is **explicitly contingent** on U2 verification that the constrained field drives the relevant column axis/members; if not, filter correctness and remaining layout/axis cause must be separately evidenced.
- **Uniform rule (all filters, incl. report-level):** (1) referenced field not in PivotTable available fields → fail with clear contextual error; (2) none of specified values present/available/matchable for that field → fail with clear contextual error; (3) at least one present/matchable → apply matched subset and complete without error. No 2411 skip-and-continue; behavior changes vs legacy strict-missing and report-skip cases are intentional per Q2-G4/Q2-G5/Q2-G9.
- **Generic semantics:** no hardcoded Week/Month/Year/period field logic in filter engine; `max` = maximum among all available numeric values of the referenced field (Q2-G1); range bounds allow numeric literals and `min`/`max`, `start > end` after resolution fails (Q2-G2); ranges require numeric-convertible values/bounds, non-numeric ranges fail and non-numeric selections use list form (Q2-G3); duplicate fields fail (Q2-G7); placement preserved/moved-to-Filters per Q2-G6.
- **Compatibility/regression (Q2-G9):** `report_m`, `report_w`, and all other configured `data_files`/reports produce correct, evidenced results on working copies under frozen compatibility decisions (including placeholder resolution for `<max_year,max_monthnum,max_weeknum>` and clause order/timing implications U3/U4/U5). Any required config migration is explicit and verified.
- **Verification discipline:** L1 `build_treport_vba.ps1` on working copy with report reviewed; **L2 manual VBE `Debug → Compile VBAProject` on working copy required to claim compile**; L3 functional checks on copies when logic/run behavior changes. No originals corrupted, no gitignored `data/theiss/*` committed.

### 11.5 Risks and contingencies (U1-U10) and verification mapping

| ID | Contingency / risk | Impact | Verification / resolution (mapped tasks) |
|---|---|---|---|
| **U1** | Live field name/caption resolution differs from configured strings or cache names (e.g. `Week of Year` identity as VBA/Excel sees it). | Field-missing false failures or wrong field targeted. | Confirm pivot field lookup semantics against available fields on target Excel/working copy; adjust lookup/matching if evidenced. **T1,T3,T6,T7**. |
| **U2 (critical)** | Filtering `Week of Year` may not constrain column axis driven by different week/time fields (`Week Number`/`Start of Week` observed in R-1), cache/layout state may still yield extra members/columns 54–350. | Q2-G8 acceptance (53 intended periods, no 54–350) may not be achieved by filter-only fix. | L3 evidence on working copy: compare intended vs actual column/axis members and field driving expansion; record causality and scoped follow-up if needed. **T6,T7**. |
| **U3** | `max` resolution timing/source and clause order: global-per-field in current pivot context vs legacy Year-first/ordered behavior. | Different selected maxima/ranges, especially multi-clause configs. | Freeze order/timing rule; verify `report_m/w` and `report_wclip` results; migrate clause order/config if required. **T2,T4,T6,T7**. |
| **U4** | Legacy Year-context impact: period `max` historically scoped after Year selection; generic order-driven resolution may change results. | Regression in `report_m`/`report_w` sheet names/data if not aligned. | Explicit regression comparison on copies; confirm placeholder representative values. **T2,T4,T6,T7**. |
| **U5** | Placeholder contract for representative values (scalar/range/max) and alias mapping completeness. | Wrong `report_sheet` names or unresolved placeholders. | Finalize mapping rule at freeze; unit-like functional checks via L3 outputs. **T2,T4,T6,T7**. |
| **U6** | Report schema generality: only `articleLevel4Name_clean`/`targetAudience` columns modeled today. | Limits future generic report-level fields; minimal-compat change may leave schema coupled. | Keep minimal compatible schema now; document evolution need only if evidenced/required. **T5,T6**. |
| **U7** | Numeric/locale conversion (`IsNumeric`/`CDbl`/decimal separators) and numeric-equivalence behavior. | Misclassification of numeric values/ranges or match failures. | Implement deterministic, locale-safe numeric handling consistent with Q2-G3; verify with real item formats. **T1,T3,T7**. |
| **U8** | `Week of Year` item typing in pivot/cache (e.g. `202501` vs `202501.0`/text forms). | Requested range values fail to match despite being present. | Robust numeric/string equivalence in normalization/matching; L3 verification against `report_wclip`. **T1,T3,T6,T7**. |
| **U9** | Existing configs may rely on 2411 skip-on-none-match or strict-missing-token rejection behavior. | Intentional behavior changes surface as new failures under uniform rule. | Regression sweep across all `data_files`/reports on copies; classify failures as spec-aligned corrections vs migration needs. **T3,T5,T6,T7**. |
| **U10** | Placement/multi-select feasibility for arbitrary fields/orientations/pivot types when moved or filtered in place. | Apply/placement failures not previously encountered. | Exercise generic placement on configured fields in copies; fail clearly if impossible per Q2-G6. **T3,T6,T7**. |

**Freeze condition reminder:** B4 implementation execution begins only after B3 freeze (approach/compatibility decisions finalized with U1-U4,U8 addressed sufficiently), acceptance criteria locked for Q2-G8 (with U2 contingency explicit) and Q2-G9 regression scope, and L1/L2/L3 plan confirmed executable on working copies.

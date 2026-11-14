# Execution plan: treport (001_treport)

- **Plan ID:** 001-treport-execution-plan
- **Replaces:** `.ai/plans/001-treport-execution-plan.md` (this document is the corrected replacement content)
- **Recommended filename:** `.ai/plans/001-treport-execution-plan.md` (unchanged)
- **Based on:** `.ai/problems/001_treport_spec.md` — **rev. 3** (2026-11-14, 267 lines), all clarification options resolved by Product Owner as “all recommended”
- **Fresh code/artifact context (Auditor, repo vs rev. 3):** `.ai/problems/001_treport_spec.md` rev. 3 present. Local (gitignored) artifacts observed: `data/theiss/treport.xlsm` with `run`/`settings`, 4 ListObjects matching rev. 3 schema/values (`general`: `input_path,output_path,delimeter`; `reports`: 6 rows with `articleLevel4Name_clean,targetAudience,name` incl. trailing space `DOLGIT `; `data_files`: 3 rows with `filters`/placeholder tokens; `placeholders`: `max_weeknum,max_monthnum,max_year`), `xl/vbaProject.bin` present, **no exported VBA source anywhere**, `run` has one button “Button 1” bound to `[0]!RunXlconcat` (no controls named `Refresh_data`/`Create_report` evidenced). No `xl/connections.xml`/pivots in `treport.xlsm`. `data/theiss/base/report_wclip.xlsx`, `monthly/report_m.xlsx`, `weekly/report_w.xlsx` exist with sheet `clean_data_full`, PQ connections/query tables/customXml and pivot cache/table parts; pivot caches show fields incl. `articleLevel4Name_clean`, `targetAudience`, `Year`/`Years` (partly), `Month Number`, `Week Number`, but exact pivot selection/orientation/captions/types, `Years` availability in all files, and presented ranges not fully verified (OQ-T1). PQ connections include user/machine-specific paths (portability/refresh risk). No `data/theiss/output/` directory, no golden outputs. `mko_treport/src/{scripts,vbs,xls}` empty. No implementation/tests/CI/tooling. Known mismatches/gaps: `data_files` row `report_m` has `group_rows=TRUE` (conflicts D13, R4); single `RunXlconcat` control vs FR-RUN-1; VBA opaque (OQ-T4); refresh behavior unknown (OQ-T2); pivot/filter/copy fidelity unknown (OQ-T1,T3); OQ-B7 golden baseline missing; config trailing space in one report name.
- **Planning intent:** dependency-safe, minimal-scope, incremental decomposition into execution blocks aligned to rev. 3. **No final implementation approach is chosen where technical uncertainty exists** (VBA architecture/refactor approach, exact PQ refresh completion technique, pivot identification/copy strategy, and related platform-dependent mechanics remain alternatives pending OQ-T1–T4 evidence). No code or existing audit-phase files are changed by this plan.
- **Execution posture:** **Implementation is in-scope to begin per plan where gates allow** (spec planning gate items 1–3 satisfied per §14). Blocks/phases that can start now vs gated are marked explicitly below. Correctness sign-off remains blocked until OQ-B7 golden baseline and OQ-T1–T4 validations are closed.

## Scope and guardrails

- **In scope:** decompose, sequence, and begin implementation/validation work to satisfy spec rev. 3 §§4–7/14 within the existing `treport.xlsm` Excel-macro model and configured `settings`/`run`/ListObjects (`general`, `reports`, `data_files`, `placeholders`) for operations `Refresh_data` and `Create_report`, including deterministic policies and verification aligned to D1–D17.
- **In scope (rev. 3 semantics to realize):** config reading/validation (FR-CFG); deterministic data-file resolution with precedence `.xlsx,.xlsm,.xlsb` (D2); access modes (D1: `Create_report` read-only never saved; `Refresh_data` writable refresh PQ/RefreshAll+save+close); `filters` DSL parse and dependency-aware `min/max` resolution and `max_*` placeholder derivation (D5,D6,D10,D11); data-file/report filter staging and placement (D4,D7,D8,D9); pivot snapshot extraction and static paste (D15); formatting/layout with `group_rows` fail-fast policy (D13); output assembly, placeholder substitution, naming/collision/idempotency (D3,D11,D14); fail-fast contextual diagnostics incl. no-data (D12); dropped parameters not required (D17).
- **Out of scope / non-goals:** architecture redesign or migration off Excel/VBA; non-Excel runtime/scheduler/UI redesign; speculative abstractions or invented requirements; source data/schema/content or upstream PQ/data-model changes; implementing or choosing approaches that remain technically uncertain before evidence; `group_rows` color-band/outline behavior until an explicit deterministic rule is approved (D13); pixel-perfect parity without an agreed golden output (D18/OQ-B7); BI/business meaning beyond configured names/expressions.
- **Method:** semantic anchors only (workbook/ListObjects/sheets/operations/fields/contracts/modules by responsibility) — no line numbers or positional assumptions. Tasks/blocks are atomic, independently reviewable/executable where feasible, and ordered by dependency/risk. Prefer isolated changes, incremental evolution, and backward-compatible/stable task boundaries.
- **Execution constraint (this interaction):** never change any code or existing audit-phase files. Plan may prescribe implementation/config actions (including edits to local gitignored artifacts) to be performed in subsequent implementation steps where permitted.

## Resolved decisions and open gates

### Confirmed product decisions (D1–D17) — planning baseline fixed

| ID | Decision (rev. 3) | Planning/implementation implication |
|---|---|---|
| D1 | `Create_report`: data files read-only, never saved. `Refresh_data`: writable (not read-only), refresh PQ/RefreshAll, save, close. | Enforce mode-specific open/close/save and state-safe cleanup (B6,B7,B12). |
| D2 | `data_file` base name resolved in `input_path` to `<data_file>.<ext>` with fixed precedence **`.xlsx`, `.xlsm`, `.xlsb`**; **fail fast** if no match or >1 match. | Deterministic resolution and contextual errors (B4,B6). |
| D3 | One output workbook per `reports` row named `reports.name` + `.xlsx`; overwrite if exists on successful run; sheets in `data_files` order, recreated/replaced deterministically for idempotent reruns. | Output lifecycle/idempotency (B11). |
| D4 | `data_files.filters` contains period/data fields as configured; `articleLevel4Name_clean` and `targetAudience` controlled via `reports`. | Filter responsibilities and staging (B5,B8). |
| D5 | `filters` grammar: clauses `|`; `<field> = <expr>` (spaces around `=`,`|` trimmed; field names may contain spaces). Expr: scalar literal or `min`/`max`; range `(start-end)` **inclusive**, bounds numeric or `min`/`max`, `-` range operator; list `(v1<delimeter>v2…)` via `general.delimeter`. Exact-after-trim match to pivot captions/items. Disambiguation: contains `general.delimeter` → list; else contains `-` between two bound tokens → range; else scalar. | Strict parser/validator/disambiguation (B5). |
| D6 | Duplicated `Week Number=(1-max)|Years=max` example line treated as typo/ignored. | Stale ТЗ text must not override semantics (B5). |
| D7 | Placement applies to **all** filter fields (`data_files.filters` and `reports.articleLevel4Name_clean`, `reports.targetAudience`): filter in place if in Filters/Rows/Columns, else add to Filters; no move. | Per-field placement behavior (B8). |
| D8 | Per `data_file` apply `data_files.filters` once and keep; per `reports` row apply/reapply that row’s `articleLevel4Name_clean` and `targetAudience` overriding prior report-row selection for those fields, without moving placed fields; no leak between report rows. | Stage/reset/isolation per report row (B8,B11,B12). |
| D9 | Multi-value selection split by `general.delimeter`, tokens trimmed, all exactly matching pivot items selected (exact after trim). Applies to `articleLevel4Name_clean` and any list expressions in `filters`; `targetAudience` same exact/trim principle. | Deterministic matching (B5,B8). |
| D10 | Resolve `min`/`max` per `data_files` row against available values of referenced fields in pivot on `data_sheet` of that `data_file` (actual data). Deterministic and dependency-aware: when `Years` present, resolve `Years` bound first, then resolve dependent period fields within that resolved `Years` context (e.g. `Month Number`/`Week Number` min/max in resolved year). Ranges **inclusive** both ends; comparisons **numeric**. | Core resolution algorithm and validation evidence (B5,B8,B13). |
| D11 | `max_year` = resolved max of `Years`; `max_weeknum` = resolved max of `Week Number`; `max_monthnum` = resolved max of `Month Number` for that `data_files` row in same resolved context. Substitution target **`report_sheet` only**. Token syntax `<name>` as in workbook, substituted as resolved values **without zero-padding**. Referenced but underivable ⇒ **fail fast**. | Placeholder derivation/substitution and naming (B5,B11,B12). |
| D12 | **Fail fast** with clear contextual error for missing/invalid config or paths, missing/unresolvable data file, missing `data_sheet`, missing pivot, missing required filter field, invalid/unparseable/unresolvable `filters`, or referenced-but-underivable placeholder. Filter yielding **no data** ⇒ **fail fast** with context. | Unified error taxonomy and outcomes (B4–B13). |
| D13 | Blanks = default/no-op for `freeze_row/freeze_col/zoom/paste_row/paste_col/group_rows`. `group_rows` has no approved deterministic rule; **`group_rows=TRUE` must fail fast** until explicit deterministic rule approved. Blank/false = no-op. | Validation/formatting policy and R4 remediation (B1,B4,B10,B12). |
| D14 | Sanitize Excel-invalid sheet-name characters, enforce **31-character** limit, **fail fast** on sheet-name or report-file-name collision after substitution/sanitization. | Naming/sanitization/collision handling (B11,B12). |
| D15 | Copy pivot presented data range incl. row/column labels and grand totals as shown after filtering; exclude pivot object/page-filter area and PQ connections; paste static **values + formats** (+ column widths if needed). | Extraction/paste fidelity (B9,B13). |
| D16 | Excel/VBA solution in `treport.xlsm`; behavior aligned to revised ТЗ intent reconciled with current workbook structure; `Refresh_data` and `Create_report` in scope. | Solution boundary and reconciliation (all blocks). |
| D17 | `new_file_delimeter` and `keep_col_hidden` dropped (absent rev. 2) and must not be required/processed. | Config contract/validation (B4). |

### Open/product-validation gates

| ID | Status | Recommended path (PO: use recommended options) | Blocks / effect |
|---|---|---|---|
| **OQ-B7 (D18)** | **Pending** (does not block planning) | **Obtain/provide ≥1 golden expected output for a fixed configuration/period** on target Excel/OS with representative `data/theiss` dataset (and document the fixed config/period, environment, and input state). Track as pending baseline. | **Required for correctness sign-off** (values/formatting/layout/copy/fidelity). Structural checks alone are insufficient. Can start acquisition/capture now (B3,B13). |
| **R4 / D13 config conflict** | **Known, must be resolved** | **Set `group_rows` to blank/false for `data_files` row `report_m` in `data/theiss/treport.xlsm` as part of implementation** to comply with D13 (recommended), unless/until an explicit deterministic rule is approved (then update spec/policy before relying on it). | Current config will fail fast under D13-compliant behavior. Resolve before acceptance; can be scheduled early (B1) with backup/safe edit of local artifact. |
| **VBA source missing** | **Known maintainability gap** | **Export/version VBA modules for maintainability where feasible** as part of implementation (recommended), but **do not assume behavior** from absence/presence of exports. | Supports OQ-T4/B2 reconciliation and future maintenance; opportunistic/non-blocking except where inspection access requires it. |

### Technical validation gates (pending evidence; do not choose final approach)

| ID | Uncertainty | Planning treatment | Effect |
|---|---|---|---|
| **OQ-T1** | Actual pivot(s), field captions/types/orientations, PQ/connection topology in `report_m/report_w/report_wclip` and real inputs (incl. `Years` availability in all files, `Week Number`, `Month Number`, `articleLevel4Name_clean`, `targetAudience`), presented ranges. | Auditor investigation required before finalizing filter/copy behavior. Data-source details for “actual data”/available values (D10) must be validated; algorithmic resolution rules are fixed by spec. **Pivot identification/filtering strategy not chosen** pending evidence. | Gates correctness-sensitive implementation details in B5,B8,B9. |
| **OQ-T2** | Synchronous PQ refresh + pivot cache refresh completeness on target Excel/build/query types. | Validate on target environment; **exact PQ refresh completion technique not chosen** pending evidence. | Gates B7 finalization/validation. |
| **OQ-T3** | Copy-range boundaries and formatting fidelity against golden output (labels/totals/page area, widths/heights/hidden/outline, value/date fidelity). | Golden-based verification; **pivot/copy strategy not chosen** pending evidence. | Gates B9 fidelity/sign-off. |
| **OQ-T4** | Reconcile `xl/vbaProject.bin` behavior with rev. 3 spec (binary opaque, no exports). | Inspect/reconcile in Excel/VBA IDE or approved extraction; **VBA architecture/refactor approach not chosen** pending evidence. | May constrain any block; informs maintainability/export action. |

## Local artifact and gitignore constraints (must be accounted for)

- `data/theiss/*` and `data/theiss/treport.xlsm` are **gitignored local artifacts** (not repo-tracked). Plan, evidence, and validation must not assume repository availability or CI visibility of these files.
- Baseline capture, R4 config remediation, VBA export (if performed), golden-output acquisition, and test runs must be handled explicitly in the target/local environment and recorded reproducibly (paths, fixed config/period, environment, file hashes/timestamps where useful) without committing ignored data.
- `mko_treport/src/{scripts,vbs,xls}` are currently empty; any future implementation placement must follow project conventions and workspace-cleanliness rules when executed (not part of this plan edit).
- No `data/theiss/output/` currently exists; outputs are created per D3/FR-FILE-5 during runs and remain local unless explicitly managed.

## Agent role key

- **Auditor:** deeper code/architecture/artifact investigation where uncertainty or complexity exists.
- **Researcher:** modern/best-practice and viable-approach/platform-behavior guidance where multiple approaches or Excel/PQ/pivot behavior matter.
- **Planner:** detailed pre-implementation design, sequencing, and testing/verification shaping.
- **Validator:** independent plan/review/verification gate when implementation or correctness risk is high.
- **High-risk rule:** blocks marked **High risk** require **all four (Auditor, Researcher, Planner, Validator)**.

## Logical execution blocks (B0–B14)

### B0. Baseline evidence & validation setup
- **Objective:** establish inspectable baseline aligned to rev. 3: confirm local artifact inventory (`treport.xlsm`, `base/report_wclip.xlsx`, `monthly/report_m.xlsx`, `weekly/report_w.xlsx`, `docs/description.docx`), record gitignored/local-only status, capture target-environment prerequisites relevant to OQ-B7/OQ-T2 (Windows/Excel), and define reproducible evidence collection for OQ-T1–T4 and golden baseline acquisition without modifying repo-tracked/audit-phase files.
- **Status:** **Can start now.**
- **Depends on:** —
- **Key risks:** Impl: medium (access/tooling limits to inspect binary workbook/PQ internals). Rollout: low. Regression: low. Compatibility: high (PQ/macro-trust, Excel build differences, missing golden). Evidence drift if local artifacts change without recording fixed state.
- **Alternatives / paths (not chosen):** in-Excel/VBA IDE inspection and/or approved module/metadata export vs package/XML inspection only; selection depends on available tooling/access and must not assume behavior.
- **Required agents:** **Auditor, Planner, Validator**
- **Gate / exit criteria:** evidence inventory and validation checklist ready (artifact locations/local-only status, fixed config/period candidates, environment capture template, OQ-T1–T4/B7 blockers and owners/steps); no assumptions recorded as facts.

### B1. Spec alignment, config compliance & maintainability actions
- **Objective:** align execution baseline to rev. 3 decisions D1–D17; confirm config-contract implications (D17 dropped params not required); schedule/perform **R4 remediation (recommended): set `group_rows` to blank/false for `data_files` row `report_m` in `data/theiss/treport.xlsm`** to comply with D13 (unless/until deterministic rule approved); and **export/version VBA modules for maintainability where feasible** (recommended) without assuming behavior.
- **Status:** **Can start now** (alignment/planning immediately; local workbook edit and VBA export are implementation/maintainability actions to perform when safe access is available, with backup of the gitignored artifact).
- **Depends on:** B0 (artifact identification/state capture).
- **Key risks:** Impl: low-medium (binary workbook edit/export access). Rollout: medium if R4 left unresolved (D13-compliant runs will fail fast). Regression: low-medium (unintended config change) — mitigate with backup, minimal targeted edit, and post-edit verification of ListObject values. Compatibility: low.
- **Alternatives / paths (not chosen for R4):** approve and specify a deterministic `group_rows` rule and update spec/policy instead of clearing the flag — **not chosen now** (recommended path is config compliance with D13). VBA architecture/refactor approach **not chosen**; export is for inspectability/maintenance only.
- **Required agents:** **Planner, Validator, Auditor**
- **Gate / exit criteria:** rev. 3 contract baseline acknowledged in work items; R4 remediation completed and recorded (or explicitly tracked with owner before acceptance) so config is D13-compliant; any VBA export captured/versioned where feasible with note that behavior remains subject to OQ-T4 reconciliation.

### B2. Technical investigation — pivots, fields, PQ, VBA (OQ-T1, OQ-T4)
- **Objective:** determine/confirm actual pivot object(s) on `data_sheet` (`clean_data_full`) per data file, field captions/types/orientations and availability for `Years`, `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience`; presented data ranges/totals/page-filter areas; PQ/connection/query topology and refresh dependencies; control bindings/entry points (observed single `RunXlconcat`); and reconcile observable `xl/vbaProject.bin` behavior with rev. 3 where possible.
- **Status:** **Can start now** (investigation). **High risk.**
- **Depends on:** B0.
- **Key risks:** Impl: high (opaque VBA, OLAP/Data Model possibilities, package evidence incomplete). Rollout: high. Regression: high (wrong pivot/field/range assumptions). Compatibility: high (Excel object model/PQ/field naming differences across files/builds).
- **Alternatives / paths (not chosen):** Excel UI/VBA IDE/manual field audit vs package/XML/connection-part inspection vs approved module export/decompilation aids — evidence path chosen per access; **pivot identification/filtering/copy strategy and VBA architecture/refactor approach not chosen** pending findings.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** semantic map recorded (per data file: target pivot, field identity/captions/types/orientations, `Years` presence/values context, range/totals characteristics, PQ/connection dependencies, control/entry behavior) or explicit unknowns with impact/risk; no implementation behavior assumed.

### B3. Golden baseline & target-environment confirmation (OQ-B7, prereqs for OQ-T2/T3)
- **Objective:** confirm/document target Excel version & OS and representative `data/theiss` input state; **obtain/provide ≥1 golden expected output for a fixed configuration/period** (recommended, OQ-B7/D18) including the exact resolved period/filters/config used; establish reproducible capture conditions for refresh/copy-fidelity validation.
- **Status:** **Can start now** (coordination, environment capture, baseline preparation). **Correctness sign-off blocked** until golden baseline is available and accepted.
- **Depends on:** B0; informed by B1,B2 where possible.
- **Key risks:** Impl: medium. Rollout: high (no correctness oracle). Regression: high. Compatibility: high (Excel build/PQ behavior). Golden mismatch if fixed config/period/input state not frozen.
- **Alternatives / paths (not chosen):** PO-provided approved output vs collaboratively captured run once behavior/evidence validated — baseline must be explicitly designated as authoritative for the fixed case; not chosen here.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** target environment documented; fixed configuration/period and representative dataset state recorded; golden output(s) available and traceable, or explicitly tracked as blocking correctness sign-off (B13).

### B4. Config contract & validation (FR-CFG, T1)
- **Objective:** implement deterministic config reading/validation for ListObjects `general`, `reports`, `data_files`, `placeholders` per rev. 3 schema: required/non-empty fields, defaults/blanks/no-op semantics, D17 dropped params not required/processed, D13 `group_rows=TRUE` fail-fast until rule approved, and fail-fast contextual errors (D12) for invalid/missing/unresolvable config.
- **Status:** **Can start now** (design/implementation of schema and validation rules independent of runtime pivot internals).
- **Depends on:** B1 (spec alignment, R4 awareness), B2 (observed structure/control nuances as available).
- **Key risks:** Impl: medium. Rollout: medium. Regression: medium (over/under-validation). Compatibility: low-medium. False assumptions about optionality/types if not tied to spec.
- **Alternatives / paths (not chosen):** validation strictness/diagnostic granularity details may vary, but must remain deterministic and spec-compliant (no silent undefined behavior, no required dropped params).
- **Required agents:** **Planner, Validator** (+ **Auditor** if structure ambiguities remain)
- **Gate / exit criteria:** config schema/validation matrix implemented or specified with deterministic outcomes/errors covering FR-CFG-1..6, D12,D13,D17; `group_rows=TRUE` rejected per D13; R4-compliant config accepted.

### B5. Filters DSL parsing & dependency-aware resolution (FR-FLT-1/2/3, FR-NAME-1 inputs, T2)
- **Objective:** implement deterministic `filters` parsing and resolution per `data_files` row: DSL grammar/disambiguation (D5,D6), multi-value/list handling via `general.delimeter`, dependency-aware `min/max` resolution (D10: resolve `Years` bound first if present, then dependent period fields within resolved `Years` context; numeric inclusive comparisons; against available values of referenced fields in pivot on `data_sheet` of that `data_file`), and placeholder derivation/substitution inputs (D11: `max_year/max_weeknum/max_monthnum` as resolved maxima in same context; substitution in `report_sheet` only as `<name>` without zero-padding; referenced but underivable ⇒ fail fast).
- **Status:** **Can start now** (parser/resolver design and implementation of spec-fixed algorithm and error handling). **High risk** for correctness of “available values” sourcing until OQ-T1 evidence.
- **Depends on:** B4, B1; informed/validated by B2 (field identity/types/value domains, OQ-T1).
- **Key risks:** Impl: high (grammar edge cases, dependency ordering, numeric coercion/comparison, context scoping). Rollout: medium-high. Regression: high (wrong period/filter/placeholder resolution). Compatibility: medium-high (pivot/cache item representation).
- **Alternatives / paths (not chosen):** concrete source of “available values” in pivot data and extraction mechanics (cache/items/rendered/domain APIs) — **not chosen** pending OQ-T1; must satisfy D10 semantics when evidenced. No deviation from D5/D6/D10/D11 rules.
- **Required agents:** **Planner, Validator, Auditor, Researcher** (all, due to algorithmic/platform risk)
- **Gate / exit criteria:** deterministic parse/resolve/derive behavior specified/implemented for D5,D6,D10,D11 with contextual fail-fast errors (D12); resolved values/logging available for verification; value-source assumptions explicitly tied to OQ-T1 evidence.

### B6. File resolution & access/session policy (FR-FILE, T3)
- **Objective:** implement deterministic data-file resolution in `input_path` per D2 (`<data_file>.<ext>`, precedence `.xlsx,.xlsm,.xlsb`, fail fast if none or >1) and mode-specific access/session behavior per D1 (`Create_report` read-only never save; `Refresh_data` writable refresh+save+close), including `output_path` readiness (FR-FILE-5) and state-safe cleanup on success/failure.
- **Status:** **Can start now**.
- **Depends on:** B4, B1, B2 (path/lock realities as observed).
- **Key risks:** Impl: high. Rollout: high. Regression: high (accidental source mutation, locks/leaks, wrong file selected). Compatibility: medium-high (Windows paths/locking/permissions).
- **Alternatives / paths (not chosen):** session/workbook lifecycle patterns (single-app vs per-file isolation, centralized cleanup/finally structure) and lock-retry posture — multiple viable within Excel/VBA constraints; **not chosen** as a final architecture beyond deterministic, spec-compliant behavior.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** deterministic D2 resolution and D1 access rules enforced with contextual fail-fast errors (D12); no source save in `Create_report`; cleanup/state restoration behavior defined/testable.

### B7. Refresh workflow (FR-REF, T4; OQ-T2)
- **Objective:** implement reliable `Refresh_data` flow: for each resolved data file in deterministic order open writable, refresh Power Query and dependent pivot data so persisted results are current, save, and close; fail fast with contextual errors without corrupting state.
- **Status:** **Partially startable** (design, sequencing, error handling, and instrumentation can start now; completion/validation **gated by OQ-T2** evidence on target Excel). **High risk.**
- **Depends on:** B6, B2, B3 (target environment/prereqs).
- **Key risks:** Impl: high. Rollout: high. Regression: high (stale persisted data). Compatibility: high (async/background queries, PQ/connection types, Excel build differences).
- **Alternatives / paths (not chosen):** **exact PQ refresh completion technique** (e.g. query/background-refresh synchronization strategies, pivot cache refresh ordering/waits) — platform-dependent, **not chosen** pending OQ-T2 evidence.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** evidenced refresh-completion criteria satisfied on target environment (OQ-T2); no stale-save path accepted; failures reported with file/stage context (D12).

### B8. Filter application, staging & placement (FR-FLT-4/5/6/7, T5; D4,D7,D8,D9,D12; OQ-T1)
- **Objective:** apply resolved filters deterministically: per `data_file` apply `data_files.filters` once and keep (D4,D8); per `reports` row apply/reapply `articleLevel4Name_clean` and `targetAudience` overriding prior report-row selection for those fields with no leak (D8); enforce placement for all filter fields in-place if in Filters/Rows/Columns else add to Filters with no move (D7); exact-after-trim matching and multi-value selection via `general.delimeter` (D9,D5); fail fast on missing sheet/pivot/field, unresolvable expression/value, or no-data outcomes (D12).
- **Status:** **Partially startable** (spec-aligned design/state machine and non-pivot-specific logic can start now; pivot-specific behavior **gated by OQ-T1**). **High risk.**
- **Depends on:** B5, B4, B6, B2 (OQ-T1), B7 context where refresh state matters.
- **Key risks:** Impl: high. Rollout: high. Regression: high (incorrect filtering, cross-report leakage, stale state). Compatibility: high (pivot orientation, regular vs OLAP/Data Model, field item availability/types).
- **Alternatives / paths (not chosen):** **pivot identification and filter-application strategy** (object-model APIs/mechanics across orientations/OLAP, manual-update batching, item selection methods) — multiple viable, **not chosen** pending OQ-T1 evidence.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** field-by-field, stage-correct, placement-correct behavior confirmed against OQ-T1 findings; per-report isolation/no-leak and fail-fast no-data behavior verified.

### B9. Pivot snapshot extraction & static paste (FR-RPT-3/4, T6; D15; OQ-T3)
- **Objective:** capture the pivot’s presented data range after all applicable filters (including row/column labels and grand totals as shown; excluding pivot object/page-filter area and PQ connections) and paste at `paste_row/paste_col` as static **values + formats** (+ column widths if needed) with no pivot objects or PQ connections in output.
- **Status:** **Partially startable** (design and scaffolding can start now; bounds/fidelity finalization **gated by OQ-T1/OQ-T3 and golden baseline**). **High risk.**
- **Depends on:** B8, B2, B3 (golden/fidelity reference).
- **Key risks:** Impl: high. Rollout: medium-high. Regression: high (wrong/misaligned range, lost totals/labels, fidelity loss). Compatibility: medium-high (layout modes, hidden/outline/size behavior, value/date formatting).
- **Alternatives / paths (not chosen):** **copy strategy / range-detection approach** (API-provided ranges vs rendered/bounded detection) and fidelity trade-offs — **not chosen** pending OQ-T1/T3 evidence and golden comparison.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** copy-range and paste contract matches D15 and is validated against golden output (OQ-T3/OQ-B7) on representative data.

### B10. Formatting & layout (FR-RPT-5/6, T7; D13)
- **Objective:** apply deterministic formatting per `data_files`: freeze panes (`freeze_row/freeze_col`), zoom, hide gridlines, and row/column sizing/state as required; blanks as no-op/default. Enforce D13: `group_rows=TRUE` must fail fast until explicit deterministic rule approved; blank/false no-op.
- **Status:** **Can start now** (policy and implementation for defined settings), with visual/fidelity checks gated by OQ-T3/golden.
- **Depends on:** B4, B9 (target sheet/content context), B1 (R4 compliance).
- **Key risks:** Impl: medium. Rollout: medium. Regression: medium (visual/structural drift). Compatibility: medium (Excel rendering/outline/conditional behavior).
- **Alternatives / paths (not chosen):** any color-band/outline grouping for `group_rows` is out of scope unless/until a deterministic rule is approved and reflected in spec/policy — **not chosen**.
- **Required agents:** **Planner, Validator, Auditor** (+ **Researcher** if platform-specific outline/sizing constraints require investigation)
- **Gate / exit criteria:** formatting acceptance criteria met for configured cases; no undefined `group_rows` behavior; D13 fail-fast preserved for `TRUE`.

### B11. Report assembly, naming, placeholders & idempotency (FR-RPT-1/2/7, FR-NAME, T8; D3,D11,D14)
- **Objective:** create one output workbook per `reports` row (`reports.name` + `.xlsx`, overwrite on successful run), create/replace target sheets per `data_files` order with `report_sheet` names after placeholder substitution in `report_sheet` only (D11), sanitize Excel-invalid characters, enforce 31-char limit, fail fast on collisions (D14), ensure deterministic idempotent reruns (D3), save to `output_path`, and release resources per D1.
- **Status:** **Can start now**.
- **Depends on:** B5 (resolved placeholders), B6, B9, B10, B4.
- **Key risks:** Impl: high. Rollout: high. Regression: high (overwrites, collisions, ordering, naming drift, incomplete cleanup). Compatibility: medium (filesystem locks, Excel name limits/invalid chars).
- **Alternatives / paths (not chosen):** internal sheet-recreation strategy for idempotency (delete/recreate vs replace) — multiple viable if deterministic and spec-compliant; **not chosen** as a mandated micro-approach.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** end-to-end output lifecycle deterministic and idempotent per D3/D11/D14; correct names/order/paths; no pivot/PQ artifacts in outputs (with B9).

### B12. Diagnostics, run control & fail-fast error taxonomy (FR-RUN, T9)
- **Objective:** provide explicit success/failure outcomes and fail-fast, contextual errors (operation, config row/name, file path, sheet, field, filter/expression, value as applicable) across config, resolution, refresh, filtering, extraction, naming, and save; ensure run controls/state-safe execution and cleanup without unnecessary UI disruption.
- **Status:** **Can start now** (incrementally alongside other blocks) and must consolidate before sign-off.
- **Depends on:** B4–B11 (cross-cutting consolidation).
- **Key risks:** Impl: medium-high. Rollout: medium. Regression: medium (missing context, swallowed errors, leaked state). Compatibility: medium (Excel UI/application state behavior).
- **Alternatives / paths (not chosen):** messaging/surfacing mechanism within Excel-macro constraints — keep minimal, deterministic, and consistent; no speculative UX redesign. Control-binding reconciliation (e.g. observed `RunXlconcat` vs named `Refresh_data`/`Create_report`) depends on B2/OQ-T4 findings and must not be assumed.
- **Required agents:** **Planner, Validator, Auditor** (+ **Researcher** if Excel control/event constraints require guidance)
- **Gate / exit criteria:** unified error taxonomy/messages mapped to D12/D13/D14 and FR-RUN outcomes; state restored on completion/failure; failures identify actionable context.

### B13. Verification & acceptance (T10; D18/OQ-B7, OQ-T1–T4)
- **Objective:** define and execute verification scenarios and acceptance checks: config validation (incl. D13/D17), D2 resolution and D1 access modes, refresh completeness (OQ-T2), `filters` parse/disambiguation and dependency-aware `min/max`/`max_*` resolution (D5,D6,D10,D11), staging/placement/matching/no-leak/no-data (D7,D8,D9,D12), copy/format fidelity (D15,OQ-T3), naming/sanitization/collision/idempotency (D3,D14), and golden comparison for fixed configuration/period (OQ-B7).
- **Status:** **Can start now** (scenario matrix, test design, structural checks, instrumentation). **Correctness sign-off gated** by OQ-B7 golden baseline and OQ-T1–T4 evidence/closure.
- **Depends on:** B3, B4–B12 (implemented/available behaviors), B2 evidence.
- **Key risks:** Impl: medium. Rollout: high. Regression: high. Compatibility: high. False confidence from structural-only checks without golden parity.
- **Alternatives / paths (not chosen):** structural validation vs golden-value/format parity — structural checks are necessary but insufficient for correctness sign-off (D18/Q12).
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** executable verification evidence recorded; pass/fail per scenario; OQ-T1–T4 tracked/closed; golden-based correctness sign-off achieved for fixed configuration/period or explicitly pending with blockers.

### B14. Independent review & release-readiness gate
- **Objective:** consolidated independent review that implementation sequencing and delivered behavior satisfy rev. 3 spec (§§4–7,14), preserve D1–D17, respect R4/D13 compliance, do not rely on unchosen approaches beyond evidenced constraints, and have required validation/sign-off artifacts (or explicit accepted limitations).
- **Status:** **Gated** — perform when implementation and verification reach reviewable completeness (after B4–B13 progress as applicable).
- **Depends on:** B1–B13 outputs/evidence and current spec baseline.
- **Key risks:** Impl: low. Rollout: high. Regression: high (missed dependency/gate or unsupported assumption). Compatibility: medium.
- **Alternatives / paths:** N/A (review gate).
- **Required agents:** **Validator** (lead), with **Planner, Auditor, Researcher** as needed for contested/high-risk findings
- **Gate / exit criteria:** dependency-safe readiness confirmed for release/sign-off or clearly bounded by open OQ-B7/OQ-T1–T4 items; no unresolved contradiction with D1–D17/R4 policy.

## Execution order and dependency DAG

### Textual DAG
1. **B0** (baseline evidence & setup) — **start now**.
2. **B1** (spec alignment, R4 compliance, VBA export where feasible) — **start now**, depends on B0.
3. **B2** (technical investigation OQ-T1,OQ-T4) — **start now**, depends on B0. *(high risk)*
4. **B3** (golden baseline & target environment OQ-B7; prereqs OQ-T2/T3) — **start now**, depends on B0, informed by B1,B2. *(sign-off gated)*
5. **B4** (config contract & validation, T1) — **start now**, depends on B1,B2.
6. **B5** (filters DSL parse & dependency-aware resolution, T2) — **start now**, depends on B4,B1; validated by B2/OQ-T1. *(algorithm fixed; data-source details evidence-gated)*
7. **B6** (file resolution & access/session, T3) — **start now**, depends on B4,B1,B2.
8. **B7** (refresh workflow, T4) — **partially startable**, depends on B6,B2,B3; **gated by OQ-T2**. *(high risk)*
9. **B8** (filter application/staging/placement, T5) — **partially startable**, depends on B5,B4,B6,B2 (and refresh context as relevant); **gated by OQ-T1** for pivot-specific behavior. *(high risk)*
10. **B9** (pivot snapshot extraction & paste, T6) — **partially startable**, depends on B8,B2,B3; **gated by OQ-T1/OQ-T3/golden**. *(high risk)*
11. **B10** (formatting & layout, T7) — **start now** (after inputs), depends on B4,B9,B1.
12. **B11** (report assembly/naming/placeholders/idempotency, T8) — **start now**, depends on B5,B6,B9,B10,B4.
13. **B12** (diagnostics/run control/error taxonomy, T9) — **start now** incrementally, consolidates B4–B11.
14. **B13** (verification & acceptance, T10) — **start now** for planning/structural checks; **sign-off gated** by OQ-B7,OQ-T1–T4, depends on B3,B4–B12,B2.
15. **B14** (independent review & release-readiness) — **gated**, depends on B1–B13.

### Safe parallelization opportunities (when prerequisites met)
- **B1 ∥ B2 ∥ B3** after/alongside B0 (coordination-heavy but dependency-compatible).
- **B4** can proceed once B1 alignment is clear, incorporating B2 observations as they arrive.
- **B5 ∥ B6** after B4 (different concerns: resolution logic vs file/session handling), with B5 validated against B2 findings.
- **B7** design and **B5/B8/B11** non-pivot-specific design can progress in parallel where interfaces are clear; B7 completion remains OQ-T2-gated.
- **B10, B11, B12** can advance incrementally with B9 as extraction contracts stabilize.
- **B13** scenario design runs early and continuously; execution evidence accrues as gates close.

## Cross-cutting risks and containment

| Risk / area | Impact | Containment / gate |
|---|---|---|
| **VBA opacity & unverified runtime behavior** (R7, OQ-T4) | Hidden logic diverges from rev. 3; wrong assumptions in implementation | B2 investigation/reconciliation; optional VBA export for maintainability (recommended) without assuming behavior; no final VBA/refactor approach chosen prematurely. |
| **Pivot topology, field identity & `Years` availability** (R1, R3, OQ-T1) | Wrong target pivot/field, incorrect `min/max`/`max_*` or filter application | B2 evidence before finalizing pivot-specific behavior; D10 algorithm fixed but value-source mechanics validated; fail fast on missing/ambiguous fields (D12). |
| **PQ refresh completeness/timing** (R2, OQ-T2) | Stale persisted data after `Refresh_data` | B7 validation on target Excel; exact completion technique not chosen until evidenced; no stale-save acceptance. |
| **Copy-range/format fidelity** (R5, OQ-T3) | Output differs from intended presented data/layout | B9/B13 golden-based validation; copy strategy not chosen until evidenced. |
| **`group_rows` policy/config conflict** (R4, D13) | Current config (`report_m` TRUE) fails under compliant behavior | **Recommended remediation in B1:** set `group_rows` blank/false for `report_m` in local `treport.xlsm`; enforce D13 fail-fast for TRUE until rule approved. |
| **No-data and error handling gaps** (D12) | Silent wrong outputs or unclear failures | Fail-fast contextual errors for all specified cases incl. no-data; unified taxonomy in B12. |
| **Naming/collision and idempotency defects** (D3, D14) | Overwrites, collisions, unstable reruns | Deterministic sanitize/31-char/fail-fast collisions, ordered recreation/replacement, overwrite-on-success policy. |
| **Locked files/permissions/environment variance** (R6, R9) | Run failures, partial outputs, non-reproducible behavior | D1/D12 fail-fast with context; target env documented (B3); cleanup/state restoration. |
| **Performance on large pivots/many items** (R8) | Long/unresponsive runs | Benchmark with representative data during B13; optimize only with evidence and without changing semantics (technical concern, not in this plan’s chosen approach). |
| **Local-only gitignored artifacts** (`data/theiss/*`, `treport.xlsm`) | Non-reproducible evidence, accidental omission from validation, or unnoticed edits | Explicit inventory/capture in B0/B3/B13; record fixed config/period/input/golden state; treat R4/VBA-export edits as controlled local actions with backup/traceability. |
| **Spec/observed artifact discrepancies** (FR-RUN-1 control naming, stale ТЗ text, trailing space in report name) | Confusion or incorrect assumptions | Reconcile via B2/B4/B12 against rev. 3 and observed workbook; preserve configured values exactly where required; follow D6 for stale duplicate example. |

## Definition of Ready alignment (spec §14)

| §14 item | Status | Notes |
|---|---|---|
| **1. Business decisions/clarifications closed (D1–D17)** | **Satisfied** | PO selected all recommended options; OQ-1a–d, OQ-2a–c, OQ-3a–b, OQ-4, OQ-B1, OQ-B3–B6 resolved in rev. 3. |
| **2. Scope & requirements stable (§§4–7)** | **Satisfied** | Includes `filters` grammar (D5), dependency-aware `min/max` and `max_*` derivation (D10,D11), placement/staging (D7,D8), deterministic policies (D2,D12–D14), and related FRs. |
| **3. Deterministic policies fixed** | **Satisfied** | Extension precedence, `min/max` resolution algorithm, placeholder mapping/scope/syntax/no-padding, no-data fail-fast, collision fail-fast, `group_rows` fail-fast while undefined. |
| **4. Validation baseline available (OQ-B7)** | **Pending (does not block planning)** | Target Excel/OS + representative dataset + ≥1 golden expected output for fixed configuration/period required for correctness sign-off; **recommended path:** obtain/provide golden baseline (B3/B13). |
| **5. Technical validations tracked (§13.2/§10)** | **Tracked, pending closure** | OQ-T1–T4 treated as implementation/verification gates (B2,B7,B8,B9,B13). |

- **Planning gate:** **items 1–3 completed (satisfied).** Item 4 pending but does not block planning; **items 4–5 must close before correctness sign-off**.
- **Implementation start:** **allowed now** where block prerequisites/gates permit (see B0–B13 statuses). Gated technical choices must remain open until OQ-T1–T4 evidence.
- **Acceptance note (R4):** current config contains `group_rows=TRUE` for `report_m` (conflicts D13). **Must be addressed as part of implementation/acceptance** — recommended: set to blank/false in `data/theiss/treport.xlsm` (B1), unless/until a deterministic rule is approved and reflected in spec/policy.

## Expected deliverables (when blocks execute)

- Evidence records for B0/B2/B3 (artifact/topology/PQ/control findings, target environment, fixed config/period, golden baseline or tracked gap).
- D13-compliant local config state for R4 (or approved-rule alternative with spec/policy update) and, where feasible, versioned VBA exports for maintainability (behavior not assumed).
- Dependency-safe implementation increments covering B4–B12 semantics (semantic anchors, deterministic behavior, fail-fast diagnostics), with approaches left open only where OQ-T1–T4 require evidence.
- Verification artifacts and sign-off package for B13/B14: scenario results, OQ-T1–T4 closure evidence, and golden-based correctness confirmation for fixed configuration/period (or explicit pending blockers).

---

**Brief notes (plan handoff):**
- **Recommended filename remains:** `.ai/plans/001-treport-execution-plan.md` (this is the complete corrected replacement plan).
- **Local/gitignored artifacts:** `data/theiss/*` and `data/theiss/treport.xlsm` are gitignored and local-only; the plan explicitly accounts for controlled local evidence capture, R4 remediation in `treport.xlsm`, optional VBA export, golden baseline storage/traceability, and reproducible validation without assuming repo-tracked availability.
- **Recommended options applied (per PO):** OQ-B7 tracked with recommended path to obtain/provide golden expected output for fixed configuration/period; R4 resolved via recommended config change (`report_m.group_rows` blank/false) in implementation/acceptance unless a deterministic rule is later approved; VBA modules to be exported/versioned for maintainability where feasible without assuming behavior; file resolution follows D2 (`.xlsx,.xlsm,.xlsb` precedence). No final implementation approach chosen where OQ-T1–T4 technical uncertainty remains.

## B2. Technical investigation — findings/status (OQ-T1, OQ-T4) — updated 2026-11-14 (exported VBA reconciled)

- **Status:** Investigation deepened (read-only). Exported VBA source in `mko_treport/src/vbs/export/{xlconcat_macro.bas,ThisWorkbook.cls,Sheet1.cls,Sheet2.cls}` analyzed and reconciled with package/XML evidence from `.ai/tmp/b2_extract/{treport_xlsm,report_m_xlsx,report_w_xlsx,report_wclip_xlsx}` and local gitignored artifacts `data/theiss/{treport.xlsm,monthly/report_m.xlsx,weekly/report_w.xlsx,base/report_wclip.xlsx}`. No implementation behavior assumed beyond exported source/XML facts; no final pivot/filter/copy/refresh/VBA approach chosen.

- **Evidence base (VBA):** `xlconcat_macro.bas` (module `xlconcat_macro`, 1455 lines, `APP_NAME="xlconcat"`). Entry point `Public Sub RunXlconcat()` assigned to `treport.xlsm` `run` sheet (`Sheet2`) control `Button 1` / caption "Run Macro" (`macro="[0]!RunXlconcat"`, `xl/drawings/drawing1.xml`, `xl/drawings/vmlDrawing1.vml`, `xl/ctrlProps/ctrlProp1.xml`, `xl/worksheets/sheet1.xml`). `ThisWorkbook.cls`, `Sheet1.cls` (`settings`, code name Sheet1), `Sheet2.cls` (`run`, code name Sheet2) contain no procedures. `vbaProject.bin` ASCII scan: HIT `RunXlconcat,run,settings,delimeter,ReadOnly,Save,output,input_path`; MISS `Refresh_data,Create_report,general,reports,data_files,placeholders,articleLevel4Name_clean,targetAudience,filters,max_*,PivotTable,PivotCache,RefreshAll,Power,Query,clean_data_full,report_m/w/wclip`.

- **Confirmed VBA architecture/behavior (`xlconcat_macro`):** generic workbook/worksheet concatenator, not treport rev.3 reporting. Reads config from sheet `settings`, `ListObject("settings")` with `parameter,value` columns via `GetSettingsTable`/`FindSettingRow`/`GetSettingString`/`GetSettingLong`/`GetSettingBool` (params: `input_path`,`output_path`,`sheets_interval`,`workbook_interval`,`ignore_col_mismatch`,`new_file_delimeter`,`keep_col_hidden`; defaults/validation as coded). Reads optional `ListObject("exclude_names")` via `ReadExcludeNames`/`CleanName`. Resolves input as folder (`NormalizeFolderPath`, `FileExists` via FSO) and output as file path (`NormalizeFilePath`, `GetFileFormat`); fails fast if input missing/empty, output empty, intervals negative, output equals macro workbook (`PathsEqual`), output file already exists. Collects `.xlsx`/`.xls` in input folder alphabetically (`CollectInputFiles`/`SortCollection`, excludes `~$` and macro workbook) — no per-`data_files` base-name/extension precedence (D2) and no `input_path`/`output_path` semantics from `general` table. Opens every source `Workbooks.Open(..., ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)`; never refreshes PQ/pivots; never saves source files. Creates one output workbook/sheet `plans`, appends visible non-excluded sheets sequentially using `GetCopyRange` (UsedRange or `Find` last cell when rows > `LARGE_USERANGE_THRESHOLD`), `PasteSpecial xlPasteValuesAndNumberFormats` + `xlPasteFormats`, internal-formula handling (`CopyFormulasWithoutExternalLinks`/`HasExternalReference`/`ShiftFormulaR1C1`), row/column/zoom/layout copying (`CopyHiddenRows`,`CopyRowHeights`,`CopyColumnWidths`,`CopyZoom`,`CopyColumnHidden`), optional file delimiter row, column-count and Excel-row-limit checks; saves/closes output (`SaveAs`). Error/state handling via saved app settings, `On Error GoTo FatalError`/`CleanUp`, MsgBox diagnostics — no rev.3 contextual taxonomy (operation/config row/file/sheet/field/filter/value), no per-report isolation, no deterministic report lifecycle.

- **Confirmed treport config/artifacts (`treport.xlsm`, decoded shared strings):** `general` A1:B4 `input_path,output_path,delimeter` (`&`, paths formula-resolved in saved state to `C:\py_exp\mko_treport\data\theiss\` and `...\output`); `reports` D1:F7 6 rows (`articleLevel4Name_clean`,`targetAudience`,`name`) including `DOLGIT ` trailing space and category token `ЗУБНАЯ ПАСТА&ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`; `data_files` D10:M13 3 rows (`report_m|clean_data_full|помесячное сравнение|Month Number=(1-max)|Years=max|1,1,blank,blank,0.7,TRUE`, `report_w|clean_data_full|сравн конкур_<max_weeknum>w_<max_year>_прямая|Week Number=(1-max)|Years=max|1,1,2,2,blank,blank`, `report_wclip|clean_data_full|хроно по неделям|Week Number=(1-max)|Years=max|1,1,blank,blank,blank,blank`); `placeholders` A10:B13 `max_weeknum,max_monthnum,max_year`. `exclude_names` ListObject not observed in treport config; no `Refresh_data`/`Create_report` controls.

- **Confirmed data workbook/pivot/PQ evidence (per file, XML):** each of `monthly/report_m.xlsx`, `weekly/report_w.xlsx`, `base/report_wclip.xlsx` has sheet `clean_data_full` and `PivotTable1` on it with external cache `pivotCache/pivotCacheDefinition1.xml`, `cacheSource type="external" connectionId="2"` = `Query - clean_data_full` (Mashup, type=5). `connections.xml` identical in extracted size/content shape across files (17 connections, `background="1"`, includes user/machine-specific `odcFile="C:\Users\NMatveichev\..."` paths). Pivot layouts/field availability differ:
  - `report_m`: location `A8:T426`, page fields `6->researchDate,8->Week Number,7->Start of Week,0->Competitor,1->targetAudience,10->articleLevel4Name_clean`, cache has `Year` (2026 only), `Month Number` (3-10), `Week Number` (irregular set incl. 9,14,18,23,27,31,36,40), `articleLevel4Name_clean`, `targetAudience`; **no `Years` cache field**.
  - `report_w`: location `A7:V77`, page fields `8->Week Number,7->Start of Week,0->Competitor,1->targetAudience,10->articleLevel4Name_clean`, cache has `Year` (2026), `Month Number` (3-10), `Week Number` (14-43), `articleLevel4Name_clean`, `targetAudience`; **no `Years` cache field**.
  - `report_wclip`: location `A8:AP273`, page fields `35->Years,8->researchDate,0->Competitor,10->articleLevel4Name_clean,1->targetAudience,21->adTypeName`, cache has `Year` (2024-2026), `Month Number` (1-12), `Week Number` (1-53), `articleLevel4Name_clean`, `targetAudience`, and grouped `Years` field (base 4, groupBy years, items `<01/01/2024,2024,2025,2026,>02/11/2026`).
  Decoded cache items (wclip sample, UTF-8): `targetAudience` includes `W 50+`,`W 30-50`; `articleLevel4Name_clean` includes `ЗУБНАЯ ПАСТА` but `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as a distinct cache item in inspected extracts (potential token mismatch vs config split by `&` remains a risk).

- **OQ-T4 reconciliation (VBA source now visible):** exported source and control bindings confirm implementation is `xlconcat_macro.RunXlconcat` concatenation workflow, **not** rev.3 treport `Refresh_data`/`Create_report` workflow. Absence in exported source/strings of rev.3 tables/fields/operations and pivot/PQ/report logic is strong evidence of functional divergence, but authoritative runtime behavior of the current `vbaProject.bin` still requires Excel/VBA IDE or approved decompilation only if binary may differ from exports (exports taken as reconciled source snapshot here; runtime parity unverified in this read-only pass).

- **OQ-T1/T2/T3 (partially informed, still open):** pivot/cache topology, field names/captions, page-field placement, `Year` vs `Years` availability, presented locations, and PQ async/background/user-path indicators are evidenced from package XML; runtime pivot object identity/selection semantics, grouped `Years` member mapping/filter behavior, exact presented copy bounds excluding page-filter/pivot-object/totals-as-shown, PQ M definitions/load targets/refresh completion, and formatting/copy fidelity remain unverified and require target Excel + golden baseline.

- **Risks/impacts for rev.3 (confirmed/high):** (1) full functional gap — exported macro does not implement D1–D17/FR-CFG/FILE/REF/FLT/RPT/NAME/RUN; `Refresh_data`/`Create_report` absent. (2) Config/schema mismatch — macro reads `settings` parameter/value list and `exclude_names`, not rev.3 ListObjects/fields/DSL. (3) Access/path/output divergence — always read-only source opens, folder-wide concat, output-must-not-exist, single `plans` sheet output. (4) Field mismatch — `data_files.filters` reference `Years` but `report_m/report_w` caches expose `Year` (singular); `report_wclip` has grouped `Years` page field (mapping/semantics unconfirmed). (5) Possible category token mismatch — config split token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not seen in inspected `articleLevel4Name_clean` cache items. (6) PQ refresh/portability — async/background Mashup connections with user-specific `odcFile` paths. (7) Pivot/layout variability — different locations/page/row/col structures across files affect D7/D8/D15 implementation and verification. (8) R4 config conflict remains (`report_m.group_rows=TRUE`, D13 fail-fast). No implementation approach chosen.


## Execution progress/status (updated 2026-11-14, implementation started)

> Status notes are factual progress records. Scope, decisions (D1–D17), gates (OQ-B7 pending golden/sign-off, OQ-T1–T4 pending evidence/closure as marked), and “no final implementation approach chosen where technical uncertainty exists” remain unchanged.

- **B0 Baseline evidence & validation setup — completed (evidence/checklist).** Auditor produced `{context_a}` with artifact inventory (gitignored/local-only), hashes/paths/state, environment capture template, fixed config/period capture tuple, and OQ-T1–T4/OQ-B7 blocker/owner/step checklist. Key observations: Excel install evidenced (`C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE` present), target Windows/Excel not fully documented for sign-off; `data/theiss/output/` absent; PQ connections/user-specific paths and workbook absPaths noted.
- **B1 Spec alignment, config compliance & maintainability actions — R4 remediated, VBA exports versioned.** 
  - R4 (D13, recommended): `data/theiss/treport.xlsm` (gitignored) updated so `data_files.report_m.group_rows` (M11, `xl/worksheets/sheet2.xml`) changed from `<c r="M11" s="1" t="b"><v>1</v></c>` to `<c r="M11" s="1"/>`; `M12`/`M13` unchanged blank/no-op; only `sheet2.xml` member changed (zip/XML/integrity verified). Distinct pre-change backup created: `data/theiss/treport.xlsm.bak_b1_pre_r4` (byte-identical to pre-change state). Prior identical backup `treport.xlsm.bak_20261114_b1_r4` also present. **Not committed** (gitignored).
  - VBA maintainability (recommended, no behavior assumed): opaque binary snapshot `mko_treport/src/vbs/vbaProject.bin` added/committed (`81ce32e chore: add VBA binary snapshot for B1 R4/D13 maintainability`); readable modules exported read-only via Excel COM from `treport.xlsm` and committed (`a8895aa2bb1260a56f00262676a6f412928fa239 chore(vba): export treport VBA modules for maintainability (B1)`): `mko_treport/src/vbs/export/{Sheet1.cls,Sheet2.cls,ThisWorkbook.cls,xlconcat_macro.bas}`.
- **B2 Technical investigation (OQ-T1, OQ-T4) — deepened with exported VBA source; gates remain open per plan.** Package/XML + VBA-source reconciliation completed (Auditor `{context_a}`; plan appendix retained). Confirmed: single control `RunXlconcat`/`xlconcat_macro.bas` is a legacy concatenation macro (reads `settings` parameter/value incl. `new_file_delimeter/keep_col_hidden/...`, enumerates `.xlsx/.xls`, opens sources read-only, no PQ refresh, no pivot/filtering, no rev.3 `general/reports/data_files/placeholders/filters/max_*` handling, single output workbook/sheet `plans`) — **full functional mismatch** with rev.3 D1–D17/FRs (OQ-T4 substantially reconciled; residual runtime-vs-export parity unverified). Pivot/PQ evidence: `PivotTable1` on `clean_data_full` per data file (external cache, `Query - clean_data_full`, 17 Mashup connections, async/background, user-specific paths); field `Years` cache/grouped page field present only in `report_wclip` while `report_m`/`report_w` expose `Year` (singular) but `data_files.filters` use `Years=max`; report category token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as distinct cache item in inspected extracts. **OQ-T1, OQ-T2, OQ-T3 remain open (evidence/runtime-gated); D10 algorithm fixed by spec but value-source/member semantics unconfirmed.**
- **B3 Golden baseline & target-environment confirmation — partially started, blocked for sign-off.** Known target: Windows 10–11, Excel LTS 2021+, dataset `data/theiss` (PO). Local machine: Office16 Excel binaries present; full target build/macro-trust/PQ details and golden expected output for fixed config/period **not yet provided** (OQ-B7). Recommended path tracked: obtain/provide ≥1 golden output for fixed configuration/period.
- **Next (DAG, gates respected):** proceed with implementation-capable design/implementation where allowed (B4 config contract/validation, B5 filters DSL/parse+dependency-aware resolution design/impl per spec-fixed algorithm with OQ-T1 data-source assumptions explicit, B6 file resolution/access/session, B10/B11/B12 foundations) while keeping OQ-T1/T2/T3-dependent pivot/refresh/copy behaviors evidence-gated; do not claim correctness sign-off without golden/OQ closures.


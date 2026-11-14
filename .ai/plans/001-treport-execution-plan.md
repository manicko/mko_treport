# Execution plan: treport (001_treport)

- **Plan ID:** 001-treport-execution-plan
- **Replaces:** `.ai/plans/001-treport-execution-plan.md` (this document is the corrected replacement content)
- **Recommended filename:** `.ai/plans/001-treport-execution-plan.md` (unchanged)
- **Based on (current baseline):** `.ai/problems/001_treport_spec.md` — **rev. 4** (2026-11-14), all clarification options resolved by Product Owner as “all recommended”, and including the **PO-confirmed mandatory VBA-only implementation constraint** (no Python or any other programming/scripting language for implementation or automation).
- **Revision lineage (as determinable):** rev. 4 is the current baseline and supersedes rev. 3 as the planning/implementation baseline (rev. 3 predated the VBA-only constraint). Earlier revision sequence is not fully reconstructable from available documents; rev. 3 may be referenced below only as historical context where useful, never as the current baseline.
- **Fresh code/artifact context (Auditor, repo/history, reconciled to rev. 4):** spec rev. 4 present. Local (gitignored) artifacts observed historically/reported: `data/theiss/treport.xlsm` with `run`/`settings`, 4 ListObjects matching the rev. 4 schema/values (`general`: `input_path,output_path,delimeter`; `reports`: 6 rows with `articleLevel4Name_clean,targetAudience,name` incl. trailing space `DOLGIT `; `data_files`: 3 rows with `filters`/placeholder tokens; `placeholders`: `max_weeknum,max_monthnum,max_year`), `xl/vbaProject.bin` present, `run` has one button “Button 1” bound to `[0]!RunXlconcat` (no controls named/bound to `Refresh_data`/`Create_report` evidenced). No `xl/connections.xml`/pivots in `treport.xlsm`. `data/theiss/base/report_wclip.xlsx`, `monthly/report_m.xlsx`, `weekly/report_w.xlsx` exist with sheet `clean_data_full`, PQ connections/query tables/customXml and pivot cache/table parts; pivot caches show fields incl. `articleLevel4Name_clean`, `targetAudience`, `Year`/`Years` (partly), `Month Number`, `Week Number`, but exact pivot selection/orientation/captions/types, per-file period-field identity (`Year` vs `Years`, grouped `Years`), presented ranges, and copy fidelity remain unverified (OQ-T1/OQ-T3). PQ connections include user/machine-specific paths (portability/refresh risk). No `data/theiss/output/` directory evidenced; no golden outputs (OQ-B7). Repository artifacts: `mko_treport/src/vbs/vbaProject.bin` (snapshot) and exported VBA modules in `mko_treport/src/vbs/export/{Sheet1.cls,Sheet2.cls,ThisWorkbook.cls,xlconcat_macro.bas}` are present/committed for maintainability (behavior subject to OQ-T4; runtime-vs-export parity unverified). `mko_treport/src/scripts/` currently contains no treport implementation modules (non-compliant Python treport implementation modules were committed in 4e60165, 148b668, 3d7b4c2, 93e1ae4 and removed in e040c8c); tracked non-VBA test artifacts under `mko_treport/tests/treport_*` remain present and are **not rev. 4-compliant implementation evidence** (disposition/reconciliation pending clarification, spec §13.3). No CI/tooling for VBA acceptance. Known mismatches/gaps/risks: baseline `data_files.report_m.group_rows=TRUE` conflicts with D13/R4 (local remediation in gitignored `treport.xlsm` reported but **unverified in repo/for acceptance**); single legacy `RunXlconcat` control/entry point vs required `Refresh_data`/`Create_report` (FR-RUN-1, disposition needs clarification); VBA runtime behavior/refresh/copy fidelity/pivot-field identity/available values not fully verified (OQ-T1–T4); OQ-B7 golden baseline missing; `data_files.filters` use `Years` while inspected caches show `Year`/grouped `Years` differences (needs validation/alignment, no guessed mapping); configured category token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as distinct cache item in inspected extracts (fail-fast/mismatch risk, needs clarification); trailing space in `reports.name` `DOLGIT ` (whitespace handling needs clarification).
- **Planning intent:** dependency-safe, minimal-scope, incremental decomposition into execution blocks aligned to rev. 4. **No final implementation approach is chosen where technical uncertainty exists** (VBA architecture/refactor/entry-point structure, exact PQ refresh completion technique, pivot identification/filter/copy strategy, grouped-field/member mapping mechanics, and related platform-dependent mechanics remain alternatives pending OQ-T1–T4 evidence and required clarifications). No source code files are changed by this plan; only these two `.md` documents are edited.
- **Execution posture:** **Only VBA-compliant implementation work within `treport.xlsm` is in-scope to begin per plan where gates allow** (spec planning gate items 1–3 satisfied per §14). **Non-VBA implementation/automation work is prohibited and must not proceed** under this plan until reconciled/resolved per clarification. Correctness sign-off remains blocked until OQ-B7 golden baseline and OQ-T1–T4 validations/required clarifications are closed.

## Scope and guardrails

- **In scope:** decompose, sequence, and enable **VBA-only** implementation/validation work (Excel VBA macros/modules in `treport.xlsm`) to satisfy spec rev. 4 §§4–7/14 within the existing `treport.xlsm` Excel-macro model and configured `settings`/`run`/ListObjects (`general`, `reports`, `data_files`, `placeholders`) for operations `Refresh_data` and `Create_report`, including deterministic policies and verification aligned to D1–D17.
- **In scope (rev. 4 semantics to realize, VBA-only):** config reading/validation (FR-CFG); deterministic data-file resolution with precedence `.xlsx,.xlsm,.xlsb` (D2); access modes (D1: `Create_report` read-only never saved; `Refresh_data` writable refresh PQ/RefreshAll+save+close); `filters` DSL parse and dependency-aware `min/max` resolution and `max_*` placeholder derivation (D5,D6,D10,D11) with per-file verified field identities (no guessed mapping); data-file/report filter staging and placement (D4,D7,D8,D9); pivot snapshot extraction and static paste (D15); formatting/layout with `group_rows` fail-fast policy (D13); output assembly, placeholder substitution, naming/collision/idempotency (D3,D11,D14) with deterministic whitespace handling once clarified; fail-fast contextual diagnostics incl. no-data (D12); dropped parameters not required (D17); `run` controls/entry points reconciled for `Refresh_data`/`Create_report` per clarified requirements.
- **Out of scope / non-goals:** **any implementation or automation in Python or any other programming/scripting language (hard exclusion, rev. 4)**; architecture redesign or migration off Excel/VBA; non-Excel runtime/scheduler/UI redesign beyond required `run` controls/entry-point reconciliation; speculative abstractions or invented requirements; source data/schema/content or upstream PQ/data-model changes; implementing or choosing approaches that remain technically uncertain before evidence; `group_rows` color-band/outline behavior until an explicit deterministic rule is approved (D13); pixel-perfect parity without an agreed golden output (D18/OQ-B7); BI/business meaning beyond configured names/expressions.
- **Implementation placement / artifact policy (VBA-only):** implementation logic must reside in **VBA within `treport.xlsm`**. `mko_treport/src/vbs/` (or another explicitly approved path) may hold **versioned exports/snapshots only** for maintainability/inspection (behavior not assumed; runtime-vs-export parity subject to OQ-T4). **Placing treport implementation or automation in `mko_treport/src/scripts/` in Python or any other language is prohibited.** Tracked non-VBA artifacts that conflict with rev. 4 must be documented as discrepancies/open reconciliation actions (see V-002,V-003 and status); this plan does not authorize deleting or modifying source code files.
- **Method:** semantic anchors only (workbook/ListObjects/sheets/operations/fields/contracts/modules/procedures by responsibility) — no line numbers or positional assumptions. Tasks/blocks are atomic, independently reviewable/executable where feasible, and ordered by dependency/risk. Prefer isolated changes, incremental evolution, and backward-compatible/stable task boundaries.
- **Execution constraint (this interaction):** never change any source code files (`.py`, `.bas`, `.cls`, `.vba`, `.ts`, `.js`, `.sql`, etc.) or existing audit-phase files outside the two specified `.md` documents. Plan may prescribe VBA implementation/config actions (including edits to local gitignored artifacts) to be performed in subsequent compliant implementation steps where permitted and clarified.

## Resolved decisions and open gates

### Confirmed product decisions (D1–D17) — planning baseline fixed (rev. 4)

| ID | Decision (rev. 4, unchanged from confirmed semantics) | Planning/implementation implication (VBA-only) |
|---|---|---|
| D1 | `Create_report`: data files read-only, never saved. `Refresh_data`: writable (not read-only), refresh PQ/RefreshAll, save, close. | Enforce mode-specific open/close/save and state-safe cleanup in VBA (B6,B7,B12). |
| D2 | `data_file` base name resolved in `input_path` to `<data_file>.<ext>` with fixed precedence **`.xlsx`, `.xlsm`, `.xlsb`**; **fail fast** if no match or >1 match. | Deterministic resolution and contextual errors in VBA (B4,B6). |
| D3 | One output workbook per `reports` row named `reports.name` + `.xlsx`; overwrite if exists on successful run; sheets in `data_files` order, recreated/replaced deterministically for idempotent reruns. | Output lifecycle/idempotency in VBA (B11). |
| D4 | `data_files.filters` contains period/data fields as configured; `articleLevel4Name_clean` and `targetAudience` controlled via `reports`. | Filter responsibilities and staging (B5,B8). |
| D5 | `filters` grammar: clauses `|`; `<field> = <expr>` (spaces around `=`,`|` trimmed; field names may contain spaces). Expr: scalar literal or `min`/`max`; range `(start-end)` **inclusive**, bounds numeric or `min`/`max`, `-` range operator; list `(v1<delimeter>v2…)` via `general.delimeter`. Exact-after-trim match to pivot captions/items. Disambiguation: contains `general.delimeter` → list; else contains `-` between two bound tokens → range; else scalar. | Strict VBA parser/validator/disambiguation (B5). |
| D6 | Duplicated `Week Number=(1-max)|Years=max` example line treated as typo/ignored. | Stale ТЗ text must not override semantics (B5). |
| D7 | Placement applies to **all** filter fields (`data_files.filters` and `reports.articleLevel4Name_clean`, `reports.targetAudience`): filter in place if in Filters/Rows/Columns, else add to Filters; no move. | Per-field placement behavior (B8). |
| D8 | Per `data_file` apply `data_files.filters` once and keep; per `reports` row apply/reapply that row’s `articleLevel4Name_clean` and `targetAudience` overriding prior report-row selection for those fields, without moving placed fields; no leak between report rows. | Stage/reset/isolation per report row (B8,B11,B12). |
| D9 | Multi-value selection split by `general.delimeter`, tokens trimmed, all exactly matching pivot items selected (exact after trim). Applies to `articleLevel4Name_clean` and any list expressions in `filters`; `targetAudience` same exact/trim principle. | Deterministic matching (B5,B8). Missing required match must fail fast (D12); config/data alignment for observed token risk needs clarification (V-008, OQ-T1). |
| D10 | Resolve `min`/`max` per `data_files` row against available values of referenced fields in pivot on `data_sheet` of that `data_file` (actual data). Deterministic and dependency-aware: when `Years` present, resolve `Years` bound first, then resolve dependent period fields within that resolved `Years` context. Ranges **inclusive** both ends; comparisons **numeric**. Referenced field identity/caption must be verified per file and resolvable/matchable; do not guess mapping (evidence: configured `Years` vs cache `Year`/grouped `Years`; needs OQ-T1 clarification/alignment). | Core resolution algorithm fixed; per-file field identity/value-source and grouped-`Years` semantics evidence-gated (B2,B5,B8,B13). |
| D11 | `max_year` = resolved max of `Years`; `max_weeknum` = resolved max of `Week Number`; `max_monthnum` = resolved max of `Month Number` for that `data_files` row in same resolved context. Substitution target **`report_sheet` only**. Token syntax `<name>` as in workbook, substituted as resolved values **without zero-padding**. Referenced but underivable ⇒ **fail fast**. Derivations depend on verified field identities/semantics (no guess). | Placeholder derivation/substitution and naming (B5,B11,B12). |
| D12 | **Fail fast** with clear contextual error for missing/invalid config or paths, missing/unresolvable data file, missing `data_sheet`, missing pivot, missing required filter field, invalid/unparseable/unresolvable `filters`, or referenced-but-underivable placeholder. Filter yielding **no data** ⇒ **fail fast** with context. | Unified error taxonomy and outcomes (B4–B13). |
| D13 | Blanks = default/no-op for `freeze_row/freeze_col/zoom/paste_row/paste_col/group_rows`. `group_rows` has no approved deterministic rule; **`group_rows=TRUE` must fail fast** until explicit deterministic rule approved. Blank/false = no-op. R4 baseline observed `report_m.group_rows=TRUE` conflict; local remediation reported but **unverified in repo/for acceptance** (pending verification). | Validation/formatting policy and R4 compliance verification (B1,B4,B10,B12,B13). |
| D14 | Sanitize Excel-invalid sheet-name characters, enforce **31-character** limit, **fail fast** on sheet-name or report-file-name collision after substitution/sanitization. Leading/trailing whitespace handling for configured names before/within substitution/naming/collision is not explicitly specified and needs clarification (V-020). | Naming/sanitization/collision handling (B11,B12). |
| D15 | Copy pivot presented data range incl. row/column labels and grand totals as shown after filtering; exclude pivot object/page-filter area and PQ connections; paste static **values + formats** (+ column widths if needed). | Extraction/paste fidelity (B9,B13). |
| D16 | **VBA-only** solution in `treport.xlsm` (Excel VBA macros/modules); **no Python or any other programming/scripting language**. Behavior aligned to revised ТЗ intent reconciled with current workbook structure; `Refresh_data` and `Create_report` in scope. | Solution boundary and reconciliation (all blocks). |
| D17 | `new_file_delimeter` and `keep_col_hidden` dropped (absent rev. 2) and must not be required/processed. | Config contract/validation (B4). |

### Open/product-validation gates and clarifications

| ID | Status | Recommended path / notes | Blocks / effect |
|---|---|---|---|
| **OQ-B7 (D18)** | **Pending** (does not block planning) | **Obtain/provide ≥1 golden expected output for a fixed configuration/period** on target Excel/OS with representative `data/theiss` dataset (and document fixed config/period, environment, input state). | **Required for correctness sign-off** (values/formatting/layout/copy/fidelity). Structural checks insufficient (B3,B13). |
| **R4 / D13 config conflict** | **Pending verification / needs clarification** | Baseline observed: `data_files.report_m.group_rows=TRUE` (conflicts D13). **Reported**: remediated locally to blank/false in gitignored `data/theiss/treport.xlsm` (recommended). **Status (consistent):** resolved locally in gitignored artifact — **unverified in repo/for acceptance, requires verification**. | Must be compliant before acceptance; verification required (B1,B3,B13). Do not assert verified. |
| **VBA source/runtime parity** | **Known maintainability/gap context** | Export/version VBA modules for maintainability where feasible (done), but **do not assume behavior**; runtime-vs-export parity unverified (OQ-T4). | Supports reconciliation/maintenance; residual risk tracked. |
| **FR-RUN-1 controls/entry points (V-006, V-014)** | **Needs clarification (open)** | Required end state: `run` sheet provides invokable VBA controls/entry points for `Refresh_data` and `Create_report`. Current: one button bound to legacy `[0]!RunXlconcat` (concatenation, non-compliant). **Clarification needed:** control names/captions/bindings and VBA entry-point signatures, and disposition of `RunXlconcat` (remove, replace, or retain separately). | Implementation/acceptance reconciliation required (B2,B4,B12,B13,B14). |
| **Period field identity / `Years` vs `Year` / grouped `Years` (V-007, V-018, OQ-T1)** | **Needs clarification / pending evidence** | Configured `filters` use `Years`; evidence indicates `report_m`/`report_w` caches expose `Year` (singular) while `report_wclip` has grouped `Years`. Authoritative per-file names/captions and grouped-`Years` semantics/mapping unconfirmed. D5/D10/D12 require verified, resolvable/matchable fields/values or fail-fast; **no guessed mapping**. Validate/align `data_files.filters`/`max_*` once evidenced. | Gates filter/placeholder behavior (B2,B5,B8,B13). |
| **Category token alignment (V-008, OQ-T1)** | **Needs clarification / pending evidence** | Configured token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as distinct cache item in inspected extracts. D9/D12 require exact-after-trim match or fail-fast; **no mapping assumed**. | Config/data alignment and acceptance risk (B2,B5,B8,B13). |
| **Whitespace handling for configured names (V-020)** | **Needs clarification** | Deterministic handling of leading/trailing whitespace in `reports.name`, `report_sheet`, and other configured string values before/within substitution, naming, sanitization, and collision checks not explicitly specified (observed trailing space in `DOLGIT `). Define and apply consistently. | Deterministic naming/substitution/collision behavior (B4,B11,B12). |
| **VBA-only reconciliation of non-compliant artifacts (V-002, V-003)** | **Needs clarification (open)** | Rev. 4 mandates VBA-only implementation. Prior committed non-VBA Python treport implementation modules (4e60165,148b668,3d7b4c2,93e1ae4) removed from `mko_treport/src/scripts/` in e040c8c, but related tracked non-VBA artifacts remain under `mko_treport/tests/treport_*` (observed tracked dirs include `treport_config`, `treport_files`, `treport_filters`, `treport_filter_application`, `treport_refresh`) and are **not rev. 4-compliant implementation evidence**. Disposition/reconciliation and acceptance impact require clarification; this plan cannot delete/modify source code. | Progress/compliance accounting and release readiness (B1,B13,B14). |

### Technical validation gates (pending evidence; do not choose final approach)

| ID | Uncertainty | Planning treatment | Effect |
|---|---|---|---|
| **OQ-T1** | Actual pivot(s), field captions/types/orientations, presented ranges, PQ/connection topology; per-file `Year`/`Years`/grouped-`Years` identity/semantics, `Week Number`, `Month Number`, `articleLevel4Name_clean`, `targetAudience` cache items. | D10 algorithm fixed but available-value sourcing/member semantics must be validated; **pivot identification/filtering/copy strategy not chosen** pending evidence. No guessed mapping. | Gates B5,B8,B9,B11. |
| **OQ-T2** | Synchronous PQ refresh + pivot cache refresh completeness on target Excel. | Validate on target environment; **exact PQ refresh completion technique not chosen** pending evidence. | Gates B7 finalization/validation. |
| **OQ-T3** | Copy-range boundaries and formatting fidelity against golden output. | Golden-based verification; **pivot/copy strategy not chosen** pending evidence. | Gates B9 fidelity/sign-off. |
| **OQ-T4** | Reconcile `treport.xlsm` VBA behavior with rev. 4. Exported modules/snapshot exist but runtime-vs-export parity unverified. | Inspect/reconcile in Excel/VBA IDE or approved extraction as needed; **VBA architecture/refactor/entry-point approach not chosen** pending evidence/clarification. | Constrains blocks; informs B2,B12,FR-RUN-1. |

## Local artifact and gitignore constraints (must be accounted for)

- `data/theiss/*` and `data/theiss/treport.xlsm` are **gitignored local artifacts** (not repo-tracked). Plan, evidence, and validation must not assume repository availability or CI visibility.
- Baseline capture, R4 compliance verification/remediation evidence, VBA export, golden-output acquisition, and test runs must be handled in target/local environment and recorded reproducibly (paths, fixed config/period, environment, hashes/timestamps where useful) without committing ignored data.
- `mko_treport/src/vbs/` contains tracked VBA snapshot/export artifacts for maintainability (not behavior proof). `mko_treport/src/scripts/` currently contains no treport implementation modules; `mko_treport/src/xls/` is empty. Tracked non-VBA test artifacts under `mko_treport/tests/treport_*` exist and require disposition/reconciliation clarification (cannot be modified/deleted by this plan).
- No `data/theiss/output/` currently evidenced; outputs are created per D3/FR-FILE-5 during runs and remain local unless explicitly managed.

## Agent role key

- **Auditor:** deeper code/architecture/artifact investigation where uncertainty or complexity exists.
- **Researcher:** modern/best-practice and viable-approach/platform-behavior guidance where multiple approaches or Excel/PQ/pivot/VBA behavior matter.
- **Planner:** detailed pre-implementation design, sequencing, and testing/verification shaping.
- **Validator:** independent plan/review/verification gate when implementation or correctness risk is high.
- **High-risk rule:** blocks marked **High risk** require **all four (Auditor, Researcher, Planner, Validator)**.

## R1–R9 traceability mapping (spec §10)

| Risk | Summary | Plan blocks / gates |
|---|---|---|
| **R1** | Pivot topology & field identity (pivot selection, captions/types/orientations, per-file `Year`/`Years`/grouped-`Years`, regular vs OLAP/Data Model) | B0, B2 (OQ-T1), B5, B8, B9, B13; V-007/V-018 |
| **R2** | PQ refresh completeness/timing across Excel builds/query types | B0, B2, B3, B7 (OQ-T2), B13 |
| **R3** | `min`/`max` data source/member details consistent with D10 (incl. grouped-`Years` semantics) | B2 (OQ-T1), B5, B8, B13; V-007/V-018 |
| **R4** | `group_rows` config conflict / D13 compliance verification | B0, B1, B4, B10, B12, B13, B14; pending verification |
| **R5** | Copy fidelity (range bounds, labels/totals/page area, widths/heights/hidden/outline, value/date fidelity) | B2, B3, B9 (OQ-T3), B10, B13 |
| **R6** | Locked/open file handling and filesystem permissions | B3, B6, B7, B11, B12, B13 |
| **R7** | VBA logic opaque/runtime behavior and parity with rev. 4 | B0, B1, B2 (OQ-T4), B12, B14; V-006/V-014 |
| **R8** | Performance on large pivots/many items | B2, B8, B9, B13 (benchmark; no semantics change) |
| **R9** | Environment prerequisites (Excel version, Windows, macro trust, PQ enabled, paths) | B0, B3, B7, B13, B14 |

## Logical execution blocks (B0–B14)

### B0. Baseline evidence & validation setup
- **Objective:** establish inspectable baseline aligned to rev. 4: confirm local artifact inventory, record gitignored/local-only status, capture target-environment prerequisites relevant to OQ-B7/OQ-T2, and define reproducible evidence collection for OQ-T1–T4, OQ-B7, R4 verification, and required clarifications without modifying source code outside permitted local evidence handling.
- **Status:** **Completed (evidence/checklist, factual).** Prior audit recorded artifact inventory/state, environment capture template, fixed config/period tuple, and blocker checklist. Full target sign-off details, golden baseline, and authoritative R4 compliance evidence remain pending.
- **Depends on:** —
- **Key risks:** Impl: medium. Rollout: low. Regression: low. Compatibility: high (PQ/macro-trust, Excel build differences, missing golden). Evidence drift if local artifacts change without recording fixed state.
- **Alternatives / paths (not chosen):** Excel/VBA IDE inspection and/or approved export vs package/XML inspection only; must not assume behavior.
- **Required agents:** **Auditor, Planner, Validator**
- **Gate / exit criteria:** evidence inventory and validation checklist ready; no assumptions recorded as facts.

### B1. Spec alignment, config compliance & maintainability actions (VBA-only context)
- **Objective:** align execution baseline to rev. 4 decisions D1–D17 and hard VBA-only constraint; maintain D13/R4 compliance posture (reported local remediation in gitignored `treport.xlsm` to blank/false for `report_m.group_rows`, but **unverified in repo/for acceptance** — verification required); maintain/export VBA modules/snapshots for maintainability where feasible without assuming behavior; and document reconciliation status of non-compliant non-VBA artifacts (V-002,V-003) as open clarification/disposition actions.
- **Status:** **Partially complete / pending verification and clarification.** Spec/plan alignment to rev. 4 corrected in this document set. R4 remediation reported as resolved locally in gitignored artifact but **unverified in repo/for acceptance**. VBA snapshot/exports committed for maintainability (OQ-T4 parity unverified). Disposition/reconciliation of remaining non-compliant non-VBA artifacts (incl. tracked `mko_treport/tests/treport_*`) is **needs clarification**.
- **Depends on:** B0.
- **Key risks:** Impl: low-medium. Rollout: medium-high if R4 compliance unverified or non-VBA artifacts misinterpreted as compliant progress. Regression: low-medium. Compatibility: low.
- **Alternatives / paths (not chosen):** approve deterministic `group_rows` rule and update spec/policy — not chosen now. VBA architecture/refactor/entry-point approach not chosen.
- **Required agents:** **Planner, Validator, Auditor**
- **Gate / exit criteria:** rev. 4/VBA-only baseline acknowledged; R4 compliance verified with evidence for authoritative acceptance artifact/config (or explicitly tracked pending); non-compliant non-VBA artifacts documented with open disposition/clarification; export/snapshot parity caveat noted.

### B2. Technical investigation — pivots, fields, PQ, VBA, controls (OQ-T1, OQ-T4, V-006–V-008, V-014, V-018)
- **Objective:** determine/confirm per data file on `data_sheet` (`clean_data_full`): actual pivot object(s), field captions/types/orientations and availability for period/category/audience fields (incl. authoritative `Year`/`Years`/grouped-`Years` identity and semantics, `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience` and cache-item presence for configured tokens), presented ranges/totals/page-filter areas, PQ/connection/query topology and refresh dependencies; control bindings/entry points and reconcile observable VBA behavior with rev. 4 where possible.
- **Status:** **Deepened (evidence gathered), gates remain open.** Exported VBA-source and package/XML reconciliation completed. `xlconcat_macro.RunXlconcat` confirmed legacy/non-compliant for rev. 4 operations (runtime-vs-export parity unverified). Pivot/field/cache/PQ indicators partially evidenced; authoritative identities/semantics, runtime behavior, and copy/refresh fidelity remain unverified (OQ-T1–T4). Clarifications needed.
- **Depends on:** B0.
- **Key risks:** Impl: high. Rollout: high. Regression: high. Compatibility: high.
- **Alternatives / paths (not chosen):** inspection/extraction path per access; **pivot identification/filtering/copy strategy and VBA architecture/refactor/entry-point approach not chosen** pending findings/clarification.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** per-file semantic map or explicit unknowns with impact/risk and required clarifications; no guessed behavior/mapping.

### B3. Golden baseline & target-environment confirmation (OQ-B7, prereqs for OQ-T2/T3)
- **Objective:** confirm/document target Excel/OS and representative input state; obtain/provide ≥1 golden expected output for fixed configuration/period (OQ-B7/D18) including resolved period/filters/config; establish reproducible capture conditions and R4 compliance verification evidence.
- **Status:** **Partially started, blocked for sign-off.** Target class known; full target build/macro-trust/PQ details, fixed authoritative config/period/R4 evidence, and golden expected output not provided.
- **Depends on:** B0; informed by B1,B2.
- **Key risks:** Impl: medium. Rollout: high. Regression: high. Compatibility: high.
- **Alternatives / paths (not chosen):** PO-provided or collaboratively captured authoritative baseline once VBA-compliant behavior/evidence validated.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** target environment documented; fixed configuration/period and representative dataset (incl. verified R4-compliant state) recorded; golden output(s) traceable or explicitly tracked as blocking sign-off.

### B4. Config contract & validation (FR-CFG, T1) — VBA-only
- **Objective:** implement deterministic **VBA** config reading/validation for `general`, `reports`, `data_files`, `placeholders` per rev. 4: required/non-empty fields, defaults/blanks/no-op, D17 dropped params not required/processed, D13 `group_rows=TRUE` fail-fast until rule approved, and fail-fast contextual errors (D12). Include deterministic whitespace handling for configured names/values (V-020) once clarified, consistent with D5/D9 where applicable.
- **Status:** **Not started as rev. 4-compliant VBA implementation.** Non-compliant Python config artifacts committed (4e60165) and removed from `mko_treport/src/scripts/` (e040c8c); related tracked tests may remain — **not compliant** with rev. 4, disposition pending clarification.
- **Implementation start condition:** may start only for **VBA-compliant** work in `treport.xlsm` after applicable B1/B2 inputs and required clarifications are sufficiently resolved; no Python/other-language implementation permitted.
- **Depends on:** B1, B2, V-020 as applicable.
- **Key risks:** Impl: medium. Rollout: medium. Regression: medium. Compatibility: low-medium.
- **Required agents:** **Planner, Validator** (+ **Auditor** if needed)
- **Gate / exit criteria:** VBA config schema/validation with deterministic outcomes/errors covering FR-CFG-1..6, D12,D13,D17; `group_rows=TRUE` rejected per D13; R4-compliant config accepted; whitespace handling applied per clarification or tracked open.

### B5. Filters DSL parsing & dependency-aware resolution (FR-FLT-1/2/3, FR-NAME-1 inputs, T2) — VBA-only
- **Objective:** implement deterministic **VBA** parsing/resolution per D5,D6,D10,D11 with dependency-aware `min/max`, placeholder derivations in same resolved context, and fail-fast on unresolvable/underivable cases. **No guessed field mappings**: validate/align with verified per-file field identities/captions/semantics (V-007,V-018,OQ-T1). Account for category token alignment risk (V-008).
- **Status:** **Not started as rev. 4-compliant VBA implementation.** Non-compliant Python artifacts committed (148b668) and removed from `mko_treport/src/scripts/` (e040c8c); related tracked tests may remain — **not compliant**, disposition pending clarification.
- **Implementation start condition:** VBA-compliant work only; correctness-sensitive value-source/member semantics and field alignment gated by OQ-T1/clarifications.
- **Depends on:** B4, B1, B2, V-007/V-008/V-018.
- **Key risks:** Impl: high. Rollout: medium-high. Regression: high. Compatibility: medium-high.
- **Required agents:** **Planner, Validator, Auditor, Researcher** (all)
- **Gate / exit criteria:** deterministic VBA parse/resolve/derive behavior for D5,D6,D10,D11 with contextual fail-fast errors; resolved values/logging available; assumptions tied to OQ-T1 evidence and resolved clarifications.

### B6. File resolution & access/session policy (FR-FILE, T3) — VBA-only
- **Objective:** implement deterministic **VBA** D2 resolution and D1 mode-specific access/session behavior, including `output_path` readiness and state-safe cleanup.
- **Status:** **Not started as rev. 4-compliant VBA implementation.** Non-compliant Python artifacts committed (3d7b4c2) and removed from `mko_treport/src/scripts/` (e040c8c); related tracked tests may remain — **not compliant**, disposition pending clarification.
- **Implementation start condition:** VBA-compliant work only.
- **Depends on:** B4, B1, B2.
- **Key risks:** Impl: high. Rollout: high. Regression: high. Compatibility: medium-high.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** deterministic D2/D1 enforcement in VBA with contextual fail-fast errors; no source save in `Create_report`; cleanup/state restoration defined/testable.

### B7. Refresh workflow (FR-REF, T4; OQ-T2) — VBA-only
- **Objective:** implement reliable **VBA** `Refresh_data` flow (open writable, refresh PQ/dependent pivot data to persisted current state, save, close) with fail-fast contextual errors and no state corruption.
- **Status:** **Not started as rev. 4-compliant VBA implementation; gated by OQ-T2.** Non-compliant Python artifacts committed (93e1ae4) and removed from `mko_treport/src/scripts/` (e040c8c); related tracked tests may remain — **not compliant**, disposition pending clarification.
- **Implementation start condition:** VBA-compliant work only; completion/validation gated by OQ-T2.
- **Depends on:** B6, B2, B3.
- **Key risks:** Impl: high. Rollout: high. Regression: high. Compatibility: high.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** evidenced refresh-completion criteria satisfied on target environment (OQ-T2); no stale-save path; contextual failure reporting.

### B8. Filter application, staging & placement (FR-FLT-4/5/6/7, T5; OQ-T1) — VBA-only
- **Objective:** apply resolved filters deterministically in **VBA** per D4,D7,D8,D9,D12 with verified field identities/semantics, per-report isolation, and fail-fast on missing/unresolvable fields/values/matches or no-data outcomes.
- **Status:** **Not started as rev. 4-compliant VBA implementation; pivot-specific behavior gated by OQ-T1/clarifications.**
- **Depends on:** B5, B4, B6, B2, V-007/V-008/V-018, B7 context as relevant.
- **Key risks:** Impl: high. Rollout: high. Regression: high. Compatibility: high.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** stage/placement/matching/isolation behavior confirmed against OQ-T1 findings and resolved clarifications; fail-fast cases verified.

### B9. Pivot snapshot extraction & static paste (FR-RPT-3/4, T6; D15; OQ-T3) — VBA-only
- **Objective:** capture presented pivot data range after filtering and paste static values+formats (+ column widths if needed) per D15, with no pivot/PQ artifacts in output.
- **Status:** **Not started as rev. 4-compliant VBA implementation; fidelity gated by OQ-T1/OQ-T3/golden.**
- **Depends on:** B8, B2, B3.
- **Key risks:** Impl: high. Rollout: medium-high. Regression: high. Compatibility: medium-high.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** copy-range/paste contract matches D15 and validated against golden output.

### B10. Formatting & layout (FR-RPT-5/6, T7; D13) — VBA-only
- **Objective:** apply deterministic **VBA** formatting (freeze/zoom/gridlines/sizing/state) with blanks as no-op; enforce D13 `group_rows` policy; verify R4-compliant state for acceptance.
- **Status:** **Not started as rev. 4-compliant VBA implementation.** Fidelity checks gated by OQ-T3/golden.
- **Depends on:** B4, B9, B1.
- **Key risks:** Impl: medium. Rollout: medium. Regression: medium. Compatibility: medium.
- **Required agents:** **Planner, Validator, Auditor** (+ **Researcher** if needed)
- **Gate / exit criteria:** formatting criteria met; D13 fail-fast preserved for `TRUE`; no undefined `group_rows` behavior.

### B11. Report assembly, naming, placeholders & idempotency (FR-RPT-1/2/7, FR-NAME, T8; D3,D11,D14) — VBA-only
- **Objective:** deterministic **VBA** output workbook/sheet lifecycle, placeholder substitution in `report_sheet` only, sanitization/31-char/collision fail-fast, idempotent reruns, save/cleanup per D1/D3/D11/D14. Apply clarified whitespace policy (V-020).
- **Status:** **Not started as rev. 4-compliant VBA implementation.**
- **Depends on:** B5, B6, B9, B10, B4, V-020 clarification.
- **Key risks:** Impl: high. Rollout: high. Regression: high. Compatibility: medium.
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** deterministic idempotent outputs with correct names/order/paths; clarified whitespace policy applied; no pivot/PQ artifacts.

### B12. Diagnostics, run control & fail-fast error taxonomy (FR-RUN, T9) — VBA-only
- **Objective:** unified **VBA** fail-fast contextual diagnostics and state-safe run outcomes; reconcile required `run` controls/entry points for `Refresh_data` and `Create_report` (FR-RUN-1, V-006/V-014) and disposition of legacy `RunXlconcat` once clarified.
- **Status:** **Not started as rev. 4-compliant VBA implementation.** Incremental with B4–B11, consolidated before sign-off.
- **Depends on:** B4–B11, B2/OQ-T4, FR-RUN-1 clarifications.
- **Key risks:** Impl: medium-high. Rollout: medium. Regression: medium. Compatibility: medium.
- **Required agents:** **Planner, Validator, Auditor** (+ **Researcher** if needed)
- **Gate / exit criteria:** unified error taxonomy/outcomes per D12–D14/FR-RUN; required controls/entry points reconciled per clarification; state restored on completion/failure.

### B13. Verification & acceptance (T10; D18/OQ-B7, OQ-T1–T4) — target Excel/VBA context
- **Objective:** define and execute verification scenarios and acceptance checks in **target Excel/VBA runtime context** consistent with rev. 4 (config/D1/D2, refresh OQ-T2, filters/resolution/field identity D5,D6,D10,D11/V-007,V-018, staging/matching/no-leak/no-data/token alignment D7–D9/D12/V-008, copy/format D15/OQ-T3, naming/collision/idempotency/whitespace D3,D14,V-020, controls FR-RUN-1/V-006,V-014, R4 verified compliance, golden comparison OQ-B7).
- **Status:** **Scenario planning can start; execution/sign-off gated.** Blocked by OQ-B7, OQ-T1–T4, R4 verification, and required clarifications.
- **Depends on:** B3, B4–B12, B2, clarified open items.
- **Key risks:** Impl: medium. Rollout: high. Regression: high. Compatibility: high. False confidence from structural-only or non-compliant/non-VBA evidence.
- **Verification evidence note (V-019):** sign-off evidence must be obtained in target Excel/VBA context consistent with VBA-only implementation. External/non-VBA tooling may support analysis but must not be treated as compliant implementation evidence unless explicitly permitted (no additional prohibition implied beyond rev. 4 intent).
- **Required agents:** **Auditor, Researcher, Planner, Validator** (all)
- **Gate / exit criteria:** executable evidence in target context; pass/fail per scenario; R4 verified; OQ-T1–T4 and required clarifications closed or explicitly accepted with rationale; golden-based sign-off achieved or explicitly pending.

### B14. Independent review & release-readiness gate
- **Objective:** independent review that delivered behavior satisfies rev. 4 (D1–D17, VBA-only, §§4–7/13/14), R4 compliance verified, controls/fields/naming clarifications reconciled, no reliance on unchosen approaches beyond evidence, and required validation/sign-off artifacts present (or explicit accepted limitations). Non-compliant artifact disposition documented/resolved as required.
- **Status:** **Gated** — after VBA-compliant implementation and verification reach reviewable completeness.
- **Depends on:** B1–B13.
- **Key risks:** Impl: low. Rollout: high. Regression: high. Compatibility: medium.
- **Required agents:** **Validator** (lead), with **Planner, Auditor, Researcher** as needed
- **Gate / exit criteria:** dependency-safe readiness confirmed or clearly bounded by open items; no unresolved contradiction with rev. 4/D1–D17/R4/VBA-only policy.

## Execution order and dependency DAG

### Textual DAG (rev. 4, VBA-only compliant path)
1. **B0** (baseline evidence & setup) — **completed (evidence/checklist)**.
2. **B1** (spec alignment, R4 compliance verification posture, VBA export/maintainability, non-VBA reconciliation tracking) — **partially complete / pending verification and clarification**, depends on B0.
3. **B2** (technical investigation OQ-T1,OQ-T4, controls/field/token clarifications) — **deepened, gates open**, depends on B0. *(high risk)*
4. **B3** (golden baseline & target environment OQ-B7; prereqs OQ-T2/T3; R4 verification evidence) — **partially started**, depends on B0, informed by B1,B2. *(sign-off gated)*
5. **B4** (VBA config contract & validation, T1) — **not started as compliant VBA implementation**, depends on B1,B2 and applicable clarifications.
6. **B5** (VBA filters DSL parse & dependency-aware resolution, T2) — **not started as compliant VBA implementation**, depends on B4,B1; validated by B2/OQ-T1 and V-007/V-008/V-018. *(evidence-gated details)*
7. **B6** (VBA file resolution & access/session, T3) — **not started as compliant VBA implementation**, depends on B4,B1,B2.
8. **B7** (VBA refresh workflow, T4) — **not started; completion/validation gated by OQ-T2**, depends on B6,B2,B3. *(high risk)*
9. **B8** (VBA filter application/staging/placement, T5) — **not started; gated by OQ-T1/clarifications**, depends on B5,B4,B6,B2 and refresh context as relevant. *(high risk)*
10. **B9** (VBA pivot snapshot extraction & paste, T6) — **not started; gated by OQ-T1/OQ-T3/golden**, depends on B8,B2,B3. *(high risk)*
11. **B10** (VBA formatting & layout, T7) — **not started as compliant VBA implementation**, depends on B4,B9,B1.
12. **B11** (VBA report assembly/naming/placeholders/idempotency, T8) — **not started as compliant VBA implementation**, depends on B5,B6,B9,B10,B4,V-020 clarification.
13. **B12** (VBA diagnostics/run control/error taxonomy, T9) — **not started as compliant VBA implementation**, incremental with B4–B11 and FR-RUN-1 reconciliation.
14. **B13** (verification & acceptance in target Excel/VBA context, T10) — **scenario planning can start; execution/sign-off gated**, depends on B3,B4–B12,B2 and closures.
15. **B14** (independent review & release-readiness) — **gated**, depends on B1–B13.

### Safe parallelization opportunities (when prerequisites met, VBA-compliant only)
- **B1 ∥ B2 ∥ B3** evidence/clarification/verification activities.
- **B4** once B1 alignment is clear and applicable clarifications addressed, incorporating B2 observations as they arrive.
- **B5 ∥ B6** after B4 (different concerns), with B5 constrained by B2/OQ-T1 findings.
- **B7** design and **B5/B8/B11** non-pivot-specific VBA design can progress in parallel where interfaces are clear; completion remains gated.
- **B10, B11, B12** can advance incrementally with B9 as extraction contracts stabilize.
- **B13** scenario design runs early and continuously; execution evidence accrues as gates/clarifications close.

## Cross-cutting risks and containment

| Risk / area | Impact | Containment / gate |
|---|---|---|
| **VBA-only constraint and non-compliant artifacts** (rev. 4, V-002,V-003) | Non-compliant implementation/tests mistaken for progress; acceptance confusion | Hard VBA-only scope/non-goal/placement/start-condition rules; document and seek clarification on disposition of tracked non-VBA artifacts; no source-code deletion/modification by this plan. |
| **VBA opacity & unverified runtime behavior** (R7, OQ-T4) | Hidden logic diverges from rev. 4; wrong assumptions | B2 investigation/reconciliation; exports/snapshots without assuming behavior/parity; no final VBA/refactor/entry-point approach chosen prematurely. |
| **Pivot topology, field identity & period semantics** (R1, R3, OQ-T1, V-007,V-008,V-018) | Wrong target/field, incorrect `min/max`/`max_*`, missing matches, misapplied filters | B2 evidence; D10 algorithm fixed but value-source/member mechanics and per-file identities validated/aligned; fail fast (D12); no guessed mapping. |
| **PQ refresh completeness/timing** (R2, OQ-T2) | Stale persisted data after `Refresh_data` | B7 validation on target Excel; exact completion technique not chosen until evidenced; no stale-save acceptance. |
| **Copy-range/format fidelity** (R5, OQ-T3) | Output differs from intended presented data/layout | B9/B13 golden-based validation; copy strategy not chosen until evidenced. |
| **`group_rows` policy/compliance verification** (R4, D13, V-004) | Non-compliant config or misreported status | Reported local remediation unverified; require verified evidence of authoritative compliant state before acceptance; enforce D13 fail-fast for TRUE until rule approved. |
| **`run` controls/entry-point reconciliation** (FR-RUN-1, V-006,V-014) | Required operations not invokable; legacy macro confusion | B2/B12 reconciliation tasks; clarification required; no assumption of compliant binding. |
| **No-data and error handling gaps** (D12) | Silent wrong outputs or unclear failures | Fail-fast contextual errors for specified cases incl. no-data and missing required matches; unified taxonomy in B12. |
| **Naming/collision/idempotency/whitespace defects** (D3, D14, V-020) | Overwrites, collisions, unstable reruns, inconsistent names | Deterministic sanitize/31-char/collision fail-fast, idempotent lifecycle; whitespace handling clarified and applied consistently. |
| **Locked files/permissions/environment variance** (R6, R9) | Run failures, partial outputs, non-reproducible behavior | D1/D12 fail-fast with context; target env documented (B3); cleanup/state restoration. |
| **Performance on large pivots/many items** (R8) | Long/unresponsive runs | Benchmark during B13; optimize only with evidence and without changing semantics. |
| **Local-only gitignored artifacts** (`data/theiss/*`, `treport.xlsm`) | Non-reproducible evidence or unnoticed edits | Explicit inventory/capture and verified state recording; controlled local actions with backup/traceability. |
| **Spec/observed artifact discrepancies** (stale ТЗ text, trailing space, Year/Years, category token) | Confusion or incorrect assumptions | Reconcile against rev. 4 and evidence; D6 for stale example; track clarifications explicitly. |

## Definition of Ready alignment (spec §14, rev. 4)

| §14 item | Status | Notes |
|---|---|---|
| **1. Business decisions/clarifications closed (D1–D17)** | **Satisfied** | PO selected all recommended options; OQ-1a–d, OQ-2a–c, OQ-3a–b, OQ-4, OQ-B1, OQ-B3–B6 resolved. Rev. 4 adds mandatory VBA-only constraint (D16/A1). |
| **2. Scope & requirements stable (§§4–7)** | **Satisfied (with open implementation clarifications tracked)** | VBA-only boundary and core semantics stable. Open clarifications (FR-RUN-1 controls, field identity/alignment, token/whitespace, artifact reconciliation) documented in spec §13.3 and must be resolved or explicitly accepted where they affect implementation/acceptance. |
| **3. Deterministic policies fixed** | **Satisfied (subject to clarified name-whitespace policy)** | Core deterministic policies fixed. Deterministic whitespace handling for configured names requires clarification (V-020) and must be defined consistently before claiming final naming determinism. |
| **4. Validation baseline available (OQ-B7)** | **Pending (does not block planning)** | Target Excel/OS + representative dataset + ≥1 golden expected output for fixed configuration/period required for correctness sign-off (B3,B13). |
| **5. Technical validations tracked (§13.2/§10)** | **Tracked, pending closure** | OQ-T1–T4 gates plus R4 compliance verification and documented clarifications. |

- **Planning gate:** items 1–3 satisfied as baseline; open clarifications in spec §13.3 must be resolved or explicitly accepted where they affect implementation/acceptance. Item 4 pending but does not block planning; **items 4–5 and required clarifications/R4 verification must close before correctness sign-off**.
- **Implementation start:** **allowed only for work compliant with rev. 4 VBA-only constraint** where block prerequisites/gates permit. **Non-VBA implementation/automation work is prohibited** under this plan until reconciled/resolved per clarification.
- **Acceptance note (R4):** baseline config observed `group_rows=TRUE` for `report_m` (conflicts D13). A local remediation in gitignored `data/theiss/treport.xlsm` is reported but **unverified in repo/for acceptance**; compliance must be verified before acceptance, or an approved deterministic rule must be reflected in spec/policy.

## Expected deliverables (when blocks execute, VBA-compliant)

- Evidence records for B0/B2/B3 (artifact/topology/PQ/control findings, target environment, fixed config/period, R4 verification evidence, golden baseline or tracked gap, clarification outcomes).
- Verified D13-compliant authoritative config state for R4 (or approved-rule alternative with spec/policy update) and, where feasible, versioned VBA exports/snapshots for maintainability (behavior/parity not assumed).
- Dependency-safe **VBA** implementation increments in `treport.xlsm` covering B4–B12 semantics, with approaches left open only where OQ-T1–T4/clarifications require evidence.
- Documented reconciliation/disposition outcomes for non-compliant non-VBA artifacts (per clarification), without unauthorized source-code modification.
- Verification artifacts and sign-off package for B13/B14 in target Excel/VBA context (scenario results, OQ-T1–T4/R4/clarification closure evidence, golden-based confirmation or explicit pending blockers).

## Execution progress/status (current, normalized; rev. 4 baseline)

> Status notes are factual progress records aligned to spec rev. 4. Confirmed decisions (D1–D17), VBA-only constraint, gates (OQ-B7 pending golden/sign-off; OQ-T1–T4 pending evidence/closure; R4 pending verification; FR-RUN-1/field/token/whitespace/artifact-reconciliation clarifications pending), and “no final implementation approach chosen where technical uncertainty exists” remain in force.

### Validation traceability (in-session; N-003)
- Prior in-session validation/planning runs (this workspace/session lineage, not embedded as standalone report files): validation 1 `ses_e24a365b4ffeZyl6DWd7Qekjis` (reported Not aligned, V-001..V-020), planner correction 1 `ses_e249c8a66ffevRL7YZe0ZgfpdH` (updated spec/plan to rev. 4 baseline), validation 2 `ses_e24810fadffeH5EKlS1hrS24ja` (reported Not aligned; many V-001..V-020 addressed/partially addressed but some untraceable without prior report text; remaining/new N-001–N-003).
- **Traceability limitation:** the detailed definitions/text of V-001..V-020 are **not embedded in workspace documents**. Definitive closure of items whose closure depended on that prior report (notably V-001, V-005, V-009–V-013, V-015–V-017 as reported by validation 2) cannot be independently evidenced from repository docs alone without that report; this plan records their handling as addressed in rev. 4 corrections where determinable but does not claim undocumented closures. New findings N-001–N-003 are addressed by the corresponding current-doc corrections/re-validation cycle.

- **B0 Baseline evidence & validation setup — completed (evidence/checklist).** Artifact inventory (gitignored/local-only), hashes/paths/state, environment capture template, fixed config/period capture tuple, and blocker checklist produced. Full target sign-off details, authoritative fixed config/R4 evidence, and golden baseline remain pending.

- **B1 Spec alignment, config compliance & maintainability actions — partially complete; R4 pending verification; reconciliation pending clarification.**
  - **Spec/plan alignment (rev. 4):** corrected in this document set to enforce VBA-only baseline and document open clarifications.
  - **R4 (D13):** **reported** resolved locally in gitignored `data/theiss/treport.xlsm` (recommended config change to blank/false for `report_m.group_rows`) with pre-change backups recorded in prior status notes; **not committed** (gitignored). **Current status:** unverified in repository/for acceptance — authoritative compliant state requires verification (pending). Prior notes claiming “R4 remediated” must be read as “reported local remediation, unverified”.
  - **VBA maintainability:** opaque binary snapshot `mko_treport/src/vbs/vbaProject.bin` committed (`81ce32e`); readable modules exported and committed (`a8895aa`): `mko_treport/src/vbs/export/{Sheet1.cls,Sheet2.cls,ThisWorkbook.cls,xlconcat_macro.bas}`. Exports/snapshot support inspection/maintenance only; **runtime-vs-export parity unverified** (OQ-T4 residual).
  - **VBA-only reconciliation:** non-compliant Python treport implementation modules were committed in 4e60165, 148b668, 3d7b4c2, 93e1ae4 and removed from `mko_treport/src/scripts/` in e040c8c. `mko_treport/src/scripts/` currently contains no treport implementation modules, but tracked non-VBA test artifacts remain under `mko_treport/tests/treport_*` (observed tracked dirs include `treport_config`, `treport_files`, `treport_filters`, `treport_filter_application`, `treport_refresh`). **Disposition/reconciliation and acceptance impact are needs clarification** (V-002,V-003); these artifacts are not rev. 4-compliant implementation evidence.

- **B2 Technical investigation (OQ-T1, OQ-T4, controls/fields) — deepened with exported VBA source; gates/clarifications remain open.** Package/XML + VBA-source reconciliation completed (appendix retained). Confirmed: single control `RunXlconcat`/`xlconcat_macro.bas` is a legacy concatenation macro, **not** rev. 4 `Refresh_data`/`Create_report` (full functional mismatch; runtime-vs-export parity unverified). Pivot/PQ evidence partially identified per data file; field/caption/semantics evidence indicates `Year`/grouped-`Years` differences vs configured `Years` and a possible missing configured category token in inspected extracts. **OQ-T1, OQ-T2, OQ-T3 remain open; OQ-T4 substantially informed but not fully closed; FR-RUN-1/field/token clarifications required. No implementation approach chosen.**

- **B3 Golden baseline & target-environment confirmation — partially started, blocked for sign-off.** Known target class: Windows 10–11, Excel LTS 2021+, dataset `data/theiss`. Local Office16 presence evidenced; full target build/macro-trust/PQ details, verified fixed config/R4 state, and golden expected output for fixed configuration/period **not yet provided** (OQ-B7 pending).

- **B4–B7 prior non-VBA implementation artifacts/commits (historical) — NOT compliant with rev. 4 VBA-only constraint; pending reconciliation/disposition.**
  - **B4 / 4e60165 (`feat(treport-config): add B4 config contract validation`):** added Python config modules under `mko_treport/src/scripts/treport_config/` and tests under `mko_treport/tests/treport_config/`. Implementation modules removed from `src/scripts` in e040c8c; tracked tests may remain. **Status:** not rev. 4-compliant implementation; not acceptable as compliant implementation/acceptance evidence. Disposition needs clarification.
  - **B5 / 148b668 (`feat(treport-filters): add B5 filters DSL parsing and resolution core`):** added Python filter modules/tests. Removed from `src/scripts` in e040c8c; tracked tests may remain. **Status:** not rev. 4-compliant; pending reconciliation/disposition.
  - **B6 / 3d7b4c2 (`feat(treport-files): add B6 file resolution and session policy core`):** added Python file/session modules/tests. Removed from `src/scripts` in e040c8c; tracked tests may remain. **Status:** not rev. 4-compliant; pending reconciliation/disposition.
  - **B7 / 93e1ae4 (`feat(treport-refresh): add B7 refresh orchestration core`):** added Python refresh modules/tests. Removed from `src/scripts` in e040c8c; tracked tests may remain. **Status:** not rev. 4-compliant; pending reconciliation/disposition.
  - **No VBA-compliant B4–B7 implementation exists as evidence in this plan.**

- **B8–B12 — not started as rev. 4-compliant VBA implementation.** Dependent on B2/B4–B7 compliant VBA progress and open gates/clarifications.

- **B13–B14 — gated.** Scenario planning may proceed; execution/sign-off/readiness blocked by missing VBA-compliant implementation evidence, OQ-B7/OQ-T1–T4 closure, R4 verification, and required clarifications/reconciliation outcomes.

- **Next (DAG, VBA-only, gates respected):** (1) resolve/obtain clarifications and evidence for R4 verification, FR-RUN-1 controls/entry-point disposition, per-file field identities/grouped-`Years` semantics, category token alignment, whitespace handling, and non-VBA artifact disposition (B1–B3/B13 planning inputs); (2) continue B2/OQ-T1–T4 investigations and golden baseline acquisition; (3) start/continue only **VBA-compliant** design/implementation in `treport.xlsm` for B4–B12 where prerequisites and evidence permit, keeping pivot/refresh/copy/control behaviors gated as required; (4) do **not** perform or resume non-VBA implementation/automation and do not claim correctness sign-off without target Excel/VBA evidence and required closures.

## B2. Technical investigation — findings/status (OQ-T1, OQ-T4, controls/fields) — normalized to rev. 4

- **Status:** Investigation deepened (read-only). Exported VBA source in `mko_treport/src/vbs/export/{xlconcat_macro.bas,ThisWorkbook.cls,Sheet1.cls,Sheet2.cls}` analyzed and reconciled with package/XML evidence and local gitignored artifacts. No implementation behavior assumed beyond exported source/XML facts; runtime-vs-export parity unverified; no final pivot/filter/copy/refresh/VBA/entry-point approach chosen.

- **Evidence base (VBA):** `xlconcat_macro.bas` (`APP_NAME="xlconcat"`). Entry point `Public Sub RunXlconcat()` assigned to `run` sheet control (“Button 1”, `[0]!RunXlconcat`). `ThisWorkbook.cls`, `Sheet1.cls` (`settings`), `Sheet2.cls` (`run`) contain no procedures. Binary scan: HIT `RunXlconcat,run,settings,delimeter,ReadOnly,Save,output,input_path`; MISS rev. 4 operations/tables/fields `Refresh_data,Create_report,general,reports,data_files,placeholders,articleLevel4Name_clean,targetAudience,filters,max_*,PivotTable,PivotCache,RefreshAll,Power,Query,clean_data_full,report_m/w/wclip`.

- **Confirmed VBA architecture/behavior (`xlconcat_macro`, legacy):** generic workbook/worksheet concatenator, **not** rev. 4 treport reporting. Reads sheet `settings` `ListObject("settings")` (`parameter,value`) and optional `ListObject("exclude_names")`; uses params incl. `new_file_delimeter,keep_col_hidden`; resolves input folder/output file (not rev. 4 `general`/D2 semantics), collects `.xlsx`/`.xls` alphabetically, opens sources read-only, never refreshes PQ/pivots, never saves sources, writes single output workbook/sheet `plans` via value/format copy logic, saves output (fails if output exists). Diagnostics/state handling are not the rev. 4 contextual taxonomy. **Functional mismatch with rev. 4 D1–D17/FR-CFG/FILE/REF/FLT/RPT/NAME/RUN is confirmed for this implementation; runtime parity with current binary unverified.**

- **Confirmed treport config/artifacts (evidence snapshot):** `general` `input_path,output_path,delimeter`; `reports` 6 rows (`articleLevel4Name_clean,targetAudience,name`) including trailing space in `DOLGIT ` and split token `ЗУБНАЯ ПАСТА&ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`; `data_files` 3 rows with `filters` using `Month Number=(1-max)|Years=max` / `Week Number=(1-max)|Years=max` and baseline-observed `report_m.group_rows=TRUE` (R4 compliance pending verification as noted above); `placeholders` `max_weeknum,max_monthnum,max_year`. No evidenced `Refresh_data`/`Create_report` controls.

- **Confirmed data workbook/pivot/PQ evidence (per file, XML, partial):** each data file has `clean_data_full` and `PivotTable1` with external cache `Query - clean_data_full` (Mashup, async/background connections, user/machine-specific paths). Layouts/field availability differ:
  - `report_m`: cache has `Year` (singular), `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience`; **no `Years` cache field observed**.
  - `report_w`: cache has `Year` (singular), `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience`; **no `Years` cache field observed**.
  - `report_wclip`: cache has `Year`, grouped `Years` field/items, `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience`.
  Decoded cache items (inspected extracts): `targetAudience` includes expected values; configured token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as distinct `articleLevel4Name_clean` cache item (risk/mismatch, needs clarification).

- **OQ-T4 reconciliation status:** substantially informed by exported source/bindings (legacy `RunXlconcat` only; no rev. 4 entry points/logic evidenced), but **not fully closed** due to unverified runtime-vs-export parity and unresolved FR-RUN-1 control/entry-point disposition.

- **OQ-T1/T2/T3 status:** partially informed, still open. Per-file pivot/cache/PQ indicators are evidenced; runtime pivot object identity/selection semantics, authoritative period-field identity/semantics (incl. grouped `Years` mapping), available-value sourcing, exact presented copy bounds, PQ refresh completion, and formatting/copy fidelity remain unverified and require target Excel + golden baseline.

- **Risks/impacts for rev. 4 (confirmed/high):** (1) no compliant VBA implementation for `Refresh_data`/`Create_report` and D1–D17/FRs evidenced; legacy macro is functionally divergent. (2) Run-control/entry-point reconciliation required (FR-RUN-1). (3) Config/field executability risk: `filters` reference `Years` while caches indicate `Year`/grouped-`Years` differences — validate/align, no guessed mapping. (4) Possible category token/data mismatch requiring fail-fast or clarified alignment. (5) PQ refresh/portability/async behavior unverified. (6) Pivot/layout variability affects D7/D8/D15 and verification. (7) R4 compliance remains pending verification despite reported local remediation. (8) Prior non-VBA artifacts are not compliant evidence and need disposition clarification. No implementation approach chosen.

---
**Brief notes (plan handoff):**
- **Recommended filename remains:** `.ai/plans/001-treport-execution-plan.md` (this is the complete corrected replacement plan).
- **Baseline and constraint:** spec **rev. 4** is authoritative; **VBA-only implementation** is mandatory (no Python/other programming/scripting languages). This plan prohibits non-VBA implementation/automation and documents non-compliant historical artifacts as open reconciliation items.
- **Local/gitignored artifacts:** `data/theiss/*` and `data/theiss/treport.xlsm` are gitignored/local-only; the plan requires verified, reproducible evidence capture for R4 compliance, golden baseline, and validation without assuming repo-tracked availability.
- **Unresolved (must remain explicit):** OQ-B7 golden baseline; OQ-T1–T4 technical validations; R4 authoritative compliance verification; FR-RUN-1 control names/bindings/entry points and `RunXlconcat` disposition; per-file `Year`/`Years`/grouped-`Years` identity/semantics and `filters`/`max_*` alignment; category token alignment; configured-name whitespace handling policy; disposition/reconciliation of remaining non-VBA artifacts (incl. tracked tests). No guessing.

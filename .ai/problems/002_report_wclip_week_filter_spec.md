# Analytical specification: report_wclip weekly report shows extra columns (suspected incorrect filters) — **generalized filter model**

- **Spec ID:** 002_report_wclip_week_filter
- **Status:** Draft analytical specification (awaiting Product Owner decisions on ambiguous business/behavior questions in §6); not an implementation plan by itself.
- **Date:** 2026-11-16 (rev. 2 — incorporates PO direction: filters must be fully generic, not hardcoded to Week/Year/Month)
- **Supersedes/extends:** This document refines the earlier draft intent for P02 by replacing hardcoded period-field assumptions with a **general, field-driven filter model** defined below.
- **Source request / context:**
  - `.ai/problems/P02.md` (original problem: `report_wclip` built incorrectly; hypothesis: `filters` wrong; symptom: extra columns 54–350 in `output/DOLGIT W30-50.xlsx`, sheet `хроно по неделям 2025`)
  - User clarification (2026-11-16): “Filters must not be hardcoded to weeks/months/years. User specifies an arbitrary field and an arbitrary value or range. Key: the field must exist in the PivotTable; if not, error. Principle…” (see §2.2 and §3).
  - Related baseline: `.ai/problems/001_treport_spec.md` (rev. 6) — some earlier decisions describe `data_files.filters` examples with `Week Number`/`Year`/`Month Number`; those are **examples/legacy snapshots**, not the target design. The target design is the **general model** in this spec.
- **Primary evidence reviewed in-repo:** `P02.md`, `001_treport_spec.md`, `docs/02-build-and-deployment/vba-build.md`, VBA exports `modTreportFilters.bas`, `modTreportRun.bas`, `modTreportPivot.bas`, `modTreportConfig.bas` (structure/logic scan). No changes made to VBA/workbooks in this analysis.
- **Out-of-repo/local Excel artifacts (not opened/edited here):** `data/theiss/treport.xlsm`, `data/theiss/base/report_wclip.xlsx`, `data/theiss/output/DOLGIT W30-50.xlsx` (sheet `хроно по неделям 2025`), related base/monthly/weekly sources. All Excel facts below that depend on those files are **to be verified** by evidence tasks (R-1/R-2).

---

## 1. Problem statement (normalized)

- Reports are built from `treport.xlsm` settings and VBA macros. The `report_wclip` path produces incorrect output: in `output/DOLGIT W30-50.xlsx`, sheet `хроно по неделям 2025`, extra columns appear (reported **54–350**) instead of the intended column set implied by correct filtering.
- Root-cause investigation must start from Excel evidence (workbooks/pivots/config) and code paths, but the **filter subsystem must be generalized**: it must not assume specific fields like “Week/Month/Year” are special beyond being names present in the PivotTable.
- The user-provided concrete example from P02 (“field `Week of Year`, range `202501–202553`, weeks 1–53 of 2025”) is treated as **one possible instance** of the general model, to be validated against the actual pivot field name and value format in `report_wclip.xlsx` (see §6).

---

## 2. Business goal, scope, constraints, assumptions

### 2.1 Business goal
- Restore correct reporting for `report_wclip` (and agreed scope) so outputs contain only intended columns/rows per correctly applied filters, with **no unintended extra period/value columns** caused by missing/incorrect filtering.
- Establish a **flexible, generic filtering mechanism** configurable per `data_files` row: users specify arbitrary pivot fields and filter expressions; the system applies them deterministically and fails clearly when misconfigured or when referenced fields do not exist in the PivotTable.

### 2.2 Scope (in scope)
- Investigate and specify root cause of extra columns in the reported output (Excel evidence + VBA trace).
- Specify the **general filter model** (this spec) as the target behavior for `data_files.filters` (and clarify interactions with `reports.articleLevel4Name_clean` / `targetAudience` only where relevant to the defect; note: current code treats report-level `articleLevel4Name_clean` with a special “skip report if no matches” rule in some paths—PO must confirm whether that exception remains or is generalized; see §6).
- Define conceptual tasks, required Researcher investigations, PO questions, risks, and Definition of Ready for implementation.

### 2.3 Out of scope
- Implementing code/workbook changes in this analysis.
- Changing source data except via documented working-copy workflows during later implementation/verification.
- Redesign beyond what is needed except adopting the general filter model stated here.

### 2.4 Constraints (project)
- VBA export sources under `mko_treport/src/vbs/export/`; any future changes follow repo build/verification workflow (`build_treport_vba.ps1`, backups/working copies, L2 manual VBE Compile, L3 if logic changes) per project rules/docs.
- Final runtime remains **VBA-only in Excel** for end users (auxiliary tools allowed for build/inspection only, not runtime dependencies).
- **No hardcoding** of field names or semantics to “week/month/year” in the filter engine design. All field references come from configuration and must resolve against the PivotTable’s **available fields** (see §3).
- Do not invent Excel/pivot facts; unresolved items go to Research tasks and PO questions.

### 2.5 Assumptions
- A1. The reported symptom (extra columns 54–350 in the named sheet/output) is real and reproducible in the current environment/context (to be confirmed by CT-1/R-1 on working copies).
- A2. The fix may require changes to `data_files.filters` text, and/or VBA parsing/application logic, and/or both, guided by evidence and PO decisions.
- A3. “PivotTable” in this spec means the primary pivot used for the given `data_sheet` row in the `data_file` workbook (consistent with current behavior: e.g. `PivotTable1`/primary selection logic in code)—exact resolution rules unchanged unless stated otherwise.

---

## 3. Confirmed facts and current code signals (evidence-based)

### 3.1 Confirmed (repo)
- VBA filter parsing/resolution is implemented in `mko_treport/src/vbs/export/modTreportFilters.bas`; application helpers in `modTreportPivot.bas`; orchestration in `modTreportRun.bas`.
- Current implementation contains **hardcoded assumptions** in logic (to be removed/generalized by design): e.g. `FindClauseIndex(pf, "Week Number")`, `FindClauseIndex(pf, "Year")`, resolution of `MaxWeek`/`MaxMonth` tied to those names, and `ResolveMinMaxAndYear` builds a `ResolvedContext` with year/week/month maxima. This is inconsistent with the new general model and must be refactored conceptually (details in Research tasks).
- `modTreportPivot.GetPivotFieldRequired` treats a field as available if present in `pt.PivotFields` (P01/D7 spirit in comments); missing → error. Placement rules (active vs inactive → move to Filters/page area in current behavior for some cases) exist today—behavior for the general model must be clarified (see §6).
- `data_files.filters` is read as required (non-empty) in config loading (per `modTreportConfig.bas` behavior described in prior spec). Multiple clauses joined by `|`.

### 3.2 Confirmed from P02 (user report)
- Suspected bad filter context for `report_wclip`: example text mentioned “`Week of Year=(202501-202553)`” in P02 narrative; this is a **hypothesis/example** under the old mental model. Under the new general model, the user must be able to express any field+expression that matches the pivot.

---

## 4. Target filter model (general, authoritative for this spec)

### 4.1 Core principle (must hold)
1. **No hardcoded field semantics in the filter engine.** The engine must not special-case “week”, “month”, “year”, or any business-named field except through generic parsing rules below and the fact that all referenced fields must exist in the PivotTable.
2. **Field source of truth:** For each `data_files` row, filters apply to the PivotTable on `data_sheet` in the opened `data_file` workbook. A filter field is valid **only if it exists in that PivotTable’s available fields** (i.e. discoverable via the PivotTable API/model as an available/pivot field in that context). If any referenced filter field does **not** exist → **fail fast** with a clear error (file, sheet, data_file row, field name).
3. **User-controlled expressions:** For each clause the user specifies:  
   **`<field name> = <expression>`**  
   where `<field name>` is the exact name/caption as recognized by the PivotTable field lookup rules of the implementation (to be defined precisely after evidence; see §6), and `<expression>` is one of the forms in §4.3.
4. **Multiple filters:** Clauses are joined by the pipe character **`|`** (same separator style as before unless PO changes it). All clauses in the set apply (logical AND in effect for pivot selection constraints) unless explicitly stated otherwise by PO.
5. **`general.delimeter` role:** The existing `general.delimeter` is used to split **multi-value lists inside an expression** (form `(value1<delim>value2…)`). It is **not** the clause separator (clauses are `|`). This matches current behavior and must remain explicit; any change requires PO decision.

### 4.2 Grammar (normative description)
- A `filters` string is a sequence of one or more clauses separated by `|`.
- Whitespace around `|` and around `=` should be trimmed for parsing (consistent with current `Trim`/normalization behavior), unless PO requires otherwise.
- Each clause: `<field> = <expr>` where `<field>` is the field identifier string (may contain spaces) after trimming.
- Empty/invalid clauses, duplicate handling, and error messages: follow fail-fast rules (see §4.6).

### 4.3 Expression forms (`<expr>`)
The implementation must support these forms (general model). Exact parsing disambiguation rules must be documented in implementation spec; below are the required semantics.

| Form | Syntax (conceptual) | Meaning | Notes |
|------|---------------------|---------|-------|
| **Single value** | `value` | Select/filter to that single value (literal match to a pivot item/value representation as defined by implementation). | For non-numeric text values, exact match after trim per existing normalization rules unless PO states case-insensitivity policy (clarify in §6). |
| **Numeric maximum token** | `max` | Use the **maximum available numeric value** for that field in the current resolution context (as defined by implementation after applying any prior constraints/ordering rules the system already uses). | **Only valid when the field’s comparable values are numeric** (or when the system can deterministically derive a numeric max). If `max` is used on a non-numeric field → **fail fast** with clear error. |
| **Range (inclusive)** | `(start-end)` | Inclusive range from `start` to `end`. | Applies to **numeric** ranges in the general model unless PO explicitly requests string/lexicographic ranges (risky) — see §6. `start` and `end` may be numeric literals or the token `max` **only on the right end in practice**? Or allowed in general? Current code allows `min`/`max` as bound tokens in some forms; PO must confirm allowed bound forms (see §6 Q2-G1). Comparisons numeric when both bounds resolve numerically. |
| **Multi-value list** | `(v1<delim>v2<delim>...)` using `general.delimeter` | Select/include all listed values that are valid/matchable per rules below. | Tokens trimmed. Matching rules: **prefer exact match to pivot item caption/name after trim** (consistent with current behavior) unless evidence shows different model; any unmatched **required** list token that makes the selection impossible should follow fail-fast unless a specific exception is approved (current code has a special exception only for report-level `articleLevel4Name_clean` multi-value “skip report if none match” in some paths — this must not be silently carried to `data_files.filters`; PO decides in §6). |

**Disambiguation guidance (to be made precise in code spec):**  
- If the expression contains `general.delimeter` → treat as **list** (after stripping optional outer parentheses if the current grammar requires them; current behavior uses parentheses around lists in examples like `(1-12)` ranges vs lists `(v1&v2)`). PO/code must align with existing accepted forms or document migration.  
- If the expression is of the form `( ... - ... )` with a single `-` separating two bound tokens in the range context → treat as **range** (current behavior).  
- Otherwise → **single value** (or the literal `max` token as a special single-value/resolution case).

### 4.4 Numeric `max` semantics (critical)
- `max` means “the maximum value among **available/eligible values of that field** in the context used for resolution”, resolved deterministically at runtime against the PivotTable field data/items as implemented.
- The **context** for “eligible values” must be defined (current code had dependency-aware behavior tied to `Year` then resolving `Month Number`/`Week Number` within that year context). Under the **general model**, there is **no implicit dependency graph** unless:
  - (a) the implementation derives it from clause order is explicitly defined, OR
  - (b) PO defines explicit dependency rules, OR  
  **(Recommended default)** `max` is resolved from the **full set of available numeric items for that pivot field** on the `data_sheet` pivot at the time of resolution (before applying other clauses of the same `data_files.filters` batch in a way that would exclude them?) — this must be clarified in §6 to avoid surprising behavior. The previous “resolve year first, then month/week within year” was a hardcoded special case and is **not allowed** in the general model unless explicitly reintroduced as a configured rule (not recommended).

### 4.5 What is NOT allowed (anti-patterns)
- Hardcoding any logic like “if field is Week Number then…” / “if field is Year then…” in the filter engine. All behavior flows from field existence + expression parsing + generic item/value matching/resolution.
- Assuming a particular calendar encoding (`202501` etc.) inside the engine; that is data-driven and appears only as literal values in config when used.
- Silently ignoring a referenced field that is not in the PivotTable (must error).

### 4.6 Error handling (general)
- Any referenced filter field **not present in the PivotTable’s available fields** → **fail fast** with contextual error: operation, `data_file`, `data_sheet`, `filters` text/context, offending field name.
- Invalid/unparseable clause or expression → **fail fast** (contextual).
- `max` used but field has no numeric values / cannot be resolved → **fail fast**.
- Range with `start > end` after resolution → **fail fast** (or clearly defined behavior — PO may prefer auto-swap; **not recommended** without explicit approval; default fail fast).
- Expression references values that do not match any pivot items and the clause is required → **fail fast** (exceptions only if explicitly approved by PO, e.g. no approved exception for `data_files.filters` beyond what is specified).
- Filter combination yields **no data** after applying all `data_files.filters` and then report-level filters as per current orchestration rules → behavior follows current agreed rules in `001`/code context unless changed by PO (currently fail fast for no-data in `Create_report` path in many cases; report-level `articleLevel4Name_clean` has a special “skip report if none match” warn-and-continue in current code — see §6 for whether to keep/generalize).

### 4.7 Interaction with report-level filters (`reports` table) — must be clarified
- `reports` currently has `articleLevel4Name_clean`, `targetAudience`, `name`. Current code applies report-level filters with special handling for `articleLevel4Name_clean` (matched subset if any match; if none match → warn and skip that report, continue others) and strict behavior for `targetAudience` in practice? (per comments in `modTreportFilters`/`modTreportRun`).  
- **PO decision required (§6):** Do we keep this special-case behavior for `articleLevel4Name_clean`/`targetAudience`, or do we generalize it to “any report-level multi-value field uses the same rules as configured”? The new general model applies primarily to `data_files.filters`; report-level fields are a separate surface area—do not silently broaden scope without approval.

---

## 5. Examples (illustrative; adapt to real pivot field names/values after Excel evidence)

> These examples assume the field names and value formats exist in the target PivotTable and use `general.delimeter = &` where lists appear (illustrative). Always verify against actual pivot items.

### 5.1 Single value (non-numeric)
- `Brand = Acme`
- `Region = EMEA`

### 5.2 Single numeric value
- `StoreId = 42`

### 5.3 Numeric `max`
- `Week Number = max`  (if such a field exists; means maximum numeric week item available in the resolution context defined by implementation/PO)
- `PriorityLevel = max` (only if numeric)

### 5.4 Numeric range (inclusive)
- `Week Number = (1-53)`
- `PriceBand = (10-99)`
- `Week Number = (1-max)` (from 1 to maximum numeric available for that field in defined context)

### 5.5 Multi-value list (using `general.delimeter = &`)
- `Category = (A&B&C)`
- `Segment = (Retail&Online)`

### 5.6 Multiple clauses joined by `|`
- `Year = 2025 | Week Number = (1-53)`
- `Region = EMEA | Brand = Acme | Week Number = max`
- `articleLevel4Name_clean` is **not** in `data_files.filters` by current design (it’s in `reports`) — but if PO later allows generic report filters, the same syntax applies.

### 5.7 Example tied to P02’s reported case (hypothetical, pending Excel confirmation)
If Excel shows a pivot field literally named **`Week of Year`** with coded numeric/string items like `202501`..`202553`, then under the **general model** (no special code for the name):  
- `data_files.filters = Week of Year = (202501-202553)`  
If instead Excel shows separate fields **`Year`** and **`Week Number`** (numeric 1–53):  
- `data_files.filters = Year = 2025 | Week Number = (1-53)` (or `Year = max | Week Number = (1-53)` depending on rules)  
**These are mutually exclusive possibilities until R-1 confirms the real model.** The engine must support either case as long as the named fields exist in the PivotTable.

---

## 6. Gray areas and Product Owner questions (resolve before implementation)

### 6.1 Gray areas (separate from confirmed requirements)
- GA-G1. **`max` resolution context:** Should `max` be computed from **all available numeric items** of that field in the pivot at filter-application time, ignoring other `data_files.filters` clauses? (Recommended default for general model.) Or must it respect some dependency order? Current hardcoded behavior resolved year first then week/month within year context—this is removed conceptually; PO must approve any new implicit dependency rules.
- GA-G2. **Range bounds with `max`:** Is `max` allowed only as the **end** bound of a range (e.g. `(1-max)`) or also as start (unusual)? Current behavior allowed resolving end from clause in some paths; confirm allowed forms.
- GA-G3. **Data types:** Do we support ranges only for numeric fields? (Recommended: yes, fail fast otherwise.) String/lexicographic ranges are error-prone in pivots—avoid unless required.
- GA-G4. **Matching rules:** Exact match after trim (current). Case sensitivity? (Current appears case-sensitive in string comparisons in code.) PO preference?
- GA-G5. **Multiple clauses:** Are all clauses ANDed? Can the same field appear twice in `data_files.filters`? (Current code rejects duplicate field names in parsed clauses—see `modTreportFilters`: duplicate → fail 2312). Confirm this remains the rule.
- GA-G6. **Report-level filters:** Keep current special rule for `articleLevel4Name_clean` (skip report if no configured values match any pivot items, warn, continue others) vs strict fail-fast? Keep strict for `targetAudience`? This is separate from `data_files.filters` general model but affects the same defect workflow.
- GA-G7. **Placement/orientation of filter fields in PivotTable:** Current code may move inactive fields to page/Filters area and sort in some cases; for a **generic** set of user-chosen fields, is this behavior still desired? Or must the code only filter fields that are already placed appropriately and fail if not placeable? (P01/D7 rules in 001 were specific; general model needs clarification.)
- GA-G8. **“Field name” identity:** Must match the PivotField name/caption exactly as seen by Excel API in the current implementation? Are localized names a concern? (Documented assumption: match per current pivot field lookup behavior.)
- GA-G9. **Interaction with `Refresh_data`:** Does filter definition change require refreshed data? (Usually no for filter logic, but evidence may show stale pivots—out of scope of model, but note for verification.)

### 6.2 Product Owner questions (answer required)

**Q2-G1 — `max` resolution context (required)**
- Options:
  1. **(Recommended)** Compute `max` (and numeric bounds resolution for ranges, where applicable) from **all available numeric items of that field** in the current PivotTable state for the `data_files` row, **independent** of other clauses in the same `filters` string (no implicit cross-field dependency like old Year→Week logic).
  2. Allow **ordered/dependency rules** defined in config (not recommended now). Describe if needed: ____
  3. Other: ____
- Your decision: ____

**Q2-G2 — Allowed bound forms in ranges**
- Options:
  1. **(Recommended)** `(start-end)` where `start` and `end` are numeric literals **or** `end` may be `max`; `start` may be `min` only if you explicitly want it (current code had support patterns in parts of logic?) — clarify: allow `start = min`? Allow both `min` and `max`? Default recommendation: allow numeric literals and allow `max` as **end** only (and optionally `min` as start meaning “smallest available”) **only if** clearly defined; otherwise only numeric and `max` at end to reduce ambiguity.
  2. Allow only numeric literals (no `min`/`max` tokens in ranges) — users write explicit numbers.
  3. Allow any combination of numeric, `min`, `max` (define semantics): ____
- Your decision: ____

**Q2-G3 — Numeric-only ranges**
- Options:
  1. **(Recommended)** Ranges `(start-end)` are valid **only for fields whose items are numeric** (or resolvable numerically). Non-numeric fields with range expressions → **fail fast**.
  2. Allow string/lexicographic ranges (not recommended): ____
- Your decision: ____

**Q2-G4 — Unmatched list values in `data_files.filters`**
- Options:
  1. **(Recommended)** Any required value/token that cannot be matched to an available pivot item in a `data_files.filters` clause → **fail fast** (consistent with strict current behavior for `data_files.filters`).
  2. Ignore unmatched tokens silently (risky): ____
  3. Warn and continue if at least one matches (like current `articleLevel4Name_clean` report-skip exception?) — **not recommended for `data_files.filters`** unless justified: ____
- Your decision: ____

**Q2-G5 — Report-level `articleLevel4Name_clean` / `targetAudience` exceptions (keep or change?)**
- Options:
  1. **(Recommended, minimal change)** Keep current behavior: for `articleLevel4Name_clean`, if none of the configured multi-value tokens match → warn and **skip that report** (continue others). For `targetAudience`, strict fail-fast if required value not matchable (as today). `data_files.filters` remains strict per Q2-G4.
  2. Make everything strict (no report-skip exception) — may break existing configs; requires review.
  3. Generalize exceptions (describe): ____
- Your decision: ____

**Q2-G6 — Field placement/orientation when a filter field is not in Filters area**
- Options:
  1. **(Recommended, preserve flexibility but be safe)** For any referenced field that exists in available fields but is placed in Rows/Columns (active) — keep placement, apply filtering in place (per current P01 intent for active fields). For fields that are available but not placed (hidden/inactive) — move to Filters/page area and apply selection (current behavior in code paths that do this), with clear errors if not possible.
  2. **Strict:** Only allow filtering fields that are already in the Filters/page area; otherwise fail fast with instructions to configure the pivot. (More predictable, less “magic”, may break current workflows.)
  3. Other behavior: ____
- Your decision: ____

**Q2-G7 — Same field twice in `filters`**
- Options:
  1. **(Recommended)** Disallow duplicates → **fail fast** (current behavior). One clause per field per `data_files` row.
  2. Allow last-wins or merge semantics (complex, not recommended): ____
- Your decision: ____

**Q2-G8 — For the reported `report_wclip` defect, what is the confirmed intended filter after Excel evidence? (to be filled after R-1)**
- Once R-1 identifies actual pivot field names and value formats, answer with the exact `data_files.filters` string you want, e.g.:
  - `Week of Year = (202501-202553)`  OR
  - `Year = 2025 | Week Number = (1-53)`  OR
  - Other: ____

**Q2-G9 — Scope of fix for this defect**
- Options:
  1. **(Recommended)** Fix only the reported `report_wclip` path/output (`DOLGIT W30-50.xlsx`, sheet `хроно по неделям 2025`) and ensure the general model is specified/implemented without breaking `report_m`/`report_w` behavior except where configs must change.
  2. Broader refactor of all `data_files.filters` to the new model in the same change (higher risk; needs full regression plan).
- Your decision: ____

---

## 7. Required Researcher tasks (must be launched)

Each Researcher task must require: investigate current architecture relevant to the problem; investigate modern best practices for the general model in this Excel VBA context; identify feasible approaches (top 2–3 if multiple) with **recommended option and rationale**; cite evidence and state what was not verified.

- **R-1 (Excel evidence, highest priority):** Confirm in working copies: `treport.xlsm` `data_files` row for `report_wclip` (exact `filters`, `data_sheet`, `report_sheet`), `reports` rows tied to the reported output, and in `base/report_wclip.xlsx` pivot/cache **all available field names/captions** (especially any week/year/period fields), item formats/ranges, and the column layout of current wrong output vs expected (map columns 54–350 to pivot axes/fields). Also inspect `report_m.xlsx`, `report_w.xlsx` for comparison.
- **R-2 (VBA trace + refactor impact):** Trace `Create_report` path for `report_wclip` through `modTreportRun`/`Filters`/`Pivot`/`Extract`/`Output`; identify every hardcoded special case (`Week Number`, `Year`, `MaxWeek`/`MaxMonth` resolution, dependency ordering) and propose a **minimal refactor** to support the general model (§4) while preserving current behavior where configs remain compatible (or documenting required config migrations).
- **R-3 (Design options):** Top 2–3 approaches to implement general `data_files.filters` (parsing, `max`, ranges, lists, matching, placement/orientation, error handling) with recommendation. Must address Q2-G1–Q2-G7 trade-offs and backward compatibility with existing `report_m`/`report_w` configs.
- **R-4 (conditional):** If VBA changes proposed, build/deployment checklist using `build_treport_vba.ps1`, module preservation, L2/L3 steps, and verification plan (no golden workbook).

---

## 8. Research review guidance
Reject/rework if outputs lack evidence, ignore the **no-hardcoding** requirement, omit top 2–3 approaches + recommendation when applicable, or contradict confirmed constraints (VBA-only runtime, backups/working copies, modular exports).

---

## 9. Consolidated analytical specification

### 9.1 Problem statement
`report_wclip` output contains unintended extra columns (reported 54–350) on `хроно по неделям 2025` in `DOLGIT W30-50.xlsx`, likely due to incorrect or misapplied filters. The filter system must be **generalized**: arbitrary pivot fields and expressions (`value`, `(range)`, list via `general.delimeter`), numeric `max`, multiple clauses joined by `|`, with **hard fail if a referenced field is not in the PivotTable**. The old `Week/Year/Month` hardcoded logic is not acceptable as the target design.

### 9.2 Confirmed requirements (target)
- CR-G1. `data_files.filters` is parsed as `|`-separated clauses `<field> = <expr>` with generic semantics in §4.
- CR-G2. Any field referenced in `data_files.filters` must exist in the PivotTable’s available fields for that `data_sheet`; otherwise fail fast with context.
- CR-G3. No hardcoded special treatment of any field name in the generic filter engine (business meanings are expressed only through config + pivot data).
- CR-G4. Support single value, numeric `max` (with rules per PO Q2-G1/Q2-G2/Q2-G3), inclusive numeric range `(start-end)`, and multi-value list `(v1<delim>v2…)` using `general.delimeter`.
- CR-G5. Fix the reported `report_wclip` defect per PO scope/acceptance after evidence (Q2-G8/Q2-G9).

### 9.3 Conceptual tasks
- CT-1 Excel evidence (R-1), CT-2 VBA trace/refactor mapping (R-2), CT-3 option design (R-3), CT-4 build/verify plan (R-4 if needed), CT-5 PO decision capture, CT-6 implementation handoff (after DoR).

### 9.4 Risks
- **Rsk-G1:** Over-refactoring breaks existing `report_m`/`report_w` configs. Mitigation: R-2 compatibility analysis + regression smoke checks.
- **Rsk-G2:** `max`/dependency semantics changes alter behavior unexpectedly. Mitigation: PO decision Q2-G1 + explicit tests.
- **Rsk-G3:** Excel field naming/caption differences cause lookup failures. Mitigation: document lookup rules from evidence; fail fast with clear messages.
- **Rsk-G4:** Range/list parsing edge cases. Mitigation: define grammar precisely in implementation spec + tests on copies.

### 9.5 Open questions
- All items in §6.2 (Q2-G1–Q2-G9) plus OQ-2A–OQ-2E from the prior draft remain open until answered/evidenced.

### 9.6 Definition of Ready (DoR) for implementation
Ready when: PO answers §6.2 recorded; R-1/R-2 evidence complete and root cause confirmed; approved design (R-3, R-4 if any) selected; acceptance criteria for agreed scope defined (including Q2-G8 exact filter string and expected output column behavior); verification workflow (build, L2, L3 if applicable) planned with backups/working copies; regression plan for other `data_files` rows included.

---

## Appendix A — P02 source (verbatim)
(unchanged from original P02; see `.ai/problems/P02.md`)

## Appendix B — Key repo references
- `.ai/problems/P02.md`, `.ai/problems/001_treport_spec.md`
- `mko_treport/src/vbs/export/modTreportFilters.bas`, `modTreportPivot.bas`, `modTreportRun.bas`, `modTreportConfig.bas`
- `mko_treport/src/scripts/build_treport_vba.ps1`, `docs/02-build-and-deployment/vba-build.md`
# Execution plan: 003 — report_wclip week filter / generalized `data_files.filters` model

- **Plan ID:** 003-report-wclip-week-filter
- **Status:** Audit-reviewed (2026-11-16). VBA exports show provisional generic-filter refactor direction in places, but **Spec 001 rev.7 is not fully implemented** (critical: `report_sheet` placeholders still legacy in code/config flow). Runtime verification (L1–L3, U1–U10) and workbook-level checks remain pending.
- **Date:** 2026-11-16 (revised from audit findings M1–M4, C1–C3)
- **Source specs/plans:** `.ai/problems/002_report_wclip_week_filter_spec.md` (rev.2), `.ai/problems/001_treport_spec.md` (**rev.7**, normative). Plan lineage: prior 003 drafts (B1/B2/B3) are superseded where they conflict with rev.7 and this audit-aligned plan.
- **Scope/constraints:** Semantic/task intent only. No speculative implementation. Runtime VBA-only in Excel; auxiliary tools non-runtime. VBA exports under `mko_treport/src/vbs/export/*`. Build via `mko_treport/src/scripts/build_treport_vba.ps1` on working copies/backups only. **L2 manual VBE `Debug → Compile VBAProject` required** to claim compile; L3 on copies for logic/run changes.

## 1. Current state (code-truth summary, audit-aligned)

**Realized in VBA exports (provisional, subject to runtime evidence):**
- `modTreportFilters`: generic clause parsing/classification/validation (`|`, `field = expr`, list/range/scalar, numeric-convertible/bound rules, duplicate-field rejection), per-field numeric-domain pre-resolution in current model, clause application in declared order, shared selection helpers for data/report filters.
- `modTreportPivot`: field lookup by `PivotFields`/available-field presence in current flow, placement intent (keep active in place; move hidden/unplaced to Filters/page with best-effort sort), field-scoped selection applying matched subset when any specified value matches and failing when none matchable in current implementation.
- `modTreportRun`: `Refresh_data` / `Create_report` orchestration, state capture/restore, apply `data_files.filters` once per `data_files` row then per-report filters; no warn-and-skip report path for `articleLevel4Name_clean` none-match observed in current flow (uniform fail direction). Report workbook/sheet assembly, sanitization/collision checks, static snapshot copy + formatting (group_rows no-op).
- `modTreportOutput` + `modTreportRun`/`modTreportFilters`: **legacy placeholder model still in effect** — templates/substitution use fixed `<max_year>`, `<max_monthnum>`, `<max_weeknum>` (or equivalent aliases) derived from filter/context representatives in current implementation. This **does not satisfy Spec 001 rev.7**.

**Audit gaps / corrections to prior narrative (M4, C3):**
- **D9 condition (1) missing-field fail-fast must be explicit:** for any applied filter (`data_files.filters`, `reports.articleLevel4Name_clean`, `reports.targetAudience`), if the configured field name is **not found among PivotTable available fields** → **fail fast** with clear context. Prior “Current state” summaries must not omit this condition even if implementation direction appears uniform elsewhere.
- **D9 conditions (2)/(3):** field exists but **none** of specified configured values matchable → **fail fast** (no warn-and-skip of reports); if **at least one** matchable → apply **matched subset** and proceed (unmatched configured tokens ignored).
- **D10:** `min`/`max` resolve against available numeric values of **the referenced field only** in the relevant pivot context for the row — **no** cross-field dependency-aware year-first scoping.
- **Placeholders / “relevant pivot/context” (M3, C3):** rev.7 requires placeholders `max_<PivotTable field name>` in `report_sheet` resolved to the **maximum numeric value** of that named field computed from available numeric-convertible items of that field in the relevant pivot/data context used for the processed `data_files` row **at the time/place resolution is performed by the implementation** (implementation-defined sequencing). Resolution is **independent of `data_files.filters`** clause derivation/representatives (not derived from `data_files.filters`). Unresolvable → fail fast.
- **D2 divergence (M2):** **Spec 001 D2 is normative** (base name in `input_path`, fixed precedence `.xlsx`, `.xlsm`, `.xlsb`; fail fast if no match or >1 match). Current code behavior is reported to diverge (e.g. recursive search, no precedence selection, fail on >1 match). This is a **known compliance gap** and must be resolved before claiming D2 compliance (do not treat divergence as an accepted spec change).

**Not evidenced / pending:** L1/L2/L3 results on working copies; `run` sheet controls/bindings (FR-RUN-1); FR-DEP repeatable build/import proof; live Excel confirmation of field/caption resolution, numeric/locale/item behavior, axis/layout effects (U2), and regression across all configured rows/reports.

## 2. Decisions locked (rev.7, must be reflected in implementation/plan)

- **D9 (uniform, all filters):** (1) missing field in PivotTable available fields → fail fast; (2) no matchable specified values → fail fast (no report-skip); (3) ≥1 matchable → apply matched subset.
- **D10:** `min`/`max` per referenced field only; no cross-field/year-first dependency scoping.
- **D11 / placeholders:** `report_sheet` only; tokens `<max_<PivotTable field name>>` (e.g. `<max_Year>`, `<max_Week of Year>`). Value = maximum numeric value of that named field in the relevant pivot/data context for the processed `data_files` row at implementation-defined resolution time/place, from available numeric-convertible items of that field. **Not derived from `data_files.filters`**. Fail fast if unresolvable. Legacy fixed names `max_weeknum/max_monthnum/max_year` are **not** the target model.
- **D2:** normative as in spec; do not silently weaken precedence/no-match/>1-match rules.

## 3. Required work (actionable, dependency-aware)

### T-REV7-PLACEHOLDERS — Implement rev.7 `max_<field>` placeholder contract (critical, M1/C2/C3)
**Goal:** replace legacy fixed placeholder handling with spec-compliant `max_<PivotTable field name>` resolution in `report_sheet`, independent of `data_files.filters`.
- **Semantic changes:** in `modTreportOutput` / `modTreportRun` (and remove/refactor dependent alias plumbing in `modTreportFilters`/`ResolvedContext` as needed), detect all `max_<…>` tokens in `report_sheet` templates and resolve each by the named PivotTable field (deterministic field lookup consistent with project rules) against the relevant pivot/data context for the processed `data_files` row at a defined resolution point. Compute maximum from available numeric-convertible items of that field in that context; fail fast with clear context if field missing/unusable, no numeric values, or maximum not determinable.
- **Remove legacy coupling:** do not source placeholder values from filter representatives/`data_files.filters` resolution chains. Remove or neutralize legacy `<max_year>/<max_monthnum>/<max_weeknum>` substitution/alias derivation as the primary mechanism (no silent fallback to legacy semantics except where explicitly required by a separately approved migration step — none assumed).
- **Template/config migration (mandatory, M1):** audit and update **every** `data_files.report_sheet` template (and any other in-scope placeholder-bearing naming targets per spec — `report_sheet` only) in affected workbooks/configs so legacy tokens are replaced with the correct `max_<PivotTable field name>` tokens matching actual pivot field names/captions in use. Treat this as an explicit checklist item per row/template; do not leave “may require changing” ambiguity. Record before/after template strings in verification evidence.
- **Preserve:** `Trim`, sanitization, 31-char limit, collision/idempotency fail-fast rules per D14/FR-NAME/FR-RPT; fail-fast on unresolved placeholders (D11/D12).

### T-D9-ALIGN — Enforce/confirm uniform D9 behavior (including missing-field case)
- Ensure implementation and error reporting cover **all three** D9 conditions for every applied filter: (1) field not in PivotTable available fields → fail fast; (2) field exists but none of specified values matchable → fail fast; (3) ≥1 matchable → matched subset applied. No warn-and-skip of reports for “no matches”.
- Keep D10 field-only `min`/`max` resolution (no cross-field/year-first dependency). Do not reintroduce legacy special-casing that contradicts D9/D10.

### T-D2-CONFORMANCE — Data-file resolution per spec D2 (M2, known gap)
- **Normative:** resolve `data_files.data_file` base name in `general.input_path` to `<data_file>.<ext>` with fixed precedence `.xlsx`, `.xlsm`, `.xlsb`; fail fast if no match or >1 match (deterministic).
- **Required action:** implement/conform code and tests/checks to D2 behavior and remove/avoid recursive or precedence-ambiguous resolution that violates spec. Until this is done and verified, plan/status must state “D2 not compliant in code; do not claim D2 compliance”.

### T-U2-REPORT_WCLIP — `report_wclip` filter/axis correctness (evidence-driven)
- Validate that intended `data_files.filters` for `report_wclip` (per decided model/config) constrains output to the intended period set (e.g. 2025 week period) without unintended extra columns (reported symptom 54–350). **U2:** determine with evidence whether filtering the target field constrains the axis driving expansion or whether axis/layout handling is also required; apply only minimal, evidence-backed, non-speculative changes if needed.

### T-VERIFY — Build and runtime verification (mandatory gates)
- **L1:** run `mko_treport/src/scripts/build_treport_vba.ps1` on a working copy; capture/report path and summary (SUCCESS/ISSUES/ERROR, VBOM/trust notes). No compile claim from L1 alone.
- **L2 (required for compile claim):** manual VBE **Debug → Compile VBAProject** on the working copy; record file path/date-time and pass/fail. If not done, state limitation explicitly.
- **L3 (on copies only, required for logic/run changes):** run `Refresh_data` / `Create_report` as applicable on copies; verify outputs and that sources unchanged per project rules.
- **Workbook-level:** confirm FR-RUN-1 (`run` sheet exactly two controls/entry points `Refresh_data`/`Create_report`, same-named VBA procedures, legacy `RunXlconcat` removed) — cannot be proven from exports alone.
- **FR-DEP (if exercised):** single-step auxiliary build/import on working copy is repeatable/idempotent, preserves module/document identities, no duplicates/corruption; document compile/validation outcome.

## 4. Regression / acceptance checks (audit-aligned, explicit)

Execute and record results (L3/evidence where applicable):

1. **Placeholder migration (M1/C2):** every `data_files.report_sheet` template in scope uses only rev.7 `max_<PivotTable field name>` tokens (no legacy `max_weeknum/max_monthnum/max_year`) and resolves to the correct max of the named field per D11.
2. **Placeholder independence (M3/C3, D11):** for each processed `data_files` row, `max_<field>` values are computed from the named field in the relevant pivot/data context at implementation-defined resolution time/place and **not** derived from `data_files.filters` clause derivation/representatives; changing filter representatives/ranges (where configuration allows equivalent field maxima scenarios) must not be the source of placeholder values.
3. **D9 uniform rule:** (a) configured filter field missing from PivotTable available fields → fail fast with clear context; (b) field exists, no specified values matchable → fail fast; (c) ≥1 matchable → matched subset applied, run proceeds. Applies to all `data_files.filters` fields and `reports.articleLevel4Name_clean`, `reports.targetAudience`. **No warn-and-skip** of reports for no-match cases.
4. **D10:** `min`/`max` in `data_files.filters` resolve from the referenced field only (no cross-field/year-first dependency).
5. **D2 compliance (M2):** data-file resolution uses fixed precedence `.xlsx/.xlsm/.xlsb` in `input_path`, deterministic; no match → fail fast; >1 match → fail fast (recursive/ambiguous behavior not accepted).
6. **No-data / fail-fast outcomes:** filter yielding no data after applicable selections and other D12 conditions produce contextual fail-fast errors as specified.
7. **Naming/sanitization/collisions (D14/FR-NAME/FR-RPT):** `Trim` rules preserved; Excel-invalid chars sanitized; 31-char limit enforced; collisions fail fast after substitution/sanitization/trimming.
8. **`report_wclip` correctness:** intended week/period selection produces expected output columns/structure (no unintended 54–350 expansion) with evidence for axis behavior (U2).
9. **Regression across configs:** all `data_files` rows and all `reports` rows produce correct, deterministic outputs (names/sheets/values/formatting) per spec on actual source data/copies; existing rows (`report_m`, `report_w`) remain valid after template/migration changes.
10. **Run controls (FR-RUN-1):** `Refresh_data`/`Create_report` present and correctly bound; legacy `RunXlconcat` removed.

## 5. Out of scope

- Changing D2 meaning in spec (normative) or introducing non-spec requirements.
- Reintroducing warn-and-skip for filter no-match cases or cross-field/year-first `min`/`max` dependency.
- Redesign beyond rev.7 placeholder contract, D2/D9 alignment, and evidence-backed `report_wclip`/verification needs.
- Committing gitignored `data/theiss/*` artifacts or adding runtime dependencies outside VBA/Excel.

## 6. References

- `.ai/problems/001_treport_spec.md` (rev.7, normative)
- `.ai/problems/002_report_wclip_week_filter_spec.md` (rev.2)
- `mko_treport/src/vbs/export/*`
- `mko_treport/src/scripts/build_treport_vba.ps1`
- `docs/02-build-and-deployment/vba-build.md`

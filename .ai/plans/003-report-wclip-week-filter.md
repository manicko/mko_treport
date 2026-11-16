# Execution plan: 003 — report_wclip week filter / generalized `data_files.filters` model

- **Plan ID:** 003-report-wclip-week-filter  
- **Status:** Audit + implementation state reviewed (2026-11-16). Code in `mko_treport/src/vbs/export/` reflects a **provisional generic-filter refactor direction** (uniform rules, global numeric domain pre-resolution in current model). **Spec 001 rev.7 placeholder contract is not implemented in code** (still legacy `<max_year>/<max_monthnum>/<max_weeknum>` alias flow). Runtime verification (L1–L3, U1–U10) and several **decision conflicts** remain open.  
- **Date:** 2026-11-16 (this revision)  
- **Source specs/plans:** `.ai/problems/002_report_wclip_week_filter_spec.md` (rev.2), `.ai/problems/001_treport_spec.md` (**rev.7**, placeholders: `max_<PivotTable field name>` in `report_sheet`, max of that named field in relevant pivot/context for processed `data_files` row, **independent of `data_files.filters`**). Plan lineage references `.ai/plans/003-report-wclip-week-filter.md` prior sections (B1/B2/B3 drafts) where noted.  
- **Authoring note:** semantic/task intent only; do not add speculative implementation. Follow repo rules: VBA exports under `mko_treport/src/vbs/export/*`, runtime VBA-only in Excel, build via `mko_treport/src/scripts/build_treport_vba.ps1`, backups/working copies only, **L2 manual VBE `Debug → Compile VBAProject` required** to claim compile, L3 on copies for logic/run changes.

## Current state (from audit, code-truth summary)

**Realized in VBA exports (high level, architecture):**
- `modTreportFilters`: generic clause parsing/classification/validation (`|`, `field = expr`, list/range/scalar, numeric-convertible/bound rules, duplicate-field rejection), **global per-field numeric domains** pre-resolved before apply (`ResolveFilterContext`/`FieldStates`), clause application in declared order in current flow, **uniform selection** behavior via shared helpers (`ApplyDataFileFilters`, `ApplyReportFilters`). Legacy hardcoded `Year/Month Number/Week Number` + `ResolveMinMaxAndYear`/fixed `MaxYear/MaxMonth/MaxWeek` resolution path removed in favor of this model.
- `modTreportPivot`: field availability via `PivotFields` (`GetPivotFieldRequired`), placement (`EnsureFilterPlacement`: keep active in place; move hidden/unplaced to Filters/page with best-effort sort), field-scoped selection with **fail if none of specified values matchable** and **apply matched subset if any match**, no `ClearAllFilters`.
- `modTreportRun`: `Refresh_data` / `Create_report` entrypoints, state capture/restore, orchestration uses generic filters → context → apply once per `data_files` row then report filters; **no warn+skip report path for `articleLevel4Name_clean` none-match** (uniform fail). Report workbook/sheet assembly, sanitization/collision checks, static snapshot copy + formatting as before (group_rows no-op).
- `modTreportOutput`: legacy placeholders only **`<max_year>`, `<max_monthnum>`, `<max_weeknum>`** substituted from caller-supplied values.
- `modTreportFilters` + `modTreportRun`: **legacy placeholder compatibility** — representatives recorded during apply and aliases derived for `Year`/`Month Number`/`Week Number` (`DeriveLegacyPlaceholderAliases`), passed into `SubstitutePlaceholders`. This is **not** the rev.7 model (not tied to arbitrary `max_<field name>` resolved independently of `data_files.filters`).
- `modTreportConfig`: `reports`/`data_files`/`general`/`placeholders` reading unchanged in schema intent for this area except behavior changes above; `placeholders` table content not used by the legacy resolver beyond existing structure.

**Not evidenced / pending (from audit):** L1/L2/L3 results on working copies; workbook `run` sheet buttons/bindings vs FR-RUN-1; FR-DEP repeatability proof; live Excel confirmation of field identity/caption resolution, U2 axis causality for `report_wclip` extra columns, and regression across all configured rows.

## What remains to implement / verify (only what is needed)

### 1) Spec 001 rev.7 placeholder contract (critical) — implement
- **Requirement (rev.7, authoritative):** in `report_sheet`, placeholders are **`max_<PivotTable field name>`** as **`<max_<field name>>`**. Value = **maximum numeric value** of that named PivotTable field in the relevant pivot/context for the processed `data_files` row. **Not derived from `data_files.filters`**. Unresolvable → **fail fast** with clear context. `report_sheet` only; substitution/sanitization/naming/collision rules per D14/FR-NAME (spaces inside preserved by `Trim` rules as specified).
- **Approach (minimal, semantic):**
  - In `modTreportOutput`/call path (`modTreportRun`), parse `report_sheet` templates for all `max_<…>` tokens (or equivalent scan) and resolve each by that exact/normalized field name against the relevant PivotTable available fields (use same deterministic field-lookup rules agreed for filters; document resolution rule). Compute the **maximum** from available numeric values of that field in the pivot context used for the row (define whether this is “all available numeric items of that field in current pivot state after any required refresh/application context” consistently with project numeric rules: convertible-to-number, locale-safe per implementation constraints; fail if no numeric values / unresolvable).
  - Replace legacy fixed trio handling (`<max_year>/<max_monthnum>/<max_weeknum>`) and the alias-derivation chain in `modTreportFilters`/`ResolvedContext` that feeds placeholder substitution for this purpose (remove or refactor so placeholders are **not** sourced from filter representatives). Preserve existing `report_sheet` normalization/sanitization/collision/fail-fast behavior.
  - **Config/migration impact:** existing `report_m`/`report_w` templates using legacy tokens may require changing to `max_<…>` (e.g. to the actual field names present in those pivots). Assess and apply explicitly (config updates in `treport.xlsm`/workbooks per workflow, not invented here); verify sheet-name outcomes in L3.
  - **Do not** reintroduce hardcoded field-name semantics in the filter engine for this feature except the documented token form `max_<field name>`.

### 2) Decisions recorded (user, rev. 7) — align code/spec/plan
- **Filter matching (D9, user-decided):** **Uniform rule** for all filters (`data_files.filters`, `reports.articleLevel4Name_clean`, `reports.targetAudience`): (1) field name not in PivotTable **available fields** → **fail fast**. (2) field exists but **none** of specified values matchable → **fail fast** (no warn-and-skip of reports). (3) if **≥1** matchable → apply **matched subset**, proceed. This supersedes any earlier “warn and skip report” on no matches for `articleLevel4Name_clean`.
- **`min`/`max` in `data_files.filters` (D10, user-decided):** resolve against available numeric values of **the referenced field only** (named in that clause), in the relevant pivot context for the row. **No** cross-field dependency-aware year-first scoping in this rule set.
- **Placeholders (D11, rev. 7):** `max_<PivotTable field name>` in `report_sheet` only; value = maximum numeric value of that named field in the relevant pivot/context for the processed `data_files` row; **not derived from `data_files.filters`**. Placeholders are defined by pivot field names in use (legacy fixed placeholder descriptions in workbook/config commentary are superseded). If unresolvable → **fail fast**.
- **File resolution (D2):** current code behavior (recursive search, fail on >1 match, no extension precedence pick) is noted as a divergence from strict D2 text in some contexts; confirm desired strict behavior separately if needed.

### 3) Runtime / verification (T6 completion + T7)
- **L1:** run `mko_treport/src/scripts/build_treport_vba.ps1` on a working copy, review report (note VBOM/trust/caveats). No compile claim from L1 alone.
- **L2 (required):** manual **VBE `Debug → Compile VBAProject`** on working copy; record result.
- **L3 (on copies only):** `Create_report`/`Refresh_data` as relevant. Verify **`report_wclip`** (Q2-G8): apply intended `filters` context for `Week of Year = (202501–202553)` per decided model; acceptance = intended 2025 week period set in output (no unintended extra columns 54–350 as in reported symptom) — **U2**: confirm whether filtering `Week of Year` constrains the column axis driving expansion or whether axis/layout also requires adjustment; record evidence and make only evidence-backed, minimal scoped changes if needed (no speculative hardcoding).
- **Regression:** all `data_files`/`reports` rows (incl. `report_m`, `report_w`) with outputs compared appropriately to expectations after any config/token/migration changes.
- **U1/U7/U8/U10:** verify field name/caption resolution as VBA sees it, numeric/locale/item typing/matching (`Week of Year`/numeric forms), placement/multi-select feasibility for arbitrary fields used.
- **FR-RUN-1:** in the actual workbook confirm exactly two controls/entry points `Refresh_data` and `Create_report` bound to same-named VBA procedures and removal of legacy `RunXlconcat` (not provable from exports).
- **FR-DEP (if in scope):** verify single-step auxiliary build/import idempotence, module/document integrity, compile-clean result on working copies as per `docs/02-build-and-deployment/vba-build.md`.

## Out of scope (this plan)
- Redesign beyond the rev.7 placeholder contract + resolving the listed decision conflicts + required verification; inventing Excel/workbook facts; committing gitignored `data/theiss/*`; runtime dependencies outside VBA/Excel.

## References (current)
- `.ai/problems/001_treport_spec.md` (rev.7)
- `.ai/problems/002_report_wclip_week_filter_spec.md`
- `mko_treport/src/vbs/export/*` (authoritative sources)
- `docs/02-build-and-deployment/vba-build.md`
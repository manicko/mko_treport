# Analytical specification: treport (Excel macro reporting automation)

- **Spec ID:** 001_treport
- **Status:** Ready for implementation planning, pending validation baseline (OQ-B7: target Excel/OS, representative data, golden output) and technical validations (§13)
- **Date:** 2026-11-14 (rev. 4 — rev. 2 changes finalized; all clarification options resolved as recommended; added VBA-only implementation constraint)
- **Source request:** `docs/description.docx` (ТЗ “treport”, revised 2026-11-14)
- **Primary artifact in scope:** `data/theiss/treport.xlsm` (macro-enabled workbook, contains `xl/vbaProject.bin`; VBA source not present in repo)
- **Related observed data/artifacts (evidence):** data/theiss/base/ (eport_wclip.xlsx, CategoryData*.xlsx, clean_data/, aw_data/, dict/), data/theiss/monthly/ (eport_m.xlsx), data/theiss/weekly/ (eport_w.xlsx); mko_treport/src/{scripts,vbs,xls} exist but are empty.
- **Implementation constraint (PO-confirmed):** Project must use **VBA only** (Excel VBA macros/modules in `treport.xlsm`). **No Python or any other programming/scripting language** may be used for implementation or automation.

> This document specifies **what** must be built/achieved (business behavior and acceptance), not **how** to implement it in code/architecture beyond constraints implied by confirmed product decisions and the existing Excel/VBA artifact. **Implementation is constrained to VBA only; Python and any other programming languages are not permitted.** **Facts, confirmed decisions, assumptions, and open technical validations are separated explicitly.**

## 1. Problem statement

Reporting requires (1) refreshing source Excel data files that contain Power Query and/or pivot-table data and (2) producing periodical Excel report workbooks by extracting filtered pivot results with required formatting. This is performed by a macro in `treport.xlsm` driven by configuration tables on the `settings` sheet.

**Revision context (2026-11-14):** filter logic and configuration were revised and are now finalized by PO decision:
- `general`: `report_week`/`report_year` removed; `category_delimeter` renamed to `delimeter` (used in `reports` and `data_files`).
- `reports`: filter columns renamed to `articleLevel4Name_clean`, `targetAudience` (aligned to data-file field names).
- `data_files`: added `filters` column; period/field filtering is expressed per data-file row via a `filters` DSL with `min`/`max` tokens.
- `placeholders`: renamed to `max_weeknum`, `max_monthnum`, `max_year` (derived from `filters`/data), substituted in `report_sheet` (e.g. `<max_weeknum>`, `<max_year>`).
- PO confirmed all recommended clarifications for `filters` grammar/resolution, placeholder scope/format, filter placement/staging, and deterministic policies (OQ-1..OQ-4, OQ-B1, OQ-B3–OQ-B6). **OQ-B7** (target Excel/OS, representative data, golden output) remains required input for correctness sign-off; technical validations remain in §13.

**Noted ТЗ wording issues (non-blocking, resolved by confirmed decisions):** revised ТЗ constraint text still mentions old field names `report_week\report_year\ category \ta`; a `Week Number=(1-max)|Years=max` example line appears duplicated. These are treated as stale/typo per OQ-1d and do not override confirmed semantics.

## 2. Goals and non-goals

### 2.1 Business goals
- Enable a repeatable, configuration-driven run to **refresh data files** and **generate report workbooks** for the configured period/filters and audience/category selections.
- Reduce manual effort and copy/paste errors; ensure deterministic outputs for the same resolved inputs/configuration and data state.
- Provide clear, actionable failure outcomes when required inputs/config/data are missing, invalid, ambiguous, or yield no data.

### 2.2 Non-goals (out of scope)
- Designing or changing system/software architecture beyond the agreed Excel-macro solution in `treport.xlsm`.
- Implementing VBA/code, refactoring, or modifying repository files/data as part of this analysis.
- Implementing or automating the solution in Python or any other programming/scripting language (VBA only is permitted).
- Pixel-perfect parity with an unspecified historical output unless validated against an agreed golden output (OQ-B7).
- Defining BI/business meaning of metrics/categories/TA/field semantics beyond names and filter expressions present in configuration and data.

## 3. Confirmed facts (evidence-based)

| Area | Confirmed fact | Evidence |
|---|---|---|
| Revised ТЗ | Describes `settings`/`run`, tables `general,reports,data_files,placeholders`; `delimeter` in `general` used in `reports` and `data_files`; `reports` columns `articleLevel4Name_clean`, `targetAudience`, `name`; `data_files` includes `filters`; workflow applies data-file-level filters (`data_files.filters`) then report-level filters (`reports`); placeholders usable in `data_files` for substitution. States `data_files` read-only mode (see Q1 resolution) and `Refresh_data` refreshes PQ/RefreshAll and saves. | `docs/description.docx` (extracted, rev. 2) |
| `filters` principle (ТЗ) | `<имя поля> = значение или (диапазон значений) или (значение 1 <разделитель> значение 2) | <имя поля 2> =....` Examples: `Month Number=(1-max)|Years=max`, `Week Number=(1-max)|Years=max`. `max` — maximum, `min` — minimum. | `docs/description.docx` (rev. 2) |
| Workbook container | `treport.xlsm` exists; sheets `run` (`Sheet2`), `settings` (`Sheet1`); 4 table parts; drawings/vml/control props; `xl/vbaProject.bin` present. | `data/theiss/treport.xlsm`; `xl/workbook.xml`, `xl/worksheets/*`, package parts |
| `general` table | ListObject `general` (`A1:B4`, `parameter,value`): `input_path` = workbook folder formula, `output_path` = `B2&"output"`, `delimeter` = `&`. (`report_week`, `report_year`, `category_delimeter`, `new_file_delimeter`, `keep_col_hidden` not present.) | `xl/tables/table1.xml`, `xl/worksheets/sheet2.xml`, `xl/sharedStrings.xml` |
| `reports` table | ListObject `reports` (`D1:F7`, `articleLevel4Name_clean,targetAudience,name`), 6 rows: (`ЗУБНАЯ ПАСТА&ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`,`W 50+`,`LACALUT ACTIV`), (`ЗУБНАЯ ПАСТА&ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`,`W 30-50`,`LACALUT ACTIV MEDICAL WHITE`), (`СРЕДСТВА ОТ БОЛИ В МЫШЦАХ И СУСТАВАХ`,`W 50+`,`DOLGIT`), (`СРЕДСТВА ОТ БОЛИ В МЫШЦАХ И СУСТАВАХ`,`W 30-50`,`DOLGIT `), (`СРЕДСТВА ОТ КАШЛЯ`,`W 30-50`,`HEDELIX`), (`СЛАБИТЕЛЬНЫЕ ПРЕПАРАТЫ`,`W 50+`,`REGULAX`). | `xl/tables/table2.xml`, sheet2/sharedStrings |
| `data_files` table | ListObject `data_files` (`D10:M13`, `data_file,data_sheet,report_sheet,filters,paste_row,paste_col,freeze_row,freeze_col,zoom,group_rows`), 3 rows: `report_m`/`clean_data_full`/`помесячное сравнение`/`Month Number=(1-max)|Years=max`/1,1,blank,blank,0.7,TRUE; `report_w`/`clean_data_full`/`сравн конкур_<max_weeknum>w_<max_year>_прямая`/`Week Number=(1-max)|Years=max`/1,1,2,2,blank,blank; `report_wclip`/`clean_data_full`/`хроно по неделям`/`Week Number=(1-max)|Years=max`/1,1,blank,blank,blank,blank. | `xl/tables/table4.xml`, sheet2/sharedStrings |
| `placeholders` table | ListObject `placeholders` (`A10:B13`, `placeholder,description`): `max_weeknum` — “неделя отчета рассчитывается на основе filters в data_file”; `max_monthnum` — “месяц  отчета рассчитывается на основе filters и данных в data_file”; `max_year` — “год отчета рассчитывается на основе filters и данных в data_file”. | `xl/tables/table3.xml`, sheet2/sharedStrings |
| Data artifacts | Candidate files exist: `data/theiss/monthly/report_m.xlsx`, `data/theiss/weekly/report_w.xlsx`, `data/theiss/base/report_wclip.xlsx` (unique by observed exact name). No `data/theiss/output` folder observed. Candidate xlsx packages show pivot/connection/query indicators; field-level schemas/captions not fully extracted. | Directory listings / package inspection |
| Repo state | No exported VBA modules; `mko_treport/src/{scripts,vbs,xls}` empty; `data/*` git-ignored. | Repo scan |

## 4. Scope

### 4.1 In scope
1. **Configuration reading & validation** from `treport.xlsm` `settings` ListObjects `general`, `reports`, `data_files`, `placeholders` per revised schema and confirmed rules.
2. **`Refresh_data`**: locate configured data files in `input_path`, refresh Power Query/pivot data as required, persist refreshed data, and report outcome.
3. **`Create_report`**: per `data_files` row (in order), open source, resolve `filters`/`min`/`max` and placeholders deterministically, apply `data_files.filters` (data-file level) then per-`reports` filters (`articleLevel4Name_clean`, `targetAudience`) with confirmed placement/staging; generate output workbooks/sheets per `reports` and `data_files` order with required copy/format/naming; save outputs and report outcome.
4. **Deterministic rules** for file resolution, `filters` grammar/interpretation, `min`/`max` resolution, multi-value matching via `general.delimeter`, copy range, placeholder substitution/naming, idempotent overwrite, and fail-fast error behavior.
5. **Operational/user-visible outcomes**: clear success/failure messaging sufficient to identify failing operation, config row/name, file path, sheet, field, filter/expression, or value.

### 4.2 Out of scope
- Any non-Excel runtime, scheduler/orchestrator, UI redesign, or changes to source data content/pipelines.
- Any implementation or automation using Python or any programming/scripting language other than VBA.
- Recovery/merge of concurrent edits or multi-user locking workflows beyond deterministic failure on locked/unavailable resources.
- Historical-data correction, metric definition changes, or new report types not representable by existing config model.
- Defining BI/business meaning of metrics/categories/TA/field semantics beyond configured names and filter expressions.
- `group_rows` color-band grouping behavior until an explicit deterministic rule is approved (see D11).

## 5. Constraints

- **Artifact/architecture:** Solution operates within the existing 	report.xlsm Excel-macro model and configured settings/un/ListObjects.
- **Implementation language (project constraint):** The project must use **VBA only** (Excel VBA macros/modules within `treport.xlsm`). **No Python or any other programming or scripting language** may be used to implement or automate Refresh_data, Create_report, or any related logic.
- **Access modes (D1):** `Create_report` — data files opened **read-only, never saved**. `Refresh_data` — data files opened **writable (not read-only)** to refresh PQ/RefreshAll and **save**, then closed.
- **Processing order:** `data_files` rows processed in configured order; output sheets follow `data_files` order.
- **Filter placement (D7):** for every filter field applied (all fields in `data_files.filters` and `reports.articleLevel4Name_clean`, `reports.targetAudience`), if the field already exists in Filters/Rows/Columns filter it in place without moving it; if absent, add to Filters and filter there.
- **Filter staging (D8):** for each `data_file`, apply `data_files.filters` once and keep them applied; for each `reports` row apply/reapply that row’s `articleLevel4Name_clean` and `targetAudience` (overriding prior report-row selection for those fields) without moving placed fields; one report row’s report-level filters must not leak into another.
- **No invention:** unspecified semantics must not be fabricated; unresolved technical validations are recorded in §13.
- **Determinism:** identical resolved inputs/config/data state must yield deterministic resolution, naming, and structure.

## 6. Confirmed product decisions (PO: all recommended, 2026-11-14)

| ID | Decision | Confirmed choice |
|---|---|---|
| D1 (Q1) | Access modes | `Create_report`: data files read-only, never saved. `Refresh_data`: data files writable (not read-only) to refresh PQ/RefreshAll and save, then close. |
| D2 (Q4, OQ-B3) | Data-file resolution | `data_file` is base name resolved in `input_path` to `<data_file>.<ext>` with fixed precedence **`.xlsx`, `.xlsm`, `.xlsb`**; **fail fast** if no match or >1 match. |
| D3 (Q5) | Outputs & idempotency | One output workbook per `reports` row named `reports.name` + `.xlsx`; overwrite if exists on successful run; sheets created in `data_files` order and recreated/replaced deterministically for idempotent reruns. |
| D4 (rev. 2, OQ-4) | Filter responsibilities | `data_files.filters` contains period/data fields (e.g. `Years`, `Week Number`, `Month Number`) as configured; `articleLevel4Name_clean` and `targetAudience` are controlled via `reports` table. |
| D5 (OQ-1a) | `filters` grammar | Clauses separated by `|`; each `<field> = <expr>` (spaces around `=`, `|` trimmed; field names may contain spaces). `<expr>`: (i) **scalar** — literal or `min`/`max` (e.g. `Years=max`); (ii) **range** — `(start-end)` **inclusive**, operands are numeric literals or `min`/`max`, `-` is range operator (e.g. `(1-max)`, `(1-12)`); (iii) **list** — `(v1<delimeter>v2…)` using `general.delimeter` (`&`). Field and literal values match pivot field/item captions **exactly after trim**. Disambiguation: expression containing `general.delimeter` → list; else containing `-` between two bound tokens → range; else scalar. Multi-value lists in `filters` use `general.delimeter`. |
| D6 (OQ-1d) | ТЗ duplicate example | Duplicated `Week Number=(1-max)|Years=max` example line treated as typo/ignored. |
| D7 (OQ-1b, Q6 placement) | Placement rule scope | Applies to **all** filter fields: every field in `data_files.filters` and `reports.articleLevel4Name_clean`, `reports.targetAudience`. Filter in place if in Filters/Rows/Columns, else add to Filters (no move). |
| D8 (OQ-1c) | Filter staging/reset | Per `data_file`: apply `data_files.filters` once and keep applied; per `reports` row apply/reapply that row’s `articleLevel4Name_clean` and `targetAudience` (override prior report-row selection for those fields) without moving placed fields; no leak between report rows. |
| D9 (Q6, OQ-1a) | Value matching | Multi-value selection split by `general.delimeter`, tokens trimmed, all exactly matching pivot items selected (exact to captions after trim). Applies to `articleLevel4Name_clean` and any list expressions in `filters`; `targetAudience` applied per configured value with same exact/trim matching principle. |
| D10 (OQ-2a, OQ-2c; supersedes Q2/Q3) | `min`/`max` resolution & bounds | Per `data_files` row, resolve against available values of referenced fields in the pivot on `data_sheet` of that `data_file` (actual data). Resolution is deterministic and dependency-aware: when `Years` is present, resolve `Years` bound first (e.g. `Years=max` → max year in data), then resolve dependent period fields within that resolved `Years` context (e.g. `Month Number` `min/max` = min/max month available in resolved year; `Week Number` `min/max` = min/max week available in resolved year). Ranges are **inclusive** on both ends; comparisons are **numeric**. |
| D11 (OQ-2b, OQ-3a, OQ-3b) | `max_*` placeholders | `max_year` = resolved max of `Years`; `max_weeknum` = resolved max of `Week Number`; `max_monthnum` = resolved max of `Month Number` for that `data_files` row in the same resolved context. Substitution target: **`report_sheet` only**. Token syntax `<name>` as in workbook (`<max_weeknum>`, `<max_monthnum>`, `<max_year>`), substituted as resolved values **without zero-padding**. If a placeholder is **referenced** but its corresponding value cannot be deterministically derived for that row, **fail fast** with a clear config/data error. |
| D12 (Q8, OQ-B4) | Errors & no-data | **Fail fast** with clear contextual error for missing/invalid config or paths, missing/unresolvable data file, missing `data_sheet`, missing pivot, missing required filter field, invalid/unparseable/unresolvable `filters`, or referenced-but-underivable placeholder. Filter yielding **no data** → **fail fast** with contextual error (file/sheet/filters/report). |
| D13 (Q9, OQ-B5) | Formatting defaults & `group_rows` | Blanks = default/no-op for `freeze_row/freeze_col/zoom/paste_row/paste_col/group_rows`. `group_rows` has no approved deterministic color-band rule; therefore **`group_rows=TRUE` must fail fast** with a clear error until an explicit deterministic rule is approved. Blank/false = no-op. (Current workbook config must be brought into compliance before acceptance — noted as config/validation issue.) |
| D14 (Q10, OQ-B6) | Naming & collisions | Sanitize Excel-invalid sheet-name characters, enforce **31-character** limit, resolve collisions deterministically: **fail fast** on sheet-name or report-file-name collision after substitution/sanitization. |
| D15 (Q7) | Copy range | Copy pivot presented data range incl. row/column labels and grand totals as shown after filtering; exclude pivot object/page-filter area and PQ connections; paste static **values + formats** (+ column widths if needed). |
| D16 (Q11) | Scope/source of truth & language | **VBA-only** solution in `treport.xlsm` (Excel VBA macros/modules); **no Python or any other programming/scripting language**. Behavior aligned to revised ТЗ intent reconciled with current workbook structure; `Refresh_data` and `Create_report` in scope. |
| D17 (OQ-B1) | Removed parameters | `new_file_delimeter` and `keep_col_hidden` are dropped (absent in rev. 2) and must not be required/processed. |
| D18 (Q12, OQ-B7) | Acceptance baseline | Correctness sign-off requires target Excel/OS context, representative `input_path` dataset, and ≥1 golden expected output for a fixed configuration/period (baseline details to be provided). |

## 7. Confirmed functional requirements

### 7.1 Configuration model (FR-CFG)
- **FR-CFG-1:** Read settings from ListObjects `general`, `reports`, `data_files`, `placeholders` on the `settings` sheet of `treport.xlsm`.
- **FR-CFG-2:** `general` provides `input_path`, `output_path`, `delimeter`. `input_path` and `output_path` are mandatory and non-empty; `delimeter` is mandatory when multi-value splitting is required by config. Missing/empty/invalid required values → fail-fast validation error (D12).
- **FR-CFG-3:** `reports` defines report rows with columns `articleLevel4Name_clean`, `targetAudience`, `name` (one output workbook per row) (D3, D16).
- **FR-CFG-4:** `data_files` defines ordered source-processing rows with columns `data_file`, `data_sheet`, `report_sheet`, `filters`, `paste_row`, `paste_col`, `freeze_row`, `freeze_col`, `zoom`, `group_rows`; row order is authoritative for processing and output sheet order (D3, D8).
- **FR-CFG-5:** `placeholders` defines substitution parameters `max_weeknum`, `max_monthnum`, `max_year` with derivation basis per D11.
- **FR-CFG-6:** Blank optional fields follow default/no-op where defined (D13). Dropped parameters (`new_file_delimeter`, `keep_col_hidden`) must not be required/processed (D17). Invalid/unsupported/unresolvable expressions or values must not fail silently → fail fast per D12.

### 7.2 File resolution and access (FR-FILE)
- **FR-FILE-1:** Resolve each `data_files.data_file` (base name) within `general.input_path` to `<data_file>.<ext>` checking extensions in fixed precedence **`.xlsx`, `.xlsm`, `.xlsb`** (D2). Outcome must be deterministic.
- **FR-FILE-2:** If no match or more than one match exists, fail fast with an error listing configured name, searched path, and cause (D2, D12).
- **FR-FILE-3:** `Create_report` opens data files read-only and never saves them (D1).
- **FR-FILE-4:** `Refresh_data` opens data files writable (not read-only), performs refresh, saves, and closes (D1). Unavailability/lock is a clear fail-fast error (D12).
- **FR-FILE-5:** Ensure `output_path` exists (create if missing) before writing outputs; inability to create/write fails fast (D12).

### 7.3 Refresh operation (FR-REF)
- **FR-REF-1:** For each resolved data file (deterministic order), open writable, refresh Power Query and dependent pivot data so results are persisted, save, and close (D1).
- **FR-REF-2:** Refresh must complete before save/close; asynchronous/background refresh must not leave persisted data stale (technical validation §13, OQ-T2).
- **FR-REF-3:** On refresh failure, fail fast with contextual error (file, stage); do not leave application/workbook state corrupted (D12).

### 7.4 Filter parsing, resolution, and application (FR-FLT)
- **FR-FLT-1 (parse `filters`):** For each `data_files` row, parse `filters` per D5: clauses separated by `|`; each `<field> = <expr>` (trim spaces around `=`, `|`; field names may contain spaces). `<expr>` is scalar (literal or `min`/`max`), inclusive range `(start-end)` with numeric/`min`/`max` bounds and `-` operator, or list `(v1<delimeter>v2…)` using `general.delimeter`. Disambiguation: contains `general.delimeter` → list; else contains `-` between two bound tokens → range; else scalar. Empty/invalid/unparseable expression or duplicate/ambiguous clause → fail fast (D12).
- **FR-FLT-2 (resolve `min`/`max`):** Per D10, for each `data_files` row resolve `min`/`max` against available values of referenced fields in the pivot on `data_sheet` of that `data_file` (actual data). Deterministic, dependency-aware: resolve `Years` bound first when present, then resolve dependent period fields within that resolved `Years` context (e.g. `Month Number` and/or `Week Number` `min/max` within resolved year). Ranges inclusive on both ends; comparisons numeric. Unresolvable bound/value → fail fast (D12).
- **FR-FLT-3 (resolve placeholders):** Per D11, derive `max_year`/`max_weeknum`/`max_monthnum` as resolved maxima of `Years`/`Week Number`/`Month Number` in the same resolved context for that `data_files` row. If a placeholder is referenced in `report_sheet` but cannot be deterministically derived → fail fast (D12).
- **FR-FLT-4 (apply data-file filters):** For each `data_file`, apply resolved `data_files.filters` once to the pivot on `data_sheet` and keep applied for all `reports` rows processed against that `data_file` (D8). `data_files.filters` contains period/data fields as configured (D4).
- **FR-FLT-5 (apply report filters):** For each `reports` row, apply `articleLevel4Name_clean` and `targetAudience`. Multi-value `articleLevel4Name_clean` split by `general.delimeter`, tokens trimmed, all exactly matching pivot items selected (D9). `targetAudience` applied per configured value with exact-after-trim matching. Report-level filters are reapplied per report row (override prior report-row selection for those fields) without leaking between report rows (D8).
- **FR-FLT-6 (placement):** For every filter field applied (all `data_files.filters` fields and `articleLevel4Name_clean`, `targetAudience`), if field exists in Filters/Rows/Columns filter in place without moving it; if absent, add to Filters and filter there (D7).
- **FR-FLT-7 (field identity & outcomes):** Filter field names and literal values must correspond/match pivot field and item captions exactly after trim (D5, D9). Missing `data_sheet`, missing pivot, missing required filter field, or unresolvable expression/value → fail fast with context (D12). Filter yielding **no data** → fail fast with contextual error (D12).

### 7.5 Report generation and output (FR-RPT)
- **FR-RPT-1:** Create one output workbook per `reports` row named `reports.name` + `.xlsx`. If exists, overwrite on successful run. Lock/overwrite failures → fail fast (D3, D12).
- **FR-RPT-2:** For each `data_files` row in order, create a target sheet in each relevant report workbook named from `report_sheet` after placeholder substitution, in `data_files` order. For idempotency, target sheets are recreated/replaced deterministically (D3, D11).
- **FR-RPT-3:** Copy range = pivot’s presented data after all applicable filters: include row/column labels and grand totals as shown; exclude pivot object/page-filter area and PQ connections (D15).
- **FR-RPT-4:** Paste at `paste_row/paste_col` (blank = default/no-op per D13) as static **values + formats** (+ column widths if needed). Output must not contain pivot objects or PQ connections (D15).
- **FR-RPT-5:** Apply formatting per `data_files`: freeze panes (`freeze_row/freeze_col`), zoom, hide gridlines, and row/column sizing/state as required; blanks are no-op/default (D13).
- **FR-RPT-6:** `group_rows=TRUE` → **fail fast** with clear error until an explicit deterministic rule is approved; blank/false = no-op (D13).
- **FR-RPT-7:** Save each report workbook to `output_path` and release resources; close source files per access-mode rules (D1).

### 7.6 Naming and placeholders (FR-NAME)
- **FR-NAME-1:** Substitute placeholders in `report_sheet` only (D11). Token syntax `<max_weeknum>`, `<max_monthnum>`, `<max_year>`; values are resolved per D10/D11 and inserted without zero-padding.
- **FR-NAME-2:** Sanitize final sheet/report names for Excel-invalid characters, enforce **31-character** limit, and **fail fast** on collision after substitution/sanitization (D14). Empty/invalid resulting names → fail fast (D12).

### 7.7 Run behavior and diagnostics (FR-RUN)
- **FR-RUN-1:** `Refresh_data` and `Create_report` invokable from `run` sheet controls as provided.
- **FR-RUN-2:** Execute without unnecessary UI disruption and restore application state on completion/failure.
- **FR-RUN-3:** Success and failure outcomes explicit. All errors fail fast and identify context (operation, config row/name, file path, sheet, field, filter/expression, value as applicable) (D12).

## 8. Conceptual development tasks (for planning decomposition)

| ID | Task | Purpose | Expected outcome | Dependencies |
|---|---|---|---|---|
| T1 | Config contract & validation | Implement revised schema and validation (incl. D2, D13 `group_rows`=TRUE reject, D17 dropped params, required fields). | Config schema, required/optional/defaults, validation/error matrix. | D1–D17 |
| T2 | `filters` parse & deterministic resolution | Parse DSL (D5) and resolve `min`/`max` + `max_*` per dependency-aware rule (D10,D11). | Deterministic parse/eval/resolve spec and behavior per `data_files` row. | T1, D5,D10,D11 |
| T3 | Data-file resolution & access policy | Deterministic extension resolution and mode-specific open/close/save. | Resolution + access rules per operation. | T1, D1,D2 |
| T4 | Refresh workflow | Reliable refresh-and-persist behavior for PQ/pivot sources. | Refresh sequence, completion criteria, failure handling. | T3, D1; OQ-T2 |
| T5 | Filter application & staging | Apply data-file filters once, then per-report filters with placement/reset rules; exact matching. | Stage/order/placement/reset behavior and no-data handling. | T1,T2, D4,D7,D8,D9,D12; OQ-T1 |
| T6 | Pivot snapshot extraction & paste | Static correct data+format capture/placement without pivot/PQ artifacts. | Copy-range and paste/width/height behavior. | T5, D15; OQ-T3 |
| T7 | Formatting & layout | Freeze/zoom/gridlines/row-col sizing/state; enforce `group_rows` policy. | Formatting acceptance criteria. | T1,T6, D13 |
| T8 | Report assembly, naming & idempotency | Workbook/sheet creation, ordering, overwrite/replace, placeholder substitution, collision handling. | End-to-end output lifecycle and naming rules. | T1,T2,T6, D3,D11,D14 |
| T9 | Diagnostics & run control | Fail-fast contextual errors and state-safe execution. | Error taxonomy/messages and run outcomes. | T2–T8, D12 |
| T10 | Verification & acceptance | Golden-data scenarios and sign-off criteria. | Test scenarios and Definition of Done inputs. | T1–T9, D18, OQ-T1–T4 |

## 9. Research summary (evidence-backed synthesis)

> Prior in-session research (artifact/state audit and VBA/PQ/pivot approach research) remains applicable; rev. 3 fixes semantics (filters DSL, dependency-aware `min/max`, staging, deterministic policies).

### 9.1 Current-architecture findings
- `treport.xlsm` contains `run`/`settings`, 4 ListObjects and `xl/vbaProject.bin` (VBA not source-reviewed). Rev. 2 schema/values confirmed: `general`(`input_path,output_path,delimeter`), `reports`(`articleLevel4Name_clean,targetAudience,name`), `data_files`(+`filters`, tokens `<max_weeknum>/<max_year>`), `placeholders`(`max_weeknum,max_monthnum,max_year`).
- Period selection is per `data_files` row via `filters` (e.g. `Month Number=(1-max)|Years=max`, `Week Number=(1-max)|Years=max`).
- `mko_treport/src/{scripts,vbs,xls}` empty; no exported VBA; `report_m/report_w/report_wclip` present with pivot/query indicators, field-level schemas unverified.
- **Config note:** current `data_files` row `report_m` has `group_rows=TRUE`; under D13 this is a fail-fast condition until a rule is approved (must be resolved in config or policy before acceptance).

### 9.2 Feasible implementation approaches
| Option | Description | Trade-offs |
|---|---|---|
| **A. Straight-through procedural VBA** | Procedural flows mirroring spec. | Low initial effort; higher regression risk for DSL parse/resolve, staging/reset, determinism, and state/error handling. |
| **B. Modular component-based VBA** | Separated config/validation, file resolution, `filters` parser+resolver, pivot/filter service, range export, formatting, report builder, centralized cleanup. | Better maintainability/correctness for DSL + placement/staging rules; moderate effort; Excel OM limits remain. |
| **C. Plan-then-execute with per-file isolation & auditable logging (on B)** *(recommended)* | Preflight + resolve (paths, parsed `filters`, dependency-aware `min/max`/`max_*`, names/collisions, filter plan) then isolated execution with structured errors; optional validate/dry-run. | Strongest determinism/auditability for D5–D14; predictable reruns and clear fail-fast diagnostics. Moderate effort. |

### 9.3 Recommended approach
- **Recommendation:** **Option C on modular foundation B**, aligned to D1–D17.
- **Rationale:** confirmed `filters` DSL and dependency-aware `min/max`/placeholder resolution (D5,D10,D11), strict staging/placement (D7,D8), and fail-fast policies (D12–D14) require deterministic pre-resolution, explicit validation, and per-file isolation to avoid silent mis-filtering, stale state across report rows, and ambiguous naming/output.

### 9.4 Key technical practices
- **`filters` handling:** strict validated parser per D5; resolve `min/max` dependency-aware per D10 against deterministic data sources; derive/log `max_*` per D11.
- **State & access:** enforce D1 modes; snapshot/restore app state; no unnecessary UI selection.
- **PQ/pivot refresh:** synchronous completion before pivot cache refresh/save (validate OQ-T2).
- **Filter staging/placement:** per D8 apply data-file filters once, reapply report filters per report row; per D7 filter in place or add to Filters without moving.
- **Static copy & formatting:** presented range values+formats (+widths) per D15; row heights/outline handling; enforce D13.
- **Naming:** substitute in `report_sheet` per D11; sanitize/limit/collision fail-fast per D14.

## 10. Risks and known constraints

| ID | Risk / unknown | Impact | Mitigation / validation |
|---|---|---|---|
| R1 | **Pivot topology & field identity** (pivot selection, exact captions/types/orientations for `articleLevel4Name_clean`, `targetAudience`, `Years`, `Week Number`, `Month Number`; regular vs OLAP/Data Model) | Filters/copy may target wrong object or fail | Validate on `report_m/report_w/report_wclip` and real inputs (OQ-T1); fail fast on missing/ambiguous fields (D12). |
| R2 | **PQ refresh completeness/timing** across Excel builds/query types | Stale persisted data after `Refresh_data` | Validate synchronous completion + pivot cache refresh on target Excel (OQ-T2). |
| R3 | **`min`/`max` data source details** (pivot cache items vs visible/other values) consistent with D10 | Incorrect period/filters/`max_*` | Confirm against actual data behavior during OQ-T1/T3; log resolved values. |
| R4 | **`group_rows` config conflict**: `report_m` currently `group_rows=TRUE` but D13 requires fail-fast until rule approved | `Create_report` will fail on current config unless changed | Resolve by setting `group_rows` blank/false for affected rows or by approving a deterministic rule and updating D13/spec. |
| R5 | **Copy fidelity** (range bounds incl. totals/page area, widths/heights/hidden/outline, value/date fidelity) | Output differs from intended | Golden-output validation (OQ-T3). |
| R6 | **Locked/open file handling** (sources/outputs in use) and filesystem permissions | Run failures or blocked overwrite | Fail fast with contextual errors (D12); validate in target environment. |
| R7 | **VBA logic opaque** (`xl/vbaProject.bin` not source-reviewed) | Divergence from spec/current behavior | Inspect/reconcile in Excel/VBA IDE or approved extraction (OQ-T4). |
| R8 | **Performance** on large pivots/many items | Long/unresponsive runs | Benchmark with representative data; batch/manual-update/progress (technical). |
| R9 | **Environment prerequisites** (Excel version, Windows, macro trust, PQ enabled, paths) | Inconsistent or failing runs | Provide/confirm baseline (D18, OQ-B7). |

## 11. Assumptions

- **A1:** `treport` remains and must be implemented as a **VBA-only** Excel solution in `treport.xlsm`; **no Python or any other programming/scripting language** is permitted (D16).
- **A2:** `articleLevel4Name_clean` and `targetAudience` are pivot filter fields matching those captions/names in source data; `Years`, `Week Number`, `Month Number` are period fields used in `filters` where configured (to be validated OQ-T1).
- **A3:** `min`/`max` and comparable values in period fields are numeric and compared numerically; ranges are inclusive (D10).
- **A4:** Each `data_files` row that is executed to produce named sheets with placeholders provides resolvable `max_*` values as required by its `filters`/referenced tokens; otherwise execution fails fast by design (D11,D12).
- **A5:** Correctness of values/formatting/layout is judged against an agreed golden output for a fixed configuration/period (D18).

## 12. Out of scope (explicit)

- Implementation, code changes, VBA authoring/refactoring, or repository/data modifications.
- New architecture, non-Excel tooling/pipelines, scheduling/orchestration services.
- Use of Python or any programming/scripting language other than VBA for any part of the solution.
- Changes to source data preparation, metric definitions, or upstream PQ/data models.
- `group_rows` color-band grouping until an explicit deterministic rule is approved (D13).
- Performance SLAs beyond correctness on representative agreed data unless separately defined.

## 13. Open items (non-blocking for planning, required for correctness sign-off)

### 13.1 Product/acceptance input
- **OQ-B7 (D18):** Provide/confirm target Excel version & OS, representative `input_path` dataset, and ≥1 **golden** expected output for a fixed configuration/period.

### 13.2 Technical validations
- **OQ-T1:** Identify actual pivot(s), field captions/types/orientations, and PQ/connection topology in `report_m/report_w/report_wclip` and real inputs (incl. `Years`, `Week Number`, `Month Number`, `articleLevel4Name_clean`, `targetAudience`).
- **OQ-T2:** Verify synchronous PQ refresh + pivot cache refresh completeness on target Excel.
- **OQ-T3:** Confirm copy-range boundaries and formatting fidelity against golden output.
- **OQ-T4:** Inspect/reconcile `xl/vbaProject.bin` behavior with this spec where relevant.

## 14. Definition of Ready (for implementation planning)

Ready for implementation planning when:
1. **Business decisions/clarifications closed:** D1–D17 confirmed (satisfied — PO selected all recommended). Resolved: OQ-1a–d, OQ-2a–c, OQ-3a–b, OQ-4, OQ-B1, OQ-B3–B6.
2. **Scope & requirements stable:** §§4–7 agreed, including `filters` grammar (D5), dependency-aware `min`/`max` and `max_*` derivation (D10,D11), placement/staging (D7,D8), and deterministic policies (D2,D12–D14).
3. **Deterministic policies fixed:** extension precedence, `min`/`max` resolution algorithm, placeholder mapping/scope/syntax/no-padding, no-data outcome (fail-fast), collision handling (fail-fast), `group_rows` handling (fail-fast while undefined).
4. **Validation baseline available:** target Excel/OS and representative dataset with golden output (OQ-B7) — *pending input*.
5. **Technical validations tracked:** §13.2/§10 treated as implementation/verification tasks.

**Planning gate:** items 1–3 completed (satisfied). Item 4 pending but does not block planning; items 4–5 must close before correctness sign-off. **Note:** current config contains `group_rows=TRUE` (R4) which conflicts with D13 and must be addressed (config change or approved rule) as part of implementation/acceptance.

## 15. Traceability summary (rev. 3)

| ТЗ / change / decision | Spec coverage |
|---|---|
| `general`: `input_path,output_path,delimeter`; removed `report_week/year` and dropped extras | FR-CFG-2, §3, D2,D17 |
| `reports`: `articleLevel4Name_clean,targetAudience,name` | FR-CFG-3, FR-FLT-5/6/7, D4,D8,D9 |
| `data_files`: +`filters`, order, format columns | FR-CFG-4, FR-FLT-1/2/4/6/7, FR-RPT, D4–D8,D10 |
| `filters` DSL + `min/max` + dependency-aware resolution | FR-FLT-1/2/3, D5,D6,D10,D11, §9.4 |
| Placeholders `max_weeknum/max_monthnum/max_year`, `<name>` in `report_sheet`, no padding | FR-CFG-5, FR-NAME-1, D11 |
| Staging: data-file filters once, report filters per report row | FR-FLT-4/5, D8, §5 |
| Placement: in-place or add to Filters, all filter fields | FR-FLT-6, D7, §5 |
| Access modes (Create read-only; Refresh writable+save) | FR-FILE-3/4, FR-REF, D1 |
| File resolution `.xlsx/.xlsm/.xlsb`, fail if none/>1 | FR-FILE-1/2, D2 |
| Output per report, overwrite, idempotent sheets in order | FR-RPT-1/2/7, D3 |
| Copy presented range, exclude pivot/page/PQ, static values+formats | FR-RPT-3/4, D15 |
| Formatting defaults; `group_rows=TRUE` fail-fast until rule approved | FR-RPT-5/6, D13, R4 |
| Fail-fast errors incl. no-data and collisions | FR-*, D12,D14 |
| Acceptance requires target Excel/OS + representative data + golden output | D18, §13.1, §14 |

---
*End of specification (rev. 4).*

# Analytical specification: treport (Excel macro reporting automation)

- **Spec ID:** 001_treport
- **Status:** Ready for implementation planning (business decisions confirmed; technical validations required — see Risks / Open technical validations)
- **Date:** 2026-11-14
- **Source request:** `docs/description.docx` (ТЗ “treport”)
- **Primary artifact in scope:** `data/theiss/treport.xlsm` (macro-enabled workbook, contains `xl/vbaProject.bin`; VBA source not present in repo)
- **Related observed data/artifacts (evidence):** `data/theiss/base/` (`report_wclip.xlsx`, `CategoryData*.xlsx`, `clean_data/`, `raw_data/`, `dict/`), `data/theiss/monthly/` (`report_m.xlsx`), `data/theiss/weekly/` (`report_w.xlsx`); `mko_treport/src/{scripts,vbs,xls}` exist but are empty.

> This document specifies **what** must be built/achieved (business behavior and acceptance), not **how** to implement it in code/architecture beyond the constraints implied by the chosen product decisions and existing Excel/VBA artifact. Facts, assumptions, decisions, and open items are separated explicitly.

## 1. Problem statement

Reporting requires (1) refreshing source Excel data files that contain Power Query and/or pivot-table data and (2) producing periodical Excel report workbooks by extracting filtered pivot results with required formatting. Today this is intended to be performed by a macro in `treport.xlsm` driven by configuration tables, but the behavior must be unambiguously specified so implementation/refinement can proceed without further business analysis.

**Observed gap:** the ТЗ is partially ambiguous/truncated in places (e.g. copy-range research note, grouping-by-color detail, some formatting defaults and parameter descriptions), and the workbook contains parameters/values not fully described in the ТЗ (`new_file_delimeter`, `keep_col_hidden`, `paste_row/paste_col`, `group_rows`, placeholder descriptions, `report_week/year = last`). Repository evidence of internal VBA logic and full runtime behavior is limited (no exported VBA; binary inspection constraints). Business decisions below resolve the product ambiguities for specification purposes.

## 2. Goals and non-goals

### 2.1 Business goals
- Enable a repeatable, configuration-driven run to **refresh data files** and **generate report workbooks** for a selected reporting period and audience/category filters.
- Reduce manual effort and copy/paste errors; ensure deterministic outputs for the same resolved inputs/configuration.
- Provide clear, actionable failure outcomes when required inputs/config/data are missing or invalid.

### 2.2 Non-goals (out of scope for this specification)
- Designing or changing system/software architecture beyond the agreed Excel-macro solution in `treport.xlsm`.
- Implementing VBA/code, refactoring, or modifying repository files/data as part of this analysis.
- Migrating to Python/CLI/services, changing data pipelines, or altering source data schemas/content.
- Pixel-perfect parity with an unspecified historical output unless validated against an agreed golden output (see Definition of Ready).
- Defining BI/business meaning of metrics/categories/TA beyond names/filters present in configuration and data.

## 3. Confirmed facts (evidence-based)

| Area | Confirmed fact | Evidence |
|---|---|---|
| Request source | ТЗ describes `treport.xlsm` macro with `settings`/`run`, tables `general,reports,data_files,placeholders`, buttons `Refresh_data`/`Create_report`, workflows, constraints. | `docs/description.docx` (extracted text) |
| Workbook container | `treport.xlsm` exists; contains sheets `run`, `settings`; 4 table parts; drawings/vml/control props; `xl/vbaProject.bin` present. | `data/theiss/treport.xlsm`; `xl/workbook.xml`, `xl/worksheets/sheet1.xml`, `xl/worksheets/sheet2.xml`, `xl/tables/table1-4.xml`, package parts listing |
| Config tables (structure) | `general(parameter,value)`; `reports(category,ta,name)`; `placeholders(placeholder,description)`; `data_files(data_file,data_sheet,report_sheet,paste_row,paste_col,freeze_row,freeze_col,zoom,group_rows)`. | `xl/tables/table1.xml`–`table4.xml` |
| Config snapshot (workbook) | `general`: `input_path` = workbook folder formula, `output_path` = `B2&"output"`, `category_delimeter` = `&`, `report_week` = `last`, `report_year` = `last`, plus `new_file_delimeter`, `keep_col_hidden` (TRUE/FALSE). `reports`: 6 rows (e.g. `ЗУБНАЯ ПАСТА&ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` / `W 50+` / `LACALUT ACTIV`, etc.). `data_files`: 3 rows — `report_m`/`clean_data_full`/`помесячное сравнение`, `report_w`/`clean_data_full`/`сравн конкур_<weeknum>w_<year>_прямая` (group_rows TRUE), `report_wclip`/`clean_data_full`/`хроно по неделям`. `placeholders`: `weeknum`, `monthnum`, `year` with descriptions. | `xl/worksheets/sheet2.xml`, `xl/sharedStrings.xml` |
| Data artifacts present | Candidate files by base name exist: `data/theiss/monthly/report_m.xlsx`, `data/theiss/weekly/report_w.xlsx`, `data/theiss/base/report_wclip.xlsx` (unique by exact name in observed locations). No `data/theiss/output` folder observed. | Directory listings / repo scan |
| Source workbook indicators | Inspected candidate xlsx files contain pivot cache/table, connections/query-table and customXml indicators (complexity varies; e.g. `report_wclip` heavier). Exact pivot fields/schemas not fully extracted. | Package part-name inspection (research) |
| Repo state | No exported VBA modules; `mko_treport/src/{scripts,vbs,xls}` empty; `data/*` git-ignored. | Repo scan, `.gitignore` |

## 4. Scope

### 4.1 In scope
1. **Configuration reading & validation** from `treport.xlsm` `settings` sheet ListObjects: `general`, `reports`, `data_files`, `placeholders` (including parameters present in workbook; unspecified extras must be handled deterministically per decisions).
2. **`Refresh_data`** operation: locate configured data files in `input_path`, refresh their Power Query/pivot data as required, persist refreshed data, and report outcome.
3. **`Create_report`** operation: resolve reporting period/filters, create/overwrite output report workbooks per `reports`, generate sheets in `data_files` order from filtered pivot data, apply required formatting/placement/naming, save outputs, and report outcome.
4. **Deterministic rules** for file resolution, “last” period resolution, filter semantics/placement, multi-value matching, copy range, placeholders/naming, idempotent overwrite, and fail-fast error behavior as specified below.
5. **Operational/user-visible outcomes**: clear success/failure messaging sufficient to identify the failing file/sheet/field/config element (conceptual requirement).

### 4.2 Out of scope
- Any non-Excel runtime, scheduler/orchestrator, UI redesign, or changes to source data content/pipelines.
- Recovery/merge of concurrent edits, multi-user locking workflows beyond deterministic failure on locked resources.
- Historical-data correction, metric definition changes, or new report types not representable by existing config model (unless separately requested).

## 5. Constraints

- **Artifact/architecture constraint:** Solution operates within the existing `treport.xlsm` Excel-macro model and configured `settings`/`run`/ListObjects.
- **Access modes (PO-confirmed):** `Create_report` — data files opened **read-only, never saved**. `Refresh_data` — data files opened **writable (not read-only)** to refresh and **save**, then closed. (Resolves ТЗ conflict in favor of a mode-specific policy.)
- **Processing order:** `data_files` rows processed in configured order; output sheets created in the same order.
- **Filter placement:** if a filter field already exists in Filters/Rows/Columns, filter it in place without moving it; if absent, add to Filters and filter there.
- **No invention:** unspecified business meanings must not be fabricated; unresolved technical feasibilities are recorded as validations/risks, not assumed behavior.
- **Determinism:** identical resolved inputs/config/data state must yield deterministic resolution, naming, and structure.

## 6. Product Owner decisions (confirmed — all recommended)

| ID | Decision area | Confirmed choice (recommended) |
|---|---|---|
| Q1 | Read-only vs refresh-save | `Create_report`: read-only, never save data files. `Refresh_data`: writable open to refresh PQ/RefreshAll and save, then close (only mutating operation on data files). |
| Q2 | `report_week`/`report_year = "last"` | Resolve from **actual data** in relevant `data_file/data_sheet`: `report_year` = maximum available year; `report_week` = maximum available week within that resolved year. Resolution must be **deterministic** when multiple data files differ. |
| Q3 | Week boundary | Include weeks **<= `report_week`** and exclude weeks **> `report_week`** within selected `report_year` (`report_week` included). |
| Q4 | Data file name resolution | `data_file` is base name; match `<data_file>.<supported_ext>` (supported extensions explicitly listed, e.g. `.xlsx/.xlsm/.xlsb`) with a **defined precedence**; **error** if none or if more than one match. |
| Q5 | Output/overwrite/idempotency | One output workbook per `reports` row, name `reports.name` (+ `.xlsx` unless otherwise specified). If target exists, **overwrite** on successful run. Sheets created in `data_files` order; target sheets **recreated/replaced** for idempotency. |
| Q6 | Multi-value filter matching | `category` may contain multiple values separated by `general.category_delimeter`; select **all listed** matching items. Matching is **exact** to pivot item captions after trimming whitespace. `ta` is treated as a filter field per config with the same matching/placement principles. |
| Q7 | Copied data range | Copy pivot’s presented data range including **row/column labels and grand totals** as shown after filtering; **exclude** pivot object/page-filter area and any PQ connections. Paste **values + formats** (and column widths if needed) as static content. |
| Q8 | Errors / empty filter result | **Fail fast** with a clear, contextual error for: missing/invalid config or required paths, missing/unresolvable data file, missing `data_sheet`, missing pivot table, or missing required filter field. A filter yielding **no data** must produce an **explicit, defined outcome** (policy fixed at implementation/spec sign-off: fail-fast preferred for correctness, or an explicit marked-empty sheet — must be deterministic and documented). |
| Q9 | Formatting & extra params | Blanks = default/no-op for `freeze_row/freeze_col/zoom/paste_row/paste_col/group_rows`. `group_rows=TRUE` requires a **deterministic** rule (color-band grouping “от белого в центре” is ambiguous and must be explicitly defined before relied upon). Extra workbook params not specified in ТЗ (`new_file_delimeter`, `keep_col_hidden`) must be **explicitly specified or explicitly ignored** with deterministic behavior. |
| Q10 | Placeholders & naming | `weeknum`, `monthnum`, `year` resolved from the resolved report period and substituted at least in `report_sheet`. Handle invalid Excel sheet-name characters, **31-char** limit, and collisions **deterministically**. Token syntax/formatting must be explicit and consistent. |
| Q11 | Scope/source of truth | Behavior aligned to **ТЗ intent** reconciled with **current workbook structure** (actual table columns/rows/params included once confirmed). Both `Create_report` and `Refresh_data` are in scope. |
| Q12 | Acceptance context | Validation requires target Excel/OS context and, for correctness sign-off, representative `input_path` data plus at least one **golden** expected output for a fixed `report_week/report_year` (filters/values/formatting). Structural-only acceptance is insufficient for correctness claims. |

## 7. Confirmed functional requirements

### 7.1 Configuration model (FR-CFG)
- **FR-CFG-1:** Read settings from ListObjects `general`, `reports`, `data_files`, `placeholders` on the `settings` sheet of `treport.xlsm`.
- **FR-CFG-2:** `general` provides at minimum: `input_path`, `output_path`, `category_delimeter`, `report_week`, `report_year`. `input_path` and `output_path` are mandatory and non-empty; missing/empty required values cause a clear validation error.
- **FR-CFG-3:** `reports` defines report generation rows with at least `category`, `ta`, `name` (one output workbook per row per Q5/Q11).
- **FR-CFG-4:** `data_files` defines ordered source-processing rows with at least `data_file`, `data_sheet`, `report_sheet` and formatting/placement fields (`paste_row`, `paste_col`, `freeze_row`, `freeze_col`, `zoom`, `group_rows`) as present in workbook; row order is authoritative for processing/sheet order.
- **FR-CFG-5:** `placeholders` defines available substitution parameters (at least `weeknum`, `monthnum`, `year`) and their meaning; resolution follows Q10/Q2.
- **FR-CFG-6:** Workbook parameters not described in ТЗ (`new_file_delimeter`, `keep_col_hidden`) and any blank optional fields must follow a documented deterministic policy (use, ignore, or reject) — see Open items/Risks; silent undefined behavior is not acceptable.

### 7.2 File resolution and access (FR-FILE)
- **FR-FILE-1:** Resolve each `data_files.data_file` (base name) within `general.input_path` to `<data_file>.<supported_ext>` using a defined extension set and precedence (Q4). Case-insensitive matching is acceptable on Windows but outcome must be deterministic.
- **FR-FILE-2:** If no match or more than one match exists for a configured data file, fail fast with an error listing the configured name, searched path, and cause (Q4/Q8).
- **FR-FILE-3:** `Create_report` opens data files read-only and never saves them (Q1).
- **FR-FILE-4:** `Refresh_data` opens data files writable (not read-only), performs refresh, saves, and closes (Q1). File/lock unavailability must result in a clear failure.
- **FR-FILE-5:** Ensure `output_path` exists (create if missing, per ТЗ intent) before writing outputs; inability to create/write fails fast (Q8).

### 7.3 Refresh operation (FR-REF)
- **FR-REF-1:** For each resolved data file (in deterministic order), open, refresh Power Query and dependent pivot data so refreshed results are persisted, save, and close (Q1).
- **FR-REF-2:** Refresh must complete before save/close; asynchronous/background refresh states must not leave persisted data stale (technical validation required).
- **FR-REF-3:** On per-file refresh failure, fail fast with a clear error identifying the file and stage (Q8); application/workbook state must not be left corrupted by the run.

### 7.4 Period and filter resolution (FR-FLT)
- **FR-FLT-1:** Resolve `report_year` and `report_week` from `general`: explicit values used as provided; value `last` resolved from actual data in relevant `data_file/data_sheet` as `max(year)` and `max(week)` within that year, deterministically across data files (Q2). Resolved values are the reporting period for the run and for placeholder substitution.
- **FR-FLT-2:** Apply period filter as **<= `report_week` included, > `report_week` excluded** within selected `report_year` (Q3).
- **FR-FLT-3:** Apply `category` and `ta` filters from `reports` rows as configured. Multi-value `category` is split by `category_delimeter`, tokens trimmed, and all exactly matching pivot items selected (Q6). `ta` follows configured field/matching semantics.
- **FR-FLT-4:** Respect filter placement: filter in place if the field is already in Filters/Rows/Columns; if absent, add to Filters then filter (ТЗ/Q6). Do not move fields that are already placed.
- **FR-FLT-5:** Missing `data_sheet`, missing pivot, or missing required filter field is a fail-fast error with context (Q8). Filter resulting in no data yields the defined explicit outcome (Q8, deterministic).

### 7.5 Report generation and output (FR-RPT)
- **FR-RPT-1:** Create one output workbook per `reports` row named `reports.name` (+ `.xlsx` unless specified). If it exists, overwrite on successful run (Q5). Overwrite/locking failures are clear errors (Q8).
- **FR-RPT-2:** For each data file in `data_files` order, create a target sheet in each relevant report workbook named from `report_sheet` after placeholder substitution (Q10), placed/ordered according to `data_files` sequence. For idempotent reruns, target sheets are recreated/replaced deterministically (Q5).
- **FR-RPT-3:** Determine the copy range from the pivot’s presented data after filtering: include row/column labels and grand totals as shown; exclude the pivot object/page-filter area and PQ connections (Q7).
- **FR-RPT-4:** Paste into the target sheet at `paste_row/paste_col` (blank = default per Q9) as static **values + formats** (and column widths if needed). Output must not contain pivot objects or PQ connections from the source.
- **FR-RPT-5:** Apply formatting per `data_files`: freeze panes (`freeze_row/freeze_col`), zoom, hide gridlines, and row/column sizing/state as required by ТЗ; blanks are no-op/default (Q9).
- **FR-RPT-6:** `group_rows=TRUE` must apply a deterministic grouping/outline rule. Because color-band grouping (“от белого в центре”) is not fully defined, a specific, approved rule is required before this behavior can be accepted as correct (Q9) — see Open items.
- **FR-RPT-7:** Save each report workbook to `output_path` and release resources; opened source files must be closed according to mode (Q1).

### 7.6 Naming and placeholders (FR-NAME)
- **FR-NAME-1:** Resolve `year`, `weeknum`, `monthnum` from the resolved report period and substitute into `report_sheet` (at minimum) per Q10. `monthnum` derivation must follow an explicit defined rule (see Open items).
- **FR-NAME-2:** Final sheet/report names must be sanitized for Excel invalid characters, limited to 31 characters, and collisions resolved by a deterministic, documented policy (Q10). Empty/invalid resulting names are errors.

### 7.7 Run behavior and diagnostics (FR-RUN)
- **FR-RUN-1:** `Refresh_data` and `Create_report` are invokable from `run` sheet controls as provided by the workbook.
- **FR-RUN-2:** Runs execute without unnecessary UI disruption and restore application state on completion/failure (quality requirement).
- **FR-RUN-3:** Success and failure outcomes are explicit. Errors are fail-fast and identify context (operation, configured row/name, file path, sheet, field/pivot/filter/value as applicable) per Q8.

## 8. Conceptual development tasks (for planning decomposition)

| ID | Task | Purpose | Expected outcome | Dependencies |
|---|---|---|---|---|
| T1 | Config contract & validation | Finalize meaning/defaults/validation of all config fields (incl. workbook extras). | Documented config schema and validation/error matrix. | PO decisions; confirm extra-param policy |
| T2 | Data-file resolution & access policy | Deterministic path/extension resolution and mode-specific open/close/save behavior. | Resolution rules + access-mode rules per operation. | T1, Q1/Q4 |
| T3 | Refresh workflow | Reliable refresh-and-persist behavior for PQ/pivot sources. | Defined refresh sequence, completion criteria, failure handling. | T2, Q1; R2 validation |
| T4 | Period resolution (`last`) & placeholders | Deterministic `year/weeknum/monthnum` derivation and substitution. | Resolved-period algorithm and placeholder formatting/naming rules. | T1, Q2/Q10; define `monthnum` & aggregation |
| T5 | Filter engine semantics | Apply year/week(<=)/category/ta with placement and multi-select rules across orientations. | Filter behavior spec per field type/location and no-data outcome. | T1,T4, Q3/Q6/Q8; R2/R3 validation |
| T6 | Pivot snapshot extraction & paste | Static, correct data+format capture and placement without pivot/PQ artifacts. | Copy-range definition and paste/width/height behavior. | T5, Q7; golden validation |
| T7 | Formatting, layout & grouping | Freeze/zoom/gridlines/row-col sizing/outline with deterministic grouping. | Formatting acceptance criteria; `group_rows` rule fixed. | T1,T6, Q9; define `group_rows` |
| T8 | Report assembly & idempotency | Workbook/sheet creation, ordering, overwrite/replace, naming/collision handling. | End-to-end output lifecycle rules. | T1,T4,T6, Q5/Q10 |
| T9 | Diagnostics & run control | Fail-fast errors, messages, and state-safe execution. | Error taxonomy/messages and run outcomes. | T2–T8, Q8 |
| T10 | Verification & acceptance | Golden-data scenarios and sign-off criteria. | Test scenarios and Definition of Done inputs. | T1–T9, Q12, representative+golden data |

## 9. Research summary (evidence-backed synthesis)

> Research tasks executed in-session: artifact/state inspection (R1-equivalent, partial due to binary/VBA/tool limits) and VBA/automation approach research (R2/R3-equivalent). Findings are actionable but include items requiring validation against representative files and target Excel.

### 9.1 Current-architecture findings
- **Macro workbook:** `data/theiss/treport.xlsm` contains `run`/`settings`, 4 ListObjects with expected names/columns, control/drawing artifacts, and `xl/vbaProject.bin`; macro logic is not decompiled/reviewable from repo (no exported VBA). This supports the Excel/VBA-in-place architecture (Q11).
- **Config reality vs ТЗ:** workbook includes params/values beyond a minimal reading of ТЗ (e.g. `new_file_delimeter`, `keep_col_hidden`, `paste_row/paste_col`, `group_rows`, `report_week/year=last`, specific report/data/placeholder rows). These must be governed explicitly (Q9/Q11).
- **Data files:** `report_m/report_w/report_wclip` exist in `monthly/weekly/base` with pivot/query/connection indicators; `report_wclip` appears structurally more complex. Exact pivot names, field captions/types, and query topology are not confirmed here.
- **Gaps:** no `output` folder; no golden outputs; VBA behavior and several data semantics unverified.

### 9.2 Feasible implementation approaches (top options, compatible with Excel/VBA)

| Option | Description | Trade-offs |
|---|---|---|
| **A. Straight-through procedural VBA** | Single procedural flows mirroring spec steps with minimal abstraction. | **Pros:** low initial effort, direct fit. **Cons:** lower maintainability; higher regression risk for Q1/Q2/Q4/Q5/Q8/Q10 determinism; harder to optimize/validate for large/complex pivots and PQ timing. |
| **B. Modular component-based VBA** | Separate responsibilities (config/validation, file resolution, period resolution, app-state/workbook session, pivot/filter service, range export, formatting, report builder) with centralized cleanup. | **Pros:** high maintainability; clearer enforcement of access modes and rules; easier targeted optimization and review. **Cons:** moderate upfront effort; still bounded by Excel OM limits. |
| **C. Plan-then-execute with per-file isolation & auditable logging (on B)** | Two-phase run: preflight + resolve + immutable execution plan (paths, resolved period, sanitized names/collisions, filter specs) then isolated execution (esp. per data file) with structured progress/errors; optional validate/dry-run (no writes). | **Pros:** strongest determinism/auditability and operational reliability; surfaces Q8 issues early; predictable reruns (Q5). **Cons:** moderate effort and small planning overhead; Excel/OM unknowns remain (PQ topology, OLAP, color-grouping). |

### 9.3 Recommended approach (clear recommendation)
- **Recommendation:** **Option C implemented on the modular foundation of Option B**.
- **Rationale:** best aligns with confirmed constraints — deterministic “last” resolution and naming (Q2,Q10), strict mode-specific read/write behavior (Q1), ordered/idempotent outputs (Q5), exact/placed filtering (Q3,Q6), fail-fast contextual errors (Q8), and the need for preflight/validation against representative and golden data (Q12). It reduces common Excel automation failures (stale refresh state, partial UI/state corruption, ambiguous file/name resolution) while keeping the solution within the existing `treport.xlsm` macro architecture.

### 9.4 Key technical practices to apply (from research, to inform planning)
- **State & access:** snapshot/restore application state; avoid UI selection; enforce Q1 at open boundary; track macro-opened workbooks; never save sources in `Create_report`.
- **PQ/pivot refresh:** force synchronous completion before dependent pivot refresh/save (e.g. disable background query where applicable, wait for async queries) and refresh pivot caches after data is current; treat ordering/completion as a validation point (Q1/FR-REF).
- **Pivot/filtering:** identify pivot/fields deterministically; filter in place by orientation without moving fields; support multi-select exact-match (Q6) and week threshold `<=` (Q3) with manual-update/batching considerations; handle “no items visible” explicitly (Q8).
- **“Last” resolution:** compute from actual data in relevant `data_file/data_sheet` per a single documented aggregation rule across files (Q2); log resolved period.
- **Static copy:** capture presented range (labels/totals, no pivot/page/PQ) and paste static values+formats (+widths) (Q7); copy row heights/hidden state efficiently (run-length) and apply deterministic outline/grouping.
- **Naming:** sanitize invalid chars, enforce 31-char limit, deterministic collision policy, and explicit placeholder formatting (Q10).

## 10. Risks and key unknowns (requiring validation)

| ID | Risk / unknown | Impact | Mitigation / validation needed |
|---|---|---|---|
| R1 | **VBA logic opaque** (`xl/vbaProject.bin` not decompiled/exported) | Current behavior may diverge from ТЗ/this spec | Inspect in Excel/VBA IDE or approved extraction; reconcile before/with implementation; consider exporting modules for review. |
| R2 | **Pivot topology & field identity** (which pivot, exact field captions/types for year/week/category/ta, orientation, regular vs OLAP/Data Model) | Filters/copy may target wrong object or fail | Validate on representative `report_m/report_w/report_wclip` (and real inputs); fix deterministic identification rules. |
| R3 | **PQ refresh completeness/timing** across Excel builds and query types | Stale persisted data after `Refresh_data` | Validate synchronous refresh + completion waiting + cache refresh on target Excel; confirm all relevant queries/load modes refreshed. |
| R4 | **“Last” aggregation across multiple data files** and definition of “actual data” (cache/items vs source) | Inconsistent period resolution | Fix one deterministic aggregation rule and data source; log and test with differing files. |
| R5 | **`monthnum` derivation** (week can span months/years) | Incorrect placeholder/sheet naming | Define explicit business/technical rule or source for `monthnum`; add boundary tests. |
| R6 | **`group_rows` color-band rule** (“от белого в центре”) undefined; `DisplayFormat`/conditional color limits | Non-deterministic or infeasible grouping | Obtain explicit rule (inspected column(s), color→level mapping, boundaries) or approve deterministic alternative (e.g. outline-level copy); do not guess. |
| R7 | **No-data outcome policy** (fail-fast vs marked-empty sheet) not fixed to a single option | Misleading outputs or unexpected aborts | Select and document one deterministic policy; reflect in acceptance cases. |
| R8 | **Sheet-name collision policy** and placeholder token/format details | Overwritten/confusing sheets or Excel errors | Fix deterministic collision handling and exact token syntax/formatting. |
| R9 | **Extra/ambiguous config params** (`new_file_delimeter`, `keep_col_hidden`) and exact `ta` semantics per workbook vs ТЗ text | Silent misbehavior or config drift | Confirm meaning: specify or explicitly ignore; document. |
| R10 | **Copy fidelity** (range bounds incl. totals/page area, number/date fidelity, widths/heights/hidden/outline, conditional formats appearance) | Output differs from intended | Validate against golden outputs on representative filtered cases. |
| R11 | **File resolution details** (supported extensions and precedence order) and locked/open files | Wrong file picked or run failures | Fix ordered extension list/precedence; define locked-file failure behavior. |
| R12 | **Performance on large files/many pivot items** | Long/unresponsive runs | Benchmark with representative data; apply batching/manual-update and progress feedback (technical). |
| R13 | **Environment prerequisites** (Excel version, Windows, macro trust, PQ enabled, local/network paths) | Non-runnable or inconsistent behavior | Document and validate target environment (Q12). |

## 11. Assumptions

- **A1:** `treport` remains an Excel/VBA solution in `treport.xlsm` (Q11).
- **A2:** Supported data-file extensions and precedence will be explicitly defined (recommended set includes `.xlsx`, `.xlsm`, `.xlsb`) and applied per Q4.
- **A3:** “Last” aggregation rule across multiple data files is a single deterministic max-based rule as stated in Q2 unless a different explicit rule is provided.
- **A4:** `ta` in `reports` is a pivot filter field applied with the same placement/matching principles as other configured filters; exact field caption(s) will be confirmed from data (R2).
- **A5:** One output workbook per `reports` row; `reports.name` is the output base name (Q5).
- **A6:** Placeholder tokens are the configured names `weeknum`, `monthnum`, `year` as present in the workbook; exact delimiter/syntax/zero-padding will be standardized and documented (Q10).
- **A7:** A clear, deterministic no-data outcome and collision policy will be fixed (recommendation: fail-fast for no-data/collisions for correctness unless stakeholders require marked-empty output) (Q8/Q10).
- **A8:** `group_rows` will be implemented only under an explicitly defined deterministic rule (color-band rule to be specified) or treated as no-op/blocked until specified (Q9).

## 12. Open questions

### 12.1 Business/product (must be confirmed before acceptance of affected behavior)
- **OQ-B1 (Q8 detail):** For “filter yields no data”, is the accepted outcome **fail-fast error** (recommended) or a **generated sheet with explicit ‘NO DATA’ marker/context**?
- **OQ-B2 (Q10 detail):** Exact **sheet-name collision** policy — fail-fast (recommended) or deterministic auto-suffix?
- **OQ-B3 (Q9):** Approved deterministic rule for `group_rows=TRUE` (color-band mapping, source column(s), “от белого в центре”) or decision to disable/ignore until defined.
- **OQ-B4 (Q10):** Definitive definition/format of `monthnum` relative to resolved `report_week/report_year` (and token syntax/zero-padding).
- **OQ-B5 (Q9):** Business meaning and desired handling of `new_file_delimeter` and `keep_col_hidden` (specify behavior or explicitly ignore).
- **OQ-B6 (Q4):** Confirm the ordered list and precedence of supported file extensions for `data_file` resolution.
- **OQ-B7 (Q12):** Provide/confirm target Excel version & OS, representative input dataset, and at least one **golden** output for a fixed period to sign off correctness.

### 12.2 Technical validations (not business ambiguity, but must be closed for implementation confidence)
- **OQ-T1:** Identify actual pivot(s), field captions/types/orientations, and PQ/connection topology in `report_m/report_w/report_wclip` (and real inputs).
- **OQ-T2:** Verify synchronous PQ refresh + pivot cache refresh completeness on target Excel.
- **OQ-T3:** Confirm copy-range boundaries and formatting fidelity (totals, widths/heights/hidden/outline, value/date fidelity) against golden output.
- **OQ-T4:** Inspect/reconcile `xl/vbaProject.bin` behavior with this spec where relevant.

## 13. Out of scope (explicit)

- Implementation, code changes, VBA authoring/refactoring, or repository/data modifications.
- New architecture, non-Excel tooling/pipelines, scheduling/orchestration services.
- Changes to source data preparation, metric definitions, or upstream PQ/data models.
- Performance SLAs beyond “works correctly on representative agreed data” unless separately defined.

## 14. Definition of Ready (for implementation planning)

This specification is **ready for implementation planning** when the following are satisfied:

1. **Business decisions confirmed:** Q1–Q12 accepted as recorded (satisfied — PO selected all recommended) **and** the remaining product choices in §12.1 (OQ-B1–B7) are answered/fixed (policy selections and golden-acceptance inputs provided).
2. **Scope & requirements stable:** in-scope operations (`Refresh_data`, `Create_report`), config contract, access modes, period/filter/copy/naming/idempotency rules, and fail-fast/error expectations are agreed per §§4–7.
3. **Deterministic policies fixed:** extension precedence (Q4), “last” aggregation rule (Q2), `monthnum` rule and token format (Q10), no-data outcome (Q8), sheet-name collision handling (Q10), and `group_rows` / extra-param handling (Q9).
4. **Validation baseline available:** target Excel/OS confirmed and at least one representative dataset with a golden expected output for a fixed period is available (Q12).
5. **Known risks acknowledged:** technical validations (§12.2, §10) are tracked as implementation/verification tasks with owners, not treated as unknown business requirements.

**Planning gate (minimum to start implementation work):** §14 items 1 (including OQ-B1–B6 decisions) and 2–3 completed; items 4–5 may run in parallel as explicit validation tasks but must be closed before correctness sign-off.

## 15. Traceability summary

| ТЗ / need | Spec coverage |
|---|---|
| Settings tables `general/reports/data_files/placeholders` | FR-CFG, §3, §6 (Q11) |
| `Refresh_data` (refresh PQ/RefreshAll, save/close) | FR-REF, FR-FILE, Q1 |
| `Create_report` (reports per row, sheets per `data_files` order) | FR-RPT, FR-NAME, Q5/Q10 |
| Read-only handling of data files | FR-FILE, Q1 (mode-specific resolution) |
| Filter placement & period/category/ta rules | FR-FLT, Q2/Q3/Q6 |
| Copy without pivot/PQ, values+formats | FR-RPT, Q7 |
| Formatting (paste/freeze/zoom/gridlines/row-col/grouping) | FR-RPT, Q9 |
| Missing files/config errors | FR-FILE/FR-FLT/FR-RUN, Q8 |
| Placeholders in names/sheets | FR-NAME, Q10 |

---
*End of specification.*

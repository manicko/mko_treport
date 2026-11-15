# treport (001_treport, rev.5) — Implementation Status, Problems, and Next Steps

- **Report date:** 2026-11-15 (workspace evidence collected ~2026-11-14T22:09Z–2026-11-15T07:35Z)
- **Repository:** `C:\py_exp\mko_treport`
- **Baseline documents:** `.ai/problems/001_treport_spec.md` (rev. 5, 2026-11-14), `.ai/plans/001-treport-execution-plan.md` (rev. 5)
- **Purpose:** factual status of VBA implementation work performed in this session, observed problems/blockers, and an ordered path to a functional, verifiable rev.5 solution.

## 1. Executive summary

- **Done (verified in artifacts/COM):** B1 control surface is implemented in the actual `data/theiss/treport.xlsm` working runtime artifact (exactly two `run` controls `Refresh_data`/`Create_report` bound to same-named VBA procedures; legacy `RunXlconcat` control/binding removed from `run`). VBA runtime modules for config, file resolution, filters/min-max resolution, pivot/filter handling, refresh scaffold, extraction/formatting/output assembly, state/diagnostics, and real `Refresh_data`/`Create_report` entry points were implemented and imported into the workbook; tracked exports/bin/workbook copy were updated and committed.
- **Not done / not proven:** no successful end-to-end `Create_report` or `Refresh_data` run has been confirmed on actual source data. Latest automated `Create_report` probe (`.ai/tmp/treport_live_test3.xlsm` → `.ai/tmp/live_output3`) remains **`settings!Z10=RUNNING`** with **no output files** and Excel processes still present (`5520`, `19828` observed) — outcome unknown (possible long run or hang). B2 runtime validations (OQ-T1–T4), B13 functional verification, and B14 review are incomplete.
- **Key blockers:** unresolved config collision `DOLGIT` vs `DOLGIT ` (D14 fail-fast) in canonical reports config, unconfirmed end-to-end execution/performance on large source data (especially `report_wclip.xlsx` ~130 MB), and lack of runtime-validated PQ refresh/pivot presented-range fidelity.

## 2. Baseline and constraints (rev.5, must remain true)

- **Runtime:** final solution VBA-only inside `treport.xlsm`; end user must not need Python/PowerShell/BAT/external runtime. Auxiliary dev/inspection/test tools are allowed but non-runtime.
- **Run controls:** exactly two buttons/entry points `Refresh_data` and `Create_report`; legacy `RunXlconcat`/obsolete workflow removed from `run`.
- **Decisions fixed:** D1–D18 (access modes, deterministic file resolution `.xlsx/.xlsm/.xlsb`, one output per report, filter DSL/staging/placement, `min/max` + `max_*` dependency-aware resolution, `Trim`, generalized `articleLevel4Name_clean` multi-value behavior (matched subset if any configured token matches; if none match warn and skip that report, continue others, no whole-run failure), fail-fast diagnostics, presented-range static copy, naming/sanitization/31-char/collision rules, no golden workbook, `group_rows` deferred/no-op).
- **Evidence-based alignment:** `data_files.filters` use `Year` for all rows (user-updated, aligned); `report_wclip` has grouped `Years` in addition to base `Year` — runtime behavior must be validated, no guessed mapping.
- **Local artifacts:** `data/theiss` workbooks are local/gitignored; modify on working copies/backups, keep originals, do not commit large Excel/generated outputs.

## 3. Work completed (factual)

### 3.1 B1 — Control/workbook preparation (runtime workbook)
- Canonical working baseline selected: `data/theiss/treport.xlsm` (SHA256 `3394bbf42e9833c2526b6d98ffc0cb57220c3fa0a9d5fe15f78e343b6b9a622e`, 78439 B at baseline). Compared with `mko_treport/src/xls/treport.xlsm` (SHA256 `f6445f75…`, 78432 B, older). Backups created at `.ai/tmp/b1_backup_20261115-005533/`.
- Excel 16 COM (Office16) used on working copy to: import `modTreportRun`, remove legacy run shape/control bound to `RunXlconcat`, add two Form Controls on `run` (`Refresh_data` → `modTreportRun.Refresh_data`, `Create_report` → `modTreportRun.Create_report`), save.
- COM-verified: `run` has 2 shapes with correct names/OnAction/text; `VBComponents` include `modTreportRun`; `Refresh_data`/`Create_report` procedures present; settings ListObjects `general,reports,placeholders,data_files` present/unchanged. Legacy binding absent from `run`; `xlconcat_macro` module remains in project/exports but not bound to `run`.
- Modified workbook synced to `data/theiss/treport.xlsm` and `mko_treport/src/xls/treport.xlsm`; `vbaProject.bin` refreshed from workbook package; VBA components exported.
- **Parity note (OQ-T4):** exports are auxiliary; control shapes/bindings are authoritative in the workbook binary and may not be fully represented in `.bas/.cls`.

### 3.2 B2 — Technical investigation (partial)
- Read-only auditor/researcher analysis produced: pivot/field/layout/member evidence for `report_m/report_w/report_wclip` (base `Year` in all; `report_wclip` has grouped `Years`; `Month Number`, `Week Number`, `articleLevel4Name_clean`, `targetAudience` present; `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` not observed as distinct cache item in inspected evidence (historical example; consistent with generalized rev.6 handling when other configured tokens match)), PQ/external-cache presence, and VBA technique guidance for sync refresh/filter/copy/state handling. Local research artifact: `.ai/tmp/context_b2_research.md` (referenced in session).
- Direct COM/ZIP inspections (read-only) performed against `data/theiss/monthly/report_m.xlsx`, `weekly/report_w.xlsx`, `base/report_wclip.xlsx` and workbook config.
- **Not completed:** runtime confirmation of PQ completion/persistence, actual presented-range bounds/fidelity, and `Year`-vs-grouped-`Years` behavior in `report_wclip` during live filtering/extraction (OQ-T1–T3).

### 3.3 B4–B12 — VBA implementation (initial pass, imported to workbook)
Modules added/implemented in `mko_treport/src/vbs/export/` and imported into workbook/project:
- `modTreportConfig.bas` — `Norm/Trim`, config types, ListObject/sheet lookup, required schema/value validation (`general/reports/data_files/placeholders`), contextual `Fail`, input/output path normalization helpers, `group_rows` read/no-op.
- `modTreportFiles.bas` — deterministic base-name resolution with precedence `.xlsx,.xlsm,.xlsb`; **recursive match search under `input_path` subfolders** (supports `monthly/weekly/base` layout) with ambiguity/fail-fast checks; read-only/writable open, close/save helpers (D1,D2).
- `modTreportFilters.bas` — filter DSL parse/classify (`|`, `field = expr`, scalar/range/list, `min/max`, inclusive numeric ranges, `&` lists), duplicate/empty checks, dependency-aware resolution (resolve/apply `Year` context first when present, derive `max_monthnum`/`max_weeknum`), placeholder-bound resolution helpers, report filter splitting. **Historical note:** token-specific optional detection for `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` existed in earlier snapshots; current exports implement generalized `articleLevel4Name_clean` matched-subset/report-skip behavior (rev.6).
- `modTreportPivot.bas` — sheet/pivot/field discovery, numeric/trim member matching, placement (keep orientation if set, else page/filter) (D7), selection application/reset with generalized `articleLevel4Name_clean` matched-subset/report-skip signaling (rev.6) and reduced forced refresh during toggling, pivot data-empty heuristic.
- `modTreportRefresh.bas` — refresh scaffold: set non-background where possible, `RefreshAll`, `CalculateUntilAsyncQueriesDone`, wait/timeout loops, pivot cache/table refresh attempts (B7, OQ-T2 pending validation).
- `modTreportExtract.bas` — presented range selection (`TableRange1` fallback), static copy/paste (values+number formats, formats, column widths, row heights), sheet formatting (gridlines off, freeze panes, zoom handling).
- `modTreportOutput.bas` — placeholder substitution in `report_sheet` only (no zero-padding), `Trim`/sanitize names, 31-char/collision checks, output workbook/sheet create/replace, save/overwrite helpers.
- `modTreportState.bas` — app state capture/apply/restore.
- `modTreportRun.bas` — real `Refresh_data`/`Create_report` procedures: config-driven per-data-file processing in order, file/report filter staging, per-report snapshot into per-report workbooks/sheets, contextual fail-fast error handling and state cleanup, success messages (UI) / result signaling (in test harness variants).

### 3.4 Auxiliary/cleanup
- Stale auxiliary Python treport tests targeting removed `mko_treport/src/scripts/treport_*` modules removed from `mko_treport/tests/` (commit `cf2c4e3`). `treport_filter_application` was empty.
- Source workbooks backed up before risky tests: `.ai/tmp/source_backup_20261115-013539/` (`report_m.xlsx`, `report_w.xlsx`, `report_wclip.xlsx`).

## 4. Git and workspace state (observed)

### 4.1 Relevant commits
- `57ae40e` — `fix(treport): improve file resolution, optional filter handling, and pivot refresh stability` (historical; later generalized rev.6 `articleLevel4Name_clean` rule applied in source exports)
- `cf2c4e3` — `chore(treport): remove stale auxiliary Python tests for removed runtime modules`
- `fe2c05b` — `feat(treport): implement VBA config, refresh, and report generation pipeline`
- `dba8639` — `feat(treport): complete B1 run controls in workbook`
- `03aedbb` — `chore(treport): add treport workbook binary`
- `544108d` — `feat(treport): add rev5 run entry-point stubs (B1)`
- `2027281` — `Update treport spec and execution plan in .ai`

### 4.2 Current working tree (`git status --short`)
- ` M .ai/plans/001-treport-execution-plan.md` — modified (not staged by this report run; contains earlier plan/task edits).
- ` D mko_treport/src/__init__.py` — deleted (pre-existing working-tree state; not part of this implementation).
- ` D mko_treport/src/vbs/export/xlconcat_macro.bas` — **deleted in working tree but present in HEAD** (unintended risk; legacy module intentionally kept unbound for reference — see Problems).
- `?? .ai/tasks/` — untracked.
- `?? .ai/plans/002-treport-implementation-status-and-next-steps.md` — this report (new).

## 5. Runtime test evidence (this session)

| Test / action | Environment/artifact | Result | Notes |
|---|---|---|---|
| B1 workbook control/VBA inspection | Excel 16 COM on modified workbook | **Pass (structural)** | 2 controls/bindings correct, legacy run binding removed, procs present, settings tables intact. |
| `modTreportConfig.Norm` / `ReadConfig` probes | Excel COM, temp copies | **Pass** | Config read correctly (6 reports, 3 data_files, `&` delimiter, expected filters/placeholders). |
| `Create_report` probe (silent harness) | temp workbook, output to `.ai/tmp/*` | **Fail (expected by D14)** | `DOLGIT` vs `DOLGIT ` collided after Trim/sanitization → fail-fast reported correctly. |
| `Create_report` probe | temp workbook | **Fail (fixed)** | `report_m/report_w/report_wclip` not found in flat `input_path`; fixed by recursive subfolder search. |
| `Create_report` probe | temp workbook | **Fail (fixed)** | `PivotTable.RefreshTable` failed (1004) in early path; mitigated by reducing forced refresh/update during filtering. |
| `Create_report` probe | temp workbook | **Fail (fixed, historical)** | Optional token `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` misdetected as required-missing due to Cyrillic literal handling; fixed in that snapshot. Superseded by rev.6 generalized `articleLevel4Name_clean` matched-subset/report-skip behavior. |
| `Create_report` latest probe | `.ai/tmp/treport_live_test3.xlsm`, output `.ai/tmp/live_output3` | **Inconclusive / running** | `settings!Z10=RUNNING`, `live_output3` empty, Excel PIDs `5520,19828` still present at report time. No confirmed success/failure. |
| `Refresh_data` end-to-end | — | **Not run** | Not executed to avoid uncontrolled writable changes; requires controlled validation. |
| Manual UI button click / visible Excel verification | — | **Not recorded** | Controls verified via COM object model only. |

**Source-data safety:** `report_m.xlsx`, `report_w.xlsx`, `report_wclip.xlsx` were backed up to `.ai/tmp/source_backup_20261115-013539/` before tests. `Create_report` implementation opens sources read-only; no intentional source mutation recorded.

## 6. Problems and blockers (prioritized)

### 6.1 Critical — blocks functional acceptance

| ID | Problem | Evidence/impact | Affected |
|---|---|---|---|
| P1 | **No confirmed successful end-to-end `Create_report` run.** Latest probe still `Z10=RUNNING` with empty output and lingering Excel processes. | `.ai/tmp/treport_live_test3.xlsm` (`Z10=RUNNING`), `.ai/tmp/live_output3` empty, PIDs `5520,19828`. Cannot claim correct report generation, naming/placeholders, sheet order, filter correctness, or copy fidelity. | B8–B11, B13, FR-RPT/FR-FLT/FR-NAME |
| P2 | **Canonical config has report name collision after Trim (D14).** `reports` row 3 = `DOLGIT`, row 4 = `DOLGIT ` (trailing space) → same sanitized file name `DOLGIT`. | Confirmed by fail-fast error in probe and by config evidence. `Create_report` must fail fast on collision per spec (correct behavior), so a successful run on canonical config is impossible without disambiguation. | FR-CFG-3, D14, B11, B13 |
| P3 | **`Refresh_data` runtime behavior unvalidated (OQ-T2).** PQ + dependent pivot synchronous completion/persistence before save/close not proven on target workbooks/build. | `modTreportRefresh` is scaffolded but not executed end-to-end in controlled test. Large files and environment-specific PQ connections increase risk. | B7, B13, FR-REF, D1 |

### 6.2 High — significant correctness/risk gaps

| ID | Problem | Evidence/impact | Affected |
|---|---|---|---|
| P4 | **`report_wclip` Year/grouped-Years runtime semantics unconfirmed (V-007/V-018/OQ-T1/R1/R3).** Config aligned to `Year`; workbook also has grouped `Years` field. | Static evidence only. Risk of wrong period selection/`max_year`/`max_weeknum` in `сравн конкур_<max_weeknum>w_<max_year>_прямая`-style outputs or incorrect filter application. | B5,B8,B11,B13 |
| P5 | **Presented-range/copy fidelity unvalidated (OQ-T3/R5).** `TableRange1` fallback and values/formats/widths/row-heights/static-only output not compared to actual presented pivot state. | No completed outputs to inspect; page-filter exclusion and grand-total/label boundaries are layout-dependent. | B9,B10,B13, FR-RPT-3/4/5 |
| P6 | **Pivot update stability/performance under real data not proven.** Prior 1004 `RefreshTable` fixed by reducing refresh; large `report_wclip` (~130 MB) may still cause very long runs or hangs (consistent with running probe). | Long-running/unfinished probe, no timeout telemetry in main UI path. | B5,B8,B9,B13,R8 |
| P7 | **VBA export/workbook parity and accidental tracked deletion risk.** `mko_treport/src/vbs/export/xlconcat_macro.bas` is deleted in working tree (HEAD still has it); exports are auxiliary while workbook is authoritative. | `git status` shows ` D mko_treport/src/vbs/export/xlconcat_macro.bas`. Risk of losing reference export or committing unintended deletion. | B1/B12 hygiene, OQ-T4 |

### 6.3 Medium — quality/operational gaps

| ID | Problem | Evidence/impact | Affected |
|---|---|---|---|
| P8 | **No automated regression suite aligned to VBA runtime.** Old Python tests removed (correctly stale); no replacement auxiliary verification harness committed. | Harder to detect regressions in DSL/naming/resolution logic. | B13, maintainability |
| P9 | **`IsPivotDataEmpty`/no-data detection is heuristic.** May misclassify presented-empty vs structurally present data body in some layouts. | Implemented but not runtime-verified across required report/filter combinations. | FR-FLT-7, B8,B13 |
| P10 | **Diagnostics/progress in main UI path limited to MsgBox/status-bar behavior (not fully audited).** Long operations may appear hung without detailed user-visible progress. | Relevant to observed running probe; main implementation shows success/fail MsgBox but progress granularity not validated. | FR-RUN-2/3, B12 |
| P11 | **Working-tree hygiene unrelated/pre-existing changes remain (`.ai/plans/001-treport-execution-plan.md`, `mko_treport/src/__init__.py`, `?? .ai/tasks/`).** | Risk of accidental staging/commit mixing concerns. | delivery hygiene |

## 7. Block status (rev.5 B0–B14)

| Block | Status | Notes |
|---|---|---|
| B0 Baseline/evidence | **Done (baseline)** / env recording partial | Artifacts, backups, Excel 16 available; exact target build/PQ prereqs for validation runs should be recorded during B13. |
| B1 Run controls & workbook prep | **Done (runtime, structurally verified)** | 2 controls/bindings in `treport.xlsm`, legacy run binding removed, modules imported/exported/synced. OQ-T4 parity caveat remains. |
| B2 Investigation/alignment | **Partially done** | Strong static evidence; runtime OQ-T1–T4 confirmations pending. Config aligned to `Year` for all files. |
| B3 Golden baseline | **N/A** | D18: no golden workbook. |
| B4 Config contract/validation | **Implemented, unproven end-to-end** | VBA module in workbook; probe-read passed. |
| B5 Filters parse & min/max/placeholders | **Implemented, unproven end-to-end** | Dependency-aware logic present; `report_wclip`/real-member behavior needs runtime proof. |
| B6 File resolution & access/session | **Implemented, partially fixed** | Subfolder-aware resolution added; D1 modes implemented; needs full run proof. |
| B7 Refresh workflow | **Implemented (scaffold), unvalidated** | PQ sync/persist not runtime-proven. |
| B8 Filter application/staging/placement | **Implemented, unproven** | D4,D7–D9/V-008 logic present; runtime correctness/leak/no-data behavior pending. |
| B9 Pivot snapshot extraction & paste | **Implemented, unproven** | Static copy path present; presented-range/fidelity pending (OQ-T3). |
| B10 Formatting/layout | **Implemented, unproven** | Freeze/zoom/gridlines/row heights; `group_rows` no-op by design. |
| B11 Assembly/naming/placeholders/idempotency | **Implemented, blocked by config collision / unproven** | D3/D11/D14 logic present; canonical `DOLGIT` collision prevents success. |
| B12 Diagnostics/run control/state safety | **Implemented (initial), partially proven** | Fail-fast/context/state restore paths present; long-run UX/progress and full runtime paths pending. |
| B13 Functional verification (target Excel) | **Not complete** | No successful end-to-end `Create_report`; `Refresh_data` untested. |
| B14 Independent review/readiness | **Not started** | Gated on B13 evidence. |

## 8. Ordered next steps (recommended execution path)

### Step 0 — Stabilize workspace (immediate, no behavior change)
1. **Resolve running probe state:** check/terminate stale Excel automation instances from test runs if safe; record final `Z10`/output for `.ai/tmp/treport_live_test3.xlsm` and archive logs/artifacts in `.ai/tmp`.
2. **Restore unintended working-tree deletion:** restore `mko_treport/src/vbs/export/xlconcat_macro.bas` from HEAD unless intentionally removed with documented rationale (keep unbound).
3. **Do not mix unrelated changes:** leave `.ai/plans/001-treport-execution-plan.md`, `mko_treport/src/__init__.py`, `.ai/tasks/` handling explicit.
**Exit:** clean, attributable working set for treport runtime fixes.

### Step 1 — Diagnose B13 runtime failure/hang (highest priority)
1. Re-run `Create_report` on a **temp copy** of `data/theiss/treport.xlsm` with silent/result logging and bounded timeout/checkpoints (per data_file/report stage).
2. Use **test-only config adjustments**: disambiguate report 4 name (`DOLGIT ` → e.g. `DOLGIT_30_50`) to satisfy D14, keep all other semantics intact; redirect `output_path` to `.ai/tmp/verification_output_*`.
3. Capture exact failing stage: file open, pivot locate, year/period resolve, apply data filters, apply report filters, snapshot copy, save. Record Excel build/PQ state if available.
4. If `report_wclip` dominates time, run a **controlled subset** (by temp config copy) to isolate per-file behavior, then full run.
**Exit:** reproducible pass/fail evidence with stage-level cause (or confirmed completion).

### Step 2 — Fix runtime defects in VBA (iterative, B4–B12)
Address defects found in Step 1, likely areas:
- Pivot refresh/update sequencing and `ManualUpdate` safety (avoid 1004/hangs) while preserving correctness.
- `Year` resolution/application for `report_wclip` with grouped `Years` present (validate base `Year` field path; escalate only genuine residual business ambiguity with concrete alternatives).
- Member matching/availability in resolved-year context for `Month Number`/`Week Number` ranges and `min/max`.
- No-data detection accuracy, generalized `articleLevel4Name_clean` matched-subset/report-skip behavior (V-008, rev.6), report-filter reset/leak prevention (D8).
- Presented-range boundaries (`TableRange1`/page-filter exclusion), static copy fidelity, and formatting application without pivot objects/PQ artifacts.
- Progress/status/timeout-safe behavior and contextual errors.
**Per fix:** edit `mko_treport/src/vbs/export/*.bas`, import into backed-up workbook copy, save, re-export components and refresh `mko_treport/src/vbs/vbaProject.bin`, sync tracked `mko_treport/src/xls/treport.xlsm` as needed, commit tracked code/export changes atomically.

### Step 3 — Resolve canonical config collision (P2)
- Decide and apply an explicit, spec-compliant name for row 4 to avoid Trim collision (e.g. distinct `reports.name`) in the working `treport.xlsm` config copy used for validation; document the change and rationale. Preserve `Trim` behavior (D14) — do not weaken collision detection.
**Exit:** config allows full 6-report generation without collision.

### Step 4 — B13 functional verification (`Create_report`)
Run on target Excel against actual source data (backed-up originals) and verify:
- Correct data files resolved/opened read-only and closed without save.
- Per-file filters applied once, per-report filters reapplied without leakage; generalized `articleLevel4Name_clean` multi-value rule: if any configured token matches apply matched subset, if none match warn and skip that report (continue other reports, no whole-run failure).
- Correct `min/max` and `max_*` (`max_year`, `max_weeknum`, `max_monthnum`) per D10/D11 and placeholder substitution in `report_sheet` only.
- One `.xlsx` per report, deterministic overwrite/idempotent rerun, 3 sheets per workbook in `data_files` order, correct names/sanitization/31-char behavior.
- Static values+formats (+widths/row heights as required), no pivot objects/PQ connections, formatting (freeze/zoom/gridlines), `group_rows` no-op.
- Fail-fast cases (missing/invalid config, bad filter, no-data, name collision) produce contextual errors.

### Step 5 — B13 functional verification (`Refresh_data`)
- Run on backed-up working copies with writable access: PQ/pivot refresh completes before save, data persisted, files closed correctly; record behavior/prereqs (OQ-T2).
- Confirm `Create_report` can run read-only on refreshed/current data as expected.

### Step 6 — B14 review, docs, and readiness
- Independent Validator pass against rev.5 (D1–D18), runtime boundary, controls, and recorded evidence.
- Update affected documentation only if behavior/contracts changed (per repo doc rules); keep changes minimal/relevant.
- Record final workbook hashes, module/export parity notes, Excel build/prereqs, and verification outcomes.

## 9. Suggested auxiliary verification aids (non-runtime, optional)

- A silent COM harness script (local `.ai/tmp`) that invokes `Refresh_data`/`Create_report` on a temp workbook copy, captures stage/result text, and enforces timeout — for repeatable debugging without UI `MsgBox` blocking.
- Lightweight OpenXML/value comparison checks on generated outputs vs source presented ranges for key cells/totals/names (auxiliary evidence, not a golden baseline).

## 10. Definition of done for implementation completion

- `treport.xlsm` provides working VBA-only `Refresh_data` and `Create_report` with required run controls and no legacy run binding.
- B13 passes functionally against actual source data in target Excel (including resolved P2 collision handling, `report_wclip` behavior, generalized `articleLevel4Name_clean` matched-subset/report-skip behavior, `Trim`, naming/idempotency, static copy/format, refresh persistence as applicable).
- Evidence recorded locally (environment, hashes, outputs, observations); tracked VBA exports/bin/workbook copy synced as appropriate; no external runtime dependency introduced.
- B14 independent validation accepted with no unresolved critical/high defects.
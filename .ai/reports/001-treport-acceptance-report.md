# Independent Acceptance Audit Report — treport VBA Solution

- **Task:** `001-treport-acceptance`
- **Audit date/time (environment UTC):** 2026-11-15T07:23:52Z (audit executed 2026-11-15 local session)
- **Auditor role:** Independent Validator / Acceptance Reviewer (read-only audit; no implementation files modified)
- **Report path:** `.ai/reports/001-treport-acceptance-report.md`
- **Authoritative baseline:** `.ai/problems/001_treport_spec.md` **rev.5** (2026-11-14) and `.ai/plans/001-treport-execution-plan.md` **rev.5** (2026-11-14)
- **Audit method:** Static inspection of repository/git state, exported VBA modules, workbook package/XML contents (read-only zip/XML parsing), parsed pivot-cache evidence (`.ai/tmp/pivots_parsed.json`), and local evidence artifacts in `.ai/tmp`. Subagents used: `auditor` (static VBA/spec compliance review) and `researcher` (environment/package/pivot evidence). **No live Excel runtime execution was performed in this audit session.**

## 1. Executive summary

Exported VBA modules in `mko_treport/src/vbs/export/` implement substantial rev.5-oriented logic: configuration reading/validation with `Norm/Trim` (`modTreportConfig.bas`), file resolution and access modes (`modTreportFiles.bas`), refresh orchestration with synchronous-wait attempts (`modTreportRefresh.bas`), filter DSL parse/classify and dependency-aware resolution/application (`modTreportFilters.bas`, `modTreportPivot.bas`), presented-range snapshot copy and formatting (`modTreportExtract.bas`), output/naming/placeholder handling (`modTreportOutput.bas`), application-state restore (`modTreportState.bas`), and required run entry points `Public Sub Refresh_data()` / `Public Sub Create_report()` (`modTreportRun.bas`).

Current workbook packages `data/theiss/treport.xlsm` and `mko_treport/src/xls/treport.xlsm` are identical (SHA256 `3F8E76888E94B5C204F7DE8AA47E46F242B3350CB21E8F6985417B5697A6667F`, 126672 bytes, 2026-11-15 01:55:57) and statically show two `run`-sheet controls (`Refresh_data`, `Create_report`) bound to `modTreportRun.Refresh_data` / `modTreportRun.Create_report` with no `RunXlconcat` binding in inspected UI XML/VML.

**The solution is not accepted as complete and working.** Principal blockers:

- **Critical functional defect (static):** `modTreportRun.ProcessDataFile` tracks sheet-name collisions in one dictionary for the entire data-file pass across all report workbooks, while sheet names are derived from the same data-file `report_sheet` template. This causes false fail-fast `2612` on the second report and blocks normal multi-report generation (FR-RPT-1/2, D3, D14, B11).
- **Delivered config not runnable to success as-is (spec-correct fail-fast):** `reports` rows 3 and 4 are `DOLGIT` and `DOLGIT ` (trailing space). After Trim/sanitization these collide and `Create_report` would fail-fast with `2611` before processing (D14). This is expected behavior but prevents successful end-to-end output from the current workbook configuration.
- **Workbook VBA parity not demonstrated (OQ-T4/R7):** embedded `xl/vbaProject.bin` inside both current workbook copies contains legacy strings (`RunXlconcat`, `xlconcat_macro`) and unexpected module-name strings (`modTreportExtractL`, `modTreportOutput2`, `modTreportStatea`, `modTreportStateJ`) alongside expected modules. Exports do not contain `xlconcat_macro.bas` (deleted in git) and do not show these unexpected module names. Standalone `mko_treport/src/vbs/vbaProject.bin` (SHA256 `845884F65E3250B130704B76423A861B738E5FD790DF73C1B421787237A4AFD2`, 254976 bytes) also contains legacy markers. Compile/runtime equivalence to exports is unverified.
- **Runtime acceptance not performed (D18/B13, OQ-T1–T4):** no Excel execution of `Refresh_data`/`Create_report`, no proof of synchronous PQ refresh/persisted save, no proof of pivot filtering/copy/formatting/output correctness, and no required end-to-end scenarios executed. No generated report outputs were observed in `data/theiss/output` or `.ai/tmp/*output*`.
- **Spec-conformance concern:** file resolution searches recursively into subfolders and does not fail-fast when matches exist across different allowed extensions (best precedence is selected), which conflicts with a strict reading of FR-FILE-1/D2 ("within `input_path`", "fail fast if >1 match").
- **Non-evidence test artifacts:** `mko_treport/tests/treport_*` contains only Python `__pycache__/*.pyc` bytecode (no `.py` sources) and does not provide evidence of VBA compliance.

**Conclusion:** the repository contains partial implementation/documentation and locally modified artifacts, but it does **not** demonstrate a delivered, working rev.5 VBA report-generation solution. Deferred `group_rows` (D13) and absence of a golden workbook (D18) are correctly excluded as failure causes.

## 2. Final acceptance decision

**FAIL**

**Decision basis**
- Core report-generation behavior is statically defective/blocked for the required multi-report case (D-01).
- Actual workbook controls/procedures and end-to-end VBA behavior have not been verified in Excel runtime (task rule: do not issue PASS if actual VBA behavior has not been verified or core report-generation functionality is missing/incorrectly unproven).
- Workbook VBA project parity is not demonstrated and contains legacy/unexpected artifacts (OQ-T4/R7).
- Required runtime functional acceptance (D18/B13, OQ-T1–T4) is incomplete.

**Not treated as failure (PO-resolved):** absence of a golden expected-output workbook (D18) and `group_rows` deferred/no-op (D13/R4).

## 3. Specification and plan revision identification

| Artifact | Identification | Evidence |
|---|---|---|
| Specification | `001_treport`, **rev.5** (2026-11-14), "Ready for implementation planning and VBA implementation; technical validations/evidence tasks tracked in §13" | `.ai/problems/001_treport_spec.md` (rev.5 header, §§1–15, D1–D18, FR-CFG/FILE/REF/FLT/RPT/NAME/RUN, §13 resolved/open items) |
| Execution plan | `001-treport-execution-plan`, **rev.5 baseline** aligned to spec rev.5 | `.ai/plans/001-treport-execution-plan.md` (scope/guardrails, B0–B14, readiness, gates) |
| Key approved decisions audited | VBA-only final runtime (auxiliary dev tools allowed, non-runtime), target Windows 10+/Excel 2021+, no golden workbook, `group_rows` deferred/no-op, exactly two `run` controls/procedures `Refresh_data`/`Create_report`, `Trim` policy, optional `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ`, `Year` used for all `data_files` filters (evidence-driven V-007/V-018), local/gitignored workbook handling | Spec §§2,5,6,13.1,13.3,15; Plan "Resolved PO decisions and current evidence (rev. 5)" |

## 4. Repository and workbook inventory

### 4.1 Git state (audit snapshot)
- **Status:** `M .ai/plans/001-treport-execution-plan.md`, `D mko_treport/src/__init__.py`, `D mko_treport/src/vbs/export/xlconcat_macro.bas`, `?? .ai/tasks/`
- **Recent commits:** `57ae40e fix(treport): improve file resolution, optional filter handling, and pivot refresh stability`; `cf2c4e3 chore(treport): remove stale auxiliary Python tests for removed runtime modules`; `fe2c05b feat(treport): implement VBA config, refresh, and report generation pipeline`; `dba8639 feat(treport): complete B1 run controls in workbook`; `544108d feat(treport): add rev5 run entry-point stubs (B1)`.
- **Observation:** commit messages/status claims are not treated as acceptance evidence.

### 4.2 Exported VBA modules and binaries (repository)
| Path | Size | Modified | Notes |
|---|---:|---|---|
| `mko_treport/src/vbs/export/modTreportConfig.bas` | 9056 | 2026-11-15 01:33:20 | Config read/validation, `Norm/Trim`, folder helpers |
| `mko_treport/src/vbs/export/modTreportRun.bas` | 8619 | 2026-11-15 01:33:20 | `Refresh_data`, `Create_report`, orchestration/error handling |
| `mko_treport/src/vbs/export/modTreportFiles.bas` | 5410 | 2026-11-15 01:46:01 | Resolution, open/save/close, access modes |
| `mko_treport/src/vbs/export/modTreportRefresh.bas` | 3004 | 2026-11-15 01:33:20 | Refresh workflow/waits |
| `mko_treport/src/vbs/export/modTreportFilters.bas` | 15808 | 2026-11-15 01:55:57 | DSL parse/classify, min/max/year resolution, apply filters |
| `mko_treport/src/vbs/export/modTreportPivot.bas` | 8124 | 2026-11-15 01:50:42 | Pivot access, placement, selection, no-data check |
| `mko_treport/src/vbs/export/modTreportExtract.bas` | 2833 | 2026-11-15 01:50:42 | Presented range copy, static paste, formatting |
| `mko_treport/src/vbs/export/modTreportOutput.bas` | 5585 | 2026-11-15 01:33:20 | Sanitization, placeholders, sheet/workbook creation/save |
| `mko_treport/src/vbs/export/modTreportState.bas` | 1405 | 2026-11-15 01:33:20 | Application state capture/restore |
| `mko_treport/src/vbs/export/Sheet1.cls` | 219 | 2026-11-15 01:33:20 | Minimal class export |
| `mko_treport/src/vbs/export/Sheet2.cls` | 219 | 2026-11-15 01:33:20 | Minimal class export |
| `mko_treport/src/vbs/export/ThisWorkbook.cls` | 225 | 2026-11-15 01:33:20 | Minimal class export |
| `mko_treport/src/vbs/vbaProject.bin` | 254976 | 1980-01-01 00:00:00 | SHA256 `845884F65E3250B130704B76423A861B738E5FD790DF73C1B421787237A4AFD2`; legacy/unexpected module strings observed |
| `mko_treport/src/vbs/export/xlconcat_macro.bas` | deleted | git `D` | Absent from exports; legacy must not be UI-accessible |

### 4.3 Workbooks and source data (local)
| Path | Size | Modified | SHA256 / note |
|---|---:|---|---|
| `data/theiss/treport.xlsm` | 126672 | 2026-11-15 01:55:57 | `3F8E7688...67A6667F`; working control workbook (config paths point to `data/theiss`) |
| `mko_treport/src/xls/treport.xlsm` | 126672 | 2026-11-15 01:55:57 | Identical to above |
| `data/theiss/treport.xlsm.bak_20261114_b1_r4` | 78443 | 2026-11-14 16:57:30 | Local backup |
| `data/theiss/treport.xlsm.bak_b1_pre_r4` | 78443 | 2026-11-14 16:57:30 | Local backup |
| `data/theiss/monthly/report_m.xlsx` | 3577065 | 2026-11-14 16:58:51 | `718F9793...3CC889A5` |
| `data/theiss/weekly/report_w.xlsx` | 5767351 | 2026-11-06 12:09:26 | `F4BB589C...83EA5754` |
| `data/theiss/base/report_wclip.xlsx` | 130222326 | 2026-11-15 00:26:26 | `A3CBCE00...E3892CC1` |
| `.ai/tmp/b2_extract/treport_xlsm/xl/vbaProject.bin` | 154624 | zip epoch | `3A755274...9A83986A` (historical extracted snapshot) |
| `.ai/tmp/b1_backup_20261115-005533/treport_data_theiss_original.xlsm` | 78439 | 2026-11-15 00:26:52 | Pre-change backup (evidence artifact) |

### 4.4 Test environment (available evidence)
| Item | Value | Verification |
|---|---|---|
| OS/platform | Windows 11, PowerShell session (environment) | Static |
| Excel executable | `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE` present | Static path check |
| Excel/Office registration (registry) | ClickToRun `ProductReleaseIds=ProPlus2021Volume`, `AudienceData=Production::LTSC2021`, `Platform=x64`, `VersionToReport=16.0.14332.20615`, `ClientVersionToReport=16.0.14332.20604`, `UpdateChannel=http://officecdn.microsoft.com/pr/5030841d-c919-4594-8d2d-84ae4f96e58e` | Auxiliary registry query (static) |
| COM ProgID | `Excel.Application` resolves to `Microsoft.Office.Interop.Excel.ApplicationClass` | Static type resolution only (no instance created) |
| Live Excel/VBA execution | **Not performed** | N/A |
| Macro/PQ runtime prerequisites | **Not exercised/recorded in runtime** (macro trust, VBA project access, PQ/Mashup availability, connection-path viability) | Not verified |
| Inspected workbook config | `general.input_path=C:\py_exp\mko_treport\data\theiss\`, `output_path=C:\py_exp\mko_treport\data\theiss\output`, `delimeter=&`; 6 report rows; 3 data_files (`report_m`, `report_w`, `report_wclip`, sheet `clean_data_full`, filters `Month Number=(1-max)|Year=max`, `Week Number=(1-max)|Year=max`, `Week Number=(1-max)|Year=max`) | Static package/shared strings/tables |

### 4.5 Auxiliary evidence and test artifacts (non-runtime)
- `.ai/tmp/pivots_parsed.json`, `.ai/tmp/b2_extract/**`, `.ai/tmp/inspect_pivots*.py`, `.ai/tmp/b2_pivots.ps1` — package/pivot-cache inspection evidence.
- `.ai/tmp/b1_evidence_20261115-005533/B1_evidence.md` — prior local B1 note (reported UI/binding/parity not fully programmatically verified in that context).
- `.ai/tmp/live_output*`, `.ai/tmp/output`, `.ai/tmp/test_output`, `data/theiss/output` — directories present; **no generated report files observed**.
- `mko_treport/tests/treport_config|files|filters|refresh/**/__pycache__/*.pyc` only — stale Python bytecode artifacts, no `.py` sources; **not VBA compliance evidence**.

## 5. B0–B14 completion matrix

| Block | Objective (rev.5) | Status | Evidence | Assessment |
|---|---|---|---|---|
| B0 | Environment, artifact baseline & evidence capture | **PARTIAL** | Inventory, hashes/sizes, Excel install/registry data collected; local/gitignored workbooks and backups observed | Static baseline available; runtime prerequisites not exercised |
| B1 | `run` controls and runtime-boundary alignment (FR-RUN-1) | **PARTIAL / NOT VERIFIED** | Current workbook package UI shows 2 controls `Refresh_data`/`Create_report` bound to `modTreportRun.*`, no `RunXlconcat` in UI XML/VML; exports define both procedures; embedded VBA binaries contain legacy/unexpected module strings | UI appears compliant statically; compile, click execution, and VBA project parity not proven |
| B2 | Technical investigation & config alignment (OQ-T1–T4, V-007/V-018) | **PARTIAL** | Config uses `Year=max` for all files; cache evidence: `Year` present in all sources, grouped `Years` also present in `report_wclip`; PQ/connection structures observed | Evidence-driven alignment partially supported; runtime confirmation pending |
| B3 | Golden baseline | **N/A** | D18 (no golden workbook required) | Correctly excluded |
| B4 | VBA config contract & validation | **PARTIAL** | `modTreportConfig.ReadConfig/Norm/Fail`; required ListObjects/columns/values, Trim, `group_rows` stored but no-op in flow | Implemented in exports; runtime correctness unverified |
| B5 | Filters parse & dependency-aware resolution | **PARTIAL** | `ParseAndClassify/ClassifyExpr/ResolveMinMaxAndYear/ResolveClauseMax`; numeric/inclusive handling, year-first context | Logic present; runtime member semantics unverified |
| B6 | File resolution & access/session policy | **PARTIAL / CONCERN** | Precedence `.xlsx/.xlsm/.xlsb`, read-only/writable open, save/close, output/input checks; recursive subfolder search and cross-extension ambiguity handling deviate from strict D2/FR-FILE-1 | Implemented but spec-conformance concern (D-04) |
| B7 | Refresh workflow | **PARTIAL / NOT VERIFIED** | Disable background query attempts, `RefreshAll`, `CalculateUntilAsyncQueriesDone`, polling timeout, pivot cache/table refresh, save/close | Best-effort logic; OQ-T2 not runtime-validated |
| B8 | Filter application, staging & placement | **PARTIAL** | Data-file filters once, per-report category/TA reapplied, in-place/add-to-Filters placement, optional token skip with >=1 valid required | Static logic largely aligned; runtime unverified |
| B9 | Pivot snapshot extraction & static paste | **PARTIAL / NOT VERIFIED** | `GetPresentedRange` (`TableRange1` preferred), paste values+number formats/formats/column widths/row heights | Implemented; presented-range/fidelity unverified (OQ-T3) |
| B10 | Formatting & layout | **PARTIAL** | Freeze/zoom/gridlines, blanks/no-op, `group_rows` ignored | Implemented; runtime unverified |
| B11 | Assembly, naming, placeholders & idempotency | **FAIL** | One workbook per report, placeholder substitution, Trim/sanitize/31-char/filename collision checks; **sheet collision tracked across all report workbooks per data-file pass** | Critical blocker for multi-report generation (D-01) |
| B12 | Diagnostics, run control & state safety | **PARTIAL** | Contextual `Fail` taxonomy/messages, state restore, explicit outcomes, legacy removed from exports/UI | Runtime/parity incomplete (D-02) |
| B13 | Verification & acceptance (functional, target Excel) | **NOT VERIFIED / FAIL** | No runtime scenarios executed; no outputs observed; static blockers identified | Acceptance criteria not demonstrated |
| B14 | Independent review & readiness | **FAIL** | This report | Not ready for delivery |

## 6. Full requirement traceability matrix

| Requirement ID | Expected behavior | Evidence | Verification method | Result | Finding | Severity | Recommendation |
|---|---|---|---|---|---|---|---|
| D16, §2.2 | VBA-only final runtime; no Python/PowerShell/BAT external runtime dependency for user | Exports use VBA/Excel OM and in-process COM (`Scripting.FileSystemObject`, `Scripting.Dictionary`); no shell/python/ps/bat process calls found; `.ai/tmp` scripts not referenced by VBA | Static inspection | **PASS (boundary)** | No external process runtime dependency found in exported VBA | Low | Validate on clean target Excel host |
| D18 | No golden workbook; validate against functional reqs and actual source data | No golden baseline required/used; runtime functional validation not executed | Process/static review | **NOT VERIFIED** | Required functional acceptance not performed | High | Execute full B13 against actual source data on working copies |
| D17 | Dropped params not required | No `new_file_delimeter`/`keep_col_hidden` usage in exports | Static inspection | **PASS** | Not required/processed | Low | None |
| FR-CFG-1 | Read `general,reports,data_files,placeholders` from `settings` | `GetSettingsWs`, `GetLO` for required ListObjects | Static inspection | **PASS** | Correct targets | Low | Runtime check |
| FR-CFG-2 | `input_path`,`output_path`,`delimeter` mandatory/non-empty | `ReadConfig` fails 2104–2106 | Static inspection | **PASS** | Enforced | Low | Negative tests |
| FR-CFG-3, D14 | Reports schema, one output per row, Trim | `CellText/Norm`, required checks 2110–2112, per-report workbook creation | Static inspection | **PASS** | Implemented | Low | Runtime check |
| FR-CFG-4, D3, D13 | Ordered `data_files`, required fields, `group_rows` no-op | Ordered load/validation 2113–2117; `GroupRows` not used in processing | Static inspection | **PASS** | Implemented | Low | Runtime check |
| FR-CFG-5, D11 | Placeholder definitions/derivation support | Placeholder table read; `ResolveMinMaxAndYear`; `SubstitutePlaceholders` | Static inspection | **PASS** | Implemented | Medium | Runtime check |
| FR-CFG-6, D12 | Defaults/no-op and fail-fast validation | Layout defaults and widespread contextual `Fail` usage | Static inspection | **PASS** | Implemented pattern | Medium | Edge-case runtime tests |
| V-020, D14 | Trim before match/substitution/naming/sanitization/collision | `Norm` used across config, filters, file/output helpers | Static inspection | **PASS** | Broadly applied | Low | Runtime edge cases |
| FR-FILE-1, D2 | Resolve in `input_path` with precedence `.xlsx,.xlsm,.xlsb` | Ordered ext precedence implemented | Static inspection | **CONCERN** | Recursive subfolder search extends scope beyond strict "within `input_path`" | Medium | Align scope to spec or obtain PO clarification (D-04) |
| FR-FILE-2, D2 | Fail-fast if no match or >1 match | 2201 when none; 2202 when >1 at same best precedence; cross-extension coexistence selects best precedence without fail-fast | Static inspection | **FAIL** | Strict ">1 match" ambiguity rule not fully enforced | Medium | Fail on any >1 matching candidate as specified (D-04) |
| FR-FILE-3, D1 | `Create_report` opens sources read-only, never saves | `OpenWorkbookReadOnly`, `CloseNoSave` in normal/error paths | Static inspection | **PASS** | Implemented | High | Runtime verify no source save/timestamp change |
| FR-FILE-4, D1 | `Refresh_data` writable, refresh, save, close | `OpenWorkbookWritable` rejects read-only/locked, refresh, `SaveAndClose` | Static inspection | **PASS** | Implemented | High | Runtime persistence/lock tests |
| FR-FILE-5 | Ensure/create output path | `EnsureOutputFolder` in both operations | Static inspection | **PASS** | Implemented | Medium | Permission tests |
| FR-REF-1 | Refresh PQ and dependent pivot data per file, save/close | Ordered loop, refresh call, save/close | Static inspection | **PASS** | Implemented | High | Runtime check |
| FR-REF-2, OQ-T2 | Synchronous completion before save/close; no stale persisted data | Disable background query attempts, `RefreshAll`, `CalculateUntilAsyncQueriesDone`, polling timeout, cache/table refresh | Static inspection | **NOT VERIFIED** | Best-effort; environment/PQ dependent | High | Validate in target Excel and record evidence |
| FR-REF-3 | Contextual fail-fast on refresh failure | 2501/2502 and cleanup handlers | Static inspection | **CONCERN** | Implemented; runtime behavior unverified | Medium | Failure-injection tests |
| FR-FLT-1, D5,D6 | Parse `filters` DSL and disambiguate | Clause parsing, trim, validation, list/range/scalar classification, duplicate detection | Static inspection | **PASS** | Implemented | Medium | Edge-case tests |
| FR-FLT-2, D10, V-007/V-018 | Dependency-aware min/max in resolved year context; numeric/inclusive; no guessed mapping | Resolves/applies Year first then month/week; inclusive ranges; config uses `Year` for all files | Static + cache evidence | **CONCERN** | Logic aligned; runtime behavior for `report_wclip` (base `Year` and grouped `Years` coexist) unverified | High | Runtime-resolve/log per file |
| FR-FLT-3, D11 | Derive `max_*`; fail-fast if referenced but underivable | `ResolvedContext` and placeholder fail-fast 2605–2607 | Static inspection | **PASS** | Implemented | Medium | Runtime check |
| FR-FLT-4, D8 | Apply data-file filters once and keep applied | `ApplyDataFileFilters` before report loop | Static inspection | **PASS** | Implemented (Year also applied during resolution) | Medium | Runtime no-leak check |
| FR-FLT-5, D8,D9,V-008 | Per-report filters reapplied without leak; exact-after-trim; optional `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` skip-if-absent with >=1 valid required | `ApplyReportFilters`, `ApplyFieldSelection(allowOptionalSkip:=True)` for category, `IsOptionalSkipToken`; caches show token absent and `ЗУБНАЯ ПАСТА` present | Static + cache evidence | **CONCERN** | Static logic correct; runtime matching/visibility unverified | High | Runtime scenario |
| FR-FLT-6, D7 | Filter in place or add to Filters, no move | `EnsureFilterPlacement` preserves existing orientation except hidden/unset | Static inspection | **PASS** | Implemented | Medium | Runtime layout check |
| FR-FLT-7, D12 | Missing sheet/pivot/field/unresolvable/no-data fail-fast (except optional skip) | 2401–2403, 2406–2410, parse/resolve errors | Static inspection | **CONCERN** | Coverage present; runtime unverified | Medium | Negative scenarios |
| FR-RPT-1, D3 | One output workbook per report, overwrite on success | Per-report workbook creation, sanitized filename, SaveAs to output folder | Static inspection | **CONCERN** | Blocked by config/implementation defects in current state | Critical | Fix blockers, verify overwrite/idempotency |
| FR-RPT-2, D3,D11 | Sheets per data_files in order, substituted names, deterministic replace | Per-data-file loop, substitution/sanitization, `AddReportSheet` replaces existing in target workbook | Static inspection | **FAIL** | **Cross-workbook sheet collision tracking causes false 2612** (D-01) | Critical | Scope collision detection to each output workbook |
| FR-RPT-3, D15, OQ-T3 | Presented range incl. labels/totals, exclude page-filter/pivot/PQ | `GetPresentedRange` prefers `TableRange1`, fallback `TableRange2` | Static inspection | **NOT VERIFIED** | Fallback/layout-dependent; no runtime fidelity proof | High | Validate in Excel for all sources |
| FR-RPT-4, D15 | Static values+formats (+ column widths/row heights) at configured position | PasteSpecial values+number formats, formats, column widths; row-height copy | Static inspection | **PASS** | Implemented | High | Runtime verify |
| FR-RPT-5, D13 | Freeze/zoom/gridlines/sizing/state; blanks no-op; `group_rows` ignored | `ApplySheetFormatting`, no grouping logic | Static inspection | **PASS** | Implemented | Medium | Runtime verify |
| FR-RPT-6, D13 | `group_rows` not implemented/non-blocking | No-op | Static inspection | **PASS** | Correctly deferred | Low | None |
| FR-RPT-7, D1 | Save outputs, release resources, close sources per rules | `SaveOutputWorkbook`, `CloseNoSave`, `SaveAndClose` | Static inspection | **CONCERN** | Implemented; success not demonstrated | Medium | Runtime verify |
| FR-NAME-1, D11 | Placeholder substitution in `report_sheet` only; no zero-padding; fail-fast if underivable | Substitution on `ReportSheet` with numeric string output and 2605–2607 | Static inspection | **PASS** | Implemented | Medium | Runtime verify |
| FR-NAME-2, D14 | Trim/sanitize/31-char/collision fail-fast for names | `SanitizeSheetName/FileName`, report-name collision pre-check 2611, 2602/2612 checks; current config has `DOLGIT`/`DOLGIT ` collision | Static + config evidence | **CONCERN** | Rules implemented; report-name fail-fast would correctly trigger on current config; sheet collision scope defective (D-01) | High | Fix D-01; normalize config for runnable acceptance run |
| FR-RUN-1 | Exactly two `run` controls `Refresh_data`/`Create_report` bound to same-named VBA procedures; legacy removed | UI package shows two controls/bindings and no legacy UI binding; exports define procedures; embedded VBA binary contains legacy/unexpected strings | Static package + code + string scan | **CONCERN** | UI appears correct; workbook VBA parity/compile/runtime not verified | Critical | Rebuild/sync VBA project, compile, runtime-test buttons (D-02) |
| FR-RUN-2 | Restore application state, minimal UI disruption | `CaptureState/ApplyFastState/RestoreState` in success/error paths | Static inspection | **PASS** | Implemented pattern | Medium | Runtime observation |
| FR-RUN-3, D12 | Explicit outcomes and contextual fail-fast errors | MsgBox outcomes, contextual `Fail`, error handlers | Static inspection | **PASS** | Implemented | Medium | Runtime validation |
| A6, §13.3 | Local Excel handling, backups/working copies, no large commits required | Local workbooks/data and backups observed; audit artifacts limited to report/temp evidence paths | Static inspection | **PASS** | Consistent with policy | Low | Continue preserving originals |

## 7. OQ-T1–T4 verification results

| ID | Required validation | Evidence | Method | Result | Notes |
|---|---|---|---|---|---|
| OQ-T1 | Actual PivotTables/fields/captions/orientations/members/ranges/totals/page-filter areas; Year/Years/Month/Week/category/TA behavior; `Year`-based filtering correctness for all files incl. `report_wclip` | Parsed pivot-cache evidence (`.ai/tmp/pivots_parsed.json`, auxiliary). `report_m`: Year [2026], Month 3–10 (max 10), Week {9,14,18,23,31,27,36,40}. `report_w`: Year [2026], Month 3–10, Week 14–43 (max 43). `report_wclip`: base Year [2024,2026,2025] (max 2026), grouped `Years` field present (groupBy years, items `<01/01/2024,2024,2025,2026,>02/11/2026`), Month 1–12, Week 1–53. All sources include `ЗУБНАЯ ПАСТА` and target audiences `W 30-50`, `W 50+`; none include distinct `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` as cache item. Workbook config uses `Year=max` for all 3 data_files. Full pivot table definitions/presented ranges/page-filter layout/totals not comprehensively extracted in this audit. | Auxiliary package/cache inspection (static) | **PARTIAL** | Field/member presence and likely maxima derivable statically (m: Year 2026/Month 10; w: Year 2026/Week 43; base: Year max 2026, Week cache up to 53 but dependent-on-year runtime resolution unverified). Runtime pivot object selection, presented layout, and correct `Year`-based behavior in `report_wclip` remain unverified. |
| OQ-T2 | Synchronous PQ refresh + pivot cache refresh completeness and persisted-save behavior on target Excel | `modTreportRefresh.RefreshWorkbookData` disables background query flags where possible, calls `RefreshAll`, `CalculateUntilAsyncQueriesDone`, polls refreshing flags (300s timeout), refreshes pivot caches/tables. Source packages indicate external/Mashup connections with environment-specific `odcFile` paths. No execution performed. | Static inspection | **NOT VERIFIED** | Critical runtime validation outstanding; connection portability may affect refresh. |
| OQ-T3 | Presented copy-range boundaries and formatting/value fidelity against actual source pivot state | `GetPresentedRange` prefers `TableRange1` with fallback `TableRange2`; `CopyPivotSnapshot` pastes values+number formats, formats, column widths and copies row heights. No runtime comparison performed. | Static inspection | **NOT VERIFIED** | Fallback/layout-dependent; page-filter exclusion and fidelity must be proven in Excel. |
| OQ-T4 | VBA runtime-vs-export parity; legacy `RunXlconcat` removed/replaced | Exports contain rev.5 modules and no `xlconcat_macro.bas`; workbook UI shows `Refresh_data`/`Create_report`; embedded `xl/vbaProject.bin` in both workbook copies contains `RunXlconcat`/`xlconcat_macro` and unexpected module-name strings; standalone `mko_treport/src/vbs/vbaProject.bin` differs and contains legacy markers. No VBA IDE compile/runtime inspection. | Static comparison/string scan | **FAIL (not demonstrated)** | Parity and legacy cleanup not proven; embedded VBA project not reconciled with exports. |

## 8. FR-RUN-1 control and VBA entry-point verification

| Check | Expected | Evidence | Method | Result | Finding |
|---|---|---|---|---|---|
| `run` sheet controls | Exactly two required controls: `Refresh_data`, `Create_report` | `data/theiss/treport.xlsm` package: `xl/worksheets/sheet1.xml` and `xl/drawings/vmlDrawing1.vml` define two controls/macros `Refresh_data` and `Create_report`; two `xl/ctrlProps/*` entries; no `RunXlconcat`/`Run Macro` in inspected UI XML/VML | Static package inspection | **PARTIAL PASS** | UI artifacts appear compliant in current package |
| Control bindings | Bound to corresponding VBA procedures | VML FmlaMacro/Object bindings: `modTreportRun.Refresh_data`, `modTreportRun.Create_report` | Static package inspection | **PARTIAL PASS** | Static binding present; click/runtime invocation not verified |
| Legacy control/workflow | Old `RunXlconcat` not accessible via `run` | No legacy UI binding in inspected current package; `xlconcat_macro.bas` deleted from exports (git). Historical extracted snapshot (`.ai/tmp/b2_extract/treport_xlsm`) shows prior legacy binding. | Static package + git | **PARTIAL PASS (UI) / CONCERN (project)** | UI appears clean; embedded VBA binary still contains legacy identifiers |
| VBA procedures exist | `Public Sub Refresh_data()` and `Public Sub Create_report()` implemented | `mko_treport/src/vbs/export/modTreportRun.bas` defines both procedures with workflow logic, contextual errors, and state restore | Static code inspection | **PASS (exports)** | Workbook-embedded parity unverified |
| Compile and runtime behavior | Both compile and work correctly when invoked | Not executed; no VBA IDE evidence | Not performed | **NOT VERIFIED** | Required runtime acceptance evidence missing |

## 9. End-to-end test results

No end-to-end scenarios were executed in Excel runtime during this audit. Results below are static predictions/evidence-based assessments.

| Scenario | Expected result | Method | Result | Evidence / finding |
|---|---|---|---|---|
| `Refresh_data` | Source workbooks refresh, save, close successfully | Not executed | **NOT VERIFIED** | Writable-open/refresh/wait/save/close logic present (`modTreportRun.Refresh_data`, `modTreportRefresh`); environment-specific PQ connections and OQ-T2 unverified |
| `Create_report` | Correct output workbooks generated | Not executed | **NOT VERIFIED / PREDICTED FAIL** | Current config would fail-fast 2611 (`DOLGIT` vs `DOLGIT ` collision after Trim). If config corrected, implementation likely fails 2612 on second report due to sheet-collision scoping defect (D-01). No outputs observed in `data/theiss/output` or `.ai/tmp/*output*`. |
| Repeated report generation | Deterministic results and correct overwrite behavior | Not executed | **NOT VERIFIED** | Overwrite/save logic present, but success path blocked/unproven |
| Missing optional category `ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ` | Skip if absent, no failure when >=1 valid remains | Static + cache evidence | **CONCERN (static PASS)** | Optional-skip logic present; token absent in all source caches while `ЗУБНАЯ ПАСТА` present; runtime unverified |
| At least one valid category remains | Correct selection/fail-fast if none | Static | **CONCERN** | 2407 logic present; runtime unverified |
| `Trim` | Leading/trailing spaces handled correctly | Static + config evidence | **CONCERN** | Trim applied broadly; current config trailing-space report name triggers correct filename collision fail-fast (2611) |
| Different report rows | No filter leakage | Static | **CONCERN** | Category/TA reapplied per report row; runtime unverified |
| Different period fields | Correct per-file period resolution | Static + cache evidence | **CONCERN** | Expected maxima from caches: report_m Year 2026/Month 10; report_w Year 2026/Week 43; report_wclip Year max 2026 and Week cache 1–53 (dependent-on-year runtime resolution unverified). Expected report_w sheet: `сравн конкур_43w_2026_прямая`; report_m: `помесячное сравнение`; report_wclip: `хроно по неделям` (subject to runtime resolution/behavior). |
| Missing required file | Clear failure | Static | **CONCERN** | 2201 implemented; not executed |
| Invalid required field/filter | Clear failure | Static | **CONCERN** | Validation/fail-fast paths implemented; not executed |
| No matching data | Clear failure | Static | **CONCERN** | 2410 implemented; not executed |
| Naming collision | Correct fail-fast behavior | Static + config evidence | **PASS (predicted, report filenames)** | 2611 would correctly trigger on current `DOLGIT`/`DOLGIT ` config; sheet-collision handling has separate scope defect (D-01) |
| Source protection in `Create_report` | Sources not modified/saved | Static | **CONCERN** | Read-only open + close-no-save implemented; runtime unverified |

## 10. Defects and severity classifications

| ID | Severity | Requirement/Block | Description | Evidence | Recommendation |
|---|---|---|---|---|---|
| D-01 | **Critical** | FR-RPT-1/2, FR-NAME-2, D3, D14, B11 | `ProcessDataFile` uses one `usedNames` dictionary for the entire data-file pass across all report workbooks, but sheet names are computed from the same data-file `report_sheet` template/context. Second report triggers false 2612 "Sheet name collision in output workbook...", blocking normal multi-report generation. | `mko_treport/src/vbs/export/modTreportRun.bas` (`ProcessDataFile`: `usedNames` created once per data file; checked/added inside report loop while writing to different `outBooks(r)`) | Track collisions per output workbook namespace (or rely on per-workbook replace with within-workbook uniqueness checks); add regression for 6 reports × 3 data-file sheets. |
| D-02 | **Critical** | FR-RUN-1, OQ-T4, R7, B1/B12 | Embedded workbook VBA projects are not demonstrably synchronized with clean exports and contain legacy/unexpected artifacts: `RunXlconcat`/`xlconcat_macro` and unexpected module-name strings (`modTreportExtractL`, `modTreportOutput2`, `modTreportStatea`, `modTreportStateJ`) observed in `xl/vbaProject.bin` strings for current workbook copies. Standalone `mko_treport/src/vbs/vbaProject.bin` differs and also contains legacy markers. Exports lack `xlconcat_macro.bas` and these unexpected module names. | Workbook package string scans, hashes/sizes, export inventory, git deletion of `xlconcat_macro.bas` | Rebuild VBA project in working workbook from reviewed source: remove legacy/stray modules, keep only intended modules/procedures, compile in VBA IDE, verify button execution, re-export synchronized `.bas/.cls` and `vbaProject.bin`. |
| D-03 | **High** | D18, B13, OQ-T1–T4, FR-RUN-1 runtime | No runtime/functional acceptance evidence. Controls/compile, PQ synchronous persistence, per-file filtering (incl. `report_wclip` Year vs grouped Years context), copy/format fidelity, output correctness, and all B13 scenarios are unverified. No generated outputs observed. | Audit record; environment inspected but Excel not executed | Perform controlled runtime acceptance on working copies in target Excel and record evidence (build, macro/PQ prerequisites, outcomes, comparisons). |
| D-04 | **High/Medium** | FR-FILE-1, FR-FILE-2, D2, B6 | File resolution searches recursively into subfolders and does not fail-fast when matches exist across different allowed extensions (best precedence selected). | `mko_treport/src/vbs/export/modTreportFiles.bas` (`CollectMatchingFilesInFolder` recursion; ambiguity logic) | Align to spec (scope within `input_path`, fail-fast on any >1 match) or obtain explicit PO clarification and update baseline docs. |
| D-05 | **Medium** | D14, runnable config state | Current config has deterministic report filename collision after Trim (`DOLGIT` / `DOLGIT `), so correct fail-fast 2611 prevents successful-run evidence on delivered config. | Workbook `reports` table/shared strings; `Create_report` pre-check | Normalize config for acceptance run or validate as intentional negative case. |
| D-06 | **Medium** | Evidence quality / auxiliary tests | `mko_treport/tests/treport_*` contains only Python `__pycache__/*.pyc` (no sources), not VBA compliance evidence and potentially misleading. | `mko_treport/tests/**` inventory | Exclude/quarantine from acceptance claims; do not cite as VBA verification. |
| D-07 | **Medium** | FR-RPT-3, OQ-T3, R5 | `GetPresentedRange` fallback to `TableRange2` may include page-field area; runtime fidelity unverified. | `mko_treport/src/vbs/export/modTreportExtract.bas` | Validate presented bounds/exclusions and copy fidelity in Excel. |
| D-08 | **Low** | FR-RPT structure/idempotency hygiene | `CleanupDefaultSheets` is effectively a no-op; extra default sheets may remain in output workbooks. | `mko_treport/src/vbs/export/modTreportOutput.bas` | Confirm required final structure and remove unused default sheets deterministically if needed. |

## 11. Runtime versus static verification distinctions

| Verification type | Performed | Covered in this audit | Not covered / limitations |
|---|---|---|---|
| **Actual Excel runtime** | **No** | None | Macro execution, button clicks, PQ refresh timing/persistence, pivot filtering/layout results, copy/format fidelity, output generation, overwrite/idempotency in practice, lock/permission behavior |
| **Static inspection** | **Yes** | Exported VBA source logic, git state, workbook package/UI XML/VML, config tables/shared strings, file hashes/sizes/metadata, embedded binary string markers | Cannot prove compilation, true binding behavior, Excel OM runtime semantics, PQ engine behavior, actual pivot presented results |
| **Auxiliary (non-runtime)** | **Yes** | Zip/XML parsing, pivot-cache parsing (`.ai/tmp/pivots_parsed.json`), registry/path probes for Excel presence, review of existing local evidence artifacts | Not a substitute for runtime; some evidence derives from prior local artifacts/scripts not re-executed here |
| **Not verified** | **Explicit** | FR-RUN-1 runtime operation/compile, OQ-T1 runtime, OQ-T2, OQ-T3, OQ-T4 runtime parity, all B13 scenarios, performance/lock cases | Requires controlled execution on working copies in target Excel |

**Reason runtime not executed in this audit:** safe runtime validation requires controlled execution on working copies (preserving originals), confirmed macro trust/VBA access, and handling of potentially environment-specific Power Query connections; Excel presence was evidenced statically but no controlled, non-destructive runtime run was performed in this session.

## 12. Remaining blockers and precise corrective actions

1. **Fix D-01 (critical):** scope sheet-name collision detection to each output workbook in `modTreportRun.ProcessDataFile` (identical `report_sheet` names across different report workbooks must be allowed), while preserving within-workbook collision fail-fast.
2. **Fix D-02 / OQ-T4 / FR-RUN-1 (critical):** rebuild and synchronize the workbook VBA project from reviewed exports/source (remove `RunXlconcat`/`xlconcat_macro` remnants and stray duplicate modules), compile in VBA IDE, verify `run` controls invoke correct procedures, and re-export synchronized modules and `vbaProject.bin`.
3. **Address D-04 (high):** align file resolution scope/ambiguity handling with FR-FILE-1/D2 or obtain explicit PO clarification and update spec/plan.
4. **Prepare runnable acceptance config (D-05):** resolve `DOLGIT`/`DOLGIT ` report-name collision for successful-run evidence (or document as intentional negative case) without changing Trim semantics.
5. **Execute runtime acceptance on working copies (D-03):** run `Refresh_data` and all required `Create_report` scenarios in target Excel; record Excel build/macro/PQ prerequisites; verify persisted refresh, per-file `Year`/period resolution (including `report_wclip`), optional-category skip behavior, staging/no-leak, presented-range copy/fidelity, formatting, naming/placeholders/collisions, overwrite/idempotency, and negative scenarios; compare outputs against actual source presented data per D18.
6. **Clean evidence hygiene (D-06):** exclude stale `mko_treport/tests/treport_*` `__pycache__` bytecode from compliance claims.
7. **Verify output details (D-07, D-08):** confirm no page-filter leakage in copied ranges and deterministic final sheet structure.

## 13. Final recommendation

**Recommendation: reject acceptance (FAIL).** The delivery shows meaningful VBA implementation progress and a partially correct static control setup, but a critical multi-report generation defect, unresolved VBA project parity, and absent runtime validation prevent concluding that an approved rev.5-compliant, working VBA report-generation solution has been delivered. Re-audit after D-01/D-02 remediation and a complete, evidence-backed Excel runtime acceptance pass on working copies.

**Overall acceptance result: FAIL**

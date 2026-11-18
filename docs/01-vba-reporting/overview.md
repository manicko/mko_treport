---
id: vba-reporting-overview
domain: vba-reporting
tags:
  - vba
  - excel
  - reporting
  - refresh
  - architecture
related:
  - treport-configuration-model
  - treport-filters-dsl
  - treport-pipeline-dataflow
  - vba-build-automation
---

## Purpose

Describes the treport VBA reporting subsystem: a configuration-driven Excel macro solution for refreshing source data and generating static report workbooks. This is a high-level architecture and behavior overview based on current VBA exports and specifications; runtime outcomes are not fully verified (see [Runtime Verification Status](#runtime-verification-status)).

## Main Concepts

- **VBA-only runtime**: end-user workflows run natively in Excel VBA within the macro-enabled workbook (`treport.xlsm`). No external Python/PowerShell/BAT runtime is required for normal `Refresh_data` / `Create_report` execution.
- **Entry points**: two user-facing operations are provided by `modTreportRun`:
  - `Refresh_data` — refreshes configured source workbooks and persists refreshed data.
  - `Create_report` — produces static output report workbooks from configured pivots/data.
- **Configuration-driven behavior**: report inputs, file mappings, filters, layout options, and placeholder values are driven by `settings` workbook tables (see [Configuration model](./configuration-model.md)).
- **Static report outputs**: report generation captures presented pivot data and writes static values/formats to output sheets/workbooks (see [Pipeline and data flow](./pipeline-and-dataflow.md)).
- **Auxiliary build tooling**: modular VBA sources can be assembled into a workbook copy by a separate build script; this is a development/deployment aid, not an end-user runtime dependency (see [VBA build automation](../02-build-and-deployment/vba-build.md)).

## Scope and Boundaries

**In scope (implemented/defined at high level):**
- Reading and validating configuration from `settings` ListObjects (including required workbook tables/fields and optional P04/P07 extension columns when present).
- Resolving source workbooks, opening in appropriate access mode, and running refresh/report workflows.
- Parsing/applying filter DSL and derived period/`report_sheet` placeholder resolution (including offset forms).
- Applying pivot filters and extracting static report snapshots into output workbooks.
- Preparing output target sheets (plain or `template`-based from `ThisWorkbook`), transferring snapshots (`only_values`, `skip_header`) with configurable insertion/row-preparation modes (`insert_rows`, including template-scoped rework semantics), and applying layout/formatting.
- Per-output worksheet name planning and preflight for report generation: resolve/sanitize effective base names, group same-base subreports case-insensitively in original order, assign suffix-aware planned names, and fail fast on length/collision issues before affected worksheet creation.
- Two-stage `Create_report` assembly for multi-subreport groups: prepare each subreport on its own worksheet, then append prepared temps into the base/final worksheet with configured spacing/limits/bounds checks and required transfer/lifecycle behavior.
- Bounds-scoped Outline row grouping driven by `data_files.group_rows` with `general.rows_collapsed` control.
- Output workbook finalization (including temp cleanup after successful combine, removal of unproduced default sheets on successful completion paths with produced-sheet guards) and save/collision handling.
- Deterministic validation/fail-fast diagnostics and application-state restore/cleanup intent.

**Out of scope:**
- External runtime services or non-VBA end-user execution dependencies.

**Implemented extensions in current exports (P01/P04/P05/P07/P09, runtime-unverified):** `max-N` offsets in `data_files.filters` and `report_sheet` placeholders with deterministic bounds/domain validation; P01 combine-on-one-worksheet support (including `data_files.gap_rows`, required `general.report_rows_limit`, suffix-aware planned naming/collision preflight, two-stage prepare-then-combine flow, per-subreport/bounds validations, transfer/lifecycle handling, and produced-sheet/finalization guards); default-sheet cleanup with produced-sheet guards; template-based output sheets with `only_values`, `insert_rows` (legacy/non-template and template-scoped anchor rework) and `skip_header` snapshot handling; and `group_rows`/`rows_collapsed` Outline grouping. End-to-end behavior remains subject to manual L2 compile and L3 runtime verification on working copies (dataset-specific correctness, including combine/fidelity/limit/collision scenarios and `report_wclip` pivot/period semantics, not proven in this audit).

## Component Overview

Primary VBA modules under `mko_treport/src/vbs/export/` (authoritative code snapshot for this overview):

| Module | Responsibility (high-level) |
|---|---|
| `modTreportRun.bas` | Orchestration, entry points (`Refresh_data`, `Create_report`), scan/plan/preflight, Stage 1 preparation, Stage 2 combine/finalization sequencing, per-output lifecycle guards, and user-facing outcomes/status. |
| `modTreportConfig.bas` | Configuration reading, normalization/Trim, schema/value validation (including P01 config fields), path checks, contextual fail-fast. |
| `modTreportFiles.bas` | Source workbook resolution and workbook open/close/save session handling. |
| `modTreportRefresh.bas` | Best-effort synchronous refresh of connections/queries/pivot caches before persisted save. |
| `modTreportFilters.bas` | Filter DSL parsing/classification, min/max and period/placeholder resolution, data/report filter application. |
| `modTreportPivot.bas` | Pivot/field lookup, availability and placement handling, field selection, no-data heuristics. |
| `modTreportExtract.bas` | Presented-range detection, Stage 1 snapshot copy/formatting/outline application, and Stage 2 combine validations/append transfer helpers. |
| `modTreportOutput.bas` | Placeholder substitution, effective-name planning/suffix-aware validation and collision preflight, sheet/workbook preparation, temp-finalization/cleanup/save helpers. |
| `modTreportState.bas` | Capture/apply/restore Excel application state around runs. |
| `Sheet1.cls`, `Sheet2.cls`, `ThisWorkbook.cls` | Workbook/sheet document modules (minimal exported class shells). |

Data/control flow details are documented in [Pipeline and data flow](./pipeline-and-dataflow.md).

## Runtime Verification Status

- **Workbook binaries not inspectable in this audit**: `data/theiss/treport.xlsm`, `mko_treport/src/xls/treport.xlsm`, and `mko_treport/src/vbs/vbaProject.bin` cannot be reliably inspected for embedded VBA, control bindings, or live configuration from text tooling.
- **Runtime unproven**: end-to-end `Refresh_data`/`Create_report` execution, Power Query/pivot persistence, pivot filter behavior, P01 combine/naming/limit/bounds/transfer/cleanup behavior, snapshot/output fidelity, and build/compile acceptance are not evidenced by this documentation pass.
- **Test evidence limitation**: `mko_treport/tests/` currently contains only `__pycache__` artifacts (no `.py` sources), so it is not evidentiary for current VBA behavior.
- **Source of truth for this doc**: current exported VBA modules/scripts and specification/planning context; discrepancies and open verification items are tracked outside governed docs (e.g. `.ai/audit/problems`, `.ai/reports`).

## Related Docs

- [Configuration model](./configuration-model.md)
- [Filters DSL](./filters-dsl.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)
- [VBA build automation](../02-build-and-deployment/vba-build.md)
- [Documentation maintenance rules](../00-overview/doc-maintenance-rules.md)

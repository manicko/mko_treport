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

Describes the treport VBA reporting subsystem: a configuration-driven Excel macro solution for refreshing source data and generating static report workbooks. This is a high-level architecture and behavior overview based on current exported VBA sources; runtime outcomes are not fully verified (see [Runtime Verification Status](#runtime-verification-status)).

## Main Concepts

- **VBA-only runtime**: end-user workflows run natively in Excel VBA within the macro-enabled workbook (	report.xlsm). No external Python/PowerShell/BAT runtime is required for normal Refresh_data / Create_report execution.
- **Entry points**: two user-facing operations are provided by modTreportRun:
  - Refresh_data — refreshes configured source workbooks and persists refreshed data.
  - Create_report — produces static output report workbooks from configured pivots/data.
- **Configuration-driven behavior**: report inputs, file mappings, filters, layout options, naming/placeholders, and combine/limit settings are driven by settings workbook tables (see [Configuration model](./configuration-model.md)).
- **Static report outputs**: report generation captures presented pivot data and writes static values/formats to output sheets/workbooks (see [Pipeline and data flow](./pipeline-and-dataflow.md)).
- **Multi-subreport assembly (P01, runtime-unverified)**: for each output workbook, subreports that resolve to the same effective eport_sheet base name can be prepared and combined into one final worksheet using planned naming, per-subreport/bounds validation, and controlled finalization (see [Pipeline and data flow](./pipeline-and-dataflow.md)).
- **Auxiliary build tooling**: modular VBA sources can be assembled into a workbook copy by a separate build script; this is a development/deployment aid, not an end-user runtime dependency (see [VBA build automation](../02-build-and-deployment/vba-build.md)).

## Scope and Boundaries

**In scope (implemented/defined at high level in exported VBA; runtime-unverified unless stated):**
- Reading and validating configuration from settings ListObjects (including required workbook tables/fields and optional extension columns when present).
- Resolving source workbooks, opening in appropriate access mode, and running refresh/report workflows.
- Parsing/applying filter DSL and derived period/eport_sheet placeholder resolution (including offset forms).
- Applying pivot filters and extracting static report snapshots into output workbooks.
- Per-output worksheet planning/preflight for report generation: resolve/sanitize effective base names, group same-base subreports case-insensitively in deterministic emission order, assign suffix-aware planned names, and fail fast on name-length/collision issues before affected worksheet creation (see [Configuration model](./configuration-model.md), [Pipeline and data flow](./pipeline-and-dataflow.md)).
- Preparing output target sheets (plain or 	emplate-based from ThisWorkbook) and transferring snapshots (only_values, skip_header) with configurable insertion/row-preparation modes (insert_rows, including legacy/non-template and template-scoped rework semantics).
- Two-stage Create_report assembly for multi-subreport groups: Stage 1 prepares each subreport on its own worksheet preserving existing preparation/formatting/grouping behavior; Stage 2 appends prepared members into the base/final worksheet with configured spacing (gap_rows), per-subreport row-limit and projected Excel-bounds checks, defined transfer profile, and controlled temp/finalization behavior.
- Bounds-scoped Outline row grouping driven by data_files.group_rows with general.rows_collapsed control.
- Output workbook finalization and save guards: temp cleanup/reconciliation only after successful combine, removal of unproduced non-conflicting default sheets on successful completion paths, verification of final-produced sheets, and no-save-on-failure handling.
- Deterministic validation/fail-fast diagnostics and application-state restore/cleanup intent.

**Out of scope:**
- External runtime services or non-VBA end-user execution dependencies.

**Implemented extensions in current exports (runtime-unverified):** derived max-N offsets in data_files.filters and eport_sheet placeholders with deterministic bounds/domain validation; multi-subreport combine/assembly support (including data_files.gap_rows, required general per-subreport row limit accepted as eport_rows_limit or eport_rows_limit_general, suffix-aware planned naming/collision preflight, two-stage prepare-then-combine flow, per-subreport/bounds/content validations, transfer/lifecycle handling, and produced-sheet/finalization guards); default-sheet cleanup with produced-sheet guards; template-based output sheets with only_values, insert_rows (legacy/non-template and template-scoped anchor rework) and skip_header snapshot handling; and group_rows/ows_collapsed Outline grouping. End-to-end behavior remains subject to manual L2 compile and L3 runtime verification on working copies (dataset-specific correctness, including combine/fidelity/limit/collision scenarios and pivot/period semantics, not proven in this audit).

## Component Overview

Primary VBA modules under mko_treport/src/vbs/export/ (authoritative code snapshot for this overview):

| Module | Responsibility (high-level) |
|---|---|
| modTreportRun.bas | Orchestration, entry points (Refresh_data, Create_report), run planning/preflight sequencing, Stage 1 preparation and Stage 2 combine/finalization sequencing, per-output lifecycle/registry guards, and user-facing outcomes/status. |
| modTreportConfig.bas | Configuration reading, normalization/Trim, schema/value validation (including gap_rows and required general row-limit parameter), path checks, contextual fail-fast. |
| modTreportFiles.bas | Source workbook resolution and workbook open/close/save session handling. |
| modTreportRefresh.bas | Best-effort synchronous refresh of connections/queries/pivot caches before persisted save. |
| modTreportFilters.bas | Filter DSL parsing/classification, min/max and period/placeholder resolution, data/report filter application. |
| modTreportPivot.bas | Pivot/field lookup, availability and placement handling, field selection, no-data heuristics. |
| modTreportExtract.bas | Presented-range detection, Stage 1 snapshot copy/formatting/outline application, and Stage 2 combine validations/append transfer helpers. |
| modTreportOutput.bas | Placeholder substitution, effective-name planning/suffix-aware validation and collision preflight, sheet/workbook preparation, temp-finalization/cleanup/save helpers. |
| modTreportState.bas | Capture/apply/restore Excel application state around runs. |
| Sheet1.cls, Sheet2.cls, ThisWorkbook.cls | Workbook/sheet document modules (minimal exported class shells). |

Data/control flow details are documented in [Pipeline and data flow](./pipeline-and-dataflow.md).

## Runtime Verification Status

- **Workbook binaries not inspectable in this audit**: data/theiss/treport.xlsm, mko_treport/src/xls/treport.xlsm, and mko_treport/src/vbs/vbaProject.bin cannot be reliably inspected for embedded VBA, control bindings, or live configuration from text tooling.
- **Runtime unproven**: end-to-end Refresh_data/Create_report execution, Power Query/pivot persistence, pivot filter behavior, multi-subreport combine/assembly (grouping, planning/preflight, Stage 1/2 sequencing, append transfer fidelity including row heights/hidden/outline restoration, per-subreport row-limit/projected-bounds/content checks, temp lifecycle/finalization, and save/cleanup guards), snapshot/output fidelity, and build/compile acceptance are not evidenced by this documentation pass.
- **Test evidence limitation**: mko_treport/tests/ currently contains only __pycache__ artifacts (no .py sources), so it is not evidentiary for current VBA behavior.
- **Source of truth for this doc**: current exported VBA modules/scripts and verified implementation summaries; discrepancies and open verification items are tracked outside governed docs (e.g. .ai/audit/problems, .ai/reports).

## Related Docs

- [Configuration model](./configuration-model.md)
- [Filters DSL](./filters-dsl.md)
- [Pipeline and data flow](./pipeline-and-dataflow.md)
- [VBA build automation](../02-build-and-deployment/vba-build.md)
- [Documentation maintenance rules](../00-overview/doc-maintenance-rules.md)

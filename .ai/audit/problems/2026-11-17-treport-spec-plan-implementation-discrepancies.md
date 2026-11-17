# treport spec/plan vs implementation discrepancies (audit)

---
domain: vba-reporting
status: open
created: 2026-11-17
sources:
  - .ai/problems/001_treport_spec.md
  - .ai/plans/003-report-wclip-week-filter.md
  - .ai/plans/005-p01-max-offset-default-sheet-grouping.md
  - mko_treport/src/vbs/export/*.bas
  - docs/01-vba-reporting/*.md
---

## Purpose

Records discrepancies between planned/specified behavior and the current VBA export implementation, and verification gaps that affect confidence. These findings are separate from governed documentation updates except where policy/behavior must be described accurately for implemented functionality.

Evidence basis: exported VBA sources (`mko_treport/src/vbs/export/*.bas`, `*.cls`) and governed docs as of 2026-11-17. Workbook binaries/control bindings, manual VBE compile (L2), and end-to-end runtime behavior (L3) were **not** verified in this audit.

## Discrepancies

### AUD-P01-01 — Spec 001 rev.7 scope vs P01-era implemented behavior (exports)

- **Requirement/spec:** `.ai/problems/001_treport_spec.md` (rev.7) defines config-driven VBA refresh/report workflows, D2 recursive deterministic source resolution, generic filter/pivot pipeline with D9 uniform fail-fast/matched-subset, `report_sheet` `max_<PivotTable field>` placeholders resolved pre-filter/independent of `data_files.filters`, static output assembly/lifecycle, and D13 stating `group_rows` deferred/no-op. Spec is a baseline, not proof of implementation.
- **Observed implementation (exports):** Core mechanisms aligned in part (entry points in `modTreportRun.bas`; D2-style resolution in `modTreportFiles.ResolveDataFilePath`; D9-style selection via `modTreportPivot.ApplyFieldSelection`/`ApplyReportFilters`; pre-filter placeholder resolution in `modTreportRun.ResolveReportSheetPlaceholderMaxValues` / `modTreportOutput.bas`). Exports also include P01 extensions beyond rev.7 text: `max-N` offsets in `data_files.filters` scalars/ranges and `<max-N_field>` placeholders (`modTreportFilters.bas`, `modTreportOutput.bas`); `general.rows_collapsed` and bounds-scoped Outline grouping driven by `data_files.group_rows` (`modTreportConfig.bas`, `modTreportExtract.ApplyOutlineGrouping`, `modTreportRun.bas`) superseding D13 deferral; default-sheet cleanup with produced-sheet guards on successful output paths (`modTreportOutput.CleanupDefaultSheets`, `modTreportRun.bas`).
- **Gap/impact:** Spec rev.7 narrative does not reflect P01 extensions or D13 supersession; runtime/L2/L3 and workbook binding/embedded parity remain unverified. Documentation updated to describe implemented export behavior with runtime-unverified caveats.
- **Relevant files:** `.ai/problems/001_treport_spec.md`, `modTreportRun.bas`, `modTreportConfig.bas`, `modTreportFiles.bas`, `modTreportFilters.bas`, `modTreportPivot.bas`, `modTreportExtract.bas`, `modTreportOutput.bas`.

### AUD-P01-02 — Plan 003 alignment/closure not evidenced end-to-end

- **Plan reference:** `.ai/plans/003-report-wclip-week-filter.md` targets rev.7 placeholders, D9/D2 conformance, and `report_wclip` correctness (U2) with L1/L2/L3 evidence gates.
- **Observed implementation (exports):** D2 resolution, pre-filter/independent `max_<field>` placeholder handling, D9 uniform selection behavior, and offset/filter machinery are present in exports, but dataset-specific `report_wclip` correctness (e.g. `Year` vs grouped `Years`, period/axis identity, pivot-member-dependent min/max in context) and end-to-end runtime outcomes are not evidenced in this audit.
- **Gap/impact:** Plan closure and runtime correctness for `report_wclip` remain unproven pending working-copy L2 compile and L3 functional checks on relevant data.
- **Relevant files:** `.ai/plans/003-report-wclip-week-filter.md`, `modTreportFilters.bas`, `modTreportPivot.bas`, `modTreportRun.bas`, `modTreportFiles.bas`.

### AUD-P01-03 — Plan 005 P01 implementation claimed, L2/L3 pending

- **Plan reference:** `.ai/plans/005-p01-max-offset-default-sheet-grouping.md` (CR-1..CR-5): default-sheet cleanup, `max-N` offsets for filters/placeholders, Outline grouping (`group_rows`) and `general.rows_collapsed`.
- **Observed implementation (exports):** Corresponding code paths exist (`modTreportFilters.bas`, `modTreportOutput.bas`, `modTreportExtract.ApplyOutlineGrouping`, `modTreportConfig.bas`, `modTreportRun.bas`). Plan/log context indicates B1 implemented with L1 on a working copy in prior execution, but L2 manual compile and L3 functional/runtime verification are not evidenced in this audit.
- **Gap/impact:** Runtime correctness (cleanup guards in real outputs, offset/domain edge cases, Outline behavior/collapse, interaction with real pivots) remains unverified.
- **Relevant files:** `.ai/plans/005-p01-max-offset-default-sheet-grouping.md`, `modTreportRun.bas`, `modTreportConfig.bas`, `modTreportFilters.bas`, `modTreportExtract.bas`, `modTreportOutput.bas`, `mko_treport/src/scripts/build_treport_vba.ps1`.

### AUD-DOC-01 — Prior warn-and-skip description vs uniform D9 fail-fast (resolved in governed docs)

- **Reference:** Older/historical behavior described warn-and-skip for generalized multi-value `articleLevel4Name_clean` on no matches in some contexts; governed docs previously reflected that wording.
- **Observed implementation (exports):** `modTreportPivot.ApplyFieldSelection` / report filter application enforce uniform D9 behavior: missing/unavailable required field or no matchable configured values => fail fast with context; >=1 matchable => apply matched subset. No warn-and-skip path for no-match cases.
- **Resolution:** `docs/01-vba-reporting/filters-dsl.md` and `docs/01-vba-reporting/pipeline-and-dataflow.md` updated to describe uniform D9 fail-fast/matched-subset as implemented in exports (runtime still unverified).
- **Relevant files:** `modTreportPivot.bas`, `modTreportFilters.bas`, `modTreportRun.bas`, `docs/01-vba-reporting/filters-dsl.md`, `docs/01-vba-reporting/pipeline-and-dataflow.md`.

## Verification notes (not discrepancies by themselves)

- `Refresh_data`/`Create_report` workbook control bindings and removal of legacy `RunXlconcat` cannot be confirmed from text exports/binaries.
- Build script compile is best-effort and VBOM/Excel-environment dependent (`docs/02-build-and-deployment/vba-build.md`); no successful L2/L3 acceptance evidenced here.
- `mko_treport/src/xls/treport.xlsm` and other workbook artifacts may differ from exports; embedded VBA parity not evidenced.

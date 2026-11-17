---
id: vba-build-automation
domain: build-and-deployment
tags:
  - vba
  - build
  - deployment
  - powershell
  - excel-com
related:
  - vba-reporting-overview
  - treport-pipeline-dataflow
---

## Purpose

Documents the auxiliary automation used to assemble VBA code from exported modules into the tracked treport Excel workbook, and (optionally) to synchronize exports/binary artifacts or deploy a data copy. This is development/deployment tooling, not a runtime dependency for end-user `Refresh_data` / `Create_report` execution.

## Script and responsibilities

- **Script:** `mko_treport/src/scripts/build_treport_vba.ps1`
- **VBA sources (text, source of truth for code):** `mko_treport/src/vbs/export/` (`.bas` and `.cls`)
- **Default target (tracked working workbook):** `mko_treport/src/xls/treport.xlsm`

### Default behavior (no sync, no deploy)

- Builds managed VBA modules from `mko_treport/src/vbs/export/` into `mko_treport/src/xls/treport.xlsm` **in place**.
- Creates a **timestamped backup** of the target workbook in `.ai/tmp/` (by default) before modification and writes a build report there.
- **Does not** update `mko_treport/src/vbs/export/*.bas|*.cls` and **does not** update `mko_treport/src/vbs/vbaProject.bin` unless `-SyncExports` is specified explicitly.
- **Does not** modify `data\theiss\treport.xlsm` unless `-DeployToDataTheiss` is specified explicitly.

### Build/apply details

- Reads managed `.bas`/`.cls` exports, preserves module identities (no merge), and treats document modules separately from standard/class/userform modules.
- Rebuilds non-document managed components by removing them and re-adding standard modules from `.bas` exports with `CodeModule` body injection from stripped UTF-8 source (avoids `VBComponents.Import` for std modules).
- Synchronizes document-module code (`Sheet1`, `Sheet2`, `ThisWorkbook`) by updating `CodeModule` only (no worksheet/workbook component replacement).
- Prevents duplicate modules and removes obsolete legacy component `xlconcat_macro` when detectable.
- Attempts VBA compile/validation **best-effort** (programmatic compile may not be confirmable in all Excel/COM contexts).

### Optional explicit actions

- **`-SyncExports` (opt-in):** after a successful build/save, refreshes managed `.bas/.cls` files in the export directory via `VBComponent.Export()` from the single built workbook and extracts `xl/vbaProject.bin` to `mko_treport/src/vbs/vbaProject.bin`. Use intentionally (typically after successful manual VBE compile/L2 on a verified build). `VBComponent.Export` output is not post-normalized by the script (format/EOL/encoding may vary by VBE environment), and `vbaProject.bin` is a volatile binary artifact.
- **`-DeployToDataTheiss` (opt-in):** after a successful build/save (and optional sync), deletes `data\theiss\treport.xlsm` if present and copies the built/tracked `mko_treport/src/xls/treport.xlsm` to `data\theiss\treport.xlsm` (creates the directory if needed). The data copy is a derived/runtime copy and is **not** used as a sync source.
- **`-TargetWorkbook <path>`:** build an explicit workbook in place (still with backup/report). Default is the tracked working workbook above.

## Prerequisites and constraints

- **Windows + Excel COM automation:** script relies on local Excel application automation.
- **VBA project access trust (VBOM):** required for import/sync operations (Excel Trust Center: "Trust access to the VBA project object model").
- **Environment-dependent:** depends on local Office/Excel installation, macro security/trust configuration, file locks, and permissions.
- **Auxiliary only:** required for build/sync/deploy workflows, not for end-user report refresh/generation inside Excel.

## Idempotency and workflow intent

- The script is designed for repeatable (idempotent) assembly of managed VBA modules without duplicating components.
- Single tracked working workbook workflow: `mko_treport/src/xls/treport.xlsm` is the live workbook for settings/templates and receives current VBA via build. No separate "canonical" workbook copy is maintained by default.
- Routine builds should rely on backups (and git history for tracked artifacts) rather than implicit sync of exports/bin. Use `-SyncExports` only as an explicit, reviewed step from a verified build.

## Status and verification

- Successful end-to-end build, compile-clean state, and runtime acceptance are environment-dependent and must be verified per run (see build report; programmatic compile is best-effort). Manual VBE compile (L2) is required to claim compilation; functional checks (L3) are required when behavior changes.
- Usage details and supported options are defined by the script parameters/help; this document does not prescribe a CI pipeline or claim unattended execution support.

## Related docs

- [VBA reporting overview](../01-vba-reporting/overview.md)
- [Pipeline and data flow](../01-vba-reporting/pipeline-and-dataflow.md)
- [Documentation maintenance rules](../00-overview/doc-maintenance-rules.md)

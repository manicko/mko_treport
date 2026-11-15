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

Documents the auxiliary automation used to assemble and synchronize the modular treport VBA project into an Excel macro workbook copy. This is development/deployment tooling, not a runtime dependency for end-user `Refresh_data` / `Create_report` execution.

## Script and Responsibilities

- **Script:** `mko_treport/src/scripts/build_treport_vba.ps1`
- **Inputs (high-level):** exported VBA modules/classes in `mko_treport/src/vbs/export/` and a source workbook.
- **Intended responsibilities (as implemented in script):**
  - Prepare a working workbook copy (or operate in-place per script options) with backup handling.
  - Assemble/import managed VBA components from exports into the workbook VBA project (preserving module identities; idempotent intent).
  - Synchronize managed exports and embedded VBA artifacts (including `vbaProject.bin` handling as applicable) and remove obsolete/legacy components when detectable.
  - Attempt compile/validation steps and generate a build/verification report (auxiliary artifact output path controlled by script parameters).
  - Scan/report run-binding/control consistency where supported.

## Prerequisites and Constraints

- **Windows + Excel COM automation:** script relies on local Excel application automation.
- **VBA project access trust:** access to the VBA project object model (VBOM/trust setting) is required for import/sync operations.
- **Environment-dependent:** success depends on local Office/Excel installation, macro security/trust configuration, file locks, and permissions.
- **Auxiliary only:** required for build/sync workflows, not for end-user report refresh/generation inside Excel.

## Idempotency and Usage Intent

- The script is designed for repeatable (idempotent) assembly/sync of managed VBA modules without duplicating components.
- It targets a working copy workflow to avoid unsafe direct edits to canonical artifacts; exact parameter behavior should be taken from the script itself when executed.
- Usage details and supported options are defined by the script parameters/help; this document does not prescribe a CI pipeline or claim unattended execution support.

## Status and Verification

- **Current status:** script exists in the repository working tree and contains the build/sync/compile/report logic described above.
- **Verification:** successful end-to-end build, compile-clean workbook parity, and runtime acceptance have **not** been evidenced in this documentation audit.
- **Known gaps:** workbook binaries (`data/theiss/treport.xlsm`, `mko_treport/src/xls/treport.xlsm`, `mko_treport/src/vbs/vbaProject.bin`) are not inspectable here; parity between exports and embedded VBA remains an open verification concern tracked outside governed docs.

## Related Docs

- [VBA reporting overview](../01-vba-reporting/overview.md)
- [Pipeline and data flow](../01-vba-reporting/pipeline-and-dataflow.md)
- [Documentation maintenance rules](../00-overview/doc-maintenance-rules.md)

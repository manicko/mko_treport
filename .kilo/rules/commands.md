# Agent instructions (VBA / Excel) — English

## Context: “cannot compile” when Excel is installed
- **Excel installed does not mean** an external process can compile/check VBA. Common blockers: **Trust Center** must allow **Trust access to the VBA project object model (VBOM)**; macro/security policies (Protected View, blocked macros from internet, disabled content); Office edition/policies differ by machine; **no stable, universal COM API** to run “Compile VBAProject” identically across Office builds/locales (scripts may try VBE CommandBars/menu/reflection and still be best-effort).
- Many “I cannot compile” messages come from: missing VBOM trust, running without an interactive Excel session that permits the action, or confusing **VBA syntax compile** (L2) with **runtime/functional** checks (L3). Distinguish these in every report.

## Repository facts (this project)
- VBA export sources (edit these): `mko_treport/src/vbs/export/*.bas`, `*.cls`
- Workbooks: `data/theiss/treport.xlsm`, `mko_treport/src/xls/treport.xlsm` (treat as sensitive; use backups/working copies)
- Build/sync script: `mko_treport/src/scripts/build_treport_vba.ps1` (see `docs/02-build-and-deployment/vba-build.md`)
- Build reports: `.ai/tmp/build_treport_vba_*.txt`

## Required workflow when you change VBA exports
1. **Backup / working copy first.** Never edit canonical workbooks blindly for verification. Prefer:
   ```powershell
   .\mko_treport\src\scripts\build_treport_vba.ps1 `
     -SourceWorkbook "data\theiss\treport.xlsm" `
     -WorkCopyPath ".ai\tmp\treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm" `
     -BackupDir ".ai\tmp" `
     -ReportPath ".ai\tmp\build_treport_vba_<YYYYMMDD_HHMMSS>.txt"
   ```
   Adjust parameters to match the script/help if needed. Always record the report path.

2. **L1 — automated sync/import (script):** review the generated report for import/legacy/duplicate issues and the **Compile/validation** note (best-effort). If it reports missing VBOM access, state clearly: enable **Trust access to the VBA project object model** in Excel Trust Center (Macro Settings), close Excel, retry. Note corporate policies may block this.

3. **L2 — manual compile in Excel (required for “compiles” claim):** open the **working copy** `.xlsm` in Excel → VBA Editor → **Debug → Compile VBAProject**.  
   - If clean: record file path, date/time, and “L2 passed (manual VBE Compile)”.  
   - If errors: paste/fix from exports, re-run L1 on a new working copy, repeat until resolved (or report blockers precisely).

4. **L3 — functional checks (when logic/run behavior changes):** test `Refresh_data` / `Create_report` on **copies** of data/workbooks only (do not corrupt originals). Verify expected outputs and that source files remain unchanged per project rules.

5. **Syncing canonical artifacts:** only after L2 (and L3 if applicable) consider syncing exports/bin via script options (e.g. default sync behavior per docs) — do **not** overwrite canonical state without an explicit task/approval in context.

## What to report after VBA edits (always)
- List changed files under `mko_treport/src/vbs/export/`.
- Exact `build_treport_vba.ps1` command used (or why not run).
- Path to build report and short summary (SUCCESS/ISSUES/ERROR; note VBOM issues if any).
- L2 result: done (with file/time) / not done (state this is a limitation, not proof of compilation).
- L3 result if performed, or “not performed”.

## Do not
- Claim “compiled successfully” based only on the script’s best-effort `VBProject.Compile` call without L2 when a user asks for verification.
- Run or request running macros on production/original data without working copies and explicit need.

## Reference
- Full analysis report (created 2026-11-15): `.ai/reports/vba-compile-and-verification-report-2026-11-15.md`
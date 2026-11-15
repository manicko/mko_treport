<#
.SYNOPSIS
One-step VBA build/import utility for treport.xlsm (FR-DEP).

.DESCRIPTION
Assembles the current VBA project from modular sources in mko_treport/src/vbs/export/
into a working copy of treport.xlsm. Auxiliary dev/deployment tooling only (not a
runtime dependency). Idempotent: repeated runs do not create duplicates.

Behavior (FR-DEP):
- Single command/script invocation.
- Reads all current VBA source modules from the export directory (9 standard .bas,
  3 document .cls). Preserves individual module identities; never merges modules and
  never treats document modules as ordinary standard modules.
- Prepares a working copy of the workbook (source unchanged by default) with backup.
- Imports/updates all required standard modules (remove-by-name then Import).
- Synchronizes document-module code safely (Sheet1/Sheet2/ThisWorkbook) by updating
  CodeModule only (no worksheet/workbook component replacement).
- Prevents duplicate modules and removes obsolete legacy component 'xlconcat_macro'.
- Attempts VBA compile/validation (best-effort); reports prerequisite if programmatic
  access to the VBA project object model is not trusted.
- Saves resulting workbook. Exports updated vbaProject.bin and individual source
  exports from the resulting workbook when SyncExports is enabled.
- Produces a clear build/verification report (stdout and file in .ai/tmp/).

.PARAMETER SourceWorkbook
Path to source treport.xlsm. Default: data\theiss\treport.xlsm (resolved from repo root).

.PARAMETER ExportDir
Directory containing VBA source exports (.bas/.cls). Default: mko_treport\src\vbs\export.

.PARAMETER WorkCopyPath
Path for the working/output workbook copy. Default: <BackupDir>\treport_working_<ts>.xlsm.
If -InPlace is specified, the target workbook itself is updated (after backup).

.PARAMETER BackupDir
Directory for backups and default work copy/report. Default: .ai\tmp.

.PARAMETER InPlace
Update TargetWorkbook in place (a backup is always created first). Mutually exclusive
with specifying a custom WorkCopyPath that differs from TargetWorkbook.

.PARAMETER TargetWorkbook
Workbook to update when -InPlace is used. Default: same as SourceWorkbook.

.PARAMETER SyncExports
When set (default), export vbaProject.bin to mko_treport\src\vbs\vbaProject.bin and
refresh individual .bas/.cls files in ExportDir from the resulting workbook.

.PARAMETER NoSyncExports
Disable export/bin sync back to source tree.

.PARAMETER ReportPath
Explicit report file path. Default: <BackupDir>\build_treport_vba_<ts>.txt.

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1 -SourceWorkbook data\theiss\treport.xlsm -WorkCopyPath .ai\tmp\treport_test.xlsm -BackupDir .ai\tmp
#>
[CmdletBinding()]
param(
    [string]$SourceWorkbook,
    [string]$ExportDir,
    [string]$WorkCopyPath,
    [string]$BackupDir,
    [string]$TargetWorkbook,
    [switch]$InPlace,
    [switch]$NoSyncExports,
    [string]$ReportPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Resolve-RepoRoot {
    # script is at mko_treport/src/scripts/build_treport_vba.ps1 -> repo root is 3 levels up
    return (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
}

function Get-Timestamp { return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
function Get-Stamp     { return (Get-Date -Format 'yyyyMMdd_HHmmss') }

function Write-Step([ref]$Lines, [string]$Message) {
    $Lines.Value += $Message
    Write-Host $Message
}

function Release-ComObject([ref]$Obj) {
    try {
        if ($null -ne $Obj.Value) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Obj.Value) } catch { }
            $Obj.Value = $null
        }
    } catch { }
}

function Strip-VbaHeader([string]$Content) {
    # Remove VERSION/Attribute VB_* header block from exported .cls/.bas text so it can
    # be injected into a CodeModule. Keep everything from first non-header line onward.
    $text = $Content -replace "`r`n", "`n"
    $lines = $text -split "`n"
    $start = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match '^VERSION\s' -or $line -match '^BEGIN(\s|\r?$)' -or $line -match '^END(\s|\r?$)' -or $line -match '^\s*MultiUse\s*=') {
            $start = $i + 1
            continue
        }
        if ($line -match '^Attribute\s+VB_') {
            $start = $i + 1
            continue
        }
        # stop at first real code/procedure/comment line
        break
    }
    if ($start -ge $lines.Count) { return '' }
    $body = ($lines[$start..($lines.Count - 1)] -join "`r`n")
    # trim leading blank lines only
    return $body.TrimStart("`r", "`n")
}

function Get-ComponentInfo($VbProj) {
    $list = @()
    $comps = $VbProj.VBComponents
    $count = $comps.Count
    for ($i = 1; $i -le $count; $i++) {
        $c = $comps.Item($i)
        $t = [int]$c.Type
        $typeName = switch ($t) { 1 { 'StdModule' } 2 { 'Class' } 3 { 'UserForm' } 100 { 'Document' } default { "Type$t" } }
        $list += ("{0} ({1})" -f $c.Name, $typeName)
    }
    return $list
}

function Try-GetRunBinding($Workbook) {
    # Best-effort: detect a shape/button whose OnAction references 'run'
    try {
        $hits = @()
        $sheets = $Workbook.Worksheets
        for ($si = 1; $si -le $sheets.Count; $si++) {
            $ws = $sheets.Item($si)
            try {
                $shapes = $ws.Shapes
                for ($hi = 1; $hi -le $shapes.Count; $hi++) {
                    $sh = $shapes.Item($hi)
                    $onAction = $null
                    try { $onAction = [string]$sh.OnAction } catch { $onAction = $null }
                    if ($onAction -and ($onAction -match '(?i)(^|[!\.])run(\b|$)')) {
                        $hits += ("{0}!{1} -> {2}" -f $ws.Name, $sh.Name, $onAction)
                    }
                }
            } catch { }
        }
        if ($hits.Count -gt 0) { return ($hits -join '; ') }
        return 'not detected (best-effort scan found no shape OnAction referencing run)'
    } catch {
        return ('not detectable: ' + $_.Exception.Message)
    }
}

$report = @()
$excel = $null
$workbook = $null
$vbProj = $null
$zipArchive = $null
$exitCode = 0

try {
    $repoRoot = Resolve-RepoRoot

    if ([string]::IsNullOrWhiteSpace($SourceWorkbook)) { $SourceWorkbook = Join-Path $repoRoot 'data\theiss\treport.xlsm' }
    if ([string]::IsNullOrWhiteSpace($ExportDir))      { $ExportDir      = Join-Path $repoRoot 'mko_treport\src\vbs\export' }
    if ([string]::IsNullOrWhiteSpace($BackupDir))      { $BackupDir      = Join-Path $repoRoot '.ai\tmp' }
    if ([string]::IsNullOrWhiteSpace($TargetWorkbook)) { $TargetWorkbook = $SourceWorkbook }

    $SyncExports = -not $NoSyncExports.IsPresent
    $ts = Get-Timestamp
    $stamp = Get-Stamp

    if (-not (Test-Path -LiteralPath $SourceWorkbook)) { throw "Source workbook not found: $SourceWorkbook" }
    if (-not (Test-Path -LiteralPath $ExportDir))      { throw "Export directory not found: $ExportDir" }
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

    # Discover source modules (preserve individual identities; do not merge)
    $stdFiles = @(Get-ChildItem -LiteralPath $ExportDir -Filter '*.bas' -File | Sort-Object Name)
    $docFiles = @(Get-ChildItem -LiteralPath $ExportDir -Filter '*.cls' -File | Sort-Object Name)
    $docNames = @('Sheet1', 'Sheet2', 'ThisWorkbook')
    $missingDoc = @($docNames | Where-Object { -not ($docFiles | Where-Object { $_.BaseName -eq $_ }) })
    # above Where-Object uses $_ twice incorrectly; compute properly
    $missingDoc = @()
    foreach ($dn in $docNames) {
        if (-not ($docFiles | Where-Object { $_.BaseName -eq $dn })) { $missingDoc += $dn }
    }

    $sourceResolved = (Resolve-Path -LiteralPath $SourceWorkbook).Path
    $exportResolved = (Resolve-Path -LiteralPath $ExportDir).Path
    $backupResolved = (Resolve-Path -LiteralPath $BackupDir).Path

    # Determine target/work copy and backup
    $backupPath = Join-Path $backupResolved ("treport_backup_{0}.xlsm" -f $stamp)
    if ($InPlace) {
        $effectiveTarget = (if ([string]::IsNullOrWhiteSpace($TargetWorkbook)) { $sourceResolved } else { $TargetWorkbook })
        $workPath = [System.IO.Path]::GetFullPath($effectiveTarget)
    } else {
        if ([string]::IsNullOrWhiteSpace($WorkCopyPath)) {
            $workPath = Join-Path $backupResolved ("treport_working_{0}.xlsm" -f $stamp)
        } else {
            $workPath = [System.IO.Path]::GetFullPath($WorkCopyPath)
        }
    }

    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path $backupResolved ("build_treport_vba_{0}.txt" -f $stamp)
    } else {
        $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    }

    Write-Step ([ref]$report) "=== build_treport_vba (FR-DEP) ==="
    Write-Step ([ref]$report) ("Timestamp: {0}" -f $ts)
    Write-Step ([ref]$report) ("Repo root: {0}" -f $repoRoot)
    Write-Step ([ref]$report) ("Source workbook: {0}" -f $sourceResolved)
    Write-Step ([ref]$report) ("Export dir: {0}" -f $exportResolved)
    Write-Step ([ref]$report) ("Backup dir: {0}" -f $backupResolved)
    Write-Step ([ref]$report) ("Target/work copy: {0}" -f $workPath)
    Write-Step ([ref]$report) ("InPlace: {0}" -f [bool]$InPlace)
    Write-Step ([ref]$report) ("SyncExports: {0}" -f $SyncExports)
    Write-Step ([ref]$report) ("Standard modules found ({0}): {1}" -f $stdFiles.Count, (($stdFiles | ForEach-Object { $_.Name }) -join ', '))
    Write-Step ([ref]$report) ("Document modules found ({0}): {1}" -f $docFiles.Count, (($docFiles | ForEach-Object { $_.Name }) -join ', '))
    if ($missingDoc.Count -gt 0) {
        Write-Step ([ref]$report) ("WARNING: missing expected document module sources: {0}" -f ($missingDoc -join ', '))
    }

    # Prepare backup and working copy (never modify source unless explicitly InPlace target)
    if ($InPlace) {
        $targetResolved = (Resolve-Path -LiteralPath $workPath).Path
        Copy-Item -LiteralPath $targetResolved -Destination $backupPath -Force
        Write-Step ([ref]$report) ("Backup created: {0} (from in-place target)" -f $backupPath)
        # work on target in place (already backed up)
        $workPath = $targetResolved
    } else {
        # backup source for traceability, then create working copy
        Copy-Item -LiteralPath $sourceResolved -Destination $backupPath -Force
        $workDir = Split-Path -Parent $workPath
        if ($workDir -and -not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Force -Path $workDir | Out-Null }
        Copy-Item -LiteralPath $sourceResolved -Destination $workPath -Force
        $workPath = (Resolve-Path -LiteralPath $workPath).Path
        Write-Step ([ref]$report) ("Backup created: {0} (source snapshot)" -f $backupPath)
        Write-Step ([ref]$report) ("Working copy prepared: {0}" -f $workPath)
    }

    # Start Excel COM
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw ("Failed to start Excel.Application COM automation. Is Excel installed? Details: {0}" -f $_.Exception.Message)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.EnableEvents = $false } catch { }
    try { $excel.AskToUpdateLinks = $false } catch { }

    $workbook = $excel.Workbooks.Open($workPath)
    try {
        $vbProj = $workbook.VBProject
    } catch {
        throw ("Cannot access VBA project object model. Enable 'Trust access to the VBA project object model' in Excel Trust Center (Macro Settings). Details: {0}" -f $_.Exception.Message)
    }

    # Record components before
    $before = @(Get-ComponentInfo $vbProj)
    Write-Step ([ref]$report) ("Components before ({0}): {1}" -f $before.Count, ($before -join ', '))

    $imported = @()
    $syncedDocs = @()
    $removedLegacy = @()
    $removedDupes = @()
    $issues = @()

    # Remove legacy obsolete module if present
    $legacyNames = @('xlconcat_macro')
    foreach ($legacy in $legacyNames) {
        $foundLegacy = @()
        $comps = $vbProj.VBComponents
        for ($i = $comps.Count; $i -ge 1; $i--) {
            $c = $comps.Item($i)
            if ($c.Name -ieq $legacy) { $foundLegacy += $c }
        }
        foreach ($c in $foundLegacy) {
            $t = [int]$c.Type
            if ($t -eq 1 -or $t -eq 2) { # standard/class removable
                $vbProj.VBComponents.Remove($c)
                $removedLegacy += $legacy
            } else {
                $issues += ("Legacy component '{0}' found but not a removable standard/class module (type {1}); left unchanged" -f $legacy, $t)
            }
        }
    }

    # Remove any accidental duplicate-named components (keep none for .bas re-import targets; handled below)
    # Standard modules: remove existing by name if present then Import (idempotent, prevents dupes)
    foreach ($f in $stdFiles) {
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        # remove all existing components with same name (standard/class/userform types) to avoid duplicates
        $comps = $vbProj.VBComponents
        for ($i = $comps.Count; $i -ge 1; $i--) {
            $c = $comps.Item($i)
            if ($c.Name -ieq $modName) {
                $ct = [int]$c.Type
                if ($ct -eq 1 -or $ct -eq 2 -or $ct -eq 3) {
                    $vbProj.VBComponents.Remove($c)
                } else {
                    $issues += ("Name collision: '{0}' exists as document/type {1}; cannot replace as standard module" -f $modName, $ct)
                }
            }
        }
        $null = $vbProj.VBComponents.Import($f.FullName)
        $imported += $modName
    }

    # Document modules: update CodeModule safely (do NOT replace components)
    foreach ($f in $docFiles) {
        $docName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $comp = $null
        try {
            $comp = $vbProj.VBComponents.Item($docName)
        } catch {
            $issues += ("Document module '{0}' not found in workbook; skipped (component not created/replaced)" -f $docName)
            $comp = $null
        }
        if ($null -ne $comp) {
            $raw = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            $body = Strip-VbaHeader $raw
            $cm = $comp.CodeModule
            $lineCount = [int]$cm.CountOfLines
            if ($lineCount -gt 0) { $cm.DeleteLines(1, $lineCount) }
            if (-not [string]::IsNullOrEmpty($body)) { $cm.AddFromString($body) }
            $syncedDocs += $docName
        }
    }

    # Detect duplicate component names after operations
    $nameCounts = @{}
    $compsAfterOps = $vbProj.VBComponents
    for ($i = 1; $i -le $compsAfterOps.Count; $i++) {
        $nm = $compsAfterOps.Item($i).Name
        if (-not $nameCounts.ContainsKey($nm)) { $nameCounts[$nm] = 0 }
        $nameCounts[$nm]++
    }
    $dupeNames = @($nameCounts.Keys | Where-Object { $nameCounts[$_] -gt 1 })
    if ($dupeNames.Count -gt 0) {
        # attempt to remove extra non-document duplicates by name
        foreach ($dn in $dupeNames) {
            $instances = @()
            for ($i = 1; $i -le $vbProj.VBComponents.Count; $i++) { $instances += $vbProj.VBComponents.Item($i) }
            $same = @($instances | Where-Object { $_.Name -ieq $dn })
            if ($same.Count -gt 1) {
                # keep first document instance if any, else keep first; remove rest if removable
                $keepIndex = 0
                for ($k = 0; $k -lt $same.Count; $k++) { if ([int]$same[$k].Type -eq 100) { $keepIndex = $k; break } }
                for ($k = 0; $k -lt $same.Count; $k++) {
                    if ($k -eq $keepIndex) { continue }
                    $ct = [int]$same[$k].Type
                    if ($ct -ne 100) {
                        try { $vbProj.VBComponents.Remove($same[$k]); $removedDupes += ("{0} (extra)" -f $dn) } catch { $issues += ("Could not remove duplicate '{0}': {1}" -f $dn, $_.Exception.Message) }
                    } else {
                        $issues += ("Duplicate document component name '{0}' remains (document components cannot be removed)" -f $dn)
                    }
                }
            }
        }
    }

    # Compile/validate (best-effort)
    $compileStatus = 'not attempted'
    $compileOk = $false
    try {
        # Some Excel versions don't expose Compile; use late-bound call
        $vbProj.GetType().InvokeMember('Compile', [System.Reflection.BindingFlags]::InvokeMethod, $null, $vbProj, $null) | Out-Null
        $compileStatus = 'compiled (Compile invoked successfully)'
        $compileOk = $true
    } catch {
        $compileStatus = ("compile not confirmed: {0}" -f $_.Exception.Message)
        $issues += "VBA compile could not be confirmed programmatically (may be unsupported in this Excel/COM context); manual compile check recommended in VBE (Debug > Compile)."
    }

    # Save workbook
    $workbook.Save()
    Write-Step ([ref]$report) "Workbook saved."

    # Extract vbaProject.bin from saved xlsm (Open Packaging Convention zip)
    $vbaBinExported = $false
    $vbaBinPath = $null
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $vbaBinPath = Join-Path $repoRoot 'mko_treport\src\vbs\vbaProject.bin'
        if ($SyncExports) {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($workPath)
            $zipArchive = $zip
            $entry = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/vbaProject.bin' } | Select-Object -First 1
            if ($null -ne $entry) {
                $destDir = Split-Path -Parent $vbaBinPath
                if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Force -Path $destDir | Out-Null }
                $inStream = $entry.Open()
                try {
                    $outStream = [System.IO.File]::Open($vbaBinPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                    try { $inStream.CopyTo($outStream) } finally { $outStream.Dispose() }
                } finally { $inStream.Dispose() }
                $vbaBinExported = $true
                Write-Step ([ref]$report) ("vbaProject.bin exported: {0}" -f $vbaBinPath)
            } else {
                $issues += "xl/vbaProject.bin not found inside saved workbook package; bin export skipped."
                Write-Step ([ref]$report) "WARNING: xl/vbaProject.bin not found in workbook package."
            }
        }
    } catch {
        $issues += ("Failed to export vbaProject.bin: {0}" -f $_.Exception.Message)
        Write-Step ([ref]$report) ("WARNING: vbaProject.bin export failed: {0}" -f $_.Exception.Message)
    } finally {
        if ($null -ne $zipArchive) { try { $zipArchive.Dispose() } catch { }; $zipArchive = $null }
    }

    # Export individual components back to export dir (sync) from resulting workbook
    $exportedBack = @()
    if ($SyncExports) {
        $compsFinalForExport = $vbProj.VBComponents
        for ($i = 1; $i -le $compsFinalForExport.Count; $i++) {
            $c = $compsFinalForExport.Item($i)
            $nm = $c.Name
            $ct = [int]$c.Type
            $ext = $null
            if ($ct -eq 1) { $ext = '.bas' }
            elseif ($ct -eq 2) { $ext = '.cls' }
            elseif ($ct -eq 100) { $ext = '.cls' }
            if ($null -ne $ext) {
                # only export components that correspond to managed sources or are known doc modules
                $srcFile = Join-Path $exportResolved ($nm + $ext)
                $isManaged = (Test-Path -LiteralPath $srcFile)
                if ($isManaged) {
                    try { $c.Export($srcFile); $exportedBack += ($nm + $ext) } catch { $issues += ("Failed to export component {0}: {1}" -f $nm, $_.Exception.Message) }
                }
            }
        }
        Write-Step ([ref]$report) ("Source exports refreshed ({0}): {1}" -f $exportedBack.Count, ($exportedBack -join ', '))
    } else {
        Write-Step ([ref]$report) "SyncExports disabled: vbaProject.bin and source exports not updated."
    }

    # Final component inventory and run binding
    $after = @(Get-ComponentInfo $vbProj)
    $runBinding = Try-GetRunBinding $workbook

    # Close workbook (already saved)
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null

    # Build final report
    $report += ""
    $report += "--- Summary ---"
    $report += ("Modules imported (standard .bas, remove-then-import): {0}" -f (($imported | Sort-Object -Unique) -join ', '))
    $report += ("Document modules synced (CodeModule update): {0}" -f (($syncedDocs | Sort-Object -Unique) -join ', '))
    $report += ("Legacy removed: {0}" -f ($(if ($removedLegacy.Count -gt 0) { ($removedLegacy | Sort-Object -Unique) -join ', ' } else { 'none' })))
    $report += ("Duplicates removed: {0}" -f ($(if ($removedDupes.Count -gt 0) { ($removedDupes -join ', ') } else { 'none' })))
    $report += ("Compile/validation: {0}" -f $compileStatus)
    $report += ("vbaProject.bin exported: {0}" -f ($(if ($vbaBinExported) { $vbaBinPath } elseif ($SyncExports) { 'no' } else { 'skipped (NoSyncExports)' })))
    $report += ("Run control binding (best-effort): {0}" -f $runBinding)
    $report += ("Final components ({0}): {1}" -f $after.Count, ($after -join ', '))

    # Simple consistency checks
    $expectedStd = @($stdFiles | ForEach-Object { $_.BaseName })
    $expectedDoc = @('Sheet1','Sheet2','ThisWorkbook')
    $finalNames = @()
    # re-open not needed; we captured $after strings like "Name (Type)" - parse names
    foreach ($entry in $after) { if ($entry -match '^(.+?)\s+\(') { $finalNames += $Matches[1] } }
    $missingExpected = @()
    foreach ($e in ($expectedStd + $expectedDoc)) { if ($finalNames -notcontains $e) { $missingExpected += $e } }
    $hasLegacy = ($finalNames -contains 'xlconcat_macro')
    $report += ("Expected components present: {0}" -f ($(if ($missingExpected.Count -eq 0) { 'yes' } else { ('NO, missing: ' + ($missingExpected -join ', ')) })))
    $report += ("Legacy xlconcat_macro present: {0}" -f ($(if ($hasLegacy) { 'YES (unexpected)' } else { 'no' })))
    $dupeFinal = @($nameCounts.Keys | Where-Object { $nameCounts[$_] -gt 1 })
    $report += ("Duplicate component names after build: {0}" -f ($(if ($dupeFinal.Count -eq 0) { 'none' } else { ($dupeFinal -join ', ') })))

    if ($issues.Count -gt 0) {
        $report += ""
        $report += "--- Issues / Prerequisites ---"
        foreach ($iss in $issues) { $report += ("- " + $iss) }
    } else {
        $report += ""
        $report += "--- Issues / Prerequisites ---"
        $report += "- none"
    }

    $overallOk = ($missingExpected.Count -eq 0 -and -not $hasLegacy -and $dupeFinal.Count -eq 0)
    $report += ""
    $report += ("RESULT: {0}" -f ($(if ($overallOk) { 'SUCCESS' } else { 'COMPLETED WITH ISSUES' })))
    $report += ("Report file: {0}" -f $ReportPath)

    # Write report file
    $reportText = ($report -join "`r`n")
    $reportDir = Split-Path -Parent $ReportPath
    if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
    [System.IO.File]::WriteAllText($ReportPath, $reportText, [System.Text.Encoding]::UTF8)

    if (-not $overallOk) { $exitCode = 1 }
}
catch {
    $errMsg = $_.Exception.Message
    $detail = $_ | Out-String
    $report += ""
    $report += "!!! ERROR !!!"
    $report += $errMsg
    $report += $detail
    try {
        if ($workbook -ne $null) { $workbook.Close($false) }
    } catch { }
    try {
        if ($excel -ne $null) { $excel.Quit() }
    } catch { }
    $reportText = ($report -join "`r`n")
    try {
        $backupResolved2 = if ($BackupDir) { $BackupDir } else { (Join-Path (Get-Location) '.ai\tmp') }
        if (-not (Test-Path -LiteralPath $backupResolved2)) { New-Item -ItemType Directory -Force -Path $backupResolved2 | Out-Null }
        $failReport = if ($ReportPath) { $ReportPath } else { Join-Path $backupResolved2 ("build_treport_vba_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss')) }
        [System.IO.File]::WriteAllText($failReport, $reportText, [System.Text.Encoding]::UTF8)
        Write-Host ("Report file: {0}" -f $failReport)
    } catch { }
    Write-Host $reportText
    exit 1
}
finally {
    if ($null -ne $zipArchive) { try { $zipArchive.Dispose() } catch { }; $zipArchive = $null }
    # release COM objects in reverse order
    $refWb = [ref]$workbook; Release-ComObject $refWb; $workbook = $refWb.Value
    $refVb = [ref]$vbProj;   Release-ComObject $refVb; $vbProj = $refVb.Value
    $refXl = [ref]$excel;    Release-ComObject $refXl; $excel = $refXl.Value
    try { [GC]::Collect() } catch { }
    try { [GC]::WaitForPendingFinalizers() } catch { }
}

exit $exitCode

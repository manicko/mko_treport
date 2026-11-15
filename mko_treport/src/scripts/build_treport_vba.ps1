<#
.SYNOPSIS
One-step VBA build/import utility for treport.xlsm (FR-DEP).

.DESCRIPTION
Assembles the current VBA project from modular sources in mko_treport/src/vbs/export/
into treport.xlsm workbooks. Auxiliary dev/deployment tooling only (not a runtime
dependency). Idempotent: repeated runs do not create duplicates.

Default behavior:
- Updates BOTH canonical workbooks in place (after backup), using mko_treport/src/vbs/export
  as the source of VBA modules:
  1) data\theiss\treport.xlsm
  2) mko_treport\src\xls\treport.xlsm
- Syncs exports/bin back from the second (last) updated workbook when SyncExports is enabled
  (default), so mko_treport/src/vbs/vbaProject.bin and managed .bas/.cls in ExportDir match it.
- A single Excel COM session is used to apply the same exported modules to both workbooks.

Behavior (FR-DEP):
- Single command/script invocation.
- Reads all current VBA source modules from the export directory (9 standard .bas,
  3 document .cls). Preserves individual module identities; never merges modules and
  never treats document modules as ordinary standard modules.
- Creates timestamped backups for each workbook updated.
- Imports/updates all required standard modules (remove-by-name then Import).
- Synchronizes document-module code safely (Sheet1/Sheet2/ThisWorkbook) by updating
  CodeModule only (no worksheet/workbook component replacement).
- Prevents duplicate modules and removes obsolete legacy component 'xlconcat_macro'.
- Attempts VBA compile/validation (best-effort) on the last workbook processed; reports
  prerequisite if programmatic access to the VBA project object model is not trusted.
- Exports updated vbaProject.bin and individual source exports from the final workbook
  when SyncExports is enabled.
- Produces a clear build/verification report (stdout and file in .ai/tmp/).

.PARAMETER SourceWorkbook
Optional override for the primary workbook path (first in the default pair). When
-NoDefaultBoth is set and no targets are customized, this is the single workbook updated.
Default when using -NoDefaultBoth without other overrides: data\theiss\treport.xlsm.

.PARAMETER ExportDir
Directory containing VBA source exports (.bas/.cls). Default: mko_treport\src\vbs\export.

.PARAMETER WorkCopyPath
When set (and not using default dual in-place), creates/uses a working copy path for the
single-target flow. Ignored in default dual in-place mode.

.PARAMETER BackupDir
Directory for backups and report. Default: .ai\tmp.

.PARAMETER InPlace
In single-target mode, update the target in place (backup created). In default mode
dual targets are always updated in place.

.PARAMETER TargetWorkbook
Single-target mode: workbook to update when -InPlace is used or when -NoDefaultBoth.
Default: same as SourceWorkbook.

.PARAMETER NoDefaultBoth
Update only one workbook (legacy/single behavior) instead of the default pair
(data\theiss\treport.xlsm and mko_treport\src\xls\treport.xlsm).

.PARAMETER AdditionalWorkbook
Extra workbook path(s) to update after the primary targets (in place, backed up). Can be
combined with defaults; duplicates are ignored.

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
.\mko_treport\src\scripts\build_treport_vba.ps1 -NoDefaultBoth -SourceWorkbook mko_treport\src\xls\treport.xlsm -InPlace
#>
[CmdletBinding()]
param(
    [string]$SourceWorkbook,
    [string]$ExportDir,
    [string]$WorkCopyPath,
    [string]$BackupDir,
    [string]$TargetWorkbook,
    [switch]$InPlace,
    [switch]$NoDefaultBoth,
    [string[]]$AdditionalWorkbook,
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

function Apply-TreportVbaToOpenWorkbook {
    param(
        [Parameter(Mandatory = $true)] $VbProj,
        [Parameter(Mandatory = $true)] $ExportDirPath,
        [Parameter(Mandatory = $true)] $StdFilesList,
        [Parameter(Mandatory = $true)] $DocFilesList,
        [Parameter(Mandatory = $true)] $MissingDocList,
        [ref]$ReportRef
    )

    $before = @(Get-ComponentInfo $VbProj)
    Write-Step $ReportRef ("Components before ({0}): {1}" -f $before.Count, ($before -join ', '))

    $imported = @()
    $syncedDocs = @()
    $removedLegacy = @()
    $removedDupes = @()
    $issuesLocal = @()

    # Remove legacy obsolete module if present
    $legacyNames = @('xlconcat_macro')
    foreach ($legacy in $legacyNames) {
        $foundLegacy = @()
        $comps = $VbProj.VBComponents
        for ($i = $comps.Count; $i -ge 1; $i--) {
            $c = $comps.Item($i)
            if ($c.Name -ieq $legacy) { $foundLegacy += $c }
        }
        foreach ($c in $foundLegacy) {
            $t = [int]$c.Type
            if ($t -eq 1 -or $t -eq 2) {
                $VbProj.VBComponents.Remove($c)
                $removedLegacy += $legacy
            } else {
                $issuesLocal += ("Legacy component '{0}' found but not a removable standard/class module (type {1}); left unchanged" -f $legacy, $t)
            }
        }
    }

    foreach ($f in $StdFilesList) {
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $comps = $VbProj.VBComponents
        for ($i = $comps.Count; $i -ge 1; $i--) {
            $c = $comps.Item($i)
            if ($c.Name -ieq $modName) {
                $ct = [int]$c.Type
                if ($ct -eq 1 -or $ct -eq 2 -or $ct -eq 3) {
                    $VbProj.VBComponents.Remove($c)
                } else {
                    $issuesLocal += ("Name collision: '{0}' exists as document/type {1}; cannot replace as standard module" -f $modName, $ct)
                }
            }
        }
        $null = $VbProj.VBComponents.Import($f.FullName)
        $imported += $modName
    }

    foreach ($f in $DocFilesList) {
        $docName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $comp = $null
        try {
            $comp = $VbProj.VBComponents.Item($docName)
        } catch {
            $issuesLocal += ("Document module '{0}' not found in workbook; skipped (component not created/replaced)" -f $docName)
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

    $nameCounts = @{}
    $compsAfterOps = $VbProj.VBComponents
    for ($i = 1; $i -le $compsAfterOps.Count; $i++) {
        $nm = $compsAfterOps.Item($i).Name
        if (-not $nameCounts.ContainsKey($nm)) { $nameCounts[$nm] = 0 }
        $nameCounts[$nm]++
    }
    $dupeNames = @($nameCounts.Keys | Where-Object { $nameCounts[$_] -gt 1 })
    if ($dupeNames.Count -gt 0) {
        foreach ($dn in $dupeNames) {
            $instances = @()
            for ($i = 1; $i -le $VbProj.VBComponents.Count; $i++) { $instances += $VbProj.VBComponents.Item($i) }
            $same = @($instances | Where-Object { $_.Name -ieq $dn })
            if ($same.Count -gt 1) {
                $keepIndex = 0
                for ($k = 0; $k -lt $same.Count; $k++) { if ([int]$same[$k].Type -eq 100) { $keepIndex = $k; break } }
                for ($k = 0; $k -lt $same.Count; $k++) {
                    if ($k -eq $keepIndex) { continue }
                    $ct = [int]$same[$k].Type
                    if ($ct -ne 100) {
                        try { $VbProj.VBComponents.Remove($same[$k]); $removedDupes += ("{0} (extra)" -f $dn) } catch { $issuesLocal += ("Could not remove duplicate '{0}': {1}" -f $dn, $_.Exception.Message) }
                    } else {
                        $issuesLocal += ("Duplicate document component name '{0}' remains (document components cannot be removed)" -f $dn)
                    }
                }
            }
        }
    }

    $compileStatusLocal = 'not attempted'
    $compileOkLocal = $false
    try {
        $VbProj.GetType().InvokeMember('Compile', [System.Reflection.BindingFlags]::InvokeMethod, $null, $VbProj, $null) | Out-Null
        $compileStatusLocal = 'compiled (Compile invoked successfully)'
        $compileOkLocal = $true
    } catch {
        $compileStatusLocal = ("compile not confirmed: {0}" -f $_.Exception.Message)
        $issuesLocal += "VBA compile could not be confirmed programmatically (may be unsupported in this Excel/COM context); manual compile check recommended in VBE (Debug > Compile)."
    }

    $after = @(Get-ComponentInfo $VbProj)
    $finalNamesList = @()
    foreach ($entry in $after) { if ($entry -match '^(.+?)\s+\(') { $finalNamesList += $Matches[1] } }

    $expectedStdNames = @($StdFilesList | ForEach-Object { $_.BaseName })
    $expectedDocNames = @('Sheet1', 'Sheet2', 'ThisWorkbook')
    $missingExpected = @()
    foreach ($e in ($expectedStdNames + $expectedDocNames)) {
        if ($finalNamesList -notcontains $e) { $missingExpected += $e }
    }
    $hasLegacy = ($finalNamesList -contains 'xlconcat_macro')
    $dupeFinal = @($nameCounts.Keys | Where-Object { $nameCounts[$_] -gt 1 })

    return [pscustomobject]@{
        Imported = @($imported | Sort-Object -Unique)
        SyncedDocs = @($syncedDocs | Sort-Object -Unique)
        RemovedLegacy = @($removedLegacy | Sort-Object -Unique)
        RemovedDupes = @($removedDupes)
        Issues = @($issuesLocal)
        CompileStatus = $compileStatusLocal
        CompileOk = $compileOkLocal
        AfterNamesJoined = ($after -join ', ')
        AfterCount = $after.Count
        MissingExpected = @($missingExpected)
        HasLegacy = [bool]$hasLegacy
        DupeFinal = @($dupeFinal)
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

    if ([string]::IsNullOrWhiteSpace($ExportDir))      { $ExportDir      = Join-Path $repoRoot 'mko_treport\src\vbs\export' }
    if ([string]::IsNullOrWhiteSpace($BackupDir))      { $BackupDir      = Join-Path $repoRoot '.ai\tmp' }

    $SyncExports = -not $NoSyncExports.IsPresent
    $ts = Get-Timestamp
    $stamp = Get-Stamp

    if (-not (Test-Path -LiteralPath $ExportDir))      { throw "Export directory not found: $ExportDir" }
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

    $stdFiles = @(Get-ChildItem -LiteralPath $ExportDir -Filter '*.bas' -File | Sort-Object Name)
    $docFiles = @(Get-ChildItem -LiteralPath $ExportDir -Filter '*.cls' -File | Sort-Object Name)
    $docNames = @('Sheet1', 'Sheet2', 'ThisWorkbook')
    $missingDoc = @()
    foreach ($dn in $docNames) {
        if (-not ($docFiles | Where-Object { $_.BaseName -eq $dn })) { $missingDoc += $dn }
    }

    $exportResolved = (Resolve-Path -LiteralPath $ExportDir).Path
    $backupResolved = (Resolve-Path -LiteralPath $BackupDir).Path

    $workbookTargets = New-Object System.Collections.ArrayList

    if ($NoDefaultBoth) {
        if ([string]::IsNullOrWhiteSpace($SourceWorkbook)) { $SourceWorkbook = Join-Path $repoRoot 'data\theiss\treport.xlsm' }
        if ([string]::IsNullOrWhiteSpace($TargetWorkbook)) { $TargetWorkbook = $SourceWorkbook }
        $t = [System.IO.Path]::GetFullPath($TargetWorkbook)
        if (-not (Test-Path -LiteralPath $t)) { throw "Target workbook not found: $t" }
        [void]$workbookTargets.Add($t)
    } else {
        $defaultTargets = @(
            (Join-Path $repoRoot 'data\theiss\treport.xlsm'),
            (Join-Path $repoRoot 'mko_treport\src\xls\treport.xlsm')
        )
        if (-not [string]::IsNullOrWhiteSpace($SourceWorkbook)) {
            $so = [System.IO.Path]::GetFullPath($SourceWorkbook)
            if (-not ($defaultTargets | Where-Object { $_ -ieq $so })) { $defaultTargets = @($so) + $defaultTargets }
        }
        if (-not [string]::IsNullOrWhiteSpace($TargetWorkbook)) {
            $to = [System.IO.Path]::GetFullPath($TargetWorkbook)
            if (-not ($defaultTargets | Where-Object { $_ -ieq $to })) { $defaultTargets = @($to) + $defaultTargets }
        }
        foreach ($dt in $defaultTargets) {
            $rp = [System.IO.Path]::GetFullPath($dt)
            if (-not (Test-Path -LiteralPath $rp)) { throw "Target workbook not found: $rp" }
            $already = $false
            foreach ($ex in $workbookTargets) { if ($ex -ieq $rp) { $already = $true; break } }
            if (-not $already) { [void]$workbookTargets.Add($rp) }
        }
    }

    if ($AdditionalWorkbook) {
        foreach ($aw in $AdditionalWorkbook) {
            if ([string]::IsNullOrWhiteSpace($aw)) { continue }
            $rp = [System.IO.Path]::GetFullPath($aw)
            if (-not (Test-Path -LiteralPath $rp)) { throw "Additional workbook not found: $rp" }
            $already = $false
            foreach ($ex in $workbookTargets) { if ($ex -ieq $rp) { $already = $true; break } }
            if (-not $already) { [void]$workbookTargets.Add($rp) }
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
    Write-Step ([ref]$report) ("Export dir: {0}" -f $exportResolved)
    Write-Step ([ref]$report) ("Backup dir: {0}" -f $backupResolved)
    Write-Step ([ref]$report) ("Mode: {0}" -f ($(if ($NoDefaultBoth) { 'single (NoDefaultBoth)' } else { 'default dual in-place: data\\theiss\\treport.xlsm + mko_treport\\src\\xls\\treport.xlsm (+ any overrides/additions)' })))
    Write-Step ([ref]$report) ("SyncExports: {0}" -f $SyncExports)
    Write-Step ([ref]$report) ("Standard modules found ({0}): {1}" -f $stdFiles.Count, (($stdFiles | ForEach-Object { $_.Name }) -join ', '))
    Write-Step ([ref]$report) ("Document modules found ({0}): {1}" -f $docFiles.Count, (($docFiles | ForEach-Object { $_.Name }) -join ', '))
    if ($missingDoc.Count -gt 0) {
        Write-Step ([ref]$report) ("WARNING: missing expected document module sources: {0}" -f ($missingDoc -join ', '))
    }
    Write-Step ([ref]$report) ("Workbooks to update ({0}):" -f $workbookTargets.Count)
    foreach ($wt in $workbookTargets) { Write-Step ([ref]$report) ("  - {0}" -f $wt) }
    if (-not [string]::IsNullOrWhiteSpace($WorkCopyPath) -and $NoDefaultBoth) {
        Write-Step ([ref]$report) "NOTE: -WorkCopyPath is ignored in default dual mode; use -NoDefaultBoth for single work-copy workflows."
    }

    $targetResults = @()
    $lastWorkbookPath = $null
    $vbaBinExported = $false
    $vbaBinPath = Join-Path $repoRoot 'mko_treport\src\vbs\vbaProject.bin'

    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw ("Failed to start Excel.Application COM automation. Is Excel installed? Details: {0}" -f $_.Exception.Message)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.EnableEvents = $false } catch { }
    try { $excel.AskToUpdateLinks = $false } catch { }

    $idx = 0
    foreach ($wt in $workbookTargets) {
        $idx++
        $wbPathResolved = [System.IO.Path]::GetFullPath($wt)
        $lastWorkbookPath = $wbPathResolved
        $backupPath = Join-Path $backupResolved ("treport_backup_{0}_{1:D2}_{2}.xlsm" -f $stamp, $idx, [System.IO.Path]::GetFileNameWithoutExtension($wbPathResolved))
        Copy-Item -LiteralPath $wbPathResolved -Destination $backupPath -Force
        Write-Step ([ref]$report) ("")
        Write-Step ([ref]$report) ("--- Target {0}/{1}: {2} ---" -f $idx, $workbookTargets.Count, $wbPathResolved)
        Write-Step ([ref]$report) ("Backup created: {0}" -f $backupPath)

        $workbook = $null
        $vbProj = $null
        try {
            $workbook = $excel.Workbooks.Open($wbPathResolved)
        } catch {
            throw ("Failed to open workbook '{0}': {1}" -f $wbPathResolved, $_.Exception.Message)
        }
        try {
            $vbProj = $workbook.VBProject
        } catch {
            try { $workbook.Close($false) } catch { }
            $workbook = $null
            throw ("Cannot access VBA project object model for '{0}'. Enable 'Trust access to the VBA project object model' in Excel Trust Center (Macro Settings). Details: {1}" -f $wbPathResolved, $_.Exception.Message)
        }

        $applyRes = Apply-TreportVbaToOpenWorkbook -VbProj $vbProj -ExportDirPath $exportResolved -StdFilesList $stdFiles -DocFilesList $docFiles -MissingDocList $missingDoc -ReportRef ([ref]$report)

        $workbook.Save()
        $runBinding = Try-GetRunBinding $workbook
        Write-Step ([ref]$report) ("Workbook saved: {0}" -f $wbPathResolved)
        Write-Step ([ref]$report) ("Run control binding (best-effort): {0}" -f $runBinding)

        $targetResults += [pscustomobject]@{
            Path = $wbPathResolved
            Backup = $backupPath
            Apply = $applyRes
        }

        try { $workbook.Close($false) } catch { }
        $workbook = $null
        $vbProj = $null
    }

    # Export vbaProject.bin and managed exports from the last workbook (primary sync source)
    if ($SyncExports -and $lastWorkbookPath) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            $zip = [System.IO.Compression.ZipFile]::OpenRead($lastWorkbookPath)
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
                Write-Step ([ref]$report) ("vbaProject.bin exported (from last target): {0}" -f $vbaBinPath)
            } else {
                Write-Step ([ref]$report) "WARNING: xl/vbaProject.bin not found inside last workbook package; bin export skipped."
            }
        } catch {
            Write-Step ([ref]$report) ("WARNING: vbaProject.bin export failed: {0}" -f $_.Exception.Message)
        } finally {
            if ($null -ne $zipArchive) { try { $zipArchive.Dispose() } catch { }; $zipArchive = $null }
        }

        try {
            $wbEx = $excel.Workbooks.Open($lastWorkbookPath)
            try {
                $vpEx = $wbEx.VBProject
                $compsFinalForExport = $vpEx.VBComponents
                $exportedBack = @()
                for ($i = 1; $i -le $compsFinalForExport.Count; $i++) {
                    $c = $compsFinalForExport.Item($i)
                    $nm = $c.Name
                    $ct = [int]$c.Type
                    $ext = $null
                    if ($ct -eq 1) { $ext = '.bas' }
                    elseif ($ct -eq 2) { $ext = '.cls' }
                    elseif ($ct -eq 100) { $ext = '.cls' }
                    if ($null -ne $ext) {
                        $srcFile = Join-Path $exportResolved ($nm + $ext)
                        if (Test-Path -LiteralPath $srcFile) {
                            try { $c.Export($srcFile); $exportedBack += ($nm + $ext) } catch { Write-Step ([ref]$report) ("WARNING: Failed to export component {0}: {1}" -f $nm, $_.Exception.Message) }
                        }
                    }
                }
                Write-Step ([ref]$report) ("Source exports refreshed from last target ({0}): {1}" -f $exportedBack.Count, ($exportedBack -join ', '))
            } finally {
                try { $wbEx.Close($false) } catch { }
            }
        } catch {
            Write-Step ([ref]$report) ("WARNING: Could not refresh source exports from last target: {0}" -f $_.Exception.Message)
        }
    } elseif (-not $SyncExports) {
        Write-Step ([ref]$report) "SyncExports disabled: vbaProject.bin and source exports not updated."
    }

    $report += ""
    $report += "--- Per-target summary ---"
    $overallOkAll = $true
    $allIssues = @()
    foreach ($tr in $targetResults) {
        $a = $tr.Apply
        $report += ("Target: {0}" -f $tr.Path)
        $report += ("  Backup: {0}" -f $tr.Backup)
        $report += ("  Modules imported (standard .bas): {0}" -f (($a.Imported -join ', ')))
        $report += ("  Document modules synced: {0}" -f (($a.SyncedDocs -join ', ')))
        $report += ("  Legacy removed: {0}" -f ($(if ($a.RemovedLegacy.Count -gt 0) { ($a.RemovedLegacy -join ', ') } else { 'none' })))
        $report += ("  Duplicates removed: {0}" -f ($(if ($a.RemovedDupes.Count -gt 0) { ($a.RemovedDupes -join ', ') } else { 'none' })))
        $report += ("  Compile/validation: {0}" -f $a.CompileStatus)
        $report += ("  Final components ({0}): {1}" -f $a.AfterCount, $a.AfterNamesJoined)
        $report += ("  Expected components present: {0}" -f ($(if ($a.MissingExpected.Count -eq 0) { 'yes' } else { ('NO, missing: ' + ($a.MissingExpected -join ', ')) })))
        $report += ("  Legacy xlconcat_macro present: {0}" -f ($(if ($a.HasLegacy) { 'YES (unexpected)' } else { 'no' })))
        $report += ("  Duplicate component names after build: {0}" -f ($(if ($a.DupeFinal.Count -eq 0) { 'none' } else { ($a.DupeFinal -join ', ') })))
        if ($a.Issues.Count -gt 0) {
            $report += "  Issues:"
            foreach ($iss in $a.Issues) { $report += ("    - " + $iss); $allIssues += ("[{0}] {1}" -f $tr.Path, $iss) }
        }
        if ($a.MissingExpected.Count -gt 0 -or $a.HasLegacy -or $a.DupeFinal.Count -gt 0) { $overallOkAll = $false }
    }

    $report += ""
    $report += "--- Global ---"
    $report += ("vbaProject.bin exported (from last target): {0}" -f ($(if ($vbaBinExported) { $vbaBinPath } elseif ($SyncExports) { 'no' } else { 'skipped (NoSyncExports)' })))

    if ($allIssues.Count -gt 0) {
        $report += ""
        $report += "--- Issues / Prerequisites (aggregated) ---"
        foreach ($iss in $allIssues) { $report += ("- " + $iss) }
    } else {
        $report += ""
        $report += "--- Issues / Prerequisites ---"
        $report += "- none"
    }

    $report += ""
    $report += ("RESULT: {0}" -f ($(if ($overallOkAll) { 'SUCCESS' } else { 'COMPLETED WITH ISSUES' })))
    $report += ("Report file: {0}" -f $ReportPath)

    $reportText = ($report -join "`r`n")
    $reportDir = Split-Path -Parent $ReportPath
    if ($reportDir -and -not (Test-Path -LiteralPath $reportDir)) { New-Item -ItemType Directory -Force -Path $reportDir | Out-Null }
    [System.IO.File]::WriteAllText($ReportPath, $reportText, [System.Text.Encoding]::UTF8)

    if (-not $overallOkAll) { $exitCode = 1 }
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

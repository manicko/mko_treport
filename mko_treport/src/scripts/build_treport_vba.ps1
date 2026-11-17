<# 
.SYNOPSIS
One-step VBA build/import utility for treport.xlsm.

.DESCRIPTION
Assembles the current VBA project from modular sources in mko_treport/src/vbs/export/
into the tracked treport workbook. Auxiliary dev/deployment tooling only (not a runtime
dependency).

Default behavior (safe, no sync):
- Builds VBA modules from mko_treport/src/vbs/export into mko_treport/src/xls/treport.xlsm
  IN PLACE (after creating a timestamped backup in BackupDir).
- Does NOT update mko_treport/src/vbs/export/*.bas|*.cls and does NOT update
  mko_treport/src/vbs/vbaProject.bin unless -SyncExports is specified explicitly.
- Does NOT touch data\theiss\treport.xlsm unless -DeployToDataTheiss is specified explicitly.

Build/apply behavior:
- Reads managed VBA source modules from ExportDir (.bas and .cls). Preserves individual
  module identities; never merges modules and never treats document modules as ordinary
  standard modules.
- Creates a timestamped backup of the target workbook before modification.
- Rebuilds managed standard/class/userform components as non-document modules by removing
  them and re-adding standard modules from .bas exports (CodeModule body injection from
  stripped UTF-8 source; avoids VBComponents.Import for std modules).
- Synchronizes document-module code safely (Sheet1/Sheet2/ThisWorkbook) by updating
  CodeModule only (no worksheet/workbook component replacement).
- Prevents duplicate modules and removes obsolete legacy component 'xlconcat_macro'.
- Attempts VBA compile/validation (best-effort) on the built workbook; reports prerequisite
  if programmatic access to the VBA project object model is not trusted.

Sync behavior (-SyncExports, explicit opt-in only):
- After a successful build/save, refreshes managed .bas/.cls files in ExportDir via
  VBComponent.Export() from the built workbook and extracts xl/vbaProject.bin to
  mko_treport/src/vbs/vbaProject.bin.
- Export is performed from the single built target workbook (unambiguous source).
- Note: VBE Export may normalize formatting/line endings/case and does not apply
  post-export encoding normalization here; vbaProject.bin is a volatile binary artifact.
  Use sync intentionally (typically after successful manual VBE compile/L2).

Deploy behavior (-DeployToDataTheiss, explicit opt-in only):
- After a successful build/save (and optional sync), replaces data\theiss\treport.xlsm
  with a copy of the built mko_treport\src\xls\treport.xlsm (removes existing data copy
  if present, creates directory if needed). The data copy is not used as a sync source.

.PARAMETER TargetWorkbook
Workbook to build in place. Default: mko_treport\src\xls\treport.xlsm (tracked working workbook).
A timestamped backup is always created before modification.

.PARAMETER ExportDir
Directory containing VBA source exports (.bas/.cls). Default: mko_treport\src\vbs\export.

.PARAMETER BackupDir
Directory for backups and report. Default: .ai\tmp.

.PARAMETER SyncExports
Explicit opt-in: export vbaProject.bin to mko_treport\src\vbs\vbaProject.bin and refresh
managed .bas/.cls files in ExportDir from the built workbook after a successful save.
Default: disabled.

.PARAMETER DeployToDataTheiss
Explicit opt-in: after successful build/save, delete data\theiss\treport.xlsm if present and
copy the built mko_treport\src\xls\treport.xlsm to data\theiss\treport.xlsm.
Default: disabled.

.PARAMETER ReportPath
Explicit report file path. Default: <BackupDir>\build_treport_vba_<ts>.txt.

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1
# Default: build VBA into tracked mko_treport\src\xls\treport.xlsm (backup in .ai\tmp), no sync, no deploy.

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1 -SyncExports
# Build and then explicitly refresh exports/bin from the built workbook.

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1 -DeployToDataTheiss
# Build tracked workbook and then replace data\theiss\treport.xlsm with a copy of it.

.EXAMPLE
.\mko_treport\src\scripts\build_treport_vba.ps1 -TargetWorkbook .ai\tmp\treport_check.xlsm
# Build into an explicit workbook copy/path (backup/report still in BackupDir by default).
#>
[CmdletBinding()]
param(
    [string]$TargetWorkbook,
    [string]$ExportDir,
    [string]$BackupDir,
    [switch]$SyncExports,
    [switch]$DeployToDataTheiss,
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

    # Remove ALL non-document modules (Std/Class/UserForm) for a clean rebuild.
    # Document modules (Sheet/Workbook, type 100) cannot be removed and are synced by CodeModule update.
    # This prevents stale/corrupted leftovers (e.g. Module1 with BOM artifacts) and duplicates.
    $compsToRemove = @()
    for ($i = 1; $i -le $VbProj.VBComponents.Count; $i++) {
        $c = $VbProj.VBComponents.Item($i)
        if ([int]$c.Type -ne 100) { $compsToRemove += $c }
    }
    foreach ($c in $compsToRemove) {
        $nm = [string]$c.Name
        $ct = [int]$c.Type
        try {
            $VbProj.VBComponents.Remove($c)
            if ($nm -ieq 'xlconcat_macro') { $removedLegacy += $nm }
            else { $removedDupes += ("{0} (removed-prebuild)" -f $nm) }
        } catch {
            $issuesLocal += ("Could not remove module '{0}' (type {1}) before rebuild: {2}" -f $nm, $ct, $_.Exception.Message)
        }
    }

    # Re-add managed standard modules from export (.bas) as fresh components and inject
    # stripped CodeModule bodies (avoids .Import carrying binary/BOM artifacts).
    foreach ($f in $StdFilesList) {
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        try {
            $newComp = $VbProj.VBComponents.Add(1) # vbext_ct_StdModule
            $newComp.Name = $modName
            $rawStd = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            $bodyStd = Strip-VbaHeader $rawStd
            $cmStd = $newComp.CodeModule
            $lcStd = [int]$cmStd.CountOfLines
            if ($lcStd -gt 0) { $cmStd.DeleteLines(1, $lcStd) }
            if (-not [string]::IsNullOrEmpty($bodyStd)) { $cmStd.AddFromString($bodyStd) }
            $imported += $modName
        } catch {
            $issuesLocal += ("Failed to add standard module '{0}': {1}" -f $modName, $_.Exception.Message)
        }
    }

    # Synchronize document modules safely (Sheet1/Sheet2/ThisWorkbook) by updating CodeModule only.
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
            try {
                $raw = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
                $body = Strip-VbaHeader $raw
                $cm = $comp.CodeModule
                $lineCount = [int]$cm.CountOfLines
                if ($lineCount -gt 0) { $cm.DeleteLines(1, $lineCount) }
                if (-not [string]::IsNullOrEmpty($body)) { $cm.AddFromString($body) }
                $syncedDocs += $docName
            } catch {
                $issuesLocal += ("Failed to sync document module '{0}': {1}" -f $docName, $_.Exception.Message)
            }
        }
    }

    # Post-check: duplicates (defensive; should not occur after full non-document rebuild).
    $nameCounts = @{}
    for ($i = 1; $i -le $VbProj.VBComponents.Count; $i++) {
        $nm = $VbProj.VBComponents.Item($i).Name
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

    if ([string]::IsNullOrWhiteSpace($ExportDir)) { $ExportDir = Join-Path $repoRoot 'mko_treport\src\vbs\export' }
    if ([string]::IsNullOrWhiteSpace($BackupDir)) { $BackupDir = Join-Path $repoRoot '.ai\tmp' }
    if ([string]::IsNullOrWhiteSpace($TargetWorkbook)) { $TargetWorkbook = Join-Path $repoRoot 'mko_treport\src\xls\treport.xlsm' }

    $ts = Get-Timestamp
    $stamp = Get-Stamp

    if (-not (Test-Path -LiteralPath $ExportDir)) { throw "Export directory not found: $ExportDir" }
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

    $targetResolved = [System.IO.Path]::GetFullPath($TargetWorkbook)
    $defaultTrackedTarget = [System.IO.Path]::GetFullPath((Join-Path $repoRoot 'mko_treport\src\xls\treport.xlsm'))
    $isDefaultTrackedTarget = ($targetResolved -ieq $defaultTrackedTarget)

    if (-not (Test-Path -LiteralPath $targetResolved)) { throw "Target workbook not found: $targetResolved" }

    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path $backupResolved ("build_treport_vba_{0}.txt" -f $stamp)
    } else {
        $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    }

    $vbaBinPath = Join-Path $repoRoot 'mko_treport\src\vbs\vbaProject.bin'
    $dataDeployTarget = Join-Path $repoRoot 'data\theiss\treport.xlsm'

    Write-Step ([ref]$report) "=== build_treport_vba ==="
    Write-Step ([ref]$report) ("Timestamp: {0}" -f $ts)
    Write-Step ([ref]$report) ("Repo root: {0}" -f $repoRoot)
    Write-Step ([ref]$report) ("Export dir: {0}" -f $exportResolved)
    Write-Step ([ref]$report) ("Backup dir: {0}" -f $backupResolved)
    Write-Step ([ref]$report) ("Target workbook (in-place): {0}" -f $targetResolved)
    Write-Step ([ref]$report) ("Default tracked target: {0} ({1})" -f $defaultTrackedTarget, ($(if ($isDefaultTrackedTarget) { 'yes' } else { 'no' })))
    Write-Step ([ref]$report) ("SyncExports (explicit): {0}" -f $SyncExports.IsPresent)
    Write-Step ([ref]$report) ("DeployToDataTheiss (explicit): {0}" -f $DeployToDataTheiss.IsPresent)
    Write-Step ([ref]$report) ("Standard modules found ({0}): {1}" -f $stdFiles.Count, (($stdFiles | ForEach-Object { $_.Name }) -join ', '))
    Write-Step ([ref]$report) ("Document modules found ({0}): {1}" -f $docFiles.Count, (($docFiles | ForEach-Object { $_.Name }) -join ', '))
    if ($missingDoc.Count -gt 0) {
        Write-Step ([ref]$report) ("WARNING: missing expected document module sources: {0}" -f ($missingDoc -join ', '))
    }

    $backupPath = Join-Path $backupResolved ("treport_backup_{0}_{1}.xlsm" -f $stamp, [System.IO.Path]::GetFileNameWithoutExtension($targetResolved))
    Copy-Item -LiteralPath $targetResolved -Destination $backupPath -Force
    Write-Step ([ref]$report) ("Backup created: {0}" -f $backupPath)

    $targetResult = $null
    $vbaBinExported = $false

    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw ("Failed to start Excel.Application COM automation. Is Excel installed? Details: {0}" -f $_.Exception.Message)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.EnableEvents = $false } catch { }
    try { $excel.AskToUpdateLinks = $false } catch { }

    Write-Step ([ref]$report) ""
    Write-Step ([ref]$report) ("--- Build target: {0} ---" -f $targetResolved)

    try {
        $workbook = $excel.Workbooks.Open($targetResolved)
    } catch {
        throw ("Failed to open workbook '{0}': {1}" -f $targetResolved, $_.Exception.Message)
    }
    try {
        $vbProj = $workbook.VBProject
    } catch {
        try { $workbook.Close($false) } catch { }
        $workbook = $null
        throw ("Cannot access VBA project object model for '{0}'. Enable 'Trust access to the VBA project object model' in Excel Trust Center (Macro Settings). Details: {1}" -f $targetResolved, $_.Exception.Message)
    }

    $applyRes = Apply-TreportVbaToOpenWorkbook -VbProj $vbProj -ExportDirPath $exportResolved -StdFilesList $stdFiles -DocFilesList $docFiles -MissingDocList $missingDoc -ReportRef ([ref]$report)

    $workbook.Save()
    $runBinding = Try-GetRunBinding $workbook
    Write-Step ([ref]$report) ("Workbook saved: {0}" -f $targetResolved)
    Write-Step ([ref]$report) ("Run control binding (best-effort): {0}" -f $runBinding)

    $targetResult = [pscustomobject]@{
        Path = $targetResolved
        Backup = $backupPath
        Apply = $applyRes
    }

    try { $workbook.Close($false) } catch { }
    $workbook = $null
    $vbProj = $null

    # Optional explicit sync: exports/bin from the single built workbook (unambiguous source)
    if ($SyncExports) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
            $zip = [System.IO.Compression.ZipFile]::OpenRead($targetResolved)
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
                Write-Step ([ref]$report) ("vbaProject.bin exported (from built target): {0}" -f $vbaBinPath)
            } else {
                Write-Step ([ref]$report) "WARNING: xl/vbaProject.bin not found inside built workbook package; bin export skipped."
            }
        } catch {
            Write-Step ([ref]$report) ("WARNING: vbaProject.bin export failed: {0}" -f $_.Exception.Message)
        } finally {
            if ($null -ne $zipArchive) { try { $zipArchive.Dispose() } catch { }; $zipArchive = $null }
        }

        try {
            $wbEx = $excel.Workbooks.Open($targetResolved)
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
                Write-Step ([ref]$report) ("Source exports refreshed from built target ({0}): {1}" -f $exportedBack.Count, ($exportedBack -join ', '))
                Write-Step ([ref]$report) "NOTE: VBComponent.Export output is not post-normalized by this script (format/EOL/encoding may vary by VBE environment)."
            } finally {
                try { $wbEx.Close($false) } catch { }
            }
        } catch {
            Write-Step ([ref]$report) ("WARNING: Could not refresh source exports from built target: {0}" -f $_.Exception.Message)
        }
    } else {
        Write-Step ([ref]$report) "SyncExports: disabled (default). mko_treport/src/vbs/export and mko_treport/src/vbs/vbaProject.bin were not updated."
    }

    # Optional explicit deploy: replace data\theiss\treport.xlsm with a copy of tracked working workbook
    if ($DeployToDataTheiss) {
        $deploySource = $defaultTrackedTarget
        if (-not (Test-Path -LiteralPath $deploySource)) {
            throw ("DeployToDataTheiss requested, but tracked workbook not found: {0}" -f $deploySource)
        }
        $dataDir = Split-Path -Parent $dataDeployTarget
        if (-not (Test-Path -LiteralPath $dataDir)) { New-Item -ItemType Directory -Force -Path $dataDir | Out-Null }
        if (Test-Path -LiteralPath $dataDeployTarget) { Remove-Item -LiteralPath $dataDeployTarget -Force }
        Copy-Item -LiteralPath $deploySource -Destination $dataDeployTarget -Force
        Write-Step ([ref]$report) ("DeployToDataTheiss: replaced '{0}' with a copy of '{1}'" -f $dataDeployTarget, $deploySource)
        Write-Step ([ref]$report) "NOTE: data copy is a derived/runtime copy and is not used as a sync source."
    } else {
        Write-Step ([ref]$report) "DeployToDataTheiss: disabled (default). data\\theiss\\treport.xlsm was not modified."
    }

    $report += ""
    $report += "--- Summary ---"
    $a = $targetResult.Apply
    $report += ("Target: {0}" -f $targetResult.Path)
    $report += ("Backup: {0}" -f $targetResult.Backup)
    $report += ("Modules imported (standard .bas): {0}" -f (($a.Imported -join ', ')))
    $report += ("Document modules synced: {0}" -f (($a.SyncedDocs -join ', ')))
    $report += ("Legacy removed: {0}" -f ($(if ($a.RemovedLegacy.Count -gt 0) { ($a.RemovedLegacy -join ', ') } else { 'none' })))
    $report += ("Duplicates removed: {0}" -f ($(if ($a.RemovedDupes.Count -gt 0) { ($a.RemovedDupes -join ', ') } else { 'none' })))
    $report += ("Compile/validation: {0}" -f $a.CompileStatus)
    $report += ("Final components ({0}): {1}" -f $a.AfterCount, $a.AfterNamesJoined)
    $report += ("Expected components present: {0}" -f ($(if ($a.MissingExpected.Count -eq 0) { 'yes' } else { ('NO, missing: ' + ($a.MissingExpected -join ', ')) })))
    $report += ("Legacy xlconcat_macro present: {0}" -f ($(if ($a.HasLegacy) { 'YES (unexpected)' } else { 'no' })))
    $report += ("Duplicate component names after build: {0}" -f ($(if ($a.DupeFinal.Count -eq 0) { 'none' } else { ($a.DupeFinal -join ', ') })))
    $report += ("vbaProject.bin exported (from built target): {0}" -f ($(if ($vbaBinExported) { $vbaBinPath } elseif ($SyncExports) { 'no' } else { 'skipped (SyncExports disabled)' })))
    $report += ("DeployToDataTheiss: {0}" -f ($(if ($DeployToDataTheiss) { 'completed' } else { 'skipped (disabled)' })))

    $overallOk = $true
    if ($a.Issues.Count -gt 0) {
        $report += ""
        $report += "--- Issues / Prerequisites ---"
        foreach ($iss in $a.Issues) { $report += ("- " + $iss) }
    } else {
        $report += ""
        $report += "--- Issues / Prerequisites ---"
        $report += "- none"
    }

    if ($a.MissingExpected.Count -gt 0 -or $a.HasLegacy -or $a.DupeFinal.Count -gt 0) { $overallOk = $false }

    $report += ""
    $report += ("RESULT: {0}" -f ($(if ($overallOk) { 'SUCCESS' } else { 'COMPLETED WITH ISSUES' })))
    $report += ("Report file: {0}" -f $ReportPath)

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

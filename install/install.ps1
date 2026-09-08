# install.ps1
#
# Version 1.0 / 2026-09-08
# Version 1.1 / 2026-09-08: skip template copy when the book already sits next to
#   tmp_index.html in the project root (Copy-Item cannot overwrite itself).
# Version 1.3 / 2026-09-08: backups go to a dedicated bak\ folder instead of next
#   to the workbook (project rule docs/rules.md: never beside the working book and
#   never into build\, build is handed over as a whole). Books inside ProjectRoot
#   -> <ProjectRoot>\bak, books outside -> bak\ next to the book.
# Version 1.2 / 2026-09-08: also refresh Power Query queries from src\powerquery.
#   Before v1.2 the patch touched VBA only, so changed M-code (fnNormalizeFields v6)
#   never reached the production book and report generation kept failing. Rules:
#   existing queries are updated when the formula differs from the .pq file;
#   missing pipeline queries (fn*, Query-ImportJSON, qExistingData) are ADDED;
#   qDiagImport is never added (diagnostics only); prmSourcePath (runtime state,
#   holds the path picked by the user) is never touched.
# Applies the v7.1 patch to existing ReportMTO workbooks in place (Path A):
#   - replaces VBA modules modMain and modContentMTO (Remove + Import,
#     source re-encoded UTF-8 -> ANSI 1251, as VBE reads .bas only as ANSI);
#   - copies the v2.0 template tmp_index.html next to each patched workbook;
#   - makes timestamped backups of the workbook and the old template first.
# Data (tbDATA), Variable sheet, buttons and Power Query connections are untouched.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1
#       -> patches every *.xlsm in .\build (default target)
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target .\path.xlsm
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target .\some\folder
#
# Script is ASCII-only on purpose: PowerShell 5.1 reads .ps1 without BOM as ANSI.
# All cyrillic strings live in the VBA sources, never in this script.

param(
    [string]$Target = "build",
    [string]$ProjectRoot = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($ProjectRoot -eq "") { $ProjectRoot = Split-Path -Parent $scriptRoot }
if (-not [System.IO.Path]::IsPathRooted($Target)) {
    $Target = Join-Path $ProjectRoot $Target
} else {
    $Target = [System.IO.Path]::GetFullPath($Target)
}

$vbaDir = Join-Path $ProjectRoot "src\vba"
$tplSrc = Join-Path $ProjectRoot "tmp_index.html"

if (-not (Test-Path $tplSrc)) { Write-Output "NO_TEMPLATE: $tplSrc"; exit 1 }

# Resolve target to a list of workbooks: folder -> all *.xlsm inside, file -> itself.
$books = @()
$targetItem = Get-Item $Target -ErrorAction SilentlyContinue
if ($null -ne $targetItem) {
    if ($targetItem.PSIsContainer) {
        $books = @(Get-ChildItem -Path $Target -Filter *.xlsm -File)
    } else {
        $books = @($targetItem)
    }
}
if ($null -eq $targetItem) { Write-Output "NOT_FOUND: $Target"; exit 1 }
if ($books.Count -eq 0) { Write-Output "NO_XLSM: $Target"; exit 1 }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ansi = [System.Text.Encoding]::GetEncoding(1251)
$tempDir = $env:TEMP
$patched = 0

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($bookItem in $books) {
    $bookPath = $bookItem.FullName
    try {
        $bookDir = Split-Path -Parent $bookPath

        # Backup workbook and the old template (rollback = restore both). Backups go
        # to a dedicated bak\ folder (project rule): never next to the working book
        # and never into build\. Books inside ProjectRoot -> <ProjectRoot>\bak,
        # books outside -> bak\ next to the book.
        $bakDir = Join-Path $ProjectRoot "bak"
        $rootFull = [System.IO.Path]::GetFullPath($ProjectRoot)
        $bookFull = [System.IO.Path]::GetFullPath($bookPath)
        if (-not $bookFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $bakDir = Join-Path $bookDir "bak"
        }
        New-Item -ItemType Directory -Force -Path $bakDir | Out-Null
        Copy-Item $bookPath (Join-Path $bakDir ($bookItem.Name + ".bak_" + $stamp)) -Force
        $oldTpl = Join-Path $bookDir "tmp_index.html"
        if (Test-Path $oldTpl) { Copy-Item $oldTpl (Join-Path $bakDir ("tmp_index.html.bak_" + $stamp)) -Force }

        $wb = $excel.Workbooks.Open($bookPath, 0, $false)
        $vp = $wb.VBProject
        foreach ($mod in @("modMain", "modContentMTO")) {
            try { $vp.VBComponents.Remove($vp.VBComponents.Item($mod)) } catch { }
            $src = Join-Path $vbaDir ($mod + ".bas")
            $utf8 = [System.IO.File]::ReadAllText($src)
            $ansiTemp = Join-Path $tempDir ($mod + "_install_ansi.bas")
            [System.IO.File]::WriteAllText($ansiTemp, $utf8, $ansi)
            $vp.VBComponents.Import($ansiTemp) | Out-Null
        }

        # Refresh Power Query queries from src\powerquery (v1.2).
        $pqDir = Join-Path $ProjectRoot "src\powerquery"
        foreach ($pqFile in Get-ChildItem -Path $pqDir -Filter *.pq) {
            $pqName = [System.IO.Path]::GetFileNameWithoutExtension($pqFile.Name)
            if ($pqName -eq "qDiagImport") { continue }
            $pqFormula = [System.IO.File]::ReadAllText($pqFile.FullName)
            $qFound = $false
            foreach ($q in $wb.Queries) { if ($q.Name -eq $pqName) { $qFound = $true } }
            if ($qFound) {
                if ($wb.Queries.Item($pqName).Formula -ne $pqFormula) {
                    $wb.Queries.Item($pqName).Formula = $pqFormula
                    Write-Output "PQ_UPDATED $pqName"
                }
            } else {
                $wb.Queries.Add($pqName, $pqFormula) | Out-Null
                Write-Output "PQ_ADDED $pqName"
            }
        }

        $wb.Save()
        $wb.Close($false)

        # Template next to the workbook (GenerateReport reads it by ThisWorkbook.Path).
        # Skip when source and destination are the same file (book in project root).
        if ($oldTpl -ne $tplSrc) { Copy-Item $tplSrc $oldTpl -Force }

        $patched += 1
        Write-Output "PATCHED $bookPath"
    } catch {
        Write-Output "FAILED $bookPath : $($_.Exception.Message)"
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "DONE $patched"
if ($patched -eq $books.Count) { exit 0 } else { exit 1 }

# run_v8.0_full.ps1
# Version 1.0 / 2026-09-10
# Full rollout of report v8.0 followed by a data load and a report run.
# Order follows docs\rules.md ([e2e/poryadok]): build first, compile, only then prod.
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Steps - each prints a marker, the script stops at the first failure:
#   STEP1  install VBA + Power Query + template into build\ReportMTO v7.0.xlsm
#   STEP2  force a full VBA compile of that book (tools\compile_check.ps1)
#   STEP3  install the same sources into the production workbook
#   STEP4  load every data\*.json into the production workbook one by one (upsert)
#   STEP5  modContentMTO.BuildPivots
#   STEP6  modMain.GenerateReport - the closing MsgBox is dismissed by a SendKeys job
#
# Usage from the project root:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\run_v8.0_full.ps1 > tools\run_v8.0_full.log 2>&1
#   powershell ... -File tools\run_v8.0_full.ps1 -SkipBuild        (straight to prod)
#   powershell ... -File tools\run_v8.0_full.ps1 -SkipLoad         (no JSON reload)

param(
    [string]$Book = "ReportMTO.xlsm",
    [switch]$SkipBuild,
    [switch]$SkipLoad
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
if (-not [System.IO.Path]::IsPathRooted($Book)) {
    $Book = [System.IO.Path]::GetFullPath((Join-Path $root $Book))
}
$buildBook = Join-Path $root "build\ReportMTO v7.0.xlsm"
$dataDir   = Join-Path $root "data"
$resultDir = Join-Path $root "result"

function Say([string]$s) { Write-Output ((Get-Date -Format "HH:mm:ss") + "  " + $s) }
function Die([string]$s) { Say ("DONE_FAIL " + $s); exit 1 }

Say ("RUN_START book=" + $Book)
if (-not (Test-Path $Book)) { Die ("book not found: " + $Book) }

# A stale Excel lock file means the workbook is open (or Excel crashed on it).
$lock = Join-Path (Split-Path -Parent $Book) ("~$" + (Split-Path -Leaf $Book))
if (Test-Path $lock) { Say ("WARN lock file present: " + $lock + " - close Excel or delete it") }

# ---------------------------------------------------------------- STEP1 / STEP2
if (-not $SkipBuild) {
    if (Test-Path $buildBook) {
        Say "STEP1 install into build"
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "install\install.ps1") -Target $buildBook
        if ($LASTEXITCODE -ne 0) { Die ("install into build returned " + $LASTEXITCODE) }

        Say "STEP2 compile check on the build book"
        $cc = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "tools\compile_check.ps1") -Book $buildBook
        $cc | ForEach-Object { Say ("  " + $_) }
        if (($cc -join " ") -notmatch "COMPILE_OK") { Die "VBA compile failed - see COMPILE_FAIL above" }
    } else {
        Say ("STEP1 SKIP - no build book at " + $buildBook)
    }
} else {
    Say "STEP1/STEP2 SKIP (-SkipBuild)"
}

# ---------------------------------------------------------------- STEP3
Say "STEP3 install into the production workbook"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "install\install_prod.ps1") -Target $Book
if ($LASTEXITCODE -ne 0) { Die ("install_prod returned " + $LASTEXITCODE) }

# ---------------------------------------------------------------- STEP4..STEP6
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$wb = $null
$code = 0

try {
    $wb = $excel.Workbooks.Open($Book, 0, $false)
    Say "OPENED"
    $wsData = $wb.Sheets.Item("tbDATA")
    $lo = $wsData.ListObjects.Item("tbDATA")
    Say ("ROWS_START " + [int]$lo.ListRows.Count)

    if (-not $SkipLoad) {
        $files = @(Get-ChildItem -Path (Join-Path $dataDir "*.json") | Sort-Object Name)
        Say ("STEP4 load " + $files.Count + " json file(s)")
        if ($files.Count -eq 0) { Say "WARN no json files in data\" }
        foreach ($f in $files) {
            $t0 = Get-Date
            $rowsBefore = [int]$lo.ListRows.Count
            $safe = $f.FullName -replace '"', '""'
            $wb.Queries.Item("prmSourcePath").Formula =
                '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
            $qt = $lo.QueryTable
            $qt.BackgroundQuery = $false
            $null = $qt.Refresh($false)
            $rowsAfter = [int]$lo.ListRows.Count
            $sec = [int]((Get-Date) - $t0).TotalSeconds
            Say ("  LOADED " + $f.Name + " | rows " + $rowsBefore + " -> " + $rowsAfter + " | " + $sec + " s")
        }
        Say ("ROWS_AFTER_LOAD " + [int]$lo.ListRows.Count)
        $wb.Save()
    } else {
        Say "STEP4 SKIP (-SkipLoad)"
    }

    Say "STEP5 BuildPivots"
    $null = $excel.Run("modContentMTO.BuildPivots")
    Say "  BuildPivots ok"

    Say "STEP6 GenerateReport"
    $startTime = Get-Date
    $excelPid = (Get-Process EXCEL | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    # The run ends with a MsgBox; a background job presses ENTER until it is gone.
    $job = Start-Job -ArgumentList $excelPid -ScriptBlock {
        param($pidExcel)
        Start-Sleep -Seconds 45
        $wsh = New-Object -ComObject WScript.Shell
        for ($k = 1; $k -le 120; $k++) {
            Start-Sleep -Seconds 10
            try {
                $null = $wsh.AppActivate($pidExcel)
                Start-Sleep -Milliseconds 400
                $wsh.SendKeys("{ENTER}")
            } catch { }
        }
    }
    try {
        $null = $excel.Run("modMain.GenerateReport")
    } finally {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }

    $latest = Get-ChildItem $resultDir -Filter "Report_*.html" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $latest) {
        Say "STEP6 FAIL - no fresh Report_*.html in result\"
        $code = 1
    } else {
        Say ("RESULT_FILE " + $latest.FullName + " | " + [int]($latest.Length / 1024) + " KB")
    }
    $wb.Save()
} catch {
    Say ("EXCEPTION " + $_.Exception.Message)
    $code = 1
} finally {
    try { if ($null -ne $wb) { $wb.Close($true) } } catch { }
    try { $excel.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    # Kill only the EXCEL processes this script started.
    foreach ($p in Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ($before -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
}

if ($code -eq 0) { Say "DONE_OK" } else { Say "DONE_FAIL see markers above" }
exit $code

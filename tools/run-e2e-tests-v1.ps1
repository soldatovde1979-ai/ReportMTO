# run-e2e-tests-v1.ps1
#
# Version 1.0 / 2026-09-07
# End-to-end test pipeline for ReportMTO build book:
#   STEP1 - load tests/test_sppr_tablet_v1.json into empty tbDATA, expect 25 rows,
#           then run tests/modSelfTest.bas SelfTest (all CHECK=1, DONE)
#   STEP2 - reload the same JSON, expect still 25 rows (upsert no-growth)
#   STEP3 - load tests/test_upsert_same_keys_v1.json, expect 25 rows and
#           ZN-002 arm=PK on both rows (modSelfTest.SelfTestUpsert)
#   STEP4 - run modMain.DebugGenerateOffline (MsgBox closed via SendKeys),
#           check the newest debug_*.html: no '{{', has '[offline]', has 'drill-dump'
# Exit code: 0 = all passed, 1 = at least one step failed.
#
# Source is ASCII-only on purpose: PowerShell 5.1 reads .ps1 without BOM as ANSI
# and cyrillic string literals break parsing. All cyrillic comparisons live in VBA.
#
# Scenario and rules for adding new tests: tests/testing-e2e-v1.md

param(
    [string]$BookPath = "build\ReportMTO v7.0.xlsm",
    [string]$ResultFolder = "C:\Projects\ReportMTO\result"
)

$ErrorActionPreference = "Stop"
$tempDir = $env:TEMP
$workBook = Join-Path $tempDir "ReportMTO_e2e.xlsm"
$tplPath = Join-Path $tempDir "tmp_index.html"
$selfTestBas = (Resolve-Path "tests\modSelfTest.bas").Path
$json1 = (Resolve-Path "tests\test_sppr_tablet_v1.json").Path
$json2 = (Resolve-Path "tests\test_upsert_same_keys_v1.json").Path
$tplSrc = (Resolve-Path "tmp_index.html").Path
$resultLog = Join-Path $tempDir "selftest_result.txt"
$log = Join-Path $tempDir "e2e_result.txt"

$fails = 0

function StepResult([string]$name, [bool]$ok, [string]$detail) {
    if ($ok) {
        $line = "PASS $name"
    } else {
        $line = "FAIL $name"
        $script:fails += 1
    }
    if ($detail -ne "") { $line = $line + " | " + $detail }
    Add-Content -Path $log -Value $line
    Write-Output $line
}

function RefreshSource([string]$jsonPath) {
    $safe = $jsonPath -replace '"', '""'
    $script:wb.Queries.Item("prmSourcePath").Formula = '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
    $lo = $script:wsData.ListObjects.Item("tbDATA")
    $qt = $lo.QueryTable
    $qt.BackgroundQuery = $false
    $null = $qt.Refresh($false)
    return [int]$lo.ListRows.Count
}

function CheckSelfTestLines([string[]]$markers) {
    # All CHECK lines must end with '=1'; no '=0', no 'CHECK:VBA_ERROR'; all markers present.
    $lines = Get-Content -Path $resultLog -Encoding Default -ErrorAction SilentlyContinue
    if ($null -eq $lines -or $lines.Count -eq 0) { return $false }
    foreach ($ln in $lines) {
        if ($ln -match "CHECK:.*=0$") { return $false }
        if ($ln -match "CHECK:VBA_ERROR") { return $false }
    }
    foreach ($m in $markers) {
        $found = $false
        foreach ($ln in $lines) { if ($ln -eq $m) { $found = $true } }
        if (-not $found) { return $false }
    }
    return $true
}

function BadLines() {
    $bad = @()
    $lines = Get-Content -Path $resultLog -Encoding Default -ErrorAction SilentlyContinue
    if ($null -eq $lines) { return "no selftest_result.txt" }
    foreach ($ln in $lines) {
        if ($ln -match "CHECK:.*=0$" -or $ln -match "CHECK:VBA_ERROR") { $bad += $ln }
    }
    if ($bad.Count -eq 0) { return "" } else { return ($bad -join " ;; ") }
}

if (Test-Path $log) { Remove-Item $log -Force }
if (Test-Path $resultLog) { Remove-Item $resultLog -Force }
Write-Output "E2E START (book: $BookPath)"

if (-not (Test-Path $BookPath)) {
    StepResult "PREP_BOOK" $false "book not found: $BookPath"
    exit 1
}
Copy-Item $BookPath $workBook -Force
Copy-Item $tplSrc $tplPath -Force
Write-Output "COPIED"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$script:wb = $null
$script:wsData = $null

try {
    $script:wb = $excel.Workbooks.Open($workBook, 0, $false)
    Write-Output "OPENED"

    # Import test module (UTF-8 -> ANSI 1251, same conversion as build script).
    $vbProj = $wb.VBProject
    try { $vbProj.VBComponents.Remove($vbProj.VBComponents.Item("modSelfTest")) } catch { }
    $utf8Text = [System.IO.File]::ReadAllText($selfTestBas)
    $ansi = [System.Text.Encoding]::GetEncoding(1251)
    $ansiTemp = Join-Path $tempDir "modSelfTest_ansi.bas"
    [System.IO.File]::WriteAllText($ansiTemp, $utf8Text, $ansi)
    $vbProj.VBComponents.Import($ansiTemp) | Out-Null
    $wb.Save()
    Write-Output "MODULE_IMPORTED"

    $script:wsData = $wb.Sheets.Item("tbDATA")
    $lo = $wsData.ListObjects.Item("tbDATA")
    if ($lo.ListRows.Count -gt 0) { $lo.DataBodyRange.Delete() }

    # STEP1: load test JSON #1, expect 25 rows; run SelfTest.
    $rows1 = RefreshSource $json1
    $step1a = ($rows1 -eq 25)
    StepResult "STEP1_ROWS" $step1a "rows=$rows1 expected=25"

    $null = $excel.Run("modSelfTest.SelfTest")
    $step1b = CheckSelfTestLines @("DONE")
    StepResult "STEP1_SELFTEST" $step1b (BadLines)

    # STEP2: reload the same JSON, expect still 25 rows.
    $rows2 = RefreshSource $json1
    StepResult "STEP2_RELOAD" ($rows2 -eq 25) "rows=$rows2 expected=25"

    # STEP3: load upsert JSON, expect 25 rows and ZN-002 arm=PK; run SelfTestUpsert.
    $rows3 = RefreshSource $json2
    $step3a = ($rows3 -eq 25)
    StepResult "STEP3_ROWS" $step3a "rows=$rows3 expected=25"

    $null = $excel.Run("modSelfTest.SelfTestUpsert")
    $step3b = CheckSelfTestLines @("UPSERT_START", "CHECK:UPSERT_ARM_PK=1", "DONE_UPSERT")
    StepResult "STEP3_UPSERT" $step3b (BadLines)

    # STEP4: DebugGenerateOffline (MsgBox auto-closed by SendKeys background job).
    $startTime = Get-Date
    $excelPid = (Get-Process EXCEL | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    $job = Start-Job -ArgumentList $excelPid -ScriptBlock {
        param($pidExcel)
        Start-Sleep -Seconds 60
        $wsh = New-Object -ComObject WScript.Shell
        for ($k = 1; $k -le 18; $k++) {
            Start-Sleep -Seconds 10
            try {
                $null = $wsh.AppActivate($pidExcel)
                Start-Sleep -Milliseconds 500
                $wsh.SendKeys("{ENTER}")
            } catch { }
        }
    }
    try {
        $null = $excel.Run("modMain.DebugGenerateOffline")
    } finally {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
    $latest = Get-ChildItem $ResultFolder -Filter "debug_*.html" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $latest) {
        StepResult "STEP4_OFFLINE" $false "no fresh debug_*.html in $ResultFolder"
    } else {
        $t = [System.IO.File]::ReadAllText($latest.FullName)
        $ok4 = (-not $t.Contains("{{")) -and $t.Contains("[offline]") -and $t.Contains("drill-dump")
        StepResult "STEP4_OFFLINE" $ok4 ("file=" + $latest.Name + " placeholder=" + $t.Contains("{{") + " offline=" + $t.Contains("[offline]") + " dump=" + $t.Contains("drill-dump"))
    }
} finally {
    if ($null -ne $script:wb) { try { $wb.Close($false) } catch { } }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output ""
if ($fails -eq 0) {
    Write-Output "E2E RESULT: ALL PASSED"
    exit 0
} else {
    Write-Output ("E2E RESULT: FAILED (" + $fails + " step(s))")
    exit 1
}

# reimport_open_book.ps1
# Version 1.0 / 2026-09-10
# Re-imports the given VBA modules into the ALREADY OPEN workbook (attaches via
# GetActiveObject). Re-encodes UTF-8 -> ANSI 1251 before Import (VBE reads .bas as ANSI).
# Used when the book is open in Excel and cannot be re-opened by install.ps1.
#
# Usage from the project root:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\reimport_open_book.ps1
#   powershell ... -File tools\reimport_open_book.ps1 -Modules modMain,modContentMTO

param(
    [string[]]$Modules = @("modMain", "modContentMTO"),
    [string]$BookName = "ReportMTO.xlsm"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
$vbaDir = Join-Path $root "src\vba"
$ansi = [System.Text.Encoding]::GetEncoding(1251)
$tempDir = $env:TEMP

function Say([string]$s) { Write-Output ((Get-Date -Format "HH:mm:ss") + "  " + $s) }
function Die([string]$s) { Say ("DONE_FAIL " + $s); exit 1 }

Say "RUN_START (attach to running Excel)"

$excel = $null
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Die "no running Excel instance found"
}

$wb = $null
try {
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq $BookName) { $wb = $w; break }
    }
    if ($null -eq $wb) { Die ("workbook not open: " + $BookName) }

    Say ("ATTACHED book=" + $wb.FullName)
    $vp = $wb.VBProject

    foreach ($mod in $Modules) {
        $src = Join-Path $vbaDir ($mod + ".bas")
        if (-not (Test-Path $src)) { Say ("SKIP " + $mod + " - source not found"); continue }

        $utf8 = [System.IO.File]::ReadAllText($src)
        $ansiTemp = Join-Path $tempDir ($mod + "_reimport_ansi.bas")
        [System.IO.File]::WriteAllText($ansiTemp, $utf8, $ansi)

        $curComp = $null
        try { $curComp = $vp.VBComponents.Item($mod) } catch { }
        if ($null -ne $curComp) {
            $vp.VBComponents.Remove($curComp)
            Say ("REMOVED " + $mod)
        }
        $vp.VBComponents.Import($ansiTemp) | Out-Null
        Say ("IMPORTED " + $mod)
    }

    $wb.Save()
    Say "SAVED"
    Say "DONE_OK"
} catch {
    Say ("EXCEPTION " + $_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}

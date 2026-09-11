# apply_vba_tmp.ps1 - re-imports the given VBA modules into a workbook
# (Remove + Import + Save), sources re-encoded UTF-8 -> ANSI 1251 (VBE reads
# .bas only as ANSI, project rule). ASCII-only.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\apply_vba_tmp.ps1 `
#       -Book "build\ReportMTO v7.0.xlsm" -Modules modContentMTO,modAggregate

param(
    [string]$Book = "build\ReportMTO v7.0.xlsm",
    [string[]]$Modules = @("modContentMTO", "modAggregate")
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
if (-not [System.IO.Path]::IsPathRooted($Book)) {
    $Book = [System.IO.Path]::GetFullPath((Join-Path $root $Book))
}

$ansiEnc = [System.Text.Encoding]::GetEncoding(1251)
$tempFiles = @()
foreach ($m in $Modules) {
    $src = Join-Path (Join-Path $root "src\vba") ($m + ".bas")
    $ansi = Join-Path $env:TEMP ($m + "_ansi.bas")
    $t = [IO.File]::ReadAllText($src)
    [IO.File]::WriteAllText($ansi, $t, $ansiEnc)
    $tempFiles += $ansi
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($Book, 0, $false)
    $vbProj = $wb.VBProject
    $null = $vbProj.VBComponents.Count
    foreach ($m in $Modules) {
        $comp = $vbProj.VBComponents.Item($m)
        $vbProj.VBComponents.Remove($comp)
        $ansi = Join-Path $env:TEMP ($m + "_ansi.bas")
        $vbProj.VBComponents.Import($ansi) | Out-Null
        $lines = $vbProj.VBComponents.Item($m).CodeModule.CountOfLines
        Write-Output ("REIMPORT " + $m + " OK lines=" + $lines)
    }
    $wb.Save()
    Write-Output ("SAVED " + $Book)
    $wb.Close($false)
} catch {
    Write-Output ("REIMPORT FAILED: " + $_.Exception.Message)
    try { $wb.Close($false) } catch { }
    throw
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    foreach ($f in $tempFiles) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
}

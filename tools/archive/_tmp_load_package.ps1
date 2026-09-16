# _tmp_load_package.ps1 - ВРЕМЕННЫЙ: пакетная загрузка data в корневую книгу (ТЗ v1.0, задача 9).
# Ставит DEBUG=2 перед загрузкой (замечание владельца), запускает LoadPackage, печатает
# строки tbDATA, тайминг и хвост ReportMTO.log. Удалить после использования.
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$book = "D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm"
$dataFolder = "D:\GOOGLEDISK\PROJECTs\ReportMTO\data"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Open($book, 0, $false)
try {
    $ws = $null
    foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "Variable") { $ws = $sh } }
    if ($null -eq $ws) { throw "no Variable sheet" }
    $lo = $ws.ListObjects.Item("tblVariable")
    $keyCol = $lo.ListColumns.Item("Key").DataBodyRange
    $valCol = $lo.ListColumns.Item("Value").DataBodyRange

    # DATA/SOURCE_FOLDER - прочитать и показать
    $src = ""
    for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
        if ($keyCol.Cells.Item($r, 1).Text -eq "DATA/SOURCE_FOLDER") {
            $src = $valCol.Cells.Item($r, 1).Text
        }
    }
    Write-Output ("SOURCE_FOLDER=[" + $src + "]")

    # DEBUG = 2 (замечание владельца: перед загрузкой DEBUG=2)
    $set = $false
    for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
        if ($keyCol.Cells.Item($r, 1).Text -eq "DEBUG") {
            $valCol.Cells.Item($r, 1).Value2 = "2"
            $set = $true
        }
    }
    if (-not $set) {
        $lo.ListRows.Add() | Out-Null
        $n = $lo.ListRows.Count
        $lo.ListColumns.Item("Key").DataBodyRange.Cells.Item($n, 1).Value2 = "DEBUG"
        $lo.ListColumns.Item("Value").DataBodyRange.Cells.Item($n, 1).Value2 = "2"
    }
    Write-Output "DEBUG_SET=2"
    $wb.Save()

    $t0 = Get-Date
    Write-Output ("LOAD_START " + $t0.ToString("HH:mm:ss"))
    try {
        $ret = $xl.Run("modMain.LoadPackage")
        Write-Output ("LOAD_RET=" + $ret)
    } catch {
        Write-Output ("LOAD_ERR: " + $_.Exception.Message)
    }
    $el = ((Get-Date) - $t0).TotalSeconds
    Write-Output ("LOAD_SECONDS=" + [math]::Round($el, 1))

    $wsData = $null
    foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "tbDATA") { $wsData = $sh } }
    if ($null -ne $wsData) {
        $tlo = $wsData.ListObjects.Item("tbDATA")
        Write-Output ("TBDATA_ROWS=" + $tlo.ListRows.Count)
    }
    $wb.Save()
} finally {
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

Write-Output "--- TAIL ReportMTO.log ---"
$log = Join-Path (Split-Path $book) "ReportMTO.log"
if (Test-Path $log) {
    Get-Content $log -Tail 15 -Encoding UTF8 | ForEach-Object { Write-Output $_ }
} else {
    Write-Output "NO ReportMTO.log"
}
Write-Output "LOAD_PACKAGE_DONE"

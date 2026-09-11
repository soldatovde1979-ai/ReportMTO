# 8.1.0__variable_keys.ps1
# requires-from: 0.0.0
# Заводит ключи листа Variable, появившиеся к версии 8.1.0:
#   DATA/KEEP_WEEKS = 52 - глубина хранения tbDATA в неделях (ретеншн, qKeepWeeks.pq).
# Идемпотентна: существующие ключи не трогает, значения не перезаписывает.
#
# Контракт миграции (см. docs\version-guide.md):
#   - принимает открытую книгу $Workbook и путь $BookPath;
#   - НЕ сохраняет и НЕ закрывает книгу - это делает release.ps1;
#   - пишет в вывод строки вида "MIG ..." ;
#   - при непоправимой ошибке бросает исключение (throw) - релиз остановится.
#
# Файл UTF-8 с BOM: PowerShell 5.1 иначе читает кириллицу как ANSI.

param(
    $Workbook,
    [string]$BookPath
)

$ErrorActionPreference = "Stop"

$keys = @(
    @{ Key = "DATA/KEEP_WEEKS"; Value = 52 }
)

$ws = $null
foreach ($sh in $Workbook.Worksheets) { if ($sh.Name -eq "Variable") { $ws = $sh } }
if ($null -eq $ws) { throw "на книге нет листа Variable" }

$lo = $null
foreach ($l in $ws.ListObjects) { if ($l.Name -eq "tblVariable") { $lo = $l } }
if ($null -eq $lo) { throw "на листе Variable нет таблицы tblVariable" }

foreach ($item in $keys) {
    $name = $item.Key
    $found = $false
    $keyCol = $lo.ListColumns.Item("Key").DataBodyRange
    for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
        if ($keyCol.Cells.Item($r, 1).Text -eq $name) { $found = $true }
    }
    if ($found) {
        Write-Output ("MIG SKIP " + $name + " - ключ уже есть, значение не трогаем")
    } else {
        $lo.ListRows.Add() | Out-Null
        $n = $lo.ListRows.Count
        $lo.ListColumns.Item("Key").DataBodyRange.Cells.Item($n, 1).Value2 = $name
        $lo.ListColumns.Item("Value").DataBodyRange.Cells.Item($n, 1).Value2 = $item.Value
        Write-Output ("MIG ADD " + $name + " = " + $item.Value)
    }
}

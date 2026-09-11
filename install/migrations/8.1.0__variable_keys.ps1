# 8.1.0__variable_keys.ps1
# requires-from: 0.0.0
#
# Версия 2.1 от 11.09.2026: запись через .Value2 на DataBodyRange падала с
#   "Unable to cast System.Double to System.String". Пишем через координаты листа
#   и свойство .Formula - оно принимает строку, приведения типов нет.
# Версия 2.0 от 11.09.2026: первый боевой прогон упал с
#   "Unable to cast object of type 'System.Int32' to type 'System.String'".
#   Точное место было не видно - вся миграция шла одним куском. Переписано:
#   каждый шаг печатает метку MIG STEP, все значения приводятся к типу явно,
#   добавление строки идёт через ListRows.Add, а при отказе - прямой записью
#   под таблицу (ListObject расширяется сам).
# Версия 1.0 от 11.09.2026: исходный вариант.
#
# Заводит ключи листа Variable, появившиеся к версии 8.1.0:
#   DATA/KEEP_WEEKS = 52 - глубина хранения tbDATA в неделях (ретеншн, qKeepWeeks.pq).
# Идемпотентна: существующие ключи не трогает, значения не перезаписывает.
#
# Контракт: принимает открытую книгу, НЕ сохраняет и НЕ закрывает её - это делает
# release.ps1. Файл UTF-8 с BOM.

param(
    $Workbook,
    [string]$BookPath
)

$ErrorActionPreference = "Stop"

Write-Output "MIG STEP 1 - ищу лист Variable"
$ws = $null
foreach ($sh in $Workbook.Worksheets) { if ([string]$sh.Name -eq "Variable") { $ws = $sh } }
if ($null -eq $ws) { throw "на книге нет листа Variable" }

Write-Output "MIG STEP 2 - ищу таблицу tblVariable"
$lo = $null
foreach ($l in $ws.ListObjects) { if ([string]$l.Name -eq "tblVariable") { $lo = $l } }
if ($null -eq $lo) { throw "на листе Variable нет таблицы tblVariable" }

Write-Output "MIG STEP 3 - читаю существующие ключи"
$keyCol = $lo.ListColumns.Item("Key").DataBodyRange
$rowCount = [int]$keyCol.Rows.Count
$existing = @()
for ($r = 1; $r -le $rowCount; $r++) {
    $existing += [string]$keyCol.Cells.Item($r, 1).Text
}
Write-Output ("MIG STEP 3 - ключей в таблице: " + $rowCount)

# Ключи этой версии: имя (строка) и значение (число).
# Text - то, что реально записывается в ячейку (Formula принимает строку).
$wanted = @(
    @{ Name = "DATA/KEEP_WEEKS"; Text = "52" }
)

foreach ($item in $wanted) {
    $name = [string]$item.Name

    if ($existing -contains $name) {
        Write-Output ("MIG SKIP " + $name + " - ключ уже есть, значение не трогаем")
        continue
    }

    Write-Output ("MIG STEP 4 - добавляю строку " + $name)
    $added = $false
    try {
        $lo.ListRows.Add() | Out-Null
        $added = $true
        Write-Output "MIG STEP 4 - строка добавлена через ListRows.Add"
    } catch {
        Write-Output ("MIG STEP 4 - ListRows.Add отказал (" + $_.Exception.Message + "), пишу прямо под таблицу")
    }

    # Пишем ВСЕГДА через координаты листа и свойство Formula: оно принимает строку,
    # и никакого приведения типов в COM не происходит. Присвоение .Value2 на
    # DataBodyRange падало с "Unable to cast System.Double to System.String" -
    # два прогона 11.09.2026, шаг MIG STEP 5.
    $firstRow = [int]$lo.Range.Row
    $rowsTotal = [int]$lo.Range.Rows.Count
    $col = [int]$lo.Range.Column
    if ($added) {
        $target = $firstRow + $rowsTotal - 1      # последняя строка таблицы, её только что добавили
    } else {
        $target = $firstRow + $rowsTotal          # строка под таблицей, Excel расширит сам
    }

    Write-Output ("MIG STEP 5a - строка " + $target + ", колонка ключа " + $col)
    $ws.Cells.Item($target, $col).Formula = $name
    Write-Output "MIG STEP 5b - имя ключа записано"
    $ws.Cells.Item($target, ($col + 1)).Formula = [string]$item.Text
    Write-Output "MIG STEP 5c - значение записано"

    Write-Output ("MIG ADD " + $name + " = " + [string]$item.Text)
}

Write-Output "MIG STEP 6 - готово"

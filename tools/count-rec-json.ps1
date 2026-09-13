# count-rec-json.ps1
# Считает количество записей и уникальных записей в JSON-выгрузке МТО (массив записей).
#
# Ключ уникальности — по data.md §2.1 / fnComputeKey.pq, но дата заменена на ГОД:
#   Key = number | year(date) | ready_for | direction
# где year(date) — первые 4 символа даты в формате yyyy-MM-ddTHH:mm:ss.
# null / отсутствующие поля в ключе — пустая строка (как Record.FieldOrDefault в PQ).
#
# Использование:
#   powershell -NoProfile -File tools/count-rec-json.ps1 -Json "data\file.json"

param(
    [Parameter(Mandatory = $true)]
    [string]$Json
)

$ErrorActionPreference = 'Stop'

# Вывод кириллицы в терминалах с OEM-кодировкой (cmd/старый PS)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$fullPath = Resolve-Path -Path $Json
$raw = Get-Content -Path $fullPath -Raw -Encoding UTF8
# PS 5.1: разворачивание JSON-массива только через переменную + @(), иначе Count = 1
$parsed  = $raw | ConvertFrom-Json
$records = @($parsed)

$keys = @{}
foreach ($r in $records) {
    $number = [string]$r.number

    $year = ''
    if ($null -ne $r.date) {
        $d = [string]$r.date
        if ($d.Length -ge 4) { $year = $d.Substring(0, 4) }
    }

    $readyFor  = [string]$r.ready_for
    $direction = [string]$r.direction

    $key = "$number|$year|$readyFor|$direction"
    if (-not $keys.ContainsKey($key)) { $keys[$key] = 1 }
}

$total  = $records.Count
$unique = $keys.Count
$dups   = $total - $unique

Write-Output ("Файл: {0}" -f $fullPath)
Write-Output ("Всего записей: {0}" -f $total)
Write-Output ("Уникальных записей (Key = number|year|ready_for|direction): {0}" -f $unique)
Write-Output ("Дублей: {0}" -f $dups)

# count-rec-json.ps1
# Считает количество записей и уникальных записей в JSON-выгрузках МТО.
#
# Ключ уникальности — по data.md §2.1 / fnComputeKey.pq, но дата заменена на ГОД:
#   Key = number | year(date) | ready_for | direction
# где year(date) — первые 4 символа даты в формате yyyy-MM-ddTHH:mm:ss.
# null / отсутствующие поля в ключе — пустая строка (как Record.FieldOrDefault в PQ).
#
# Один файл:
#   powershell -NoProfile -File tools/count-rec-json.ps1 -Json "data\file.json"
# Вся папка (общий накопительный ключ по всем файлам — пересечения выгрузок схлопываются):
#   powershell -NoProfile -File tools/count-rec-json.ps1 -Folder "data"

param(
    [string]$Json,
    [string]$Folder
)

$ErrorActionPreference = 'Stop'

# Вывод кириллицы в терминалах с OEM-кодировкой (cmd/старый PS)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

if (-not $Json -and -not $Folder) {
    Write-Output 'Укажите -Json <файл> или -Folder <папка>'
    exit 1
}

if ($Json) {
    $files = @(Get-Item -Path $Json)
} else {
    $files = @(Get-ChildItem -Path $Folder -Filter '*.json' | Sort-Object Name)
}

# Накопительный словарь ключей: действует весь прогон, между файлами одной обработки
$globalKeys = @{}
$sumRecs = 0

foreach ($f in $files) {
    $raw = Get-Content -Path $f.FullName -Raw -Encoding UTF8
    # PS 5.1: разворачивание JSON-массива только через переменную + @(), иначе Count = 1
    $parsed  = $raw | ConvertFrom-Json
    $records = @($parsed)

    $sumRecs += $records.Count
    $newCount = 0

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
        if (-not $globalKeys.ContainsKey($key)) {
            $globalKeys[$key] = 1
            $newCount++
        } else {
            $globalKeys[$key] = [int]$globalKeys[$key] + 1
        }
    }

    Write-Output ("{0}: записей = {1}, новых ключей = {2}" -f $f.Name, $records.Count, $newCount)
}

$unique = $globalKeys.Count
Write-Output ''
Write-Output ("Всего записей (сумма по файлам): {0}" -f $sumRecs)
Write-Output ("Уникальных записей (общий Key по всем файлам): {0}" -f $unique)
Write-Output ("Пересечений между файлами: {0}" -f ($sumRecs - $unique))

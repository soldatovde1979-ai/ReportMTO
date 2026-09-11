# diag_verify_tmp.ps1
# Временная диагностика (read-only): сравнивает VBA-модули и формулы Power
# Query в книге с исходниками и печатает позицию первого отличия и коды
# символов в этой позиции. Ничего не удаляет и не сохраняет.
# Файл в UTF-8 с BOM (требование PowerShell 5.1 для кириллицы).
#
# Запуск:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\diag_verify_tmp.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\diag_verify_tmp.ps1 -Book "ReportMTO.xlsm"

param(
    [string]$Book = "ReportMTO.xlsm",
    [string]$Root = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($Root -eq "") { $Root = Split-Path -Parent $scriptRoot }
$bookPath = [System.IO.Path]::GetFullPath((Join-Path $Root $Book))

function Normalize-Text([string]$text) {
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $lines = $text -split "`n"
    $out = ""
    foreach ($line in $lines) { $out += $line.TrimEnd() + "`n" }
    return $out.TrimEnd("`n")
}

function Show-FirstDiff([string]$src, [string]$cur, [string]$label) {
    $n = [Math]::Min($src.Length, $cur.Length)
    $i = 0
    while ($i -lt $n -and $src[$i] -eq $cur[$i]) { $i++ }
    if ($i -eq $n -and $src.Length -eq $cur.Length) {
        Write-Output ($label + " - СОВПАДАЕТ")
    } else {
        Write-Output ($label + " - РАЗЛИЧАЕТСЯ: позиция " + $i + ", длина исходника " + $src.Length + ", длина в книге " + $cur.Length)
        $out = "  коды символов:"
        if ($i -lt $src.Length) { $out += " исходник=" + [int][char]$src[$i] }
        if ($i -lt $cur.Length) { $out += " книга=" + [int][char]$cur[$i] }
        Write-Output $out
        if ($i -gt 0) {
            $s1 = [Math]::Max(0, $i - 30)
            $l1 = [Math]::Min(60, $src.Length - $s1)
            $l2 = [Math]::Min(60, $cur.Length - $s1)
            Write-Output ("  контекст исходника: ..." + $src.Substring($s1, $l1).Replace("`n", "|") + "...")
            Write-Output ("  контекст в книге:   ..." + $cur.Substring($s1, $l2).Replace("`n", "|") + "...")
        }
    }
}

Write-Output "DIAG_BOOK $bookPath - открываю только для чтения"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($bookPath, 0, $true)

$ansi = [System.Text.Encoding]::GetEncoding(1251)
$vp = $wb.VBProject

foreach ($mod in @("modMain", "modContentMTO")) {
    $src = Join-Path $Root ("src\vba\" + $mod + ".bas")
    $utf8 = [System.IO.File]::ReadAllText($src)
    $srcText = Normalize-Text ($ansi.GetString($ansi.GetBytes($utf8)))
    $curComp = $null
    try { $curComp = $vp.VBComponents.Item($mod) } catch { }
    if ($null -eq $curComp) { Write-Output ("VBA " + $mod + " - МОДУЛЬ ОТСУТСТВУЕТ В КНИГЕ"); continue }
    $curText = Normalize-Text $curComp.CodeModule.Lines(1, $curComp.CodeModule.CountOfLines)
    Show-FirstDiff $srcText $curText ("VBA " + $mod)
}

$pqDir = Join-Path $Root "src\powerquery"
foreach ($pqFile in Get-ChildItem -Path $pqDir -Filter *.pq) {
    $pqName = [System.IO.Path]::GetFileNameWithoutExtension($pqFile.Name)
    $bytes = [System.IO.File]::ReadAllBytes($pqFile.FullName)
    $firstByte = -1
    if ($bytes.Length -gt 0) { $firstByte = $bytes[0] }
    Write-Output ("PQ_FILE " + $pqName + " - первый байт файла=" + $firstByte + " (239 = UTF-8 BOM)")
    $pqFormula = [System.IO.File]::ReadAllText($pqFile.FullName)
    $bookFormula = $null
    foreach ($q in $wb.Queries) { if ($q.Name -eq $pqName) { $bookFormula = $q.Formula } }
    if ($null -eq $bookFormula) { Write-Output ("PQ " + $pqName + " - ЗАПРОС ОТСУТСТВУЕТ В КНИГЕ"); continue }
    Show-FirstDiff $pqFormula $bookFormula ("PQ " + $pqName)
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "DIAG_DONE"

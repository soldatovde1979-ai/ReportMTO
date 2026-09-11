# release.ps1
#
# Version 1.0 / 11.09.2026: единая точка раскатки с версионированием.
#
# Что делает по шагам:
#   1) считает хеши исходников (src\vba\*.bas, src\powerquery\*.pq, tmp_index.html)
#      и сравнивает с install\release.state - какой набор исходников был установлен
#      в прошлый раз;
#   2) если исходники не менялись - версия остаётся прежней (повторный прогон
#      не должен накручивать номер);
#   3) ставит книгу через install\install_prod.ps1 (он вызывает install.ps1);
#   4) прогоняет tools\compile_check.ps1;
#   5) ТОЛЬКО если оба шага вернули 0: поднимает версию (по умолчанию патч,
#      -Bump minor|major - руками при смене контракта), пишет BUILD/VERSION,
#      BUILD/DATE, BUILD/SRC_MD5 на лист Variable книги, прогоняет недостающие
#      миграции из install\migrations и добавляет запись в CHANGELOG.md.
#
# Имя файла книги НЕ меняется: версия живёт внутри книги и в CHANGELOG.
#
# Использование (из корня проекта, книга должна быть закрыта):
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\release.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\release.ps1 -DryRun
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\release.ps1 -Bump minor
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\release.ps1 -Target ".\build\ReportMTO v7.0.xlsm"
#
# Файл UTF-8 с BOM: PowerShell 5.1 иначе читает кириллицу как ANSI.

param(
    [string]$Target = ".\ReportMTO.xlsm",
    [ValidateSet("patch", "minor", "major")]
    [string]$Bump = "patch",
    [switch]$DryRun,
    [switch]$SkipCompile
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptRoot
$versionFile = Join-Path $scriptRoot "VERSION"
$stateFile = Join-Path $scriptRoot "release.state"
$migDir = Join-Path $scriptRoot "migrations"
$changelog = Join-Path $root "CHANGELOG.md"

function Fail([string]$msg) { Write-Output ("RELEASE_FAIL " + $msg); exit 1 }

# ---------------------------------------------------------------- версия
function Read-BuildVersion() {
    if (-not (Test-Path $versionFile)) { Fail ("нет файла " + $versionFile) }
    $lines = Get-Content -LiteralPath $versionFile -Encoding UTF8
    foreach ($l in $lines) {
        $t = $l.Trim()
        if ($t -ne "") {
            if ($t -notmatch "^\d+\.\d+\.\d+$") { Fail ("первая строка VERSION не версия вида 1.2.3: " + $t) }
            return $t
        }
    }
    Fail "файл VERSION пуст"
}

function Read-VersionDescription() {
    $lines = Get-Content -LiteralPath $versionFile -Encoding UTF8
    $seen = $false
    $out = @()
    foreach ($l in $lines) {
        if (-not $seen) { if ($l.Trim() -ne "") { $seen = $true }; continue }
        $out += $l
    }
    return ($out -join " ").Trim()
}

function Step-Version([string]$ver, [string]$kind) {
    $p = $ver.Split(".")
    $ma = [int]$p[0]; $mi = [int]$p[1]; $pa = [int]$p[2]
    if ($kind -eq "major") { $ma += 1; $mi = 0; $pa = 0 }
    elseif ($kind -eq "minor") { $mi += 1; $pa = 0 }
    else { $pa += 1 }
    return ("" + $ma + "." + $mi + "." + $pa)
}

function Compare-Version([string]$a, [string]$b) {
    $pa = $a.Split("."); $pb = $b.Split(".")
    for ($i = 0; $i -lt 3; $i++) {
        $x = [int]$pa[$i]; $y = [int]$pb[$i]
        if ($x -lt $y) { return -1 }
        if ($x -gt $y) { return 1 }
    }
    return 0
}

# ---------------------------------------------------------------- хеши исходников
function Get-SourceMap() {
    $map = New-Object System.Collections.Specialized.OrderedDictionary
    $files = @()
    $files += Get-ChildItem -Path (Join-Path $root "src\vba") -Filter *.bas -File | Sort-Object Name
    $files += Get-ChildItem -Path (Join-Path $root "src\powerquery") -Filter *.pq -File | Sort-Object Name
    $tpl = Join-Path $root "tmp_index.html"
    if (Test-Path $tpl) { $files += Get-Item -LiteralPath $tpl }
    foreach ($f in $files) {
        $h = (Get-FileHash -Algorithm MD5 -LiteralPath $f.FullName).Hash
        $map[$f.Name] = $h
    }
    return $map
}

function Get-MapDigest($map) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($k in $map.Keys) { [void]$sb.Append($k); [void]$sb.Append(" "); [void]$sb.Append($map[$k]); [void]$sb.Append("`n") }
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $hash = $md5.ComputeHash($bytes)
    return (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
}

function Read-State() {
    $state = @{ Version = ""; Digest = ""; Files = @{} }
    if (-not (Test-Path $stateFile)) { return $state }
    foreach ($l in (Get-Content -LiteralPath $stateFile -Encoding UTF8)) {
        $t = $l.Trim()
        if ($t -eq "" -or $t.StartsWith("#")) { continue }
        $parts = $t -split " ", 3
        if ($parts[0] -eq "version") { $state.Version = $parts[1] }
        elseif ($parts[0] -eq "digest") { $state.Digest = $parts[1] }
        elseif ($parts[0] -eq "file") { $state.Files[$parts[1]] = $parts[2] }
    }
    return $state
}

function Write-State([string]$ver, $map, [string]$digest) {
    $lines = @()
    $lines += "# Служебный файл install\release.ps1 - состояние последнего успешного релиза."
    $lines += "# Руками не править: отсюда берётся ответ на вопрос «менялись ли исходники»."
    $lines += ("version " + $ver)
    $lines += ("digest " + $digest)
    foreach ($k in $map.Keys) { $lines += ("file " + $k + " " + $map[$k]) }
    [IO.File]::WriteAllLines($stateFile, $lines, (New-Object System.Text.UTF8Encoding($false)))
}

# ---------------------------------------------------------------- лист Variable
function Set-VariableKey($listObject, [string]$key, $value) {
    $keyCol = $listObject.ListColumns.Item("Key").DataBodyRange
    for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
        if ($keyCol.Cells.Item($r, 1).Text -eq $key) {
            $listObject.ListColumns.Item("Value").DataBodyRange.Cells.Item($r, 1).Value2 = $value
            return "SET"
        }
    }
    $listObject.ListRows.Add() | Out-Null
    $n = $listObject.ListRows.Count
    $listObject.ListColumns.Item("Key").DataBodyRange.Cells.Item($n, 1).Value2 = $key
    $listObject.ListColumns.Item("Value").DataBodyRange.Cells.Item($n, 1).Value2 = $value
    return "ADD"
}

function Get-VariableKey($listObject, [string]$key) {
    $keyCol = $listObject.ListColumns.Item("Key").DataBodyRange
    for ($r = 1; $r -le $keyCol.Rows.Count; $r++) {
        if ($keyCol.Cells.Item($r, 1).Text -eq $key) {
            return $listObject.ListColumns.Item("Value").DataBodyRange.Cells.Item($r, 1).Text
        }
    }
    return ""
}

# ================================================================== старт
$full = $Target
if (-not [System.IO.Path]::IsPathRooted($full)) { $full = Join-Path $root $full }
$full = [System.IO.Path]::GetFullPath($full)
if (-not (Test-Path $full)) { Fail ("книга не найдена: " + $full) }

$curVersion = Read-BuildVersion
$map = Get-SourceMap
$digest = Get-MapDigest $map
$state = Read-State

$changedFiles = @()
foreach ($k in $map.Keys) {
    $old = ""
    if ($state.Files.ContainsKey($k)) { $old = $state.Files[$k] }
    if ($old -ne $map[$k]) { $changedFiles += $k }
}
foreach ($k in $state.Files.Keys) { if (-not $map.Contains($k)) { $changedFiles += ($k + " (удалён)") } }

$sourcesChanged = ($digest -ne $state.Digest)
$newVersion = $curVersion
if ($sourcesChanged -and $state.Digest -ne "") { $newVersion = Step-Version $curVersion $Bump }

Write-Output ("RELEASE_TARGET " + $full)
Write-Output ("RELEASE_VERSION_CURRENT " + $curVersion + " - версия из install\VERSION")
if ($state.Digest -eq "") {
    Write-Output "RELEASE_FIRST_RUN - прошлого состояния нет, версия не поднимается, фиксируем текущую"
} elseif ($sourcesChanged) {
    Write-Output ("RELEASE_SOURCES_CHANGED " + $changedFiles.Count + " - изменились: " + ($changedFiles -join ", "))
    Write-Output ("RELEASE_VERSION_NEXT " + $newVersion + " - поднимаем (" + $Bump + ")")
} else {
    Write-Output "RELEASE_SOURCES_SAME - исходники не менялись, версия остаётся прежней"
}

if ($DryRun) {
    Write-Output "RELEASE_DRYRUN - книга не тронута, версия не записана"
    exit 0
}

# ---------------------------------------------------------------- установка
Write-Output "--- ШАГ 1: установка модулей и запросов"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scriptRoot "install_prod.ps1") -Target "$full"
if ($LASTEXITCODE -ne 0) { Fail ("установщик вернул " + $LASTEXITCODE + " - версия не изменена") }

if (-not $SkipCompile) {
    Write-Output "--- ШАГ 2: компиляция VBA"
    $compileOut = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "tools\compile_check.ps1") -Book "$full"
    $compileOut | ForEach-Object { Write-Output $_ }
    $okLine = $compileOut | Where-Object { $_ -like "COMPILE_OK*" }
    if ($null -eq $okLine) { Fail "компиляция не прошла - версия не изменена" }
}

# ---------------------------------------------------------------- книга: версия и миграции
Write-Output "--- ШАГ 3: версия и миграции в книге"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$applied = @()
try {
    $wb = $excel.Workbooks.Open($full, 0, $false)
    try {
        $ws = $null
        foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "Variable") { $ws = $sh } }
        if ($null -eq $ws) { throw "на книге нет листа Variable" }
        $lo = $null
        foreach ($l in $ws.ListObjects) { if ($l.Name -eq "tblVariable") { $lo = $l } }
        if ($null -eq $lo) { throw "на листе Variable нет таблицы tblVariable" }

        $bookVersion = (Get-VariableKey $lo "BUILD/VERSION").Trim()
        if ($bookVersion -notmatch "^\d+\.\d+\.\d+$") { $bookVersion = "0.0.0" }
        Write-Output ("BOOK_VERSION " + $bookVersion + " - версия, записанная в книге до этого прогона")

        if (Test-Path $migDir) {
            $migs = Get-ChildItem -Path $migDir -Filter *.ps1 -File | Where-Object { $_.Name -match "^(\d+\.\d+\.\d+)__" } |
                Sort-Object @{ Expression = { [version]($_.Name -replace '^(\d+\.\d+\.\d+)__.*$', '$1') } }
            foreach ($m in $migs) {
                $mv = ($m.Name -replace '^(\d+\.\d+\.\d+)__.*$', '$1')
                if ((Compare-Version $mv $bookVersion) -le 0) {
                    Write-Output ("MIG_SKIP " + $m.Name + " - версия книги уже " + $bookVersion)
                    continue
                }
                $reqFrom = "0.0.0"
                foreach ($line in (Get-Content -LiteralPath $m.FullName -Encoding UTF8 -TotalCount 20)) {
                    if ($line -match "^#\s*requires-from:\s*(\d+\.\d+\.\d+)") { $reqFrom = $Matches[1] }
                }
                if ((Compare-Version $bookVersion $reqFrom) -lt 0) {
                    throw ("миграция " + $m.Name + " требует книгу не ниже " + $reqFrom + ", а в книге " + $bookVersion + " - поставьте промежуточную версию")
                }
                Write-Output ("MIG_RUN " + $m.Name + " - применяем")
                & $m.FullName -Workbook $wb -BookPath $full | ForEach-Object { Write-Output ("  " + $_) }
                $applied += $m.Name
            }
        }

        $stampDate = (Get-Date -Format "dd.MM.yyyy HH:mm")
        Write-Output ("BUILD_VERSION " + (Set-VariableKey $lo "BUILD/VERSION" $newVersion) + " " + $newVersion)
        Write-Output ("BUILD_DATE " + (Set-VariableKey $lo "BUILD/DATE" $stampDate) + " " + $stampDate)
        Write-Output ("BUILD_SRC_MD5 " + (Set-VariableKey $lo "BUILD/SRC_MD5" $digest) + " " + $digest)

        $wb.Save()
    } finally {
        $wb.Close($false)
    }
} catch {
    Write-Output ("BOOK_ERROR " + $_.Exception.Message)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Fail "модули установлены, но версия/миграции в книгу не записаны - смотрите BOOK_ERROR выше"
}
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# ---------------------------------------------------------------- VERSION, state, CHANGELOG
if ($newVersion -ne $curVersion) {
    $desc = Read-VersionDescription
    $lines = @()
    $lines += $newVersion
    $lines += ("Собрано " + (Get-Date -Format "dd.MM.yyyy") + ". Изменены: " + ($changedFiles -join ", "))
    $lines += ""
    $lines += "Первая строка этого файла - версия сборки, всё ниже - описание."
    $lines += "Версию меняет install\release.ps1, руками править не нужно."
    [IO.File]::WriteAllLines($versionFile, $lines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output ("VERSION_WRITTEN " + $newVersion)
}

Write-State $newVersion $map $digest
Write-Output "STATE_WRITTEN install\release.state"

$entry = @()
$entry += ("## " + $newVersion + " " + [char]0x2014 + " " + (Get-Date -Format "dd.MM.yyyy"))
$entry += ""
if ($changedFiles.Count -gt 0) { $entry += ("- Изменены исходники: " + ($changedFiles -join ", ")) }
else { $entry += "- Исходники не менялись, переустановка в книгу" }
if ($applied.Count -gt 0) { $entry += ("- Миграции: " + ($applied -join ", ")) }
$entry += ("- Книга: " + $full)
$entry += ("- Хеш исходников: " + $digest)
$entry += ""

if (Test-Path $changelog) {
    $old = Get-Content -LiteralPath $changelog -Encoding UTF8
    $idx = -1
    for ($i = 0; $i -lt $old.Count; $i++) { if ($old[$i] -like "## *") { $idx = $i; break } }
    if ($idx -lt 0) { $new = $old + "" + $entry }
    elseif ($idx -eq 0) { $new = ($entry + $old) }
    else { $new = ($old[0..($idx - 1)] + $entry + $old[$idx..($old.Count - 1)]) }
} else {
    $head = @("# Журнал версий ReportMTO", "", "Формат: Keep a Changelog, новая запись сверху. Файл ведёт install\release.ps1.", "")
    $new = $head + $entry
}
[IO.File]::WriteAllLines($changelog, $new, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "CHANGELOG_WRITTEN CHANGELOG.md"

Write-Output ("RELEASE_OK " + $newVersion + " - книга " + $full)
exit 0

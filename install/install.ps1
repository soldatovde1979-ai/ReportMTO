# install.ps1
#
# Единая точка установки И релиза ReportMTO (объединение install + release, ТЗ v1.0).
#
# История установщика:
#   Version 1.0 / 2026-09-08 ... Version 1.10 / 2026-09-13 (см. git-историю):
#   v1.10 - список VBA-модулей больше не захардкожен: устанавливаются ВСЕ *.bas
#           из src\vba (как уже сделано для *.pq); тот же список в пост-верификации.
#   v1.2  - обновление Power Query-запросов из src\powerquery (qDiagImport не ставится,
#           prmSourcePath - рабочее состояние - не трогается).
#   v1.3  - бэкапы в отдельную bak\: книги внутри ProjectRoot -> <ProjectRoot>\bak,
#           книги вне - bak\ рядом с книгой.
#   v1.4  - по-модульный отчёт VBA/PQ/TPL и обязательная пост-верификация; fail -> exit 1.
#   v1.6  - русские пояснения; файл UTF-8 с BOM (PowerShell 5.1).
#   v1.7  - сравнение VBA без Attribute-строк, PQ с trim хвостовых пробелов.
#
# История релиза (был install\release.ps1, встроен сюда):
#   v1.0 / 11.09.2026 - единая точка раскатки с версионированием.
#   v1.1 / 11.09.2026 - защита от отката (книга новее исходников -> RELEASE_FAIL,
#           осознанный откат -Force, BOOK_MD5_MISMATCH при равных версиях).
#
# Version 2.0 / 2026-09-13 - ОБЪЕДИНЕНИЕ (ТЗ v1.0, задача 5):
#   - релизный поток (release.ps1) и обёртка install_prod.ps1 встроены в этот файл;
#   - флаг -InstallOnly: только установка + верификация (для build-книг и повторных
#     установок), без версии/миграций/CHANGELOG;
#   - документация (README.md, docs\**\*.md, install\*.md) включена в дайджест исходников:
#     правки документации поднимают патч; отдельной строкой в CHANGELOG не пишутся;
#   - запись в CHANGELOG не дублируется при неизменных исходниках (RELEASE_SOURCES_SAME);
#   - install\release.ps1 и install\install_prod.ps1 перенесены в tools\archive\.
#
# Режимы:
#   1) Полный релиз (по умолчанию, -Target - ОДНА книга):
#        дайджест исходников -> защита от отката -> установка -> компиляция ->
#        версия + миграции -> VERSION/release.state/CHANGELOG.
#      Версия фиксируется ТОЛЬКО при нулевом коде установки и компиляции.
#   2) -InstallOnly (установка без версионирования; -Target - книга ИЛИ папка):
#        бэкап -> VBA -> PQ -> шаблон -> верификация -> exit с кодом.
#
# Использование (из корня проекта, книга должна быть закрыта):
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -DryRun
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Bump minor
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target ".\build" -InstallOnly
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target ".\build\ReportMTO v7.0.xlsm" -InstallOnly
#
# Файл UTF-8 с BOM: PowerShell 5.1 иначе читает кириллицу как ANSI.

param(
    [string]$Target = ".\ReportMTO.xlsm",
    [string]$ProjectRoot = "",
    [ValidateSet("patch", "minor", "major")]
    [string]$Bump = "patch",
    [switch]$DryRun,
    [switch]$SkipCompile,
    [switch]$Force,
    [switch]$InstallOnly
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($ProjectRoot -eq "") { $ProjectRoot = Split-Path -Parent $scriptRoot }
if (-not [System.IO.Path]::IsPathRooted($Target)) {
    $Target = [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot $Target))
} else {
    $Target = [System.IO.Path]::GetFullPath($Target)
}
Write-Output "TARGET $Target - цель установки (полный путь)"

$vbaDir      = Join-Path $ProjectRoot "src\vba"
$pqDir       = Join-Path $ProjectRoot "src\powerquery"
$tplSrc      = Join-Path $ProjectRoot "tmp_index.html"
$versionFile = Join-Path $scriptRoot "VERSION"
$stateFile   = Join-Path $scriptRoot "release.state"
$migDir      = Join-Path $scriptRoot "migrations"
$changelog   = Join-Path $ProjectRoot "CHANGELOG.md"

if (-not (Test-Path $tplSrc)) { Write-Output "NO_TEMPLATE: $tplSrc - нет шаблона tmp_index.html в корне проекта"; exit 1 }

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
# v2.0: в дайджест включена документация (README.md, docs\**\*.md, install\*.md):
# правки документации поднимают патч; отдельной строкой в CHANGELOG не пишутся.
# Служебные файлы (VERSION, release.state) - не .md и в дайджест не входят.
function Get-SourceMap() {
    $map = New-Object System.Collections.Specialized.OrderedDictionary
    $files = @()
    $files += Get-ChildItem -Path $vbaDir -Filter *.bas -File | Sort-Object Name
    $files += Get-ChildItem -Path $pqDir -Filter *.pq -File | Sort-Object Name
    if (Test-Path $tplSrc) { $files += Get-Item -LiteralPath $tplSrc }
    $readme = Join-Path $ProjectRoot "README.md"
    if (Test-Path $readme) { $files += Get-Item -LiteralPath $readme }
    $docsDir = Join-Path $ProjectRoot "docs"
    if (Test-Path $docsDir) { $files += Get-ChildItem -Path $docsDir -Filter *.md -File -Recurse | Sort-Object FullName }
    $files += Get-ChildItem -Path $scriptRoot -Filter *.md -File | Sort-Object Name
    foreach ($f in $files) {
        $h = (Get-FileHash -Algorithm MD5 -LiteralPath $f.FullName).Hash
        $key = $f.FullName.Substring($ProjectRoot.Length + 1).Replace("\", "/")
        $map[$key] = $h
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
    $lines += "# Служебный файл install\install.ps1 - состояние последнего успешного релиза."
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

# Читает BUILD/VERSION и BUILD/SRC_MD5 из книги ДО установки: только так можно
# поймать откат (на этой машине исходники старее, чем то, что уже стоит в книге).
# Книга открывается только на чтение. Сообщения внутри не печатаются - иначе они
# попали бы в возвращаемое значение вместе с хеш-таблицей.
function Read-BookBuildInfo([string]$path) {
    $info = @{ Version = "0.0.0"; Md5 = ""; Warn = "" }
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    try {
        $wb = $xl.Workbooks.Open($path, 0, $true)
        try {
            $ws = $null
            foreach ($sh in $wb.Worksheets) { if ($sh.Name -eq "Variable") { $ws = $sh } }
            if ($null -eq $ws) { throw "нет листа Variable" }
            $lo = $null
            foreach ($l in $ws.ListObjects) { if ($l.Name -eq "tblVariable") { $lo = $l } }
            if ($null -eq $lo) { throw "нет таблицы tblVariable" }
            $v = (Get-VariableKey $lo "BUILD/VERSION").Trim()
            if ($v -match "^\d+\.\d+\.\d+$") { $info.Version = $v }
            $info.Md5 = (Get-VariableKey $lo "BUILD/SRC_MD5").Trim()
        } finally {
            $wb.Close($false)
        }
    } catch {
        $info.Warn = $_.Exception.Message
    } finally {
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    }
    return $info
}

# Normalize text for comparison: one EOL style, no trailing spaces per line,
# no trailing line breaks. VBE re-encodes on import, so a raw compare would
# report a false DIFF on every run.
function Normalize-Text([string]$text) {
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`r", "`n"
    $lines = $text -split "`n"
    $out = ""
    foreach ($line in $lines) { $out += $line.TrimEnd() + "`n" }
    return $out.TrimEnd("`n")
}

# Print where two texts first differ and the char codes there, so the log
# shows WHY verification failed: 10/13 = line endings, 65279 = BOM, 63 = '?'
# (lost encoding), 32 = space, 9 = tab.
function Show-DiffDetail([string]$src, [string]$cur, [string]$label) {
    $n = [Math]::Min($src.Length, $cur.Length)
    $i = 0
    while ($i -lt $n -and $src[$i] -eq $cur[$i]) { $i++ }
    $out = "  " + $label + " - первое отличие на позиции " + $i + " (длина исходника " + $src.Length + ", длина в книге " + $cur.Length + ")"
    if ($i -lt $src.Length) { $out += ", исходник: код символа " + [int][char]$src[$i] }
    if ($i -lt $cur.Length) { $out += ", в книге: код символа " + [int][char]$cur[$i] }
    Write-Output $out
}

# Drop "Attribute VB_Name = ..." lines from a .bas source before comparing:
# CodeModule.Lines never returns attribute lines, so they are the only
# difference between an imported module and its source file.
function Remove-AttributeLines([string]$text) {
    return $text -replace '(?m)^[ \t]*Attribute VB_Name[ \t]*=[^\r\n]*(\r\n|\r|\n)?', ''
}

# ================================================================== РЕЛИЗНЫЙ ДАЙДЖЕСТ
# В режиме -InstallOnly пропускается целиком: установка ничего не фиксирует.
$sourcesChanged = $false
$changedFiles = @()
$curVersion = ""
$newVersion = ""
$digest = ""
$map = $null
$state = $null
$forced = $false
$book = $null

if (-not $InstallOnly) {
    $full = $Target
    if (-not (Test-Path $full)) { Fail ("книга не найдена: " + $full) }
    if ((Get-Item -LiteralPath $full).PSIsContainer) {
        Fail "режим релиза требует ОДНУ книгу (-Target файл). Для папки используйте -InstallOnly"
    }

    $curVersion = Read-BuildVersion
    $map = Get-SourceMap
    $digest = Get-MapDigest $map
    $state = Read-State

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

    # --- защита от отката: что уже стоит в книге против того, что лежит в исходниках
    $book = Read-BookBuildInfo $full
    if ($book.Warn -ne "") {
        Write-Output ("BOOK_READ_WARN " + $book.Warn + " - версию книги прочитать не удалось, проверка пропущена")
    } else {
        Write-Output ("BOOK_VERSION_BEFORE " + $book.Version + " - версия, которая уже стоит в книге")
        $cmp = Compare-Version $book.Version $curVersion
        if ($cmp -gt 0) {
            if (-not $Force) {
                $msg = "книга НОВЕЕ исходников: в книге " + $book.Version + ", в install\VERSION " + $curVersion
                $msg = $msg + " - сделайте git pull. Осознанный откат: тот же запуск с флагом -Force"
                Fail $msg
            }
            $forced = $true
            Write-Output ("FORCE_DOWNGRADE " + $book.Version + " -> " + $curVersion + " - откат по флагу -Force, записываем в журнал")
        } elseif ($cmp -eq 0 -and $book.Md5 -ne "" -and $book.Md5 -ne $digest) {
            Write-Output ("BOOK_MD5_MISMATCH " + $book.Md5 + " - в книге та же версия " + $book.Version + ", но другое содержимое: похоже, релиз делали с другой машины. Сверьте git перед установкой")
        }
    }

    if ($DryRun) {
        Write-Output "RELEASE_DRYRUN - книга не тронута, версия не записана"
        exit 0
    }
}

# ================================================================== УСТАНОВКА
# Resolve target to a list of workbooks: folder -> all *.xlsm inside, file -> itself.
$books = @()
$targetItem = Get-Item $Target -ErrorAction SilentlyContinue
if ($null -ne $targetItem) {
    if ($targetItem.PSIsContainer) {
        $books = @(Get-ChildItem -Path $Target -Filter *.xlsm -File)
    } else {
        $books = @($targetItem)
    }
}
if ($null -eq $targetItem) { Write-Output "NOT_FOUND: $Target - цель не найдена"; exit 1 }
if ($books.Count -eq 0) { Write-Output "NO_XLSM: $Target - в указанной папке нет xlsm-книг"; exit 1 }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ansi = [System.Text.Encoding]::GetEncoding(1251)
$tempDir = $env:TEMP
$patched = 0
$verifyFailed = 0
# Все *.bas из src\vba (v1.10) - единый список для установки и верификации.
$vbaMods = @(Get-ChildItem -Path $vbaDir -Filter *.bas -File | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($bookItem in $books) {
    $bookPath = $bookItem.FullName
    try {
        $bookDir = Split-Path -Parent $bookPath

        # Backup workbook and the old template (rollback = restore both). Backups go
        # to a dedicated bak\ folder (project rule): never next to the working book
        # and never into build\. Books inside ProjectRoot -> <ProjectRoot>\bak,
        # books outside -> bak\ next to the book.
        $bakDir = Join-Path $ProjectRoot "bak"
        $rootFull = [System.IO.Path]::GetFullPath($ProjectRoot)
        $bookFull = [System.IO.Path]::GetFullPath($bookPath)
        if (-not $bookFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            $bakDir = Join-Path $bookDir "bak"
        }
        New-Item -ItemType Directory -Force -Path $bakDir | Out-Null
        Copy-Item $bookPath (Join-Path $bakDir ($bookItem.Name + ".bak_" + $stamp)) -Force
        $oldTpl = Join-Path $bookDir "tmp_index.html"
        if (Test-Path $oldTpl) { Copy-Item $oldTpl (Join-Path $bakDir ("tmp_index.html.bak_" + $stamp)) -Force }

        $wb = $excel.Workbooks.Open($bookPath, 0, $false)
        $vp = $wb.VBProject
        $verifyFail = ""

        # VBA modules: add when missing, replace when different, skip when same.
        foreach ($mod in $vbaMods) {
            $src = Join-Path $vbaDir ($mod + ".bas")
            $utf8 = [System.IO.File]::ReadAllText($src)
            $ansiText = $ansi.GetString($ansi.GetBytes($utf8))
            $srcText = Normalize-Text (Remove-AttributeLines $ansiText)
            $curComp = $null
            try { $curComp = $vp.VBComponents.Item($mod) } catch { }
            $curText = ""
            if ($null -ne $curComp -and $curComp.CodeModule.CountOfLines -gt 0) {
                $curText = $curComp.CodeModule.Lines(1, $curComp.CodeModule.CountOfLines)
            }
            $curText = Normalize-Text $curText
            if ($null -eq $curComp) {
                $ansiTemp = Join-Path $tempDir ($mod + "_install_ansi.bas")
                [System.IO.File]::WriteAllText($ansiTemp, $utf8, $ansi)
                $vp.VBComponents.Import($ansiTemp) | Out-Null
                Write-Output "VBA_ADDED $mod - модуля не было, добавлен из исходника"
            } elseif ($curText -eq $srcText) {
                Write-Output "VBA_SAME $mod - совпадает с исходником, пропущен"
            } else {
                $vp.VBComponents.Remove($curComp)
                $ansiTemp = Join-Path $tempDir ($mod + "_install_ansi.bas")
                [System.IO.File]::WriteAllText($ansiTemp, $utf8, $ansi)
                $vp.VBComponents.Import($ansiTemp) | Out-Null
                Write-Output "VBA_DIFF $mod - различается, переустановлен из исходника"
                Show-DiffDetail $srcText $curText ("VBA " + $mod)
            }
        }

        # Refresh Power Query queries from src\powerquery (v1.2).
        foreach ($pqFile in Get-ChildItem -Path $pqDir -Filter *.pq) {
            $pqName = [System.IO.Path]::GetFileNameWithoutExtension($pqFile.Name)
            if ($pqName -eq "qDiagImport") { Write-Output "PQ_SKIP $pqName - диагностический запрос, не устанавливается"; continue }
            $pqFormula = [System.IO.File]::ReadAllText($pqFile.FullName)
            $qFound = $false
            $qFormula = $null
            foreach ($q in $wb.Queries) { if ($q.Name -eq $pqName) { $qFound = $true; $qFormula = $q.Formula } }
            if ($qFound) {
                if ($qFormula.TrimEnd() -ne $pqFormula.TrimEnd()) {
                    $wb.Queries.Item($pqName).Formula = $pqFormula
                    Write-Output "PQ_DIFF $pqName - формула различается, обновлена из исходника"
                    Show-DiffDetail $pqFormula $qFormula ("PQ " + $pqName)
                } else {
                    Write-Output "PQ_SAME $pqName - совпадает, пропущен"
                }
            } else {
                $wb.Queries.Add($pqName, $pqFormula) | Out-Null
                Write-Output "PQ_ADDED $pqName - запроса не было, добавлен из исходника"
            }
        }

        $wb.Save()

        # Mandatory verification: installed code/formulas must match the sources.
        foreach ($mod in $vbaMods) {
            $src = Join-Path $vbaDir ($mod + ".bas")
            $utf8 = [System.IO.File]::ReadAllText($src)
            $ansiText = $ansi.GetString($ansi.GetBytes($utf8))
            $srcText = Normalize-Text (Remove-AttributeLines $ansiText)
            $curComp = $null
            try { $curComp = $vp.VBComponents.Item($mod) } catch { }
            $curText = ""
            if ($null -ne $curComp -and $curComp.CodeModule.CountOfLines -gt 0) {
                $curText = $curComp.CodeModule.Lines(1, $curComp.CodeModule.CountOfLines)
            }
            $curText = Normalize-Text $curText
            if ($curText -eq $srcText) {
                Write-Output "VBA_VERIFY_OK $mod - код в книге совпал с исходником"
            } else {
                Write-Output "VBA_VERIFY_FAIL $mod - КОД В КНИГЕ НЕ СОВПАЛ С ИСХОДНИКОМ"
                Show-DiffDetail $srcText $curText ("VBA " + $mod)
                if ($verifyFail -ne "") { $verifyFail += ", " }
                $verifyFail += "VBA:$mod"
            }
        }
        foreach ($pqFile in Get-ChildItem -Path $pqDir -Filter *.pq) {
            $pqName = [System.IO.Path]::GetFileNameWithoutExtension($pqFile.Name)
            if ($pqName -eq "qDiagImport") { continue }
            $pqFormula = [System.IO.File]::ReadAllText($pqFile.FullName)
            $qFormula = $null
            foreach ($q in $wb.Queries) {
                if ($q.Name -eq $pqName) { $qFormula = $q.Formula }
            }
            if ($null -ne $qFormula -and $qFormula.TrimEnd() -eq $pqFormula.TrimEnd()) {
                Write-Output "PQ_VERIFY_OK $pqName - формула в книге совпала с исходником"
            } else {
                Write-Output "PQ_VERIFY_FAIL $pqName - ФОРМУЛА В КНИГЕ НЕ СОВПАЛА С ИСХОДНИКОМ"
                if ($null -eq $qFormula) {
                    Write-Output "  PQ $pqName - запрос отсутствует в книге"
                } else {
                    Show-DiffDetail $pqFormula $qFormula ("PQ " + $pqName)
                }
                if ($verifyFail -ne "") { $verifyFail += ", " }
                $verifyFail += "PQ:$pqName"
            }
        }

        $wb.Close($false)

        # Template next to the workbook (GenerateReport reads it by ThisWorkbook.Path).
        # Skip when source and destination are the same file (book in project root)
        # or when the file content is already identical.
        $tplSame = ($oldTpl -eq $tplSrc)
        if ((-not $tplSame) -and (Test-Path $oldTpl)) {
            $tplSame = ((Get-FileHash $oldTpl -Algorithm MD5).Hash -eq (Get-FileHash $tplSrc -Algorithm MD5).Hash)
        }
        if ($tplSame) {
            Write-Output "TPL_SAME tmp_index.html - совпадает, пропущен"
        } else {
            Copy-Item $tplSrc $oldTpl -Force
            Write-Output "TPL_DIFF tmp_index.html - различается, скопирован из исходника"
        }

        # Verify the template next to the workbook against the source.
        $tplHash = ""
        if (Test-Path $oldTpl) { $tplHash = (Get-FileHash $oldTpl -Algorithm MD5).Hash }
        $srcHash = (Get-FileHash $tplSrc -Algorithm MD5).Hash
        if ($tplHash -eq $srcHash) {
            Write-Output "TPL_VERIFY_OK tmp_index.html - шаблон рядом с книгой совпал с исходником"
        } else {
            Write-Output "TPL_VERIFY_FAIL tmp_index.html - ШАБЛОН НЕ СОВПАЛ С ИСХОДНИКОМ (MD5: книга=$tplHash, исходник=$srcHash)"
            if ($verifyFail -ne "") { $verifyFail += ", " }
            $verifyFail += "TPL:tmp_index.html"
        }

        $patched += 1
        if ($verifyFail -eq "") {
            Write-Output "VERIFY_OK $bookPath - ПРОВЕРКА ПРОЙДЕНА: все модули, запросы и шаблон совпали с исходниками"
        } else {
            Write-Output "VERIFY_FAIL $bookPath - ПРОВЕРКА НЕ ПРОЙДЕНА, расхождения: $verifyFail (подробности в строках *_VERIFY_FAIL выше)"
            $verifyFailed += 1
        }
        Write-Output "PATCHED $bookPath - книга обработана"
    } catch {
        Write-Output "FAILED $bookPath - ОШИБКА: $($_.Exception.Message)"
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "DONE $patched - обработано книг: $patched"

$installCode = 0
if (($patched -eq $books.Count) -and ($verifyFailed -eq 0)) { $installCode = 0 } else { $installCode = 1 }
Write-Output ("INSTALL_EXIT " + $installCode + " - код возврата установки")

if ($InstallOnly) {
    if ($installCode -eq 0) { Write-Output "PROD_OK - установка выполнена и подтверждена проверкой" }
    else { Write-Output "PROD_FAIL - проверка не подтвердила установку, см. строки VERIFY_FAIL выше" }
    exit $installCode
}

if ($installCode -ne 0) { Fail ("установщик вернул " + $installCode + " - версия не изменена") }

# ================================================================== КОМПИЛЯЦИЯ
if (-not $SkipCompile) {
    Write-Output "--- ШАГ 2: компиляция VBA"
    $compileOut = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ProjectRoot "tools\compile_check.ps1") -Book "$full"
    $compileOut | ForEach-Object { Write-Output $_ }
    $okLine = $compileOut | Where-Object { $_ -like "COMPILE_OK*" }
    if ($null -eq $okLine) { Fail "компиляция не прошла - версия не изменена" }
}

# ================================================================== КНИГА: ВЕРСИЯ И МИГРАЦИИ
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

# ================================================================== VERSION, state, CHANGELOG
if ($newVersion -ne $curVersion) {
    $desc = Read-VersionDescription
    $lines = @()
    $lines += $newVersion
    $lines += ("Собрано " + (Get-Date -Format "dd.MM.yyyy") + ". Изменены: " + ($changedFiles -join ", "))
    $lines += ""
    $lines += "Первая строка этого файла - версия сборки, всё ниже - описание."
    $lines += "Версию меняет install\install.ps1, руками править не нужно."
    [IO.File]::WriteAllLines($versionFile, $lines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output ("VERSION_WRITTEN " + $newVersion)
}

Write-State $newVersion $map $digest
Write-Output "STATE_WRITTEN install\release.state"

# v2.0: при неизменных исходниках запись в CHANGELOG не дублируется (двойной
# «8.1.3» от 12.09.2026). Запись пишется только при изменении исходников,
# осознанном откате (-Force) или применённых миграциях.
if (-not $sourcesChanged -and -not $forced -and ($applied.Count -eq 0)) {
    Write-Output "CHANGELOG_SKIP - исходники не менялись, запись не дублируется"
} else {
    $entry = @()
    $entry += ("## " + $newVersion + " " + [char]0x2014 + " " + (Get-Date -Format "dd.MM.yyyy"))
    $entry += ""
    # v3.0: строки «Изменены исходники/документация» убраны — их даёт git;
    # CHANGELOG оставляет только факты раскатки (миграции, -Force, книга, хеш).
    if ($applied.Count -gt 0) { $entry += ("- Миграции: " + ($applied -join ", ")) }
    if ($forced) { $entry += ("- ВНИМАНИЕ: установлено с -Force поверх более новой книги " + $book.Version) }
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
        $head = @("# Журнал версий ReportMTO", "", "Формат: Keep a Changelog, новая запись сверху. Файл ведёт install\install.ps1.", "")
        $new = $head + $entry
    }
    [IO.File]::WriteAllLines($changelog, $new, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "CHANGELOG_WRITTEN CHANGELOG.md"
}

Write-Output ("RELEASE_OK " + $newVersion + " - книга " + $full)
exit 0

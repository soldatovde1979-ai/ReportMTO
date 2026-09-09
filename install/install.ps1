# install.ps1
#
# Version 1.0 / 2026-09-08
# Version 1.1 / 2026-09-08: skip template copy when the book already sits next to
#   tmp_index.html in the project root (Copy-Item cannot overwrite itself).
# Version 1.3 / 2026-09-08: backups go to a dedicated bak\ folder instead of next
#   to the workbook (project rule docs/rules.md: never beside the working book and
#   never into build\, build is handed over as a whole). Books inside ProjectRoot
#   -> <ProjectRoot>\bak, books outside -> bak\ next to the book.
# Version 1.2 / 2026-09-08: also refresh Power Query queries from src\powerquery.
#   Before v1.2 the patch touched VBA only, so changed M-code (fnNormalizeFields v6)
#   never reached the production book and report generation kept failing. Rules:
#   existing queries are updated when the formula differs from the .pq file;
#   missing pipeline queries (fn*, Query-ImportJSON, qExistingData) are ADDED;
#   qDiagImport is never added (diagnostics only); prmSourcePath (runtime state,
#   holds the path picked by the user) is never touched.
# Version 1.4 / 2026-09-09: per-module report for VBA/PQ/template (VBA_ADDED,
#   VBA_DIFF, VBA_SAME, PQ_ADDED, PQ_DIFF, PQ_SAME, PQ_SKIP qDiagImport,
#   TPL_DIFF, TPL_SAME) and mandatory post-install verification against the
#   sources (VBA/PQ/TPL_VERIFY_OK|FAIL, VERIFY_OK|FAIL); verify fail -> exit 1.
# Version 1.5 / 2026-09-09: print the resolved full path of the target
#   (TARGET marker) so the report is unambiguous.
# Version 1.6 / 2026-09-09: human-readable Russian messages after every marker
#   and diff details (position + char codes) on mismatch; file saved as UTF-8
#   with BOM (required by PowerShell 5.1 for cyrillic strings).
# Version 1.7 / 2026-09-09: fix false mismatches found by diagnostics:
#   VBA compare drops "Attribute VB_Name = ..." lines (CodeModule.Lines never
#   returns them), PQ compare trims trailing whitespace (Excel drops the final
#   line break of a stored formula on Save).
# Version 1.8 / 2026-09-09: set console output to UTF-8 so the Russian
#   messages are readable in the terminal.
# Applies the v7.1 patch to existing ReportMTO workbooks in place (Path A):
#   - replaces VBA modules modMain, modContentMTO and modAggregate (Remove + Import,
#     source re-encoded UTF-8 -> ANSI 1251, as VBE reads .bas only as ANSI);
#   - copies the v2.0 template tmp_index.html next to each patched workbook;
#   - makes timestamped backups of the workbook and the old template first.
# Data (tbDATA), Variable sheet, buttons and Power Query connections are untouched.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1
#       -> patches every *.xlsm in .\build (default target)
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target .\path.xlsm
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install.ps1 -Target .\some\folder
#
# Script is UTF-8 with BOM: PowerShell 5.1 reads .ps1 without BOM as ANSI, so
# the BOM is required for the Russian messages in this script.

param(
    [string]$Target = "build",
    [string]$ProjectRoot = ""
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

$vbaDir = Join-Path $ProjectRoot "src\vba"
$tplSrc = Join-Path $ProjectRoot "tmp_index.html"

if (-not (Test-Path $tplSrc)) { Write-Output "NO_TEMPLATE: $tplSrc - нет шаблона tmp_index.html в корне проекта"; exit 1 }

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
        foreach ($mod in @("modMain", "modContentMTO", "modAggregate")) {
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
        $pqDir = Join-Path $ProjectRoot "src\powerquery"
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
        foreach ($mod in @("modMain", "modContentMTO", "modAggregate")) {
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
if (($patched -eq $books.Count) -and ($verifyFailed -eq 0)) { exit 0 } else { exit 1 }

# run_report.ps1
# Version 1.0 / 2026-09-18
#
# Запуск книги и формирование отчета (исполнитель run.bat из корня проекта).
#
# Открывает книгу в СОБСТВЕННОМ COM-инстансе Excel (НЕ GetActiveObject - правило
# docs\rules.md: при нескольких EXCEL.EXE цепляние к первому инстансу падает),
# выполняет modContentMTO.BuildPivots + modMain.GenerateReport, контролирует,
# что в result\ появился свежий Report_*.html, сохраняет книгу (если она не
# занята другим процессом), закрывает книгу и Excel. Чужие EXCEL.EXE не трогает.
#
# Использование (из корня проекта):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\run_report.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\run_report.ps1 -Book "C:\...\ReportMTO.xlsm"
#
# Файл UTF-8 с BOM: PowerShell 5.1 иначе читает кириллицу как ANSI.

param(
    [string]$Book = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot

function Say([string]$s) { Write-Output ((Get-Date -Format "HH:mm:ss") + "  " + $s) }

# ---------------------------------------------------------------- поиск книги
if ($Book -eq "") {
    $candidate = Join-Path $projectRoot "ReportMTO.xlsm"
    if (Test-Path -LiteralPath $candidate) {
        $Book = $candidate
        Say ("Книга: корневая " + $Book)
    } else {
        $latest = Get-ChildItem -LiteralPath (Join-Path $projectRoot "build") -Filter "ReportMTO*.xlsm" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -eq $latest) {
            Say "RUN_FAIL книга не найдена: нет ReportMTO.xlsm в корне проекта и нет build\ReportMTO*.xlsm"
            exit 1
        }
        $Book = $latest.FullName
        Say ("Книга: " + $Book + " (корневой нет, взята самая свежая из build\)")
    }
}

try {
    $Book = (Resolve-Path -LiteralPath $Book).Path
} catch {
    Say ("RUN_FAIL книга не найдена: " + $Book + " (" + $_.Exception.Message + ")")
    exit 1
}
Say ("BOOK " + $Book)

$resultDir = Join-Path $projectRoot "result"
if (-not (Test-Path -LiteralPath $resultDir)) {
    Say "RUN_FAIL нет папки result\ в корне проекта"
    exit 1
}

# lock-файл рядом с книгой = книга, вероятно, открыта в другом окне Excel.
# Отчет все равно соберется, но Save упадет - предупреждаем заранее.
$lock = Join-Path (Split-Path -Parent $Book) ("~$" + [System.IO.Path]::GetFileName($Book))
if (Test-Path -LiteralPath $lock) {
    Say "ВНИМАНИЕ: рядом с книгой есть lock-файл - книга, вероятно, открыта в другом окне Excel."
    Say "Отчет будет собран, но сохранить книгу не удастся."
}

Say "RUN_START"
$startTime = Get-Date

$excel = $null
$wb = $null

try {
    Say "Создаю собственный COM-инстанс Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1   # msoAutomationSecurityLow: макросы неподписанной книги

    Say "Открываю книгу (UpdateLinks=0)..."
    $wb = $excel.Workbooks.Open($Book, 0, $false)

    Say "BuildPivots"
    $null = $excel.Run("modContentMTO.BuildPivots")
    Say "  BuildPivots ok"

    Say "GenerateReport (это минуты)"
    $null = $excel.Run("modMain.GenerateReport")
    Say "  GenerateReport returned"

    $fresh = Get-ChildItem -LiteralPath $resultDir -Filter "Report_*.html" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $fresh) {
        Say "RUN_FAIL в result\ нет свежего Report_*.html за время прогона"
    } else {
        Say ("RESULT_FILE " + $fresh.FullName + " | " + [int]($fresh.Length / 1024) + " KB")
    }

    try {
        $wb.Save()
        Say "SAVED"
    } catch {
        Say "ВНИМАНИЕ: сохранить книгу не удалось (вероятно, она открыта в другом окне Excel): " + $_.Exception.Message
    }

    if ($null -eq $fresh) { exit 1 }
    Say "RUN_OK"
} catch {
    Say ("RUN_FAIL " + $_.Exception.Message)
    exit 1
} finally {
    # Close без аргументов сохранения: Save уже сделан выше (или не удался - тогда
    # повторно не сохраняем, чтобы не получить диалог). Close обернут в try/catch:
    # книга могла быть закрыта человеком вручную (Visible = true).
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

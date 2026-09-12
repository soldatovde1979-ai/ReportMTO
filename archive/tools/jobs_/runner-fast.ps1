# runner-fast.ps1
# Version 1.0 / 11.09.2026
#
# Быстрая очередь: tools\jobs\fast-queue -> tools\jobs\fast-out.
# Отдельная задача планировщика ReportMTO-Runner-Fast, раз в минуту.
# Смысл: длинная задача (загрузка, отчёт) блокирует свою очередь целиком -
# планировщик не запускает вторую копию одной задачи. Эта очередь живёт
# независимо и умеет ровно то, что нужно в такой момент: спросить состояние,
# снять снимок экрана, поставить флаг мягкой остановки, аварийно снять Excel.
#
# Тяжёлые задачи здесь запрещены намеренно: они снова всё заблокируют.
# Файл UTF-8 с BOM.

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$jobsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent (Split-Path -Parent $jobsRoot)
. (Join-Path $jobsRoot "jobs-lib.ps1")

$queue = Join-Path $jobsRoot "fast-queue"
$outDir = Join-Path $jobsRoot "fast-out"
$state = Join-Path $jobsRoot "state"
$processed = Join-Path $state "processed-fast"
foreach ($d in @($queue, $outDir, $state, $processed)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$runnerLog = Join-Path $state "runner-fast.log"
$lockFile = Join-Path $state "runner-fast.lock"

# Пульс быстрой очереди пишется ВСЕГДА: он и есть честный признак «помощник жив»,
# в отличие от пульса длинной очереди, который замирает на время работы.
Set-Content -LiteralPath (Join-Path $state "heartbeat-fast.txt") -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ok") -Encoding UTF8

if (Test-Path $lockFile) {
    $age = (Get-Date) - (Get-Item $lockFile).LastWriteTime
    if ($age.TotalMinutes -lt 10) { exit 0 }
    Remove-Item -LiteralPath $lockFile -Force
}

$jobs = @(Get-ChildItem -Path $queue -Filter *.job -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($jobs.Count -eq 0) { exit 0 }

Set-Content -LiteralPath $lockFile -Value ("PID " + $PID + " " + (Get-Date -Format "HH:mm:ss")) -Encoding UTF8

# Только короткое и аварийное.
$allowed = @{
    "ping"       = "^$"
    "stop"       = "^$"
    "kill"       = "^\s*confirm\s*$"
    "diag"       = "^\s*(rdp|env|git|session|all)?\s*$"
    "screenshot" = "^$"
    "wakescreen" = "^$"
}

try {
    Invoke-JobQueue $jobs $queue $outDir $processed $allowed $runnerLog
} finally {
    if (Test-Path $lockFile) { Remove-Item -LiteralPath $lockFile -Force }
}

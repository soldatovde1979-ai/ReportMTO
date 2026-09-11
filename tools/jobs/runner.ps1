# runner.ps1
# Version 3.0 / 11.09.2026: код задач вынесен в jobs-lib.ps1, здесь осталась
#   только длинная очередь. Быстрая - в runner-fast.ps1, отдельной задачей
#   планировщика: пока идёт длинная задача, вторая копия ЭТОЙ задачи не
#   запускается, и без второй очереди помощник глух (загрузка 11.09 шла 5 часов).
# Ранние версии 1.0-2.0 - в истории git.
#
# Длинная очередь: tools\jobs\queue -> tools\jobs\out.
# Запускается задачей планировщика ReportMTO-Runner раз в 2 минуты.
# Файл UTF-8 с BOM.

param([switch]$Once)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$jobsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent (Split-Path -Parent $jobsRoot)
. (Join-Path $jobsRoot "jobs-lib.ps1")

$queue = Join-Path $jobsRoot "queue"
$outDir = Join-Path $jobsRoot "out"
$state = Join-Path $jobsRoot "state"
$processed = Join-Path $state "processed"
foreach ($d in @($queue, $outDir, $state, $processed)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$runnerLog = Join-Path $state "runner.log"
$lockFile = Join-Path $state "runner.lock"

Set-Content -LiteralPath (Join-Path $state "heartbeat.txt") -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ok") -Encoding UTF8

if (Test-Path $lockFile) {
    $age = (Get-Date) - (Get-Item $lockFile).LastWriteTime
    if ($age.TotalMinutes -lt 360) { exit 0 }
    Add-Content -LiteralPath $runnerLog -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  Замок старше 6 часов - снимаю") -Encoding UTF8
    Remove-Item -LiteralPath $lockFile -Force
}

$jobs = @(Get-ChildItem -Path $queue -Filter *.job -File -ErrorAction SilentlyContinue | Sort-Object Name)
if ($jobs.Count -eq 0) { exit 0 }

Set-Content -LiteralPath $lockFile -Value ("PID " + $PID + " " + (Get-Date -Format "HH:mm:ss")) -Encoding UTF8

# Длинные задачи: работа с книгой и репозиторием.
$allowed = @{
    "gitsync"    = "^$"
    "release"    = "^(\s*(-DryRun|-Force|-SkipCompile|-Bump\s+(patch|minor|major)))*\s*$"
    "compile"    = "^$"
    "report"     = "^$"
    "load"       = "^\s*[A-Za-z0-9_\-.]{0,60}\s*$"
    "closeexcel" = "^$"
    "aikey"      = "^$"
    "diag"       = "^\s*(rdp|env|git|session|all)?\s*$"
    "installfast" = "^$"
    "rebuild"     = "^$"
}

try {
    Invoke-JobQueue $jobs $queue $outDir $processed $allowed $runnerLog
} finally {
    if (Test-Path $lockFile) { Remove-Item -LiteralPath $lockFile -Force }
}

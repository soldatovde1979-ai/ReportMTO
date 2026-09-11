$ErrorActionPreference = 'Stop'
$tail = Get-Content 'ReportMTO_log.txt' -Encoding UTF8 | Select-Object -Last 8
[IO.File]::WriteAllLines((Join-Path 'tools' 'log_tail_now.txt'), $tail, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ('tail written: ' + $tail.Count + ' lines')

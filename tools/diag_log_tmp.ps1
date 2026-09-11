$ErrorActionPreference = 'Stop'
$log = (Resolve-Path 'ReportMTO_log.txt').Path
$all = Get-Content $log -Encoding UTF8
Write-Output ('LINES TOTAL: ' + $all.Count)

$tail = $all | Select-Object -Last 60
[IO.File]::WriteAllLines((Join-Path 'tools' 'log_tail_tmp.txt'), $tail, (New-Object System.Text.UTF8Encoding($false)))
Write-Output 'TAIL WRITTEN to tools\log_tail_tmp.txt'

Write-Output '---REPORT FILES REFS (last 10)---'
$m = $all | Select-String -Pattern 'Report_[0-9]+_[0-9]+\.html' | Select-Object -Last 10
foreach ($x in $m) { Write-Output ('L' + $x.LineNumber + ': ' + $x.Line.Substring(0, [Math]::Min(120, $x.Line.Length))) }

Write-Output '---ERROR-LIKE REFS (last 10)---'
$e = $all | Select-String -Pattern 'Err|error|fail|Failed|-2147' | Select-Object -Last 10
foreach ($x in $e) { Write-Output ('L' + $x.LineNumber + ': ' + $x.Line.Substring(0, [Math]::Min(140, $x.Line.Length))) }

$ErrorActionPreference = 'Stop'

Write-Output '---ROOT BOOKS AND LOGS---'
foreach ($p in @('ReportMTO.xlsm', 'ReportMTO_log.txt', 'tmp_index.html', 'build\ReportMTO v7.0.xlsm', 'build\ReportMTO_log.txt', 'build\tmp_index.html')) {
    if (Test-Path $p) {
        $f = Get-Item $p
        Write-Output ($p + ' | ' + $f.Length + ' | ' + $f.LastWriteTime)
    } else {
        Write-Output ($p + ' | MISSING')
    }
}

Write-Output '---ROOT LOG TAIL (ReportMTO_log.txt)---'
if (Test-Path 'ReportMTO_log.txt') {
    Get-Content 'ReportMTO_log.txt' -Tail 15 | ForEach-Object { Write-Output $_ }
}

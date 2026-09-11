$ErrorActionPreference = 'Stop'

Write-Output '---EXCEL PROCS---'
$ps = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($ps) {
    $ps | Select-Object Id, MainWindowTitle | Format-Table -AutoSize | Out-String | Write-Output
} else {
    Write-Output 'NONE'
}

Write-Output '---FILES---'
Write-Output ('book: ' + (Test-Path 'build\ReportMTO v7.0.xlsm'))
Write-Output ('tpl:  ' + (Test-Path 'build\tmp_index.html'))
Write-Output ('log:  ' + (Test-Path 'build\ReportMTO_log.txt'))
Write-Output ('result dir: ' + (Test-Path 'result'))

Write-Output '---RESULT DIR---'
Get-ChildItem result -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String | Write-Output

Write-Output '---DATA DIR---'
Get-ChildItem data -ErrorAction SilentlyContinue | Select-Object Name, Length | Format-Table -AutoSize | Out-String | Write-Output

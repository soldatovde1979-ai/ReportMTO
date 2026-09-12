$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path bak | Out-Null
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$bk = Join-Path 'bak' ("ReportMTO_v7.0_" + $ts + ".xlsm")
$bt = Join-Path 'bak' ("tmp_index_" + $ts + ".html")
if (-not (Test-Path $bk)) { Copy-Item 'build\ReportMTO v7.0.xlsm' $bk -Force }
if (-not (Test-Path $bt)) { Copy-Item 'build\tmp_index.html' $bt -Force }
Write-Output ('BACKUP: ' + $bk)
Write-Output ('BACKUP: ' + $bt)
Get-ChildItem bak | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize | Out-String | Write-Output

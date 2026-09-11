# load_prod_tmp.ps1
# Temporary one-shot loader: loads data\sppr_tablet_20260908_113800_60924rec.json
# into the production book (upsert via the standard pipeline) and rebuilds pivots.
# Backs the book up into bak\ first (project rule). ASCII-only literals.
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\load_prod_tmp.ps1

$ErrorActionPreference = "Stop"
$root = (Resolve-Path ".").Path
$bookPath = (Resolve-Path "ReportMTO.xlsm").Path
$jsonPath = (Resolve-Path "data\sppr_tablet_20260908_113800_60924rec.json").Path

# Backup into bak\ (outside git), timestamped.
$bakDir = Join-Path $root "bak"
New-Item -ItemType Directory -Force -Path $bakDir | Out-Null
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
Copy-Item $bookPath (Join-Path $bakDir ("ReportMTO.xlsm.bak_" + $stamp)) -Force
Write-Output ("BACKUP bak\ReportMTO.xlsm.bak_" + $stamp)

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$job = Start-Job -ScriptBlock {
    param($sec, $ids)
    Start-Sleep -Seconds $sec
    foreach ($p in Get-Process EXCEL -ErrorAction SilentlyContinue) {
        if ($ids -notcontains $p.Id) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
} -ArgumentList 1800, $before

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($bookPath, 0, $false)

$lo = $wb.Sheets.Item("tbDATA").ListObjects.Item("tbDATA")
Write-Output ("ROWS_BEFORE=" + $lo.ListRows.Count)

$safe = $jsonPath -replace '"', '""'
$wb.Queries.Item("prmSourcePath").Formula = '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
$qt = $lo.QueryTable
$qt.BackgroundQuery = $false
$null = $qt.Refresh($false)
Write-Output ("ROWS_AFTER_REFRESH=" + $lo.ListRows.Count)

$null = $excel.Run("modContentMTO.BuildPivots")
$wb.Save()
Write-Output "PIVOTS_AND_SAVE_OK"

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Stop-Job $job -ErrorAction SilentlyContinue
Remove-Job $job -Force -ErrorAction SilentlyContinue
Write-Output "LOAD_DONE"

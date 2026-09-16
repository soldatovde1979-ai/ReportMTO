$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$tempDir = $env:TEMP
$workBook = Join-Path $tempDir "ReportMTO_check.xlsm"
$json1 = (Resolve-Path "tests\test_sppr_tablet_v1.json").Path
$e2eDataDir = Join-Path $tempDir "ReportMTO_e2e_data"

if (Test-Path $e2eDataDir) { Remove-Item $e2eDataDir -Recurse -Force }
New-Item -ItemType Directory -Path $e2eDataDir | Out-Null
Copy-Item $json1 $e2eDataDir -Force

Copy-Item "build\ReportMTO v7.0.xlsm" $workBook -Force
$wb = $excel.Workbooks.Open($workBook, 0, $false)

$wsData = $wb.Sheets.Item("tbDATA")
$lo = $wsData.ListObjects.Item("tbDATA")
if ($lo.ListRows.Count -gt 0) { $null = $lo.DataBodyRange.Delete() }

$safe = $e2eDataDir -replace '"', '""'
$wb.Queries.Item("prmSourcePath").Formula = '"' + $safe + '" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
$qt = $lo.QueryTable
$qt.BackgroundQuery = $false
$null = $qt.Refresh($false)

Write-Host "ROWS:" $lo.ListRows.Count

$excel.Run("modContentMTO.BuildPivots")
$d = $excel.Run("modContentMTO.BuildPlaceholders", "S3", "S4", "S5")

Write-Host "---FACTS---"
Write-Host $d.Item('FACTS')
Write-Host "---BLOCK_TIME_HIST---"
Write-Host $d.Item('BLOCK_TIME_HIST')
Write-Host "---BLOCK_WEEKS_DGM---"
Write-Host $d.Item('BLOCK_WEEKS_DGM')
Write-Host "---BLOCK_PEOPLE_DGM---"
Write-Host $d.Item('BLOCK_PEOPLE_DGM')
Write-Host "---BLOCK_PEOPLE_DENT---"
Write-Host $d.Item('BLOCK_PEOPLE_DENT')
Write-Host "---BLOCK_SIGNSTAT_DGM---"
Write-Host $d.Item('BLOCK_SIGNSTAT_DGM')
Write-Host "---KPI_UNSIGNED---"
Write-Host $d.Item('KPI_UNSIGNED')
Write-Host "---BLOCK_UNSIGNED_SOURCE---"
Write-Host $d.Item('BLOCK_UNSIGNED_SOURCE')
Write-Host "---BLOCK_POSTS_DGM---"
Write-Host $d.Item('BLOCK_POSTS_DGM')
Write-Host "---BLOCK_POSTS_DENT---"
Write-Host $d.Item('BLOCK_POSTS_DENT')

$wb.Close($false)
$excel.Quit()

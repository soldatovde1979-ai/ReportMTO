$ErrorActionPreference = 'Stop'
$book = (Resolve-Path 'ReportMTO.xlsm').Path
$src = (Resolve-Path 'src\vba\modContentMTO.bas').Path
$ansi = Join-Path $env:TEMP 'modContentMTO_ansi.bas'

$t = [IO.File]::ReadAllText($src)
[IO.File]::WriteAllText($ansi, $t, [Text.Encoding]::GetEncoding(1251))
Write-Output ('ANSI copy: ' + $ansi + ' bytes=' + (Get-Item $ansi).Length)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
try {
    $wb = $excel.Workbooks.Open($book)
    $vbProj = $wb.VBProject
    $null = $vbProj.VBComponents.Count
    $comp = $vbProj.VBComponents.Item('modContentMTO')
    $vbProj.VBComponents.Remove($comp)
    $vbProj.VBComponents.Import($ansi) | Out-Null
    $wb.Save()
    $lines = $vbProj.VBComponents.Item('modContentMTO').CodeModule.CountOfLines
    Write-Output ('REIMPORT OK, lines=' + $lines)
    $wb.Close($false)
} catch {
    Write-Output ('REIMPORT FAILED: ' + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
    throw
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

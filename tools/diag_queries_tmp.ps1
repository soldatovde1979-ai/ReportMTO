$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
$book = Join-Path $root 'build\ReportMTO v7.0.xlsm'
$out  = Join-Path $root 'data\diag_queries.txt'
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('BOOK: ' + $book)
$procs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
[void]$sb.AppendLine('EXCEL processes: ' + @($procs).Count)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($book, 0, $true)
try {
    [void]$sb.AppendLine('QUERY COUNT: ' + $wb.Queries.Count)
    foreach ($q in $wb.Queries) {
        [void]$sb.AppendLine('=== QUERY: ' + $q.Name + ' ===')
        [void]$sb.AppendLine([string]$q.Formula)
    }
    [void]$sb.AppendLine('CONNECTION COUNT: ' + $wb.Connections.Count)
    foreach ($c in $wb.Connections) {
        [void]$sb.AppendLine('--- CONNECTION: ' + $c.Name + ' | type=' + $c.Type + ' ---')
        try {
            [void]$sb.AppendLine('OLEDB: ' + $c.OLEDBConnection.Connection)
        } catch {
            [void]$sb.AppendLine('OLEDB: n/a')
        }
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
}
[System.IO.File]::WriteAllText($out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
Write-Output ('SAVED ' + $out)

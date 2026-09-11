$ErrorActionPreference = 'Stop'
$book = (Resolve-Path 'ReportMTO.xlsm').Path
$out = (Resolve-Path '.').Path + '\tools\diag_book_tmp.txt'
$lines = New-Object System.Collections.Generic.List[string]

function AddLine($s) { $lines.Add([string]$s) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3

try {
    $wb = $excel.Workbooks.Open($book, $null, $true)

    AddLine ('SHEETS: ' + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ', '))

    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq 'Variable') {
            foreach ($lo in $ws.ListObjects) {
                AddLine ('VAR TABLE: ' + $lo.Name + ' rows=' + $lo.DataBodyRange.Rows.Count)
                $vals = $lo.DataBodyRange.Value2
                for ($i = 1; $i -le $lo.DataBodyRange.Rows.Count; $i++) {
                    $k = $vals.GetValue($i, 1)
                    $v = $vals.GetValue($i, 2)
                    if ($null -ne $k -and [string]$k -ne '') {
                        if ([string]$k -match 'KEY|SECRET|TOKEN') { $v = '<HIDDEN len=' + ([string]$v).Length + '>' }
                        AddLine ('VAR[' + $i + ']: ' + $k + ' = ' + $v)
                    }
                }
            }
        }
        if ($ws.Name -eq 'tbDATA') {
            foreach ($lo in $ws.ListObjects) {
                AddLine ('tbDATA TABLE: ' + $lo.Name + ' rows=' + $lo.DataBodyRange.Rows.Count + ' cols=' + $lo.DataBodyRange.Columns.Count)
            }
        }
        if ($ws.Name -eq 'Logs') {
            foreach ($lo in $ws.ListObjects) {
                AddLine ('LOGS TABLE: ' + $lo.Name + ' rows=' + $lo.DataBodyRange.Rows.Count)
                $n = $lo.DataBodyRange.Rows.Count
                if ($n -gt 0) {
                    $start = [Math]::Max(1, $n - 2)
                    $r = $ws.Range($ws.Cells($start + 1, 1), $ws.Cells($n + 1, 5))
                    $lv = $r.Value2
                    for ($i = 1; $i -le 3; $i++) {
                        AddLine ('LOGLAST[' + $i + ']: ' + $lv.GetValue($i,1) + ' | ' + $lv.GetValue($i,2) + ' | ' + $lv.GetValue($i,3) + ' | ' + $lv.GetValue($i,4) + ' | ' + $lv.GetValue($i,5))
                    }
                }
            }
        }
    }

    AddLine '---QUERIES---'
    foreach ($q in $wb.Queries) {
        $f = ''
        try { $f = [string]$q.Formula } catch { $f = '<ERR>' }
        if ($f.Length -gt 200) { $f = $f.Substring(0, 200) + '...' }
        AddLine ('Q: ' + $q.Name + ' => ' + $f)
    }

    AddLine '---CONNECTIONS---'
    foreach ($c in $wb.Connections) {
        AddLine ('C: ' + $c.Name + ' (' + $c.Type + ')')
    }

    $wb.Close($false)
} catch {
    AddLine ('EXCEPTION: ' + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

[IO.File]::WriteAllLines($out, $lines, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ('WROTE ' + $lines.Count + ' lines to ' + $out)

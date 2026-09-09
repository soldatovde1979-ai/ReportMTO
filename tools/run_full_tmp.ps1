$ErrorActionPreference = 'Stop'
$book = (Resolve-Path 'ReportMTO.xlsm').Path
$logPath = (Resolve-Path 'ReportMTO_log.txt').Path
$watchFile = (Resolve-Path '.').Path + '\tools\watch_full_tmp.txt'
$w = New-Object System.Collections.Generic.List[string]

function Watch($s) { $w.Add((Get-Date -Format 'HH:mm:ss') + ' | ' + $s) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    Watch ('Opening ' + $book)
    $wb = $excel.Workbooks.Open($book)
    Start-Sleep -Seconds 3

    $start = Get-Date
    $pidExcel = $excel.Hwnd

    $null = $excel.Application.OnTime([System.DateTime]::Now.AddSeconds(1), 'GenerateReport')
    Watch ('STARTED GenerateReport via OnTime at ' + $start.ToString('HH:mm:ss'))

    $wshell = New-Object -ComObject WScript.Shell
    $deadline = $start.AddMinutes(30)
    $done = $false
    $tick = 0

    while ((Get-Date) -lt $deadline -and -not $done) {
        Start-Sleep -Seconds 20
        $tick++
        $elapsed = [int]((Get-Date) - $start).TotalSeconds

        $fresh = Get-ChildItem 'result' -Filter 'Report_*.html' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt $start }
        if ($fresh) {
            Watch ('DONE at ' + $elapsed + 's: ' + ($fresh.Name -join ', '))
            $done = $true
        }

        $procAlive = (Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0
        if (-not $procAlive) { Watch ('EXCEL DIED at ' + $elapsed + 's'); $done = $true }

        if ($tick % 2 -eq 0 -or $done) {
            $tail = Get-Content $logPath -Encoding UTF8 | Select-Object -Last 2
            Watch ('t=' + $elapsed + 's tail: ' + ($tail -join ' /// '))
        }
    }

    if (-not $done) { Watch ('TIMEOUT after 30 min') }
    else {
        for ($i = 0; $i -lt 4; $i++) {
            $wshell.AppActivate($pidExcel) | Out-Null
            $wshell.SendKeys('{ENTER}')
            Start-Sleep -Milliseconds 800
        }
        Watch ('Saving and closing book')
        $wb.Save()
        $wb.Close($false)
    }
} catch {
    Watch ('EXCEPTION: ' + $_.Exception.Message)
} finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

[IO.File]::WriteAllLines($watchFile, $w, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ('WATCH WROTE ' + $w.Count + ' lines')

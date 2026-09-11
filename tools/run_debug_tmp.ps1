$ErrorActionPreference = 'Stop'
$book = (Resolve-Path 'ReportMTO.xlsm').Path
$logPath = (Resolve-Path 'ReportMTO_log.txt').Path
$watchFile = (Resolve-Path '.').Path + '\tools\watch_debug_tmp.txt'
$w = New-Object System.Collections.Generic.List[string]

function Watch($s) { $w.Add((Get-Date -Format 'HH:mm:ss') + ' | ' + $s) }

# Ключ ИИ: только факт наличия, без значения
$envKey = $env:AI_API_KEY
Watch ('AI_API_KEY env: ' + $(if ($envKey) { 'PRESENT len=' + $envKey.Length } else { 'ABSENT' }))
$keyFile = Join-Path $env:APPDATA 'ReportMTO\deepseek.key'
Watch ('deepseek.key file: ' + $(if (Test-Path $keyFile) { 'PRESENT len=' + (Get-Item $keyFile).Length } else { 'ABSENT' }))

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

try {
    Watch ('Opening ' + $book)
    $wb = $excel.Workbooks.Open($book)
    Start-Sleep -Seconds 3

    $start = Get-Date
    $before = (Get-Content $logPath -Encoding UTF8).Count
    $pidExcel = $excel.Hwnd

    $null = $excel.Application.OnTime([System.DateTime]::Now.AddSeconds(1), 'DebugGenerateOffline')
    Watch ('STARTED DebugGenerateOffline via OnTime at ' + $start.ToString('HH:mm:ss'))

    $wshell = New-Object -ComObject WScript.Shell
    $deadline = $start.AddMinutes(45)
    $done = $false
    $tick = 0

    while ((Get-Date) -lt $deadline -and -not $done) {
        Start-Sleep -Seconds 20
        $tick++
        $elapsed = [int]((Get-Date) - $start).TotalSeconds

        # Новый debug-отчёт в result после старта?
        $fresh = Get-ChildItem 'result' -Filter 'debug_*.html' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt $start }
        if ($fresh) {
            Watch ('DONE at ' + $elapsed + 's: ' + ($fresh.Name -join ', '))
            $done = $true
        }

        # Excel жив?
        $procAlive = (Get-Process -Id (Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -ErrorAction SilentlyContinue).Count -gt 0
        if (-not $procAlive) { Watch ('EXCEL DIED at ' + $elapsed + 's'); $done = $true }

        if ($tick % 3 -eq 0 -or $done) {
            $tail = Get-Content $logPath -Encoding UTF8 | Select-Object -Last 2
            Watch ('t=' + $elapsed + 's tail: ' + ($tail -join ' /// '))
        }
    }

    if (-not $done) { Watch ('TIMEOUT after 45 min') }
    else {
        # Закрыть MsgBox (появляется после записи «Сохранён»)
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

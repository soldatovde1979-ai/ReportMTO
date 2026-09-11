$ErrorActionPreference = 'Stop'
$keyPath = Join-Path $env:APPDATA 'ReportMTO\deepseek.key'
if (-not (Test-Path $keyPath)) { Write-Output 'KEY FILE MISSING'; exit 1 }
$key = (Get-Content $keyPath -Raw -Encoding UTF8).Trim()
Write-Output ('key len=' + $key.Length + ' prefix=' + $key.Substring(0, [Math]::Min(3, $key.Length)) + '...')

$bodyFile = Join-Path $env:TEMP 'ds_probe_body.json'
[IO.File]::WriteAllText($bodyFile, '{"model":"deepseek-chat","messages":[{"role":"user","content":"ping"}],"max_tokens":5}', (New-Object System.Text.ASCIIEncoding))

$resp = & curl.exe -s -m 40 -w "`nHTTP_CODE:%{http_code}" -X POST "https://api.deepseek.com/chat/completions" -H "Content-Type: application/json" -H ("Authorization: Bearer " + $key) --data-binary "@$bodyFile" 2>&1
Write-Output ('RESPONSE: ' + ($resp -join ' '))

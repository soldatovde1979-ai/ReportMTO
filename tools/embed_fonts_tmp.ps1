$ErrorActionPreference = 'Stop'
$marker = 'EMBEDDED FONTS'

function Get-B64($p) {
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path $p).Path)
    return [Convert]::ToBase64String($bytes)
}

$bRegular = Get-B64 'tools\fonts\segoeui.woff2'
$bSemibold = Get-B64 'tools\fonts\segoeuib.woff2'
$bConsolas = Get-B64 'tools\fonts\consola.woff2'

$css = "<style id='embedded-fonts'>`n" +
"/* EMBEDDED FONTS: Segoe UI / Segoe UI Semibold / Consolas (woff2 base64, 09.09.2026) */`n" +
"@font-face{font-family:'Segoe UI';font-style:normal;font-weight:400;src:url(data:font/woff2;base64,$bRegular) format('woff2')}`n" +
"@font-face{font-family:'Segoe UI Semibold';font-style:normal;font-weight:600;src:url(data:font/woff2;base64,$bSemibold) format('woff2')}`n" +
"@font-face{font-family:'Consolas';font-style:normal;font-weight:400;src:url(data:font/woff2;base64,$bConsolas) format('woff2')}`n" +
"</style>`n"

foreach ($path in @('tmp_index.html', 'build\tmp_index.html')) {
    $full = (Resolve-Path $path).Path
    $html = [IO.File]::ReadAllText($full)
    if ($html.Contains($marker)) {
        Write-Output ($path + ': marker present - skip (idempotent)')
        continue
    }
    $idx = $html.IndexOf('<style>')
    if ($idx -lt 0) { throw ('no <style> in ' + $path) }
    $html = $html.Substring(0, $idx) + $css + $html.Substring($idx)
    [IO.File]::WriteAllText($full, $html, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output ($path + ': embedded, size=' + (Get-Item $full).Length)
}

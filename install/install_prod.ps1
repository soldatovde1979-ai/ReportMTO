# install_prod.ps1
#
# Version 1.0 / 2026-09-08
# Runs install.ps1 against the production workbook in the project root.
# Default target: .\ReportMTO.xlsm (override with -Target <file>).
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install_prod.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install_prod.ps1 -Target ".\ReportMTO v7.0.xlsm"
#
# ASCII-only on purpose (PowerShell 5.1 ANSI parsing without BOM).

param(
    [string]$Target = ".\ReportMTO.xlsm"
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$install = Join-Path $scriptRoot "install.ps1"

& powershell -NoProfile -ExecutionPolicy Bypass -File $install -Target "$Target"
exit $LASTEXITCODE

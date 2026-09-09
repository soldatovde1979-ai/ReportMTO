# install_prod.ps1
#
# Version 1.0 / 2026-09-08
# Version 1.1 / 2026-09-09: print target before the run and exit code / result
#   after; per-module report and verification come from install.ps1.
# Version 1.2 / 2026-09-09: resolve and print the full path of the target
#   workbook (relative targets are resolved against the project root, same as
#   install.ps1) and pass the resolved path to install.ps1.
# Version 1.3 / 2026-09-09: human-readable Russian messages; UTF-8 with BOM.
# Version 1.4 / 2026-09-09: set console output to UTF-8 so the Russian
#   messages are readable in the terminal.
# Runs install.ps1 against the production workbook in the project root.
# Default target: .\ReportMTO.xlsm (override with -Target <file>).
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install_prod.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File install\install_prod.ps1 -Target ".\ReportMTO v7.0.xlsm"
#
# UTF-8 with BOM: required by PowerShell 5.1 for the Russian messages below.

param(
    [string]$Target = ".\ReportMTO.xlsm"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$install = Join-Path $scriptRoot "install.ps1"

# Resolve the target the same way install.ps1 does and print the full path,
# so the report shows exactly which file will be patched.
$fullTarget = $Target
if (-not [System.IO.Path]::IsPathRooted($fullTarget)) {
    $projectRoot = Split-Path -Parent $scriptRoot
    $fullTarget = Join-Path $projectRoot $fullTarget
}
$fullTarget = [System.IO.Path]::GetFullPath($fullTarget)

Write-Output "PROD_TARGET $fullTarget - цель установки (полный путь)"
Write-Output "RUNNING $install - запуск установщика"
& powershell -NoProfile -ExecutionPolicy Bypass -File $install -Target "$fullTarget"
$code = $LASTEXITCODE
Write-Output "INSTALL_EXIT $code - код возврата установщика"
if ($code -eq 0) { Write-Output "PROD_OK - установка выполнена и подтверждена проверкой" } else { Write-Output "PROD_FAIL - проверка не подтвердила установку, см. строки VERIFY_FAIL выше" }
exit $code

# apply_v8.0.ps1
# Version 1.0 / 2026-09-10
# One-shot rollout of report v8.0 (template v4.0, 8 slides) into a workbook:
#   1) install\install.ps1  - VBA modules (incl. modContentZone, modContentDisc),
#                             Power Query from src\powerquery, template tmp_index.html;
#   2) tools\compile_check.ps1 - forces a full VBA compile on a temp copy.
# Nothing is generated here: after this script reload the JSON in the workbook
# (the tbDATA rows must be rebuilt by the refreshed M code) and press the report button.
# ASCII-only by project rule (PowerShell 5.1 ANSI parsing).
#
# Usage from the project root:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\apply_v8.0.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\apply_v8.0.ps1 -Book ".\build\ReportMTO v7.0.xlsm"

param(
    [string]$Book = "ReportMTO.xlsm"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $root
if (-not [System.IO.Path]::IsPathRooted($Book)) {
    $Book = [System.IO.Path]::GetFullPath((Join-Path $root $Book))
}
if (-not (Test-Path $Book)) { Write-Output ("BOOK_MISSING " + $Book); exit 1 }

Write-Output ("APPLY_TARGET " + $Book)

Write-Output "--- STEP 1: install (VBA + Power Query + template)"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "install\install.ps1") -Target $Book
if ($LASTEXITCODE -ne 0) { Write-Output "APPLY_FAIL install returned $LASTEXITCODE"; exit 1 }

Write-Output "--- STEP 2: VBA compile check"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "tools\compile_check.ps1") -Book $Book

Write-Output "--- DONE. Next: open the workbook, reload the JSON, then press the report button."

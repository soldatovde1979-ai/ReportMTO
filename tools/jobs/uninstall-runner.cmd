@echo off
chcp 65001 >nul
rem uninstall-runner.cmd - снимает задачу планировщика ReportMTO-Runner.
schtasks /Delete /TN "ReportMTO-Runner" /F
echo.
echo  Помощник снят.
pause

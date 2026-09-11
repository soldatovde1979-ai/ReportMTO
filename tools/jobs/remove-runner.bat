@echo off
chcp 65001 >nul
rem Снимает задачу планировщика ReportMTO-Runner. Запускать только осознанно.
schtasks /Delete /TN "ReportMTO-Runner" /F
echo.
echo  Задача планировщика удалена.
pause

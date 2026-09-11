@echo off
chcp 65001 >nul
rem install-runner.cmd - Version 1.0 / 11.09.2026
rem Ставит задачу планировщика ReportMTO-Runner: раз в 2 минуты проверяет очередь
rem заданий tools\jobs\queue. Права администратора не нужны.

echo.
echo  Установка помощника ReportMTO...
echo.

schtasks /Create /TN "ReportMTO-Runner" /TR "powershell -NoProfile -ExecutionPolicy Bypass -File \"%~dp0runner.ps1\" -Once" /SC MINUTE /MO 2 /F
if errorlevel 1 goto fail

schtasks /Run /TN "ReportMTO-Runner" >nul 2>&1

echo.
echo  ГОТОВО. Помощник установлен и запущен.
echo  Это окно можно закрыть.
echo.
pause
exit /b 0

:fail
echo.
echo  НЕ ПОЛУЧИЛОСЬ. Покажите это окно Дмитрию.
echo.
pause
exit /b 1

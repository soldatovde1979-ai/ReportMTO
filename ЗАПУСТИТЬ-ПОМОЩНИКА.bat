@echo off
chcp 65001 >nul
title Запуск помощника ReportMTO
echo.
echo   ==========================================
echo    Запуск помощника ReportMTO
echo   ==========================================
echo.
echo   Ставлю задачу планировщика...
echo.

schtasks /Create /TN "ReportMTO-Runner" /TR "powershell -NoProfile -ExecutionPolicy Bypass -File \"%~dp0tools\jobs\runner.ps1\" -Once" /SC MINUTE /MO 2 /F
if errorlevel 1 goto fail

echo.
echo   Запускаю первый раз...
schtasks /Run /TN "ReportMTO-Runner"
if errorlevel 1 goto fail

echo.
echo   Проверяю, что задача создана:
schtasks /Query /TN "ReportMTO-Runner"
if errorlevel 1 goto fail

echo.
echo   ==========================================
echo    ГОТОВО. Помощник работает.
echo    Можно закрыть это окно.
echo   ==========================================
echo.
pause
exit /b 0

:fail
echo.
echo   ==========================================
echo    НЕ ПОЛУЧИЛОСЬ.
echo    Сфотографируйте это окно и отправьте Дмитрию.
echo   ==========================================
echo.
pause
exit /b 1

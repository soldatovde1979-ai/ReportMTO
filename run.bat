@echo off
chcp 65001 >nul
title Запуск ReportMTO: формирование отчета
echo.
echo   ==========================================
echo    Запуск книги и формирование отчета
echo   ==========================================
echo.
echo   Открываю книгу, строю сводные и отчет.
echo   Это может занять несколько минут.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\run_report.ps1"
if errorlevel 1 goto fail

echo.
echo   ==========================================
echo    ГОТОВО. Отчет собран в папке result\
echo   ==========================================
echo.
pause
exit /b 0

:fail
echo.
echo   ==========================================
echo    НЕ ПОЛУЧИЛОСЬ. Подробности выше в окне.
echo   ==========================================
echo.
pause
exit /b 1

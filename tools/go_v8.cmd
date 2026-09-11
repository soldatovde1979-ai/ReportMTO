@echo off
rem go_v8.cmd - Version 2.0 / 2026-09-10: compile check of the production workbook.
rem Writes a marker first so a failed launch can be told from a failed script.
echo STARTED %DATE% %TIME% > "D:\GOOGLEDISK\PROJECTs\ReportMTO\tools\_go_marker.txt"
powershell -NoProfile -ExecutionPolicy Bypass -File "D:\GOOGLEDISK\PROJECTs\ReportMTO\tools\compile_check.ps1" -Book "D:\GOOGLEDISK\PROJECTs\ReportMTO\ReportMTO.xlsm" > "D:\GOOGLEDISK\PROJECTs\ReportMTO\tools\_compile_check.log" 2>&1
echo FINISHED %DATE% %TIME% >> "D:\GOOGLEDISK\PROJECTs\ReportMTO\tools\_go_marker.txt"

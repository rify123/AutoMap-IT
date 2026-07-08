@echo off
REM AutoMap IT
REM Copyright (c) 2026 Aziz Arbiine
REM Todos los derechos reservados.
REM Autor: Aziz Arbiine
REM LinkedIn: https://www.linkedin.com/in/aziz-arbiine
REM Queda prohibida la copia, distribucion, modificacion o atribucion falsa sin autorizacion escrita.

setlocal
echo Cerrando procesos de AutoMap IT...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*Start-AutoMapIT.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
echo Listo.
pause

@echo off
REM AutoMap IT
REM Copyright (c) 2026 Aziz Arbiine
REM Todos los derechos reservados.
REM Autor: Aziz Arbiine
REM LinkedIn: https://www.linkedin.com/in/aziz-arbiine
REM Queda prohibida la copia, distribucion, modificacion o atribucion falsa sin autorizacion escrita.

setlocal
cd /d "%~dp0"
title AutoMap IT - Portal PowerShell
echo.
echo Iniciando AutoMap IT...
echo PowerShell abrira el portal en el primer puerto libre desde 8780.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-AutoMapIT.ps1" -Port 8780
echo.
echo AutoMap IT se ha detenido.
pause

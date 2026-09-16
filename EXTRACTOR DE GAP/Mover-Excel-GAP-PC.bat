@echo off
setlocal
cd /d "%~dp0"

echo === Mover Excel GAP a CARGA\EXCEL ===
echo === Este .bat se corre EN TU PC REAL (no en la VM) ===
echo.

where python >nul 2>nul
if errorlevel 1 goto :SIN_PYTHON

python mover_excel_gap_pc.py %*
if errorlevel 1 goto :FALLO

echo.
echo Listo. El Excel ya esta en CARGA\EXCEL
pause
goto :FIN

:SIN_PYTHON
echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
echo Marca la casilla "Add Python to PATH" durante la instalacion.
pause
exit /b 1

:FALLO
echo.
echo Algo fallo - revisa el mensaje de arriba.
pause
exit /b 1

:FIN
endlocal

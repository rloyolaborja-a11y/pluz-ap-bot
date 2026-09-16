@echo off
setlocal
cd /d "%~dp0"

echo === Extraer GAP (SDAPeru / Extractor datos SAP) ===
echo === Este .bat se corre DENTRO del escritorio virtual ===
echo.

where python >nul 2>nul
if errorlevel 1 goto :SIN_PYTHON

python -c "import pywinauto" >nul 2>nul
if errorlevel 1 goto :SIN_PYWINAUTO

python extraer_gap_vm.py
if errorlevel 1 goto :FALLO

echo.
echo Listo. El Excel ya esta en la carpeta OneDrive "GAP".
pause
goto :FIN

:SIN_PYTHON
echo No se encontro Python instalado EN ESTA MAQUINA VIRTUAL.
echo Instalalo desde https://www.python.org/downloads/
echo Marca la casilla "Add Python to PATH" durante la instalacion.
pause
exit /b 1

:SIN_PYWINAUTO
echo Falta instalar pywinauto. Abre una consola (CMD) en esta carpeta
echo DENTRO de la VM y ejecuta:
echo     pip install pywinauto pillow
pause
exit /b 1

:FALLO
echo.
echo Algo fallo - revisa el mensaje de arriba y la carpeta "diagnosticos"
echo dentro de GAP_RPA_Excel, en tu carpeta de usuario DE LA VM.
pause
exit /b 1

:FIN
endlocal

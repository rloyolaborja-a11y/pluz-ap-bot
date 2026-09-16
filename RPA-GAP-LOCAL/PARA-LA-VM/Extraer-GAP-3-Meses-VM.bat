@echo off
setlocal
cd /d "%~dp0"

echo === Extraer GAP - 3 MESES (SDAPeru / Extractor datos SAP) ===
echo === Este .bat se corre DENTRO del escritorio virtual ===
echo.
echo Hace 3 extracciones seguidas de 30 dias cada una (mes actual, -1, -2)
echo reusando la misma sesion de SDAPeru. Deja R_m0_*.xls, R_m1_*.xls y
echo R_m2_*.xls en la carpeta OneDrive "GAP".
echo.
echo Si CUALQUIERA de las 3 falla, se detiene: la data de 3 meses tiene que
echo estar completa.
echo.

where python >nul 2>nul
if errorlevel 1 goto :SIN_PYTHON

python -c "import pywinauto" >nul 2>nul
if errorlevel 1 goto :SIN_PYWINAUTO

python extraer_gap_vm.py --meses 3
if errorlevel 1 goto :FALLO

echo.
echo Listo. Los 3 Excel ya estan en la carpeta OneDrive "GAP".
echo En tu PC real: Mover-Excel-GAP-PC.bat --meses 3  (o Iniciar-GAP-PC.bat --meses 3)
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
echo NO uses los R_m*.xls que hayan quedado a medias.
pause
exit /b 1

:FIN
endlocal

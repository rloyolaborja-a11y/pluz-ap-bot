@echo off
setlocal
cd /d "%~dp0"

echo === Prueba de la Pantalla 1 del robot UIA (Clase de orden + Periodo + Layout) ===
echo Version en construccion -- todavia NO es el bot final (falta Pantalla 2 y 3).
echo Va a llenar la pantalla de verdad y esperar tu confirmacion ANTES de
echo clickear 'Ejecutar', para que puedas revisar que los datos quedaron bien.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado.
  pause
  exit /b 1
)

python -c "import pywinauto" >nul 2>nul
if errorlevel 1 (
  echo Falta pywinauto. Abre una consola en esta carpeta y ejecuta:
  echo     pip install pywinauto
  pause
  exit /b 1
)

python descargar_excel_sap_uia.py

echo.
pause

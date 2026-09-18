@echo off
setlocal
cd /d "%~dp0"

echo === Prueba de la Pantalla 1 de SAP con pywinauto (sin CDP) ===
echo (Version en construccion -- no es el bot final todavia.)
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

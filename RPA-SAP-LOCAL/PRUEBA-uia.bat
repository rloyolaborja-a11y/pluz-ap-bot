@echo off
setlocal
cd /d "%~dp0"

echo === Prueba de factibilidad: pywinauto + UI Automation contra SAP ===
echo (Esto NO es el bot real, es solo para probar una idea. No toca nada.)
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

python _prueba_uia.py

echo.
pause

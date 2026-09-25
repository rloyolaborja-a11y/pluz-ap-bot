@echo off
setlocal
cd /d "%~dp0"

echo === Diagnostico visual de la Pantalla 1 de SAP (para terminar el robot UIA) ===
echo Esto NO es el bot real, no descarga nada -- solo saca una foto marcada
echo con numeros para identificar los campos de Clase de orden/Periodo/Layout.
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

python -c "import PIL" >nul 2>nul
if errorlevel 1 (
  echo Falta Pillow. Abre una consola en esta carpeta y ejecuta:
  echo     pip install Pillow
  pause
  exit /b 1
)

python diagnostico_pantalla1.py

echo.
pause

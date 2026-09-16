@echo off
setlocal
cd /d "%~dp0"

echo === Publicar HISTORICO de Atendidas AP (meses ya cerrados) ===
echo.
echo Esto se corre UNA SOLA VEZ para cargar los meses pasados (Enero a Agosto).
echo De ahi en adelante cada mes se archiva solo, automaticamente.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python publicar_historico.py
if errorlevel 1 (
  echo.
  echo La publicacion NO se completo del todo - revisa el detalle de arriba.
  pause
  exit /b 1
)

echo.
pause

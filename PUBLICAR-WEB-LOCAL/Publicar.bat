@echo off
setlocal
cd /d "%~dp0"

echo === Publicar sitio web (Pendientes AP + Atendidas AP) ===
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

if not exist "config.json" (
  echo No existe config.json todavia.
  echo Copia config.ejemplo.json, renombralo a config.json y pega tu token de GitHub.
  pause
  exit /b 1
)

if not exist "SITIO\index.html" (
  echo No encuentro SITIO\index.html
  echo Descomprime el zip del portal DENTRO de la carpeta SITIO\ antes de publicar.
  pause
  exit /b 1
)

python publicar_sitio.py
if errorlevel 1 (
  echo.
  echo La publicacion NO se completo del todo - revisa el detalle de arriba.
  pause
  exit /b 1
)

echo.
echo Listo.
pause

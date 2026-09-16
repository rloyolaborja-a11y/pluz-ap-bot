@echo off
setlocal
cd /d "%~dp0"

echo === Reporte de Atendidas AP ===
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python procesar_diario.py
if errorlevel 1 (
  echo.
  echo El procesamiento fallo. Revisa el mensaje de arriba.
  pause
  exit /b 1
)

if not exist "config.json" (
  echo.
  echo AVISO: no existe config.json, asi que no se va a publicar automaticamente.
  echo Copia config.ejemplo.json, renombralo a config.json y completa tu token de GitHub.
  pause
  exit /b 0
)

echo.
python feed.py
if errorlevel 1 (
  echo.
  echo La publicacion NO se completo del todo - revisa el detalle de arriba.
  pause
  exit /b 1
)

if exist "HISTORICO" (
  echo.
  echo Publicando HISTORICO (meses ya archivados)...
  python publicar_historico.py
  if errorlevel 1 (
    echo.
    echo AVISO: no se pudo publicar el HISTORICO esta vez - no es grave, se
    echo reintenta solo la proxima vez que corras esto, o corre
    echo Publicar-Historico.bat aparte.
  )
)

echo.
echo Listo.
pause

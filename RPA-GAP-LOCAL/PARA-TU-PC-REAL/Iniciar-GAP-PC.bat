@echo off
setlocal
cd /d "%~dp0"

echo === Iniciar GAP (manda la senal a la VM y espera el Excel) ===
echo.
echo IMPORTANTE: en la VM tiene que estar corriendo Vigilar-GAP-VM.bat
echo (con SDAPeru ya abierto y logueado) para que esto funcione solo.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python iniciar_y_esperar_gap_pc.py %*
if errorlevel 1 (
  echo.
  echo Algo fallo o se agoto el tiempo de espera - revisa el mensaje de arriba.
  pause
  exit /b 1
)

echo.
echo Listo.
pause

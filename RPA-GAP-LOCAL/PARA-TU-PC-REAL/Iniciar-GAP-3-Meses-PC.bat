@echo off
setlocal
cd /d "%~dp0"

echo === Iniciar GAP - 3 MESES (manda la senal a la VM y espera los 3 Excel) ===
echo.
echo Manda una senal a la VM con MESES=3. La VM (con el Vigilante corriendo)
echo hace 3 extracciones de 30 dias reusando la misma ventana del Extractor,
echo y este script espera a que lleguen LOS 3 Excel por OneDrive y los copia
echo a ..\CARGA\EXCEL\ (R_m0_*.xls, R_m1_*.xls, R_m2_*.xls).
echo.
echo IMPORTANTE: en la VM tiene que estar corriendo Vigilar-GAP-VM.bat
echo (version fix22 o mas nueva, con SDAPeru ya abierto y logueado).
echo.
echo Si no llegan los 3, NO se copia nada parcial -- te avisa y podes correr
echo   Mover-Excel-GAP-PC.bat --meses 3   cuando ya esten sincronizados.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python iniciar_y_esperar_gap_pc.py --meses 3
if errorlevel 1 (
  echo.
  echo Algo fallo o se agoto el tiempo de espera - revisa el mensaje de arriba.
  pause
  exit /b 1
)

echo.
echo Listo. Los 3 Excel ya estan en ..\CARGA\EXCEL\
pause

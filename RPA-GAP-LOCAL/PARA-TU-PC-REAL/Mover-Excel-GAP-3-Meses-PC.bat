@echo off
setlocal
cd /d "%~dp0"

echo === Mover Excel GAP - 3 MESES (OneDrive "GAP" -> CARGA\EXCEL) ===
echo.
echo Copia los 3 Excel MAS RECIENTES de la carpeta OneDrive "GAP" a
echo ..\CARGA\EXCEL\ (borra los que hubiera antes). Usar despues de correr
echo Extraer-GAP-3-Meses-VM.bat en la VM y esperar a que OneDrive sincronice.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  pause
  exit /b 1
)

python mover_excel_gap_pc.py --meses 3
if errorlevel 1 (
  echo.
  echo Algo fallo - revisa el mensaje de arriba.
  pause
  exit /b 1
)

echo.
echo Listo. Los 3 Excel ya estan en ..\CARGA\EXCEL\
pause

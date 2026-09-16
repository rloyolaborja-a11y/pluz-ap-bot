@echo off
setlocal
cd /d "%~dp0"

echo === Descargar Excel de SAP - 3 MESES (IW39 / ZM06) ===
echo.
echo Descarga 3 ventanas de 30 dias (mes actual, -1, -2) en una sola corrida,
echo reusando el mismo navegador y login. Deja SAP_m0.xlsx, SAP_m1.xlsx y
echo SAP_m2.xlsx en ..\CARGA\SAP\ (borra lo que hubiera antes ahi).
echo.
echo Si CUALQUIERA de las 3 falla, el proceso se detiene: la data de 3 meses
echo tiene que estar completa.
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python -c "import playwright" >nul 2>nul
if errorlevel 1 (
  echo Falta instalar Playwright. Abre una consola en esta carpeta y ejecuta:
  echo     pip install playwright
  echo     playwright install chromium
  pause
  exit /b 1
)

python descargar_excel_sap.py --meses 3
if errorlevel 1 (
  echo.
  echo Algo fallo - revisa el mensaje de arriba y la carpeta "diagnosticos".
  echo NO uses los Excel que hayan quedado a medias.
  pause
  exit /b 1
)

echo.
echo Listo. Los 3 Excel ya estan en ..\CARGA\SAP\
pause

@echo off
setlocal
cd /d "%~dp0"

echo === Descargar Excel de SAP (IW39 / ZM06) ===
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

python -c "import pywinauto" >nul 2>nul
if errorlevel 1 (
  echo AVISO: pywinauto no esta instalado. Solo es un respaldo por si
  echo aparece el cuadro nativo de Windows "Guardar como" ^(normalmente no
  echo aparece^). Si quieres tenerlo: pip install pywinauto
  echo.
)

python descargar_excel_sap.py
if errorlevel 1 (
  echo.
  echo Algo fallo - revisa el mensaje de arriba y la carpeta "diagnosticos".
  pause
  exit /b 1
)

echo.
echo Listo. El Excel ya esta en ..\CARGA\SAP\
pause

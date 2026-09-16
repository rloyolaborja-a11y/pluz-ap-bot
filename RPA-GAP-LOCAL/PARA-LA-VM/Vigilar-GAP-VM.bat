@echo off
setlocal
cd /d "%~dp0"

echo === Vigilar GAP (arranca solo cuando llega la senal desde tu PC real) ===
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca la casilla "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python -c "import pywinauto" >nul 2>nul
if errorlevel 1 (
  echo Falta instalar pywinauto. Abre una consola en esta carpeta y ejecuta:
  echo     pip install pywinauto
  pause
  exit /b 1
)

echo Deja SDAPeru abierto y con sesion iniciada antes de seguir.
echo Esta ventana se va a quedar corriendo, vigilando -- podes minimizarla.
echo Para parar: Ctrl+C, o cerrar esta ventana.
echo.

python vigilar_gap_vm.py
pause

@echo off
setlocal
cd /d "%~dp0"

echo ############################################################
echo #   REPORTE AP - CORRIDA COMPLETA (SAP + GAP + Pendientes + Atendidas)
echo ############################################################
echo.
echo ANTES DE SEGUIR:
echo   - En la VM tiene que estar corriendo Vigilar-GAP-VM.bat
echo     (con SDAPeru abierto y logueado).
echo   - Si algun paso falla, se DETIENE todo y te dice donde mirar.
echo.
echo Si NO queres descargar ahora (ya bajaste los Excel a mano), corre en
echo su lugar:  Ejecutar-Todo.bat --saltar-descargas
echo.
pause

where python >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
  echo (marca "Add Python to PATH" durante la instalacion^)
  pause
  exit /b 1
)

python ejecutar_todo.py %*
set RC=%errorlevel%

echo.
if "%RC%"=="0" (
  echo === Termino OK ===
) else (
  echo === Termino con ERROR - revisa el resumen y el log de arriba ===
)
pause
exit /b %RC%

@echo off
setlocal
cd /d "%~dp0"

echo === Importar mes(es) historico(s) de Atendidas AP (con todas las columnas) ===
echo.

where python >nul 2>nul
if errorlevel 1 goto :SIN_PYTHON

if "%~1"=="" (
  echo No se arrastro ningun archivo encima -- se van a buscar Excel dentro de
  echo la carpeta EXCEL-MESES-ANTERIORES\ en su lugar.
  echo.
  python importar_mes_historico.py
) else (
  python importar_mes_historico.py %*
)
if errorlevel 1 goto :FALLO

echo.
echo Listo. Ahora corre Publicar-Historico.bat para subir estos meses al sitio.
pause
goto :FIN

:SIN_PYTHON
echo No se encontro Python instalado. Instalalo desde https://www.python.org/downloads/
echo Marca la casilla "Add Python to PATH" durante la instalacion.
pause
exit /b 1

:FALLO
echo.
echo Algo fallo - revisa el mensaje de arriba.
pause
exit /b 1

:FIN
endlocal

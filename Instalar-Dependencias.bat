@echo off
echo === Instalando librerias de Python necesarias para el bot ===
echo (pandas, openpyxl, requests, playwright, pywinauto)
echo.
python -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto :error

echo.
echo === Instalando el navegador que usa Playwright para SAP ===
python -m playwright install chromium
if errorlevel 1 goto :error

echo.
echo Listo. Ya se puede correr el panel (PANEL-CONTROL-LOCAL\Panel-AP.bat).
pause
exit /b 0

:error
echo.
echo Algo fallo -- revisa el mensaje de arriba. Si dice que no encuentra
echo "python" o "pip", instala Python desde python.org primero (marca la
echo casilla "Add python.exe to PATH" durante la instalacion).
pause
exit /b 1

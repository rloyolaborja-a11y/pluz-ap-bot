@echo off
cd /d "%~dp0"

rem Resolver pythonw REAL (evitando el alias de la Microsoft Store).
set "PYW="
for /d %%D in ("%LOCALAPPDATA%\Python\pythoncore-*") do if exist "%%D\pythonw.exe" set "PYW=%%D\pythonw.exe"
if not defined PYW for /f "delims=" %%P in ('python -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))" 2^>nul') do set "PYW=%%P"
if not defined PYW set "PYW=pythonw"

rem Abre el Panel como ventana de aplicacion (sin consola).
start "" "%PYW%" "%~dp0panel.py"
exit /b 0

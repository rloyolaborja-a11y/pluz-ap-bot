@echo off
setlocal
echo Sacando el Panel de Control AP (la carpeta y los archivos quedan)...
schtasks /delete /tn "PanelAP_Tick" /f 2>nul
del "%USERPROFILE%\Desktop\Panel de Control AP.lnk" 2>nul
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Panel de Control AP.lnk" 2>nul
del "%~dp0tick_launcher.vbs" 2>nul
echo Listo: se quito la tarea programada y los accesos directos.
echo (Los .bat de siempre siguen funcionando igual.)
pause

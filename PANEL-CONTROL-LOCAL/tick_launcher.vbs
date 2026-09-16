' Generado por Instalar-Panel.bat - lanza tick.py sin ninguna ventana.
' El tercer parametro True = ESPERAR a que tick.py termine: si tick.py corre
' el pipeline (7 min), la Tarea de Windows lo mantiene vivo todo ese rato en
' vez de matarlo al salir wscript.
Set sh = CreateObject("WScript.Shell")
sh.Run """C:\Users\P721725611\AppData\Local\Python\pythoncore-3.14-64\pythonw.exe"" ""C:\Users\P721725611\Desktop\Programaciones\1. REPORTE\PANEL-CONTROL-LOCAL\tick.py""", 0, True

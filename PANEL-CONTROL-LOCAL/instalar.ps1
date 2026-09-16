# Instala el Panel de Control AP: Tarea de Windows (tick cada 5 min) + accesos directos.
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "============================================"
Write-Host "  Instalar Panel de Control AP"
Write-Host "============================================`n"

# --- 1. Resolver pythonw REAL (NO el alias de la Microsoft Store) -------------
function Resolver-Pythonw {
    foreach ($cmd in @('python', 'py')) {
        try {
            $p = (& $cmd -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))" 2>$null)
            if ($p) { $p = $p.Trim() }
            if ($p -and (Test-Path $p) -and ($p -notmatch 'WindowsApps')) { return $p }
        } catch {}
    }
    $g = Get-Command pythonw -All -ErrorAction SilentlyContinue |
         Where-Object { $_.Source -and $_.Source -notmatch 'WindowsApps' } |
         Select-Object -First 1
    if ($g) { return $g.Source }
    return $null
}

$pyw = Resolver-Pythonw
if (-not $pyw) {
    Write-Host "ERROR: no encontre un pythonw.exe de verdad." -ForegroundColor Red
    Write-Host "Instala Python desde python.org (no la version de la Microsoft Store)."
    Read-Host "`nEnter para cerrar"; exit 1
}
Write-Host "Python:  $pyw"

$tick = Join-Path $dir 'tick.py'
$bat  = Join-Path $dir 'Panel-AP.bat'

# Lanzador VBS: ejecuta tick.py 100% oculto (ni un parpadeo de consola,
# aunque el Programador de tareas se porte raro).
$vbs = Join-Path $dir 'tick_launcher.vbs'
@"
' Generado por Instalar-Panel.bat - lanza tick.py sin ninguna ventana.
' El tercer parametro True = ESPERAR a que tick.py termine: si tick.py corre
' el pipeline (7 min), la Tarea de Windows lo mantiene vivo todo ese rato en
' vez de matarlo al salir wscript.
Set sh = CreateObject("WScript.Shell")
sh.Run """$pyw"" ""$tick""", 0, True
"@ | Set-Content -Path $vbs -Encoding ascii

# --- 2. Tarea de Windows: revisa cada 5 min si toca correr -------------------
Unregister-ScheduledTask -TaskName 'PanelAP_Tick' -Confirm:$false -ErrorAction SilentlyContinue
$act = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\wscript.exe" `
        -Argument ('//nologo //B "{0}"' -f $vbs) -WorkingDirectory $dir
$trg = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5)
$set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$prn = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName 'PanelAP_Tick' -Action $act -Trigger $trg -Settings $set `
        -Principal $prn -Description 'Panel de Control AP: revisa cada 5 min si toca una corrida programada.' | Out-Null
Write-Host "[OK] Tarea 'PanelAP_Tick' creada (corre solo con tu sesion de Windows abierta)."

# --- 2b. Tarea de Windows: "clic fantasma" (corre siempre, desde el logon) --
# (2026-09-12) Simula actividad minima (mueve el mouse 1px, solo si la PC ya
# lleva un rato inactiva) para que un timeout de inactividad de la pagina
# del portal Oracle Secure Desktops no la cierre mientras la automatizacion
# corre sola. Ver clic_fantasma_pc.py.
$fantasma = Join-Path $dir 'clic_fantasma_pc.py'
$vbsFant = Join-Path $dir 'clic_fantasma_launcher.vbs'
@"
' Generado por instalar.ps1 - lanza clic_fantasma_pc.py sin ninguna ventana.
' Corre para siempre (no hay tercer parametro True -- no hace falta esperar).
Set sh = CreateObject("WScript.Shell")
sh.Run """$pyw"" ""$fantasma""", 0, False
"@ | Set-Content -Path $vbsFant -Encoding ascii

Unregister-ScheduledTask -TaskName 'PanelAP_ClicFantasma' -Confirm:$false -ErrorAction SilentlyContinue
$actF = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\wscript.exe" `
        -Argument ('//nologo //B "{0}"' -f $vbsFant) -WorkingDirectory $dir
$trgF = New-ScheduledTaskTrigger -AtLogOn
$setF = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero)
$prnF = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName 'PanelAP_ClicFantasma' -Action $actF -Trigger $trgF -Settings $setF `
        -Principal $prnF -Description 'Panel de Control AP: mueve el mouse 1px cuando esta inactiva, para no perder la sesion del portal VDI.' | Out-Null
Write-Host "[OK] Tarea 'PanelAP_ClicFantasma' creada (corre siempre desde que iniciás sesión)."

# --- 3. Accesos directos (Escritorio + menu Inicio) ------------------------
# Apuntan directo a pythonw.exe -> no hay .bat ni cmd de por medio, cero parpadeo.
$panelpy = Join-Path $dir 'panel.py'
$ico = Join-Path $dir 'panel.ico'
if (-not (Test-Path $ico)) { $ico = "$env:SystemRoot\System32\imageres.dll,171" }
$w = New-Object -ComObject WScript.Shell
foreach ($d in @([Environment]::GetFolderPath('Desktop'),
                 (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'))) {
    $s = $w.CreateShortcut((Join-Path $d 'Panel de Control AP.lnk'))
    $s.TargetPath       = $pyw
    $s.Arguments        = ('"{0}"' -f $panelpy)
    $s.WorkingDirectory = $dir
    $s.IconLocation     = $ico
    $s.WindowStyle      = 7   # minimizada
    $s.Save()
}
Write-Host "[OK] Acceso directo en el Escritorio y en el menu Inicio."

Write-Host "`nListo. Abri 'Panel de Control AP' desde el Escritorio."
Write-Host "Para sacar todo:  Desinstalar-Panel.bat"
Read-Host "`nEnter para cerrar"

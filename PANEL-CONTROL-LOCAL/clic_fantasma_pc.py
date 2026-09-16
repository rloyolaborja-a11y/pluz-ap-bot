#!/usr/bin/env python3
"""Clic fantasma -- PC real.

Mueve el mouse 1px (y lo devuelve) cada pocos minutos, SOLO si la PC ya
lleva un rato sin actividad real, para que un timeout de inactividad de
la pagina web del portal Oracle Secure Desktops (u otra sesion sensible a
inactividad) no la cierre mientras la automatizacion corre sin que haya
nadie tocando el mouse/teclado.

No hace nada mas: no clickea nada, no cambia el foco de ninguna ventana.
Corre para siempre en segundo plano.

2026-09-16: en PCs donde una politica de IT bloquea las Tareas Programadas
"al iniciar sesion" (ej. la de la jefa -- Register-ScheduledTask daba
"Acceso denegado" solo con ESTE disparador, no con el de PanelAP_Tick, que
es por horario), instalar.ps1 usa un disparador REPETIDO cada 1 minuto en
vez de "al iniciar sesion" para lograr el mismo efecto ("siempre corriendo")
sin ese permiso especial. Como eso significa que Windows puede intentar
lanzar este script de nuevo aunque ya haya una copia corriendo (el proceso
anterior nunca "termina" para que Windows sepa que seguir esperando), este
script ahora se fija solo si YA hay otra copia viva (mismo mecanismo de
PID que ya usa panel.py) y si es asi, se cierra al toque sin hacer nada.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import os
import time
from pathlib import Path

INTERVALO_SEG = 120          # cada cuanto revisa/actua
UMBRAL_INACTIVIDAD_SEG = 60  # solo actua si la PC YA estaba inactiva este rato
                              # (si hay actividad real, no hace falta simular nada)

PID_PATH = Path(__file__).resolve().parent / "_clic_fantasma.pid"


def _pid_vivo(pid: int) -> bool:
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return False
    ctypes.windll.kernel32.CloseHandle(h)
    return True


def _otra_instancia_viva() -> bool:
    if not PID_PATH.exists():
        return False
    try:
        pid_anterior = int(PID_PATH.read_text(encoding="utf-8").strip())
    except (ValueError, OSError):
        return False
    if pid_anterior == os.getpid():
        return False
    try:
        return _pid_vivo(pid_anterior)
    except Exception:
        return False


def _tomar_pid():
    try:
        PID_PATH.write_text(str(os.getpid()), encoding="utf-8")
    except OSError:
        pass


def _soltar_pid():
    try:
        if PID_PATH.exists() and PID_PATH.read_text(encoding="utf-8").strip() == str(os.getpid()):
            PID_PATH.unlink()
    except OSError:
        pass


class _LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.wintypes.UINT), ("dwTime", ctypes.wintypes.DWORD)]


def _segundos_inactiva() -> float:
    info = _LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(_LASTINPUTINFO)
    if not ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info)):
        return 0.0
    tick_actual = ctypes.windll.kernel32.GetTickCount()
    return max(0.0, (tick_actual - info.dwTime) / 1000.0)


def _nudge():
    pt = ctypes.wintypes.POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    ctypes.windll.user32.SetCursorPos(pt.x + 1, pt.y)
    time.sleep(0.05)
    ctypes.windll.user32.SetCursorPos(pt.x, pt.y)


def correr():
    print(f"Clic fantasma activo (reviso cada {INTERVALO_SEG}s; actua solo si "
          f"la PC lleva {UMBRAL_INACTIVIDAD_SEG}s+ inactiva).")
    while True:
        try:
            if _segundos_inactiva() >= UMBRAL_INACTIVIDAD_SEG:
                _nudge()
        except Exception:
            pass
        time.sleep(INTERVALO_SEG)


if __name__ == "__main__":
    if _otra_instancia_viva():
        pass  # ya hay una copia corriendo -- no hacer nada, cerrar en silencio
    else:
        _tomar_pid()
        try:
            correr()
        except KeyboardInterrupt:
            pass
        finally:
            _soltar_pid()

#!/usr/bin/env python3
"""Clic fantasma -- PC real.

Mueve el mouse 1px (y lo devuelve) cada pocos minutos, SOLO si la PC ya
lleva un rato sin actividad real, para que un timeout de inactividad de
la pagina web del portal Oracle Secure Desktops (u otra sesion sensible a
inactividad) no la cierre mientras la automatizacion corre sin que haya
nadie tocando el mouse/teclado.

No hace nada mas: no clickea nada, no cambia el foco de ninguna ventana.
Corre para siempre en segundo plano (pensado para lanzarse al iniciar
sesion de Windows, vía Tarea Programada -- ver instalar.ps1).
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import time

INTERVALO_SEG = 120          # cada cuanto revisa/actua
UMBRAL_INACTIVIDAD_SEG = 60  # solo actua si la PC YA estaba inactiva este rato
                              # (si hay actividad real, no hace falta simular nada)


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
    try:
        correr()
    except KeyboardInterrupt:
        pass

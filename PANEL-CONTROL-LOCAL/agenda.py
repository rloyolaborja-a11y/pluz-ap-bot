# -*- coding: utf-8 -*-
"""Cálculo de la programación automática (Fase 2).

Lo comparten panel.py (para mostrar "Próxima corrida: …") y tick.py (para
decidir si toca lanzar una corrida ahora).

Config esperada en panel_config.json -> "programacion":
    activa            bool
    intervalo_horas   1..24   (cada cuántas horas)
    minuto            0..59    (a qué minuto arranca; 50 = 10 min antes de la hora)
    hora_desde        0..24    (ventana horaria; 0 y 24 => todo el día)
    hora_hasta        0..24
"""

from datetime import datetime, timedelta


def _slots_horas(intervalo):
    intervalo = max(1, min(24, int(intervalo or 1)))
    return [h for h in range(24) if h % intervalo == 0]


def _en_ventana(h, hd, hh):
    hd = int(hd) % 24
    hh_raw = int(hh)
    if hh_raw >= 24:
        hh = 24
    else:
        hh = hh_raw % 24
    if hd == hh or (hd == 0 and hh == 24):
        return True                       # todo el día
    if hd < hh:
        return hd <= h < hh
    return h >= hd or h < hh               # ventana que cruza medianoche


def proxima_corrida(cfg, ref=None):
    """Primer datetime > ref en que corresponde correr. None si no está activa."""
    p = (cfg or {}).get("programacion", {}) or {}
    if not p.get("activa"):
        return None
    ref = ref or datetime.now()
    minuto = max(0, min(59, int(p.get("minuto", 50))))
    hd = p.get("hora_desde", 0)
    hh = p.get("hora_hasta", 24)
    slots = set(_slots_horas(p.get("intervalo_horas", 1)))

    base0 = ref.replace(minute=minuto, second=0, microsecond=0)
    for d in range(6):
        base = base0 + timedelta(days=d)
        for h in range(24):
            if h in slots and _en_ventana(h, hd, hh):
                cand = base.replace(hour=h)
                if cand > ref:
                    return cand
    return None


def _parse_iso(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def _slot_reciente(cfg, ref):
    """El slot programado más reciente <= ref (dentro de la ventana), o None."""
    p = cfg["programacion"]
    minuto = max(0, min(59, int(p.get("minuto", 50))))
    hd = p.get("hora_desde", 0)
    hh = p.get("hora_hasta", 24)
    slots = set(_slots_horas(p.get("intervalo_horas", 1)))
    base0 = ref.replace(minute=minuto, second=0, microsecond=0)
    for d in range(3):
        base = base0 - timedelta(days=d)
        for h in range(23, -1, -1):
            if h in slots and _en_ventana(h, hd, hh):
                cand = base.replace(hour=h)
                if cand <= ref:
                    return cand
    return None


def debe_correr(cfg, ultima_ejecucion_iso, ref=None, tolerancia_min=20):
    """Devuelve el datetime del slot a correr si toca AHORA, o None.

    Dispara cuando el slot programado más reciente todavía no se corrió
    ('ultima_ejecucion_iso' es anterior a ese slot) y no pasó tanto tiempo
    como para considerarlo perdido (PC apagada)."""
    p = (cfg or {}).get("programacion", {}) or {}
    if not p.get("activa"):
        return None
    ref = ref or datetime.now()
    slot = _slot_reciente(cfg, ref)
    if slot is None or (ref - slot) > timedelta(minutes=tolerancia_min):
        return None
    ancla = _parse_iso(ultima_ejecucion_iso)
    if ancla is not None and ancla >= slot:
        return None
    return slot

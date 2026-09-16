# -*- coding: utf-8 -*-
"""
historico_pendientes.py — "Libreta" acumulada de pendientes NO DT.

POR QUÉ EXISTE
--------------
El GAP/SAP que se descarga solo cubre ~1 mes (o 3, con la ventana ampliada).
Un pendiente NO DT reportado antes de esa ventana ya no viene en el Excel, y
sin esto se perdería del reporte aunque el poste siga malogrado.

Esta "libreta" (bd_nodt_hist.json) NO se borra: cada corrida la actualiza
comparándola con lo que trajo el Excel de ese día.

REGLA DE ACTUALIZACIÓN (una fila = un "SAP - Número de reclamo")
--------------------------------------------------------------
Por cada caso ya anotado en la libreta:
  - Vino en el Excel y SIGUE pendiente  -> se actualiza con lo nuevo y se le
    quita la marca "sin confirmar".
  - Vino en el Excel pero YA NO está pendiente (pasó a ATENDIDA / anulada /
    reasignada)                          -> se RETIRA de la libreta (cierre
    confirmado; si fue atendido, pasa a la página de Atendidas).
  - NO vino en el Excel (es más viejo que la ventana descargada) -> se
    CONSERVA, pero marcado "_sinConfirmar": no se puede verificar si sigue
    abierto o si ya lo cerraron fuera de ventana.
Los pendientes NO DT nuevos del Excel que no estaban en la libreta se agregan.

Solo NO DT. Los DT siguen saliendo 100% de la ventana actual, sin libreta.
"""

import json
import os
import re
from datetime import datetime

ARCHIVO = "bd_nodt_hist.json"


def _clave(valor):
    """Misma normalización que clave_busqueda() de procesar_diario.py, para
    que el cruce por 'SAP - Número de reclamo' sea consistente."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.upper()


def ruta_archivo(hist_dir):
    return os.path.join(hist_dir, ARCHIVO)


def cargar(hist_dir):
    """Devuelve la lista de filas de la libreta (o [] si todavía no existe)."""
    ruta = ruta_archivo(hist_dir)
    if not os.path.exists(ruta):
        return []
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
        return datos.get("rows", []) or []
    except Exception:
        return []


def guardar(hist_dir, rows, updated_at):
    os.makedirs(hist_dir, exist_ok=True)
    ruta = ruta_archivo(hist_dir)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump({"updated_at": updated_at, "rows": rows}, f, ensure_ascii=False)
    return len(rows)


def mezclar(acumulado, nuevas_pendientes, reclamos_en_gap, ahora_iso):
    """Aplica la regla de actualización descrita arriba.

    acumulado            : filas de la libreta actual (dicts ya "planos"/JSON).
    nuevas_pendientes    : pendientes NO DT vigentes de esta corrida (mismos
                           dicts, ya deshidratados a JSON).
    reclamos_en_gap      : set con la clave normalizada de TODOS los
                           'SAP - Número de reclamo' que aparecen en el Excel
                           GAP de esta corrida (cualquier estado).
    ahora_iso            : marca de tiempo de la corrida.

    Devuelve (filas_resultado, stats).
    """
    por_id = {}
    for r in (acumulado or []):
        por_id[_clave(r.get("SAP - Número de reclamo"))] = dict(r)

    ids_frescas = set()
    for r in (nuevas_pendientes or []):
        k = _clave(r.get("SAP - Número de reclamo"))
        if not k:
            continue
        ids_frescas.add(k)
        fila = dict(r)
        fila["_sinConfirmar"] = False
        fila["_ultimaVezVisto"] = ahora_iso
        por_id[k] = fila

    reclamos_en_gap = reclamos_en_gap or set()
    resultado = []
    retirados = 0
    marcados_ahora = 0
    for k, fila in por_id.items():
        if k in ids_frescas:
            resultado.append(fila)
            continue
        if k in reclamos_en_gap:
            # Se vio en el Excel pero no está entre los pendientes vigentes
            # -> ya se cerró / atendió / reasignó: se retira de la libreta.
            retirados += 1
            continue
        # No se vio en el Excel: se conserva, marcado "sin confirmar".
        if not fila.get("_sinConfirmar"):
            marcados_ahora += 1
        fila = dict(fila)
        fila["_sinConfirmar"] = True
        resultado.append(fila)

    stats = {
        "en_libreta_antes": len(acumulado or []),
        "pendientes_frescos": len(ids_frescas),
        "retirados": retirados,
        "sin_confirmar_marcados_ahora": marcados_ahora,
        "sin_confirmar_total": sum(1 for r in resultado if r.get("_sinConfirmar")),
        "total": len(resultado),
    }
    return resultado, stats

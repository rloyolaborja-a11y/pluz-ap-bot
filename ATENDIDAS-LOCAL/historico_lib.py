# -*- coding: utf-8 -*-
"""
historico_lib.py — Lógica COMPARTIDA para archivar un mes ya cerrado de
Atendidas AP en HISTORICO\\ (reducido + "_completa", versión admin y por
contratista), CON MERGE: una fila nueva con el mismo "SAP - Número de
reclamo" reemplaza a la que ya estaba archivada, pero el resto de filas
que ya había se conservan (no se pisa el mes entero cada vez).

La usan dos lugares distintos, para que ambos archiven exactamente igual:
  - importar_mes_historico.py — cuando le pasas a mano el Excel de un mes
    ya cerrado (arrastrándolo sobre Importar-Mes-Historico.bat, o desde la
    carpeta EXCEL-MESES-ANTERIORES\\).
  - procesar_diario.py — AUTOMÁTICAMENTE, cada vez que corres Ejecutar.bat
    y el GAB del día trae filas de un mes que ya no es el mes calendario
    actual (por ejemplo, los primeros días de setiembre siguen trayendo
    algunos días de agosto). Así, con solo correr Ejecutar.bat día a día
    como siempre, cada mes se va completando y archivando solo — sin tener
    que conseguir después el Excel completo de ese mes por separado.
"""

import json
import os
from datetime import datetime

MESES_LARGOS = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]


def iso_js(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{dt.microsecond // 1000:03d}Z"


def limpia(v):
    try:
        import pandas as pd
        if isinstance(v, float) and pd.isna(v):
            return None
    except ImportError:
        pass
    return v


def _cargar_rows(ruta):
    if not os.path.exists(ruta):
        return []
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
        return datos.get("rows", []) or []
    except Exception:
        return []


def mezclar_por_reclamo(existentes, nuevas):
    """Igual que mezclarFilasHistorico() en assets/js/procesar_ap.js — junta
    lo que ya estaba archivado con las filas nuevas; si un mismo reclamo
    (SAP - Número de reclamo) aparece en ambos, gana la versión nueva."""
    mapa = {}
    for r in (existentes or []):
        mapa[str(r.get("SAP - Número de reclamo"))] = r
    for r in (nuevas or []):
        mapa[str(r.get("SAP - Número de reclamo"))] = r
    return list(mapa.values())


def guardar_json(ruta, rows, updated_at):
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    filas = [{k: limpia(v) for k, v in r.items()} for r in rows]
    payload = {"updated_at": updated_at, "rows": filas}
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return len(filas)


def actualizar_index(historico_dir, clave, updated_at):
    ruta = os.path.join(historico_dir, "index.json")
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as f:
            indice = json.load(f)
    else:
        indice = {"meses": []}
    anio, mes = clave.split("-")
    label = f"{MESES_LARGOS[int(mes)]} {anio}"
    meses = [m for m in indice.get("meses", []) if m.get("clave") != clave]
    meses.append({"clave": clave, "label": label, "updated_at": updated_at})
    meses.sort(key=lambda m: m["clave"], reverse=True)
    indice["meses"] = meses
    os.makedirs(historico_dir, exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(indice, f, ensure_ascii=False)


def archivar_mes(historico_dir, clave, rows_completas_nuevas, id_login_fn, ahora_iso=None):
    """Archiva/actualiza UN mes (clave "YYYY-MM") en historico_dir, con
    MERGE contra lo que ya hubiera archivado ahí (ver mezclar_por_reclamo).
    rows_completas_nuevas: filas con TODAS las columnas (incl. "Poste -
    Referencias") de ese mes en el archivo que se está procesando ahora —
    el reducido (sin esa columna) y los recortes por contratista se derivan
    solos a partir de esto, así que solo hace falta pasar un juego de filas.
    id_login_fn: función que recibe el valor de la columna "Contratista" y
    devuelve el login de contratista (COBRA/LARI/PA/NORTE) o None — pásale
    id_login_de_contratista de procesar_diario.py (o el mismo import que ya
    uses en el script que llama a esto).
    Devuelve un dict con la cantidad de filas escritas, para poder loguearlo.
    """
    os.makedirs(historico_dir, exist_ok=True)
    ahora_iso = ahora_iso or iso_js(datetime.utcnow())
    anio, mes = clave.split("-")
    label = f"{MESES_LARGOS[int(mes)]} {anio}"

    # Base para el merge: primero lo que ya hubiera en el archivo "_completa"
    # (todas las columnas); si un reclamo solo existe en el archivo reducido
    # de antes (meses que se archivaron alguna vez SIN "_completa" — por
    # ejemplo, los que ya estaban armados desde antes de que existiera este
    # archivado con todas las columnas), se conserva igual, aunque le falte
    # "Poste - Referencias" (mejor eso que perder la fila). Encima de todo
    # eso, ganan siempre las filas nuevas de esta corrida.
    ruta_completa = os.path.join(historico_dir, f"{clave}_completa.json")
    ruta_reducido = os.path.join(historico_dir, f"{clave}.json")
    base = mezclar_por_reclamo(_cargar_rows(ruta_reducido), _cargar_rows(ruta_completa))
    merged_completa = mezclar_por_reclamo(base, rows_completas_nuevas)
    merged_publicas = [{k: v for k, v in r.items() if k != "Poste - Referencias"} for r in merged_completa]

    n_pub = guardar_json(os.path.join(historico_dir, f"{clave}.json"), merged_publicas, ahora_iso)
    n_comp = guardar_json(ruta_completa, merged_completa, ahora_iso)
    resultado = {"clave": clave, "label": label, "publicas": n_pub, "completas": n_comp, "contratistas": []}

    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_pub_c = [r for r in merged_publicas if id_login_fn(r.get("Contratista")) == contratista]
        filas_comp_c = [r for r in merged_completa if id_login_fn(r.get("Contratista")) == contratista]
        if not filas_pub_c and not filas_comp_c:
            continue
        n_pub_c = guardar_json(os.path.join(historico_dir, f"{clave}_{contratista.lower()}.json"), filas_pub_c, ahora_iso)
        n_comp_c = guardar_json(os.path.join(historico_dir, f"{clave}_completa_{contratista.lower()}.json"), filas_comp_c, ahora_iso)
        resultado["contratistas"].append((contratista, n_pub_c, n_comp_c))

    actualizar_index(historico_dir, clave, ahora_iso)
    return resultado

# -*- coding: utf-8 -*-
"""
procesar_diario.py — Reporte de Veredas AP (versión local)

Fuente de datos: los mismos Excel que ya usa Pendientes AP — NO hace falta
descargar nada aparte:
  - ..\\CARGA\\SAP\\   -> SAP_m0/m1/m2.xlsx (o 1 solo Excel), ZM06 (IW39).
    Ya trae todas las columnas que necesita Veredas (Grupo hojas ruta,
    Distrito, Fecha de creación, Clase actividad PM, etc.) — no hizo falta
    tocar el robot de SAP.
  - ..\\CARGA\\EXCEL\\ -> el GAP, SOLO para completar "Distrito" cuando viene
    vacío en el SAP (los casos "TRAMO SUB AP" no traen distrito propio).

QUÉ ES "VEREDAS"
-----------------
Del SAP se queda con las filas de "Grupo hojas ruta" = APEMEVER (2026-09-04,
confirmado con la usuaria comparando contra el Power BI de Veredas). De esas,
"pendiente" es todo lo que NO esté en Status de usuario CER (atendido) o CAN
(cancelado) — no hay separación tipo DT/NO DT, es un solo conjunto.

COLUMNAS
--------
Se publican las columnas del SAP tal cual vienen (por ahora, sin agregar
calculadas) + "Distrito" ya completado con el cruce contra el GAP cuando
hacía falta + "Contratista" (mapeado desde "Pto.tbjo.responsable").

AISLAMIENTO POR CONTRATISTA
----------------------------
Igual que Pendientes/Atendidas: se publica bd_actual.json (todo, lo ve Pluz)
y un bd_actual_<contratista>.json por cada uno de COBRA/LARI/PA/NORTE, solo
con sus propias filas.
"""

import glob
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timezone, timedelta

try:
    import pandas as pd
except ImportError:
    print("Falta la librería pandas. Ejecuta: pip install pandas openpyxl")
    sys.exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARGA_DIR = os.path.join(BASE_DIR, "..", "CARGA")
SAP_DIR = os.path.join(CARGA_DIR, "SAP")
GAP_DIR = os.path.join(CARGA_DIR, "EXCEL")
SALIDA_DIR = os.path.join(BASE_DIR, "SALIDA")
MAPA_CONTRATISTA_PATH = os.path.join(BASE_DIR, "REQUISITO", "mapa_contratista.json")
PENDIENTE_PUBLICAR = os.path.join(SALIDA_DIR, "bd_actual.json")
# (2026-09-23) Lista de ODMs "a vigilar" para el bloque 2 de SAP -- ver el
# mismo mecanismo en PENDIENTES-LOCAL/procesar_diario.py. Veredas viene
# DIRECTO del SAP (cada fila YA es una orden con su "Orden"), asi que aca no
# hay que esperar a que se genere el ODM -- el unico motivo por el que un
# pendiente de Veredas se "escapa" es que su mes quedo fuera del rango
# descargado (bloque 1). descargar_excel_sap.py lee y combina este archivo
# con el de Pendientes para armar la lista completa del bloque 2.
ODM_VIGILAR_PATH = os.path.join(BASE_DIR, "pendientes_odm_vigilar.json")

GRUPO_HOJAS_RUTA_VEREDAS = "APEMEVER"
ESTADOS_EXCLUIDOS = {"CER", "CAN"}  # atendido / cancelado -> ya no es pendiente

CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]

# Columnas del SAP tal cual vienen (mismo orden en que las trae el Excel) +
# las que se agregan/completan en el proceso.
COLUMNAS_SAP = [
    "Clase de orden", "Nivel Tensión", "Orden", "Status de usuario",
    "Sistema Origen", "Orden sistema origen", "Pto.tbjo.responsable",
    "Fecha de creación", "Fecha fin real", "Área de empresa",
    "CeCo responsable", "Equipo", "Denominación objeto", "Elemento PEP",
    "Clase actividad PM", "Grupo planificación", "Distrito",
    "Status de sistema", "Grupo hojas ruta",
]
COLUMNAS_CALCULADAS = ["Contratista", "Mes", "Día"]
BD_COLUMNS = COLUMNAS_SAP + COLUMNAS_CALCULADAS

MESES_CORTOS = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}


def clave_busqueda(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.upper()


def normalizar_encabezado(v):
    texto = "" if v is None else str(v)
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip()


def iso_js(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S") + ".000Z"


def limpia(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if v is pd.NaT:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return iso_js(v)
    return v


def todos_los_excel(carpeta, exts=(".xlsx", ".xls", ".xlsm")):
    candidatos = []
    for ext in exts:
        candidatos.extend(glob.glob(os.path.join(carpeta, f"*{ext}")))
    candidatos = [c for c in candidatos if not os.path.basename(c).startswith("~$")]
    return sorted(candidatos, key=os.path.getmtime, reverse=True)


_RE_SLOT_MES = re.compile(r"[_.]?m(\d+)[_.]", re.I)


def _ordenar_por_slot(paths):
    """m0 (mes actual) primero, para que al deduplicar por 'Orden' gane el
    dato más reciente. Sin slot (1 Excel manual), por fecha de archivo."""
    def clave(p):
        m = _RE_SLOT_MES.search(os.path.basename(p))
        return (0, int(m.group(1))) if m else (1, -os.path.getmtime(p))
    return sorted(paths, key=clave)


def leer_sap_muchos(paths):
    paths = _ordenar_por_slot(paths)
    dfs = []
    for p in paths:
        df = pd.read_excel(p, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
    antes = len(df)
    df = df[~df["Orden"].map(clave_busqueda).duplicated(keep="first")].reset_index(drop=True)
    if antes != len(df):
        print(f"SAP: {antes} filas en {len(paths)} Excel(s) -> {len(df)} tras unir y quitar repetidos entre meses.")
    return df


def construir_lookup_distrito_gap(gap_paths):
    """{'SAP - Número de reclamo' normalizado -> 'Poste - Distrito'}, para
    completar el Distrito de Veredas cuando el SAP lo trae vacío (casos
    'TRAMO SUB AP', que no tienen distrito propio en el SAP)."""
    lookup = {}
    for path in gap_paths:
        df_raw = pd.read_excel(path, skiprows=1, header=None, dtype=object)
        if df_raw.empty:
            continue
        header_real = df_raw.iloc[0].tolist()
        body = df_raw.iloc[1:].reset_index(drop=True)
        header_norm = [normalizar_encabezado(h) for h in header_real]
        try:
            idx_reclamo = header_norm.index(normalizar_encabezado("SAP - Número de reclamo"))
            idx_distrito = header_norm.index(normalizar_encabezado("Poste - Distrito"))
        except ValueError:
            print(f"  aviso: {os.path.basename(path)} no tiene las columnas esperadas para el cruce de Distrito.")
            continue
        for _, fila in body.iterrows():
            k = clave_busqueda(fila.iloc[idx_reclamo])
            if not k or k in lookup:
                continue
            valor = fila.iloc[idx_distrito]
            if valor is not None and not (isinstance(valor, float) and pd.isna(valor)) and str(valor).strip():
                lookup[k] = str(valor).strip()
    return lookup


def parse_fecha(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (datetime, pd.Timestamp)):
        return pd.Timestamp(valor).to_pydatetime().replace(tzinfo=None)
    s = str(valor).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def procesar(df_sap, lookup_distrito_gap, mapa_contratista):
    faltan = [c for c in ("Orden", "Grupo hojas ruta", "Status de usuario",
                           "Pto.tbjo.responsable", "Orden sistema origen") if c not in df_sap.columns]
    if faltan:
        raise ValueError(f"Al SAP le faltan columnas: {', '.join(faltan)}")

    f1 = df_sap[df_sap["Grupo hojas ruta"].astype(str).str.strip() == GRUPO_HOJAS_RUTA_VEREDAS]
    f2 = f1[~f1["Status de usuario"].astype(str).str.strip().isin(ESTADOS_EXCLUIDOS)]

    filas = []
    completados_por_gap = 0
    for _, fila in f2.iterrows():
        base = {c: fila.get(c) for c in COLUMNAS_SAP}

        distrito = base.get("Distrito")
        distrito_vacio = distrito is None or (isinstance(distrito, float) and pd.isna(distrito)) or not str(distrito).strip()
        if distrito_vacio:
            k = clave_busqueda(fila.get("Orden sistema origen"))
            desde_gap = lookup_distrito_gap.get(k)
            if desde_gap:
                base["Distrito"] = desde_gap
                completados_por_gap += 1

        base["Contratista"] = mapa_contratista.get(clave_busqueda(fila.get("Pto.tbjo.responsable")), "SIN ASIGNAR")

        f_creacion = parse_fecha(fila.get("Fecha de creación"))
        base["Fecha de creación"] = f_creacion
        base["Fecha fin real"] = parse_fecha(fila.get("Fecha fin real"))
        base["Mes"] = MESES_CORTOS[f_creacion.month] if f_creacion else None
        base["Día"] = f_creacion.strftime("%Y-%m-%d") if f_creacion else None

        filas.append(base)

    stats = {
        "total_sap": len(df_sap), "veredas": len(f1), "pendientes": len(f2),
        # Sobre f2 (pendientes), que es lo que de verdad se publica -- antes
        # esto se calculaba sobre f1 (veredas SIN filtrar CER/CAN) y no
        # cuadraba con "completados_por_gap" (que sí corre solo sobre f2).
        "distrito_vacio_en_sap": int((f2["Distrito"].isna() | (f2["Distrito"].astype(str).str.strip() == "")).sum()),
        "completados_por_gap": completados_por_gap,
    }
    return filas, stats


def guardar_excel_completo(filas, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    df = pd.DataFrame(filas, columns=BD_COLUMNS)
    ruta = os.path.join(SALIDA_DIR, f"bd_completa_{ahora.strftime('%Y-%m-%d_%H%M')}.xlsx")
    df.to_excel(ruta, index=False)
    return ruta


def _filas_a_json(filas, updated_at):
    return {"updated_at": updated_at, "rows": [{k: limpia(v) for k, v in f.items()} for f in filas]}


def guardar_json_publicar(filas, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    payload = _filas_a_json(filas, updated_at)
    with open(PENDIENTE_PUBLICAR, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return PENDIENTE_PUBLICAR, len(payload["rows"])


def guardar_json_publicar_contratistas(filas, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    resultados = []
    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_c = [f for f in filas if f.get("Contratista") == contratista]
        payload = _filas_a_json(filas_c, updated_at)
        ruta = os.path.join(SALIDA_DIR, f"bd_actual_{contratista.lower()}.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        resultados.append((contratista, ruta, len(payload["rows"])))
    return resultados


# (2026-09-23) Igual fix que en PENDIENTES-LOCAL/procesar_diario.py: el
# bloque 1 de SAP baja los ultimos RANGO_DIAS_BLOQUE_1 dias RODANTES desde
# hoy, NO "el mes calendario actual" -- hay que comparar contra eso, no
# contra mes/año.
RANGO_DIAS_BLOQUE_1 = 30


def generar_lista_odm_vigilar(filas, ahora, ruta=ODM_VIGILAR_PATH):
    """Igual criterio que Pendientes (ver PENDIENTES-LOCAL/procesar_diario.py):
    los pendientes de Veredas cuya 'Fecha de creación' cae FUERA de la
    ventana rodante de RANGO_DIAS_BLOQUE_1 dias no van a aparecer en el
    bloque 1 de SAP -- hay que vigilarlos por su Orden para que el bloque 2
    los busque en la corrida siguiente. Todos los pendientes de Veredas YA
    tienen Orden (son filas de SAP), asi que a diferencia de Pendientes no
    hace falta filtrar por "sin ODM"."""
    corte_bloque_1 = ahora - timedelta(days=RANGO_DIAS_BLOQUE_1)
    ordenes = []
    vistos = set()
    for fila in filas:
        orden = fila.get("Orden")
        if orden is None or (isinstance(orden, float) and pd.isna(orden)):
            continue
        f_creacion = fila.get("Fecha de creación")
        if not isinstance(f_creacion, (datetime, pd.Timestamp)):
            continue
        if f_creacion >= corte_bloque_1:
            continue  # dentro de la ventana rodante de 30 dias -- ya lo trae el bloque 1
        clave = clave_busqueda(orden)
        if clave == "" or clave in vistos:
            continue
        vistos.add(clave)
        ordenes.append(clave)

    payload = {"generado": iso_js(ahora), "ordenes": ordenes}
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return ruta, len(ordenes)


def main():
    print("=== Reporte de Veredas AP — procesamiento local ===")
    sap_paths = todos_los_excel(SAP_DIR)
    gap_paths = todos_los_excel(GAP_DIR)
    if not sap_paths:
        print(f"No se encontró ningún Excel en ..\\CARGA\\SAP\\ ({SAP_DIR}).")
        sys.exit(1)
    if not os.path.exists(MAPA_CONTRATISTA_PATH):
        print(f"Falta REQUISITO\\mapa_contratista.json ({MAPA_CONTRATISTA_PATH}).")
        sys.exit(1)

    print(f"SAP ({len(sap_paths)}): {', '.join(os.path.basename(p) for p in sap_paths)}")
    if gap_paths:
        print(f"GAP ({len(gap_paths)}, solo para completar Distrito): {', '.join(os.path.basename(p) for p in gap_paths)}")
    else:
        print("AVISO: no hay ningún Excel en ..\\CARGA\\EXCEL\\ — los Distrito vacíos del SAP van a quedar sin completar.")

    with open(MAPA_CONTRATISTA_PATH, encoding="utf-8") as f:
        mapa_contratista = json.load(f)

    df_sap = leer_sap_muchos(sap_paths)
    lookup_distrito = construir_lookup_distrito_gap(gap_paths) if gap_paths else {}

    filas, stats = procesar(df_sap, lookup_distrito, mapa_contratista)
    print(f"Filas: SAP total={stats['total_sap']} -> Veredas={stats['veredas']} -> "
          f"pendientes (sin CER/CAN)={stats['pendientes']}")
    print(f"Distrito vacío en el SAP: {stats['distrito_vacio_en_sap']} "
          f"-> completados cruzando con el GAP: {stats['completados_por_gap']} "
          f"-> siguen sin distrito: {stats['distrito_vacio_en_sap'] - stats['completados_por_gap']}")

    ahora = datetime.utcnow()

    ruta_odm, n_odm = generar_lista_odm_vigilar(filas, ahora)
    print(f"Lista de ODMs a vigilar (bloque 2 SAP, próxima corrida): {n_odm} orden(es) -> {ruta_odm}")

    ruta_excel = guardar_excel_completo(filas, ahora)
    print(f"Excel (solo local): {ruta_excel}")

    # (2026-09-22) "Última actualización" = hora del Excel del GAP mas
    # reciente (mismo criterio que Pendientes/Atendidas, para que las 3
    # paginas muestren la MISMA hora). Si no hay GAP disponible (es
    # opcional aca, solo se usa para completar Distrito), se cae a la hora
    # real de ahora como respaldo.
    ultima_actualizacion = (
        datetime.utcfromtimestamp(os.path.getmtime(gap_paths[0])) if gap_paths else ahora
    )

    ruta_json, n = guardar_json_publicar(filas, ultima_actualizacion)
    print(f"Listo para publicar, dashboard/exportar ({n} filas): {ruta_json}")

    print("Generando archivos recortados por contratista (aislamiento de datos)...")
    for contratista, ruta, n in guardar_json_publicar_contratistas(filas, ultima_actualizacion):
        print(f"  {contratista}: {n} filas: {ruta}")

    print("Ejecuta feed.py (o Ejecutar.bat, que ya lo hace) para publicar.")


if __name__ == "__main__":
    main()

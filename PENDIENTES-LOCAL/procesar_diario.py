# -*- coding: utf-8 -*-
"""
procesar_diario.py — Reporte de Pendientes AP (versión local)

Reproduce EXACTO el pipeline de assets/js/procesar.js del reporte web:
- Lee el Excel de ..\\CARGA\\EXCEL\\ (el más reciente que encuentre) y el de ..\\CARGA\\SAP\\.
- Cruza contra REQUISITO\\data_requisito.json (mismos mapas que usa la página web).
- Calcula Contratista, Tipo DT, Intervalo, Motivo Reclamo, LEGAL, PLAZO, etc.
- Escribe en SALIDA\\:
    - bd_completa_AAAA-MM-DD_HHMM.xlsx  -> Excel con TODAS las columnas (incluye
      datos del cliente). Se queda solo en tu computadora, nunca se sube.
    - bd_actual.json                    -> el paquete "público" (sin datos del
      cliente) que feed.py va a publicar en GitHub para todos los visitantes
      del sitio.

Requiere: pandas, openpyxl  (pip install pandas openpyxl)
"""

import json
import glob
import os
import sys
import re
import unicodedata
from datetime import datetime, timezone, timedelta

try:
    import pandas as pd
except ImportError:
    print("Falta la librería pandas. Ejecuta: pip install pandas openpyxl requests")
    sys.exit(1)

import historico_pendientes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARGA_DIR = os.path.join(BASE_DIR, "..", "CARGA")
GAP_DIR = os.path.join(CARGA_DIR, "EXCEL")
SAP_DIR = os.path.join(CARGA_DIR, "SAP")
SALIDA_DIR = os.path.join(BASE_DIR, "SALIDA")
REQUISITO_PATH = os.path.join(BASE_DIR, "REQUISITO", "data_requisito.json")
PENDIENTE_PUBLICAR = os.path.join(SALIDA_DIR, "bd_actual.json")
# "Libreta" acumulada de pendientes NO DT (ver historico_pendientes.py).
HIST_NODT_DIR = os.path.join(BASE_DIR, "HISTORICO-NODT")

# Mismo orden en que aparecen estas columnas en el Excel grande (reducido)
# que se sube hoy — así el Excel que se descarga en SALIDA\ sale con las
# columnas en el mismo orden en que el usuario ya las conoce. La lectura
# sigue siendo por NOMBRE (ver leer_por_nombre_de_columna), así que este
# orden no afecta en nada cómo se leen los datos.
GAP_COLUMNS = [
    "SAP - Número de reclamo", "SAP - Cod.SalesForce", "SAP - Motivo de reclamo",
    "SAP - Fecha de registro", "SAP - Fecha Estimada", "SAP - Estado",
    "Poste - Distrito", "Poste - Número", "SAP - Zona", "SAP - DT de Ingreso",
    "Poste - Referencias", "Cliente - Suministro", "Poste - SED", "Cliente - SED",
    "Cliente - Nombre", "Cliente - Apellido Paterno", "Cliente - Apellido Materno",
    "Cliente - Teléfono 1", "Cliente - Teléfono 2",
    # 2026-09-04: se agrega para poder armar "RELACIONADA" (ver más abajo) —
    # en Atendidas esta misma columna se usa para ELIMINAR la fila entera;
    # en Pendientes NO se elimina nada, se deja como filtro (Todos/Sin
    # relacionadas/Solo relacionadas) para que el usuario decida.
    "SAP - SAP Maestra",
    # Estas dos no vienen en el Excel reducido actual — se quedan en None,
    # igual que antes, por si algún día vuelven a incluirse.
    "SAP - Sucursal", "SAP MOVIL - Codigo Cuadrilla",
]

SAP_REQUIRED_COLUMNS = [
    "Orden", "Orden sistema origen", "Status de usuario",
    "Pto.tbjo.responsable", "Status de sistema",
]

BD_COLUMNS = [c for c in GAP_COLUMNS if c != "Cliente - SED"] + [
    "Fecha", "Contratista", "Tipo DT", "Duración", "Horas Transcurridas",
    "Día", "Intervalo", "Motivo Reclamo", "U.O. DISTRITO", "UO", "LEGAL",
    "ODM", "Status de usuario", "CONTRA", "FECHA2", "PLAZO", "CON ODM",
    "REASIGNADO SAP", "Fecha Interno", "PLAZO INTERNO", "CERRAR POR GAP",
    # "SI" si "SAP - SAP Maestra" trae dato (el caso está relacionado a otro) —
    # filtro Todos/Sin relacionadas/Solo relacionadas en Dashboard/Mapa/Buscar.
    "RELACIONADA",
    # NO DT: marca de la "libreta" acumulada (historico_pendientes.py) — True
    # si el caso ya no viene en el Excel descargado y no se pudo confirmar si
    # sigue abierto. Siempre False/None para DT y para NO DT visto en la
    # ventana actual.
    "_sinConfirmar",
]

# Motivos con plazo interno propio (no el "Fecha Estimada" que manda SAP):
# se vencen a 1 día calendario desde el registro (lámparas/AP encendido/SED
# fuera de servicio) o a 2 días calendario (zona sin AP).
MOTIVOS_1DC = {"AP ENCENDIDO", "LAMPARA APAGADA", "SED AP FUERA DE SERVICIO"}
MOTIVOS_2DC = {"ZONA SIN AP"}

# Columnas "públicas" que sí se publican en GitHub (sin datos del cliente) —
# igual que COLS_BD en assets/js/app.js.
COLS_BD_SLIM = [
    "Fecha", "Día", "Contratista", "SAP - Motivo de reclamo", "SAP - Estado",
    "REASIGNADO SAP", "Poste - Distrito", "Intervalo", "PLAZO", "CON ODM", "LEGAL",
    "Tipo DT", "Fecha Interno", "PLAZO INTERNO", "CERRAR POR GAP", "RELACIONADA", "_sinConfirmar",
]

# (2026-09-08) El corte entre las 2 pestañas del portal ahora es por la
# columna LEGAL (mapa_legal en data_requisito.json), no por un set de motivos:
#   LEGAL == "SI"  -> pestaña "LEGAL"    (antes "DT": las 9 con plazo legal)
#   LEGAL != "SI"  -> pestaña "NO LEGAL" (antes "NO DT": Poste en mal estado,
#                     Mantenimiento, Mejoramiento — y cualquier motivo nuevo
#                     sin clasificar)
# Mismo criterio que en assets/js/app.js (esLegal).
ESTADOS_PENDIENTE = {"ASIGNADA", "INGRESADA"}

# Contratistas con clave propia (aislamiento de datos) — igual que
# CONTRATISTAS_CON_CLAVE en assets/js/app.js. Los casos "sin asignar" (o de
# cualquier otro contratista) simplemente no entran en ninguno de estos
# archivos recortados — solo quedan en bd_actual.json/bd_completa.json, que
# lee únicamente el admin (Pluz).
CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]

MESES_CORTOS = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
MESES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
          7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}


def clave_busqueda(valor):
    if valor is None or valor == "":
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto.upper()


# Filtro ELIMINAR (2026-08-29, aplica a las 2 páginas): revisa TODAS las
# columnas del GAP en busca de frases que indican que el caso ya se manejó
# por otro lado (Salesforce, ODM, relacionado) — igual que en Atendidas.
_TILDES_PEND = str.maketrans({
    "Á": "A", "À": "A", "Â": "A", "Ä": "A", "Ã": "A",
    "É": "E", "È": "E", "Ê": "E", "Ë": "E",
    "Í": "I", "Ì": "I", "Î": "I", "Ï": "I",
    "Ó": "O", "Ò": "O", "Ô": "O", "Ö": "O", "Õ": "O",
    "Ú": "U", "Ù": "U", "Û": "U", "Ü": "U",
    "Ñ": "N",
})


def _norm_filtro(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    t = str(v).strip().upper().translate(_TILDES_PEND)
    t = re.sub(r"[.,;:/\\\-_()\[\]{}\"']", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _contiene_palabra_completa_filtro(texto, palabra):
    return f" {palabra} " in f" {texto} "


FRASES_ELIMINAR = [
    "SALESFORCE", "SALEFORCE", "SALESFOR", "SALES FORCE", "SALE FORCE", "SELFORCE", "SALFORCE",
    "SE ATIENDE CON ODM", "SE ATENDIO CON ODM", "ATENDIDO CON ODM",
    "ATENDIDO CON LA DENUNCIA", "ATENDIDO CON DENUNCIA",
    "TOMANORMAL RELACIONADO", "TOMA NORMAL RELACIONADO", "SE RELACIONA",
    "SE ATENDIO CON SF",
]
PALABRAS_COMPLETAS_ELIMINAR = ["RELACIONADO", "RELACIONADA", "RELACIONAR"]
_CODIGO_SF_RE_PEND = re.compile(r"^SF\d+$")


def fila_debe_eliminarse(valores):
    texto_completo = " ".join(_norm_filtro(v) for v in valores)
    if any(frase in texto_completo for frase in FRASES_ELIMINAR):
        return True
    if any(_contiene_palabra_completa_filtro(texto_completo, p) for p in PALABRAS_COMPLETAS_ELIMINAR):
        return True
    if any(_CODIGO_SF_RE_PEND.match(tok) for tok in texto_completo.split(" ")):
        return True
    # 2026-09-09: cualquier palabra con "FORCE" adentro = Salesforce mal
    # escrito ("LASAFORCE", etc.).
    if "FORCE" in texto_completo:
        return True
    return False


_RE_DMY = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$")
_RE_YMD = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$")


def parse_fecha_excel(valor):
    """Replica parseFechaExcel() de procesar.js. Devuelve datetime "naive" en UTC
    (mismo criterio que la web: la hora del Excel se trata tal cual, sin zona horaria)."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, datetime):
        return valor.replace(tzinfo=None)
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime().replace(tzinfo=None)
    if isinstance(valor, (int, float)):
        # Excel serial date, base 1899-12-30 (igual que la web)
        from datetime import timedelta
        base = datetime(1899, 12, 30)
        return base + timedelta(days=float(valor))
    s = str(valor).strip()
    m = _RE_DMY.match(s)
    if m:
        d, mo, y, h, mi, se = m.groups()
        if len(y) == 2:
            y = "20" + y
        return datetime(int(y), int(mo), int(d), int(h or 0), int(mi or 0), int(se or 0))
    m = _RE_YMD.match(s)
    if m:
        y, mo, d, h, mi, se = m.groups()
        return datetime(int(y), int(mo), int(d), int(h or 0), int(mi or 0), int(se or 0))
    return None


def fmt_fecha(dt, meses):
    if dt is None:
        return None
    return f"{dt.day:02d}-{meses[dt.month]}"


def iso_js(dt):
    """Serializa un datetime exactamente como JSON.stringify(Date) en JS:
    'YYYY-MM-DDTHH:mm:ss.000Z'."""
    if dt is None:
        return None
    return dt.strftime("%Y-%m-%dT%H:%M:%S") + ".000Z"


def buscar_periodo(horas, tabla_ordenada):
    """Busca en [[HORAS, PERIODO], ...] (ordenada asc por HORAS) el valor con
    HORAS <= horas más cercano — equivalente a merge_asof / búsqueda binaria."""
    if horas is None:
        return None
    lo, hi, res = 0, len(tabla_ordenada) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        if tabla_ordenada[mid][0] <= horas:
            res = tabla_ordenada[mid][1]
            lo = mid + 1
        else:
            hi = mid - 1
    return res


def aplicar_mapa(valor, mapa):
    k = clave_busqueda(valor)
    return mapa.get(k) if k != "" else None


def archivo_mas_reciente(carpeta, exts=(".xlsx", ".xls", ".xlsm")):
    candidatos = todos_los_excel(carpeta, exts)
    return candidatos[0] if candidatos else None


def todos_los_excel(carpeta, exts=(".xlsx", ".xls", ".xlsm")):
    """Todos los Excel de la carpeta, del MÁS NUEVO al más viejo. Con la
    ventana de 3 meses, CARGA\\EXCEL\\ y CARGA\\SAP\\ pueden tener un Excel
    por mes (mes actual, -1, -2); si hay uno solo, funciona igual que antes."""
    candidatos = []
    for ext in exts:
        candidatos.extend(glob.glob(os.path.join(carpeta, f"*{ext}")))
    candidatos = [c for c in candidatos if not os.path.basename(c).startswith("~$")]
    return sorted(candidatos, key=os.path.getmtime, reverse=True)


_RE_SLOT_MES = re.compile(r"[_.]?m(\d+)[_.]", re.I)


def _ordenar_por_slot(paths):
    """Ordena para que, al deduplicar, GANE el mes MÁS RECIENTE. Los archivos
    de slot (SAP_m0, R_m0_...) traen m0 = mes actual, m1 = mes -1, etc., así
    que m0 va primero. El robot los genera en ese mismo orden, así que su
    mtime NO sirve (m0 es el más viejo por fecha de archivo). Para Excel sin
    slot (uso manual), se cae a mtime (más nuevo primero)."""
    def clave(p):
        m = _RE_SLOT_MES.search(os.path.basename(p))
        return (0, int(m.group(1))) if m else (1, -os.path.getmtime(p))
    return sorted(paths, key=clave)


def leer_gap_muchos(paths):
    """Une varios GAP (uno por mes) en un solo DataFrame. Un mismo
    'SAP - Número de reclamo' repetido entre meses se queda con la aparición
    del mes MÁS RECIENTE."""
    paths = _ordenar_por_slot(paths)
    dfs = [leer_gap(p) for p in paths]
    df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
    antes = len(df)
    df = df[~df["SAP - Número de reclamo"].map(clave_busqueda).duplicated(keep="first")].reset_index(drop=True)
    if antes != len(df):
        print(f"GAP: {antes} filas en {len(paths)} Excel(s) -> {len(df)} tras unir y quitar repetidos entre meses.")
    return df


def leer_sap_muchos(paths):
    """Une varios SAP (uno por mes). El deduplicado por ODM lo hace
    construir_sap_lookup (se queda con la primera aparición); por eso el mes
    más reciente (m0) va primero."""
    paths = _ordenar_por_slot(paths)
    dfs = [leer_sap(p) for p in paths]
    return pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]


def normalizar_encabezado(v):
    """Para comparar encabezados de Excel sin que importen espacios de
    sobra, mayúsculas/minúsculas ni acentos (ej. 'SAP - DT Atención ' con
    espacio final sigue matcheando 'SAP - DT Atención')."""
    texto = "" if v is None else str(v)
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip()


def leer_por_nombre_de_columna(path, columnas_requeridas):
    """Lee el GAP/GAB buscando cada columna por su NOMBRE real en el
    archivo (fila 1 = título del reporte, fila 2 = encabezados), en vez de
    asumir una posición fija. Así, un Excel con columnas recortadas o
    reordenadas (como el reducido que ya no trae todas las columnas, o las
    trae en otro orden) igual se lee bien, siempre que las columnas que SÍ
    trae tengan el nombre esperado. Las que falten quedan en None (se
    manejan con los mismos fallbacks que ya existen más abajo, ej. Zona ->
    U.O. DISTRITO)."""
    df_raw = pd.read_excel(path, skiprows=1, header=None, dtype=object)
    header_real = df_raw.iloc[0].tolist() if len(df_raw) else []
    body = df_raw.iloc[1:].reset_index(drop=True)
    header_norm = [normalizar_encabezado(h) for h in header_real]
    data = {}
    for col in columnas_requeridas:
        objetivo = normalizar_encabezado(col)
        idx = header_norm.index(objetivo) if objetivo in header_norm else -1
        data[col] = body.iloc[:, idx] if idx != -1 else pd.Series([None] * len(body))
    return pd.DataFrame(data)


def leer_gap(path):
    return leer_por_nombre_de_columna(path, GAP_COLUMNS)


def leer_sap(path):
    df = pd.read_excel(path, dtype=object)
    faltantes = [c for c in SAP_REQUIRED_COLUMNS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Falta columna en SAP: {', '.join(faltantes)}")
    return df


def construir_sap_lookup(df_sap):
    tiene_odm_real = "ODM REAL" in df_sap.columns
    lookup = {}
    for _, row in df_sap.iterrows():
        estado_sistema = str(row.get("Status de sistema") or "").strip().upper()
        if estado_sistema.endswith("VER"):
            continue
        clave_modelo = row.get("ODM REAL") if tiene_odm_real else row.get("Orden sistema origen")
        k = clave_busqueda(clave_modelo)
        if k == "" or k in lookup:
            continue
        lookup[k] = {
            "Orden": row.get("Orden"),
            "Status de usuario": row.get("Status de usuario"),
            "Pto.tbjo.responsable": row.get("Pto.tbjo.responsable"),
        }
    return lookup


def procesar(df_gap, df_sap, requisito):
    sap_lookup = construir_sap_lookup(df_sap)

    # Todos los Nº de reclamo que trae el GAP de esta corrida (cualquier
    # estado) — lo usa la "libreta" NO DT para decidir qué casos viejos ya
    # se cerraron (los vio el Excel) y cuáles quedaron fuera de ventana.
    reclamos_en_gap = {
        clave_busqueda(v) for v in df_gap["SAP - Número de reclamo"].tolist()
        if clave_busqueda(v)
    }

    mapa_contratista = requisito["mapa_contratista"]
    tabla_periodo = requisito["tabla_periodo"]
    mapa_motivo = requisito["mapa_motivo"]
    # (2026-09-09) "Tipo DT" y "LEGAL" ahora salen del MOTIVO (tabla oficial de
    # distribución), NO de "SAP - DT de Ingreso" ni de un mapa aparte. Un
    # motivo sin clasificar cae en "NO DT" (=> LEGAL "NO").
    mapa_deficiencia = requisito.get("mapa_deficiencia", {})
    mapa_distrito = requisito["mapa_distrito"]
    mapa_zona = requisito["mapa_zona"]
    mapa_sed = requisito["mapa_sed"]
    mapa_poste = requisito["mapa_poste"]

    ahora = datetime.utcnow()

    # 2026-09-04: BUG encontrado por el usuario -- "hoy_medianoche" (para
    # decidir PLAZO=DP/FDP) se calculaba con "ahora" en UTC real, pero "SAP -
    # Fecha Estimada"/"SAP - Fecha de registro" vienen del Excel tal cual, sin
    # zona horaria (ver parse_fecha_excel: es la hora de Lima/Perú "pelada").
    # Perú es UTC-5, así que entre las 19:00 y 23:59 hora Perú, UTC ya cayó
    # en el día siguiente: hoy_medianoche quedaba en el día de MAÑANA y
    # cualquier pendiente con vencimiento HOY (aún no vencido) se marcaba
    # "FDP" (vencido) por error.
    #
    # 2026-09-04 (corrección del mismo día -- el primer arreglo se pasó de
    # rosca): NO se puede simplemente restarle 5 horas a "ahora" y usar ESE
    # mismo "ahora" para todo -- "ahora" también se usa para "updated_at"
    # (sale tal cual en bd_actual.json/bd_completa.json) y esa fecha SÍ es un
    # instante UTC real de verdad (el navegador la convierte a hora de Lima
    # al mostrarla, ver fmtFechaHora en app.js, que fuerza timeZone:
    # "America/Lima" -- a diferencia de "Fecha"/"Fecha Interno", que son
    # Lima "disfrazada" y se muestran forzando timeZone:"UTC"). Correrle 5
    # horas a "ahora" globalmente hacía que "Última modificación de datos"
    # se mostrara 5 horas ATRASADA (el navegador le restaba OTRAS 5 horas
    # encima). Por eso "ahora" se deja intacto (UTC real, como siempre) y se
    # arma una variable APARTE, "ahora_lima", solo para hoy_medianoche --
    # mismo patrón que ya usa ATENDIDAS-LOCAL/procesar_diario.py (ahora_lima
    # ahí también queda acotado a un solo uso puntual, nunca reemplaza al
    # "ahora" que se serializa en updated_at).
    ahora_lima = ahora - timedelta(hours=5)
    hoy_medianoche = datetime(ahora_lima.year, ahora_lima.month, ahora_lima.day)

    errores = 0
    eliminadas_por_filtro = 0
    filas_bd = []
    for _, fila in df_gap.iterrows():
        if fila_debe_eliminarse(fila.to_dict().values()):
            eliminadas_por_filtro += 1
            continue
        fecha_registro = parse_fecha_excel(fila["SAP - Fecha de registro"])
        fecha_estimada = parse_fecha_excel(fila["SAP - Fecha Estimada"])
        if fecha_registro is None or fecha_estimada is None:
            errores += 1
            continue

        clave_reclamo = clave_busqueda(fila["SAP - Número de reclamo"])
        sap_match = sap_lookup.get(clave_reclamo)
        odm = sap_match["Orden"] if sap_match and sap_match.get("Orden") is not None else "SIN ODM"
        status_usuario = sap_match["Status de usuario"] if sap_match and sap_match.get("Status de usuario") is not None else "SIN ODM"
        contra0 = sap_match.get("Pto.tbjo.responsable") if sap_match else ""
        if contra0 is None or (isinstance(contra0, float) and pd.isna(contra0)):
            contra0 = ""

        horas_transcurridas = int((ahora - fecha_registro).total_seconds() // 3600)
        # "Tipo DT" = deficiencia según el motivo (DT1..DT6 / "NO DT").
        tipo_dt = aplicar_mapa(fila["SAP - Motivo de reclamo"], mapa_deficiencia) or "NO DT"
        # "LEGAL" va SIEMPRE de la mano: todo DT* es LEGAL "SI", "NO DT" es "NO".
        legal = "SI" if str(tipo_dt).upper().startswith("DT") else "NO"
        intervalo = buscar_periodo(horas_transcurridas, tabla_periodo)
        motivo_reclamo = aplicar_mapa(fila["SAP - Motivo de reclamo"], mapa_motivo)
        uo_distrito = aplicar_mapa(fila["Poste - Distrito"], mapa_distrito)
        uo = aplicar_mapa(fila["SAP - Zona"], mapa_zona)

        contra = contra0
        if not contra:
            # (2026-09-09) SED del POSTE, no del cliente: "Poste - SED" viene
            # lleno en ~99.6% de los pendientes; "Cliente - SED" en ~67%.
            # "Cliente - SED" queda solo de respaldo por si el poste no trae.
            contra = aplicar_mapa(fila["Poste - SED"], mapa_sed)
            if not contra:
                contra = aplicar_mapa(fila["Cliente - SED"], mapa_sed)
            if not contra:
                contra = aplicar_mapa(fila["Poste - Número"], mapa_poste)
            if not contra:
                contra = ""
        contratista = aplicar_mapa(contra, mapa_contratista)

        plazo = "DP" if fecha_estimada > hoy_medianoche else "FDP"
        con_odm = "SIN ODM" if odm == "SIN ODM" else "CON ODM"
        # (2026-09-08) Reasignado se toma del GAP (columna "SAP - Estado" ==
        # "REASIGNADA"), no del SAP ("Status de usuario" == "REA"). La columna
        # sigue llamándose "REASIGNADO SAP" para no romper filtros/exportables.
        reasignado = "SI" if normalizar_encabezado(fila.get("SAP - Estado")) == "REASIGNADA" else "NO"
        # La orden SAP (ODM) ya se cerró (Status de usuario = "CER") aunque el
        # reclamo (GAP) todavía figure "asignada" — bandera para filtrar esos
        # casos en el Mapa y en Buscar y exportar, sin ocultarlos del conteo.
        cerrar_por_gap = "SI" if status_usuario == "CER" else "NO"
        relacionada = "SI" if clave_busqueda(fila.get("SAP - SAP Maestra")) else "NO"

        motivo_norm = normalizar_encabezado(fila["SAP - Motivo de reclamo"])
        fecha_venc_interno = None
        plazo_interno = None
        if motivo_norm in MOTIVOS_1DC or motivo_norm in MOTIVOS_2DC:
            dias = 1 if motivo_norm in MOTIVOS_1DC else 2
            fecha_venc_interno = fecha_registro + timedelta(days=dias)
            plazo_interno = "DP" if fecha_venc_interno > hoy_medianoche else "FDP"

        fila_bd = {c: fila[c] for c in GAP_COLUMNS}
        fila_bd["SAP - Fecha de registro"] = fecha_registro
        fila_bd["SAP - Fecha Estimada"] = fecha_estimada
        fila_bd["Fecha"] = fmt_fecha(fecha_estimada, MESES_CORTOS)
        fila_bd["Contratista"] = contratista
        fila_bd["Tipo DT"] = tipo_dt
        fila_bd["Duración"] = int((ahora - fecha_registro).total_seconds() * 1000)
        fila_bd["Horas Transcurridas"] = horas_transcurridas
        fila_bd["Día"] = fmt_fecha(fecha_registro, MESES_CORTOS)
        fila_bd["Intervalo"] = intervalo
        fila_bd["Motivo Reclamo"] = motivo_reclamo
        fila_bd["U.O. DISTRITO"] = uo_distrito
        fila_bd["UO"] = uo
        fila_bd["LEGAL"] = legal
        fila_bd["ODM"] = odm
        fila_bd["Status de usuario"] = status_usuario
        fila_bd["CONTRA"] = contra
        fila_bd["FECHA2"] = MESES[fecha_registro.month]
        fila_bd["PLAZO"] = plazo
        fila_bd["CON ODM"] = con_odm
        fila_bd["REASIGNADO SAP"] = reasignado
        fila_bd["Fecha Interno"] = fmt_fecha(fecha_venc_interno, MESES_CORTOS) if fecha_venc_interno else None
        fila_bd["PLAZO INTERNO"] = plazo_interno
        fila_bd["CERRAR POR GAP"] = cerrar_por_gap
        fila_bd["RELACIONADA"] = relacionada
        # Fecha cruda (no es columna de BD_COLUMNS, solo sirve para ordenar
        # cronológicamente el eje del gráfico "FECHA VENCIMIENTO INTERNO").
        fila_bd["_fechaVencInternoRaw"] = fecha_venc_interno
        filas_bd.append(fila_bd)

    if errores > 0:
        raise ValueError(f"GAP contiene {errores} filas con fechas inválidas")

    asignar_contratista_por_concurrencia(filas_bd)

    return filas_bd, ahora, reclamos_en_gap


def asignar_contratista_por_concurrencia(filas_bd):
    """Casos 'sin asignar' (muy pocos, la excepción — cuando Cuadrilla/SED/
    Poste no mapean a ningún contratista conocido): se les asigna el
    contratista que más atiende ese mismo distrito, según los propios datos
    de esta corrida (concurrencia). Si el distrito no tiene ningún caso ya
    asignado en esta corrida, el caso se queda sin asignar (solo lo ve el
    admin/Pluz)."""
    concurrencia_por_distrito = {}
    for f in filas_bd:
        if not f.get("Contratista"):
            continue
        distrito = f.get("Poste - Distrito")
        if not distrito:
            continue
        m = concurrencia_por_distrito.setdefault(distrito, {})
        m[f["Contratista"]] = m.get(f["Contratista"], 0) + 1

    def contratista_mas_frecuente(distrito):
        m = concurrencia_por_distrito.get(distrito)
        if not m:
            return None
        return max(m.items(), key=lambda kv: kv[1])[0]

    for f in filas_bd:
        if not f.get("Contratista"):
            asignado = contratista_mas_frecuente(f.get("Poste - Distrito"))
            if asignado:
                f["Contratista"] = asignado


def guardar_excel_completo(filas_bd, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    df = pd.DataFrame(filas_bd, columns=BD_COLUMNS)
    nombre = f"bd_completa_{ahora.strftime('%Y-%m-%d_%H%M')}.xlsx"
    ruta = os.path.join(SALIDA_DIR, nombre)
    df.to_excel(ruta, index=False)
    return ruta


def guardar_json_publicar(filas_bd, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    slim = []
    for fila in filas_bd:
        o = {c: fila.get(c) for c in COLS_BD_SLIM}
        fe = fila.get("SAP - Fecha Estimada")
        o["_fechaTS"] = int(fe.replace(tzinfo=timezone.utc).timestamp() * 1000) if fe else None
        fi = fila.get("_fechaVencInternoRaw")
        o["_fechaInternoTS"] = int(fi.replace(tzinfo=timezone.utc).timestamp() * 1000) if fi else None
        # "Día" (fecha de registro) también necesita su TS para ordenar
        # cronológicamente el filtro "Fecha de inicio" en el Dashboard.
        fr = fila.get("SAP - Fecha de registro")
        o["_diaTS"] = int(fr.replace(tzinfo=timezone.utc).timestamp() * 1000) if isinstance(fr, (pd.Timestamp, datetime)) else None
        # Convierte cualquier valor no serializable (NaN, Timestamp, etc.)
        for k, v in list(o.items()):
            if isinstance(v, float) and pd.isna(v):
                o[k] = None
            elif isinstance(v, (pd.Timestamp, datetime)):
                o[k] = iso_js(v)
        o["_sinConfirmar"] = bool(o.get("_sinConfirmar"))
        slim.append(o)
    payload = {"updated_at": updated_at, "rows": slim}
    with open(PENDIENTE_PUBLICAR, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return PENDIENTE_PUBLICAR, len(slim)


def limpia_valor(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return iso_js(v)
    return v


def guardar_json_completo(filas_bd, ahora):
    """BD completa (con datos del cliente) en JSON, para que feed.py la
    publique — es lo que alimenta 'Buscar y exportar' para todos los que
    entren al sitio, sin que tengan que reprocesar nada."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    filas = [{c: limpia_valor(fila.get(c)) for c in BD_COLUMNS} for fila in filas_bd]
    payload = {"updated_at": updated_at, "rows": filas}
    ruta = os.path.join(SALIDA_DIR, "bd_completa.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return ruta, len(filas)


def _filas_del_contratista(filas_bd, contratista):
    return [f for f in filas_bd if str(f.get("Contratista") or "").strip().upper() == contratista]


def guardar_json_publicar_contratistas(filas_bd, ahora):
    """Un bd_actual_<contratista>.json por cada contratista con clave propia
    — SOLO sus filas, ninguna de los demás ni de 'sin asignar'."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    resultados = []
    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_c = _filas_del_contratista(filas_bd, contratista)
        slim = []
        for fila in filas_c:
            o = {c: fila.get(c) for c in COLS_BD_SLIM}
            fe = fila.get("SAP - Fecha Estimada")
            o["_fechaTS"] = int(fe.replace(tzinfo=timezone.utc).timestamp() * 1000) if fe else None
            fi = fila.get("_fechaVencInternoRaw")
            o["_fechaInternoTS"] = int(fi.replace(tzinfo=timezone.utc).timestamp() * 1000) if fi else None
            fr = fila.get("SAP - Fecha de registro")
            o["_diaTS"] = int(fr.replace(tzinfo=timezone.utc).timestamp() * 1000) if isinstance(fr, (pd.Timestamp, datetime)) else None
            for k, v in list(o.items()):
                if isinstance(v, float) and pd.isna(v):
                    o[k] = None
                elif isinstance(v, (pd.Timestamp, datetime)):
                    o[k] = iso_js(v)
            o["_sinConfirmar"] = bool(o.get("_sinConfirmar"))
            slim.append(o)
        payload = {"updated_at": updated_at, "rows": slim}
        ruta = os.path.join(SALIDA_DIR, f"bd_actual_{contratista.lower()}.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        resultados.append((contratista, ruta, len(slim)))
    return resultados


def guardar_json_completo_contratistas(filas_bd, ahora):
    """Un bd_completa_<contratista>.json por cada contratista con clave
    propia — es lo que alimenta 'Buscar y exportar' y el Mapa para ESE
    contratista, con SOLO sus propios datos de cliente."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    resultados = []
    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_c = _filas_del_contratista(filas_bd, contratista)
        filas = [{c: limpia_valor(fila.get(c)) for c in BD_COLUMNS} for fila in filas_c]
        payload = {"updated_at": updated_at, "rows": filas}
        ruta = os.path.join(SALIDA_DIR, f"bd_completa_{contratista.lower()}.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        resultados.append((contratista, ruta, len(filas)))
    return resultados


# ---------------------------------------------------------------------------
# "Libreta" acumulada de pendientes NO DT
# ---------------------------------------------------------------------------

_RE_ISO_JS = re.compile(r"^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})")


def _parse_iso_js(s):
    """Convierte 'YYYY-MM-DDTHH:MM:SS.000Z' (lo que escribe iso_js) de vuelta
    a datetime naive. Cae a parse_fecha_excel para otros formatos."""
    if not isinstance(s, str) or not s:
        return None
    m = _RE_ISO_JS.match(s)
    if m:
        y, mo, d, h, mi, se = (int(x) for x in m.groups())
        return datetime(y, mo, d, h, mi, se)
    return parse_fecha_excel(s)


def _es_legal(fila):
    return normalizar_encabezado(fila.get("LEGAL")) == "SI"


def _es_nolegal(fila):
    return not _es_legal(fila)


def _es_pendiente_vigente(fila):
    """Mismo criterio que filtrarPagina() en assets/js/app.js: reclamo no
    reasignado y en estado ASIGNADA/INGRESADA."""
    if normalizar_encabezado(fila.get("REASIGNADO SAP")) != "NO":
        return False
    return normalizar_encabezado(fila.get("SAP - Estado")) in ESTADOS_PENDIENTE


def _deshidratar_nodt(fila):
    """Pasa una fila de filas_bd (con datetimes) a un dict 100% JSON, para
    guardarlo en la libreta. Conserva _fechaVencInternoRaw como texto ISO."""
    d = {c: limpia_valor(fila.get(c)) for c in BD_COLUMNS}
    fv = fila.get("_fechaVencInternoRaw")
    if isinstance(fv, (pd.Timestamp, datetime)):
        d["_fechaVencInternoRaw"] = iso_js(fv)
    elif isinstance(fv, str):
        d["_fechaVencInternoRaw"] = fv
    else:
        d["_fechaVencInternoRaw"] = None
    d["_sinConfirmar"] = bool(fila.get("_sinConfirmar", False))
    return d


def _rehidratar_nodt(d):
    """Inversa de _deshidratar_nodt: deja la fila de la libreta con la misma
    forma que una fila fresca de filas_bd (datetimes reales) para que el
    resto del pipeline (slim, completa, _fechaTS, etc.) la trate igual."""
    f = dict(d)
    for col in ("SAP - Fecha de registro", "SAP - Fecha Estimada"):
        v = f.get(col)
        if isinstance(v, str):
            f[col] = _parse_iso_js(v)
    fv = f.get("_fechaVencInternoRaw")
    f["_fechaVencInternoRaw"] = _parse_iso_js(fv) if isinstance(fv, str) else fv
    f["_sinConfirmar"] = bool(f.get("_sinConfirmar", False))
    return f


def aplicar_libreta_nodt(filas_bd, reclamos_en_gap, ahora):
    """LEGAL: todas las pendientes vigentes de la ventana descargada (3 meses)
    — antes solo mostraba ~1 mes. NO LEGAL: pasa por la "libreta" acumulada
    para no perder pendientes más viejos que el Excel. Devuelve la lista final
    de filas a publicar."""
    filas_legal = [f for f in filas_bd if _es_legal(f) and _es_pendiente_vigente(f)]
    nolegal_pendientes = [f for f in filas_bd if _es_nolegal(f) and _es_pendiente_vigente(f)]

    # La libreta arrastra casos de cuando el corte era por motivo (varios de
    # esos motivos ahora son LEGAL). Se sacan de una vez: la libreta solo
    # guarda NO LEGAL de acá en adelante.
    acumulado = [r for r in historico_pendientes.cargar(HIST_NODT_DIR) if not _es_legal(r)]
    frescas = [_deshidratar_nodt(f) for f in nolegal_pendientes]
    merged, stats = historico_pendientes.mezclar(
        acumulado, frescas, reclamos_en_gap, iso_js(ahora))
    historico_pendientes.guardar(HIST_NODT_DIR, merged, iso_js(ahora))

    print(f"LEGAL: {len(filas_legal)} pendientes en la ventana. "
          f"Libreta NO LEGAL: {stats['en_libreta_antes']} antes -> "
          f"{stats['pendientes_frescos']} vistos pendientes en esta corrida, "
          f"{stats['retirados']} retirados (cerrados/atendidos), "
          f"{stats['sin_confirmar_total']} sin confirmar. Total libreta: {stats['total']}.")

    nodt_finales = [_rehidratar_nodt(d) for d in merged]
    return filas_legal + nodt_finales


def main():
    print("=== Reporte de Pendientes AP — procesamiento local ===")
    gap_paths = todos_los_excel(GAP_DIR)
    sap_paths = todos_los_excel(SAP_DIR)
    if not gap_paths:
        print(f"No se encontró ningún Excel en ..\\CARGA\\EXCEL\\ ({GAP_DIR}).")
        sys.exit(1)
    if not sap_paths:
        print(f"No se encontró ningún Excel en ..\\CARGA\\SAP\\ ({SAP_DIR}).")
        sys.exit(1)
    if not os.path.exists(REQUISITO_PATH):
        print(f"Falta REQUISITO\\data_requisito.json ({REQUISITO_PATH}).")
        sys.exit(1)

    print(f"GAP ({len(gap_paths)}): {', '.join(os.path.basename(p) for p in gap_paths)}")
    print(f"SAP ({len(sap_paths)}): {', '.join(os.path.basename(p) for p in sap_paths)}")
    with open(REQUISITO_PATH, encoding="utf-8") as f:
        requisito = json.load(f)

    df_gap = leer_gap_muchos(gap_paths)
    df_sap = leer_sap_muchos(sap_paths)
    filas_bd, ahora, reclamos_en_gap = procesar(df_gap, df_sap, requisito)
    print(f"Filas procesadas (ventana descargada): {len(filas_bd)}")

    # LEGAL: todas las pendientes de la ventana de 3 meses. NO LEGAL: pasa por
    # la "libreta" acumulada para no perder pendientes más viejos que el Excel.
    filas_bd = aplicar_libreta_nodt(filas_bd, reclamos_en_gap, ahora)
    print(f"Filas a publicar (LEGAL ventana + NO LEGAL libreta): {len(filas_bd)}")

    ruta_excel = guardar_excel_completo(filas_bd, ahora)
    print(f"Excel completo (con datos de cliente, solo local): {ruta_excel}")

    ruta_json, n = guardar_json_publicar(filas_bd, ahora)
    print(f"Listo para publicar, dashboard ({n} filas): {ruta_json}")

    ruta_completa, n2 = guardar_json_completo(filas_bd, ahora)
    print(f"Listo para publicar, BD completa/Buscar y exportar ({n2} filas): {ruta_completa}")

    print("Generando archivos recortados por contratista (aislamiento de datos)...")
    for contratista, ruta, n in guardar_json_publicar_contratistas(filas_bd, ahora):
        print(f"  {contratista}: dashboard ({n} filas): {ruta}")
    for contratista, ruta, n in guardar_json_completo_contratistas(filas_bd, ahora):
        print(f"  {contratista}: BD completa ({n} filas): {ruta}")

    print("Ejecuta feed.py (o Ejecutar.bat, que ya lo hace) para subir todo a GitHub.")


if __name__ == "__main__":
    main()

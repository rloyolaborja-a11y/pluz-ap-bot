# -*- coding: utf-8 -*-
"""
procesar_diario.py — Reporte de Atendidas AP (versión local)

Reproduce EXACTO el pipeline de assets/js/procesar_ap.js del reporte web:
- Lee el Excel de ..\\CARGA\\EXCEL\\ (el más reciente que encuentre).
- Filtra Estado = ATENDIDA y Motivo de reclamo en (Lámpara apagada, Zona sin
  AP, Sed AP fuera de servicio).
- Deduplica por Cod.SalesForce (se queda con la primera aparición).
- Clasifica Actividad/Rubro con MAESTROS\\criterios.json (misma lógica que la
  macro ClasificarActividadRubro), calcula Días hábiles/Rango contra
  MAESTROS\\feriados.json, y Contratista contra MAESTROS\\tecnicos.json.
- Escribe en SALIDA\\:
    - bd_completa_AAAA-MM-DD_HHMM.xlsx  -> Excel con TODAS las columnas
      (incluye "Poste - Referencias"). Solo referencia local, no se publica.
    - bd_actual.json                    -> paquete para el Dashboard.
    - bd_completa.json                  -> BD completa (con "Poste - Referencias")
      para que feed.py la publique y "Buscar y exportar" esté siempre activo
      con los últimos datos, sin que nadie tenga que reprocesar nada.

Requiere: pandas, openpyxl  (pip install pandas openpyxl)
"""

import json
import glob
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone

try:
    import pandas as pd
except ImportError:
    print("Falta la librería pandas. Ejecuta: pip install pandas openpyxl requests")
    sys.exit(1)

import historico_lib


def clave_mes_actual_lima():
    """"YYYY-MM" del mes calendario actual en hora de Lima (UTC-5, sin
    horario de verano) — mismo criterio que mesActualLima() en app.js /
    procesar_ap.js, para que "mes actual" signifique lo mismo en la página
    y acá."""
    ahora_lima = datetime.utcnow() - timedelta(hours=5)
    return ahora_lima.strftime("%Y-%m")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARGA_DIR = os.path.join(BASE_DIR, "..", "CARGA")
GAB_DIR = os.path.join(CARGA_DIR, "EXCEL")
SALIDA_DIR = os.path.join(BASE_DIR, "SALIDA")
MAESTROS_DIR = os.path.join(BASE_DIR, "MAESTROS")
HISTORICO_DIR = os.path.join(BASE_DIR, "HISTORICO")
PENDIENTE_PUBLICAR = os.path.join(SALIDA_DIR, "bd_actual.json")
# (2026-09-16) Misma tabla motivo->DT que usa Pendientes, para poder filtrar
# Atendidas por LEGAL/NO LEGAL y Tipo DT (Mapa/Buscar y exportar) -- pedido
# del usuario, mismo criterio, no uno nuevo.
REQUISITO_PATH = os.path.join(BASE_DIR, "..", "PENDIENTES-LOCAL", "REQUISITO", "data_requisito.json")

# (2026-09-16) Estos 3 son los ÚNICOS motivos que entran al DASHBOARD (sus
# gráficos LÁMPARA/ZONA SIN AP y la clasificación Actividad/Rubro son
# específicos de ellos) -- pero desde ahora YA NO filtran qué entra a la
# base completa: el Mapa y "Buscar y exportar" necesitan TODOS los motivos
# atendidos, no solo estos 3. Cada fila lleva "_dashboardValido" para que el
# Dashboard siga mostrando solo lo de siempre sin tener que tocarlo.
MOTIVOS_VALIDOS = {"LAMPARA APAGADA", "ZONA SIN AP", "SED AP FUERA DE SERVICIO"}


def cargar_mapa_deficiencia():
    """motivo (normalizado) -> "DT1".."DT6"/"NO DT", desde el mismo
    data_requisito.json que usa Pendientes (mapa_deficiencia)."""
    try:
        with open(REQUISITO_PATH, encoding="utf-8") as f:
            requisito = json.load(f)
    except (FileNotFoundError, ValueError):
        return {}
    return {norm(k): v for k, v in (requisito.get("mapa_deficiencia") or {}).items()}


def cargar_mapa_zona():
    """"SAP - Zona" (normalizado) -> UO (ej. "NORTE"), del mismo
    data_requisito.json que usa Pendientes (mapa_zona)."""
    try:
        with open(REQUISITO_PATH, encoding="utf-8") as f:
            requisito = json.load(f)
    except (FileNotFoundError, ValueError):
        return {}
    return {norm(k): v for k, v in (requisito.get("mapa_zona") or {}).items()}
MESES_CORTOS = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

# Mismo orden en que aparecen estas columnas en el Excel grande (reducido)
# que se sube hoy — así el Excel exportado sale en el mismo orden en que el
# usuario ya las conoce. La lectura sigue siendo por NOMBRE (ver
# leer_por_nombre_de_columna), así que este orden no afecta en nada cómo se
# leen los datos.
GAB_COLUMNS = [
    "SAP - Número de reclamo", "SAP - Cod. SAP", "SAP - Cod.SalesForce", "SAP - Motivo de reclamo",
    "SAP - Fecha de registro", "SAP - Fecha de Atención", "SAP - Fecha Estimada", "SAP - Estado",
    "SAP - Zona", "Poste - Distrito", "SAP - Observaciones en Atención", "SAP - SAP Maestra",
    "Poste - Referencias", "SAP - Motivo falla técnica", "Técnico Ejecutor",
    # Estas dos no vienen en el Excel reducido actual — se quedan en None,
    # igual que antes, por si algún día vuelven a incluirse.
    "SAP - Sucursal", "SAP MOVIL - Codigo Cuadrilla",
]
# "SAP - Cod. SAP" = la ODM (código que empieza con 600...). 2026-09-05: se
# agregó porque la pestaña "Correcciones" la muestra como columna "ODM".

COLS_CALCULADAS = [
    "Días hábiles", "Rango", "Actividad", "Rubro", "Contratista",
    "Dias calendario", "Rango calendario", "Dia Atención", "Día", "Nombre del mes", "EJE X",
    "PLAZO", "UO",
]
BD_COLUMNS_COMPLETAS = GAB_COLUMNS + COLS_CALCULADAS
BD_COLUMNS_PUBLICAS = [c for c in GAB_COLUMNS if c != "Poste - Referencias"] + COLS_CALCULADAS

# Contratistas con clave propia (aislamiento de datos) — igual que en
# PENDIENTES-LOCAL/procesar_diario.py y assets/js/app.js.
CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]

# El campo "Contratista" del reporte se muestra tal cual viene en
# MAESTROS/tecnicos.json ("PA PERÚ", "COBRA NCH", etc.) — eso NO se toca.
# Pero para decidir a qué login de contratista (COBRA/LARI/PA/NORTE) le
# corresponde cada fila (por ejemplo, al armar bd_actual_<id>.json para el
# acceso aislado de cada contratista), "PA PERÚ" es en realidad PA y
# "COBRA NCH" es en realidad NORTE. "DOMINION" no equivale a ninguno de
# los 4, así que no entra en ningún login.
EQUIVALENCIA_LOGIN_CONTRATISTA = {
    "PA PERÚ": "PA",
    "PA PERU": "PA",
    "COBRA NCH": "NORTE",
}


def id_login_de_contratista(valor):
    clave = norm(valor)
    return EQUIVALENCIA_LOGIN_CONTRATISTA.get(clave, clave)


# Normalización "fuerte" (2026-08-29, misma que en procesar_ap.js): mayúsculas,
# sin tildes/Ñ, signos de puntuación a espacio, espacios repetidos colapsados.
_TILDES = str.maketrans({
    "Á": "A", "À": "A", "Â": "A", "Ä": "A", "Ã": "A",
    "É": "E", "È": "E", "Ê": "E", "Ë": "E",
    "Í": "I", "Ì": "I", "Î": "I", "Ï": "I",
    "Ó": "O", "Ò": "O", "Ô": "O", "Ö": "O", "Õ": "O",
    "Ú": "U", "Ù": "U", "Û": "U", "Ü": "U",
    "Ñ": "N",
})
import re as _re
_PUNTUACION_RE = _re.compile(r"[.,;:/\\\-_()\[\]{}\"']")
_ESPACIOS_RE = _re.compile(r"\s+")


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    t = str(v).strip().upper().translate(_TILDES)
    t = _PUNTUACION_RE.sub(" ", t)
    t = _ESPACIOS_RE.sub(" ", t).strip()
    return t


def contiene_palabra_completa(texto, palabra):
    return f" {palabra} " in f" {texto} "


def contiene_tab(texto):
    return contiene_palabra_completa(texto, "TAB") or contiene_palabra_completa(texto, "TAB.")


def contiene_sub_valido(texto):
    return any(tok == "SUB" or tok.startswith("SUBT") for tok in texto.split(" "))


def clave_dedup(v):
    if v is None or v == "":
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    t = str(v).strip()
    if t.endswith(".0"):
        t = t[:-2]
    return t.upper()


def parse_fecha(valor):
    if valor is None or valor == "":
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, datetime):
        return valor.replace(tzinfo=None)
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime().replace(tzinfo=None)
    if isinstance(valor, (int, float)):
        base = datetime(1899, 12, 30)
        return base + timedelta(days=float(valor))
    import re
    s = str(valor).strip()
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$", s)
    if m:
        d, mo, y, h, mi, se = m.groups()
        if len(y) == 2:
            y = "20" + y
        return datetime(int(y), int(mo), int(d), int(h or 0), int(mi or 0), int(se or 0))
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$", s)
    if m:
        y, mo, d, h, mi, se = m.groups()
        return datetime(int(y), int(mo), int(d), int(h or 0), int(mi or 0), int(se or 0))
    return None


def solo_fecha(dt):
    return datetime(dt.year, dt.month, dt.day)


def dias_habiles(freg_medianoche, fate_medianoche, feriados_habil_set):
    cnt = 0
    d = freg_medianoche + timedelta(days=1)
    while d <= fate_medianoche:
        if d.strftime("%Y-%m-%d") in feriados_habil_set:
            cnt += 1
        d += timedelta(days=1)
    return cnt


def rango_de(dh):
    if dh is None:
        return None
    if dh <= 1:
        return "1D"
    if dh == 2:
        return "2D"
    if dh == 3:
        return "3D"
    return "Más de 3D"


# 2026-09-05: se reemplazó la clasificación por la del macro v2
# (Clasificacion_AP_Reestructurada_v2), validado por el jefe del equipo.
# Cambios grandes vs. la versión anterior (criterios.json):
#  - Lo que no matchea ninguna regla ya NO copia el Motivo: queda como
#    "No se encuentra" (esos casos se corrigen a mano en la pestaña
#    "Correcciones" de Atendidas — ver aplicar_correcciones más abajo).
#  - El Motivo se compara EXACTO (no "contiene"), salvo la regla de SED
#    que sí es "contiene".
#  - Reglas de normalización ("TODO NORMAL", "AP NORMALIZADO", etc.) suben
#    a prioridad 1/2 y le ganan a las de palabra clave.
#  - Rubro: solo "Lámpara" -> "Lámparas"; todo lo demás -> "Zona sin AP".
# Misma lógica, 1 a 1, que clasificar() en assets/js/procesar_ap.js.
#
# Cada regla: (tipo, motivo, condiciones_obs, actividad)
#   tipo "M"  -> Motivo EXACTO + alguna condición de Observaciones
#   tipo "MC" -> Motivo CONTIENE la frase (no mira Observaciones)
#   tipo "O"  -> solo condiciones de Observaciones
# Están en orden de prioridad: la PRIMERA que matchea gana.
REGLAS_V2 = [
    ("M", "LAMPARA APAGADA", "TODO NORMAL|AP NORMALIZADO|AP CONFORME|AP NORMAL|TODO NORMALIZADO|KTK|MALA DIRECCION|FALTAN DATOS", "Lámpara"),
    ("M", "ZONA SIN AP", "TODO NORMAL|AP NORMALIZADO|AP NORMAL|AP CONFORME|TODO NORMALIZADO|MALA DIRECCION|FALTAN DATOS", "Tablero AP"),
    ("MC", "SED AP FUERA DE SERVICIO", None, "SED AP FUERA DE SERVICIO"),
    ("O", None, "NA2XY|N2XY|NA3XY|SUB|CABLE EN CORTO|CABLE CORTO|C/V|C/ V|S/V|S/ V", "Red Subt."),
    ("M", "ZONA SIN AP", "TODO NORMAL|MALA DIRECCION|CONFORME|FALTAN DATOS|MAYOR REFERENCIA", "Tablero AP"),
    ("O", None, "SED|SECIONADOR|SECCIONADOR|COMUNICACION|FUSIBLE|CONTACTOR|FOTO|TAB|NH|TABLERO|CABLE DE SALIDA CKTO|CKTO|RELOJ|MEDIDOR|CONECTOR", "Tablero AP"),
    ("O", None, "CONECTOR|AUTO|AEREO|AEREA|ACOMETIDA|CORTA CIRCUITO|CNX|C/ACOMETIDA| CONEXION A PASANTE|CONECTORES", "Red Aérea"),
    ("O", None, "LUM|LAMP|LAM|REACT|IGNI|CONDEN|DIFUS|EQUI|CAMBIO X DEFECTO|F. CONTACTO|F.CONTACTO|FALSO CONTACTO|X CORROSION|KTK|DRIVER|PASANTE|CONECCION EN CORTOCIRCUITO|LED|PASTORAL", "Lámpara"),
    ("M", "LAMPARA APAGADA", "TODO NORMAL|CONFORME|FALTAN DATOS|MAYOR REFERENCIA|MALA DIRECCION", "Lámpara"),
    ("O", None, "SE RENOVO DEL POSTE", "Red Aérea"),
    ("M", "LAMPARA APAGADA", "OPERATIVO|OPERATIVA", "Lámpara"),
    ("M", "ZONA SIN AP", "OPERATIVO|OPERATIVA", "Tablero AP"),
    ("O", None, "SE NORMALIZO MONTAJE", "Red Aérea"),
    ("M", "LAMPARA APAGADA", "ENCONTRO NORMALIZADO", "Lámpara"),
    ("M", "ZONA SIN AP", "ENCONTRO NORMALIZADO", "Tablero AP"),
    ("O", None, "AP NORMAL|CORTACIRCUITO", "Lámpara"),
    ("O", None, "2X16|3X35|3X70", "Red Aérea"),
    ("O", None, "SE NORMALIZO SOBRECARGA DEL POSTE", "Red Aérea"),
    ("O", None, "CONCENTRICO", "Red Aérea"),
    ("M", "LAMPARA APAGADA", "SE NORMALIZO CON SAP", "Lámpara"),
    ("O", None, "SE ENDEREZO DEL POSTE", "Lámpara"),
    ("M", "ZONA SIN AP", "SE NORMALIZO CON SAP", "Tablero AP"),
    ("M", "ZONA SIN AP", "EMPAL", "Red Subt."),
    ("O", None, "CONTACTOR|FOTO|TAB|NH", "Lámpara"),
]

NO_ENCONTRADO = "No se encuentra"


def _cond_coincide(texto_norm, cond_norm):
    if cond_norm == "TAB":
        return contiene_palabra_completa(texto_norm, "TAB")
    if cond_norm == "SUB":
        return contiene_sub_valido(texto_norm)
    return cond_norm in texto_norm


def _alguna_cond(texto_norm, lista_condiciones):
    if not texto_norm:
        return False
    for cond in lista_condiciones.split("|"):
        cn = norm(cond)
        if cn and _cond_coincide(texto_norm, cn):
            return True
    return False


def clasificar(obs, motivo_reclamo):
    """Devuelve (actividad, rubro) segun el macro v2. Si nada matchea ->
    ("No se encuentra", "No se encuentra")."""
    o = norm(obs)
    m = norm(motivo_reclamo)
    for (tipo, motivo, cond, act) in REGLAS_V2:
        if tipo == "M":
            if m == norm(motivo) and _alguna_cond(o, cond):
                return act, rubro_de(act)
        elif tipo == "MC":
            if norm(motivo) in m:
                return act, rubro_de(act)
        elif tipo == "O":
            if _alguna_cond(o, cond):
                return act, rubro_de(act)
    return NO_ENCONTRADO, NO_ENCONTRADO


def rubro_de(actividad):
    return "Lámparas" if actividad == "Lámpara" else "Zona sin AP"


# Filtro ELIMINAR (2026-08-29, aplica a las 2 páginas): revisa TODAS las
# columnas del GAB en busca de frases que indican que el caso ya se manejó
# por otro lado — ver el mismo filtro en procesar_ap.js.
FRASES_ELIMINAR = [
    "SALESFORCE", "SALEFORCE", "SALESFOR", "SALES FORCE", "SALE FORCE", "SELFORCE", "SALFORCE",
    "SE ATIENDE CON ODM", "SE ATENDIO CON ODM", "ATENDIDO CON ODM",
    "ATENDIDO CON LA DENUNCIA", "ATENDIDO CON DENUNCIA",
    "TOMANORMAL RELACIONADO", "TOMA NORMAL RELACIONADO", "SE RELACIONA",
    "SE ATENDIO CON SF",
]
PALABRAS_COMPLETAS_ELIMINAR = ["RELACIONADO", "RELACIONADA", "RELACIONAR"]
_CODIGO_SF_RE = _re.compile(r"^SF\d+$")


def tiene_codigo_sf(texto):
    # 2026-09-05 (macro v2): además de "SF123..." también elimina si aparece
    # un token "SF" solo (así lo hace ContieneCodigoPrefijo del macro v2).
    return any(tok == "SF" or _CODIGO_SF_RE.match(tok) for tok in texto.split(" "))


def fila_debe_eliminarse(fila_dict):
    valores = fila_dict.values()
    texto_completo = " ".join(norm(v) for v in valores)
    if any(frase in texto_completo for frase in FRASES_ELIMINAR):
        return True
    if any(contiene_palabra_completa(texto_completo, p) for p in PALABRAS_COMPLETAS_ELIMINAR):
        return True
    if tiene_codigo_sf(texto_completo):
        return True
    # 2026-09-09: cualquier palabra con "FORCE" adentro = Salesforce mal
    # escrito ("LASAFORCE", etc.). Ninguna nota de campo en español lleva
    # "force" dentro de una palabra, así que es seguro.
    if "FORCE" in texto_completo:
        return True
    # (2026-09-01 a 2026-09-16) Antes, si "SAP - SAP Maestra" tenía dato, la
    # fila se eliminaba del todo -- por eso nunca aparecía ninguna
    # "relacionada" en Atendidas (había 7,410 casos así en el GAB crudo, cero
    # en la base publicada). Ahora se deja de eliminar: se marca con
    # "RELACIONADA" (ver procesar(), igual que en Pendientes) y se excluye
    # del Dashboard vía "_dashboardValido" -- así el Dashboard sigue dando
    # los mismos números de siempre, pero el Mapa/Buscar y exportar SÍ ven
    # estos casos y pueden filtrar por relacionada.
    return False


def archivo_mas_reciente(carpeta, exts=(".xlsx", ".xls", ".xlsm")):
    todos = todos_los_excel(carpeta, exts)
    return todos[0] if todos else None


def todos_los_excel(carpeta, exts=(".xlsx", ".xls", ".xlsm")):
    candidatos = []
    for ext in exts:
        candidatos.extend(glob.glob(os.path.join(carpeta, f"*{ext}")))
    candidatos = [c for c in candidatos if not os.path.basename(c).startswith("~$")]
    return sorted(candidatos, key=os.path.getmtime, reverse=True)


_RE_SLOT_MES = re.compile(r"[_.]?m(\d+)[_.]", re.I)


def _ordenar_por_slot(paths):
    """Con la ventana de 3 meses, CARGA\\EXCEL\\ trae R_m0/m1/m2. m0 = mes
    actual va PRIMERO para que, al deduplicar por Cod.SalesForce (primera
    aparición gana), quede la versión del mes más reciente. Sin slot (uso
    manual con 1 Excel), se ordena por fecha de archivo, más nuevo primero."""
    def clave(p):
        m = _RE_SLOT_MES.search(os.path.basename(p))
        return (0, int(m.group(1))) if m else (1, -os.path.getmtime(p))
    return sorted(paths, key=clave)


def leer_gab_muchos(paths):
    """Une varios GAB (uno por mes) en un solo DataFrame, en orden de slot
    (m0 primero). El deduplicado por Cod.SalesForce lo hace procesar()."""
    paths = _ordenar_por_slot(paths)
    dfs = [leer_gab(p) for p in paths]
    if len(dfs) == 1:
        return dfs[0]
    df = pd.concat(dfs, ignore_index=True)
    print(f"GAB: {len(df)} filas en {len(paths)} Excel(s) unidos (dedup por Cod.SalesForce en el proceso).")
    return df


def normalizar_encabezado(v):
    """Para comparar encabezados de Excel sin que importen espacios de
    sobra, mayúsculas/minúsculas ni acentos."""
    texto = "" if v is None else str(v)
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto).strip()


def leer_gab(path):
    # fila 1 = título del reporte, fila 2 = encabezados reales -> se busca
    # cada columna por su NOMBRE (no por posición), para que un Excel con
    # columnas recortadas o reordenadas se siga leyendo bien.
    df_raw = pd.read_excel(path, skiprows=1, header=None, dtype=object)
    header_real = df_raw.iloc[0].tolist() if len(df_raw) else []
    body = df_raw.iloc[1:].reset_index(drop=True)
    header_norm = [normalizar_encabezado(h) for h in header_real]
    data = {}
    for col in GAB_COLUMNS:
        objetivo = normalizar_encabezado(col)
        idx = header_norm.index(objetivo) if objetivo in header_norm else -1
        data[col] = body.iloc[:, idx] if idx != -1 else pd.Series([None] * len(body))
    return pd.DataFrame(data)


def iso_js(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S") + ".000Z"


def asignar_contratista_por_concurrencia(rows):
    """Reasigna por distrito de concurrencia los casos que quedaron
    'SIN ASIGNAR' (ver comentario de uso más abajo)."""
    concurrencia_por_distrito = {}
    for r in rows:
        c = r.get("Contratista")
        if not c or c == "SIN ASIGNAR":
            continue
        distrito = r.get("Poste - Distrito")
        if not distrito:
            continue
        m = concurrencia_por_distrito.setdefault(distrito, {})
        m[c] = m.get(c, 0) + 1

    def contratista_mas_frecuente(distrito):
        m = concurrencia_por_distrito.get(distrito)
        if not m:
            return None
        return max(m.items(), key=lambda kv: kv[1])[0]

    for r in rows:
        if r.get("Contratista") == "SIN ASIGNAR":
            asignado = contratista_mas_frecuente(r.get("Poste - Distrito"))
            if asignado:
                r["Contratista"] = asignado


ACTIVIDAD_ELIMINAR = "ELIMINAR"


def procesar(df, maestros, correcciones=None, mapa_deficiencia=None, mapa_zona=None):
    """correcciones: dict { str(nro_reclamo): (actividad, rubro) } con lo YA
    VALIDADO por Pluz en la pestaña "Correcciones" de Atendidas.
      - Actividad normal -> le gana al macro para siempre.
      - Actividad == "ELIMINAR" -> esa fila se saca del reporte publicado
        (no cuenta en ningún gráfico ni total).
    mapa_deficiencia: motivo normalizado -> "DT1".."DT6"/"NO DT" (ver
    cargar_mapa_deficiencia) -- para "Tipo DT"/"LEGAL" (Mapa/Buscar y exportar)."""
    if df.empty:
        raise ValueError("El archivo GAB no tiene filas de datos.")

    correcciones = correcciones or {}
    mapa_deficiencia = mapa_deficiencia or {}
    mapa_zona = mapa_zona or {}
    feriados_habil = {f["fecha"] for f in maestros["feriados"] if f.get("estado") == "Hábil"}
    mapa_tecnicos = {norm(t["tecnico"]): t["contratista"] for t in maestros["tecnicos"]}

    # 1) Estado = ATENDIDA
    f1 = df[df["SAP - Estado"].apply(norm) == "ATENDIDA"]
    # 2) (2026-09-16) YA NO se filtra por motivo acá -- el Dashboard sigue
    # acotado a MOTIVOS_VALIDOS (ver "_dashboardValido" más abajo), pero el
    # Mapa y "Buscar y exportar" necesitan CUALQUIER motivo atendido.
    f2 = f1
    # 3) Deduplicar por Cod.SalesForce, primera aparición
    vistos = set()
    filas3 = []
    for _, fila in f2.iterrows():
        k = clave_dedup(fila["SAP - Cod.SalesForce"])
        if k in vistos:
            continue
        vistos.add(k)
        filas3.append(fila)

    # 4) Filtro ELIMINAR (2026-08-29): se excluyen del reporte publicado las
    # filas que traigan Salesforce/relacionado/ODM en cualquiera de sus columnas.
    eliminadas_por_filtro = 0
    filas4 = []
    for fila in filas3:
        if fila_debe_eliminarse(fila.to_dict()):
            eliminadas_por_filtro += 1
            continue
        filas4.append(fila)

    rows_publicas = []
    rows_completas = []
    errores_fecha = 0

    eliminadas_manual = 0
    for fila in filas4:
        nro = fila["SAP - Número de reclamo"]
        _correc = correcciones.get(str(nro).strip()) if nro is not None else None
        if _correc and norm(_correc[0]) == ACTIVIDAD_ELIMINAR:
            eliminadas_manual += 1
            continue
        sucursal = fila["SAP - Sucursal"]
        motivo = fila["SAP - Motivo de reclamo"]
        distrito = fila["Poste - Distrito"]
        referencias = fila["Poste - Referencias"]
        obs = fila["SAP - Observaciones en Atención"]
        sap_maestra = fila["SAP - SAP Maestra"]
        motivo_falla = fila["SAP - Motivo falla técnica"]
        cuadrilla = fila["SAP MOVIL - Codigo Cuadrilla"]
        codsf = fila["SAP - Cod.SalesForce"]
        cod_sap = fila["SAP - Cod. SAP"]  # ODM (600...)
        freg_raw = fila["SAP - Fecha de registro"]
        fate_raw = fila["SAP - Fecha de Atención"]
        festim_raw = fila["SAP - Fecha Estimada"]
        zona_raw = fila["SAP - Zona"]
        estado = fila["SAP - Estado"]
        tecnico = fila["Técnico Ejecutor"]

        freg = parse_fecha(freg_raw)
        fate = parse_fecha(fate_raw)
        if freg is None or fate is None:
            errores_fecha += 1
            continue

        freg_med = solo_fecha(freg)
        fate_med = solo_fecha(fate)
        dh = dias_habiles(freg_med, fate_med, feriados_habil)
        rango = rango_de(dh)
        actividad, rubro = clasificar(obs, motivo)
        # Corrección manual validada (pestaña "Correcciones") — le gana al macro.
        correc = correcciones.get(str(nro).strip()) if nro is not None else None
        if correc:
            actividad, rubro = correc
        contratista = mapa_tecnicos.get(norm(tecnico), "SIN ASIGNAR")
        dias_cal = (fate_med - freg_med).days
        rango_cal = "1 DC" if dias_cal <= 1 else ("2 DC" if dias_cal <= 2 else "> 2 DC")
        dia_atencion_iso = fate_med.strftime("%Y-%m-%d")
        dia = fate_med.day
        nombre_mes = MESES_CORTOS[fate_med.month]
        eje_x = f"{dia} {nombre_mes}"
        tipo_dt = mapa_deficiencia.get(norm(motivo), "NO DT")
        legal = "SI" if str(tipo_dt).upper().startswith("DT") else "NO"
        relacionada = "SI" if norm(sap_maestra) != "" else "NO"
        dashboard_valido = norm(motivo) in MOTIVOS_VALIDOS and relacionada == "NO"
        # PLAZO (DP/FDP): a diferencia de Pendientes (compara la fecha límite
        # contra HOY, porque el caso sigue abierto), acá el caso ya se
        # atendió -- se compara la fecha límite ("SAP - Fecha Estimada")
        # contra la fecha real en que se atendió. Si el Excel no trae esa
        # columna (viene vacía en algunos reducidos), queda "SIN DATO" en vez
        # de reventar o inventar un valor.
        festim = parse_fecha(festim_raw)
        plazo = ("DP" if festim >= fate else "FDP") if festim is not None else "SIN DATO"
        uo = mapa_zona.get(norm(zona_raw), None)

        base = {
            "SAP - Número de reclamo": nro,
            "SAP - Cod. SAP": cod_sap,
            "SAP - Sucursal": sucursal,
            "SAP - Motivo de reclamo": motivo,
            "Poste - Distrito": distrito,
            "SAP - Observaciones en Atención": obs,
            "SAP - SAP Maestra": sap_maestra,
            "SAP - Motivo falla técnica": motivo_falla,
            "SAP MOVIL - Codigo Cuadrilla": cuadrilla,
            "SAP - Cod.SalesForce": codsf,
            "SAP - Fecha de registro": iso_js(freg),
            "SAP - Fecha de Atención": iso_js(fate),
            "SAP - Estado": estado,
            "Técnico Ejecutor": tecnico,
            "Días hábiles": dh,
            "Rango": rango,
            "Actividad": actividad,
            "Rubro": rubro,
            "Contratista": contratista,
            "Dias calendario": dias_cal,
            "Rango calendario": rango_cal,
            "Dia Atención": dia_atencion_iso,
            "Día": dia,
            "Nombre del mes": nombre_mes,
            "EJE X": eje_x,
            "Tipo DT": tipo_dt,
            "LEGAL": legal,
            "PLAZO": plazo,
            "UO": uo,
            "RELACIONADA": relacionada,
            "_dashboardValido": dashboard_valido,
        }
        rows_publicas.append(base)
        rows_completas.append({"Poste - Referencias": referencias, **base})

    # Casos "SIN ASIGNAR" (muy pocos, la excepción — técnico que no está en
    # MAESTROS/tecnicos.json): se les asigna el contratista que más atiende
    # ese mismo distrito, según los propios datos de esta corrida
    # (concurrencia). rows_publicas y rows_completas están en el mismo orden
    # (un append de cada uno por fila), así que se replica el mismo
    # resultado en ambos por índice.
    asignar_contratista_por_concurrencia(rows_publicas)
    for i, r in enumerate(rows_publicas):
        rows_completas[i]["Contratista"] = r["Contratista"]

    if errores_fecha > 0:
        print(f"AVISO: {errores_fecha} filas se descartaron por fechas inválidas.")

    stats = {
        "total": len(df), "tras_estado": len(f1), "tras_motivo": len(f2),
        "tras_dedup": len(filas3), "eliminadas_por_filtro": eliminadas_por_filtro,
        "tras_filtro_eliminar": len(filas4), "eliminadas_manual": eliminadas_manual,
        "final": len(rows_publicas), "errores_fecha": errores_fecha,
    }
    return rows_publicas, rows_completas, datetime.utcnow(), stats


def limpia(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    return v


def guardar_excel_completo(rows_completas, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    df = pd.DataFrame(rows_completas, columns=BD_COLUMNS_COMPLETAS)
    nombre = f"bd_completa_{ahora.strftime('%Y-%m-%d_%H%M')}.xlsx"
    ruta = os.path.join(SALIDA_DIR, nombre)
    df.to_excel(ruta, index=False)
    return ruta


def guardar_json_publicar(rows_publicas, ahora):
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    filas = [{k: limpia(v) for k, v in r.items()} for r in rows_publicas]
    payload = {"updated_at": updated_at, "rows": filas}
    with open(PENDIENTE_PUBLICAR, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return PENDIENTE_PUBLICAR, len(filas)


def guardar_json_completo(rows_completas, ahora):
    """BD completa (con 'Poste - Referencias', datos del cliente) en JSON,
    para que feed.py la publique — es lo que alimenta 'Buscar y exportar'
    para todos los que entren al sitio, sin que tengan que reprocesar nada."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    filas = [{k: limpia(v) for k, v in r.items()} for r in rows_completas]
    payload = {"updated_at": updated_at, "rows": filas}
    ruta = os.path.join(SALIDA_DIR, "bd_completa.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    return ruta, len(filas)


def _filas_del_contratista(rows, contratista):
    return [r for r in rows if id_login_de_contratista(r.get("Contratista")) == contratista]


def guardar_json_publicar_contratistas(rows_publicas, ahora):
    """Un bd_actual_<contratista>.json por cada contratista con clave propia
    — SOLO sus filas, ninguna de los demás ni de 'sin asignar'."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    resultados = []
    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_c = [{k: limpia(v) for k, v in r.items()} for r in _filas_del_contratista(rows_publicas, contratista)]
        payload = {"updated_at": updated_at, "rows": filas_c}
        ruta = os.path.join(SALIDA_DIR, f"bd_actual_{contratista.lower()}.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        resultados.append((contratista, ruta, len(filas_c)))
    return resultados


def guardar_json_completo_contratistas(rows_completas, ahora):
    """Un bd_completa_<contratista>.json por cada contratista con clave
    propia — es lo que alimenta 'Buscar y exportar' para ESE contratista,
    con SOLO sus propios datos de cliente."""
    os.makedirs(SALIDA_DIR, exist_ok=True)
    updated_at = iso_js(ahora)
    resultados = []
    for contratista in CONTRATISTAS_CON_CLAVE:
        filas_c = [{k: limpia(v) for k, v in r.items()} for r in _filas_del_contratista(rows_completas, contratista)]
        payload = {"updated_at": updated_at, "rows": filas_c}
        ruta = os.path.join(SALIDA_DIR, f"bd_completa_{contratista.lower()}.json")
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        resultados.append((contratista, ruta, len(filas_c)))
    return resultados


def cargar_correcciones_validadas(script_url):
    """Baja de la Apps Script (Google Sheet "Correcciones AP") las correcciones
    manuales de Actividad/Rubro YA VALIDADAS por Pluz. Devuelve
    { str(reclamo): (actividad, rubro) }.

    Si NO se puede leer (sin internet, Apps Script caída, etc.) LANZA una
    excepción a propósito — mejor abortar la corrida que publicar los datos
    sin las correcciones (haría que casos ya arreglados vuelvan a
    "No se encuentra" sin que nadie se dé cuenta)."""
    import urllib.request
    url = script_url.rstrip("/") + "?path=correcciones&t=" + str(int(datetime.utcnow().timestamp()))
    with urllib.request.urlopen(url, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if not isinstance(data, dict) or "rows" not in data:
        raise RuntimeError(f"Respuesta inesperada al leer correcciones: {str(data)[:200]}")
    out = {}
    for r in data["rows"]:
        if str(r.get("validado", "")).strip().upper() != "SI":
            continue
        rec = str(r.get("reclamo", "")).strip()
        act = (r.get("actividad") or "").strip()
        rub = (r.get("rubro") or "").strip()
        if rec and act:
            out[rec] = (act, rub or rubro_de(act))
    return out


def _reclamos_con_correccion(script_url):
    """Set de TODOS los reclamos que tienen alguna fila en la hoja
    Correcciones (validada o no). Para el índice de la pestaña Correcciones."""
    import urllib.request
    url = script_url.rstrip("/") + "?path=correcciones&t=" + str(int(datetime.utcnow().timestamp()))
    with urllib.request.urlopen(url, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return {str(r.get("reclamo", "")).strip() for r in data.get("rows", []) if str(r.get("reclamo", "")).strip()}


def generar_indice_correcciones(script_url):
    """Escribe SALIDA/correcciones_indice.json: por cada mes (actual +
    HISTORICO), SOLO los casos 'No se encuentra' + los que ya tienen alguna
    corrección. La pestaña Correcciones del sitio lee ESTE archivo en vez de
    bajar todos los meses completos (que eran miles de filas)."""
    NO_ENC = "No se encuentra"
    try:
        con_correc = _reclamos_con_correccion(script_url)
    except Exception as e:
        print(f"  (aviso: no se pudo leer la hoja de correcciones para el índice: {e})")
        con_correc = set()

    casos = []

    def _extraer(rows, mes):
        for row in rows:
            # (2026-09-16) bd_completa.json ahora trae TODOS los motivos (para
            # Mapa/Buscar y exportar) -- pero Actividad/Rubro (REGLAS_V2) solo
            # sabe clasificar los 3 motivos de siempre, así que cualquier otro
            # motivo cae siempre en "No se encuentra" e inundaba Correcciones
            # con casos que ni siquiera le importan al Dashboard. Correcciones
            # sigue siendo solo para lo que SÍ es "_dashboardValido".
            if not row.get("_dashboardValido", True):
                continue
            act = str(row.get("Actividad") or "").strip()
            rec = str(row.get("SAP - Número de reclamo") or "").strip()
            if act == NO_ENC or (rec and rec in con_correc):
                casos.append({
                    "__mes": mes,
                    "SAP - Número de reclamo": rec,
                    "SAP - Cod. SAP": row.get("SAP - Cod. SAP") or "",
                    "SAP - Motivo de reclamo": row.get("SAP - Motivo de reclamo") or "",
                    "Contratista": row.get("Contratista") or "",
                    "SAP - Observaciones en Atención": row.get("SAP - Observaciones en Atención") or "",
                    "Actividad": act,
                    "Rubro": row.get("Rubro") or "",
                })

    try:
        with open(os.path.join(SALIDA_DIR, "bd_completa.json"), encoding="utf-8") as f:
            _extraer(json.load(f).get("rows", []), clave_mes_actual_lima())
    except Exception:
        pass
    if os.path.isdir(HISTORICO_DIR):
        for nombre in sorted(os.listdir(HISTORICO_DIR)):
            if not re.match(r"^\d{4}-\d{2}_completa\.json$", nombre):
                continue
            try:
                with open(os.path.join(HISTORICO_DIR, nombre), encoding="utf-8") as f:
                    _extraer(json.load(f).get("rows", []), nombre[:7])
            except Exception:
                pass

    ruta = os.path.join(SALIDA_DIR, "correcciones_indice.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump({"generado": datetime.utcnow().isoformat() + "Z", "casos": casos}, f, ensure_ascii=False)
    print(f"Índice de Correcciones: {len(casos)} caso(s) -> {ruta}")


def main():
    print("=== Reporte de Atendidas AP — procesamiento local ===")
    gab_paths = todos_los_excel(GAB_DIR)
    if not gab_paths:
        print(f"No se encontró ningún Excel en ..\\CARGA\\EXCEL\\ ({GAB_DIR}).")
        sys.exit(1)

    faltan_maestros = [n for n in ("tecnicos.json", "feriados.json")
                        if not os.path.exists(os.path.join(MAESTROS_DIR, n))]
    if faltan_maestros:
        print(f"Faltan archivos en MAESTROS\\: {', '.join(faltan_maestros)}")
        sys.exit(1)

    # Correcciones manuales validadas (Google Sheet). Si falla, se aborta.
    try:
        with open(os.path.join(BASE_DIR, "config.json"), encoding="utf-8") as f:
            script_url = json.load(f)["script_url"]
        correcciones = cargar_correcciones_validadas(script_url)
        print(f"Correcciones manuales validadas: {len(correcciones)}")
    except Exception as e:
        print(f"NO se pudieron leer las correcciones manuales: {e}")
        print("Se aborta la corrida (no se publica nada) para no perder correcciones ya hechas.")
        sys.exit(1)

    print(f"GAB ({len(gab_paths)}): {', '.join(os.path.basename(p) for p in gab_paths)}")
    maestros = {}
    for clave, nombre in (("tecnicos", "tecnicos.json"), ("feriados", "feriados.json")):
        with open(os.path.join(MAESTROS_DIR, nombre), encoding="utf-8") as f:
            maestros[clave] = json.load(f)

    mapa_deficiencia = cargar_mapa_deficiencia()
    mapa_zona = cargar_mapa_zona()
    print(f"Tabla motivo->Tipo DT (de Pendientes): {len(mapa_deficiencia)} motivo(s)" if mapa_deficiencia
          else "AVISO: no se pudo leer data_requisito.json -- Tipo DT/LEGAL van a salir 'NO DT'/'NO' para todo.")

    # (2026-09-22) "Última actualización" que se ve en la página = hora del
    # Excel del GAP mas reciente (el mismo criterio que Pendientes/Veredas,
    # para que las 3 paginas muestren la MISMA hora) -- no la hora en que
    # termino de correr todo el proceso. "ahora" sigue siendo la hora real
    # para todo lo que es CALCULO (duracion, archivado de meses, etc.).
    ultima_actualizacion = datetime.utcfromtimestamp(os.path.getmtime(gab_paths[0]))

    df_gab = leer_gab_muchos(gab_paths)
    rows_publicas, rows_completas, ahora, stats = procesar(df_gab, maestros, correcciones, mapa_deficiencia, mapa_zona)
    print(f"Filas: total={stats['total']} -> ATENDIDA={stats['tras_estado']} -> "
          f"motivo válido={stats['tras_motivo']} -> únicas={stats['tras_dedup']} -> final={stats['final']}")

    ruta_excel = guardar_excel_completo(rows_completas, ahora)
    print(f"Excel completo (con Poste - Referencias, solo local): {ruta_excel}")

    # 2026-09-02: separa lo del mes calendario actual (va al dashboard/"Buscar
    # y exportar" en vivo, como siempre) de lo de meses ya pasados que todavía
    # traiga el GAB (el robot exporta ~30 días, así que a inicios de mes
    # siempre trae la cola del mes anterior) — esos meses pasados se archivan
    # solos en HISTORICO\ (ver historico_lib.archivar_mes), en vez de quedar
    # mezclados en la BD "en vivo" para siempre perderse cuando se salgan de
    # la ventana de 30 días. Así, con solo correr Ejecutar.bat día a día, cada
    # mes se completa y archiva automáticamente apenas cierra.
    clave_actual = clave_mes_actual_lima()
    actual_publicas, actual_completas = [], []
    pasado_por_mes = {}  # "YYYY-MM" -> [filas completas]
    for pub, comp in zip(rows_publicas, rows_completas):
        dia = pub.get("Dia Atención")
        clave = str(dia)[:7] if dia else None
        if not clave or clave == clave_actual:
            # (2026-09-16) bd_actual*.json (Dashboard) sigue siendo SOLO los 3
            # motivos de siempre -- bd_completa*.json (Mapa/Buscar y exportar)
            # ya lleva TODOS los motivos, filtrados más abajo (ver
            # "_dashboardValido" en procesar()).
            if pub.get("_dashboardValido", True):
                actual_publicas.append(pub)
            actual_completas.append(comp)
        else:
            pasado_por_mes.setdefault(clave, []).append(comp)

    if pasado_por_mes:
        print(f"Archivando en HISTORICO\\ {len(pasado_por_mes)} mes(es) que ya cerraron "
              f"({', '.join(sorted(pasado_por_mes))})...")
        ahora_iso = historico_lib.iso_js(ahora)
        for clave in sorted(pasado_por_mes):
            res = historico_lib.archivar_mes(HISTORICO_DIR, clave, pasado_por_mes[clave], id_login_de_contratista, ahora_iso)
            print(f"  {clave} ({res['label']}): {res['publicas']} filas archivadas (ya mezcladas con lo que hubiera antes)")

    ruta_json, n = guardar_json_publicar(actual_publicas, ultima_actualizacion)
    print(f"Listo para publicar - dashboard, mes actual ({n} filas): {ruta_json}")

    ruta_completa, n_completa = guardar_json_completo(actual_completas, ultima_actualizacion)
    print(f"Listo para publicar - BD completa/Buscar y exportar, mes actual ({n_completa} filas): {ruta_completa}")

    print("Generando archivos recortados por contratista (aislamiento de datos)...")
    for contratista, ruta, n in guardar_json_publicar_contratistas(actual_publicas, ultima_actualizacion):
        print(f"  {contratista}: dashboard ({n} filas): {ruta}")
    for contratista, ruta, n in guardar_json_completo_contratistas(actual_completas, ultima_actualizacion):
        print(f"  {contratista}: BD completa ({n} filas): {ruta}")

    generar_indice_correcciones(script_url)

    print("Ejecuta feed.py (o Ejecutar.bat, que ya lo hace) para subir todo a GitHub.")


if __name__ == "__main__":
    main()

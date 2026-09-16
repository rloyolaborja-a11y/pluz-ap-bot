# -*- coding: utf-8 -*-
"""
feed.py — Publica SALIDA\\bd_actual.json (dashboard) y SALIDA\\bd_completa.json
(BD completa, con datos de cliente, para "Buscar y exportar"), más las
versiones recortadas por contratista, en tu repositorio de GitHub, a través
de la Apps Script de Google (el "puente" hacia GitHub) — tu computadora
nunca le habla a github.com directamente, le habla a Google, y es Google
quien empuja el cambio.

Esto evita que un bloqueo de red/firewall/antivirus corporativo hacia
GitHub afecte esta publicación, porque Google casi nunca está bloqueado.

Lee la configuración desde config.json (ese archivo NO se sube a ningún
lado — solo vive en tu computadora, y ya no guarda ningún token de
GitHub, solo la URL de la Apps Script y una clave/SECRET).

Solo necesita tener Python instalado (desde python.org) — NO requiere
instalar ninguna librería adicional (no hace falta "pip install" de nada),
usa únicamente lo que ya viene incluido con Python.

IMPORTANTE (2026-08-27): todos los archivos de esta publicación (dashboard,
BD completa y el par de cada contratista — hasta 10 archivos) se suben
como UN SOLO commit (se le pide todo junto a la Apps Script, que arma un
único commit con la Git Data API de GitHub) — antes se mandaba un commit
POR ARCHIVO, y como cada commit dispara un despliegue nuevo en Vercel, eso
agotaba rápido el límite diario de despliegues de Vercel. Con un solo
commit por publicación, cada corrida de Ejecutar.bat cuenta como UN solo
despliegue.
"""

import base64
import json
import os
import ssl
import sys
import time
import datetime
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _contexto_ssl():
    """Arma un contexto HTTPS que además de los certificados que trae Python
    (certifi), suma los certificados "de confianza" que Windows ya tiene
    instalados en la PC. Esto es a propósito (2026-08-29): en redes de
    empresa con antivirus/firewall que inspecciona HTTPS (un "proxy MITM"),
    Windows y el navegador SÍ confían en el certificado que pone la empresa,
    pero Python (con su propio set de certificados, separado del de Windows)
    no lo conoce y la conexión fallaba con un error de certificado — por eso
    "Procesar y actualizar" desde la página (que usa el navegador) funcionaba
    pero Ejecutar.bat/Publicar.bat (que usan Python) no."""
    ctx = ssl.create_default_context()
    if sys.platform == "win32":
        try:
            for tienda in ("CA", "ROOT"):
                for cert_der, encoding, trust in ssl.enum_certificates(tienda):
                    if not trust:
                        continue
                    try:
                        ctx.load_verify_locations(cadata=ssl.DER_cert_to_PEM_cert(cert_der))
                    except ssl.SSLError:
                        pass
        except Exception:
            pass
    return ctx


# 2026-09-16: BUG encontrado -- Google Apps Script (script.google.com) casi
# siempre responde a /exec con un 302 hacia una URL de ejecución real
# (script.googleusercontent.com). El manejador de redirecciones POR DEFECTO
# de urllib (HTTPRedirectHandler) sigue ese 302 pero, para un POST, reenvía
# la redirección como GET (así lo hace también cualquier navegador desde
# hace décadas) -- se pierde el body entero (accion:"datos" + los archivos).
# El síntoma es justo el que apareció en un log real: un intento de los 3
# reintentos falló con "Falta el parámetro 'path' (ej: ?path=...)" -- ese es
# el mensaje de error de doGet(), no de doPost(): la Apps Script SÍ recibió
# la llamada, pero como GET sin body, así que cayó en el otro handler. Se
# arma un opener con un manejador de redirección propio que, si el pedido
# original era POST, vuelve a mandarlo como POST (mismo body) hacia la nueva
# URL, en vez de dejar que se degrade a GET.
# 2026-09-16: algunos proxies/antivirus de red corporativos distinguen
# tráfico "de script" (el User-Agent por defecto de Python, "Python-urllib/
# 3.x") del de un navegador real y lo tratan distinto (a veces bloqueándolo)
# -- probamos con un User-Agent de navegador común por si acaso.
_USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
               "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


class _ConservarPostAlRedirigir(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if req.get_method() == "POST" and code in (301, 302, 303, 307, 308):
            return urllib.request.Request(
                newurl, data=req.data, headers=req.headers, method="POST"
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


_SSL_CONTEXT = _contexto_ssl()
_OPENER = urllib.request.build_opener(
    _ConservarPostAlRedirigir, urllib.request.HTTPSHandler(context=_SSL_CONTEXT)
)
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
SALIDA_DIR = os.path.join(BASE_DIR, "SALIDA")
DATA_PATH = os.path.join(SALIDA_DIR, "bd_actual.json")
DATA_COMPLETA_PATH = os.path.join(SALIDA_DIR, "bd_completa.json")

# Contratistas con clave propia (aislamiento de datos) — igual que en
# procesar_diario.py y assets/js/app.js.
CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]


def cargar_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"Falta config.json ({CONFIG_PATH}). Copia config.ejemplo.json y complétalo.")
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    faltan = [k for k in ("script_url", "secret", "path") if not cfg.get(k)]
    if faltan:
        print(f"Faltan datos en config.json: {', '.join(faltan)}")
        sys.exit(1)
    if not cfg.get("path_completa"):
        cfg["path_completa"] = cfg["path"].replace("bd_actual.json", "bd_completa.json")
    return cfg


# ---------------------------------------------------------------------------
# 2026-09-16: publicación de "datos" DIRECTO a GitHub (repo aparte, sin
# Vercel conectado: rloyolaborja-a11y/pluz-ap-datos) en vez de Google Drive
# vía Apps Script. Motivo: un antivirus/DLP corporativo de esta PC bloquea
# las SUBIDAS (POST) hacia *.googleusercontent.com (adonde redirige la Apps
# Script al ejecutar) -- confirmado con pruebas reales, ver conversación del
# 2026-09-16. Mismo mecanismo que ya usa la Apps Script para "sitio" (Git
# Data API: blob por archivo + tree + commit + mover la rama), replicado acá
# en Python. Se intenta GitHub primero; si falla (ej. en OTRA PC donde fuera
# GitHub el bloqueado, no Google), cae al camino viejo de Apps Script como
# respaldo -- ver publicar_lote() más abajo.
def _gh_fetch(url, token, method="GET", payload=None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", "Bearer " + token)
    req.add_header("Accept", "application/vnd.github+json")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with _OPENER.open(req, timeout=60) as resp:
            cuerpo = resp.read()
            return resp.status, (json.loads(cuerpo) if cuerpo else {})
    except urllib.error.HTTPError as e:
        cuerpo = e.read()
        try:
            return e.code, json.loads(cuerpo)
        except ValueError:
            return e.code, {"message": cuerpo[:300].decode("utf-8", "replace")}


def _publicar_commit_github(cfg, archivos, mensaje):
    """archivos: lista de (path_remoto, contenido_bytes). Un solo commit,
    igual que publicarUnSoloCommit() en Code.gs. Maneja el caso del repo
    vacío (sin ningún commit todavía): ahí no hay rama que leer, así que el
    primer commit se arma sin padre ni base_tree."""
    token = cfg.get("github_token")
    owner = cfg.get("github_datos_owner")
    repo = cfg.get("github_datos_repo")
    branch = cfg.get("github_datos_branch", "main")
    if not token or not owner or not repo:
        raise RuntimeError("Falta github_token/github_datos_owner/github_datos_repo en config.json.")

    base = f"https://api.github.com/repos/{owner}/{repo}"

    codigo, ref = _gh_fetch(f"{base}/git/refs/heads/{branch}", token)
    if codigo == 409:
        # Repo 100% vacío (sin ningún commit): la Git Data API entera (ni
        # siquiera crear un blob) no funciona hasta que exista un primer
        # commit -- lo único que sirve para crear ESE primer commit en un
        # repo vacío es la Contents API (PUT contents/{path}). Se usa el
        # primer archivo de esta publicación para ese arranque; después la
        # rama ya existe y se sigue con el flujo normal de abajo para
        # (re)subir TODOS los archivos (incluido ese primero) de una vez.
        primer_path, primer_contenido = archivos[0]
        codigo_boot, resultado_boot = _gh_fetch(
            f"{base}/contents/{primer_path}", token, "PUT", {
                "message": mensaje + " (primer commit, arranque del repo)",
                "content": base64.b64encode(primer_contenido).decode("ascii"),
                "branch": branch,
            })
        if codigo_boot not in (200, 201):
            raise RuntimeError(f"GitHub contents (arranque repo vacío) -> {codigo_boot}: {resultado_boot}")
        codigo, ref = _gh_fetch(f"{base}/git/refs/heads/{branch}", token)

    commit_actual_sha = None
    tree_base_sha = None
    if codigo == 200:
        commit_actual_sha = ref["object"]["sha"]
        codigo2, commit_actual = _gh_fetch(f"{base}/git/commits/{commit_actual_sha}", token)
        if codigo2 != 200:
            raise RuntimeError(f"GitHub git/commits -> {codigo2}: {commit_actual}")
        tree_base_sha = commit_actual["tree"]["sha"]
    elif codigo != 404:
        raise RuntimeError(f"GitHub git/refs -> {codigo}: {ref}")
    # codigo == 404 -> la rama todavía no existe (caso raro, ya cubierto el
    # de repo vacío arriba) -> primer commit, sin padre/base_tree.

    entradas_arbol = []
    for path_remoto, contenido in archivos:
        codigo_b, blob = _gh_fetch(f"{base}/git/blobs", token, "POST", {
            "content": base64.b64encode(contenido).decode("ascii"), "encoding": "base64",
        })
        if codigo_b not in (200, 201):
            raise RuntimeError(f"GitHub git/blobs ({path_remoto}) -> {codigo_b}: {blob}")
        entradas_arbol.append({"path": path_remoto, "mode": "100644", "type": "blob", "sha": blob["sha"]})

    payload_tree = {"tree": entradas_arbol}
    if tree_base_sha:
        payload_tree["base_tree"] = tree_base_sha
    codigo_t, tree_nuevo = _gh_fetch(f"{base}/git/trees", token, "POST", payload_tree)
    if codigo_t not in (200, 201):
        raise RuntimeError(f"GitHub git/trees -> {codigo_t}: {tree_nuevo}")

    payload_commit = {"message": mensaje, "tree": tree_nuevo["sha"]}
    if commit_actual_sha:
        payload_commit["parents"] = [commit_actual_sha]
    codigo_c, commit_nuevo = _gh_fetch(f"{base}/git/commits", token, "POST", payload_commit)
    if codigo_c not in (200, 201):
        raise RuntimeError(f"GitHub git/commits -> {codigo_c}: {commit_nuevo}")

    if commit_actual_sha:
        codigo_r, resultado_ref = _gh_fetch(f"{base}/git/refs/heads/{branch}", token, "PATCH", {
            "sha": commit_nuevo["sha"], "force": False,
        })
    else:
        codigo_r, resultado_ref = _gh_fetch(f"{base}/git/refs", token, "POST", {
            "ref": f"refs/heads/{branch}", "sha": commit_nuevo["sha"],
        })
    if codigo_r not in (200, 201):
        raise RuntimeError(f"GitHub mover rama -> {codigo_r}: {resultado_ref}")

    return {"ok": True, "commit": commit_nuevo["sha"]}


# ---------------------------------------------------------------------------
# 2026-09-04: publicación por "slots" alternados (a/b) -- BUG real encontrado
# en producción: bd_actual.json y bd_completa.json (y sus 8 variantes por
# contratista, 10 archivos en total) se pisaban en el MISMO lugar en cada
# publicación, y como se suben uno por uno (varios minutos en total), quien
# tuviera el reporte abierto justo en esa ventana podía leer una MEZCLA de
# archivos viejos y nuevos (ej. el gráfico del Dashboard con la BD ya
# actualizada, pero "Buscar y exportar" todavía con la BD completa vieja).
#
# Ahora cada publicación completa se sube a un slot que NADIE está leyendo
# ("a" o "b", alternando) y, recién cuando TODOS esos archivos ya están
# arriba, se mueve un punterito (data/latest.json) al slot nuevo -- ese
# punterito es un archivo chiquito, así que "pisarlo" es casi instantáneo,
# sin la ventana de varios minutos que sí tenían los archivos grandes. El
# sitio primero lee el puntero y con eso sabe qué slot leer (ver
# cargarPunteroSlot en assets/js/app.js) -- nunca puede agarrar una mezcla,
# porque el slot que NO está señalado ahora mismo no se toca hasta la
# SIGUIENTE publicación (dos corridas más adelante).
# ---------------------------------------------------------------------------

def _puntero_path(cfg):
    base = cfg["path"].rsplit("/", 1)[0]  # ej. "pendientes/data"
    return f"{base}/latest.json"


def _con_slot(path_remoto, slot):
    base, nombre = path_remoto.rsplit("/", 1)
    return f"{base}/{slot}/{nombre}"


def _leer_puntero(cfg):
    """Lee data/latest.json para saber qué slot está vigente ahora mismo.
    Primero intenta el repo de datos en GitHub (raw.githubusercontent.com,
    lectura pública, sin token) y si no hay nada ahí (repo recién creado,
    o publicaciones viejas que quedaron en Drive) cae al doGet de la Apps
    Script de Google. Si no existe todavía o algo falla en los dos, se
    asume slot "b" -- así esta publicación arranca escribiendo en "a"."""
    owner = cfg.get("github_datos_owner")
    repo = cfg.get("github_datos_repo")
    branch = cfg.get("github_datos_branch", "main")
    if owner and repo:
        url_raw = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{_puntero_path(cfg)}"
        try:
            with urllib.request.urlopen(url_raw, timeout=15, context=_SSL_CONTEXT) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            if data.get("slot") in ("a", "b"):
                return data["slot"]
        except Exception:
            pass

    import urllib.parse
    url = cfg["script_url"] + "?" + urllib.parse.urlencode({"path": _puntero_path(cfg)})
    try:
        with urllib.request.urlopen(url, timeout=30, context=_SSL_CONTEXT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("slot") in ("a", "b"):
            return data["slot"]
    except Exception:
        pass
    return "b"


REINTENTOS = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4


def _publicar_lote_intento(cfg, archivos, mensaje):
    """archivos: lista de (path_remoto, contenido_bytes)."""
    files_payload = [
        {"path": path_remoto, "content_base64": base64.b64encode(contenido).decode("ascii")}
        for path_remoto, contenido in archivos
    ]
    # accion:"datos" -> se guarda en Drive vía la Apps Script, SIN pasar por
    # GitHub, así que esta publicación de datos ya NO consume el cupo de
    # despliegues de Vercel (2026-08-29).
    body = {"secret": cfg["secret"], "message": mensaje, "files": files_payload, "accion": "datos"}
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(cfg["script_url"], data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("User-Agent", _USER_AGENT)
    try:
        with _OPENER.open(req, timeout=120) as resp:
            resultado = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"La Apps Script respondió {e.code}: {e.read()[:400]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"No se pudo conectar con la Apps Script de Google: {e.reason}")

    if not resultado.get("ok"):
        raise RuntimeError(resultado.get("error", "La Apps Script rechazó la publicación."))
    return resultado


# Tamaño máximo aproximado (bytes, ya en base64) por llamada a la Apps
# Script. Descubrimos (2026-08-29, probando con datos reales) que mandar
# TODOS los archivos juntos en un solo POST (dashboard + BD completa + el
# par de cada contratista, que puede pasar de 30-50MB) hace que la llamada
# se cuelgue y termine en "timed out" — como esto va a Drive (accion:
# "datos") y ya NO dispara ningún despliegue en Vercel, no hay ninguna
# desventaja en partirlo en varias llamadas más chicas.
TAMANO_MAX_POR_LLAMADA = 3 * 1024 * 1024  # ~3MB de contenido (antes de base64) por llamada


def _agrupar_por_tamano(archivos):
    grupos = []
    grupo_actual = []
    tamano_actual = 0
    for path_remoto, contenido in archivos:
        tamano = len(contenido)
        if tamano > TAMANO_MAX_POR_LLAMADA:
            if grupo_actual:
                grupos.append(grupo_actual)
                grupo_actual, tamano_actual = [], 0
            grupos.append([(path_remoto, contenido)])
            continue
        if grupo_actual and tamano_actual + tamano > TAMANO_MAX_POR_LLAMADA:
            grupos.append(grupo_actual)
            grupo_actual, tamano_actual = [], 0
        grupo_actual.append((path_remoto, contenido))
        tamano_actual += tamano
    if grupo_actual:
        grupos.append(grupo_actual)
    return grupos


# 2026-09-05: los envíos se mandaban uno por uno, esperando a que cada uno
# termine antes de mandar el siguiente -- con ~9 envíos (el caso normal de
# Pendientes), eso es lo que hacía que "Pendientes: publicar" solita tardara
# 3-4 minutos. Cada envío escribe en un archivo DISTINTO (no hay nada
# compartido entre ellos), así que no hay ningún riesgo en mandarlos en
# paralelo -- el puntero (ver publicar_lote_atomico más abajo, en main())
# igual se sigue moviendo recién cuando ESTA función entera ya terminó
# (todos los envíos, no solo el primero), así que la garantía de
# atomicidad del slot no cambia en nada.
MAX_ENVIOS_EN_PARALELO = 4


def _subir_grupo_con_reintentos(cfg, grupo, mensaje, i, total):
    ultimo_error = None
    for intento in range(1, REINTENTOS + 1):
        try:
            return _publicar_lote_intento(cfg, grupo, mensaje)
        except Exception as e:
            ultimo_error = e
            etiqueta = f"envío {i}/{total}, " if total > 1 else ""
            print(f"  {etiqueta}intento {intento}/{REINTENTOS} falló: {e}")
            if intento < REINTENTOS:
                time.sleep(ESPERA_ENTRE_REINTENTOS_SEG)
    raise ultimo_error


def _publicar_lote_apps_script(cfg, archivos, mensaje):
    grupos = _agrupar_por_tamano(archivos)
    if len(grupos) > 1:
        print(f"  ({len(grupos)} envíos, hasta {MAX_ENVIOS_EN_PARALELO} en paralelo)")
    resultado_final = None
    with ThreadPoolExecutor(max_workers=MAX_ENVIOS_EN_PARALELO) as pool:
        futuros = [
            pool.submit(_subir_grupo_con_reintentos, cfg, grupo, mensaje, i, len(grupos))
            for i, grupo in enumerate(grupos, start=1)
        ]
        for futuro in as_completed(futuros):
            resultado_final = futuro.result()  # relanza la excepción si ese envío falló del todo
    return resultado_final


def publicar_lote(cfg, archivos, mensaje):
    """Intenta GitHub directo primero (ver _publicar_commit_github) y, si
    falla, cae al camino viejo de Apps Script/Drive -- por si algún día es
    al revés (GitHub bloqueado, Google no) en otra PC/red."""
    try:
        return _publicar_commit_github(cfg, archivos, mensaje)
    except Exception as e_github:
        print(f"  GitHub directo falló ({e_github}), probando por Apps Script de Google...")
        try:
            return _publicar_lote_apps_script(cfg, archivos, mensaje)
        except Exception as e_apps_script:
            raise RuntimeError(f"GitHub: {e_github} | Apps Script: {e_apps_script}")


def main():
    if not os.path.exists(DATA_PATH):
        print(f"No existe {DATA_PATH}. Corre primero procesar_diario.py.")
        sys.exit(1)
    cfg = cargar_config()
    mensaje = f"Actualizar datos (local) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"

    slot_anterior = _leer_puntero(cfg)
    slot_nuevo = "a" if slot_anterior == "b" else "b"
    print(f"Publicando en el slot '{slot_nuevo}' (el vigente hasta ahora es '{slot_anterior}') ...")

    archivos = []  # lista de (path_remoto, contenido_bytes)

    with open(DATA_PATH, "rb") as f:
        contenido = f.read()
    n_filas = len(json.loads(contenido).get("rows", []))
    path_slot = _con_slot(cfg["path"], slot_nuevo)
    print(f"Dashboard: {n_filas} filas ({path_slot})")
    archivos.append((path_slot, contenido))

    if os.path.exists(DATA_COMPLETA_PATH):
        with open(DATA_COMPLETA_PATH, "rb") as f:
            contenido_completa = f.read()
        n_completa = len(json.loads(contenido_completa).get("rows", []))
        path_completa_slot = _con_slot(cfg["path_completa"], slot_nuevo)
        print(f"BD completa/Buscar y exportar: {n_completa} filas ({path_completa_slot})")
        archivos.append((path_completa_slot, contenido_completa))
    else:
        print(f"AVISO: no existe {DATA_COMPLETA_PATH} — 'Buscar y exportar' no se actualizará esta vez.")

    # Archivos recortados por contratista (aislamiento de datos): cada uno
    # tiene su propia clave y su propio par de archivos, con SOLO sus filas.
    for contratista in CONTRATISTAS_CON_CLAVE:
        sufijo = f"_{contratista.lower()}"
        ruta_c = os.path.join(SALIDA_DIR, f"bd_actual{sufijo}.json")
        ruta_completa_c = os.path.join(SALIDA_DIR, f"bd_completa{sufijo}.json")
        path_remoto_c = _con_slot(cfg["path"].replace("bd_actual.json", f"bd_actual{sufijo}.json"), slot_nuevo)
        path_remoto_completa_c = _con_slot(cfg["path_completa"].replace("bd_completa.json", f"bd_completa{sufijo}.json"), slot_nuevo)

        if os.path.exists(ruta_c):
            with open(ruta_c, "rb") as f:
                contenido_c = f.read()
            n_c = len(json.loads(contenido_c).get("rows", []))
            print(f"{contratista}: {n_c} filas ({path_remoto_c})")
            archivos.append((path_remoto_c, contenido_c))
        else:
            print(f"AVISO: no existe {ruta_c} — {contratista} no verá datos nuevos esta vez.")

        if os.path.exists(ruta_completa_c):
            with open(ruta_completa_c, "rb") as f:
                contenido_completa_c = f.read()
            n_completa_c = len(json.loads(contenido_completa_c).get("rows", []))
            print(f"{contratista} BD completa: {n_completa_c} filas ({path_remoto_completa_c})")
            archivos.append((path_remoto_completa_c, contenido_completa_c))

    print()
    print(f"Publicando {len(archivos)} archivo(s) en el slot '{slot_nuevo}' ...")
    try:
        publicar_lote(cfg, archivos, mensaje)
    except Exception as e:
        print()
        print(f"No se pudo publicar: {e}")
        print(f"El puntero NO se movió (sigue en '{slot_anterior}') -- el sitio sigue viendo la versión anterior, nada quedó a medias.")
        print("Vuelve a correr Ejecutar.bat para reintentar.")
        sys.exit(1)

    # Recién ahora, con TODO el slot nuevo ya arriba y confirmado, se mueve
    # el punterito -- es un archivo chiquito y es el ÚNICO que se pisa en su
    # lugar de siempre, así que no hay ventana real de varios minutos donde
    # alguien pueda leer una mezcla (ver el comentario grande más arriba).
    updated_at = json.loads(contenido).get("updated_at")
    contenido_puntero = json.dumps({"slot": slot_nuevo, "updated_at": updated_at}).encode("utf-8")
    try:
        publicar_lote(cfg, [(_puntero_path(cfg), contenido_puntero)], mensaje + " (puntero)")
    except Exception as e:
        print()
        print(f"AVISO: los datos ya se subieron al slot '{slot_nuevo}', pero no se pudo mover el puntero: {e}")
        print("El sitio sigue mostrando la versión anterior. Vuelve a correr Ejecutar.bat para reintentar.")
        sys.exit(1)

    print()
    print(f"Publicado (slot '{slot_nuevo}'). El sitio va a tardar 1-2 minutos en mostrar los datos nuevos.")


if __name__ == "__main__":
    main()

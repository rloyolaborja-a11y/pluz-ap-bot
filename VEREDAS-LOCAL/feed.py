# -*- coding: utf-8 -*-
"""
feed.py — Publica SALIDA\\bd_actual.json y los recortes por contratista de
Veredas AP, a través de la Apps Script de Google (mismo "puente" que ya usan
Pendientes y Atendidas). Ver LEEME-APPS-SCRIPT.txt en la carpeta de arriba
("1. REPORTE\\") para armar esa Apps Script (se hace una sola vez y es
compartida por las 3 carpetas).

Solo necesita Python instalado — no requiere "pip install" de nada.
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

# 2026-09-20: publicar con GIT DE VERDAD (add/commit/push) como primer
# intento -- mismo motivo/patron que ya se uso para el sitio web
# (PUBLICAR-WEB-LOCAL/publicar_sitio.py): las llamadas HTTP crudas a la API
# de GitHub (mas abajo, _publicar_commit_github) son justo lo que el
# antivirus/firewall corporativo corta de vez en cuando ("WinError 10053,
# se ha anulado una conexion establecida por el software en su equipo
# host"); el ejecutable git.exe parece tratarse distinto. Si el modulo
# compartido no esta disponible (por ejemplo, todavia no se clono
# DATOS-GITHUB-LOCAL/pluz-ap-datos en esta PC), se sigue con los caminos de
# siempre sin romper nada.
sys.path.insert(0, os.path.join(BASE_DIR, "..", "DATOS-GITHUB-LOCAL"))
try:
    from publicar_datos_git import publicar_commit_git
    _GIT_DISPONIBLE = True
except Exception:
    _GIT_DISPONIBLE = False


def _contexto_ssl():
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


# 2026-09-16: mismo arreglo que en pendientes/feed.py -- Google Apps Script
# responde a /exec con un 302 hacia script.googleusercontent.com, y el
# manejador de redirección POR DEFECTO de urllib degrada un POST a GET al
# seguirlo (se pierde el body: accion:"datos" + archivos). Síntoma real
# visto en un log: "Falta el parámetro 'path'" -- ese es el error de
# doGet(), prueba de que la llamada llegó como GET sin body. Se usa un
# opener que, si el pedido era POST, lo reenvía como POST hacia la nueva URL.
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

CONTRATISTAS_CON_CLAVE = ["COBRA", "LARI", "PA", "NORTE"]

REINTENTOS = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4
TAMANO_MAX_POR_LLAMADA = 3 * 1024 * 1024


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
    return cfg


# 2026-09-16: publicación de "datos" DIRECTO a GitHub (repo aparte, sin
# Vercel conectado: rloyolaborja-a11y/pluz-ap-datos) en vez de Google Drive
# vía Apps Script -- ver el comentario grande en pendientes/feed.py. Se
# intenta GitHub primero; si falla, cae al camino viejo de Apps Script.
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
    vacío (sin ningún commit todavía) con la Contents API para el primero."""
    token = cfg.get("github_token")
    owner = cfg.get("github_datos_owner")
    repo = cfg.get("github_datos_repo")
    branch = cfg.get("github_datos_branch", "main")
    if not token or not owner or not repo:
        raise RuntimeError("Falta github_token/github_datos_owner/github_datos_repo en config.json.")

    base = f"https://api.github.com/repos/{owner}/{repo}"

    codigo, ref = _gh_fetch(f"{base}/git/refs/heads/{branch}", token)
    if codigo == 409:
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
# 2026-09-04: publicación por "slots" alternados (a/b) -- mismo esquema que
# Pendientes/Atendidas (ver el comentario grande en esos feed.py): cada
# publicación completa se sube a un slot que nadie está leyendo, y recién al
# final se mueve un punterito (data/latest.json) al slot nuevo, para que
# nadie pueda leer una mezcla de archivos viejos/nuevos a medio publicar.
# ---------------------------------------------------------------------------

def _puntero_path(cfg):
    base = cfg["path"].rsplit("/", 1)[0]  # ej. "veredas/data"
    return f"{base}/latest.json"


def _con_slot(path_remoto, slot):
    base, nombre = path_remoto.rsplit("/", 1)
    return f"{base}/{slot}/{nombre}"


def _leer_puntero(cfg):
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


def _publicar_lote_intento(cfg, archivos, mensaje):
    files_payload = [
        {"path": path_remoto, "content_base64": base64.b64encode(contenido).decode("ascii")}
        for path_remoto, contenido in archivos
    ]
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


def _agrupar_por_tamano(archivos):
    grupos, grupo_actual, tamano_actual = [], [], 0
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


# 2026-09-05: los envíos se mandaban uno por uno -- cada uno escribe en un
# archivo DISTINTO (nada compartido entre ellos), así que no hay ningún
# riesgo en mandarlos en paralelo. El puntero se sigue moviendo recién
# cuando esta función entera termina (todos los envíos), así que la
# atomicidad del slot no cambia.
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
            resultado_final = futuro.result()
    return resultado_final


def publicar_lote(cfg, archivos, mensaje):
    """Intenta git real primero (si el clon local esta disponible), despues
    la API HTTP de GitHub directo, y si todo eso falla cae al camino viejo
    de Apps Script/Drive."""
    if _GIT_DISPONIBLE:
        try:
            return publicar_commit_git(archivos, mensaje)
        except Exception as e_git:
            print(f"  git real falló ({e_git}), probando por la API HTTP de GitHub...")
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
    mensaje = f"Actualizar datos Veredas (local) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"

    slot_anterior = _leer_puntero(cfg)
    slot_nuevo = "a" if slot_anterior == "b" else "b"
    print(f"Publicando en el slot '{slot_nuevo}' (el vigente hasta ahora es '{slot_anterior}') ...")

    archivos = []
    with open(DATA_PATH, "rb") as f:
        contenido = f.read()
    n_filas = len(json.loads(contenido).get("rows", []))
    path_slot = _con_slot(cfg["path"], slot_nuevo)
    print(f"Dashboard/exportar: {n_filas} filas ({path_slot})")
    archivos.append((path_slot, contenido))

    for contratista in CONTRATISTAS_CON_CLAVE:
        ruta_c = os.path.join(SALIDA_DIR, f"bd_actual_{contratista.lower()}.json")
        path_remoto_c = _con_slot(cfg["path"].replace("bd_actual.json", f"bd_actual_{contratista.lower()}.json"), slot_nuevo)
        if os.path.exists(ruta_c):
            with open(ruta_c, "rb") as f:
                contenido_c = f.read()
            n_c = len(json.loads(contenido_c).get("rows", []))
            print(f"{contratista}: {n_c} filas ({path_remoto_c})")
            archivos.append((path_remoto_c, contenido_c))

    print()
    print(f"Publicando {len(archivos)} archivo(s) en el slot '{slot_nuevo}' ...")
    try:
        publicar_lote(cfg, archivos, mensaje)
    except Exception as e:
        print(f"\nNo se pudo publicar: {e}")
        print(f"El puntero NO se movió (sigue en '{slot_anterior}') -- nada quedó a medias.")
        print("Vuelve a correr Ejecutar.bat para reintentar.")
        sys.exit(1)

    updated_at = json.loads(contenido).get("updated_at")
    contenido_puntero = json.dumps({"slot": slot_nuevo, "updated_at": updated_at}).encode("utf-8")
    try:
        publicar_lote(cfg, [(_puntero_path(cfg), contenido_puntero)], mensaje + " (puntero)")
    except Exception as e:
        print(f"\nAVISO: los datos ya se subieron al slot '{slot_nuevo}', pero no se pudo mover el puntero: {e}")
        print("El sitio sigue mostrando la versión anterior. Vuelve a correr Ejecutar.bat para reintentar.")
        sys.exit(1)

    print(f"\nPublicado (slot '{slot_nuevo}'). El sitio va a tardar 1-2 minutos en mostrar los datos nuevos.")


if __name__ == "__main__":
    main()

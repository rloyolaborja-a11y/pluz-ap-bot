# -*- coding: utf-8 -*-
"""
publicar_historico.py — Publica los archivos de la carpeta HISTORICO\\
(los meses ya cerrados de Atendidas AP: 2026-01.json, 2026-02.json, ...,
más index.json con la lista) en tu repositorio, por el MISMO camino que
feed.py usa para bd_actual.json: Apps Script de Google -> Drive, acción
"datos" — así que NO pasa por GitHub/Vercel ni gasta cupo de despliegues,
y usa el mismo config.json que ya tienes (no hace falta configurar nada
aparte).

Por qué existe este script aparte de Publicar.bat (el de PUBLICAR-WEB-LOCAL):
Publicar.bat publica el CÓDIGO del sitio y a propósito NUNCA toca ninguna
carpeta llamada "data" (para no pisar datos en vivo) — pero
atendidas/data/historico/ vive justamente adentro de esa carpeta "data".
Este script sí sabe publicar ahí, igual que feed.py ya publica
atendidas/data/bd_actual.json.

Cuándo correrlo: UNA SOLA VEZ, para cargar los meses pasados que Claude ya
armó (ver HISTORICO\\). De ahí en adelante, cada mes que se cierre se va a
archivar solo, automáticamente, cada vez que proceses un Excel nuevo desde
la página — no hace falta volver a correr esto salvo que quieras corregir
o agregar un mes viejo a mano.

Cómo usarlo:
1. Doble clic en Publicar-Historico.bat (en esta misma carpeta).

Solo necesita Python instalado (desde python.org), igual que feed.py — no
hace falta instalar nada más.
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _contexto_ssl():
    """Mismo comentario que en feed.py: suma los certificados de Windows a
    los que ya trae Python, para que funcione igual en redes de empresa con
    inspección HTTPS."""
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
# seguirlo (se pierde el body). Se usa un opener que, si el pedido era POST,
# lo reenvía como POST hacia la nueva URL.
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
HISTORICO_DIR = os.path.join(BASE_DIR, "HISTORICO")


def cargar_config():
    if not os.path.exists(CONFIG_PATH):
        print(f"Falta config.json ({CONFIG_PATH}).")
        print("Es el mismo que usa feed.py — si Ejecutar.bat ya te publica los")
        print("datos del día a día, este archivo ya existe aquí mismo.")
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    faltan = [k for k in ("script_url", "secret") if not cfg.get(k)]
    if faltan:
        print(f"Faltan datos en config.json: {', '.join(faltan)}")
        sys.exit(1)
    return cfg


# 2026-09-16: publicación DIRECTO a GitHub (repo aparte, sin Vercel
# conectado: rloyolaborja-a11y/pluz-ap-datos) en vez de Google Drive vía
# Apps Script -- ver el comentario grande en pendientes/feed.py. Se intenta
# GitHub primero; si falla, cae al camino viejo de Apps Script.
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


REINTENTOS = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4


def _publicar_lote_intento(cfg, archivos, mensaje):
    """archivos: lista de (path_remoto, contenido_bytes)."""
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


# Cada mes pesa varios MB — se manda cada archivo en su propia llamada (no
# todos juntos) para no arriesgarse a que una sola llamada gigante se cuelgue.
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


def _publicar_lote_apps_script(cfg, archivos, mensaje):
    grupos = _agrupar_por_tamano(archivos)
    if len(grupos) > 1:
        print(f"  (se parte en {len(grupos)} envíos más chicos para evitar que se cuelgue la conexión)")
    resultado_final = None
    for i, grupo in enumerate(grupos, start=1):
        ultimo_error = None
        logrado = False
        for intento in range(1, REINTENTOS + 1):
            try:
                resultado_final = _publicar_lote_intento(cfg, grupo, mensaje)
                logrado = True
                break
            except Exception as e:
                ultimo_error = e
                etiqueta = f"envío {i}/{len(grupos)}, " if len(grupos) > 1 else ""
                print(f"  {etiqueta}intento {intento}/{REINTENTOS} falló: {e}")
                if intento < REINTENTOS:
                    time.sleep(ESPERA_ENTRE_REINTENTOS_SEG)
        if not logrado:
            raise ultimo_error
    return resultado_final


def publicar_lote(cfg, archivos, mensaje):
    """Intenta GitHub directo primero (un solo commit, sin necesidad de
    partir en grupos chicos) y, si falla, cae al camino viejo de Apps
    Script/Drive."""
    try:
        return _publicar_commit_github(cfg, archivos, mensaje)
    except Exception as e_github:
        print(f"  GitHub directo falló ({e_github}), probando por Apps Script de Google...")
        try:
            return _publicar_lote_apps_script(cfg, archivos, mensaje)
        except Exception as e_apps_script:
            raise RuntimeError(f"GitHub: {e_github} | Apps Script: {e_apps_script}")


def main():
    print("=== Publicar HISTÓRICO de Atendidas AP (meses ya cerrados) ===")
    print()
    if not os.path.isdir(HISTORICO_DIR):
        print(f"No existe la carpeta HISTORICO\\ ({HISTORICO_DIR}).")
        sys.exit(1)
    cfg = cargar_config()

    nombres = sorted(f for f in os.listdir(HISTORICO_DIR) if f.endswith(".json"))
    if not nombres:
        print("La carpeta HISTORICO\\ no tiene ningún archivo .json.")
        sys.exit(1)

    archivos = []  # lista de (path_remoto, contenido_bytes)
    for nombre in nombres:
        ruta_local = os.path.join(HISTORICO_DIR, nombre)
        with open(ruta_local, "rb") as f:
            contenido = f.read()
        if nombre == "index.json":
            n_meses = len(json.loads(contenido).get("meses", []))
            print(f"index.json: {n_meses} mes(es) en la lista")
        else:
            n_filas = len(json.loads(contenido).get("rows", []))
            print(f"{nombre}: {n_filas} filas")
        archivos.append((f"atendidas/data/historico/{nombre}", contenido))

    mensaje = f"Archivar histórico de Atendidas AP (local) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    print()
    print(f"Publicando {len(archivos)} archivo(s)…")
    try:
        publicar_lote(cfg, archivos, mensaje)
    except Exception as e:
        print()
        print(f"No se pudo publicar: {e}")
        print("Vuelve a correr Publicar-Historico.bat para reintentar.")
        sys.exit(1)

    print()
    print("Listo. Los meses archivados ya deberían verse en el desplegable")
    print('"Mes" de Atendidas AP (puede tardar unos segundos en reflejarse).')


if __name__ == "__main__":
    main()

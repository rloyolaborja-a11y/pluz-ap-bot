# -*- coding: utf-8 -*-
"""
publicar_sitio.py — Publica los archivos de la carpeta SITIO\\ (todo el
código de la página web: index.html, assets\\, pendientes\\, atendidas\\)
en tu repositorio de GitHub con GIT DE VERDAD (add/commit/push), igual que
lo haría cualquiera desde VS Code.

Cómo usarlo cada vez que te mande una versión nueva del portal:
1. Descomprime el zip que te doy DENTRO de la carpeta SITIO\\ de aquí,
   reemplazando lo que haya (dile que sí a "reemplazar" / "sobrescribir").
2. Doble clic en Publicar.bat.

IMPORTANTE (2026-09-16): antes esto pasaba por una Apps Script de Google
como "puente" hacia GitHub (para no depender de github.com directo) -- se
abandonó ese camino porque un antivirus/DLP corporativo bloquea las SUBIDAS
hacia *.googleusercontent.com (adonde redirige la Apps Script al
ejecutar), confirmado con pruebas reales. GitHub, en cambio, sí es
alcanzable desde acá, y usar `git` real es más simple y más estándar que
reimplementar la Git Data API a mano en Python. SITIO\\ ahora es un
repositorio git de verdad (`git init` ya corrido, remoto "origin" ya
apuntando a GitHub, credenciales guardadas en SITIO\\.git\\credentials-pluz
-- ESE archivo NO se publica en ningún lado, vive solo en tu PC dentro de
.git\\, igual de "solo local" que config.json).

IMPORTANTE: este script NO toca ninguna carpeta llamada "data" (ahí viven
bd_actual.json, bd_completa.json, data_requisito.json, criterios.json,
feriados.json, tecnicos.json) para no pisar datos ya publicados desde la
web o desde PENDIENTES-LOCAL / ATENDIDAS-LOCAL. Esos archivos se siguen
actualizando solo con "Procesar y actualizar" (en la web) o con el feed.py
de cada carpeta local.
"""

import hashlib
import os
import re
import subprocess
import sys
import time
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SITIO_DIR = os.path.join(BASE_DIR, "SITIO")

# Carpetas que NUNCA se publican desde aquí (datos en vivo, se manejan aparte).
CARPETAS_EXCLUIDAS = {"data"}
ARCHIVOS_EXCLUIDOS = {"config.json", "config.ejemplo.json", "Publicar.bat",
                       "publicar_sitio.py", "LEEME.txt", "LEEME-APPS-SCRIPT.txt",
                       ".DS_Store", "PON-AQUI-EL-PORTAL.txt"}


def _git(*args, timeout=120):
    """Corre un comando git DENTRO de SITIO\\ y devuelve (codigo, salida)."""
    resultado = subprocess.run(
        ["git"] + list(args), cwd=SITIO_DIR, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
    )
    salida = (resultado.stdout or "") + (resultado.stderr or "")
    return resultado.returncode, salida.strip()


# ---------------------------------------------------------------------------
# Cache-busting automático: en cada .html, los <script src=...js> y
# <link href=...css> de assets locales llevan ?v=<hash del archivo>. Así el
# navegador SIEMPRE toma la última versión sin tener que subir el número a
# mano (y sin romper caché cuando el archivo no cambió).
# ---------------------------------------------------------------------------
_RE_ASSET = re.compile(r'(src|href)="([^"?]+\.(?:js|css))(?:\?v=[^"]*)?"')


def _hash8(ruta):
    try:
        with open(ruta, "rb") as f:
            return hashlib.sha1(f.read()).hexdigest()[:8]
    except OSError:
        return None


def actualizar_cache_busting():
    cambios = 0
    for root, dirs, files in os.walk(SITIO_DIR):
        dirs[:] = [d for d in dirs if d not in CARPETAS_EXCLUIDAS and not d.startswith(".")]
        for nombre in files:
            if not nombre.endswith(".html"):
                continue
            ruta_html = os.path.join(root, nombre)
            with open(ruta_html, encoding="utf-8") as f:
                html = f.read()

            def _sub(m, _root=root):
                attr, rel = m.group(1), m.group(2)
                if rel.startswith(("http", "//", "data:")):
                    return m.group(0)
                h = _hash8(os.path.normpath(os.path.join(_root, rel)))
                return m.group(0) if not h else f'{attr}="{rel}?v={h}"'

            nuevo = _RE_ASSET.sub(_sub, html)
            if nuevo != html:
                with open(ruta_html, "w", encoding="utf-8") as f:
                    f.write(nuevo)
                cambios += 1
    if cambios:
        print(f"  cache-busting: {cambios} HTML actualizado(s)")


def listar_archivos():
    if not os.path.isdir(SITIO_DIR):
        print(f"No existe la carpeta SITIO\\ ({SITIO_DIR}). Descomprime ahí el zip del portal.")
        sys.exit(1)
    archivos = []
    for root, dirs, files in os.walk(SITIO_DIR):
        dirs[:] = [d for d in dirs if d not in CARPETAS_EXCLUIDAS and not d.startswith(".")]
        for nombre in files:
            if nombre in ARCHIVOS_EXCLUIDOS or nombre.startswith("."):
                continue
            ruta_local = os.path.join(root, nombre)
            ruta_repo = os.path.relpath(ruta_local, SITIO_DIR).replace("\\", "/")
            archivos.append((ruta_local, ruta_repo))
    return archivos


REINTENTOS = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4


def publicar_con_git(archivos, mensaje):
    """archivos: lista de (ruta_local, ruta_repo) -- ver listar_archivos().
    Hace exactamente lo que harías desde VS Code: git add de cada archivo
    (respetando la misma exclusión de siempre de la carpeta "data"), commit
    y push. Si no hay nada nuevo para subir, no falla -- simplemente avisa."""
    codigo, salida = _git("rev-parse", "--is-inside-work-tree")
    if codigo != 0:
        raise RuntimeError(
            f"SITIO\\ todavía no es un repositorio git (o el remoto no está configurado): {salida}")

    codigo, salida = _git("add", "--", *[ruta_repo for _, ruta_repo in archivos])
    if codigo != 0:
        raise RuntimeError(f"git add falló: {salida}")

    codigo, salida = _git("diff", "--cached", "--quiet")
    if codigo == 0:
        return {"publicado": False}

    codigo, salida = _git("commit", "-m", mensaje)
    if codigo != 0:
        raise RuntimeError(f"git commit falló: {salida}")

    ultimo_error = None
    for intento in range(1, REINTENTOS + 1):
        codigo, salida = _git("push", "origin", "HEAD:main", timeout=180)
        if codigo == 0:
            return {"publicado": True, "commit": salida}
        ultimo_error = salida
        print(f"    intento {intento}/{REINTENTOS} falló: {salida}")
        if intento < REINTENTOS:
            time.sleep(ESPERA_ENTRE_REINTENTOS_SEG)
    raise RuntimeError(f"git push falló después de {REINTENTOS} intentos: {ultimo_error}")


def main():
    print("=== Publicar sitio web (Pendientes AP + Atendidas AP) — vía git/GitHub directo ===")
    actualizar_cache_busting()
    archivos = listar_archivos()
    if not archivos:
        print("La carpeta SITIO\\ está vacía. Descomprime ahí el zip del portal antes de publicar.")
        sys.exit(1)

    print(f"Archivos a revisar: {len(archivos)}")
    mensaje = f"Actualizar sitio (local) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"

    try:
        resultado = publicar_con_git(archivos, mensaje)
    except Exception as e:
        print()
        print(f"No se pudo publicar: {e}")
        print("Vuelve a correr Publicar.bat para reintentar.")
        sys.exit(1)

    print()
    if not resultado["publicado"]:
        print("No había cambios nuevos para publicar (el sitio ya está al día).")
        return
    print("Publicado (1 commit, 1 despliegue en Vercel).")
    print("Listo. Vercel va a redesplegar solo en 1-2 minutos.")


if __name__ == "__main__":
    main()

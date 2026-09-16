# -*- coding: utf-8 -*-
"""
actualizar_requisito.py — Publica REQUISITO\\data_requisito.json (los mapas
de Contratista, Distrito, Zona, Motivo, Legal, SED, Poste, etc.) en
pendientes/data/data_requisito.json del repositorio del SITIO.

A diferencia de bd_actual.json / bd_completa.json (que se publican solos
cada vez que corres Ejecutar.bat), data_requisito.json NO cambia todos los
días — solo corre este script cuando Claude te pida actualizar esta tabla
(por ejemplo, para agregar un código de contratista nuevo).

2026-09-16: antes esto iba por la Apps Script de Google (bloqueada por un
antivirus/DLP corporativo, ver feed.py) -- ahora, como este archivo en
realidad vive DENTRO del repo del sitio (no en el repo aparte de datos),
se copia directo a PUBLICAR-WEB-LOCAL\\SITIO\\pendientes\\data\\ y se
publica con `git` de verdad (add/commit/push), igual que publicar_sitio.py.

Uso: doble clic no funciona con .py directamente — ábrelo con
"python actualizar_requisito.py" desde una consola en esta carpeta.
"""

import os
import shutil
import subprocess
import sys
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REQUISITO_PATH = os.path.join(BASE_DIR, "REQUISITO", "data_requisito.json")
SITIO_DIR = os.path.join(BASE_DIR, "..", "PUBLICAR-WEB-LOCAL", "SITIO")
PATH_REPO = "pendientes/data/data_requisito.json"


def _git(*args, timeout=120):
    resultado = subprocess.run(
        ["git"] + list(args), cwd=SITIO_DIR, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
    )
    salida = (resultado.stdout or "") + (resultado.stderr or "")
    return resultado.returncode, salida.strip()


def main():
    if not os.path.exists(REQUISITO_PATH):
        print(f"No existe {REQUISITO_PATH}.")
        sys.exit(1)
    if not os.path.isdir(os.path.join(SITIO_DIR, ".git")):
        print(f"{SITIO_DIR} todavía no es un repositorio git (ver publicar_sitio.py).")
        sys.exit(1)

    destino = os.path.join(SITIO_DIR, *PATH_REPO.split("/"))
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    shutil.copyfile(REQUISITO_PATH, destino)

    mensaje = f"Actualizar data_requisito (local) - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    print(f"Publicando {PATH_REPO} ...")

    codigo, salida = _git("add", "--", PATH_REPO)
    if codigo != 0:
        print(f"No se pudo publicar: git add falló: {salida}")
        sys.exit(1)

    codigo, salida = _git("diff", "--cached", "--quiet")
    if codigo == 0:
        print("No había cambios (data_requisito.json ya está igual en el sitio).")
        return

    codigo, salida = _git("commit", "-m", mensaje)
    if codigo != 0:
        print(f"No se pudo publicar: git commit falló: {salida}")
        sys.exit(1)

    codigo, salida = _git("push", "origin", "HEAD:main", timeout=180)
    if codigo != 0:
        print(f"No se pudo publicar: git push falló: {salida}")
        print("Vuelve a correr este script para reintentar.")
        sys.exit(1)

    print("Publicado. El sitio va a tardar 1-2 minutos en tomar la tabla nueva.")


if __name__ == "__main__":
    main()

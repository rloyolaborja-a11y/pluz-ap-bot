# -*- coding: utf-8 -*-
"""
publicar_datos_git.py — Publica archivos en el repo `pluz-ap-datos` con GIT
DE VERDAD (add/commit/push), igual que ya se hace para el sitio web
(PUBLICAR-WEB-LOCAL/publicar_sitio.py).

MOTIVO (2026-09-20): la publicacion de datos (feed.py de Pendientes/
Atendidas/Veredas) usaba llamadas HTTP directas a la API de GitHub con
Python (`requests`/`urllib`) -- eso es justo lo que el antivirus/firewall
corporativo interrumpe de vez en cuando (visto en vivo: "WinError 10053, se
ha anulado una conexion establecida por el software en su equipo host").
Cuando migramos el sitio web al mismo patron de `git` real hace unas
semanas, dejo de fallar por bloqueo de red -- la teoria es que la seguridad
de la empresa trata distinto al ejecutable `git.exe` (conocido/confiable)
que a conexiones HTTP crudas hechas desde Python. Se replica el mismo
patron aca para los datos.

Este modulo es COMPARTIDO por los 3 feed.py (Pendientes/Atendidas/Veredas)
-- todos publican al MISMO repo `pluz-ap-datos`, cada uno en su propia
subcarpeta (pendientes/, atendidas/, veredas/). Por eso el clon local vive
UNA sola vez aca, en DATOS-GITHUB-LOCAL/pluz-ap-datos, y no adentro de cada
carpeta de reporte.
"""

from __future__ import annotations

import subprocess
import time
from pathlib import Path

CLON_DIR = Path(__file__).resolve().parent / "pluz-ap-datos"

REINTENTOS_PUSH = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4


def _git(*args, timeout=120):
    resultado = subprocess.run(
        ["git"] + list(args), cwd=str(CLON_DIR), capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
    )
    salida = (resultado.stdout or "") + (resultado.stderr or "")
    return resultado.returncode, salida.strip()


def _verificar_repo() -> None:
    if not CLON_DIR.is_dir():
        raise RuntimeError(
            f"No existe el clon del repo de datos en {CLON_DIR}. "
            "Correr: git clone https://github.com/rloyolaborja-a11y/pluz-ap-datos.git "
            "dentro de DATOS-GITHUB-LOCAL\\, y configurar user.name/user.email/"
            "credential.helper (ver LEEME de esta carpeta)."
        )
    codigo, salida = _git("rev-parse", "--is-inside-work-tree")
    if codigo != 0:
        raise RuntimeError(f"{CLON_DIR} no es un repositorio git valido: {salida}")


def publicar_commit_git(archivos, mensaje: str) -> dict:
    """archivos: lista de (path_remoto, contenido_bytes) -- path_remoto
    relativo a la raiz del repo (ej. "atendidas/data/a/bd_actual.json").
    Escribe cada archivo en el clon local, hace add/commit/push. Si no hay
    nada nuevo para subir, no falla -- avisa {"publicado": False}."""
    _verificar_repo()

    # Traer lo ultimo antes de escribir -- los 3 reportes (y otras PCs)
    # publican al mismo repo, mejor partir de la punta actual para no
    # terminar con un push rechazado por "no fast-forward".
    codigo, salida = _git("pull", "--rebase", "origin", "main", timeout=60)
    if codigo != 0:
        raise RuntimeError(f"git pull (antes de publicar) fallo: {salida}")

    rutas_repo = []
    for path_remoto, contenido in archivos:
        destino = CLON_DIR / path_remoto
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_bytes(contenido)
        rutas_repo.append(path_remoto)

    codigo, salida = _git("add", "--", *rutas_repo)
    if codigo != 0:
        raise RuntimeError(f"git add fallo: {salida}")

    codigo, salida = _git("diff", "--cached", "--quiet")
    if codigo == 0:
        return {"publicado": False}

    codigo, salida = _git("commit", "-m", mensaje)
    if codigo != 0:
        raise RuntimeError(f"git commit fallo: {salida}")

    ultimo_error = None
    for intento in range(1, REINTENTOS_PUSH + 1):
        codigo, salida = _git("push", "origin", "HEAD:main", timeout=120)
        if codigo == 0:
            return {"publicado": True, "commit": salida}
        ultimo_error = salida
        print(f"    (git push) intento {intento}/{REINTENTOS_PUSH} fallo: {salida}")
        if intento < REINTENTOS_PUSH:
            time.sleep(ESPERA_ENTRE_REINTENTOS_SEG)
    raise RuntimeError(f"git push fallo despues de {REINTENTOS_PUSH} intentos: {ultimo_error}")

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

import os
import subprocess
import sys
import time
from pathlib import Path

CLON_DIR = Path(__file__).resolve().parent / "pluz-ap-datos"

# (2026-09-20) Sin esto, CADA comando de git (add/commit/push/pull, uno por
# uno) abre su propia ventanita de consola en Windows y la cierra enseguida
# -- confirmado por la usuaria viendolo parpadear en pantalla durante una
# corrida "silenciosa". CREATE_NO_WINDOW hace que corran ocultos, sin abrir
# ninguna ventana.
_FLAGS_SIN_VENTANA = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

REINTENTOS_PUSH = 3
ESPERA_ENTRE_REINTENTOS_SEG = 4


def _git(*args, timeout=120):
    resultado = subprocess.run(
        ["git"] + list(args), cwd=str(CLON_DIR), capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
        creationflags=_FLAGS_SIN_VENTANA,
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


_LOCK_PATH = CLON_DIR.parent / ".publicar.lock"
_LOCK_TIMEOUT_SEG = 180  # si un candado queda "colgado" mas que esto, se pisa


def _adquirir_lock() -> None:
    """Candado de archivo simple -- necesario porque Pendientes/Atendidas/
    Veredas ahora publican EN PARALELO (procesos de Python separados, no
    hilos del mismo proceso) al MISMO clon local. Sin esto, dos `git`
    corriendo a la vez sobre la misma carpeta se pisan entre si (o tiran
    "Unable to create '.git/index.lock'"). Si dos procesos llegan juntos,
    el segundo espera a que el primero termine, en vez de fallar."""
    t0 = time.time()
    while True:
        try:
            fd = os.open(str(_LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return
        except FileExistsError:
            if time.time() - t0 > _LOCK_TIMEOUT_SEG:
                # Candado viejo (algun proceso murio sin liberarlo) -- se pisa
                # para no bloquear publicaciones para siempre.
                try:
                    _LOCK_PATH.unlink()
                except Exception:
                    pass
                continue
            time.sleep(1)


def _liberar_lock() -> None:
    try:
        _LOCK_PATH.unlink()
    except Exception:
        pass


def publicar_commit_git(archivos, mensaje: str) -> dict:
    """archivos: lista de (path_remoto, contenido_bytes) -- path_remoto
    relativo a la raiz del repo (ej. "atendidas/data/a/bd_actual.json").
    Escribe cada archivo en el clon local, hace add/commit/push. Si no hay
    nada nuevo para subir, no falla -- avisa {"publicado": False}."""
    _verificar_repo()
    _adquirir_lock()
    try:
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
    finally:
        _liberar_lock()

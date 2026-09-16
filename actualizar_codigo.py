# -*- coding: utf-8 -*-
"""
actualizar_codigo.py — Trae los últimos cambios del código del bot (Python,
panel, RPA) desde el repositorio de GitHub (rloyolaborja-a11y/pluz-ap-bot),
con `git pull`. NO toca tus datos, tus config.json ni ninguna clave -- eso
nunca vive en este repo (ver .gitignore).

Uso: botón "Actualizar código del bot" en el panel, o doble clic en
Actualizar-Codigo.bat.
"""

import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _git(*args, timeout=60):
    resultado = subprocess.run(
        ["git"] + list(args), cwd=BASE_DIR, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=timeout,
    )
    salida = (resultado.stdout or "") + (resultado.stderr or "")
    return resultado.returncode, salida.strip()


def main():
    print("=== Actualizar código del bot (git pull) ===")
    if not os.path.isdir(os.path.join(BASE_DIR, ".git")):
        print("Esta carpeta todavía no es un repositorio git.")
        sys.exit(1)

    codigo, salida = _git("pull", "--ff-only", "origin", "main", timeout=120)
    print(salida)
    if codigo != 0:
        print()
        print("No se pudo actualizar. Si dice algo de 'cambios locales' o")
        print("'divergent branches', avisa a Claude para revisarlo -- no se")
        print("tocó nada más.")
        sys.exit(1)

    if "Already up to date" in salida or "ya está actualizado" in salida.lower():
        print()
        print("Ya estabas al día, no había nada nuevo.")
    else:
        print()
        print("Listo, código actualizado.")


if __name__ == "__main__":
    main()

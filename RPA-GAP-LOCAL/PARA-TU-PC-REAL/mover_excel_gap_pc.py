#!/usr/bin/env python3
"""Toma el Excel del GAP mas reciente de la carpeta OneDrive "GAP"
(sincronizada desde el escritorio virtual) y lo copia a ..\\CARGA\\EXCEL,
reemplazando cualquier Excel viejo que hubiera ahi.

ESTE SCRIPT SE EJECUTA EN TU PC REAL (no en la VM) -- es el segundo paso,
despues de correr Extraer-GAP-VM.bat dentro del escritorio virtual y
esperar a que OneDrive termine de sincronizar el archivo nuevo.

COMO SE USA
-----------
1. Doble clic en Mover-Excel-GAP-PC.bat (en esta misma carpeta, EN TU PC
   REAL).
2. Listo: el Excel mas reciente de la carpeta OneDrive "GAP" queda copiado
   en CARGA\\EXCEL.
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

# ----------------------------------------------------------------------
# Configuracion basica
# ----------------------------------------------------------------------

# 2026-09-16: antes esta ruta traia el usuario de Windows de ESTA PC
# ("P721725611") escrito a mano -- funcionaba solo aca. Ahora se detecta
# sola (mismo mecanismo que ya usa vigilar_gap_vm.py del lado de la VM) para
# que esta misma carpeta sirva tal cual en cualquier otra PC (ej. la de
# respaldo), sin tener que tocar el codigo -- PERO la detección automática
# asume que la carpeta se llama exactamente "GAP" dentro de OneDrive. Si en
# otra PC se llama distinto (o el detector no la encuentra por lo que sea),
# se puede pegar la ruta exacta en "ruta_carpeta_gap.txt" (mismo lugar que
# este archivo) -- una sola línea con la ruta completa, nada más. Si ese
# archivo existe y apunta a una carpeta real, gana siempre por sobre el
# detector automático.
RUTA_MANUAL_TXT = Path(__file__).resolve().parent / "ruta_carpeta_gap.txt"


def _leer_ruta_manual():
    if not RUTA_MANUAL_TXT.exists():
        return None
    texto = RUTA_MANUAL_TXT.read_text(encoding="utf-8").strip()
    if not texto:
        return None
    ruta = Path(texto)
    if not ruta.is_dir():
        print(f"AVISO: 'ruta_carpeta_gap.txt' apunta a una carpeta que no existe:\n  {ruta}")
        print("Se ignora y se intenta detectar sola.")
        return None
    return ruta


def _detectar_carpeta_onedrive_gap() -> Path:
    manual = _leer_ruta_manual()
    if manual:
        return manual
    candidatos_base = []
    for var in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive"):
        v = os.environ.get(var)
        if v:
            candidatos_base.append(Path(v))
    for base in candidatos_base:
        posible = base / "GAP"
        if posible.is_dir():
            return posible
    try:
        for posible in Path("C:/Users").glob("*/OneDrive*/GAP"):
            if posible.is_dir():
                return posible
    except Exception:
        pass
    raise SystemExit(
        "ERROR: no encuentro la carpeta compartida 'GAP' de OneDrive en esta "
        "PC. Revisa que OneDrive este sincronizado y que la carpeta se llame "
        "'GAP' (dentro de alguna carpeta que empiece con 'OneDrive' en tu "
        "carpeta de usuario)."
    )


CARPETA_ONEDRIVE_GAP = _detectar_carpeta_onedrive_gap()
# CARGA\EXCEL vive siempre 3 niveles arriba de este archivo
# (RPA-GAP-LOCAL\PARA-TU-PC-REAL\este_archivo.py) -- calculado relativo al
# propio script, no a una ruta fija, para que funcione en cualquier PC.
CARPETA_DESTINO = Path(__file__).resolve().parent.parent.parent / "CARGA" / "EXCEL"

EXTENSIONES_EXCEL = (".xls", ".xlsx")

# Cuanto esperar (segundos) a que OneDrive termine de bajar/sincronizar el
# archivo antes de darlo por listo (por si el archivo mas nuevo todavia se
# esta escribiendo cuando corres este script muy pegado al de la VM).
ESPERA_SINCRONIZACION_SEG = 5


def archivo_excel_mas_reciente(carpeta: Path) -> Path:
    if not carpeta.exists():
        print(f"ERROR: no encuentro la carpeta OneDrive 'GAP':\n  {carpeta}")
        print("Revisa que la ruta en CARPETA_ONEDRIVE_GAP (arriba en este")
        print("archivo) sea exactamente la misma que ves en el Explorador de")
        print("Windows.")
        raise SystemExit(1)

    candidatos = [
        p for p in carpeta.iterdir()
        if p.is_file() and p.suffix.lower() in EXTENSIONES_EXCEL
    ]
    if not candidatos:
        print(f"ERROR: no hay ningun Excel (.xls/.xlsx) en:\n  {carpeta}")
        print("¿Ya corriste Extraer-GAP-VM.bat dentro de la VM y esperaste a")
        print("que OneDrive terminara de sincronizar?")
        raise SystemExit(1)

    return max(candidatos, key=lambda p: p.stat().st_mtime)


def limpiar_destino(carpeta: Path) -> None:
    carpeta.mkdir(parents=True, exist_ok=True)
    for p in carpeta.iterdir():
        if p.is_file() and p.suffix.lower() in EXTENSIONES_EXCEL:
            try:
                p.unlink()
            except Exception as exc:
                print(f"  (aviso: no pude borrar {p.name}: {exc})")


def _meses_pedidos() -> int:
    for i, a in enumerate(sys.argv):
        if a == "--meses" and i + 1 < len(sys.argv):
            try:
                return max(1, min(6, int(sys.argv[i + 1])))
            except ValueError:
                return 1
        if a.startswith("--meses="):
            try:
                return max(1, min(6, int(a.split("=", 1)[1])))
            except ValueError:
                return 1
    return 1


def correr() -> None:
    meses = _meses_pedidos()
    print(f"Buscando el/los {meses} Excel mas reciente(s) en:\n  {CARPETA_ONEDRIVE_GAP}\n")
    time.sleep(ESPERA_SINCRONIZACION_SEG)

    if not CARPETA_ONEDRIVE_GAP.exists():
        archivo_excel_mas_reciente(CARPETA_ONEDRIVE_GAP)  # imprime el error y sale
    candidatos = sorted(
        (p for p in CARPETA_ONEDRIVE_GAP.iterdir()
         if p.is_file() and p.suffix.lower() in EXTENSIONES_EXCEL),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    if len(candidatos) < meses:
        print(f"ERROR: se esperaban {meses} Excel y hay {len(candidatos)} en:\n  {CARPETA_ONEDRIVE_GAP}")
        print("¿Ya corriste la extraccion en la VM y espero OneDrive a sincronizar?")
        raise SystemExit(1)

    elegidos = candidatos[:meses]
    for p in elegidos:
        print(f"Encontrado: {p.name} (modificado {time.ctime(p.stat().st_mtime)})")

    print(f"\nLimpiando Excel(s) viejos en:\n  {CARPETA_DESTINO}")
    limpiar_destino(CARPETA_DESTINO)

    for origen in elegidos:
        destino = CARPETA_DESTINO / origen.name
        destino.write_bytes(origen.read_bytes())
        print(f"  copiado: {destino.name}")

    print("\nListo.")


if __name__ == "__main__":
    try:
        correr()
    except SystemExit as exc:
        print(f"\n{exc}")
        sys.exit(1)

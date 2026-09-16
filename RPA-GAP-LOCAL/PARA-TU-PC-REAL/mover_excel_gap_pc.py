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

import sys
import time
from pathlib import Path

# ----------------------------------------------------------------------
# Configuracion basica -- ajusta aqui si alguna de las dos rutas cambia
# ----------------------------------------------------------------------

CARPETA_ONEDRIVE_GAP = Path(
    r"C:\Users\P721725611\OneDrive - Pluz Energía Perú S.A.A\GAP"
)
CARPETA_DESTINO = Path(
    r"C:\Users\P721725611\Desktop\Programaciones\1. REPORTE\CARGA\EXCEL"
)

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

#!/usr/bin/env python3
"""Vigila la carpeta compartida de OneDrive "GAP" (vista desde ESTA VM,
Oracle Secure Desktops) esperando una SEÑAL que manda el robot de la PC
real (iniciar_y_esperar_gap_pc.py, en PARA-TU-PC-REAL) -- en cuanto
aparece, arranca sola la extraccion completa (el mismo flujo probado de
extraer_gap_vm.py), sin que haga falta estar manejando el mouse dentro de
la VM en ese momento.

REQUISITO: SDAPeru ya tiene que estar abierto, con sesion iniciada (login
a mano, como siempre) y el arbol "Gestion de Alumbrado Publico" visible --
este robot NUNCA toca contrasenas ni abre SDAPeru por su cuenta. Si la
sesion vence mientras este script esta vigilando, la extraccion va a
fallar con el mismo error de siempre ("no encuentro la ventana..."),
queda anotado en diagnosticos, y el script SIGUE vigilando igual (no se
cae) -- solo hay que volver a iniciar sesion en SDAPeru y mandar la senal
de nuevo.

COMO SE USA
-----------
1. Dentro de la VM, con SDAPeru ya abierto y logueado: doble clic en
   Vigilar-GAP-VM.bat (en esta misma carpeta).
2. Se queda corriendo en una consola, revisando la carpeta de senales cada
   pocos segundos -- dejala abierta (se puede minimizar).
3. Cuando alguien corra Iniciar-GAP-PC.bat en la PC real, esta consola lo
   detecta sola y arranca la extraccion -- imprime todo igual que
   extraer_gap_vm.py, y al terminar vuelve a quedar vigilando.
4. Para parar de vigilar: Ctrl+C en esta consola, o cerrarla nomas.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import json
import os
import sys
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import extraer_gap_vm as gap  # reusa TODO el flujo de extraccion ya probado

INTERVALO_VIGILANCIA_SEG = 15
SUBCARPETA_SENAL = "_senal_inicio"

# (2026-09-12) "Clic fantasma": mientras este script vigila sin que haya
# NADIE tocando el mouse/teclado de verdad dentro de esta sesion remota, es
# posible que un timeout de inactividad del lado del ESCRITORIO REMOTO (no
# de la pagina web del portal, que es otra capa aparte, del lado de la PC
# real) la desconecte -- son sistemas VDI tipicos, Oracle Secure Desktops
# entre ellos. No sabemos con certeza si ese timeout existe ni cuanto dura,
# asi que probamos: cada 2 minutos, mueve el mouse 1px y lo devuelve --
# actividad minima e invisible, pero suficiente para contar como "no
# inactivo" ante cualquier detector estandar.
INTERVALO_CLIC_FANTASMA_SEG = 120

# Se pausa mientras hay una extraccion real corriendo (pywinauto manejando
# SDAPeru) -- para no arriesgar mover el mouse en medio de un hover/drag y
# meter ruido en la automatizacion de verdad.
_EXTRAYENDO = threading.Event()


def _clic_fantasma_loop():
    while True:
        if not _EXTRAYENDO.is_set():
            try:
                pt = ctypes.wintypes.POINT()
                ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                ctypes.windll.user32.SetCursorPos(pt.x + 1, pt.y)
                time.sleep(0.05)
                ctypes.windll.user32.SetCursorPos(pt.x, pt.y)
            except Exception:
                pass
        time.sleep(INTERVALO_CLIC_FANTASMA_SEG)


def _detectar_carpeta_onedrive_gap() -> Path:
    """Encuentra la carpeta compartida 'GAP' de OneDrive SIN depender de
    saber de antemano el nombre de usuario de esta VM (puede ser distinto
    al de la PC real -- ya lo vimos en los diagnosticos: la VM usa
    'DesktopUser' o similar, no 'P721725611'). Primero prueba las
    variables de entorno que Windows define para OneDrive, y si no las
    encuentra, busca cualquier carpeta 'GAP' dentro de
    C:\\Users\\*\\OneDrive*\\."""
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
        "VM. Revisa que OneDrive este sincronizado y que la carpeta se llame "
        "'GAP' (dentro de alguna carpeta que empiece con 'OneDrive' en tu "
        "carpeta de usuario). Si la ruta es rara, avisame la ruta exacta "
        "(la ves en el Explorador de Windows) para dejarla fija en el script."
    )


def vigilar() -> None:
    carpeta_gap = _detectar_carpeta_onedrive_gap()
    carpeta_senal = carpeta_gap / SUBCARPETA_SENAL
    carpeta_senal.mkdir(parents=True, exist_ok=True)
    print(f"Carpeta OneDrive 'GAP' detectada en:\n  {carpeta_gap}")
    print(f"Vigilando senales de inicio en:\n  {carpeta_senal}")
    print(f"(reviso cada {INTERVALO_VIGILANCIA_SEG}s -- Ctrl+C para parar)\n")

    threading.Thread(target=_clic_fantasma_loop, daemon=True).start()
    print(f"Clic fantasma activado (mueve el mouse 1px cada {INTERVALO_CLIC_FANTASMA_SEG}s) "
          "-- para que un timeout de inactividad del escritorio remoto no corte la sesion.\n")

    # (2026-09-11) Los diagnosticos (foto de pantalla + .txt cuando algo
    # falla) por defecto se guardan en el disco LOCAL de la VM
    # (%USERPROFILE%\GAP_RPA_Excel\diagnosticos) -- invisibles desde la PC
    # real. Como YA estamos corriendo dentro de la vigilancia automatica
    # (el flujo que realmente usa el panel), los mandamos en cambio a la
    # MISMA carpeta compartida de OneDrive "GAP" -- asi sincronizan solos y
    # se pueden revisar desde la PC real sin entrar a la VM.
    carpeta_diag = carpeta_gap / "_diagnosticos"
    gap.DIAG_BASE_DIR = carpeta_gap
    gap.DIAGNOSTICS_DIR = carpeta_diag
    print(f"Diagnosticos (si algo falla) se guardan en:\n  {carpeta_diag}\n"
          "(esa carpeta sincroniza con OneDrive -- se puede revisar desde la PC real)\n")

    # (2026-09-12) LATIDO en archivo (no solo en consola): antes, si la VM
    # se apagaba (o esta consola se cerraba, o la sesion de SDAPeru quedaba
    # inservible), la PC real no tenia forma de saberlo hasta agotar los 20
    # minutos completos esperando Excels que nunca iban a llegar. Ahora se
    # escribe un latido en la MISMA carpeta OneDrive "GAP" --
    # iniciar_y_esperar_gap_pc.py lo revisa y, si esta viejo, avisa de
    # entrada que la VM no esta respondiendo en vez de esperar a ciegas.
    #
    # (2026-09-14) BUG encontrado: el latido se escribia solo entre vueltas
    # del loop principal -- pero una extraccion real (gap.correr) puede
    # tardar varios minutos, y durante ESE rato el loop esta ocupado (no
    # llega a la linea que escribe el latido). Resultado: el latido se
    # quedaba "viejo" justo mientras la VM SI estaba trabajando de verdad,
    # y iniciar_y_esperar_gap_pc.py cortaba la espera pensando que la VM
    # no respondia -- aunque los 3 Excel iban a llegar bien poco despues.
    # Arreglo: el latido ahora es un hilo APARTE, independiente de lo que
    # este haciendo el loop principal (extrayendo o no) -- late cada 15s
    # SIEMPRE, mientras este proceso siga vivo.
    ruta_latido = carpeta_gap / "_latido_vm.json"

    def _latido_loop():
        while True:
            try:
                ruta_latido.write_text(
                    json.dumps({"ultimo": datetime.now().isoformat(timespec="seconds")}),
                    encoding="utf-8",
                )
            except Exception:
                pass
            time.sleep(INTERVALO_VIGILANCIA_SEG)

    threading.Thread(target=_latido_loop, daemon=True).start()

    # (2026-09-03) IMPORTANTE: antes esto se quedaba MUDO mientras no
    # encontraba ninguna senal -- si algo andaba mal (por ejemplo, OneDrive
    # tardando en sincronizar el archivo de la PC real hacia esta VM, o
    # algun problema de ruta) no habia forma de saberlo sin adivinar. Ahora
    # cada ~1 minuto imprime un "laton" (heartbeat) mostrando TODO lo que
    # hay en la carpeta en ese momento (no solo los .trigger), para poder
    # ver de un vistazo si OneDrive esta poniendo algo ahi o no.
    ticks_por_laton = max(1, round(60 / INTERVALO_VIGILANCIA_SEG))
    contador = 0

    while True:
        try:
            # (2026-09-03) ".txt" en vez de ".trigger" -- ver el comentario
            # en iniciar_y_esperar_gap_pc.py (mandar_senal()): la extension
            # rara ".trigger" es sospechosa de que la bloquee alguna
            # politica de SharePoint/OneDrive sin avisar nada. El prefijo
            # "INICIAR_" evita confundir esto con cualquier otro .txt que
            # alguien deje suelto en esta carpeta.
            senales = sorted(carpeta_senal.glob("INICIAR_*.txt"))
        except Exception as exc:
            print(f"  (aviso: no pude revisar la carpeta de senales: {exc})")
            senales = []

        contador += 1
        if not senales and contador % ticks_por_laton == 0:
            try:
                todo_lo_que_hay = sorted(p.name for p in carpeta_senal.iterdir())
            except Exception as exc:
                todo_lo_que_hay = [f"(no se pudo listar la carpeta: {exc})"]
            print(
                f"  ({datetime.now():%H:%M:%S} sigo vigilando, nada nuevo todavia -- "
                f"contenido actual de la carpeta: {todo_lo_que_hay or '(vacia)'})"
            )

        if senales:
            senal = senales[0]
            print(f"\n=== Senal detectada: {senal.name} -- arrancando extraccion ===")
            # La senal puede traer "MESES=3" en su contenido (la manda asi
            # iniciar_y_esperar_gap_pc.py --meses 3). Sin eso -> modo clasico.
            meses = 1
            try:
                txt = senal.read_text(encoding="utf-8", errors="ignore")
                import re as _re
                m = _re.search(r"MESES\s*=\s*(\d+)", txt)
                if m:
                    meses = max(1, min(6, int(m.group(1))))
            except Exception:
                pass
            try:
                senal.unlink()
            except Exception as exc:
                print(f"  (aviso: no pude borrar la senal {senal.name}: {exc} -- sigo igual)")

            if meses > 1:
                ventanas = gap.ventanas_meses(meses)
                print(f"Modo 3 meses: {meses} ventanas de {gap.RANGO_DIAS} dias.")
                for d, h, s in ventanas:
                    print(f"  {s}: {d} a {h}")
            else:
                hasta_dt = datetime.now()
                desde_dt = hasta_dt - timedelta(days=gap.RANGO_DIAS)
                ventanas = [(desde_dt.strftime("%d/%m/%Y"), hasta_dt.strftime("%d/%m/%Y"), None)]
                print(f"Rango: {ventanas[0][0]} a {ventanas[0][1]}")
            print(f"Version del script de extraccion: {gap.VERSION_SCRIPT}")
            _EXTRAYENDO.set()
            try:
                gap.correr(ventanas)
            except SystemExit as exc:
                print(f"\n{exc}")
            except Exception as exc:  # red de seguridad final -- nunca dejar caer el vigilante
                gap.guardar_diagnostico("error_inesperado_vigilancia", error=exc)
                print(f"\nFallo inesperado: {type(exc).__name__}: {exc}")
            finally:
                _EXTRAYENDO.clear()
            print("\n=== Vuelvo a vigilar ===\n")

        time.sleep(INTERVALO_VIGILANCIA_SEG)


if __name__ == "__main__":
    try:
        vigilar()
    except KeyboardInterrupt:
        print("\nVigilancia detenida por el usuario.")

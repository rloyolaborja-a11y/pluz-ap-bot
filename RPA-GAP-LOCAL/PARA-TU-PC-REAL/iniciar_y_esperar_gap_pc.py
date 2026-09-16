#!/usr/bin/env python3
"""Arranca la extraccion del GAP mandando una SEÑAL a la VM (a traves de
la carpeta compartida de OneDrive "GAP") y espera a que el Excel nuevo
aparezca ahi para copiarlo solo a CARGA\\EXCEL -- todo en un solo paso,
SIN tener que entrar vos a la VM.

ESTE SCRIPT SE EJECUTA EN TU PC REAL (no en la VM).

REQUISITO: en la VM tiene que estar corriendo Vigilar-GAP-VM.bat (con
SDAPeru ya abierto y logueado) -- si no esta corriendo ahi, la senal se
va a quedar esperando sin que nadie la lea, y este script se va a quedar
esperando el Excel hasta que se agote el tiempo (avisa igual, no se cae
en silencio).

COMO SE USA
-----------
1. Doble clic en Iniciar-GAP-PC.bat (en esta misma carpeta, EN TU PC
   REAL).
2. Espera -- puede tardar varios minutos (la VM tiene que leer la senal,
   hacer la extraccion completa, y OneDrive tiene que sincronizar el
   archivo de vuelta a tu PC). El programa va avisando en que va.
3. Al terminar, el Excel ya esta copiado en CARGA\\EXCEL.

Si se agota el tiempo de espera (por ejemplo, no habia nadie vigilando en
la VM, o tardo mas de lo normal), este script te avisa bien claro y podes
correr Mover-Excel-GAP-PC.bat mas tarde a mano, en cuanto el archivo ya
haya llegado -- sin tener que repetir el envio de la senal.
"""

from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# ----------------------------------------------------------------------
# Configuracion basica -- mismas rutas que ya usa mover_excel_gap_pc.py
# ----------------------------------------------------------------------

# 2026-09-16: detectada sola en vez de escrita a mano -- ver el mismo
# comentario en mover_excel_gap_pc.py. Si la carpeta no se llama "GAP" en
# esta PC, se puede pegar la ruta exacta en "ruta_carpeta_gap.txt" (mismo
# lugar que este archivo, una sola línea) -- gana siempre sobre el detector.
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
CARPETA_DESTINO = Path(__file__).resolve().parent.parent.parent / "CARGA" / "EXCEL"

SUBCARPETA_SENAL = "_senal_inicio"
EXTENSIONES_EXCEL = (".xls", ".xlsx")

ESPERA_MAX_SEG = 20 * 60  # 20 minutos -- VM + ida y vuelta de OneDrive puede tardar
INTERVALO_POLL_SEG = 5
SEGUNDOS_ESTABLE = 3  # cuanto tiene que dejar de crecer el archivo para darlo por listo

# (2026-09-12) Vigilar-GAP-VM.bat (version actualizada) escribe un "latido"
# en esta misma carpeta compartida cada ~15s mientras esta vigilando. Si ese
# latido esta viejo, es que la VM esta apagada, la sesion se cerro, o nadie
# dejo corriendo Vigilar-GAP-VM.bat -- en vez de esperar los 20 minutos
# completos a ciegas (como pasaba antes), avisamos fuerte y bien temprano.
# Umbral generoso (5 min) para no confundir con el retraso normal de
# sincronizacion de OneDrive.
RUTA_LATIDO_VM = CARPETA_ONEDRIVE_GAP / "_latido_vm.json"
UMBRAL_LATIDO_SEG = 5 * 60
VM_NO_RESPONDE = False  # lo pone en True esperar_excels_nuevos() si corta la espera por esto


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


def _purgar_senales_viejas(carpeta_senal: Path, minutos: int = 2) -> None:
    """Borra señales INICIAR_*.txt que quedaron sin consumir hace rato.

    (2026-09-11) Si una corrida anterior mandó la señal y después se
    canceló del lado de la PC (por ejemplo: falló el SAP y se detuvo la
    corrida), esa señal se queda ahí -- Vigilar-GAP-VM.bat no tiene forma de
    saber que ya no hace falta, y más tarde la procesa igual (encolada),
    aunque nadie la esté esperando. Si Vigilar-GAP-VM.bat está corriendo
    normalmente, revisa cada 15s, así que una señal que sigue sin tocar
    después de este ratito casi seguro es basura de un intento anterior
    (o directamente no hay nadie vigilando en la VM) -- se limpia antes de
    mandar la señal nueva, para no acumular corridas fantasma en la VM."""
    limite = time.time() - minutos * 60
    try:
        viejas = list(carpeta_senal.glob("INICIAR_*.txt"))
    except Exception:
        return
    for p in viejas:
        try:
            if p.stat().st_mtime < limite:
                p.unlink()
                print(f"  (se limpió una señal vieja sin consumir: {p.name})")
        except Exception:
            pass


def mandar_senal(meses: int = 1) -> None:
    carpeta_senal = CARPETA_ONEDRIVE_GAP / SUBCARPETA_SENAL
    carpeta_senal.mkdir(parents=True, exist_ok=True)
    _purgar_senales_viejas(carpeta_senal)
    # (2026-09-03) Extension ".txt" en vez de ".trigger": la extension rara
    # ".trigger" es sospechosa de que algunas politicas corporativas de
    # SharePoint/OneDrive (bloqueo de tipos de archivo) la frenen SIN avisar
    # nada -- el archivo queda "subido" en tu PC (por eso mostraba el visto
    # verde) pero nunca se replica a otros dispositivos. ".txt" es un tipo
    # de archivo comun que ninguna politica razonable bloquea.
    nombre = f"INICIAR_{datetime.now():%Y%m%d_%H%M%S}.txt"
    (carpeta_senal / nombre).write_text(
        "Senal para arrancar la extraccion del GAP -- generada automaticamente "
        f"desde la PC real el {datetime.now():%d/%m/%Y %H:%M:%S}.\n"
        f"MESES={meses}\n",
        encoding="utf-8",
    )
    print(f"Senal enviada: {nombre}  (MESES={meses})")
    print(f"  (en: {carpeta_senal})")


def latido_vm_edad_seg() -> float | None:
    """Segundos desde el ultimo latido de Vigilar-GAP-VM.bat, o None si el
    archivo no existe (VM con version vieja del script) o no se pudo leer."""
    try:
        d = json.loads(RUTA_LATIDO_VM.read_text(encoding="utf-8"))
        ultimo = datetime.fromisoformat(d["ultimo"])
        return (datetime.now() - ultimo).total_seconds()
    except Exception:
        return None


def _excels_actuales() -> set[Path]:
    if not CARPETA_ONEDRIVE_GAP.is_dir():
        return set()
    return {
        p for p in CARPETA_ONEDRIVE_GAP.iterdir()
        if p.is_file() and p.suffix.lower() in EXTENSIONES_EXCEL
    }


def _estable(p: Path, memoria: dict) -> bool:
    """True cuando el archivo dejo de crecer por SEGUNDOS_ESTABLE seguidos."""
    try:
        tam = p.stat().st_size
    except FileNotFoundError:
        return False
    prev_tam, estable_desde = memoria.get(p, (-1, None))
    if tam > 0 and tam == prev_tam:
        if estable_desde is None:
            estable_desde = time.time()
        elif time.time() - estable_desde >= SEGUNDOS_ESTABLE:
            return True
    else:
        estable_desde = None
    memoria[p] = (tam, estable_desde)
    return False


def esperar_excels_nuevos(vistos_antes: set, cuantos: int) -> list:
    """Espera a que aparezcan 'cuantos' Excel NUEVOS (que no estaban antes) y
    a que cada uno deje de crecer. Devuelve la lista de los que logro
    confirmar (puede ser menos de 'cuantos' si se agota el tiempo)."""
    print(
        f"\nEsperando {cuantos} Excel nuevo(s) en:\n  {CARPETA_ONEDRIVE_GAP}\n"
        f"(hasta {ESPERA_MAX_SEG // 60} minutos -- la VM tiene que detectar la "
        "senal, extraer cada ventana, y OneDrive sincronizar de vuelta...)"
    )
    fin = time.time() + ESPERA_MAX_SEG
    memoria = {}
    confirmados = set()
    vistos_alguna_vez = set()
    avisado_vm = False
    ultimo_chequeo_vm = time.time()  # primer chequeo recien a los 30s (dar margen a OneDrive)
    # (2026-09-15) BUG encontrado: el latido (_latido_vm.json) viaja por
    # OneDrive igual que los Excel de 3-4MB -- si OneDrive está ocupado
    # sincronizando esos archivos grandes justo en ese momento, puede atrasar
    # la sincronización del latido (chiquito) y parecer "viejo" aunque la VM
    # esté respondiendo y entregando todo bien (caso real: entregó los 3
    # Excel en <3 min y la corrida igual cortó por "no responde"). Ahora el
    # corte por latido viejo SOLO se dispara si, ADEMÁS, no llegó NINGÚN
    # archivo nuevo en ese mismo rato -- si hay archivos llegando, es
    # evidencia directa de que la VM está viva, sin importar qué diga el
    # latido (que puede estar retrasado por la propia sincronización).
    ultimo_progreso = time.time()
    while time.time() < fin:
        nuevos = _excels_actuales() - vistos_antes
        if nuevos - vistos_alguna_vez:
            ultimo_progreso = time.time()
        vistos_alguna_vez |= nuevos
        for p in nuevos:
            if p not in confirmados and _estable(p, memoria):
                confirmados.add(p)
                ultimo_progreso = time.time()
                print(f"  ({len(confirmados)}/{cuantos}) listo: {p.name}")
        if len(confirmados) >= cuantos:
            return sorted(confirmados, key=lambda p: p.stat().st_mtime)

        ahora = time.time()
        if ahora - ultimo_chequeo_vm >= 30:
            ultimo_chequeo_vm = ahora
            edad = latido_vm_edad_seg()
            sin_progreso_seg = ahora - ultimo_progreso
            if edad is not None and edad > UMBRAL_LATIDO_SEG and sin_progreso_seg > UMBRAL_LATIDO_SEG:
                if avisado_vm:
                    print(
                        f"\n  (!) La VM no responde hace {int(edad // 60)} min (según su último "
                        f"latido) Y tampoco llegó ningún archivo nuevo en {int(sin_progreso_seg // 60)} "
                        "min. Probablemente está apagada, la sesión de SDAPeru se cerró, o nadie dejó "
                        "corriendo Vigilar-GAP-VM.bat. Corto la espera acá para no perder los 20 "
                        "minutos completos a ciegas."
                    )
                    global VM_NO_RESPONDE
                    VM_NO_RESPONDE = True
                    return sorted(confirmados, key=lambda p: p.stat().st_mtime)
                avisado_vm = True
                print(
                    f"\n  (!) Aviso: el último latido de la VM tiene {int(edad // 60)} min y no hay "
                    "archivos nuevos. Reviso de nuevo en 30s antes de darla por caída."
                )
            else:
                avisado_vm = False
        time.sleep(INTERVALO_POLL_SEG)
    return sorted(confirmados, key=lambda p: p.stat().st_mtime)


def limpiar_destino(carpeta: Path) -> None:
    carpeta.mkdir(parents=True, exist_ok=True)
    for p in carpeta.iterdir():
        if p.is_file() and p.suffix.lower() in EXTENSIONES_EXCEL:
            try:
                p.unlink()
            except Exception as exc:
                print(f"  (aviso: no pude borrar {p.name}: {exc})")


def correr() -> None:
    if not CARPETA_ONEDRIVE_GAP.is_dir():
        print(f"ERROR: no encuentro la carpeta OneDrive 'GAP':\n  {CARPETA_ONEDRIVE_GAP}")
        print("Revisa que la ruta en CARPETA_ONEDRIVE_GAP (arriba en este archivo)")
        print("sea exactamente la misma que ves en el Explorador de Windows.")
        raise SystemExit(1)

    meses = _meses_pedidos()
    vistos_antes = _excels_actuales()
    mandar_senal(meses)
    excels = esperar_excels_nuevos(vistos_antes, meses)

    if len(excels) < meses:
        if VM_NO_RESPONDE:
            print(
                f"\nCORTADO TEMPRANO: la VM dejó de mandar latido (no respondió). "
                f"Llegaron {len(excels)} de {meses} Excel esperados.\n"
                "Revisa la VM: probablemente está apagada, se cerró la sesión de\n"
                "SDAPeru, o se cerró la consola de Vigilar-GAP-VM.bat. Una vez\n"
                "resuelto ahí, volvé a correr esto (o el paso del GAP en el panel)."
            )
        else:
            print(
                f"\nSe agoto el tiempo de espera: llegaron {len(excels)} de {meses} "
                "Excel esperados.\n"
                "Revisa que Vigilar-GAP-VM.bat estuviera corriendo en la VM (con\n"
                "SDAPeru ya logueado) y que la sesion de SDAPeru no haya vencido.\n"
                "La data de 3 meses tiene que estar COMPLETA -- no se copia nada\n"
                "parcial. Cuando esten los archivos, corre Mover-Excel-GAP-PC.bat."
            )
        raise SystemExit(1)

    print(f"\nLlegaron los {len(excels)} Excel. Limpiando CARGA\\EXCEL:\n  {CARPETA_DESTINO}")
    limpiar_destino(CARPETA_DESTINO)
    for excel in excels:
        destino = CARPETA_DESTINO / excel.name
        destino.write_bytes(excel.read_bytes())
        print(f"  copiado: {destino.name}")
    print("\nListo.")


if __name__ == "__main__":
    try:
        correr()
    except SystemExit as exc:
        if str(exc):
            print(f"\n{exc}")
        sys.exit(1)

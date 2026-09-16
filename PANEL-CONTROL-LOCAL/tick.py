# -*- coding: utf-8 -*-
"""
tick.py — el "reloj" de la programación automática (Fase 2).

Lo corre una Tarea de Windows cada 5 minutos (la crea Instalar-Panel.bat,
a través de tick_launcher.vbs que ESPERA a que este script termine).

En cada tick:
  1. Deja una marca de "el tick corrió" (tick_latido.json).
  2. Lee panel_config.json. Si la programación está desactivada -> nada.
  3. Pregunta a agenda.py si toca correr (o si toca un reintento).
  4. Si toca: corre ejecutar_todo.py ACÁ MISMO (síncrono), guardando el log
     y el resumen en la carpeta runs igual que el botón "Ejecutar ahora" —
     así aparece en el Historial. NO depende de que el panel esté abierto.

Uso manual (para probar):  python tick.py            (respeta el horario)
                           python tick.py --forzar   (corre igual, ya)
"""

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
import agenda  # noqa: E402
import panel   # noqa: E402  (solo para reutilizar helpers/rutas — no levanta servidor)

CONFIG_PATH = BASE / "panel_config.json"
AGENDA_ESTADO_PATH = BASE / "agenda_estado.json"
TICK_LATIDO_PATH = BASE / "tick_latido.json"
TICK_LOG = BASE / "tick.log"


def latido():
    try:
        TICK_LATIDO_PATH.write_text(
            json.dumps({"ultimo": datetime.now().isoformat(timespec="seconds")}),
            encoding="utf-8")
    except Exception:
        pass


def log(msg):
    linea = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    try:
        viejo = TICK_LOG.read_text(encoding="utf-8").splitlines() if TICK_LOG.exists() else []
        TICK_LOG.write_text("\n".join(viejo[-400:] + [linea]) + "\n", encoding="utf-8")
    except Exception:
        pass


def leer_json(p):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8-sig"))
    except (FileNotFoundError, ValueError):
        return {}


def set_estado(**kv):
    e = leer_json(AGENDA_ESTADO_PATH)
    e.update(kv)
    try:
        AGENDA_ESTADO_PATH.write_text(json.dumps(e, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def correr_pipeline(cfg, reintento):
    """Corre ejecutar_todo.py de forma síncrona. Escribe el log y el .json en
    la carpeta runs con el MISMO formato que usa el panel para el Historial.
    Devuelve (id, estado)."""
    rid = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = panel.RUNS_DIR / f"run_{rid}.log"
    meta_path = panel.RUNS_DIR / f"run_{rid}.json"
    marcador = panel.RUNS_DIR / "_corriendo.json"
    detener_path = panel.RUNS_DIR / "_detener.txt"
    args = panel.flags_ejecutar_todo(cfg)

    env = dict(os.environ)
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PANEL_SIN_VENTANA"] = "1"
    pyw = panel._pythonw() or sys.executable
    inicio = time.time()
    lineas = []
    lock = threading.Lock()
    for _p in (marcador, detener_path):
        try:
            _p.unlink()
        except Exception:
            pass
    try:
        marcador.write_text(json.dumps({
            "id": rid, "origen": "programada" + (" (reintento)" if reintento else ""),
            "inicio": datetime.fromtimestamp(inicio).isoformat(timespec="seconds"),
        }), encoding="utf-8")
    except Exception:
        pass

    estado = "error"
    detenido = False
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            def w(s):
                with lock:
                    lineas.append(s)
                    fh.write(s + "\n")
                    fh.flush()

            w("=" * 64)
            w(f"  Panel de Control — corrida automática ({'reintento' if reintento else 'programada'})  {rid}")
            w(f"  ejecutar_todo.py {' '.join(args)}")
            w("=" * 64)
            try:
                p = subprocess.Popen(
                    [pyw, str(panel.ORQUESTADOR), *args],
                    cwd=str(panel.EJECUTAR_TODO_DIR), env=env,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception as exc:
                w(f"ERROR al lanzar ejecutar_todo.py: {exc}")
            else:
                th = threading.Thread(
                    target=lambda: [w(x.rstrip("\n")) for x in p.stdout], daemon=True)
                th.start()
                fin_limite = time.time() + 45 * 60   # backstop
                while p.poll() is None:
                    time.sleep(2)
                    pedido = ""
                    try:
                        if detener_path.exists():
                            pedido = detener_path.read_text(encoding="utf-8").strip()
                    except Exception:
                        pass
                    if pedido == rid or time.time() > fin_limite:
                        detenido = (pedido == rid)
                        w("(tick: corrida DETENIDA a pedido)" if detenido
                          else "(tick: pasó de 45 min — corte de seguridad)")
                        try:
                            subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                                           capture_output=True)
                        except Exception:
                            pass
                        break
                th.join(timeout=10)
                try:
                    p.wait(timeout=15)
                except Exception:
                    pass
                estado = "detenido" if detenido else ("ok" if p.returncode == 0 else "error")
    finally:
        for _p in (marcador, detener_path):
            try:
                _p.unlink()
            except Exception:
                pass

    fin = time.time()
    paso = next((s.strip()[4:].strip() for s in reversed(lineas) if s.strip().startswith(">>> ")), None)
    meta = {
        "id": rid, "origen": "programada", "estado": estado,
        "inicio": datetime.fromtimestamp(inicio).isoformat(timespec="seconds"),
        "fin": datetime.fromtimestamp(fin).isoformat(timespec="seconds"),
        "dur_seg": round(fin - inicio),
        "reportes": cfg.get("reportes"), "meses": cfg.get("meses"),
        "prueba": bool(cfg.get("prueba")), "silencioso": bool(cfg.get("silencioso")),
        "titulo": "Corrida automática" + (" (reintento)" if reintento else ""), "paso": paso,
    }
    try:
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return rid, estado


def main():
    latido()
    forzar = "--forzar" in sys.argv
    cfg = leer_json(CONFIG_PATH)
    est = leer_json(AGENDA_ESTADO_PATH)

    reintento = False
    if forzar:
        slot_iso = datetime.now().isoformat(timespec="seconds")
    else:
        activa = bool((cfg.get("programacion") or {}).get("activa"))
        if not activa:
            return  # programación apagada -> ni corridas ni reintentos
        slot = agenda.debe_correr(cfg, est.get("ultima_ejecucion"))
        if slot is not None:
            slot_iso = slot.isoformat(timespec="seconds")
        else:
            rein = est.get("reintentar_en")
            try:
                if rein and datetime.now() >= datetime.fromisoformat(rein):
                    slot_iso = est.get("ultima_ejecucion") or datetime.now().isoformat(timespec="seconds")
                    reintento = True
                else:
                    return  # no toca nada
            except ValueError:
                return

    # Candado global: si el panel tiene una corrida manual (o una acción de
    # SAP/GAP) en curso, NO arrancamos otra encima — dos Chromium sobre el
    # mismo perfil de SAP se pisan. No anclamos el slot: el próximo tick (5
    # min) lo vuelve a intentar, y entra dentro de la tolerancia de 20 min.
    ok_lock, msg_lock = panel.tomar_lock("programada", titulo="Corrida automática")
    if not ok_lock:
        log(f"no arranco todavía: {msg_lock}")
        return

    # Anclar el slot ANTES de correr: si el proceso se corta a la mitad, el
    # próximo tick no lo vuelve a lanzar en bucle.
    kv = {"ultima_ejecucion": slot_iso, "reintentar_en": None}
    if not reintento:
        kv["reintentos"] = 0
    set_estado(**kv)

    etiqueta = "REINTENTO" if reintento else f"slot {slot_iso}"
    log(f"toca correr ({etiqueta}) — corriendo ejecutar_todo…")
    try:
        rid, estado = correr_pipeline(cfg, reintento)
    except Exception as exc:
        log(f"ERROR corriendo el pipeline: {exc}")
        return
    finally:
        panel.soltar_lock()
    log(f"corrida {rid}: {estado}")

    if estado == "error":
        reintentos = int(leer_json(AGENDA_ESTADO_PATH).get("reintentos", 0))
        if reintentos < 1:
            cuando = (datetime.now() + timedelta(seconds=600)).isoformat(timespec="seconds")
            set_estado(reintentar_en=cuando, reintentos=reintentos + 1)
            log("se reintenta en ~10 min")
        else:
            set_estado(reintentar_en=None, reintentos=0)
        try:
            panel.notificar_windows("Reporte AP — falló la corrida automática",
                                    "Abrí el Panel de Control y mirá el Terminal / Historial.")
        except Exception:
            pass
    else:
        set_estado(reintentar_en=None, reintentos=0)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
panel.py — Panel de Control del Reporte AP.

Servidor web local (solo librería estándar de Python) que reemplaza a los
.bat: desde el navegador elegís qué reportes correr, cuántos meses, lo lanzás
con un botón, ves el terminal en vivo y el historial de corridas.

    Fase 1 (LISTO):  Ejecutar ahora + terminal en vivo + historial.
    Fase 2 (LISTO):  Programación automática (tick.py + Tarea de Windows).
    Fase 3 (LISTO):  Modo silencioso del SAP (navegador fuera de pantalla).
    Fase 4:          Acciones sueltas + panel de salud.

Uso:  doble clic en Panel-AP.bat   (o:  python panel.py)
Abre:  http://localhost:8765
"""

import json
import os
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import agenda

BASE = Path(__file__).resolve().parent
RAIZ = BASE.parent                       # "1. REPORTE"
EJECUTAR_TODO_DIR = RAIZ / "EJECUTAR-TODO"
ORQUESTADOR = EJECUTAR_TODO_DIR / "ejecutar_todo.py"
SAP_DIR = RAIZ / "RPA-SAP-LOCAL"
GAP_PC_DIR = RAIZ / "RPA-GAP-LOCAL" / "PARA-TU-PC-REAL"
RUTA_GAP_TXT = GAP_PC_DIR / "ruta_carpeta_gap.txt"
ATENDIDAS_DIR = RAIZ / "ATENDIDAS-LOCAL"
PUBLICAR_WEB_DIR = RAIZ / "PUBLICAR-WEB-LOCAL"
CARGA_SAP = RAIZ / "CARGA" / "SAP"
CARGA_GAP = RAIZ / "CARGA" / "EXCEL"
RUNS_DIR = BASE / "runs"
CONFIG_PATH = BASE / "panel_config.json"
AGENDA_ESTADO_PATH = BASE / "agenda_estado.json"   # { "ultima_ejecucion", "reintentar_en", "reintentos" }
TICK_LATIDO_PATH = BASE / "tick_latido.json"       # { "ultimo": iso } — lo escribe tick.py en cada corrida
UI_PATH = BASE / "ui.html"
PUERTO = 8765
TICK_TASK = "PanelAP_Tick"                          # nombre de la Tarea de Windows

TIMEOUT_CORRIDA_SEG = 45 * 60      # backstop: si una corrida pasa de esto, se corta
RETENCION_RUNS_DIAS = 45           # se borran los logs de corridas más viejas
REINTENTO_PROGRAMADA_SEG = 600     # 10 min tras un fallo automático, entre reintento y reintento
# (2026-09-20) Antes era un UNICO reintento -- el bloqueo de red de la
# empresa a veces se corta por varios minutos seguidos (no un solo
# instante), asi que un solo reintento no siempre alcanzaba para que la
# conexion ya estuviera libre otra vez. Se suben a 3 reintentos (o sea,
# hasta 4 intentos en total contando el primero), separados 10 min cada uno.
MAX_REINTENTOS_PROGRAMADA = 3
TICK_ATRASO_SEG = 2 * 3600         # si el tick no corrió en este tiempo, avisar

RUNS_DIR.mkdir(exist_ok=True)

DEFAULTS = {
    "reportes": ["pendientes", "atendidas", "veredas"],
    "meses": 3,
    "prueba": False,
    "silencioso": False,
    # Fase 2 (todavía no se usa) — se deja escrito para no migrar después.
    "programacion": {
        "activa": False,
        "intervalo_horas": 1,
        "minuto": 50,
        "hora_desde": 0,
        "hora_hasta": 24,
    },
}


# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
def cargar_config():
    cfg = json.loads(json.dumps(DEFAULTS))  # copia profunda
    try:
        guardado = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        for k, v in guardado.items():
            if k == "programacion" and isinstance(v, dict):
                cfg["programacion"].update(v)
            else:
                cfg[k] = v
    except (FileNotFoundError, ValueError):
        pass
    return cfg


def guardar_config(cfg):
    limpio = {}
    for k in DEFAULTS:
        limpio[k] = cfg.get(k, DEFAULTS[k])
    CONFIG_PATH.write_text(json.dumps(limpio, ensure_ascii=False, indent=2), encoding="utf-8")
    return limpio


def agenda_estado():
    try:
        return json.loads(AGENDA_ESTADO_PATH.read_text(encoding="utf-8-sig"))
    except (FileNotFoundError, ValueError):
        return {}


def set_agenda_estado(**kv):
    e = agenda_estado()
    e.update(kv)
    AGENDA_ESTADO_PATH.write_text(json.dumps(e, ensure_ascii=False, indent=2), encoding="utf-8")
    return e


def flags_ejecutar_todo(cfg):
    """Flags para ejecutar_todo.py según la config. Lo usan la corrida del
    botón (Corrida._args) y la corrida programada (tick.py)."""
    args = []
    reps = cfg.get("reportes") or DEFAULTS["reportes"]
    if set(reps) != set(DEFAULTS["reportes"]):
        args += ["--reportes", ",".join(reps)]
    args += ["--meses", str(int(cfg.get("meses", 3)))]
    if cfg.get("silencioso"):
        args.append("--silencioso")
    if cfg.get("prueba"):
        args.append("--sin-publicar")
    return args


# (2026-09-23) A pedido de la usuaria: mostrar en el panel cuántos meses de
# GAP hacen falta de verdad para no perder confirmación de ningún pendiente
# viejo. IMPORTANTE (corregido en vivo con la usuaria): NO alcanza con mirar
# solo los "LEGAL" -- un NO LEGAL viejo que ya no entra en el rango del GAP
# NO se pierde de la página, pero queda marcado "_sinConfirmar" PARA SIEMPRE
# (ver historico_pendientes.mezclar): nunca más se puede confirmar si de
# verdad sigue abierto o ya se cerró. Por eso la métrica correcta es la edad
# del caso MÁS VIEJO que la corrida actual todavía puede confirmar bien
# (LEGAL, que siempre se reconfirma; o NO LEGAL con _sinConfirmar=False) --
# ese es el mínimo de meses de GAP para no empezar a acumular "sin confirmar".
def meses_gap_sugeridos():
    ruta = RAIZ / "PENDIENTES-LOCAL" / "SALIDA" / "bd_completa.json"
    try:
        rows = json.loads(ruta.read_text(encoding="utf-8"))["rows"]
    except Exception:
        return {"disponible": False}

    def _parse(s):
        try:
            return datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S")
        except Exception:
            return None

    activos = [r for r in rows if r.get("LEGAL") == "SI" or not r.get("_sinConfirmar")]
    fechas = [f for f in (_parse(r.get("SAP - Fecha de registro")) for r in activos) if f]
    if not fechas:
        return {"disponible": False}

    dias = (datetime.utcnow() - min(fechas)).days
    meses = max(1, -(-dias // 30))  # ceil sin importar math
    sin_confirmar = sum(1 for r in rows if r.get("_sinConfirmar"))
    return {
        "disponible": True,
        "dias_caso_mas_viejo": dias,
        "meses_sugeridos": meses,
        "casos_sin_confirmar": sin_confirmar,
    }


def tick_latido_iso():
    try:
        return json.loads(TICK_LATIDO_PATH.read_text(encoding="utf-8-sig")).get("ultimo")
    except (FileNotFoundError, ValueError):
        return None


_ultima_limpieza = [0.0]


def limpiar_runs():
    """Borra logs de corridas más viejas que RETENCION_RUNS_DIAS. Como mucho
    una vez por hora."""
    if time.time() - _ultima_limpieza[0] < 3600:
        return
    _ultima_limpieza[0] = time.time()
    corte = time.time() - RETENCION_RUNS_DIAS * 86400
    for p in RUNS_DIR.glob("run_*"):
        try:
            if p.stat().st_mtime < corte:
                p.unlink()
        except Exception:
            pass


# Windows: correr procesos de consola SIN que parpadee una ventana negra.
_NO_WIN = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _si_oculto():
    if os.name != "nt":
        return None
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0  # SW_HIDE
    return si


def _pythonw():
    """pythonw.exe real (mismo dir que el intérprete actual), o None."""
    p = Path(sys.executable).with_name("pythonw.exe")
    return str(p) if p.exists() else None


def _run_oculto(cmd, **kw):
    return subprocess.run(cmd, creationflags=_NO_WIN, startupinfo=_si_oculto(),
                          capture_output=True, text=True, **kw)


# ---------------------------------------------------------------------------
# Candado global de corrida  (runs/_lock.json)
# ---------------------------------------------------------------------------
# UNA sola corrida "pesada" a la vez, sin importar quién la lanzó: el botón
# "Ejecutar ahora", una acción de SAP/GAP, o el tick automático. Antes se
# podían pisar (dos Chromium sobre el mismo perfil de SAP -> el navegador se
# caía a mitad, "Target page has been closed"). El candado es un archivito
# que las dos vías (panel.py y tick.py) crean antes de arrancar y borran al
# terminar. Si el que lo tiene ya murió (pid muerto) o pasó de una hora, se
# considera vencido y se puede pisar.
LOCK_PATH = RUNS_DIR / "_lock.json"
LOCK_VENCE_SEG = 65 * 60  # más que el timeout duro de una corrida (45 min)

# (2026-09-11) Candado de INSTANCIA UNICA del panel en sí (distinto del
# candado de corridas de arriba). Antes, "una sola instancia" dependía de
# que ThreadingHTTPServer no pudiera bindear el puerto si ya había otro
# panel escuchando -- pero en Windows, HTTPServer nace con
# allow_reuse_address=True, y con eso un SEGUNDO proceso puede bindear el
# MISMO puerto sin que salte OSError. Resultado real visto en producción:
# 3 procesos panel.py vivos a la vez, cada uno con su propio estado en
# memoria (ACTUAL, config cacheada, etc.) -- cuál de los 3 responde en un
# momento dado es impredecible, así que la ventana podía mostrar info
# vieja de un proceso "fantasma" que nunca se enteró de nada nuevo. Este
# archivo de PID (con el mismo criterio _pid_vivo() que ya usa el candado
# de corridas) es la fuente de verdad real: si el PID adentro sigue vivo,
# YA hay un panel; si no, se pisa y sigue.
PANEL_PID_PATH = BASE / "_panel.pid"


def _otra_instancia_del_panel_viva():
    try:
        pid = int(PANEL_PID_PATH.read_text(encoding="utf-8-sig").strip())
    except (FileNotFoundError, ValueError):
        return False
    if pid == os.getpid():
        return False
    return _pid_vivo(pid)


def _tomar_pid_panel():
    try:
        PANEL_PID_PATH.write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        pass


def _soltar_pid_panel():
    try:
        if int(PANEL_PID_PATH.read_text(encoding="utf-8-sig").strip()) == os.getpid():
            PANEL_PID_PATH.unlink()
    except Exception:
        pass


def _pid_vivo(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if os.name == "nt":
        try:
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if h:
                ctypes.windll.kernel32.CloseHandle(h)
                return True
            return False
        except Exception:
            return True  # ante la duda, asumir vivo (no pisar)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except Exception:
        return True


def lock_info():
    """dict del candado si hay uno VIGENTE; None si no hay o está vencido
    (y en ese caso borra el archivo de paso)."""
    try:
        d = json.loads(LOCK_PATH.read_text(encoding="utf-8-sig"))
    except (FileNotFoundError, ValueError):
        return None
    try:
        edad = time.time() - float(d.get("ts", 0))
    except (TypeError, ValueError):
        edad = 0
    if edad > LOCK_VENCE_SEG or not _pid_vivo(d.get("pid")):
        try:
            LOCK_PATH.unlink()
        except Exception:
            pass
        return None
    return d


def tomar_lock(origen, titulo=None):
    """Intenta crear el candado. Devuelve (True, None) o (False, mensaje)."""
    dueno = lock_info()
    if dueno is not None:
        desde = (dueno.get("inicio") or "")[11:16]
        que = dueno.get("titulo") or dueno.get("origen") or "otra corrida"
        return False, (f"Ya hay una corrida en curso ({que}"
                       + (f", desde las {desde}" if desde else "") + "). "
                       "Esperá a que termine.")
    try:
        RUNS_DIR.mkdir(parents=True, exist_ok=True)
        LOCK_PATH.write_text(json.dumps({
            "pid": os.getpid(), "origen": origen, "titulo": titulo,
            "inicio": datetime.now().isoformat(timespec="seconds"),
            "ts": time.time(),
        }, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        _log(f"no se pudo escribir el candado: {exc}")
        return True, None  # que no bloquee por un problema de disco
    return True, None


def soltar_lock():
    try:
        LOCK_PATH.unlink()
    except FileNotFoundError:
        pass
    except Exception as exc:
        _log(f"no se pudo borrar el candado: {exc}")


_tick_cache = {"t": 0.0, "v": False}


def tick_instalado():
    if os.name != "nt":
        return False
    ahora = time.time()
    if ahora - _tick_cache["t"] < 30:
        return _tick_cache["v"]
    try:
        v = _run_oculto(["schtasks", "/query", "/tn", TICK_TASK]).returncode == 0
    except Exception:
        v = False
    _tick_cache.update(t=ahora, v=v)
    return v


def notificar_windows(titulo, texto):
    """Globo de notificación de Windows (sin módulos extra)."""
    if os.name != "nt":
        return
    ps = (
        "$ErrorActionPreference='SilentlyContinue';"
        "Add-Type -AssemblyName System.Windows.Forms,System.Drawing;"
        "$n=New-Object System.Windows.Forms.NotifyIcon;"
        "$n.Icon=[System.Drawing.SystemIcons]::Warning;$n.Visible=$true;"
        f"$n.ShowBalloonTip(15000,'{titulo}','{texto}',"
        "[System.Windows.Forms.ToolTipIcon]::Error);"
        "Start-Sleep -Seconds 16;$n.Dispose()"
    )
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                         creationflags=_NO_WIN, startupinfo=_si_oculto())
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Corrida (una sola a la vez)
# ---------------------------------------------------------------------------
class Corrida:
    def __init__(self, cfg, origen, comando=None, cwd=None, titulo=None, bloquea=False):
        self.id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.cfg = cfg
        self.origen = origen                       # "manual" | "programada" | "accion"
        self.bloquea = bloquea                     # True => tomó el candado global
        self.comando = comando                     # None => corrida normal (ejecutar_todo)
        self.cwd = Path(cwd) if cwd else EJECUTAR_TODO_DIR
        self.titulo = titulo
        self.log_path = RUNS_DIR / f"run_{self.id}.log"
        self.meta_path = RUNS_DIR / f"run_{self.id}.json"
        self.proc = None
        self.pid = None
        self.inicio = time.time()
        self.fin = None
        self.estado = "corriendo"                  # corriendo | ok | error | detenido
        self.lineas = []
        self._lock = threading.Lock()
        self._timer = None
        self._fh = open(self.log_path, "w", encoding="utf-8")

    # -- construcción del comando -------------------------------------------
    def _args(self):
        return flags_ejecutar_todo(self.cfg)

    def _escribir(self, texto):
        with self._lock:
            self.lineas.append(texto)
            self._fh.write(texto + "\n")
            self._fh.flush()

    def _paso_actual(self):
        for t in reversed(self.lineas[-400:]):
            s = t.strip()
            if s.startswith(">>> "):
                return s[4:].strip()
        return None

    # -- ejecución ---------------------------------------------------------
    def arrancar(self):
        cmd = self.comando or ([str(ORQUESTADOR)] + self._args())
        self._escribir("=" * 64)
        etiqueta = self.titulo or "corrida completa"
        self._escribir(f"  Panel de Control — {etiqueta}  ({self.origen})  {self.id}")
        self._escribir(f"  {' '.join(str(c) for c in cmd)}")
        self._escribir(f"  (en {self.cwd})")
        self._escribir("=" * 64)

        env = dict(os.environ)
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"

        creationflags = 0
        if os.name == "nt":
            # Grupo propio => se puede matar todo el árbol; sin ventana negra.
            creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | _NO_WIN
            # el proceso y TODOS sus hijos (SAP, GAP, feed, procesar) heredan
            # esto y no abren consola.  El navegador del SAP es aparte (Fase 3).
            env["PANEL_SIN_VENTANA"] = "1"

        pyw = _pythonw() or sys.executable
        try:
            self.proc = subprocess.Popen(
                [pyw, *[str(c) for c in cmd]],
                cwd=str(self.cwd), env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", bufsize=1,
                creationflags=creationflags, startupinfo=_si_oculto(),
            )
        except Exception as exc:
            self._escribir(f"ERROR al lanzar el proceso: {exc}")
            self._terminar("error")
            return
        self.pid = self.proc.pid
        self._timer = threading.Timer(TIMEOUT_CORRIDA_SEG, self._timeout)
        self._timer.daemon = True
        self._timer.start()
        threading.Thread(target=self._leer, daemon=True).start()

    def _leer(self):
        try:
            for linea in self.proc.stdout:
                self._escribir(linea.rstrip("\n"))
        except Exception as exc:
            self._escribir(f"(panel: se cortó la lectura del log: {exc})")
        self.proc.wait()
        if self.estado == "detenido":
            self._terminar("detenido")
        elif self.estado == "corriendo":
            self._terminar("ok" if self.proc.returncode == 0 else "error")

    def _matar_proc(self):
        if self.proc and self.proc.poll() is None:
            try:
                if os.name == "nt":
                    _run_oculto(["taskkill", "/F", "/T", "/PID", str(self.pid)])
                else:
                    self.proc.terminate()
            except Exception as exc:
                self._escribir(f"(panel: no se pudo matar el proceso: {exc})")

    def _timeout(self):
        if self.estado != "corriendo":
            return
        self._escribir(f"(panel: la corrida pasó de {TIMEOUT_CORRIDA_SEG // 60} min — "
                       f"se corta por seguridad. Revisá qué la dejó colgada.)")
        self._matar_proc()
        self._terminar("error")

    def _terminar(self, estado):
        if self.fin is not None:
            return
        self.estado = estado
        self.fin = time.time()
        if self._timer:
            self._timer.cancel()
        if self.bloquea:
            soltar_lock()

        # Reintento único de la corrida automática tras un fallo.
        if self.origen == "programada":
            est = agenda_estado()
            if estado == "ok":
                set_agenda_estado(reintentar_en=None, reintentos=0)
            elif estado == "error":
                if int(est.get("reintentos", 0)) < MAX_REINTENTOS_PROGRAMADA:
                    cuando = datetime.now() + timedelta(seconds=REINTENTO_PROGRAMADA_SEG)
                    set_agenda_estado(reintentar_en=cuando.isoformat(timespec="seconds"),
                                      reintentos=int(est.get("reintentos", 0)) + 1)
                    self._escribir(f"(panel: se reintenta la corrida automática en "
                                   f"~{REINTENTO_PROGRAMADA_SEG // 60} min)")
                else:
                    set_agenda_estado(reintentar_en=None, reintentos=0)

        try:
            self._fh.close()
        except Exception:
            pass
        self.meta_path.write_text(json.dumps(self.meta(), ensure_ascii=False, indent=2), encoding="utf-8")

        if estado == "error":
            quien = {"programada": "automática", "accion": "acción"}.get(self.origen, "manual")
            notificar_windows("Reporte AP — falló la corrida",
                              f"{self.titulo or ('Corrida ' + quien)}: abrí el Panel y mirá el Terminal.")

    def detener(self):
        self.estado = "detenido"
        self._matar_proc()

    # -- vistas ----------------------------------------------------------
    def meta(self):
        dur = (self.fin or time.time()) - self.inicio
        return {
            "id": self.id,
            "origen": self.origen,
            "estado": self.estado,
            "inicio": datetime.fromtimestamp(self.inicio).isoformat(timespec="seconds"),
            "fin": datetime.fromtimestamp(self.fin).isoformat(timespec="seconds") if self.fin else None,
            "dur_seg": round(dur),
            "reportes": self.cfg.get("reportes"),
            "meses": self.cfg.get("meses"),
            "prueba": bool(self.cfg.get("prueba")),
            "silencioso": bool(self.cfg.get("silencioso")),
            "titulo": self.titulo or ("Corrida: " + ", ".join(self.cfg.get("reportes") or [])),
            "paso": self._paso_actual(),
        }


ACTUAL = None            # type: Corrida | None
ARRANQUE_LOCK = threading.Lock()
ULTIMO_PING = time.time()   # último /api/estado — si el navegador se cierra, deja de llegar


def iniciar_corrida(cfg, origen="manual", comando=None, cwd=None, titulo=None, bloquea=False):
    global ACTUAL
    with ARRANQUE_LOCK:
        if ACTUAL is not None and ACTUAL.estado == "corriendo":
            return None, "Ya hay una corrida en curso."
        if bloquea:
            ok, msg = tomar_lock(origen, titulo=titulo or "corrida completa")
            if not ok:
                return None, msg
        c = Corrida(cfg, origen, comando=comando, cwd=cwd, titulo=titulo, bloquea=bloquea)
        ACTUAL = c
    c.arrancar()
    return c, None


# ---------------------------------------------------------------------------
# Acciones sueltas (lo que antes eran .bat aparte)
# ---------------------------------------------------------------------------
ACCIONES = {
    "actualizar-codigo": {
        "titulo": "Actualizar código del bot",
        "cwd": RAIZ, "cmd": [str(RAIZ / "actualizar_codigo.py")]},
    "publicar-sitio": {
        "titulo": "Publicar solo el sitio (código)",
        "cwd": PUBLICAR_WEB_DIR, "cmd": ["publicar_sitio.py"]},
    "descargar-sap": {
        # (2026-09-23) Fijo en 1 mes + bloque de ODMs -- ver MESES_SAP en
        # EJECUTAR-TODO/ejecutar_todo.py.
        "titulo": "Descargar SAP (1 mes + bloque ODMs)",
        "cwd": SAP_DIR, "cmd": ["descargar_excel_sap.py", "--meses", "1"],
        "silencio": True, "bloquea": True},
    "descargar-gap": {
        "titulo": "Traer GAP de la VM (3 meses)",
        "cwd": GAP_PC_DIR, "cmd": ["iniciar_y_esperar_gap_pc.py", "--meses", "3"],
        "bloquea": True},
    "importar-historico": {
        "titulo": "Importar mes(es) al histórico (Atendidas)",
        "cwd": ATENDIDAS_DIR, "cmd": ["importar_mes_historico.py"]},
    "publicar-historico": {
        "titulo": "Publicar histórico (Atendidas)",
        "cwd": ATENDIDAS_DIR, "cmd": ["publicar_historico.py"]},
    "login-sap": {
        "titulo": "Iniciar sesión en SAP (ventana visible)",
        "cwd": SAP_DIR, "cmd": ["descargar_excel_sap.py", "--solo-login"], "bloquea": True},
}


def iniciar_accion(clave):
    a = ACCIONES.get(clave)
    if not a:
        return None, "Acción desconocida."
    cmd = list(a["cmd"])
    if a.get("silencio") and cargar_config().get("silencioso"):
        cmd.append("--silencioso")
    return iniciar_corrida(cargar_config(), "accion", comando=cmd,
                           cwd=a["cwd"], titulo=a["titulo"], bloquea=a.get("bloquea", False))


# ---------------------------------------------------------------------------
# Historial
# ---------------------------------------------------------------------------
def historial(limite=40):
    metas = []
    for p in sorted(RUNS_DIR.glob("run_*.json"), reverse=True)[:limite]:
        try:
            metas.append(json.loads(p.read_text(encoding="utf-8")))
        except ValueError:
            pass
    # Si la corrida actual TODAVIA ESTA CORRIENDO y aun no escribio su meta,
    # meterla arriba. (2026-09-15) BUG: antes se metia igual aunque ACTUAL ya
    # hubiera TERMINADO hace rato -- si después corrieron corridas de tick.py
    # (proceso aparte) mas nuevas, esta linea las tapaba igual y el historial
    # mostraba una corrida vieja primero. Mismo criterio que estado_general()
    # y leer_log(): ACTUAL solo manda mientras esta "corriendo" de verdad.
    if ACTUAL is not None and ACTUAL.estado == "corriendo" and (
            not metas or metas[0].get("id") != ACTUAL.id):
        metas.insert(0, ACTUAL.meta())
    # Corrida automática (tick.py) en curso — todavía no tiene .json.
    tk = _tick_corriendo()
    if tk and (not metas or metas[0].get("id") != tk["id"]):
        metas.insert(0, {"id": tk["id"], "origen": tk.get("origen", "programada"),
                         "estado": "corriendo", "inicio": tk.get("inicio"), "fin": None,
                         "dur_seg": None, "titulo": "Corrida automática", "paso": tk.get("paso")})
    return metas


def _tick_corriendo():
    """dict del marcador si hay una corrida automática (tick.py) en curso."""
    p = RUNS_DIR / "_corriendo.json"
    try:
        d = json.loads(p.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return None
    # el log crece mientras corre -> sacar el paso actual de ahí
    lp = RUNS_DIR / f"run_{d.get('id')}.log"
    if lp.exists():
        try:
            for ln in reversed(lp.read_text(encoding="utf-8", errors="replace").splitlines()[-300:]):
                if ln.strip().startswith(">>> "):
                    d["paso"] = ln.strip()[4:].strip()
                    break
        except Exception:
            pass
    return d


def leer_log(run_id, desde=0):
    # Mismo criterio que estado_general(): si no se pidió un run_id puntual,
    # ACTUAL solo manda mientras esté EN CURSO. Si ya terminó y hay una
    # corrida automática (tick.py, proceso aparte) corriendo, esa gana --
    # si no, el panel se quedaba mostrando el log de la corrida manual vieja
    # mientras la programada corría en segundo plano.
    usar_actual = ACTUAL is not None and (
        run_id == ACTUAL.id or (run_id in (None, "") and ACTUAL.estado == "corriendo")
    )
    if usar_actual:
        with ACTUAL._lock:
            total = len(ACTUAL.lineas)
            nuevas = ACTUAL.lineas[desde:]
        return {"id": ACTUAL.id, "lineas": nuevas, "total": total,
                "estado": ACTUAL.estado, "corriendo": ACTUAL.estado == "corriendo"}
    # corrida pasada (o automática del tick en curso) -> del archivo
    tk = _tick_corriendo()
    if not run_id and tk:
        run_id = tk["id"]
    if not run_id:
        metas = historial(1)
        if not metas:
            return {"id": None, "lineas": [], "total": 0, "estado": "—", "corriendo": False}
        run_id = metas[0]["id"]
    p = RUNS_DIR / f"run_{run_id}.log"
    if not p.exists():
        return {"id": run_id, "lineas": [], "total": 0, "estado": "—", "corriendo": False}
    todas = p.read_text(encoding="utf-8", errors="replace").splitlines()
    corriendo_tk = bool(tk and tk["id"] == run_id)
    estado = "corriendo" if corriendo_tk else "—"
    mp = RUNS_DIR / f"run_{run_id}.json"
    if mp.exists() and not corriendo_tk:
        try:
            estado = json.loads(mp.read_text(encoding="utf-8")).get("estado", "—")
        except ValueError:
            pass
    return {"id": run_id, "lineas": todas[desde:], "total": len(todas),
            "estado": estado, "corriendo": corriendo_tk}


def estado_general():
    cfg = cargar_config()
    tk = _tick_corriendo()
    # Prioridad: 1) corrida EN CURSO de este proceso (manual/API), 2) corrida
    # EN CURSO de tick.py (proceso aparte -- se entera por _corriendo.json),
    # 3) recién si no hay nada corriendo AHORA, la última que quedó en ACTUAL.
    # Antes, si ACTUAL tenía una corrida manual ya TERMINADA, tapaba para
    # siempre a una corrida automática que arrancara después (el panel se
    # quedaba mostrando la corrida pasada mientras la programada corría).
    if ACTUAL is not None and ACTUAL.estado == "corriendo":
        corr = ACTUAL.meta()
    elif tk:
        corr = {"id": tk["id"], "origen": tk.get("origen", "programada"), "estado": "corriendo",
                "inicio": tk.get("inicio"), "titulo": "Corrida automática", "paso": tk.get("paso")}
    elif ACTUAL is not None:
        corr = ACTUAL.meta()
    else:
        corr = None
    ultima = None
    for m in historial(8):
        if m.get("fin"):
            ultima = m
            break
    prox = agenda.proxima_corrida(cfg)
    est = agenda_estado()
    activa = bool(cfg.get("programacion", {}).get("activa"))
    instalado = tick_instalado()
    latido = tick_latido_iso()
    atrasado = False
    if activa and instalado and latido:
        try:
            atrasado = (datetime.now() - datetime.fromisoformat(latido)).total_seconds() > TICK_ATRASO_SEG
        except ValueError:
            pass
    return {"config": cfg, "corrida": corr,
            "corriendo": (ACTUAL is not None and ACTUAL.estado == "corriendo") or (tk is not None),
            "ultima": ultima,
            "programacion": {
                **cfg.get("programacion", {}),
                "proxima": prox.isoformat(timespec="minutes") if prox else None,
                "tick_instalado": instalado,
                "ultima_ejecucion": est.get("ultima_ejecucion"),
                "reintentar_en": est.get("reintentar_en"),
                "tick_latido": latido,
                "tick_atrasado": atrasado,
            }}


# ---------------------------------------------------------------------------
# Panel de salud
# ---------------------------------------------------------------------------
_EXCEL_EXT = (".xls", ".xlsx", ".xlsm", ".mhtml", ".mht")
_salud_cache = {"t": 0.0, "v": {}}


def _excel_info(carpeta):
    try:
        fs = [p for p in carpeta.glob("*") if p.suffix.lower() in _EXCEL_EXT and p.is_file()]
    except Exception:
        fs = []
    if not fs:
        return {"n": 0, "mas_nuevo": None, "horas": None}
    mt = max(p.stat().st_mtime for p in fs)
    return {"n": len(fs),
            "mas_nuevo": datetime.fromtimestamp(mt).isoformat(timespec="minutes"),
            "horas": round((time.time() - mt) / 3600, 1)}


def _publicacion_una(carpeta_local):
    cfg = json.loads((carpeta_local / "config.json").read_text(encoding="utf-8"))
    base = cfg["path"].rsplit("/", 1)[0]
    url = cfg["script_url"] + "?" + urllib.parse.urlencode({"path": base + "/latest.json"})
    with urllib.request.urlopen(url, timeout=6) as r:
        d = json.loads(r.read().decode("utf-8"))
    return d.get("updated_at")


def _publicaciones():
    out = {}
    for nombre, carpeta in (("pendientes", RAIZ / "PENDIENTES-LOCAL"),
                            ("atendidas", RAIZ / "ATENDIDAS-LOCAL"),
                            ("veredas", RAIZ / "VEREDAS-LOCAL")):
        try:
            out[nombre] = _publicacion_una(carpeta)
        except Exception:
            out[nombre] = None
    return out


def salud():
    ahora = time.time()
    if ahora - _salud_cache["t"] < 90:
        return _salud_cache["v"]
    v = {"excel_sap": _excel_info(CARGA_SAP),
         "excel_gap": _excel_info(CARGA_GAP),
         "publicaciones": _publicaciones()}
    _salud_cache.update(t=ahora, v=v)
    return v


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        cuerpo = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(cuerpo)

    def _leer_cuerpo(self):
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except ValueError:
            return {}

    # -- GET -------------------------------------------------------------
    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        ruta = u.path

        if ruta in ("/", "/index.html"):
            try:
                html = UI_PATH.read_text(encoding="utf-8")
            except FileNotFoundError:
                html = "<h1>Falta ui.html</h1>"
            cuerpo = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(cuerpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(cuerpo)
            return

        if ruta == "/api/estado":
            global ULTIMO_PING
            ULTIMO_PING = time.time()
            return self._json(estado_general())

        if ruta == "/api/config":
            return self._json(cargar_config())

        if ruta == "/api/meses-gap-sugeridos":
            return self._json(meses_gap_sugeridos())

        if ruta == "/api/gap-ruta":
            actual = ""
            if RUTA_GAP_TXT.exists():
                try:
                    actual = RUTA_GAP_TXT.read_text(encoding="utf-8").strip()
                except OSError:
                    pass
            return self._json({"ruta": actual})

        if ruta == "/api/historial":
            return self._json({"corridas": historial()})

        if ruta == "/api/salud":
            return self._json(salud())

        if ruta == "/api/log":
            run_id = (q.get("id") or [None])[0]
            desde = int((q.get("desde") or ["0"])[0] or 0)
            return self._json(leer_log(run_id, desde))

        if ruta == "/api/log-crudo":
            run_id = (q.get("id") or [None])[0]
            data = leer_log(run_id, 0)
            texto = "\n".join(data["lineas"]).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Disposition",
                             f'attachment; filename="corrida_{data.get("id") or "log"}.txt"')
            self.send_header("Content-Length", str(len(texto)))
            self.end_headers()
            self.wfile.write(texto)
            return

        self._json({"error": "ruta no encontrada"}, 404)

    # -- POST ----------------------------------------------------------
    def do_POST(self):
        ruta = urlparse(self.path).path
        cuerpo = self._leer_cuerpo()

        if ruta == "/api/config":
            cfg = cargar_config()
            activa_antes = bool(cfg.get("programacion", {}).get("activa"))
            for k in ("reportes", "meses", "prueba", "silencioso"):
                if k in cuerpo:
                    cfg[k] = cuerpo[k]
            if isinstance(cuerpo.get("programacion"), dict):
                cfg["programacion"].update(cuerpo["programacion"])
            guardado = guardar_config(cfg)
            activa_ahora = bool(guardado.get("programacion", {}).get("activa"))
            if not activa_antes and activa_ahora:
                # Al ACTIVAR: anclar 'ahora' para no disparar un slot viejo del día.
                set_agenda_estado(ultima_ejecucion=datetime.now().isoformat(timespec="seconds"),
                                  reintentar_en=None, reintentos=0)
            elif activa_antes and not activa_ahora:
                # Al DESACTIVAR: cancelar cualquier reintento pendiente.
                set_agenda_estado(reintentar_en=None, reintentos=0)
            return self._json(guardado)

        if ruta == "/api/ejecutar":
            cfg = cargar_config()
            for k in ("reportes", "meses", "prueba", "silencioso"):
                if k in cuerpo:
                    cfg[k] = cuerpo[k]
            guardar_config(cfg)
            if not cfg.get("reportes"):
                return self._json({"error": "Elegí al menos un reporte."}, 400)
            c, err = iniciar_corrida(cfg, "manual", bloquea=True)
            if err:
                return self._json({"error": err}, 409)
            return self._json({"ok": True, "id": c.id})

        if ruta == "/api/gap-ruta":
            # 2026-09-16: para no tener que crear el .txt a mano -- se pega
            # la ruta acá y el panel escribe el mismo archivo que ya leen
            # mover_excel_gap_pc.py / iniciar_y_esperar_gap_pc.py (gana
            # siempre sobre el detector automático). Vacío = borrar el
            # archivo (volver a detectar sola).
            texto = str(cuerpo.get("ruta") or "").strip()
            if texto:
                ruta_obj = Path(texto)
                if not ruta_obj.is_dir():
                    return self._json({"error": f"Esa carpeta no existe: {texto}"}, 400)
                RUTA_GAP_TXT.write_text(texto, encoding="utf-8")
            elif RUTA_GAP_TXT.exists():
                RUTA_GAP_TXT.unlink()
            return self._json({"ok": True, "ruta": texto})

        if ruta == "/api/accion":
            c, err = iniciar_accion(cuerpo.get("accion", ""))
            if err:
                return self._json({"error": err}, 409 if "curso" in err else 400)
            return self._json({"ok": True, "id": c.id})

        if ruta == "/api/ejecutar-programada":
            # La llama tick.py. Usa la config guardada tal cual.
            cfg = cargar_config()
            if not cfg.get("reportes"):
                return self._json({"error": "Sin reportes configurados."}, 400)
            slot = cuerpo.get("slot")  # iso del slot, para anclar
            es_reintento = bool(cuerpo.get("reintento"))
            c, err = iniciar_corrida(cfg, "programada", bloquea=True)
            if err:
                return self._json({"error": err}, 409)
            kv = {"ultima_ejecucion": slot or datetime.now().isoformat(timespec="seconds"),
                  "reintentar_en": None}
            if not es_reintento:
                kv["reintentos"] = 0   # slot nuevo: el contador de reintentos arranca de cero
            set_agenda_estado(**kv)
            return self._json({"ok": True, "id": c.id})

        if ruta == "/api/detener":
            if ACTUAL is not None and ACTUAL.estado == "corriendo":
                ACTUAL.detener()
                return self._json({"ok": True})
            # Corrida automática (tick.py) — el panel no la lanzó, así que se
            # le deja una señal en un archivo y tick.py la corta.
            tk = _tick_corriendo()
            if tk:
                try:
                    (RUNS_DIR / "_detener.txt").write_text(tk["id"], encoding="utf-8")
                except Exception:
                    pass
                return self._json({"ok": True, "mensaje": "Deteniendo la corrida automática…"})
            return self._json({"error": "No hay ninguna corrida en curso."}, 409)

        if ruta == "/api/salir":
            self._json({"ok": True})
            _soltar_pid_panel()
            threading.Timer(0.3, lambda: os._exit(0)).start()
            return

        self._json({"error": "ruta no encontrada"}, 404)


URL = f"http://localhost:{PUERTO}"


def _buscar_edge():
    import shutil
    for cand in (
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
    ):
        if os.path.isfile(cand):
            return cand
    return shutil.which("msedge") or shutil.which("chrome")


def abrir_ventana_app():
    """Abre el panel como ventana de aplicación (sin barras ni pestañas).
    NO se espera este proceso: cuando Edge ya está abierto, el lanzador
    delega en la instancia existente y termina enseguida — atar la vida del
    panel a eso lo mataba y dejaba la ventana con 'conexión rechazada'."""
    navegador = _buscar_edge()
    if not navegador:
        webbrowser.open(URL)
        return
    perfil = BASE / ".ventana-app"
    try:
        subprocess.Popen([
            navegador,
            f"--app={URL}",
            f"--user-data-dir={perfil}",
            "--no-first-run", "--no-default-browser-check",
            "--window-size=1220,840",
        ])
    except Exception:
        webbrowser.open(URL)


def _watchdog(gracia):
    """Apaga el panel sólo si NADIE lo usa por MUCHÍSIMO rato (gracia grande).
    La forma normal de cerrarlo es el botón 'Salir' o cerrar la ventana (la
    UI avisa con sendBeacon). Esto es sólo una red por si algo quedó colgado.
    Nunca apaga con una corrida (manual o del tick) en curso."""
    while True:
        time.sleep(60)
        limpiar_runs()
        if (ACTUAL is not None and ACTUAL.estado == "corriendo") or _tick_corriendo():
            continue
        if time.time() - ULTIMO_PING > gracia:
            _soltar_pid_panel()
            os._exit(0)


def _log(msg):
    """print tolerante a consolas raras (pythonw / cp1252 / stdout None)."""
    try:
        sys.stdout.write(msg + "\n")
        sys.stdout.flush()
    except Exception:
        pass


def main():
    solo_servidor = "--server-only" in sys.argv

    if _otra_instancia_del_panel_viva():
        # Ya hay un panel corriendo (de verdad -- PID vivo) — solo abrí/traé
        # la ventana, NO levantar un servidor nuevo encima.
        _log(f"El panel ya esta corriendo en {URL}")
        if not solo_servidor:
            abrir_ventana_app()
        return

    try:
        srv = ThreadingHTTPServer(("127.0.0.1", PUERTO), Handler)
    except OSError:
        # Por si acaso (puerto ocupado por otra cosa, o carrera al arrancar
        # dos panel.py al mismo tiempo) — mismo comportamiento de siempre.
        _log(f"El panel ya esta corriendo en {URL}")
        if not solo_servidor:
            abrir_ventana_app()
        return

    _tomar_pid_panel()

    _log(f"Panel de Control AP  ->  {URL}")

    if not solo_servidor and "--no-ventana" not in sys.argv:
        # dar un instante a que el socket esté escuchando, después abrir la ventana
        threading.Timer(0.6, abrir_ventana_app).start()
    # Red de seguridad: sólo se apaga solo tras 12 h SIN que nadie lo toque
    # (ni ventana, ni tick). Lo normal es cerrarlo con "Salir" o cerrando la
    # ventana (la UI avisa). Antes eran 5 min y se apagaba si la ventana se
    # quedaba minimizada un rato (el navegador frena el setInterval).
    threading.Thread(target=_watchdog, args=(12 * 3600,), daemon=True).start()

    try:
        srv.serve_forever()   # bloquea acá — el panel vive mientras corre esto
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
ejecutar_todo.py — Un solo comando que corre TODO el reporte AP de punta a
punta:

  1. Descarga SAP + Descarga GAP  (3 meses, EN PARALELO)
       SAP -> ..\\CARGA\\SAP\\SAP_m0..m2.xlsx   (corre en esta PC)
       GAP -> ..\\CARGA\\EXCEL\\R_m0..m2.xls    (corre en la VM, acá se espera)
  2. Pendientes AP : procesar + publicar
  3. Atendidas AP  : procesar + publicar
  4. Veredas AP    : procesar + publicar (usa los MISMOS Excel de SAP/GAP de
                      arriba, no descarga nada aparte)

REGLAS
------
- Si CUALQUIER paso falla, se DETIENE todo ahí (la data de 3 meses tiene que
  estar completa) y muestra una pantalla de resumen diciendo qué paso falló,
  dónde está el log y dónde mirar los diagnósticos.
- El paso del GAP necesita que en la VM ya esté corriendo Vigilar-GAP-VM.bat
  (con SDAPeru abierto y logueado). Eso es responsabilidad tuya — este script
  no lo puede levantar.
- Todo lo que se ve en pantalla queda guardado en EJECUTAR-TODO\\logs\\.

USO
---
  Doble clic en Ejecutar-Todo.bat
  (o: python ejecutar_todo.py [opciones])

Opciones:
  --saltar-sap        No descargar SAP (usa lo que ya haya en CARGA\\SAP\\).
  --saltar-gap        No descargar GAP (usa lo que ya haya en CARGA\\EXCEL\\).
  --saltar-descargas  = --saltar-sap --saltar-gap (solo procesa y publica).
  --solo-pendientes   Corre SOLO Pendientes (no Atendidas ni Veredas).
  --solo-atendidas    Corre SOLO Atendidas (no Pendientes ni Veredas).
  --solo-veredas      Corre SOLO Veredas (no Pendientes ni Atendidas).
"""

import os
import subprocess
import sys
import threading
import time
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(BASE_DIR)  # "1. REPORTE"
LOGS_DIR = os.path.join(BASE_DIR, "logs")

SAP_DIR = os.path.join(RAIZ, "RPA-SAP-LOCAL")
GAP_PC_DIR = os.path.join(RAIZ, "RPA-GAP-LOCAL", "PARA-TU-PC-REAL")
PENDIENTES_DIR = os.path.join(RAIZ, "PENDIENTES-LOCAL")
ATENDIDAS_DIR = os.path.join(RAIZ, "ATENDIDAS-LOCAL")
VEREDAS_DIR = os.path.join(RAIZ, "VEREDAS-LOCAL")

MESES = 3

# El Panel de Control corre esto con PANEL_SIN_VENTANA=1 para que ni este
# proceso ni sus hijos (SAP, GAP, feed...) abran una ventana de consola.
_CREATIONFLAGS = 0
if os.name == "nt" and os.environ.get("PANEL_SIN_VENTANA"):
    _CREATIONFLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class Log:
    def __init__(self, ruta):
        self.f = open(ruta, "w", encoding="utf-8")
        self.ruta = ruta

    def write(self, texto, tambien_consola=True):
        self.f.write(texto)
        self.f.flush()
        if tambien_consola:
            try:
                sys.stdout.write(texto)
            except UnicodeEncodeError:
                sys.stdout.write(texto.encode("ascii", "replace").decode("ascii"))
            sys.stdout.flush()

    def linea(self, texto=""):
        self.write(texto + "\n")

    def close(self):
        self.f.close()


def correr_paso(log, titulo, script, cwd, args=None):
    """Corre un script .py como subproceso, mostrando su salida en vivo y
    guardándola en el log. Devuelve (ok: bool, segundos: float)."""
    args = args or []
    log.linea()
    log.linea("=" * 70)
    log.linea(f">>> {titulo}")
    log.linea(f"    {os.path.basename(script)} {' '.join(args)}   (en {cwd})")
    log.linea("=" * 70)

    if not os.path.exists(os.path.join(cwd, script)):
        log.linea(f"    ERROR: no existe {os.path.join(cwd, script)}")
        return False, 0.0

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"

    t0 = time.time()
    try:
        proc = subprocess.Popen(
            [sys.executable, script, *args],
            cwd=cwd, env=env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
            creationflags=_CREATIONFLAGS,
        )
    except Exception as exc:
        log.linea(f"    ERROR al lanzar el script: {exc}")
        return False, time.time() - t0

    for linea in proc.stdout:
        log.write("    " + linea)
    proc.wait()
    dur = time.time() - t0

    ok = proc.returncode == 0
    log.linea(f"    -> {'OK' if ok else 'FALLO (codigo ' + str(proc.returncode) + ')'}  ({dur:.0f}s)")
    return ok, dur


def _stream(proc, prefijo, log, lock):
    for linea in proc.stdout:
        with lock:
            log.write(f"    {prefijo}{linea}")


def correr_pasos_en_paralelo(log, specs):
    """specs: [{titulo, script, cwd, args, prefijo}, ...]. Lanza todos a la
    vez (SAP corre en esta PC, GAP corre en la VM y acá solo se espera, así
    que no se estorban), mezcla su salida con prefijo, y espera a todos.
    Devuelve [(titulo, ok, seg), ...] en el mismo orden de specs."""
    log.linea()
    log.linea("=" * 70)
    log.linea(">>> " + "  +  ".join(s["titulo"] for s in specs) + "   (EN PARALELO)")
    log.linea("=" * 70)

    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUNBUFFERED"] = "1"
    lock = threading.Lock()

    lanzados = []  # (spec, proc|None, t0, hilo|None)
    for s in specs:
        ruta = os.path.join(s["cwd"], s["script"])
        if not os.path.exists(ruta):
            with lock:
                log.linea(f"    {s['prefijo']}ERROR: no existe {ruta}")
            lanzados.append((s, None, time.time(), None))
            continue
        try:
            p = subprocess.Popen(
                [sys.executable, s["script"], *s.get("args", [])],
                cwd=s["cwd"], env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", bufsize=1,
                creationflags=_CREATIONFLAGS,
            )
        except Exception as exc:
            with lock:
                log.linea(f"    {s['prefijo']}ERROR al lanzar: {exc}")
            lanzados.append((s, None, time.time(), None))
            continue
        th = threading.Thread(target=_stream, args=(p, s["prefijo"], log, lock), daemon=True)
        th.start()
        lanzados.append((s, p, time.time(), th))

    resultados = []
    for s, p, t0, th in lanzados:
        if p is None:
            resultados.append((s["titulo"], False, 0.0))
            continue
        p.wait()
        if th:
            th.join(timeout=10)
        dur = time.time() - t0
        ok = p.returncode == 0
        with lock:
            log.linea(f"    -> {s['prefijo']}{'OK' if ok else 'FALLO (codigo ' + str(p.returncode) + ')'}  ({dur:.0f}s)")
        resultados.append((s["titulo"], ok, dur))
    return resultados


def main():
    argv = set(sys.argv[1:])
    saltar_sap = "--saltar-sap" in argv or "--saltar-descargas" in argv
    saltar_gap = "--saltar-gap" in argv or "--saltar-descargas" in argv
    solo_pend = "--solo-pendientes" in argv
    solo_aten = "--solo-atendidas" in argv
    solo_ver = "--solo-veredas" in argv

    # --- Opciones que usa el Panel de Control (PANEL-CONTROL-LOCAL) --------
    lista = sys.argv[1:]

    def _valor(flag, defecto=None):
        if flag in lista:
            i = lista.index(flag)
            if i + 1 < len(lista):
                return lista[i + 1]
        return defecto

    # --meses N : cuántos meses bajar/procesar (1..6). Default MESES.
    meses_sel = MESES
    try:
        v = _valor("--meses")
        if v is not None:
            meses_sel = max(1, min(6, int(v)))
    except ValueError:
        pass

    # --reportes pendientes,atendidas,veredas : subconjunto a correr.
    rep = _valor("--reportes")
    if rep:
        pedidos = {x.strip().lower() for x in rep.split(",") if x.strip()}
        solo_pend = "pendientes" in pedidos
        solo_aten = "atendidas" in pedidos
        solo_ver = "veredas" in pedidos

    silencioso = "--silencioso" in argv
    publicar = "--sin-publicar" not in argv  # modo prueba: procesa pero no publica

    algun_solo = solo_pend or solo_aten or solo_ver

    os.makedirs(LOGS_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log = Log(os.path.join(LOGS_DIR, f"corrida_{stamp}.log"))

    log.linea("############################################################")
    log.linea(f"#  REPORTE AP - corrida completa   {datetime.now():%d/%m/%Y %H:%M:%S}")
    log.linea(f"#  Log: {log.ruta}")
    log.linea("############################################################")
    if saltar_sap or saltar_gap or solo_pend or solo_aten:
        log.linea(f"#  Opciones: {' '.join(sorted(argv))}")

    # Aviso sobre la VM (solo si se va a descargar GAP).
    if not saltar_gap:
        log.linea()
        log.linea(">>> ANTES DE SEGUIR: en la VM tiene que estar corriendo")
        log.linea("    Vigilar-GAP-VM.bat (SDAPeru abierto y logueado).")
        log.linea("    Si no lo está, el paso del GAP va a fallar y se detiene todo.")

    pasos = []  # (titulo, ok, seg, saltado)

    def registrar(titulo, ok, seg, saltado=False):
        pasos.append((titulo, ok, seg, saltado))

    def abortar(titulo_fallo):
        log.linea()
        log.linea("!" * 70)
        log.linea(f"!  SE DETUVO: fallo el paso '{titulo_fallo}'.")
        log.linea("!  No se siguió con los pasos siguientes (la data quedaría incompleta).")
        log.linea("!" * 70)
        resumen(pasos, log, ok_total=False)
        log.close()
        sys.exit(1)

    # ---- Pasos 1-2: descargas SAP + GAP EN PARALELO ----
    # SAP corre en esta PC (su propio Chromium); GAP corre en la VM y acá solo
    # se espera el archivo por OneDrive -- no compiten, así que van juntas.
    _sap_args = ["--meses", str(meses_sel)] + (["--silencioso"] if silencioso else [])
    SPEC_SAP = {"titulo": f"Descarga SAP ({meses_sel} meses)", "script": "descargar_excel_sap.py",
                "cwd": SAP_DIR, "args": _sap_args, "prefijo": "[SAP] "}
    SPEC_GAP = {"titulo": f"Descarga GAP ({meses_sel} meses)", "script": "iniciar_y_esperar_gap_pc.py",
                "cwd": GAP_PC_DIR, "args": ["--meses", str(meses_sel)], "prefijo": "[GAP] "}

    a_correr = []
    if not saltar_sap:
        a_correr.append(SPEC_SAP)
    else:
        log.linea("\n(— se salta la descarga de SAP —)")
        registrar(SPEC_SAP["titulo"], True, 0, saltado=True)
    if not saltar_gap:
        a_correr.append(SPEC_GAP)
    else:
        log.linea("\n(— se salta la descarga de GAP —)")
        registrar(SPEC_GAP["titulo"], True, 0, saltado=True)

    if len(a_correr) == 2:
        resultados_desc = correr_pasos_en_paralelo(log, a_correr)
    elif len(a_correr) == 1:
        s = a_correr[0]
        ok, seg = correr_paso(log, s["titulo"], s["script"], s["cwd"], s["args"])
        resultados_desc = [(s["titulo"], ok, seg)]
    else:
        resultados_desc = []

    fallo_desc = None
    for titulo, ok, seg in resultados_desc:
        registrar(titulo, ok, seg)
        if not ok and fallo_desc is None:
            fallo_desc = titulo
    if fallo_desc:
        if "SAP" in fallo_desc:
            log.linea("    Diagnósticos SAP: %USERPROFILE%\\SAP_RPA_Excel\\diagnosticos")
        if "GAP" in fallo_desc:
            log.linea("    Revisá que Vigilar-GAP-VM.bat estuviera corriendo en la VM,")
            log.linea("    y los diagnósticos DENTRO de la VM: %USERPROFILE%\\GAP_RPA_Excel\\diagnosticos")
        abortar(fallo_desc)

    # ---- Pasos 3-4: Pendientes ----
    if solo_pend or not algun_solo:
        ok, seg = correr_paso(log, "Pendientes AP — procesar", "procesar_diario.py", PENDIENTES_DIR)
        registrar("Pendientes: procesar", ok, seg)
        if not ok:
            abortar("Pendientes: procesar")

        if not publicar:
            log.linea("\n    (modo prueba: --sin-publicar — se salta la publicación)")
            registrar("Pendientes: publicar", True, 0, saltado=True)
        elif os.path.exists(os.path.join(PENDIENTES_DIR, "config.json")):
            ok, seg = correr_paso(log, "Pendientes AP — publicar", "feed.py", PENDIENTES_DIR)
            registrar("Pendientes: publicar", ok, seg)
            if not ok:
                abortar("Pendientes: publicar")
        else:
            log.linea("\n    (no hay config.json en PENDIENTES-LOCAL — se salta la publicación)")
            registrar("Pendientes: publicar", True, 0, saltado=True)
    else:
        registrar("Pendientes: procesar", True, 0, saltado=True)
        registrar("Pendientes: publicar", True, 0, saltado=True)

    # ---- Pasos 5-6: Atendidas ----
    if solo_aten or not algun_solo:
        ok, seg = correr_paso(log, "Atendidas AP — procesar", "procesar_diario.py", ATENDIDAS_DIR)
        registrar("Atendidas: procesar", ok, seg)
        if not ok:
            abortar("Atendidas: procesar")

        if not publicar:
            log.linea("\n    (modo prueba: --sin-publicar — se salta la publicación)")
            registrar("Atendidas: publicar", True, 0, saltado=True)
        elif os.path.exists(os.path.join(ATENDIDAS_DIR, "config.json")):
            ok, seg = correr_paso(log, "Atendidas AP — publicar", "feed.py", ATENDIDAS_DIR)
            registrar("Atendidas: publicar", ok, seg)
            if not ok:
                abortar("Atendidas: publicar")
        else:
            log.linea("\n    (no hay config.json en ATENDIDAS-LOCAL — se salta la publicación)")
            registrar("Atendidas: publicar", True, 0, saltado=True)
    else:
        registrar("Atendidas: procesar", True, 0, saltado=True)
        registrar("Atendidas: publicar", True, 0, saltado=True)

    # ---- Pasos 7-8: Veredas (usa los mismos SAP/GAP de arriba) ----
    if solo_ver or not algun_solo:
        ok, seg = correr_paso(log, "Veredas AP — procesar", "procesar_diario.py", VEREDAS_DIR)
        registrar("Veredas: procesar", ok, seg)
        if not ok:
            abortar("Veredas: procesar")

        if not publicar:
            log.linea("\n    (modo prueba: --sin-publicar — se salta la publicación)")
            registrar("Veredas: publicar", True, 0, saltado=True)
        elif os.path.exists(os.path.join(VEREDAS_DIR, "config.json")):
            ok, seg = correr_paso(log, "Veredas AP — publicar", "feed.py", VEREDAS_DIR)
            registrar("Veredas: publicar", ok, seg)
            if not ok:
                abortar("Veredas: publicar")
        else:
            log.linea("\n    (no hay config.json en VEREDAS-LOCAL — se salta la publicación)")
            registrar("Veredas: publicar", True, 0, saltado=True)
    else:
        registrar("Veredas: procesar", True, 0, saltado=True)
        registrar("Veredas: publicar", True, 0, saltado=True)

    resumen(pasos, log, ok_total=True)
    log.close()


def resumen(pasos, log, ok_total):
    log.linea()
    log.linea("======================= RESUMEN =======================")
    for titulo, ok, seg, saltado in pasos:
        if saltado:
            estado = "[OMITIDO]"
        elif ok:
            estado = "[OK]     "
        else:
            estado = "[FALLO]  "
        t = f"{seg/60:.0f}m{seg%60:02.0f}s" if seg >= 60 else f"{seg:.0f}s"
        log.linea(f"  {estado} {titulo:<32} {t:>8}")
    log.linea("======================================================")
    if ok_total:
        log.linea("  TODO OK. El sitio (Vercel/Drive) muestra los datos nuevos en 1-2 min.")
    else:
        log.linea("  QUEDÓ A MEDIAS. Revisá el paso [FALLO] de arriba.")
    log.linea(f"  Log completo: {log.ruta}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelado por el usuario.")
        sys.exit(1)

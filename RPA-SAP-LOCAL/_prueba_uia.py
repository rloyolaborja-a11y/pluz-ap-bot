#!/usr/bin/env python3
"""PRUEBA de factibilidad -- NO es parte del bot, no lo corre el panel.

Objetivo: confirmar si, lanzando Edge como un programa NORMAL (sin el
protocolo de automatizacion CDP que usa Playwright -- el sospechoso de que
el antivirus mate el navegador), pywinauto puede igual "ver" y encontrar los
controles de SAP (checkboxes, botones, campos) a traves del backend "uia"
(UI Automation de Windows), que lee directo el arbol de accesibilidad del
navegador en vez de hablarle al motor de renderizado por CDP.

Si esto funciona, tiene sentido reescribir todo el robot con pywinauto puro.
Si SAP no expone bien sus controles a UI Automation, hay que buscar otro
camino ANTES de invertir en la reescritura grande.

COMO USARLO
-----------
1. python _prueba_uia.py
2. Se abre una ventana de Edge nueva (perfil separado, de prueba).
3. Andá a mano al portal de SAP, entrá a IW39 ("Visualizar ordenes PM") hasta
   la pantalla con los checkboxes "Pendiente", "En tratam.", "concluido",
   "Hist." y el campo "Clase de orden".
4. Volvé a esta consola y apreta ENTER.
5. El script va a decir que SI o NO pudo encontrar cada control, y con que
   nombre/tipo lo encontro (o no encontro nada).
"""

import subprocess
import sys
import time
from pathlib import Path

try:
    from pywinauto import Desktop
    from pywinauto.application import Application
except ImportError:
    print("Falta pywinauto. Instalalo con: pip install pywinauto")
    sys.exit(1)

PERFIL_PRUEBA = Path.home() / "SAP_RPA_Excel" / "perfil_prueba_uia"
PORTAL_URL = "https://pluz-peru-portal-prd.workzonehr.cfapps.br10.hana.ondemand.com/site#workzone-home&/home"


def encontrar_msedge() -> str:
    import shutil
    for candidato in (
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ):
        if Path(candidato).exists():
            return candidato
    encontrado = shutil.which("msedge")
    if encontrado:
        return encontrado
    print("No encontre msedge.exe -- si esta en otro lado, avisa la ruta.")
    sys.exit(1)


def main():
    PERFIL_PRUEBA.mkdir(parents=True, exist_ok=True)
    msedge = encontrar_msedge()

    print(f"Abriendo Edge NORMAL (sin CDP, sin --remote-debugging-port): {msedge}")
    # IMPORTANTE: esto es un lanzamiento de Edge COMO USUARIO NORMAL -- NO
    # pasa --remote-debugging-port ni ningun flag de automatizacion. Windows
    # y el antivirus deberian verlo igual que si el usuario le hubiera dado
    # doble clic al icono.
    proceso = subprocess.Popen([
        msedge,
        f"--user-data-dir={PERFIL_PRUEBA}",
        PORTAL_URL,
    ])
    time.sleep(3)

    input(
        "\n>>> Anda a mano a IW39 en SAP hasta la pantalla con los checkboxes "
        "'concluido'/'Hist.' y el campo 'Clase de orden'.\n"
        ">>> Cuando esa pantalla este en primer plano, volve aca y apreta ENTER... "
    )

    print("\nVentanas de Edge (msedge.exe) que encontre:")
    candidatas = []
    for w in Desktop(backend="win32").windows(visible_only=True):
        try:
            if _nombre_proceso(w.process_id()) != "msedge.exe":
                continue
        except Exception:
            continue
        candidatas.append(w)
        print(f"  [{len(candidatas) - 1}] '{w.window_text()}'  (pid={w.process_id()}, mio={w.process_id() == proceso.pid})")

    if not candidatas:
        print("No encontre ninguna ventana de msedge.exe abierta. Revisa que Edge siga abierto.")
        return

    if len(candidatas) == 1:
        ventana_sap = candidatas[0]
    else:
        idx = input(f"\nCual es la de SAP? Escribi el numero [0-{len(candidatas) - 1}]: ").strip()
        ventana_sap = candidatas[int(idx)]

    print(f"Ventana encontrada: '{ventana_sap.window_text()}' -- conectando con UI Automation...")
    app = Application(backend="uia").connect(handle=ventana_sap.handle)
    ventana_uia = app.window(handle=ventana_sap.handle)

    print("\n--- Buscando 'concluido' ---")
    _buscar_por_texto(ventana_uia, "concluido")

    print("\n--- Buscando 'Hist.' ---")
    _buscar_por_texto(ventana_uia, "Hist.")

    print("\n--- Buscando 'Clase de orden' ---")
    _buscar_por_texto(ventana_uia, "Clase de orden")

    print("\n--- Buscando boton 'Ejecutar' ---")
    _buscar_por_texto(ventana_uia, "Ejecutar")

    print("\nListo. Copiame TODO lo que salio arriba.")


def _nombre_proceso(pid: int) -> str:
    import ctypes
    from ctypes import wintypes

    handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
    if not handle:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(260)
        tam = wintypes.DWORD(260)
        ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(tam))
        return Path(buf.value).name.lower() if ok else ""
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _buscar_por_texto(ventana_uia, texto: str):
    # (2026-09-18) La version de pywinauto instalada no acepta title_re como
    # parametro de descendants() (distinto de lo que dice la doc general) --
    # se trae TODO el arbol y se filtra a mano por texto, en minuscula.
    texto_buscado = texto.lower()
    try:
        todos = ventana_uia.descendants()
    except Exception as exc:
        print(f"  ERROR trayendo el arbol completo: {exc}")
        return
    encontrados = []
    for el in todos:
        try:
            nombre = el.window_text() or ""
        except Exception:
            continue
        if texto_buscado in nombre.lower():
            encontrados.append(el)
    if not encontrados:
        print(f"  NO encontre nada con texto parecido a '{texto}'. (Arbol completo tenia {len(todos)} elementos)")
        return
    print(f"  Encontre {len(encontrados)} elemento(s):")
    for el in encontrados[:5]:
        try:
            print(f"    - tipo={el.element_info.control_type!r} nombre={el.window_text()!r}")
        except Exception:
            print("    - (no se pudo leer)")


if __name__ == "__main__":
    main()

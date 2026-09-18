#!/usr/bin/env python3
"""Reescritura del robot de SAP con pywinauto + UI Automation, SIN pasar por
el protocolo de automatizacion (CDP) que usa Playwright -- la sospecha
confirmada es que algo de seguridad de la empresa mata el navegador cuando
lo detecta controlado por ese protocolo, sin importar el navegador (Chrome,
Chromium o Edge) ni el clic puntual.

ESTADO (2026-09-18, PAUSADO -- retomar cuando haya tiempo): en construccion,
PANTALLA POR PANTALLA. NO esta conectado al bot real todavia (el bot sigue
usando descargar_excel_sap.py con Playwright/Chromium mientras tanto).

Lo que YA se probo y funciona:
  - Abrir Edge NORMAL (subprocess.Popen, sin --remote-debugging-port) y
    conectarlo con pywinauto backend="uia" SI funciona -- SAP se deja
    controlar sin pasar por CDP. Ver abrir_edge_y_conectar() (identifica la
    ventana nueva comparando handles de msedge.exe antes/despues de
    lanzarla, NO por PID de Popen -- Edge a veces reusa un proceso ya
    corriendo y ese PID nunca tiene ventana propia).
  - Los checkboxes 'concluido' y 'Hist.' SI tienen nombre accesible -> se
    encuentran bien por nombre (_esperar_control_por_nombre).
  - UI Automation de Windows tira de vez en cuando un COMError pasajero
    ("Un evento no pudo invocar a ninguno de los subscriptores") sin que
    haya nada mal -- hay que reintentar (ver _con_reintento).
  - Los textos de SAP a veces usan '\xa0' (espacio irrompible) en vez de
    espacio normal -- hay que normalizar antes de comparar (_normalizar_texto).

Lo que NO funciono / donde quedo trabado:
  - Los campos de TEXTO (Clase de orden, Periodo, Layout, y en general TODOS
    los inputs de esta pantalla) NO tienen nombre accesible NI etiqueta
    cercana -- SAP no expone esa info a UI Automation aca (a diferencia de
    los checkboxes). Buscarlos por nombre/texto es imposible.
  - Se intento identificarlos por POSICION (indice dentro de la lista de
    todos los 'Edit' con rectangulo valido, ver _campos_edit_visibles) --
    la pantalla real es: fila 1 "Orden" (desde/hasta), fila 2 "Clase de
    orden" (desde/hasta), fila 3 "Ubicacion tecnica", fila 4 "Equipo", fila 5
    "Material", fila 6 "Numero de serie", fila 7 "Dat.adic.disposit.", fila 8
    "Notificacion", fila 9 "Pto.tbjo.responsable", fila 10
    "Ce.p.pto.trabajo" ... (sigue bajando hasta llegar a "Periodo" y
    "Layout", no llegamos a ver esas dos filas en la captura). Los indices
    que se probaron (INDICE_CLASE_ORDEN=3, INDICE_PERIODO_DESDE=21,
    INDICE_PERIODO_HASTA=22) estaban CORRIDOS -- el valor de prueba "ZM06"
    termino en el campo "Orden: a:" (fila 1, columna derecha) en vez de
    "Clase de orden" (fila 2). Hay que volver a contar con cuidado,
    correlacionando una captura de pantalla REAL (con las etiquetas
    visibles) contra el volcado de _campos_edit_visibles en el MISMO
    momento, fila por fila.
  - Falta totalmente resolver 'Layout' (no se llego a probar el Tab-desde-
    Periodo que se dejo armado en el codigo, ver el bloque "DIAGNOSTICO:
    foco despues de cada Tab").

PROXIMO PASO sugerido al retomar: correr esto de nuevo, SIN llenar nada a
mano, y en vez de confiar en indices ya calculados, usar el volcado de
_volcar_diagnostico_etiqueta / _campos_edit_visibles JUNTO con una captura
de pantalla tomada en el MISMO instante (ImageGrab.grab(), ya esta el
codigo) para contar los indices bien desde cero, fila por fila, sin apurar.
Una vez que Clase de orden/Periodo/Layout esten confirmados, falta TODA la
Pantalla 2 (lista de resultados + filtro por Fecha de creacion) y la
Pantalla 3 (exportar a Excel) -- ninguna de las dos esta ni empezada.

Se corre a mano (PRUEBA-uia-pantalla1.bat) para ir verificando cada parte
contra el SAP real antes de conectarlo al bot de verdad.
"""

from __future__ import annotations

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

PORTAL_URL = "https://pluz-peru-portal-prd.workzonehr.cfapps.br10.hana.ondemand.com/site#workzone-home&/home"
IW39_URL_DIRECTA = "https://pluz-peru-portal-prd.workzonehr.cfapps.br10.hana.ondemand.com/site#MaintenanceOrder-displayList?sap-app-origin-hint=&sap-ui-app-id-hint=s4hana_EA13B4961465613FB41F616A1868F01D&sap-ui-tech-hint=GUI"

PROFILE_DIR = Path.home() / "SAP_RPA_Excel" / "perfil_edge_uia"
CLASE_ORDEN = "ZM06"
RANGO_DIAS = 30

TIMEOUT_ESPERA_VENTANA_SEG = 20
TIMEOUT_ESPERA_CONTROL_SEG = 60

LOG_PATH = Path.home() / "SAP_RPA_Excel" / "prueba_uia_log.txt"


class _Tee:
    """Duplica todo lo que se imprime por consola tambien a un archivo, para
    no depender de que alguien copie/pegue a mano lo que salio -- el archivo
    queda en LOG_PATH despues de correr el script."""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, dato):
        for s in self.streams:
            try:
                s.write(dato)
                s.flush()
            except Exception:
                pass

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except Exception:
                pass


def _activar_log_a_archivo() -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    archivo = open(LOG_PATH, "w", encoding="utf-8")
    sys.stdout = _Tee(sys.__stdout__, archivo)
    sys.stderr = _Tee(sys.__stderr__, archivo)
    print(f"(Todo esto tambien se esta guardando en: {LOG_PATH})\n")


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


def _encontrar_msedge() -> str:
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
    raise RuntimeError("No se encontro msedge.exe en las rutas conocidas.")


def abrir_edge_y_conectar(url: str):
    """Abre Edge NORMAL (como si el usuario le diera doble clic -- sin
    --remote-debugging-port ni ningun flag de automatizacion) y devuelve el
    WindowSpecification (backend UIA) de la ventana que abrimos.

    Identifica la ventana correcta comparando que ventanas de msedge.exe
    HABIA ANTES de lanzarlo contra las que hay DESPUES -- la diferencia es
    la que acabamos de abrir. Esto es mas confiable que buscar por el PID
    exacto del proceso que lanzamos con Popen: Edge a veces reutiliza un
    proceso ya corriendo (por su mecanismo interno de instancia unica) y esa
    PID nunca llega a tener una ventana propia, aunque la pestana/ventana SI
    se abra bien (confirmado en la PC de la usuaria: el PID de Popen nunca
    aparecio, pero Edge abrio la pagina igual en otro proceso).

    Importante para no confundirla con OTRAS ventanas de Edge que la
    usuaria ya tenga abiertas (su navegacion normal, 'Secure Desktops' de
    Oracle, el Panel de Control que tambien es una app de Edge, etc.): se
    guarda el conjunto de handles ANTES de lanzar, y solo se acepta un
    handle que NO estuviera en ese conjunto."""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    msedge = _encontrar_msedge()

    def _handles_msedge():
        handles = set()
        for w in Desktop(backend="win32").windows(visible_only=True):
            try:
                if _nombre_proceso(w.process_id()) == "msedge.exe":
                    handles.add(w.handle)
            except Exception:
                continue
        return handles

    handles_antes = _handles_msedge()
    proceso = subprocess.Popen([msedge, f"--user-data-dir={PROFILE_DIR}", url])

    deadline = time.time() + TIMEOUT_ESPERA_VENTANA_SEG
    while time.time() < deadline:
        nuevos = _handles_msedge() - handles_antes
        if nuevos:
            handle_nuevo = next(iter(nuevos))
            app = Application(backend="uia").connect(handle=handle_nuevo)
            return app.window(handle=handle_nuevo), proceso
        time.sleep(0.5)
    raise RuntimeError(
        f"No aparecio ninguna ventana NUEVA de msedge.exe dentro de {TIMEOUT_ESPERA_VENTANA_SEG}s."
    )


def _con_reintento(func, intentos=4, espera=1.0):
    """UI Automation de Windows a veces tira un COMError pasajero
    ('Un evento no pudo invocar a ninguno de los subscriptores') sin que
    haya nada realmente mal -- confirmado en la PC de la usuaria. Reintenta
    unas pocas veces antes de darlo por error de verdad."""
    ultimo_exc = None
    for intento in range(intentos):
        try:
            return func()
        except Exception as exc:
            ultimo_exc = exc
            if intento < intentos - 1:
                time.sleep(espera)
    raise ultimo_exc


def _todos_los_descendientes(ventana, forzar_refresco=True):
    return _con_reintento(lambda: ventana.descendants())


def _normalizar_texto(s: str) -> str:
    """SAP renderiza algunas etiquetas con espacios NO estandar (ej. '\xa0'
    -- espacio irrompible -- en vez de un espacio normal, confirmado viendo
    el titulo de la ventana: 'Visualizar\xa0ordenes\xa0PM...'). Una busqueda
    con espacio comun nunca matchea eso. Se normalizan todos los espacios
    raros (\xa0,  , ​, tabs) a un espacio normal antes de comparar."""
    import unicodedata

    s = s.replace("\xa0", " ").replace(" ", " ").replace("​", "").replace("\t", " ")
    s = unicodedata.normalize("NFKC", s)
    return " ".join(s.split())


def _buscar_control_por_nombre(ventana, nombre: str, tipos=None, exacto=False):
    """Busca un control cuyo Name (accesible) coincida con 'nombre'. Si
    'tipos' se pasa (ej. ('CheckBox',)), filtra tambien por control_type."""
    objetivo = _normalizar_texto(nombre).lower()
    for el in _todos_los_descendientes(ventana):
        try:
            n = (el.window_text() or "").strip()
        except Exception:
            continue
        if not n:
            continue
        n_low = _normalizar_texto(n).lower()
        coincide = (n_low == objetivo) if exacto else (objetivo in n_low)
        if not coincide:
            continue
        if tipos is not None:
            try:
                if el.element_info.control_type not in tipos:
                    continue
            except Exception:
                continue
        return el
    return None


def _esperar_control_por_nombre(ventana, nombre: str, tipos=None, exacto=False, timeout=TIMEOUT_ESPERA_CONTROL_SEG):
    deadline = time.time() + timeout
    while time.time() < deadline:
        el = _buscar_control_por_nombre(ventana, nombre, tipos=tipos, exacto=exacto)
        if el is not None:
            return el
        time.sleep(1)
    raise RuntimeError(f"No encontre el control '{nombre}' (tipos={tipos}) en {timeout}s.")


def _campo_cercano_a_etiqueta(ventana, etiqueta: str, tipos_campo=("Edit",)):
    """SAP GUI clasico no siempre pone el texto de la etiqueta como Name del
    campo -- el campo suele estar SUELTO al lado de un elemento de texto
    (Text/Static/Pane) que dice la etiqueta. Buscamos el/los elemento(s) de
    texto con ese nombre, y de ahi el campo de tipo 'Edit' mas cercano A LA
    DERECHA en la misma fila (mismo rango vertical aproximado)."""
    objetivo = _normalizar_texto(etiqueta).lower()
    etiquetas = []
    for el in _todos_los_descendientes(ventana):
        try:
            n = (el.window_text() or "").strip()
        except Exception:
            continue
        if n and objetivo in _normalizar_texto(n).lower():
            etiquetas.append(el)

    campos = []
    for el in _todos_los_descendientes(ventana):
        try:
            if el.element_info.control_type in tipos_campo:
                campos.append(el)
        except Exception:
            continue

    mejor = None
    mejor_dist = None
    for et in etiquetas:
        try:
            r_et = et.rectangle()
        except Exception:
            continue
        centro_y_et = (r_et.top + r_et.bottom) / 2
        for campo in campos:
            try:
                r_campo = campo.rectangle()
            except Exception:
                continue
            centro_y_campo = (r_campo.top + r_campo.bottom) / 2
            if r_campo.left < r_et.left:
                continue  # el campo tiene que estar a la derecha de la etiqueta
            if abs(centro_y_campo - centro_y_et) > 15:
                continue  # no esta en la misma fila (tolerancia de 15px)
            dist = r_campo.left - r_et.left
            if mejor_dist is None or dist < mejor_dist:
                mejor, mejor_dist = campo, dist
    return mejor


def _volcar_diagnostico_etiqueta(ventana, texto_parcial: str) -> None:
    """Cuando no se encuentra un campo, imprime TODO lo que tenga ese texto
    parcial en el nombre (etiquetas) y TODOS los campos editables cercanos
    en altura, con su rectangulo -- para poder ver a ojo por que no matcheo
    la busqueda automatica (nombre distinto, sin rectangulo, tipo raro,
    etc.) sin tener que adivinar."""
    print(f"\n  --- DIAGNOSTICO: elementos con '{texto_parcial}' en el nombre ---")
    encontrados = 0
    etiquetas_rect = []
    for el in _todos_los_descendientes(ventana):
        try:
            n = (el.window_text() or "").strip()
        except Exception:
            continue
        if _normalizar_texto(texto_parcial).lower() in _normalizar_texto(n).lower() and n:
            encontrados += 1
            try:
                r = el.rectangle()
                print(f"    nombre={n!r} tipo={el.element_info.control_type!r} rect={r}")
                etiquetas_rect.append(r)
            except Exception:
                print(f"    nombre={n!r} tipo={el.element_info.control_type!r} (sin rectangulo)")
    if encontrados == 0:
        print(f"    (NADA tenia '{texto_parcial}' en el nombre)")

    if etiquetas_rect:
        print(f"  --- Campos editables (Edit/Document/Pane) cerca de esa altura ---")
        for el in _todos_los_descendientes(ventana):
            try:
                tipo = el.element_info.control_type
            except Exception:
                continue
            if tipo not in ("Edit", "Document", "Pane"):
                continue
            try:
                r = el.rectangle()
            except Exception:
                continue
            for r_et in etiquetas_rect:
                centro_et = (r_et.top + r_et.bottom) / 2
                centro_campo = (r.top + r.bottom) / 2
                if abs(centro_campo - centro_et) <= 30:
                    try:
                        n = el.window_text()
                    except Exception:
                        n = "?"
                    print(f"    tipo={tipo!r} nombre={n!r} rect={r}")
                    break
    else:
        # No encontramos NINGUNA etiqueta -- volcamos TODOS los campos
        # editables de la ventana entera, y una foto de pantalla, para poder
        # identificar el correcto a ojo por posicion en vez de por nombre.
        print("  --- No hubo ninguna etiqueta para guiarse -- listando TODOS los Edit/Document/Pane de la ventana ---")
        for i, el in enumerate(_todos_los_descendientes(ventana)):
            try:
                tipo = el.element_info.control_type
            except Exception:
                continue
            if tipo not in ("Edit", "Document", "Pane"):
                continue
            try:
                r = el.rectangle()
                n = el.window_text()
            except Exception:
                continue
            print(f"    #{i} tipo={tipo!r} nombre={n!r} rect={r}")

        try:
            from PIL import ImageGrab

            ruta_foto = LOG_PATH.parent / "prueba_uia_pantalla.png"
            ImageGrab.grab().save(ruta_foto)
            print(f"  (foto de pantalla completa guardada en: {ruta_foto})")
        except Exception as exc:
            print(f"  (no se pudo sacar la foto de pantalla: {exc})")
    print("  --- fin diagnostico ---\n")


# (2026-09-18) Indices confirmados a partir de una corrida real llenada a
# mano: en la lista de TODOS los campos 'Edit' con rectangulo valido (no
# (0,0,0,0)), en el orden en que aparecen, el campo #3 (contando desde 0) es
# 'Clase de orden' y los campos #21/#22 son 'Periodo desde'/'Periodo hasta'.
# Ninguno de estos campos tiene nombre accesible ni etiqueta cercana (SAP no
# expone esa info a UI Automation en esta pantalla) -- la POSICION es la
# unica pista estable que encontramos.
INDICE_CLASE_ORDEN = 3
INDICE_PERIODO_DESDE = 21
INDICE_PERIODO_HASTA = 22


def _campos_edit_visibles(ventana):
    """Lista de campos 'Edit' con rectangulo valido, en el orden en que
    aparecen en el arbol de accesibilidad (orden estable mientras la
    estructura de la pantalla no cambie)."""
    campos = []
    for el in _todos_los_descendientes(ventana):
        try:
            if el.element_info.control_type != "Edit":
                continue
            r = el.rectangle()
            if r.width() <= 0 or r.height() <= 0:
                continue
        except Exception:
            continue
        campos.append(el)
    return campos


def _elemento_con_foco():
    """Devuelve informacion legible del elemento que tiene el foco de
    teclado AHORA MISMO (via UI Automation), para diagnosticar a donde nos
    lleva un Tab sin tener que adivinar a ciegas."""
    try:
        from pywinauto.uia_defines import IUIA
        from pywinauto.uia_element_info import UIAElementInfo
        from pywinauto.controls.uiawrapper import UIAWrapper

        elem = IUIA().iuia.GetFocusedElement()
        info = UIAElementInfo(elem)
        wrapper = UIAWrapper(info)
        try:
            r = wrapper.rectangle()
        except Exception:
            r = None
        try:
            nombre = wrapper.window_text()
        except Exception:
            nombre = "?"
        return f"tipo={info.control_type!r} nombre={nombre!r} rect={r}"
    except Exception as exc:
        return f"(no se pudo leer el foco: {exc})"


def llenar_pantalla_seleccion(ventana, desde: str, hasta: str) -> None:
    print("Marcando 'concluido'...")
    chk_concluido = _esperar_control_por_nombre(ventana, "concluido", tipos=("CheckBox",))
    if not _con_reintento(lambda: chk_concluido.get_toggle_state()):
        _con_reintento(lambda: chk_concluido.click_input())

    print("Marcando 'Hist.'...")
    chk_hist = _esperar_control_por_nombre(ventana, "Hist.", tipos=("CheckBox",), exacto=True)
    if not _con_reintento(lambda: chk_hist.get_toggle_state()):
        _con_reintento(lambda: chk_hist.click_input())

    campos = _con_reintento(lambda: _campos_edit_visibles(ventana))
    print(f"  ({len(campos)} campos 'Edit' visibles encontrados)")
    necesarios = max(INDICE_CLASE_ORDEN, INDICE_PERIODO_DESDE, INDICE_PERIODO_HASTA)
    if len(campos) <= necesarios:
        raise RuntimeError(
            f"Esperaba al menos {necesarios + 1} campos 'Edit' visibles, "
            f"pero solo encontre {len(campos)}. La pantalla debe ser distinta "
            "a la esperada -- revisa la foto/log."
        )

    print(f"Clase de orden = {CLASE_ORDEN}  (campo #{INDICE_CLASE_ORDEN})...")
    campo_clase = campos[INDICE_CLASE_ORDEN]
    campo_clase.click_input()
    campo_clase.type_keys("^a")
    campo_clase.type_keys(CLASE_ORDEN, with_spaces=True)

    print(f"Periodo = {desde} a {hasta}  (campos #{INDICE_PERIODO_DESDE}/#{INDICE_PERIODO_HASTA})...")
    campo_desde = campos[INDICE_PERIODO_DESDE]
    campo_desde.click_input()
    campo_desde.type_keys("^a")
    campo_desde.type_keys(desde, with_spaces=True)
    campo_hasta = campos[INDICE_PERIODO_HASTA]
    campo_hasta.click_input()
    campo_hasta.type_keys("^a")
    campo_hasta.type_keys(hasta, with_spaces=True)

    # (2026-09-18) Layout: por ahora NO sabemos su indice (esta pantalla
    # tiene cientos de campos ocultos de "seleccion multiple" entre Periodo
    # y Layout, contar el indice exacto a mano es muy propenso a error).
    # Probamos algo distinto: Tab desde el campo de Periodo-hasta (donde SI
    # sabemos pararnos) -- el orden de tabulacion del teclado deberia saltar
    # directo a los controles VISIBLES de verdad, sin pasar por las filas
    # ocultas. Se imprime que quedo enfocado despues de cada Tab para poder
    # ajustar la cantidad si hace falta.
    print("\n  --- DIAGNOSTICO: foco despues de cada Tab desde Periodo-hasta ---")
    for i in range(1, 6):
        campo_hasta.type_keys("{TAB}")
        time.sleep(0.3)
        print(f"    Tab #{i}: {_elemento_con_foco()}")
    print("  --- fin diagnostico de Tabs ---\n")


def clickear_ejecutar(ventana) -> None:
    print("Ejecutando...")
    boton = _esperar_control_por_nombre(ventana, "Ejecutar", tipos=("Button",))
    boton.click_input()


def main():
    from datetime import datetime, timedelta

    _activar_log_a_archivo()

    hasta_dt = datetime.now()
    desde_dt = hasta_dt - timedelta(days=RANGO_DIAS)
    desde = desde_dt.strftime("%d.%m.%Y")
    hasta = hasta_dt.strftime("%d.%m.%Y")

    print(f"Abriendo Edge (normal, sin CDP) en IW39 directo...")
    ventana, proceso = abrir_edge_y_conectar(IW39_URL_DIRECTA)
    titulo = _con_reintento(lambda: ventana.window_text())
    print(f"Ventana conectada: '{titulo}'")

    input(
        "\n>>> Si SAP pide login, hacelo a mano en esa ventana. Cuando estes "
        "en la pantalla de IW39 con los checkboxes, volve aca y apreta ENTER... "
    )

    llenar_pantalla_seleccion(ventana, desde, hasta)
    print("\nPantalla 1 llenada. Reviso los datos EN LA VENTANA antes de seguir.")
    input(">>> Si esta todo bien, apreta ENTER para clickear 'Ejecutar' (o Ctrl+C para cortar aca)... ")
    clickear_ejecutar(ventana)

    print("\nListo -- deberia estar armando la lista de resultados. Sigo despues con la Pantalla 2.")


if __name__ == "__main__":
    main()

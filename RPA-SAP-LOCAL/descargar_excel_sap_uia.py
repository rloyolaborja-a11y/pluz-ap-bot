#!/usr/bin/env python3
"""Reescritura del robot de SAP con pywinauto + UI Automation, SIN pasar por
el protocolo de automatizacion (CDP) que usa Playwright -- la sospecha
confirmada es que algo de seguridad de la empresa mata el navegador cuando
lo detecta controlado por ese protocolo, sin importar el navegador (Chrome,
Chromium o Edge) ni el clic puntual.

ESTADO (2026-09-24, EN CONSTRUCCION -- Pantalla 1 resuelta): NO esta
conectado al bot real todavia (el bot sigue usando descargar_excel_sap.py
con Playwright/Chromium mientras tanto), pero la Pantalla 1 completa
(checkboxes + Clase de orden + Periodo + Layout) ya funciona de punta a
punta contra SAP real.

Lo que YA se probo y funciona:
  - Abrir Edge NORMAL (subprocess.Popen, sin --remote-debugging-port) y
    conectarlo con pywinauto backend="uia" SI funciona -- SAP se deja
    controlar sin pasar por CDP. Ver abrir_edge_y_conectar() (identifica la
    ventana nueva comparando handles de msedge.exe antes/despues de
    lanzarla, NO por PID de Popen -- Edge a veces reusa un proceso ya
    corriendo y ese PID nunca tiene ventana propia).
  - (2026-09-24) abrir_edge_y_conectar() lanza Edge con el flag
    --force-renderer-accessibility. SIN esto, Chromium decidia el solo
    (mal) si construir el arbol de accesibilidad completo del contenido
    web -- confirmado con diagnostico_pantalla1.py: sin el flag, pywinauto
    solo veia 1 elemento 'Edit' en TODA la ventana (la barra de
    direcciones), sin importar que tan visible estuviera SAP en pantalla.
    Con el flag, aparecen los 160+ campos reales de la pantalla. Este flag
    NO habilita CDP ni control remoto, solo afecta como Chromium arma su
    propio arbol interno -- no deberia disparar la misma deteccion del
    antivirus que hizo abandonar CDP.
  - Los checkboxes 'concluido' y 'Hist.' SI tienen nombre accesible -> se
    encuentran bien por nombre (_esperar_control_por_nombre).
  - UI Automation de Windows tira de vez en cuando un COMError pasajero
    ("Un evento no pudo invocar a ninguno de los subscriptores") sin que
    haya nada mal -- hay que reintentar (ver _con_reintento).
  - Los textos de SAP a veces usan '\xa0' (espacio irrompible) en vez de
    espacio normal -- hay que normalizar antes de comparar (_normalizar_texto).
  - (2026-09-24) Los indices de los campos de TEXTO (Clase de orden,
    Periodo, Layout) quedaron confirmados usando diagnostico_pantalla1.py
    -- un script aparte que saca una foto de la pantalla real con un
    numero dibujado encima de cada campo 'Edit', para identificar el indice
    correcto con solo mirar la foto en vez de contar a mano (que fue justo
    donde se trabo el intento anterior). Ver INDICE_CLASE_ORDEN,
    INDICE_PERIODO_DESDE/HASTA e INDICE_LAYOUT mas abajo para los valores y
    la explicacion completa.
  - (2026-09-24) Layout (indice 162, muy abajo en la pantalla, fuera de lo
    visible sin scroll) se llena con set_focus() en vez de click_input():
    pedirle el foco a un elemento de contenido web via UI Automation hace
    que Chromium lo scrollee solo hasta que quede visible, evitando tener
    que calcular ni mover el scroll a mano o depender de coordenadas de
    pantalla que podrian estar fuera de vista.
  - (2026-09-24) IMPORTANTE -- llenar_pantalla_seleccion() y
    clickear_ejecutar() ya NO usan click_input()/type_keys() (que simulan
    mouse/teclado REALES y le mueven el cursor y le interrumpen lo que este
    haciendo en su PC a quien lo corra -- confirmado, se quejo de esto
    probando el robot). Se cambio a llamadas puras de UI Automation que NO
    tocan mouse/teclado: set_edit_text() (IUIAutomationValuePattern.SetValue,
    la misma API que usa un lector de pantalla), toggle()
    (IUIAutomationTogglePattern.Toggle) e invoke()
    (IUIAutomationInvokePattern.Invoke). Con esto el robot deberia poder
    correr sin robarle el mouse a la usuaria mientras hace otra cosa.
    PENDIENTE DE PROBAR: confirmar que corriendo asi, de verdad no se mueve
    el mouse ni interrumpe.

Lo que NO esta hecho todavia:
  - Pantalla 2 (lista de resultados + filtro por Fecha de creacion) y
    Pantalla 3 (exportar a Excel) -- ninguna de las dos esta ni empezada.
    main() todavia corta despues de clickear 'Ejecutar'.

Se corre a mano (DIAGNOSTICO-Pantalla1-SAP.bat para el diagnostico visual,
o este mismo archivo directo) para ir verificando cada parte contra el SAP
real antes de conectarlo al bot de verdad.
"""

from __future__ import annotations

import ctypes
import subprocess
import sys
import time
from pathlib import Path

import comtypes

try:
    from pywinauto import Desktop
    from pywinauto.application import Application
    from pywinauto.uia_defines import get_elem_interface
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
    # (2026-09-24) --force-renderer-accessibility: SIN esto, Chromium decide
    # el mismo si construye o no el arbol de accesibilidad completo del
    # contenido web (deteccion automatica de "hay un lector de pantalla/
    # cliente UIA escuchando"), y esa deteccion resulto poco confiable con la
    # pantalla de SAP -- confirmado con diagnostico_pantalla1.py: con Edge
    # completamente visible y la pantalla de SAP cargada, pywinauto solo veia
    # 1 elemento 'Edit' (la barra de direcciones), ninguno de los campos de
    # SAP. Este flag fuerza esa construccion desde el arranque, sin
    # depender de la deteccion. A diferencia de --remote-debugging-port, NO
    # habilita CDP ni ningun protocolo de control remoto -- solo cambia como
    # Chromium arma su propio arbol de accesibilidad interno -- no deberia
    # disparar la misma deteccion del antivirus que causo abandonar CDP.
    proceso = subprocess.Popen([
        msedge,
        f"--user-data-dir={PROFILE_DIR}",
        "--force-renderer-accessibility",
        url,
    ])

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


# (2026-09-24) Indices RECONFIRMADOS con diagnostico_pantalla1.py (foto con
# numeros dibujados encima de cada campo, ver ese script) -- los indices
# viejos (3/21/22, de 2026-09-18) estaban CORRIDOS, como ya avisaban las
# notas de mas arriba. Ademas se necesito el flag
# --force-renderer-accessibility en abrir_edge_y_conectar() para que
# Chromium exponga TODOS los campos de la pantalla via UI Automation --
# sin ese flag, solo se veia 1 campo (la barra de direcciones de Edge) sin
# importar que tan visible estuviera la pantalla de SAP.
# En la lista de TODOS los campos 'Edit' con rectangulo valido (no
# (0,0,0,0)), en el orden en que aparecen:
#   #4   = Clase de orden (desde)
#   #24  = Periodo desde
#   #25  = Periodo hasta
#   #162 = Layout
# Ninguno de estos campos tiene nombre accesible ni etiqueta cercana (SAP no
# expone esa info a UI Automation en esta pantalla) -- la POSICION es la
# unica pista estable que encontramos.
INDICE_CLASE_ORDEN = 4
INDICE_PERIODO_DESDE = 24
INDICE_PERIODO_HASTA = 25
INDICE_LAYOUT = 162
LAYOUT = "/PRT"


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
    # (2026-09-24) .toggle() llama a IUIAutomationTogglePattern.Toggle() --
    # cambia el estado del checkbox via COM, sin clic de mouse real (a
    # diferencia de click_input(), que si mueve el cursor de verdad).
    print("Marcando 'concluido'...")
    chk_concluido = _esperar_control_por_nombre(ventana, "concluido", tipos=("CheckBox",))
    if not _con_reintento(lambda: chk_concluido.get_toggle_state()):
        _con_reintento(lambda: chk_concluido.toggle())

    print("Marcando 'Hist.'...")
    chk_hist = _esperar_control_por_nombre(ventana, "Hist.", tipos=("CheckBox",), exacto=True)
    if not _con_reintento(lambda: chk_hist.get_toggle_state()):
        _con_reintento(lambda: chk_hist.toggle())

    campos = _con_reintento(lambda: _campos_edit_visibles(ventana))
    print(f"  ({len(campos)} campos 'Edit' visibles encontrados)")
    necesarios = max(INDICE_CLASE_ORDEN, INDICE_PERIODO_DESDE, INDICE_PERIODO_HASTA, INDICE_LAYOUT)
    if len(campos) <= necesarios:
        raise RuntimeError(
            f"Esperaba al menos {necesarios + 1} campos 'Edit' visibles, "
            f"pero solo encontre {len(campos)}. La pantalla debe ser distinta "
            "a la esperada -- revisa la foto/log."
        )

    # (2026-09-24) IMPORTANTE: se usa set_edit_text() en vez de
    # click_input()+type_keys(). click_input()/type_keys() simulan un CLIC Y
    # TECLADO REALES (mueven el mouse de la usuaria de verdad y le
    # interrumpen lo que este haciendo en su PC en ese momento -- confirmado,
    # se quejo de esto probando el robot). set_edit_text() en cambio llama
    # directo a IUIAutomationValuePattern.SetValue() -- la MISMA API que usa
    # un lector de pantalla para escribir por la persona -- sin tocar el
    # mouse ni el teclado real, y sin robarle el foco a la ventana que la
    # usuaria tenga abierta en ese momento. set_edit_text(texto), sin indicar
    # pos_start/pos_end, reemplaza TODO el contenido previo del campo (igual
    # que Ctrl+A y despues escribir).
    #   PERO set_edit_text() exige que el elemento este "visible" (no
    #   offscreen DENTRO DE SU PROPIA PAGINA -- esto no tiene nada que ver
    #   con que otra ventana lo tape en el escritorio, eso no afecta a UI
    #   Automation). La pantalla arranca con el scroll arriba del todo, asi
    #   que cualquier campo mas abajo que "Clase de orden" arranca offscreen
    #   -- confirmado, fallo justo en Periodo con ElementNotVisible. Por eso
    #   TODOS los campos pasan primero por set_focus() (tambien pura UI
    #   Automation, sin mouse/teclado real) para que Chromium los scrollee
    #   solo antes de escribirles.
    def _esperar_que_quede_visible(campo, timeout=5.0):
        """El scroll que dispara set_focus() no es instantaneo -- en vez de
        adivinar cuanto dormir, se chequea el estado REAL del elemento
        (is_visible(), que lee la propiedad IsOffscreen de UI Automation)
        en un loop corto hasta confirmar que ya esta visible, o hasta
        timeout. Determinista: en cuanto esta visible, sigue de una, sin
        esperar de mas."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                if campo.is_visible():
                    return
            except Exception:
                pass
            time.sleep(0.1)
        raise RuntimeError("El campo no quedo visible a tiempo despues de pedirle el foco (set_focus).")

    def _campo_fresco(indice):
        """NO reusa el objeto 'campo' guardado en la lista 'campos' de mas
        arriba -- al hacerle scroll para traer un campo lejano a la vista,
        SAP puede reciclar/recrear esa parte de la pagina, y la referencia
        vieja a ese elemento queda invalida, aunque el campo #indice siga
        siendo el mismo logicamente. Por eso se vuelve a pedir la lista de
        campos FRESCA justo antes de escribir."""
        campo = _campos_edit_visibles(ventana)[indice]
        campo.set_focus()
        _esperar_que_quede_visible(campo)
        # Se vuelve a pedir de nuevo aca: el paso anterior (esperar a que
        # quede visible) puede el mismo haber disparado otro reciclado.
        return _campos_edit_visibles(ventana)[indice]

    def _valor_actual(indice):
        try:
            return _campos_edit_visibles(ventana)[indice].get_value()
        except Exception:
            return None

    def _quedo_escrito(indice, texto, espera=0.35):
        # (2026-09-24) CLAVE: escribir "por atras" (ValuePattern o
        # LegacyIAccessible) cambia lo que se VE en el campo, pero SAP
        # (SAPUI5/Fiori) tiene su PROPIO modelo de datos interno separado del
        # DOM -- si no se disparan los eventos de teclado reales que SAP
        # escucha, su modelo interno nunca se entera del cambio, y SAP
        # redibuja el campo desde su modelo (el valor viejo) enseguida,
        # BORRANDO lo que se acababa de escribir -- confirmado a ojo (se veia
        # "ZM06" aparecer y desaparecer solo). set_edit_text()/SetValue() NO
        # tiran error cuando esto pasa (la llamada en si funciona bien),
        # asi que la UNICA forma de saber si de verdad quedo es volver a leer
        # el campo despues de darle un instante a SAP para redibujar, y
        # comparar.
        time.sleep(espera)
        return _valor_actual(indice) == texto

    def _por_teclado(indice, texto):
        # (2026-09-24) type_keys() SI dispara los eventos de teclado reales
        # que SAP necesita para actualizar su modelo interno -- por eso es la
        # unica via que persiste de verdad. A diferencia de click_input(), no
        # mueve el mouse, pero SI manda teclado real a nivel de Windows, que
        # va a la ventana que este en PRIMER PLANO en ese instante -- si la
        # usuaria esta trabajando en otra ventana en ese momento, el texto se
        # escribiria ahi por error. Por eso se trae Edge al frente JUSTO
        # antes de escribir y se devuelve el foco a lo que la usuaria tuviera
        # abierto apenas se termina -- la interrupcion queda acotada a los
        # campos que de verdad lo necesitan, no a toda la corrida.
        user32 = ctypes.windll.user32
        user32.GetForegroundWindow.restype = ctypes.c_void_p
        user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
        anterior = user32.GetForegroundWindow()
        user32.SetForegroundWindow(ventana.handle)
        try:
            campo = _campo_fresco(indice)
            campo.type_keys("^a{DEL}", set_foreground=False)
            campo.type_keys(texto, with_spaces=True, set_foreground=False)
        finally:
            if anterior:
                user32.SetForegroundWindow(anterior)

    def _llenar(indice, texto):
        # Intento 1: ValuePattern (set_edit_text) -- el mas silencioso.
        try:
            _con_reintento(lambda: _campo_fresco(indice).set_edit_text(texto), intentos=3, espera=0.5)
            if _quedo_escrito(indice, texto):
                return
            print("    (SAP borro el valor solo -- ValuePattern no le avisa a SAP del cambio, probando otra via)")
        except comtypes.COMError:
            pass

        # Intento 2: LegacyIAccessible -- la interfaz VIEJA de accesibilidad
        # (la que usaban los lectores de pantalla antes de UI Automation).
        # Sigue siendo pura COM, cero simulacion de input.
        print("    (probando por LegacyIAccessible)")
        try:
            def via_legacy():
                campo = _campo_fresco(indice)
                iface = get_elem_interface(campo.element_info.element, "LegacyIAccessible")
                iface.SetValue(texto)
            _con_reintento(via_legacy, intentos=3, espera=0.5)
            if _quedo_escrito(indice, texto):
                return
            print("    (SAP tambien borro este -- probando por teclado)")
        except comtypes.COMError:
            pass

        # Intento 3 (ultimo recurso): teclado real. Ver _por_teclado() arriba
        # para la explicacion del breve parpadeo de foreground que esto
        # implica.
        _con_reintento(lambda: _por_teclado(indice, texto), intentos=3, espera=0.5)
        if not _quedo_escrito(indice, texto):
            raise RuntimeError(
                f"El campo #{indice} no acepto '{texto}' ni siquiera por teclado -- "
                "revisar a mano si es el campo correcto."
            )

    print(f"Clase de orden = {CLASE_ORDEN}  (campo #{INDICE_CLASE_ORDEN})...")
    _llenar(INDICE_CLASE_ORDEN, CLASE_ORDEN)

    print(f"Periodo = {desde} a {hasta}  (campos #{INDICE_PERIODO_DESDE}/#{INDICE_PERIODO_HASTA})...")
    _llenar(INDICE_PERIODO_DESDE, desde)
    _llenar(INDICE_PERIODO_HASTA, hasta)

    print(f"Layout = {LAYOUT}  (campo #{INDICE_LAYOUT})...")
    _llenar(INDICE_LAYOUT, LAYOUT)


def clickear_ejecutar(ventana) -> None:
    # (2026-09-24) .invoke() llama a IUIAutomationInvokePattern.Invoke() --
    # activa el boton via COM, sin clic de mouse real.
    print("Ejecutando...")
    boton = _esperar_control_por_nombre(ventana, "Ejecutar", tipos=("Button",))
    boton.invoke()


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

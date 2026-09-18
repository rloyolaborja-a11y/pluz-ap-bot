#!/usr/bin/env python3
"""Extrae automaticamente el reporte GAP (Extractor de Datos SAP, dentro
de "Sistema de Distribucion - Modulos Locales" / SDAPeru).

ESTE SCRIPT SE EJECUTA DENTRO DEL ESCRITORIO VIRTUAL (Oracle Secure
Desktops), NO en tu PC real. Reproduce a mano alzada el flujo descrito en
"RPA GAP.docx":

1. Asume que SDAPeru YA esta abierto y con sesion iniciada, mostrando el
   arbol "Gestion de Alumbrado Publico" (login/usuario/clave y la
   verificacion de Oracle son pasos manuales tuyos -- este robot nunca
   toca contrasenas).
2. Doble clic en "GAP : Extractor De Datos" > "Extractor datos SAP".
3. Si la ventana abre con datos de una extraccion anterior (en vez del
   formulario en blanco), la cierra (X) y la vuelve a abrir -- tal como
   describe la OBSERVACION del Word.
4. "Abrir Consulta" -> selecciona REPORTE_AP -> Abrir.
5. Pone las fechas "Desde"/"Hasta" (mismo criterio que el robot de SAP:
   ultimos RANGO_DIAS dias hasta hoy, formato DD/MM/AAAA).
6. Click en "Leer", espera a que cargue, click en "Extraer".
7. En el dialogo "Grabar Archivo" (que ya abre por defecto en la carpeta
   OneDrive "GAP" compartida), escribe un nombre con fecha/hora y Guardar.
8. Espera a que termine y cierra la ventana del Extractor (Salir), dejando
   SDAPeru otra vez en el arbol -- listo para la proxima corrida.

COMO SE USA
-----------
1. Abre el escritorio virtual y entra a SDAPeru con tu usuario/clave (a
   mano, como siempre).
2. Deja visible la pantalla del arbol "Gestion de Alumbrado Publico".
3. Doble clic en Extraer-GAP-VM.bat (en esta misma carpeta, DENTRO de la
   VM).
4. Espera a que el programa avise "Listo: ...". El Excel queda en la
   carpeta OneDrive "GAP" -- se sincronizara solo a tu PC real. De ahi,
   corre (en tu PC real) Mover-Excel-GAP-PC.bat para llevarlo a
   CARGA\\EXCEL.

NOTA IMPORTANTE: los selectores estan escritos a partir de capturas de
pantalla del Word, no de una sesion en vivo dentro de la VM -- es muy
probable que la primera corrida real necesite ajustar 1-2 cosas (un boton
o campo que la ventana llame distinto). Cuando algo falla, el programa
guarda automaticamente una foto de pantalla en una carpeta "diagnosticos"
(dentro de GAP_RPA_Excel, en tu carpeta de usuario DENTRO de la VM).
Mandame esa foto y lo corregimos, igual que hicimos con el de SAP.
"""

from __future__ import annotations

import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

try:
    from pywinauto import Application, Desktop
    from pywinauto.controls.common_controls import TreeViewWrapper
    from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
    from pywinauto.timings import TimeoutError as PywinautoTimeoutError
except ImportError:
    print("Falta instalar pywinauto. Abre una consola (CMD) en esta carpeta")
    print("DENTRO de la VM y ejecuta:")
    print("    pip install pywinauto")
    sys.exit(1)

# (2026-09-03) Marca de version -- se imprime apenas arranca el programa
# (ver bloque __main__ mas abajo) para poder confirmar de un vistazo,
# mirando la consola, que se esta corriendo ESTA version del script y no
# una copia vieja -- mismo patron que se agrego en descargar_excel_sap.py
# despues de perder tiempo sin poder distinguir, a partir de un log, si la
# usuaria estaba corriendo el fix nuevo o uno viejo.
VERSION_SCRIPT = "2026-09-03 fix23 (modo 3 meses: cierra el 'AVISO OK' de exportacion -que aparece con retraso- antes de la siguiente ventana)"


# ----------------------------------------------------------------------
# Configuracion basica (mismo patron que descargar_excel_sap.py)
# ----------------------------------------------------------------------

RANGO_DIAS = 30  # ultimos N dias desde hoy -- cambiar aqui si hace falta otro default
NOMBRE_CONSULTA = "REPORTE_AP"


def ventanas_meses(n):
    """n ventanas de 30 dias contiguas hacia atras desde hoy (mismo criterio
    que el robot de SAP): m0 = [hoy-30, hoy], m1 = [hoy-61, hoy-31], ...
    Devuelve [(desde, hasta, slot), ...] con fechas en formato DD/MM/AAAA."""
    hoy = datetime.now()
    ventanas = []
    for k in range(n):
        hasta_dt = hoy - timedelta(days=k * 31)
        desde_dt = hasta_dt - timedelta(days=RANGO_DIAS)
        ventanas.append((desde_dt.strftime("%d/%m/%Y"), hasta_dt.strftime("%d/%m/%Y"), f"m{k}"))
    return ventanas

DIAG_BASE_DIR = Path.home() / "GAP_RPA_Excel"
DIAGNOSTICS_DIR = DIAG_BASE_DIR / "diagnosticos"

TITULO_VENTANA_PRINCIPAL = "Sistema de Distribuci"  # match parcial (acentos)
TITULO_EXTRACTOR = "Extractor de Datos SAP"
TITULO_LEER_ARCHIVO = "Leer Archivo"
TITULO_GRABAR_ARCHIVO = "Grabar Archivo"

ESPERA_CORTA = 1.5
ESPERA_LARGA_MAX = 120  # segundos maximos esperando a que "Leer" termine

# Slot de la ventana que se esta extrayendo ahora mismo ("m0"/"m1"/... en modo
# 3 meses, None en modo clasico). Lo usa click_extraer_y_guardar para nombrar
# el archivo. Se fija en correr(); modulo-global para no tener que pasarlo por
# el wrapper _paso().
_SLOT_ACTUAL = None


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------

def guardar_diagnostico(etiqueta: str, texto_extra: str | None = None, error: Exception | None = None) -> None:
    """Guarda una foto de pantalla completa para poder diagnosticar un
    fallo a distancia (igual que hace el robot de SAP con HTML).

    IMPORTANTE (2026-09-03, mismo aprendizaje que en descargar_excel_sap.py):
    la foto y el .txt de texto_extra se guardan cada uno en su PROPIO
    try/except, independiente uno del otro -- si uno falla (por ejemplo, no
    se pudo tomar la captura), el otro se guarda igual. Ademas, si se pasa
    'error' (una excepcion de Python), se guarda un .txt CHIQUITO adicional
    con el tipo y mensaje del error en si -- ese es el que SIEMPRE tiene que
    quedar, pase lo que pase con la foto o con texto_extra, porque es lo
    minimo indispensable para poder diagnosticar a distancia."""
    try:
        DIAGNOSTICS_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        print(f"  (no se pudo preparar la carpeta de diagnosticos: {exc})")
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    try:
        from PIL import ImageGrab

        ruta = DIAGNOSTICS_DIR / f"error_{etiqueta}_{ts}.png"
        ImageGrab.grab().save(ruta)
        print(f"  (diagnostico guardado en: {ruta})")
    except Exception as exc:  # PIL puede faltar o fallar el grab
        print(f"  (no se pudo guardar la foto de diagnostico: {exc})")

    if texto_extra:
        try:
            ruta_txt = DIAGNOSTICS_DIR / f"error_{etiqueta}_{ts}.txt"
            ruta_txt.write_text(texto_extra, encoding="utf-8")
            print(f"  (detalle guardado en: {ruta_txt})")
        except Exception as exc:
            print(f"  (no se pudo guardar el detalle de texto: {exc})")

    if error is not None:
        try:
            ruta_err = DIAGNOSTICS_DIR / f"error_{etiqueta}_{ts}_excepcion.txt"
            ruta_err.write_text(
                f"Etapa: {etiqueta}\nError: {type(error).__name__}: {error}",
                encoding="utf-8",
            )
            print(f"  (detalle del error guardado en: {ruta_err})")
        except Exception as exc:
            print(f"  (no se pudo guardar el .txt del error: {exc})")


def _normalizar_texto(v) -> str:
    """Quita acentos, espacios de mas y diferencias de mayusculas/minusculas,
    para comparar texto de la interfaz sin que un detalle asi haga fallar la
    busqueda."""
    import unicodedata

    v = (str(v) if v is not None else "").strip()
    v = unicodedata.normalize("NFKD", v)
    v = "".join(c for c in v if not unicodedata.combining(c))
    return " ".join(v.split()).casefold()


def _buscar_item_arbol(arbol, texto_buscado):
    """Recorre TODO el arbol (sin importar cual sea la raiz real -- por si
    tiene un nodo raiz invisible distinto al que se ve en pantalla) buscando
    el primer item cuyo texto coincida con texto_buscado, ignorando
    acentos/mayusculas/espacios. Devuelve el item (para hacerle click) o
    None si no lo encuentra. De paso arma la lista de todo lo que SI
    encontro, para poder diagnosticar si falla."""
    objetivo = _normalizar_texto(texto_buscado)
    vistos = []

    def visitar(item, nivel):
        try:
            texto_item = item.text()
        except Exception:
            texto_item = None
        vistos.append("  " * nivel + (texto_item or "(sin texto)"))
        if texto_item is not None and _normalizar_texto(texto_item) == objetivo:
            return item
        try:
            item.expand()
        except Exception:
            pass
        try:
            hijos = item.children()
        except Exception:
            hijos = []
        for hijo in hijos:
            encontrado = visitar(hijo, nivel + 1)
            if encontrado is not None:
                return encontrado
        return None

    try:
        raices = arbol.roots()
    except Exception:
        raices = []
    resultado = None
    for raiz in raices:
        resultado = visitar(raiz, 0)
        if resultado is not None:
            break
    return resultado, "\n".join(vistos)


def _volcar_controles(ventana, limite=300):
    """Lista TODOS los controles (clase + texto) que cuelgan de 'ventana',
    para poder identificar a ojo cual es el arbol de verdad cuando no
    resulto ser un SysTreeView32 estandar (algunas apps viejas de VB6 usan
    un control de arbol de otro fabricante, con otro nombre de clase)."""
    lineas = []
    try:
        descendientes = ventana.descendants()
    except Exception as exc:
        return f"(no se pudo listar los controles: {exc})"
    for i, ctrl in enumerate(descendientes[:limite]):
        try:
            clase = ctrl.class_name()
        except Exception:
            clase = "?"
        try:
            texto = ctrl.window_text()
        except Exception:
            texto = ""
        try:
            habilitado = ctrl.is_enabled()
        except Exception:
            habilitado = "?"
        lineas.append(f"[{i}] clase={clase!r} texto={texto!r} habilitado={habilitado}")
    if len(descendientes) > limite:
        lineas.append(f"... ({len(descendientes) - limite} controles mas, no listados)")
    return "\n".join(lineas) if lineas else "(la ventana no tiene ningun control hijo)"


def _listar_ventanas_abiertas():
    """Lista TODAS las ventanas de nivel superior que hay abiertas en el
    escritorio ahora mismo (titulo + clase + si esta visible/habilitada) --
    sirve para diagnosticar si el doble clic SI abrio algo, pero con otro
    titulo al que el programa esta esperando, o si quedo una ventana vieja
    tapando/bloqueando la nueva."""
    lineas = []
    try:
        ventanas = Desktop(backend="win32").windows()
    except Exception as exc:
        return f"(no se pudo listar las ventanas abiertas: {exc})"
    for w in ventanas:
        try:
            titulo = w.window_text()
        except Exception:
            titulo = "?"
        try:
            clase = w.class_name()
        except Exception:
            clase = "?"
        try:
            visible = w.is_visible()
        except Exception:
            visible = "?"
        lineas.append(f"titulo={titulo!r} clase={clase!r} visible={visible}")
    return "\n".join(lineas) if lineas else "(no hay ninguna ventana de nivel superior abierta)"


def conectar_ventana_principal():
    """Busca la ventana ya abierta de SDAPeru (NO la abre ni inicia
    sesion -- eso es manual)."""
    try:
        ventana = Desktop(backend="win32").window(title_re=f".*{TITULO_VENTANA_PRINCIPAL}.*")
        ventana.wait("exists visible ready", timeout=10)
        return ventana
    except (ElementNotFoundError, PywinautoTimeoutError):
        print("ERROR: No encuentro la ventana 'Sistema de Distribucion - Modulos")
        print("Locales' abierta. Abre SDAPeru e inicia sesion primero (a mano),")
        print("deja visible el arbol 'Gestion de Alumbrado Publico', y vuelve a")
        print("correr este programa.")
        raise SystemExit(1)


def abrir_extractor_datos_sap(ventana_principal) -> None:
    """Doble clic en el nodo del arbol 'Extractor datos SAP'. Si la
    ventana que abre no es el formulario en blanco (porque quedo abierta
    con la ultima extraccion), la cierra y reintenta -- tal como describe
    la OBSERVACION del Word."""
    TEXTO_NODO = "Extractor datos SAP"

    # Importante (2026-09-02): el doble clic se hace con COORDENADAS DE
    # PANTALLA. Si la ventana de SDAPeru no esta al frente (por ejemplo,
    # porque quedo la consola negra o el Explorador de "diagnosticos"
    # tapandola encima -- como suele pasar mientras se esta revisando algo),
    # el clic cae sobre lo que SI este al frente en ese punto de la
    # pantalla, no sobre el arbol, aunque el rectangulo calculado sea
    # correcto (por eso podia salir siempre el mismo rectangulo y aun asi
    # fallar). Por eso, antes de cada intento, forzamos que la ventana
    # principal quede al frente.
    try:
        if ventana_principal.is_minimized():
            ventana_principal.restore()
    except Exception:
        pass
    try:
        ventana_principal.set_focus()
    except Exception as exc:
        print(f"  (aviso: no se pudo poner la ventana de SDAPeru al frente: {exc})")
    time.sleep(ESPERA_CORTA / 2)

    # (2026-09-18) BUG real encontrado: si una corrida anterior se corto a
    # mitad (por ejemplo la VM se quedo sin responder) y dejo abierta una
    # ventana vieja de "Extractor de Datos SAP" sin cerrar, el doble clic de
    # aca abre OTRA ventana nueva con el MISMO titulo -- y la busqueda de mas
    # abajo (Desktop().window(title=TITULO_EXTRACTOR)) revienta con
    # "ElementAmbiguousError: hay 2 elementos que matchean" porque no sabe
    # cual de las dos usar. Se cierran de entrada todas las que ya existan
    # ANTES de hacer el doble clic, para asegurar que despues quede una sola.
    try:
        viejas = Desktop(backend="win32").windows(title=TITULO_EXTRACTOR, visible_only=True)
    except Exception:
        viejas = []
    for vieja in viejas:
        try:
            print("  (habia una ventana vieja de 'Extractor de Datos SAP' abierta de antes -- cerrandola...)")
            vieja.close()
            time.sleep(ESPERA_CORTA)
        except Exception as exc:
            print(f"  (aviso: no se pudo cerrar la ventana vieja del Extractor: {exc})")

    for intento in range(2):
        try:
            if ventana_principal.is_minimized():
                ventana_principal.restore()
            ventana_principal.set_focus()
            time.sleep(ESPERA_CORTA / 3)
        except Exception:
            pass
        # El arbol es un SysTreeView32 nativo -- en el backend "win32" eso
        # se busca por class_name, NO por control_type="TreeView" (esa
        # propiedad es de UI Automation y no existe para este tipo de
        # control, por eso nunca lo encontraba). Los items de ese arbol
        # tampoco son ventanas hijas buscables por separado -- se acceden
        # SIEMPRE a traves del propio control del arbol, nunca con
        # child_window. En vez de armar una ruta exacta desde la raiz
        # (fragil: basta que la raiz real tenga un nombre distinto al que
        # se ve en pantalla, o un espacio de mas, para que falle entero),
        # recorremos TODO el arbol buscando el texto -- ver _buscar_item_arbol.
        #
        # La ventana puede tener MAS DE UN control de arbol (paneles
        # internos, controles ocultos, etc.) -- pedir "el" (singular) con
        # child_window puede devolver uno vacio si hay varios. Por eso
        # probamos con TODOS los que haya, hasta encontrar el item en
        # alguno. Esta app (SDAPeru, VB6) NO usa el SysTreeView32 estandar
        # de Windows -- su arbol tiene la clase "TreeViewWndClass" (se
        # descubrio viendo el volcado de controles) -- pero por dentro sigue
        # respondiendo a los mismos mensajes de Windows que un TreeView
        # comun, asi que igual se puede envolver como TreeViewWrapper a la
        # fuerza (TreeViewWrapper(...) en vez de dejar que pywinauto elija
        # el wrapper solo por el nombre de clase, que es lo que hace
        # child_window/descendants normalmente).
        try:
            controles_arbol = ventana_principal.descendants(class_name="SysTreeView32")
            controles_arbol += ventana_principal.descendants(class_name="TreeViewWndClass")
        except Exception:
            controles_arbol = []

        arboles = []
        for ctrl in controles_arbol:
            try:
                arboles.append(TreeViewWrapper(ctrl.handle))
            except Exception:
                pass

        nodo = None
        listados = []
        for i, arbol in enumerate(arboles):
            try:
                nodo, listado_arbol = _buscar_item_arbol(arbol, TEXTO_NODO)
            except Exception:
                nodo, listado_arbol = None, "(error leyendo este arbol)"
            listados.append(f"--- Arbol #{i + 1} de {len(arboles)} ---\n{listado_arbol}")
            if nodo is not None:
                break
        if listados:
            listado_arbol = "\n\n".join(listados)
        else:
            # No hay ningun SysTreeView32 -- puede que esta app use un
            # control de arbol de otro fabricante (comun en apps viejas de
            # VB6). Volcamos TODOS los controles de la ventana para poder
            # identificar a ojo cual es el arbol de verdad.
            listado_arbol = (
                "No se encontro ningun control de clase SysTreeView32 en la "
                "ventana. Puede que el arbol sea un control de otro tipo. "
                "Listado de TODOS los controles de la ventana:\n\n"
                + _volcar_controles(ventana_principal)
            )

        if nodo is not None:
            # Diagnostico extra (2026-09-02): esta app usa un control de
            # arbol no estandar (TreeViewWndClass) -- no es seguro que
            # responda igual que un SysTreeView32 real a los mensajes de
            # "dame el rectangulo de este item" que pywinauto usa para saber
            # DONDE hacer clic. Por eso: primero lo seleccionamos con
            # select() (manda un mensaje directo, sin depender de
            # coordenadas en pantalla) para asegurarnos de que quede
            # resaltado, y despues del doble clic con mouse guardamos una
            # foto de inmediato (sin esperar), para poder ver que paso
            # exactamente en pantalla justo despues del clic.
            try:
                nodo.select()
            except Exception as exc:
                print(f"  (aviso: no se pudo seleccionar el item antes del clic: {exc})")
            rect_texto = "(no se pudo leer)"
            try:
                rect = nodo.client_rect()
                rect_texto = str(rect)
                print(f"  (item encontrado, rectangulo en pantalla: {rect})")
            except Exception as exc:
                print(f"  (aviso: no se pudo leer el rectangulo del item: {exc})")
            nodo.click_input(double=True)
            # Antes solo se guardaba la foto -- ahora tambien un .txt con el
            # rectangulo usado para el clic y TODAS las ventanas que hay
            # abiertas justo despues (por si el clic SI abrio algo pero con
            # otro titulo, o quedo una ventana vieja tapando la nueva).
            guardar_diagnostico(
                f"justo_despues_del_clic_intento{intento}",
                texto_extra=(
                    f"Rectangulo del item usado para el doble clic: {rect_texto}\n\n"
                    f"Ventanas abiertas justo despues del clic:\n{_listar_ventanas_abiertas()}"
                ),
            )
        else:
            guardar_diagnostico(
                "no_encontro_nodo_extractor_sap",
                texto_extra=(
                    f"Buscando: \"{TEXTO_NODO}\"\n\n"
                    f"Items que SI encontro en el arbol:\n{listado_arbol}"
                ),
            )
            raise SystemExit(
                "ERROR: no encontre el item 'Extractor datos SAP' en el arbol "
                "de 'Sistema de Distribucion - Modulos Locales'. Revisa la foto "
                "y el .txt en diagnosticos (el .txt lista todo lo que SI pudo "
                "leer del arbol, para comparar)."
            )

        time.sleep(ESPERA_CORTA * 2)

        try:
            ventana_extractor = Desktop(backend="win32").window(title=TITULO_EXTRACTOR)
            ventana_extractor.wait("exists visible ready", timeout=10)
        except ElementAmbiguousError:
            # Defensa extra: si de todos modos quedaron 2, nos quedamos con
            # la ULTIMA (la mas nueva, la que recien abrio nuestro doble
            # clic) en vez de reventar sin abrir ninguna.
            print("  (aviso: habia mas de una ventana 'Extractor de Datos SAP' -- uso la mas nueva)")
            candidatas = Desktop(backend="win32").windows(title=TITULO_EXTRACTOR, visible_only=True)
            ventana_extractor = candidatas[-1]
            ventana_extractor.wait("exists visible ready", timeout=10)
        except (ElementNotFoundError, PywinautoTimeoutError):
            guardar_diagnostico(
                "no_abrio_extractor",
                texto_extra=(
                    f"Buscando ventana con titulo exacto: \"{TITULO_EXTRACTOR}\"\n\n"
                    f"Ventanas abiertas en este momento:\n{_listar_ventanas_abiertas()}"
                ),
            )
            raise SystemExit(
                "ERROR: no se abrio la ventana 'Extractor de Datos SAP' tras el "
                "doble clic. Revisa la foto Y el .txt en diagnosticos (el .txt "
                "lista todas las ventanas que hay abiertas -- si el Extractor "
                "SI abrio pero con otro titulo, ahi deberia aparecer)."
            )

        # Si ya trae datos de una extraccion anterior (boton "Movimientos"
        # visible/habilitado y una grilla con filas), la cerramos y
        # reabrimos para partir del formulario en blanco.
        tiene_datos_previos = False
        try:
            boton_movimientos = _buscar_boton(ventana_extractor, "Movimientos")
            if boton_movimientos is not None and boton_movimientos.is_enabled():
                tiene_datos_previos = True
        except Exception:
            pass

        if tiene_datos_previos and intento == 0:
            print("  (el Extractor abrio con datos de la corrida anterior -- cerrando y reabriendo...)")
            ventana_extractor.close()
            time.sleep(ESPERA_CORTA)
            continue

        return ventana_extractor

    raise SystemExit("ERROR: el Extractor de Datos SAP sigue mostrando datos viejos tras reintentar.")


def _confirmar_dialogo_archivo(dialogo, campo_nombre, texto_boton) -> None:
    """Confirma un dialogo 'Abrir'/'Guardar' de Windows (el mismo que usa
    cualquier programa, no algo propio de SDAPeru). Estos dialogos modernos
    (estilo Explorador) a veces no dejan encontrar el boton por su clase
    real via automatizacion -- pero SIEMPRE aceptan Enter con el foco en el
    campo del nombre, que hace exactamente lo mismo que hacer clic en el
    boton. Por eso Enter es el metodo PRINCIPAL aca, y el clic al boton
    queda solo como respaldo por si Enter no alcanza a cerrar el dialogo."""
    try:
        campo_nombre.type_keys("{ENTER}")
    except Exception as exc:
        print(f"  (aviso: fallo enviar Enter en el dialogo: {exc})")

    time.sleep(ESPERA_CORTA)
    if not dialogo.exists():
        return  # Enter ya cerro el dialogo -- listo.

    # Sigue abierto -- probamos con el boton como respaldo.
    try:
        dialogo.child_window(title=texto_boton, control_type="Button").click_input()
    except Exception as exc:
        print(f"  (aviso: no se pudo hacer clic en el boton '{texto_boton}': {exc})")


def abrir_consulta_reporte_ap(ventana_extractor) -> None:
    def _boton_abrir_consulta():
        b = ventana_extractor.child_window(title="Abrir\nConsulta", control_type="Button")
        if not b.exists():
            b = ventana_extractor.child_window(best_match="Abrir Consulta")
        return b

    # Antes de tocar nada: si quedo abierto el "AVISO OK" de la exportacion
    # anterior (modo 3 meses), tapa el Extractor y el clic en "Abrir
    # Consulta" no llega. Lo cerramos primero.
    _cerrar_aviso_ok(reintentos=10, espera=0.5)
    try:
        ventana_extractor.set_focus()
    except Exception:
        pass

    dialogo = None
    for intento in range(2):
        try:
            _boton_abrir_consulta().click_input()
        except Exception as exc:
            print(f"  (aviso: fallo el clic en 'Abrir Consulta' (intento {intento}): {exc})")
        time.sleep(ESPERA_CORTA)
        try:
            dialogo = Desktop(backend="win32").window(title=TITULO_LEER_ARCHIVO)
            dialogo.wait("exists visible ready", timeout=10)
            break
        except (ElementNotFoundError, PywinautoTimeoutError):
            dialogo = None
            if intento == 0:
                print("  ('Leer Archivo' no abrio -- cierro avisos que puedan estar tapando y reintento...)")
                _cerrar_aviso_ok(reintentos=8, espera=0.5)
                try:
                    ventana_extractor.set_focus()
                except Exception:
                    pass

    if dialogo is None:
        guardar_diagnostico("no_abrio_leer_archivo")
        raise SystemExit("ERROR: no se abrio el dialogo 'Leer Archivo'. Revisa la foto en diagnosticos.")

    campo_nombre = dialogo.child_window(class_name="Edit", found_index=0)
    campo_nombre.set_edit_text(NOMBRE_CONSULTA)
    time.sleep(ESPERA_CORTA / 2)
    _confirmar_dialogo_archivo(dialogo, campo_nombre, "Abrir")
    time.sleep(ESPERA_CORTA)


def _buscar_boton(ventana, texto: str):
    """Busca un boton por su texto, sin importar si la ventana lo expone
    con el '&' del atajo de teclado delante -- esta app SI lo hace tal
    cual (confirmado con el volcado de controles: '&Leer', '&Extraer',
    '&Salir', etc., en vez de "Leer"/"Extraer"/"Salir" a secas), por eso
    una busqueda por titulo exacto sin el '&' nunca los encontraba (y
    ademas, al no estar en un try/except, esto reventaba el programa entero
    con un error de pywinauto en vez de mostrar un mensaje claro). Prueba
    varias formas hasta que una encuentre el boton, y devuelve el control
    (SIN hacerle clic) o None si ninguna funciono."""
    intentos = [
        lambda: ventana.child_window(title=texto, control_type="Button"),
        lambda: ventana.child_window(title="&" + texto, control_type="Button"),
        lambda: ventana.child_window(title_re=r"&?" + re.escape(texto) + r"\s*$", control_type="Button"),
        lambda: ventana.child_window(best_match=texto),
    ]
    for construir in intentos:
        try:
            boton = construir()
            if boton.exists():
                return boton
        except Exception:
            continue
    return None


def _es_campo_texto(clase: str) -> bool:
    """True si 'clase' es una clase de control que se comporta como un
    campo de texto editable. Ademas del "Edit" nativo de Windows, esta app
    (SDAPeru, hecha en Visual Basic 6) usa el control propio de VB6
    "ThunderRT6TextBox" para TODOS sus campos de texto -- confirmado
    revisando el volcado de controles de la ventana 'Extractor de Datos
    SAP' (los campos de fecha ahi son justamente de esa clase, no "Edit").
    Se compara sin distinguir mayusculas por si alguna otra pantalla de la
    misma app usa una variante con otra capitalizacion."""
    c = (clase or "").casefold()
    return c == "edit" or "thundertextbox" in c or "thunderrt6textbox" in c or c.endswith("textbox")


def _buscar_campo_por_etiqueta(ventana, texto_etiqueta):
    """Busca un campo de texto que este pegado a una etiqueta (con el texto
    indicado) -- sirve para apps viejas de VB6 donde el campo no tiene
    ningun nombre/auto_id que pywinauto pueda leer, pero SI tiene un rotulo
    visible al lado, tal como se ve en pantalla. Devuelve el campo mas
    cercano (misma altura, a la derecha) a esa etiqueta, o None."""
    objetivo = _normalizar_texto(texto_etiqueta)
    try:
        controles = ventana.descendants()
    except Exception:
        return None

    etiquetas, edits = [], []
    for c in controles:
        try:
            clase = c.class_name()
        except Exception:
            clase = ""
        if _es_campo_texto(clase):
            edits.append(c)
            continue
        try:
            texto = c.window_text()
        except Exception:
            texto = ""
        texto_norm = _normalizar_texto(texto) if texto else ""
        # "in" en vez de "==": algunas etiquetas traen ':' o espacios de
        # mas al final (ej. "Total de registros a extraer :") que un
        # match exacto no perdonaria.
        if texto_norm and (objetivo in texto_norm or texto_norm in objetivo):
            etiquetas.append(c)

    if not etiquetas or not edits:
        return None

    try:
        r_etq = etiquetas[0].rectangle()
    except Exception:
        return None

    def distancia(edit):
        try:
            r = edit.rectangle()
        except Exception:
            return float("inf")
        dy = abs(r.top - r_etq.top)
        dx = r.left - r_etq.right
        if dx < -5:  # el Edit queda a la izquierda de la etiqueta -- raro
            dx += 10000
        return dy * 3 + abs(dx)

    return sorted(edits, key=distancia)[0]


def _buscar_valor_por_etiqueta(ventana, texto_etiqueta):
    """Como _buscar_campo_por_etiqueta, pero SIN limitarse a controles de
    texto editables -- para un valor de SOLO LECTURA (ej. 'Total de
    registros a extraer') es comun que la app lo muestre con un Label o
    Static en vez de un campo editable, y ese no lo reconoceria
    _es_campo_texto. Devuelve el control (de cualquier clase, menos los
    Frame -- para no terminar "encontrando" el marco contenedor) mas
    cercano a la etiqueta, o None."""
    objetivo = _normalizar_texto(texto_etiqueta)
    try:
        controles = ventana.descendants()
    except Exception:
        return None

    etiquetas, candidatos = [], []
    for c in controles:
        try:
            texto = c.window_text()
        except Exception:
            texto = ""
        texto_norm = _normalizar_texto(texto) if texto else ""
        if texto_norm and (objetivo in texto_norm or texto_norm in objetivo):
            etiquetas.append(c)
            continue
        try:
            clase = c.class_name()
        except Exception:
            clase = ""
        if "frame" not in clase.casefold():
            candidatos.append(c)

    if not etiquetas or not candidatos:
        return None

    try:
        r_etq = etiquetas[0].rectangle()
    except Exception:
        return None

    def distancia(c):
        try:
            r = c.rectangle()
        except Exception:
            return float("inf")
        dy = abs(r.top - r_etq.top)
        dx = r.left - r_etq.right
        if dx < -5:
            dx += 10000
        return dy * 3 + abs(dx)

    return sorted(candidatos, key=distancia)[0]


def llenar_fechas(ventana_extractor, desde: str, hasta: str) -> None:
    grupo = ventana_extractor.child_window(title_re="Fecha de Registro del Reclamo.*")
    campo_desde = ventana_extractor.child_window(auto_id="Desde", control_type="Edit")
    campo_hasta = ventana_extractor.child_window(auto_id="Hasta", control_type="Edit")

    if not (campo_desde.exists() and campo_hasta.exists()):
        # Alternativa 1 por posicion: dentro del grupo "Fecha de Registro
        # del Reclamo", el primer campo de texto es Desde y el segundo es
        # Hasta. IMPORTANTE: no filtramos con control_type="Edit" -- esta
        # app (VB6) usa la clase propia "ThunderRT6TextBox" para sus campos
        # de texto, no el "Edit" nativo de Windows, asi que ese filtro
        # nunca los encontraba (confirmado con el volcado de controles).
        # Pedimos TODOS los descendientes del grupo y filtramos nosotros
        # mismos por clase con _es_campo_texto (que reconoce ambas).
        if grupo.exists():
            try:
                todos = grupo.descendants()
            except Exception:
                todos = []
        else:
            todos = []
        edits = []
        for c in todos:
            try:
                clase = c.class_name()
            except Exception:
                clase = ""
            if _es_campo_texto(clase):
                edits.append(c)
        if len(edits) >= 2:
            campo_desde, campo_hasta = edits[0], edits[1]
        else:
            # Alternativa 2: buscar el campo Edit mas cercano a la etiqueta
            # de texto "Desde"/"Hasta" que se ve en pantalla (esta app no
            # expone auto_id, y el grupo "Fecha de Registro del Reclamo" a
            # veces tampoco se encuentra por su titulo exacto -- pero la
            # etiqueta de texto SI esta ahi, junto al campo, en cualquier
            # caso).
            campo_desde = _buscar_campo_por_etiqueta(ventana_extractor, "Desde")
            campo_hasta = _buscar_campo_por_etiqueta(ventana_extractor, "Hasta")
            if campo_desde is None or campo_hasta is None:
                guardar_diagnostico(
                    "no_encontre_campos_fecha",
                    texto_extra="Listado de controles de la ventana:\n\n" + _volcar_controles(ventana_extractor),
                )
                raise SystemExit(
                    "ERROR: no encontre los campos de fecha 'Desde'/'Hasta'. Revisa "
                    "la foto y el .txt en diagnosticos."
                )

    _escribir_fecha(campo_desde, desde)
    _escribir_fecha(campo_hasta, hasta)


def _escribir_fecha(campo, fecha_ddmmaaaa: str) -> None:
    """Escribe una fecha (formato "DD/MM/AAAA") en un campo de fecha con
    mascara automatica -- estos campos ponen las barras "/" ellos solos a
    medida que escribes los numeros (igual que cualquier campo de fecha de
    Windows). Si le escribimos el texto CON las barras, la propia mascara
    las duplica/desordena contra lo que ya hubiera -- eso fue justo lo que
    paso en la prueba anterior: quedo "01/0802026" en vez de "03/08/2026"
    (mezcla de lo viejo con lo nuevo). Por eso aca:
      1) Nos paramos al INICIO del campo y borramos TODO lo que ya tuviera
         con varios DELETE seguidos (de sobra, no importa si el campo
         quedaba con menos texto que eso).
      2) Escribimos SOLO los digitos (DDMMAAAA, sin las barras) -- el
         campo las pone solo, tal como haria una persona escribiendo a
         mano."""
    solo_digitos = fecha_ddmmaaaa.replace("/", "")
    try:
        campo.set_focus()
    except Exception:
        pass
    try:
        campo.click_input()
    except Exception as exc:
        print(f"  (aviso: no se pudo hacer clic en el campo antes de escribir: {exc})")
    try:
        campo.type_keys("{HOME}" + "{DELETE}" * 14)
    except Exception as exc:
        print(f"  (aviso: no se pudo limpiar el campo antes de escribir: {exc})")
    campo.type_keys(solo_digitos)


def click_leer_y_esperar(ventana_extractor) -> None:
    # OJO: el boton se llama "&Leer" (con el '&' del atajo de teclado) en
    # el volcado de controles, no "Leer" a secas -- por eso antes esto
    # fallaba con un error de pywinauto sin guardar ningun diagnostico
    # (el .child_window(title="Leer", ...) nunca lo encontraba, y no
    # estaba dentro de un try/except que avisara con un mensaje claro).
    boton_leer = _buscar_boton(ventana_extractor, "Leer")
    if boton_leer is None:
        guardar_diagnostico(
            "no_encontre_boton_leer",
            texto_extra="Listado de controles de la ventana:\n\n" + _volcar_controles(ventana_extractor),
        )
        raise SystemExit(
            "ERROR: no encontre el boton 'Leer'. Revisa la foto y el .txt en diagnosticos."
        )
    # IMPORTANTE (2026-09-02): el "Total de registros a extraer" que se ve
    # en pantalla NO es un control real -- lo confirmamos con el volcado
    # COMPLETO de la ventana (54 controles) mientras ya se veia el numero
    # en pantalla: ninguno de los 54 tenia ese texto ni el numero. Esta app
    # lo dibuja directo sobre el formulario (con Form.Print de VB6), no con
    # un Label/Edit que pywinauto pueda leer -- por eso NINGUNA busqueda
    # por etiqueta o por clase lo iba a encontrar nunca.
    #
    # En vez de eso usamos una senal que SI es un control real: el boton
    # "Extraer" esta deshabilitado mientras "Leer" esta trabajando, y se
    # habilita solo cuando termina de cargar la grilla -- eso si lo puede
    # leer pywinauto con is_enabled().
    boton_extraer = _buscar_boton(ventana_extractor, "Extraer")
    extraer_habilitado_antes = False
    if boton_extraer is not None:
        try:
            extraer_habilitado_antes = boton_extraer.is_enabled()
        except Exception:
            extraer_habilitado_antes = False

    boton_leer.click_input()

    if boton_extraer is None:
        # No lo encontramos -- no hay forma de detectar el fin, esperamos
        # un tiempo fijo de margen en vez de quedarnos sin hacer nada.
        print("  (aviso: no encontre el boton 'Extraer' para esperar a que se habilite -- espero un tiempo fijo.)")
        time.sleep(10)
        return

    if extraer_habilitado_antes:
        # "Extraer" YA estaba habilitado antes de apretar "Leer" (tipico en
        # modo 3 meses: quedo asi de la ventana anterior). Igual, al apretar
        # "Leer" el boton se DESHABILITA mientras trabaja y se vuelve a
        # habilitar al terminar -- asi que esperamos a ver ese ciclo:
        # primero que se apague, y despues (como en el caso normal) que se
        # vuelva a encender. Si nunca lo vemos apagarse en ~8s, asumimos que
        # la lectura fue instantanea/cacheada y seguimos con un margen corto.
        vio_deshabilitado = False
        t0 = time.time()
        while time.time() - t0 < 8:
            try:
                if not boton_extraer.is_enabled():
                    vio_deshabilitado = True
                    break
            except Exception:
                pass
            time.sleep(0.5)
        if not vio_deshabilitado:
            print("  ('Extraer' nunca se deshabilito tras 'Leer' -- lectura instantanea, sigo con margen corto.)")
            time.sleep(4)
            return
        print("  ('Leer' esta trabajando ('Extraer' se deshabilito) -- espero a que termine...)")
        # cae al bucle normal de abajo, que espera a que se vuelva a habilitar.

    inicio = time.time()
    while time.time() - inicio < ESPERA_LARGA_MAX:
        try:
            if boton_extraer.is_enabled():
                time.sleep(1)  # margen chiquito para que termine de pintar la grilla
                return
        except Exception:
            pass
        time.sleep(1)

    # No detectamos que se habilitara a tiempo -- guardamos diagnostico
    # mas NO frenamos el programa entero: seguimos igual (click_extraer_y_
    # guardar intentara el clic de todas formas; si de verdad no estaba
    # listo, el paso siguiente -- esperar el dialogo "Grabar Archivo" --
    # va a fallar con su propio mensaje claro, en vez de quedarnos
    # atascados aca para siempre por una deteccion que puede estar mal).
    guardar_diagnostico(
        "leer_no_termino",
        texto_extra="Listado de controles de la ventana:\n\n" + _volcar_controles(ventana_extractor),
    )
    print(
        "  (aviso: 'Extraer' no parece haberse habilitado a tiempo -- "
        "sigo igual por si la deteccion fallo; si no arranca bien, revisa "
        "diagnosticos.)"
    )


def click_extraer_y_guardar(ventana_extractor) -> str:
    boton_extraer = _buscar_boton(ventana_extractor, "Extraer")
    if boton_extraer is None:
        guardar_diagnostico(
            "no_encontre_boton_extraer",
            texto_extra="Listado de controles de la ventana:\n\n" + _volcar_controles(ventana_extractor),
        )
        raise SystemExit(
            "ERROR: no encontre el boton 'Extraer'. Revisa la foto y el .txt en diagnosticos."
        )
    boton_extraer.click_input()
    time.sleep(ESPERA_CORTA)

    try:
        dialogo = Desktop(backend="win32").window(title=TITULO_GRABAR_ARCHIVO)
        dialogo.wait("exists visible ready", timeout=15)
    except (ElementNotFoundError, PywinautoTimeoutError):
        guardar_diagnostico("no_abrio_grabar_archivo")
        raise SystemExit("ERROR: no se abrio el dialogo 'Grabar Archivo'. Revisa la foto en diagnosticos.")

    nombre_archivo = f"R_{_SLOT_ACTUAL}_{datetime.now():%d.%m.%H.%M.%S}" if _SLOT_ACTUAL else f"R_{datetime.now():%d.%m.%H.%M}"
    campo_nombre = dialogo.child_window(class_name="Edit", found_index=0)
    campo_nombre.set_edit_text(nombre_archivo)
    time.sleep(ESPERA_CORTA / 2)
    _confirmar_dialogo_archivo(dialogo, campo_nombre, "Guardar")

    # Espera a que el dialogo se cierre solo (indica que termino de grabar).
    inicio = time.time()
    while time.time() - inicio < ESPERA_LARGA_MAX:
        if not dialogo.exists():
            break
        time.sleep(2)
    else:
        guardar_diagnostico("grabar_no_termino")
        raise SystemExit("ERROR: 'Grabar Archivo' no parece haber terminado. Revisa la foto en diagnosticos.")

    # El "AVISO OK" de exportacion exitosa aparece con RETRASO -- si no lo
    # cerramos, queda tapando el Extractor y la ventana SIGUIENTE (modo 3
    # meses) no puede hacer "Abrir Consulta". Insistimos hasta ~15s.
    if not _cerrar_aviso_ok(reintentos=25, espera=0.6):
        print("  (no aparecio el aviso de exportacion todavia -- se revisara de nuevo antes del siguiente paso)")

    return nombre_archivo


def _cerrar_aviso_ok(reintentos: int = 1, espera: float = 0.5) -> bool:
    """Cierra el mensaje 'AVISO OK' que aparece justo despues de exportar
    ("Los datos de la grilla fueron exportados satisfactoriamente...") -- si
    se queda abierto, tapa la ventana del Extractor y NINGUN clic siguiente
    (Salir, Abrir Consulta, etc.) le llega. IMPORTANTE (2026-09-03): ese
    aviso aparece con RETRASO (unos segundos despues de que 'Grabar Archivo'
    se cierra), asi que un solo intento inmediato muchas veces no lo
    encuentra. Por eso ahora se puede reintentar: reintentos x espera
    segundos, cerrando cualquier aviso que aparezca en esa ventana.
    Devuelve True si en algun momento cerro uno."""
    cerro_alguno = False
    libres_seguidas = 0
    for _ in range(max(1, reintentos)):
        encontrado = False
        try:
            dialogo = Desktop(backend="win32").window(title_re=".*AVISO.*")
            if dialogo.exists() and dialogo.is_visible():
                encontrado = True
                print("  (cerrando el aviso de exportacion exitosa...)")
                boton = _buscar_boton(dialogo, "Aceptar")
                if boton is not None:
                    boton.click_input()
                else:
                    try:
                        dialogo.type_keys("{ENTER}")
                    except Exception:
                        dialogo.close()
                cerro_alguno = True
                time.sleep(ESPERA_CORTA / 2)
        except Exception:
            pass
        if encontrado:
            libres_seguidas = 0
        else:
            libres_seguidas += 1
            # Si ya cerramos alguno y van 2 chequeos seguidos sin nada mas,
            # damos por hecho que no va a reaparecer -- no seguir esperando.
            if cerro_alguno and libres_seguidas >= 2:
                break
            time.sleep(espera)
    return cerro_alguno


def cerrar_extractor(ventana_extractor) -> None:
    # (2026-09-03) La version anterior hacia UN solo intento: click en
    # "Salir" (o .close() si no encontraba el boton) y ya -- sin verificar
    # si la ventana REALMENTE se cerro. En la prueba real la ventana se
    # quedo abierta igual (el "Listo:" final se imprimio bien, pero el
    # Extractor seguia en pantalla) -- probablemente porque, aunque
    # _buscar_boton() SI encuentra un control que matchea "Salir" (o su
    # variante con '&'), el click_input() por coordenadas de pantalla no le
    # llego de verdad (mismo tipo de problema que tuvimos hace un tiempo
    # con el doble clic en el arbol: si la ventana no esta al frente, el
    # clic cae en otro lado). Ahora: forzamos foco en la ventana antes de
    # cada intento, y REVISAMOS que la ventana haya desaparecido de
    # verdad -- si no, reintentamos (clic de nuevo, despues .close()) hasta
    # 3 veces, y si de plano no se cierra, guardamos diagnostico (no
    # frenamos el programa entero por esto -- el archivo ya quedo
    # guardado, que es lo importante; solo avisamos que quedo abierta).
    for intento in range(3):
        # Por si el aviso de "exportado satisfactoriamente" sigue abierto
        # (o aparecio recien, por ejemplo como confirmacion de "Salir") --
        # lo cerramos antes de cada intento, si no tapa la ventana del
        # Extractor y el clic no le llega a hacer nada.
        _cerrar_aviso_ok()

        if not ventana_extractor.exists():
            return  # ya se cerro -- nada mas que hacer.

        try:
            if ventana_extractor.is_minimized():
                ventana_extractor.restore()
            ventana_extractor.set_focus()
        except Exception:
            pass
        time.sleep(ESPERA_CORTA / 3)

        try:
            boton_salir = _buscar_boton(ventana_extractor, "Salir")
            if boton_salir is not None:
                boton_salir.click_input()
            else:
                ventana_extractor.close()
        except Exception:
            try:
                ventana_extractor.close()
            except Exception:
                pass

        time.sleep(ESPERA_CORTA)
        _cerrar_aviso_ok()  # el clic en Salir puede abrir OTRO aviso de confirmacion

        if not ventana_extractor.exists():
            return

    # Se agotaron los 3 intentos y la ventana sigue abierta -- no es grave
    # (el Excel ya se guardo antes de llegar aca), pero lo dejamos anotado
    # para poder revisar por que "Salir" no la cierra.
    guardar_diagnostico(
        "extractor_no_se_cerro",
        texto_extra="Listado de controles de la ventana:\n\n" + _volcar_controles(ventana_extractor),
    )
    print(
        "  (aviso: la ventana del Extractor sigue abierta tras varios intentos de "
        "'Salir' -- el archivo ya se guardo, pero revisa diagnosticos para ver por "
        "que no se cerro sola. La proxima corrida igual la detecta y la cierra "
        "antes de empezar, como siempre.)"
    )


# ----------------------------------------------------------------------
# Flujo principal
# ----------------------------------------------------------------------

def _paso(nombre_paso: str, funcion, *args):
    """Corre un paso del flujo (una de las funciones de mas abajo) y, si
    revienta con CUALQUIER excepcion que no sea un SystemExit ya controlado
    (esas ya guardan su propio diagnostico especifico y su propio mensaje
    claro), guarda un diagnostico generico ANTES de dejarla pasar -- mismo
    aprendizaje que en descargar_excel_sap.py: un error de pywinauto no
    previsto (por ejemplo, una ventana que se cerro justo en mal momento, o
    un control que cambio de golpe) no puede terminar en un traceback crudo
    sin ninguna foto ni .txt guardados, porque eso deja sin forma de
    diagnosticar a distancia. Con SystemExit no hace falta nada mas: la
    funcion que lo lanzo ya guardo su propia foto/txt mas especifico."""
    try:
        return funcion(*args)
    except SystemExit:
        raise
    except Exception as exc:
        guardar_diagnostico(f"error_inesperado_en_{nombre_paso}", error=exc)
        raise SystemExit(
            f"ERROR inesperado en el paso '{nombre_paso}': {type(exc).__name__}: {exc}\n"
            "Revisa la foto y el .txt en diagnosticos."
        ) from exc


def _extraer_una_ventana(ventana_principal, desde: str, hasta: str, slot, ventana_extractor=None):
    """Hace UNA extraccion. Si 'ventana_extractor' viene dado, REUSA esa
    ventana ya abierta (no vuelve a abrir el Extractor -- eso evita el
    abrir/cerrar/reabrir 'en vano' que hace abrir_extractor_datos_sap cuando
    la ventana trae datos de la vuelta anterior). Devuelve (nombre, ventana)
    -- NO cierra el Extractor: de eso se encarga correr() al final."""
    global _SLOT_ACTUAL
    _SLOT_ACTUAL = slot

    if ventana_extractor is None:
        print("Abriendo 'Extractor datos SAP'...")
        ventana_extractor = _paso("abrir_extractor_datos_sap", abrir_extractor_datos_sap, ventana_principal)
    else:
        print("Reusando la ventana del Extractor ya abierta (sin cerrar/reabrir)...")

    print(f"Abriendo consulta '{NOMBRE_CONSULTA}'...")
    _paso("abrir_consulta_reporte_ap", abrir_consulta_reporte_ap, ventana_extractor)

    print(f"Periodo = {desde} a {hasta}...")
    _paso("llenar_fechas", llenar_fechas, ventana_extractor, desde, hasta)

    print("Leyendo...")
    _paso("click_leer_y_esperar", click_leer_y_esperar, ventana_extractor)

    print("Extrayendo y guardando...")
    nombre = _paso("click_extraer_y_guardar", click_extraer_y_guardar, ventana_extractor)
    return nombre, ventana_extractor


def correr(ventanas) -> list:
    """ventanas: lista de (desde, hasta, slot). slot=None -> modo clasico (un
    solo Excel). slot='m0'/'m1'/... -> modo 3 meses (un Excel por ventana,
    con nombre R_m0_..., R_m1_..., etc.). Reusa la misma sesion de SDAPeru
    Y la misma ventana del Extractor para todas las ventanas -- solo se abre
    una vez y se cierra al final. Si una ventana falla, corta (SystemExit)
    -- la data de 3 meses tiene que estar completa."""
    print("Conectando con la ventana de SDAPeru ya abierta...")
    ventana_principal = _paso("conectar_ventana_principal", conectar_ventana_principal)

    nombres = []
    ventana_extractor = None
    for i, (desde, hasta, slot) in enumerate(ventanas, start=1):
        if len(ventanas) > 1:
            print(f"\n===== Ventana {i}/{len(ventanas)} ({slot}): {desde} a {hasta} =====")
        nombre, ventana_extractor = _extraer_una_ventana(
            ventana_principal, desde, hasta, slot, ventana_extractor)
        nombres.append(nombre)
        print(f"  guardado: {nombre}.xls")
        if i < len(ventanas):
            time.sleep(ESPERA_CORTA * 2)

    print("Cerrando el Extractor...")
    if ventana_extractor is not None:
        _paso("cerrar_extractor", cerrar_extractor, ventana_extractor)

    print(f"\nListo: se guardo(aron) {len(nombres)} archivo(s) en la carpeta OneDrive 'GAP':")
    for n in nombres:
        print(f"  {n}.xls")
    print("Se sincroniza(n) solo a tu PC real (Iniciar-GAP-PC.bat / Mover-Excel-GAP-PC.bat).")
    return nombres


def _arg_valor(nombre, default=None):
    for i, a in enumerate(sys.argv):
        if a == nombre and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
        if a.startswith(nombre + "="):
            return a.split("=", 1)[1]
    return default


if __name__ == "__main__":
    print("=== Extraer GAP (Extractor de Datos SAP) ===")
    print(f"Version del script: {VERSION_SCRIPT}")

    meses = _arg_valor("--meses")
    if meses:
        try:
            n = max(1, min(6, int(meses)))
        except ValueError:
            sys.exit("--meses necesita un numero (ej. --meses 3).")
        ventanas = ventanas_meses(n)
        print(f"Modo 3 meses: {n} ventanas de {RANGO_DIAS} dias.")
        for d, h, s in ventanas:
            print(f"  {s}: {d} a {h}")
    else:
        hasta_dt = datetime.now()
        desde_dt = hasta_dt - timedelta(days=RANGO_DIAS)
        desde = desde_dt.strftime("%d/%m/%Y")
        hasta = hasta_dt.strftime("%d/%m/%Y")
        if "--preguntar-fechas" in sys.argv:
            entrada_desde = input(f"Fecha 'desde' (DD/MM/AAAA) [{desde}]: ").strip()
            entrada_hasta = input(f"Fecha 'hasta' (DD/MM/AAAA) [{hasta}]: ").strip()
            desde = entrada_desde or desde
            hasta = entrada_hasta or hasta
        ventanas = [(desde, hasta, None)]
        print(f"Rango: {desde} a {hasta}")

    print(f"Diagnosticos en: {DIAGNOSTICS_DIR}\n")

    try:
        correr(ventanas)
    except SystemExit as exc:
        print(f"\n{exc}")
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit("\nCancelado por el usuario.")

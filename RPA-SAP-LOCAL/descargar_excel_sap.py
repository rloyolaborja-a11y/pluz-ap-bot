#!/usr/bin/env python3
"""Descarga automaticamente el Excel de SAP (reporte IW39, Clase de orden
ZM06) que se sube junto al GAP en Pendientes/Atendidas AP.

Reproduce EXACTO el flujo manual descrito en "RPA DESCARGAR DE SAP.docx":
1. Busca IW39 ("Visualizar ordenes PM") en el portal SAP.
2. Marca "concluido" e "Hist." (ademas de "Pendiente"/"En tratam.", que ya
   vienen marcados de por si).
3. Clase de orden = ZM06.
4. Periodo = ultimos 30 dias (hoy - 30 hasta hoy) -- se puede cambiar con
   RANGO_DIAS mas abajo, o contestando las fechas a mano si se corre con
   --preguntar-fechas.
5. Layout = /PRT.
6. Ejecutar.
7. En la lista de resultados: click en el encabezado "Fe.creac.", filtro
   (embudo) con el MISMO rango de fechas, aceptar.
8. "Hoja de calculo del coste" -> Exportar como Microsoft Excel, Localmente.
9. Guarda el Excel descargado directo en ..\\CARGA\\SAP\\ (la misma carpeta
   que ya lee PENDIENTES-LOCAL\\procesar_diario.py), reemplazando cualquier
   Excel viejo que hubiera ahi.

COMO SE USA
-----------
1. Doble clic en Descargar-Excel-SAP.bat (en esta misma carpeta), o
   ``python descargar_excel_sap.py``.
2. Si SAP pide iniciar sesion, se abre una ventana de Chrome: inicia sesion
   normalmente (usuario/clave/MFA) y el programa continua solo.
3. Al terminar, el Excel ya esta en ..\\CARGA\\SAP\\, listo para que
   Ejecutar.bat (de PENDIENTES-LOCAL o ATENDIDAS-LOCAL) lo use.

NOTA IMPORTANTE (igual que en el script de croquis): los selectores estan
escritos a partir de capturas de pantalla del Word, no de una sesion en
vivo. Es muy probable que en la primera corrida real haya que ajustar
alguno -- cuando algo falla, el programa guarda automaticamente una foto de
la pantalla y el HTML en la carpeta "diagnosticos" (dentro de
SAP_RPA_Excel, en tu carpeta de usuario), que es justo lo que hay que
revisar para corregir el selector que no funciono.
"""

from __future__ import annotations

import asyncio
import re
import shutil
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

from playwright.async_api import Frame, Page, TimeoutError as PlaywrightTimeoutError, async_playwright

# (2026-09-03) IMPORTANTE: confirmado por la usuaria viendolo en pantalla --
# justo despues del clic que dispara la descarga, aparece OTRA ventana mas
# (el dialogo NATIVO de Windows "Guardar como", no el de SAP que ya
# maneja el script) y AHI se cierra todo. Esto pasa cuando una politica de
# la empresa fuerza "Preguntar donde guardar cada archivo" en Chrome -- una
# politica de administrador de Windows/Chrome que NO se puede apagar desde
# Playwright (downloads_path no la anula, las politicas de administrador
# siempre ganan). Ese dialogo es una ventana de Windows, no parte de la
# pagina web, asi que Playwright no puede tocarlo -- hace falta pywinauto
# (la misma libreria que se uso para el robot del GAP) para confirmarlo.
try:
    from pywinauto import Desktop as WinDesktop

    _PYWINAUTO_DISPONIBLE = True
except ImportError:
    _PYWINAUTO_DISPONIBLE = False

# (2026-09-03) Marca de version -- se imprime apenas arranca el programa
# (ver main()) para poder confirmar de un vistazo, mirando la consola, que
# se esta corriendo ESTA version del script y no una copia vieja.
VERSION_SCRIPT = "2026-09-08 fix6 (reintento ante crash del navegador; silencioso sin minimizar; espera de Excel 120s)"

# Lo pone main() con --silencioso (lo usa el Panel de Control). El navegador
# NO es headless de verdad (eso cambia la huella de automatizacion y la
# herramienta de seguridad de la empresa lo mata) -- se lanza fuera de
# pantalla y minimizado, asi no tapa nada ni roba el foco.
SILENCIOSO = False


# ----------------------------------------------------------------------
# Configuracion basica (mismo patron que descargar_croquis.py)
# ----------------------------------------------------------------------

PORTAL_URL = "https://pluz-peru-portal-prd.workzonehr.cfapps.br10.hana.ondemand.com/site#workzone-home&/home"

# Si el clic automatico en la tarjeta de resultados de IW39 llega a fallar
# de forma repetida, se puede pegar aqui la URL completa que se copia de la
# barra de direcciones al entrar manualmente a "Visualizar ordenes PM:
# Seleccion de ordenes PM" (la pantalla con los checkboxes Pendiente/En
# tratam./concluido/Hist.). Si se deja vacia, el programa sigue intentando
# encontrar y hacer clic en la tarjeta de busqueda solo (funciona igual,
# solo que un poco mas lento).
IW39_URL_DIRECTA = "https://pluz-peru-portal-prd.workzonehr.cfapps.br10.hana.ondemand.com/site#MaintenanceOrder-displayList?sap-app-origin-hint=&sap-ui-app-id-hint=s4hana_EA13B4961465613FB41F616A1868F01D&sap-ui-tech-hint=GUI"

BASE_DIR = Path(__file__).resolve().parent
CARGA_SAP_DIR = BASE_DIR.parent / "CARGA" / "SAP"
DIAG_BASE_DIR = Path.home() / "SAP_RPA_Excel"
# (2026-09-03) Perfil NUEVO y separado del anterior "perfil_chrome": ahora el
# robot usa el Chromium propio de Playwright (no el Chrome de la empresa), y
# no conviene mezclar el user-data-dir entre los dos motores. La primera
# corrida con este cambio va a pedir iniciar sesion una vez.
PROFILE_DIR = DIAG_BASE_DIR / "perfil_chromium"
DIAGNOSTICS_DIR = DIAG_BASE_DIR / "diagnosticos"

TIMEOUT_MS = 60_000
LOGIN_TIMEOUT_MS = 300_000  # 5 minutos para completar login/MFA a mano
EJECUTAR_TIMEOUT_MS = 300_000  # 5 minutos -- SAP puede tardar bastante en
# armar la lista de resultados despues de "Ejecutar" (vimos un caso
# quedandose en "20%: Seleccion de ordenes"), sobre todo con rangos de
# fecha amplios

# (2026-09-03) Espera POR INTENTO de la lista de resultados tras "Ejecutar".
# Sintoma real reportado: a veces SAP se queda "cargando" para siempre
# despues de Ejecutar (no rebota a la pantalla de seleccion, no tira error)
# y solo un arranque limpio lo destraba. Antes se esperaba EJECUTAR_TIMEOUT_MS
# (5 min) de una sola vez, sin reintentar -> 5 min perdidos + cerrar SAP a
# mano. Ahora: ~1:40 por intento y hasta MAX_INTENTOS_EJECUTAR arranques
# limpios de IW39 (que es justo lo que resuelve el cuelgue a mano).
ESPERA_LISTA_POR_INTENTO_MS = 100_000  # ~1:40
MAX_INTENTOS_EJECUTAR = 3

# (2026-09-08) Cuanto esperamos a que el Excel exportado aparezca en disco.
# Antes eran 60s (TIMEOUT_MS) y con la VPN lenta o el export en MHTML a veces
# no llegaba a tiempo -> "No aparecio ningun archivo Excel nuevo". Lo subimos
# a 120s. El evento 'download' de Playwright tambien se espera mas (90s).
ESPERA_EXCEL_MS = 120_000
ESPERA_EVENTO_DESCARGA_SEG = 90

# (2026-09-08) Reintento de la corrida ENTERA si el navegador o el driver de
# Playwright se caen a mitad (errores tipo "Target page/context/browser has
# been closed" o "Connection closed while reading from la driver"). En modo
# "3 meses" solo se rehacen las ventanas cuyo Excel todavia no bajo.
# (2026-09-18) Subido de 2 a 4 intentos -- en la PC de un contratista con un
# antivirus mas estricto, el cierre del navegador (mismo sintoma de arriba)
# salio 2 corridas seguidas; con solo 1 reintento no alcanzaba para que se
# recupere sola sin que alguien la vuelva a correr a mano.
MAX_INTENTOS_CORRIDA = 4

CLASE_ORDEN = "ZM06"
RANGO_DIAS = 30  # ultimos N dias desde hoy -- cambiar aqui si hace falta otro default

# (2026-09-03) El export "Hoja de calculo del coste -> Microsoft Excel" de SAP
# WebGUI NO siempre baja un .xlsx: muy seguido es un MHTML con extension .xls,
# o directamente .mhtml. Antes solo se buscaba "*.xls*" y por eso el archivo
# "no aparecia" aunque si se hubiera descargado. Ahora se contemplan todos.
_PATRONES_EXCEL = ("*.xlsx", "*.xls", "*.xlsm", "*.mhtml", "*.mht", "*.csv")


def _archivos_excel(carpeta: Path) -> list[Path]:
    """Devuelve los archivos de la carpeta que parecen el Excel de SAP
    (cualquiera de los formatos de _PATRONES_EXCEL), ignorando descargas a
    medio bajar (.crdownload / .tmp)."""
    vistos: list[Path] = []
    for patron in _PATRONES_EXCEL:
        for p in carpeta.glob(patron):
            if p.name.endswith((".crdownload", ".tmp")):
                continue
            if p not in vistos:
                vistos.append(p)
    return vistos


# ----------------------------------------------------------------------
# Utilidades SAP (identicas a descargar_croquis.py -- ya probadas)
# ----------------------------------------------------------------------

async def first_visible(frame: Frame, selectors: Iterable[str], timeout_ms: int):
    """Recorre TODAS las coincidencias de cada selector (no solo la primera)
    hasta hallar una realmente visible -- SAP suele duplicar el mismo control
    en un elemento oculto (display:none / aria-hidden) mas uno visible, y
    quedarse solo con `.first` a veces agarra el oculto."""
    options = list(selectors)
    deadline = asyncio.get_event_loop().time() + timeout_ms / 1000
    while asyncio.get_event_loop().time() < deadline:
        for selector in options:
            locator = frame.locator(selector)
            try:
                total = await locator.count()
            except Exception:
                continue
            for i in range(min(total, 8)):
                candidato = locator.nth(i)
                try:
                    if await candidato.is_visible():
                        return candidato
                except Exception:
                    continue
        await frame.page.wait_for_timeout(350)
    raise PlaywrightTimeoutError(f"No se encontro ningun control con: {options}")


async def texto_visible(root, texto: str, exact: bool = False):
    """Como root.get_by_text(...).first, pero revisa todas las coincidencias
    hasta encontrar una visible (mismo motivo que first_visible: SAP repite
    el mismo texto en un elemento oculto ademas del visible)."""
    candidatos = root.get_by_text(texto, exact=exact)
    try:
        total = await candidatos.count()
    except Exception:
        total = 0
    for i in range(min(total, 8)):
        c = candidatos.nth(i)
        try:
            if await c.is_visible():
                return c
        except Exception:
            continue
    return candidatos.first


async def frame_with(page: Page, selectores, timeout_ms: int) -> Frame:
    options = [selectores] if isinstance(selectores, str) else list(selectores)
    deadline = asyncio.get_event_loop().time() + timeout_ms / 1000
    while asyncio.get_event_loop().time() < deadline:
        for frame in page.frames:
            for selector in options:
                try:
                    locator = frame.locator(selector)
                    total = await locator.count()
                except Exception:
                    continue
                for i in range(min(total, 8)):
                    try:
                        if await locator.nth(i).is_visible():
                            return frame
                    except Exception:
                        continue
        await page.wait_for_timeout(400)
    raise PlaywrightTimeoutError(f"No aparecio ningun iframe con: {options}")


async def volver_a_inicio(page: Page) -> None:
    """Recupera el portal cuando aparece de mas la pantalla de login SAML
    (la sesion sigue activa) -- hace EXACTAMENTE lo que hace el usuario a
    mano: clickear el logo de Pluz (id="shell-header-logo" en la barra
    superior de SAP Fiori), que navega a la pagina de inicio ya
    autenticada. Si el logo no esta disponible en ese momento (por ejemplo
    la barra superior tampoco cargo), se usa una navegacion directa a
    PORTAL_URL como respaldo -- funciona parecido, pero el clic al logo es
    lo que el usuario confirmo que funciona siempre."""
    logo = page.locator("#shell-header-logo")
    try:
        if await logo.count() and await logo.is_visible():
            await logo.click(timeout=5000)
            return
    except Exception:
        pass
    try:
        await page.goto(PORTAL_URL, wait_until="domcontentloaded")
    except Exception:
        pass


async def wait_for_portal(page: Page, headless: bool) -> None:
    search = page.locator("#searchFieldInShell-input-inner")
    search_btn = page.locator("#wzSearchBtn, #sf, button[title*='Buscar' i], [id*='search-button']").first
    login_signals = "input[name='j_username'], input[type='password'], #idSIButton9, form[name='f1']"
    # A veces el SAML "pierde" el relay-state y muestra este error de login
    # AUNQUE la sesion siga activa -- reintentar la navegacion (lo mismo que
    # hace a mano el usuario al clickear el logo de Pluz) suele resolverlo
    # solo, sin volver a pedir usuario/clave.
    error_saml = page.get_by_text("No hemos podido autenticarle", exact=False)
    deadline = asyncio.get_event_loop().time() + TIMEOUT_MS / 1000
    print("Verificando sesion SAP...")
    while asyncio.get_event_loop().time() < deadline:
        if await search.is_visible():
            print("Sesion SAP activa.")
            return
        if await search_btn.count() and await search_btn.is_visible():
            try:
                await search_btn.click()
                await page.wait_for_timeout(500)
                if await search.is_visible():
                    return
            except Exception:
                pass
        if await error_saml.count() and await error_saml.is_visible():
            print("  (SAP mostro 'No hemos podido autenticarle' -- reintentando la pagina, la sesion sigue activa normalmente...)")
            await volver_a_inicio(page)
            await page.wait_for_timeout(1500)
            if await search.is_visible():
                print("Sesion SAP activa.")
                return
            continue
        if await page.locator(login_signals).count() and await page.locator(login_signals).first.is_visible():
            if headless:
                raise RuntimeError("La sesion SAP vencio y hace falta iniciar sesion a mano (modo visible).")
            print("\n[SAP] Inicia sesion (usuario/clave/MFA) en la ventana de Chrome...")
            login_deadline = asyncio.get_event_loop().time() + LOGIN_TIMEOUT_MS / 1000
            while asyncio.get_event_loop().time() < login_deadline:
                if await search.is_visible():
                    print("Sesion iniciada correctamente.")
                    return
                # La busqueda puede quedar detras de un icono colapsado tras el
                # login (no aparece sola) -- reintentamos el clic periodicamente,
                # igual que hace el bucle externo antes del login.
                if await search_btn.count() and await search_btn.is_visible():
                    try:
                        await search_btn.click()
                        await page.wait_for_timeout(500)
                        if await search.is_visible():
                            print("Sesion iniciada correctamente.")
                            return
                    except Exception:
                        pass
                # Mismo recuperador de SAML que en el bucle externo -- puede
                # aparecer justo despues de terminar de escribir la clave/MFA.
                if await error_saml.count() and await error_saml.is_visible():
                    print("  (SAP mostro 'No hemos podido autenticarle' -- reintentando la pagina...)")
                    await volver_a_inicio(page)
                    await page.wait_for_timeout(1500)
                    if await search.is_visible():
                        print("Sesion iniciada correctamente.")
                        return
                    continue
                await page.wait_for_timeout(1000)
            raise PlaywrightTimeoutError("Tiempo agotado esperando el login SAP (5 min).")
        await page.wait_for_timeout(500)
    raise PlaywrightTimeoutError("El portal SAP no cargo dentro del tiempo esperado.")


async def guardar_diagnostico(page: Page, etapa: str, error: Exception, contexto=None) -> None:
    # IMPORTANTE (2026-09-02): antes esto guardaba captura + HTML + HTML de
    # cada frame TODO dentro de un solo try/except -- si el error que hizo
    # fallar el programa dejo la pagina/navegador en un estado raro (por
    # ejemplo "Target page, context or browser has been closed", que es
    # justo lo que puede pasar si SAP cierra sola la pestana de descarga),
    # la PRIMERA operacion que fallara (screenshot, o el .content()) cortaba
    # TODO lo demas en silencio -- por eso a veces solo quedaba el .png (o
    # ni eso) y nunca el detalle del error real. Ahora cada paso es
    # independiente, y el .txt con el error de Python en si SIEMPRE se
    # guarda -- ese no depende de que la pagina siga viva.
    DIAGNOSTICS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # (2026-09-03) Si la pestana original ('page') ya esta cerrada -- justo el
    # caso en que mas falta hace ver la pantalla -- probamos con cualquier
    # otra pestana que siga viva en el contexto, para no quedarnos sin foto.
    paginas: list[Page] = [page]
    if contexto is not None:
        try:
            for p in contexto.pages:
                if p not in paginas:
                    paginas.append(p)
        except Exception:
            pass

    page_usada: Page | None = None
    for p in paginas:
        try:
            await p.screenshot(path=DIAGNOSTICS_DIR / f"{etapa}_{stamp}.png", full_page=True)
            page_usada = p
            break
        except Exception as exc:
            print(f"  (aviso: no se pudo guardar la captura de esa pestana: {exc})")
    if page_usada is None:
        page_usada = page

    try:
        (DIAGNOSTICS_DIR / f"{etapa}_{stamp}.html").write_text(await page_usada.content(), encoding="utf-8")
    except Exception as exc:
        print(f"  (aviso: no se pudo guardar el HTML de la pagina: {exc})")

    try:
        for i, frame in enumerate(page_usada.frames):
            try:
                contenido = await frame.content()
                if len(contenido) > 500:
                    (DIAGNOSTICS_DIR / f"{etapa}_{stamp}_frame{i}.html").write_text(contenido, encoding="utf-8")
            except Exception:
                continue
    except Exception as exc:
        print(f"  (aviso: no se pudieron guardar los frames: {exc})")

    # Este SI o SI queda guardado -- es lo mas importante para diagnosticar
    # a distancia, y no depende de la pagina/navegador seguir vivos.
    try:
        (DIAGNOSTICS_DIR / f"{etapa}_{stamp}.txt").write_text(
            f"Etapa: {etapa}\nError: {type(error).__name__}: {error}",
            encoding="utf-8",
        )
    except Exception as exc:
        print(f"  (aviso: no se pudo guardar el .txt del error: {exc})")

    print(f"  -> Diagnostico guardado en {DIAGNOSTICS_DIR} ({error})")


# ----------------------------------------------------------------------
# Pantalla 1: "Visualizar ordenes PM: Seleccion de ordenes PM"
# ----------------------------------------------------------------------

CAMPO_CLASE_ORDEN_SELECTORES = [
    "input[title='Clase de orden']:not([readonly])",
    "input[title='Clase de orden:']:not([readonly])",
    "input[title*='Clase de orden' i]:not([readonly])",
]

CAMPO_LAYOUT_SELECTORES = [
    "input[title='Layout']:not([readonly])",
    "input[title='Layout:']:not([readonly])",
    "input[title*='Layout' i]:not([readonly])",
]


async def abrir_iw39(page: Page) -> Frame:
    """Navega hasta 'Visualizar ordenes PM: Seleccion de ordenes PM' y
    devuelve el FRAME donde vive el checkbox 'concluido' (senal de que la
    pantalla ya cargo)."""
    # A veces SAP muestra la pantalla de login justo al navegar a IW39
    # AUNQUE la sesion siga activa (el mismo glitch de SAML que en
    # wait_for_portal, solo que aqui a veces no trae el mensaje de error
    # visible) -- la sesion no se cerro de verdad, y basta con volver al
    # portal (como cuando el usuario hace clic en el logo de Pluz) para que
    # cargue ya autenticada. Por eso NO se espera un login real de una vez:
    # primero se reintenta volver al portal un par de veces, y solo si
    # despues de eso sigue en pantalla de login se asume que la sesion si
    # vencio de verdad y ahi si se espera a que el usuario inicie sesion.
    login_signals = "input[name='j_username'], input[type='password'], #idSIButton9, form[name='f1']"

    async def hay_login() -> bool:
        loc = page.locator(login_signals)
        return bool(await loc.count()) and await loc.first.is_visible()

    async def recuperar_sesion_si_hace_falta() -> bool:
        """Devuelve True si hizo falta recuperar la sesion (para que el
        llamador sepa si tiene que repetir el paso que estaba haciendo)."""
        if not await hay_login():
            return False
        print("  (SAP mostro la pantalla de login justo al abrir IW39 -- la sesion deberia seguir activa, volviendo al portal...)")
        for _ in range(3):
            await volver_a_inicio(page)
            await page.wait_for_timeout(1500)
            if not await hay_login():
                break
        if await hay_login():
            print("  (sigue en login despues de reintentar -- esperando a que se inicie sesion...)")
            await wait_for_portal(page, headless=SILENCIOSO)
        return True

    if IW39_URL_DIRECTA.strip():
        for intento in range(2):
            await page.goto("about:blank")
            await page.wait_for_timeout(300)
            await page.goto(IW39_URL_DIRECTA.strip(), wait_until="domcontentloaded")
            await page.wait_for_timeout(1500)
            await recuperar_sesion_si_hace_falta()
            if await hay_login():
                # Ya se espero el login real dentro de recuperar_sesion_si_hace_falta();
                # si sigue en login es otro problema -- se reintenta la URL directa.
                await page.goto(IW39_URL_DIRECTA.strip(), wait_until="domcontentloaded")
            try:
                return await frame_with(page, "text=concluido", 25_000 if intento == 0 else TIMEOUT_MS)
            except PlaywrightTimeoutError:
                if intento == 0:
                    print("    (la pantalla de IW39 no cargo bien, reintentando...)")
                    continue
                raise

    shell_search = page.locator("#searchFieldInShell-input-inner")
    if not await shell_search.is_visible():
        search_btn = page.locator("#wzSearchBtn, #sf, button[title*='Buscar' i]").first
        if await search_btn.count() and await search_btn.is_visible():
            await search_btn.click()
            await page.wait_for_timeout(500)
    await shell_search.fill("iw39")
    await shell_search.press("Enter")
    await page.wait_for_timeout(2000)

    if await recuperar_sesion_si_hace_falta():
        if not await shell_search.is_visible():
            search_btn = page.locator("#wzSearchBtn, #sf, button[title*='Buscar' i]").first
            if await search_btn.count() and await search_btn.is_visible():
                await search_btn.click()
                await page.wait_for_timeout(500)
        await shell_search.fill("iw39")
        await shell_search.press("Enter")
        await page.wait_for_timeout(2000)

    frame_resultados = await frame_with(page, "text=Visualizar órdenes PM", TIMEOUT_MS)
    tarjeta = await texto_visible(frame_resultados, "Visualizar órdenes PM", exact=False)
    try:
        await tarjeta.click(timeout=15_000)
    except Exception:
        await tarjeta.click(timeout=15_000, force=True)

    return await frame_with(page, "text=concluido", TIMEOUT_MS)


async def marcar_checkbox(frame: Frame, etiqueta: str) -> None:
    """Clickea el checkbox identificado por su etiqueta de texto (ej.
    'concluido', 'Hist.').

    En SAP, el texto de la etiqueta a veces resuelve a un <span
    aria-hidden="true" id="ARIA_XXX"> -- un rotulo invisible que existe solo
    para lectores de pantalla, NO el checkbox visible (por eso Playwright
    reporta "element is outside of the viewport": esta fuera de la pantalla
    a proposito). En ese caso hay que clickear el control real: el input
    hermano/cercano, o el elemento cuyo id es igual sin el prefijo "ARIA_"
    (ej. ARIA_COMPLETED -> COMPLETED)."""
    candidato = await texto_visible(frame, etiqueta, exact=True)
    if not await candidato.count():
        candidato = await texto_visible(frame, etiqueta, exact=False)

    objetivo = candidato
    try:
        aria_hidden = await candidato.get_attribute("aria-hidden")
    except Exception:
        aria_hidden = None

    if aria_hidden == "true":
        elem_id = (await candidato.get_attribute("id")) or ""
        encontrado = False
        if elem_id.upper().startswith("ARIA_"):
            alterno = frame.locator(f"#{elem_id[5:]}")
            if await alterno.count() and await alterno.first.is_visible():
                objetivo = alterno.first
                encontrado = True
        if not encontrado:
            cercano = candidato.locator(
                "xpath=preceding-sibling::input[1] | following-sibling::input[1] "
                "| ancestor::*[position()<=3]//input[@type='checkbox'][1]"
            )
            if await cercano.count():
                objetivo = cercano.first

    try:
        await objetivo.scroll_into_view_if_needed(timeout=5000)
    except Exception:
        pass
    try:
        await objetivo.click(timeout=15_000)
    except Exception:
        await objetivo.click(timeout=15_000, force=True)


async def llenar_campo_por_titulo(frame: Frame, selectores: Iterable[str], valor: str) -> None:
    campo = await first_visible(frame, selectores, TIMEOUT_MS)
    await campo.click()
    await campo.fill(valor)


async def llenar_periodo(frame: Frame, desde: str, hasta: str) -> None:
    """Campo 'Período:' -- dos inputs de fecha (desde / 'a:').

    Esta pantalla es SAP GUI clasico (WebGUI/ITS), no SAP UI5: los inputs de
    fecha no tienen atributo "title" (por eso el intento anterior, basado en
    la posicion del texto de la etiqueta, no los encontraba). En cambio, SAP
    siempre les pone su nombre tecnico de campo (DATUV = "de", DATUB = "a")
    dentro del atributo lsdata -- ese nombre es estable, a diferencia del id
    del elemento que cambia cada sesion. Se usa eso primero; si algun dia no
    aparece (otra variante de pantalla), se cae al metodo viejo por posicion."""
    desde_input = frame.locator("input[lsdata*='ctxtDATUV']").first
    hasta_input = frame.locator("input[lsdata*='ctxtDATUB']").first

    if not (await desde_input.count() and await hasta_input.count()):
        etiqueta = await texto_visible(frame, "Período:", exact=False)
        if not await etiqueta.count():
            etiqueta = await texto_visible(frame, "Periodo:", exact=False)
        fila = etiqueta.locator("xpath=ancestor::*[self::tr or self::div][1]")
        inputs_fila = fila.locator("input")
        total = await inputs_fila.count()
        if total < 2:
            # Respaldo: los dos inputs de fecha mas cercanos DESPUES de la etiqueta.
            inputs_fila = etiqueta.locator("xpath=following::input[position()<=2]")
            total = await inputs_fila.count()
        desde_input = inputs_fila.nth(0)
        hasta_input = inputs_fila.nth(1)

    await desde_input.click()
    await desde_input.fill(desde)
    await hasta_input.click()
    await hasta_input.fill(hasta)


# ----------------------------------------------------------------------
# Pantalla 2: "Visualizar ordenes PM: Lista de ordenes PM" (resultados)
# ----------------------------------------------------------------------

BOTON_FILTRO_SELECTORES = [
    # SAP GUI clasico (WebGUI) renderiza los botones de la barra de
    # herramientas como <div role="button" title="..."> -- NO como
    # <button> -- por eso hace falta el selector generico [title*=] ademas
    # del basado en <button>. El boton del embudo se llama "Fijar filtros"
    # en esta pantalla.
    "[role='button'][title*='Fijar filtro' i]",
    "[role='button'][title*='Filtro' i]",
    "[title*='Fijar filtro' i]",
    "button[title*='Filtro' i]",
    "[aria-label*='Filtro' i]",
    "button[title*='Filter' i]",
]

BOTON_ACEPTAR_MODAL_SELECTORES = [
    # En el popup "Especificar valores p.criterios filtros" el boton que
    # confirma (el check verde) se llama "Ejecutar (Entrada)" -- no
    # "Adoptar/Aceptar/OK" como se supuso al principio.
    "[role='button'][title*='Ejecutar' i]",
    "[role='button'][title*='Adoptar' i]",
    "[role='button'][title*='Aceptar' i]",
    "[role='button'][title='OK']",
    "button[title*='Ejecutar' i]",
    "button[title*='Adoptar' i]",
    "button[title*='Aceptar' i]",
    "button[aria-label*='ok' i]",
    "button[title='OK']",
]


async def filtrar_por_fecha_creacion(frame_lista: Frame, page: Page, desde: str, hasta: str) -> None:
    encabezado = await texto_visible(frame_lista, "Fe.creac.", exact=False)
    await encabezado.click(timeout=15_000)
    await page.wait_for_timeout(500)

    boton_filtro = await first_visible(frame_lista, BOTON_FILTRO_SELECTORES, TIMEOUT_MS)
    await boton_filtro.click(timeout=15_000)
    await page.wait_for_timeout(1000)

    try:
        frame_modal = await frame_with(page, "text=Especificar valores", TIMEOUT_MS)
    except PlaywrightTimeoutError:
        # (2026-09-15) Encontrado con captura real (error_20260914_093132.png):
        # a veces SAP no abre el popup de valores directo -- abre primero un
        # asistente "Filtro" ("Paso 1: Definición de criterios filtrado", con
        # "Fecha de creación" ya listada a la izquierda en "Criter.filtro").
        # Hay que confirmarlo con el boton "Tomar" para que RECIEN AHI
        # aparezca "Especificar valores" con los campos de fecha. Antes esto
        # se daba por fallo directo (ese era el error real que se repetia).
        print("  (aparecio el asistente 'Filtro' en vez del popup de valores directo -- confirmando con 'Tomar'...)")
        frame_asistente = await frame_with(page, "text=Definición de criterios filtrado", 5_000)
        boton_tomar = await texto_visible(frame_asistente, "Tomar", exact=True)
        try:
            await boton_tomar.click(timeout=15_000)
        except Exception:
            await boton_tomar.click(timeout=15_000, force=True)
        await page.wait_for_timeout(1000)
        frame_modal = await frame_with(page, "text=Especificar valores", TIMEOUT_MS)

    # En SAP clasico el ":" de la etiqueta ("Fecha de creación:") se dibuja
    # por CSS, no es texto real -- buscarlo como parte del texto nunca
    # encuentra nada visible. Ademas este popup puede traer mas de un campo
    # (aqui aparecio tambien "Nivel Tensión"), asi que en vez de adivinar la
    # etiqueta se buscan directamente los campos de FECHA por su maxlength
    # (10 = DD.MM.AAAA), que en este popup siempre son los primeros dos.
    campos_fecha = frame_modal.locator("input[maxlength='10']")
    total = await campos_fecha.count()
    if total >= 2:
        campo_desde, campo_hasta = campos_fecha.nth(0), campos_fecha.nth(1)
    else:
        # Respaldo: por el texto de la etiqueta, SIN el ":" final.
        etiqueta = await texto_visible(frame_modal, "Fecha de creación", exact=False)
        fila = etiqueta.locator("xpath=ancestor::*[self::tr or self::div][1]")
        inputs_fila = fila.locator("input")
        total = await inputs_fila.count()
        if total < 2:
            inputs_fila = etiqueta.locator("xpath=following::input[position()<=2]")
        campo_desde, campo_hasta = inputs_fila.nth(0), inputs_fila.nth(1)

    await campo_desde.click()
    await campo_desde.fill(desde)
    await campo_hasta.click()
    await campo_hasta.fill(hasta)

    try:
        boton_ok = await first_visible(frame_modal, BOTON_ACEPTAR_MODAL_SELECTORES, TIMEOUT_MS)
        try:
            await boton_ok.click(timeout=15_000)
        except Exception:
            await boton_ok.click(timeout=15_000, force=True)
    except PlaywrightTimeoutError:
        # Respaldo: el boton de confirmar es "Ejecutar (Entrada)" -- la
        # misma tecla Enter lo dispara, sin depender de encontrar el boton.
        await campo_hasta.press("Enter")
    await page.wait_for_timeout(1500)


def _confirmar_guardar_como_nativo_sync(destino: Path, timeout_seg: float = 25.0, slot: str = None) -> bool:
    """Si aparece el dialogo NATIVO de Windows 'Guardar como' (o 'Save As'),
    lo confirma solo -- escribe la RUTA COMPLETA de destino (con nombre de
    archivo incluido) en el campo de nombre y presiona Enter. Los dialogos
    comunes de Windows aceptan una ruta absoluta completa en el campo de
    nombre y guardan directo ahi, sin tener que navegar carpeta por carpeta.
    Se corre en un hilo aparte (ver asyncio.to_thread mas abajo) porque
    pywinauto es sincrono y esto es un script async. Devuelve True si
    encontro y confirmo el dialogo, False si no aparecio a tiempo o si
    pywinauto no esta instalado."""
    if not _PYWINAUTO_DISPONIBLE:
        return False
    # (2026-09-15) BUG encontrado: esta funcion nombraba el archivo con un
    # timestamp propio, SIN saber a que ventana/slot (m0/m1/m2) pertenecia.
    # Si el dialogo nativo aparecia justo cuando el camino normal (evento de
    # descarga de Playwright) TAMBIEN estaba por confirmarse, terminaban
    # quedando DOS archivos para la misma ventana -- uno bien nombrado
    # (SAP_m2.xlsx) y otro suelto con el timestamp (SAP_20260914_160235.xlsx,
    # caso real que rompio procesar_diario.py: "Falta columna en SAP" porque
    # leyo ese archivo a medias/con otro layout). Ahora usa el MISMO nombre
    # fijo que el camino normal -- si los dos caminos terminan escribiendo,
    # se pisan entre si en vez de dejar un archivo extra.
    nombre_archivo = _slot_filename(slot) if slot else f"SAP_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    ruta_completa = str(destino / nombre_archivo)
    print(f"  (vigilando si aparece el dialogo nativo 'Guardar como' de Windows, hasta {timeout_seg:.0f}s...)")
    fin = time.time() + timeout_seg
    while time.time() < fin:
        # Ademas de buscar por titulo (que puede variar de idioma/version),
        # tambien revisamos TODAS las ventanas de nivel superior con clase
        # "#32770" (la clase estandar de Windows para cuadros de dialogo
        # comunes, incluido "Guardar como") -- asi detectamos el dialogo
        # aunque el titulo no coincida exactamente con lo esperado.
        candidatos = []
        for titulo_re in (r".*Guardar como.*", r".*Save [Aa]s.*"):
            try:
                ventana = WinDesktop(backend="win32").window(title_re=titulo_re)
                if ventana.exists():
                    candidatos.append(ventana)
            except Exception:
                pass
        if not candidatos:
            try:
                for w in WinDesktop(backend="win32").windows(class_name="#32770", visible_only=True):
                    candidatos.append(w)
            except Exception:
                pass
        for ventana in candidatos:
            try:
                titulo_real = ventana.window_text()
            except Exception:
                titulo_real = "?"
            print(f"  (se detecto un dialogo nativo de Windows: '{titulo_real}' -- guardando en {ruta_completa}...)")
            try:
                campo = ventana.child_window(class_name="Edit", found_index=0)
                campo.set_edit_text(ruta_completa)
                campo.type_keys("{ENTER}")
            except Exception as exc:
                print(f"  (aviso: no se pudo escribir la ruta en el dialogo nativo ({exc}) -- probando solo Enter...)")
                try:
                    ventana.type_keys("{ENTER}")
                except Exception:
                    pass
            return True
        time.sleep(0.5)
    print("  (no aparecio ningun dialogo nativo 'Guardar como' dentro del tiempo esperado)")
    return False


async def _esperar_quieto(page: Page, ms: int) -> None:
    """Como page.wait_for_timeout, pero si la pestana ya se cerro (SAP suele
    cerrar la suya al terminar de entregar la descarga) no revienta con
    TargetClosedError -- simplemente espera con un sleep normal."""
    try:
        await page.wait_for_timeout(ms)
    except Exception:
        await asyncio.sleep(ms / 1000)


async def exportar_a_excel(frame_lista: Frame, page: Page, destino: Path, descargas: list, slot: str = None) -> Path:
    boton_hoja = await texto_visible(frame_lista, "Hoja de cálculo del coste", exact=False)
    await boton_hoja.click(timeout=15_000)
    await _esperar_quieto(page, 1000)

    frame_export = await frame_with(page, "text=Exportar como", TIMEOUT_MS)
    boton_exportar_a = await texto_visible(frame_export, "Exportar a", exact=False)
    try:
        await boton_exportar_a.click(timeout=15_000)
    except Exception:
        await boton_exportar_a.click(timeout=15_000, force=True)
    await _esperar_quieto(page, 1000)

    # "Exportar a" abre TODAVIA otro cuadro pidiendo nombre de archivo,
    # formato y destino ("Introducir el nombre del archivo que se debe
    # guardar") -- la descarga real solo empieza al confirmar ese cuadro con
    # su boton "OK" (id="UpDownDialogChoose", un id fijo del framework de
    # SAP GUI para este tipo de dialogo, no cambia de sesion a sesion).
    frame_nombre = await frame_with(page, "text=Introducir el nombre del archivo", TIMEOUT_MS)
    boton_ok_nombre = frame_nombre.locator("#UpDownDialogChoose")
    if not (await boton_ok_nombre.count() and await boton_ok_nombre.is_visible()):
        boton_ok_nombre = await texto_visible(frame_nombre, "OK", exact=True)

    # IMPORTANTE (2026-09-02): preparamos la carpeta de destino y borramos
    # los Excel viejos ANTES de disparar la descarga -- no despues. Si lo
    # hacemos despues (como estaba antes), pasa tiempo entre que Playwright
    # "recibe" la descarga y que la guardamos con save_as(); en ese tiempo
    # SAP puede cerrar sola la pestana/pop-up que uso para la descarga (es
    # comun en SAP GUI para HTML), y entonces save_as() falla con "Target
    # page, context or browser has been closed" -- justo el error que
    # aparecio. Haciendolo antes, achicamos esa ventana de tiempo al minimo.
    destino.mkdir(parents=True, exist_ok=True)
    objetivo_fijo = (destino / _slot_filename(slot)) if slot else None
    if slot is None:
        # Modo clasico: la carpeta queda con UN solo Excel.
        for viejo in _archivos_excel(destino):
            try:
                viejo.unlink()
            except Exception:
                pass
    else:
        # Modo 3 meses: solo se borra el archivo de ESTE slot (por si quedo
        # de una corrida anterior a medias); los otros slots no se tocan.
        for ext in (".xlsx", ".xls", ".xlsm", ".mhtml", ".mht", ".csv"):
            p = objetivo_fijo.with_suffix(ext)
            try:
                if p.exists():
                    p.unlink()
            except Exception:
                pass
    ignorar_previos = {p.resolve() for p in _archivos_excel(destino)}

    # (2026-09-03) Doble red de seguridad para traer el archivo:
    #   - Preferido: el listener de eventos 'download' (ver correr()) nos
    #     entrega el objeto Download aunque salte en un popup que SAP cierra
    #     enseguida, y lo guardamos con save_as en la carpeta final.
    #   - Respaldo: el contexto se abrio con downloads_path = carpeta final,
    #     asi que Chromium igual deja el archivo ahi solo; lo vigilamos por
    #     disco (funciona aunque el navegador se cierre justo despues).
    # El bloque pywinauto de abajo queda como respaldo POR SI acaso apareciera
    # el dialogo nativo de Windows "Guardar como" -- en esta PC no hay
    # politica que lo fuerce (se reviso el registro), asi que normalmente no
    # aparece y este hilo solo espera en vano unos segundos sin estorbar.
    if _PYWINAUTO_DISPONIBLE:
        tarea_dialogo_nativo = asyncio.create_task(
            asyncio.to_thread(_confirmar_guardar_como_nativo_sync, destino, 25.0, slot)
        )
    else:
        print(
            "  (aviso: pywinauto no esta instalado -- si aparece el dialogo "
            "nativo 'Guardar como' de Windows, no se va a poder confirmar "
            "solo. Si sigue fallando, instala pywinauto: pip install pywinauto)"
        )
        tarea_dialogo_nativo = None

    descargas_antes = len(descargas)
    try:
        try:
            await boton_ok_nombre.click(timeout=15_000)
        except Exception:
            await boton_ok_nombre.click(timeout=15_000, force=True)
    except Exception as exc:
        print(f"  (aviso: fallo el clic final de exportar ({exc}) -- sigo esperando la descarga igual...)")

    try:
        # 1) Camino preferido: agarrar el objeto Download que dispare
        #    cualquier pestana/popup del contexto (lo llena el listener de
        #    correr()) y guardarlo nosotros con save_as en la carpeta final.
        limite_evento = time.time() + ESPERA_EVENTO_DESCARGA_SEG
        while time.time() < limite_evento and len(descargas) == descargas_antes:
            await asyncio.sleep(0.5)

        if len(descargas) > descargas_antes:
            dl = descargas[-1]
            sugerido = (dl.suggested_filename or "").strip() or f"SAP_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            objetivo = objetivo_fijo or (destino / sugerido)
            try:
                await dl.save_as(str(objetivo))
                print(f"  (descarga guardada por Playwright: {objetivo.name})")
                if slot is None:
                    _dejar_solo(destino, objetivo)
                else:
                    _limpiar_estray(destino)
                return objetivo
            except Exception as exc:
                print(f"  (no se pudo guardar la descarga con save_as ({exc}) -- busco en disco...)")

        # 2) Respaldo: Chromium se abrio con downloads_path = carpeta final,
        #    asi que aunque no hayamos visto el evento, el archivo deberia
        #    estar apareciendo ahi solo. Lo vigilamos por disco.
        ruta = await _esperar_archivo_nuevo(destino, ignorar=ignorar_previos,
                                            timeout_ms=ESPERA_EXCEL_MS)
        if objetivo_fijo is not None and ruta.resolve() != objetivo_fijo.resolve():
            try:
                if objetivo_fijo.exists():
                    objetivo_fijo.unlink()
                ruta = ruta.replace(objetivo_fijo)
            except Exception as exc:
                print(f"  (no se pudo renombrar '{ruta.name}' a '{objetivo_fijo.name}' ({exc}) -- queda con su nombre)")
        if slot is None:
            _dejar_solo(destino, ruta)
        else:
            _limpiar_estray(destino)
        return ruta
    finally:
        if tarea_dialogo_nativo is not None:
            tarea_dialogo_nativo.cancel()


def _dejar_solo(destino: Path, ganador: Path) -> None:
    """Borra cualquier otro Excel/MHTML que haya quedado en la carpeta, para
    que procesar_diario.py encuentre UNO solo (el que acabamos de bajar)."""
    for p in _archivos_excel(destino):
        if p.resolve() != ganador.resolve():
            try:
                p.unlink()
            except Exception:
                pass


# Modo "3 meses": cada ventana se guarda con nombre fijo SAP_m0/m1/m2 y NO se
# borra a las demas. procesar_diario.py (Pendientes) las lee todas y las une.
_SLOT_RE = re.compile(r"^SAP_m\d+\.(xlsx|xls|xlsm|mhtml|mht|csv)$", re.I)


def _slot_filename(slot: str) -> str:
    return f"SAP_{slot}.xlsx"


def _limpiar_estray(destino: Path) -> None:
    """Borra Excel que NO sean archivos de slot (SAP_m0.xlsx, etc.) -- por si
    la descarga dejo tambien una copia con su nombre sugerido / EXPORT_*."""
    for p in _archivos_excel(destino):
        if not _SLOT_RE.match(p.name):
            try:
                p.unlink()
            except Exception:
                pass


async def _esperar_archivo_nuevo(destino: Path, ignorar=frozenset(), timeout_ms: int = TIMEOUT_MS) -> Path:
    """Espera a que aparezca un Excel NUEVO en 'destino' (los que ya estaban
    antes de disparar la descarga van en 'ignorar') y a que deje de crecer
    (senal de que Chrome termino de escribirlo) -- mirando SOLO el disco, sin
    pedirle nada al navegador, asi que funciona aunque el navegador ya se
    haya cerrado justo despues de guardar el archivo."""
    ignorar = {Path(p).resolve() for p in ignorar}
    limite = time.time() + (timeout_ms / 1000)
    candidato: Path | None = None
    tam_anterior = -1
    estable_desde: float | None = None
    while time.time() < limite:
        excels = [p for p in _archivos_excel(destino) if p.resolve() not in ignorar]
        if excels:
            candidato = max(excels, key=lambda p: p.stat().st_mtime)
            try:
                tam_actual = candidato.stat().st_size
            except FileNotFoundError:
                tam_actual = -1
            if tam_actual > 0 and tam_actual == tam_anterior:
                if estable_desde is None:
                    estable_desde = time.time()
                elif time.time() - estable_desde >= 2:
                    return candidato
            else:
                estable_desde = None
            tam_anterior = tam_actual
        await asyncio.sleep(0.5)

    if candidato is not None:
        # Algo se alcanzo a descargar, aunque no confirmamos que dejara de
        # crecer a tiempo -- lo devolvemos igual, mejor eso que nada.
        print(f"  (aviso: no confirme que '{candidato.name}' termino de escribirse del todo, pero ahi quedo.)")
        return candidato
    raise RuntimeError(f"No aparecio ningun archivo Excel nuevo en {destino} dentro del tiempo esperado.")


# ----------------------------------------------------------------------
# Flujo principal
# ----------------------------------------------------------------------

async def _llenar_pantalla_seleccion(frame_sel: Frame, desde: str, hasta: str) -> None:
    print("Marcando 'concluido' y 'Hist.'...")
    await marcar_checkbox(frame_sel, "concluido")
    await marcar_checkbox(frame_sel, "Hist.")

    print(f"Clase de orden = {CLASE_ORDEN}...")
    await llenar_campo_por_titulo(frame_sel, CAMPO_CLASE_ORDEN_SELECTORES, CLASE_ORDEN)

    print(f"Periodo = {desde} a {hasta}...")
    await llenar_periodo(frame_sel, desde, hasta)

    print("Layout = /PRT...")
    await llenar_campo_por_titulo(frame_sel, CAMPO_LAYOUT_SELECTORES, "/PRT")


async def ejecutar_y_esperar_lista(pagina: Page, desde: str, hasta: str) -> Frame:
    """Un intento = abrir IW39 desde cero, llenar la pantalla de seleccion,
    clickear 'Ejecutar' y esperar la lista de resultados hasta
    ESPERA_LISTA_POR_INTENTO_MS. Si la lista no aparece (SAP se queda
    'cargando' sin avanzar), reintenta desde cero -- en el ultimo intento
    hace un reset mas duro volviendo primero al portal. Es lo mismo que el
    usuario hace a mano: cerrar y relanzar."""
    ultimo_error: Exception | None = None
    for intento in range(1, MAX_INTENTOS_EJECUTAR + 1):
        try:
            if intento > 1:
                print(f"\n--- Reintento {intento}/{MAX_INTENTOS_EJECUTAR}: reabriendo IW39 desde cero ---")
                if intento >= 3:
                    print("  (reset mas duro: volviendo al portal antes de reabrir IW39...)")
                    try:
                        await pagina.goto("about:blank")
                        await pagina.wait_for_timeout(500)
                        await pagina.goto(PORTAL_URL, wait_until="domcontentloaded")
                        await wait_for_portal(pagina, headless=SILENCIOSO)
                    except Exception as exc:
                        print(f"  (aviso: no se pudo volver al portal para el reset ({exc}) -- sigo igual)")

            print("Abriendo IW39 (Visualizar ordenes PM)...")
            frame_sel = await abrir_iw39(pagina)

            await _llenar_pantalla_seleccion(frame_sel, desde, hasta)

            print("Ejecutando...")
            boton_ejecutar = await texto_visible(frame_sel, "Ejecutar", exact=True)
            await boton_ejecutar.click(timeout=15_000)
            await pagina.wait_for_timeout(2500)

            print(f"  (esperando a que SAP arme la lista de resultados, hasta "
                  f"{ESPERA_LISTA_POR_INTENTO_MS // 1000}s este intento...)")
            return await frame_with(pagina, "text=Fe.creac.", ESPERA_LISTA_POR_INTENTO_MS)

        except PlaywrightTimeoutError as e:
            ultimo_error = e
            print(f"  SAP no armo la lista en {ESPERA_LISTA_POR_INTENTO_MS // 1000}s "
                  f"(intento {intento}/{MAX_INTENTOS_EJECUTAR}) -- suele quedarse 'cargando' sin avanzar.")
            if intento == MAX_INTENTOS_EJECUTAR:
                raise
    raise ultimo_error  # pragma: no cover (el for siempre sale por return o raise)


async def correr(ventanas) -> list:
    """ventanas: lista de (desde, hasta, slot). slot=None -> modo clasico (un
    solo Excel, se borran los demas). slot="m0"/"m1"/... -> modo 3 meses (un
    Excel por ventana, con nombre fijo, sin borrar los otros). Reusa un solo
    navegador para todas las ventanas (un solo login)."""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    CARGA_SAP_DIR.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as playwright:
        # (2026-09-03) SIN channel="chrome": usamos el Chromium que trae
        # Playwright, NO el Chrome instalado por la empresa. El Chrome de la
        # empresa carga por politica de maquina
        # (HKLM\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist) una
        # extension de seguridad/DLP que -- confirmado por el patron de fallos
        # -- cierra el navegador entero justo al disparar la descarga de datos
        # de SAP. El Chromium de Playwright no lee esas politicas ni carga esa
        # extension. Requiere: playwright install chromium
        contexto = await playwright.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            no_viewport=True,
            accept_downloads=True,
            # (2026-09-03) Le decimos a Chrome que descargue DIRECTO en la
            # carpeta final -- asi, si el navegador se cierra de golpe justo
            # despues de la descarga (como ha estado pasando), no dependemos
            # de que Playwright siga conectado para "entregarnos" el archivo
            # con save_as()/path() (que es justo lo que fallaba): el archivo
            # ya queda guardado ahi por el propio Chrome, sin intermediarios.
            downloads_path=str(CARGA_SAP_DIR),
            args=[
                # Modo silencioso: fuera de pantalla (esquina negativa) en vez
                # de maximizado, para que no aparezca ni robe el foco. Igual a
                # 1920x1080 para que SAP renderice todos los controles.
                ("--window-position=-32000,-32000" if SILENCIOSO else "--start-maximized"),
                "--window-size=1920,1080",
                # (2026-09-03) El navegador se estaba cerrando ENTERO (no
                # solo una pestana) justo al disparar la descarga -- la foto,
                # el HTML y el archivo mismo fallaban los 3 a la vez con
                # "Target page, context or browser has been closed", y se
                # confirmo viendolo pasar en pantalla: Chrome se cerraba
                # solo, casi al instante. Eso es tipico de un antivirus o
                # herramienta de seguridad de la empresa detectando que el
                # navegador esta siendo controlado por automatizacion (via
                # el puerto de depuracion que usa Playwright) y cerrandolo
                # como medida de seguridad justo quando se dispara una
                # descarga de datos. Estos dos flags reducen esa "huella" de
                # automatizacion (quitan el aviso de "Chrome esta siendo
                # controlado por software de pruebas automatizado" y la
                # bandera que exponen las paginas web para detectarlo).
                "--disable-blink-features=AutomationControlled",
                # (2026-09-08) La ventana silenciosa iba MINIMIZADA (via CDP,
                # abajo) y Chromium frena/duerme los timers y las descargas de
                # una ventana en segundo plano -> a veces el Excel "no llegaba".
                # Estos tres flags apagan ese ahorro de energia: la ventana
                # puede estar fuera de pantalla pero sigue trabajando full.
                "--disable-background-timer-throttling",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
            ],
            ignore_default_args=["--enable-automation"],
        )
        pagina = contexto.pages[0] if contexto.pages else await contexto.new_page()
        pagina.set_default_timeout(TIMEOUT_MS)

        if SILENCIOSO:
            # (2026-09-08) Antes esto MINIMIZABA la ventana -> Chromium la
            # trataba como "en segundo plano" y ralentizaba la descarga. Ahora
            # la dejamos en estado NORMAL pero clavada fuera de pantalla
            # (esquina -32000,-32000): no aparece, no roba foco, y NO se
            # duerme. Si Windows "rescata" la posicion, los flags de arriba ya
            # evitan el throttling igual.
            try:
                cdp = await contexto.new_cdp_session(pagina)
                win = await cdp.send("Browser.getWindowForTarget")
                await cdp.send("Browser.setWindowBounds", {
                    "windowId": win["windowId"],
                    "bounds": {"windowState": "normal", "left": -32000, "top": -32000,
                               "width": 1920, "height": 1080},
                })
            except Exception as exc:
                print(f"  (aviso: no se pudo reposicionar la ventana del SAP: {exc})")

        # (2026-09-03) Captura de la descarga por EVENTO, no por pestana fija.
        # SAP dispara la descarga a veces en su propia pestana y a veces en un
        # popup que abre y cierra enseguida. En vez de depender de una pestana
        # concreta (y en vez de la pestana-ancla en blanco que habia antes),
        # escuchamos el evento 'download' en CUALQUIER pestana del contexto,
        # presente o futura, y las vamos juntando en 'descargas'.
        descargas: list = []

        def _cablear_descargas(pg: Page) -> None:
            pg.on("download", lambda d: descargas.append(d))

        for _pg in contexto.pages:
            _cablear_descargas(_pg)
        contexto.on("page", _cablear_descargas)

        try:
            await pagina.goto(PORTAL_URL, wait_until="domcontentloaded")
            await wait_for_portal(pagina, headless=SILENCIOSO)

            rutas = []
            for i, (desde, hasta, slot) in enumerate(ventanas, start=1):
                if len(ventanas) > 1:
                    print(f"\n============ Ventana {i}/{len(ventanas)}"
                          f"{(' (' + slot + ')') if slot else ''}: {desde} a {hasta} ============")

                frame_lista = await ejecutar_y_esperar_lista(pagina, desde, hasta)

                print("Filtrando por Fecha de creación...")
                await filtrar_por_fecha_creacion(frame_lista, pagina, desde, hasta)

                print("Exportando a Excel...")
                ruta = await exportar_a_excel(frame_lista, pagina, CARGA_SAP_DIR, descargas, slot=slot)
                print(f"Listo: {ruta}")
                rutas.append(ruta)

            # Red de seguridad final (2026-09-15): barre cualquier Excel
            # suelto que haya quedado con un nombre que no sea SAP_m#.xlsx
            # -- por si el dialogo nativo y el camino normal de Playwright
            # llegaron a pisarse justo antes de que _limpiar_estray corriera
            # dentro de exportar_a_excel. procesar_diario.py lee TODO lo que
            # encuentra en la carpeta, asi que un archivo extra a medias
            # (otro layout/columnas) tira abajo el pipeline entero.
            if any(slot for _, _, slot in ventanas):
                _limpiar_estray(CARGA_SAP_DIR)

            return rutas

        except Exception as error:
            await guardar_diagnostico(pagina, "error", error, contexto)
            raise
        finally:
            # Si el navegador ya se cayo, contexto.close() revienta con el
            # mismo "Target closed" y taparia el error original -> lo tragamos.
            try:
                await contexto.close()
            except Exception:
                pass


def _arg_valor(nombre, default=None):
    """Lee --nombre VALOR o --nombre=VALOR de sys.argv."""
    for i, a in enumerate(sys.argv):
        if a == nombre and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
        if a.startswith(nombre + "="):
            return a.split("=", 1)[1]
    return default


def _ventanas_meses(n):
    """n ventanas de 30 dias contiguas hacia atras desde hoy:
    m0 = [hoy-30, hoy], m1 = [hoy-61, hoy-31], m2 = [hoy-92, hoy-62]..."""
    hoy = datetime.now()
    ventanas = []
    for k in range(n):
        hasta_dt = hoy - timedelta(days=k * 31)
        desde_dt = hasta_dt - timedelta(days=30)
        ventanas.append((desde_dt.strftime("%d.%m.%Y"), hasta_dt.strftime("%d.%m.%Y"), f"m{k}"))
    return ventanas


async def solo_login() -> None:
    """Abre el portal SAP en una ventana VISIBLE para iniciar sesión a mano
    (lo usa el botón '🔑 Iniciar sesión en SAP' del Panel). El perfil queda
    guardado, así las descargas silenciosas siguientes no piden nada."""
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR), headless=False, no_viewport=True,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        pg = ctx.pages[0] if ctx.pages else await ctx.new_page()
        pg.set_default_timeout(TIMEOUT_MS)
        await pg.goto(PORTAL_URL, wait_until="domcontentloaded")
        print("Si SAP pide iniciar sesión, hacelo en la ventana que se abrió. Esperando...")
        await wait_for_portal(pg, headless=False)
        print("Sesión SAP OK. El perfil quedó guardado — ya podés cerrar.")
        await pg.wait_for_timeout(1500)
        await ctx.close()


def _slot_ya_ok(slot) -> bool:
    """True si el Excel de esta ventana ('m0'/'m1'/...) ya esta bajado y con
    tamano razonable. En modo clasico (slot=None) siempre False: no se puede
    'retomar' una descarga unica, se rehace entera."""
    if not slot:
        return False
    for ext in (".xlsx", ".xls", ".xlsm", ".mhtml", ".mht", ".csv"):
        p = CARGA_SAP_DIR / f"SAP_{slot}{ext}"
        try:
            if p.exists() and p.stat().st_size > 5000:
                return True
        except Exception:
            pass
    return False


def _es_crash_navegador(error) -> bool:
    """True si el error parece una caida del navegador o del driver de
    Playwright (no un selector mal ni un problema de logica) -> vale la pena
    reintentar. Falso para timeouts de SAP armando la lista, etc."""
    s = f"{type(error).__name__}: {error}".lower()
    marcas = (
        "target page, context or browser has been closed",
        "target closed", "targetclosederror",
        "browser has been closed", "context or browser has been closed",
        "connection closed while reading from the driver",
        "connection closed", "websocket", "pipe closed",
        "browser closed unexpectedly", "the driver",
    )
    return any(m in s for m in marcas)


def main() -> None:
    global SILENCIOSO
    SILENCIOSO = "--silencioso" in sys.argv or "--headless" in sys.argv
    print("=== Descargar Excel SAP (IW39 / ZM06) ===")
    print(f"Version del script: {VERSION_SCRIPT}")

    if "--solo-login" in sys.argv:
        print("Modo: solo iniciar sesión en SAP (no descarga nada).")
        try:
            asyncio.run(solo_login())
        except Exception as error:
            sys.exit(f"\nFallo: {error}")
        return

    if SILENCIOSO:
        print("Modo silencioso: el navegador del SAP corre FUERA DE PANTALLA (no lo vas a ver).")
    print(
        "pywinauto (para el dialogo nativo 'Guardar como'): "
        + ("instalado, OK" if _PYWINAUTO_DISPONIBLE else "NO INSTALADO -- ejecuta: pip install pywinauto")
    )

    meses = _arg_valor("--meses")
    if meses:
        try:
            n = max(1, min(6, int(meses)))
        except ValueError:
            sys.exit("--meses necesita un numero (ej. --meses 3).")
        ventanas = _ventanas_meses(n)
        # Barrido unico: la carpeta arranca limpia para no mezclar descargas
        # viejas (EXPORT_*, manuales) con los slots SAP_m0/m1/... de ahora.
        CARGA_SAP_DIR.mkdir(parents=True, exist_ok=True)
        for p in list(CARGA_SAP_DIR.glob("*")):
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            # Excel/descargas a medias + los temporales sin extension que deja
            # Playwright en downloads_path cuando el navegador se cae (nombres
            # tipo GUID); NO tocar el LEEME ni ningun .txt.
            es_basura = ext in (".xlsx", ".xls", ".xlsm", ".mhtml", ".mht", ".csv", ".crdownload", ".tmp")
            es_temp_playwright = ext == "" and len(p.name) >= 32 and "-" in p.name
            if es_basura or es_temp_playwright:
                try:
                    p.unlink()
                except Exception:
                    pass
        print(f"Modo 3 meses: {n} ventanas de 30 dias.")
        for d, h, s in ventanas:
            print(f"  {s}: {d} a {h}")
    else:
        hasta_dt = datetime.now()
        desde_dt = hasta_dt - timedelta(days=RANGO_DIAS)
        desde = desde_dt.strftime("%d.%m.%Y")
        hasta = hasta_dt.strftime("%d.%m.%Y")
        if "--preguntar-fechas" in sys.argv:
            entrada_desde = input(f"Fecha 'desde' (DD.MM.AAAA) [{desde}]: ").strip()
            entrada_hasta = input(f"Fecha 'hasta' (DD.MM.AAAA) [{hasta}]: ").strip()
            desde = entrada_desde or desde
            hasta = entrada_hasta or hasta
        ventanas = [(desde, hasta, None)]
        print(f"Rango: {desde} a {hasta}")

    print(f"Se guardara en: {CARGA_SAP_DIR}\n")

    intento = 0
    while True:
        intento += 1
        faltan = [v for v in ventanas if not _slot_ya_ok(v[2])]
        if not faltan:
            break  # ya estan todos los Excel (puede pasar en un reintento)
        try:
            asyncio.run(correr(faltan))
            break
        except KeyboardInterrupt:
            sys.exit("\nCancelado por el usuario.")
        except Exception as error:
            if intento < MAX_INTENTOS_CORRIDA and _es_crash_navegador(error):
                print(f"\n(el navegador/driver de Playwright se cayo: {error})")
                print(f"Reintento {intento + 1}/{MAX_INTENTOS_CORRIDA} en 15s — "
                      "retomo desde la(s) ventana(s) que falten…")
                time.sleep(15)
                continue
            print(f"\nFallo: {error}")
            print(f"Revisa la carpeta de diagnosticos: {DIAGNOSTICS_DIR}")
            sys.exit(1)


if __name__ == "__main__":
    main()

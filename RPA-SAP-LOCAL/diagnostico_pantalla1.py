#!/usr/bin/env python3
"""Diagnostico visual de la Pantalla 1 de SAP (IW39) para terminar el robot
UIA (ver descargar_excel_sap_uia.py).

Por que existe esto
--------------------
El robot UIA se trabo porque los campos de texto de esta pantalla (Clase de
orden, Periodo, Layout) no tienen nombre accesible -- solo se pueden ubicar
por POSICION (su indice en la lista de todos los 'Edit' visibles). Contar esa
posicion a mano, comparando un volcado de texto contra una captura de
pantalla aparte, es lento y con el numero de campos de esta pantalla (20+)
es facil equivocarse un indice (ya paso una vez, ver notas en
descargar_excel_sap_uia.py).

Que hace este script
---------------------
1. Abre Edge NORMAL (sin CDP) en IW39, igual que el robot real.
2. Vos llegas a mano a la pantalla con los checkboxes 'concluido'/'Hist.' y
   apretas ENTER aca.
3. El script saca una foto de la pantalla y le DIBUJA ENCIMA un numero al
   lado de cada campo 'Edit' visible -- el mismo numero es su indice en la
   lista que usa el robot (INDICE_CLASE_ORDEN, etc.).
4. Guarda esa foto en SAP_RPA_Excel/diagnostico_pantalla1.png.

Con la foto abierta, solo hay que MIRAR que numero quedo al lado de 'Clase
de orden', al lado de 'Periodo' (dos campos: desde/hasta) y al lado de
'Layout' -- sin contar nada a mano. Mandale esa foto (o los 4 numeros) a
quien este terminando el robot.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

try:
    from pywinauto import Desktop
except ImportError:
    print("Falta pywinauto. Instalalo con: pip install pywinauto")
    sys.exit(1)

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Falta Pillow. Instalalo con: pip install Pillow")
    sys.exit(1)

# Reusa la logica ya probada de abrir_edge_y_conectar() del robot en
# construccion, para no duplicarla.
sys.path.insert(0, str(Path(__file__).parent))
from descargar_excel_sap_uia import (  # noqa: E402
    IW39_URL_DIRECTA,
    abrir_edge_y_conectar,
    _campos_edit_visibles,
    _con_reintento,
)

SALIDA_DIR = Path.home() / "SAP_RPA_Excel"


def _dpi_awareness() -> None:
    """Sin esto, en pantallas con escalado (125%, 150%, etc.) las
    coordenadas que da pywinauto (en pixeles 'reales') no coinciden con las
    de la foto que saca Pillow (en pixeles 'escalados') y los numeros
    quedan dibujados en el lugar equivocado."""
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    _dpi_awareness()
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)

    print("Abriendo Edge (normal, sin CDP) en IW39 directo...")
    ventana, _proceso = abrir_edge_y_conectar(IW39_URL_DIRECTA)
    titulo = _con_reintento(lambda: ventana.window_text())
    print(f"Ventana conectada: '{titulo}'")

    input(
        "\n>>> Si SAP pide login, hacelo a mano en esa ventana.\n"
        ">>> Paso 1: llega a la pantalla de IW39 con los checkboxes "
        "'concluido'/'Hist.' (no hace falta llenar nada).\n"
        ">>> Paso 2: hace clic en el campo 'Clase de orden' y apreta la tecla "
        "Tab varias veces hasta ver en pantalla 'Periodo' y 'Layout'.\n"
        ">>> Recien CUANDO YA VEAS 'Periodo' y 'Layout' en pantalla, volve "
        "aca y apreta ENTER (este es el UNICO Enter que hace falta)... "
    )

    # (2026-09-24) Al apretar ENTER en la consola, la consola queda al FRENTE
    # tapando a Edge -- y Chrome/Edge "podan" el arbol de accesibilidad de una
    # pestana que queda detras de otra ventana (para ahorrar recursos), asi
    # que UI Automation deja de ver los campos de SAP (confirmado: la primera
    # corrida encontro solo 1 campo, la barra de direcciones).
    #   Intento 1: ventana.set_focus() para traer Edge al frente -- fallo,
    #   Windows bloquea que un proceso le "robe" el foco a la consola activa.
    #   Intento 2: minimizar la consola via GetConsoleWindow() -- fallo
    #   TAMBIEN, porque con Windows Terminal (el default en Windows 11)
    #   GetConsoleWindow() devuelve el handle de una ventanita FANTASMA de
    #   compatibilidad, no la ventana visible de verdad -- confirmado, la
    #   consola siguio de pie sin minimizarse.
    #   Intento 3: GetForegroundWindow() justo despues del input() -- en ESE
    #   momento, sea cual sea el programa de terminal (cmd, Windows Terminal,
    #   PowerShell, VS Code, etc.), su ventana es garantizado la que tiene el
    #   foco. Este intento TAMBIEN fallo en silencio -- causa real: por
    #   defecto ctypes asume que toda funcion de Windows devuelve/recibe un
    #   entero de 32 bits (c_int) si no se le dice lo contrario. HWND es un
    #   puntero de 64 bits en Windows moderno -- sin declarar restype/argtypes
    #   como c_void_p, GetForegroundWindow() devolvia el handle TRUNCADO a 32
    #   bits (un numero corrupto), y ShowWindow() con ese handle invalido no
    #   hace nada (falla callado, sin excepcion). Se corrige declarando los
    #   tipos correctos antes de llamarlas.
    print("Minimizando esta terminal para que Edge quede al frente...")
    import ctypes

    SW_MINIMIZE = 6
    SW_RESTORE = 9
    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
    user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
    terminal_hwnd = user32.GetForegroundWindow()
    ok = user32.ShowWindow(terminal_hwnd, SW_MINIMIZE)
    print(f"  (hwnd terminal={terminal_hwnd}, ShowWindow devolvio={ok})")
    time.sleep(1.5)
    _con_reintento(lambda: ventana.set_focus())
    time.sleep(1.5)

    print("Levantando la lista de campos 'Edit' visibles...")
    campos = _con_reintento(lambda: _campos_edit_visibles(ventana))
    print(f"  ({len(campos)} campos encontrados)")

    print("Sacando foto de la pantalla completa...")
    from PIL import ImageGrab

    foto = ImageGrab.grab().convert("RGB")
    draw = ImageDraw.Draw(foto)

    for i, campo in enumerate(campos):
        try:
            r = campo.rectangle()
        except Exception:
            continue
        # Caja roja alrededor del campo.
        draw.rectangle([r.left, r.top, r.right, r.bottom], outline=(255, 0, 0), width=2)
        # Numero en un recuadro amarillo, arriba a la izquierda del campo,
        # para que se lea facil incluso con campos chicos o pegados.
        texto = str(i)
        pos = (r.left, max(0, r.top - 16))
        ancho_aprox = 8 * len(texto) + 4
        draw.rectangle([pos[0], pos[1], pos[0] + ancho_aprox, pos[1] + 14], fill=(255, 255, 0))
        draw.text((pos[0] + 2, pos[1] + 1), texto, fill=(0, 0, 0))

    ruta_foto = SALIDA_DIR / "diagnostico_pantalla1.png"
    foto.save(ruta_foto)

    ruta_txt = SALIDA_DIR / "diagnostico_pantalla1.txt"
    with open(ruta_txt, "w", encoding="utf-8") as f:
        for i, campo in enumerate(campos):
            try:
                r = campo.rectangle()
                f.write(f"#{i}: rect={r}\n")
            except Exception:
                f.write(f"#{i}: (sin rectangulo)\n")

    # Restauramos la terminal (la habiamos minimizado para que Edge quedara
    # al frente) para que el resultado final se vea sin tener que buscarla en
    # la barra de tareas.
    ctypes.windll.user32.ShowWindow(terminal_hwnd, SW_RESTORE)
    ctypes.windll.user32.SetForegroundWindow(terminal_hwnd)

    print(f"\nListo. Foto guardada en:\n  {ruta_foto}")
    print(f"Detalle en texto:\n  {ruta_txt}")
    print(
        "\n>>> Abri la foto y anota el numero que quedo al lado de:\n"
        "      - 'Clase de orden'\n"
        "      - 'Periodo' (dos campos: desde / hasta)\n"
        "      - 'Layout'\n"
        "    Mandanos esos 4 numeros (o la foto) para terminar el robot."
    )


if __name__ == "__main__":
    main()

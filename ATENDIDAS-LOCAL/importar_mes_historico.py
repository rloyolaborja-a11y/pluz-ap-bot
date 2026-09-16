# -*- coding: utf-8 -*-
"""
importar_mes_historico.py — Reconstruye mes(es) ya cerrados de "Atendidas AP"
con TODAS las columnas (incluida "Poste - Referencias"), para que también
se puedan ver/exportar en "Buscar y exportar" (no solo en los gráficos del
Dashboard).

Por qué hace falta esto para meses VIEJOS: desde 2026-09-02, cada corrida
normal de Ejecutar.bat ya archiva SOLA los meses que se le van cerrando
(ver historico_lib.py y procesar_diario.py) — pero eso solo cubre los
meses que hayan pasado por el GAB del día a día mientras el robot ya
estaba corriendo. Para meses de ANTES de eso, hace falta darle el Excel
(GAB) original de ese mes una sola vez, con este script.

CÓMO SE USA (2 formas, las dos hacen lo mismo)
-----------------------------------------------
A) La más simple: deja tus Excel de meses viejos sueltos dentro de la
   carpeta EXCEL-MESES-ANTERIORES\\ (al lado de este script) y solo haz
   doble clic en Importar-Mes-Historico.bat — sin arrastrar nada. Procesa
   TODOS los Excel que encuentre ahí y, a cada uno que logre importar, lo
   mueve a EXCEL-MESES-ANTERIORES\\YA-IMPORTADOS\\ para que no se vuelva a
   procesar la próxima vez (si algo sale mal con alguno, lo deja donde
   estaba para que lo revises).
B) Arrastra uno o varios Excel directo sobre Importar-Mes-Historico.bat
   (vengan de donde vengan) — se procesan esos, y NO se mueven de donde
   estaban (esto es para un uso puntual, fuera de la carpeta de arriba).

El <mes> (AAAA-MM) se calcula solo, a partir de la columna "SAP - Fecha de
Atención" de cada fila — si el Excel que le pasas trae más de un mes
mezclado, esto los separa y arma un archivo por cada uno, automáticamente.
Si un mes ya tenía datos archivados (por ejemplo, porque el archivado
automático de Ejecutar.bat ya alcanzó a cubrir parte de ese mes), esto no
los pisa: los MEZCLA (misma fila por "SAP - Número de reclamo" -> gana la
más nueva).

IMPORTANTE: el Excel de cada mes debe cubrir el MES CALENDARIO COMPLETO
(del día 1 al último día de ese mes) — no el export de 30 días del robot
del GAP, que no llega a cubrir el mes entero.

Después de importar, corre Publicar-Historico.bat (como ya lo has hecho
antes) para subir los archivos nuevos/actualizados al sitio — o espera a
la próxima vez que corras Ejecutar.bat, que también lo hace por ti.

Requiere: pandas, openpyxl (los mismos que ya necesita procesar_diario.py).
"""

import glob
import json
import os
import shutil
import sys
from collections import defaultdict
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORICO_DIR = os.path.join(BASE_DIR, "HISTORICO")
MAESTROS_DIR = os.path.join(BASE_DIR, "MAESTROS")
ENTRADA_DIR = os.path.join(BASE_DIR, "EXCEL-MESES-ANTERIORES")
YA_IMPORTADOS_DIR = os.path.join(ENTRADA_DIR, "YA-IMPORTADOS")

sys.path.insert(0, BASE_DIR)
from procesar_diario import procesar, leer_gab, id_login_de_contratista, cargar_correcciones_validadas  # noqa: E402
import historico_lib  # noqa: E402


def cargar_maestros():
    faltan = [n for n in ("tecnicos.json", "feriados.json")
              if not os.path.exists(os.path.join(MAESTROS_DIR, n))]
    if faltan:
        print(f"Faltan archivos en MAESTROS\\: {', '.join(faltan)}")
        sys.exit(1)
    maestros = {}
    for clave, nombre in (("tecnicos", "tecnicos.json"), ("feriados", "feriados.json")):
        with open(os.path.join(MAESTROS_DIR, nombre), encoding="utf-8") as f:
            maestros[clave] = json.load(f)
    return maestros


def cargar_correcciones():
    """Igual que en procesar_diario.py — las correcciones manuales validadas
    también se aplican al re-importar meses viejos. Si falla, se aborta."""
    try:
        with open(os.path.join(BASE_DIR, "config.json"), encoding="utf-8") as f:
            script_url = json.load(f)["script_url"]
        corr = cargar_correcciones_validadas(script_url)
        print(f"Correcciones manuales validadas: {len(corr)}")
        return corr
    except Exception as e:
        print(f"NO se pudieron leer las correcciones manuales: {e}")
        print("Se aborta (no se importa nada) para no perder correcciones ya hechas.")
        sys.exit(1)


def agrupar_por_mes(rows_completas):
    """{'2026-01': [filas completas], ...}"""
    por_mes = defaultdict(list)
    for r in rows_completas:
        dia = r.get("Dia Atención")
        if not dia or not isinstance(dia, str) or len(dia) < 7:
            continue
        por_mes[dia[:7]].append(r)
    return por_mes


def importar_archivo(gab_path, maestros, correcciones=None):
    print(f"\n=== Procesando: {os.path.basename(gab_path)} ===")
    try:
        _rows_publicas, rows_completas, _ahora, stats = procesar(leer_gab(gab_path), maestros, correcciones)
    except Exception as e:
        print(f"  ERROR procesando este archivo: {e}")
        return False
    print(f"  Filas: total={stats['total']} -> ATENDIDA={stats['tras_estado']} -> "
          f"motivo válido={stats['tras_motivo']} -> únicas={stats['tras_dedup']} -> final={stats['final']}")
    if not rows_completas:
        print("  No quedó ninguna fila útil en este archivo -- nada que importar.")
        return False

    ahora_iso = historico_lib.iso_js(datetime.utcnow())
    por_mes = agrupar_por_mes(rows_completas)
    for clave in sorted(por_mes):
        res = historico_lib.archivar_mes(HISTORICO_DIR, clave, por_mes[clave], id_login_de_contratista, ahora_iso)
        print(f"  {clave} ({res['label']}): {res['publicas']} filas (reducido+completo, ya mezclado con lo archivado antes)")
        for contratista, n_pub_c, n_comp_c in res["contratistas"]:
            print(f"    {contratista}: {n_pub_c} filas")
    return True


def main():
    maestros = cargar_maestros()
    correcciones = cargar_correcciones()

    archivos = sys.argv[1:]
    modo_carpeta = not archivos
    if modo_carpeta:
        os.makedirs(ENTRADA_DIR, exist_ok=True)
        os.makedirs(YA_IMPORTADOS_DIR, exist_ok=True)
        archivos = sorted(
            p for ext in ("*.xlsx", "*.xls", "*.xlsm")
            for p in glob.glob(os.path.join(ENTRADA_DIR, ext))
        )
        if not archivos:
            print(f"No hay ningún Excel en {ENTRADA_DIR}\\ para importar.")
            print("Coloca ahí el/los Excel de los meses ya cerrados, o arrastra uno")
            print("directo sobre Importar-Mes-Historico.bat.")
            sys.exit(1)
        print(f"Encontrados {len(archivos)} archivo(s) en EXCEL-MESES-ANTERIORES\\:")
        for a in archivos:
            print(f"  - {os.path.basename(a)}")

    for gab_path in archivos:
        if not os.path.exists(gab_path):
            print(f"\nNo existe: {gab_path}")
            continue
        ok = importar_archivo(gab_path, maestros, correcciones)
        if ok and modo_carpeta:
            try:
                destino = os.path.join(YA_IMPORTADOS_DIR, os.path.basename(gab_path))
                if os.path.exists(destino):
                    nombre, ext = os.path.splitext(os.path.basename(gab_path))
                    destino = os.path.join(YA_IMPORTADOS_DIR, f"{nombre}_{datetime.now().strftime('%H%M%S')}{ext}")
                shutil.move(gab_path, destino)
                print(f"  Movido a YA-IMPORTADOS\\{os.path.basename(destino)}")
            except Exception as e:
                print(f"  AVISO: no se pudo mover el archivo a YA-IMPORTADOS ({e}) — no es grave, ya se importó igual.")

    print("\nListo. Corre Publicar-Historico.bat para subir estos meses al sitio")
    print("(o espera a la próxima vez que corras Ejecutar.bat, que también lo hace).")


if __name__ == "__main__":
    main()

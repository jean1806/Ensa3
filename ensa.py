import os
import sys
import shutil
import tempfile
import json

if getattr(sys, 'frozen', False):
    local_appdata = os.environ.get("LOCALAPPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Local"))
    ms_playwright_dir = os.path.join(local_appdata, "ms-playwright")
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ms_playwright_dir

import re
import html
import unicodedata
import random
import asyncio
import threading
import subprocess
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

def running_in_pyinstaller():
    return getattr(sys, 'frozen', False)

if not running_in_pyinstaller():
    def pip_install_if_missing(package_name, import_name=None):
        try:
            __import__(import_name or package_name)
        except Exception:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            __import__(import_name or package_name)

    def ensure_playwright_ready():
        pip_install_if_missing("playwright")
        try:
            subprocess.run([sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
                check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

    try:
        pip_install_if_missing("openpyxl")
        pip_install_if_missing("PyPDF2")
        ensure_playwright_ready()
    except Exception:
        pass

# ── Archivo de credenciales ───────────────────────────────────────────────────
# Línea 1: usuario ENSA (DNI/RUC/CE)
# Línea 2: contraseña ENSA
CREDENTIALS_FILE = Path(__file__).parent / "credenciales.txt"
# ── Archivos de correo ────────────────────────────────────────────────────────
CORREO_TEMPLATE_FILE      = Path(__file__).parent / "correo.txt"
DESTINATARIOS_CORREO_FILE = Path(__file__).parent / "destinatarios_correo.txt"

CORREO_TEMPLATE_DEFAULT = (
    "Asunto: Notificaciones OSINERGMIN\n"
    "Cuerpo:\n"
    "Estimados,\n\n"
    "Se les informa que se han descargado y clasificado las notificaciones de OSINERGMIN.\n\n"
    "Por favor revisar la carpeta correspondiente para atender los documentos.\n\n"
    "Saludos."
)

def _crear_archivos_correo_si_no_existen():
    if not CORREO_TEMPLATE_FILE.exists():
        try:
            CORREO_TEMPLATE_FILE.write_text(CORREO_TEMPLATE_DEFAULT, encoding="utf-8")
        except Exception:
            pass
    if not DESTINATARIOS_CORREO_FILE.exists():
        try:
            DESTINATARIOS_CORREO_FILE.write_text(
                "# Un correo por línea. Líneas con # son comentarios.\n# ejemplo@empresa.com\n",
                encoding="utf-8"
            )
        except Exception:
            pass

def leer_template_correo():
    _crear_archivos_correo_si_no_existen()
    try:
        texto = CORREO_TEMPLATE_FILE.read_text(encoding="utf-8")
        asunto, cuerpo_lineas, en_cuerpo = "", [], False
        for linea in texto.splitlines():
            if not en_cuerpo and linea.lower().startswith("asunto:"):
                asunto = linea[len("asunto:"):].strip()
            elif not en_cuerpo and linea.lower().startswith("cuerpo:"):
                en_cuerpo = True
                resto = linea[len("cuerpo:"):].strip()
                if resto:
                    cuerpo_lineas.append(resto)
            elif en_cuerpo:
                cuerpo_lineas.append(linea)
        return asunto, "\n".join(cuerpo_lineas)
    except Exception:
        return "Notificaciones OSINERGMIN", ""

def leer_destinatarios_correo():
    _crear_archivos_correo_si_no_existen()
    try:
        lines = DESTINATARIOS_CORREO_FILE.read_text(encoding="utf-8").splitlines()
        return [l.strip() for l in lines if l.strip() and not l.strip().startswith("#")]
    except Exception:
        return []

def enviar_correo_outlook(asunto, cuerpo, destinatarios, log_cb=None):
    if not destinatarios:
        if log_cb:
            log_cb("⚠ Sin destinatarios en destinatarios_correo.txt")
        return False
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = asunto
        mail.Body    = cuerpo
        mail.To      = "; ".join(destinatarios)
        mail.Send()
        if log_cb:
            log_cb(f"✅ Correo enviado a: {', '.join(destinatarios)}")
        return True
    except ImportError:
        if log_cb:
            log_cb("❌ pywin32 no instalado. Instalar con: pip install pywin32")
        return False
    except Exception as e:
        if log_cb:
            log_cb(f"❌ Error enviando correo: {e}")
        return False

def leer_credenciales_ensa():
    """Lee líneas 1 y 2 del archivo de credenciales para ENSA."""
    usuario, password = "", ""
    if CREDENTIALS_FILE.exists():
        try:
            lineas = CREDENTIALS_FILE.read_text(encoding="utf-8").splitlines()
            if len(lineas) >= 1:
                usuario = lineas[0].strip()
            if len(lineas) >= 2:
                password = lineas[1].strip()
        except Exception:
            pass
    return usuario, password

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN ENSA
# ══════════════════════════════════════════════════════════════════════════════
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".osinergmin_config.json")
FORZAR_SIN_UE = False  # True para evitar mover a UE y mantener todo en Sin UE (útil para debug)
MOVER_A_UE = True   # True para clasificar en su carpeta de unidad empresarial

PROYECTO_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA_DESCARGAS_DEFAULT = os.path.join(PROYECTO_DIR, "descargas")


def get_documentos_download_folder():
    # Carpeta predeterminada: carpeta 'descargas' dentro del proyecto
    return CARPETA_DESCARGAS_DEFAULT

def _leer_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def _escribir_config(clave: str, valor):
    config = _leer_config()
    config[clave] = valor
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"⚠ No se pudo guardar config ({clave}): {e}")

def _normalizar_ruta(ruta):
    try:
        return os.path.normpath(ruta)
    except Exception:
        return ruta


def _carpeta_interna(ruta):
    if not ruta:
        ruta_n = CARPETA_DESCARGAS_DEFAULT
    else:
        ruta_n = os.path.expanduser(str(ruta))
        if not os.path.isabs(ruta_n):
            ruta_n = os.path.join(CARPETA_DESCARGAS_DEFAULT, ruta_n)
        ruta_n = _normalizar_ruta(ruta_n)
    ruta_n = os.path.abspath(ruta_n)
    os.makedirs(ruta_n, exist_ok=True)
    return ruta_n


def cargar_ruta_guardada():
    ruta = _leer_config().get('carpeta_descarga')
    if not ruta:
        ruta = get_documentos_download_folder()
    else:
        try:
            ruta = os.path.abspath(os.path.expanduser(str(ruta)))
            if not os.path.exists(ruta):
                os.makedirs(ruta, exist_ok=True)
        except Exception:
            ruta = get_documentos_download_folder()
    return _normalizar_ruta(ruta)


def guardar_ruta(ruta):
    global CARPETA_DESCARGA
    try:
        ruta_n = os.path.abspath(os.path.expanduser(str(ruta)))
        os.makedirs(ruta_n, exist_ok=True)
        CARPETA_DESCARGA = _normalizar_ruta(ruta_n)
        _escribir_config('carpeta_descarga', CARPETA_DESCARGA)
    except Exception:
        pass


CARPETA_DESCARGA = _normalizar_ruta(cargar_ruta_guardada())
os.makedirs(CARPETA_DESCARGA, exist_ok=True)

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

def sanitize_filename(nombre: str) -> str:
    nombre = re.sub(r'[<>:"/\\|?*]+', "_", nombre)
    return nombre.strip().strip(".")

async def espera_humana(page, min_sec=0.5, max_sec=1.0):
    tiempo = random.uniform(min_sec, max_sec) * 1000
    await page.wait_for_timeout(tiempo)

async def mover_mouse_a_elemento(page, elemento):
    box = await elemento.bounding_box()
    if box:
        x = box["x"] + box["width"] / 2
        y = box["y"] + box["height"] / 2
        await page.mouse.move(x, y, steps=random.randint(8, 15))

async def guardar_descarga(download, carpeta_destino, registrar_cb, log_cb, nombre_preferido=None):
    try:
        os.makedirs(carpeta_destino, exist_ok=True)
    except OSError as e:
        msg = f"❌ No se pudo crear la carpeta de descarga: {carpeta_destino}\nError: {e}"
        log_cb(msg)
        raise RuntimeError(msg)

    sugerido = download.suggested_filename
    nombre = sanitize_filename(nombre_preferido or sugerido or "archivo")
    ruta = os.path.join(carpeta_destino, nombre)
    # Sobrescribir si existe, no añadir (1)
    try:
        await download.save_as(ruta)
        ruta = _normalizar_ruta(ruta)
        registrar_cb(ruta)
        log_cb(f"✅ Guardado: {ruta}")
        return ruta
    except OSError as e:
        if (hasattr(e, "winerror") and e.winerror == 112) or "No space" in str(e) or "espacio" in str(e).lower():
            msg = "❌ ¡Carpeta llena o insuficiente espacio en disco/nube!"
        else:
            msg = f"❌ No se pudo guardar la descarga: {e}"
        log_cb(msg)
        raise RuntimeError(msg)

def leer_registros_desde_disco(carpeta_destino):
    registros = []
    if not os.path.isdir(carpeta_destino):
        return registros
    
    # Hacer walk recursivo para encontrar todos los .meta.json en cualquier nivel
    for root, dirs, files in os.walk(carpeta_destino):
        # Saltar la carpeta 'excel'
        if os.path.basename(root) == 'excel':
            continue
        
        if '.meta.json' in files:
            meta_path = os.path.join(root, '.meta.json')
            try:
                with open(meta_path, 'r', encoding='utf-8') as f:
                    meta = json.load(f)
                archivos = meta.get('archivos')
                if isinstance(archivos, list):
                    archivos = [f for f in archivos if isinstance(f, str) and f.strip()]
                else:
                    archivos = []
                if not archivos:
                    archivos = [f for f in files if f != '.meta.json' and os.path.isfile(os.path.join(root, f))]
                suministros = meta.get('suministros', [])
                suministro_principal = meta.get('suministro_principal', '')
                carpeta_destino_meta = meta.get('carpeta_destino', '')
                
                # Si no está en metadatos, inferir desde la ruta
                if not carpeta_destino_meta:
                    carpeta_rel = os.path.normpath(os.path.relpath(root, carpeta_destino))
                    # El primer componente de la ruta es el UE
                    parts = carpeta_rel.split(os.sep)
                    carpeta_destino_meta = parts[0] if parts and parts[0] != '.' else ''

                base = {
                    'expediente':          meta.get('expediente', ''),
                    'codigo_notificacion': meta.get('codigo_notificacion', os.path.basename(root)),
                    'unidad_operativa':    meta.get('unidad_operativa', ''),
                    'procedimiento':       meta.get('procedimiento', ''),
                    'asunto':              meta.get('asunto', ''),
                    'fecha_notificacion':  meta.get('fecha_notificacion', ''),
                    'suministros':         ', '.join(suministros) if suministros else '',
                    'suministro_principal': suministro_principal,
                    'carpeta_destino':     carpeta_destino_meta,
                }
                
                # Crear un registro por cada archivo (no agrupar)
                if archivos:
                    for archivo in sorted(archivos):
                        registros.append({**base, 'archivo': archivo})
                # No agregamos filas vacías si no hay archivos reales
            except Exception:
                pass
    
    return registros


def crear_excel_reporte(registros_sesion, carpeta_destino, fecha_descarga, tipo_descarga, identificador=None):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None

    index = {}
    # Cargar registros desde disco primero (son los más actualizados)
    for r in leer_registros_desde_disco(carpeta_destino):
        index[(r['codigo_notificacion'], r['archivo'])] = r
    # Luego, para cada registro de sesión, actualizar SOLO si no existe en disco
    # o actualizar solo ciertos campos que pudo haber cambiado en sesión
    for r in registros_sesion:
        key = (r['codigo_notificacion'], r['archivo'])
        if key in index:
            # Si ya existe en disco, preferir valores de disco (son más recientes)
            # Solo actualizar expediente, procedimiento, asunto, fecha si vinieron de sesión
            if r.get('expediente') and not index[key].get('expediente'):
                index[key]['expediente'] = r['expediente']
            if r.get('procedimiento') and not index[key].get('procedimiento'):
                index[key]['procedimiento'] = r['procedimiento']
            if r.get('asunto') and not index[key].get('asunto'):
                index[key]['asunto'] = r['asunto']
            if r.get('fecha_notificacion') and not index[key].get('fecha_notificacion'):
                index[key]['fecha_notificacion'] = r['fecha_notificacion']
        else:
            # Si no existe en disco, agregar el registro de sesión
            index[key] = r
    todos = [r for r in sorted(index.values(), key=lambda r: (r.get('codigo_notificacion',''), r.get('archivo',''))) if r.get('archivo')]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte"
    ws['A1'] = f"Excel Osinergmin ({fecha_descarga})"
    ws['A1'].font = Font(bold=True, size=12)

    headers = ['Número de Expediente', 'Número de Notificación', 'Agente Supervisado',
               'Procedimiento', 'Asunto', 'Fecha de Notificación', 'Nombre del Archivo', 'Carpeta destino', 'Suministros detectados']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, registro in enumerate(todos, start=4):
        ws.cell(row=row_idx, column=1).value = registro.get('expediente', '')
        ws.cell(row=row_idx, column=2).value = registro.get('codigo_notificacion', '')
        unidad_operativa_valor = registro.get('unidad_operativa', '')
        if isinstance(unidad_operativa_valor, str) and unidad_operativa_valor.strip().upper() == 'SIN UE':
            unidad_operativa_valor = ''
        ws.cell(row=row_idx, column=3).value = unidad_operativa_valor
        ws.cell(row=row_idx, column=4).value = registro.get('procedimiento', '')
        ws.cell(row=row_idx, column=5).value = registro.get('asunto', '')
        ws.cell(row=row_idx, column=6).value = registro.get('fecha_notificacion', '')
        ws.cell(row=row_idx, column=7).value = registro.get('archivo', '')
        carpeta_destino_valor = registro.get('carpeta_destino') or 'Sin UE'
        ws.cell(row=row_idx, column=8).value = carpeta_destino_valor
        suministros_valor = registro.get('suministros') or registro.get('suministro_principal', '')
        if not suministros_valor:
            suministros_valor = "suministro no encontrado"
        ws.cell(row=row_idx, column=9).value = suministros_valor

    for col in range(1, len(headers) + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 5, 60)

    last_row = 3 + len(todos)
    ws.auto_filter.ref = f"A3:I{last_row}"
    carpeta_excel = os.path.join(carpeta_destino, "excel")
    os.makedirs(carpeta_excel, exist_ok=True)

    if tipo_descarga == "no_leidos":
        nombre_excel = "reporte_no_leidos.xlsx"
    elif tipo_descarga == "fecha":
        nombre_excel = "reporte_fecha.xlsx"
    elif tipo_descarga == "hora":
        nombre_excel = "reporte_hora.xlsx"
    else:
        nombre_excel = "reporte_expediente.xlsx"

    ruta_excel = os.path.join(carpeta_excel, nombre_excel)
    wb.save(ruta_excel)
    return ruta_excel

def crear_excel_sin_ue(registros_sesion, carpeta_destino, fecha_descarga, tipo_descarga, identificador=None):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None

    index = {}
    # Cargar registros desde disco primero (son los más actualizados)
    for r in leer_registros_desde_disco(carpeta_destino):
        index[(r['codigo_notificacion'], r['archivo'])] = r
    # Luego, para cada registro de sesión, actualizar SOLO si no existe en disco
    # o actualizar solo ciertos campos que pudo haber cambiado en sesión
    for r in registros_sesion:
        key = (r['codigo_notificacion'], r['archivo'])
        if key in index:
            # Si ya existe en disco, preferir valores de disco (son más recientes)
            # Solo actualizar expediente, procedimiento, asunto, fecha si vinieron de sesión
            if r.get('expediente') and not index[key].get('expediente'):
                index[key]['expediente'] = r['expediente']
            if r.get('procedimiento') and not index[key].get('procedimiento'):
                index[key]['procedimiento'] = r['procedimiento']
            if r.get('asunto') and not index[key].get('asunto'):
                index[key]['asunto'] = r['asunto']
            if r.get('fecha_notificacion') and not index[key].get('fecha_notificacion'):
                index[key]['fecha_notificacion'] = r['fecha_notificacion']
        else:
            # Si no existe en disco, agregar el registro de sesión
            index[key] = r
    todos = [r for r in sorted(index.values(), key=lambda r: (r.get('codigo_notificacion',''), r.get('archivo',''))) if r.get('archivo')]

    # Agrupar por expediente
    expedientes = {}
    for registro in todos:
        exp = registro.get('expediente', '')
        if exp:
            if exp not in expedientes:
                expedientes[exp] = []
            expedientes[exp].append(registro)

    # Filtrar expedientes que tienen al menos un registro sin UE
    expedientes_sin_ue = {}
    for exp, regs in expedientes.items():
        tiene_sin_ue = any(r.get('unidad_operativa', '').strip() in ['', 'Sin UE'] for r in regs)
        if tiene_sin_ue:
            expedientes_sin_ue[exp] = regs

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expedientes Sin UE"
    ws['A1'] = f"Excel Expedientes Sin UE ({fecha_descarga})"
    ws['A1'].font = Font(bold=True, size=12)

    headers = ['Número de Expediente', 'Número de Notificación', 'Unidad Operativa',
               'Procedimiento', 'Asunto', 'Fecha de Notificación', 'Nombre del Archivo', 'Carpeta destino', 'Suministros detectados']
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row_idx = 4
    for exp in sorted(expedientes_sin_ue.keys()):
        regs = expedientes_sin_ue[exp]
        # Separar con suministro y sin suministro
        con_suministro = [r for r in regs if r.get('unidad_operativa', '').strip() not in ['', 'Sin UE']]
        sin_suministro = [r for r in regs if r.get('unidad_operativa', '').strip() in ['', 'Sin UE']]
        
        # Agregar filas para con suministro
        for registro in con_suministro:
            ws.cell(row=row_idx, column=1).value = exp
            ws.cell(row=row_idx, column=2).value = registro.get('codigo_notificacion', '')
            ws.cell(row=row_idx, column=3).value = registro.get('unidad_operativa', '')
            ws.cell(row=row_idx, column=4).value = registro.get('procedimiento', '')
            ws.cell(row=row_idx, column=5).value = registro.get('asunto', '')
            ws.cell(row=row_idx, column=6).value = registro.get('fecha_notificacion', '')
            ws.cell(row=row_idx, column=7).value = registro.get('archivo', '')
            carpeta_destino_valor = registro.get('carpeta_destino') or 'Sin UE'
            ws.cell(row=row_idx, column=8).value = carpeta_destino_valor
            suministros_valor = registro.get('suministros', '') or registro.get('suministro_principal', '')
            if not suministros_valor:
                suministros_valor = "suministro no encontrado"
            ws.cell(row=row_idx, column=9).value = suministros_valor
            row_idx += 1
        
        # Agregar filas para sin suministro
        for registro in sin_suministro:
            ws.cell(row=row_idx, column=1).value = exp
            ws.cell(row=row_idx, column=2).value = registro.get('codigo_notificacion', '')
            ws.cell(row=row_idx, column=3).value = registro.get('unidad_operativa', '')
            ws.cell(row=row_idx, column=4).value = registro.get('procedimiento', '')
            ws.cell(row=row_idx, column=5).value = registro.get('asunto', '')
            ws.cell(row=row_idx, column=6).value = registro.get('fecha_notificacion', '')
            ws.cell(row=row_idx, column=7).value = registro.get('archivo', '')
            carpeta_destino_valor = registro.get('carpeta_destino') or 'Sin UE'
            ws.cell(row=row_idx, column=8).value = carpeta_destino_valor
            suministros_valor = registro.get('suministros', '') or registro.get('suministro_principal', '')
            if not suministros_valor:
                suministros_valor = "suministro no encontrado"
            ws.cell(row=row_idx, column=9).value = suministros_valor
            row_idx += 1

    for col in range(1, len(headers) + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 5, 60)

    last_row = max(3, row_idx - 1)
    ws.auto_filter.ref = f"A3:I{last_row}"
    carpeta_excel = os.path.join(carpeta_destino, "excel")
    os.makedirs(carpeta_excel, exist_ok=True)

    nombre_excel = "reporte_sin_ue.xlsx"
    ruta_excel = os.path.join(carpeta_excel, nombre_excel)
    wb.save(ruta_excel)
    return ruta_excel

def guardar_progreso(carpeta_descarga, tipo, identificador, notificaciones_procesadas, archivos_descargados, pagina_actual):
    try:
        archivo_progreso = os.path.join(carpeta_descarga, ".progreso.json")
        progreso = {
            'tipo': tipo,
            'identificador': identificador,
            'notificaciones_procesadas': list(notificaciones_procesadas),
            'archivos_descargados': list(archivos_descargados),
            'pagina_actual': pagina_actual,
            'fecha_guardado': datetime.now().isoformat()
        }
        with open(archivo_progreso, 'w') as f:
            json.dump(progreso, f, indent=2)
        return archivo_progreso
    except Exception as e:
        print(f"⚠ Error guardando progreso: {e}")
        return None

def cargar_progreso(carpeta_descarga):
    try:
        archivo_progreso = os.path.join(carpeta_descarga, ".progreso.json")
        if os.path.exists(archivo_progreso):
            with open(archivo_progreso, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"⚠ Error cargando progreso: {e}")
    return None

def limpiar_progreso(carpeta_descarga):
    try:
        archivo_progreso = os.path.join(carpeta_descarga, ".progreso.json")
        if os.path.exists(archivo_progreso):
            os.remove(archivo_progreso)
    except:
        pass

CONFIG_HORA_FILE = os.path.join(os.path.expanduser("~"), ".osinergmin_ultima_hora.json")

def cargar_ultima_hora():
    if os.path.exists(CONFIG_HORA_FILE):
        try:
            with open(CONFIG_HORA_FILE, 'r') as f:
                data = json.load(f)
                return datetime.fromisoformat(data['ultima_descarga'])
        except Exception:
            pass
    return None

def guardar_ultima_hora(dt):
    try:
        with open(CONFIG_HORA_FILE, 'w') as f:
            json.dump({'ultima_descarga': dt.isoformat()}, f)
    except Exception as e:
        print(f"⚠ Error guardando última hora: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA PDF
# ══════════════════════════════════════════════════════════════════════════════

def html_unescape_deep(s: str) -> str:
    if s is None:
        return ""
    return html.unescape(html.unescape(s))

def limpiar_texto(t: str) -> str:
    t = html_unescape_deep(t or "")
    t = unicodedata.normalize("NFKC", t)
    t = t.replace("Nº", "N°").replace("N.o", "N°").replace("N.°", "N°").replace("N°.", "N°")
    return re.sub(r"[\s\t]+", " ", t).strip()

def normalizar_pdf(t: str) -> str:
    if not t:
        return ""
    t = unicodedata.normalize("NFKC", t).upper()
    t = t.replace("–", "-").replace("—", "-").replace("−", "-").replace("／", "/")
    return re.sub(r"[ \t]+", " ", t).strip()

# Regex para detectar "Asunto :" en oficios (como segunda línea de encabezado)
RE_OFICIO_ASUNTO_LINEA = re.compile(
    r"^(?:OFICIO|OFIC\.?)\s+(?:N[°º]|NRO\.?|NO\.?)\s+[\d\-/A-Z]+",
    re.IGNORECASE | re.MULTILINE
)
# Patrones de cuerpo típicos de oficios de traslado/denuncia/requerimiento
RE_OFICIO_CUERPO = re.compile(
    r"(?:"
    r"Mediante\s+la\s+presente\s+(?:le\s+saludo|nos\s+dirigimos|me\s+dirijo)|"
    r"Nos\s+dirigimos\s+a\s+usted\s+para\s+comunicarle|"
    r"Me\s+dirijo\s+a\s+usted|"
    r"traslado\s+de\s+denuncia|"
    r"requerimiento\s+de\s+(?:informaci[oó]n|base\s+de\s+datos)|"
    r"fiscalizaci[oó]n\s+(?:especifica|de\s+campo)"
    r")",
    re.IGNORECASE
)
RE_CABECERA_OFICIO = re.compile(
    r"""
    \bOFICIO\s+(?:NRO\.?|N°|Nº|NO\.?)
    \s+(?P<codigo>\d{2,6}-\d{4}-[A-Z0-9][A-Z0-9\-/\.]*)
    (?:\s+(?P<sede>[A-ZÁÉÍÓÚÜÑ]{2,}))?
    \s+EXPEDIENTE\s*:\s*(?P<exp>\d{5,})\b
    """,
    re.IGNORECASE | re.VERBOSE
)
RE_OFICIO_COD = re.compile(
    r"\bOFICIO\s+(?:NRO\.?|N[°º]|NO\.?)\s+(?P<codigo>.+?)\s*Expediente",
    re.IGNORECASE
)
RE_ASUNTO_OFICIO = re.compile(
    r"Asunto\s*:\s*(.+?)$",
    re.IGNORECASE | re.MULTILINE
)
RE_EXPEDIENTE_SUELTO = re.compile(
    r"EXPEDIENTE\s*[N°º]*\s*:?\s*(?P<exp>\d{10,})",
    re.IGNORECASE
)
RE_RESOLUCION_JARU = re.compile(
    r"(?:OSINERGMIN\s+)?(?:N[°º]|NRO\.?|NO\.?)\s*(?P<codigo>\d{2,6}-\d{4}[\s\-/][A-Z0-9][A-Z0-9\-/\.\s]*JARU[A-Z0-9\-/\.]*)",
    re.IGNORECASE
)
RE_TIENE_JARU  = re.compile(r"\bJARU\b", re.IGNORECASE)
RE_MATERIA     = re.compile(r"MATERIAS?\s*:\s*(.+?)$", re.IGNORECASE | re.MULTILINE)
RE_RESOLUCION_FUERZA = re.compile(
    r"(?:OSINERGMIN\s+)?N[°º]\s*(?P<codigo>\d{2,6}-\d{4}-[A-Z0-9][A-Z0-9\-/\.]*)(?:\s+(?P<sede>[A-ZÁÉÍÓÚÜÑ]{2,6}))?",
    re.IGNORECASE
)
RE_FALLO = re.compile(
    r"ART[IÍ]CULO\s+1[°º][\.\-\s].*?(FUNDADA|INFUNDADA|IMPROCEDENTE)",
    re.IGNORECASE | re.DOTALL
)
RE_ART2 = re.compile(r"ART[IÍ]CULO\s+2[°º][\.\-\s]", re.IGNORECASE)

# ── Unidades Empresariales JARU ───────────────────────────────────────────────
UNIDADES_EMPRESARIALES = ["PIURA", "SULLANA", "TUMBES", "PAITA", "TALARA", "ALTO PIURA", "SECHURA"]

# ── Lista negra de palabras que no deben considerarse como ubicaciones ──────────
PALABRAS_NEGRA_UE = [
    "EMPRESA", "TRANSPORTES", "CARGA", "SERVICIOS", "ELECTRICIDAD", "GAS",
    "CONCESIONARIA", "OSINERGMIN", "SUMINISTRO", "RESOLUCIÓN", "EXPEDIENTE",
    "CARTA", "NOTIFICACIÓN", "DOCUMENTO", "ADMINISTRATIVO", "PROCEDIMIENTO",
    "FOLIO", "BERNARDO", "MONTEAGUDO", "ATENCIÓN", "CIUDADANO", "OSIVIRTUAL"
]

# Mapeo de localidades/distritos/ciudades a su Unidad Empresarial correspondiente
# Clave: nombre en mayúsculas (sin tildes normalizadas); Valor: nombre UE
MAPA_LOCALIDAD_UE = {
    # PIURA
    "PIURA":          "PIURA",
    "CASTILLA":       "PIURA",
    "CATACAOS":       "PIURA",
    "CURA MORI":      "PIURA",
    "LA UNION":       "PIURA",
    "LAS LOMAS":      "PIURA",
    "MIGUEL CHECA":   "PIURA",
    "RINCONADA LLICUAR": "PIURA",
    # SULLANA
    "SULLANA":        "SULLANA",
    "TAMBOGRANDE":    "SULLANA",
    "MARCAVELICA":    "SULLANA",
    "QUERECOTILLO":   "SULLANA",
    "SALITRAL":       "SULLANA",
    "MIGUEL CHECA":   "SULLANA",
    "LANCONES":       "SULLANA",
    "AYABACA":        "SULLANA",
    "PAIMAS":         "SULLANA",
    "SÍCCHEZ":        "SULLANA",
    "SICCHEZ":        "SULLANA",
    "SUYO":           "SULLANA",
    "LAGUNAS":        "SULLANA",
    "MONTERO":        "SULLANA",
    "JILILI":         "SULLANA",
    "SAPILLICA":      "SULLANA",
    "FRÍAS":          "SULLANA",
    "FRIAS":          "SULLANA",
    # TUMBES
    "TUMBES":         "TUMBES",
    "ZARUMILLA":      "TUMBES",
    "ZORRITOS":       "TUMBES",
    "AGUAS VERDES":   "TUMBES",
    "CORRALES":       "TUMBES",
    "SAN JACINTO":    "TUMBES",
    "PAMPAS DE HOSPITAL": "TUMBES",
    "LA CRUZ":        "TUMBES",
    "CASITAS":        "TUMBES",
    # PAITA
    "PAITA":          "PAITA",
    "AMOTAPE":        "PAITA",
    "ARENAL":         "PAITA",
    "COLAN":          "PAITA",
    "COLÁN":          "PAITA",
    "LA HUACA":       "PAITA",
    "TAMARINDO":      "PAITA",
    "VICHAYAL":       "PAITA",
    # TALARA
    "TALARA":         "TALARA",
    "MANCORA":        "TALARA",
    "MÁNCORA":        "TALARA",
    "NEGRITOS":       "TALARA",
    "EL ALTO":        "TALARA",
    "LOS ORGANOS":    "TALARA",
    "LOS ÓRGANOS":    "TALARA",
    "ORGANOS":        "TALARA",
    "ÓRGANOS":        "TALARA",
    "LA BREA":        "TALARA",
    "LOBITOS":        "TALARA",
    "EL PORVENIR":    "TALARA",
    "PARIÑAS":        "TALARA",
    # ALTO PIURA
    "CHULUCANAS":     "ALTO PIURA",
    "MORROPON":       "ALTO PIURA",
    "MORROPÓN":       "ALTO PIURA",
    "HUANCABAMBA":    "ALTO PIURA",
    "CHALACO":        "ALTO PIURA",
    "CANCHAQUE":      "ALTO PIURA",
    "BUENOS AIRES":   "ALTO PIURA",
    "YAMANGO":        "ALTO PIURA",
    # SECHURA
    "SECHURA":        "SECHURA",
    "BERNAL":         "SECHURA",
    "CRISTO NOS VALGA": "SECHURA",
    "RINCONADA LLICUAR": "SECHURA",
    "LaVICE":           "SECHURA",
}

def _normalizar_localidad(s: str) -> str:
    """Normaliza una cadena para comparación: mayúsculas, sin tildes, sin espacios extra."""
    s = unicodedata.normalize("NFKD", s.upper())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()

# ── Mapeo de Oficinas Regionales → Provincias ─────────────────────────────────
# Para detectar provincia desde códigos de oficina regional en PDFs de cargo/trámite
MAPA_OFICINA_REGIONAL = {
    # TALARA
    "ENOSATALARA2": "TALARA",
    "ENOSATALARA":  "TALARA",
    "TALARA":       "TALARA",
    "OR TALARA":    "TALARA",
    "TALARA2":      "TALARA",
    
    # TUMBES
    "TUMBES":       "TUMBES",
    "OR TUMBES":    "TUMBES",
    
    # SULLANA
    "SULLANA":      "SULLANA",
    "OR SULLANA":   "SULLANA",
    "ENOSASULLANA": "SULLANA",
    "ENOSASULLANA2": "SULLANA",
    
    # PAITA
    "PAITA":        "PAITA",
    "OR PAITA":     "PAITA",
    "ENOSAPAIT":    "PAITA",
    "ENOSAPITA":    "PAITA",
    
    # PIURA (CENTRAL)
    "PIURA":        "PIURA",
    "OR PIURA":     "PIURA",
    "ENOSAPIR":     "PIURA",
    "ENOSACF":      "PIURA",  # Central Finance/Castilla o similar
    "ENOSACF2":     "PIURA",
    "ENOSACENT":    "PIURA",
    
    # ALTO PIURA
    "ALTO PIURA":   "ALTO PIURA",
    "ENOSACHUL":    "ALTO PIURA",
    "ENOSACHULUCANAS": "ALTO PIURA",
    
    # SECHURA
    "SECHURA":      "SECHURA",
    "ENOSASEC":     "SECHURA",
}

# ── Patrones para detectar provincia/ciudad desde el encabezado del documento ──
# Detecta: "Miraflores, 20 de marzo del 2026" o "Tumbes, 23 de marzo del 2026"
RE_CIUDAD_ENCABEZADO = re.compile(
    r"^(?:Crédito de Documento Ingresado)?\s*"
    r"([A-ZÁÉÍÓÚÜÑ][A-Za-záéíóúüñ\s]{1,30}?),\s*"
    r"(?:\d{1,2}|(?:0?[1-9]|[12]\d|3[01]))\s+de\s+"
    r"(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|setiembre|septiembre|octubre|noviembre|diciembre)",
    re.IGNORECASE | re.MULTILINE
)

# Detecta códigos de oficina regional como "ENOSATALARA2", "STOR", etc.
RE_OFICINA_REGIONAL = re.compile(
    r"(?:^|\n|Oficina\s+de\s*:|Remitente\s*:)\s*"
    r"(?P<oficina>ENOSA\s*[A-Z0-9]{2,}|STOR(?:\s+[A-Z]{2,})?|OR\s+[A-Z]{2,})",
    re.IGNORECASE | re.MULTILINE
)

# ── Captura el texto de "Ubicación del suministro" hasta la siguiente etiqueta conocida o fin de línea
RE_UBICACION_SUMINISTRO = re.compile(
    r"Ubicaci[oó]n\s+del\s+suministro(?:\s+y\s+domicilio\s+procesal)?\s*:\s*(.+?)"
    r"(?=\n\s*(?:Domicilio\s+procesal|Resoluci[oó]n\s+impugnada|Materia\s*:|Suministro\s*:|Recurrente\s*:|Empresa\s+distribuidora)|$)",
    re.IGNORECASE | re.DOTALL
)

# Regex para el campo "Ubicación" en PDFs de cargo/inconformidad de Osinergmin.
# PyPDF2 puede extraer el texto en orden inverso ("dirección Ubicación") o normal.
RE_UBICACION_CARGO = re.compile(
    r"(?:^|\n)\s*Ubicaci[oó]n\s+(.+?)\s*(?:\n|$)"
    r"|(?:^|\n)\s*(.+?)\s+Ubicaci[oó]n\s*(?:\n|$)",
    re.IGNORECASE | re.MULTILINE
)

# ══════════════════════════════════════════════════════════════════════════════
# ── Funciones para detectar provincia desde encabezado y oficina regional ──────
# ══════════════════════════════════════════════════════════════════════════════

def detectar_provincia_desde_encabezado(texto: str) -> str:
    """
    Detecta la provincia desde la ciudad en el encabezado del documento.
    Ejemplo: "Tumbes, 20 de marzo del 2026" → "TUMBES"
    Retorna la provincia detectada o empty string si no la encuentra.
    """
    for m in RE_CIUDAD_ENCABEZADO.finditer(texto[:500]):
        ciudad = m.group(1).strip().upper()
        # Si es Miraflores (sede central), no es provincia específica
        if "MIRAFLORES" in ciudad:
            return ""
        # Evitar falso positivo con la referencia a PIURA como departamento genérico en encabezado
        if ciudad == "PIURA":
            return ""
        # Normalizar y buscar en el mapeo
        ciudad_norm = _normalizar_localidad(ciudad)
        for clave_norm, ue in MAPA_LOCALIDAD_UE.items():
            if ciudad_norm == clave_norm:
                return ue
    return ""

def detectar_provincia_desde_oficina_regional(texto: str) -> str:
    """
    Detecta la provincia desde el código de oficina regional.
    Ejemplo: "ENOSATALARA2" → "TALARA"
    Ejemplo: "OR TUMBES" → "TUMBES"
    Retorna la provincia (UE) detectada o empty string si no la encuentra.
    """
    for m in RE_OFICINA_REGIONAL.finditer(texto[:1000]):
        oficina = m.group("oficina").strip().upper()
        oficina_norm = _normalizar_localidad(oficina)
        oficina_norm_sin_esp = oficina_norm.replace(" ", "")
        for clave_norm, provincia in MAPA_OFICINA_REGIONAL.items():
            if clave_norm == oficina_norm:
                return provincia
            if clave_norm == oficina_norm_sin_esp:
                return provincia
            if oficina_norm.startswith(clave_norm) or oficina_norm_sin_esp.startswith(clave_norm):
                return provincia
            if clave_norm.startswith(oficina_norm) or clave_norm.startswith(oficina_norm_sin_esp):
                return provincia
    return ""

# ══════════════════════════════════════════════════════════════════════════════

def detectar_ue_desde_texto(texto: str) -> str:
    """
    Dado el texto de un PDF de resolución JARU, busca la 'Ubicación del suministro'
    y determina a qué Unidad Empresarial pertenece usando el MAPA_LOCALIDAD_UE.

    Revisa todas las partes de la dirección (separadas por coma o guión),
    de más específico a más general, ignorando el departamento (último campo).
    """
    mapa_norm = {_normalizar_localidad(k): v for k, v in MAPA_LOCALIDAD_UE.items()}

    for m in RE_UBICACION_SUMINISTRO.finditer(texto):
        fragmento = m.group(1)
        # Unir líneas seguidas (puede haber salto de línea en medio de la dirección)
        fragmento = re.sub(r"\s+", " ", fragmento).strip()[:400]
        # Intentar separar por coma primero; si hay al menos 2 partes, usar esa jerarquía.
        partes_coma = [p.strip() for p in fragmento.split(",") if p.strip()]
        if len(partes_coma) >= 2:
            partes = partes_coma
        else:
            # Separar por guión largo o guión simple rodeado de espacios o al final
            partes = [p.strip() for p in re.split(r"\s*[-–]\s*", fragmento) if p.strip()]
        # Quitar el último campo (departamento)
        candidatos = partes[:-1] if len(partes) > 1 else partes
        # Revisar de más específico (último) a más general
        for parte in reversed(candidatos):
            norm = _normalizar_localidad(parte)
            if norm in mapa_norm:
                return mapa_norm[norm]
            # Coincidencia de palabra completa dentro de la parte
            for clave_norm, ue in mapa_norm.items():
                if re.search(rf"\b{re.escape(clave_norm)}\b", norm):
                    return ue
    return "Sin UE"

# Regex para detectar "provincia de X" / "distrito de X" en fuerza mayor (sección 1.1)
RE_DISTRITO_PROVINCIA_FM = re.compile(
    r"(?:en\s+el\s+)?(?:distrito|provincia)\s+de\s+"
    r"([A-Za-záéíóúüñÁÉÍÓÚÜÑ][A-Za-záéíóúüñÁÉÍÓÚÜÑ\s]{1,30}?)"
    r"(?:\s*[,;]|\s+departamento)",
    re.IGNORECASE
)


def detectar_ue_desde_texto_fm(texto: str) -> str:
    """
    Para fuerza mayor: busca primero 'Ubicación del suministro' (por si acaso),
    luego busca 'distrito de X' / 'provincia de X' en el texto del antecedente.
    Usa el MAPA_LOCALIDAD_UE para mapear a UE.
    """
    mapa_norm = {_normalizar_localidad(k): v for k, v in MAPA_LOCALIDAD_UE.items()}

    # 1º intento: "Ubicación del suministro" (igual que JARU)
    ue = detectar_ue_desde_texto(texto)
    if ue != "Sin UE":
        return ue

    # 2º intento: "distrito de X, provincia de Y" en las primeras 3000 chars
    fragmento = re.sub(r"\s+", " ", texto[:3000])
    candidatos = []
    for m in RE_DISTRITO_PROVINCIA_FM.finditer(fragmento):
        candidatos.append(m.group(1).strip())

    # Buscar de más específico (distrito) a menos (provincia)
    for candidato in candidatos:
        norm = _normalizar_localidad(candidato)
        if norm in mapa_norm:
            return mapa_norm[norm]
        for clave_norm, ue_val in mapa_norm.items():
            if re.search(rf"\b{re.escape(clave_norm)}\b", norm):
                return ue_val

    return "Sin UE"

def detectar_ue_desde_texto_cargo(texto: str) -> str:
    """
    Extrae la UE desde el campo 'Ubicación' de un PDF de cargo/inconformidad de Osinergmin.

    PyPDF2 puede extraer el texto en distintos órdenes:
      - Normal:     "Ubicación  dirección"
      - Invertido:  "dirección Ubicación"  (puede no haber espacio entre ambos)
      - Multilinea: la dirección se parte en dos líneas y la etiqueta queda pegada al final

    Estrategia: localizar la palabra "Ubicación" y analizar una ventana de texto
    alrededor (líneas previas + texto posterior) para reconstruir la dirección completa.
    """
    mapa_norm = {_normalizar_localidad(k): v for k, v in MAPA_LOCALIDAD_UE.items()}

    def _es_palabra_basura(frag: str) -> bool:
        """Verifica si el fragmento contiene palabras de la lista negra."""
        frag_norm = frag.upper()
        for palabra in PALABRAS_NEGRA_UE:
            if palabra in frag_norm:
                return True
        return False

    def _buscar_ue_en_fragmento(frag: str) -> str:
        frag = re.sub(r"\s+", " ", frag).strip()
        
        # Filtrar si contiene palabras de la lista negra
        if _es_palabra_basura(frag):
            return "Sin UE"
        
        # Parte geográfica: todo lo que está después del último guión relevante
        partes_dash = re.split(r"\s*-\s*", frag)
        geo = partes_dash[-1] if len(partes_dash) > 1 else frag
        partes = [p.strip() for p in geo.split(",") if p.strip()]
        candidatos = partes[1:] if len(partes) >= 3 else partes
        for parte in reversed(candidatos):
            if _es_palabra_basura(parte):
                continue
            n = _normalizar_localidad(parte)
            if n in mapa_norm:
                return mapa_norm[n]
            for k, v in mapa_norm.items():
                if re.search(rf"\b{re.escape(k)}\b", n):
                    return v
        n_total = _normalizar_localidad(frag)
        for k, v in mapa_norm.items():
            if re.search(rf"\b{re.escape(k)}\b", n_total):
                return v
        return "Sin UE"

    for m in re.finditer(r"Ubicaci[oó]n", texto, re.IGNORECASE):
        pos = m.start()

        # Caso 1: etiqueta ANTES de la dirección
        despues = texto[m.end(): m.end() + 300]
        linea_despues = despues.split("\n")[0].strip()
        if linea_despues:
            ue = _buscar_ue_en_fragmento(linea_despues)
            if ue != "Sin UE":
                return ue

        # Caso 2: etiqueta DESPUÉS de la dirección (puede estar pegada sin espacio)
        # Tomar las 2 líneas anteriores y unirlas (la dirección puede romperse en dos)
        antes = texto[max(0, pos - 400): pos]
        lineas_antes = [l.strip() for l in antes.split("\n") if l.strip()]
        if lineas_antes:
            fragmento_antes = " ".join(lineas_antes[-2:])
            # Eliminar basura de celdas vecinas que PyPDF2 puede mezclar al inicio
            fragmento_antes = re.sub(
                r"^.*?(?:Inconformidad|Descripci[oó]n|Asunto|Motivo|Sector|Canal|"
                r"Fecha|Exp\.?|Nro\.?|Suministro|Solicitud|DNI|Tel[eé]fono|"
                r"Correo|Empresa|atenci[oó]n)\s*",
                "", fragmento_antes, flags=re.IGNORECASE
            ).strip()
            if fragmento_antes:
                ue = _buscar_ue_en_fragmento(fragmento_antes)
                if ue != "Sin UE":
                    return ue

    return "Sin UE"
def _map_ue_a_unidad_negocio(ue: str) -> str:
    if not ue or ue == "Sin UE":
        return "Sin UE"
    ue_norm = ue.strip().upper()

    # No usar valores demasiado cortos que suelen aparecer como "A" o "P" de regex basura.
    if len(ue_norm) <= 2:
        return "Sin UE"

    # Si ya es una unidad empresarial conocida (nombre completo), mantenerla.
    if ue_norm in UNIDADES_EMPRESARIALES:
        return ue_norm

    # Intentar mapear por localidad conocida (mayúscula/normalizada).
    localidad_norm = _normalizar_localidad(ue_norm)
    if localidad_norm in MAPA_LOCALIDAD_UE:
        return MAPA_LOCALIDAD_UE[localidad_norm]

    # Devolver el nombre original en mayúsculas (fallback razonable).
    return ue_norm


def detectar_ue_carpeta(carpeta_notif: str) -> str:
    """
    Lee los PDFs descargados en carpeta_notif y detecta la UE (o carpeta especial).
    Estrategia:
      1. Recopilar tipos de todos los PDFs.
      2. Si hay cargo → "otros" (todos los cargos van aquí, sin importar contenido)
      3. Si hay sancion, coactiva, informe_final, inicio_proceso_sancionador, informe_fiscalizacion → "otros"
      4. Para otros casos, detectar UE normalmente.
    """
    try:
        todos = sorted(os.listdir(carpeta_notif))
    except Exception:
        return "Sin UE"

    pdfs_cargo = [f for f in todos if f.lower().startswith("cargo") and f.lower().endswith(".pdf")]
    pdfs_otros = [f for f in todos if f not in pdfs_cargo and f.lower().endswith(".pdf")]

    tipos = set()
    for archivo in pdfs_cargo + pdfs_otros:
        ruta_pdf = os.path.join(carpeta_notif, archivo)
        texto = extraer_texto_pdf(ruta_pdf, max_paginas=10 if archivo.lower().startswith("cargo") else None)
        if not texto.strip():
            continue
        if re.search(r"CARGO\s*DE\s*DOCUMENTO\s*INGRESADO", texto, re.IGNORECASE):
            # Priorizar cargo incluso si el detect_tipo_documento falla por formatos raros
            return "otros"
        tipo = detectar_tipo_documento(texto)
        tipos.add(tipo)

    # Todos los cargos van a "otros"
    if "cargo" in tipos:
        return "otros"

    # Sanción y coactiva preservan su propia carpeta (no van al "otros").
    if "sancion" in tipos:
        return "Resolución de sanción"
    if "coactiva" in tipos:
        return "Coactiva"

    # Otros documentos sancionadores/informes van a "otros".
    if any(t in ("informe_final", "inicio_proceso_sancionador", "informe_fiscalizacion") for t in tipos):
        return "otros"

    # Para otros casos, detectar UE normalmente
    ue_fallback = "Sin UE"
    for archivo in pdfs_cargo + pdfs_otros:
        ruta_pdf = os.path.join(carpeta_notif, archivo)
        texto = extraer_texto_pdf(ruta_pdf, max_paginas=10 if archivo.lower().startswith("cargo") else None)
        if not texto.strip():
            continue
        tipo = detectar_tipo_documento(texto)

        if tipo in ("jaru", "resolucion"):
            continue

        # ════ Nuevas detecciones (alta prioridad) ════
        # Para cargos, detectar desde oficina regional PRIMERO
        if tipo == "cargo":
            ue_oficina = detectar_provincia_desde_oficina_regional(texto)
            if ue_oficina:
                return _map_ue_a_unidad_negocio(ue_oficina)
        
        ue_encabezado = detectar_provincia_desde_encabezado(texto)
        if ue_encabezado:
            return _map_ue_a_unidad_negocio(ue_encabezado)
        
        ue_oficina = detectar_provincia_desde_oficina_regional(texto)
        if ue_oficina:
            return _map_ue_a_unidad_negocio(ue_oficina)

        # ════ Detecciones originales (fallback) ════
        if tipo == "cargo":
            ue = detectar_ue_desde_texto_cargo(texto)
            if ue != "Sin UE":
                return _map_ue_a_unidad_negocio(ue)
            continue

        if tipo in ("fuerza_mayor", "oficio"):
            ue = detectar_ue_desde_texto_fm(texto)
        else:
            continue

        if ue != "Sin UE" and ue_fallback == "Sin UE":
            ue_fallback = ue

    return _map_ue_a_unidad_negocio(ue_fallback)

def _actualizar_meta_carpeta(ruta_notif: str, ue: str):
    try:
        meta_path = os.path.join(ruta_notif, '.meta.json')
        if os.path.exists(meta_path):
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
        else:
            meta = {}
        meta['carpeta_destino'] = ue
        with open(meta_path, 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def mover_notif_a_ue(carpeta_notif: str, carpeta_base: str, ue: str) -> str:
    """
    Mueve la carpeta de notificación a carpeta_base/ue/codigo_notif.
    Devuelve la nueva ruta.

    Para notificaciones de cargo, usa carpeta_base/otros/cargo/codigo_notif.
    """
    if FORZAR_SIN_UE:
        ue = "Sin UE"

    if not ue:
        ue = "Sin UE"

    carpeta_ue = _safe_subfolder_path(carpeta_base, ue)
    nombre = os.path.basename(carpeta_notif)

    # Colocar los cargos en subdirectorio 'otros/cargo' si corresponde.
    if ue == "otros":
        try:
            archivos = os.listdir(carpeta_notif)
            tiene_cargo = any(f.lower().startswith("cargo") and f.lower().endswith(".pdf") for f in archivos)
            if tiene_cargo:
                carpeta_ue = _safe_subfolder_path(carpeta_ue, "cargo")
        except Exception:
            pass

    destino = os.path.abspath(os.path.join(carpeta_ue, nombre))

    if not _es_subruta_valida(carpeta_base, destino):
        carpeta_ue = os.path.abspath(os.path.join(carpeta_base, "Sin UE"))
        os.makedirs(carpeta_ue, exist_ok=True)
        destino = os.path.abspath(os.path.join(carpeta_ue, nombre))

    if os.path.abspath(carpeta_notif) == destino:
        _actualizar_meta_carpeta(carpeta_notif, ue)
        return carpeta_notif

    if os.path.exists(destino):
        # Si la carpeta destino ya existe, mover contenido sin perder archivos.
        try:
            for item in os.listdir(carpeta_notif):
                origen_item = os.path.join(carpeta_notif, item)
                destino_item = os.path.join(destino, item)
                if os.path.exists(destino_item):
                    # Si colisiona, usar ruta única para no borrar.
                    destino_item = _get_unique_path(destino_item)
                shutil.move(origen_item, destino_item)
            try:
                os.rmdir(carpeta_notif)
            except Exception:
                pass
            _actualizar_meta_carpeta(destino, ue)
            return destino
        except Exception:
            return carpeta_notif

    try:
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        shutil.move(carpeta_notif, destino)
        _actualizar_meta_carpeta(destino, ue)
        return destino
    except Exception:
        return carpeta_notif


def _get_unique_path(path):
    """Devuelve una ruta no existente a partir de path, sin repetidos como '(1) (1)'."""
    if not os.path.exists(path):
        return path

    carpeta = os.path.dirname(path)
    base = os.path.basename(path)
    nombre, ext = os.path.splitext(base)

    m = re.match(r"^(.*) \((\d+)\)$", nombre)
    if m:
        nombre_base = m.group(1)
        contador = int(m.group(2))
    else:
        nombre_base = nombre
        contador = 0

    while True:
        contador += 1
        candidato = os.path.join(carpeta, f"{nombre_base} ({contador}){ext}")
        if not os.path.exists(candidato):
            return candidato


def _es_subruta_valida(carpeta_base: str, ruta: str) -> bool:
    try:
        ruta_abs = os.path.abspath(ruta)
        base_abs = os.path.abspath(carpeta_base)
        return os.path.commonpath([ruta_abs, base_abs]) == base_abs
    except Exception:
        return False


def _safe_subfolder_path(base_folder: str, subfolder_name: str) -> str:
    """Devuelve una ruta válida dentro de base_folder."""
    if not subfolder_name:
        subfolder_name = "Sin UE"

    subfolder_safe = sanitize_filename(str(subfolder_name).strip())
    m = re.match(r"^(.*) \((\d+)\)$", subfolder_safe)
    if m:
        subfolder_safe = m.group(1)
    if subfolder_safe in ("", ".", ".."):
        subfolder_safe = "Sin UE"

    carpeta_base_abs = os.path.abspath(base_folder)
    carpeta_destino = os.path.abspath(os.path.join(carpeta_base_abs, subfolder_safe))

    if not _es_subruta_valida(carpeta_base_abs, carpeta_destino):
        carpeta_destino = os.path.join(carpeta_base_abs, "Sin UE")
    carpeta_destino = os.path.abspath(carpeta_destino)

    os.makedirs(carpeta_destino, exist_ok=True)
    return carpeta_destino


def extraer_texto_pdf(ruta_pdf: str, max_paginas: int = 10) -> str:
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(ruta_pdf)
    except Exception:
        return ""
    partes = []
    for i, page in enumerate(reader.pages, start=1):
        if max_paginas is not None and i > max_paginas:
            break
        partes.append(page.extract_text() or "")
    return "\n".join(partes)

def detectar_tipo_documento(texto_raw: str) -> str:
    txt = normalizar_pdf(texto_raw)
    # PDFs de cargo/inconformidad de Osinergmin (tienen campo "Ubicación" aprovechable)
    if re.search(r"DETALLE\s+DE\s+LA\s+INCONFORMIDAD", txt[:400], re.IGNORECASE):
        return "cargo"
    # También tratar como cargo los documentos de tipo "CARGO DE DOCUMENTO INGRESADO"
    if re.search(r"CARGO\s*DE\s*DOCUMENTO\s*INGRESADO", txt, re.IGNORECASE):
        return "cargo"
    txt_sin_saltos = re.sub(r"\s+", " ", txt)
    if RE_OFICIO_COD.search(txt_sin_saltos[:600]):
        return "oficio"
    for linea in txt.splitlines()[:5]:
        if RE_CABECERA_OFICIO.search(linea):
            return "oficio"
    # Detectar por encabezado "OFICIO N° XXX" en cualquier línea inicial
    for linea in txt.splitlines()[:10]:
        if RE_OFICIO_ASUNTO_LINEA.search(linea):
            return "oficio"
    # Detectar por cuerpo típico de oficios (traslado, requerimiento, fiscalización)
    if RE_OFICIO_CUERPO.search(texto_raw[:1500]):
        return "oficio"
    primeras = txt_sin_saltos[:1000]
    es_resolucion = bool(re.search(r"RESOLUCI[OÓ]N", primeras, re.IGNORECASE))
    if not es_resolucion:
        return "otro"
    # Proceso sancionador / Resolución de sanción
    if re.search(
        r"PROCESO\s+SANCIONADOR|RESOLUCI[OÓ]N\s+DE\s+SANCI[OÓ]N|INICIO\s+DE\s+PROCEDIMIENTO\s+SANCIONADOR",
        txt_sin_saltos[:3000], re.IGNORECASE
    ):
        return "sancion"
    # Fuerza mayor: cubre resoluciones originales, recursos de reconsideración y apelación
    if re.search(
        r"ASUNTO\s*:\s*EVALUACI[OÓ]N\s+DE\s+SOLICITUD\s+DE\s+FUERZA\s+MAYOR",
        txt_sin_saltos[:2000], re.IGNORECASE
    ) or re.search(
        r"(?:CALIFICACI[OÓ]N\s+DE\s+SOLICITUDES?\s+DE\s+EXCLUSI[OÓ]N|"
        r"SOLICITUD\s+DE\s+(?:CALIFICACI[OÓ]N\s+DE\s+)?FUERZA\s+MAYOR|"
        r"RECURSO\s+DE\s+(?:RECONSIDERACI[OÓ]N|APELACI[OÓ]N).*?FUERZA\s+MAYOR|"
        r"FUERZA\s+MAYOR.*?RECURSO\s+DE\s+(?:RECONSIDERACI[OÓ]N|APELACI[OÓ]N))",
        txt_sin_saltos[:2000], re.IGNORECASE
    ):
        return "fuerza_mayor"
    # JARU: contiene "JARU" en el código de resolución o en el encabezado
    if (RE_RESOLUCION_JARU.search(txt[:1000]) or RE_TIENE_JARU.search(txt[:800])
            or RE_TIENE_JARU.search(txt_sin_saltos[:800])):
        return "jaru"
    if re.search(r"MATERIA\s*:", txt[:1000], re.IGNORECASE):
        return "resolucion"
    if re.search(r"COACTIVA", txt, re.IGNORECASE):
        return "coactiva"
    if re.search(r"INFORME\s+FINAL", txt, re.IGNORECASE):
        return "informe_final"
    if re.search(r"INICIO\s+DE\s+PROCESO\s+SANCIONADOR", txt, re.IGNORECASE):
        return "inicio_proceso_sancionador"
    if re.search(r"INFORME\s+(?:DE\s+)?FISCALIZACIÓN|INFORME\s+FISCALIZADOR|INFORME\s+DE\s+FISCALIZADOR", txt, re.IGNORECASE):
        return "informe_fiscalizacion"
    return "otro"

def extraer_oficio(texto_raw: str) -> dict:
    txt_norm = normalizar_pdf(texto_raw)
    resultado = {"codigo": "", "expediente": "", "asunto_pdf": ""}
    for linea in txt_norm.splitlines():
        m = RE_CABECERA_OFICIO.search(linea)
        if m:
            sede = (m.group("sede") or "").strip()
            codigo = m.group("codigo").strip()
            resultado["codigo"] = f"{codigo} {sede}".strip() if sede else codigo
            resultado["expediente"] = m.group("exp")
            break
    if not resultado["codigo"]:
        txt_sin_saltos = re.sub(r"\s+", " ", txt_norm)
        m_cod = RE_OFICIO_COD.search(txt_sin_saltos)
        if m_cod:
            resultado["codigo"] = m_cod.group("codigo").strip()
    m_asunto = RE_ASUNTO_OFICIO.search(texto_raw)
    if m_asunto:
        linea1 = m_asunto.group(1).strip()
        pos_fin = m_asunto.end()
        resto = texto_raw[pos_fin:].lstrip("\r\n")
        siguiente = resto.split("\n")[0].strip() if resto else ""
        es_etiqueta = ":" in siguiente[:40]
        es_cuerpo   = len(siguiente) > 35 and siguiente[:1].isupper()
        if siguiente and not es_etiqueta and not es_cuerpo:
            asunto = (linea1 + " " + siguiente).strip()
        else:
            asunto = linea1
        resultado["asunto_pdf"] = limpiar_texto(asunto)
    if not resultado["expediente"]:
        m_exp = RE_EXPEDIENTE_SUELTO.search(txt_norm)
        if m_exp:
            resultado["expediente"] = m_exp.group("exp")
    return resultado

def extraer_jaru(texto_raw: str) -> dict:
    txt_norm = normalizar_pdf(texto_raw)
    txt_sin_saltos = re.sub(r"\s+", " ", txt_norm)
    resultado = {"codigo": "", "expediente": "", "materia": ""}
    m = RE_RESOLUCION_JARU.search(txt_sin_saltos)
    if m:
        codigo = m.group("codigo").strip()
        codigo = re.sub(r"(\d{4})\s+([A-Z])", r"\1-\2", codigo)
        resultado["codigo"] = codigo
    m_mat = RE_MATERIA.search(texto_raw)
    if not m_mat:
        m_mat = RE_MATERIA.search(re.sub(r"\s+", " ", texto_raw))
    if m_mat:
        resultado["materia"] = limpiar_texto(re.sub(r"\s+", " ", m_mat.group(1)))
    m_exp = RE_EXPEDIENTE_SUELTO.search(txt_sin_saltos)
    if m_exp:
        resultado["expediente"] = m_exp.group("exp")
    return resultado

def extraer_resolucion(texto_raw: str) -> dict:
    txt_norm = normalizar_pdf(texto_raw)
    txt_sin_saltos = re.sub(r"\s+", " ", txt_norm)
    resultado = {"codigo": "", "materia": "", "expediente": ""}
    RE_RES_COD = re.compile(
        r"RESOLUCI[O\xd3]N\s+(?:\w+\s+){0,3}?(?:N[\xb0\xba]|N\u00ba|NRO\.?)\s*(?P<codigo>[\d][\d\-/]+)",
        re.IGNORECASE
    )
    m_cod = RE_RES_COD.search(texto_raw)
    if m_cod:
        resultado["codigo"] = m_cod.group("codigo").strip()
    RE_MAT = re.compile(
        r"MATERIA\s*:\s*(.+?)(?=\s*(?:OBLIGADO|DOMICILIO|ENTIDAD|OSINERGMIN|\Z))",
        re.IGNORECASE
    )
    m_mat = RE_MAT.search(txt_sin_saltos)
    if m_mat:
        resultado["materia"] = limpiar_texto(m_mat.group(1))
    RE_EXP_12 = re.compile(r"N[°º]\s*(?P<exp>\d{12})", re.IGNORECASE)
    m_exp = RE_EXP_12.search(texto_raw)
    if m_exp:
        resultado["expediente"] = m_exp.group("exp")
    else:
        m_exp2 = RE_EXPEDIENTE_SUELTO.search(txt_sin_saltos)
        if m_exp2:
            resultado["expediente"] = m_exp2.group("exp")
    return resultado

def extraer_fuerza_mayor(texto_raw: str) -> dict:
    txt_norm = normalizar_pdf(texto_raw)
    resultado = {"codigo": "", "sede": "", "expediente": "", "fallo": ""}
    RE_FUERZA_ORIG = re.compile(
        r"OSINERGMIN\s+(?:N[°º]|NO\.?|NRO\.?)\s*(?P<codigo>\d{2,6}-\d{4}-[A-Z0-9][A-Z0-9\-/\.]*)(?:\s+(?P<sede>[A-ZÁÉÍÓÚÜÑ]{2,10}))?",
        re.IGNORECASE
    )
    m = RE_FUERZA_ORIG.search(texto_raw[:1000])
    if not m:
        m = RE_FUERZA_ORIG.search(re.sub(r"\s+", " ", txt_norm)[:1000])
    if m:
        resultado["codigo"] = m.group("codigo")
        resultado["sede"]   = (m.group("sede") or "").strip()
    else:
        m = RE_RESOLUCION_FUERZA.search(txt_norm[:1000])
        if m:
            resultado["codigo"] = m.group("codigo")
            resultado["sede"]   = (m.group("sede") or "").strip()
    m_exp = RE_EXPEDIENTE_SUELTO.search(txt_norm)
    if m_exp:
        resultado["expediente"] = m_exp.group("exp")
    m_art2  = RE_ART2.search(texto_raw)
    limite  = m_art2.start() if m_art2 else len(texto_raw)
    m_fallo = RE_FALLO.search(texto_raw[:limite])
    if m_fallo:
        resultado["fallo"] = m_fallo.group(1).upper()
    return resultado

def construir_asunto(tipo: str, datos: dict) -> str:
    exp    = datos.get("expediente", "") or "N/D"
    codigo = datos.get("codigo", "")
    if tipo == "oficio":
        asunto_pdf = datos.get("asunto_pdf", "")
        partes = [f"Oficio {codigo}"] if codigo else []
        if asunto_pdf:
            partes.append(asunto_pdf)
        partes.append(f"Expediente: {exp}")
        return " \u2013 ".join(partes)
    elif tipo == "jaru":
        materia = datos.get("materia", "")
        partes  = [f"RESOLUCIÓN Nº {codigo}"] if codigo else ["RESOLUCIÓN"]
        if materia:
            partes.append(materia)
        partes.append(f"Expediente N° {exp}")
        return " - ".join(partes)
    elif tipo == "resolucion":
        materia = datos.get("materia", "")
        partes  = [f"RESOLUCIÓN Nº {codigo}"] if codigo else ["RESOLUCIÓN"]
        if materia:
            partes.append(materia)
        partes.append(f"Expediente N° {exp}")
        return " - ".join(partes)
    elif tipo == "fuerza_mayor":
        fallo = datos.get("fallo", "")
        sede  = datos.get("sede", "")
        cod_completo = f"{codigo} {sede}".strip() if sede else codigo
        base  = f"RESOLUCIÓN DE FUERZA MAYOR Nº {cod_completo}" if cod_completo else "RESOLUCIÓN DE FUERZA MAYOR"
        partes = [base]
        if fallo:
            partes.append(fallo)
        partes.append(f"Expediente N° {exp}")
        return " - ".join(partes)
    elif tipo == "sancion":
        materia = datos.get("materia", "")
        partes  = [f"RESOLUCIÓN DE SANCIÓN Nº {codigo}"] if codigo else ["RESOLUCIÓN DE SANCIÓN"]
        if materia:
            partes.append(materia)
        partes.append(f"Expediente N° {exp}")
        return " - ".join(partes)
    else:
        return f"Expediente: {exp}"

CUERPO_FMT = (
    "Estimados,\n\n"
    "Por medio del presente, se adjuntan los PDF para su atención\n"
    "y/o conocimiento, correspondientes al expediente N.° {expediente}.\n\n"
    "Asunto: {asunto_json}"
)

def construir_asunto_y_cuerpo(meta_path: str, carpeta: str) -> tuple:
    expediente_json = ""
    asunto_json     = ""
    if os.path.isfile(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            expediente_json = (meta.get("expediente") or "").strip()
            asunto_json     = limpiar_texto(meta.get("asunto") or "")
            if expediente_json and not re.fullmatch(r"\d{6,}", expediente_json):
                expediente_json = ""
        except Exception:
            pass

    PRIORIDAD = {"oficio": 4, "jaru": 3, "resolucion": 2, "fuerza_mayor": 1, "sancion": 1, "otro": 0}
    tipo_final      = "otro"
    datos_final     = {"expediente": expediente_json or "N/D"}
    prioridad_actual = 0

    for archivo in sorted(os.listdir(carpeta)):
        if not archivo.lower().endswith(".pdf"):
            continue
        ruta_pdf = os.path.join(carpeta, archivo)
        texto    = extraer_texto_pdf(ruta_pdf, max_paginas=5)
        if not texto.strip():
            continue
        tipo = detectar_tipo_documento(texto)
        if tipo == "oficio":
            datos = extraer_oficio(texto)
        elif tipo == "jaru":
            datos = extraer_jaru(texto)
        elif tipo == "resolucion":
            datos = extraer_resolucion(texto)
        elif tipo == "fuerza_mayor":
            datos = extraer_fuerza_mayor(texto)
        elif tipo == "sancion":
            datos = extraer_resolucion(texto)
        else:
            datos = {}
        if not datos.get("expediente") and expediente_json:
            datos["expediente"] = expediente_json
        tiene_codigo = bool(datos.get("codigo") or datos.get("asunto_pdf"))
        p = PRIORIDAD.get(tipo, 0)
        if tiene_codigo and p > prioridad_actual:
            tipo_final      = tipo
            datos_final     = datos
            prioridad_actual = p
        if tipo_final == "oficio" and tiene_codigo:
            break

    asunto_final    = construir_asunto(tipo_final, datos_final)
    exp_para_cuerpo = datos_final.get("expediente", expediente_json or "N/D")
    cuerpo_final    = CUERPO_FMT.format(
        expediente=exp_para_cuerpo,
        asunto_json=asunto_json or asunto_final
    )
    return tipo_final, asunto_final, cuerpo_final

# ══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización OSINERGMIN")
        self.root.geometry("1000x660")
        self.root.resizable(True, True)
        self._archivos_descargados = set()
        self._notificaciones_procesadas = set()

        frm_top = ttk.Frame(root)
        frm_top.pack(fill="x", padx=12, pady=10)
        ttk.Label(frm_top, text="Automatización OSINERGMIN", font=("Segoe UI", 13, "bold")).pack(side="left")
        self.lbl_time = ttk.Label(frm_top, text="")
        self.lbl_time.pack(side="right")
        self._tick_clock()

        frm_cfg = ttk.Frame(root)
        frm_cfg.pack(fill="x", padx=12, pady=8)
        ttk.Label(frm_cfg, text="Carpeta descargas:").grid(row=0, column=0, sticky="w", padx=(0,8))
        self.lbl_path = ttk.Label(frm_cfg, text=CARPETA_DESCARGA, foreground="blue")
        self.lbl_path.grid(row=0, column=1, sticky="w", padx=(0, 8))
        ttk.Button(frm_cfg, text="Cambiar ruta", command=self.on_cambiar_ruta).grid(row=0, column=2, sticky="w")


        ttk.Label(frm_cfg, text="Filtro:").grid(row=1, column=0, sticky="w", padx=(0,8), pady=(6,0))
        self.lbl_filtro = ttk.Label(frm_cfg, text="")
        self.lbl_filtro.grid(row=1, column=1, sticky="w", pady=(6,0))

        frm_prog = ttk.Frame(root)
        frm_prog.pack(fill="x", padx=12, pady=12)
        self.lbl_step = ttk.Label(frm_prog, text="Listo para iniciar…")
        self.lbl_step.pack(anchor="w")
        self.progress = ttk.Progressbar(frm_prog, orient="horizontal", length=820, mode="determinate")
        self.progress.pack(fill="x", pady=6)
        self.lbl_percent = ttk.Label(frm_prog, text="0%")
        self.lbl_percent.pack(anchor="e")

        frm_log = ttk.Frame(root)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0,12))

        frm_log_main = ttk.Frame(frm_log)
        frm_log_main.pack(side="left", fill="both", expand=True)
        ttk.Label(frm_log_main, text="Log principal:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.txt = tk.Text(frm_log_main, height=18, wrap="word")
        self.txt.pack(fill="both", expand=True)
        scroll_main = ttk.Scrollbar(frm_log_main, command=self.txt.yview)
        scroll_main.pack(side="right", fill="y")
        self.txt.config(yscrollcommand=scroll_main.set)

        frm_log_suministros = ttk.Frame(frm_log)
        frm_log_suministros.pack(side="left", fill="both", expand=True, padx=(8,0))
        ttk.Label(frm_log_suministros, text="Log suministros:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.txt_suministros = tk.Text(frm_log_suministros, height=18, wrap="word")
        self.txt_suministros.pack(fill="both", expand=True)
        scroll_sum = ttk.Scrollbar(frm_log_suministros, command=self.txt_suministros.yview)
        scroll_sum.pack(side="right", fill="y")
        self.txt_suministros.config(yscrollcommand=scroll_sum.set)

        frm_btn = ttk.Frame(root)
        frm_btn.pack(fill="x", padx=12, pady=(0,12))

        # Botón ENSA (azul)
        self.btn_start = ttk.Button(frm_btn, text="▼ Descargar OSINERGMIN", command=self.on_start)
        self.btn_start.pack(side="left")

        # Botón cancelar
        self.btn_cancel = ttk.Button(frm_btn, text="Cancelar", command=self.on_cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=(12,0))

        self._worker_thread = None
        self._cancel_event = threading.Event()
        self._total_objetivo = 0
        self._procesadas = 0
        self._registros_descarga = []
        self._fecha_inicio = None
        self._carpeta_descarga_actual = None
        self._tipo_descarga_actual = None
        self._identificador_actual = None

    def on_cambiar_ruta(self):
        global CARPETA_DESCARGA
        nueva_ruta = filedialog.askdirectory(
            title="Seleccionar carpeta para guardar archivos (dentro del proyecto recomendado)",
            initialdir=CARPETA_DESCARGA
        )
        if nueva_ruta:
            CARPETA_DESCARGA = _carpeta_interna(nueva_ruta)
            os.makedirs(CARPETA_DESCARGA, exist_ok=True)
            guardar_ruta(CARPETA_DESCARGA)
            self.lbl_path.config(text=CARPETA_DESCARGA)
            self.log(f"✅ Ruta de descarga actualizada: {CARPETA_DESCARGA}")

    def _tick_clock(self):
        self.lbl_time.config(text=datetime.now().strftime("%Y-%m-%d %H:%M"))
        self.root.after(60000, self._tick_clock)

    def ui(self, fn, *args, **kwargs):
        self.root.after(0, lambda: fn(*args, **kwargs))

    def _should_auto_scroll(self, widget):
        # Evita el salto cuando el usuario está leyendo en una posición antigua del log.
        first, last = widget.yview()
        return last >= 0.95

    def log(self, msg):
        def _write():
            at_bottom = self._should_auto_scroll(self.txt)
            self.txt.insert("end", msg + "\n")
            if at_bottom:
                self.txt.see("end")
        self.ui(_write)

    def set_step(self, text):
        self.ui(self.lbl_step.config, text=text)

    def set_progress(self, value_percent):
        value_percent = max(0, min(100, float(value_percent)))
        def _upd():
            self.progress["value"] = value_percent
            self.lbl_percent.config(text=f"{value_percent:.1f}%")
        self.ui(_upd)

    def log_suministros(self, msg):
        def _write():
            if hasattr(self, 'txt_suministros'):
                at_bottom = self._should_auto_scroll(self.txt_suministros)
                self.txt_suministros.insert("end", msg + "\n")
                if at_bottom:
                    self.txt_suministros.see("end")
            else:
                at_bottom = self._should_auto_scroll(self.txt)
                self.txt.insert("end", "[SUMINISTROS] " + msg + "\n")
                if at_bottom:
                    self.txt.see("end")
        self.ui(_write)

    def registrar_descarga(self, ruta_archivo):
        pass

    def registrar_en_excel(self, expediente, codigo_notificacion, unidad_operativa, procedimiento, asunto, fecha_notificacion, nombre_archivo, carpeta_destino=None):
        # No inferimos carpeta_destino desde 'unidad_operativa' porque ese campo
        # corresponde al Agente Supervisado de la página y puede no ser la UE.
        if carpeta_destino is None:
            carpeta_destino = ''

        registro = {
            'expediente': expediente,
            'codigo_notificacion': codigo_notificacion,
            'unidad_operativa': unidad_operativa,
            'procedimiento': procedimiento,
            'asunto': asunto,
            'fecha_notificacion': fecha_notificacion,
            'archivo': nombre_archivo,
            'carpeta_destino': carpeta_destino,
            'suministros': '',
            'suministro_principal': '',
        }
        self._registros_descarga.append(registro)
        if self._carpeta_descarga_actual:
            carpeta_notif = os.path.join(self._carpeta_descarga_actual, sanitize_filename(codigo_notificacion))
            meta_path = os.path.join(carpeta_notif, ".meta.json")
            try:
                if os.path.exists(meta_path):
                    with open(meta_path, 'r', encoding='utf-8') as f:
                        meta = json.load(f)
                else:
                    meta = {
                        'expediente': expediente,
                        'codigo_notificacion': codigo_notificacion,
                        'unidad_operativa': unidad_operativa,
                        'procedimiento': procedimiento,
                        'asunto': asunto,
                        'fecha_notificacion': fecha_notificacion,
                        'carpeta_destino': carpeta_destino,
                        'archivos': []
                    }
                if not isinstance(meta.get('archivos', []), list):
                    meta['archivos'] = []
                meta['carpeta_destino'] = carpeta_destino
                if nombre_archivo not in meta['archivos']:
                    meta['archivos'].append(nombre_archivo)
                with open(meta_path, 'w', encoding='utf-8') as f:
                    json.dump(meta, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

    def show_error_window(self, title, error_text):
        def _show():
            win = tk.Toplevel(self.root)
            win.title(title)
            win.geometry("780x420")
            ttk.Label(win, text=title, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=10, pady=10)
            txt_err = tk.Text(win, wrap="word")
            txt_err.pack(fill="both", expand=True, padx=10, pady=(0,10))
            txt_err.insert("end", error_text)
            txt_err.see("end")
            txt_err.config(state="disabled")
            ttk.Button(win, text="Cerrar", command=win.destroy).pack(pady=(0,12))
        self.ui(_show)

    # ── ENSA / OSINERGMIN ────────────────────────────────────────────────────
    def on_start(self):
        if self._worker_thread and self._worker_thread.is_alive():
            return

        # Leer credenciales ENSA del archivo (líneas 1 y 2)
        dni_guardado, clave_guardada = leer_credenciales_ensa()

        if dni_guardado and clave_guardada:
            usar_guardadas = messagebox.askyesno(
                "Credenciales",
                f"¿Usar credenciales guardadas?\n\nUsuario: {dni_guardado}\n\n"
                "(Las credenciales están en las líneas 1 y 2 de credenciales.txt)",
                parent=self.root
            )
            if usar_guardadas:
                dni = dni_guardado
                clave = clave_guardada
            else:
                dni = simpledialog.askstring("Credencial", "Ingrese su DNI/RUC/CE:", parent=self.root)
                clave = simpledialog.askstring("Credencial", "Ingrese su contraseña:", show="*", parent=self.root)
        else:
            dni = simpledialog.askstring("Credencial", "Ingrese su DNI/RUC/CE:", parent=self.root)
            clave = simpledialog.askstring("Credencial", "Ingrese su contraseña:", show="*", parent=self.root)

        if not dni or not clave:
            self.ui(messagebox.showerror, "Faltan datos", "Debe ingresar credenciales.")
            return

        opciones = ["Descargar NO LEÍDOS", "Descargar por Nro. Expediente", "Descargar por Fecha", "Descargar por Hora"]
        respuesta = self._mostrar_menu(opciones)

        if respuesta is None:
            return

        descargar_por_expediente = False
        descargar_por_fecha = False
        descargar_por_hora = False
        nro_expediente = None
        fecha_inicio = None
        fecha_fin = None

        if respuesta == 0:
            self.lbl_filtro.config(text="NO LEÍDOS (N)")
        elif respuesta == 1:
            descargar_por_expediente = True
            self.lbl_filtro.config(text="LEÍDOS")
            nro_expediente = simpledialog.askstring("Nro. Expediente", "Ingrese el Nro. de Expediente a buscar:", parent=self.root)
            if not nro_expediente:
                self.ui(messagebox.showerror, "Faltan datos", "Debe ingresar el Nro. de Expediente.")
                return
        elif respuesta == 2:
            descargar_por_fecha = True
            self.lbl_filtro.config(text="LEÍDOS")
            fecha_inicio = self._pedir_fecha("Fecha de Inicio (DD/MM/YYYY)")
            if not fecha_inicio:
                return
            fecha_fin = self._pedir_fecha("Fecha de Fin (DD/MM/YYYY)")
            if not fecha_fin:
                return
        elif respuesta == 3:
            descargar_por_hora = True
            ahora = datetime.now()
            ultima = cargar_ultima_hora()
            if ultima is None:
                desde = ahora.replace(hour=0, minute=0, second=0, microsecond=0)
                hasta = ahora
                msg = (f"Primera descarga por hora.\n\n"
                       f"Se descargará desde:\n"
                       f"  {desde.strftime('%d/%m/%Y 00:00:00')}\n"
                       f"hasta:\n"
                       f"  {hasta.strftime('%d/%m/%Y %H:%M:%S')}\n\n¿Continuar?")
            else:
                from datetime import timedelta
                desde = ultima + timedelta(seconds=1)
                hasta = ahora
                msg = (f"Última descarga: {ultima.strftime('%d/%m/%Y %H:%M:%S')}\n\n"
                       f"Se descargará desde:\n"
                       f"  {desde.strftime('%d/%m/%Y %H:%M:%S')}\n"
                       f"hasta:\n"
                       f"  {hasta.strftime('%d/%m/%Y %H:%M:%S')}\n\n¿Continuar?")
            if not messagebox.askyesno("Descargar por Hora", msg, parent=self.root):
                return
            fecha_inicio = desde.strftime("%d/%m/%Y %H:%M:%S")
            fecha_fin    = hasta.strftime("%d/%m/%Y %H:%M:%S")
            self.lbl_filtro.config(text=f"POR HORA: {desde.strftime('%d/%m %H:%M:%S')} → {hasta.strftime('%d/%m %H:%M:%S')}")

        self._cancel_event.clear()
        self._total_objetivo = 0
        self._procesadas = 0
        self._registros_descarga = []
        self._fecha_inicio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if respuesta == 0:
            self._tipo_descarga_actual = "no_leidos"
            self._identificador_actual = None
        elif respuesta == 1:
            self._tipo_descarga_actual = "expediente"
            self._identificador_actual = nro_expediente
        elif respuesta == 2:
            self._tipo_descarga_actual = "fecha"
            self._identificador_actual = f"{fecha_inicio}_{fecha_fin}"
        elif respuesta == 3:
            self._tipo_descarga_actual = "hora"
            self._identificador_actual = f"{fecha_inicio}_{fecha_fin}"

        self.set_step("Iniciando…")
        self.set_progress(0)
        self.btn_start.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.log(f"📂 Archivos se guardarán en: {CARPETA_DESCARGA}")

        self._worker_thread = threading.Thread(
            target=self._run_async_job_wrapper,
            args=(dni, clave, descargar_por_expediente, nro_expediente, descargar_por_fecha, fecha_inicio, fecha_fin, descargar_por_hora),
            daemon=True
        )
        self._worker_thread.start()

    def _mostrar_menu(self, opciones):
        resultado = [None]

        def _show():
            win = tk.Toplevel(self.root)
            win.title("Seleccionar Descarga")
            win.geometry("400x280")
            win.transient(self.root)
            win.grab_set()

            win.update_idletasks()
            x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
            y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
            win.geometry(f"+{x}+{y}")

            ttk.Label(win, text="¿Cómo desea descargar?", font=("Segoe UI", 12, "bold")).pack(pady=10)

            for idx, opcion in enumerate(opciones):
                def crear_callback(i):
                    def callback():
                        resultado[0] = i
                        win.destroy()
                    return callback
                ttk.Button(win, text=opcion, command=crear_callback(idx), width=30).pack(pady=5)

            ttk.Button(win, text="Cancelar", command=win.destroy, width=30).pack(pady=5)
            win.wait_window()

        self.ui(_show)
        while self.root.winfo_exists():
            try:
                self.root.update()
                if resultado[0] is not None:
                    break
            except:
                break

        return resultado[0]

    def _pedir_fecha(self, titulo):
        resultado = [None]

        def _show():
            win = tk.Toplevel(self.root)
            win.title(titulo)
            win.geometry("300x150")
            win.transient(self.root)
            win.grab_set()

            win.update_idletasks()
            x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
            y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
            win.geometry(f"+{x}+{y}")

            ttk.Label(win, text=titulo, font=("Segoe UI", 11, "bold")).pack(pady=10)

            frm = ttk.Frame(win)
            frm.pack(pady=10)

            hoy = datetime.now()
            dia_hoy = f"{hoy.day:02d}"
            mes_hoy = f"{hoy.month:02d}"
            año_hoy = str(hoy.year)

            ttk.Label(frm, text="Día:").grid(row=0, column=0)
            dia_var = tk.StringVar(value=dia_hoy)
            ttk.Combobox(frm, textvariable=dia_var, values=[f"{i:02d}" for i in range(1, 32)], width=5, state="readonly").grid(row=0, column=1, padx=5)

            ttk.Label(frm, text="Mes:").grid(row=0, column=2)
            mes_var = tk.StringVar(value=mes_hoy)
            ttk.Combobox(frm, textvariable=mes_var, values=[f"{i:02d}" for i in range(1, 13)], width=5, state="readonly").grid(row=0, column=3, padx=5)

            ttk.Label(frm, text="Año:").grid(row=0, column=4)
            año_var = tk.StringVar(value=año_hoy)
            ttk.Combobox(frm, textvariable=año_var, values=[str(i) for i in range(2020, 2031)], width=5, state="readonly").grid(row=0, column=5, padx=5)

            def aceptar():
                resultado[0] = f"{dia_var.get()}/{mes_var.get()}/{año_var.get()}"
                win.destroy()

            ttk.Button(win, text="Aceptar", command=aceptar).pack(side="left", padx=5, pady=10)
            ttk.Button(win, text="Cancelar", command=win.destroy).pack(side="right", padx=5, pady=10)
            win.wait_window()

        self.ui(_show)
        while self.root.winfo_exists():
            try:
                self.root.update()
                if resultado[0] is not None:
                    break
            except:
                break

        return resultado[0]

    def on_cancel(self):
        if self._worker_thread and self._worker_thread.is_alive():
            self._cancel_event.set()
            self.log("⏹ Cancelando…")
        else:
            self.btn_cancel.config(state="disabled")

    def _run_async_job_wrapper(self, dni, clave, descargar_por_expediente=False, nro_expediente=None, descargar_por_fecha=False, fecha_inicio=None, fecha_fin=None, descargar_por_hora=False):
        try:
            asyncio.run(self._run_async(dni, clave, descargar_por_expediente, nro_expediente, descargar_por_fecha, fecha_inicio, fecha_fin, descargar_por_hora))
        except Exception:
            tb = traceback.format_exc()
            self.log("❌ Error crítico. Abriendo detalles…")
            self.show_error_window("Error en la ejecución", tb)
        finally:
            self.ui(self.btn_start.config, state="normal")
            self.ui(self.btn_cancel.config, state="disabled")

    async def _run_async(self, dni, clave, descargar_por_expediente=False, nro_expediente=None, descargar_por_fecha=False, fecha_inicio=None, fecha_fin=None, descargar_por_hora=False):
        self._cancel_event.clear()
        intentos = 0
        max_intentos = 3

        while intentos < max_intentos:
            try:
                self.set_step("Cargando Playwright/Chromium…")
                async with async_playwright() as p:
                    try:
                        browser = await p.chromium.launch(headless=True, slow_mo=200)
                    except Exception as e:
                        try:
                            local_appdata = os.environ.get("LOCALAPPDATA", os.path.join(os.path.expanduser("~"), "AppData", "Local"))
                            chromium_path = os.path.join(local_appdata, "ms-playwright", "chromium-1208", "chrome-win64", "chrome.exe")
                            if not os.path.exists(chromium_path):
                                self.log("Chromium no encontrado. Descargando…")
                                self.set_step("Instalando Chromium…")
                                os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(local_appdata, "ms-playwright")
                                subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
                                self.log("✅ Chromium instalado. Intentando de nuevo…")
                            browser = await p.chromium.launch(headless=True, slow_mo=200, executable_path=chromium_path)
                        except Exception as e2:
                            msg = "No se pudo lanzar Chromium.\nIntente correr: playwright install chromium"
                            self.log(msg)
                            self.ui(messagebox.showerror, "Playwright/Chromium", msg)
                            return

                    context = await browser.new_context(accept_downloads=True)
                    page = await context.new_page()
                    self.set_step("Abriendo la página de acceso…")
                    await page.goto("https://osivirtual.osinergmin.gob.pe/autenticacion/acceso-sistema")
                    await espera_humana(page)
                    await page.fill("#documentoIdentidad", dni)
                    await espera_humana(page)
                    await page.fill("#contrasena", clave)
                    await espera_humana(page)
                    boton_ingresar = await page.wait_for_selector("text=Ingresar")
                    await mover_mouse_a_elemento(page, boton_ingresar)
                    await espera_humana(page)
                    await boton_ingresar.click()

                    login_ok = False
                    login_tiempo = 25_000

                    try:
                        await asyncio.wait_for(
                            page.wait_for_url(re.compile(r"./autenticacion/principal(?:\?.)?$"), timeout=login_tiempo),
                            timeout=login_tiempo/1000+3)
                        login_ok = True
                    except PlaywrightTimeoutError:
                        login_error = None
                        for selector in ["div.text-danger", ".invalid-feedback", ".alert-danger", ".alert-warning"]:
                            try:
                                el = await page.query_selector(selector)
                                if el:
                                    t = (await el.inner_text()).strip()
                                    if t:
                                        login_error = t
                                        break
                            except Exception:
                                pass
                        msg = login_error or "Usuario o contraseña incorrecto."
                        self.set_step("Error de credenciales.")
                        self.ui(messagebox.showerror, "Error de login", msg)
                        self.ui(self.btn_start.config, state="normal")
                        self.ui(self.btn_cancel.config, state="disabled")
                        await context.close()
                        await browser.close()
                        return

                    if not login_ok:
                        self.set_step("Error de credenciales.")
                        self.ui(messagebox.showerror, "Error de login", "No fue posible acceder.")
                        self.ui(self.btn_start.config, state="normal")
                        self.ui(self.btn_cancel.config, state="disabled")
                        await context.close()
                        await browser.close()
                        return

                    await espera_humana(page)

                    self.set_step("Abriendo Casilla Electrónica del SNE…")
                    casilla = await page.wait_for_selector("text=Casilla Electrónica del SNE", timeout=20000)
                    await mover_mouse_a_elemento(page, casilla)
                    await espera_humana(page)

                    pages_antes = len(context.pages)
                    await casilla.click()
                    await page.wait_for_timeout(1000)

                    if len(context.pages) > pages_antes:
                        new_page = context.pages[-1]
                        await new_page.wait_for_load_state()
                    else:
                        new_page = page
                    await espera_humana(new_page)

                    # Iniciar procesamiento de suministros en segundo plano mientras descarga sigue
                    self._suministros_bg_task = asyncio.create_task(self._procesar_suministros_en_segundo_plano())

                    if descargar_por_expediente and nro_expediente:
                        await self._procesar_expediente(new_page, nro_expediente)
                        self.log("⏳ Iniciando procesamiento de suministros (expediente)…")
                        try:
                            await self._procesar_suministros()
                        except Exception as e:
                            self.log(f"⚠ Error en _procesar_suministros: {e}")
                        await context.close()
                        await browser.close()
                        if not self._cancel_event.is_set():
                            await self._mostrar_dialogo_excel("expediente", nro_expediente)
                        return

                    if descargar_por_fecha and fecha_inicio and fecha_fin:
                        await self._procesar_por_fecha(new_page, fecha_inicio, fecha_fin)
                        self.log("⏳ Iniciando procesamiento de suministros (fecha)…")
                        try:
                            await self._procesar_suministros()
                        except Exception as e:
                            self.log(f"⚠ Error en _procesar_suministros: {e}")
                        await context.close()
                        await browser.close()
                        if not self._cancel_event.is_set():
                            await self._mostrar_dialogo_excel("fecha", f"{fecha_inicio}_{fecha_fin}")
                        return

                    if descargar_por_hora and fecha_inicio and fecha_fin:
                        await self._procesar_por_hora(new_page, fecha_inicio, fecha_fin)
                        self.log("⏳ Iniciando procesamiento de suministros (hora)…")
                        try:
                            await self._procesar_suministros()
                        except Exception as e:
                            self.log(f"⚠ Error: {e}")
                        await context.close()
                        await browser.close()
                        if not self._cancel_event.is_set():
                            await self._mostrar_dialogo_excel("hora", f"{fecha_inicio}_{fecha_fin}")
                        return

                    await self._procesar_no_leidos(new_page)
                    self.log("⏳ Iniciando procesamiento de suministros… (esto puede tardar)")
                    try:
                        await self._procesar_suministros()
                    except Exception as e:
                        self.log(f"⚠ Error en _procesar_suministros: {e}")
                        self.log(traceback.format_exc())
                    await context.close()
                    await browser.close()

                    if not self._cancel_event.is_set():
                        await self._mostrar_dialogo_excel("no_leidos", None)

                    return
            except Exception as e:
                import traceback
                intentos += 1
                tb = traceback.format_exc()
                self.log(f"⚠ Error en intento {intentos}/{max_intentos}: {str(e)[:100]}")

                if intentos < max_intentos:
                    def _preguntar():
                        return messagebox.askyesno(
                            "Error en la descarga",
                            f"Ocurrió un error.\n\n¿Desea continuar desde donde se quedó?\n\n(Procesadas: {self._procesadas}/{self._total_objetivo})",
                            parent=self.root
                        )

                    reintentar = self.ui(_preguntar)
                    retry_time = 0
                    while retry_time < 10 and self.root.winfo_exists():
                        try:
                            self.root.update()
                            retry_time += 0.1
                        except:
                            break

                    if not reintentar:
                        self.log("❌ Descarga cancelada por el usuario.")
                        self.show_error_window("Error", tb)
                        return

                    self.log(f"🔄 Reintentando… (intento {intentos + 1}/{max_intentos})")
                    await asyncio.sleep(3)
                else:
                    self.log("❌ Error crítico después de varios intentos.")
                    self.show_error_window("Error crítico", tb)
                    return
            finally:
                try:
                    await self._finalizar_procesamiento_suministros()
                except Exception as e:
                    self.log_suministros(f"⚠ Error en cleanup de suministros: {e}")

    # ── Procesar expediente ──────────────────────────────────────────────────
    async def _procesar_expediente(self, new_page, nro_expediente):
        import re as _re
        self.set_step(f"Buscando Expediente: {nro_expediente}")
        exp_input = await new_page.wait_for_selector('#expedienteSigedNotificacion', timeout=20000)
        await exp_input.fill("")
        await espera_humana(new_page)
        await exp_input.fill(nro_expediente)
        await espera_humana(new_page)
        select_element = await new_page.wait_for_selector("#leidoNotificacion", timeout=20000)
        await espera_humana(new_page)
        await select_element.select_option("")
        buscar_btn = await new_page.wait_for_selector("#buscar-boton", timeout=20000)
        await mover_mouse_a_elemento(new_page, buscar_btn)
        await espera_humana(new_page)
        await buscar_btn.click()
        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
        pager_text = await new_page.inner_text('#notificaciones-pager_right')
        m = _re.search(r'\d+ de (\d+)', pager_text)
        total_registros = int(m.group(1)) if m else 0
        self._total_objetivo = max(1, total_registros)
        self._procesadas = 0
        self.log(f"🔎 Total de notificaciones encontradas: {self._total_objetivo}")
        carpeta_expediente = os.path.join(CARPETA_DESCARGA, sanitize_filename(nro_expediente))
        os.makedirs(carpeta_expediente, exist_ok=True)
        self._carpeta_descarga_actual = carpeta_expediente
        progreso_anterior = cargar_progreso(carpeta_expediente)
        if progreso_anterior:
            self.log(f"📋 Progreso anterior encontrado")
            respuesta = [None]
            def _preguntar_continuar():
                respuesta[0] = messagebox.askyesno(
                    "Descarga Incompleta",
                    f"¿Continuar desde donde se quedó?\n\nNotificaciones: {len(progreso_anterior['notificaciones_procesadas'])}\nÚltima página: {progreso_anterior.get('pagina_actual', 1)}",
                    parent=self.root
                )
            self.ui(_preguntar_continuar)
            while respuesta[0] is None and self.root.winfo_exists():
                try:
                    self.root.update()
                    await asyncio.sleep(0.05)
                except:
                    break
            if respuesta[0] is True:
                self._notificaciones_procesadas = set(progreso_anterior['notificaciones_procesadas'])
                self._archivos_descargados = set(progreso_anterior['archivos_descargados'])
                self._procesadas = len(self._notificaciones_procesadas)
                self.log(f"✅ Continuando desde donde se quedó...")
            elif respuesta[0] is False:
                limpiar_progreso(carpeta_expediente)
                for item in os.listdir(carpeta_expediente):
                    ruta_item = os.path.join(carpeta_expediente, item)
                    if item not in [".progreso.json", "excel"]:
                        try:
                            if os.path.isdir(ruta_item):
                                shutil.rmtree(ruta_item)
                            else:
                                os.remove(ruta_item)
                        except:
                            pass
                self._notificaciones_procesadas = set()
                self._archivos_descargados = set()
                self._procesadas = 0
                self.log(f"🔄 Iniciando descarga desde cero...")
        else:
            self._notificaciones_procesadas = set()
            self._archivos_descargados = set()
            self._procesadas = 0
        pagina_actual = 1
        while True:
            if self._cancel_event.is_set():
                break
            try:
                await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
            except asyncio.TimeoutError:
                self.log("⏳ Tiempo de espera expirado. Reinentando...")
                continue
            filas = await new_page.locator('table#notificaciones-grid tr.jqgrow').all()
            if not filas:
                break
            filas_procesadas_en_esta_pagina = 0
            for fila in filas:
                if self._cancel_event.is_set():
                    break
                try:
                    expediente_cell = fila.locator('td[aria-describedby="notificaciones-grid_expedienteSigedNotificacion"]').first
                    codigo_cell = fila.locator('td[aria-describedby="notificaciones-grid_codigoNotificacion"]').first
                    unidad_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreUnidadOperativa"]').first
                    procedimiento_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreProcedimiento"]').first
                    asunto_cell = fila.locator('td[aria-describedby="notificaciones-grid_asuntoNotificacion"]').first
                    fecha_cell = fila.locator('td[aria-describedby="notificaciones-grid_fechaNotificacion"]').first
                    expediente_text = (await expediente_cell.inner_text(timeout=30000)).strip()
                    codigo_notificacion = sanitize_filename((await codigo_cell.inner_text(timeout=30000)).strip())
                    unidad_text = (await unidad_cell.inner_text(timeout=30000)).strip()
                    procedimiento_text = (await procedimiento_cell.inner_text(timeout=30000)).strip()
                    asunto_text = (await asunto_cell.inner_text(timeout=30000)).strip()
                    fecha_text = (await fecha_cell.inner_text(timeout=30000)).strip()
                except PlaywrightTimeoutError:
                    continue
                except Exception:
                    continue
                if codigo_notificacion in self._notificaciones_procesadas:
                    continue
                filas_procesadas_en_esta_pagina += 1
                i = self._procesadas + 1
                pagina_actual = (self._procesadas // 10) + 1
                self.log(f"➡ {i}/{self._total_objetivo} [Pag {pagina_actual}]")
                self.set_step(f"[Pag {pagina_actual}] {i}/{self._total_objetivo}…")
                carpeta_notif = os.path.join(carpeta_expediente, codigo_notificacion)
                os.makedirs(carpeta_notif, exist_ok=True)
                try:
                    lupa = fila.locator('img[title="Lectura de Notificación"]').first
                    await mover_mouse_a_elemento(new_page, await lupa.element_handle())
                    await espera_humana(new_page)
                    await lupa.click()
                    try:
                        await new_page.wait_for_url(_re.compile(r"./notificacion/verDetalle(?:\?.)?$"), timeout=45000)
                    except:
                        await new_page.wait_for_load_state("domcontentloaded", timeout=45000)
                    await espera_humana(new_page)
                    try:
                        constancia_link = await new_page.wait_for_selector("#verConstanciaNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, constancia_link)
                        await espera_humana(new_page)
                        async with new_page.expect_download() as ev1:
                            await constancia_link.click()
                        descarga = await ev1.value
                        ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                        self._archivos_descargados.add(os.path.basename(ruta))
                        self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                    except Exception:
                        pass
                    try:
                        doc_link = await new_page.wait_for_selector("#verDocumentosNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, doc_link)
                        await espera_humana(new_page)
                        await doc_link.click()
                        await new_page.wait_for_selector('a[title="Descargar archivo"]', timeout=30000)
                        enlaces_docs = new_page.locator('a[title="Descargar archivo"]')
                        total_docs = await enlaces_docs.count()
                        for idx in range(total_docs):
                            if self._cancel_event.is_set():
                                break
                            enlace = enlaces_docs.nth(idx)
                            try:
                                async with new_page.expect_download() as ev2:
                                    await enlace.click()
                                descarga = await ev2.value
                                ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                                self._archivos_descargados.add(os.path.basename(ruta))
                                self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                            except Exception:
                                pass
                        try:
                            cerrar = await new_page.wait_for_selector('span.ui-icon-closethick', timeout=8000)
                            await cerrar.click()
                        except:
                            pass
                    except Exception:
                        pass
                    try:
                        regresar = await new_page.wait_for_selector('#regresar-boton', timeout=20000)
                        await mover_mouse_a_elemento(new_page, regresar)
                        await espera_humana(new_page)
                        await regresar.click()
                    except Exception:
                        pass
                    # ── Clasificar por Unidad Empresarial ─────────────────────
                    try:
                        ue = detectar_ue_carpeta(carpeta_notif)
                        if MOVER_A_UE:
                            carpeta_notif = mover_notif_a_ue(carpeta_notif, carpeta_expediente, ue)
                            self.log(f"   📂 UE: {ue}")
                        else:
                            self.log(f"   📂 UE detectada (no movida): {ue}")
                        _actualizar_meta_carpeta(carpeta_notif, ue)
                        # Actualizar carpeta_destino en registros
                        for reg in self._registros_descarga:
                            if reg['codigo_notificacion'] == codigo_notificacion:
                                reg['carpeta_destino'] = ue
                                break
                    except Exception:
                        pass
                    self._notificaciones_procesadas.add(codigo_notificacion)
                    guardar_progreso(carpeta_expediente, "expediente", nro_expediente, self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
                    await new_page.wait_for_timeout(200)
                    self._procesadas += 1
                    self.set_progress(min(100, (self._procesadas / self._total_objetivo) * 100))
                except Exception:
                    guardar_progreso(carpeta_expediente, "expediente", nro_expediente, self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    continue
            pager_text = await new_page.inner_text('#notificaciones-pager_right')
            m = _re.search(r'(\d+) de (\d+)', pager_text)
            total_registros = int(m.group(2)) if m else 0
            if self._procesadas < total_registros:
                if filas_procesadas_en_esta_pagina == 0:
                    self.log(f"📄 Página ya procesada, avanzando…")
                try:
                    next_btn = await new_page.wait_for_selector('#next_notificaciones-pager', timeout=10000)
                    await next_btn.click()
                    await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
                except asyncio.TimeoutError:
                    self.log("⏳ Tiempo de espera expirado al avanzar página. ¿Reintentar?")
                    if self._preguntar_reintentar():
                        continue
                    else:
                        break
                except:
                    break
            else:
                break
        if not self._cancel_event.is_set():
            limpiar_progreso(carpeta_expediente)
            self.set_step("📊 Completado…")
            self.set_progress(100)
            self.log("✅ Completado.")

    async def _procesar_por_fecha(self, new_page, fecha_inicio, fecha_fin):
        import re as _re
        self.set_step(f"Buscando por fecha: {fecha_inicio} - {fecha_fin}")
        # Esperar a que la página esté completamente cargada
        await new_page.wait_for_load_state("networkidle", timeout=60000)
        fecha_inicio_input = await new_page.wait_for_selector('#fechaNotificacionInicio', timeout=60000)
        await fecha_inicio_input.fill(fecha_inicio)
        await espera_humana(new_page)
        fecha_fin_input = await new_page.wait_for_selector('#fechaNotificacionFin', timeout=60000)
        await fecha_fin_input.fill(fecha_fin)
        await espera_humana(new_page)
        select_element = await new_page.wait_for_selector("#leidoNotificacion", timeout=60000)
        await select_element.select_option("Todos")
        buscar_btn = await new_page.wait_for_selector("#buscar-boton", timeout=60000)
        await buscar_btn.click()
        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
        pager_text = await new_page.inner_text('#notificaciones-pager_right')
        m = _re.search(r'\d+ de (\d+)', pager_text)
        total_registros = int(m.group(1)) if m else 0
        self._total_objetivo = max(1, total_registros)
        self._procesadas = 0
        self.log(f"🔎 Total de notificaciones encontradas: {self._total_objetivo}")
        carpeta_fecha = os.path.join(CARPETA_DESCARGA, f"FECHA_{fecha_inicio}_{fecha_fin}".replace("/", "-"))
        os.makedirs(carpeta_fecha, exist_ok=True)
        self._carpeta_descarga_actual = carpeta_fecha
        progreso_anterior = cargar_progreso(carpeta_fecha)
        if progreso_anterior:
            respuesta = [None]
            def _preguntar_continuar():
                respuesta[0] = messagebox.askyesno(
                    "Descarga Incompleta",
                    f"¿Continuar desde donde se quedó?\n\nNotificaciones: {len(progreso_anterior['notificaciones_procesadas'])}\nÚltima página: {progreso_anterior.get('pagina_actual', 1)}",
                    parent=self.root
                )
            self.ui(_preguntar_continuar)
            while respuesta[0] is None and self.root.winfo_exists():
                try:
                    self.root.update()
                    await asyncio.sleep(0.05)
                except:
                    break
            if respuesta[0] is True:
                self._notificaciones_procesadas = set(progreso_anterior['notificaciones_procesadas'])
                self._archivos_descargados = set(progreso_anterior['archivos_descargados'])
                self._procesadas = len(self._notificaciones_procesadas)
            elif respuesta[0] is False:
                limpiar_progreso(carpeta_fecha)
                for item in os.listdir(carpeta_fecha):
                    ruta_item = os.path.join(carpeta_fecha, item)
                    if item not in [".progreso.json", "excel"]:
                        try:
                            if os.path.isdir(ruta_item): shutil.rmtree(ruta_item)
                            else: os.remove(ruta_item)
                        except: pass
                self._notificaciones_procesadas = set()
                self._archivos_descargados = set()
                self._procesadas = 0
        else:
            self._notificaciones_procesadas = set()
            self._archivos_descargados = set()
            self._procesadas = 0
        pagina_actual = 1
        while True:
            if self._cancel_event.is_set():
                break
            try:
                await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
            except asyncio.TimeoutError:
                self.log("⏳ Tiempo de espera expirado. Reinentando...")
                continue
            filas = await new_page.locator('table#notificaciones-grid tr.jqgrow').all()
            if not filas:
                break
            filas_procesadas_en_esta_pagina = 0
            for fila in filas:
                if self._cancel_event.is_set():
                    break
                try:
                    expediente_cell = fila.locator('td[aria-describedby="notificaciones-grid_expedienteSigedNotificacion"]').first
                    codigo_cell = fila.locator('td[aria-describedby="notificaciones-grid_codigoNotificacion"]').first
                    unidad_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreUnidadOperativa"]').first
                    procedimiento_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreProcedimiento"]').first
                    asunto_cell = fila.locator('td[aria-describedby="notificaciones-grid_asuntoNotificacion"]').first
                    fecha_cell = fila.locator('td[aria-describedby="notificaciones-grid_fechaNotificacion"]').first
                    expediente_text = (await expediente_cell.inner_text(timeout=30000)).strip()
                    codigo_notificacion = sanitize_filename((await codigo_cell.inner_text(timeout=30000)).strip())
                    unidad_text = (await unidad_cell.inner_text(timeout=30000)).strip()
                    procedimiento_text = (await procedimiento_cell.inner_text(timeout=30000)).strip()
                    asunto_text = (await asunto_cell.inner_text(timeout=30000)).strip()
                    fecha_text = (await fecha_cell.inner_text(timeout=30000)).strip()
                except PlaywrightTimeoutError:
                    continue
                except Exception:
                    continue
                if codigo_notificacion in self._notificaciones_procesadas:
                    continue
                filas_procesadas_en_esta_pagina += 1
                i = self._procesadas + 1
                pagina_actual = (self._procesadas // 10) + 1
                self.log(f"➡ {i}/{self._total_objetivo} [Pag {pagina_actual}]")
                self.set_step(f"[Pag {pagina_actual}] {i}/{self._total_objetivo}…")
                carpeta_notif = os.path.join(carpeta_fecha, codigo_notificacion)
                os.makedirs(carpeta_notif, exist_ok=True)
                try:
                    lupa = fila.locator('img[title="Lectura de Notificación"]').first
                    await mover_mouse_a_elemento(new_page, await lupa.element_handle())
                    await espera_humana(new_page)
                    await lupa.click()
                    try:
                        await new_page.wait_for_url(_re.compile(r"./notificacion/verDetalle(?:\?.)?$"), timeout=45000)
                    except:
                        await new_page.wait_for_load_state("domcontentloaded", timeout=45000)
                    await espera_humana(new_page)
                    try:
                        constancia_link = await new_page.wait_for_selector("#verConstanciaNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, constancia_link)
                        await espera_humana(new_page)
                        async with new_page.expect_download() as ev1:
                            await constancia_link.click()
                        descarga = await ev1.value
                        ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                        self._archivos_descargados.add(os.path.basename(ruta))
                        self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                    except Exception:
                        pass
                    try:
                        doc_link = await new_page.wait_for_selector("#verDocumentosNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, doc_link)
                        await espera_humana(new_page)
                        await doc_link.click()
                        await new_page.wait_for_selector('a[title="Descargar archivo"]', timeout=30000)
                        enlaces_docs = new_page.locator('a[title="Descargar archivo"]')
                        total_docs = await enlaces_docs.count()
                        for idx in range(total_docs):
                            if self._cancel_event.is_set():
                                break
                            try:
                                async with new_page.expect_download() as ev2:
                                    await enlaces_docs.nth(idx).click()
                                descarga = await ev2.value
                                ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                                self._archivos_descargados.add(os.path.basename(ruta))
                                self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                            except: pass
                        try:
                            cerrar = await new_page.wait_for_selector('span.ui-icon-closethick', timeout=8000)
                            await cerrar.click()
                        except: pass
                    except: pass
                    try:
                        regresar = await new_page.wait_for_selector('#regresar-boton', timeout=20000)
                        await mover_mouse_a_elemento(new_page, regresar)
                        await espera_humana(new_page)
                        await regresar.click()
                    except: pass
                    # ── Clasificar por Unidad Empresarial ─────────────────────
                    try:
                        ue = detectar_ue_carpeta(carpeta_notif)
                        if MOVER_A_UE:
                            carpeta_notif = mover_notif_a_ue(carpeta_notif, carpeta_fecha, ue)
                            self.log(f"   📂 UE: {ue}")
                        else:
                            self.log(f"   📂 UE detectada (no movida): {ue}")
                        _actualizar_meta_carpeta(carpeta_notif, ue)
                        # Actualizar carpeta_destino en todos los registros de esta notificación
                        for reg in self._registros_descarga:
                            if reg['codigo_notificacion'] == codigo_notificacion:
                                reg['carpeta_destino'] = ue
                    except Exception:
                        pass
                    self._notificaciones_procesadas.add(codigo_notificacion)
                    guardar_progreso(carpeta_fecha, "fecha", f"{fecha_inicio}_{fecha_fin}", self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
                    await new_page.wait_for_timeout(200)
                    self._procesadas += 1
                    self.set_progress(min(100, (self._procesadas / self._total_objetivo) * 100))
                except Exception:
                    guardar_progreso(carpeta_fecha, "fecha", f"{fecha_inicio}_{fecha_fin}", self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    continue
            try:
                pager_text = await new_page.inner_text('#notificaciones-pager_right', timeout=15000)
                m = _re.search(r'(\d+) de (\d+)', pager_text)
                total_registros = int(m.group(2)) if m else self._total_objetivo
            except Exception:
                self.log("⚠ No se encontró el paginador '#notificaciones-pager_right'. Usando total anterior.")
                total_registros = self._total_objetivo

            if self._procesadas < total_registros:
                if filas_procesadas_en_esta_pagina == 0:
                    self.log(f"📄 Página ya procesada, avanzando…")
                try:
                    next_btn = await new_page.wait_for_selector('#next_notificaciones-pager', timeout=10000)
                    await next_btn.click()
                    await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
                except asyncio.TimeoutError:
                    self.log("⏳ Tiempo de espera expirado al avanzar página. ¿Reintentar?")
                    if self._preguntar_reintentar():
                        continue
                    else:
                        break
                except: break
            else:
                break
        if not self._cancel_event.is_set():
            limpiar_progreso(carpeta_fecha)
            self.set_step("📊 Completado…")
            self.set_progress(100)
            self.log("✅ Completado.")

    async def _procesar_por_hora(self, new_page, fecha_hora_inicio: str, fecha_hora_fin: str):
        import re as _re
        from datetime import timedelta as _td
        self.set_step(f"Buscando por hora: {fecha_hora_inicio} → {fecha_hora_fin}")
        dt_desde = datetime.strptime(fecha_hora_inicio, "%d/%m/%Y %H:%M:%S")
        dt_hasta = datetime.strptime(fecha_hora_fin,    "%d/%m/%Y %H:%M:%S")
        fecha_str_desde = dt_desde.strftime("%d/%m/%Y")
        fecha_str_hasta = dt_hasta.strftime("%d/%m/%Y")

        async def fill_fecha(selector, valor):
            await new_page.wait_for_selector(selector, timeout=20000)
            await new_page.evaluate(f"""
                (function() {{
                    var el = document.querySelector('{selector}');
                    if (!el) return;
                    el.value = '{valor}';
                    el.dispatchEvent(new Event('input',  {{bubbles: true}}));
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                    el.dispatchEvent(new Event('blur',   {{bubbles: true}}));
                }})();
            """)
            await espera_humana(new_page)

        await fill_fecha('#fechaNotificacionInicio', fecha_str_desde)
        await fill_fecha('#fechaNotificacionFin',    fecha_str_hasta)
        await espera_humana(new_page)
        try:
            sel_leido = await new_page.wait_for_selector("#leidoNotificacion", timeout=5000)
            await sel_leido.select_option("")
        except: pass
        buscar_btn = await new_page.wait_for_selector("#buscar-boton", timeout=20000)
        await buscar_btn.click()
        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
        pager_text = await new_page.inner_text('#notificaciones-pager_right')
        m = _re.search(r'\d+ de (\d+)', pager_text)
        total_registros = int(m.group(1)) if m else 0
        self._total_objetivo = max(1, total_registros)
        self._procesadas = 0
        self.log(f"🔎 Total en rango de días: {self._total_objetivo} (se filtrarán por hora)")
        tag = (f"{dt_desde.strftime('%Y-%m-%d_%H-%M-%S')}"
               f"_a_{dt_hasta.strftime('%Y-%m-%d_%H-%M-%S')}")
        carpeta_hora = os.path.join(CARPETA_DESCARGA, f"HORA_{tag}")
        os.makedirs(carpeta_hora, exist_ok=True)
        self._carpeta_descarga_actual = carpeta_hora
        self._notificaciones_procesadas = set()
        self._archivos_descargados = set()
        pagina_actual = 1
        while True:
            if self._cancel_event.is_set():
                break
            try:
                await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
            except asyncio.TimeoutError:
                self.log("⏳ Tiempo de espera expirado. Reinentando...")
                continue
            except:
                await new_page.wait_for_timeout(2000)
            filas = await new_page.locator('table#notificaciones-grid tr.jqgrow').all()
            if not filas:
                break
            filas_procesadas_en_esta_pagina = 0
            for fila in filas:
                if self._cancel_event.is_set():
                    break
                try:
                    expediente_cell    = fila.locator('td[aria-describedby="notificaciones-grid_expedienteSigedNotificacion"]').first
                    codigo_cell        = fila.locator('td[aria-describedby="notificaciones-grid_codigoNotificacion"]').first
                    unidad_cell        = fila.locator('td[aria-describedby="notificaciones-grid_nombreUnidadOperativa"]').first
                    procedimiento_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreProcedimiento"]').first
                    asunto_cell        = fila.locator('td[aria-describedby="notificaciones-grid_asuntoNotificacion"]').first
                    fecha_cell         = fila.locator('td[aria-describedby="notificaciones-grid_fechaNotificacion"]').first
                    expediente_text     = (await expediente_cell.inner_text(timeout=15000)).strip()
                    codigo_notificacion = sanitize_filename((await codigo_cell.inner_text(timeout=15000)).strip())
                    unidad_text         = (await unidad_cell.inner_text(timeout=15000)).strip()
                    procedimiento_text  = (await procedimiento_cell.inner_text(timeout=15000)).strip()
                    asunto_text         = (await asunto_cell.inner_text(timeout=15000)).strip()
                    fecha_text          = (await fecha_cell.inner_text(timeout=15000)).strip()
                except Exception:
                    continue
                if codigo_notificacion in self._notificaciones_procesadas:
                    continue
                dt_fila = None
                for fmt in ("%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
                    try:
                        dt_fila = datetime.strptime(fecha_text.strip(), fmt)
                        break
                    except: continue
                if dt_fila is not None and not (dt_desde <= dt_fila <= dt_hasta):
                    continue
                filas_procesadas_en_esta_pagina += 1
                i = self._procesadas + 1
                self.log(f"➡ [Pag {pagina_actual}] {codigo_notificacion} ({fecha_text})")
                self.set_step(f"[Pag {pagina_actual}] {i} procesando…")
                carpeta_notif = os.path.join(carpeta_hora, codigo_notificacion)
                os.makedirs(carpeta_notif, exist_ok=True)
                try:
                    lupa = fila.locator('img[title="Lectura de Notificación"]').first
                    await mover_mouse_a_elemento(new_page, await lupa.element_handle())
                    await espera_humana(new_page)
                    await lupa.click()
                    try:
                        await new_page.wait_for_url(_re.compile(r"./notificacion/verDetalle(?:\?.)?$"), timeout=45000)
                    except:
                        await new_page.wait_for_load_state("domcontentloaded", timeout=45000)
                    await espera_humana(new_page)
                    try:
                        constancia_link = await new_page.wait_for_selector("#verConstanciaNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, constancia_link)
                        await espera_humana(new_page)
                        async with new_page.expect_download() as ev1:
                            await constancia_link.click()
                        descarga = await ev1.value
                        ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                        self._archivos_descargados.add(os.path.basename(ruta))
                        self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                    except: pass
                    try:
                        doc_link = await new_page.wait_for_selector("#verDocumentosNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, doc_link)
                        await espera_humana(new_page)
                        await doc_link.click()
                        await new_page.wait_for_selector('a[title="Descargar archivo"]', timeout=30000)
                        enlaces_docs = new_page.locator('a[title="Descargar archivo"]')
                        total_docs   = await enlaces_docs.count()
                        for idx in range(total_docs):
                            if self._cancel_event.is_set(): break
                            try:
                                async with new_page.expect_download() as ev2:
                                    await enlaces_docs.nth(idx).click()
                                descarga = await ev2.value
                                ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                                self._archivos_descargados.add(os.path.basename(ruta))
                                self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                            except: pass
                        try:
                            cerrar = await new_page.wait_for_selector('span.ui-icon-closethick', timeout=8000)
                            await cerrar.click()
                        except: pass
                    except: pass
                    try:
                        regresar = await new_page.wait_for_selector('#regresar-boton', timeout=20000)
                        await mover_mouse_a_elemento(new_page, regresar)
                        await espera_humana(new_page)
                        await regresar.click()
                    except: pass
                    # ── Clasificar por Unidad Empresarial ─────────────────────
                    try:
                        ue = detectar_ue_carpeta(carpeta_notif)
                        if MOVER_A_UE:
                            carpeta_notif = mover_notif_a_ue(carpeta_notif, carpeta_hora, ue)
                            self.log(f"   📂 UE: {ue}")
                        else:
                            self.log(f"   📂 UE detectada (no movida): {ue}")
                        _actualizar_meta_carpeta(carpeta_notif, ue)
                        # Actualizar carpeta_destino en todos los registros de esta notificación
                        for reg in self._registros_descarga:
                            if reg['codigo_notificacion'] == codigo_notificacion:
                                reg['carpeta_destino'] = ue
                    except Exception:
                        pass
                    self._notificaciones_procesadas.add(codigo_notificacion)
                    try:
                        await new_page.wait_for_load_state("domcontentloaded", timeout=30000)
                    except: pass
                    try:
                        await new_page.wait_for_selector("#notificaciones-grid", state="visible", timeout=30000)
                    except: pass
                    try:
                        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=60000)
                    except: pass
                    pager_check = await new_page.inner_text('#notificaciones-pager_right')
                    m_check = _re.search(r'(\d+)\s+de\s+\d+', pager_check)
                    current_pager_page = int(m_check.group(1)) if m_check else 1
                    if current_pager_page < pagina_actual:
                        for _ in range(pagina_actual - current_pager_page):
                            try:
                                nb = await new_page.wait_for_selector('#next_notificaciones-pager', timeout=10000)
                                await nb.click()
                                await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
                            except: break
                    await new_page.wait_for_timeout(300)
                    self._procesadas += 1
                    self.set_progress(min(99, (self._procesadas / self._total_objetivo) * 100))
                except Exception:
                    continue
            pager_text = await new_page.inner_text('#notificaciones-pager_right')
            m = _re.search(r'(\d+) de (\d+)', pager_text)
            total_registros = int(m.group(2)) if m else 0
            if self._procesadas < total_registros:
                pagina_actual = int(m.group(1)) if m else pagina_actual
                if filas_procesadas_en_esta_pagina == 0:
                    self.log(f"📄 Página {pagina_actual} sin registros en rango, avanzando…")
                try:
                    next_btn = await new_page.wait_for_selector('#next_notificaciones-pager', timeout=10000)
                    await next_btn.click()
                    await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
                    pagina_actual += 1
                except asyncio.TimeoutError:
                    self.log("⏳ Tiempo de espera expirado al avanzar página. ¿Reintentar?")
                    if self._preguntar_reintentar():
                        continue
                    else:
                        break
                except: break
            else:
                break
        if not self._cancel_event.is_set():
            guardar_ultima_hora(dt_hasta)
            self.log(f"💾 Última hora guardada: {dt_hasta.strftime('%d/%m/%Y %H:%M:%S')}")
            self.set_step("📊 Completado…")
            self.set_progress(100)
            self.log("✅ Completado.")

    async def _procesar_no_leidos(self, new_page):
        import re as _re
        self.set_step("Aplicando filtro: NO LEÍDAS…")
        select_element = await new_page.wait_for_selector("#leidoNotificacion", timeout=20000)
        await select_element.select_option("N")
        buscar_btn = await new_page.wait_for_selector("#buscar-boton", timeout=20000)
        await buscar_btn.click()
        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
        pager_text_total = await new_page.inner_text('#notificaciones-pager_right')
        m_total = _re.search(r'\d+ de (\d+)', pager_text_total)
        total_registros = int(m_total.group(1)) if m_total else 0
        if total_registros == 0:
            total_inicial = await new_page.evaluate("""
                () => {
                  return document.querySelectorAll('table#notificaciones-grid tr.jqgrow').length || 0;
                }
            """)
            total_registros = int(total_inicial)
        self._total_objetivo = max(1, total_registros)
        self.log(f"🔎 Total de notificaciones encontradas: {self._total_objetivo}")
        fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        carpeta_no_leidos = os.path.join(CARPETA_DESCARGA, f"NO_LEÍDOS_{fecha_hora}")
        os.makedirs(carpeta_no_leidos, exist_ok=True)
        self._carpeta_descarga_actual = carpeta_no_leidos
        progreso_anterior = cargar_progreso(carpeta_no_leidos)
        if progreso_anterior:
            respuesta = [None]
            def _preguntar_continuar():
                respuesta[0] = messagebox.askyesno(
                    "Descarga Incompleta",
                    f"¿Continuar desde donde se quedó?\n\nNotificaciones: {len(progreso_anterior['notificaciones_procesadas'])}\nÚltima página: {progreso_anterior.get('pagina_actual', 1)}",
                    parent=self.root
                )
            self.ui(_preguntar_continuar)
            while respuesta[0] is None and self.root.winfo_exists():
                try:
                    self.root.update()
                    await asyncio.sleep(0.05)
                except: break
            if respuesta[0] is True:
                self._notificaciones_procesadas = set(progreso_anterior['notificaciones_procesadas'])
                self._archivos_descargados = set(progreso_anterior['archivos_descargados'])
                self._procesadas = len(self._notificaciones_procesadas)
                self.log(f"✅ Continuando desde donde se quedó...")
            elif respuesta[0] is False:
                limpiar_progreso(carpeta_no_leidos)
                for item in os.listdir(carpeta_no_leidos):
                    ruta_item = os.path.join(carpeta_no_leidos, item)
                    if item != ".progreso.json":
                        try:
                            if os.path.isdir(ruta_item): shutil.rmtree(ruta_item)
                            else: os.remove(ruta_item)
                        except: pass
                self._notificaciones_procesadas = set()
                self._archivos_descargados = set()
                self._procesadas = 0
                self.log(f"🔄 Iniciando descarga desde cero...")
        else:
            self._notificaciones_procesadas = set()
            self._archivos_descargados = set()
            self._procesadas = 0
        pagina_actual = 1
        while True:
            if self._cancel_event.is_set():
                break
            try:
                await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
            except asyncio.TimeoutError:
                self.log("⏳ Tiempo de espera expirado. Reinentando...")
                continue
            except PlaywrightTimeoutError:
                await new_page.wait_for_timeout(2000)
                continue
            filas = await new_page.locator('table#notificaciones-grid tr.jqgrow').all()
            if not filas:
                break
            filas_procesadas_en_esta_pagina = 0
            for fila in filas:
                if self._cancel_event.is_set():
                    break
                try:
                    expediente_cell = fila.locator('td[aria-describedby="notificaciones-grid_expedienteSigedNotificacion"]').first
                    codigo_cell = fila.locator('td[aria-describedby="notificaciones-grid_codigoNotificacion"]').first
                    unidad_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreUnidadOperativa"]').first
                    procedimiento_cell = fila.locator('td[aria-describedby="notificaciones-grid_nombreProcedimiento"]').first
                    asunto_cell = fila.locator('td[aria-describedby="notificaciones-grid_asuntoNotificacion"]').first
                    fecha_cell = fila.locator('td[aria-describedby="notificaciones-grid_fechaNotificacion"]').first
                    expediente_text = (await expediente_cell.inner_text(timeout=30000)).strip()
                    codigo_notificacion = sanitize_filename((await codigo_cell.inner_text(timeout=30000)).strip())
                    unidad_text = (await unidad_cell.inner_text(timeout=30000)).strip()
                    procedimiento_text = (await procedimiento_cell.inner_text(timeout=30000)).strip()
                    asunto_text = (await asunto_cell.inner_text(timeout=30000)).strip()
                    fecha_text = (await fecha_cell.inner_text(timeout=30000)).strip()
                except PlaywrightTimeoutError:
                    continue
                except Exception:
                    continue
                if codigo_notificacion in self._notificaciones_procesadas:
                    continue
                filas_procesadas_en_esta_pagina += 1
                i = self._procesadas + 1
                pagina_actual = (self._procesadas // 10) + 1
                self.log(f"➡ {i}/{self._total_objetivo} [Pag {pagina_actual}]")
                self.set_step(f"[Pag {pagina_actual}] {i}/{self._total_objetivo}…")
                carpeta_notif = os.path.join(carpeta_no_leidos, codigo_notificacion)
                os.makedirs(carpeta_notif, exist_ok=True)
                try:
                    lupa = fila.locator('img[title="Lectura de Notificación"]').first
                    await mover_mouse_a_elemento(new_page, await lupa.element_handle())
                    await espera_humana(new_page)
                    await lupa.click()
                    try:
                        await new_page.wait_for_url(_re.compile(r"./notificacion/verDetalle(?:\?.)?$"), timeout=45000)
                    except:
                        await new_page.wait_for_load_state("domcontentloaded", timeout=45000)
                    await espera_humana(new_page)
                    try:
                        constancia_link = await new_page.wait_for_selector("#verConstanciaNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, constancia_link)
                        await espera_humana(new_page)
                        async with new_page.expect_download() as ev1:
                            await constancia_link.click()
                        descarga = await ev1.value
                        ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                        self._archivos_descargados.add(os.path.basename(ruta))
                        self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                    except Exception:
                        pass
                    try:
                        doc_link = await new_page.wait_for_selector("#verDocumentosNotificacion-link", timeout=20000)
                        await mover_mouse_a_elemento(new_page, doc_link)
                        await espera_humana(new_page)
                        await doc_link.click()
                        await new_page.wait_for_selector('a[title="Descargar archivo"]', timeout=30000)
                        enlaces_docs = new_page.locator('a[title="Descargar archivo"]')
                        total_docs = await enlaces_docs.count()
                        for idx in range(total_docs):
                            if self._cancel_event.is_set(): break
                            enlace = enlaces_docs.nth(idx)
                            try:
                                async with new_page.expect_download() as ev2:
                                    await enlace.click()
                                descarga = await ev2.value
                                ruta = await guardar_descarga(descarga, carpeta_notif, self.registrar_descarga, self.log)
                                self._archivos_descargados.add(os.path.basename(ruta))
                                self.registrar_en_excel(expediente_text, codigo_notificacion, unidad_text, procedimiento_text, asunto_text, fecha_text, os.path.basename(ruta))
                            except: pass
                        try:
                            cerrar = await new_page.wait_for_selector('span.ui-icon-closethick', timeout=8000)
                            await cerrar.click()
                        except: pass
                    except: pass
                    try:
                        regresar = await new_page.wait_for_selector('#regresar-boton', timeout=20000)
                        await mover_mouse_a_elemento(new_page, regresar)
                        await espera_humana(new_page)
                        await regresar.click()
                    except: pass
                    # ── Clasificar por Unidad Empresarial ─────────────────────
                    try:
                        ue = detectar_ue_carpeta(carpeta_notif)
                        if MOVER_A_UE:
                            carpeta_notif = mover_notif_a_ue(carpeta_notif, carpeta_no_leidos, ue)
                            self.log(f"   📂 UE: {ue}")
                        else:
                            self.log(f"   📂 UE detectada (no movida): {ue}")
                        _actualizar_meta_carpeta(carpeta_notif, ue)
                        # Actualizar carpeta_destino en todos los registros de esta notificación
                        for reg in self._registros_descarga:
                            if reg['codigo_notificacion'] == codigo_notificacion:
                                reg['carpeta_destino'] = ue
                    except Exception:
                        pass
                    self._notificaciones_procesadas.add(codigo_notificacion)
                    guardar_progreso(carpeta_no_leidos, "no_leidos", None, self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    try:
                        await new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000)
                    except: pass
                    await new_page.wait_for_timeout(200)
                    self._procesadas += 1
                    self.set_progress(min(100, (self._procesadas / self._total_objetivo) * 100))
                except Exception:
                    guardar_progreso(carpeta_no_leidos, "no_leidos", None, self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)
                    continue
            pager_text = await new_page.inner_text('#notificaciones-pager_right')
            m = _re.search(r'(\d+) de (\d+)', pager_text)
            total_registros = int(m.group(2)) if m else 0
            if self._procesadas < total_registros:
                if filas_procesadas_en_esta_pagina == 0:
                    self.log(f"📄 Página ya procesada, avanzando…")
                try:
                    next_btn = await new_page.wait_for_selector('#next_notificaciones-pager', timeout=10000)
                    await next_btn.click()
                    await asyncio.wait_for(new_page.wait_for_selector("#load_notificaciones-grid", state="hidden", timeout=180000), timeout=200)
                except asyncio.TimeoutError:
                    self.log("⏳ Tiempo de espera expirado al avanzar página. ¿Reintentar?")
                    if self._preguntar_reintentar():
                        continue
                    else:
                        break
                except: break
            else:
                break
            def _preguntar_reintentar(self):
                respuesta = [None]
                def _show():
                    respuesta[0] = messagebox.askyesno(
                        "Tiempo de espera expirado",
                        "¿Desea reintentar la descarga desde donde se quedó?",
                        parent=self.root
                    )
                self.ui(_show)
                while respuesta[0] is None and self.root.winfo_exists():
                    try:
                        self.root.update()
                    except:
                        break
                return respuesta[0] is True
        if not self._cancel_event.is_set():
            limpiar_progreso(carpeta_no_leidos)
            self.set_step("📊 Completado…")
            self.set_progress(100)
            self.log("✅ Completado.")
        else:
            guardar_progreso(carpeta_no_leidos, "no_leidos", None, self._notificaciones_procesadas, self._archivos_descargados, pagina_actual)

    async def _ir_a_siguiente_pagina_notificaciones(self, new_page):
        """Avanza a la siguiente página de notificaciones en casos de paginación JS."""
        selectores = [
            '#next_notificaciones-pager',
            '#notificaciones-pager a.ui-pg-button',
            '#notificaciones-pager a:has-text("Siguiente")',
            '#notificaciones-pager a:has-text("Next")',
            '#notificaciones-pager-left',
        ]
        for sel in selectores:
            try:
                btn = await new_page.query_selector(sel)
                if btn:
                    try:
                        enabled = await btn.is_enabled()
                    except:
                        enabled = True
                    if enabled:
                        await mover_mouse_a_elemento(new_page, btn)
                        await espera_humana(new_page)
                        await btn.click()
                        return True
            except Exception:
                pass

        # fallback: click en cualquier botón con texto Next/Siguiente dentro del pager
        try:
            encontrado = await new_page.evaluate("""
                () => {
                    const pager = document.querySelector('#notificaciones-pager') || document.querySelector('[id$="-pager"]');
                    if (!pager) return false;
                    const cand = Array.from(pager.querySelectorAll('a, button')).find(e => /siguiente|next/i.test(e.innerText));
                    if (!cand) return false;
                    cand.click();
                    return true;
                }
            """)
            return bool(encontrado)
        except Exception:
            return False

    async def _consultar_recibo(self, suministro):
        # Funcion desactivada: no se consulta la plataforma externa.
        # Solo devolvemos 'desconocido' para que el procesamiento de suministros
        # quede centrado en el número detectado en PDF.
        self.log_suministros(f"🔍 [SUMINISTROS] _consultar_recibo desactivado para suministro: {suministro}")
        return "desconocido"

    async def _procesar_suministros(self):
        self.log_suministros("🔍 [SUMINISTROS] Iniciando procesamiento de suministros en carpetas Sin UE…")
        carpeta_base = self._carpeta_descarga_actual or CARPETA_DESCARGA
        carpeta_base = _normalizar_ruta(carpeta_base)
        self.log_suministros(f"🔍 [SUMINISTROS] Carpeta base: {carpeta_base}")
        if not os.path.exists(carpeta_base):
            self.log_suministros(f"❌ [SUMINISTROS] Carpeta no existe: {carpeta_base}")
            return
        self.log_suministros(f"✅ [SUMINISTROS] Carpeta existe. Escaneando…")

        carpetas_notif = []
        for root, dirs, files in os.walk(carpeta_base):
            # No procesar contenido de carpeta excel
            if os.path.basename(root).lower() == 'excel':
                dirs[:] = []
                continue
            if '.meta.json' in files:
                carpetas_notif.append(root)

        self.log_suministros(f"🔍 [SUMINISTROS] Carpetas con .meta.json encontradas: {len(carpetas_notif)}")

        for ruta_notif in sorted(carpetas_notif):
            if self._cancel_event.is_set():
                break
            nombre_notif = os.path.basename(ruta_notif)
            if ruta_notif == carpeta_base:
                continue
            meta_path = os.path.join(ruta_notif, '.meta.json')

            unidad_operativa = ''
            suministros_actuales = []
            try:
                with open(meta_path, 'r', encoding='utf-8') as f:
                    meta = json.load(f)
                unidad_operativa = meta.get('unidad_operativa', '')
                suministros_actuales = meta.get('suministros', [])
            except Exception as e:
                self.log_suministros(f"⚠ [SUMINISTROS] Error leyendo metadata {meta_path}: {e}")

            # Reprocesar si no se tienen suministros grabados o si es carpeta Sin UE
            if unidad_operativa.strip() and suministros_actuales:
                continue

            self.log_suministros(f"📂 Procesando carpeta (UE={unidad_operativa or 'Sin UE'}): {ruta_notif}")
            suministros_encontrados = set()
            pdfs_en_carpeta = [f for f in os.listdir(ruta_notif) if f.lower().endswith('.pdf')]
            self.log_suministros(f"  PDFs encontrados: {pdfs_en_carpeta}")
            for archivo in pdfs_en_carpeta:
                ruta_pdf = os.path.join(ruta_notif, archivo)
                try:
                    from PyPDF2 import PdfReader
                    reader = PdfReader(ruta_pdf)
                    texto = ""
                    for page in reader.pages:
                        page_text = page.extract_text() or ""
                        texto += page_text + "\n"
                    # Log para debugging
                    self.log_suministros(f"  [DEBUG] Primeras 300 chars de {archivo}: {texto[:300]}")
                    # Patrón simple y permisivo para detectar suministros
                    matches = re.findall(
                        r"(?:Suministro|Código\s+Suministro)[\s:°NnRro\-()]*([0-9]{7,8})",
                        texto,
                        flags=re.IGNORECASE
                    )
                    self.log_suministros(f"  Escaneando {archivo}: {len(matches)} suministro(s) encontrado(s) = {matches}")
                    for m in matches:
                        suministros_encontrados.add(m.strip())
                except Exception as e:
                    self.log_suministros(f"Error procesando {ruta_pdf}: {e}")

            if not suministros_encontrados:
                self.log_suministros(f"  ❌ No se encontraron suministros en {nombre_notif}")
                suministros_meta = []
            else:
                suministros_meta = sorted(suministros_encontrados)

            # Guardar información de suministros en metadata para reporte
            try:
                if os.path.exists(meta_path):
                    with open(meta_path, 'r', encoding='utf-8') as f:
                        meta = json.load(f)
                else:
                    meta = {}
                meta['suministros'] = suministros_meta
                meta['suministro_principal'] = suministros_meta[0] if suministros_meta else ''
                with open(meta_path, 'w', encoding='utf-8') as f:
                    json.dump(meta, f, ensure_ascii=False, indent=2)
            except Exception as e:
                self.log_suministros(f"⚠ Error guardando suministros en meta {meta_path}: {e}")

            if not suministros_encontrados:
                ue_fallback = detectar_ue_carpeta(ruta_notif)
                if ue_fallback and ue_fallback != "Sin UE" and MOVER_A_UE:
                    carpeta_destino = _safe_subfolder_path(carpeta_base, ue_fallback)
                    nueva_ruta = os.path.join(carpeta_destino, nombre_notif)
                    try:
                        if os.path.exists(nueva_ruta):
                            # Fusionar contenido en la carpeta existente sin renombrar al estilo (1) (2)...
                            for item in os.listdir(ruta_notif):
                                origen_item = os.path.join(ruta_notif, item)
                                destino_item = os.path.join(nueva_ruta, item)
                                if os.path.exists(destino_item):
                                    destino_item = _get_unique_path(destino_item)
                                shutil.move(origen_item, destino_item)
                            try:
                                os.rmdir(ruta_notif)
                            except Exception:
                                pass
                            self.log_suministros(f"Clasificado en fallback {ue_fallback}: fusionado en {nueva_ruta}")
                        else:
                            shutil.move(ruta_notif, nueva_ruta)
                            self.log_suministros(f"Clasificado en fallback {ue_fallback}: {nueva_ruta}")
                    except Exception as e:
                        self.log_suministros(f"Error al mover carpeta {ruta_notif} a {nueva_ruta}: {e}")
                else:
                    self.log_suministros(f"Clasificado en fallback {ue_fallback}: no movido")
                continue

            self.log_suministros(f"  ✅ Total suministros encontrados: {len(suministros_encontrados)}")
            self.log_suministros(f"  📌 Números de suministro: {', '.join(suministros_meta)}")

            # No se procesa más en la plataforma externa ni se clasifica en esta fase.
            # Dejar en Sin UE para revisión posterior.
            continue

    async def _procesar_suministros_en_segundo_plano(self):
        self._cancel_event.clear()
        while not self._cancel_event.is_set():
            try:
                await self._procesar_suministros()
            except asyncio.CancelledError:
                break
            except Exception as e:
                self.log_suministros(f"⚠ [SUMINISTROS] Error en bg: {e}")
            await asyncio.sleep(5)

    async def _finalizar_procesamiento_suministros(self):
        if not hasattr(self, '_suministros_bg_task') or self._suministros_bg_task is None:
            return
        self.log_suministros("🛑 Finalizando procesamiento de suministros en segundo plano...")
        # Esperar múltiples ciclos para asegurar que procesa todo completamente
        # El loop en background procesa cada 5 segundos, así que esperar 30+ segundos 
        for _ in range(35):
            if self._suministros_bg_task.done():
                break
            await asyncio.sleep(1)
        self._cancel_event.set()
        self._suministros_bg_task.cancel()
        try:
            await self._suministros_bg_task
        except asyncio.CancelledError:
            pass
        except Exception as e:
            self.log_suministros(f"⚠ Error al detener bg suministros: {e}")

    async def _mostrar_dialogo_excel(self, tipo_descarga, identificador):
        def _show():
            resp = messagebox.askyesno(
                "Generar Reporte",
                "¿Desea generar archivo Excel con los datos descargados?",
                parent=self.root
            )
            if resp:
                self.set_step("📊 Generando Excel…")
                fecha_desc = self._fecha_inicio or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if self._carpeta_descarga_actual:
                    ruta = crear_excel_reporte(
                        self._registros_descarga,
                        self._carpeta_descarga_actual,
                        fecha_desc,
                        tipo_descarga,
                        identificador
                    )
                    if ruta:
                        self.log(f"✅ Excel guardado: {ruta}")
                        self.set_step("🎉 ¡Finalizado!")
                        messagebox.showinfo("✅ Reporte Generado", f"Guardado en:\n{ruta}")
                    else:
                        self.log("❌ No se pudo generar el Excel")
            else:
                self.set_step("🎉 ¡Finalizado!")
                self.log("📋 Excel no fue generado.")
            # ── Preguntar si enviar correo ────────────────────────────────
            destinatarios = leer_destinatarios_correo()
            if not destinatarios:
                self.log("📧 Sin destinatarios configurados (edita destinatarios_correo.txt)")
                return
            resp_correo = messagebox.askyesno(
                "Enviar correo",
                f"¿Desea enviar un correo de notificación?\n\nDestinatarios ({len(destinatarios)}):\n" +
                "\n".join(f"  • {d}" for d in destinatarios[:5]) +
                ("\n  ..." if len(destinatarios) > 5 else ""),
                parent=self.root
            )
            if resp_correo:
                self._enviar_correo_notificacion()
            else:
                self.log("📧 Correo no enviado.")
        self.ui(_show)

    def _enviar_correo_notificacion(self):
        try:
            asunto_tmpl, cuerpo_tmpl = leer_template_correo()
            destinatarios = leer_destinatarios_correo()
            fecha_str   = datetime.now().strftime("%d/%m/%Y %H:%M")
            total_str   = str(self._procesadas)
            carpeta_str = self._carpeta_descarga_actual or CARPETA_DESCARGA
            def _fmt(s):
                return s.replace("{fecha}", fecha_str).replace("{total}", total_str).replace("{carpeta}", carpeta_str)
            asunto = _fmt(asunto_tmpl) if asunto_tmpl else f"Notificaciones OSINERGMIN - {fecha_str}"
            cuerpo = _fmt(cuerpo_tmpl)
            self.log("📧 Enviando correo…")
            enviar_correo_outlook(asunto, cuerpo, destinatarios, self.log)
        except Exception as e:
            self.log(f"⚠ Error al enviar correo: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.call("source", "sun-valley.tcl")
        ttk.Style().theme_use("sun-valley-dark")
    except:
        pass
    app = App(root)
    root.mainloop() 

#!/usr/bin/env python3
"""
Lee archivos de Ofertas/ con formato estandar (ELEIA, Iberdrola, etc.)
y llena o CREA plantillas en MONOPUNTO/ y MULTIPUNTO/.

Formato de oferta esperado:
  - Hoja: PETICION OFERTAS
  - Columnas: B=Cliente, D=Codigo, I=CUPS, J=Tarifa, Z-AE=P1-P6, AF=FEE, AG=Coste
  - Datos desde fila 8
  - Codigo aparece 1 vez = MONOPUNTO, >1 vez = MULTIPUNTO
"""

import os
import sys
import re
import threading
from datetime import date
from collections import defaultdict

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    sys.stderr.write(
        "Falta la dependencia 'openpyxl'. Instálela con:  pip install -r requirements.txt\n"
        f"Detalle: {exc}\n"
    )
    sys.exit(1)


# ---------------------------------------------------------------------------
# LOG Y CANCELACION
# ---------------------------------------------------------------------------

# levels: info | ok | warn | error
_LOG_HOOK = None

# Se activa para que el usuario pueda detener un proceso largo.
CANCELAR = threading.Event()


def set_log_hook(fn):
    """Registra un callback (nivel, mensaje) para el log de la interfaz.
    Los mensajes se siguen mostrando por consola (comportamiento original)."""
    global _LOG_HOOK
    _LOG_HOOK = fn


def log(mensaje, nivel="info"):
    """Escribe un mensaje y lo notifica al hook de la interfaz."""
    print(mensaje, end="")
    if _LOG_HOOK is not None:
        try:
            _LOG_HOOK(nivel, mensaje)
        except Exception:
            pass


def hubo_cancelacion():
    return CANCELAR.is_set()


def obtener_base_dir():
    """Obtiene el directorio base, compatible con PyInstaller."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = obtener_base_dir()
OFERTAS_DIR = os.path.join(BASE_DIR, "Ofertas")
MONOPUNTO_DIR = os.path.join(BASE_DIR, "MONOPUNTO")
MULTIPUNTO_DIR = os.path.join(BASE_DIR, "MULTIPUNTO")

# Columnas en ofertas (1-based)
OF_CLIENTE, OF_CODIGO, OF_CUPS, OF_TARIFA = 2, 4, 9, 10
OF_P1, OF_P2, OF_P3, OF_P4, OF_P5, OF_P6 = 26, 27, 28, 29, 30, 31
OF_FEE, OF_COSTE = 32, 33

# Columnas plantilla MONOPUNTO
MP = {
    'COMERCIALIZADORA': 1, 'Recibida': 2, 'Vencimiento': 3, 'Validez': 4,
    'Renovable': 5, 'P1': 6, 'P2': 7, 'P3': 8, 'P4': 9, 'P5': 10, 'P6': 11,
    'Indice': 12, 'Ajuste': 13, 'Indexado': 14,
}
MP_HEADERS = ['COMERCIALIZADORA', 'Recibida', 'Vencimiento', 'Validez', 'Renovable',
              'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'Indice', 'Ajuste', 'Indexado']

# Columnas plantilla MULTIPUNTO
MT = {
    'Cliente': 1, 'COMERCIALIZADORA': 2, 'Recibida': 3, 'Vencimiento': 4,
    'Validez': 5, 'Renovable': 6, 'P1': 7, 'P2': 8, 'P3': 9, 'P4': 10,
    'P5': 11, 'P6': 12, 'Indice': 13, 'Ajuste': 14, 'Indexado': 15, 'Tarifa': 16,
}
MT_HEADERS = ['Cliente', 'COMERCIALIZADORA', 'Recibida', 'Vencimiento', 'Validez',
              'Renovable', 'P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'Indice', 'Ajuste',
              'Indexado', 'Tarifa']


def extract_comercializadora(filename):
    """Extrae nombre antes de ' -' del nombre del archivo."""
    name = filename.replace('.xlsx', '')
    # Captura todo antes de " -" (guion con espacio) o "PETICION"
    match = re.match(r'^(.+?)\s*-\s*(?:PETICION|Oferta|Ofertas)', name, re.IGNORECASE)
    if match:
        return match.group(1).strip().upper()
    # Fallback: todo antes de "PETICION"
    match = re.match(r'^(.+?)\s*(?:PETICION)', name, re.IGNORECASE)
    if match:
        return match.group(1).strip().upper()
    return name.upper()


def normalize(name):
    orig = name.upper().strip()
    name = re.sub(r',\s*S\.?L\.?U?\.?$', '', orig)
    name = re.sub(r',\s*S\.?L\.?$', '', name)
    name = re.sub(r'\s*\.$', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name if name else orig


def name_matches(a, b):
    tn = normalize(a)
    on = normalize(b)
    if not tn or not on:
        return False
    return tn in on or on in tn


def safe_float(v):
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def limpiar_nombre_cliente(raw_name):
    """Versión limpia del nombre del cliente para nombre de archivo."""
    name = raw_name.strip()
    name = re.sub(r',\s*S\.?L\.?U?\.?$', '', name)
    name = re.sub(r',\s*S\.?L\.?$', '', name)
    name = re.sub(r'\s*\.$', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    if not name:
        name = raw_name.strip()
    return name


def extraer_codigo_principal(codigos):
    """Toma el prefijo numerico del primer codigo. Ej: '212/001' -> '212'."""
    if not codigos:
        return '000'
    c = str(codigos[0])
    match = re.match(r'^(\d+)', c)
    return match.group(1) if match else '000'


def extraer_prefijo(codigo):
    """Extrae el prefijo numerico de un codigo. Ej: '212/001' -> '212'."""
    c = str(codigo).strip()
    match = re.match(r'^(\d+)', c)
    return match.group(1) if match else c


def detectar_formato(ws):
    """Verifica si la hoja tiene formato estandar (columnas Z-AE con precios).
    Escanea hasta 200 filas para detectar archivos con datos mas abajo.
    Acepta archivos con al menos 1 precio valido en cualquier columna (P1-P6 o FEE/Coste)."""
    max_row_check = min(ws.max_row, 200)
    found_prices = 0
    for row in ws.iter_rows(min_row=8, max_row=max_row_check, min_col=OF_P1, max_col=OF_COSTE):
        for cell in row:
            if cell.value is not None:
                try:
                    val = float(cell.value)
                    if val > 0:
                        found_prices += 1
                        if found_prices >= 1:
                            return True
                except (ValueError, TypeError):
                    pass
    return False


# ---------------------------------------------------------------------------
# LECTURA DE OFERTAS
# ---------------------------------------------------------------------------

def leer_ofertas(fnames=None, filename_to_com=None, ofertas_dir=None):
    """
    Lee los archivos en Ofertas/ que tengan formato estandar.
    Si `fnames` se especifica, solo procesa esos archivos.
    `filename_to_com`: dict {nombre_archivo: comercializadora_override}
    `ofertas_dir`: directorio donde buscar archivos de oferta (default: OFERTAS_DIR)
    Retorna (clientes_mp, clientes_mt):
      {eff_key: {
          '_nombre': str,
          '_codigos': [str],
          '_ofertas': [{'archivo', 'com', 'tarifas': {tarifa: precios}}]
      }}
    """
    if ofertas_dir is None:
        ofertas_dir = OFERTAS_DIR
    # Acumuladores globales (cruce de archivos)
    todos_clientes = {}       # prefijo -> {_nombre, _codigos: set}
    todos_precios = defaultdict(list)  # (prefijo, archivo, tarifa) -> lista de dicts

    for fname in sorted(os.listdir(ofertas_dir)):
        if not fname.endswith('.xlsx'):
            continue
        if fnames is not None and fname not in fnames:
            continue

        if hubo_cancelacion():
            log("Proceso cancelado por el usuario.\n", "warn")
            break

        if filename_to_com and fname in filename_to_com:
            comercializadora = filename_to_com[fname].strip().upper()
        else:
            comercializadora = extract_comercializadora(fname)
        filepath = os.path.join(ofertas_dir, fname)
        log(f"\nProcesando: {fname} -> Comercializadora: {comercializadora}\n")

        try:
            wb = load_workbook(filepath, data_only=True)
        except Exception as e:
            log(f"  Error al abrir: {e}\n", "error")
            continue

        if 'PETICION OFERTAS' not in wb.sheetnames:
            log("  Hoja 'PETICION OFERTAS' no encontrada, saltando\n", "warn")
            continue

        ws = wb['PETICION OFERTAS']

        if not detectar_formato(ws):
            log("  Formato no reconocido (sin precios en col Z-AE), saltando\n", "warn")
            continue

        log("  Formato detectado correctamente\n", "ok")
        # PASE 1: agrupar filas por Codigo
        rows_by_codigo = defaultdict(list)
        current_cliente = None
        max_col = ws.max_column

        for row in ws.iter_rows(min_row=8, max_col=min(max_col, OF_COSTE)):
            vals = [c.value for c in row]
            if len(vals) < OF_COSTE:
                continue
            codigo = vals[OF_CODIGO - 1]
            cliente = vals[OF_CLIENTE - 1]

            if cliente is not None:
                current_cliente = str(cliente).strip()

            if codigo is not None:
                codigo_str = str(codigo).strip()
                if codigo_str in rows_by_codigo:
                    # Mismo bloque de suministro (filas de periodo)
                    rows_by_codigo[codigo_str].append({
                        'vals': vals,
                        'cliente': current_cliente or rows_by_codigo[codigo_str][0]['cliente'],
                    })
                elif cliente is not None:
                    # Nuevo bloque con nombre de cliente en la fila
                    rows_by_codigo[codigo_str].append({
                        'vals': vals,
                        'cliente': str(cliente).strip(),
                    })
                # else: nuevo codigo sin nombre de cliente -> no atribuible, se omite

        # PASE 2: extraer precios y acumular
        codigos_por_cliente = defaultdict(set)
        for codigo, entries in rows_by_codigo.items():
            cliente = entries[0]['cliente']
            if cliente:
                codigos_por_cliente[cliente].add(codigo)

        for codigo, entries in rows_by_codigo.items():
            cliente = entries[0]['cliente']
            if not cliente:
                continue

            header_row = entries[0]['vals']
            tarifa = str(header_row[OF_TARIFA - 1]).strip() if header_row[OF_TARIFA - 1] else ''

            p1 = safe_float(header_row[OF_P1 - 1])
            p2 = safe_float(header_row[OF_P2 - 1])
            p3 = safe_float(header_row[OF_P3 - 1])
            p4 = safe_float(header_row[OF_P4 - 1])
            p5 = safe_float(header_row[OF_P5 - 1])
            p6 = safe_float(header_row[OF_P6 - 1])
            fee = safe_float(header_row[OF_FEE - 1])
            coste = safe_float(header_row[OF_COSTE - 1])

            if p1 == 0 and p2 == 0 and p3 == 0 and p4 == 0 and p5 == 0 and p6 == 0 and fee == 0 and coste == 0:
                continue

            # Acumular info del cliente - agrupar por PREFIJO del codigo
            prefijo = extraer_prefijo(codigo)

            if prefijo not in todos_clientes:
                todos_clientes[prefijo] = {
                    '_nombre': cliente,
                    '_codigos': set(),
                }
            todos_clientes[prefijo]['_codigos'].add(codigo)

            key = (prefijo, fname, tarifa)
            todos_precios[key].append({
                'P1': p1, 'P2': p2, 'P3': p3, 'P4': p4, 'P5': p5, 'P6': p6,
                'FEE': fee, 'Coste': coste,
                'codigo': codigo,
                'com': comercializadora,
            })

        n_codigos = len(rows_by_codigo)
        n_clientes = len(codigos_por_cliente)
        log(f"  {n_codigos} suministros, {n_clientes} clientes\n")

    # Clasificar: MONOPUNTO = 1 codigo, MULTIPUNTO = >1 codigos
    clientes_mp = {}
    clientes_mt = {}

    for cliente_norm, info in todos_clientes.items():
        codigos = list(info['_codigos'])
        cant = len(codigos)
        target = clientes_mp if cant == 1 else clientes_mt

        # Cada archivo conserva su propia oferta aunque comparta comercializadora.
        ofertas_por_archivo = {}
        for (cn, archivo, tarifa), entries in todos_precios.items():
            if cn != cliente_norm:
                continue
            oferta = ofertas_por_archivo.setdefault(archivo, {
                'archivo': archivo,
                'com': entries[0]['com'],
                'tarifas': {},
            })
            entry = entries[0]
            oferta['tarifas'][tarifa] = {
                'P1': entry['P1'], 'P2': entry['P2'], 'P3': entry['P3'],
                'P4': entry['P4'], 'P5': entry['P5'], 'P6': entry['P6'],
                'FEE': entry['FEE'], 'Coste': entry['Coste'],
            }

        target[cliente_norm] = {
            '_nombre': info['_nombre'],
            '_codigos': codigos,
            '_ofertas': sorted(
                ofertas_por_archivo.values(),
                key=lambda oferta: (oferta['com'].upper(), oferta['archivo'].lower())),
        }

    return clientes_mp, clientes_mt


# ---------------------------------------------------------------------------
# CREACION DE PLANTILLAS
# ---------------------------------------------------------------------------

def auto_ajustar_columnas(ws):
    """Ajusta el ancho de las columnas al contenido."""
    from openpyxl.utils import get_column_letter
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            val = str(cell.value) if cell.value is not None else ''
            length = sum(2 if ord(c) > 127 else 1 for c in val)
            if length > max_len:
                max_len = length
        adjusted = min(max_len + 2, 45)
        ws.column_dimensions[col_letter].width = max(adjusted, 8)


def crear_libro_plantilla(headers):
    """Crea un Workbook nuevo con hoja 'Plantilla' y fila de encabezados."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Plantilla'
    for col_idx, h in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=h)
    return wb, ws


def _datos_para_com(com, fechas_com, archivo=None):
    """Returns offer dates, preserving compatibility with maps keyed by retailer.

    None signals: don't touch columns (preserve existing).
    """
    if fechas_com is None:
        return date.today(), None, None, 'No'
    f = fechas_com.get(archivo) if archivo else None
    if f is None:
        f = fechas_com.get(com.upper())
    if f is None:
        return None
    return (f.get('Recibida', date.today()),
            f.get('Vencimiento'),
            f.get('Validez'),
            f.get('Renovable', 'No'))


def _escribir_datos_com(ws, row_idx, cols, com, fechas_com, archivo=None):
    datos = _datos_para_com(com, fechas_com, archivo=archivo)
    if datos is None:
        return
    recibida, vencimiento, validez, renovable = datos
    # Una fecha vacia significa "no modificar": deja la celda como esta.
    for clave, valor in (('Recibida', recibida), ('Vencimiento', vencimiento), ('Validez', validez)):
        if valor is None:
            continue
        celda = ws.cell(row=row_idx, column=cols[clave])
        celda.value = valor
        celda.number_format = 'DD/MM/YYYY'
    ws.cell(row=row_idx, column=cols['Renovable']).value = renovable


def _valor_cliente(cliente):
    if cliente is None:
        return None
    try:
        return int(cliente)
    except (ValueError, TypeError):
        return cliente


def _tiene_fijos(datos):
    return any(datos.get(k, 0) for k in ('P1', 'P2', 'P3', 'P4', 'P5', 'P6'))


def _tiene_indexados(datos):
    return datos.get('FEE', 0) + datos.get('Coste', 0) > 0


def escribir_fila_fija(ws, row_idx, cols, datos, cliente=None, tarifa=None,
                       fechas_com=None):
    """Escribe una fila de precio fijo (Indexado=No)."""
    c = _valor_cliente(cliente)
    if c is not None:
        ws.cell(row=row_idx, column=cols['Cliente']).value = c
    ws.cell(row=row_idx, column=cols['COMERCIALIZADORA']).value = datos['com']
    _escribir_datos_com(ws, row_idx, cols, datos['com'], fechas_com,
                        archivo=datos.get('archivo'))
    ws.cell(row=row_idx, column=cols['P1']).value = datos['P1']
    ws.cell(row=row_idx, column=cols['P2']).value = datos['P2']
    ws.cell(row=row_idx, column=cols['P3']).value = datos['P3']
    ws.cell(row=row_idx, column=cols['P4']).value = datos['P4']
    ws.cell(row=row_idx, column=cols['P5']).value = datos['P5']
    ws.cell(row=row_idx, column=cols['P6']).value = datos['P6']
    ws.cell(row=row_idx, column=cols['Indice']).value = 0
    ws.cell(row=row_idx, column=cols['Ajuste']).value = 0
    ws.cell(row=row_idx, column=cols['Indexado']).value = 'No'
    if tarifa is not None and 'Tarifa' in cols:
        ws.cell(row=row_idx, column=cols['Tarifa']).value = tarifa


def escribir_fila_indexada(ws, row_idx, cols, datos, cliente=None, tarifa=None,
                           fechas_com=None):
    """Escribe una fila indexada (Indexado=Si)."""
    indice_val = round(datos['FEE'] + datos['Coste'], 4)
    c = _valor_cliente(cliente)
    if c is not None:
        ws.cell(row=row_idx, column=cols['Cliente']).value = c
    ws.cell(row=row_idx, column=cols['COMERCIALIZADORA']).value = datos['com']
    _escribir_datos_com(ws, row_idx, cols, datos['com'], fechas_com,
                        archivo=datos.get('archivo'))
    ws.cell(row=row_idx, column=cols['P1']).value = 0
    ws.cell(row=row_idx, column=cols['P2']).value = 0
    ws.cell(row=row_idx, column=cols['P3']).value = 0
    ws.cell(row=row_idx, column=cols['P4']).value = 0
    ws.cell(row=row_idx, column=cols['P5']).value = 0
    ws.cell(row=row_idx, column=cols['P6']).value = 0
    ws.cell(row=row_idx, column=cols['Indice']).value = indice_val
    ws.cell(row=row_idx, column=cols['Ajuste']).value = 1.5
    ws.cell(row=row_idx, column=cols['Indexado']).value = 'Si'
    if tarifa is not None and 'Tarifa' in cols:
        ws.cell(row=row_idx, column=cols['Tarifa']).value = tarifa


def crear_plantilla_monopunto(cliente_info, fechas_com=None):
    """Crea un Workbook nuevo para un cliente MONOPUNTO."""
    ofertas = cliente_info['_ofertas']

    wb, ws = crear_libro_plantilla(MP_HEADERS)
    row_idx = 2

    for oferta in ofertas:
        com = oferta['com']
        tarifas = oferta['tarifas']
        tarifa = list(tarifas.keys())[0] if tarifas else ''
        if not tarifa:
            continue
        datos = tarifas[tarifa]

        datos_con_com = dict(datos, com=com, archivo=oferta['archivo'])
        if _tiene_fijos(datos):
            escribir_fila_fija(ws, row_idx, MP, datos_con_com, fechas_com=fechas_com)
            row_idx += 1
        if _tiene_indexados(datos):
            escribir_fila_indexada(ws, row_idx, MP, datos_con_com, fechas_com=fechas_com)
            row_idx += 1

    return wb


def crear_plantilla_multipunto(cliente_info, fechas_com=None):
    """Crea un Workbook nuevo para un cliente MULTIPUNTO."""
    ofertas = cliente_info['_ofertas']

    codigo_principal = extraer_codigo_principal(cliente_info['_codigos'])

    wb, ws = crear_libro_plantilla(MT_HEADERS)
    row_idx = 2

    for oferta in ofertas:
        com = oferta['com']
        tarifas = oferta['tarifas']
        for tarifa in sorted(tarifas.keys()):
            datos = tarifas[tarifa]
            datos_con_com = dict(datos, com=com, archivo=oferta['archivo'])

            if _tiene_fijos(datos):
                escribir_fila_fija(ws, row_idx, MT, datos_con_com,
                                   cliente=codigo_principal, tarifa=tarifa,
                                   fechas_com=fechas_com)
                row_idx += 1
            if _tiene_indexados(datos):
                escribir_fila_indexada(ws, row_idx, MT, datos_con_com,
                                       cliente=codigo_principal, tarifa=tarifa,
                                       fechas_com=fechas_com)
                row_idx += 1

    return wb


def obtener_nombre_archivo_mp(cliente_info):
    """Genera el nombre de archivo para un cliente MONOPUNTO."""
    nombre_limpio = limpiar_nombre_cliente(cliente_info['_nombre'])
    codigo = extraer_codigo_principal(cliente_info['_codigos'])
    return f"{codigo} Datos Monopunto {nombre_limpio}.xlsx"


def obtener_nombre_archivo_mt(cliente_info):
    """Genera el nombre de archivo para un cliente MULTIPUNTO."""
    nombre_limpio = limpiar_nombre_cliente(cliente_info['_nombre'])
    codigo = extraer_codigo_principal(cliente_info['_codigos'])
    return f"{codigo} Datos Multipunto {nombre_limpio}.xlsx"


# ---------------------------------------------------------------------------
# ACTUALIZACION DE PLANTILLAS
# ---------------------------------------------------------------------------

def actualizar_o_crear_plantilla(cliente_info, directorio, es_monopunto, fechas_com=None):
    """
    Busca una plantilla existente para el cliente.
    Si existe: actualiza filas y agrega las que falten.
    Si no existe: la crea.
    fechas_com: {archivo: {'Recibida': date, 'Vencimiento': date, 'Validez': date}}
    """
    cliente_nombre = cliente_info['_nombre']
    ofertas = cliente_info['_ofertas']
    cols = MP if es_monopunto else MT
    headers = MP_HEADERS if es_monopunto else MT_HEADERS

    # Buscar archivo existente que coincida con el cliente
    archivo_existente = None
    for fname in os.listdir(directorio):
        if not fname.endswith('.xlsx'):
            continue

        # Extraer nombre de cliente del archivo
        basename = fname.replace('.xlsx', '')
        code_match = re.match(r'^(\d+)\s+', basename)
        file_code = code_match.group(1) if code_match else None
        basename = re.sub(r'^\d+\s+', '', basename)
        if es_monopunto:
            basename = re.sub(r'^Datos\s+Monopunto\s+', '', basename)
        else:
            basename = re.sub(r'^Datos\s+Multipunto\s+', '', basename)
        nombre_archivo = basename.strip()

        if name_matches(nombre_archivo, cliente_nombre):
            if file_code:
                our_codes = [str(c) for c in cliente_info['_codigos']]
                code_ok = any(
                    str(c).startswith(file_code) or file_code.startswith(str(c).split('/')[0])
                    for c in our_codes
                )
                if not code_ok:
                    continue
            archivo_existente = os.path.join(directorio, fname)
            break

    if es_monopunto:
        nom_archivo = obtener_nombre_archivo_mp(cliente_info)
    else:
        nom_archivo = obtener_nombre_archivo_mt(cliente_info)

    try:
        if archivo_existente:
            log(f"    Plantilla existente: {os.path.basename(archivo_existente)}\n")
            wb = load_workbook(archivo_existente)
            ws = wb['Plantilla']
            _actualizar_plantilla(ws, cols, ofertas, cliente_info, es_monopunto, fechas_com)
            output_path = archivo_existente
        else:
            output_path = os.path.join(directorio, nom_archivo)
            log(f"    Creando nueva plantilla: {nom_archivo}\n")
            if es_monopunto:
                wb = crear_plantilla_monopunto(cliente_info, fechas_com)
            else:
                wb = crear_plantilla_multipunto(cliente_info, fechas_com)
            ws = wb['Plantilla']

        auto_ajustar_columnas(ws)
        wb.save(output_path)
        log(f"    Guardado: {os.path.basename(output_path)}\n", "ok")
    except PermissionError:
        log(f"    ERROR: No se pudo guardar '{os.path.basename(output_path)}' - archivo abierto en Excel?\n", "error")
    except Exception as e:
        log(f"    ERROR al guardar: {e}\n", "error")


def _actualizar_plantilla(ws, cols, ofertas, cliente_info, es_monopunto, fechas_com=None):
    """
    Actualiza filas existentes y agrega las que falten.
    Las filas repetidas se emparejan en orden para conservar una por archivo.
    """
    codigo_principal = extraer_codigo_principal(cliente_info['_codigos'])

    # Leer filas existentes
    filas_existentes = defaultdict(list)  # (com, tarifa, indexado, renovable) -> row_idx
    max_row = ws.max_row

    for row_idx in range(2, max_row + 1):
        com_val = ws.cell(row=row_idx, column=cols['COMERCIALIZADORA']).value
        if not com_val:
            continue
        com_val = str(com_val).strip().upper()

        indexado_val = ws.cell(row=row_idx, column=cols['Indexado']).value
        indexado = str(indexado_val).strip().upper() if indexado_val else ''

        renovable_val = ws.cell(row=row_idx, column=cols['Renovable']).value
        renovable = str(renovable_val).strip().upper() if renovable_val else 'NO'

        tarifa_val = ''
        if 'Tarifa' in cols:
            tv = ws.cell(row=row_idx, column=cols['Tarifa']).value
            tarifa_val = str(tv).strip() if tv else ''

        filas_existentes[(com_val, tarifa_val, indexado, renovable)].append(row_idx)

    def renovable_para_oferta(oferta):
        if fechas_com is None:
            return 'NO'
        f = fechas_com.get(oferta['archivo'])
        if f is None:
            f = fechas_com.get(oferta['com'].upper())
        if f is None:
            return 'NO'
        return f.get('Renovable', 'No').strip().upper()

    # Actualizar cada fila una sola vez y agregar las ofertas que falten.
    combinaciones_a_agregar = []  # (oferta, tarifa, indexado)

    es_mt = 'Tarifa' in cols
    for oferta in ofertas:
        com = oferta['com']
        tarifas = oferta['tarifas']
        ren = renovable_para_oferta(oferta)
        for tarifa in sorted(tarifas.keys()):
            datos = tarifas[tarifa]
            datos_con_com = dict(datos, com=com, archivo=oferta['archivo'])
            tarifa_key = tarifa if es_mt else ''

            key_no = (com, tarifa_key, 'NO', ren)
            if filas_existentes[key_no]:
                r = filas_existentes[key_no].pop(0)
                escribir_fila_fija(ws, r, cols, datos_con_com,
                                   cliente=codigo_principal if not es_monopunto else None,
                                   tarifa=tarifa if not es_monopunto else None,
                                   fechas_com=fechas_com)
            elif _tiene_fijos(datos):
                combinaciones_a_agregar.append((oferta, tarifa, 'NO'))

            key_si = (com, tarifa_key, 'SI', ren)
            if filas_existentes[key_si]:
                r = filas_existentes[key_si].pop(0)
                escribir_fila_indexada(ws, r, cols, datos_con_com,
                                       cliente=codigo_principal if not es_monopunto else None,
                                       tarifa=tarifa if not es_monopunto else None,
                                       fechas_com=fechas_com)
            elif _tiene_indexados(datos):
                combinaciones_a_agregar.append((oferta, tarifa, 'SI'))

    # Agregar combinaciones faltantes al final
    next_row = ws.max_row + 1
    for oferta, tarifa, tipo in combinaciones_a_agregar:
        datos = oferta['tarifas'][tarifa]
        datos_con_com = dict(datos, com=oferta['com'], archivo=oferta['archivo'])

        if tipo == 'NO':
            escribir_fila_fija(ws, next_row, cols, datos_con_com,
                               cliente=codigo_principal if not es_monopunto else None,
                               tarifa=tarifa if not es_monopunto else None,
                               fechas_com=fechas_com)
        else:
            escribir_fila_indexada(ws, next_row, cols, datos_con_com,
                                   cliente=codigo_principal if not es_monopunto else None,
                                   tarifa=tarifa if not es_monopunto else None,
                                   fechas_com=fechas_com)
        next_row += 1


def procesar_clientes(clientes_dict, directorio, es_monopunto, tipo_label, fechas_com=None):
    """Procesa todos los clientes en clientes_dict.
    fechas_com: {archivo: {'Recibida': date, 'Vencimiento': date, 'Validez': date}}
    """
    log(f"\n[{tipo_label}] Procesando {len(clientes_dict)} clientes...\n")
    for cliente_norm, info in sorted(clientes_dict.items()):
        if hubo_cancelacion():
            log("  Proceso cancelado por el usuario.\n", "warn")
            return
        nom = info['_nombre']
        coms = [oferta['com'] for oferta in info['_ofertas']]
        log(f"\n  Cliente: {nom}\n")
        log(f"    Comercializadoras: {coms}\n")
        actualizar_o_crear_plantilla(info, directorio, es_monopunto, fechas_com)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("Llenador de plantillas de ofertas electricas v2")
    print("=" * 60)

    # 1. Leer ofertas
    print("\n[1/3] Leyendo ofertas...")
    ofertas_mp, ofertas_mt = leer_ofertas()

    print(f"\n--- MONOPUNTO en ofertas: {len(ofertas_mp)} ---")
    for k, v in sorted(ofertas_mp.items()):
        coms = [oferta['com'] for oferta in v['_ofertas']]
        print(f"  {v['_nombre']} (codigos: {v['_codigos']}) -> {coms}")

    print(f"\n--- MULTIPUNTO en ofertas: {len(ofertas_mt)} ---")
    for k, v in sorted(ofertas_mt.items()):
        coms = [oferta['com'] for oferta in v['_ofertas']]
        print(f"  {v['_nombre']} (codigos: {v['_codigos']}) -> {coms}")

    # 2. Procesar MONOPUNTO
    print("\n[2/3] Procesando plantillas MONOPUNTO...")
    procesar_clientes(ofertas_mp, MONOPUNTO_DIR, es_monopunto=True, tipo_label="MONOPUNTO")

    # 3. Procesar MULTIPUNTO
    print("\n[3/3] Procesando plantillas MULTIPUNTO...")
    procesar_clientes(ofertas_mt, MULTIPUNTO_DIR, es_monopunto=False, tipo_label="MULTIPUNTO")

    if hubo_cancelacion():
        log("\nProceso cancelado por el usuario.\n", "warn")
    else:
        print("\n" + "=" * 60)
        print("Proceso completado!")
        print("=" * 60)
        log("Proceso completado!\n", "ok")


if __name__ == '__main__':
    main()

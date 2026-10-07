#!/usr/bin/env python3
"""
Genera el manual de usuario en formato Word (.docx)
para la herramienta de Actualización de Precios Eléctricos.
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os
from datetime import date


def set_cell_shading(cell, color_hex):
    shading = cell._element.get_or_add_tcPr()
    shading_elem = shading.makeelement(qn('w:shd'), {
        qn('w:fill'): color_hex,
        qn('w:val'): 'clear',
    })
    shading.append(shading_elem)


def create_manual():
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(6)

    # =========================================================
    # PORTADA
    # =========================================================
    for _ in range(6):
        doc.add_paragraph()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('Manual de Usuario')
    run.font.size = Pt(28)
    run.bold = True
    run.font.color.rgb = RGBColor(0, 95, 184)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('Herramienta de Actualización de Precios Eléctricos')
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph()

    version = doc.add_paragraph()
    version.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = version.add_run(f'Versión 1.0 — {date.today().strftime("%B %Y")}')
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(120, 120, 120)

    doc.add_paragraph()

    exe_info = doc.add_paragraph()
    exe_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = exe_info.add_run('Ejecutable: ActualizarPrecios.exe')
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_page_break()

    # =========================================================
    # TABLA DE CONTENIDOS
    # =========================================================
    doc.add_heading('Tabla de Contenidos', level=1)
    toc_items = [
        ('1.', 'Introducción'),
        ('2.', 'Estructura de Carpetas y Archivos'),
        ('3.', 'Guía de Uso Paso a Paso'),
        ('4.', 'Formato de Archivo de Oferta (Entrada)'),
        ('5.', 'Formato de Plantillas de Salida'),
        ('6.', 'Clasificación: Monopunto y Multipunto'),
        ('7.', 'Listado de Comercializadoras'),
        ('8.', 'Solución de Problemas'),
    ]
    for num, item in toc_items:
        p = doc.add_paragraph()
        run = p.add_run(f'{num} {item}')
        run.font.size = Pt(12)

    doc.add_page_break()

    # =========================================================
    # 1. INTRODUCCIÓN
    # =========================================================
    doc.add_heading('1. Introducción', level=1)

    doc.add_paragraph(
        'La herramienta de Actualización de Precios Eléctricos permite procesar '
        'ofertas de energía eléctrica recibidas de distintas comercializadoras, '
        'clasificar los puntos de suministro de los clientes y generar plantillas '
        'estandarizadas con los precios actualizados.'
    )

    doc.add_paragraph('El proceso general consta de los siguientes pasos:')

    steps = [
        'El usuario ejecuta el archivo ActualizarPrecios.exe.',
        'Se muestran los archivos de oferta (.xlsx) disponibles en la carpeta configurada.',
        'El usuario selecciona las ofertas a procesar, asigna la comercializadora y las fechas correspondientes.',
        'Al hacer clic en "Procesar", la herramienta lee las ofertas, clasifica a los clientes y genera las plantillas.',
        'Los resultados se guardan en las carpetas MONOPUNTO/ y MULTIPUNTO/ dentro del directorio de salida.',
    ]
    for i, step in enumerate(steps, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. ')
        run.bold = True
        p.add_run(step)

    doc.add_page_break()

    # =========================================================
    # 2. ESTRUCTURA DE CARPETAS
    # =========================================================
    doc.add_heading('2. Estructura de Carpetas y Archivos', level=1)

    doc.add_paragraph('La herramienta utiliza la siguiente estructura de directorios:')

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table.rows[0].cells
    hdr[0].text = 'Carpeta / Archivo'
    hdr[1].text = 'Tipo'
    hdr[2].text = 'Descripción'
    for cell in hdr:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    folders = [
        ('Ofertas/', 'Entrada', 'Carpeta donde se colocan los archivos de oferta (.xlsx) recibidos de las comercializadoras.'),
        ('MONOPUNTO/', 'Salida', 'Plantillas generadas para clientes con un único punto de suministro.'),
        ('MULTIPUNTO/', 'Salida', 'Plantillas generadas para clientes con múltiples puntos de suministro.'),
        ('BACKUP/', 'Respaldo', 'Copia de seguridad de plantillas existentes.'),
        ('Comercializadoras.txt', 'Configuración', 'Lista de comercializadoras conocidas (una por línea). Se usa para auto-sugerir la comercializadora al cargar ofertas.'),
        ('ActualizarPrecios.exe', 'Ejecutable', 'Archivo ejecutable de la herramienta. No requiere instalación de Python.'),
    ]

    for name, tipo, desc in folders:
        row = table.add_row()
        row.cells[0].text = name
        row.cells[1].text = tipo
        row.cells[2].text = desc
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run('Nota: ')
    run.bold = True
    p.add_run(
        'Las carpetas MONOPUNTO/ y MULTIPUNTO/ se crean automáticamente '
        'al procesar ofertas si no existen. El directorio de salida por defecto '
        'es la misma carpeta donde se encuentra el ejecutable.'
    )

    doc.add_page_break()

    # =========================================================
    # 3. GUÍA DE USO PASO A PASO
    # =========================================================
    doc.add_heading('3. Guía de Uso Paso a Paso', level=1)

    # 3.1
    doc.add_heading('3.1 Iniciar la Herramienta', level=2)
    doc.add_paragraph(
        'Ejecutar el archivo ActualizarPrecios.exe. Se abrirá una ventana con la interfaz '
        'gráfica titulada "Gestor de Ofertas Eléctricas", organizada en cuatro bloques: '
        'CARPETAS, OFERTAS, RESULTADO y una barra inferior de acción.'
    )
    doc.add_paragraph(
        'En la esquina superior derecha está el interruptor de tema (Claro / Oscuro). '
        'La preferencia se recuerda junto con las carpetas y el tamaño de la ventana '
        'en el archivo config.json de la carpeta de la aplicación.'
    )

    # 3.2
    doc.add_heading('3.2 Seleccionar Directorios', level=2)
    doc.add_paragraph(
        'En el bloque CARPETAS se encuentran los selectores de directorio:'
    )

    dir_items = [
        ('Ofertas (origen): ', 'Carpeta donde se encuentran los archivos de oferta .xlsx. '
         'Por defecto apunta a la carpeta Ofertas/ junto al ejecutable. '
         'Se puede cambiar haciendo clic en el botón de carpeta.'),
        ('Salida (resultados): ', 'Carpeta donde se guardarán las plantillas generadas. '
         'Por defecto es la carpeta del ejecutable. Las subcarpetas MONOPUNTO/ y MULTIPUNTO/ '
         'se crean automáticamente dentro de esta ruta.'),
    ]
    for label, desc in dir_items:
        p = doc.add_paragraph()
        run = p.add_run(label)
        run.bold = True
        p.add_run(desc)

    doc.add_paragraph(
        'Junto a los selectores, los botones "Abrir MONOPUNTO" y "Abrir MULTIPUNTO" '
        'abren directamente esas carpetas de resultados en el explorador de archivos. '
        'Si aún no existen, se crean vacías.'
    )

    # 3.3
    doc.add_heading('3.3 Revisar la Tabla de Ofertas', level=2)
    doc.add_paragraph(
        'Al seleccionar el directorio de ofertas, la herramienta lista automáticamente '
        'todos los archivos .xlsx encontrados (ignora archivos temporales que empiezan con ~$). '
        'La tabla tiene UNA FILA POR ARCHIVO de oferta: no se agrupan los archivos de la '
        'misma comercializadora.'
    )
    doc.add_paragraph(
        'Sobre la tabla hay un cuadro de búsqueda. Escribir texto filtra al instante por '
        'nombre de archivo o de comercializadora, y al limpiarlo vuelven a mostrarse todos '
        'los archivos. Haciendo clic en la cabecera de una columna se alterna el orden '
        'ascendente/descendente.'
    )
    table2 = doc.add_table(rows=1, cols=3)
    table2.style = 'Light Grid Accent 1'
    hdr2 = table2.rows[0].cells
    hdr2[0].text = 'Columna'
    hdr2[1].text = 'Descripción'
    hdr2[2].text = 'Valor por defecto'
    for cell in hdr2:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    cols = [
        ('Casilla', 'Seleccionar/deseleccionar la oferta para procesar (también se '
                    'marca o desmarca haciendo clic en la celda izquierda).', 'Marcada (activa)'),
        ('Archivo', 'Nombre del archivo de oferta.', '—'),
        ('Comercializadora', 'Comercializadora del archivo. Se auto-detecta según el '
                            'nombre y se puede cambiar con el desplegable. Puede repetirse '
                            'en varios archivos.', 'Auto-detectada'),
        ('Renovable', 'Si la oferta es renovable o no.', 'No'),
        ('Recibida', 'Fecha en que se recibió la oferta.', 'Fecha de hoy'),
        ('Vencimiento', 'Fecha de vencimiento de la oferta.', 'Fin del mismo mes del '
                        'año siguiente (02/10/2026 → 31/10/2027)'),
        ('Validez', 'Fecha de validez de la oferta.', 'Viernes de esta semana '
                    '(07/10/2026 → 09/10/2026)'),
    ]
    for col_name, desc, default in cols:
        row = table2.add_row()
        row.cells[0].text = col_name
        row.cells[1].text = desc
        row.cells[2].text = default
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()
    doc.add_paragraph('Botones disponibles sobre la tabla:')
    toolbar_items = [
        'Seleccionar todas: Marca todos los checkboxes.',
        'Deseleccionar todas: Desmarca todos los checkboxes.',
        'Refrescar: Vuelve a leer la carpeta de ofertas y actualiza la tabla sin perder '
        'las fechas ni las comercializadoras ya asignadas.',
    ]
    for item in toolbar_items:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)

    # 3.4
    doc.add_heading('3.4 Editar Celdas, Fechas y Comercializadora', level=2)
    doc.add_paragraph(
        'La edición se hace directamente sobre la tabla, haciendo un clic simple en la celda '
        'correspondiente (doble clic o Intro sobre la fila seleccionada también abre el editor):'
    )
    items_34 = [
        'Casilla (primera columna): alterna entre marcada y desmarcada.',
        'Comercializadora: abre un desplegable con la lista de comercializadoras conocidas. '
        'Cada archivo tiene la suya propia; no hay grupos.',
        'Renovable: abre un desplegable con los valores "No" y "Si".',
        'Fechas Recibida, Vencimiento y Validez: abren una caja donde se puede escribir o '
        'pegar la fecha (p. ej. 01/04/2027, 2027-05-31, 30.06.27). Un botón 📅 junto a la '
        'caja abre el calendario con los botones Aceptar, Limpiar y Cancelar, además de la '
        'navegación entre meses.',
    ]
    for item in items_34:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)

    doc.add_paragraph(
        'En el editor de fecha: Intro acepta lo escrito, Tab acepta y baja a la fila siguiente '
        '(Shift+Tab a la anterior) y Esc deja la celda como estaba.'
    )
    doc.add_paragraph(
        'Copiar y pegar: Ctrl+C copia la fila enfocada al portapapeles y Ctrl+V la pega en la '
        'fila enfocada. Si el portapapeles trae las tres fechas separadas por tabulador '
        '(por ejemplo, copiadas de Excel), se rellenan Recibida, Vencimiento y Validez a la vez. '
        'Si solo hay una fecha, se aplica a la columna de fecha enfocada.'
    )

    p_note = doc.add_paragraph()
    run = p_note.add_run('Importante: ')
    run.bold = True
    run.font.color.rgb = RGBColor(180, 0, 0)
    p_note.add_run(
        'Cada archivo de oferta se procesa de forma independiente, incluso si comparte '
        'comercializadora con otro. En el Excel pueden aparecer varias filas con la misma '
        'comercializadora; cada una conserva los precios y fechas de su archivo de origen.'
    )

    p_note = doc.add_paragraph()
    run = p_note.add_run('Importante: ')
    run.bold = True
    run.font.color.rgb = RGBColor(180, 0, 0)
    p_note.add_run(
        'Una fecha vacía significa "no modificar": la herramienta deja intacta la celda de '
        'esa fecha en las plantillas ya existentes. Si se quiere borrar la fecha de una '
        'plantilla, hay que borrarla en Excel o dejar de usar la plantilla anterior.'
    )

    # 3.5
    doc.add_heading('3.5 Procesar las Ofertas', level=2)
    doc.add_paragraph(
        'Una vez configuradas todas las ofertas, hacer clic en el botón azul '
        '"Procesar" de la barra inferior. Mientras dura el proceso, ese mismo botón '
        'pasa a ser "Cancelar" y la barra de progreso muestra el avance.'
    )
    doc.add_paragraph(
        'El proceso se ejecuta en segundo plano, por lo que la ventana sigue respondiendo '
        'y se puede cancelar en cualquier momento. Todo el detalle se muestra en el bloque '
        'RESULTADO, con un código de color según el tipo de mensaje:'
    )
    steps_35 = [
        'Lectura y validación de cada archivo de oferta seleccionado.',
        'Extracción de precios por cliente y punto de suministro.',
        'Clasificación de clientes en Monopunto o Multipunto.',
        'Creación o actualización de plantillas en las carpetas de salida.',
    ]
    for i, step in enumerate(steps_35, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. ')
        run.bold = True
        p.add_run(step)

    doc.add_paragraph(
        'Al finalizar, el log muestra un resumen con el número de ofertas procesadas, '
        'clientes Monopunto y clientes Multipunto generados. Para vaciar el log, '
        'usar el botón "Limpiar log".'
    )

    # 3.6
    doc.add_heading('3.6 Consultar los Resultados', level=2)
    doc.add_paragraph('Las plantillas generadas se guardan en:')
    result_items = [
        'MONOPUNTO/: Archivos con formato "{Código} Datos Monopunto {Nombre Cliente}.xlsx"',
        'MULTIPUNTO/: Archivos con formato "{Código} Datos Multipunto {Nombre Cliente}.xlsx"',
    ]
    for item in result_items:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(item)
        run.font.name = 'Consolas'
        run.font.size = Pt(9)

    doc.add_paragraph(
        'Cada plantilla contiene una hoja llamada "Plantilla" con los precios organizados '
        'por comercializadora, incluyendo información de fechas, precios por período (P1-P6), '
        'y si es precio fijo o indexado. Para abrirlas, usar los botones "Abrir MONOPUNTO" '
        'y "Abrir MULTIPUNTO" del bloque CARPETAS.'
    )

    # 3.7
    doc.add_heading('3.7 Atajos de Teclado', level=2)
    shortcuts = [
        ('Ctrl + O', 'Elegir la carpeta de ofertas.'),
        ('Ctrl + S', 'Elegir la carpeta de salida.'),
        ('Ctrl + R', 'Refrescar la lista de ofertas.'),
        ('Ctrl + A', 'Seleccionar todas las ofertas.'),
        ('Ctrl + D', 'Deseleccionar todas las ofertas.'),
        ('Ctrl + C', 'Copiar la fila enfocada al portapapeles.'),
        ('Ctrl + V', 'Pegar en la fila enfocada (fechas, comercializadora y renovable).'),
        ('Ctrl + L', 'Limpiar el log.'),
        ('Intro', 'Procesar; sobre una celda abre su editor; en el editor de fecha, confirmar.'),
        ('Tab', 'En el editor de fecha, confirmar y saltar a la fila siguiente (Shift+Tab, '
                'a la anterior).'),
        ('Doble clic', 'Abrir el editor de la celda; sobre la casilla, alternarla.'),
        ('Flechas', 'Recorrer filas y celdas de la tabla.'),
        ('Espacio', 'Alternar la casilla de la fila enfocada.'),
        ('Esc', 'Cerrar el editor o el calendario sin guardar; cancelar el proceso en curso.'),
        ('F1', 'Ver esta ayuda.'),
    ]
    table3 = doc.add_table(rows=1, cols=2)
    table3.style = 'Light Grid Accent 1'
    hdr3 = table3.rows[0].cells
    hdr3[0].text = 'Atajo'
    hdr3[1].text = 'Acción'
    for cell in hdr3:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)
    for key, action in shortcuts:
        row = table3.add_row()
        run = row.cells[0].paragraphs[0].add_run(key)
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        run2 = row.cells[1].paragraphs[0].add_run(action)
        run2.font.size = Pt(9)

    doc.add_page_break()

    # =========================================================
    # 4. FORMATO DE ARCHIVO DE OFERTA
    # =========================================================
    doc.add_heading('4. Formato de Archivo de Oferta (Entrada)', level=1)

    doc.add_paragraph(
        'Los archivos de oferta deben ser archivos Excel (.xlsx) con un formato específico. '
        'La herramienta espera encontrar una hoja llamada "PETICION OFERTAS" en cada archivo.'
    )

    doc.add_heading('4.1 Nombre del Archivo', level=2)
    doc.add_paragraph(
        'El nombre del archivo se utiliza para extraer el nombre de la comercializadora. '
        'Se recomienda el siguiente formato:'
    )
    p = doc.add_paragraph()
    run = p.add_run('{COMERCIALIZADORA} - PETICION {PRODUCTO} {PERIODO}.xlsx')
    run.font.name = 'Consolas'
    run.font.size = Pt(10)

    doc.add_paragraph('Ejemplos:')
    examples = [
        'ELEIA - PETICION ELECTRICIDAD JUNIO PRECIOS.xlsx',
        'Iberdrola - PETICION ELECTRICIDAD JUNIO1.xlsx',
    ]
    for ex in examples:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(ex)
        run.font.name = 'Consolas'
        run.font.size = Pt(9)

    doc.add_paragraph(
        'La herramienta extrae el nombre de la comercializadora tomando el texto '
        'antes de " - PETICION" (ignorando mayúsculas/minúsculas). '
        'Si el nombre no coincide con la lista predefinida, se puede seleccionar '
        'manualmente en el desplegable.'
    )

    doc.add_heading('4.2 Estructura de la Hoja "PETICION OFERTAS"', level=2)
    doc.add_paragraph(
        'Los datos comienzan en la fila 8. Las filas 1 a 7 contienen encabezados y metadatos. '
        'Las columnas relevantes son:'
    )

    table3 = doc.add_table(rows=1, cols=4)
    table3.style = 'Light Grid Accent 1'
    hdr3 = table3.rows[0].cells
    hdr3[0].text = 'Columna Excel'
    hdr3[1].text = 'Letra'
    hdr3[2].text = 'Contenido'
    hdr3[3].text = 'Ejemplo'
    for cell in hdr3:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    input_cols = [
        ('2', 'B', 'Nombre del cliente', 'EMPRESA EJEMPLO S.L.'),
        ('4', 'D', 'Código del punto de suministro', '212/001'),
        ('9', 'I', 'Identificador CUPS', 'ES00000000000000000000'),
        ('10', 'J', 'Tarifa', '2.0TD'),
        ('26', 'Z', 'Precio período 1 (P1)', '0,1234'),
        ('27', 'AA', 'Precio período 2 (P2)', '0,1098'),
        ('28', 'AB', 'Precio período 3 (P3)', '0,0987'),
        ('29', 'AC', 'Precio período 4 (P4)', '0,0876'),
        ('30', 'AD', 'Precio período 5 (P5)', '0,0765'),
        ('31', 'AE', 'Precio período 6 (P6)', '0,0654'),
        ('32', 'AF', 'FEE (comisión de indexación)', '0,0045'),
        ('33', 'AG', 'Coste de desvíos', '0,0012'),
    ]
    for col_num, letter, content, example in input_cols:
        row = table3.add_row()
        row.cells[0].text = col_num
        row.cells[1].text = letter
        row.cells[2].text = content
        row.cells[3].text = example
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

    doc.add_heading('4.3 Formato de Filas', level=2)
    doc.add_paragraph(
        'Cada bloque de datos comienza con una fila que contiene el nombre del cliente '
        'en la columna B y el código del punto de suministro en la columna D. '
        'Si la columna B está vacía, la fila pertenece al bloque anterior '
        '(se usa para períodos de precios adicionales).'
    )

    p = doc.add_paragraph()
    run = p.add_run('Nota: ')
    run.bold = True
    p.add_run(
        'La herramienta valida que el archivo contenga al menos un precio válido '
        '(valor numérico mayor a 0) en las columnas Z a AG. Si no se encuentran '
        'precios, el archivo se salta automáticamente.'
    )

    doc.add_page_break()

    # =========================================================
    # 5. FORMATO DE PLANTILLAS DE SALIDA
    # =========================================================
    doc.add_heading('5. Formato de Plantillas de Salida', level=1)

    doc.add_paragraph(
        'Las plantillas generadas son archivos Excel (.xlsx) con una hoja llamada "Plantilla". '
        'El formato difiere según si el cliente es Monopunto o Multipunto.'
    )

    # 5.1 Monopunto
    doc.add_heading('5.1 Plantilla Monopunto', level=2)
    p = doc.add_paragraph()
    run = p.add_run('Archivo: ')
    run.bold = True
    p.add_run('"{Código} Datos Monopunto {Nombre}.xlsx"')

    doc.add_paragraph('Columnas de la hoja "Plantilla":')

    table_mp = doc.add_table(rows=1, cols=3)
    table_mp.style = 'Light Grid Accent 1'
    hdr_mp = table_mp.rows[0].cells
    hdr_mp[0].text = 'Columna'
    hdr_mp[1].text = 'Encabezado'
    hdr_mp[2].text = 'Descripción'
    for cell in hdr_mp:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    mp_cols = [
        ('A', 'COMERCIALIZADORA', 'Nombre de la comercializadora'),
        ('B', 'Recibida', 'Fecha de recepción de la oferta'),
        ('C', 'Vencimiento', 'Fecha de vencimiento'),
        ('D', 'Validez', 'Fecha de validez'),
        ('E', 'Renovable', 'Si/No'),
        ('F', 'P1', 'Precio período 1'),
        ('G', 'P2', 'Precio período 2'),
        ('H', 'P3', 'Precio período 3'),
        ('I', 'P4', 'Precio período 4'),
        ('J', 'P5', 'Precio período 5'),
        ('K', 'P6', 'Precio período 6'),
        ('L', 'Indice', 'Índice de referencia (0 si es fijo)'),
        ('M', 'Ajuste', 'Ajuste aplicado (0 si es fijo, 1.5 si es indexado)'),
        ('N', 'Indexado', 'Si/No - Indica si el precio es fijo o indexado'),
    ]
    for col_letter, header, desc in mp_cols:
        row = table_mp.add_row()
        row.cells[0].text = col_letter
        row.cells[1].text = header
        row.cells[2].text = desc
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

    # 5.2 Multipunto
    doc.add_heading('5.2 Plantilla Multipunto', level=2)
    p = doc.add_paragraph()
    run = p.add_run('Archivo: ')
    run.bold = True
    p.add_run('"{Código} Datos Multipunto {Nombre}.xlsx"')

    doc.add_paragraph('Columnas de la hoja "Plantilla":')

    table_mt = doc.add_table(rows=1, cols=3)
    table_mt.style = 'Light Grid Accent 1'
    hdr_mt = table_mt.rows[0].cells
    hdr_mt[0].text = 'Columna'
    hdr_mt[1].text = 'Encabezado'
    hdr_mt[2].text = 'Descripción'
    for cell in hdr_mt:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    mt_cols = [
        ('A', 'Cliente', 'Código numérico del punto de suministro'),
        ('B', 'COMERCIALIZADORA', 'Nombre de la comercializadora'),
        ('C', 'Recibida', 'Fecha de recepción de la oferta'),
        ('D', 'Vencimiento', 'Fecha de vencimiento'),
        ('E', 'Validez', 'Fecha de validez'),
        ('F', 'Renovable', 'Si/No'),
        ('G', 'P1', 'Precio período 1'),
        ('H', 'P2', 'Precio período 2'),
        ('I', 'P3', 'Precio período 3'),
        ('J', 'P4', 'Precio período 4'),
        ('K', 'P5', 'Precio período 5'),
        ('L', 'P6', 'Precio período 6'),
        ('M', 'Indice', 'Índice de referencia (0 si es fijo)'),
        ('N', 'Ajuste', 'Ajuste aplicado (0 si es fijo, 1.5 si es indexado)'),
        ('O', 'Indexado', 'Si/No - Indica si el precio es fijo o indexado'),
        ('P', 'Tarifa', 'Tipo de tarifa del punto de suministro'),
    ]
    for col_letter, header, desc in mt_cols:
        row = table_mt.add_row()
        row.cells[0].text = col_letter
        row.cells[1].text = header
        row.cells[2].text = desc
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

    # 5.3
    doc.add_heading('5.3 Lógica de Generación de Filas', level=2)
    doc.add_paragraph(
        'Para cada comercializadora asociada a un cliente, la herramienta genera '
        'una o dos filas según el tipo de precio:'
    )

    p = doc.add_paragraph()
    run = p.add_run('Precio fijo: ')
    run.bold = True
    p.add_run(
        'Si la oferta contiene precios en los períodos P1 a P6 (al menos uno mayor a 0), '
        'se genera una fila con Indexado="No", Indice=0 y Ajuste=0, '
        'con los valores de P1 a P6 copiados directamente de la oferta.'
    )

    p = doc.add_paragraph()
    run = p.add_run('Precio indexado: ')
    run.bold = True
    p.add_run(
        'Si la oferta contiene valores en FEE y/o Coste de desvíos (suma > 0), '
        'se genera una fila con Indexado="Si", Indice=(FEE+Coste) redondeado a 4 decimales, '
        'Ajuste=1.5 y P1 a P6 en 0.'
    )

    doc.add_paragraph(
        'Es posible que para una misma comercializadora se generen ambas filas '
        '(una fija y una indexada) si la oferta contiene ambos tipos de datos.'
    )

    doc.add_page_break()

    # =========================================================
    # 6. CLASIFICACIÓN MONO/MULTIPUNTO
    # =========================================================
    doc.add_heading('6. Clasificación: Monopunto y Multipunto', level=1)

    doc.add_paragraph(
        'La herramienta clasifica automáticamente cada cliente según el número de '
        'puntos de suministro asociados a su código:'
    )

    table_class = doc.add_table(rows=1, cols=3)
    table_class.style = 'Light Grid Accent 1'
    hdr_class = table_class.rows[0].cells
    hdr_class[0].text = 'Tipo'
    hdr_class[1].text = 'Condición'
    hdr_class[2].text = 'Resultado'
    for cell in hdr_class:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    class_data = [
        ('Monopunto', 'El prefijo del código aparece en 1 único punto de suministro', 'Se crea plantilla en MONOPUNTO/'),
        ('Multipunto', 'El prefijo del código aparece en más de 1 punto de suministro', 'Se crea plantilla en MULTIPUNTO/'),
    ]
    for tipo, cond, result in class_data:
        row = table_class.add_row()
        row.cells[0].text = tipo
        row.cells[1].text = cond
        row.cells[2].text = result
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

    doc.add_heading('6.1 El Prefijo del Código', level=2)
    doc.add_paragraph(
        'El código del punto de suministro tiene típicamente el formato "NNN/MMM" '
        '(por ejemplo, "212/001"). La herramienta extrae la parte numérica antes de la barra '
        'como prefijo único del cliente (en este caso, "212").'
    )
    doc.add_paragraph(
        'Todos los puntos de suministro que compartan el mismo prefijo se consideran '
        'parte del mismo cliente. Por ejemplo, si un cliente tiene los códigos '
        '"212/001", "212/002" y "212/003", se clasificará como Multipunto.'
    )

    doc.add_heading('6.2 Acumulación Entre Archivos', level=2)
    doc.add_paragraph(
        'La herramienta procesa todos los archivos de oferta seleccionados de forma '
        'conjunta. Esto significa que si un cliente aparece en ofertas de distintas '
        'comercializadoras con códigos diferentes bajo el mismo prefijo, '
        'se acumularán y clasificarán correctamente.'
    )

    p = doc.add_paragraph()
    run = p.add_run('Ejemplo: ')
    run.bold = True
    p.add_run(
        'Si la oferta de ELEIA incluye al cliente "EMPRESA" con código "500/001" '
        'y la oferta de Iberdrola incluye al mismo cliente con código "500/002", '
        'la herramienta lo clasificará como Multipunto (2 códigos bajo el prefijo "500").'
    )

    doc.add_page_break()

    # =========================================================
    # 7. LISTADO DE COMERCIALIZADORAS
    # =========================================================
    doc.add_heading('7. Listado de Comercializadoras', level=1)

    doc.add_paragraph(
        'La herramienta incluye una lista predefinida de comercializadoras en el archivo '
        'Comercializadoras.txt. Esta lista se utiliza para:'
    )
    items_7 = [
        'Poblar el desplegable de comercializadoras en la tabla de ofertas.',
        'Auto-sugerir la comercializadora al detectar el nombre en el archivo de oferta.',
    ]
    for item in items_7:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)

    doc.add_paragraph('El proceso de auto-sugerencia funciona de la siguiente manera:')
    auto_steps = [
        'Se extrae el nombre de la comercializadora del nombre del archivo '
        '(todo antes de " - PETICION").',
        'Se normaliza el nombre (mayúsculas, eliminación de sufijos como ", S.L.").',
        'Se busca una coincidencia parcial en la lista predefinida.',
        'Si se encuentra coincidencia, se selecciona automáticamente en el desplegable.',
        'Si no se encuentra, se muestra la primera comercializadora de la lista '
        '(el usuario puede cambiarla manualmente).',
    ]
    for i, step in enumerate(auto_steps, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'{i}. ')
        run.bold = True
        p.add_run(step)

    doc.add_paragraph(
        'La lista actual contiene 108 comercializadoras. Para agregar o modificar '
        'comercializadoras, editar el archivo Comercializadoras.txt (una por línea).'
    )

    doc.add_page_break()

    # =========================================================
    # 8. SOLUCIÓN DE PROBLEMAS
    # =========================================================
    doc.add_heading('8. Solución de Problemas', level=1)

    problems = [
        {
            'title': 'El archivo no aparece en la tabla de ofertas',
            'causes': [
                'El archivo no tiene extensión .xlsx.',
                'El archivo es un archivo temporal de Excel (empieza con ~$).',
                'El directorio de ofertas seleccionado no contiene archivos .xlsx.',
            ],
            'solution': 'Verificar que el archivo tenga extensión .xlsx y que esté en la carpeta correcta. '
                       'Usar el botón "Refrescar" para actualizar la lista.'
        },
        {
            'title': 'Mensaje: "Hoja PETICION OFERTAS no encontrada"',
            'causes': [
                'El archivo de oferta no tiene la hoja con el nombre exacto "PETICION OFERTAS".',
                'El nombre de la hoja tiene errores de ortografía o espacios adicionales.',
            ],
            'solution': 'Abrir el archivo en Excel y verificar que exista una hoja llamada exactamente '
                       '"PETICION OFERTAS". Renombrarla si es necesario.'
        },
        {
            'title': 'Mensaje: "Formato no reconocido (sin precios en col Z-AE)"',
            'causes': [
                'El archivo no contiene datos de precios en las columnas esperadas (Z a AG).',
                'Los precios están en un formato que la herramienta no reconoce.',
                'Los datos comienzan en una fila diferente a la fila 8.',
            ],
            'solution': 'Verificar que los precios estén en las columnas Z (P1) a AG (Coste) '
                       'y que los datos comiencen en la fila 8 de la hoja "PETICION OFERTAS".'
        },
        {
            'title': 'Error al guardar: "PermissionError"',
            'causes': [
                'El archivo de plantilla está abierto en Excel.',
                'Otro proceso está usando el archivo.',
            ],
            'solution': 'Cerrar el archivo en Excel y volver a procesar. '
                       'La herramienta intentará sobrescribir el archivo existente.'
        },
        {
            'title': 'La comercializadora auto-detectada es incorrecta',
            'causes': [
                'El nombre del archivo no sigue el formato esperado.',
                'La comercializadora no está en la lista predefinida.',
            ],
            'solution': 'Seleccionar manualmente la comercializadora correcta en el desplegable '
                       'correspondiente antes de procesar.'
        },
        {
            'title': 'El ejecutable no se abre o se cierra inmediatamente',
            'causes': [
                'Faltan archivos necesarios (Comercializadoras.txt).',
                'El archivo está dañado.',
            ],
            'solution': 'El ejecutable ya incluye una copia de la lista, por lo que solo sería necesario '
                       'si se ha copiado el archivo .exe sin sus recursos. Para ampliar o corregir la '
                       'lista de comercializadoras, dejar un archivo "Comercializadoras.txt" (uno por '
                       'línea) en la misma carpeta que el ejecutable: tiene prioridad sobre la interna.'
        },
        {
            'title': 'Mensaje: "Falta la dependencia ..."',
            'causes': [
                'Falta openpyxl, customtkinter o tkcalendar en el equipo.',
            ],
            'solution': 'Instalar las dependencias con:  pip install -r requirements.txt. '
                       'Si se usa el ejecutable, no es necesario: ya van incluidas.'
        },
        {
            'title': 'Una fecha vacía no borra la fecha que ya tenía la plantilla',
            'causes': [
                'Comportamiento intencionado: una fecha vacía significa "no modificar esa celda".',
            ],
            'solution': 'Si se quiere quitar la fecha de una plantilla existente, borrarla '
                       'en Excel (guardando el archivo) o dejar de procesar con esa plantilla. '
                       'Si lo que se quiere es cambiar la fecha, escribirla en la tabla antes de procesar.'
        },
        {
            'title': 'La ventana no recuerda el tema, las carpetas o el tamaño',
            'causes': [
                'El archivo config.json se ha borrado o no se puede escribir en la carpeta '
                'de la aplicación (por ejemplo, en Archivos de Programa).',
            ],
            'solution': 'La herramienta arranca con los valores por defecto: tema claro, '
                       'carpeta Ofertas/ junto al ejecutable y tamaño por defecto. '
                       'Ejecutarla desde una carpeta con permisos de escritura, como '
                       'el Escritorio o Documentos, para que guarde las preferencias.'
        },
    ]

    for prob in problems:
        doc.add_heading(prob['title'], level=2)
        p = doc.add_paragraph()
        run = p.add_run('Posibles causas:')
        run.bold = True
        for cause in prob['causes']:
            p_cause = doc.add_paragraph(style='List Bullet')
            p_cause.add_run(cause)
        p_sol = doc.add_paragraph()
        run = p_sol.add_run('Solución: ')
        run.bold = True
        p_sol.add_run(prob['solution'])

    # =========================================================
    # GUARDAR
    # =========================================================
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'Manual - ActualizarPrecios.docx')
    doc.save(output_path)
    print(f"Manual generado exitosamente: {output_path}")
    return output_path


if __name__ == '__main__':
    create_manual()

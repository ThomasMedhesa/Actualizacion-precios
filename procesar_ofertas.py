#!/usr/bin/env python3
"""
Procesador de Ofertas Eléctricas.
- Una fila por archivo de oferta
- Asignar comercializadora, Renovable y fechas de cada oferta
- Procesar y generar/actualizar las plantillas MONOPUNTO y MULTIPUNTO

Uso: python procesar_ofertas.py
"""

import ctypes
import json
import os
import queue
import sys
import threading
import tkinter as tk
from calendar import monthrange
from datetime import date, datetime, timedelta
from tkinter import filedialog, font as tkfont, messagebox, ttk


def _error_dependencia(paquete, detalle=""):
    """En el ejecutable congelado no se puede ejecutar 'pip install'."""
    texto = (f"Falta la dependencia '{paquete}'.\n\n{detalle}\n\n"
             f"Instale con:  pip install -r requirements.txt")
    sys.stderr.write(f"{paquete}: {detalle}\n")
    try:
        ctypes.windll.user32.MessageBoxW(None, texto, "Dependencia faltante", 0x10)
    except Exception:
        pass
    sys.exit(1)


try:
    import customtkinter as ctk
except ImportError as exc:
    _error_dependencia("customtkinter", str(exc))

try:
    from tkcalendar import Calendar
except ImportError as exc:
    _error_dependencia("tkcalendar", str(exc))


def obtener_base_dir():
    """Obtiene el directorio base, compatible con PyInstaller."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def activar_dpi_awareness():
    """Nitidez en monitores con escalado 125/150%."""
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


BASE_DIR = obtener_base_dir()
ICON_PATH = os.path.join(
    getattr(sys, "_MEIPASS", BASE_DIR), "assets", "actualizar_precios.ico")
sys.path.insert(0, BASE_DIR)
import llenar_planillas as proc

OFERTAS_DEF = proc.OFERTAS_DIR
SALIDA_DEF = BASE_DIR
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
RUTA_COMERCIALIZADORAS = os.path.join(BASE_DIR, "Comercializadoras.txt")


def _leer_comercializadoras():
    """Lee Comercializadoras.txt. Al empaquetado se busca tambien dentro del ejecutable,
    y si esta al lado del .exe tiene prioridad (se puede editar a mano)."""
    candidatas = [RUTA_COMERCIALIZADORAS]
    interna = getattr(sys, "_MEIPASS", None)
    if interna:
        candidatas.append(os.path.join(interna, "Comercializadoras.txt"))
    candidatas.append(os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "Comercializadoras.txt"))
    for ruta in candidatas:
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as archivo:
                    lineas = [line.strip() for line in archivo if line.strip()]
                if lineas:
                    return lineas, ruta
            except OSError as exc:
                sys.stderr.write(f"No se pudo leer Comercializadoras.txt: {exc}\n")
    return [], None


COMERCIALIZADORAS, RUTA_LEIDA = _leer_comercializadoras()

VACIO = "—"
FECHA_FMT = "%d/%m/%Y"
# Formatos que se aceptan al escribir o pegar una fecha en la celda.
FORMATOS_FECHA = ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%d.%m.%Y",
                  "%d.%m.%y", "%Y-%m-%d", "%Y/%m/%d")

COLUMNAS = {
    "archivo": "Archivo",
    "com": "Comercializadora",
    "ren": "Renovable",
    "rec": "Recibida",
    "ven": "Vencimiento",
    "val": "Validez",
}
ORDENABLES = ("archivo", "com", "ren", "rec", "ven", "val")
CAMPOS_FECHA = {"rec": "recibida", "ven": "vencimiento", "val": "validez"}
COLUMNAS_FECHA = ("rec", "ven", "val")
SIN_COMERCIALIZADORA = "SIN COMERCIALIZADORA"

COLORES_LOG = {
    "light": {"info": "#1667B1", "ok": "#1E8E4A", "warn": "#B9770E",
              "error": "#C62828", "muted": "#9A9A9A", "marca": "#1667B1"},
    "dark": {"info": "#63B3F2", "ok": "#3DDC84", "warn": "#F0B429",
             "error": "#FF6B6B", "muted": "#7C7C7C", "marca": "#63B3F2"},
}

COLORES_ARBOL = {
    "light": {"bg": "#FFFFFF", "fg": "#1A1A1A", "alt": "#F4F6F8",
              "aviso": "#B9770E", "head_bg": "#E9EDF2", "head_fg": "#33475B",
              "head_active": "#D7E0E9", "sel_bg": "#CFE4FB", "sel_fg": "#0D3B66"},
    "dark": {"bg": "#1B1D1F", "fg": "#E6E6E6", "alt": "#212427",
             "aviso": "#F0B429", "head_bg": "#2B2F33", "head_fg": "#D6DBE0",
             "head_active": "#3A4046", "sel_bg": "#2F4A66", "sel_fg": "#FFFFFF"},
}

COLOR_BOTON = "#3B8ED0"
COLOR_BOTON_HOVER = "#36719F"
COLOR_CANCELAR = "#C0392B"
COLOR_CANCELAR_HOVER = "#962D22"

AYUDA = """CÓMO SE USA

1. Elige las carpetas de ofertas (origen) y de salida (resultados).
2. La tabla tiene UNA FILA POR ARCHIVO de oferta.
   Pulsa ☑ en la fila (o Espacio con ella enfocada) para incluirla o excluirla.
3. Clic en Comercializadora, Renovable o en cualquier fecha para editarla.
   En las fechas puedes escribir o pegar dd/mm/aaaa, o pulsar el botón 📅
   para abrir el calendario. Intro confirma; Tab confirma y baja a la fila
   siguiente; Esc deja la celda como estaba.
4. Ctrl+C copia la fila enfocada y Ctrl+V la pega en la fila enfocada.
5. Pulsa Procesar y sigue el progreso en el panel de resultado.

POR DEFECTO

Recibida    hoy
Vencimiento fin del mes actual del año siguiente (02/10/2026 -> 31/10/2027)
Validez     viernes de esta semana (07/10/2026 -> 09/10/2026)
Renovable   No

Una fecha vacía significa "no modificar": en los Excel que ya existen se deja
la celda de esa fecha tal cual estaba.

ATAJOS DE TECLADO

Ctrl+O      Elegir carpeta de ofertas
Ctrl+S      Elegir carpeta de salida
Ctrl+R      Refrescar la lista de archivos
Ctrl+A      Seleccionar todos los archivos
Ctrl+D      Deseleccionar todos los archivos
Ctrl+C      Copiar la fila enfocada al portapapeles
Ctrl+V      Pegar en la fila enfocada
Intro       Procesar / abrir el editor de la celda
Tab         Confirmar la fecha y bajar a la fila siguiente
Espacio     Alternar la casilla de la fila enfocada
Esc         Cancelar el proceso en curso / cerrar el editor sin guardar
Ctrl+L      Limpiar el panel de resultado
F1          Mostrar esta ayuda
"""


def _fin_de_mes_proximo(referencia=None):
    """Ultimo dia del mes de referencia, un ano por delante.
    Si hoy es 02/10/2026 devuelve 31/10/2027."""
    ref = referencia or date.today()
    return date(ref.year + 1, ref.month, monthrange(ref.year + 1, ref.month)[1])


def _viernes_de_esta_semana(referencia=None):
    """Viernes de la semana ISO de la fecha de referencia."""
    ref = referencia or date.today()
    return ref + timedelta(days=4 - ref.weekday())


def _parsear_fecha(texto):
    """Convierte lo escrito o pegado en una fecha.

    Devuelve (fecha, error): fecha es date o None (vacio), error es str o None.
    Acepta tabuladores y saltos de linea de un pegado desde Excel."""
    limpio = str(texto or "").replace("\t", " ").replace("\n", " ")
    limpio = limpio.replace("\r", " ").strip().strip("\"'").strip()
    if not limpio or limpio == VACIO:
        return None, None
    for formato in FORMATOS_FECHA:
        try:
            return datetime.strptime(limpio, formato).date(), None
        except ValueError:
            continue
    return None, f"'{limpio}' no es una fecha válida (usa dd/mm/aaaa)"


def _normalizar_renovable(valor):
    """'si', 'Sí', 'S', '1', 'yes'... -> 'Si'. Cualquier otra cosa -> 'No'."""
    limpio = str(valor or "").strip().lower().replace(".", "")
    return "Si" if limpio in ("si", "s", "1", "y", "yes", "true", "x") else "No"


def _normalizar_com(s):
    return proc.normalize(s) if hasattr(proc, 'normalize') else s.upper().strip()


def _buscar_comercializadora(nombre_archivo):
    extraida = proc.extract_comercializadora(nombre_archivo)
    en = _normalizar_com(extraida)
    if not en:
        return extraida
    for opt in COMERCIALIZADORAS:
        on = _normalizar_com(opt)
        if en == on or (len(en) > 2 and (en in on or on in en)):
            return opt
    return extraida


def _familia_ui(root):
    """Segoe UI Variable si existe (Windows 11), si no Segoe UI."""
    try:
        disponibles = set(tkfont.families(root))
    except Exception:
        disponibles = set()
    for nombre in ("Segoe UI Variable Text", "Segoe UI", "Tahoma"):
        if nombre in disponibles:
            return nombre
    return "TkDefaultFont"


def _fmt_fecha(d):
    return d.strftime(FECHA_FMT) if isinstance(d, date) else VACIO


def _modo():
    """'light' o 'dark' en minusculas (customtkinter devuelve 'Light'/'Dark')."""
    return (ctk.get_appearance_mode() or "light").lower()


def cargar_config():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except Exception:
        return {}


def guardar_config(config):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def abrir_carpeta(ruta):
    try:
        if os.path.isdir(ruta):
            os.startfile(ruta)
            return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# WIDGETS DE APOYO
# ---------------------------------------------------------------------------

class ToolTip:
    """Ayuda emergente para widgets sin soporte nativo."""

    def __init__(self, widget, texto, delay=500):
        self.widget = widget
        self.texto = texto
        self.delay = delay
        self._tarea = None
        self._ventana = None
        widget.bind("<Enter>", self._programar, add="+")
        widget.bind("<Leave>", self._ocultar, add="+")
        widget.bind("<ButtonPress>", self._ocultar, add="+")

    def _programar(self, _evento=None):
        self._cancelar()
        try:
            self._tarea = self.widget.after(self.delay, self._mostrar)
        except tk.TclError:
            self._tarea = None

    def _cancelar(self):
        if self._tarea is not None:
            try:
                self.widget.after_cancel(self._tarea)
            except Exception:
                pass
            self._tarea = None

    def _ocultar(self, _evento=None):
        self._cancelar()
        self._destruir()

    def _destruir(self):
        if self._ventana is not None:
            try:
                self._ventana.destroy()
            except Exception:
                pass
            self._ventana = None

    def _mostrar(self):
        if self._ventana is not None:
            return
        try:
            self._ventana = ctk.CTkToplevel(self.widget)
            self._ventana.wm_overrideredirect(True)
            self._ventana.attributes("-topmost", True)
            etiqueta = ctk.CTkLabel(
                self._ventana, text=self.texto, justify="left", corner_radius=6,
                fg_color=("gray92", "gray20"), text_color=("black", "white"),
                font=("Segoe UI", 11), wraplength=330, padx=2, pady=1,
            )
            etiqueta.pack()
            self._ventana.update_idletasks()
            x = self.widget.winfo_rootx() + 14
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
            ancho, alto = self._ventana.winfo_width(), self._ventana.winfo_height()
            if x + ancho > self.widget.winfo_screenwidth() - 8:
                x = max(8, self.widget.winfo_screenwidth() - ancho - 8)
            if y + alto > self.widget.winfo_screenheight() - 8:
                y = self.widget.winfo_rooty() - alto - 8
            self._ventana.geometry(f"+{x}+{y}")
        except Exception:
            self._destruir()


class SelectorFecha(ctk.CTkToplevel):
    """Calendario emergente: el valor devuelto es una fecha o None."""

    def __init__(self, parent, titulo, inicial, al_confirmar):
        super().__init__(parent)
        self.al_confirmar = al_confirmar
        self.title(titulo)
        self.resizable(False, False)
        self.configure(fg_color=("gray95", "gray13"))
        self.transient(parent.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.bind("<Return>", lambda _e: self._aceptar())

        colores = {
            "Dark": dict(background="#2A2C2E", foreground="#E6E6E6", bordercolor="#2A2C2E",
                         normalbackground="#2A2C2E", normalforeground="#E6E6E6",
                         othermonthbackground="#222426", othermonthforeground="#707070",
                         othermonthwebackground="#222426", othermonthweforeground="#707070",
                         weekendbackground="#303336", weekendforeground="#C8C8C8"),
            "Light": dict(background="#F5F7FA", foreground="#1F2933", bordercolor="#D6DCE4",
                          normalbackground="#FFFFFF", normalforeground="#1F2933",
                          othermonthbackground="#EDF0F4", othermonthforeground="#9AA5B1",
                          othermonthwebackground="#EDF0F4", othermonthweforeground="#9AA5B1",
                          weekendbackground="#F0F3F8", weekendforeground="#4A5563"),
        }[_modo().capitalize()]
        cal = Calendar(
            self, date_pattern="dd/mm/yyyy", borderwidth=0,
            headersbackground="#3B8ED0", headersforeground="#FFFFFF",
            selectbackground="#3B8ED0", selectforeground="#FFFFFF",
            **colores)
        cal.pack(padx=10, pady=(10, 4))
        if isinstance(inicial, date):
            cal.selection_set(inicial)
        else:
            cal.selection_clear()
        self.calendario = cal

        botones = ctk.CTkFrame(self, fg_color="transparent")
        botones.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(botones, text="Aceptar", width=92, command=self._aceptar).pack(
            side="left", padx=(0, 6))
        ctk.CTkButton(botones, text="Limpiar", width=92, fg_color="gray60",
                      hover_color="gray50", command=self._limpiar).pack(side="left")
        ctk.CTkButton(botones, text="Cancelar", width=92, fg_color="gray60",
                      hover_color="gray50", command=self._cancelar).pack(side="right")

        self.update_idletasks()
        self._centrar(parent)
        self.grab_set()
        self.focus_force()

    def _centrar(self, parent):
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _seleccion(self):
        try:
            sel = self.calendario.selection_get()
        except Exception:
            return None
        if not sel:
            return None
        return date(sel.year, sel.month, sel.day)

    def _aceptar(self):
        fecha = self._seleccion()
        self.destroy()
        self.al_confirmar(fecha)

    def _limpiar(self):
        self.destroy()
        self.al_confirmar(None)

    def _cancelar(self):
        self.destroy()


class Oferta:
    """Un archivo .xlsx de oferta: una fila de la tabla con todos sus datos.

    Los valores por defecto son: Recibida hoy, Vencimiento a fin del mes
    actual del año siguiente, Validez el viernes de esta semana y Renovable
    "No". Al recargar la carpeta se reutiliza `anterior` (la Oferta previa
    del mismo archivo) para no perder lo que el usuario haya escrito a mano."""

    __slots__ = ("nombre", "comercializadora", "renovable", "recibida",
                 "vencimiento", "validez", "seleccionada")

    def __init__(self, nombre, comercializadora, anterior=None, seleccionada=None):
        self.nombre = nombre
        self.comercializadora = comercializadora or SIN_COMERCIALIZADORA
        self.renovable = anterior.renovable if anterior else "No"
        self.recibida = (anterior.recibida if anterior and anterior.recibida
                         else date.today())
        self.vencimiento = (anterior.vencimiento if anterior
                            else _fin_de_mes_proximo())
        self.validez = (anterior.validez if anterior
                        else _viernes_de_esta_semana())
        self.seleccionada = (anterior.seleccionada if anterior and seleccionada is None
                             else True if seleccionada is None else seleccionada)


# ---------------------------------------------------------------------------
# VENTANA PRINCIPAL
# ---------------------------------------------------------------------------

class OfertasView(ctk.CTk):
    def __init__(self, config=None):
        super().__init__()
        self.config = dict(config or {})
        self.title("Gestor de Ofertas Eléctricas")
        self.iconbitmap(ICON_PATH)
        self.minsize(960, 640)

        self.familia = _familia_ui(self)
        # customtkinter usa Roboto por defecto, que no existe en Windows.
        ctk.ThemeManager.theme["CTkFont"]["family"] = self.familia

        self.ofertas = {}            # nombre de archivo -> Oferta
        self.filas = {}              # iid de la tabla -> nombre de archivo
        self._editor = None
        self._orden = ("archivo", False)
        self._col_enfoque = "com"    # columna del último clic (la que abre Intro)
        self._portapapeles = None    # ultima fila copiada, para pegar con Ctrl+V
        self._visibles = 0
        self._cola_log = queue.Queue()
        self._hilo = None
        self.procesando = False
        self._tarea_config = None
        self._config_advertida = False

        self._filtro = tk.StringVar()
        self.var_dir_ofertas = tk.StringVar(value=self._ruta_guardada("dir_ofertas", OFERTAS_DEF))
        self.var_dir_salida = tk.StringVar(value=self._ruta_guardada("dir_salida", SALIDA_DEF))
        self.var_modo = tk.StringVar(
            value="Oscuro" if _modo() == "dark" else "Claro")

        self._construir_ui()
        proc.set_log_hook(self._encolar_log)

        orden = self.config.get("orden")
        if isinstance(orden, list) and len(orden) == 2 and orden[0] in ORDENABLES:
            self._orden = (orden[0], bool(orden[1]))

        self.listar_ofertas()
        self._bucle_log()
        self._aplicar_tema()
        self._atajos()
        self.geometry(self.config.get("geometry") or "1180x860")
        self.after(80, self._centrar)
        self.bind("<Configure>", self._config_pendiente)
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

    # -- arranque ----------------------------------------------------------

    def _ruta_guardada(self, clave, defecto):
        valor = self.config.get(clave)
        if isinstance(valor, str) and os.path.isdir(valor):
            return valor
        return defecto

    def _centrar(self):
        try:
            ancho, alto = 1180, 860
            if self.config.get("geometry"):
                return  # la posicion guardada manda
            x = (self.winfo_screenwidth() - ancho) // 2
            y = (self.winfo_screenheight() - alto) // 3
            self.geometry(f"{ancho}x{alto}+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

    def _atajos(self):
        enlaces = (
            ("<Control-o>", self._elegir_dir_ofertas, False),
            ("<Control-s>", self._elegir_dir_salida, False),
            ("<Control-r>", self.listar_ofertas, False),
            ("<Control-l>", self._limpiar_log, False),
            ("<Control-Return>", self.procesar, False),
            ("<Control-KP_Enter>", self.procesar, False),
            ("<F1>", self._mostrar_ayuda, False),
            ("<Escape>", self._al_escapar, False),
            ("<Return>", self._intro_atajo, True),
            ("<Control-a>", self.seleccionar_todas, True),
            ("<Control-d>", self.deseleccionar_todas, True),
            ("<Control-c>", self._copiar_fila, True),
            ("<Control-v>", self._pegar_en_fila, True),
        )
        for secuencia, funcion, fuera_de_campos in enlaces:
            self.bind_all(secuencia, self._atajo(funcion, fuera_de_campos))

    def _atajo(self, funcion, fuera_de_campos=False):
        def manejador(_evento):
            if fuera_de_campos and self._en_campo_texto():
                return None
            funcion()
            return "break"
        return manejador

    CLASES_TEXTO = ("Entry", "Text", "TEntry", "TCombobox")

    def _en_campo_texto(self):
        """True si el foco esta en algo donde se escribe texto.

        Los Entry de customtkinter son Frames que por dentro tienen un Entry
        de Tcl, asi que se sube por los padres hasta encontrarlo."""
        classes = self.CLASES_TEXTO
        try:
            widget = self.focus_get()
        except Exception:
            return False
        while widget is not None:
            if widget.winfo_class() in clases or isinstance(
                    widget, (ctk.CTkEntry, ctk.CTkComboBox, ctk.CTkOptionMenu,
                             ctk.CTkTextbox)):
                return True
            widget = getattr(widget, "master", None)
        return False

    def _foco_en_tabla(self):
        try:
            return self.focus_get() is self.tree
        except Exception:
            return False

    def _intro_atajo(self):
        """Intro procesa, salvo que el foco esté en la tabla o haya un editor abierto."""
        if self._editor is not None or self._en_campo_texto() or self._foco_en_tabla():
            return
        self.procesar()

    def _al_escapar(self):
        if self.procesando:
            self._cancelar()
        elif self._editor is not None:
            self._cerrar_editor(False)

    # -- construccion de la interfaz ---------------------------------------

    def _construir_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Cabecera ---
        cabecera = ctk.CTkFrame(self, fg_color="transparent")
        cabecera.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        cabecera.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cabecera, text="Gestor de Ofertas Eléctricas",
                     font=(self.familia, 22, "bold"), anchor="w").grid(
            row=0, column=0, sticky="w")
        ctk.CTkLabel(cabecera,
                     text="Plantillas de precios eléctricas a partir de los archivos de oferta",
                     font=(self.familia, 12), text_color=("gray40", "gray62"),
                     anchor="w").grid(row=1, column=0, sticky="w", pady=(0, 2))
        self.sw_modo = ctk.CTkSwitch(cabecera, text="Modo oscuro", variable=self.var_modo,
                                     onvalue="Oscuro", offvalue="Claro",
                                     command=self._cambiar_modo, font=(self.familia, 13))
        self.sw_modo.grid(row=0, column=1, rowspan=2, padx=(0, 12))
        ToolTip(self.sw_modo, "Alterna entre el tema claro y el oscuro")
        btn_ayuda = ctk.CTkButton(cabecera, text="?", width=36, height=36,
                                  font=(self.familia, 16, "bold"), command=self._mostrar_ayuda)
        btn_ayuda.grid(row=0, column=2, rowspan=2, sticky="ne")
        ToolTip(btn_ayuda, "Ayuda y atajos de teclado (F1)")

        # --- Carpetas ---
        self.card_rutas, contenido = self._tarjeta(
            1, "CARPETAS", "Las carpetas MONOPUNTO y MULTIPUNTO se crean solas en la de salida")
        contenido.grid_columnconfigure(1, weight=1)
        for fila, (titulo, variable, accion) in enumerate((
                ("Ofertas (origen)", self.var_dir_ofertas, self._elegir_dir_ofertas),
                ("Salida (resultados)", self.var_dir_salida, self._elegir_dir_salida))):
            ctk.CTkLabel(contenido, text=titulo, width=140, anchor="w",
                         font=(self.familia, 13)).grid(
                row=fila, column=0, sticky="w", pady=4)
            campo = ctk.CTkLabel(contenido, textvariable=variable, anchor="w", height=32,
                                 corner_radius=8, fg_color=("gray93", "gray20"),
                                 font=(self.familia, 12))
            campo.grid(row=fila, column=1, sticky="ew", padx=(0, 8), pady=4)
            ToolTip(campo, f"{titulo}\nClic para elegir la carpeta")
            ctk.CTkButton(contenido, text="Seleccionar…", width=118,
                          command=accion).grid(row=fila, column=2, pady=4)
            campo.bind("<Button-1>", lambda _e, f=accion: (f(), "break")[1])

        # --- Ofertas ---
        self.card_tabla, contenido = self._tarjeta(
            2, "OFERTAS", "Clic en ☑ para incluir un archivo · clic en una celda para editarla")
        contenido.grid_rowconfigure(1, weight=1)

        herramientas = ctk.CTkFrame(contenido, fg_color="transparent")
        herramientas.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        herramientas.grid_columnconfigure(1, weight=1)
        self.ent_filtro = ctk.CTkEntry(herramientas, textvariable=self._filtro, width=300,
                                        placeholder_text="Filtrar archivo o comercializadora…",
                                        font=(self.familia, 13))
        self.ent_filtro.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.ent_filtro.bind("<KeyRelease>", lambda _e: self._reconstruir())
        self.ent_filtro.bind("<Escape>", lambda _e: self._limpiar_filtro())
        ToolTip(self.ent_filtro, "Escribe para filtrar. Escape lo borra")
        self.lbl_contador = ctk.CTkLabel(herramientas, text="", anchor="w",
                                         text_color=("gray40", "gray62"),
                                         font=(self.familia, 12))
        self.lbl_contador.grid(row=0, column=1, sticky="w")
        self.btn_todas = ctk.CTkButton(herramientas, text="Seleccionar todas", width=140,
                                        command=self.seleccionar_todas, fg_color=("gray70", "gray35"))
        self.btn_todas.grid(row=0, column=2, padx=4)
        self.btn_ninguna = ctk.CTkButton(herramientas, text="Ninguna", width=86,
                                         command=self.deseleccionar_todas, fg_color=("gray70", "gray35"))
        self.btn_ninguna.grid(row=0, column=3, padx=4)
        ctk.CTkButton(herramientas, text="↻ Refrescar", width=104,
                      command=self.listar_ofertas, fg_color=("gray70", "gray35")).grid(
            row=0, column=4, padx=4)

        self.marco_tabla = ctk.CTkFrame(contenido, fg_color="#FFFFFF", corner_radius=10)
        self.marco_tabla.grid(row=1, column=0, sticky="nsew", pady=(0, 2))
        self.marco_tabla.grid_columnconfigure(0, weight=1)
        self.marco_tabla.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(self.marco_tabla, style="Ofertas.Treeview",
                                 columns=("com", "ren", "rec", "ven", "val"),
                                 show="tree headings", selectmode="browse", height=8)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.column("#0", width=360, minwidth=140, stretch=True, anchor="w")
        for col, ancho, ancla in (("com", 200, "w"), ("ren", 90, "center"),
                                  ("rec", 105, "center"), ("ven", 115, "center"),
                                  ("val", 105, "center")):
            self.tree.heading(col, text=COLUMNAS[col],
                              command=lambda c=col: self._ordenar_por(c))
            self.tree.column(col, width=ancho, minwidth=70, stretch=False, anchor=ancla)
        self.tree.heading("#0", text=COLUMNAS["archivo"],
                          command=lambda: self._ordenar_por("archivo"))
        ToolTip(self.tree, "Una fila por archivo de oferta.\n"
                           "Clic en ☑ para incluirlo, clic en una celda para editarla.\n"
                           "En las fechas: escribe o pega dd/mm/aaaa, o pulsa 📅.\n"
                           "Ctrl+C copia la fila enfocada y Ctrl+V la pega en otra.\n"
                           "⚠ avisa de que dos archivos son de la misma comercializadora.",
                delay=900)
        scroll_v = ctk.CTkScrollbar(self.marco_tabla, command=self.tree.yview, width=14)
        scroll_v.grid(row=0, column=1, sticky="ns")
        scroll_h = ctk.CTkScrollbar(self.marco_tabla, command=self.tree.xview,
                                    orientation="horizontal", height=14)
        scroll_h.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scroll_v.set, xscrollcommand=scroll_h.set)
        self.tree.tag_configure("duplicado", foreground="#B9770E")
        self.tree.bind("<Button-1>", self._clic_arbol)
        self.tree.bind("<Double-Button-1>", self._doble_clic_arbol)
        self.tree.bind("<space>", self._espacio_arbol)
        self.tree.bind("<Return>", lambda _e: (self._abrir_editor_actual(), "break")[1])

        # --- Resultado ---
        self.card_log, contenido = self._tarjeta(3, "RESULTADO", "Procesamiento en tiempo real")
        contenido.grid_rowconfigure(1, weight=1)
        contenido.grid_columnconfigure(0, weight=1)
        acciones = ctk.CTkFrame(contenido, fg_color="transparent")
        acciones.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        acciones.grid_columnconfigure(0, weight=1)
        self.lbl_log = ctk.CTkLabel(acciones, text="Ctrl+L limpia este panel", anchor="w",
                                    text_color=("gray45", "gray58"), font=(self.familia, 12))
        self.lbl_log.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(acciones, text="Limpiar", width=84, height=28, corner_radius=8,
                      fg_color=("gray70", "gray35"), command=self._limpiar_log).grid(
            row=0, column=1, padx=4)
        self.btn_mono = ctk.CTkButton(acciones, text="📂 MONOPUNTO", width=142, height=28,
                                      corner_radius=8,
                                      command=lambda: self._abrir_salida("MONOPUNTO"))
        self.btn_mono.grid(row=0, column=2, padx=4)
        self.btn_multi = ctk.CTkButton(acciones, text="📂 MULTIPUNTO", width=142, height=28,
                                       corner_radius=8,
                                       command=lambda: self._abrir_salida("MULTIPUNTO"))
        self.btn_multi.grid(row=0, column=3, padx=4)
        ToolTip(self.btn_mono, "Abre la carpeta MONOPUNTO de la carpeta de salida")
        ToolTip(self.btn_multi, "Abre la carpeta MULTIPUNTO de la carpeta de salida")
        self.txt_log = ctk.CTkTextbox(contenido, height=120, corner_radius=10,
                                      font=("Consolas", 12), wrap="word",
                                      fg_color=("gray97", "gray12"))
        self.txt_log.grid(row=1, column=0, sticky="nsew")
        self.txt_log.configure(state="disabled")

        # --- Barra de estado ---
        barra = ctk.CTkFrame(self, fg_color="transparent")
        barra.grid(row=4, column=0, sticky="ew", padx=18, pady=(10, 16))
        barra.grid_columnconfigure(1, weight=1)
        self.barra_progreso = ctk.CTkProgressBar(barra, width=170, height=6,
                                                  mode="indeterminate", corner_radius=3)
        self.barra_progreso.grid(row=0, column=0, sticky="w", padx=(0, 14))
        self.barra_progreso.grid_remove()
        self.lbl_estado = ctk.CTkLabel(barra, text="Listo", anchor="w",
                                        font=(self.familia, 13))
        self.lbl_estado.grid(row=0, column=1, sticky="ew")
        self.btn_procesar = ctk.CTkButton(barra, text="Procesar", width=240, height=46,
                                          corner_radius=12, font=(self.familia, 15, "bold"),
                                          command=self.procesar)
        self.btn_procesar.grid(row=0, column=2)
        ToolTip(self.btn_procesar, "Procesa las ofertas marcadas (Ctrl+Enter). "
                                   "Durante el proceso pasa a ser Cancelar (Esc)")
        self._refrescar_botones_salida()

    def _tarjeta(self, fila, titulo, subtitulo):
        tarjeta = ctk.CTkFrame(self, corner_radius=14, border_width=1,
                               border_color=("gray82", "gray25"))
        tarjeta.grid(row=fila, column=0, sticky="nsew", padx=18, pady=(0, 12))
        tarjeta.grid_columnconfigure(0, weight=1)
        cabecera = ctk.CTkFrame(tarjeta, fg_color="transparent")
        cabecera.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 0))
        cabecera.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(cabecera, text=titulo, font=(self.familia, 12, "bold"),
                     text_color=("gray35", "gray85")).grid(row=0, column=0, sticky="w")
        if subtitulo:
            ctk.CTkLabel(cabecera, text=subtitulo, font=(self.familia, 11),
                         text_color=("gray50", "gray60"), anchor="e").grid(
                row=0, column=1, sticky="e")
        contenido = ctk.CTkFrame(tarjeta, fg_color="transparent")
        contenido.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8, 14))
        contenido.grid_columnconfigure(0, weight=1)
        tarjeta.grid_rowconfigure(1, weight=1)
        return tarjeta, contenido

    # -- tema --------------------------------------------------------------

    def _cambiar_modo(self):
        ctk.set_appearance_mode("dark" if self.var_modo.get() == "Oscuro" else "light")
        self._aplicar_tema()
        self._guardar_config()

    def _aplicar_tema(self):
        modo = _modo()
        colores = COLORES_ARBOL[modo]
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("Ofertas.Treeview", background=colores["bg"],
                         fieldbackground=colores["bg"], foreground=colores["fg"],
                         rowheight=30, borderwidth=0, relief="flat", bordercolor=colores["bg"],
                         font=(self.familia, 12))
        estilo.map("Ofertas.Treeview", background=[("selected", colores["sel_bg"])],
                   foreground=[("selected", colores["sel_fg"])])
        estilo.configure("Ofertas.Treeview.Heading", background=colores["head_bg"],
                         foreground=colores["head_fg"], relief="flat", borderwidth=0,
                         padding=(8, 9), font=(self.familia, 11, "bold"))
        estilo.map("Ofertas.Treeview.Heading", background=[("active", colores["head_active"])])
        estilo.layout("Ofertas.Treeview",
                      [("Ofertas.Treeview.treearea", {"sticky": "nswe"})])
        self.marco_tabla.configure(fg_color=colores["bg"])
        self.tree.tag_configure("duplicado", foreground=colores["aviso"])
        self.tree.tag_configure("impar", background=colores["alt"])
        self.tree.tag_configure("par", background=colores["bg"])
        self._estilizar_log()

    def _estilizar_log(self):
        colores = COLORES_LOG[_modo()]
        # CTkTextbox no admite 'font' en tag_config (romperia el escalado).
        for nivel in ("info", "ok", "warn", "error", "marca"):
            self.txt_log.tag_config(nivel, foreground=colores[nivel])
        self.txt_log.tag_config("ts", foreground=colores["muted"])

    # -- log ---------------------------------------------------------------

    def _encolar_log(self, nivel, texto):
        """Callback del backend: solo encola (se invoca desde el hilo de trabajo)."""
        self._cola_log.put((nivel, texto))

    def _bucle_log(self):
        try:
            while True:
                nivel, texto = self._cola_log.get_nowait()
                if nivel == "_fin":
                    self._terminar(texto)
                else:
                    self._escribir_log(nivel, texto)
        except queue.Empty:
            pass
        except Exception:
            pass
        try:
            self.after(80, self._bucle_log)
        except tk.TclError:
            pass

    def _escribir_log(self, nivel, texto):
        self.txt_log.configure(state="normal")
        hora = datetime.now().strftime("%H:%M:%S")
        lineas = str(texto).rstrip("\n").split("\n")
        for i, linea in enumerate(lineas):
            self.txt_log.insert("end", f"{hora}  " if i == 0 else "        ", "ts")
            if linea.strip() and set(linea.strip()) <= set("=-"):
                self.txt_log.insert("end", linea + "\n", "ts")
            else:
                self.txt_log.insert("end", linea + "\n", nivel)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _aviso(self, texto, nivel="info"):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", datetime.now().strftime("%H:%M:%S") + "  " + texto + "\n", nivel)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _limpiar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    # -- carpetas ----------------------------------------------------------

    def _elegir_dir_ofertas(self):
        actual = self.var_dir_ofertas.get()
        d = filedialog.askdirectory(title="Carpeta con los archivos de oferta",
                                    initialdir=actual if os.path.isdir(actual) else BASE_DIR)
        if d:
            self.var_dir_ofertas.set(d)
            self._guardar_config()
            self.listar_ofertas()

    def _elegir_dir_salida(self):
        actual = self.var_dir_salida.get()
        d = filedialog.askdirectory(title="Carpeta donde guardar los resultados",
                                    initialdir=actual if os.path.isdir(actual) else BASE_DIR)
        if d:
            self.var_dir_salida.set(d)
            self._guardar_config()
            self._refrescar_botones_salida()

    def _abrir_salida(self, subcarpeta):
        ruta = os.path.join(self.var_dir_salida.get(), subcarpeta)
        if not abrir_carpeta(ruta):
            messagebox.showinfo("Carpeta vacía",
                                f"Todavía no existe la carpeta:\n{ruta}", parent=self)

    def _refrescar_botones_salida(self):
        base = self.var_dir_salida.get()
        for boton, sub in ((self.btn_mono, "MONOPUNTO"), (self.btn_multi, "MULTIPUNTO")):
            existe = os.path.isdir(os.path.join(base, sub))
            boton.configure(state="normal" if existe else "disabled")

    # -- listado de archivos -----------------------------------------------

    def listar_ofertas(self):
        if self.procesando:
            return
        directorio = self.var_dir_ofertas.get()
        self._cerrar_editor(False)

        encontrados = []
        if os.path.isdir(directorio):
            for fname in sorted(os.listdir(directorio)):
                if fname.endswith(".xlsx") and not fname.startswith("~$"):
                    encontrados.append(fname)

        # Cada archivo conserva sus datos al recargar la carpeta.
        previas = self.ofertas
        self.ofertas = {}
        for fname in encontrados:
            anterior = previas.get(fname)
            com = (anterior.comercializadora if anterior
                   else _buscar_comercializadora(fname).upper())
            self.ofertas[fname] = Oferta(fname, com.strip().upper(), anterior)
        self._reconstruir()
        if encontrados:
            self._aviso(f"{len(encontrados)} archivo(s) .xlsx encontrado(s) en {directorio}\n")
        else:
            self._aviso(f"No hay archivos .xlsx en {directorio}\n", "warn")

    # -- tabla -------------------------------------------------------------

    def _reconstruir(self):
        """Repinta la tabla a partir del modelo (filtro + orden)."""
        pos_y = self.tree.yview()[0] if self.tree.get_children() else 0.0
        # La fila enfocada y la seleccion se reponen: si no, despues de editar
        # una celda el Ctrl+V no tendria fila de destino.
        foco_previo = self.tree.focus()
        seleccion_previa = [i for i in self.tree.selection() if i in self.ofertas]
        self._cerrar_editor(False)

        filtro = self._filtro.get().strip().lower()
        columna, inversa = self._orden
        claves = []
        for oferta in self.ofertas.values():
            if filtro and not (filtro in oferta.nombre.lower()
                               or filtro in oferta.comercializadora.lower()):
                continue
            claves.append((self._clave_fila(oferta), oferta))
        claves.sort(key=lambda t: t[0], reverse=inversa)

        self.tree.delete(*self.tree.get_children())
        self.filas.clear()
        self._visibles = 0
        for indice, (_clave, oferta) in enumerate(claves):
            com = oferta.comercializadora or VACIO
            self.tree.insert(
                "", "end", iid=oferta.nombre,
                text=f"{'☑' if oferta.seleccionada else '☐'} {oferta.nombre}",
                values=(com, oferta.renovable, _fmt_fecha(oferta.recibida),
                        _fmt_fecha(oferta.vencimiento), _fmt_fecha(oferta.validez)),
                tags=("impar" if indice % 2 else "par",))
            self.filas[oferta.nombre] = oferta.nombre
            self._visibles += 1

        self._reponer_foco(foco_previo, seleccion_previa)
        self._actualizar_titulos()
        try:
            self.tree.yview_moveto(pos_y)
        except Exception:
            pass
        self._actualizar_contadores()

    def _reponer_foco(self, foco_previo, seleccion_previa):
        """Vuelve a dejar enfocada y seleccionada la fila que estaba, si sigue
        en pantalla (al filtrar u ordenar puede desaparecer)."""
        visibles = [i for i in seleccion_previa if self.tree.exists(i)]
        if visibles:
            self.tree.selection_set(*visibles)
        if foco_previo and self.tree.exists(foco_previo):
            try:
                self.tree.focus(foco_previo)
            except tk.TclError:
                pass

    def _clave_fila(self, oferta):
        """Clave de ordenacion de una fila (empieza por 0 o por 1 segun el tipo)."""
        columna = self._orden[0]
        nombre = oferta.nombre.lower()
        if columna == "ren":
            return (1, (oferta.renovable or "").lower(), nombre)
        if columna in CAMPOS_FECHA:
            return (1, getattr(oferta, CAMPOS_FECHA[columna]) or date.min, nombre)
        if columna == "com":
            return (0, oferta.comercializadora.lower(), nombre)
        return (0, nombre)

    def _ordenar_por(self, columna):
        if columna not in ORDENABLES:
            return
        if self._orden[0] == columna:
            self._orden = (columna, not self._orden[1])
        else:
            self._orden = (columna, False)
        self._reconstruir()
        self._guardar_config()

    def _actualizar_titulos(self):
        columna, inversa = self._orden
        flecha = " ▼" if inversa else " ▲"
        self.tree.heading("#0", text=COLUMNAS["archivo"] + (flecha if columna == "archivo" else ""))
        for col in COLUMNAS:
            if col == "archivo":
                continue
            marca = flecha if columna == col else ""
            self.tree.heading(col, text=COLUMNAS[col] + marca)

    def _actualizar_contadores(self):
        total = len(self.ofertas)
        seleccionados = len(self._seleccion())
        coms = len({o.comercializadora for o in self.ofertas.values()})
        if not total:
            texto = "No hay archivos .xlsx en la carpeta seleccionada"
        else:
            texto = (f"Mostrando {self._visibles} de {total} archivo(s) · "
                     f"{seleccionados} seleccionado(s) · {coms} comercializadora(s)")
        self.lbl_contador.configure(text=texto)
        self.btn_procesar.configure(text=f"Procesar {seleccionados} oferta(s)"
                                    if seleccionados else "Procesar")

    def _seleccion(self):
        """Ofertas marcadas, en orden alfabético (estable entre llamadas)."""
        return sorted((o for o in self.ofertas.values() if o.seleccionada),
                      key=lambda o: o.nombre.lower())

    def _oferta_de_fila(self, iid):
        if not iid:
            return None
        return self.ofertas.get(self.filas.get(iid, iid))

    def _fila_activa(self):
        iid = self.tree.focus()
        if not iid:
            seleccion = self.tree.selection()
            iid = seleccion[0] if seleccion else ""
        return iid

    def _alternar_seleccion(self, iid):
        if self.procesando:
            return
        oferta = self._oferta_de_fila(iid)
        if oferta is None:
            return
        oferta.seleccionada = not oferta.seleccionada
        marca = "☑" if oferta.seleccionada else "☐"
        texto = self.tree.item(iid, "text")
        self.tree.item(iid, text=texto.replace("☑", marca, 1).replace("☐", marca, 1))
        self._actualizar_contadores()

    def seleccionar_todas(self):
        if self.procesando:
            return
        for oferta in self.ofertas.values():
            oferta.seleccionada = True
        self._reconstruir()

    def deseleccionar_todas(self):
        if self.procesando:
            return
        for oferta in self.ofertas.values():
            oferta.seleccionada = False
        self._reconstruir()

    def _limpiar_filtro(self):
        self._filtro.set("")
        self.ent_filtro.delete(0, "end")
        self._reconstruir()

    # -- interaccion con la tabla ------------------------------------------

    def _columna_clave(self, identificador):
        """Convierte '#1' en el nombre real de la columna ('com', 'ven', ...)."""
        if not identificador:
            return None
        if not identificador.startswith("#"):
            return identificador
        try:
            return self.tree["columns"][int(identificador[1:]) - 1]
        except Exception:
            return None

    def _clic_arbol(self, evento, forzar=False):
        if self.procesando:
            return None
        region = self.tree.identify_region(evento.x, evento.y)
        if region not in ("tree", "cell"):
            return None
        iid = self.tree.identify_row(evento.y)
        col = self._columna_clave(self.tree.identify_column(evento.x))
        if not iid:
            return None
        if col in self.tree["columns"]:
            self._col_enfoque = col      # la columna en la que se hace clic
        if region == "tree":
            self._alternar_seleccion(iid)
            return None
        self._abrir_editor(iid, col, forzar=forzar)
        return None

    def _doble_clic_arbol(self, evento):
        iid = self.tree.identify_row(evento.y)
        col = self._columna_clave(self.tree.identify_column(evento.x))
        if not iid or col in (None, "#0"):
            return None
        self._abrir_editor(iid, col, forzar=True)
        return "break"

    def _espacio_arbol(self, _evento):
        if self.procesando:
            return None
        iid = self._fila_activa()
        if not iid:
            return None
        self._alternar_seleccion(iid)
        return "break"

    def _abrir_editor_actual(self):
        """Intro abre el editor de la ultima columna en la que se hizo clic."""
        iid = self._fila_activa()
        if not iid:
            return
        col = self._col_enfoque if self._col_enfoque in self.tree["columns"] else "com"
        self._abrir_editor(iid, col)

    def _descartar_editor(self):
        """Destruye los widgets del editor actual sin confirmar nada."""
        editor = self._editor
        self._editor = None
        for widget in (editor or {}).get("widgets", []):
            try:
                widget.destroy()
            except Exception:
                pass

    def _cerrar_editor(self, commit=True):
        editor = self._editor
        if not editor:
            return
        if commit and editor["tipo"] == "com":
            valor = (editor["var"].get() or "").strip().upper()
            oferta = self._oferta_de_fila(editor["iid"])
            if oferta is not None and valor and valor != editor["com_origen"]:
                oferta.comercializadora = valor
                self._aviso(f"{oferta.nombre}: comercializadora cambiada a {valor}\n")
        self._descartar_editor()
        if commit:
            self._reconstruir()

    def _abrir_editor(self, iid, col, forzar=False):
        if self._editor is not None and self._editor.get("iid") == iid and \
                self._editor.get("col") == col:
            return
        if self.procesando or self._oferta_de_fila(iid) is None:
            return
        self._cerrar_editor(True)
        caja = self.tree.bbox(iid, col)
        if not caja:
            # La fila puede estar fuera de la vista: se trae y se reintenta.
            self.tree.see(iid)
            self.update_idletasks()
            caja = self.tree.bbox(iid, col)
        if not caja:
            return
        x, y, ancho, alto = caja
        if col == "com":
            self._editor = self._editor_comercializadora(iid, x, y, ancho, alto)
        elif col == "ren":
            self._editor = self._editor_renovable(iid, x, y, ancho, alto)
        elif col in CAMPOS_FECHA:
            self._editor = self._editor_fecha(iid, col, x, y, ancho, alto)

    def _editor_comercializadora(self, iid, x, y, ancho, alto):
        oferta = self._oferta_de_fila(iid)
        com = oferta.comercializadora
        var = ctk.StringVar(value=com)
        combo = ctk.CTkComboBox(self.tree, width=max(ancho, 210), height=max(alto - 4, 24),
                                variable=var, values=COMERCIALIZADORAS,
                                font=(self.familia, 12), border_width=1,
                                dropdown_font=(self.familia, 12))
        combo.place(x=x, y=y)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._cerrar_editor(True))
        combo.bind("<Escape>", lambda _e: self._cerrar_editor(False))
        combo.bind("<FocusOut>", lambda _e: self.after(80, self._cerrar_editor, True))
        combo.focus_set()
        return {"tipo": "com", "iid": iid, "col": "com", "widgets": [combo], "var": var,
                "com_origen": com}

    def _editor_renovable(self, iid, x, y, ancho, alto):
        oferta = self._oferta_de_fila(iid)
        var = ctk.StringVar(value=oferta.renovable)

        def al_cambiar(valor=None):
            self._descartar_editor()
            self._oferta_de_fila(iid).renovable = _normalizar_renovable(valor or var.get())
            self._reconstruir()

        menu = ctk.CTkOptionMenu(self.tree, width=max(ancho, 90), height=max(alto - 4, 24),
                                 values=["No", "Si"], command=al_cambiar, variable=var,
                                 font=(self.familia, 12), dropdown_font=(self.familia, 12),
                                 corner_radius=6)
        menu.place(x=x, y=y)
        menu.bind("<Escape>", lambda _e: self._cerrar_editor(False))
        return {"tipo": "ren", "iid": iid, "col": "ren", "widgets": [menu]}

    # -- editor de fecha ----------------------------------------------------

    def _editor_fecha(self, iid, col, x, y, ancho, alto):
        """Campo de texto para escribir o pegar la fecha + boton de calendario."""
        oferta = self._oferta_de_fila(iid)
        fecha = getattr(oferta, CAMPOS_FECHA[col])
        alto_editor = max(alto - 4, 24)
        ancho_texto = max(ancho - 32, 56)
        var = ctk.StringVar(value=_fmt_fecha(fecha))

        entrada = ctk.CTkEntry(self.tree, textvariable=var, width=ancho_texto,
                               height=alto_editor, font=(self.familia, 12),
                               border_width=1, justify="center")
        entrada.place(x=x, y=y)

        boton = ctk.CTkButton(self.tree, text="📅", width=28, height=alto_editor,
                              corner_radius=6, font=(self.familia, 12), border_width=1)
        boton.place(x=x + ancho_texto + 2, y=y)
        # Se abre en Button-1 (no al soltar) para que no se cuele el commit
        # que dispara el FocusOut del campo de texto.
        boton.bind("<Button-1>",
                   lambda _e: self._abrir_calendario(iid, col, var.get()), add="+")

        entrada.bind("<Return>", lambda _e: (self._confirmar_fecha(iid, col, var.get()), "break")[1])
        entrada.bind("<KP_Enter>", lambda _e: (self._confirmar_fecha(iid, col, var.get()), "break")[1])
        entrada.bind("<Tab>", lambda _e: (self._confirmar_fecha(iid, col, var.get(), 1), "break")[1])
        entrada.bind("<Shift-Tab>", lambda _e: (self._confirmar_fecha(iid, col, var.get(), -1), "break")[1])
        entrada.bind("<Escape>", lambda _e: (self._descartar_editor(), "break")[1])
        entrada.bind("<FocusOut>",
                     lambda _e: self.after(80, self._confirmar_fecha_solo, iid, col, var))
        entrada.select_range(0, "end")
        entrada.focus_set()
        return {"tipo": "fecha", "iid": iid, "col": col, "widgets": [entrada, boton],
                "entrada": entrada, "boton": boton, "var": var}

    def _confirmar_fecha_solo(self, iid, col, var):
        """Confirmacion al perder el foco, si el editor sigue siendo este."""
        editor = self._editor
        if editor and editor.get("tipo") == "fecha" and editor.get("iid") == iid \
                and editor.get("col") == col:
            self._confirmar_fecha(iid, col, var.get())

    def _confirmar_fecha(self, iid, col, texto, mover=0):
        """Aplica lo escrito/pegado. `mover` salta a la fila siguiente (-1 anterior)."""
        oferta = self._oferta_de_fila(iid)
        nombre = oferta.nombre if oferta else iid
        fecha, error = _parsear_fecha(texto)
        self._descartar_editor()
        if error:
            self._aviso(f"{nombre} · {COLUMNAS[col]}: {error}\n", "warn")
            self._reconstruir()
            return
        if oferta is not None:
            setattr(oferta, CAMPOS_FECHA[col], fecha)
        self._reconstruir()
        if mover:
            self._abrir_editor(self._fila_movida(iid, mover), col)

    def _fila_movida(self, iid, mover):
        filas = self.tree.get_children("")
        try:
            pos = filas.index(iid) + mover
        except ValueError:
            return None
        return filas[pos] if 0 <= pos < len(filas) else None

    def _abrir_calendario(self, iid, col, texto):
        """Abre el calendario partiendo de lo que haya escrito en el campo."""
        inicial, error = _parsear_fecha(texto)
        if error:
            inicial = getattr(self._oferta_de_fila(iid), CAMPOS_FECHA[col], None)
        editor = self._editor
        self.after(0, self._descartar_si_mismo, editor)
        SelectorFecha(self, COLUMNAS[col], inicial,
                      lambda elegida: self._aplicar_fecha(iid, col, elegida))

    def _descartar_si_mismo(self, editor):
        if self._editor is editor:
            self._descartar_editor()

    def _aplicar_fecha(self, iid, col, fecha):
        self._descartar_editor()
        oferta = self._oferta_de_fila(iid)
        if oferta is not None:
            setattr(oferta, CAMPOS_FECHA[col], fecha)
        self._reconstruir()

    # -- copiar y pegar ----------------------------------------------------

    def _copiar_fila(self):
        """Ctrl+C: la fila enfocada al portapapeles, como TSV (sirve con Excel)."""
        oferta = self._oferta_de_fila(self._fila_activa())
        if oferta is None:
            return
        self._portapapeles = self._campos_de(oferta)
        texto = "\t".join(self._portapapeles)
        try:
            self.clipboard_clear()
            self.clipboard_append(texto)
        except tk.TclError:
            pass
        self._aviso(f"Copiado: {texto}\n")

    def _pegar_en_fila(self):
        """Ctrl+V: aplica en la fila enfocada lo copiado o lo que haya en el portapapeles."""
        destino = self._oferta_de_fila(self._fila_activa())
        if destino is None:
            return
        campos = self._campos_para_pegar()
        if not campos:
            self._aviso("El portapapeles está vacío\n", "warn")
            return
        cambios = self._aplicar_campos(destino, campos)
        self._reconstruir()
        self._aviso(f"{destino.nombre}: {', '.join(cambios) if cambios else 'sin cambios'}\n")

    def _campos_para_pegar(self):
        """Lo que se pega: la ultima copia de la app o el portapapeles del sistema."""
        if self._portapapeles:
            return list(self._portapapeles)
        try:
            return self._campos_de_texto(self.clipboard_get())
        except tk.TclError:
            return []

    @staticmethod
    def _campos_de_texto(texto):
        """Trocea un texto pegado (TSV de Excel, CSV o una sola fecha)."""
        if not texto:
            return []
        primera = str(texto).replace("\r\n", "\n").replace("\r", "\n").split("\n")[0]
        if "\t" in primera:
            partes = primera.split("\t")
        elif ";" in primera:
            partes = primera.split(";")
        else:
            partes = [primera]
        return [p.strip().strip('"').strip() for p in partes][:5]

    @staticmethod
    def _campos_de(oferta):
        return [oferta.comercializadora or "", oferta.renovable or "",
                _fmt_fecha(oferta.recibida), _fmt_fecha(oferta.vencimiento),
                _fmt_fecha(oferta.validez)]

    def _aplicar_campos(self, destino, campos):
        """Aplica los campos pegados segun cuantos y de que tipo sean."""
        fechas = [(_parsear_fecha(campo)[0] is not None or not campo)
                  for campo in campos]
        cambios = []
        if len(campos) >= 5:
            destino.comercializadora = campos[0].upper()
            destino.renovable = _normalizar_renovable(campos[1])
            cambios += ["comercializadora", "renovable"]
            for col, campo in zip(COLUMNAS_FECHA, campos[2:5]):
                self._poner_fecha(destino, col, campo, cambios)
        elif len(campos) == 3 and all(fechas):
            for col, campo in zip(COLUMNAS_FECHA, campos):
                self._poner_fecha(destino, col, campo, cambios)
        elif len(campos) == 1 and (fechas[0] or campos[0]):
            col = self._col_enfoque if self._col_enfoque in CAMPOS_FECHA else "ven"
            self._poner_fecha(destino, col, campos[0], cambios)
        return cambios

    @staticmethod
    def _poner_fecha(oferta, col, texto, cambios):
        fecha, error = _parsear_fecha(texto)
        if error:
            return
        setattr(oferta, CAMPOS_FECHA[col], fecha)
        cambios.append(COLUMNAS[col].lower())

    # -- procesado ---------------------------------------------------------

    def procesar(self):
        if self.procesando:
            self._cancelar()
            return
        if not COMERCIALIZADORAS:
            messagebox.showerror(
                "Falta Comercializadoras.txt",
                f"No se encontró la lista de comercializadoras:\n{RUTA_COMERCIALIZADORAS}\n\n"
                "Colócala en la misma carpeta que el ejecutable.", parent=self)
            return

        seleccionadas = self._seleccion()
        if not seleccionadas:
            messagebox.showwarning("Sin selección",
                                   "Selecciona al menos una oferta para procesar.", parent=self)
            return
        fechas_archivo, filename_to_com = self._recolectar()
        incompletas = [nombre for nombre, d in fechas_archivo.items()
                       if not d["Vencimiento"] or not d["Validez"]]
        salida = self.var_dir_salida.get()
        detalle = (f"Ofertas seleccionadas: {len(seleccionadas)}\n"
                   f"Comercializadoras: {len({o.comercializadora for o in seleccionadas})}\n"
                   f"Carpeta de salida: {salida}\n")
        if incompletas:
            detalle += (f"\n{incompletas[0]}"
                        + (f" (+{len(incompletas) - 1} más)" if len(incompletas) > 1 else "")
                        + "\nSin fecha de Vencimiento/Validez: esas columnas se dejarán "
                          "sin modificar en los Excel.\n")
        detalle += "\n¿Procesar ahora?"
        if not messagebox.askyesno("Procesar ofertas", detalle, parent=self):
            return

        ofertas_dir = self.var_dir_ofertas.get()
        proc.CANCELAR.clear()
        self.procesando = True
        self._cerrar_editor(False)
        self.btn_procesar.configure(text="Cancelar", fg_color=COLOR_CANCELAR,
                                    hover_color=COLOR_CANCELAR_HOVER)
        self.barra_progreso.grid()
        self.barra_progreso.start()
        self.lbl_estado.configure(text=f"Procesando {len(seleccionadas)} oferta(s)…")
        self._aviso(f"\nProcesando {len(seleccionadas)} oferta(s) hacia {salida}\n", "marca")
        self._hilo = threading.Thread(
            target=self._trabajo,
            args=([o.nombre for o in seleccionadas], fechas_archivo, filename_to_com,
                  ofertas_dir, salida), daemon=True)
        self._hilo.start()

    def _recolectar(self):
        """Prepara las fechas por archivo y la comercializadora asignada a cada uno."""
        fechas_archivo = {}
        filename_to_com = {}
        for oferta in self._seleccion():
            com = oferta.comercializadora.strip().upper()
            filename_to_com[oferta.nombre] = com
            fechas_archivo[oferta.nombre] = {
                "Recibida": oferta.recibida,
                "Vencimiento": oferta.vencimiento,
                "Validez": oferta.validez,
                "Renovable": oferta.renovable,
            }
        return fechas_archivo, filename_to_com

    def _trabajo(self, fnames, fechas_archivo, filename_to_com, ofertas_dir, salida_dir):
        import traceback
        resumen = {"mono": 0, "multi": 0, "cancelado": True, "total": len(fnames)}
        mono_dir = os.path.join(salida_dir, "MONOPUNTO")
        multi_dir = os.path.join(salida_dir, "MULTIPUNTO")
        try:
            ofertas_mp, ofertas_mt = proc.leer_ofertas(
                fnames=fnames, filename_to_com=filename_to_com, ofertas_dir=ofertas_dir)
            resumen["mono"] = len(ofertas_mp)
            resumen["multi"] = len(ofertas_mt)
            os.makedirs(mono_dir, exist_ok=True)
            os.makedirs(multi_dir, exist_ok=True)
            if ofertas_mp:
                proc.procesar_clientes(ofertas_mp, mono_dir, es_monopunto=True,
                                       tipo_label="MONOPUNTO", fechas_com=fechas_archivo)
            if ofertas_mt:
                proc.procesar_clientes(ofertas_mt, multi_dir, es_monopunto=False,
                                       tipo_label="MULTIPUNTO", fechas_com=fechas_archivo)
            resumen["cancelado"] = proc.hubo_cancelacion()
        except Exception as exc:
            proc.log(f"\nERROR: {exc}\n", "error")
            proc.log(traceback.format_exc(), "error")
            resumen["error"] = True
        finally:
            self._cola_log.put(("_fin", resumen))

    def _cancelar(self):
        proc.CANCELAR.set()
        self.lbl_estado.configure(text="Cancelando…")

    def _terminar(self, resumen):
        self.procesando = False
        self.barra_progreso.stop()
        self.barra_progreso.set(0)
        self.barra_progreso.grid_remove()
        self.btn_procesar.configure(text="Procesar", fg_color=COLOR_BOTON,
                                    hover_color=COLOR_BOTON_HOVER)
        self._cerrar_editor(False)
        self._refrescar_botones_salida()
        self._actualizar_contadores()
        if resumen.get("error"):
            self.lbl_estado.configure(text="Terminado con errores")
            return
        if resumen.get("cancelado"):
            self.lbl_estado.configure(text="Cancelado por el usuario")
            self._aviso("Proceso cancelado.\n", "warn")
            return
        self.lbl_estado.configure(
            text=f"Listo · {resumen['mono']} monopunto · {resumen['multi']} multipunto")
        self._aviso("\nProceso completado\n", "marca")
        self._aviso(f"Ofertas procesadas: {resumen.get('total', 0)}\n"
                    f"Clientes MONOPUNTO: {resumen['mono']}\n"
                    f"Clientes MULTIPUNTO: {resumen['multi']}\n"
                    f"Carpeta: {self.var_dir_salida.get()}\n", "ok")

    # -- configuracion y cierre --------------------------------------------

    def _config_pendiente(self, _evento=None):
        if self._tarea_config is None:
            try:
                self._tarea_config = self.after(700, self._guardar_config)
            except tk.TclError:
                self._tarea_config = None

    def _guardar_config(self):
        self._tarea_config = None
        config = {
            "dir_ofertas": self.var_dir_ofertas.get(),
            "dir_salida": self.var_dir_salida.get(),
            "geometry": self.geometry(),
            "apariencia": _modo(),
            "orden": list(self._orden),
        }
        if not guardar_config(config) and not self._config_advertida:
            self._config_advertida = True
            self._aviso("No se pudo guardar config.json (¿carpeta de solo lectura?). "
                        "Esta sesión sigue con las carpetas y el tema de siempre, pero no "
                        "se guardarán para el próximo arranque.\n", "warn")

    def cerrar(self):
        if self.procesando:
            if not messagebox.askyesno(
                    "Proceso en curso",
                    "Todavía se están escribiendo archivos.\n"
                    "Si cierras ahora, algún Excel puede quedar incompleto.\n\n¿Cerrar de todos modos?",
                    parent=self):
                return
            proc.CANCELAR.set()
        self._guardar_config()
        self.destroy()

    def _mostrar_ayuda(self):
        ventana = ctk.CTkToplevel(self)
        ventana.title("Ayuda")
        ventana.configure(fg_color=("gray95", "gray13"))
        ventana.resizable(False, False)
        ventana.transient(self)
        texto = ctk.CTkTextbox(ventana, width=520, height=430, corner_radius=10,
                               font=("Consolas", 12), wrap="word",
                               fg_color=("gray97", "gray10"))
        texto.pack(padx=14, pady=(14, 8))
        texto.insert("1.0", AYUDA)
        texto.configure(state="disabled")
        ctk.CTkButton(ventana, text="Cerrar", width=110, command=ventana.destroy).pack(pady=(0, 14))
        ventana.update_idletasks()
        try:
            x = self.winfo_rootx() + (self.winfo_width() - ventana.winfo_width()) // 2
            y = self.winfo_rooty() + (self.winfo_height() - ventana.winfo_height()) // 3
            ventana.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass


def main():
    activar_dpi_awareness()
    config = cargar_config()
    ctk.set_appearance_mode(config.get("apariencia") or "light")
    app = OfertasView(config)
    app.mainloop()


if __name__ == '__main__':
    main()

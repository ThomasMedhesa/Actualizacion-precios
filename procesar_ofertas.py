#!/usr/bin/env python3
"""
Procesador de Ofertas Eléctricas.
- Seleccionar ofertas → asignar fechas por oferta → procesar
Uso: python procesar_ofertas.py
"""

import os
import sys
import threading
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from tkcalendar import DateEntry
except ImportError:
    print("Instalando tkcalendar...")
    os.system("pip install tkcalendar")
    from tkcalendar import DateEntry

try:
    from openpyxl import load_workbook
except ImportError:
    os.system("pip install openpyxl")
    from openpyxl import load_workbook


def obtener_base_dir():
    """Obtiene el directorio base, compatible con PyInstaller."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = obtener_base_dir()
sys.path.insert(0, BASE_DIR)
import llenar_planillas as proc
OFERTAS_DEF = proc.OFERTAS_DIR
SALIDA_DEF = BASE_DIR

# Cargar lista de comercializadoras
COMERCIALIZADORAS = []
ruta_coms = os.path.join(BASE_DIR, "Comercializadoras.txt")
if os.path.exists(ruta_coms):
    with open(ruta_coms, 'r', encoding='utf-8') as f:
        COMERCIALIZADORAS = [line.strip() for line in f if line.strip()]


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


class OfertasView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.oferta_widgets = {}
        self.dir_ofertas = tk.StringVar(value=OFERTAS_DEF)
        self.dir_salida = tk.StringVar(value=SALIDA_DEF)
        self.crear_widgets()
        self.listar_ofertas()

    def crear_widgets(self):
        # --- Tema moderno ---
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        titulo = ttk.Label(self, text="Gestor de Ofertas Eléctricas",
                           font=('Segoe UI', 15, 'bold'))
        titulo.pack(pady=(10, 0))

        # --- Selectores de directorio ---
        dir_frame = ttk.LabelFrame(self, text="Directorios", padding=8)
        dir_frame.pack(fill='x', padx=16, pady=6)

        ttk.Label(dir_frame, text="Ofertas (origen):").grid(row=0, column=0, sticky='w', padx=(0, 6))
        self.lbl_dir_ofertas = ttk.Label(dir_frame, text=OFERTAS_DEF,
                                         foreground='#555')
        self.lbl_dir_ofertas.grid(row=0, column=1, sticky='w')
        ttk.Button(dir_frame, text="📂", width=3,
                   command=self._seleccionar_dir_ofertas).grid(row=0, column=2, padx=4)

        ttk.Label(dir_frame, text="Salida (resultados):").grid(row=1, column=0, sticky='w', padx=(0, 6), pady=(4, 0))
        self.lbl_dir_salida = ttk.Label(dir_frame, text=BASE_DIR,
                                        foreground='#555')
        self.lbl_dir_salida.grid(row=1, column=1, sticky='w', pady=(4, 0))
        ttk.Button(dir_frame, text="📂", width=3,
                   command=self._seleccionar_dir_salida).grid(row=1, column=2, padx=4, pady=(4, 0))

        # --- Panel dividido vertical: ofertas (arriba) + estado (abajo) ---
        paned = ttk.PanedWindow(self, orient='vertical')
        paned.pack(fill='both', expand=True, padx=16, pady=2)

        # --- Panel superior: tabla de ofertas ---
        frame_tabla = ttk.Frame(paned)
        paned.add(frame_tabla, weight=3)

        # Toolbar sobre la tabla (select all, deselect, refrescar)
        toolbar = ttk.Frame(frame_tabla)
        toolbar.pack(fill='x', pady=(0, 2))
        ttk.Button(toolbar, text="✓ Seleccionar todas",
                   command=self.seleccionar_todas).pack(side='left', padx=1)
        ttk.Button(toolbar, text="✗ Deseleccionar todas",
                   command=self.deseleccionar_todas).pack(side='left', padx=1)
        ttk.Button(toolbar, text="↻ Refrescar",
                   command=self.listar_ofertas).pack(side='left', padx=1)

        # Canvas con scroll para la tabla de ofertas
        self.canvas = tk.Canvas(frame_tabla, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_tabla, orient='vertical', command=self.canvas.yview)
        self.frame_ofertas = ttk.Frame(self.canvas)

        self.frame_ofertas.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.frame_ofertas, anchor='nw')
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.canvas.bind('<Configure>', self._redimensionar_canvas)

        # --- Panel inferior: estado ---
        frame_estado = ttk.LabelFrame(paned, text="Estado", padding=4)
        paned.add(frame_estado, weight=1)

        txt_frame = ttk.Frame(frame_estado)
        txt_frame.pack(fill='both', expand=True)

        self.txt_estado = tk.Text(txt_frame, height=7, wrap='word',
                                   state='disabled', font=('Consolas', 9),
                                   bg='#f5f5f5', relief='flat', borderwidth=4)
        scroll_estado = ttk.Scrollbar(txt_frame, orient='vertical',
                                       command=self.txt_estado.yview)
        self.txt_estado.configure(yscrollcommand=scroll_estado.set)
        self.txt_estado.pack(side='left', fill='both', expand=True)
        scroll_estado.pack(side='right', fill='y')

        # --- Botón Procesar ---
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', padx=16, pady=6)

        self.btn_procesar = tk.Button(
            btn_frame,
            text="  ▶  Procesar ofertas seleccionadas  ",
            command=self.procesar,
            bg='#005fb8', fg='white',
            font=('Segoe UI', 12, 'bold'),
            relief='flat', padx=20, pady=8, cursor='hand2',
            activebackground='#003d73', activeforeground='white'
        )
        self.btn_procesar.pack()

        self.escribir_estado("Selecciona un directorio de ofertas y haz clic en Procesar.\n")

    # ---- Directorios ----

    def _seleccionar_dir_ofertas(self):
        d = filedialog.askdirectory(title="Carpeta con archivos de ofertas",
                                    initialdir=self.dir_ofertas.get())
        if d:
            self.dir_ofertas.set(d)
            self.lbl_dir_ofertas.configure(text=d)
            self.listar_ofertas()

    def _seleccionar_dir_salida(self):
        d = filedialog.askdirectory(title="Carpeta donde guardar los resultados",
                                    initialdir=self.dir_salida.get())
        if d:
            self.dir_salida.set(d)
            self.lbl_dir_salida.configure(text=d)

    # ---- Canvas ----

    def _redimensionar_canvas(self, event):
        if hasattr(self, 'canvas_window') and self.canvas_window:
            self.canvas.itemconfig(self.canvas_window, width=event.width)

    # ---- Listado de ofertas ----

    def listar_ofertas(self):
        for child in self.frame_ofertas.winfo_children():
            child.destroy()
        self.oferta_widgets.clear()

        ofertas_dir = self.dir_ofertas.get()

        if not os.path.isdir(ofertas_dir):
            ttk.Label(self.frame_ofertas,
                      text="(La carpeta seleccionada no existe)",
                      foreground='gray').grid(row=0, column=0, padx=6, pady=4, sticky='w')
            return

        cols_headers = [
            (0, '', 30),
            (1, 'Archivo', 150),
            (2, 'Comercializadora', 200),
            (3, 'Renovable', 100),
            (4, 'Recibida', 130),
            (5, 'Vencimiento', 130),
            (6, 'Validez', 130),
        ]
        for col, hdr, w in cols_headers:
            self.frame_ofertas.columnconfigure(
                col, weight=(1 if col == 1 else 0), minsize=w)

        ttk.Label(self.frame_ofertas, text="", width=4).grid(row=0, column=0)
        ttk.Label(self.frame_ofertas, text="Archivo",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=1, padx=4, sticky='w')
        ttk.Label(self.frame_ofertas, text="Comercializadora",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=2, padx=4)
        ttk.Label(self.frame_ofertas, text="Renovable",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=3, padx=4)
        ttk.Label(self.frame_ofertas, text="Recibida",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=4, padx=4)
        ttk.Label(self.frame_ofertas, text="Vencimiento",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=5, padx=4)
        ttk.Label(self.frame_ofertas, text="Validez",
                  font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=6, padx=4)

        row = 1
        for fname in sorted(os.listdir(ofertas_dir)):
            if not fname.endswith('.xlsx') or fname.startswith('~$'):
                continue
            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(self.frame_ofertas, variable=var)
            cb.grid(row=row, column=0, padx=2, pady=2, sticky='w')

            lbl = ttk.Label(self.frame_ofertas, text=fname)
            lbl.grid(row=row, column=1, padx=4, pady=2, sticky='w')

            cbo = ttk.Combobox(self.frame_ofertas, values=COMERCIALIZADORAS,
                                state='readonly', width=22)
            sugerida = _buscar_comercializadora(fname)
            if sugerida in COMERCIALIZADORAS:
                cbo.set(sugerida)
            elif COMERCIALIZADORAS:
                cbo.set(COMERCIALIZADORAS[0])
            cbo.grid(row=row, column=2, padx=4, pady=2, sticky='ew')

            cbo_ren = ttk.Combobox(self.frame_ofertas, values=['No', 'Si'],
                                    state='readonly', width=6)
            cbo_ren.set('No')
            cbo_ren.grid(row=row, column=3, padx=4, pady=2)

            e_rec = DateEntry(self.frame_ofertas, date_pattern='dd/mm/yyyy',
                              width=12, background='#005fb8',
                              foreground='white', borderwidth=2)
            e_rec.set_date(date.today())
            e_rec.grid(row=row, column=4, padx=4, pady=2)

            e_ven = DateEntry(self.frame_ofertas, date_pattern='dd/mm/yyyy',
                              width=12, background='#005fb8',
                              foreground='white', borderwidth=2)
            e_ven.grid(row=row, column=5, padx=4, pady=2)

            e_val = DateEntry(self.frame_ofertas, date_pattern='dd/mm/yyyy',
                              width=12, background='#005fb8',
                              foreground='white', borderwidth=2)
            e_val.grid(row=row, column=6, padx=4, pady=2)

            self.oferta_widgets[fname] = {
                'var': var,
                'comercializadora': cbo,
                'renovable': cbo_ren,
                'recibida': e_rec,
                'vencimiento': e_ven,
                'validez': e_val,
            }
            row += 1

        if not self.oferta_widgets:
            ttk.Label(self.frame_ofertas,
                      text="(No hay archivos .xlsx en el directorio seleccionado)",
                      foreground='gray').grid(row=1, column=0, columnspan=5, padx=6, pady=4)

    # ---- Acciones ----

    def seleccionar_todas(self):
        for w in self.oferta_widgets.values():
            w['var'].set(True)

    def deseleccionar_todas(self):
        for w in self.oferta_widgets.values():
            w['var'].set(False)

    def escribir_estado(self, texto):
        self.txt_estado.configure(state='normal')
        self.txt_estado.insert('end', texto)
        self.txt_estado.see('end')
        self.txt_estado.configure(state='disabled')

    def procesar(self):
        seleccionadas = [(f, w) for f, w in self.oferta_widgets.items() if w['var'].get()]
        if not seleccionadas:
            messagebox.showwarning("Sin selección",
                                   "Selecciona al menos una oferta para procesar.")
            return

        fechas_com = {}
        filename_to_com = {}
        for fname, w in seleccionadas:
            com = w['comercializadora'].get().strip().upper()
            if not com:
                com = proc.extract_comercializadora(fname)
            fechas_com[com] = {
                'Recibida': w['recibida'].get_date(),
                'Vencimiento': w['vencimiento'].get_date(),
                'Validez': w['validez'].get_date(),
                'Renovable': w['renovable'].get(),
            }
            filename_to_com[fname] = com

        fnames = [f for f, _ in seleccionadas]
        self.btn_procesar.configure(state='disabled', text="  ⏳  Procesando...  ",
                                     bg='#888')

        self.escribir_estado(
            f"> Procesando {len(fnames)} oferta(s): {', '.join(fnames)}\n\n"
        )

        self.output_lines = []
        ofertas_dir = self.dir_ofertas.get()
        salida_dir = self.dir_salida.get()
        self.thread = threading.Thread(
            target=self._procesar_thread,
            args=(fnames, fechas_com, filename_to_com, ofertas_dir, salida_dir),
            daemon=True
        )
        self.after_id = None
        self.thread.start()
        self.poll_thread()

    def _procesar_thread(self, fnames, fechas_com, filename_to_com,
                          ofertas_dir, salida_dir):
        original_stdout = sys.stdout if sys.stdout is not None else open(os.devnull, 'w')
        class Captura:
            def __init__(self, lista, original):
                self.lista = lista
                self.original = original
            def write(self, text):
                self.original.write(text)
                self.lista.append(text)
            def flush(self):
                self.original.flush()
        sys.stdout = Captura(self.output_lines, original_stdout)

        try:
            ofertas_mp, ofertas_mt = proc.leer_ofertas(
                fnames=fnames, filename_to_com=filename_to_com,
                ofertas_dir=ofertas_dir
            )
            n_mp = len(ofertas_mp)
            n_mt = len(ofertas_mt)

            mono_dir = os.path.join(salida_dir, 'MONOPUNTO')
            multi_dir = os.path.join(salida_dir, 'MULTIPUNTO')
            os.makedirs(mono_dir, exist_ok=True)
            os.makedirs(multi_dir, exist_ok=True)

            if n_mp > 0:
                proc.procesar_clientes(ofertas_mp, mono_dir,
                                       es_monopunto=True, tipo_label="MONOPUNTO",
                                       fechas_com=fechas_com)
            if n_mt > 0:
                proc.procesar_clientes(ofertas_mt, multi_dir,
                                       es_monopunto=False, tipo_label="MULTIPUNTO",
                                       fechas_com=fechas_com)

            resumen = (
                f"\n"
                f"{'='*50}\n"
                f"  Proceso completado!\n"
                f"  Ofertas procesadas: {len(fnames)}\n"
                f"  Clientes MONOPUNTO: {n_mp}\n"
                f"  Clientes MULTIPUNTO: {n_mt}\n"
                f"{'='*50}\n"
            )
            self.output_lines.append(resumen)

            if n_mp > 0:
                self.output_lines.append(f"\nSalida: {mono_dir}\n")
                self.output_lines.append("MONOPUNTO:\n")
                for k, v in sorted(ofertas_mp.items()):
                    coms = list(v['_comercializadoras'].keys())
                    self.output_lines.append(
                        f"  {v['_nombre']} -> {coms}\n"
                    )
            if n_mt > 0:
                self.output_lines.append(f"\nSalida: {multi_dir}\n")
                self.output_lines.append("MULTIPUNTO:\n")
                for k, v in sorted(ofertas_mt.items()):
                    coms = list(v['_comercializadoras'].keys())
                    self.output_lines.append(
                        f"  {v['_nombre']} -> {coms}\n"
                    )

        except Exception as e:
            self.output_lines.append(f"\nERROR: {e}\n")
            import traceback
            self.output_lines.append(traceback.format_exc())
        finally:
            sys.stdout = original_stdout

    def poll_thread(self):
        if self.thread is None or not self.thread.is_alive():
            for line in self.output_lines:
                self.escribir_estado(line)
            self.output_lines.clear()
            self.btn_procesar.configure(state='normal',
                                         text="  ▶  Procesar ofertas seleccionadas  ",
                                         bg='#005fb8')
            return

        for line in self.output_lines:
            self.escribir_estado(line)
        self.output_lines.clear()
        self.after_id = self.after(100, self.poll_thread)


def main():
    root = tk.Tk()
    root.title("Gestor de Ofertas Eléctricas")
    root.geometry("950x700")
    root.minsize(750, 550)
    app = OfertasView(root)
    app.pack(fill='both', expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()

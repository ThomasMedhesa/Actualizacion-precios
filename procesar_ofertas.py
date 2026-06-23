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
from tkinter import ttk, messagebox

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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import llenar_planillas as proc

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MONOPUNTO_DIR = proc.MONOPUNTO_DIR
MULTIPUNTO_DIR = proc.MULTIPUNTO_DIR
OFERTAS_DIR = proc.OFERTAS_DIR

# Cargar lista de comercializadoras
COMERCIALIZADORAS = []
ruta_coms = os.path.join(BASE_DIR, "Comercializadoras.txt")
if os.path.exists(ruta_coms):
    with open(ruta_coms, 'r', encoding='utf-8') as f:
        COMERCIALIZADORAS = [line.strip() for line in f if line.strip()]


def _normalizar_com(s):
    return proc.normalize(s) if hasattr(proc, 'normalize') else s.upper().strip()


def _buscar_comercializadora(nombre_archivo):
    """Busca la mejor coincidencia en la lista de comercializadoras
    a partir del nombre extraído del archivo."""
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
        self.crear_widgets()
        self.listar_ofertas()

    def crear_widgets(self):
        ttk.Label(self, text="Procesador de Ofertas Eléctricas",
                  font=('Segoe UI', 14, 'bold')).pack(pady=(8, 2))
        ttk.Label(self, text="Selecciona ofertas y asigna fechas para cada una:",
                  font=('Segoe UI', 10)).pack(anchor='w', padx=16, pady=(0, 4))

        # --- Tabla de ofertas con scroll ---
        frame_tabla = ttk.Frame(self)
        frame_tabla.pack(fill='both', expand=True, padx=16, pady=2)

        canvas = tk.Canvas(frame_tabla, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_tabla, orient='vertical', command=canvas.yview)
        self.frame_ofertas = ttk.Frame(canvas)

        self.frame_ofertas.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self.canvas_window = canvas.create_window(
            (0, 0), window=self.frame_ofertas, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Que el frame interior se expanda al ancho del canvas
        canvas.bind('<Configure>', self._redimensionar_canvas)

        # --- Botones ---
        btn_sel_frame = ttk.Frame(self)
        btn_sel_frame.pack(fill='x', padx=16, pady=2)
        ttk.Button(btn_sel_frame, text="Seleccionar todas",
                   command=self.seleccionar_todas).pack(side='left', padx=2)
        ttk.Button(btn_sel_frame, text="Deseleccionar todas",
                   command=self.deseleccionar_todas).pack(side='left', padx=2)
        ttk.Button(btn_sel_frame, text="Hoy para todas",
                   command=self.hoy_para_todas).pack(side='left', padx=2)

        self.btn_procesar = ttk.Button(self, text="Procesar ofertas seleccionadas",
                                       command=self.procesar)
        self.btn_procesar.pack(pady=4)

        # --- Estado ---
        frame_estado = ttk.LabelFrame(self, text="Estado", padding=4)
        frame_estado.pack(fill='both', expand=False, padx=16, pady=(0, 6))
        self.txt_estado = tk.Text(frame_estado, height=7, width=80,
                                  wrap='word', state='disabled',
                                  font=('Consolas', 9))
        scroll_estado = ttk.Scrollbar(frame_estado, orient='vertical',
                                       command=self.txt_estado.yview)
        self.txt_estado.configure(yscrollcommand=scroll_estado.set)
        self.txt_estado.pack(side='left', fill='both', expand=True)
        scroll_estado.pack(side='right', fill='y')

    def _redimensionar_canvas(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def listar_ofertas(self):
        for child in self.frame_ofertas.winfo_children():
            child.destroy()
        self.oferta_widgets.clear()

        if not os.path.isdir(OFERTAS_DIR):
            ttk.Label(self.frame_ofertas, text="(No existe la carpeta Ofertas/)",
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
        for fname in sorted(os.listdir(OFERTAS_DIR)):
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

            e_rec = DateEntry(self.frame_ofertas, date_pattern='yyyy-mm-dd',
                              width=12, background='darkblue',
                              foreground='white', borderwidth=2)
            e_rec.set_date(date.today())
            e_rec.grid(row=row, column=4, padx=4, pady=2)

            e_ven = DateEntry(self.frame_ofertas, date_pattern='yyyy-mm-dd',
                              width=12, background='darkblue',
                              foreground='white', borderwidth=2)
            e_ven.grid(row=row, column=5, padx=4, pady=2)

            e_val = DateEntry(self.frame_ofertas, date_pattern='yyyy-mm-dd',
                              width=12, background='darkblue',
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
            ttk.Label(self.frame_ofertas, text="(No hay archivos .xlsx en Ofertas/)",
                      foreground='gray').grid(row=1, column=0, columnspan=5, padx=6, pady=4)

    def seleccionar_todas(self):
        for w in self.oferta_widgets.values():
            w['var'].set(True)

    def deseleccionar_todas(self):
        for w in self.oferta_widgets.values():
            w['var'].set(False)

    def hoy_para_todas(self):
        for w in self.oferta_widgets.values():
            w['recibida'].set_date(date.today())

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

        # Build fechas_com dict y filename_to_com
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
        self.btn_procesar.configure(state='disabled', text="Procesando...")

        self.escribir_estado(
            f"> Procesando {len(fnames)} oferta(s): {', '.join(fnames)}\n\n"
        )

        self.output_lines = []
        self.thread = threading.Thread(
            target=self._procesar_thread,
            args=(fnames, fechas_com, filename_to_com),
            daemon=True
        )
        self.after_id = None
        self.thread.start()
        self.poll_thread()

    def _procesar_thread(self, fnames, fechas_com, filename_to_com):
        original_stdout = sys.stdout
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
            ofertas_mp, ofertas_mt = proc.leer_ofertas(fnames=fnames, filename_to_com=filename_to_com)
            n_mp = len(ofertas_mp)
            n_mt = len(ofertas_mt)

            if n_mp > 0:
                proc.procesar_clientes(ofertas_mp, MONOPUNTO_DIR,
                                       es_monopunto=True, tipo_label="MONOPUNTO",
                                       fechas_com=fechas_com)
            if n_mt > 0:
                proc.procesar_clientes(ofertas_mt, MULTIPUNTO_DIR,
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
                self.output_lines.append("\nMONOPUNTO:\n")
                for k, v in sorted(ofertas_mp.items()):
                    coms = list(v['_comercializadoras'].keys())
                    self.output_lines.append(
                        f"  {v['_nombre']} -> {coms}\n"
                    )
            if n_mt > 0:
                self.output_lines.append("\nMULTIPUNTO:\n")
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
            self.btn_procesar.configure(state='normal', text="Procesar ofertas seleccionadas")
            return

        for line in self.output_lines:
            self.escribir_estado(line)
        self.output_lines.clear()
        self.after_id = self.after(100, self.poll_thread)


def main():
    root = tk.Tk()
    root.title("Gestor de Ofertas Eléctricas")
    root.geometry("900x650")
    root.minsize(700, 500)
    app = OfertasView(root)
    app.pack(fill='both', expand=True)
    root.mainloop()


if __name__ == '__main__':
    main()

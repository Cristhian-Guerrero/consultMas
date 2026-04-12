"""
Interfaz gráfica — ConsultaRUTApp
A.S. Contadores & Asesores SAS
"""

import os
import sys
import time
import threading
import webbrowser
from datetime import timedelta, datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk

from config import obtener_mensaje_profesional, obtener_tip_contable, TimeoutConfig
from core.scraper import consultar_nit, calcular_dv
from core.browser import cerrar_todos, limpiar_inactivos
from core.excel import apply_formatting


class ConsultaRUTApp(tk.Tk):

    VERSION = "V4.7 CG"

    COLORS = {
        'primary':        "#166534",
        'primary_light':  "#DCFCE7",
        'primary_dark':   '#14532D',
        'secondary':      '#64748B',
        'accent':         "#0F766E",
        'background':     '#F8FAFC',
        'surface':        '#FFFFFF',
        'text_primary':   '#1E293B',
        'text_secondary': '#64748B',
        'text_light':     '#FFFFFF',
        'success':        "#15803D",
        'warning':        '#B45309',
        'error':          '#B91C1C',
        'border':         '#E2E8F0',
    }

    FONTS = {
        'title':    ('Segoe UI', 20, 'bold'),
        'subtitle': ('Segoe UI', 11, 'bold'),
        'body':     ('Segoe UI', 10),
        'button':   ('Segoe UI', 9, 'bold'),
        'small':    ('Segoe UI', 9),
    }

    def __init__(self):
        super().__init__()
        self.title("Consulta Gestión Masiva DIAN | A.S. Contadores & Asesores SAS")
        self.ejecucion_activa = threading.Event()
        self.detener_proceso  = False
        self.tipo_consulta    = tk.StringVar(value="basica")
        self.conservar_duplicados = tk.BooleanVar(value=False)

        self.lista_nits     = []
        self.generated_file = None
        self.rows_for_excel = []
        self.tiempo_inicio  = None
        self.cronometro_activo = False

        self.geometry("1000x650")
        self.minsize(950, 630)
        self.resizable(True, False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.configure(bg=self.COLORS['background'])

        self._setup_icon()
        self._center()
        self._setup_styles()
        self._build_ui()

    # ═══ Utilidades ═══

    def resource_path(self, rel):
        try:
            base = sys._MEIPASS
        except AttributeError:
            base = os.path.abspath(".")
        return os.path.join(base, rel)

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f'{w}x{h}+{x}+{y}')

    def _setup_icon(self):
        try:
            ico = os.path.abspath(self.resource_path("dian.ico"))
            png = os.path.abspath(self.resource_path("logo.png"))
            if os.path.exists(ico):
                self.iconbitmap(default=ico)
                return
            if os.path.exists(png):
                img = Image.open(png)
                photo = ImageTk.PhotoImage(img)
                self.iconphoto(True, photo)
                self._icon_ref = photo
        except Exception as e:
            print(f"Icono no cargado: {e}")

    def _load_logo(self, max_w=300, max_h=80):
        try:
            path = self.resource_path("logo.png")
            if not os.path.exists(path):
                return None
            img = Image.open(path)
            ratio = min(max_w / img.width, max_h / img.height)
            if ratio < 1:
                img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Logo no cargado: {e}")
            return None

    # ═══ Estilos ═══

    def _setup_styles(self):
        C = self.COLORS
        F = self.FONTS
        s = ttk.Style()
        s.theme_use('clam')
        self.configure(bg=C['background'])

        s.configure('Header.TFrame',   background=C['surface'],     relief='flat')
        s.configure('Main.TFrame',     background=C['background'],  relief='flat')
        s.configure('Card.TLabelframe',background=C['surface'],     relief='solid', borderwidth=1, bordercolor=C['border'])
        s.configure('Card.TLabelframe.Label', background=C['surface'], foreground=C['primary'], font=F['subtitle'])

        s.configure('HeaderTitle.TLabel', background=C['surface'], foreground=C['primary'],        font=F['title'])
        s.configure('HeaderSub.TLabel',   background=C['surface'], foreground=C['text_secondary'], font=F['body'])
        s.configure('Body.TLabel',        background=C['surface'], foreground=C['text_primary'],   font=F['body'])
        s.configure('Status.TLabel',      background=C['surface'], foreground=C['text_secondary'], font=F['small'])

        s.configure('Consulta.TRadiobutton', background=C['surface'], foreground=C['text_primary'],
                    font=F['body'], focuscolor=C['surface'], indicatorcolor=C['surface'])
        s.map('Consulta.TRadiobutton',
              indicatorcolor=[('selected', C['primary'])],
              foreground=[('selected', C['primary'])])

        s.configure('Primary.TButton',   background=C['primary'],   foreground=C['text_light'], font=F['button'], borderwidth=0, focuscolor='none')
        s.map('Primary.TButton',   background=[('active', C['primary_dark']), ('pressed', C['primary_dark'])])
        s.configure('Secondary.TButton', background=C['secondary'], foreground=C['text_light'], font=F['button'], borderwidth=0, focuscolor='none')
        s.map('Secondary.TButton', background=[('active', '#475569')])
        s.configure('Success.TButton',   background=C['success'],   foreground=C['text_light'], font=F['button'], borderwidth=0)

        s.configure('Compact.Horizontal.TProgressbar',
                    background=C['primary'], troughcolor=C['border'], thickness=15, borderwidth=0)
        s.configure('Vertical.TScrollbar',
                    background='#CBD5E1', troughcolor='#F8FAFC', bordercolor='#F8FAFC',
                    arrowcolor=C['primary'], relief='flat')
        s.map('Vertical.TScrollbar',
              background=[('active', C['primary']), ('pressed', C['primary_dark'])],
              arrowcolor=[('active', C['primary_dark'])])

    # ═══ UI ═══

    def _build_ui(self):
        self._build_header()
        self._build_main()

    def _build_header(self):
        C = self.COLORS
        header = ttk.Frame(self, style='Header.TFrame', padding="25 18")
        header.pack(fill=tk.X)
        header.columnconfigure(1, weight=1)

        logo_container = ttk.Frame(header, style='Header.TFrame')
        logo_container.grid(row=0, column=0, sticky="nsw", padx=(0, 30), pady=8)
        logo = self._load_logo(max_w=350, max_h=90)
        if logo:
            self._logo = logo
            ttk.Label(ttk.Frame(logo_container, style='Header.TFrame', padding=0),
                      image=logo, style='Body.TLabel').pack()
            ttk.Frame(logo_container, style='Header.TFrame', padding=0).pack()
            lf = ttk.Frame(logo_container, style='Header.TFrame', padding=0)
            lf.pack()
            ttk.Label(lf, image=logo, style='Body.TLabel').pack()
        else:
            ttk.Label(logo_container, text="🏢 A.S. CONTADORES &\nASESORES SAS",
                      style='HeaderTitle.TLabel', justify=tk.CENTER).pack()

        info = ttk.Frame(header, style='Header.TFrame')
        info.grid(row=0, column=1, sticky="ew", pady=10)
        ttk.Label(info, text="Gestión Masiva DIAN",              style='HeaderTitle.TLabel').pack(anchor="w", pady=(0, 4))
        ttk.Label(info, text="Sistema automatizado de consulta Express", style='HeaderSub.TLabel').pack(anchor="w", pady=(0, 3))
        ttk.Label(info, text="Asesoría Contable • Financiera • Tributaria • Revisoría Fiscal",
                  font=self.FONTS['small'], background=C['surface'], foreground=C['accent']).pack(anchor="w")

    def _build_main(self):
        C = self.COLORS
        main = ttk.Frame(self, style='Main.TFrame', padding="15")
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)

        # ── Selector de modo ──
        sel_frame = ttk.LabelFrame(main, text="🎯 Modo de Consulta", style='Card.TLabelframe', padding="12")
        sel_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        radio_cont = ttk.Frame(sel_frame, style='Main.TFrame')
        radio_cont.pack(fill=tk.X)
        ttk.Radiobutton(radio_cont, text="📋 Express", variable=self.tipo_consulta,
                        value="basica", style='Consulta.TRadiobutton',
                        command=self._on_tipo_changed).pack(side=tk.LEFT, padx=(0, 30))

        # ── Controles ──
        ctrl_frame = ttk.LabelFrame(main, text="🔧 Controles", style='Card.TLabelframe', padding="12")
        ctrl_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        btns = ttk.Frame(ctrl_frame, style='Main.TFrame')
        btns.pack(fill=tk.X)

        ttk.Button(btns, text="📁 Cargar Excel/CSV",
                   command=self._on_cargar, style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 10))

        self.btn_detener = ttk.Button(btns, text="⏹️ Detener",
                                      command=self._detener, style='Secondary.TButton', state=tk.DISABLED)
        self.btn_detener.pack(side=tk.LEFT, padx=(0, 10))

        # Separador visual
        tk.Frame(btns, bg=C['border'], width=1, height=28).pack(side=tk.LEFT, fill=tk.Y, padx=(5, 10))

        # Toggle duplicados
        self.btn_duplicados = tk.Button(
            btns, text="◈  Duplicados: OFF",
            font=('Segoe UI', 9), relief='flat', cursor='hand2',
            bg=C['border'], fg=C['text_secondary'], padx=12, pady=5, bd=0,
            activebackground=C['border'], command=self._on_toggle_duplicados
        )
        self.btn_duplicados.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_open_excel = ttk.Button(btns, text="📊 Abrir Excel",
                                         command=self._open_excel, style='Success.TButton')

        # ── Progreso ──
        prog_frame = ttk.LabelFrame(main, text="📊 Progreso", style='Card.TLabelframe', padding="12")
        prog_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 5))

        self.lbl_status = ttk.Label(prog_frame, text="Listo para comenzar", style='Body.TLabel')
        self.lbl_status.pack(anchor="w", pady=(0, 8))

        self.lbl_tipo_activo = ttk.Label(prog_frame, text="📋 Consulta Express seleccionada", style='Status.TLabel')
        self.lbl_tipo_activo.pack(anchor="w", pady=(0, 8))

        self.progress_bar = ttk.Progressbar(prog_frame, style='Compact.Horizontal.TProgressbar', mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        self.lbl_cronometro = ttk.Label(prog_frame, text="⏱️ Tiempo: 00:00:00", style='Status.TLabel')
        self.lbl_cronometro.pack(anchor="w", pady=(0, 8))

        stats = ttk.Frame(prog_frame, style='Main.TFrame')
        stats.pack(fill=tk.X, pady=(8, 0))
        self.lbl_total    = ttk.Label(stats, text="Total: 0",      style='Status.TLabel')
        self.lbl_total.pack(anchor="w")
        self.lbl_exitosos = ttk.Label(stats, text="✅ Exitosos: 0", style='Status.TLabel')
        self.lbl_exitosos.pack(anchor="w")
        self.lbl_errores  = ttk.Label(stats, text="❌ Errores: 0",  style='Status.TLabel')
        self.lbl_errores.pack(anchor="w")

        # ── Resultados ──
        res_frame = ttk.LabelFrame(main, text="📋 Resultados del Proceso", style='Card.TLabelframe', padding="1")
        res_frame.grid(row=2, column=1, sticky="nsew", padx=(5, 0))
        txt_cont = ttk.Frame(res_frame, style='Main.TFrame', padding=1)
        txt_cont.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar = ttk.Scrollbar(txt_cont)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_resultados = tk.Text(txt_cont, wrap=tk.WORD, height=9,
                                      yscrollcommand=scrollbar.set,
                                      bg='#F8FAFC', fg=C['text_primary'],
                                      font=('Segoe UI', 9), padx=10, pady=10,
                                      relief='flat', borderwidth=0)
        self.txt_resultados.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.txt_resultados.yview)
        self._log(
            "🏢 A.S. CONTADORES & ASESORES SAS\n"
            "Sistema de Consulta Automatizada DIAN\n\n"
            "✨ 1. Seleccione tipo de consulta\n"
            "📁 2. Cargar archivo Excel/CSV (columna 'NIT')\n"
            "🚀 3. Sistema iniciará procesamiento automático\n\n"
            "Esperando configuración...\n"
        )

        # ── Footer ──
        footer = ttk.Frame(self, style='Main.TFrame', padding="8 5")
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(footer, text=f"© 2026 A.S. Contadores & Asesores SAS | {self.VERSION}",
                  style='Status.TLabel').pack()

    # ═══ Callbacks ═══

    def _on_tipo_changed(self):
        tipo = self.tipo_consulta.get()
        if tipo == "basica":
            self.lbl_tipo_activo.config(text="📋 Consulta Express seleccionada (Rápida)")
            self._add_msg("📋 Modo Express seleccionado", "INFO")
        else:
            self.lbl_tipo_activo.config(text="🔍 Consulta Detallada RUT seleccionada (Completa)")
            self._add_msg("🔍 Modo RUT Detallado seleccionado", "INFO")

    def _on_toggle_duplicados(self):
        self.conservar_duplicados.set(not self.conservar_duplicados.get())
        C = self.COLORS
        if self.conservar_duplicados.get():
            self.btn_duplicados.config(text="◈  Duplicados: ON",
                                       bg=C['warning'], fg='white', activebackground='#92400E')
            self._add_msg("◈ Duplicados activados — se procesará cada NIT por separado", "WARNING")
        else:
            self.btn_duplicados.config(text="◈  Duplicados: OFF",
                                       bg=C['border'], fg=C['text_secondary'], activebackground=C['border'])
            self._add_msg("◈ Duplicados desactivados — se eliminarán NITs repetidos", "INFO")

    def _on_close(self):
        cerrar_todos()
        self.destroy()

    # ═══ Log y progreso ═══

    def _log(self, text):
        self.txt_resultados.config(state=tk.NORMAL)
        self.txt_resultados.insert(tk.END, text)
        self.txt_resultados.see(tk.END)
        self.txt_resultados.config(state=tk.DISABLED)
        self.update()

    def _add_msg(self, message, level="INFO"):
        prefix = {"SUCCESS": "✅", "ERROR": "❌", "WARNING": "⚠️"}.get(level, "ℹ️")
        self._log(f"{prefix} {message}\n")

    def _update_stats(self, total, exitosos, errores):
        self.lbl_total.config(text=f"Total: {total}")
        self.lbl_exitosos.config(text=f"✅ Exitosos: {exitosos}")
        self.lbl_errores.config(text=f"❌ Errores: {errores}")

    def _update_progress(self, count, total, nit=None):
        self.progress_bar["value"] = count
        tipo_txt = "Express" if self.tipo_consulta.get() == "basica" else "RUT"
        if nit:
            self.lbl_status.config(text=f"Consultando {tipo_txt}: {nit} ({count}/{total})")
        else:
            self.lbl_status.config(text=f"Procesando: {count}/{total}")

    def _iniciar_cronometro(self):
        self.tiempo_inicio = time.time()
        self.cronometro_activo = True
        self._tick_cronometro()

    def _detener_cronometro(self):
        self.cronometro_activo = False

    def _tick_cronometro(self):
        if self.cronometro_activo and self.tiempo_inicio:
            elapsed = time.time() - self.tiempo_inicio
            self.lbl_cronometro.config(text=f"⏱️ Tiempo: {str(timedelta(seconds=int(elapsed)))}")
            self.after(1000, self._tick_cronometro)

    # ═══ Carga de archivo ═══

    def _reset(self):
        cerrar_todos()
        self.lista_nits = []
        self.rows_for_excel = []
        self.progress_bar["value"] = 0
        self.lbl_status.config(text="Listo para comenzar")
        self._update_stats(0, 0, 0)
        self.txt_resultados.config(state=tk.NORMAL)
        self.txt_resultados.delete("1.0", tk.END)
        self._log("🔄 Sistema reiniciado\n📁 Seleccione nuevo archivo\n\n")
        self.generated_file = None
        self.btn_open_excel.pack_forget()
        self.btn_detener.config(state=tk.DISABLED)
        self._detener_cronometro()

    def _on_cargar(self):
        self._reset()
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel o CSV",
            initialdir=os.path.expanduser("~/Documents"),
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Archivos CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not filepath:
            return

        self._add_msg(f"Archivo: {os.path.basename(filepath)}", "INFO")

        try:
            df = pd.read_csv(filepath) if filepath.endswith('.csv') else pd.read_excel(filepath)
            self._add_msg(f"Leído: {len(df)} filas", "SUCCESS")
        except Exception as e:
            self._add_msg(f"Error: {str(e)}", "ERROR")
            return

        if "NIT" not in df.columns:
            self._add_msg("Falta columna 'NIT'", "ERROR")
            return

        try:
            nits_raw = df["NIT"].dropna().astype(str).tolist()
            nits_procesados = []
            for nit in nits_raw:
                if '.' in nit:
                    nit = str(int(float(nit)))
                nit_limpio = ''.join(c for c in nit if c.isdigit())
                if nit_limpio:
                    nits_procesados.append(nit_limpio)

            nits_unicos = list(dict.fromkeys(nits_procesados))
            duplicados  = len(nits_procesados) - len(nits_unicos)

            if self.conservar_duplicados.get():
                self.lista_nits = nits_procesados
                if duplicados > 0:
                    self._add_msg(f"◈ {duplicados} duplicados conservados — se procesarán individualmente", "WARNING")
            else:
                self.lista_nits = nits_unicos
                if duplicados > 0:
                    self._add_msg(f"◈ {duplicados} duplicados eliminados", "INFO")

            if not self.lista_nits:
                self._add_msg("No hay NITs válidos", "ERROR")
                return

            total    = len(self.lista_nits)
            tipo_txt = "Express" if self.tipo_consulta.get() == "basica" else "Detallada RUT"
            self.progress_bar["maximum"] = total
            self._update_stats(total, 0, 0)
            self.lbl_status.config(text=f"Iniciando consulta {tipo_txt}: {total} NITs")
            self._add_msg(f"Procesando {total} NITs — Tipo: {tipo_txt}", "SUCCESS")
            self._add_msg("═══ INICIANDO PROCESAMIENTO ═══", "INFO")
            self._iniciar_cronometro()
            threading.Thread(target=self._procesar, daemon=True).start()

        except Exception as e:
            self._add_msg(f"Error: {str(e)}", "ERROR")

    # ═══ Proceso principal ═══

    def _detener(self):
        if self.ejecucion_activa.is_set():
            self.detener_proceso = True
            self.ejecucion_activa.clear()
            self.btn_detener.config(state=tk.DISABLED)
            self.lbl_status.config(text="Deteniendo...")
            self._add_msg("Proceso detenido", "WARNING")
            self.update()
            self._generar_excel("PARCIAL")
        else:
            messagebox.showinfo("Info", "No hay proceso activo")

    def _procesar(self):
        self.ejecucion_activa.set()
        self.detener_proceso = False
        self.btn_detener.config(state=tk.NORMAL)
        self.rows_for_excel = []

        try:
            if not self.lista_nits:
                self._add_msg("No hay NITs", "ERROR")
                return

            total            = len(self.lista_nits)
            exitosos, errores = 0, 0
            tipo             = self.tipo_consulta.get()
            tipo_txt         = "Express" if tipo == "basica" else "Detallada RUT (Completa)"

            self._add_msg(f"🚀 Procesando {total} NITs — Modo: {tipo_txt}", "SUCCESS")

            for i, nit in enumerate(self.lista_nits, 1):
                if not self.ejecucion_activa.is_set() or self.detener_proceso:
                    self._add_msg(f"Proceso interrumpido en {i}/{total}", "WARNING")
                    break

                if i % 8 == 0:
                    self._add_msg(obtener_mensaje_profesional(), "INFO")
                    limpiar_inactivos()
                if i % 4 == 1 and i > 1:
                    self._add_msg(obtener_tip_contable(), "INFO")

                self._update_progress(i, total, nit)

                resultado = consultar_nit(nit, tipo, attempt=1)
                if resultado.get("status") == "retry":
                    self._add_msg(f"🔄 Reintentando NIT {nit}", "WARNING")
                    time.sleep(1.0 if tipo == "basica" else 2.5)
                    resultado = consultar_nit(nit, tipo, attempt=2)
                if resultado.get("status") == "retry":
                    time.sleep(1.5 if tipo == "basica" else 3.5)
                    resultado = consultar_nit(nit, tipo, attempt=3)

                if resultado.get("status") == "success":
                    data = resultado["data"]

                    def limpiar(v):
                        if not v or str(v).strip() in ("", "None", "Sin inconsistencias registradas"):
                            return "-"
                        return str(v).strip()

                    if tipo == "basica":
                        fila = {
                            "NIT":              nit,
                            "DV":               data.get("dv", calcular_dv(nit)),
                            "Primer Apellido":  limpiar(data.get("primerNombre")),
                            "Segundo Apellido": limpiar(data.get("otrosNombres")),
                            "Primer Nombre":    limpiar(data.get("primerApellido")),
                            "Otros Nombres":    limpiar(data.get("segundoApellido")),
                            "Razón Social":     limpiar(data.get("razonSocial")),
                            "Fecha Consulta":   data.get("datetime", ""),
                            "Estado Consulta":  "Exitoso",
                            "Tipo de Consulta": "Express",
                            "Observaciones":    data.get("observacion", "Consulta Express exitosa"),
                        }
                    else:
                        fila = {
                            "NIT":                 nit,
                            "DV":                  limpiar(data.get("dv", calcular_dv(nit))),
                            "Primer Apellido":      limpiar(data.get("primerNombre")),
                            "Segundo Apellido":     limpiar(data.get("otrosNombres")),
                            "Primer Nombre":        limpiar(data.get("primerApellido")),
                            "Otros Nombres":        limpiar(data.get("segundoApellido")),
                            "Razón Social":         limpiar(data.get("razonSocial")),
                            "Estado del Registro":  limpiar(data.get("estado", "SIN INFORMACIÓN")),
                            "Fecha Consulta":       data.get("datetime", ""),
                            "Estado Consulta":      "Exitoso",
                            "Tipo de Consulta":     "RUT Detallado",
                            "Observaciones":        data.get("observacion", "Consulta RUT detallada exitosa"),
                        }

                    self.rows_for_excel.append(fila)
                    razon = data.get('razonSocial', '')
                    nombre = f"{data.get('primerNombre','')} {data.get('primerApellido','')}".strip()
                    display = razon if razon and razon != "-" else (nombre if nombre else "Sin datos")
                    if len(display) > 30:
                        display = display[:30] + "..."
                    self._add_msg(f"{nit}: {display}", "SUCCESS")
                    exitosos += 1

                else:
                    error_msg = resultado.get("error", "Error")
                    base_fila = {
                        "NIT": nit, "DV": "-", "Primer Apellido": "-", "Segundo Apellido": "-",
                        "Primer Nombre": "-", "Otros Nombres": "-", "Razón Social": "-",
                        "Fecha Consulta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Estado Consulta": "No Inscrito", "Tipo de Consulta": "Express" if tipo == "basica" else "RUT Detallado",
                        "Observaciones": "No Inscrito",
                    }
                    if tipo == "rut_detallado":
                        base_fila["Estado del Registro"] = "-"
                    self.rows_for_excel.append(base_fila)
                    self._add_msg(f"{nit}: Requiere verificación — {error_msg[:25]}", "ERROR")
                    errores += 1

                self._update_stats(total, exitosos, errores)
                if i < total:
                    time.sleep(TimeoutConfig.BROWSER_REST_TIME + (0.3 if tipo == "rut_detallado" else 0))

            self._add_msg("═══ PROCESO COMPLETADO ═══", "INFO")
            self._add_msg(f"📊 Total: {i} | ✅ Exitosas: {exitosos} | ⚠️ Revisión: {errores}", "INFO")

            if self.ejecucion_activa.is_set() and not self.detener_proceso:
                self._add_msg("📈 Generando reporte...", "SUCCESS")
                self._generar_excel("COMPLETO")
                self.after(1000, self._open_excel)

        except Exception as e:
            self._add_msg(f"Error en el sistema: {str(e)}", "ERROR")
        finally:
            self.ejecucion_activa.clear()
            self.detener_proceso = False
            self.btn_detener.config(state=tk.DISABLED)
            self._detener_cronometro()
            cerrar_todos()
            self.lbl_status.config(text="Proceso completado")

    # ═══ Excel ═══

    def _generar_excel(self, tipo="COMPLETO"):
        try:
            if not self.rows_for_excel:
                self._add_msg("No hay datos para Excel", "WARNING")
                return

            df_out = pd.DataFrame(self.rows_for_excel).fillna("-")

            if 'NIT' in df_out.columns:
                df_out['NIT'] = pd.to_numeric(df_out['NIT'], errors='coerce').fillna(0).astype('int64')
            if 'DV' in df_out.columns:
                df_out['DV'] = pd.to_numeric(df_out['DV'],  errors='coerce').fillna(0).astype('int64')

            tipo_consulta = self.tipo_consulta.get()
            sufijo = "express" if tipo_consulta == "basica" else "rut_detallado"
            prefijo = "resultado_completo" if tipo == "COMPLETO" else "resultado_parcial"
            fname = f"{prefijo}_{sufijo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

            with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Resultados', startrow=2)
                apply_formatting(writer.book, writer.sheets['Resultados'], df_out, tipo, tipo_consulta)

            self.generated_file = fname
            self._add_msg(f"Excel generado: {fname}", "SUCCESS")
            self._show_excel_btn()

        except Exception as e:
            self._add_msg(f"Error Excel: {str(e)}", "ERROR")
        finally:
            if tipo == "PARCIAL":
                self.lbl_status.config(text="Detenido")
                cerrar_todos()

    def _show_excel_btn(self):
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 10))
        self.update()

    def _open_excel(self):
        if self.generated_file and os.path.exists(self.generated_file):
            try:
                path = os.path.abspath(self.generated_file)
                if sys.platform.startswith('win'):
                    os.startfile(path)
                elif sys.platform.startswith('darwin'):
                    os.system(f'open "{path}"')
                else:
                    webbrowser.open(f"file:///{path}")
                self._add_msg(f"Abriendo: {self.generated_file}", "SUCCESS")
            except Exception as e:
                self._add_msg(f"Error abriendo: {str(e)}", "ERROR")
        else:
            self._add_msg("Archivo no encontrado", "ERROR")

import sys
import os

# FIX: Patch stdout/stderr for --noconsole mode (PyInstaller)
# Some libraries (numpy, catboost) try to write to stdout on import, causing crash if it's None.
class DummyWriter:
    def write(self, message): pass
    def flush(self): pass

if sys.stdout is None: sys.stdout = DummyWriter()
if sys.stderr is None: sys.stderr = DummyWriter()

import customtkinter as ctk
import threading
import pandas as pd
import webbrowser
import matplotlib.pyplot as plt
import pythoncom
import win32timezone # Necessary for Outlook Datetime parsing
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from catboost import CatBoostClassifier
from sklearn.base import BaseEstimator, ClassifierMixin
from collections import Counter
import re

# --- CONFIGURACIÓN GLOBAL ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green") # Cambiamos a Green para un look más "Matrix/Data"

# Colores Corporativos
COLOR_BG = "#121212"
COLOR_CARD = "#1E1E1E"
COLOR_ACCENT = "#00E676" # Verde Neón Brillante
COLOR_TEXT_MAIN = "#FFFFFF"
COLOR_TEXT_DIM = "#A0A0A0"

# Fuentes
FONT_MAIN = ("Segoe UI", 12)
FONT_HEADER = ("Segoe UI", 24, "bold")
FONT_KPI_VAL = ("Segoe UI", 36, "bold")
FONT_KPI_TITLE = ("Segoe UI", 11, "bold")

# --- IMPORTACIÓN DINÁMICA DE MÓDULOS ---
try:
    import importlib
    extractor = importlib.import_module("01_data_extractor")
    trainer = importlib.import_module("02_model_trainer")
    inference = importlib.import_module("03_inference_engine")
except ImportError as e:
    # Mocking
    class MockModule:
        MI_NOMBRE_MOSTRAR = ""
        MI_EMAIL_CORPORATIVO = ""
        DIAS_HISTORIAL = 0
        def generar_dataset_masivo(self): pass
        def entrenar_modelo_definitivo(self): pass
        def ejecutar_vigilancia(self): pass
    extractor = MockModule()
    trainer = MockModule()
    inference = MockModule()

# --- CLASE WRAPPER (Necesaria para Joblib) ---
class CatBoostWrapper(BaseEstimator, ClassifierMixin):
    def __init__(self, **kwargs):
        self.model = CatBoostClassifier(**kwargs)
    def fit(self, X, y):
        self.model.fit(X, y)
        self.classes_ = self.model.classes_
        return self
    def predict(self, X):
        return self.model.predict(X)
    def predict_proba(self, X):
        return self.model.predict_proba(X)
    def __sklearn_tags__(self):
        from sklearn.utils._tags import _safe_tags
        return _safe_tags(BaseEstimator(), key=None)

# --- UTILIDADES ---
class CommandRedirector:
    def __init__(self, widget, tag_parser=None):
        self.widget = widget
        self.tag_parser = tag_parser

    def write(self, text):
        if not text.strip(): return
        self.widget.insert("end", text + "\n")
        if self.tag_parser: self.tag_parser(text)
        self.widget.see("end")

    def flush(self): pass

# --- UI COMPONENTS ---

class SidebarButton(ctk.CTkFrame):
    def __init__(self, master, text, icon, command):
        super().__init__(master, fg_color="transparent", corner_radius=6, height=50)
        self.command = command
        
        self.pack_propagate(False) # Respetar altura
        
        # Layout: Icono fijo a la izquierda
        # Usamos Segoe UI Emoji para uniformidad en tamaños de iconos
        self.icon_lbl = ctk.CTkLabel(self, text=icon, width=50, font=("Segoe UI Emoji", 22), text_color="#BBBBBB", anchor="center")
        self.icon_lbl.pack(side="left", padx=(5,0))
        
        self.text_lbl = ctk.CTkLabel(self, text=text, font=("Segoe UI", 12, "bold"), text_color="#BBBBBB", anchor="w")
        self.text_lbl.pack(side="left", fill="both", expand=True)

        # Eventos para simular botón
        for w in [self, self.icon_lbl, self.text_lbl]:
            w.bind("<Button-1>", self._on_click)
            w.bind("<Enter>", self._on_hover)
            w.bind("<Leave>", self._on_leave)

        self.is_active = False

    def _on_click(self, event):
        self.command()

    def _on_hover(self, event):
        if not self.is_active:
            self.configure(fg_color="#333333")

    def _on_leave(self, event):
        if not self.is_active:
            self.configure(fg_color="transparent")

    def set_active(self, active):
        self.is_active = active
        if active:
            # Estilo Activo: Fondo sutil + Texto/Icono en Acento (Sin bordes)
            self.configure(fg_color="#2A2A2A", border_width=0)
            self.icon_lbl.configure(text_color=COLOR_ACCENT)
            self.text_lbl.configure(text_color=COLOR_ACCENT)
        else:
            self.configure(fg_color="transparent", border_width=0)
            self.icon_lbl.configure(text_color="#BBBBBB")
            self.text_lbl.configure(text_color="#BBBBBB")

class Sidebar_V3(ctk.CTkFrame):
    def __init__(self, master, command_callback):
        super().__init__(master, width=280, corner_radius=0, fg_color="#181818")
        self.command_callback = command_callback
        
        # Logo
        logo_frame = ctk.CTkFrame(self, fg_color="transparent")
        logo_frame.pack(pady=(40, 40))
        ctk.CTkLabel(logo_frame, text="MAIL", font=("Segoe UI", 26, "bold"), text_color="white").pack()
        ctk.CTkLabel(logo_frame, text="INTELLIGENCE", font=("Segoe UI", 26, "bold"), text_color=COLOR_ACCENT).pack(pady=(0,5))
        ctk.CTkLabel(logo_frame, text="AI POWERED ENGINE", font=("Segoe UI", 10, "bold"), text_color="gray").pack()

        # Botones Custom
        self.btn_monitor = self._create_btn("MONITOREO EN VIVO", "monitor", "👁️")
        self.btn_metrics = self._create_btn("MÉTRICAS Y ANÁLISIS", "metrics", "📈")
        self.btn_setup = self._create_btn("CONFIGURACIÓN", "setup", "⚙️")
        
        # Spacer
        ctk.CTkLabel(self, text="", height=50).pack(fill="y", expand=True)

        self.btn_about = self._create_btn("ACERCA DE", "about", "ℹ️")
        self.btn_about.pack(side="bottom", pady=30, padx=20, fill="x")

    def _create_btn(self, text, view, icon):
        btn = SidebarButton(self, text, icon, lambda: self.command_callback(view))
        btn.pack(fill="x", padx=20, pady=5)
        return btn

    def set_active(self, view_name):
        buttons = {'monitor': self.btn_monitor, 'setup': self.btn_setup, 'metrics': self.btn_metrics, 'about': self.btn_about}
        for name, btn in buttons.items():
            btn.set_active(name == view_name)

class KPICard_V3(ctk.CTkFrame):
    def __init__(self, master, title, value="0", icon=""):
        super().__init__(master, fg_color=COLOR_CARD, corner_radius=16, border_width=0)
        self.pack(side="left", fill="both", expand=True, padx=10)
        
        # Icono y Titulo
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=20, pady=(20,5))
        ctk.CTkLabel(head, text=icon, font=("Segoe UI", 20)).pack(side="left")
        ctk.CTkLabel(head, text=title.upper(), font=FONT_KPI_TITLE, text_color="gray").pack(side="right", pady=5)
        
        # Valor
        self.lbl_val = ctk.CTkLabel(self, text=value, font=FONT_KPI_VAL, text_color="white")
        self.lbl_val.pack(anchor="e", padx=20, pady=(0, 20))

    def update_val(self, val):
        self.lbl_val.configure(text=str(val))

# --- VISTAS ---

class MonitorView_V3(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # KPIs
        kpi_row = ctk.CTkFrame(self, fg_color="transparent")
        kpi_row.pack(fill="x", pady=(0, 25))
        
        self.card_total = KPICard_V3(kpi_row, "Escaneados", "0", "📨")
        self.card_urgent = KPICard_V3(kpi_row, "Alta Prioridad", "0", "🔥")
        self.card_low = KPICard_V3(kpi_row, "Baja Prioridad", "0", "🧊")

        # Consola "Terminal Style"
        console_frame = ctk.CTkFrame(self, fg_color="#0D0D0D", corner_radius=12, border_width=1, border_color="#333")
        console_frame.pack(fill="both", expand=True, pady=10)
        
        # Barra superior consola
        bar = ctk.CTkFrame(console_frame, height=30, fg_color="#1A1A1A", corner_radius=12)
        bar.pack(fill="x")
        ctk.CTkLabel(bar, text="  TERMINAL DE VIGILANCIA  ", font=("Consolas", 10, "bold"), text_color="gray").pack(side="left")

        self.console = ctk.CTkTextbox(console_frame, font=("Consolas", 11), fg_color="transparent", text_color="#00FF00", wrap="word")
        self.console.pack(fill="both", expand=True, padx=10, pady=10)
        self.console._textbox.tag_config("urgent", foreground="#FF3333", background="#220000", selectbackground="#FF3333")
        self.console._textbox.tag_config("review", foreground="#FF9800", background="#221100", selectbackground="#FF9800")
        self.console._textbox.tag_config("normal", foreground="#888888")

        self.btn = ctk.CTkButton(self, text="INICIAR VIGILANCIA", height=55, fg_color=COLOR_ACCENT, 
                                 text_color="black", font=("Segoe UI", 16, "bold"), hover_color="#00C853",
                                 command=self.run)
        self.btn.pack(fill="x", pady=20)
        
        # Loader
        self.loader = ctk.CTkProgressBar(self, mode="indeterminate", height=4, fg_color="#1A1A1A", progress_color=COLOR_ACCENT)
        self.loader.pack(fill="x")
        self.loader.pack_forget() # Ocultar inicial
        
        self.counts = {'total':0, 'urgent':0, 'low':0}

    def run(self):
        self.console.delete("1.0", "end")
        self.counts = {'total':0, 'urgent':0, 'low':0}
        self.update_ui()
        self.btn.configure(state="disabled", text="VIGILANDO...")
        
        # Mostrar Loader
        self.loader.pack(fill="x")
        self.loader.start()
        
        threading.Thread(target=self._thread, daemon=True).start()

    def _thread(self):
        pythoncom.CoInitialize()
        sys.stdout = CommandRedirector(self.console, self._parse)
        try: inference.ejecutar_vigilancia()
        except Exception as e: print(f"Error: {e}")
        finally:
            # Ocultar Loader
            self.loader.stop()
            self.loader.pack_forget()
            self.btn.configure(state="normal", text="REINICIAR VIGILANCIA")

    def _parse(self, text):
        if not self.winfo_exists(): return

        if "URGENTE" in text:
            self.console._textbox.tag_add("urgent", "end-2l", "end-1c")
            self.counts['urgent']+=1
        elif "REVISAR" in text:
            self.console._textbox.tag_add("review", "end-2l", "end-1c")
            self.counts['urgent']+=1 # Cuenta como prioridad
        else:
            self.console._textbox.tag_add("normal", "end-2l", "end-1c")
            if "IGNORADO" in text: self.counts['low']+=1
            
        if "[" in text: self.counts['total']+=1
        
        if self.winfo_exists():
            self.after(0, self.update_ui)

    def update_ui(self):
        if not self.winfo_exists(): return
        self.card_total.update_val(self.counts['total'])
        self.card_urgent.update_val(self.counts['urgent'])
        self.card_low.update_val(self.counts['low'])

class MetricsView_V3(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # Header
        h = ctk.CTkFrame(self, fg_color="transparent")
        h.pack(fill="x", pady=(0,20))
        ctk.CTkLabel(h, text="Métricas y Análisis de Comportamiento", font=FONT_HEADER, text_color="white").pack(side="left")
        ctk.CTkButton(h, text="Actualizar Datos", width=100, fg_color="#333", command=self.load).pack(side="right")

        self.g_container = ctk.CTkFrame(self, fg_color="transparent")
        self.g_container.pack(fill="both", expand=True)
        self.msg = ctk.CTkLabel(self.g_container, text="Cargando...", text_color="gray")
        self.msg.pack(pady=50)
        
        self.loaded = False # Cache flag

    def load(self):
        self.loaded = True
        for w in self.g_container.winfo_children(): w.destroy()
        try: df = pd.read_csv("dataset_masivo.csv", sep="|")
        except: 
            ctk.CTkLabel(self.g_container, text="No hay datos. Ejecuta la extracción primero.").pack()
            return

        # --- ROW 1: KPIs Rápidos ---
        r1 = ctk.CTkFrame(self.g_container, fg_color="transparent")
        r1.pack(fill="x", pady=(0,20))
        
        urgentes = len(df[df['TARGET_IA']==2])
        total = len(df)
        pct = (urgentes/total*100) if total else 0
        
        KPICard_V3(r1, "Total Mails", f"{total:,}", "📚").pack(side="left", padx=5)
        KPICard_V3(r1, "Tasa de Acción", f"{pct:.1f}%", "⚡").pack(side="left", padx=5)
        KPICard_V3(r1, "Respondidos", f"{urgentes:,}", "💬").pack(side="left", padx=5)

        # --- ROW 2: Graficos Principales ---
        r2 = ctk.CTkFrame(self.g_container, fg_color="transparent")
        r2.pack(fill="both", expand=True, pady=10)
        r2.grid_columnconfigure(0, weight=1); r2.grid_columnconfigure(1, weight=1); r2.grid_columnconfigure(2, weight=1)

        self._chart_wrapper(r2, 0, 0, "Distribución de Prioridad", lambda p: self._plot_pie(p, df))
        self._chart_wrapper(r2, 0, 1, "Efectividad por Contexto", lambda p: self._plot_bar_context(p, df))
        self._chart_wrapper(r2, 0, 2, "Impacto de Audiencia", lambda p: self._plot_audience_impact(p, df))
        
        # --- ROW 3: Fuentes (Dominios y Personas) ---
        r3 = ctk.CTkFrame(self.g_container, fg_color="transparent")
        r3.pack(fill="both", expand=True, pady=10)
        r3.grid_columnconfigure(0, weight=1); r3.grid_columnconfigure(1, weight=1)

        self._chart_wrapper(r3, 0, 0, "Top 5 Dominios Frecuentes", lambda p: self._plot_top_domains(p, df))
        self._chart_wrapper(r3, 0, 1, "Top Personas (Urgentes)", lambda p: self._plot_top_people(p, df))

        # --- ROW 4: Contenido y Carpetas ---
        r4 = ctk.CTkFrame(self.g_container, fg_color="transparent")
        r4.pack(fill="both", expand=True, pady=10)
        r4.grid_columnconfigure(0, weight=1); r4.grid_columnconfigure(1, weight=1)
        
        self._chart_wrapper(r4, 0, 0, "Top Carpetas Críticas", lambda p: self._plot_top_folders(p, df))
        self._chart_wrapper(r4, 0, 1, "Palabras Clave (Urgentes)", lambda p: self._plot_keywords(p, df))

    def _chart_wrapper(self, parent, r, c, title, func):
        f = ctk.CTkFrame(parent, fg_color=COLOR_CARD, corner_radius=12)
        f.grid(row=r, column=c, padx=10, pady=10, sticky="nsew")
        ctk.CTkLabel(f, text=title, font=("Segoe UI", 12, "bold"), text_color="gray").pack(pady=10)
        func(f)

    def _plot_pie(self, parent, df):
        fig, ax = plt.subplots(figsize=(4,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD)
        counts = df['TARGET_IA'].value_counts()
        ax.pie([counts.get(0,0), counts.get(1,0), counts.get(2,0)], 
               labels=['Ignorar', 'Info', 'Actuar'], autopct='%1.1f%%',
               colors=['#333', '#0091EA', COLOR_ACCENT], textprops={'color':'white'})
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _plot_bar_context(self, parent, df):
        fig, ax = plt.subplots(figsize=(4,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        t_to = df[df['Estoy_En_To']==1]['TARGET_IA'].apply(lambda x: x==2).mean()*100
        t_cc = df[df['Estoy_En_CC']==1]['TARGET_IA'].apply(lambda x: x==2).mean()*100
        
        ax.bar(['En Para', 'En Copia'], [t_to, t_cc], color=[COLOR_ACCENT, '#FFA000'])
        ax.tick_params(colors='white', which='both'); ax.spines['bottom'].set_color('gray')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['left'].set_visible(False)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _plot_audience_impact(self, parent, df):
        fig, ax = plt.subplots(figsize=(4,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        # Bining
        bins = [0, 1, 3, 10, 1000]
        labels = ['Solo Yo', '2-3', '4-10', 'Masivo']
        df['grupo'] = pd.cut(df['Total_Destinatarios'], bins=bins, labels=labels)
        
        data = df.groupby('grupo', observed=True)['TARGET_IA'].apply(lambda x: (x==2).mean()*100)
        
        ax.plot(data.index.astype(str), data.values, marker='o', color='#00E5FF', linewidth=2)
        ax.fill_between(range(len(data)), data.values, color='#00E5FF', alpha=0.1)
        
        ax.tick_params(colors='white'); ax.grid(color='#333')
        ax.spines['bottom'].set_color('gray'); ax.spines['left'].set_color('gray')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _plot_top_domains(self, parent, df):
        fig, ax = plt.subplots(figsize=(5,3), dpi=100) # Un poco más ancho
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        top = df['Dominio'].value_counts().head(5)
        top.plot(kind='barh', ax=ax, color='#0091EA')
        ax.invert_yaxis()
        ax.tick_params(colors='white'); ax.spines['bottom'].set_color('gray')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['left'].set_visible(False)
        fig.tight_layout() # Clave para que no corte etiquetas
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _plot_top_people(self, parent, df):
        fig, ax = plt.subplots(figsize=(5,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        # Filtramos urgentes (Target 2) y tomamos el remitente
        top_people = df[df['TARGET_IA'] == 2]['Remitente_ID'].value_counts().head(5)
        if top_people.empty:
            ax.text(0.5, 0.5, "Sin datos suficientes", color="white", ha="center")
        else:
            top_people.plot(kind='barh', ax=ax, color='#FF4081') # Color distinto
            ax.invert_yaxis()
        
        ax.tick_params(colors='white'); ax.spines['bottom'].set_color('gray')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['left'].set_visible(False)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _plot_top_folders(self, parent, df):
        fig, ax = plt.subplots(figsize=(5,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        # Carpetas con más urgentes
        df_u = df[df['TARGET_IA']==2]
        if df_u.empty:
             ax.text(0.5, 0.5, "Sin datos suficientes", color="white", ha="center")
        else:
            top = df_u['Carpeta_Origen'].value_counts().head(5)
            top.plot(kind='barh', ax=ax, color='#FF9800')
            ax.invert_yaxis()
            
        ax.tick_params(colors='white'); ax.spines['bottom'].set_color('gray')
        ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['left'].set_visible(False)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)





    def _plot_keywords(self, parent, df):
        # Nube de palabras simple en gráfico de barras
        fig, ax = plt.subplots(figsize=(8,3), dpi=100)
        fig.patch.set_facecolor(COLOR_CARD); ax.set_facecolor(COLOR_CARD)
        
        urgentes = df[df['TARGET_IA']==2]['Asunto'].dropna().astype(str).tolist()
        words = []
        for subj in urgentes:
            # Tokenizar simple
            w_list = re.findall(r'\w+', subj.lower())
            words.extend([w for w in w_list if len(w)>3 and w not in ['para','sobre','entre','este','fwd','re']])
            
        common = Counter(words).most_common(8) # Top 8 para llenar el ancho
        if common:
            tags, vals = zip(*common)
            ax.bar(tags, vals, color='#FF5252')
            ax.tick_params(colors='white'); ax.spines['bottom'].set_color('gray')
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['left'].set_visible(False)
        
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=parent).get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)


class SetupView_V3(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        
        # Estilo de config limpio
        self._section("1. Minería de Datos", "Configura tus credenciales y rango de extracción.", self._build_etl)
        self._section("2. Entrenamiento AI", "Entrena el cerebro con los datos extraídos.", self._build_train)

        # Console Log (Restaurado)
        ctk.CTkLabel(self, text="Registro de Operaciones:", font=("Segoe UI", 12, "bold"), text_color="gray").pack(anchor="w", pady=(10,0))
        
        # Loader Global para Setup
        self.loader = ctk.CTkProgressBar(self, mode="indeterminate", height=3, fg_color="#1A1A1A", progress_color="#F9A825")
        self.loader.pack(fill="x", pady=(5,0))
        self.loader.pack_forget()

        self.console = ctk.CTkTextbox(self, height=150, font=("Consolas", 11), fg_color="#0D0D0D", border_width=1, border_color="#333", text_color="#00FF00")
        self.console.pack(fill="both", expand=True, pady=10)

    def _section(self, title, sub, builder):
        f = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12)
        f.pack(fill="x", pady=(0,15))
        ctk.CTkLabel(f, text=title, font=("Segoe UI", 16, "bold"), text_color="white").pack(anchor="w", padx=20, pady=(15,0))
        ctk.CTkLabel(f, text=sub, font=("Segoe UI", 12), text_color="gray").pack(anchor="w", padx=20, pady=(0,10))
        builder(f)

    def _build_etl(self, parent):
        grid = ctk.CTkFrame(parent, fg_color="transparent")
        grid.pack(fill="x", padx=15, pady=10)
        
        self.entry_name = self._input(grid, "Tu Nombre", "Walter Llana")
        self.entry_email = self._input(grid, "Tu Email", "wllana@unibanca.pe")
        self.entry_days = self._input(grid, "Días", "365", width=60)
        
        self.btn_etl = ctk.CTkButton(parent, text="EJECUTAR DATA MINING", fg_color="#F9A825", text_color="black", height=40, font=("Segoe UI", 12, "bold"),
                      command=self.run_etl)
        self.btn_etl.pack(fill="x", padx=20, pady=20)

    def _input(self, parent, label, default, width=200):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(side="left", padx=5)
        ctk.CTkLabel(f, text=label, text_color="gray", font=("Segoe UI", 11)).pack(anchor="w")
        e = ctk.CTkEntry(f, width=width, fg_color="#2B2B2B", border_width=0, text_color="white")
        e.insert(0, default)
        e.pack()
        return e

    def _build_train(self, parent):
        self.btn_train = ctk.CTkButton(parent, text="ENTRENAR MODELO", fg_color="#0277BD", height=40, font=("Segoe UI", 12, "bold"),
                      command=self.run_train)
        self.btn_train.pack(fill="x", padx=20, pady=20)

    def run_etl(self): 
        extractor.MI_NOMBRE_MOSTRAR = self.entry_name.get()
        extractor.MI_EMAIL_CORPORATIVO = self.entry_email.get()
        dias = 365
        try: dias = int(self.entry_days.get())
        except: pass
        self._run_thread(lambda: extractor.generar_dataset_masivo(dias), self.btn_etl)

    def run_train(self): self._run_thread(trainer.entrenar_modelo_definitivo, self.btn_train)
    
    def _run_thread(self, target, active_btn=None):
        self.console.delete("1.0", "end")
        
        if active_btn: active_btn.configure(state="disabled")
        self.loader.pack(fill="x", pady=(5,0), before=self.console)
        self.loader.start()
        
        def task_wrapper():
            pythoncom.CoInitialize()
            original = sys.stdout
            sys.stdout = CommandRedirector(self.console)
            try:
                target()
                print("\n✅ Operación Finalizada.")
            except Exception as e:
                print(f"\n❌ Error: {e}")
            finally:
                sys.stdout = original
                self.loader.stop()
                self.loader.pack_forget()
                if active_btn: active_btn.configure(state="normal")

        threading.Thread(target=task_wrapper, daemon=True).start()

class AboutView_V3(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        f = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=20)
        f.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.5, relheight=0.6)
        
        ctk.CTkLabel(f, text="MAIL INTELLIGENCE", font=("Segoe UI", 30, "bold"), text_color=COLOR_ACCENT).pack(pady=(50,10))
        ctk.CTkLabel(f, text="Desarrollado para optimizar el flujo de trabajo\nmediante priorización predictiva.", font=("Segoe UI", 14), text_color="gray").pack()
        
        ctk.CTkButton(f, text="Ver en GitHub", width=200, height=50, fg_color="#333", hover_color="black",
                      command=lambda: webbrowser.open("https://github.com/WalterWr7/mail-intelligence-engine")).pack(pady=40)
        
        ctk.CTkLabel(f, text="v1.2.0 • Walter Llana", text_color="#555").pack(side="bottom", pady=20)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mail Intelligence Engine (V3)")
        self.geometry("1280x800")
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = Sidebar_V3(self, self.nav)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.main = ctk.CTkFrame(self, fg_color=COLOR_BG, corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew")

        self.views = {
            "monitor": MonitorView_V3(self.main),
            "metrics": MetricsView_V3(self.main),
            "setup": SetupView_V3(self.main),
            "about": AboutView_V3(self.main)
        }
        self.curr = None
        self.nav("monitor")

    def nav(self, name):
        if self.curr: self.curr.pack_forget()
        self.curr = self.views[name]
        self.curr.pack(fill="both", expand=True, padx=30, pady=30)
        self.sidebar.set_active(name)
        
        # Optimización: Solo cargar si no está cacheado
        if name == "metrics" and not self.views["metrics"].loaded: 
            self.views["metrics"].load()

    def on_closing(self):
        self.quit()
        self.destroy()
        os._exit(0) # Force kill threads

if __name__ == "__main__":
    App().mainloop()

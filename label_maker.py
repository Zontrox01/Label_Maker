import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk
from label_generator import generar_hoja_etiquetas, LabelConfig
from printer_labels import imprimir_hoja_etiquetas, exportar_hoja_pdf

def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

EXCEL_TXT = os.path.join(_get_base_dir(), "datos", "excel_labels.txt")
CONFIG_TXT = os.path.join(_get_base_dir(), "datos", "labels.txt")
PREVIEW_W = 800
PREVIEW_H = 600

class LabelMakerApp(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Generador de Etiquetas con QR")
        self.resizable(True, True)
        self.configure(bg="#f0f4f8")
                # Icono en barra de títulos y barra de tareas
        try:
            icono_path = os.path.join(_get_base_dir(), "datos", "icono.png")
            if os.path.exists(icono_path):
                img = tk.PhotoImage(file=icono_path)
                self.iconphoto(True, img)
        except Exception:
            pass
        
        self._df = None
        self._columnas = []
        self._preview_img = None
        self._config = LabelConfig()
        self.qty_entries = []
        self.qty_entries = []
                
        self._build_ui()
        self.minsize(900, 600)
        self.geometry("1200x700")  # Tamaño inicial más grande
        
        # Configurar el cierre correcto de la aplicación
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Cargar configuración guardada primero (valores por defecto)
        self._cargar_configuracion()
        
        # Configurar eventos para guardar automáticamente
        self._configurar_eventos_auto_save()
        
        # Cargar Excel por defecto si existe
        excel_default = self._cargar_excel_default()
        if excel_default:
            self._cargar_excel_desde_ruta(excel_default)
    
    def _on_closing(self):
        """Maneja el cierre correcto de la aplicación"""
        # Guardar configuración antes de cerrar
        self._guardar_configuracion()
        # Destruir la ventana
        self.destroy()
        # Salir del loop principal
        self.quit()
    
    def _configurar_eventos_auto_save(self):
        """Configura eventos para guardar la configuración automáticamente"""
        # Lista de widgets que dispararán auto-guardado
        widgets_a_guardar = [
            self.sheet_width, self.sheet_height, self.labels_per_row, self.labels_per_col,
            self.gap_h, self.margin_left, self.margin_top, self.label_width, self.label_height,
            self.qr_size
        ]
        
        # Para Entry widgets
        for widget in widgets_a_guardar:
            if widget:
                widget.bind("<KeyRelease>", lambda e: self._guardar_configuracion())
        
        # Para el Combobox de QR column
        if self.qr_column:
            self.qr_column.bind("<<ComboboxSelected>>", lambda e: self._guardar_configuracion())
        
    
    def _guardar_configuracion(self):
        """Guarda todos los parámetros actuales en el archivo labels.txt"""
        try:
            os.makedirs(os.path.dirname(CONFIG_TXT), exist_ok=True)
            with open(CONFIG_TXT, "w", encoding="utf-8") as f:
                # Configuración de hoja
                f.write(f"SHEET_WIDTH={self.sheet_width.get()}\n")
                f.write(f"SHEET_HEIGHT={self.sheet_height.get()}\n")
                
                # Diseño de etiquetas
                f.write(f"LABELS_PER_ROW={self.labels_per_row.get()}\n")
                f.write(f"LABELS_PER_COL={self.labels_per_col.get()}\n")
                f.write(f"GAP_MM={self.gap_h.get()}\n")
                f.write(f"MARGIN_LEFT={self.margin_left.get()}\n")
                f.write(f"MARGIN_TOP={self.margin_top.get()}\n")
                f.write(f"LABEL_WIDTH={self.label_width.get()}\n")
                f.write(f"LABEL_HEIGHT={self.label_height.get()}\n")
                
                # Contenido
                f.write(f"QR_COLUMN={self.qr_column.get()}\n")
                f.write(f"QR_SIZE={self.qr_size.get()}\n")
                
                # Cantidades, textos adicionales y tamaños (si hay DataFrame cargado)
                if self._df is not None and self.qty_entries:
                    cantidades = []
                    textos = []
                    font_sizes = []
                    for entry in self.qty_entries:
                        cantidades.append(entry.get())
                    for entry in self.texto_entries:
                        # Escapar caracteres especiales en el texto
                        texto = entry.get().replace("\\", "\\\\").replace("|", "\\|")
                        textos.append(texto)
                    for spin in self.fontsize_entries:
                        font_sizes.append(spin.get())
                    
                    f.write(f"QUANTITIES={'|'.join(cantidades)}\n")
                    f.write(f"EXTRA_TEXTS={'|'.join(textos)}\n")
                    f.write(f"FONT_SIZES={'|'.join(font_sizes)}\n")
                
                # Tamaño de ventana
                f.write(f"WINDOW_WIDTH={self.winfo_width()}\n")
                f.write(f"WINDOW_HEIGHT={self.winfo_height()}\n")
                
            print(f"Configuración guardada en {CONFIG_TXT}")
        except Exception as e:
            print(f"Error al guardar configuración: {e}")
    
    def _cargar_configuracion(self):
        """Carga la configuración desde el archivo labels.txt"""
        if not os.path.exists(CONFIG_TXT):
            print(f"No existe archivo de configuración, usando valores por defecto")
            return
        
        try:
            config_data = {}
            with open(CONFIG_TXT, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line:
                        key, value = line.split("=", 1)
                        config_data[key] = value
            
            # Aplicar configuración si los widgets existen
            if "SHEET_WIDTH" in config_data:
                self.sheet_width.delete(0, tk.END)
                self.sheet_width.insert(0, config_data["SHEET_WIDTH"])
            
            if "SHEET_HEIGHT" in config_data:
                self.sheet_height.delete(0, tk.END)
                self.sheet_height.insert(0, config_data["SHEET_HEIGHT"])
            
            if "LABELS_PER_ROW" in config_data:
                self.labels_per_row.delete(0, tk.END)
                self.labels_per_row.insert(0, config_data["LABELS_PER_ROW"])
            
            if "LABELS_PER_COL" in config_data:
                self.labels_per_col.delete(0, tk.END)
                self.labels_per_col.insert(0, config_data["LABELS_PER_COL"])
            
            if "GAP_MM" in config_data:
                self.gap_h.delete(0, tk.END)
                self.gap_h.insert(0, config_data["GAP_MM"])
            
            if "MARGIN_LEFT" in config_data:
                self.margin_left.delete(0, tk.END)
                self.margin_left.insert(0, config_data["MARGIN_LEFT"])
            
            if "MARGIN_TOP" in config_data:
                self.margin_top.delete(0, tk.END)
                self.margin_top.insert(0, config_data["MARGIN_TOP"])
            
            if "LABEL_WIDTH" in config_data:
                self.label_width.delete(0, tk.END)
                self.label_width.insert(0, config_data["LABEL_WIDTH"])
            
            if "LABEL_HEIGHT" in config_data:
                self.label_height.delete(0, tk.END)
                self.label_height.insert(0, config_data["LABEL_HEIGHT"])
            
            if "QR_SIZE" in config_data:
                self.qr_size.delete(0, tk.END)
                self.qr_size.insert(0, config_data["QR_SIZE"])
            
            # QR Column se aplica después de cargar el Excel
            if "QR_COLUMN" in config_data and config_data["QR_COLUMN"]:
                self._pending_qr_column = config_data["QR_COLUMN"]
            
            # Cantidades se aplican después de cargar el Excel
            if "QUANTITIES" in config_data:
                self._pending_quantities = config_data["QUANTITIES"].split("|")
            
            # Textos adicionales se aplican después de cargar el Excel
            if "EXTRA_TEXTS" in config_data:
                textos = config_data["EXTRA_TEXTS"].split("|")
                # Desescapar caracteres
                textos = [t.replace("\\|", "|").replace("\\\\", "\\") for t in textos]
                self._pending_texts = textos
            
            # Tamaños de letra se aplican después de cargar el Excel
            if "FONT_SIZES" in config_data:
                self._pending_font_sizes = config_data["FONT_SIZES"].split("|")
            
            # Tamaño de ventana
            if "WINDOW_WIDTH" in config_data and "WINDOW_HEIGHT" in config_data:
                try:
                    width = int(config_data["WINDOW_WIDTH"])
                    height = int(config_data["WINDOW_HEIGHT"])
                    if width > 100 and height > 100:  # Evitar tamaños inválidos
                        self.geometry(f"{width}x{height}")
                except:
                    pass
            
            print(f"Configuración cargada desde {CONFIG_TXT}")
        except Exception as e:
            print(f"Error al cargar configuración: {e}")
    
    def _aplicar_configuracion_pendiente(self):
        """Aplica configuración que depende del Excel (columna, cantidades, textos y tamaños)"""
        # Aplicar columna QR si estaba pendiente
        if hasattr(self, '_pending_qr_column') and self._pending_qr_column:
            if self._pending_qr_column in self._columnas:
                self.qr_column.set(self._pending_qr_column)
                print(f"Columna QR cargada: {self._pending_qr_column}")
            delattr(self, '_pending_qr_column')
        
        # Aplicar cantidades si estaban pendientes
        if hasattr(self, '_pending_quantities') and self._pending_quantities:
            for idx, entry in enumerate(self.qty_entries):
                if idx < len(self._pending_quantities):
                    entry.delete(0, tk.END)
                    entry.insert(0, self._pending_quantities[idx])
            print(f"Cantidades cargadas: {self._pending_quantities}")
            delattr(self, '_pending_quantities')
        
        # Aplicar textos adicionales si estaban pendientes
        if hasattr(self, '_pending_texts') and self._pending_texts:
            for idx, entry in enumerate(self.texto_entries):
                if idx < len(self._pending_texts):
                    entry.delete(0, tk.END)
                    entry.insert(0, self._pending_texts[idx])
            print(f"Textos cargados: {self._pending_texts}")
            delattr(self, '_pending_texts')
        
        # Aplicar tamaños de letra si estaban pendientes
        if hasattr(self, '_pending_font_sizes') and self._pending_font_sizes:
            for idx, spin in enumerate(self.fontsize_entries):
                if idx < len(self._pending_font_sizes):
                    spin.delete(0, tk.END)
                    spin.insert(0, self._pending_font_sizes[idx])
            print(f"Tamaños cargados: {self._pending_font_sizes}")
            delattr(self, '_pending_font_sizes')
        
        # Actualizar vista previa después de aplicar configuración
        self._actualizar_preview()
    
    def _build_ui(self):
        # Contenedor principal
        main_container = tk.Frame(self, bg="#f0f4f8")
        main_container.pack(fill="both", expand=True)
        
        # Configurar grid del contenedor principal
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Canvas y scrollbar
        canvas = tk.Canvas(main_container, bg="#f0f4f8", highlightthickness=0)
        h_scrollbar = ttk.Scrollbar(main_container, orient="horizontal", command=canvas.xview)
        v_scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        # Frame que contendrá todo el contenido
        scrollable_frame = tk.Frame(canvas, bg="#f0f4f8")
        
        # Forzar que el frame dentro del canvas tenga el tamaño del canvas
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Crear ventana en el canvas con ancho Y ALTO inicial
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", 
                            width=canvas.winfo_reqwidth(), height=canvas.winfo_reqheight())
        
        # Configurar el grid dentro del scrollable_frame para expansión
        scrollable_frame.grid_rowconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(0, weight=1)
        scrollable_frame.grid_columnconfigure(1, weight=3)  # Proporción izquierda:derecha
        
        # Empaquetar canvas y scrollbars
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Ahora construir la UI dentro de scrollable_frame
        self._build_scrollable_ui(scrollable_frame)
        
        # Forzar actualización del canvas cuando se redimensione la ventana
        self.bind("<Configure>", self._on_window_resize)
        # Actualizar tanto ancho como alto del frame interno
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(1, 
                                                              width=canvas.winfo_width(), 
                                                              height=canvas.winfo_height()))
    
    def _build_scrollable_ui(self, parent):
        """Construye toda la interfaz dentro del frame con scroll"""
        
        # Panel izquierdo - Configuración
        left_frame = tk.LabelFrame(parent, text="Configuración de Etiquetas", 
                                   bg="#f0f4f8", font=("Arial", 10, "bold"),
                                   padx=10, pady=10)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        
        # Configurar left_frame para expansión vertical
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        # Envolver left_frame en un canvas con scroll
        left_canvas = tk.Canvas(left_frame, bg="#f0f4f8", highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        
        left_scrollable = tk.Frame(left_canvas, bg="#f0f4f8")
        left_scrollable.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        
        left_canvas.create_window((0, 0), window=left_scrollable, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        
        # ========== TODO EL CONTENIDO DEL LEFT ==========
        # Configuración de hoja
        sheet_frame = tk.LabelFrame(left_scrollable, text="Configuración de la Hoja", 
                                   bg="#f0f4f8", font=("Arial", 9, "bold"),
                                   padx=10, pady=5)
        sheet_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(sheet_frame, text="Ancho (mm):", bg="#f0f4f8").grid(row=0, column=0, sticky="e", padx=(0, 5))
        self.sheet_width = tk.Entry(sheet_frame, width=10)
        self.sheet_width.insert(0, "115")
        self.sheet_width.grid(row=0, column=1, pady=2)
        
        tk.Label(sheet_frame, text="Alto (mm):", bg="#f0f4f8").grid(row=1, column=0, sticky="e", padx=(0, 5))
        self.sheet_height = tk.Entry(sheet_frame, width=10)
        self.sheet_height.insert(0, "171")
        self.sheet_height.grid(row=1, column=1, pady=2)
        
        # Diseño de etiquetas
        layout_frame = tk.LabelFrame(left_scrollable, text="Diseño de la Hoja", 
                                    bg="#f0f4f8", font=("Arial", 9, "bold"),
                                    padx=10, pady=5)
        layout_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(layout_frame, text="Etiquetas por fila:", bg="#f0f4f8").grid(row=0, column=0, sticky="w", pady=2)
        self.labels_per_row = tk.Entry(layout_frame, width=10)
        self.labels_per_row.insert(0, "4")
        self.labels_per_row.grid(row=0, column=1, pady=2)
        
        tk.Label(layout_frame, text="Etiquetas por columna:", bg="#f0f4f8").grid(row=1, column=0, sticky="w", pady=2)
        self.labels_per_col = tk.Entry(layout_frame, width=10)
        self.labels_per_col.insert(0, "6")
        self.labels_per_col.grid(row=1, column=1, pady=2)
        
        tk.Label(layout_frame, text="Distancia entre etiquetas (mm):", bg="#f0f4f8").grid(row=2, column=0, sticky="w", pady=2)
        self.gap_h = tk.Entry(layout_frame, width=10)
        self.gap_h.insert(0, "2")
        self.gap_h.grid(row=2, column=1, pady=2)
        
        tk.Label(layout_frame, text="Margen izquierdo (mm):", bg="#f0f4f8").grid(row=3, column=0, sticky="w", pady=2)
        self.margin_left = tk.Entry(layout_frame, width=10)
        self.margin_left.insert(0, "5")
        self.margin_left.grid(row=3, column=1, pady=2)
        
        tk.Label(layout_frame, text="Margen superior (mm):", bg="#f0f4f8").grid(row=4, column=0, sticky="w", pady=2)
        self.margin_top = tk.Entry(layout_frame, width=10)
        self.margin_top.insert(0, "8")
        self.margin_top.grid(row=4, column=1, pady=2)
        
        tk.Label(layout_frame, text="Ancho etiqueta (mm):", bg="#f0f4f8").grid(row=5, column=0, sticky="w", pady=2)
        self.label_width = tk.Entry(layout_frame, width=10)
        self.label_width.insert(0, "24")
        self.label_width.grid(row=5, column=1, pady=2)
        
        tk.Label(layout_frame, text="Alto etiqueta (mm):", bg="#f0f4f8").grid(row=6, column=0, sticky="w", pady=2)
        self.label_height = tk.Entry(layout_frame, width=10)
        self.label_height.insert(0, "24")
        self.label_height.grid(row=6, column=1, pady=2)
        
        # Contenido
        content_frame = tk.LabelFrame(left_scrollable, text="Contenido de la Etiqueta", 
                                     bg="#f0f4f8", font=("Arial", 9, "bold"),
                                     padx=10, pady=5)
        content_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(content_frame, text="Columna para el QR:", bg="#f0f4f8").grid(row=0, column=0, sticky="w", pady=2)
        self.qr_column = ttk.Combobox(content_frame, width=25, state="readonly")
        self.qr_column.grid(row=0, column=1, pady=2)
        
        tk.Label(content_frame, text="Tamaño del QR (mm):", bg="#f0f4f8").grid(row=1, column=0, sticky="w", pady=2)
        self.qr_size = tk.Entry(content_frame, width=10)
        self.qr_size.insert(0, "18")
        self.qr_size.grid(row=1, column=1, pady=2)
               
        # Cantidades
        qty_frame = tk.LabelFrame(left_scrollable, text="Cantidades por registro", 
                                 bg="#f0f4f8", font=("Arial", 9, "bold"),
                                 padx=10, pady=5)
        qty_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Canvas con scroll para las cantidades
        qty_canvas = tk.Canvas(qty_frame, bg="#f0f4f8", height=170)
        qty_scrollbar = ttk.Scrollbar(qty_frame, orient="vertical", command=qty_canvas.yview)
        self.qty_scrollable = tk.Frame(qty_canvas, bg="#f0f4f8")
        
        self.qty_scrollable.bind("<Configure>", lambda e: qty_canvas.configure(scrollregion=qty_canvas.bbox("all")))
        qty_canvas.create_window((0, 0), window=self.qty_scrollable, anchor="nw")
        qty_canvas.configure(yscrollcommand=qty_scrollbar.set)
        
        qty_canvas.pack(side="left", fill="both", expand=True)
        qty_scrollbar.pack(side="right", fill="y")
        
        # Botones (2 filas de 2 botones)
        actions_frame = tk.Frame(left_scrollable, bg="#f0f4f8")
        actions_frame.pack(fill="x", pady=(10, 0))
        
        # Configurar grid para 2 columnas
        actions_frame.grid_columnconfigure(0, weight=1)
        actions_frame.grid_columnconfigure(1, weight=1)
        
        # Fila 1
        tk.Button(actions_frame, text="📂 Cargar Excel", 
                 command=self._cargar_excel,
                 bg="#00509F", fg="white", font=("Arial", 10, "bold"),
                 relief="flat", cursor="hand2").grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        tk.Button(actions_frame, text="👁 Vista Previa", 
                 command=self._actualizar_preview,
                 bg="#00bb7e", fg="white", font=("Arial", 10, "bold"),
                 relief="flat", cursor="hand2").grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        
        # Fila 2
        tk.Button(actions_frame, text="💾 Guardar PDF", 
                 command=self._guardar_pdf,
                 bg="#9cbb00", fg="white", font=("Arial", 10, "bold"),
                 relief="flat", cursor="hand2").grid(row=1, column=0, padx=2, pady=2, sticky="ew")
                
        tk.Button(actions_frame, text="🖨 Imprimir", 
                 command=self._imprimir,
                 bg="#95009F", fg="white", font=("Arial", 10, "bold"),
                 relief="flat", cursor="hand2").grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        
        # ========== PANEL DERECHO ==========
        right_frame = tk.LabelFrame(parent, text="Vista Previa", 
                                   bg="#f0f4f8", font=("Arial", 10, "bold"),
                                   padx=10, pady=10)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        
        # Configurar right_frame para expansión
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)
        
        # Label para la vista previa (simple, sin canvas extra)
        self.preview_label = tk.Label(right_frame, bg="#cccccc", text="Vista previa aparecerá aquí")
        self.preview_label.pack(fill="both", expand=True)
        
        self.lbl_estado = tk.Label(right_frame, text="Carga un archivo Excel para empezar.",
                                  bg="#f0f4f8", fg="gray", font=("Arial", 9))
        self.lbl_estado.pack(pady=8)

    def _on_font_size_change(self):
        """Cuando cambia el tamaño de letra, actualiza preview y guarda"""
        self._guardar_configuracion()
        self._actualizar_preview()    
    
    def _on_window_resize(self, event):
        """Guarda la configuración cuando se redimensiona la ventana"""
        if event.widget == self:
            # Usar after para evitar guardar demasiado frecuentemente
            if hasattr(self, '_resize_timer'):
                self.after_cancel(self._resize_timer)
            self._resize_timer = self.after(500, self._guardar_configuracion)
    
    def _cargar_excel_default(self):
        try:
            if os.path.exists(EXCEL_TXT):
                with open(EXCEL_TXT, "r", encoding="utf-8") as f:
                    ruta = f.read().strip()
                    if os.path.exists(ruta):
                        return ruta
        except Exception:
            pass
        return None
    
    def _guardar_excel_path(self, path):
        try:
            os.makedirs(os.path.dirname(EXCEL_TXT), exist_ok=True)
            with open(EXCEL_TXT, "w", encoding="utf-8") as f:
                f.write(path)
        except Exception:
            pass
    
    def _cargar_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not path:
            return
        self._cargar_excel_desde_ruta(path)
    
    def _cargar_excel_desde_ruta(self, path):
        try:
            self._df = pd.read_excel(path, dtype=str).fillna("")
            self._columnas = list(self._df.columns)
            self._guardar_excel_path(path)
            
            # Actualizar combobox
            self.qr_column["values"] = self._columnas
            if self._columnas:
                self.qr_column.set(self._columnas[0])
            
            # Crear entradas para cantidades
            self._crear_cantidades_ui()
            
            # Aplicar configuración pendiente después de cargar el Excel
            self._aplicar_configuracion_pendiente()
            
            self.lbl_estado.config(text=f"✅ Cargado: {path.split('/')[-1]} — {len(self._df)} registros")
            self._guardar_configuracion()  # Guardar después de cargar Excel
            self._actualizar_preview()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
    
    def _crear_cantidades_ui(self):
        # Limpiar frame existente
        for widget in self.qty_scrollable.winfo_children():
            widget.destroy()
        
        if self._df is None:
            return
        
        self.qty_entries = []          # Para cantidades
        self.texto_entries = []        # Para textos adicionales
        self.fontsize_entries = []     # Para tamaños de letra
        
        # Añadir cabeceras (PRIMERO, antes de los datos)
        header_frame = tk.Frame(self.qty_scrollable, bg="#f0f4f8")
        header_frame.pack(fill="x", pady=(0, 5))
        tk.Label(header_frame, text="Registro", bg="#f0f4f8", font=("Arial", 8, "bold"), 
                width=8, anchor="w").pack(side="left", padx=(0, 5))
        tk.Label(header_frame, text="Cant", bg="#f0f4f8", font=("Arial", 8, "bold"), 
                width=4, anchor="center").pack(side="left", padx=(0, 5))
        tk.Label(header_frame, text="Texto adicional", bg="#f0f4f8", font=("Arial", 8, "bold"), 
                width=27, anchor="center").pack(side="left", padx=(0, 5))
        tk.Label(header_frame, text="Tamaño", bg="#f0f4f8", font=("Arial", 8, "bold"), 
                width=10, anchor="center").pack(side="left")
        
        # Luego los datos de cada registro
        for idx, row in self._df.iterrows():
            first_col = self._columnas[0] if self._columnas else None
            identificador = str(row[first_col]) if first_col else f"Registro {idx+1}"
            if len(identificador) > 25:
                identificador = identificador[:22] + "..."
            
            frame = tk.Frame(self.qty_scrollable, bg="#f0f4f8")
            frame.pack(fill="x", pady=2)
            
            # Columna 1: Identificador
            tk.Label(frame, text=f"{identificador}:", bg="#f0f4f8", 
                    font=("Arial", 8), width=10, anchor="w").pack(side="left", padx=(0, 5))
            
            # Columna 2: Cantidad
            entry_qty = tk.Entry(frame, width=4)
            entry_qty.insert(0, "1")
            entry_qty.pack(side="left", padx=(0, 5))
            entry_qty.bind("<KeyRelease>", lambda e: self._guardar_configuracion())
            self.qty_entries.append(entry_qty)
            
            # Columna 3: Texto adicional
            entry_texto = tk.Entry(frame, width=37)
            entry_texto.insert(0, "")
            entry_texto.pack(side="left", padx=(0, 5))
            entry_texto.bind("<KeyRelease>", lambda e: self._guardar_configuracion())
            self.texto_entries.append(entry_texto)
            
            # Columna 4: Tamaño letra (Spinbox)
            spin_fontsize = tk.Spinbox(frame, from_=8, to=30, width=3, 
                                        textvariable=tk.IntVar(value=12))
            spin_fontsize.pack(side="left")
            spin_fontsize.bind("<KeyRelease>", lambda e: self._guardar_configuracion())
            self.fontsize_entries.append(spin_fontsize)
    
    def _obtener_cantidades(self):
        cantidades = []
        for entry in self.qty_entries:
            try:
                val = int(entry.get())
                cantidades.append(max(0, val))
            except:
                cantidades.append(0)
        return cantidades
    
    def _actualizar_preview(self):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Carga un Excel primero.")
            return
        
        # Calcular total de etiquetas a generar
        cantidades = self._obtener_cantidades()
        total_etiquetas = sum(cantidades)
        
        # Calcular capacidad de la hoja
        try:
            labels_per_row = int(self.labels_per_row.get())
            labels_per_col = int(self.labels_per_col.get())
            capacidad_hoja = labels_per_row * labels_per_col
        except:
            capacidad_hoja = 0
        
        # Verificar si cabe en la hoja
        if total_etiquetas > capacidad_hoja:
            messagebox.showwarning(
                "Excede capacidad",
                f"El número total de etiquetas ({total_etiquetas}) excede la capacidad de la hoja ({capacidad_hoja}).\n"
                f"La hoja solo mostrará las primeras {capacidad_hoja} etiquetas.\n\n"
                f"Etiquetas por fila: {labels_per_row}\n"
                f"Etiquetas por columna: {labels_per_col}\n"
                f"Total que caben: {capacidad_hoja}"
            )
            # No detenemos la ejecución, solo mostramos advertencia
        
        def _generar():
            try:
                config = self._obtener_config()
                
                # Generar hoja de etiquetas en MODO PREVIEW (con borde)
                img_hoja = generar_hoja_etiquetas(self._df, config, cantidades, preview_mode=True)
                
                # Redimensionar para preview
                preview_img = img_hoja.copy()
                preview_img.thumbnail((PREVIEW_W, PREVIEW_H), Image.LANCZOS)
                photo = ImageTk.PhotoImage(preview_img)
                
                self.after(0, lambda: self._mostrar_preview(photo))
                
            except Exception as e:
                self.after(0, lambda: self.lbl_estado.config(text=f"Error: {str(e)}"))
        
        threading.Thread(target=_generar, daemon=True).start()
    
    def _mostrar_preview(self, photo):
        self._preview_img = photo
        self.preview_label.config(image=self._preview_img, text="")
        self.preview_label.image = self._preview_img
        self.lbl_estado.config(text="Vista previa actualizada")
    
    def _obtener_config(self):
        config = LabelConfig()
        
        # Configuración de hoja
        try:
            config.sheet_width_mm = float(self.sheet_width.get())
            config.sheet_height_mm = float(self.sheet_height.get())
            config.labels_per_row = int(self.labels_per_row.get())
            config.labels_per_col = int(self.labels_per_col.get())
            config.gap_mm = float(self.gap_h.get())
            config.margin_left_mm = float(self.margin_left.get())
            config.margin_top_mm = float(self.margin_top.get())
            config.label_width_mm = float(self.label_width.get())
            config.label_height_mm = float(self.label_height.get())
            config.qr_size_mm = float(self.qr_size.get())
            config.qr_column = self.qr_column.get()
            
            # Recoger textos adicionales por registro
            textos_por_registro = []
            for entry in self.texto_entries:
                textos_por_registro.append(entry.get())
            config.textos_por_registro = textos_por_registro
            
            # Recoger tamaños de letra por registro
            tamanios_por_registro = []
            for spin in self.fontsize_entries:
                try:
                    tamanios_por_registro.append(int(spin.get()))
                except:
                    tamanios_por_registro.append(12)
            config.tamanios_por_registro = tamanios_por_registro
            
        except:
            pass
        
        return config
    
    def _imprimir(self):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Carga un Excel primero.")
            return
        
        try:
            config = self._obtener_config()
            cantidades = self._obtener_cantidades()
            
            # Generar SIN borde (preview_mode=False)
            img_hoja = generar_hoja_etiquetas(self._df, config, cantidades, preview_mode=False)
            imprimir_hoja_etiquetas(img_hoja)
            
            self.lbl_estado.config(text="✅ Impresión enviada")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def _guardar_pdf(self):
        if self._df is None:
            messagebox.showwarning("Sin datos", "Carga un Excel primero.")
            return
        
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            title="Guardar hoja de etiquetas como PDF")
        if not path:
            return
        
        try:
            config = self._obtener_config()
            cantidades = self._obtener_cantidades()
            
            # Generar SIN borde (preview_mode=False)
            img_hoja = generar_hoja_etiquetas(self._df, config, cantidades, preview_mode=False)
            exportar_hoja_pdf(img_hoja, path, config)
            
            messagebox.showinfo("Éxito", f"PDF guardado en:\n{path}")
            self.lbl_estado.config(text=f"✅ PDF guardado")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana root
    app = LabelMakerApp(root)
    app.mainloop()
    root.destroy()  # Destruir root después de cerrar la app
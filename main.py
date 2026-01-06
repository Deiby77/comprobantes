# --------------------------------------------------
# IMPORTS
# --------------------------------------------------
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, simpledialog
import shutil
import os
import subprocess
import sys
import pythoncom
from win32com.shell import shell
import datetime
import pandas as pd
import threading
import unir_comprobantes_y_pagos as unir

# Colores del tema La 80 Supermercados (Negro dominante)
COLORS = {
    "bg": "#0a0a0a",           # Fondo negro profundo
    "bg_secondary": "#151515", # Fondo secundario
    "accent": "#1f1f1f",       # Acento gris oscuro
    "primary": "#e31e26",      # Rojo La 80
    "success": "#08f510",      # Verde La 80
    "warning": "#ffc107",      # Amarillo La 80
    "text": "#ffffff",         # Texto blanco
    "text_secondary": "#999999", # Texto gris
    "card": "#151515",         # Tarjetas
}


# --------------------------------------------------
# DETECCIÓN DE MODO EXE
# --------------------------------------------------
def get_base_path():
    """Obtiene la ruta base (donde está el exe o el script)"""
    if getattr(sys, 'frozen', False):
        # Ejecutándose como exe
        return os.path.dirname(sys.executable)
    else:
        # Ejecutándose como script
        return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()

# Ruta del logo (después de definir BASE_PATH)
LOGO_PATH = os.path.join(BASE_PATH, "logo.jpg")


# ENVÍO DE CORREOS
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# --------------------------------------------------
# CREDENCIALES (solo para reenvío)
# --------------------------------------------------
TU_CORREO = "carterasuperla80@gmail.com"
TU_PASSWORD = "qdao uewp cnbt kraz"

#TU_CORREO = "pruebasinformes1@gmail.com"
#TU_PASSWORD = "hevm duzl snlc nkqy"

CC_CARTERA = "liquidacionla80@gmail.com"



# --------------------------------------------------
# CONFIGURACIONES GENERALES
# --------------------------------------------------

# Excel fijo (base de datos) - relativo al exe
DEFAULT_EXCEL = os.path.join(BASE_PATH, "proveedores_correos.xlsx")

# Escritorio y carpeta Resultados
ESCRITORIO = os.path.join(os.path.expanduser("~"), "Desktop")
CARPETA_RESULTADOS = os.path.join(ESCRITORIO, "Resultados")
os.makedirs(CARPETA_RESULTADOS, exist_ok=True)

# Carpeta única por ejecución
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
CARPETA_RUN = os.path.join(CARPETA_RESULTADOS, timestamp)
os.makedirs(CARPETA_RUN, exist_ok=True)

PDF_INFORME = os.path.join(CARPETA_RUN, "Informe.pdf")
PDF_COMPROBANTE = os.path.join(CARPETA_RUN, "Comprobante.pdf")


# --------------------------------------------------
# BLOQUEAR / DESBLOQUEAR VENTANA
# --------------------------------------------------
def bloquear_ventana(root):
    root.config(cursor="wait")
    root.grab_set()

def desbloquear_ventana(root):
    root.config(cursor="")
    root.grab_release()


# --------------------------------------------------
# ACCESOS DIRECTOS
# --------------------------------------------------
def crear_acceso_directo(nombre_carpeta):
    try:
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        carpeta_destino = os.path.abspath(nombre_carpeta)
        acceso = os.path.join(escritorio, f"{os.path.basename(nombre_carpeta)}.lnk")

        shell_link = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink, None,
            pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
        )
        shell_link.SetPath(carpeta_destino)
        shell_link.SetDescription(f"Acceso directo a {nombre_carpeta}")

        persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
        persist_file.Save(acceso, 0)

        print(f"[INFO] Acceso directo creado: {acceso}")

    except Exception as e:
        print(f"[WARN] No se pudo crear acceso directo: {e}")


# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def elegir_carpeta_run(initial_dir=CARPETA_RESULTADOS):
    carpeta = filedialog.askdirectory(title="Selecciona RUN", initialdir=initial_dir)
    return carpeta or None


def localizar_excel_preferido(run_path=None):
    if os.path.exists(DEFAULT_EXCEL):
        return DEFAULT_EXCEL

    if run_path:
        candidato = os.path.join(run_path, "proveedores_correos.xlsx")
        if os.path.exists(candidato):
            return candidato

    candidato = os.path.join(os.getcwd(), "proveedores_correos.xlsx")
    if os.path.exists(candidato):
        return candidato

    messagebox.showinfo("Excel no encontrado", "Selecciona el archivo Excel.")
    return filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])


# --------------------------------------------------
# ACTUALIZAR BASE: ABRIR EXCEL
# --------------------------------------------------
def actualizar_base_de_datos(excel_path=None):
    if excel_path is None:
        excel_path = DEFAULT_EXCEL

    if not os.path.exists(excel_path):
        messagebox.showwarning("Archivo no encontrado", excel_path)
        return

    os.startfile(excel_path)

    # --------------------------------------------------
# REENVÍO REAL DE CORREOS (CORRECTO Y FINAL)
# --------------------------------------------------
def reenviar_no_enviados(run_path=None):
    print("\n==== REENVIANDO DOCUMENTOS ====\n")

    if not run_path:
        run_path = elegir_carpeta_run()
    if not run_path:
        return

    carpeta_no_enviados = os.path.join(run_path, "no_enviados")
    carpeta_enviados = os.path.join(run_path, "enviados")

    if not os.path.exists(carpeta_no_enviados):
        messagebox.showwarning("No existe carpeta no_enviados", run_path)
        return

    archivos = [
        os.path.join(carpeta_no_enviados, f)
        for f in os.listdir(carpeta_no_enviados)
        if f.lower().endswith(".docx")
    ]

    excel_path = localizar_excel_preferido(run_path)
    df = pd.read_excel(excel_path, dtype=str)
    df["NIT"] = df["NIT"].astype(str).str.replace(r"\D", "", regex=True)

    palabras_invalidas = ["nan", "none", "", "-", "sin correo"]

    fallidos = []
    enviados_count = 0

    for ruta in archivos:
        archivo = os.path.basename(ruta)
        nit = os.path.splitext(archivo)[0]

        print(f"\n➡ Procesando: {archivo} - NIT {nit}")

        row = df.loc[df["NIT"] == nit]
        correos_raw = str(row.iloc[0].get("CORREO", "")).strip().lower() if not row.empty else ""

        # Separa correos por ;
        lista_correos = [c.strip() for c in correos_raw.split(";")]

        # Filtra correos inválidos
        lista_correos = [
            c for c in lista_correos
            if c and c.lower() not in palabras_invalidas and "@" in c
        ]

        if not lista_correos:
            print(f"[WARN] {nit} sin correos válidos → No se reenvía")
            fallidos.append((nit, archivo, "Sin correo válido"))
            continue

        # CC solo si es válido
        if CC_CARTERA and CC_CARTERA.lower() not in palabras_invalidas and "@" in CC_CARTERA:
            lista_cc = [CC_CARTERA]
        else:
            lista_cc = []

        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.login(TU_CORREO, TU_PASSWORD)

            msg = MIMEMultipart()
            msg['From'] = TU_CORREO
            msg['To'] = ", ".join(lista_correos)
            if lista_cc:
                msg['Cc'] = ", ".join(lista_cc)
            msg['Subject'] = f"Reenvío - NIT {nit}"
            msg.attach(MIMEText("Adjuntamos nuevamente el comprobante.", 'plain'))

            with open(ruta, "rb") as adj:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(adj.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename=\"{archivo}\"')
                msg.attach(part)

            server.sendmail(TU_CORREO, lista_correos + lista_cc, msg.as_string())
            server.quit()

            print(f"[OK] Reenviado exitosamente: {nit}")

            os.makedirs(carpeta_enviados, exist_ok=True)
            shutil.copy(ruta, os.path.join(carpeta_enviados, archivo))
            os.remove(ruta)

            df.loc[df["NIT"] == nit, "ENVIADO"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            enviados_count += 1

        except Exception as e:
            fallidos.append((nit, archivo, str(e)))
            print(f"[ERROR] {nit}: {e}")

    df.to_excel(excel_path, index=False)

    print("\n==== REENVÍO FINALIZADO ====")
    print("Enviados:", enviados_count)
    print("Fallidos:", len(fallidos))

    messagebox.showinfo(
        "Reenvío completo",
        f"Enviados: {enviados_count}\nFallidos: {len(fallidos)}"
    )

# --------------------------------------------------
# GUI PRINCIPAL (DISEÑO MODERNO)
# --------------------------------------------------
class App:

    def __init__(self, root):
        self.root = root
        root.title("La 80 - Gestor de Comprobantes")
        root.geometry("500x680")
        root.resizable(False, False)
        root.configure(bg=COLORS["bg"])

        self.informe_path = None
        self.comprobante_path = None

        # Configurar estilos modernos
        self.setup_styles()
        
        # Frame principal con padding
        main_frame = tk.Frame(root, bg=COLORS["bg"])
        main_frame.pack(fill="both", expand=True, padx=30, pady=20)

        # ===== HEADER CON LOGO =====
        header_frame = tk.Frame(main_frame, bg=COLORS["bg"])
        header_frame.pack(fill="x", pady=(0, 15))
        
        # Cargar logo si existe
        self.logo_image = None
        try:
            if os.path.exists(LOGO_PATH):
                from PIL import Image, ImageTk
                logo = Image.open(LOGO_PATH)
                # Redimensionar logo manteniendo proporción
                logo.thumbnail((120, 80), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(logo)
                
                logo_label = tk.Label(header_frame, image=self.logo_image, bg=COLORS["bg"])
                logo_label.pack(pady=(0, 10))
        except Exception as e:
            print(f"[WARN] No se pudo cargar el logo: {e}")
        
        tk.Label(
            header_frame, 
            text="Gestor de Comprobantes",
            font=("Segoe UI", 18, "bold"),
            bg=COLORS["bg"],
            fg=COLORS["text"]
        ).pack()
        
        tk.Label(
            header_frame,
            text="Procesa y envía comprobantes automáticamente",
            font=("Segoe UI", 9),
            bg=COLORS["bg"],
            fg=COLORS["text_secondary"]
        ).pack(pady=(3, 0))

        # ===== SECCIÓN DE ARCHIVOS =====
        files_card = self.create_card(main_frame, "📁 Selección de Archivos")
        
        # Botón Informe
        self.btn_informe = self.create_file_button(
            files_card, 
            "📋 Seleccionar Informe PDF",
            self.select_informe
        )
        self.lbl_informe = tk.Label(
            files_card,
            text="No seleccionado",
            font=("Segoe UI", 9),
            bg=COLORS["card"],
            fg=COLORS["text_secondary"]
        )
        self.lbl_informe.pack(pady=(0, 10))
        
        # Botón Comprobante
        self.btn_comprobante = self.create_file_button(
            files_card,
            "💳 Seleccionar Comprobante PDF", 
            self.select_comprobante
        )
        self.lbl_comprobante = tk.Label(
            files_card,
            text="No seleccionado",
            font=("Segoe UI", 9),
            bg=COLORS["card"],
            fg=COLORS["text_secondary"]
        )
        self.lbl_comprobante.pack(pady=(0, 5))

        # ===== BOTÓN INICIAR PROCESO =====
        self.btn_iniciar = tk.Button(
            main_frame,
            text="🚀 INICIAR PROCESO",
            font=("Segoe UI", 12, "bold"),
            bg=COLORS["primary"],
            fg="#ffffff",
            activebackground="#b71c1c",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            height=2,
            command=self.cmd_iniciar_proceso
        )
        self.btn_iniciar.pack(fill="x", pady=15)
        self.add_hover_effect(self.btn_iniciar, COLORS["primary"], "#b71c1c")

        # ===== SECCIÓN DE PROGRESO =====
        progress_card = self.create_card(main_frame, "📊 Estado del Proceso")
        
        # Estado actual
        self.estado_label = tk.Label(
            progress_card,
            text="⏳ Esperando...",
            font=("Segoe UI", 11),
            bg=COLORS["card"],
            fg=COLORS["success"]
        )
        self.estado_label.pack(pady=(5, 10))
        
        # Barra de progreso estilizada
        style = ttk.Style()
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=COLORS["bg_secondary"],
            background=COLORS["success"],
            darkcolor=COLORS["success"],
            lightcolor=COLORS["success"],
            bordercolor=COLORS["bg_secondary"],
            thickness=25
        )
        
        self.progress = ttk.Progressbar(
            progress_card,
            style="Custom.Horizontal.TProgressbar",
            mode="indeterminate",
            length=380
        )
        self.progress.pack(pady=(0, 10))
        
        # Porcentaje/Detalle
        self.detalle_label = tk.Label(
            progress_card,
            text="",
            font=("Segoe UI", 9),
            bg=COLORS["card"],
            fg=COLORS["text_secondary"]
        )
        self.detalle_label.pack()

        # ===== SECCIÓN DE HERRAMIENTAS =====
        tools_card = self.create_card(main_frame, "🛠️ Herramientas")
        
        tools_frame = tk.Frame(tools_card, bg=COLORS["card"])
        tools_frame.pack(fill="x")
        
        # Botón Actualizar Base
        self.btn_actualizar = tk.Button(
            tools_frame,
            text="📊 Actualizar Base",
            font=("Segoe UI", 10),
            bg=COLORS["success"],
            fg="#ffffff",
            activebackground="#388e3c",
            relief="flat",
            cursor="hand2",
            width=18,
            height=2,
            command=self.cmd_actualizar_base
        )
        self.btn_actualizar.pack(side="left", padx=(0, 10))
        self.add_hover_effect(self.btn_actualizar, COLORS["success"], "#388e3c")
        
        # Botón Reenviar
        self.btn_reenviar = tk.Button(
            tools_frame,
            text="📨 Reenviar Fallidos",
            font=("Segoe UI", 10),
            bg=COLORS["warning"],
            fg="#000000",
            activebackground="#ffa000",
            relief="flat",
            cursor="hand2",
            width=18,
            height=2,
            command=self.cmd_reenviar
        )
        self.btn_reenviar.pack(side="right")
        self.add_hover_effect(self.btn_reenviar, COLORS["warning"], "#ffa000")

    def setup_styles(self):
        """Configura los estilos de ttk"""
        style = ttk.Style()
        style.theme_use('clam')
        
    def create_card(self, parent, title):
        """Crea una tarjeta con título"""
        card = tk.Frame(parent, bg=COLORS["card"], relief="flat")
        card.pack(fill="x", pady=8)
        
        # Título de la tarjeta
        tk.Label(
            card,
            text=title,
            font=("Segoe UI", 11, "bold"),
            bg=COLORS["card"],
            fg=COLORS["text"],
            anchor="w"
        ).pack(fill="x", padx=15, pady=(12, 8))
        
        # Línea separadora
        separator = tk.Frame(card, bg=COLORS["text_secondary"], height=1)
        separator.pack(fill="x", padx=15, pady=(0, 10))
        
        return card
    
    def create_file_button(self, parent, text, command):
        """Crea un botón de selección de archivo"""
        btn = tk.Button(
            parent,
            text=text,
            font=("Segoe UI", 10),
            bg=COLORS["bg_secondary"],
            fg=COLORS["text"],
            activebackground=COLORS["accent"],
            relief="flat",
            cursor="hand2",
            width=35,
            command=command
        )
        btn.pack(pady=5, padx=15)
        self.add_hover_effect(btn, COLORS["bg_secondary"], COLORS["accent"])
        return btn
    
    def add_hover_effect(self, button, normal_color, hover_color):
        """Agrega efecto hover a un botón"""
        button.bind("<Enter>", lambda e: button.configure(bg=hover_color))
        button.bind("<Leave>", lambda e: button.configure(bg=normal_color))

    # ----------------------------------------------
    # Selección de PDFs
    # ----------------------------------------------
    def select_informe(self):
        self.informe_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.informe_path:
            nombre = os.path.basename(self.informe_path)
            self.lbl_informe.config(text=f"✅ {nombre}", fg=COLORS["success"])
            print("[INFO] Informe:", self.informe_path)

    def select_comprobante(self):
        self.comprobante_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.comprobante_path:
            nombre = os.path.basename(self.comprobante_path)
            self.lbl_comprobante.config(text=f"✅ {nombre}", fg=COLORS["success"])
            print("[INFO] Comprobante:", self.comprobante_path)

    def actualizar_estado(self, texto, detalle=""):
        """Actualiza el estado mostrado en la interfaz"""
        self.estado_label.config(text=texto)
        self.detalle_label.config(text=detalle)
        self.root.update_idletasks()

    # ----------------------------------------------
    # INICIAR PROCESO
    # ----------------------------------------------
    def cmd_iniciar_proceso(self):
        bloquear_ventana(self.root)
        
        self.actualizar_estado("🔄 Iniciando proceso...", "Preparando archivos")
        self.progress.start(15)
        self.btn_iniciar.config(state="disabled", bg="#666666")

        def run():
            try:
                if not self.informe_path or not self.comprobante_path:
                    messagebox.showerror("Error", "Selecciona ambos PDF")
                    return

                self.actualizar_estado("📂 Copiando archivos...", "Preparando PDFs")
                shutil.copy(self.informe_path, PDF_INFORME)
                shutil.copy(self.comprobante_path, PDF_COMPROBANTE)

                print("[INFO] Archivos listos")
                self.actualizar_estado("⚙️ Procesando OCR...", "Extrayendo datos de los comprobantes")

                os.environ["RUTA_RUN"] = CARPETA_RUN
                unir.ejecutar_unir()

                self.actualizar_estado("✅ ¡Proceso completado!", "Revisa la carpeta de resultados")

            except Exception as e:
                self.actualizar_estado("❌ Error en el proceso", str(e)[:50])
                messagebox.showerror("Error", str(e))

            finally:
                desbloquear_ventana(self.root)
                self.progress.stop()
                self.btn_iniciar.config(state="normal", bg=COLORS["primary"])

        threading.Thread(target=run, daemon=True).start()

    # ----------------------------------------------
    # Abrir Excel base
    # ----------------------------------------------
    def cmd_actualizar_base(self):
        actualizar_base_de_datos(DEFAULT_EXCEL)

    # ----------------------------------------------
    # Reenvío real
    # ----------------------------------------------
    def cmd_reenviar(self):
        run = elegir_carpeta_run()
        if not run:
            return

        bloquear_ventana(self.root)

        self.actualizar_estado("📨 Reenviando correos...", "Procesando documentos fallidos")
        self.progress.start(15)
        self.btn_reenviar.config(state="disabled")

        def run_thread():
            try:
                reenviar_no_enviados(run)
                self.actualizar_estado("✅ Reenvío completado", "")
            finally:
                desbloquear_ventana(self.root)
                self.progress.stop()
                self.btn_reenviar.config(state="normal")

        threading.Thread(target=run_thread, daemon=True).start()


# --------------------------------------------------
# EJECUCIÓN PRINCIPAL
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

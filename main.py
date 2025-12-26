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

CC_CARTERA = "carteracontadola80@gmail.com"
#CC_CARTERA = "pruebasinformes1@gmail.com"


# --------------------------------------------------
# CONFIGURACIONES GENERALES
# --------------------------------------------------

# Excel fijo (base de datos)
DEFAULT_EXCEL = r"C:\comprobantes\proveedores_correos.xlsx"

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
# GUI PRINCIPAL
# --------------------------------------------------
class App:

    def __init__(self, root):
        self.root = root
        root.title("Seleccionar Archivos")
        root.geometry("420x260")
        root.resizable(False, False)

        self.informe_path = None
        self.comprobante_path = None

        tk.Label(root, text="Selecciona los archivos PDF", font=("Segoe UI", 14, "bold")).pack(pady=10)

        tk.Button(root, text="Seleccionar Informe PDF", width=34, command=self.select_informe).pack(pady=4)
        tk.Button(root, text="Seleccionar Comprobante PDF", width=34, command=self.select_comprobante).pack(pady=4)

        tk.Button(root, text="Iniciar Proceso", width=28, bg="#4CAF50", fg="white",
                  command=self.cmd_iniciar_proceso).pack(pady=8)

        tk.Button(root, text="Actualizar base de datos", width=28, bg="#2196F3", fg="white",
                  command=self.cmd_actualizar_base).pack(pady=4)

        tk.Button(root, text="Reenviar no enviados", width=28, bg="#FF9800", fg="white",
                  command=self.cmd_reenviar).pack(pady=4)
        
        # -------- BARRA DE PROGRESO --------
        self.estado_label = tk.Label(root, text="", font=("Segoe UI", 10))
        self.estado_label.pack(pady=4)

        self.progress = ttk.Progressbar(
            root,
            mode="indeterminate",
            length=300
        )
        self.progress.pack(pady=4)



    # ----------------------------------------------
    # Selección de PDFs
    # ----------------------------------------------
    def select_informe(self):
        self.informe_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.informe_path:
            print("[INFO] Informe:", self.informe_path)

    def select_comprobante(self):
        self.comprobante_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if self.comprobante_path:
            print("[INFO] Comprobante:", self.comprobante_path)

    # ----------------------------------------------
    # INICIAR PROCESO
    # ----------------------------------------------
    def cmd_iniciar_proceso(self):
        bloquear_ventana(self.root)
        
        self.estado_label.config(text="Cargando...")
        self.progress.start(10)

        def run():
            try:
                if not self.informe_path or not self.comprobante_path:
                    messagebox.showerror("Error", "Selecciona ambos PDF")
                    return

                shutil.copy(self.informe_path, PDF_INFORME)
                shutil.copy(self.comprobante_path, PDF_COMPROBANTE)

                print("[INFO] Archivos listos")

                os.environ["RUTA_RUN"] = CARPETA_RUN
                unir.ejecutar_unir()


            except Exception as e:
                messagebox.showerror("Error", str(e))

            finally:
                desbloquear_ventana(self.root)

                self.progress.stop()
                self.estado_label.config(text="Terminado")

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

        self.estado_label.config(text="Cargando...")
        self.progress.start(10)


        def run_thread():
            try:
                reenviar_no_enviados(run)
            finally:
                desbloquear_ventana(self.root)

                self.progress.stop()
                self.estado_label.config(text="Terminado")

        threading.Thread(target=run_thread, daemon=True).start()


# --------------------------------------------------
# EJECUCIÓN PRINCIPAL
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

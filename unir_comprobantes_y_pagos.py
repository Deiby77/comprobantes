# ============================================================
# IMPORTS ORDENADOS
# ============================================================

# --- Manejo de archivos y sistema ---
import os
import re
import sys
import shutil
import subprocess
from datetime import datetime


# --- Librerías externas de procesamiento PDF / DOCX / OCR ---
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
import pytesseract

# --- Procesamiento de imágenes ---
import cv2
import numpy as np
from PIL import Image, ImageChops
from difflib import SequenceMatcher
import easyocr


# --- Email ---
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# --- Excel ---
import pandas as pd


# ============================================================
# CARGAR RUTA_RUN
# ============================================================
def ejecutar_unir():
    ...


    RUTA_RUN = os.environ.get("RUTA_RUN")
    if not RUTA_RUN or not os.path.exists(RUTA_RUN):
        print("ERROR: No se recibió RUTA_RUN desde main.py o la ruta no existe.")
        sys.exit()

    # Rutas PDF
    pdf_comprobantes = os.path.join(RUTA_RUN, "Informe.pdf")
    pdf_pagos = os.path.join(RUTA_RUN, "Comprobante.pdf")

    # Carpetas internas
    CARP_RECORTES_COMPROBANTES = os.path.join(RUTA_RUN, "recortes_comprobantes")
    CARP_RECORTES_PAGOS = os.path.join(RUTA_RUN, "recortes_pagos")
    CARP_RESULTADOS = os.path.join(RUTA_RUN, "resultados")
    CARP_ENVIADOS = os.path.join(RUTA_RUN, "enviados")
    CARP_NO_ENVIADOS = os.path.join(RUTA_RUN, "no_enviados")
    CARP_NO_CONECTADOS = os.path.join(RUTA_RUN, "no_conectados")

    # Crear carpetas
    os.makedirs(CARP_RECORTES_COMPROBANTES, exist_ok=True)
    os.makedirs(CARP_RECORTES_PAGOS, exist_ok=True)
    os.makedirs(CARP_RESULTADOS, exist_ok=True)
    os.makedirs(CARP_ENVIADOS, exist_ok=True)
    os.makedirs(CARP_NO_ENVIADOS, exist_ok=True)
    os.makedirs(CARP_NO_CONECTADOS, exist_ok=True)


    # ============================================================
    # FUNCIONES BÁSICAS
    # ============================================================

    def safe_str(valor):
        if valor is None:
            return ""
        return str(valor)


    # ============================================================
    # CONFIGURAR TESSERACT
    # ============================================================

    try:
        subprocess.run(["tesseract", "--version"], check=True, stdout=subprocess.DEVNULL)
        print("Tesseract detectado automaticamente.")
    except Exception:
        posible = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
        if os.path.exists(posible):
            pytesseract.pytesseract.tesseract_cmd = posible
            print("Tesseract configurado manualmente: " + posible)
        else:
            print("ERROR: No se encontro Tesseract. Instalalo desde la pagina oficial.")
            sys.exit()

    # ============================================================
    # CONFIGURAR EASYOCR
    # ============================================================

    print("Inicializando EasyOCR...")
    easy_reader = easyocr.Reader(
        ['es'], 
        gpu=False
    )

    # ============================================================
    # VALIDACIÓN DE PDFs
    # ============================================================

    if not os.path.exists(pdf_comprobantes) or not os.path.exists(pdf_pagos):
        print("ERROR: No se encontraron los archivos requeridos:")
        print(" - " + pdf_comprobantes)
        print(" - " + pdf_pagos)
        sys.exit()


    # ============================================================
    # FUNCIONES DE IMAGEN
    # ============================================================

    def recortar_bordes(img):
        fondo = Image.new(img.mode, img.size, img.getpixel((0, 0)))
        dif = ImageChops.difference(img, fondo)
        bbox = dif.getbbox()
        return img.crop(bbox) if bbox else img


    def dividir_pagina_en_pagos(imagen, salida_base):
        gray = cv2.cvtColor(np.array(imagen), cv2.COLOR_RGB2GRAY)
        _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)

        projection = np.sum(thresh, axis=1)
        limites, dentro, inicio = [], False, 0

        for i, val in enumerate(projection):
            if val > 700 and not dentro:
                inicio, dentro = i, True
            elif val < 700 and dentro:
                limites.append((inicio, i))
                dentro = False

        recortes = []
        for j, (y1, y2) in enumerate(limites):
            if y2 - y1 > 90:
                crop = imagen.crop((0, y1, imagen.width, y2))
                path = f"{salida_base}_{j+1}.jpg"
                crop.save(path)
                recortes.append(path)

        return recortes


    def procesar_pdf(pdf, carpeta, tipo, dividir=False):
        print("Convirtiendo PDF: " + tipo + "...")
        paginas = convert_from_path(pdf, dpi=150)
        rutas = []

        for i, img in enumerate(paginas):
            img = recortar_bordes(img)
            if dividir:
                recortes = dividir_pagina_en_pagos(img, os.path.join(carpeta, f"{tipo}_{i+1}"))
                rutas.extend(recortes)
            else:
                ruta = os.path.join(carpeta, f"{tipo}_{i+1}.jpg")
                img.save(ruta)
                rutas.append(ruta)

        print(str(len(rutas)) + " imagenes generadas para " + tipo)
        return rutas


    rutas_comprobantes = procesar_pdf(pdf_comprobantes, CARP_RECORTES_COMPROBANTES, "comprobante", dividir=False)
    rutas_pagos = procesar_pdf(pdf_pagos, CARP_RECORTES_PAGOS, "pago", dividir=True)


    # ============================================================
    # OCR
    # ============================================================

    def leer_texto(imagen):
        img = cv2.imread(imagen, cv2.IMREAD_GRAYSCALE)
        img = cv2.threshold(img, 180, 255, cv2.THRESH_BINARY)[1]
        texto = pytesseract.image_to_string(img, lang="spa")
        return texto.upper().replace("\n", " ")
    



    # ============================================================
    # OCR MEJORADO (SEGUNDO INTENTO)
    # ============================================================

    def leer_texto_mejorado(imagen):
        img = cv2.imread(imagen, cv2.IMREAD_GRAYSCALE)

        # Escalar imagen (mejora OCR)
        img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

        # Reducir ruido
        img = cv2.GaussianBlur(img, (3, 3), 0)

        # Umbral adaptativo
        img = cv2.adaptiveThreshold(
            img, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31, 2
        )

        texto = pytesseract.image_to_string(
            img,
            lang="spa",
            config="--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        )

        return texto.upper().replace("\n", " ")
    

    def leer_texto_easyocr(imagen):
        resultados = easy_reader.readtext(
            imagen,
            detail=0,
            paragraph=True
        )

        texto = " ".join(resultados)
        return texto.upper().replace("\n", " ")

    # ============================================================
    # EXTRACCIÓN DE DATOS
    # ============================================================

    def extraer_datos(texto, es_pago=False):
        nombre, numero, valor, valor_raw = None, None, None, None


        m_nombre = re.search(r"(?:BENEFICIARIO|NOMBRE DE BENEFICIARIO)[:\.\s]+([A-ZÑ&.\s]{10,})", texto)
        if not m_nombre:
            m_nombre = re.search(
                r"NÚMERO DE BENEFICIARIO[:\.\s]+([A-ZÑ&.\s]+?)(?:\s+NIT|\s+DOCUMENTO|\s+CÉDULA|$)",
                texto
            )
        if not m_nombre:
            m_nombre = re.search(
                r"([A-ZÑ\s]{10,})\s+(NIT|DOCUMENTO|CÉDULA|NÚMERO DE BENEFICIARIO)",
                texto
            )

        if m_nombre:
            nombre = m_nombre.group(1).strip()
            nombre = re.sub(r"\s{2,}", " ", nombre)
            nombre = re.sub(r"(?i)\s+(NIT|DOCUMENTO|CÉDULA).*", "", nombre)

       

        if es_pago:
            # 🔥 PAGO → SOLO DOCUMENTO
            m_doc = re.search(r"DOCUMENTO\s*[:\.\s]*([0-9\.\-]+)", texto)
            if m_doc:
                numero = re.sub(r"\D", "", m_doc.group(1))

        else:
            # ✅ INFORME → SE QUEDA IGUAL
            patrones_nit = [
                r"(?:NIT|DOCUMENTO|CÉDULA|NÚMERO DE BENEFICIARIO)[:\.\s]*([0-9\.\-]+)",
                r"(?:NIT|DOCUMENTO|CÉDULA)[\s\:]*([0-9\.\-]+)",
                r"\b(90\d{7}|80\d{7}|\d{9,10})\b"
            ]

            for patron in patrones_nit:
                m_nit = re.search(patron, texto)
                if m_nit:
                    numero_raw = m_nit.group(1) if len(m_nit.groups()) > 0 else m_nit.group(0)
                    numero = re.sub(r"\D", "", numero_raw)
                    if len(numero) >= 8:
                        break

        # Normalizar ceros a la izquierda
        if numero:
            numero = numero.lstrip("0")

    
       # ---------- EXTRAER VALOR ----------
        m_valor = re.search(r"\$?\s*\d{1,3}(?:\.\d{3})+(?:,\d{2})?", texto)

        if m_valor:
            valor_raw = m_valor.group(0)
            valor_raw = valor_raw.replace("$", "").replace(" ", "")

        if valor_raw:
            valor = valor_raw.replace(".", "").replace(",", "")


        return nombre, numero, valor



    def limpiar_num(num):
        if not num:
            return ""
        return str(int(num))


    # ============================================================
    # SIMILITUD DE TEXTO
    # ============================================================

    def similitud(a, b):
        if not a or not b:
            return 0
        return SequenceMatcher(None, a, b).ratio()

    # ============================================================
    # PROCESAR CARPETAS
    # ============================================================

    def procesar_carpeta(rutas, tipo):
        datos = []

        print("Analizando " + tipo + "s...")
        for r in rutas:
            texto = leer_texto(r)
            nombre, numero, valor = extraer_datos(texto, es_pago=(tipo == "pago"))


            datos.append({"archivo": r, "nombre": nombre, "numero": numero, "valor": valor})

            print(os.path.basename(r) + " -> " + (nombre or "Sin nombre") +
                " | NIT: " + safe_str(numero) +
                " | Valor: " + safe_str(valor))

        return datos


    comprobantes = procesar_carpeta(rutas_comprobantes, "comprobante")
    pagos = procesar_carpeta(rutas_pagos, "pago")

    # ============================================================
    # AGRUPAR COMPROBANTES POR NIT
    # ============================================================

    def agrupar_por_nit(lista):
        grupos = {}
        for item in lista:
            nit = item["numero"] or "SIN_NIT"
            if nit not in grupos:
                grupos[nit] = []
            grupos[nit].append(item)
        return grupos

    # Agrupación
    grupos_comprobantes = agrupar_por_nit(comprobantes)



    # ============================================================
    # SEGUNDO INTENTO DE CONEXIÓN (NO CONECTADOS)
    # ============================================================

    def segundo_intento_conexion(pagos_sobrantes, comprobantes_no_conectados):
        nuevas_conexiones = []

        for pago in pagos_sobrantes:
            # 🔁 Intento 1: OCR mejorado con Tesseract
            texto_pago = leer_texto_mejorado(pago["archivo"])
            nom_p, num_p, val_p = extraer_datos(texto_pago, es_pago=True)

            # 🔁 Intento 2: EasyOCR SOLO si no sacó número
            if not num_p:
                texto_pago_easy = leer_texto_easyocr(pago["archivo"])
                nom_e, num_e, val_e = extraer_datos(texto_pago_easy, es_pago=True)

                nom_p = nom_p or nom_e
                num_p = num_p or num_e
                val_p = val_p or val_e


            pago["nombre"] = nom_p or pago["nombre"]
            pago["numero"] = num_p or pago["numero"]
            pago["valor"] = val_p or pago["valor"]

            for comp in comprobantes_no_conectados:
                score = 0

                # 1️⃣ Documento / NIT
                if pago["numero"] and comp["numero"]:
                    if limpiar_num(pago["numero"]) == limpiar_num(comp["numero"]):
                        score += 1

                # 2️⃣ Valor (tolerancia pequeña)
                if pago["valor"] and comp["valor"]:
                    try:
                        if abs(int(pago["valor"]) - int(comp["valor"])) <= 100:
                            score += 1
                    except:
                        pass

                # 3️⃣ Nombre parecido
                if similitud(pago["nombre"], comp["nombre"]) >= 0.85:
                    score += 1

                # ✅ CONEXIÓN SEGURA (2 DE 3)
                if score >= 2:
                    nuevas_conexiones.append((comp, pago))
                    break

        return nuevas_conexiones


    # ============================================================
    # EMPAREJAR GRUPOS DE COMPROBANTES CON PAGOS
    # ============================================================

    print("Emparejando comprobantes agrupados por NIT...")

    pares = []

    for nit, lista_comps in grupos_comprobantes.items():

        mejor_pago = None

        for pago in pagos:
            if limpiar_num(pago["numero"]) == limpiar_num(nit):
                mejor_pago = pago
                break

        if mejor_pago:
            pagos.remove(mejor_pago)

        pares.append((nit, lista_comps, mejor_pago))


    # ============================================================
    # CREAR WORD FINAL
    # ============================================================

    for nit, lista_comps, pago in pares:

        nombre_archivo = re.sub(r'[\\\/*?:"<>|]', "", nit)

        doc = Document()
        doc.add_heading(f"Comprobantes agrupados — NIT {nit}", level=1)

        # Insertar todas las páginas del comprobante
        for comp in lista_comps:
            doc.add_paragraph(f"Página del comprobante — {comp['nombre']}")
            doc.add_paragraph(f"NIT: {safe_str(comp['numero'])} | Valor: {safe_str(comp['valor'])}")
            doc.add_picture(comp["archivo"], width=Inches(6.5))
            doc.add_page_break()

        # Insertar pago si existe
        if pago:
            doc.add_heading("PAGO CORRESPONDIENTE", level=2)
            doc.add_paragraph(f"NIT: {safe_str(pago['numero'])} | Valor: {safe_str(pago['valor'])}")
            doc.add_picture(pago["archivo"], width=Inches(6.5))
        else:
            doc.add_paragraph("NO SE ENCONTRÓ PAGO.")

        path = os.path.join(CARP_RESULTADOS, f"{nombre_archivo}.docx")
        doc.save(path)
        print("Guardado: " + path)


    # ============================================================
    # RESUMEN NO CONECTADOS
    # ============================================================

    comprobantes_no_conectados = []
    for nit, lista_comps, pago in pares:
        if pago is None:
            # agregar todas las páginas del comprobante
            comprobantes_no_conectados.extend(lista_comps)

    pagos_sobrantes = pagos


    # ============================================================
    # SEGUNDO ESCANEO AUTOMÁTICO
    # ============================================================

    print("🔁 Iniciando segundo intento de conexión inteligente...")

    nuevas_conexiones = segundo_intento_conexion(
        pagos_sobrantes,
        comprobantes_no_conectados
    )

    for comp, pago in nuevas_conexiones:
        print("✔ Conectado en segundo intento:",
              comp["numero"], "↔", pago["numero"])

        pares.append((comp["numero"], [comp], pago))

        if pago in pagos_sobrantes:
            pagos_sobrantes.remove(pago)

        if comp in comprobantes_no_conectados:
            comprobantes_no_conectados.remove(comp)



    if comprobantes_no_conectados or pagos_sobrantes:
        resumen = Document()
        resumen.add_heading("REVISION MANUAL REQUERIDA", level=1)
        resumen.add_paragraph("Comprobantes sin pago o pagos sobrantes.\n")

        if comprobantes_no_conectados:
            resumen.add_heading("COMPROBANTES SIN PAGO", level=2)
            for c in comprobantes_no_conectados:
                resumen.add_paragraph("- " + (c['nombre'] or 'Sin nombre') +
                                    " | NIT: " + safe_str(c['numero']) +
                                    " | Valor: " + safe_str(c['valor']))

                run = resumen.add_paragraph().add_run()
                run.add_picture(c["archivo"], width=Inches(6.5))
            resumen.add_page_break()

        if pagos_sobrantes:
            resumen.add_heading("PAGOS SOBRANTES", level=2)
            for pago in pagos_sobrantes:
                resumen.add_paragraph("NIT: " + safe_str(pago['numero']) +
                                    " | Valor: " + safe_str(pago['valor']))

                run = resumen.add_paragraph().add_run()
                run.add_picture(pago["archivo"], width=Inches(6.5))

        resumen.save(os.path.join(CARP_NO_CONECTADOS, "NO_CONECTADOS_Y_PAGOS_SOBRANTES.docx"))

    # ============================================================    
    # ENVÍO AUTOMÁTICO GMAIL
    # ============================================================

    TU_CORREO = "carterasuperla80@gmail.com"
    TU_PASSWORD = "qdao uewp cnbt kraz"

    #TU_CORREO = "pruebasinformes1@gmail.com"
    #TU_PASSWORD = "hevm duzl snlc nkqy"

    CC_CARTERA = "carteracontadola80@gmail.com"
    #CC_CARTERA = "pruebasinformes1@gmail.com"

    fallidos = []

    try:
        print("Iniciando envio por Gmail...")

        excel_path = os.path.join(RUTA_RUN, "proveedores_correos.xlsx")
        if not os.path.exists(excel_path):
            excel_path = "proveedores_correos.xlsx"

        if not os.path.exists(excel_path):
            print("ERROR: No encontre proveedores_correos.xlsx")
        else:
            df = pd.read_excel(excel_path, dtype=str)
            df["NIT"] = df["NIT"].astype(str)
            df["NIT"] = df["NIT"].str.replace(r"\D", "", regex=True).str.strip()

            nit_to_email = dict(zip(df["NIT"], df.get("CORREO", "")))
            nit_to_name = dict(zip(df["NIT"], df.get("NOMBRE_PROVEEDOR", "")))

            enviados = 0

            for archivo in os.listdir(CARP_RESULTADOS):
                if not archivo.endswith(".docx"):
                    continue

                nit = os.path.splitext(archivo)[0]

                correos_raw = nit_to_email.get(nit, "")
                if not isinstance(correos_raw, str):
                    correos_raw = ""

                lista_correos = [c.strip() for c in correos_raw.split(";") if c.strip()]

                if not lista_correos:
                    print("FALLIDO:", nit, "-> Sin correos en Excel")
                    fallidos.append({"nit": nit, "archivo": archivo, "motivo": "Sin correo"})
                    continue

                # Agregar siempre copia a Cartera (CC)
                lista_cc = [CC_CARTERA]

                nombre = nit_to_name.get(nit, "Proveedor")

                body = f"""Buenos dias,
                
    Adjuntamos el comprobante de pago correspondiente al NIT {nit}.

    Quedamos atentos.
    cordialmente 
    Equipo de Cartera"""

                ruta_archivo = os.path.join(CARP_RESULTADOS, archivo)

                try:
                    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                    server.login(TU_CORREO, TU_PASSWORD)

                    msg = MIMEMultipart()
                    msg['From'] = TU_CORREO
                    msg['To'] = ", ".join(lista_correos)
                    msg['Cc'] = ", ".join(lista_cc)
                    # Si quieres usar copia oculta:
                    # msg['Bcc'] = ", ".join(lista_cc)
                    
                    msg['Subject'] = f"Comprobante de pago - {nombre} - NIT {nit}"
                    msg.attach(MIMEText(body, 'plain'))

                    with open(ruta_archivo, "rb") as adjunto:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(adjunto.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{archivo}"')
                        msg.attach(part)

                    server.sendmail(TU_CORREO, lista_correos + lista_cc, msg.as_string())
                    server.quit()

                    print(f"Enviado correctamente a {lista_correos} con copia a cartera")

                    enviados += 1
                    shutil.copy(ruta_archivo, os.path.join(CARP_ENVIADOS, archivo))

                except Exception as e:
                    fallidos.append({"nit": nit, "archivo": archivo, "motivo": str(e)})
                    print(f"ERROR enviando a {nit}: {e}")

            print("Envio terminado. Enviados correctamente:", enviados)

    except Exception as e:
        print("ERROR general en el envio:", e)

    # ============================================================
    # MOVER LOS QUE NO SE ENVIARON
    # ============================================================

    if fallidos:
        for f in fallidos:
            origen = os.path.join(CARP_RESULTADOS, f["archivo"])
            destino = os.path.join(CARP_NO_ENVIADOS, f["archivo"])
            if os.path.exists(origen):
                shutil.copy(origen, destino)
                print("No enviado -> movido a no_enviados: " + f["archivo"])

        print("Algunos comprobantes no se enviaron. Revisa la carpeta 'no_enviados'.")
    else:
        print("Todos los comprobantes fueron enviados correctamente.")


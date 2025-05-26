import os
import json
import pandas as pd
import tempfile
import zipfile
import time
import pythoncom
import win32com.client
import pdfplumber
import fitz  # PyMuPDF
from google import genai
from dotenv import load_dotenv
from PIL import Image
import re
from datetime import datetime
from pathlib import Path
import streamlit as st

# Cargar variables de entorno
env_path = Path(__file__).parent / ".env" if "__file__" in globals() else Path(".env")
load_dotenv(dotenv_path=env_path)

api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    raise ValueError(f"❌ ERROR: No se encontró la clave API en el archivo .env en {env_path.resolve()}")

client = genai.Client(api_key=api_key)

BATCH_SIZE = 10
MAX_RETRIES = 3

def limpiar_nombre_proveedor(file_name):
    nombre = os.path.splitext(file_name)[0]
    nombre = re.sub(r'_pagina_\d+', '', nombre)
    nombre = re.sub(r'[_\d]+', ' ', nombre)
    nombre = nombre.strip().title()
    return nombre

def extraer_archivos(archivo_zip):
    temp_dir = tempfile.mkdtemp()
    pdf_files, excel_files, image_files = [], [], []
    
    try:
        with zipfile.ZipFile(archivo_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
    except zipfile.BadZipFile:
        st.error("❌ ERROR: El archivo ZIP está corrupto o no es válido.")
        return [], [], [], temp_dir

    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
            elif file.endswith(('.xls', '.xlsx')):
                excel_files.append(os.path.join(root, file))
            elif file.endswith(('.png', '.jpg', '.jpeg')):
                image_files.append(os.path.join(root, file))

    return pdf_files, excel_files, image_files, temp_dir

def excel_a_pdf(ruta_excel, carpeta_salida):
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(ruta_excel)

        hoja = workbook.ActiveSheet or workbook.Sheets(workbook.Sheets.Count)

        hoja.PageSetup.Zoom = False
        hoja.PageSetup.FitToPagesWide = 1
        hoja.PageSetup.FitToPagesTall = False

        ruta_pdf = os.path.join(carpeta_salida, f"{os.path.basename(ruta_excel).replace('.xlsx', '.pdf')}")
        hoja.ExportAsFixedFormat(0, ruta_pdf)

        workbook.Close(False)
        excel.Quit()
        
        return ruta_pdf
    except Exception as e:
        st.warning(f"❌ ERROR al convertir {ruta_excel} a PDF: {e}")
        return None

def dividir_pdf_por_paginas(pdf_path):
    paginas = []
    doc = fitz.open(pdf_path)
    if len(doc) == 1:
        return [pdf_path]
    
    for i in range(len(doc)):
        temp_path = f"{pdf_path[:-4]}_pagina_{i+1}.pdf"
        new_pdf = fitz.open()
        new_pdf.insert_pdf(doc, from_page=i, to_page=i)
        new_pdf.save(temp_path)
        paginas.append(temp_path)
    return paginas

def generar_prompt(nombre_archivo_pdf: str, proveedor: str) -> str:
    reglas_especificas = {
        "raices.pdf": """
- Extrae datos de las columnas:
  - **Articulo**: Extrae el nombre completo del artículo.
  - **Precio por Cajón**: Extrae el precio correspondiente y añade al nombre del artículo el sufijo "por Cajón".
  - **Precio por Kilo**: Extrae el precio correspondiente y añade al nombre del artículo el sufijo "por Kilo".

- Si el nombre del artículo incluye un símbolo `$`, se debe generar **SIEMPRE** tres registros obligatorios:
  - Ejemplo: Para *Ajo GDE $300*:
    - `Ajo GDE $300 por Cajón`: Extrae el precio correspondiente o usa el precio detectado en el nombre si no está disponible.
    - `Ajo GDE $300 por Kilo`: Extrae el precio correspondiente o usa el precio detectado en el nombre si no está disponible.
    - `Ajo GDE $300 por Unidad`: Usa el precio indicado en el nombre del artículo (`$300`).
  - Los tres registros deben generarse, aunque algunos precios sean iguales o solo uno esté disponible.

- Para artículos sin símbolo `$`:
  - Renombra todos los registros con el sufijo correspondiente:
    - Si es un precio por kilo: Añade "por Kilo" al nombre.
    - Si es un precio por cajón: Añade "por Cajón" al nombre.
  - Si falta alguno de los precios (por kilo o cajón), genera el registro con el dato disponible.

- **Asegúrate de extraer todos los artículos presentes en el archivo, sin omitir ninguno.**
""",
        "bellapalta.pdf": """
- Usa únicamente la columna PRECIO POR KILO/UNI. Ignora otras columnas como Precio por Bulto.
- En la segunda página:
  - Si la primera tabla no tiene encabezados claros, extrae datos de la cuarta columna.
  - En la segunda tabla, extrae datos de la columna Precio Venta.
- No mezcles datos entre tablas diferentes; cada tabla debe procesarse por separado.
""",
        "bella_palta.pdf": """
- Usa únicamente la columna PRECIO POR KILO/UNI. Ignora otras columnas como Precio por Bulto.
- En la segunda página:
  - Si la primera tabla no tiene encabezados claros, extrae datos de la cuarta columna.
  - En la segunda tabla, extrae datos de la columna Precio Venta.
- No mezcles datos entre tablas diferentes; cada tabla debe procesarse por separado.
""",
        "soleil.pdf": """
- Usa únicamente la columna IVA INC. para el precio.
""",
        "le_soleil.pdf": """
- Usa únicamente la columna IVA INC. para el precio.
""",
        "delite.pdf": """
- Usa únicamente la columna PRECIO FINAL para el precio.
""",
        "delite_ofertas.pdf": """
- Usa únicamente la columna PRECIO FINAL para el precio.
""",
        "jumbalay.pdf": """
- Usa únicamente la columna 10% DESC + IVA para el precio.
"""
    }
    reglas = reglas_especificas.get(nombre_archivo_pdf.lower(), "- No hay reglas específicas para este archivo.\n")

    return f"""
Extrae la información de los archivos PDF indicados y devuélvela en formato JSON con los siguientes campos:

- Articulo: Nombre completo del artículo, especificando presentación o unidad si corresponde.
- Precio: Valor del precio con separador de miles (Ejemplo: 1.000 o 5.000.000).
- Proveedor: Asigna el nombre '{proveedor}' como proveedor.

---

Reglas generales para todos los archivos:
- Los encabezados del JSON deben ser exactamente: Articulo, Precio, Proveedor.
- Si el precio tiene descuentos, selecciona siempre el valor más bajo.
- Si hay varios precios disponibles, elige el que tenga IVA incluido.
- Corrige errores de formato en los precios.
- Reemplaza saltos de línea por espacios y elimina caracteres especiales.
- Extrae todos los artículos sin omitir ninguno, incluso si el precio está incompleto.

---

Ahora estás analizando el archivo: **{nombre_archivo_pdf}**

Aplica únicamente las siguientes reglas específicas:
{reglas}

---

Validación final:
- Todos los artículos deben estar presentes.
- Precios correctamente estructurados.
- El JSON final debe ser válido y sin errores antes de entregarlo.
""".strip()

def procesar_zip(archivo_zip):
    pdf_files, excel_files, image_files, temp_dir = extraer_archivos(archivo_zip)
    all_data = []

    for excel_file in excel_files:
        pdf_generado = excel_a_pdf(excel_file, temp_dir)
        if pdf_generado:
            pdf_files.append(pdf_generado)

    archivos_a_procesar = pdf_files + image_files
    for file_path in archivos_a_procesar:
        paginas_pdf = dividir_pdf_por_paginas(file_path) if file_path.endswith(".pdf") else [file_path]

        for pagina in paginas_pdf:
            file_name = os.path.basename(pagina)
            proveedor = limpiar_nombre_proveedor(file_name)
            st.info(f"📄 Procesando: {file_name} (Proveedor: {proveedor})")

            try:
                uploaded_file = client.files.upload(file=pagina, config={'display_name': file_name})
            except Exception as e:
                st.error(f"❌ ERROR al subir {file_name} a Gemini: {e}")
                continue

            prompt = generar_prompt(file_name, proveedor)

            try:
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=[prompt, uploaded_file],
                    config={"response_mime_type": "application/json"}
                )
                raw_response = response.candidates[0].content.parts[0].text
                json_data = json.loads(raw_response)
                all_data.extend(json_data)
            except Exception as e:
                st.error(f"❌ ERROR en la solicitud a Gemini para {file_name}: {e}")
                continue

    df = pd.DataFrame(all_data)
    df["Fecha"] = datetime.today().strftime("%d/%m/%Y")
    excel_path = os.path.join(temp_dir, "datos_extraidos.xlsx")
    df.to_excel(excel_path, index=False)

    return excel_path, df

# --- Interfaz Streamlit ---

st.title("📦 Procesador de ZIP con PDFs, Excels e Imágenes")
st.write("Sube un archivo ZIP que contenga documentos para procesar con Gemini.")

archivo_zip = st.file_uploader("Carga tu archivo ZIP aquí", type="zip")

if archivo_zip is not None:
    with st.spinner("Procesando archivo ZIP..."):
        try:
            temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            temp_zip.write(archivo_zip.read())
            temp_zip.close()

            excel_path, df_resultado = procesar_zip(temp_zip.name)

            st.success("✅ Procesamiento completado")
            st.write("Vista previa de los datos extraídos:")
            st.dataframe(df_resultado)

            with open(excel_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar Excel",
                    data=f,
                    file_name="datos_extraidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Ocurrió un error durante el procesamiento: {e}")

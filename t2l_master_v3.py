import os
import re
import time
import zipfile
from io import BytesIO

import streamlit as st
import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from PIL import Image 


# =========================================================
# FUNCIONES DE UTILIDAD Y PARSEO
# =========================================================

# ---------------------------------------------------------
# EXTRAER TEXTO DEL PDF (ADAPTADA: Usa objeto en memoria)
# ---------------------------------------------------------
def extract_text(uploaded_file_object):
    """Extrae texto de un objeto de archivo cargado por Streamlit (en memoria)."""
    text = ""
    with pdfplumber.open(uploaded_file_object) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                text += t + "\n"
    return text

# ---------------------------------------------------------
# PARSEAR BULTOS + KILOS (LÓGICA ORIGINAL)
# ---------------------------------------------------------
def parse_t2l(text):
    """Extrae bultos y kilos del texto del T2L."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    results = []

    i = 0
    while i < len(lines):
        line = lines[i]

        if "Number of Packages" in line:
            # BULTOS
            m_pkg = re.search(r"(\d+)", line)
            bultos = m_pkg.group(1) if m_pkg else ""

            # KILOS
            kilos = ""
            for j in range(i, min(i + 20, len(lines))):
                if "Gross Mass" in lines[j]:
                    if j + 1 < len(lines):
                        raw = lines[j + 1].strip()
                        m_g = re.search(r"([0-9.,]+)", raw)
                        if m_g:
                            v = m_g.group(1).replace(",", ".")
                            try:
                                f = float(v)
                                if f.is_integer():
                                    kilos = str(int(f)).replace(".", ",")
                                else:
                                    kilos = str(f).replace(".", ",")
                            except:
                                kilos = ""
                    break

            results.append((bultos, kilos))

        i += 1

    return results

# ---------------------------------------------------------
# GENERAR INFORME PDF (ADAPTADA: Usa buffer de BytesIO)
# ---------------------------------------------------------
def generar_informe_pdf(resumen, pdf_buffer, tiempo_total, logo_path=None):
    """Genera el informe PDF escribiendo directamente a un buffer de BytesIO."""
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4
    y = height - 50

    # Logo
    if logo_path and os.path.exists(logo_path):
        try:
            # Logo en el informe PDF
            c.drawImage(logo_path, 40, y - 40, width=160, height=40, preserveAspectRatio=True, mask='auto')
        except:
            pass

    y -= 70
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "INFORME PROCESAMIENTO T2L")
    y -= 25

    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"ñ Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
    y -= 20
    c.drawString(40, y, "© Departamento de Procesos | Bernardino Abad   Edition 2025")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Detalle por contenedor:")
    y -= 20

    c.setFont("Helvetica", 11)
    for cont, total in resumen.items():
        c.drawString(40, y, f"ñ {cont}   Total partidas: {total}") 
        y -= 18
        if y < 80:
            c.showPage()
            y = height - 80
            c.setFont("Helvetica", 11)

    y -= 10
    if y < 80:
        c.showPage()
        y = height - 80

    c.setFont("Helvetica-Oblique", 10)
    c.drawString(40, 60, "Firmado: Sistema Automatizado Departamento de Procesos | BA = ")
    c.save()

# ---------------------------------------------------------
# LIMPIEZA PARA CSV (Lógica original)
# ---------------------------------------------------------
def clean_int_str(x):
    """Limpia valores numéricos para exportación como enteros."""
    if x is None:
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s

def clean_kilos_str(x):
    """Limpia y formatea los kilos."""
    if x is None:
        return ""
    s = str(x).strip()
    if s == "":
        return ""
    s = s.replace(" ", "").replace(",", ".")
    try:
        v = float(s)
        if v.is_integer():
            return str(int(v))
        else:
            return str(v).replace(".", ",")
    except:
        return ""

# =========================================================
# FUNCIONES DE PROCESAMIENTO CENTRAL
# =========================================================

# ---------------------------------------------------------
# PROCESAR ARCHIVOS T2L Y GENERAR EXCEL/PDF (ADAPTADA)
# ---------------------------------------------------------
def procesar_t2l_streamlit(uploaded_files, sumaria, logo_path=None):
    """Procesa PDFs cargados y genera Excel y Informe PDF en BytesIO."""
    excel_output = BytesIO()
    writer = pd.ExcelWriter(excel_output, engine="openpyxl")

    resumen = {}
    t_inicio = time.time()

    for pdf_file_obj in uploaded_files:
        pdf_file_obj.seek(0) # Aseguramos que el puntero está al inicio antes de leer
        pdf_filename = pdf_file_obj.name
        
        # Contenedor desde nombre de archivo
        m_cont = re.search(r"([A-Z]{4}\d{6})", pdf_filename)
        cont = m_cont.group(1) if m_cont else "SINCONT"

        text = extract_text(pdf_file_obj)
        rows_raw = parse_t2l(text)

        final_rows = []
        sec = 1
        for b, k in rows_raw:
            final_rows.append({
                "Bultos": b,
                "Kilos": k,
                "Fijo_col3": 1,
                "Fijo_col4": "RECEPCION T2L",
                "Vacio5": "",
                "Vacio6": "",
                "Fijo_col7": "3401110000",
                "Fijo_col8": 1,
                "Contenedor": cont,
                "Fijo_col10": "ES",
                "Vacio11": "",
                "Sumaria": sumaria,
                "Orden": sec
            })
            sec += 1

        if not final_rows:
            df = pd.DataFrame([{
                "Bultos": "", "Kilos": "", "Fijo_col3": "", "Fijo_col4": "SIN PARTIDAS",
                "Vacio5": "", "Vacio6": "", "Fijo_col7": "", "Fijo_col8": "",
                "Contenedor": cont, "Fijo_col10": "", "Vacio11": "", "Sumaria": sumaria, "Orden": ""
            }])
            resumen[cont] = 0
        else:
            df = pd.DataFrame(final_rows)

            # Sumatorios
            try:
                total_bultos = sum(int(b or "0") for b, _ in rows_raw)
            except:
                total_bultos = ""

            try:
                total_kilos_val = sum(
                    float((k or "0").replace(",", "."))
                    for _, k in rows_raw
                )
                if total_kilos_val.is_integer():
                    total_kilos = str(int(total_kilos_val)).replace(".", ",")
                else:
                    total_kilos = str(total_kilos_val).replace(".", ",")
            except:
                total_kilos = ""

            df.loc[len(df)] = [
                total_bultos, total_kilos, "", "TOTAL", "", "", "", "", "", "", "", "", ""
            ]

            resumen[cont] = len(rows_raw)

        df.to_excel(writer, sheet_name=pdf_filename[:31], index=False)

    writer.close()
    excel_output.seek(0)
    excel_bytes = excel_output.read()

    t_total = time.time() - t_inicio

    # Generar Informe PDF en buffer
    pdf_buffer = BytesIO()
    # Usamos la ruta del logo para el informe PDF
    generar_informe_pdf(resumen, pdf_buffer, t_total, logo_path="imagen.png") 
    pdf_buffer.seek(0)
    informe_pdf_bytes = pdf_buffer.read()

    return excel_bytes, informe_pdf_bytes, t_total

# ---------------------------------------------------------
# GENERAR TXT INDIVIDUALES (ADAPTACIÓN DE GENERAR_ZIP_CSV)
# ---------------------------------------------------------
def generar_txt_en_memoria(uploaded_excel_file):
    """Lee el Excel revisado (en memoria) y genera un diccionario de archivos TXT."""
    
    uploaded_excel_file.seek(0) # Aseguramos que el puntero está al inicio antes de leer
    excel = pd.ExcelFile(uploaded_excel_file)
    txt_files = {}
    
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet, dtype=str)

        if df.empty:
            continue

        # quitar última fila (TOTAL)
        if len(df) > 1 and df.iloc[-1].astype(str).str.contains('TOTAL', na=False).any():
            df = df.iloc[:-1].copy()
        else:
            df = df.copy()

        # limpiar columnas numéricas
        for col in ["Bultos", "Fijo_col3", "Fijo_col7", "Fijo_col8", "Sumaria", "Orden"]:
            if col in df.columns:
                df[col] = df[col].map(clean_int_str)

        if "Kilos" in df.columns:
            df["Kilos"] = df["Kilos"].map(clean_kilos_str)

        # Generar el contenido TXT (separado por punto y coma, sin encabezado)
        txt_buffer = BytesIO()
        df.to_csv(txt_buffer, sep=";", header=False, index=False, encoding="utf-8")
        txt_buffer.seek(0)
        
        # Almacenar los bytes del archivo TXT con el nombre de la hoja
        txt_files[sheet] = txt_buffer.read()
        
    return txt_files

# =========================================================
# INTERFAZ DE STREAMLIT (ADAPTADA ESTÉTICAMENTE A DUA)
# =========================================================
def main_streamlit_app():
    # CAMBIO CLAVE: layout="wide" para usar todo el ancho de la página
    st.set_page_config(
        page_title="Procesador T2L | BA",
        page_icon="icono.ico",
        layout="wide",
    )
    
    # --- ENCABEZADO ALINEADO A LA IZQUIERDA DEL CONTENEDOR ANCHO (Estilo DUA) ---
    
    # Encabezado con logo
    try:
        logo = Image.open("imagen.png")
        st.image(logo, width=350) # Ajustado a 350px para mejor coincidencia
    except FileNotFoundError:
        st.error("No se encontró el archivo de imagen. Asegúrate de que 'imagen.png' esté en la raíz del repositorio.")

    st.markdown(
        "<h3 style='color:#132136;margin-top:-10px;'>Procesador T2L | PDF → Excel / CSV</h3>",
        unsafe_allow_html=True
    )
    
    # Subtítulo (pequeño)
    st.caption("Departamento de Aduanas - Bernardino Abad SL")
    
    # Separador que usa la aplicación DUA
    st.divider() 
    
    st.write("Siga los pasos para la extracción, revisión y exportación de partidas T2L.")

    # CSS personalizado (Lo mantenemos para forzar el primaryColor si config.toml falla)
    st.markdown("""
<style>
.stApp { background-color: #F8FAFD; } 
.stButton>button { background-color: #004C91; color: white; border-radius: 8px; padding: 0.6em 1.2em; font-weight: 600; }
h1, h2, h3, h4 { color: #004C91; }
</style>
""", unsafe_allow_html=True)

    
    # Estado de la sesión para manejar los pasos
    if 'excel_bytes' not in st.session_state:
        st.session_state.excel_bytes = None
        st.session_state.informe_pdf_bytes = None
        st.session_state.t_total = None
        st.session_state.txt_files = None # Inicializamos la variable para los TXT

    # --- 1. Cargar Archivos T2L y Sumaria ---
    st.subheader("1. Cargar PDFs T2L y Nº de Sumaria")
    
    # Uso de columnas para organizar el input de Sumaria y el uploader
    col_sumaria, col_uploader = st.columns([1, 2])
    
    with col_sumaria:
        # Input para la Sumaria
        sumaria = st.text_input("Nº de Sumaria (11 dígitos):", max_chars=11, key="sumaria_input")
    
    with col_uploader:
        # Carga de múltiples archivos PDF
        uploaded_files = st.file_uploader(
            "Sube todos los **PDFs T2L**",
            type=["pdf"],
            accept_multiple_files=True
        )
    
    # Botón de procesamiento (PASO 1)
    if st.button("Procesar Archivos T2L", type="primary", use_container_width=True):
        
        # Validaciones
        if not sumaria.isdigit() or len(sumaria) != 11:
            st.error("El Nº de Sumaria debe tener exactamente 11 dígitos numéricos.")
            return
        if not uploaded_files:
            st.warning("Por favor, sube los archivos PDF a procesar.")
            return

        # Limpiamos el estado previo
        st.session_state.excel_bytes = None
        st.session_state.informe_pdf_bytes = None
        st.session_state.t_total = None
        st.session_state.txt_files = None

        with st.spinner('ñ Procesando T2L, por favor espera...'):
            excel_bytes, informe_pdf_bytes, t_total = procesar_t2l_streamlit(
                uploaded_files, sumaria, "imagen.png" # Pasamos la ruta del logo para el informe PDF
            )

        if excel_bytes and informe_pdf_bytes:
            st.session_state.excel_bytes = excel_bytes
            st.session_state.informe_pdf_bytes = informe_pdf_bytes
            st.session_state.t_total = t_total
            st.success(f"✅ Archivos base generados. Tiempo total: **{t_total:.2f}** s")
        else:
            st.error("L No se pudo procesar la carpeta seleccionada.")


    # --- 2. Revisar y Ajustar Excel (Paso 2) ---
    if st.session_state.excel_bytes:
        st.markdown("---")
        st.subheader("2. Descargar y Revisar Excel Base")
        st.info("Descarga el Excel de trabajo, revisa y ajusta los datos, **guárdalo** y vuelve para subirlo.")
        
        col_dl_excel, col_dl_pdf = st.columns(2)
        
        with col_dl_excel:
            st.download_button(
                label="⬇️ Descargar Excel base (T2L_RESULTADO.xlsx)",
                data=st.session_state.excel_bytes,
                file_name="T2L_RESULTADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Excel con todas las partidas extraídas."
            )
        with col_dl_pdf:
            st.download_button(
                label="⬇️ Descargar Informe PDF",
                data=st.session_state.informe_pdf_bytes,
                file_name="INFORME_T2L.pdf",
                mime="application/pdf",
                help=f"Informe de proceso ({st.session_state.t_total:.2f} s)."
            )
        
        # --- 3. Cargar Excel Revisado y Generar TXT (Paso 3) ---
        st.markdown("---")
        st.subheader("3. Subir Excel Revisado y Generar Archivos TXT")
        
        revised_excel = st.file_uploader(
            "Sube aquí el **Excel revisado** (el mismo archivo después de tus cambios)",
            type=["xlsx", "xls"],
            key="revised_excel_uploader"
        )
        
        if revised_excel:
            if st.button("🎉 Generar Archivos TXT", type="secondary", use_container_width=True):
                with st.spinner('ñ Generando ficheros TXT...'):
                    
                    txt_files = generar_txt_en_memoria(revised_excel)
                    
                    if txt_files:
                        st.session_state.txt_files = txt_files
                        st.session_state.txt_count = len(txt_files)
                        st.success(f"✅ Proceso Completado. Se generaron {len(txt_files)} ficheros TXT.")
                    else:
                        st.error("No se pudieron generar archivos TXT. Verifica el formato del Excel revisado.")

        
    # --- 4. Descargar Resultados Finales Individuales (Paso Final) ---
    if 'txt_files' in st.session_state and st.session_state.txt_files:
        st.markdown("---")
        st.subheader("4. Descargar Archivos TXT Individuales")
        
        # Creamos columnas para organizar los botones de descarga
        cols = st.columns(3)
        i = 0
        
        for name, txt_bytes in st.session_state.txt_files.items():
            
            with cols[i % 3]: # Rota entre 3 columnas
                # Se eliminan los caracteres '=Á' del botón de descarga
                st.download_button(
                    label=f"⬇️ Descargar {name}.txt", 
                    data=txt_bytes,
                    file_name=f"{name}.txt",
                    mime="text/plain",
                    help=f"Contiene las partidas para el contenedor {name}.",
                    type="primary",
                    key=f"dl_txt_{name}" # Clave única para cada botón
                )
            i += 1

        st.info(f"Proceso terminado. Se generaron {st.session_state.txt_count} archivos TXT. Puedes empezar un nuevo proceso recargando la página o cambiando los inputs.")

if __name__ == "__main__":
    main_streamlit_app()

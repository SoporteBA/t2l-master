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

def extract_text(uploaded_file_object):
    """Extrae texto de un objeto de archivo cargado por Streamlit (en memoria)."""
    text = ""
    uploaded_file_object.seek(0) 
    with pdfplumber.open(uploaded_file_object) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                text += t + "\n"
    return text

def parse_t2l(text):
    """
    Extrae bultos y kilos de forma robusta para documentos multi-ítem.
    Busca todos los bultos y todos los kilos por separado y los empareja.
    """
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    
    # 1. Extraer TODOS los bultos (Number of Packages)
    bultos_encontrados = []
    for line in lines:
        if "Number of Packages:" in line:
            # Captura el número después de los dos puntos
            m = re.search(r"Number of Packages:\s*(\d+)", line)
            if m:
                bultos_encontrados.append(m.group(1))

    # 2. Extraer TODOS los kilos (Gross Mass)
    kilos_encontrados = []
    for i, line in enumerate(lines):
        if "35 Gross Mass (kg)" in line:
            # Buscamos el valor en la línea actual o en la siguiente (por si hay salto de línea)
            texto_busqueda = line
            if i + 1 < len(lines):
                texto_busqueda += " " + lines[i+1]
            
            # Busca un patrón numérico tipo 4520.0 o 4520,0
            m = re.search(r"(\d+[\.,]\d+)", texto_busqueda)
            if m:
                # Normalizamos a punto decimal para el DataFrame
                valor = m.group(1).replace(",", ".")
                kilos_encontrados.append(valor)

    # 3. Emparejar por posición
    results = []
    # Usamos el máximo de ambos para no perder datos si uno de los dos falla
    max_items = max(len(bultos_encontrados), len(kilos_encontrados))
    
    for idx in range(max_items):
        b = bultos_encontrados[idx] if idx < len(bultos_encontrados) else ""
        k = kilos_encontrados[idx] if idx < len(kilos_encontrados) else ""
        results.append((b, k))

    return results

def generar_informe_pdf(resumen, pdf_buffer, tiempo_total, logo_path=None):
    """Genera el informe PDF escribiendo directamente a un buffer de BytesIO."""
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4
    y = height - 50

    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, 40, y - 40, width=160, height=40, preserveAspectRatio=True, mask='auto')
        except:
            pass

    y -= 70
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, y, "INFORME PROCESAMIENTO T2L")
    y -= 25

    c.setFont("Helvetica", 10)
    c.drawString(40, y, f" Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
    y -= 20
    c.drawString(40, y, "© Departamento de Procesos | Bernardino Abad   Edition 2025")
    y -= 30

    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Detalle por contenedor:")
    y -= 20

    c.setFont("Helvetica", 11)
    for cont, total in resumen.items():
        c.drawString(40, y, f"- {cont}   Total partidas: {total}") 
        y -= 18
        if y < 80:
            c.showPage()
            y = height - 80
            c.setFont("Helvetica", 11)

    c.setFont("Helvetica-Oblique", 10)
    c.drawString(40, 60, "Firmado: Sistema Automatizado Departamento de Procesos | BA")
    c.save()

def clean_int_str(x):
    if x is None: return ""
    s = str(x).strip()
    if s == "" or s == "nan": return ""
    if s.endswith(".0"): s = s[:-2]
    return s

def clean_kilos_str(x):
    if x is None: return ""
    s = str(x).strip()
    if s == "" or s == "nan": return ""
    s = s.replace(" ", "")
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

def procesar_t2l_streamlit(uploaded_files, sumaria, logo_path=None):
    excel_output = BytesIO()
    writer = pd.ExcelWriter(excel_output, engine="openpyxl")

    resumen = {}
    t_inicio = time.time()
    CONTAINER_PATTERN = r"([A-Z]{4}\d{7})" 

    for pdf_file_obj in uploaded_files:
        pdf_file_obj.seek(0)
        pdf_filename = pdf_file_obj.name
        
        m_cont = re.search(CONTAINER_PATTERN, pdf_filename) 
        cont = m_cont.group(1) if m_cont else "SINCONT"

        text = extract_text(pdf_file_obj)
        rows_raw = parse_t2l(text) 

        final_rows = []
        sec = 1
        for b, k in rows_raw:
            final_rows.append({
                "Bultos": b,
                "Kilos": k,
                "Fijo_col3": "1",
                "Fijo_col4": "RECEPCION T2L",
                "Vacio5": "",
                "Vacio6": "",
                "Fijo_col7": "3401110000",
                "Fijo_col8": "1",
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
            try:
                total_bultos = sum(int(b or "0") for b, _ in rows_raw)
                total_kilos_val = sum(float(k or "0") for _, k in rows_raw)
                total_kilos = str(int(total_kilos_val)) if total_kilos_val.is_integer() else str(total_kilos_val)
            except:
                total_bultos, total_kilos = "", ""

            df["Kilos"] = pd.to_numeric(df["Kilos"], errors='coerce') 
            df["Bultos"] = pd.to_numeric(df["Bultos"], errors='coerce').astype('Int64')
            df["Orden"] = pd.to_numeric(df["Orden"], errors='coerce').astype('Int64')
            
            df.loc[len(df)] = [total_bultos, total_kilos, "", "TOTAL", "", "", "", "", "", "", "", "", ""]
            resumen[cont] = len(rows_raw)

        df.to_excel(writer, sheet_name=pdf_filename[:31], index=False)

    writer.close()
    excel_output.seek(0)
    excel_bytes = excel_output.read()

    t_total = time.time() - t_inicio
    pdf_buffer = BytesIO()
    generar_informe_pdf(resumen, pdf_buffer, t_total, logo_path="imagen.png") 
    pdf_buffer.seek(0)
    informe_pdf_bytes = pdf_buffer.read()

    return excel_bytes, informe_pdf_bytes, t_total

def generar_txt_en_memoria(uploaded_excel_file):
    uploaded_excel_file.seek(0)
    excel = pd.ExcelFile(uploaded_excel_file)
    txt_files = {}
    
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet, dtype=str)
        if df.empty: continue
        if len(df) > 1 and df.iloc[-1].astype(str).str.contains('TOTAL', na=False).any():
            df = df.iloc[:-1].copy()
        else:
            df = df.copy()

        for col in ["Bultos", "Fijo_col3", "Fijo_col7", "Fijo_col8", "Sumaria", "Orden"]:
            if col in df.columns: df[col] = df[col].map(clean_int_str)
        if "Kilos" in df.columns: df["Kilos"] = df["Kilos"].map(clean_kilos_str) 

        txt_buffer = BytesIO()
        df.to_csv(txt_buffer, sep=";", header=False, index=False, encoding="utf-8")
        txt_buffer.seek(0)
        txt_files[sheet] = txt_buffer.read()
    return txt_files

def generar_zip_desde_txt(txt_files):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename, content in txt_files.items():
            zf.writestr(f"{filename}.txt", content)
    zip_buffer.seek(0)
    return zip_buffer, len(txt_files)

# =========================================================
# INTERFAZ DE STREAMLIT
# =========================================================
def main_streamlit_app():
    st.set_page_config(page_title="Procesador T2L | BA", page_icon="icono.ico", layout="wide")
    
    try:
        logo = Image.open("imagen.png")
        st.image(logo, width=350)
    except:
        st.error("No se encontró 'imagen.png'")

    st.markdown("<h3 style='color:#132136;margin-top:-10px;'>Procesador T2L | PDF → Excel / CSV</h3>", unsafe_allow_html=True)
    st.caption("Departamento de Aduanas - Bernardino Abad SL")
    st.divider() 

    if 'excel_bytes' not in st.session_state:
        st.session_state.excel_bytes = None
        st.session_state.informe_pdf_bytes = None
        st.session_state.txt_files = None

    st.subheader("1. Cargar PDFs T2L y Nº de Sumaria")
    col_sumaria, col_uploader = st.columns([1, 2])
    
    with col_sumaria:
        sumaria = st.text_input("Nº de Sumaria (11 dígitos):", max_chars=11)
    
    with col_uploader:
        uploaded_files = st.file_uploader("Sube los PDFs T2L", type=["pdf"], accept_multiple_files=True)
    
    if st.button("Procesar Archivos T2L", type="primary", use_container_width=True):
        if not sumaria.isdigit() or len(sumaria) != 11:
            st.error("La Sumaria debe tener 11 dígitos.")
            return
        if not uploaded_files:
            st.warning("Sube archivos PDF.")
            return

        with st.spinner('Procesando...'):
            excel_bytes, informe_pdf_bytes, t_total = procesar_t2l_streamlit(uploaded_files, sumaria, "imagen.png")
            if excel_bytes:
                st.session_state.excel_bytes = excel_bytes
                st.session_state.informe_pdf_bytes = informe_pdf_bytes
                st.session_state.t_total = t_total
                st.success("✅ Procesado correctamente.")

    if st.session_state.excel_bytes:
        st.divider()
        st.subheader("2. Descargar y Revisar Excel")
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Descargar Excel", st.session_state.excel_bytes, "T2L_RESULTADO.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("⬇️ Descargar Informe PDF", st.session_state.informe_pdf_bytes, "INFORME_T2L.pdf", "application/pdf")
        
        st.divider()
        st.subheader("3. Subir Excel Revisado")
        revised_excel = st.file_uploader("Sube el Excel corregido", type=["xlsx"], key="rev_excel")
        
        if revised_excel:
            if st.button("🎉 Generar Archivos TXT", use_container_width=True):
                txt_files = generar_txt_en_memoria(revised_excel)
                if txt_files:
                    st.session_state.txt_files = txt_files
                    st.success(f"✅ Se generaron {len(txt_files)} ficheros.")

    if st.session_state.get('txt_files'):
        st.divider()
        st.subheader("4. Descargar Resultados")
        zip_buf, count = generar_zip_desde_txt(st.session_state.txt_files)
        st.download_button(f"⬇️ Descargar los {count} TXT (ZIP)", zip_buf, "Archivos_T2L.zip", "application/zip", type="primary")
        
        cols = st.columns(3)
        for i, (name, content) in enumerate(st.session_state.txt_files.items()):
            cols[i % 3].download_button(f"⬇️ {name}.txt", content, f"{name}.txt", "text/plain")

if __name__ == "__main__":
    main_streamlit_app()

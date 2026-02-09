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
# FUNCIONES DE UTILIDAD Y PARSEO (ROBUSTEZ MEJORADA)
# =========================================================

def extract_text(uploaded_file_object):
    """Extrae texto de un PDF. Devuelve string vacío si falla o es imagen."""
    text = ""
    try:
        uploaded_file_object.seek(0) 
        with pdfplumber.open(uploaded_file_object) as pdf:
            for p in pdf.pages:
                t = p.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        st.error(f"Error leyendo PDF: {e}")
    return text

def parse_t2l(text):
    """
    Parser robusto con Regex.
    Busca patrones de bultos y kilos sin importar mayúsculas o espacios extra.
    """
    if not text.strip():
        return []

    lines = [l.strip() for l in text.split("\n") if l.strip()]
    
    # 1. Extraer TODOS los bultos (Pattern: "Number of Packages" + ":" + número)
    # Es flexible con espacios y mayúsculas
    bultos_encontrados = []
    pattern_bultos = re.compile(r"Number\s+of\s+Packages\s*[:\-]?\s*(\d+)", re.IGNORECASE)
    
    for line in lines:
        m = pattern_bultos.search(line)
        if m:
            bultos_encontrados.append(m.group(1))

    # 2. Extraer TODOS los kilos (Pattern: "35 Gross Mass" + número)
    # Busca en la línea actual y en la siguiente para cubrir saltos de línea
    kilos_encontrados = []
    pattern_label_kilos = re.compile(r"35\s+Gross\s+Mass", re.IGNORECASE)
    pattern_valor_num = re.compile(r"(\d+[\.,]\d+)") # Busca 1234.56 o 1234,56

    for i, line in enumerate(lines):
        if pattern_label_kilos.search(line):
            # Combinamos con la línea siguiente para asegurar que pillamos el valor
            texto_a_buscar = line
            if i + 1 < len(lines):
                texto_a_buscar += " " + lines[i+1]
            
            m_valor = pattern_valor_num.search(texto_a_buscar)
            if m_valor:
                valor = m_valor.group(1).replace(",", ".")
                kilos_encontrados.append(valor)

    # 3. Emparejar por posición (Asegura que si hay 6 bultos y 6 pesos, se unan bien)
    results = []
    num_items = max(len(bultos_encontrados), len(kilos_encontrados))
    
    for idx in range(num_items):
        b = bultos_encontrados[idx] if idx < len(bultos_encontrados) else ""
        k = kilos_encontrados[idx] if idx < len(kilos_encontrados) else ""
        results.append((b, k))

    return results

# --- (Resto de funciones de informe y limpieza se mantienen estables) ---

def generar_informe_pdf(resumen, pdf_buffer, tiempo_total, logo_path=None):
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    width, height = A4
    y = height - 50
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, 40, y - 40, width=160, height=40, preserveAspectRatio=True, mask='auto')
        except: pass
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
    c.save()

def clean_int_str(x):
    s = str(x).strip()
    if s in ["", "nan", "None"]: return ""
    if s.endswith(".0"): s = s[:-2]
    return s

def clean_kilos_str(x):
    s = str(x).strip().replace(" ", "")
    if s in ["", "nan", "None"]: return ""
    try:
        v = float(s)
        return str(int(v)) if v.is_integer() else str(v).replace(".", ",") 
    except: return ""

# =========================================================
# PROCESAMIENTO CENTRAL
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
        
        # Alerta si el PDF no tiene texto
        if not text.strip():
            st.warning(f"⚠️ El archivo '{pdf_filename}' parece no tener texto legible (¿es un escaneo?).")

        rows_raw = parse_t2l(text) 
        final_rows = []
        sec = 1
        for b, k in rows_raw:
            final_rows.append({
                "Bultos": b, "Kilos": k, "Fijo_col3": "1", "Fijo_col4": "RECEPCION T2L",
                "Vacio5": "", "Vacio6": "", "Fijo_col7": "3401110000", "Fijo_col8": "1",
                "Contenedor": cont, "Fijo_col10": "ES", "Vacio11": "", "Sumaria": sumaria, "Orden": sec
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
            # Totales para la fila de control en Excel
            try:
                total_bultos = sum(int(b or 0) for b, _ in rows_raw if str(b).isdigit())
                total_kilos_val = sum(float(k or 0) for _, k in rows_raw)
                total_kilos = str(int(total_kilos_val)) if total_kilos_val.is_integer() else str(total_kilos_val)
            except: total_bultos, total_kilos = "", ""

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
    return excel_bytes, pdf_buffer.read(), t_total

def generar_txt_en_memoria(uploaded_excel_file):
    uploaded_excel_file.seek(0)
    excel = pd.ExcelFile(uploaded_excel_file)
    txt_files = {}
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet, dtype=str)
        if df.empty: continue
        if len(df) > 1 and df.iloc[-1].astype(str).str.contains('TOTAL', na=False).any():
            df = df.iloc[:-1].copy()
        for col in ["Bultos", "Fijo_col3", "Fijo_col7", "Fijo_col8", "Sumaria", "Orden"]:
            if col in df.columns: df[col] = df[col].map(clean_int_str)
        if "Kilos" in df.columns: df["Kilos"] = df["Kilos"].map(clean_kilos_str) 
        buf = BytesIO()
        df.to_csv(buf, sep=";", header=False, index=False, encoding="utf-8")
        buf.seek(0)
        txt_files[sheet] = buf.read()
    return txt_files

def generar_zip_desde_txt(txt_files):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename, content in txt_files.items():
            zf.writestr(f"{filename}.txt", content)
    zip_buffer.seek(0)
    return zip_buffer, len(txt_files)

# =========================================================
# INTERFAZ STREAMLIT
# =========================================================
def main_streamlit_app():
    st.set_page_config(page_title="Procesador T2L | BA", page_icon="favicon_t2l.png", layout="wide")
    
    try:
        st.image(Image.open("imagen.png"), width=350)
    except: pass

    st.markdown("<h3 style='color:#132136;'>Procesador T2L | PDF → Excel / CSV</h3>", unsafe_allow_html=True)
    st.caption("Bernardino Abad SL | Aduanas")
    st.divider() 

    if 'excel_bytes' not in st.session_state:
        st.session_state.update({'excel_bytes': None, 'pdf_bytes': None, 'txt_files': None})

    col_sum, col_up = st.columns([1, 2])
    sumaria = col_sum.text_input("Nº de Sumaria (11 dígitos):", max_chars=11)
    uploaded_files = col_up.file_uploader("Sube los PDFs T2L", type=["pdf"], accept_multiple_files=True)
    
    if st.button("Procesar Archivos T2L", type="primary", use_container_width=True):
        if len(sumaria) != 11 or not uploaded_files:
            st.error("Revisa la Sumaria (11 dígitos) y sube los archivos.")
        else:
            with st.spinner('Analizando documentos...'):
                ex, pdf, t = procesar_t2l_streamlit(uploaded_files, sumaria, "imagen.png")
                st.session_state.update({'excel_bytes': ex, 'pdf_bytes': pdf, 't': t})
                st.success(f"✅ Procesado en {t:.2f}s")

    if st.session_state.excel_bytes:
        st.divider()
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Descargar Excel", st.session_state.excel_bytes, "T2L_RESULTADO.xlsx")
        c2.download_button("⬇️ Descargar Informe PDF", st.session_state.pdf_bytes, "INFORME_T2L.pdf")
        
        rev_ex = st.file_uploader("Sube el Excel revisado", type=["xlsx"], key="rev")
        if rev_ex and st.button("🎉 Generar Archivos TXT", use_container_width=True):
            st.session_state.txt_files = generar_txt_en_memoria(rev_ex)
            st.success("Ficheros listos.")

    if st.session_state.get('txt_files'):
        st.divider()
        z_buf, z_count = generar_zip_desde_txt(st.session_state.txt_files)
        st.download_button(f"⬇️ Descargar {z_count} TXT (ZIP)", z_buf, "Archivos_T2L.zip", type="primary")
        
        cols = st.columns(3)
        for i, (name, content) in enumerate(st.session_state.txt_files.items()):
            cols[i % 3].download_button(f"📄 {name}.txt", content, f"{name}.txt")

if __name__ == "__main__":
    main_streamlit_app()

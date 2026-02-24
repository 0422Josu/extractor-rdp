import streamlit as st
import pandas as pd
from docx import Document
import re
import io
from datetime import datetime
import zipfile

st.set_page_config(page_title="Extractor RDP Pro v5.2", page_icon="⚡", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
.stApp { font-family: 'Poppins', sans-serif; background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%); }
.header-container { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 40px; border-radius: 20px; margin-bottom: 30px; box-shadow: 0 10px 40px rgba(30, 60, 114, 0.3); }
.header-title { color: white; font-size: 48px; font-weight: 700; margin-bottom: 10px; }
.header-subtitle { color: #a8c0ff; font-size: 20px; }
.metric-card { background: white; padding: 25px; border-radius: 15px; box-shadow: 0 5px 20px rgba(0,0,0,0.08); text-align: center; transition: transform 0.3s ease; }
.metric-card:hover { transform: translateY(-5px); }
.metric-number { font-size: 42px; font-weight: 700; color: #1e3c72; }
.metric-label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 1px; margin-top: 5px; }
.upload-zone { background: linear-gradient(145deg, #f8f9ff 0%, #e8f0fe 100%); border: 3px dashed #2a5298; border-radius: 20px; padding: 60px; text-align: center; transition: all 0.3s ease; }
.upload-zone:hover { background: linear-gradient(145deg, #e8f0fe 0%, #d0e3ff 100%); border-color: #1e3c72; }
.stButton>button { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; border: none; padding: 18px 50px; font-size: 20px; font-weight: 600; border-radius: 50px; box-shadow: 0 5px 20px rgba(30, 60, 114, 0.4); }
.stButton>button:hover { transform: translateY(-3px); box-shadow: 0 10px 30px rgba(30, 60, 114, 0.5); }
.sidebar-content { background: linear-gradient(180deg, #1e3c72 0%, #2a5298 100%); border-radius: 15px; padding: 25px; color: white; }
.dataframe th { background: #1e3c72 !important; color: white !important; padding: 15px !important; }
.dataframe td { padding: 12px !important; }
.footer { background: #1e3c72; color: white; padding: 25px; border-radius: 15px; text-align: center; margin-top: 50px; }
.big-icon { font-size: 70px; margin-bottom: 15px; }
.success-banner { background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; padding: 20px; border-radius: 15px; text-align: center; font-size: 20px; font-weight: 600; margin: 20px 0; }
</style>
""", unsafe_allow_html=True)

def limpiar_valor(texto):
    if not texto: return ""
    return texto.replace('\n', ' ').strip().replace(',', '')

def formatear_numero(valor_str):
    valor_str = valor_str.upper()
    numero_match = re.search(r'([0-9.]+)', valor_str)
    if not numero_match: return "0"
    numero = float(numero_match.group(1))
    if "KV" in valor_str: numero = int(numero * 1000)
    else: numero = int(numero)
    return str(numero)

def extraer_datos_archivo(doc_file, nombre_archivo):
    info = {"sub": "", "bay": "", "rel": "", "mar": "", "fec": "", "ord": "", "ser": "", "val_i": "", "val_v": "", "rechazo": "", "tc": "", "tv": "", "archivo": nombre_archivo}
    v_prim, v_sec, i_prim, i_sec = "", "", "", ""
    try:
        doc = Document(doc_file)
        texto_completo = ""
        for tabla in doc.tables:
            for fila in tabla.rows:
                celdas = [c.text.strip() for c in fila.cells]
                texto_fila = " ".join(celdas)
                texto_completo += texto_fila + " "
                for i, texto_celda in enumerate(celdas):
                    if "Substation:" == texto_celda and i + 1 < len(celdas): info["sub"] = limpiar_valor(celdas[i+1])
                    elif "Bay:" == texto_celda and i + 1 < len(celdas): info["bay"] = limpiar_valor(celdas[i+1])
                    elif "Name/description:" == texto_celda and i + 1 < len(celdas): info["rel"] = limpiar_valor(celdas[i+1])
                    elif "Manufacturer:" == texto_celda and i + 1 < len(celdas): info["mar"] = limpiar_valor(celdas[i+1])
                    elif "Serial/model number:" == texto_celda and i + 1 < len(celdas): info["ser"] = limpiar_valor(celdas[i+1])
                    elif "Additional info 1:" == texto_celda and i + 1 < len(celdas): info["ord"] = limpiar_valor(celdas[i+1])
                    if "V nom (secondary):" in texto_celda and i + 1 < len(celdas): v_sec = formatear_numero(celdas[i+1])
                    if "V primary:" in texto_celda and i + 1 < len(celdas): v_prim = formatear_numero(celdas[i+1])
                    if "I nom (secondary):" in texto_celda and i + 1 < len(celdas): i_sec = formatear_numero(celdas[i+1])
                    if "I primary:" in texto_celda and i + 1 < len(celdas): i_prim = formatear_numero(celdas[i+1])
        if v_prim and v_sec: info["tv"] = f"{v_prim}/{v_sec}"
        if i_prim and i_sec: info["tc"] = f"{i_prim}/{i_sec}"
        if "RECHAZO DE CARGA" in texto_completo.upper(): info["rechazo"] = "SI"
        f_m = re.search(r'(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})', texto_completo)
        info["fec"] = f_m.group(1).replace('.', '/') if f_m else ""
        i_m = re.findall(r'I\s*L\d\s*[|]?\s*([1-9]\d*\.?\d*)\s*A', texto_completo, re.I)
        if i_m: info["val_i"] = i_m[0]
        v_m = re.findall(r'V\s*L\d-?E?\s*[|]?\s*([1-9]\d*\.?\d*)\s*V', texto_completo, re.I)
        if v_m: info["val_v"] = v_m[0]
    except Exception as e: st.error(f"Error: {str(e)}")
    return info

def buscar_archivos_en_zip(zip_file):
    archivos = []
    try:
        with zipfile.ZipFile(zip_file, 'r') as z:
            for nombre in z.namelist():
                if nombre.lower().endswith('.docx'):
                    z.extract(nombre, '/tmp')
                    archivos.append({'nombre': nombre, 'ruta': f'/tmp/{nombre}'})
    except Exception as e: st.error(f"Error ZIP: {str(e)}")
    return archivos

st.markdown('<div class="header-container"><div class="header-title">⚡ Extractor RDP Pro v5.2</div><div class="header-subtitle">🏭 Sistema Industrial de Extracción de Protocolos TC/TV</div></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div class="sidebar-content"><h3>📋 Instrucciones</h3><p>1️⃣ Comprime tu carpeta en .ZIP<br>2️⃣ Sube el ZIP<br>3️⃣ Procesa automáticamente<br>4️⃣ Descarga Excel</p><hr><h4>📊 Campos</h4><ul><li>🏢 Subestación</li><li>⚡ Bahía</li><li>🔌 Relé</li><li>🏷️ Marca</li><li>🔢 Serie</li></ul></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.success("🟢 Sistema Activo")
    st.info(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

col1, col2 = st.columns([3, 1])
with col1:
    st.markdown('<div class="upload-zone"><div class="big-icon">📦</div><h3 style="color:#1e3c72">Sube tu archivo ZIP</h3><p>El sistema recorrerá todas las subcarpetas</p></div>', unsafe_allow_html=True)
    zip_file = st.file_uploader("", type="zip", label_visibility="collapsed")
with col2:
    st.info("**Formato:** .ZIP\n\n**Procesa:**\n- Todo automático")
    if zip_file: st.success("✅ ZIP listo")
    else: st.warning("⏳ Esperando")

st.markdown("<br>", unsafe_allow_html=True)
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    procesar_btn = st.button("🚀 PROCESAR ARCHIVOS", use_container_width=True)

if procesar_btn:
    if not zip_file: st.warning("⚠️ Sube un archivo ZIP")
    else:
        registros = []
        progreso = st.progress(0)
        archivos = buscar_archivos_en_zip(zip_file)
        if not archivos: st.error("❌ No hay .docx en el ZIP")
        else:
            total = len(archivos)
            cont_sub, cont_bay, cont_mar = set(), set(), set()
            for idx, arc in enumerate(archivos):
                try:
                    datos = extraer_datos_archivo(arc['ruta'], arc['nombre'])
                    if datos["ser"] or datos["rel"]:
                        registros.append({"Subestacion": datos["sub"], "Bahia": datos["bay"], "celda": datos["bay"], "Rele": datos["rel"], "marca": datos["mar"], "fecha": datos["fec"], "codigo_orden": datos["ord"], "serie": datos["ser"], "relacion Tc": datos["tc"], "relacion Tv": datos["tv"], "Valor_prueba": datos["val_i"], "VP_VOLTAGE": datos["val_v"], "Rechazo_Carga": datos["rechazo"], "Archivo_Origen": datos["archivo"]})
                        if datos["sub"]: cont_sub.add(datos["sub"])
                        if datos["bay"]: cont_bay.add(datos["bay"])
                        if datos["mar"]: cont_mar.add(datos["mar"])
                except: pass
                progreso.progress((idx + 1) / total)
            if registros:
                st.markdown("---")
                st.markdown(f'<div class="success-banner">✅ Procesados: {len(registros)} reportes</div>', unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: st.markdown(f'<div class="metric-card"><div class="metric-number">{len(registros)}</div><div class="metric-label">Archivos</div></div>', unsafe_allow_html=True)
                with c2: st.markdown(f'<div class="metric-card"><div class="metric-number">{len(cont_sub)}</div><div class="metric-label">Subestaciones</div></div>', unsafe_allow_html=True)
                with c3: st.markdown(f'<div class="metric-card"><div class="metric-number">{len(cont_bay)}</div><div class="metric-label">Bahías</div></div>', unsafe_allow_html=True)
                with c4: st.markdown(f'<div class="metric-card"><div class="metric-number">{len(cont_mar)}</div><div class="metric-label">Marcas</div></div>', unsafe_allow_html=True)
                st.markdown("### 📋 Vista Previa")
                df = pd.DataFrame(registros)
                st.dataframe(df, use_container_width=True, hide_index=True)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Datos_RDP')
                buffer.seek(0)
                st.markdown("---")
                col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
                with col_d2:
                    st.download_button(label="📥 DESCARGAR EXCEL MAESTRO", data=buffer, file_name="BASE_DATOS_FINAL_V52.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else: st.error("❌ Sin datos")

st.markdown('<div class="footer"><p><b>🔧 Extractor RDP Pro v5.2</b> | Sistema Industrial</p><p style="font-size:12px;opacity:0.8">© 2024</p></div>', unsafe_allow_html=True)

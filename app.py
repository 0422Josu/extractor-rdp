import streamlit as st
import pandas as pd
from docx import Document
import re
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Extractor RDP Pro v5.2", 
    page_icon="🔍", 
    layout="centered",
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS PERSONALIZADO ---
st.markdown("""
<style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton>button {
        width: 100%;
        height: 60px;
        font-size: 18px;
        font-weight: bold;
        border-radius: 10px;
    }
    .success-box {
        padding: 20px;
        border-radius: 10px;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .info-box {
        padding: 15px;
        border-radius: 10px;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
</style>
""", unsafe_allow_html=True)

# --- TÍTULO PRINCIPAL ---
st.title("🔍 Extractor RDP Pro v5.2")
st.markdown("### Sistema de Extracción de Protocolos TC/TV")
st.markdown("---")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📋 Instrucciones")
    st.info("""
    1. Selecciona archivos .docx
    2. El sistema procesará los reportes
    3. Descarga el Excel final
    
    **Campos extraídos:**
    - Subestación, Bahía, Relé
    - Marca, Serie, Código orden
    - Relaciones TC y TV
    - Valores de prueba
    """)
    st.markdown("---")
    st.markdown("**Desarrollado con Streamlit**")

# --- FUNCIONES DE EXTRACCIÓN ---

def limpiar_valor(texto):
    """Limpia caracteres especiales de un texto"""
    if not texto: return ""
    return texto.replace('\n', ' ').strip().replace(',', '')

def formatear_numero(valor_str):
    """Convierte valores a formato numérico entero"""
    valor_str = valor_str.upper()
    numero_match = re.search(r'([0-9.]+)', valor_str)
    if not numero_match: return "0"
    
    numero = float(numero_match.group(1))
    if "KV" in valor_str:
        numero = int(numero * 1000)
    else:
        numero = int(numero)
    return str(numero)

def extraer_datos_archivo(doc_file, nombre_archivo):
    """Extrae todos los datos de un archivo .docx"""
    info = {
        "sub": "", "bay": "", "rel": "", "mar": "", "fec": "", 
        "ord": "", "ser": "", "val_i": "", "val_v": "", "rechazo": "",
        "tc": "", "tv": "", "archivo": nombre_archivo
    }
    
    v_prim, v_sec, i_prim, i_sec = "", "", "", ""

    try:
        doc = Document(doc_file)
        texto_completo = ""
        
        for tabla in doc.tables:
            for fila in tabla.rows:
                celdas = [c.text.strip() for c in fila.cells]
                texto_fila = " ".join(celdas)
                texto_completo += texto_fila + " "

                # Búsqueda de datos en celdas
                for i, texto_celda in enumerate(celdas):
                    if "Substation:" == texto_celda and i + 1 < len(celdas):
                        info["sub"] = limpiar_valor(celdas[i+1])
                    elif "Bay:" == texto_celda and i + 1 < len(celdas):
                        info["bay"] = limpiar_valor(celdas[i+1])
                    elif "Name/description:" == texto_celda and i + 1 < len(celdas):
                        info["rel"] = limpiar_valor(celdas[i+1])
                    elif "Manufacturer:" == texto_celda and i + 1 < len(celdas):
                        info["mar"] = limpiar_valor(celdas[i+1])
                    elif "Serial/model number:" == texto_celda and i + 1 < len(celdas):
                        info["ser"] = limpiar_valor(celdas[i+1])
                    elif "Additional info 1:" == texto_celda and i + 1 < len(celdas):
                        info["ord"] = limpiar_valor(celdas[i+1])

                    # Extracción de valores nominales (TC/TV)
                    if "V nom (secondary):" in texto_celda and i + 1 < len(celdas):
                        v_sec = formatear_numero(celdas[i+1])
                    if "V primary:" in texto_celda and i + 1 < len(celdas):
                        v_prim = formatear_numero(celdas[i+1])
                    if "I nom (secondary):" in texto_celda and i + 1 < len(celdas):
                        i_sec = formatear_numero(celdas[i+1])
                    if "I primary:" in texto_celda and i + 1 < len(celdas):
                        i_prim = formatear_numero(celdas[i+1])

        # Construir relaciones TC/TV
        if v_prim and v_sec: info["tv"] = f"{v_prim}/{v_sec}"
        if i_prim and i_sec: info["tc"] = f"{i_prim}/{i_sec}"

        # Rechazo de Carga
        if "RECHAZO DE CARGA" in texto_completo.upper(): 
            info["rechazo"] = "SI"

        # Fecha del reporte
        f_m = re.search(r'(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})', texto_completo)
        info["fec"] = f_m.group(1).replace('.', '/') if f_m else ""

        # Voltaje y Corriente de inyección
        i_m = re.findall(r'I\s*L\d\s*[|]?\s*([1-9]\d*\.?\d*)\s*A', texto_completo, re.I)
        if i_m: info["val_i"] = i_m[0]
        
        v_m = re.findall(r'V\s*L\d-?E?\s*[|]?\s*([1-9]\d*\.?\d*)\s*V', texto_completo, re.I)
        if v_m: info["val_v"] = v_m[0]

    except Exception as e:
        st.error(f"Error procesando {nombre_archivo}: {str(e)}")
    
    return info

# --- INTERFAZ PRINCIPAL ---

st.subheader("📁 Carga de Archivos")

# Widget para subir archivos múltiples
archivos_subidos = st.file_uploader(
    "Selecciona los archivos .docx a procesar",
    type="docx",
    accept_multiple_files=True,
    help="Puedes seleccionar múltiples archivos a la vez"
)

# Botón para procesar
if st.button("🚀 PROCESAR ARCHIVOS", type="primary"):
    if not archivos_subidos:
        st.warning("⚠️ Por favor, selecciona al menos un archivo .docx")
    else:
        registros = []
        progreso = st.progress(0)
        
        total_archivos = len(archivos_subidos)
        
        for idx, archivo in enumerate(archivos_subidos):
            datos = extraer_datos_archivo(archivo, archivo.name)
            
            if datos["ser"] or datos["rel"]:
                registros.append({
                    "Subestacion": datos["sub"],
                    "Bahia": datos["bay"],
                    "celda": datos["bay"],
                    "Rele": datos["rel"],
                    "marca": datos["mar"],
                    "fecha": datos["fec"],
                    "codigo_orden": datos["ord"],
                    "serie": datos["ser"],
                    "relacion Tc": datos["tc"],
                    "relacion Tv": datos["tv"],
                    "Valor_prueba": datos["val_i"],
                    "VP_VOLTAGE": datos["val_v"],
                    "Rechazo_Carga": datos["rechazo"],
                    "Archivo_Origen": datos["archivo"]
                })
            
            # Actualizar progreso
            progreso.progress((idx + 1) / total_archivos)

        # Mostrar resultados
        if registros:
            st.markdown("---")
            st.subheader("✅ Resultados")
            
            df = pd.DataFrame(registros)
            
            # Mostrar vista previa de datos
            st.dataframe(df, use_container_width=True)
            
            # Contador de archivos procesados
            st.markdown(f"""
            <div class="success-box">
                📊 <b>Procesados exitosamente:</b> {len(registros)} reportes
            </div>
            """, unsafe_allow_html=True)
            
            # Botón de descarga Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Datos_RDP')
            
            buffer.seek(0)
            
            st.download_button(
                label="📥 DESCARGAR EXCEL MAESTRO",
                data=buffer,
                file_name="BASE_DATOS_FINAL_V52.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        else:
            st.error("❌ No se encontró información válida en los archivos.")

# --- FOOTER ---
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #888;'>"
    "🔧 Extractor RDP Pro v5.2 - Web Version | Generado con Streamlit"
    "</div>", 
    unsafe_allow_html=True
)
import streamlit as st
import os
import json
import datetime
import mercadopago
from dotenv import load_dotenv
from reviewer import evaluate_project

# Load environment variables
load_dotenv()

st.set_page_config(page_title="Revisor de Proyectos de Investigación", layout="wide")

st.title("Revisor de Proyectos de Investigación")
st.markdown("##### 🧬 Plataforma Inteligente de Evaluación Científica | **PerspectaSalud**")

# Configuración de Llaves Maestras
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
mp_access_token = st.secrets.get("MP_ACCESS_TOKEN", os.getenv("MP_ACCESS_TOKEN", ""))
APP_URL = st.secrets.get("APP_URL", "https://revisordetesis.streamlit.app/")
PRECIO_INFORME = 3.80  # Precio neuromarketing equivalente a 1 USD

# Estado de Pago
is_paid = st.session_state.get("is_paid", False)

# Si el usuario regresa de Mercado Pago, la URL tendrá un "?status=approved"
query_params = st.query_params
if "status" in query_params and query_params["status"] == "approved":
    is_paid = True
    st.session_state["is_paid"] = True
    st.query_params.clear()

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3003/3003035.png", width=100) # Un ícono temporal genérico
    st.markdown("### Acerca de PerspectaSalud")
    st.write("Somos expertos en validación metodológica y normativa de proyectos de investigación en salud. Elevamos el estándar de la ciencia académica.")
    st.markdown("---")
    
    # Bypass de Administrador (Para ti)
    admin_code = st.text_input("Acceso Institucional", type="password", placeholder="Código de acceso")
    if admin_code == "PERSPECTA2026":
        is_paid = True
        st.session_state["is_paid"] = True
        st.success("Modo Institucional Activado")

# --- MURO DE PAGO (Marketing Value Proposition) ---
if not is_paid:
    st.info(
        f"**¿Buscas la excelencia en investigación científica?**\n\n"
        f"Ya seas un **estudiante** buscando aprobar sin observaciones del jurado, o un **asesor** optimizando su tiempo de revisión, "
        f"obtén una auditoría científica de nivel doctoral. "
        f"Nuestra plataforma identifica vacíos metodológicos, bibliográficos y lógicos en cuestión de minutos, garantizando un proyecto sólido y riguroso.\n\n"
        f"**Inversión:** S/ {PRECIO_INFORME:.2f}"
    )
    
    if not mp_access_token:
        st.error("🛠️ El sistema de pagos está en mantenimiento. Vuelve más tarde.")
    else:
        sdk = mercadopago.SDK(mp_access_token)
        preference_data = {
            "items": [
                {
                    "title": "Auditoría Científica - PerspectaSalud",
                    "quantity": 1,
                    "unit_price": PRECIO_INFORME,
                    "currency_id": "PEN"
                }
            ],
            "back_urls": {
                "success": APP_URL,
                "failure": APP_URL,
                "pending": APP_URL
            },
            "auto_return": "approved"
        }
        try:
            preference_response = sdk.preference().create(preference_data)
            init_point = preference_response["response"]["init_point"]
            st.link_button("🚀 Desbloquear mi Auditoría Científica (Yape/Tarjeta)", init_point, type="primary", use_container_width=True)
        except Exception as e:
            st.error("Ocurrió un error al generar el enlace de pago.")
    st.stop()

# --- APLICACIÓN DESBLOQUEADA ---
st.success("✅ Acceso concedido. Puedes proceder con la evaluación de los documentos.")

# Crear dos columnas (Col 1: Contenido Principal (75%), Col 2: Banner Publicitario (25%))
col_main, col_ad = st.columns([3, 1])

with col_ad:
    st.markdown("### Patrocinado")
    st.image("https://images.unsplash.com/photo-1576091160399-112ba8d25d1d?w=500&q=80", caption="Asesoría Estadística Avanzada para Tesis Médicas. ¡Contáctanos hoy!")
    st.markdown("[Solicitar Asesoría](https://www.google.com.pe)")
    # Nota para Alexander: Aquí se puede inyectar el código HTML real de Google AdSense más adelante.

with col_main:
    st.markdown("### 1. Datos del Proyecto")
    project_name = st.text_input("Nombre del autor o proyecto (para la descarga):", placeholder="Ej. Coronado-Elias")

    st.markdown("### 2. Documentos del Proyecto")
    uploaded_files = st.file_uploader("Sube aquí los archivos del proyecto (PDF, DOCX, TXT, XLSX, CSV):", accept_multiple_files=True)

    normative_folder = "./normativas"
    if not os.path.exists(normative_folder):
        os.makedirs(normative_folder)

    if st.button("Iniciar Revisión", type="primary"):
        if not api_key:
            st.error("Error interno: Falta la API Key de Gemini.")
        elif not project_name:
            st.error("Por favor, ingresa el nombre del proyecto.")
        elif not uploaded_files:
            st.error("Por favor, sube al menos un documento para revisar.")
        else:
            with st.spinner("Procesando auditoría científica. Esto puede tomar un par de minutos..."):
                result = evaluate_project(uploaded_files, normative_folder, api_key)
            
            if isinstance(result, str):
                st.error(result)
            else:
                report = result.get("report", "Error recuperando el reporte.")
                
                st.success("¡Revisión completada!")
                st.markdown("### Informe de Revisión del Proyecto")
                # Al ser texto plano sin markdown, podemos usar texto preformateado o st.text
                st.text(report)
                
                # Dynamic filename
                date_str = datetime.datetime.now().strftime("%d%m%Y")
                download_file_name = f"{project_name}_{date_str}.txt"
                
                st.download_button(
                    label="Descargar Informe en TXT",
                    data=report,
                    file_name=download_file_name,
                    mime="text/plain"
                )

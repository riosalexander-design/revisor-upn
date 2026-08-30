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
st.markdown("#### Elaborado por PerspectaSalud")
st.write("Esta herramienta evalúa los documentos de investigación usando agentes especializados en metodología, bibliografía, estilo Vancouver y normativa institucional.")

# Configuración de Llaves Maestras
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
mp_access_token = st.secrets.get("MP_ACCESS_TOKEN", os.getenv("MP_ACCESS_TOKEN", ""))
APP_URL = st.secrets.get("APP_URL", "https://revisor-upn-nckbmqfpnmyfp8atlqscth.streamlit.app/")
PRECIO_INFORME = 10.00  # Puedes cambiar este precio luego

# Estado de Pago
is_paid = st.session_state.get("is_paid", False)

# Si el usuario regresa de Mercado Pago, la URL tendrá un "?status=approved"
query_params = st.query_params
if "status" in query_params and query_params["status"] == "approved":
    is_paid = True
    st.session_state["is_paid"] = True
    # Limpiar los parámetros para evitar confusiones
    st.query_params.clear()

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3003/3003035.png", width=100) # Un ícono temporal genérico
    st.markdown("### Acerca de PerspectaSalud")
    st.write("Somos expertos en validación metodológica y normativa de proyectos de investigación en salud.")
    st.markdown("---")
    
    # Bypass de Administrador (Para ti)
    admin_code = st.text_input("Admin PIN", type="password", placeholder="PIN Secreto")
    if admin_code == "PERSPECTA2026":
        is_paid = True
        st.session_state["is_paid"] = True
        st.success("Modo Admin Activado (Acceso Gratuito)")

# --- MURO DE PAGO ---
if not is_paid:
    st.warning(f"🔒 El analizador automatizado por IA se encuentra bloqueado. El costo por generación de informe profundo es de **S/ {PRECIO_INFORME:.2f}**.")
    
    if not mp_access_token:
        st.info("🛠️ El sistema de pagos está en mantenimiento. Vuelve más tarde.")
    else:
        # Generar Link de Pago Dinámico
        sdk = mercadopago.SDK(mp_access_token)
        preference_data = {
            "items": [
                {
                    "title": "Informe de Revisión - PerspectaSalud",
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
            # Botón que redirige a Mercado Pago
            st.link_button("🔓 Pagar con Yape / Tarjeta para Desbloquear", init_point, type="primary")
        except Exception as e:
            st.error("Ocurrió un error al generar el enlace de pago.")
            
    # Detenemos la aplicación aquí para que no puedan ver ni usar el formulario de subida
    st.stop()

# --- APLICACIÓN DESBLOQUEADA ---
st.success("✅ Acceso concedido. Puedes proceder con la evaluación de los documentos.")

st.markdown("### 1. Datos del Proyecto")
project_name = st.text_input("Nombre del autor o proyecto (para la descarga):", placeholder="Ej. Coronado-Elias")

st.markdown("### 2. Documentos del Proyecto")
uploaded_files = st.file_uploader("Sube aquí los archivos del proyecto (PDF, DOCX, TXT, XLSX, CSV):", accept_multiple_files=True)

# Por defecto, la carpeta de normativas estará en el mismo directorio del código
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
        with st.spinner("Los agentes están revisando la documentación. Esto puede tomar un par de minutos..."):
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

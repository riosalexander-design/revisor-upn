import streamlit as st
import os
import json
import datetime
import mercadopago
import requests
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
PRECIO_INFORME = 4.90  # Precio introductorio

# Inyectar Google Analytics (Invisible)
ga_code = """
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-J2NW5JD21D"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'G-J2NW5JD21D');
</script>
"""
import streamlit.components.v1 as components
components.html(ga_code, height=0, width=0)

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

import streamlit.components.v1 as components

# Crear dos columnas principales globales (Col 1: Contenido, Col 2: Banner Publicitario)
col_main, col_ad = st.columns([3, 1])

with col_ad:
    st.markdown("### 📊 Tráfico en Vivo")
    st.caption("Visitas por país/región")
    # Usamos un widget de FlagCounter gratuito. 
    # Alexander puede reemplazar esta URL creando su propio contador en flagcounter.com
    html_code = """
    <a href="https://info.flagcounter.com/83u"><img src="https://s11.flagcounter.com/count2/83u/bg_FFFFFF/txt_000000/border_CCCCCC/columns_2/maxflags_6/viewers_3/labels_1/pageviews_1/flags_0/percent_0/" alt="Flag Counter" border="0"></a>
    """
    components.html(html_code, height=180)
    
    st.divider()

    st.markdown("### Patrocinado")
    st.image("https://images.unsplash.com/photo-1576091160399-112ba8d25d1d?w=500&q=80", caption="Asesoría Estadística Avanzada para Tesis Médicas. ¡Contáctanos hoy!")
    st.markdown("[Solicitar Asesoría](https://www.google.com.pe)")
    # Nota para Alexander: Aquí se puede inyectar el código HTML real de Google AdSense más adelante.
    
    st.divider()
    
    st.markdown("### ✉️ Contáctanos")
    st.caption("Para soporte, reclamos o publicidad.")
    with st.form("contact_form", clear_on_submit=True):
        user_email = st.text_input("Tu correo electrónico:")
        user_message = st.text_area("Mensaje:")
        submitted = st.form_submit_button("Enviar mensaje")
        if submitted:
            contact_email = st.secrets.get("CONTACT_EMAIL", "")
            if not contact_email:
                st.error("Error: El correo de destino no está configurado.")
            elif user_email and user_message:
                try:
                    response = requests.post(f"https://formsubmit.co/ajax/{contact_email}", json={
                        "Nombre": "Usuario de PerspectaSalud",
                        "Email": user_email,
                        "Mensaje": user_message
                    })
                    if response.status_code == 200:
                        st.success("¡Mensaje recibido! Revisa tu bandeja para confirmar si es la primera vez.")
                    else:
                        st.error("Ocurrió un problema al enviar el mensaje.")
                except Exception as e:
                    st.error("Error de conexión al enviar el correo.")
            else:
                st.error("Por favor llena todos los campos.")

with col_main:
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

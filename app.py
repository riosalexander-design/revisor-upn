import streamlit as st
import os
import json
import datetime
from dotenv import load_dotenv
from reviewer import evaluate_project

# Load environment variables
load_dotenv()

st.set_page_config(page_title="Revisor de Proyectos UPN", layout="wide")

DB_FILE = "db.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_db(db_list):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db_list, f, indent=4)

db = load_db()

st.title("Revisor de Proyectos de Investigación - UPN")
st.markdown("#### Ideación y elaboración: Mg. Alexander N. Rios Rios")
st.write("Esta herramienta evalúa los documentos de investigación usando agentes especializados en metodología, bibliografía, estilo Vancouver y normativa de la Universidad Privada del Norte (UPN).")

# Se leerá de Secrets si está en la nube (st.secrets), o de .env si está local
api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

with st.sidebar:
    st.header("Configuración")
    input_api_key = st.text_input("Gemini API Key", value=api_key, type="password")
    if input_api_key:
        api_key = input_api_key
        
    st.markdown("---")
    st.markdown("### Ranking de Proyectos")
    
    # Sort DB by score descending
    db_sorted = sorted(db, key=lambda x: x.get("score", 0), reverse=True)
    
    for idx, item in enumerate(db_sorted):
        folder_name = item.get("folder_name", "Desconocido")
        score = item.get("score", 0)
        has_ethics = item.get("has_ethics", False)
        has_grave_obs = item.get("has_grave_obs_before_methodology", True)
        
        if not has_grave_obs and not has_ethics:
            color = "red"
        elif has_ethics:
            color = "blue"
        else:
            color = "black"
            
        st.markdown(f"<span style='color:{color}; font-weight:bold;'>{score}pts - {folder_name}</span>", unsafe_allow_html=True)


st.markdown("### 1. Datos del Estudiante / Proyecto")
project_name = st.text_input("Nombre del autor o proyecto (para el ranking y descarga):", placeholder="Ej. Coronado-Elias")

st.markdown("### 2. Documentos del Proyecto")
uploaded_files = st.file_uploader("Sube aquí los archivos del proyecto (PDF, DOCX, TXT, XLSX, CSV):", accept_multiple_files=True)

# Por defecto, la carpeta de normativas ahora estará en el mismo directorio del código
normative_folder = "./normativas"
if not os.path.exists(normative_folder):
    os.makedirs(normative_folder)

if st.button("Iniciar Revisión", type="primary"):
    if not api_key:
        st.error("Por favor, ingresa tu Gemini API Key en la barra lateral.")
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
                metadata = result.get("metadata", {})
                score = metadata.get("score", 0)
                
                # Update DB
                # Remove previous evaluation for this project if exists
                db = [item for item in db if item.get("folder_name") != project_name]
                
                # Append new evaluation
                db.append({
                    "folder_name": project_name,
                    "score": score,
                    "has_ethics": metadata.get("has_ethics", False),
                    "has_grave_obs_before_methodology": metadata.get("has_grave_obs_before_methodology", True),
                    "date": datetime.datetime.now().strftime("%d%m%Y")
                })
                save_db(db)
                
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

import os
import io
import fitz  # PyMuPDF
import docx
import re
import json
import time
from google import genai
from google.genai import types

import pandas as pd

def extract_text_from_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    text = ""
    try:
        if name.endswith('.pdf'):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                text += page.get_text()
        elif name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif name.endswith('.txt'):
            text = uploaded_file.read().decode('utf-8', errors='ignore')
        elif name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
            text = df.to_string()
        elif name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            text = df.to_string()
    except Exception as e:
        print(f"Error reading {name}: {e}")
    return text

def extract_text_from_local_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == '.pdf':
            doc = fitz.open(file_path)
            for page in doc:
                text += page.get_text()
        elif ext == '.docx':
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return text

def evaluate_project(uploaded_files, normative_folder, api_key, academic_level="Pregrado", work_stage="Proyecto de investigación"):
    client = genai.Client(api_key=api_key)
    
    # Extraer texto de los archivos subidos en Streamlit
    document_text = ""
    for uf in uploaded_files:
        document_text += f"\n\n--- INICIO DE DOCUMENTO: {uf.name} ---\n\n"
        document_text += extract_text_from_uploaded_file(uf)
        document_text += f"\n\n--- FIN DE DOCUMENTO: {uf.name} ---\n\n"
        
    if not document_text.strip():
        return "No se encontraron documentos válidos o están vacíos."

    # Extraer texto de la carpeta de normativas local del servidor
    institutional_normative = ""
    if os.path.isdir(normative_folder):
        for root, dirs, files in os.walk(normative_folder):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx', '.txt')):
                    path = os.path.join(root, file)
                    institutional_normative += f"\n\n--- NORMATIVA: {file} ---\n\n"
                    institutional_normative += extract_text_from_local_file(path)
    else:
        return f"No se encontró la carpeta de normativas en el servidor: {normative_folder}"
    
    sys_instruction = (
        "Eres un panel de 5 expertos (Asesor Científico de Nivel Doctoral, Metodólogo, Verificador Bibliográfico, Corrector de Estilo Vancouver y Revisor Normativo) y un Editor Académico Principal.\n"
        "Tu tarea es generar un INFORME DE REVISIÓN DEL PROYECTO DE INVESTIGACIÓN unificado, riguroso y altamente formativo, estructurado de la siguiente manera:\n\n"
        "1. REPORTE CIENTÍFICO-DOCTORAL (REALIDAD PROBLEMÁTICA, COHERENCIA LÓGICA Y ARTICULACIÓN TEÓRICA):\n"
        "- Realidad problemática en embudo: Verifica rigurosamente que la problemática se plantee desde fuentes de nivel MUNDIAL, transitando luego al nivel REGIONAL (Latinoamérica/continental), luego al nivel NACIONAL y finalmente al nivel LOCAL (si corresponde).\n"
        "- Coherencia interna integral: Evalúa la correspondencia lógica estricta entre lo que se propone, el problema general, los problemas específicos, el objetivo general, los objetivos específicos, las dimensiones del estudio y la teoría que respalda la investigación.\n"
        "- Articulación teórica y dimensiones: Las dimensiones de las variables deben insinuarse, justificarse y visibilizarse desde la realidad problemática y los antecedentes, vinculándose explícitamente con la teoría científica que sostiene la propuesta.\n"
        "- GUÍA CLARA PARA LEVANTAR OBSERVACIONES DE COHERENCIA Y TEORÍA: Cuando identifiques debilidades o vacíos en la coherencia interna o en la articulación teórica, sé sumamente claro, pedagógico y resolutivo. NO te limites a señalar la inconsistencia o falta de alineación: debes proporcionar a la estudiante alternativas y rutas directas, claras y fáciles de comprender para levantar y subsanar la observación con éxito. Presenta la alineación estructural sugerida (por ejemplo: la correspondencia exacta entre pregunta de investigación, objetivo y dimensión) y redacta el párrafo modelo de fundamentación teórica que conecte la teoría de referencia con sus variables de estudio, listo para que la estudiante lo revise, adapte e inserte directamente en su documento.\n\n"
        "2. REPORTE METODOLÓGICO Y ANÁLISIS ESTADÍSTICO:\n"
        "- Identifica debilidades metodológicas, de diseño, población, muestra, operacionalización e instrumentos.\n"
        "- Comprensibilidad según nivel académico: Explica los conceptos metodológicos o estadísticos complejos de forma pedagógica y comprensible según el nivel académico del usuario (Pregrado, Maestría o Doctorado), manteniendo siempre el lenguaje académico adecuado.\n"
        "- Naturaleza de los datos y rutas estadísticas: Revisa primero la naturaleza de las variables y los datos (categóricos/nominales/ordinales, numéricos continuos/discretos, supuestos de normalidad/homocedasticidad). A partir de allí, sugiere OBLIGATORIAMENTE LAS DOS MEJORES RUTAS de análisis estadístico para el estudio (Ruta A y Ruta B, por ejemplo: opción paramétrica vs no paramétrica, o análisis bivariado correlacional vs modelo multivariado/regresión).\n"
        "- Software estadístico: Si el estudio es de tipo cuantitativo, tus recomendaciones de análisis de datos deben enfocarse EXCLUSIVAMENTE en el uso del software SPSS, detallando los procedimientos a seguir.\n\n"
        "3. REPORTE BIBLIOGRÁFICO Y ANTECEDENTES:\n"
        "- Relevancia científica y calidad de fuentes (priorizando bases indexadas como PubMed, SciELO, Scopus, etc.).\n"
        "- Cantidad exigida de antecedentes según nivel académico: No exijas un número cerrado, sino un umbral MÍNIMO: para nivel Pregrado, exige como MÍNIMO 5 antecedentes internacionales y 5 antecedentes nacionales (al menos 10 en total); para nivel Posgrado (Maestría o Doctorado), exige con mayor profundidad académica como MÍNIMO 10 antecedentes internacionales y 10 antecedentes nacionales (al menos 20 en total).\n"
        "- Orden cronológico: Los antecedentes deben estar estrictamente organizados desde el más antiguo hasta el más reciente (orden cronológico ascendente).\n"
        "- Articulación: Los antecedentes deben guardar estrecha relación con las preguntas de investigación, los objetivos, las dimensiones de las variables y el marco teórico. Evalúa esto ponderando el nivel de investigación y el nivel académico del estudiante.\n\n"
        "4. REPORTE ESTILO VANCOUVER:\n"
        "- Revisa la alineación estricta a las normas de citación y referenciación Vancouver (numeración consecutiva según orden de aparición en el texto y formato de referencias).\n\n"
        "5. REPORTE NORMATIVA INSTITUCIONAL Y ASPECTOS ÉTICOS:\n"
        "- Cumplimiento formal institucional y estructura del documento.\n"
        "- Aspectos éticos sin solicitar comité: NUNCA solicites ni sugieras someter el trabajo a un Comité de Ética. En su lugar, indica directamente los cambios de redacción que debe implementar en el apartado ético (principios de autonomía, confidencialidad, no maleficencia, beneficencia, Declaración de Helsinki, consentimiento informado, etc.) y entrégale el texto modelo listo para insertar en su trabajo para que los aspectos éticos queden óptimos.\n\n"
        "FORMATO OBLIGATORIO PARA CADA OPORTUNIDAD DE MEJORA EN LAS 5 SECCIONES:\n"
        "Para CADA debilidad u oportunidad de mejora encontrada, mantén ESTRICTAMENTE este formato de 3 partes:\n"
        "OBSERVACIÓN: [Descripción puntual, exacta y precisa de lo que requiere ajuste]\n"
        "EXPLICACIÓN: [Explicación clara, comprensible y fundamentada, manteniendo el lenguaje académico pero adaptada al nivel académico declarado por el usuario para desglosar conceptos complejos con ejemplos]\n"
        "SUGERENCIA DE CORRECCIÓN: [Instrucción en tono ASERTIVO pero IMPERATIVO dirigida al estudiante, explicando con total claridad los pasos para levantar la observación. Incluye OBLIGATORIAMENTE un EJEMPLO FÁCIL DE COMPRENDER ajustado específicamente a su tema y variables de estudio, redactando directamente el párrafo o texto exacto sugerido para que la estudiante pueda copiarlo e insertarlo directamente en su trabajo].\n\n"
        "REGLAS ESTRICTAS DE NEGOCIO:\n"
        "- PROHIBIDO COLOCAR CALIFICACIÓN O PUNTAJE: NO asignes ninguna nota numérica, ni puntaje sobre 20 o 100, ni restes puntos en ninguna parte del informe, ya que eso desanima al estudiante. Cierra el informe con una conclusión motivadora y asertiva enfocada en cómo implementar las mejoras.\n"
        "- REDACCIÓN EN TEXTO PLANO Y SEGUNDA PERSONA: Redacta absolutamente todo el informe (fuera del bloque JSON) en TEXTO PLANO y EN SEGUNDA PERSONA ('tu investigación', 'debes precisar', etc.). NO uses formato Markdown en el cuerpo del informe (sin asteriscos de negrita, sin numerales, sin guiones para listas) para que la exportación a Word sea limpia.\n"
        "- CONFIDENCIALIDAD INSTITUCIONAL: NUNCA menciones las siglas UPN ni 'Universidad Privada del Norte'. Refiérete únicamente a 'la institución' o 'tu universidad'.\n\n"
        "Tu tarea final es MUY IMPORTANTE: Al puro final de tu respuesta, debes agregar un bloque JSON estricto con la siguiente estructura (este bloque es oculto para el sistema y nunca lo ve el estudiante, no agregues nada después):\n"
        "```json\n"
        "{\n"
        '  "project_title": "[Extrae el título exacto de la tesis/proyecto del documento. Si no está, pon \\"Título no especificado\\"]",\n'
        '  "authors": "[Extrae el nombre del autor o autores del documento. Si no está, pon \\"Autor no especificado\\"]"\n'
        "}\n"
        "```"
    )

    prompt = (
        f"DATOS DE LA SOLICITUD:\n"
        f"- Nivel académico declarado: {academic_level}\n"
        f"- Etapa del trabajo declarada: {work_stage}\n\n"
        f"Analiza los siguientes documentos del proyecto de investigación ajustando tu evaluación a este nivel académico:\n\n{document_text}\n\n"
        f"Considera la siguiente normativa/rúbrica institucional de referencia:\n{institutional_normative}"
    )

    models_to_try = [
        'gemini-3.6-flash',
        'gemini-3.5-flash',
        'gemini-3.7-flash',
        'gemini-flash-latest'
    ]
    
    last_error = ""

    for model_name in models_to_try:
        for attempt in range(3): # Reintenta 3 veces por cada modelo
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_instruction,
                        temperature=0.2,
                    )
                )
                
                raw_report = response.text
                
                # Corrección de Mojibake (Si la API devuelve texto mal decodificado en ISO-8859-1 en lugar de UTF-8)
                try:
                    raw_report = raw_report.encode('latin-1').decode('utf-8')
                except UnicodeError:
                    pass
                
                
                metadata = {
                    "score": 0,
                    "has_ethics": False,
                    "has_grave_obs_before_methodology": True
                }
                report_clean = raw_report
                
                json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_report, re.DOTALL)
                if json_match:
                    try:
                        metadata = json.loads(json_match.group(1))
                        report_clean = raw_report[:json_match.start()].strip()
                    except:
                        pass

                return {
                    "report": report_clean,
                    "metadata": metadata
                }
                
            except Exception as e:
                last_error = str(e)
                # Si el modelo no existe o está bloqueado para esta API Key (404), no perdemos tiempo reintentando, pasamos al siguiente modelo.
                if "404" in last_error or "NOT_FOUND" in last_error or "is no longer available" in last_error:
                    break 
                
                # Si es un error de saturación temporal (503), esperamos un poco y reintentamos con el mismo modelo
                time.sleep(5 * (attempt + 1))
                
    return {"report": f"Error Crítico: Servidores de IA saturados. Se intentó conectar con todos los modelos disponibles sin éxito. Último error: {last_error}", "metadata": {}}

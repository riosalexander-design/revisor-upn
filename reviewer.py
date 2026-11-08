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
        "Eres un panel de alta especialidad académica (Asesor Científico Doctoral, Metodólogo y Bioestadístico, Verificador Bibliográfico, Corrector de Estilo Vancouver y Revisor Normativo) actuando como un Asesor y Editor Académico Principal.\n"
        "Tu misión es elaborar un INFORME DE REVISIÓN Y AUDITORÍA CIENTÍFICA unificado, fluido, riguroso y formativo, dirigido directamente a la estudiante en tono de asesoría experta. Integra la evaluación en un único reporte articulado y armónico estructurado en las siguientes 5 áreas fundamentales:\n\n"
        "1. DIAGNÓSTICO DE COHERENCIA CIENTÍFICA Y ARTICULACIÓN TEÓRICA:\n"
        "- Realidad problemática en embudo: Verifica rigurosamente que la problemática se plantee desde fuentes y estadísticas de nivel MUNDIAL, transitando al nivel REGIONAL (Latinoamérica/continental), luego al nivel NACIONAL y finalmente al nivel LOCAL (contexto específico del estudio).\n"
        "- Coherencia lógica integral: Evalúa la correspondencia estricta entre la propuesta general, el problema general y específicos, el objetivo general y específicos, las dimensiones del estudio y el marco teórico que respalda la investigación.\n"
        "- Articulación teórica y dimensiones: Las dimensiones de las variables deben insinuarse, justificarse y visibilizarse desde la realidad problemática y los antecedentes, vinculándose explícitamente con la teoría científica que sostiene la propuesta.\n"
        "- GUÍA RESOLUTIVA PARA LEVANTAR OBSERVACIONES: Cuando identifiques debilidades o vacíos en la coherencia interna o en la articulación teórica, sé sumamente claro, pedagógico y resolutivo. Proporciona a la estudiante la correspondencia exacta sugerida (pregunta - objetivo - dimensión) y redacta el párrafo modelo de fundamentación teórica que conecte la teoría con sus variables, listo para copiar, adaptar e insertar directamente en su trabajo. Si este párrafo incluye datos empíricos, autores o marcos teóricos, incorpora sus llamadas de cita y entrega inmediatamente las referencias bibliográficas completas en formato Vancouver con su DOI para que las sume a su bibliografía.\n\n"
        "2. ANTECEDENTES Y REVISIÓN DE LA LITERATURA:\n"
        "- Relevancia científica y calidad de fuentes (priorizando bases indexadas como PubMed, SciELO, Scopus, Web of Science, etc.).\n"
        "- Cantidad exigida de antecedentes según nivel académico declarado: No exijas un número cerrado, sino un umbral MÍNIMO: para nivel Pregrado, exige como MÍNIMO 5 antecedentes internacionales y 5 antecedentes nacionales (al menos 10 en total); para nivel Posgrado (Maestría o Doctorado), exige con mayor profundidad científica como MÍNIMO 10 antecedentes internacionales y 10 antecedentes nacionales (al menos 20 en total).\n"
        "- Orden cronológico riguroso: Los antecedentes deben estar estrictamente organizados desde el más antiguo hasta el más reciente (orden cronológico ascendente).\n"
        "- Articulación con variables y marco teórico: Los antecedentes deben fundamentar directamente los objetivos, dimensiones y teoría del estudio.\n"
        "- Sugerencias de literatura y DOIs: Si recomiendas o citas estudios específicos para enriquecer antecedentes o sustentar la investigación, proporciona su DOI verificado o su referencia bibliográfica completa. Si no puedes certificar la existencia plena del DOI, proporciona la referencia bibliográfica detallada e indícale explícitamente a la estudiante que compruebe y corrobore la fuente en las bases de datos científicas antes de incluirla.\n\n"
        "3. METODOLOGÍA, NATURALEZA DE VARIABLES Y RUTA ESTADÍSTICA EN SPSS:\n"
        "- Claridad sobre la naturaleza de las variables: Explica con total claridad y pedagogía la naturaleza y escala de medición de cada una de las variables del estudio (ej. cualitativa nominal dicotómica/politómica, cualitativa ordinal en escala Likert, cuantitativa continua o discreta a nivel de razón/intervalo).\n"
        "- Articulación con el objetivo y las hipótesis del estudio: Determina las rutas estadísticas acordes con la intención del estudio (¿es una correlación o relación entre variables?, ¿es una comparación de grupos independientes o relacionados?, ¿es un modelo explicativo, de asociación o regresión?).\n"
        "- Planteamiento directo de las 2 mejores rutas estadísticas:\n"
        "  * RUTA A (Opción paramétrica): Procedimiento cuando las variables cuantitativas satisfacen los supuestos de distribución normal y homogeneidad de varianzas.\n"
        "  * RUTA B (Opción no paramétrica): Procedimiento robusto cuando las variables son de naturaleza ordinal, la muestra no cumple normalidad o los datos son libres de distribución.\n"
        "- Operatividad detallada paso a paso en SPSS:\n"
        "  * Verificación previa de normalidad en SPSS: Ruta exacta en el menú (ej. Analizar > Estadísticos descriptivos > Explorar > Gráficos > Gráficos con pruebas de normalidad) especificando qué estadístico observar (Kolmogorov-Smirnov con corrección Lilliefors si n > 50, o Shapiro-Wilk si n < 50) y el criterio de decisión según el p-valor (Sig. > 0.05 normalidad, Sig. < 0.05 no normalidad).\n"
        "  * Ejecución paso a paso de la prueba elegida en SPSS: Ruta exacta de menús, traslado de variables de la lista izquierda a los campos de prueba, casillas específicas a activar y regla de interpretación del resultado (p-valor / nivel de significancia bilateral vs alfa 0.05) adaptada al nivel académico declarado.\n\n"
        "4. ASPECTOS BIOÉTICOS Y RIGOR FORMAL (VANCOUVER Y NORMATIVA):\n"
        "- Aspectos bioéticos sin solicitar comité: NUNCA solicites ni sugieras someter el trabajo a un Comité de Ética institucional. En su lugar, indica de forma directa las mejoras que debe implementar en la redacción bioética (principios de autonomía, confidencialidad y anonimato, beneficencia, no maleficencia, justicia, Declaración de Helsinki, consentimiento informado formal) y entrégale el texto modelo completo listo para insertar en su trabajo para que los aspectos éticos queden óptimos y blindados.\n"
        "- Estilo de citación Vancouver: Revisa la correspondencia estricta de citación numérica consecutiva según orden de aparición en el texto y formato de referencias bibliográficas.\n"
        "- Normativa y estructura institucional: Alineación formal a las secciones y pautas metodológicas de la institución.\n\n"
        "5. SÍNTESIS Y PLAN DE ACCIÓN INMEDIATO:\n"
        "- Cierre del informe con una conclusión motivadora, asertiva y formativa que organice las prioridades de corrección del estudiante de forma secuencial.\n"
        "- PROHIBIDO COLOCAR CALIFICACIÓN O PUNTAJE: NO asignes ninguna nota numérica, ni puntaje sobre 20 o 100, ni restes puntos en ninguna parte del informe, ya que eso desanima a los estudiantes.\n\n"
        "FORMATO OBLIGATORIO PARA CADA OPORTUNIDAD DE MEJORA IDENTIFICADA:\n"
        "Para CADA debilidad, vacío u oportunidad de mejora encontrada a lo largo del informe, mantén rigurosamente esta estructura de 3 componentes:\n"
        "OBSERVACIÓN: [Descripción puntual, exacta y precisa de lo que requiere ajuste en el documento]\n"
        "EXPLICACIÓN: [Explicación clara, comprensible y fundamentada, manteniendo el rigor académico pero adaptada con ejemplos al nivel académico declarado por el usuario para desglosar conceptos complejos]\n"
        "SUGERENCIA DE CORRECCIÓN: [Instrucción en tono ASERTIVO pero IMPERATIVO dirigida a la estudiante, indicándole de forma contundente y transparente qué debe hacer para levantar la observación. Incluye OBLIGATORIAMENTE un EJEMPLO FÁCIL DE COMPRENDER ajustado con precisión a su tema y a sus variables de estudio, redactando directamente el párrafo, tabla o texto exacto sugerido para que la estudiante pueda copiarlo e insertarlo directamente en su trabajo.\n"
        "OBLIGATORIO - FUENTES PARA SU BIBLIOGRAFÍA: Si el texto sugerido menciona estudios, estadísticas, organismos (OMS, OPS, INEI, etc.) o autores teóricos, coloca las llamadas de cita en el párrafo (ej. (1), (2)) y agrega inmediatamente debajo el acápite 'FUENTES PARA TU BIBLIOGRAFÍA' con la referencia bibliográfica formal y completa en estilo Vancouver y su DOI verificado (o enlace formal) para que la estudiante no tenga citas huérfanas y las inserte directamente en su lista de referencias final].\n\n"
        "REGLAS ESTRICTAS DE NEGOCIO Y ESTILO:\n"
        "- REDACCIÓN EN TEXTO PLANO Y SEGUNDA PERSONA: Redacta absolutamente todo el informe (fuera del bloque JSON final) en TEXTO PLANO y EN SEGUNDA PERSONA ('tu investigación', 'debes precisar', 'en tu estudio', etc.). NO uses formato Markdown en el cuerpo del informe (sin asteriscos de negrita, sin numerales para títulos, sin guiones para viñetas) para garantizar que la descarga a Word (.docx) sea limpia, profesional y libre de caracteres de marcado.\n"
        "- SUGERENCIAS IMPERATIVAS Y MODELOS INSERTABLES: Cada sugerencia debe ser directiva ('Reemplaza...', 'Inserta el siguiente párrafo...', 'Ejecuta en SPSS la siguiente ruta...'), eliminando ambigüedades y facilitando el texto listo para aplicar.\n"
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

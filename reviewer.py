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
        "- Revisión del plan de análisis o propuesta del usuario: Examina primero de forma rigurosa el plan de análisis estadístico o técnica que el estudiante planteó (o omitió) en su metodología.\n"
        "- Evaluación de pertinencia con variables y objetivos: Evalúa críticamente si la propuesta del usuario es la más adecuada para la naturaleza y escala de medición de sus variables y si responde con coherencia estricta a sus objetivos e hipótesis de investigación, señalando con claridad sus aciertos o las inconsistencias metodológicas que deben corregirse.\n"
        "- Claridad sobre la naturaleza de las variables: Explica con total pedagogía la naturaleza y escala de medición de cada una de las variables del estudio (ej. cualitativa nominal dicotómica/politómica, cualitativa ordinal en escala Likert, cuantitativa continua o discreta a nivel de razón/intervalo).\n"
        "- Planteamiento directo de las 2 mejores rutas estadísticas acordes al objetivo:\n"
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
        "FORMATO OBLIGATORIO PARA LA RETROALIMENTACIÓN FORMATIVA:\n"
        "- SOBRIEDAD Y MODERACIÓN EN EL RECONOCIMIENTO DE ACIERTOS: En cada una de las 5 áreas, señala de forma objetiva, serena y profesional los aspectos que el estudiante ha desarrollado conforme a los estándares académicos (ej. 'Aspectos conformes', 'Aciertos identificados', 'Has delimitado adecuadamente tu objetivo general'). Sé moderado: QUEDA PROHIBIDO el uso de adjetivos hiperbólicos o halagos excesivos (evita palabras como 'excelente', 'maravilloso', 'brillante', 'fascinante' o elogios desmedidos). Mantén la sobriedad y distancia profesional de un comité evaluador doctoral.\n"
        "- ENFOQUE ASERTIVO Y PROFESIONAL PARA AJUSTES: Cuando un aspecto requiera corrección o ajuste, preséntalo con asertividad académica sobria, utilizando fórmulas como 'Oportunidad de mejora', 'Para mayor claridad y solidez ante el jurado', 'Ajuste metodológico necesario'. Evita tanto el lenguaje punitivo ('error grave', 'deficiente') como la sobrecarga de palabras de autoayuda o aliento excesivo. El tono debe ser riguroso, claro y constructivo.\n"
        "- ESTRUCTURA PARA CADA OPORTUNIDAD DE MEJORA:\n"
        "  OPORTUNIDAD DE MEJORA: [Descripción puntual y objetiva de lo que se debe ajustar para mayor solidez]\n"
        "  EXPLICACIÓN FORMATIVA: [Explicación clara, sobria y fundamentada según el nivel académico declarado]\n"
        "  ACCIÓN RECOMENDADA: [Instrucción en tono ASERTIVO pero IMPERATIVO ('Reemplaza...', 'Inserta...'), explicando con total claridad qué hacer. Incluye OBLIGATORIAMENTE un EJEMPLO CONCRETO ajustado a sus variables, redactando directamente el párrafo o procedimiento exacto para que la estudiante pueda insertarlo directamente en su trabajo.\n"
        "  OBLIGATORIO - FUENTES Y PREVENCIÓN DE ENLACES FALSOS: Si el texto sugerido menciona estudios o teorías, coloca las llamadas de cita (ej. (1), (2)) y agrega inmediatamente debajo el acápite 'FUENTES SUGERIDAS PARA TU BIBLIOGRAFÍA' con la referencia en estilo Vancouver. PROHIBIDO INVENTAR ENLACES DOI FICTICIOS: Si no tienes certeza fáctica absoluta del DOI exacto, NO generes una URL de DOI inventada. En su lugar, entrega la referencia bibliográfica formal y añade la 'Ecuación de búsqueda sugerida' entre comillas para que el estudiante localice el artículo en PubMed, SciELO o Google Académico (ej. Búsqueda en SciELO/PubMed: \"Gómez-Urquiza\" AND \"burnout\" AND \"urgencias\"). Recuerda siempre al estudiante corroborar personalmente el documento y los datos en las bases científicas antes de la entrega final].\n\n"
        "REGLAS ESTRICTAS DE NEGOCIO Y ESTILO:\n"
        "- REDACCIÓN EN TEXTO PLANO Y SEGUNDA PERSONA: Redacta absolutamente todo el informe (fuera del bloque JSON final) en TEXTO PLANO y EN SEGUNDA PERSONA ('tu investigación', 'debes precisar', 'en tu estudio', etc.). NO uses formato Markdown en el cuerpo del informe (sin asteriscos de negrita, sin numerales para títulos, sin guiones para viñetas) para garantizar que la descarga a Word (.docx) sea limpia, profesional y libre de caracteres de marcado.\n"
        "- TONO SOBRIO, DIRECTIVO Y SIN HALAGOS EXCESIVOS: Emplea instrucciones directas e imperativas ('Reemplaza...', 'Inserta el siguiente párrafo...', 'Ejecuta en SPSS la siguiente ruta...'), con un estilo sobrio, profesional y moderado, sin adulaciones ni entusiasmo artificial.\n"
        "- TOLERANCIA CERO A ENLACES ROTOS O DOIs ALUCINADOS: No inventes URLs largas ni códigos DOI aleatorios. Brinda citas formales reales o verificables, y proporciona términos de búsqueda directos para bases de datos indexadas para garantizar que el estudiante encuentre fuentes reales.\n"
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

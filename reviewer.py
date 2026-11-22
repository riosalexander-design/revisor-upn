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

from docx.text.paragraph import Paragraph
from docx.table import Table

def _extract_table_lines(table):
    t_lines = ["[INICIO DE TABLA]"]
    for row in table.rows:
        row_cells = []
        seen_tc = set()
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id not in seen_tc:
                seen_tc.add(tc_id)
                parts = []
                for item in cell.iter_inner_content():
                    if isinstance(item, Paragraph):
                        if item.text.strip():
                            parts.append(item.text.strip())
                    elif isinstance(item, Table):
                        parts.extend(_extract_table_lines(item))
                row_cells.append(" ".join(parts))
        if any(row_cells):
            t_lines.append(" | ".join(row_cells))
    t_lines.append("[FIN DE TABLA]")
    return t_lines

def _extract_docx_content(doc):
    lines = []
    for item in doc.iter_inner_content():
        if isinstance(item, Paragraph):
            if item.text.strip():
                lines.append(item.text)
        elif isinstance(item, Table):
            lines.extend(_extract_table_lines(item))
    return "\n".join(lines)

def _extract_pdf_content(doc):
    pages_out = []
    for idx, page in enumerate(doc, start=1):
        p_txt = f"\n--- [PÁGINA {idx}] ---\n"
        p_txt += page.get_text("text", sort=True)
        pages_out.append(p_txt)
    return "\n".join(pages_out)

_NORMATIVAS_CACHE = {}

def extract_text_from_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    text = ""
    try:
        if name.endswith('.pdf'):
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            text = _extract_pdf_content(doc)
        elif name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = _extract_docx_content(doc)
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
            text = _extract_pdf_content(doc)
        elif ext == '.docx':
            doc = docx.Document(file_path)
            text = _extract_docx_content(doc)
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

    # Extraer texto de la carpeta de normativas local del servidor (con caché en memoria)
    if normative_folder in _NORMATIVAS_CACHE:
        institutional_normative = _NORMATIVAS_CACHE[normative_folder]
    elif os.path.isdir(normative_folder):
        institutional_normative = ""
        for root, dirs, files in os.walk(normative_folder):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx', '.txt')):
                    path = os.path.join(root, file)
                    institutional_normative += f"\n\n--- NORMATIVA: {file} ---\n\n"
                    institutional_normative += extract_text_from_local_file(path)
        _NORMATIVAS_CACHE[normative_folder] = institutional_normative
    else:
        return f"No se encontró la carpeta de normativas en el servidor: {normative_folder}"
    
    sys_instruction = (
        "Eres un panel de alta especialidad académica (Asesor Científico Doctoral, Metodólogo y Bioestadístico, Auditor Clínico-Epidemiológico, Corrector de Estilo Vancouver y Revisor Normativo) actuando como Editor Académico Principal de Perspecta Salud.\n"
        "Tu misión es auditar minuciosamente el documento completo (incluyendo cada párrafo, cada tabla de resultados y cada tabla o instrumento de los anexos) y elaborar una HOJA DE RUTA DE CORRECCIONES IMPERATIVAS secuencial (página por página, desde la carátula hasta el último anexo), empleando el formato quirúrgico de contraste directo 'DONDE DICE:' / 'COLOCA:'.\n\n"
        "ESTRUCTURA OBLIGATORIA DEL REPORTE:\n\n"
        "ENCABEZADO E INSTRUCCIÓN GENERAL (LO LOGRADO Y LO QUE NO DEBES MODIFICAR):\n"
        "- Inicia con el título: HOJA DE RUTA DE CORRECCIONES IMPERATIVAS — PERSPECTA SALUD\n"
        "- Debajo, redacta el bloque 'Instrucción general y aspectos logrados:':\n"
        "  * Reconoce con sobriedad académica (sin adjetivos exagerados ni halagos como 'excelente' o 'maravilloso') los aciertos metodológicos y estructurales del trabajo.\n"
        "  * VALIDACIÓN PRIORITARIA DEL ANÁLISIS ESTADÍSTICO: Verifica primero si la persona ya realizó o propuso un análisis estadístico en su documento. SI EL ANÁLISIS ESTADÍSTICO YA REALIZADO O PROPUESTO ESTÁ BIEN (es acertado para la naturaleza de sus variables y responde a sus objetivos, ej. Rho de Spearman, Chi-cuadrado, etc.), felicítala con sobriedad en esta Instrucción General, explícale en palabras sencillas por qué esa prueba es la correcta para la escala de sus variables y sus objetivos, e indícale explícitamente: 'NO modifiques la prueba estadística [Nombre de la prueba] (es metodológicamente correcta)'. Si su elección estadística ya es correcta, NO agregues pasos de SPSS para cambiar de prueba (solo corrige en las tablas detalles de reporte como cambiar p = 0,000 por p < 0,001 o usar coma decimal).\n"
        "  * Indica también qué otros elementos bien alineados NO debe modificar por inseguridad (por ejemplo, mantener los objetivos específicos en su formulación global si en SPSS los resultados fueron procesados por puntajes globales y no desglosados por dimensiones).\n\n"
        "SECCIONES SECUENCIALES DE LA HOJA DE RUTA (Recorre el documento en su orden físico real de principio a fin, numerando correlativamente CORRECCIÓN 01, CORRECCIÓN 02, CORRECCIÓN 03...):\n\n"
        "1. PÁGINAS PRELIMINARES Y RESUMEN\n"
        "- Audita Carátula y Título (delimitación precisa de variables, población, establecimiento/lugar y año).\n"
        "- Audita Dedicatoria, Agradecimiento e Informe de Similitud (detecta páginas en blanco y corrige errores ortográficos reales como 'fé' con tilde -> 'fe' sin tilde, 'qué ilumine' -> 'que ilumine', redacción cortada).\n"
        "- Audita Resumen y Palabras Claves: precisión metodológica ('tipo básico' en lugar de 'nivel básico' cuando el nivel/alcance es correlacional/descriptivo), uso obligatorio de coma decimal en todos los porcentajes (ej. 68,3 % en vez de 68.3 %), reemplazo obligatorio de 'p = 0.000' o 'p = 0,000' por 'p < 0,001' (explicando entre paréntesis que en bioestadística nunca se reporta p = 0,000), y al menos 5 descriptores científicos (DeCS/MeSH).\n\n"
        "2. CAPÍTULO I: INTRODUCCIÓN\n"
        "- 1.1 Realidad Problemática:\n"
        "  * Estructura en embudo estricto: Verifica que transite desde el nivel MUNDIAL (OMS), al REGIONAL (OPS/Latinoamérica), luego al NACIONAL (MINSA/INEI) y aterrice obligatoriamente en el nivel LOCAL (caracterización demográfica y epidemiológica del distrito y del establecimiento o institución específica donde se ejecuta el estudio). Si falta el cierre local, redacta en 'COLOCA:' el párrafo situacional local completo listo para insertar.\n"
        "  * Rigor epidemiológico: Detecta errores conceptuales o de indicadores (por ejemplo, confundir casos absolutos de muertes con tasas por cien mil nacidos vivos, o expresiones mal redactadas como 'maternidad materna' en vez de 'mortalidad materna').\n"
        "- 1.2 Antecedentes:\n"
        "  * Umbral mínimo según nivel académico declarado: Para Pregrado, exige como MÍNIMO 5 antecedentes internacionales y 5 nacionales (al menos 10 en total); para Posgrado (Maestría/Doctorado), exige como MÍNIMO 10 internacionales y 10 nacionales (al menos 20 en total).\n"
        "  * Orden cronológico ascendente estricto: Deben ordenarse desde el más antiguo hasta el más reciente (ej. 2020 -> 2023 -> 2024 -> 2025). En 'COLOCA:', muestra la secuencia cronológica exacta ordenada de los autores que ya tiene y sugiere los estudios reales faltantes (buscados en tiempo real con google_search) con su referencia Vancouver y DOI real o enlace directo a la revista científica.\n"
        "  * Corrige erratas textuales o redundancias dentro de los antecedentes (ej. 'conocimiento alto sobre conocimiento').\n"
        "- 1.3 Bases Teóricas y Marco Conceptual:\n"
        "  * Precisión clínica y científica: Detecta errores terminológicos graves (ej. escribir 'hiperemia gravídica' en lugar de 'hiperémesis gravídica', mezclar signos clínicos distintos en una misma cita, o escribir nombres bacteriológicos como 'E-coli' en lugar de 'Escherichia coli' en cursiva).\n"
        "  * Articulación con Teoría Científica: Verifica que las variables y sus dimensiones estén respaldadas por una teoría o modelo científico clásico (ej. Modelo de Creencias en Salud de Rosenstock / Janz y Becker, Teoría de Jean Watson, Modelo de Dorothea Orem, etc.). Si falta, redacta en 'COLOCA:' el párrafo completo de articulación teórica listo para copiar e insertar al final de las Bases Teóricas.\n"
        "- 1.4 / 1.5 Justificación, Problemas, Objetivos e Hipótesis:\n"
        "  * Verifica numeración de subtítulos, signos de interrogación '?' en todos los problemas, mayúsculas en nombres propios de instituciones, y coherencia exacta entre problemas, objetivos, hipótesis y lo procesado en resultados.\n\n"
        "3. CAPÍTULO II: METODOLOGÍA\n"
        "- Tipo, alcance y diseño; Población, muestra y muestreo:\n"
        "  * Precisa 'tipo de investigación básica (o aplicada)' y 'alcance/nivel descriptivo/correlacional/explicativo', así como los meses exactos del periodo de estudio.\n"
        "  * Si el muestreo es no probabilístico por conveniencia o censal y su justificación es débil, entrega en 'COLOCA:' el párrafo técnico blindado (criterio de captación continua / saturación de flujo y cautela en la extrapolación).\n"
        "  * Verifica que no haya saltos en la numeración correlativa de citas Vancouver (ej. pasar de la cita 32 a la 34 omitiendo la 33).\n"
        "- Plan de Análisis y Tratamiento Estadístico de los Datos:\n"
        "  * SI EL ANÁLISIS ESTADÍSTICO DEL ESTUDIANTE YA ES ACERTADO: Recuerda que ya lo validaste en la Instrucción General; no agregues una corrección para cambiar de prueba ni des pasos en SPSS.\n"
        "  * SOLO SI EL ANÁLISIS ESTADÍSTICO ES INCORRECTO, INCOMPLETO O FUE OMITIDO: Crea un bloque 'CORRECCIÓN XX — Plan de Análisis Estadístico y Procesamiento en SPSS [CRÍTICO]'. En 'DONDE DICE:' cita lo que propuso u omitió. En 'COLOCA:' incluye obligatoriamente:\n"
        "    1) Explicación sencilla, con tacto pedagógico y sin jerga compleja, de por qué su propuesta anterior no se ajusta a la escala de sus variables y por qué hemos escogido UN ÚNICO tratamiento estadístico definitivo (PROHIBIDO usar la palabra 'ruta' o 'rutas' o dar alternativas condicionales), con un ejemplo práctico aplicado a sus variables.\n"
        "    2) El párrafo modelo del Plan de Análisis listo para insertar en su metodología.\n"
        "    3) El paso a paso numerado en SPSS: (1) Menú exacto en SPSS, (2) En qué casilla ubicar cada variable, (3) Qué botones y casillas marcar, y (4) Cómo interpretar el valor 'Sig.' o p-valor (< 0,05) y reportar p < 0,001 cuando SPSS muestre 0,000.\n"
        "- Aspectos Éticos (Marco bioético internacional y legal peruano vigente, SIN pedir trámite de Comité de Ética):\n"
        "  * NUNCA solicites tramitar aprobación ante un Comité de Ética ni adjuntar cartas de comité. En su lugar, en 'COLOCA:' entrega directamente el párrafo normativo blindado listo para reemplazar/insertar que integre: Declaración de Helsinki, Pautas Éticas CIOMS 2016, Informe Belmont (autonomía, beneficencia, no maleficencia y justicia), Consideraciones Éticas para la Investigación en Salud con Seres Humanos del MINSA (Resolución Ministerial N.° 233-2020-MINSA) y la Ley N.° 29733 (Ley de Protección de Datos Personales) junto con su reglamento vigente aprobado por Decreto Supremo N.° 016-2024-JUS (codificación alfanumérica de instrumentos).\n\n"
        "4. CAPÍTULO III: RESULTADOS (Aplica cuando el documento incluya capítulo o tablas de resultados)\n"
        "- Audita fila por fila cada tabla extraída del documento ([INICIO DE TABLA] / [TABLA X]) y contrástala con el cuestionario del Anexo:\n"
        "  * Detecta cálculos estadísticamente inválidos (por ejemplo, reportar una fila 'Promedio: 25 años' en la Tabla 1 cuando el instrumento del Anexo solo recogió intervalos de edad como '18-25, 26-35, 36 a más' y no la edad exacta).\n"
        "  * Exige reemplazar en todas las tablas e interpretaciones 'p = 0,000' o '0.000' por '< 0,001' y escribir '(ρ = 0,xxx; p < 0,001)'.\n"
        "  * Exige uniformizar todos los números y porcentajes con un decimal y coma decimal (ej. 50,0 %; 33,3 %; 100,0 % en lugar de mezclar '50', '33.3', '100.0').\n\n"
        "5. CAPÍTULO IV (DISCUSIÓN Y CONCLUSIONES), DECLARACIÓN DE IA Y REFERENCIAS BIBLIOGRÁFICAS\n"
        "- Discusión (si está presente): Detecta erratas ortográficas/tipográficas, exige formato '(ρ = 0,xxx; p < 0,001)' y verifica que no existan contradicciones lógicas entre lo que afirma la discusión y lo que realmente hallaron sus propios antecedentes citados (redactando en 'COLOCA:' el párrafo exacto que resuelve la contradicción).\n"
        "- Conclusiones (si están presentes): Verifica que cada conclusión responda en orden estricto a los objetivos (General, OE1, OE2, OE3...) e incorpore los valores estadísticos y porcentuales exactos obtenidos en Resultados.\n"
        "- Declaración de Uso de IA (si está presente): Indica precisar herramienta, modelo específico y proveedor (ej. 'ChatGPT, modelo GPT-4o, proveedor OpenAI').\n"
        "- Referencias Bibliográficas y Estilo Vancouver:\n"
        "  * Verifica que la lista final tenga numeración visible correlativa (1., 2., 3. ...) según orden de aparición en el texto.\n"
        "  * Corrige 'Disponible de:' por 'Disponible en:', fechas futuras erróneas y posición de las llamadas numéricas en todo el manuscrito (el número entre paréntesis debe ir ANTES del punto final y sin espacios dobles, ej. '...durante el embarazo (1).').\n"
        "  * Lista todas las fuentes nuevas incorporadas en las correcciones anteriores con su referencia completa en estilo Vancouver y su DOI real verificado o enlace directo a la revista oficial (obtenido mediante google_search, NUNCA enlaces de búsqueda de Google Scholar), incluyendo SIEMPRE la nota obligatoria: '(Verifica, lee y corrobora personalmente cada fuente y enlace en las bases científicas antes de incorporarla en tu texto y bibliografía)'.\n\n"
        "6. ANEXOS (CONSISTENCIA, OPERACIONALIZACIÓN, INSTRUMENTO Y CONSENTIMIENTO)\n"
        "- Audita exhaustivamente cada tabla y texto de los anexos:\n"
        "  * Anexo 1 (Matriz de Consistencia): Verifica que incluya todos los problemas específicos y que las dimensiones listadas coincidan con lo realmente evaluado en los objetivos y resultados.\n"
        "  * Anexo 2 (Matriz de Operacionalización de Variables): Contrasta los baremos o puntajes del Anexo 2 contra la metodología y el cuestionario del Anexo 3 (por ejemplo, detectar si en el Anexo 2 colocaron 'Nivel alto 18-20 / Nivel medio 14-17 / Nivel bajo 10-13' cuando el cuestionario real es de 10 ítems con baremo de 0 a 10 puntos). Verifica también que estén incluidas las variables sociodemográficas si corresponden al OE1.\n"
        "  * Anexo 3 (Instrumento de Recolección de Datos): Detecta plurales incorrectos en títulos ('SIGNOS DE ALARMAS' -> 'SIGNOS DE ALARMA'), ítems con opciones de respuesta duplicadas (indicando reemplazar la opción repetida por el distractor original del autor del instrumento) y encabezados desordenados o sucios en la tabla Likert ('Totalmente en desacuerdo (1)', 'En desacuerdo (2)', 'Ni de acuerdo ni en desacuerdo (3)', 'De acuerdo (4)', 'Totalmente de acuerdo (5)').\n"
        "  * Anexo de Consentimiento Informado: Adapta la concordancia gramatical de género a la población real estudiada (ej. si son gestantes, usar 'fui informada', 'para usted' en vez de 'informado (a)' o 'para él (ella)') y corrige la relación de principios bioéticos ('Confidencialidad y privacidad (principio de respeto a la autonomía)' en lugar de asociar confidencialidad con justicia).\n\n"
        "FORMATO ESTRICTO E INQUEBRANTABLE PARA CADA CORRECCIÓN:\n"
        "Cada hallazgo debe redactarse EXACTAMENTE con esta estructura de 3 líneas/bloques para que el sistema lo convierta automáticamente en tablas ejecutivas dentro del archivo Word:\n"
        "CORRECCIÓN [Número de 2 dígitos: 01, 02, 03...] — [Subtítulo claro de la sección y tema] (Página [X o rango X–Y]) [Agregar '[CRÍTICO]' al final del título si la observación compromete la aprobación metodológica, clínica, estadística o normativa]\n"
        "DONDE DICE: [Cita textual exacta entre comillas de lo que dice el documento actual, o descripción precisa de la fila/celda de la tabla con sus valores actuales. Si hay varios errores en esa misma sección, numéralos como 1., 2., 3.]\n"
        "COLOCA: [Instrucción directa, asertiva e imperativa con el texto exacto corregido listo para copiar y pegar, numerando 1., 2., 3. en correspondencia exacta con DONDE DICE, e incluyendo entre paréntesis la breve explicación pedagógica de por qué se aplica esa mejora]\n\n"
        "REGLAS ESTRICTAS DE NEGOCIO Y ESTILO:\n"
        "- PROHIBIDO COLOCAR CALIFICACIÓN O PUNTAJE: No asignes notas sobre 20 ni sobre 100.\n"
        "- PROHIBIDO SOLICITAR TRÁMITE DE COMITÉ DE ÉTICA: Nunca pidas al estudiante presentar su trabajo a un Comité de Ética ni adjuntar cartas de aprobación del CIEI; entrégale directamente el párrafo ético normativo listo para insertar.\n"
        "- CONFIDENCIALIDAD INSTITUCIONAL: NUNCA menciones las siglas UPN ni 'Universidad Privada del Norte' en ninguna parte del reporte (usa 'tu universidad' o 'la institución'), ni inventes nombres propios de asesores que no figuren en el documento subido.\n"
        "- TEXTO LIMPIO SIN MARKDOWN: No uses asteriscos (**), ni numerales (#), ni tablas Markdown (|---|) en tu respuesta; usa únicamente las etiquetas en mayúsculas 'CORRECCIÓN 01 — ...', 'DONDE DICE:' y 'COLOCA:' para que el generador nativo de Word construya las tablas.\n\n"
        "Tu tarea final es MUY IMPORTANTE: Al puro final de tu respuesta, debes agregar un bloque JSON estricto con la siguiente estructura (este bloque es oculto para el sistema y nunca lo ve el estudiante, no agregues nada después):\n"
        "```json\n"
        "{\n"
        '  "project_title": "[Extrae el título exacto de la tesis/proyecto del documento. Si no está, pon \\"Título no especificado\\"]",\n'
        '  "authors": "[Extrae el nombre del autor o autores del documento. Si no está, pon \\"Autor no especificado\\"]"\n'
        "}\n"
        "```"
    )

    prompt = (
        f"NORMATIVA Y RÚBRICA INSTITUCIONAL DE REFERENCIA:\n{institutional_normative}\n\n"
        f"=================================================================\n"
        f"DATOS DE LA SOLICITUD A AUDITAR:\n"
        f"- Nivel académico declarado: {academic_level}\n"
        f"- Etapa del trabajo declarada: {work_stage}\n\n"
        f"DOCUMENTO COMPLETO DEL PROYECTO / TESIS DEL ESTUDIANTE (INCLUYE TEXTO Y TABLAS):\n\n{document_text}\n\n"
        f"=================================================================\n"
        f"INSTRUCCIÓN FINAL OBLIGATORIA DE FORMATO:\n"
        f"Genera AHORA la 'HOJA DE RUTA DE CORRECCIONES IMPERATIVAS — PERSPECTA SALUD' recorriendo el documento página por página y capítulo por capítulo. "
        f"Inicia con 'Instrucción general y aspectos logrados:' (indicando qué está bien y qué NO debe modificar, incluyendo la validación de la prueba estadística si ya es correcta). "
        f"Luego estructura ABSOLUTAMENTE CADA CORRECCIÓN con estas 3 etiquetas exactas en líneas separadas:\n"
        f"CORRECCIÓN 01 — [Título y Sección] (Página X) [CRÍTICO]\n"
        f"DONDE DICE: [Cita textual exacta o valores actuales de la tabla con error/omisión]\n"
        f"COLOCA: [Texto exacto corregido listo para copiar e insertar + breve explicación entre paréntesis]\n"
        f"(PROHIBIDO usar el formato antiguo de 'OBSERVACIÓN / EXPLICACIÓN / ACCIÓN RECOMENDADA'; usa EXCLUSIVAMENTE 'CORRECCIÓN XX', 'DONDE DICE:' y 'COLOCA:')."
    )

    models_to_try = [
        'gemini-pro-latest',
        'gemini-3.1-pro-preview',
        'gemini-3.8-flash',
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
                        tools=[{"google_search": {}}],
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

# Contexto y Reglas del Proyecto: Perspecta Salud (Revisor de Tesis)

Este archivo (`AGENTS.md`) contiene la memoria histórica, decisiones de diseño, arquitectura y tareas pendientes del proyecto **Perspecta Salud** para que cualquier agente de Antigravity tenga el contexto completo desde el primer momento.

---

## 1. Descripción General y Arquitectura
- **Producto:** Plataforma web llamada **Perspecta Salud** (`https://revisordetesis.streamlit.app/`) que audita avances y proyectos de tesis universitarias utilizando IA (Google Gemini) basada en rúbricas académicas y entrega un informe descargable en formato Word (`.docx`).
- **Stack Principal:**
  - `app.py`: Interfaz principal en **Streamlit** (con diseño personalizado vía CSS inyectado, navegación con smooth scroll en JS, pasarela de Mercado Pago con 3 paquetes, sistema de códigos/tokens, CRM local y panel de administrador).
  - `reviewer.py`: Motor de evaluación con **Google Gemini API** y generación del documento Word con `python-docx`.
  - `normativas/`: Carpeta con documentos y rúbricas institucionales de referencia.

---

## 2. Reglas de Negocio Inquebrantables (Directivas del Usuario)
1. **Principio de UI Sin Fricción:** Mantener el formulario del estudiante lo más simple posible. **NO** dividir el campo de nombre del proyecto en múltiples campos (como "Título" y "Autor" por separado) para no darle más trabajo al cliente.
2. **Extracción Automática de Metadatos:** El título real y los autores del proyecto deben extraerse automáticamente del documento subido para incluirlos en el reporte.
3. **Nombre Exacto del Archivo Word Descargable:** El archivo `.docx` generado DEBE llamarse exactamente:
   `PerspectaSalud_(Nombre del autor o proyecto)_(fecha en formato ddmmaaaa).docx`
4. **Reglas del Prompt de Evaluación (`reviewer.py`):**
   - **Prohibido calificar o asignar puntajes:** No mostrar calificaciones numéricas ni restar puntos para no desanimar al cliente. Cierre motivador y formativo enfocado en el plan de acción.
   - **Comprensibilidad según nivel académico:** Desglosar conceptos complejos con explicaciones claras y ejemplos adaptados al nivel académico declarado (Pregrado, Maestría, Doctorado), manteniendo el rigor académico.
   - **Realidad problemática en embudo y coherencia:** Exigir estructura desde fuentes Mundiales -> Regionales -> Nacionales -> Locales. Coherencia estricta entre problema general/específicos, objetivos, dimensiones y teoría. Las dimensiones deben insinuarse y sustentarse desde la problemática y antecedentes vinculadas a la teoría. En fallas de coherencia o articulación teórica, proporcionar alternativas directas, la correspondencia terna sugerida (pregunta-objetivo-dimensión) y párrafos modelo de enlace teórico listos para insertar para levantar la observación sin ambigüedades.
   - **Antecedentes:** Exigir como umbral mínimo según nivel: al menos 5 internacionales y 5 nacionales para Pregrado (mínimo 10 en total); y al menos 10 internacionales y 10 nacionales para Posgrado (mínimo 20 en total). Siempre ordenados de forma cronológica ascendente (del más antiguo al más reciente) y articulados con preguntas, objetivos, dimensiones y teoría.
   - **Estadística (Naturaleza de variables, objetivos y rutas en SPSS):**
     1. Revisar primero el plan de análisis estadístico propuesto por el usuario en su metodología (o señalar su omisión).
     2. Evaluar críticamente si la propuesta del usuario es la más idónea para la naturaleza de sus variables y si responde con coherencia estricta a sus preguntas y objetivos de investigación.
     3. Explicar con total claridad la naturaleza y escala de medición de las variables del estudio (nominales, ordinales, intervalo/razón, continuas o discretas).
     4. Plantear obligatoriamente las dos mejores rutas: RUTA A (paramétrica con supuestos de normalidad) y RUTA B (no paramétrica para ordinales o libre distribución).
     5. Detallar la operatividad paso a paso en SPSS: comprobación previa de normalidad (Kolmogorov-Smirnov con Lilliefors para n > 50 o Shapiro-Wilk para n < 50), ruta exacta en los menús de SPSS (`Analizar > ...`), casillas a marcar y regla de decisión del p-valor según el nivel académico.
   - **Sobriedad y moderación en el reconocimiento de aciertos:** Señalar de forma objetiva y serena lo que está bien planteado ('Aspectos conformes', 'Aciertos identificados') sin adjetivos hiperbólicos ni elogios excesivos ('excelente', 'maravilloso', 'brillante'). Mantener la serenidad y distancia profesional de un comité evaluador doctoral.
   - **Enfoque asertivo y constructivo para ajustes:** Usar fórmulas como 'Oportunidad de mejora', 'Para mayor solidez ante el jurado', 'Ajuste metodológico necesario', evitando tanto el lenguaje punitivo como la autoayuda excesiva.
   - **Sugerencias imperativas con texto listo para insertar:** Cada acción correctiva debe ser directiva ('Reemplaza...', 'Inserta el siguiente párrafo...'), redactando directamente el modelo ajustado a sus variables listo para copiar y pegar.
   - **Enlaces directos de un clic (cero errores 404 y cero fricción):** Para no obligar al estudiante a copiar y pegar términos de búsqueda manuales (lo que desanima y da sensación de trabajo incompleto), el sistema entrega SIEMPRE enlaces directos de un solo clic hacia Google Académico o PubMed (`https://scholar.google.com/scholar?q=...` o `https://pubmed.ncbi.nlm.nih.gov/?term=...`) o repositorios oficiales (OMS, OPS). Estos enlaces abren directamente el artículo y su descarga en un clic sin arrojar jamás error 404.
   - **Fuentes obligatorias y corroboración preventiva:** Todo texto sugerido con datos o teorías debe incluir llamadas numéricas de cita y el bloque 'FUENTES SUGERIDAS PARA TU BIBLIOGRAFÍA' con la advertencia explícita al estudiante de corroborar personalmente la fuente en bases científicas.
   - **Aspectos éticos directos:** Prohibido solicitar o sugerir revisión por Comité de Ética. Entregar directamente el texto modelo con los principios éticos y consentimientos necesarios.
   - **Prohibido mencionar a la UPN:** En ningún lugar del reporte generado se debe mencionar a la universidad UPN (usar "la institución" o "tu universidad").
5. **Contacto en la Web:** Correo `perspectasalud@gmail.com` y WhatsApp `+51 973 152 645`.
6. **Identidad Visual (Logo):** Logotipo SVG vectorial en perspectiva isométrica de 3 columnas (`LOGO_SVG` en `app.py`).

---

## 3. Configuración Técnica, Pagos y Analíticas
- **Tarifario Multi-Paquete Implementado:**
  - 1 Evaluación: S/ 4.90 (desbloqueo inmediato).
  - 3 Evaluaciones: S/ 14.00 (desbloqueo inmediato de la 1.ª + generación automática de código con 2 usos).
  - 5 Evaluaciones: S/ 20.00 (desbloqueo inmediato de la 1.ª + generación automática de código con 4 usos).
  - Opción de QR BIM directa con WhatsApp.
- **Variables de Entorno / Streamlit Secrets (`.streamlit/secrets.toml` o `.env`):**
  - `GEMINI_API_KEY`: API Key de Google Gemini.
  - `MP_ACCESS_TOKEN`: Token de producción de Mercado Pago.
  - `ACCESS_CODE`: Contraseña maestra de administrador (desbloquea el Panel de Administrador sin pagar).
  - `PRICE`: Precio base por evaluación individual (actualmente `4.90` PEN).
- **Analíticas:**
  - **Google Analytics (GA4):** ID `G-J2NW5JD21D`, inyectado mediante `components.html` forzando `page_title: 'Perspecta Salud'` y `page_location: 'https://revisordetesis.streamlit.app/'`.
  - **FlagCounter:** Código `OA6D` en el pie de la página de inicio.
- **Persistencia de Datos (Consideración de Streamlit Cloud):**
  - Actualmente los códigos de acceso (`codes_db.json`) y el CRM de clientes (`clientes_crm.csv`) se guardan como archivos locales.
  - *Nota técnica:* Si se requiere persistencia permanente sin descargas manuales ante reinicios de Streamlit Cloud, migrar a **Google Sheets**.

---

## 4. Próximas Tareas Pendientes (Roadmap)
- Conectar el registro del CRM y la base de datos de códigos a Google Sheets vía Service Account (cuando el usuario lo solicite).

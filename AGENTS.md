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
   - **Estadística (Logros iniciales, evaluación de pertinencia de lo propuesto/realizado, tratamiento único idóneo y paso a paso en SPSS):**
     1. **Reconocimiento inicial obligatorio (`LO LOGRADO EN ESTA SECCIÓN`):** Antes de pasar a las oportunidades de mejora, iniciar siempre comentando de forma objetiva lo que el estudiante logró o planteó adecuadamente en su metodología y propuesta de manejo de datos.
     2. **Evaluación de pertinencia si el usuario ya hizo o propuso un análisis estadístico:**
        - Primero evaluar la pertinencia del análisis estadístico propuesto o realizado por la persona, indicando si está bien o si puede mejorar.
        - **Si está bien (fue acertado):** Únicamente felicitar al usuario y explicar con claridad y en palabras sencillas por qué está bien (cómo se ajusta a sus variables y responde a su pregunta de investigación). **No** indicarle qué debe hacer, **no** agregar oportunidades de mejora ni acciones recomendadas sobre la prueba estadística, y **no** incluir pasos en SPSS.
        - **Solo si no está bien (o puede mejorar / fue omitido):** Explicar en palabras sencillas por qué no es la mejor metodología a aplicar, sugerir directamente la única prueba más adecuada y explicar por qué la hemos escogido (con un ejemplo sencillo), entregar el párrafo listo para insertar y dar el paso a paso en SPSS.
     3. **Prescripción directa de un único tratamiento estadístico (PROHIBIDO usar 'ruta/rutas' o caminos condicionales):** Analizar internamente el diseño, preguntas, objetivos, hipótesis y escalas de los datos para recomendar directamente el único tratamiento estadístico más adecuado para responder a la pregunta de investigación. QUEDA PROHIBIDO usar la palabra `ruta` o `rutas`, compartir procesos deliberativos o plantear condiciones tipo "si es normal usa X y si no usa Y".
     4. **Explicación sencilla con ejemplo práctico:** Explicar en palabras sencillas qué hace la prueba elegida, por qué la hemos escogido y cómo resuelve la pregunta de investigación, incorporando obligatoriamente un ejemplo ilustrativo adaptado a las variables del estudiante si amerita.
     5. **Sugerencia y guía paso a paso en SPSS:** Indicar que si se desea realizar en IBM SPSS se deben seguir los pasos numerados detallados: menú exacto al que ingresar (`Analizar > ...`), casilla donde colocar cada variable, botones/opciones a activar y regla sencilla de interpretación del resultado (`Sig.` o p-valor < 0.05).
   - **Sobriedad y moderación en el reconocimiento de aciertos:** En cada sección iniciar siempre con el bloque `LO LOGRADO EN ESTA SECCIÓN:` señalando de forma objetiva y serena lo que está bien planteado sin adjetivos hiperbólicos ni elogios excesivos ('excelente', 'maravilloso', 'brillante'). Mantener la serenidad y distancia profesional de un comité evaluador doctoral.
   - **Enfoque asertivo y constructivo para ajustes:** Usar fórmulas como 'Oportunidad de mejora', 'Para mayor solidez ante el jurado', 'Ajuste metodológico necesario', evitando tanto el lenguaje punitivo como la autoayuda excesiva.
   - **Sugerencias imperativas con texto listo para insertar:** Cada acción correctiva debe ser directiva ('Reemplaza...', 'Inserta el siguiente párrafo...'), redactando directamente el modelo ajustado a sus variables listo para copiar y pegar.
   - **Búsqueda en tiempo real de fuentes reales (`google_search` activo):** El motor utiliza Google Search Grounding en vivo para encontrar artículos científicos reales y vigentes, entregando su enlace directo a la revista o su DOI real verificado (`https://doi.org/...`), sin usar enlaces genéricos de búsqueda de Google Scholar.
   - **Fuentes obligatorias y advertencia de revisión por el usuario:** Todo texto sugerido con datos o teorías debe incluir llamadas numéricas de cita y el bloque 'FUENTES SUGERIDAS PARA TU BIBLIOGRAFÍA', acompañado SIEMPRE de la advertencia explícita indicando al estudiante que debe revisar, leer y corroborar personalmente el artículo y su enlace antes de incorporarlo a su versión final.
   - **Aspectos éticos directos:** Prohibido solicitar o sugerir revisión por Comité de Ética. Entregar directamente el texto modelo con los principios éticos y consentimientos necesarios.
   - **Prohibido mencionar a la UPN:** En ningún lugar del reporte generado se debe mencionar a la universidad UPN (usar "la institución" o "tu universidad").
5. **Contacto en la Web:** Correo `perspectasalud@gmail.com` y WhatsApp `+51 973 152 645`.
6. **Identidad Visual (Logo):** Logotipo SVG vectorial en perspectiva isométrica de 3 columnas (`LOGO_SVG` en `app.py`).

---

## 3. Configuración Técnica, Pagos y Analíticas
- **Motor de Inteligencia Artificial (Estrategia C - Pro + Resiliencia Flash):**
  - Modelo Primario: `gemini-pro-latest` y `gemini-3.1-pro-preview` (familia Gemini Pro), que ofrece máxima profundidad analítica, razonamiento doctoral avanzado, menor alucinación y prosa académica impecable por ~S/ 0.35 por revisión (manteniendo ~92% de margen bruto).
  - Cascada de Resiliencia (Fallback): Si la API de Google reporta saturación o lentitud en los servidores Pro, el sistema conmuta automáticamente a la familia Flash (`gemini-3.8-flash`, `gemini-3.7-flash`, etc.) para asegurar que el usuario nunca experimente fallos en su entrega.
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
- **Persistencia Permanente en Base de Datos (Supabase PostgreSQL):**
  - Configurada mediante `DATABASE_URL` en Streamlit Secrets / `.env` con `psycopg2-binary`.
  - Tablas exclusivas aisladas de otros proyectos: `perspecta_codes` (códigos y usos), `perspecta_codes_log` (historial de IPs y ubicaciones) y `perspecta_crm` (registro de clientes CRM).
  - Inmune a reinicios o actualizaciones (`git push`) de Streamlit Community Cloud.

---

## 4. Próximas Tareas Pendientes (Roadmap)
- Sistema 100% operativo con motor Gemini Pro + búsqueda en tiempo real y persistencia permanente en PostgreSQL.

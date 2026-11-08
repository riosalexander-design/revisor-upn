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
     1. Explicar con total claridad la naturaleza y escala de medición de las variables del estudio (nominales, ordinales, intervalo/razón, continuas o discretas).
     2. Articular las rutas al objetivo del estudio (correlacional, comparativo, explicativo/asociativo) y a las hipótesis.
     3. Plantear obligatoriamente las dos mejores rutas: RUTA A (paramétrica con supuestos de normalidad) y RUTA B (no paramétrica para ordinales o libre distribución).
     4. Detallar la operatividad paso a paso en SPSS: comprobación previa de normalidad (Kolmogorov-Smirnov con Lilliefors para n > 50 o Shapiro-Wilk para n < 50), ruta exacta en los menús de SPSS (`Analizar > ...`), casillas a marcar y regla de decisión del p-valor según el nivel académico.
   - **Sugerencias asertivas e imperativas con texto listo para insertar:** Cada sugerencia de corrección debe redactar directamente el párrafo modelo ajustado a las variables del estudiante para que lo copie y pegue directamente en su tesis.
   - **Rigor en referencias y DOIs sugeridos:** Al sugerir estudios o literatura académica, brindar el DOI verificado existente o la referencia completa sugiriendo al estudiante comprobarla en bases indexadas cuando no se tenga certeza plena del identificador.
   - **Fuentes obligatorias para párrafos sugeridos:** Todo párrafo o texto modelo entregado al estudiante que mencione datos estadísticos, organismos (OMS, OPS, ministerios) o teorías/autores debe incluir sus llamadas de cita numéricas (Vancouver) y obligatoriamente el bloque 'FUENTES PARA TU BIBLIOGRAFÍA' con la referencia bibliográfica completa y DOI verificado para que la copie directamente a su lista de referencias y evite citas huérfanas.
   - **Corroboración obligatoria de enlaces:** Siempre debe incluirse una indicación preventiva y explícita para que el estudiante corrobore y verifique personalmente en bases de datos científicas todos los enlaces, DOIs y referencias sugeridas antes de la entrega definitiva a su asesor o jurado.
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

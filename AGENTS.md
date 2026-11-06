# Contexto y Reglas del Proyecto: Perspecta Salud (Revisor de Tesis)

Este archivo (`AGENTS.md`) contiene la memoria histórica, decisiones de diseño, arquitectura y tareas pendientes del proyecto **Perspecta Salud** para que cualquier agente de Antigravity tenga el contexto completo desde el primer momento.

---

## 1. Descripción General y Arquitectura
- **Producto:** Plataforma web llamada **Perspecta Salud** (`https://revisordetesis.streamlit.app/`) que audita avances y proyectos de tesis universitarias utilizando IA (Google Gemini) basada en rúbricas académicas y entrega un informe descargable en formato Word (`.docx`).
- **Stack Principal:**
  - `app.py`: Interfaz principal en **Streamlit** (con diseño personalizado vía CSS inyectado, navegación con smooth scroll en JS, pasarela de Mercado Pago, sistema de códigos/tokens, CRM local y panel de administrador).
  - `reviewer.py`: Motor de evaluación con **Google Gemini API** y generación del documento Word con `python-docx`.
  - `normativas/`: Carpeta con documentos y rúbricas institucionales de referencia.

---

## 2. Reglas de Negocio Inquebrantables (Directivas del Usuario)
1. **Principio de UI Sin Fricción:** Mantener el formulario del estudiante lo más simple posible. **NO** dividir el campo de nombre del proyecto en múltiples campos (como "Título" y "Autor" por separado) para no darle más trabajo al cliente.
2. **Extracción Automática de Metadatos:** El título real y los autores del proyecto deben extraerse automáticamente del documento subido para incluirlos en el reporte.
3. **Nombre Exacto del Archivo Word Descargable:** El archivo `.docx` generado DEBE llamarse exactamente:
   `PerspectaSalud_(Nombre del autor o proyecto)_(fecha en formato ddmmaaaa).docx`
4. **Reglas del Prompt de Evaluación (`reviewer.py`):**
   - **Prohibido calificar o asignar puntajes:** No mostrar calificaciones numéricas ni restar puntos para no desanimar al cliente. Cierre motivador y asertivo.
   - **Comprensibilidad según nivel académico:** Desglosar conceptos complejos con explicaciones claras y ejemplos adaptados al nivel académico declarado (Pregrado, Maestría, Doctorado), manteniendo el rigor académico.
   - **Realidad problemática en embudo y coherencia:** Exigir estructura desde fuentes Mundiales -> Regionales -> Nacionales -> Locales. Coherencia estricta entre problema general/específicos, objetivos, dimensiones y teoría. Las dimensiones deben insinuarse y sustentarse desde la problemática y antecedentes vinculadas a la teoría.
   - **Antecedentes:** Exigir como umbral mínimo según nivel: al menos 5 internacionales y 5 nacionales para Pregrado (mínimo 10 en total); y al menos 10 internacionales y 10 nacionales para Posgrado (mínimo 20 en total). Siempre ordenados de forma cronológica ascendente (del más antiguo al más reciente) y articulados con preguntas, objetivos, dimensiones y teoría.
   - **Estadística (Naturaleza de datos y 2 mejores rutas):** Evaluar primero la naturaleza de las variables/datos (nominal, ordinal, intervalo, razón; paramétrico vs no paramétrico) y sugerir OBLIGATORIAMENTE las dos mejores rutas de análisis estadístico. Si es cuantitativo, enfocado EXCLUSIVAMENTE en SPSS con procedimientos concretos.
   - **Sugerencias asertivas e imperativas con texto listo para insertar:** Cada sugerencia de corrección debe redactar directamente el párrafo modelo ajustado a las variables del estudiante para que lo copie y pegue directamente en su tesis.
   - **Aspectos éticos directos:** Prohibido solicitar o sugerir revisión por Comité de Ética. Entregar directamente el texto modelo con los principios éticos y consentimientos necesarios.
   - **Prohibido mencionar a la UPN:** En ningún lugar del reporte generado se debe mencionar a la universidad UPN (usar "la institución" o "tu universidad").
5. **Contacto en la Web:** Correo `perspectasalud@gmail.com` y WhatsApp `+51 973 152 645`.
6. **Identidad Visual (Logo):** Logotipo SVG vectorial en perspectiva isométrica de 3 columnas (`LOGO_SVG` en `app.py`).

---

## 3. Configuración Técnica, Pagos y Analíticas
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
  - *Nota técnica:* Como Streamlit Community Cloud es efímero cuando la app entra en hibernación, está sugerido a futuro migrar `clientes_crm.csv` y `codes_db.json` a **Google Sheets** si se requiere persistencia permanente sin descargas manuales.

---

## 4. Próximas Tareas Pendientes (Roadmap Inmediato)

### A. Implementar Tarifario por Paquetes (Diseño propuesto)
Permitir a los estudiantes elegir entre 3 paquetes antes de pagar con Mercado Pago:
- **1 Evaluación:** S/ 4.90
- **3 Evaluaciones:** S/ 14.00
- **5 Evaluaciones:** S/ 20.00

**Flujo técnico acordado (sin cuentas de usuario):**
1. Mostrar un selector de paquete en la vista de revisión.
2. Generar la preferencia de Mercado Pago por el monto correspondiente al paquete elegido.
3. Si el usuario compra 1 evaluación, desbloquear la revisión inmediatamente.
4. Si compra 3 o 5 evaluaciones, desbloquear su primera revisión inmediatamente **y además generar automáticamente un código secreto (ej. `PACK-XXXX`)** en `codes_db.json` con los usos restantes (2 o 4 usos), mostrándolo claramente en pantalla para que el alumno lo guarde y lo use en sus próximas entregas.

### B. Persistencia en Google Sheets (Opcional / Futuro)
Conectar el registro del CRM y la base de datos de códigos (`codes_db.json`) a Google Sheets para evitar pérdida de datos si el servidor de Streamlit Cloud se reinicia.

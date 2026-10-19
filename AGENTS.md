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
   - **Prohibido mencionar a la UPN:** En ningún lugar del reporte generado se debe mencionar a la universidad UPN.
   - **Enfoque en SPSS:** Cuando se hagan recomendaciones de análisis de datos en estudios de tipo **cuantitativo**, el reporte debe enfocarse en el uso de **SPSS**.
5. **Contacto en la Web:** En la sección "Contáctanos" siempre debe figurar la frase `"o escríbenos a perspectasalud@gmail.com"`.
6. **Identidad Visual (Logo):** El logotipo de Perspecta Salud es un SVG vectorial en perspectiva isométrica de 3 columnas (`LOGO_SVG` en `app.py`), inyectado en la barra de navegación (`nav()`) y en el pie de página (`footer`). No usar pseudo-elementos CSS antiguos.

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

import datetime
import hmac
import os
import io
import csv
import json
import random
import string
from pathlib import Path

import mercadopago
import streamlit as st
import streamlit.components.v1 as components
import docx
import requests
from dotenv import load_dotenv

from reviewer import evaluate_project

load_dotenv()
st.set_page_config(page_title="Perspecta Salud", page_icon="📘", layout="wide", initial_sidebar_state="collapsed")

# Tarifario de Paquetes: (Evaluaciones, Precio en S/)
PACKAGES = {
    "1 Evaluación — S/ 4.90": {"evals": 1, "price": 4.90, "title": "1 Evaluación - Perspecta Salud"},
    "3 Evaluaciones — S/ 14.00 (Ahorra S/ 0.70)": {"evals": 3, "price": 14.00, "title": "Paquete 3 Evaluaciones - Perspecta Salud"},
    "5 Evaluaciones — S/ 20.00 (¡Ahorra S/ 4.50!)": {"evals": 5, "price": 20.00, "title": "Paquete 5 Evaluaciones - Perspecta Salud"},
}
PRICE = 4.90

APP_URL = st.secrets.get("APP_URL", os.getenv("APP_URL", "https://revisordetesis.streamlit.app/"))
API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
MP_TOKEN = st.secrets.get("MP_ACCESS_TOKEN", os.getenv("MP_ACCESS_TOKEN", ""))
# Aceptamos ambas para que no haya problemas si usó la mía o la nueva
ACCESS_CODE = st.secrets.get("ADMIN_PASSWORD", st.secrets.get("INSTITUTIONAL_CODE", "PERSPECTA2026"))
NORMATIVAS = Path(__file__).parent / "normativas"

# Base de datos de códigos multi-uso
CODES_FILE = "codes_db.json"
def get_code_uses(code_input):
    if not os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'w') as f: json.dump({ACCESS_CODE: 99999}, f)
    try:
        with open(CODES_FILE, 'r') as f: return json.load(f).get(code_input, 0)
    except: return 0

def deduct_code_use(code_input):
    if not os.path.exists(CODES_FILE): return
    try:
        with open(CODES_FILE, 'r') as f: db = json.load(f)
        if db.get(code_input, 0) > 0 and code_input != ACCESS_CODE:
            db[code_input] -= 1
            with open(CODES_FILE, 'w') as f: json.dump(db, f)
    except: pass

def add_new_code(new_code, uses):
    if not os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'w') as f: json.dump({ACCESS_CODE: 99999}, f)
    try:
        with open(CODES_FILE, 'r') as f: db = json.load(f)
        db[new_code] = uses
        with open(CODES_FILE, 'w') as f: json.dump(db, f)
    except: pass

def generate_pack_code(extra_uses):
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
    code = f"PACK-{suffix}"
    add_new_code(code, extra_uses)
    return code

# CRM Local Database
CRM_FILE = "clientes_crm.csv"
def registrar_operacion(proyecto, nivel, etapa):
    try:
        file_exists = os.path.isfile(CRM_FILE)
        with open(CRM_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Fecha", "Proyecto", "Nivel", "Etapa"])
            date_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            writer.writerow([date_str, proyecto, nivel, etapa])
    except:
        pass

def create_docx(report_text, project_name, author_name):
    doc = docx.Document()
    
    # Agregar Encabezado (Header)
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.text = "Perspecta Salud - Revisión Académica Inteligente"
    
    # Agregar Pie de Página (Footer)
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Contacto: perspectasalud@gmail.com | Web: https://revisordetesis.streamlit.app/"
    
    # Títulos principales en la primera página
    doc.add_heading("Perspecta Salud", 0)
    doc.add_heading("Informe de Auditoría Científica", 1)
    doc.add_heading(f"Proyecto: {project_name}", 2)
    doc.add_heading(f"Autor(es): {author_name}", 3)
    doc.add_paragraph(f"Fecha: {datetime.datetime.now().strftime('%d/%m/%Y')}")
    
    for para in report_text.split("\n\n"):
        if para.strip():
            doc.add_paragraph(para.strip())
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{--navy:#072e68;--cyan:#08b8be;--ink:#112d54;--muted:#63738c;--pale:#f3f9ff;--line:#deebf4}
html{scroll-behavior:smooth}body,[class*=css]{font-family:Inter,sans-serif;color:var(--ink)}
header[data-testid=stHeader]{height:0;background:transparent}#MainMenu,footer,[data-testid=stSidebarCollapsedControl]{display:none}.block-container{max-width:1240px;padding:0 2rem 2rem}.stMainBlockContainer{padding-top:0!important}
h1,h2,h3{color:var(--navy);letter-spacing:-.03em}.anchor{scroll-margin-top:80px}.nav{height:74px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;background:#fffffff2;border-bottom:1px solid #edf3f7;margin:0 -2rem;padding:0 2rem}.brand{font-size:22px;font-weight:800;color:var(--navy)}.brand span{color:var(--cyan)}.navlinks{display:flex;align-items:center;gap:25px}.nav a{font-size:13px;font-weight:600;text-decoration:none;color:var(--ink)}.nav .cta,.button{background:var(--cyan);color:#fff;padding:13px 19px;border-radius:8px;box-shadow:0 9px 24px #08b8be2c}
.hero{display:grid;grid-template-columns:1fr 1.08fr;align-items:center;gap:38px;min-height:505px}.eyebrow{font-size:12px;color:var(--cyan);font-weight:800;letter-spacing:.11em;text-transform:uppercase}.hero h1{font-size:clamp(42px,5vw,64px);line-height:1.06;margin:12px 0 20px}.hero p,.head p{font-size:17px;line-height:1.65;color:var(--muted)}.actions{display:flex;align-items:center;gap:22px;margin:27px 0}.button{display:inline-block;text-decoration:none!important;font-weight:700}.textlink{color:var(--navy);text-decoration:none;font-weight:700}.trust{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;font-weight:600;color:#415570}.trust span:before{content:'✓';color:var(--cyan);margin-right:6px}.hero-media{height:430px;position:relative;overflow:hidden;border-radius:0 0 0 65px}.hero-media img{width:100%;height:100%;object-fit:cover}.score{position:absolute;right:24px;top:42px;background:#fffffff2;padding:20px;border-radius:16px;width:170px;box-shadow:0 16px 42px #072e6830;text-align:center}.score b{width:88px;height:88px;margin:12px auto;display:grid;place-items:center;border-radius:50%;border:10px solid var(--cyan);font-size:27px;color:var(--navy)}.score small{color:var(--muted)}
.stats{display:grid;grid-template-columns:repeat(4,1fr);background:var(--pale);border-radius:17px;padding:24px;margin:10px 0 68px}.stat{text-align:center;border-right:1px solid var(--line)}.stat:last-child{border:0}.stat strong{font-size:23px;color:var(--navy)}.stat span{display:block;color:var(--muted);font-size:12px;margin-top:4px}.section{padding:68px 0}.head{text-align:center;max-width:680px;margin:0 auto 38px}.head h2{font-size:35px;margin:0 0 10px}.steps{display:grid;grid-template-columns:repeat(4,1fr);gap:24px}.step{text-align:center;padding:10px 18px}.number{width:58px;height:58px;border-radius:50%;background:#e8f9fa;display:grid;place-items:center;margin:auto;color:var(--cyan);font-size:25px;font-weight:800}.step h3{margin:15px 0 7px}.step p{font-size:13px;line-height:1.6;color:var(--muted)}
.soft{background:linear-gradient(#f6fbff,#fff);margin:0 -2rem;padding-left:2rem;padding-right:2rem}.areas{display:grid;grid-template-columns:repeat(5,1fr);gap:13px}.card{border:1px solid var(--line);border-radius:14px;padding:20px 15px;background:#fff;box-shadow:0 8px 25px #102b5210}.card h3{font-size:14px;min-height:36px}.meter{height:7px;background:#e2edf4;border-radius:5px;overflow:hidden}.meter i{display:block;width:78%;height:100%;background:var(--cyan)}.card small{display:block;margin-top:9px;color:var(--muted)}.human{display:grid;grid-template-columns:1fr 1fr;background:#eef8ff;border-radius:20px;overflow:hidden;margin-top:52px}.human img{width:100%;height:315px;object-fit:cover}.quote{padding:45px;display:flex;flex-direction:column;justify-content:center}.quote b{font-size:58px;color:var(--cyan);line-height:.5}.quote p{font-size:20px;line-height:1.6}.quote small{color:var(--muted)}
.report{display:grid;grid-template-columns:1.15fr .85fr;gap:25px;align-items:center}.report-card{border:1px solid var(--line);border-radius:17px;padding:25px;box-shadow:0 16px 40px #102b5214}.tabs{display:flex;gap:25px;border-bottom:1px solid var(--line);padding-bottom:12px;font-size:12px}.tabs b{color:var(--cyan)}.result{display:grid;grid-template-columns:.7fr 1.3fr;gap:25px;padding-top:25px}.donut{width:105px;height:105px;border:14px solid var(--cyan);border-radius:50%;display:grid;place-items:center;margin:auto;font-size:25px;font-weight:800;color:var(--navy)}.bar{font-size:11px;margin:12px 0}.bar i{display:block;height:7px;background:linear-gradient(90deg,var(--cyan) 78%,#e5edf4 78%);border-radius:5px;margin-top:5px}.benefits{display:grid;gap:13px}.benefits div{background:var(--pale);padding:17px;border-radius:11px}.benefits b{display:block;margin-bottom:4px}.benefits span{font-size:13px;color:var(--muted)}
.price-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px}.price-card{border:1px solid var(--line);border-radius:19px;padding:30px 26px;background:#fff;box-shadow:0 15px 40px #102b5212;display:flex;flex-direction:column;justify-content:space-between;position:relative}.price-card.featured{border:2px solid var(--cyan);box-shadow:0 18px 45px #08b8be26}.badge-pop{position:absolute;top:-13px;right:22px;background:var(--cyan);color:#fff;font-size:11px;font-weight:800;padding:5px 12px;border-radius:20px;text-transform:uppercase;letter-spacing:.05em}.amount{font-size:38px;color:var(--navy);font-weight:800;margin:10px 0 4px}.checks{line-height:1.9;color:#41536d;margin:18px 0;font-size:14px}.checks div:before{content:'✓';color:var(--cyan);font-weight:900;margin-right:7px}.security{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}.security article{border:1px solid var(--line);border-radius:13px;padding:22px}.security h3{font-size:16px}.security p,.faq p{font-size:13px;line-height:1.6;color:var(--muted)}.faq{display:grid;grid-template-columns:1fr 1fr;gap:11px}.faq details{border:1px solid var(--line);border-radius:9px;padding:16px}.faq summary{font-weight:600;cursor:pointer}.final{display:flex;align-items:center;justify-content:space-between;background:var(--navy);color:white;border-radius:19px;padding:38px 48px}.final h2{color:white;margin:0 0 7px}.final p{margin:0;color:#d9e7f7}.footer{display:flex;justify-content:space-between;align-items:center;padding:35px 0;color:var(--muted);font-size:12px}.app-head{text-align:center;padding:48px 0 22px}.app-head h1{font-size:42px}.app-head p{color:var(--muted)}div.stButton>button,div.stLinkButton>a{background:var(--cyan);border:0;color:white;border-radius:8px;min-height:48px;font-weight:700}.stDownloadButton>button{background:var(--navy)!important;color:white!important}
@media(max-width:900px){.navlinks a:not(.cta){display:none}.hero,.human,.report,.price-grid{grid-template-columns:1fr}.hero{padding-top:35px}.stats,.steps{grid-template-columns:1fr 1fr}.areas{grid-template-columns:1fr 1fr}.security{grid-template-columns:1fr}.hero-media{border-radius:20px}.faq{grid-template-columns:1fr}.final{align-items:flex-start;gap:20px;flex-direction:column}}
@media(max-width:560px){.block-container{padding:0 1rem 1rem}.nav{margin:0 -1rem;padding:0 1rem}.brand{font-size:16px}.brand:before{width:27px;height:33px}.stats,.steps,.areas{grid-template-columns:1fr}.stat{border:0;border-bottom:1px solid var(--line);padding:10px}.soft{margin:0 -1rem;padding-left:1rem;padding-right:1rem}.hero h1{font-size:40px}.hero-media{height:310px}.score{width:140px;padding:12px}.result{grid-template-columns:1fr}.footer{flex-direction:column;gap:18px}.section{padding:52px 0}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# Google Analytics Inject (IFrame con sobreescritura de URL para GA4)
ga_code = """
<script async src="https://www.googletagmanager.com/gtag/js?id=G-J2NW5JD21D"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'G-J2NW5JD21D', {
    'page_title' : 'Perspecta Salud',
    'page_location': 'https://revisordetesis.streamlit.app/'
  });
</script>
"""
components.html(ga_code, height=0, width=0)

LOGO_SVG = """<svg style="margin-right:8px; margin-bottom:-4px;" width="28" height="28" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg"><polygon points="10,30 30,18 30,82 10,70" fill="#08b8be"/><polygon points="35,15 65,15 65,85 35,85" fill="#072e68"/><polygon points="46,40 54,40 54,46 60,46 60,54 54,54 54,60 46,60 46,54 40,54 40,46 46,46" fill="#ffffff"/><polygon points="70,18 90,30 90,70 70,82" fill="#08b8be"/></svg>"""

def nav():
    st.markdown(f"""<nav class="nav"><div class="brand" style="display:flex;align-items:center;">{LOGO_SVG}Perspecta <span>Salud</span></div><div class="navlinks"><a href="#inicio" onclick="document.getElementById('inicio').scrollIntoView({{behavior: 'smooth'}}); return false;">Inicio</a><a href="#como" onclick="document.getElementById('como').scrollIntoView({{behavior: 'smooth'}}); return false;">Cómo funciona</a><a href="#evaluamos" onclick="document.getElementById('evaluamos').scrollIntoView({{behavior: 'smooth'}}); return false;">Qué evaluamos</a><a href="#informe" onclick="document.getElementById('informe').scrollIntoView({{behavior: 'smooth'}}); return false;">Informe</a><a href="#empezar" onclick="document.getElementById('empezar').scrollIntoView({{behavior: 'smooth'}}); return false;">Tarifario</a><a href="#faq" onclick="document.getElementById('faq').scrollIntoView({{behavior: 'smooth'}}); return false;">Preguntas frecuentes</a><a class="cta" href="?page=review">Revisar mi proyecto →</a></div></nav>""", unsafe_allow_html=True)


def verify_payment(payment_id):
    if not MP_TOKEN or not payment_id:
        return 0
    try:
        data = mercadopago.SDK(MP_TOKEN).payment().get(str(payment_id)).get("response", {})
        if data.get("status") == "approved" and data.get("currency_id") == "PEN":
            amt = float(data.get("transaction_amount", 0))
            if amt >= 19.5:
                return 5
            elif amt >= 13.5:
                return 3
            elif amt >= 4.5:
                return 1
        return 0
    except Exception:
        return 0


def payment_link(price_amount, item_title):
    data = {"items":[{"title":item_title,"quantity":1,"unit_price":float(price_amount),"currency_id":"PEN"}],"back_urls":{"success":APP_URL,"failure":APP_URL,"pending":APP_URL},"auto_return":"approved","statement_descriptor":"PERSPECTA SALUD"}
    return mercadopago.SDK(MP_TOKEN).preference().create(data)["response"]["init_point"]


def landing():
    nav()
    st.markdown(f"""
    <main><section id="inicio" class="hero anchor"><div><div class="eyebrow">Revisión académica inteligente</div><h1>Mejora tu proyecto de investigación con una revisión inteligente</h1><p>Recibe observaciones claras, criterios de cumplimiento y recomendaciones para fortalecer tu trabajo académico.</p><div class="actions"><a class="button" href="?page=review">Revisar mi proyecto →</a><a class="textlink" href="#informe" onclick="document.getElementById('informe').scrollIntoView({{behavior: 'smooth'}}); return false;">Ver informe de ejemplo ›</a></div><div class="trust"><span>Basada en rúbrica</span><span>Informe descargable en Word</span><span>Documentos confidenciales</span></div></div><div class="hero-media"><img src="https://images.unsplash.com/photo-1541339907198-e08756dedf3f?auto=format&fit=crop&w=1200&q=86" alt="Estudiante trabajando en su investigación"><div class="score"><small>Resultado general</small><b>78</b><small>Buen avance</small></div></div></section>
    <section class="stats"><div class="stat"><strong>5</strong><span>áreas evaluadas</span></div><div class="stat"><strong>⏱</strong><span>Informe en minutos</span></div><div class="stat"><strong>24/7</strong><span>Siempre disponible</span></div><div class="stat"><strong>Desde S/ 4.90</strong><span>Planes flexibles</span></div></section>
    <section id="como" class="section anchor"><div class="head"><h2>¿Cómo funciona?</h2><p>Cuatro pasos sencillos para saber qué debes mejorar.</p></div><div class="steps"><article class="step"><div class="number">1</div><h3>Selecciona</h3><p>Elige el paquete de revisión que se adapta a tu proyecto.</p></article><article class="step"><div class="number">2</div><h3>Paga</h3><p>Realiza el pago en línea de forma segura.</p></article><article class="step"><div class="number">3</div><h3>Carga</h3><p>Sube tu proyecto y anexos relevantes.</p></article><article class="step"><div class="number">4</div><h3>Descarga</h3><p>Recibe observaciones y recomendaciones en formato Word profesional.</p></article></div></section>
    <section id="evaluamos" class="section soft anchor"><div class="head"><h2>Descubre qué debes mejorar</h2><p>Evaluamos tu proyecto con criterios académicos en cinco áreas clave.</p></div><div class="areas"><article class="card"><h3>⚛ Coherencia científica</h3><div class="meter"><i style="width:82%"></i></div><small>Problema, objetivos e hipótesis</small></article><article class="card"><h3>▣ Metodología</h3><div class="meter"><i style="width:74%"></i></div><small>Diseño, población e instrumentos</small></article><article class="card"><h3>▤ Fuentes bibliográficas</h3><div class="meter"><i style="width:76%"></i></div><small>Actualidad y pertinencia</small></article><article class="card"><h3>❞ Estilo de citación</h3><div class="meter"><i style="width:80%"></i></div><small>Aplicación del estilo solicitado</small></article><article class="card"><h3>♢ Normativa institucional</h3><div class="meter"><i></i></div><small>Estructura y requisitos formales</small></article></div><div class="human"><img src="https://images.unsplash.com/photo-1523240795612-9a054b0db644?auto=format&fit=crop&w=1200&q=86" alt="Estudiantes revisando un proyecto"><div class="quote"><b>“</b><p>Observaciones claras para que sepas por dónde empezar y qué aspectos requieren mayor atención.</p><small>Una revisión diseñada para estudiantes y asesores</small></div></div></section>
    <section id="informe" class="section anchor"><div class="head"><h2>Conoce el informe que recibirás</h2><p>Resultados organizados para convertir las observaciones en acciones concretas.</p></div><div class="report"><div class="report-card"><div class="tabs"><b>Resumen</b><span>Observaciones</span><span>Recomendaciones</span></div><div class="result"><div><div class="donut">78</div><p style="text-align:center;color:var(--cyan);font-weight:700">Buen avance</p></div><div><div class="bar">Coherencia científica<i></i></div><div class="bar">Metodología<i style="background:linear-gradient(90deg,var(--cyan) 74%,#e5edf4 74%)"></i></div><div class="bar">Fuentes bibliográficas<i style="background:linear-gradient(90deg,var(--cyan) 76%,#e5edf4 76%)"></i></div><div class="bar">Estilo de citación<i style="background:linear-gradient(90deg,var(--cyan) 80%,#e5edf4 80%)"></i></div></div></div></div><div class="benefits"><div><b>✓ Observaciones claras</b><span>Identifica los puntos que necesitan atención.</span></div><div><b>✓ Recomendaciones aplicables</b><span>Conoce qué acción realizar en cada criterio.</span></div><div><b>🔒 Documento confidencial</b><span>Solo tú tendrás acceso al informe.</span></div></div></div></section>
    <section id="empezar" class="section anchor"><div class="head"><h2>Planes y Tarifario</h2><p>Elige una revisión individual o ahorra con nuestros paquetes de seguimiento para todo tu ciclo.</p></div><div class="price-grid"><div class="price-card"><div><h3>1 Evaluación</h3><p style="color:var(--muted);font-size:14px">Ideal para una revisión puntual antes de entregar.</p><div class="amount">S/ 4.90</div><small style="color:var(--muted)">Pago único</small><div class="checks"><div>1 evaluación completa con rúbrica</div><div>Observaciones y recomendaciones</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review">Elegir 1 Evaluación →</a></div><div class="price-card featured"><div class="badge-pop">Más elegido</div><div><h3>3 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Perfecto para corregir y validar tus avances.</p><div class="amount">S/ 14.00</div><small style="color:var(--cyan);font-weight:700">S/ 4.66 por evaluación</small><div class="checks"><div>1 evaluación inmediata + código por 2 usos</div><div>Úsalas cuando quieras sin caducidad</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review">Elegir Pack de 3 →</a></div><div class="price-card"><div><h3>5 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Para acompañar toda tu tesis o compartir con tu grupo.</p><div class="amount">S/ 20.00</div><small style="color:var(--cyan);font-weight:700">¡Solo S/ 4.00 por evaluación!</small><div class="checks"><div>1 evaluación inmediata + código por 4 usos</div><div>Máximo ahorro (Ahorras S/ 4.50)</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review">Elegir Pack de 5 →</a></div></div></section>
    <section class="section"><div class="head"><h2>Tu investigación merece confidencialidad</h2></div><div class="security"><article><h3>🛡 Archivos protegidos</h3><p>Se procesan únicamente para generar la evaluación.</p></article><article><h3>🧠 Uso responsable de IA</h3><p>La IA asiste; la decisión académica final corresponde a expertos.</p></article><article><h3>◷ Eliminación programada</h3><p>Los archivos temporales no se conservan indefinidamente.</p></article></div></section>
    <section id="faq" class="section anchor"><div class="head"><h2>Preguntas frecuentes</h2></div><div class="faq"><details><summary>¿Qué proyectos pueden revisar?</summary><p>Trabajos de pregrado y posgrado compatibles con la rúbrica habilitada.</p></details><details><summary>¿Cómo funcionan los paquetes de 3 y 5 evaluaciones?</summary><p>Al pagar realizas tu primera evaluación de inmediato y el sistema te entrega en pantalla un código secreto (ej. PACK-XXXX) para usar tus evaluaciones restantes cuando quieras.</p></details><details><summary>¿Cuánto demora?</summary><p>Generalmente pocos minutos, según la extensión.</p></details><details><summary>¿Reemplaza a mi asesor?</summary><p>No. Complementa la orientación académica profesional.</p></details></div></section>
    <section class="final"><div><h2>¿Listo para mejorar tu proyecto?</h2><p>Identifica oportunidades de mejora y lleva tu investigación al siguiente nivel.</p></div><a class="button" href="?page=review">Revisar mi proyecto →</a></section>
    """, unsafe_allow_html=True)
    
    # Bloque para Iniciar / FlagCounter / Formulario
    _, center, _ = st.columns([1, 1.25, 1])
    with center:
        if st.button("Revisar mi proyecto desde S/ 4.90", type="primary", use_container_width=True):
            st.session_state.page = "review"
            st.rerun()
            
        st.divider()
        st.markdown("<div style='text-align: center'><h3>✉️ Contáctanos</h3><p>¿Dudas? ¿Asesoría Estadística Avanzada? Usa el formulario o escríbenos a perspectasalud@gmail.com</p></div>", unsafe_allow_html=True)
        with st.form("contact_form_landing"):
            u_email = st.text_input("Tu correo electrónico:")
            u_msg = st.text_area("Mensaje:")
            if st.form_submit_button("Enviar Mensaje", use_container_width=True):
                contact_email = st.secrets.get("CONTACT_EMAIL", "")
                if contact_email and u_email:
                    try:
                        requests.post(f"https://formsubmit.co/ajax/{contact_email}", json={"Email": u_email, "Mensaje": u_msg})
                        st.success("¡Mensaje enviado con éxito!")
                    except:
                        st.error("Error enviando el mensaje.")

        st.markdown("<div style='text-align: center; margin-top: 30px;'><p>Estadísticas de la plataforma:</p></div>", unsafe_allow_html=True)
        html_code = """<div style="text-align: center;"><a href="https://info.flagcounter.com/OA6D"><img src="https://s01.flagcounter.com/countxl/OA6D/bg_FFFFFF/txt_000000/border_CCCCCC/columns_2/maxflags_4/viewers_Usuarios/labels_0/pageviews_1/flags_0/percent_0/" alt="Flag Counter" border="0"></a></div>"""
        components.html(html_code, height=180)

    st.markdown(f'<footer class="footer"><div class="brand" style="display:flex;align-items:center;">{LOGO_SVG}Perspecta <span>Salud</span></div><div>Términos · Privacidad · Soporte · Contacto</div><div>© 2026 Perspecta Salud</div></footer></main><div id="activar" class="anchor"></div>', unsafe_allow_html=True)


def review():
    nav()
    st.markdown('<div class="app-head"><div class="eyebrow">Revisión académica</div><h1>Revisa tu proyecto</h1><p>Completa el proceso y recibe un informe organizado.</p></div>', unsafe_allow_html=True)
    paid = bool(st.session_state.get("paid"))
    
    with st.container(border=True):
        if not paid:
            st.subheader("1. Elige tu paquete y activa tu revisión")
            selected_pkg_label = st.radio(
                "Selecciona un plan de evaluación:",
                list(PACKAGES.keys()),
                index=0,
                help="Si eliges el paquete de 3 o 5 evaluaciones, podrás hacer tu primera revisión hoy y recibirás un código secreto en pantalla para usar las demás cuando quieras."
            )
            pkg = PACKAGES[selected_pkg_label]
            st.write(f"Total a pagar: **S/ {pkg['price']:.2f}** ({pkg['evals']} {'evaluación' if pkg['evals'] == 1 else 'evaluaciones'}). El acceso se habilita automáticamente al comprobar el pago.")
            if MP_TOKEN:
                try:
                    st.link_button(
                        f"Pagar S/ {pkg['price']:.2f} de forma segura con Mercado Pago",
                        payment_link(pkg["price"], pkg["title"]),
                        type="primary",
                        use_container_width=True
                    )
                except Exception:
                    st.error("No pudimos crear el enlace de pago. Inténtalo nuevamente.")
            else:
                st.warning("El sistema de pagos está temporalmente en mantenimiento.")
            
            with st.expander("🔑 ¿Tienes un código de paquete, promoción o acceso institucional?"):
                inst_code = st.text_input("Ingresa tu código secreto / pase", type="password")
                if st.button("Desbloquear con código", type="secondary"):
                    uses_left = get_code_uses(inst_code.strip())
                    if uses_left > 0:
                        st.session_state.paid = True
                        st.session_state.used_code = inst_code.strip()
                        if inst_code.strip() == ACCESS_CODE:
                            st.session_state.is_admin = True
                        st.rerun()
                    else:
                        st.error("Código inválido o sin usos disponibles.")
            if st.button("← Volver al inicio"): st.session_state.page="home"; st.rerun()
            return

        # ==================== APLICACIÓN DESBLOQUEADA ====================
        st.success("Acceso confirmado. Ya puedes cargar tu proyecto.")
        if st.session_state.get("generated_pack_code"):
            extra_evals = st.session_state.get("generated_pack_uses", 0)
            st.info(
                f"🎁 **¡Compraste un paquete múltiple!** Guarda ahora mismo tu código secreto: **`{st.session_state['generated_pack_code']}`**\n\n"
                f"Esta sesión cubre tu **1.ª evaluación**. Tu código **`{st.session_state['generated_pack_code']}`** tiene **{extra_evals} evaluaciones adicionales** listas para que las uses cuando quieras (ingresándolo en la opción *'¿Tienes un código de paquete?'*)."
            )
        elif st.session_state.get("used_code") and not st.session_state.get("is_admin"):
            rem = get_code_uses(st.session_state["used_code"])
            st.caption(f"🔑 Código activo: `{st.session_state['used_code']}` (Usos disponibles antes de esta revisión: {rem})")
        
        # Panel Secreto del Administrador
        if st.session_state.get("is_admin"):
            st.markdown("---")
            st.subheader("🛠️ Panel de Administrador")
            
            # Gestión de Códigos
            with st.expander("🔑 Generar códigos de acceso (Tokens)"):
                st.write("Crea contraseñas de uso limitado para que tus alumnos no tengan que pagar con tarjeta.")
                col1, col2 = st.columns(2)
                new_code_name = col1.text_input("Nuevo código", placeholder="Ej. PROMOCION2026")
                new_code_uses = col2.number_input("Número de usos", min_value=1, max_value=100, value=1)
                if st.button("Crear código", type="primary"):
                    if new_code_name.strip():
                        add_new_code(new_code_name.strip(), new_code_uses)
                        st.success(f"Código `{new_code_name.strip()}` creado con {new_code_uses} usos.")
                    else:
                        st.error("Ingresa un nombre para el código.")
                
                st.write("---")
                try:
                    with open(CODES_FILE, 'r') as f:
                        db = json.load(f)
                    if len(db) > 1:
                        st.write("**Códigos activos:**")
                        for k, v in db.items():
                            if k != ACCESS_CODE:
                                st.write(f"- `{k}`: {v} usos restantes")
                except: pass
            
            # Base de datos CRM
            try:
                with open(CRM_FILE, "r", encoding="utf-8") as f:
                    st.download_button("📥 Descargar Base de Datos CRM", f, file_name="clientes_crm.csv", mime="text/csv", use_container_width=True)
            except: pass
            st.info("💡 Como Administrador, puedes auditar otro proyecto sin tener que volver a pagar ni recargar.")
            st.divider()
        
        # EL CANDADO DE 1 SOLO USO: Si ya hay reporte, ocultamos el formulario
        if st.session_state.get("report"):
            st.warning("⚠️ **IMPORTANTE:** Descarga tu reporte en formato Word ahora. Si recargas o cierras esta página, tu sesión de pago expirará.")
            if st.session_state.get("generated_pack_code"):
                extra_evals = st.session_state.get("generated_pack_uses", 0)
                st.info(
                    f"🎁 **Recuerda guardar tu código de paquete:** **`{st.session_state['generated_pack_code']}`** "
                    f"(Te quedan **{extra_evals} evaluaciones adicionales** para usar cuando quieras)."
                )
            st.subheader("Resultado de la revisión")
            st.write(st.session_state.report)
            
            # Generar Word File
            date=datetime.datetime.now().strftime("%d%m%Y")
            client_name = st.session_state.get("client_name", st.session_state.get("project", ""))
            safe="".join(c for c in client_name if c.isalnum() or c in "-_ ").strip().replace(" ", "_")
            if not safe: safe = "Proyecto"
            
            docx_data = create_docx(st.session_state.report, st.session_state.project, st.session_state.get("author", "Autor Desconocido"))
            st.download_button(
                label="📥 Descargar Informe Completo en Word (.docx)", 
                data=docx_data, 
                file_name=f"PerspectaSalud_{safe}_{date}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                use_container_width=True
            )
            
            if st.session_state.get("is_admin"):
                st.info("💡 Como Administrador, puedes auditar otro proyecto sin tener que volver a pagar ni recargar.")
                if st.button("🔄 Auditar Nuevo Proyecto", type="primary", use_container_width=True):
                    del st.session_state.report
                    st.rerun()
            else:
                st.info("Para auditar un nuevo proyecto, debes recargar la página (F5) e iniciar una nueva sesión.")
            return

        # EL FORMULARIO
        st.subheader("2. Datos del proyecto")
        name = st.text_input("Nombre del autor o proyecto", placeholder="Ej. Apellido – título breve")
        nivel = st.selectbox("Nivel académico", ["Pregrado", "Maestría", "Doctorado"])
        etapa = st.selectbox("Etapa del trabajo", ["Proyecto de investigación", "Tesis en desarrollo", "Tesis final"])
        
        st.write("⚠️ **Límites:** Máximo 3 archivos. Peso total máximo 200 MB.")
        files = st.file_uploader("Carga los documentos", type=["pdf","docx","txt","xlsx","csv"], accept_multiple_files=True)
        consent = st.checkbox("Confirmo que tengo autorización para procesar los documentos y acepto las condiciones de privacidad.")
        
        if st.button("Iniciar revisión", type="primary", use_container_width=True):
            if not API_KEY: st.error("El motor de evaluación no está configurado.")
            elif not name.strip(): st.error("Ingresa el nombre del proyecto.")
            elif not files: st.error("Carga al menos un documento.")
            elif len(files) > 3: st.error("Por favor, sube un máximo de 3 archivos consolidados.")
            elif not consent: st.error("Debes aceptar las condiciones de privacidad.")
            elif any(f.size > 200*1024*1024 for f in files): st.error("Cada archivo debe pesar como máximo 200 MB.")
            else:
                import time
                progress_bar = st.progress(0, text="Iniciando protocolos de lectura...")
                time.sleep(0.5)
                progress_bar.progress(30, text="Extrayendo texto y analizando metodología estadística (SPSS)...")
                time.sleep(1)
                progress_bar.progress(60, text="Verificando citas, referencias y normativa institucional...")
                time.sleep(1)
                progress_bar.progress(85, text="Cruzando datos con el panel de expertos de IA...")

                with st.status("Generando reporte de nivel doctoral...", expanded=True) as status:
                    result = evaluate_project(files, str(NORMATIVAS), API_KEY)
                    status.update(label="Auditoría Completada", state="complete", expanded=False)
                
                progress_bar.progress(100, text="¡Auditoría completada!")
                time.sleep(0.5)
                progress_bar.empty()

                if isinstance(result, str): st.error(result)
                else: 
                    # Registrar en CRM
                    registrar_operacion(name, nivel, etapa)
                    
                    if st.session_state.get("used_code"):
                        deduct_code_use(st.session_state.used_code)
                    
                    st.session_state.report=result.get("report", "No fue posible recuperar el informe.")
                    st.session_state.project=result.get("metadata", {}).get("project_title", name)
                    st.session_state.author=result.get("metadata", {}).get("authors", "Autor no especificado")
                    st.session_state.client_name = name
                    st.rerun() # Recarga para ocultar el formulario

payment_id = st.query_params.get("payment_id", "")
status = st.query_params.get("status", "")
page_param = st.query_params.get("page", "")

if page_param == "review":
    st.session_state.page = "review"
    st.query_params.clear()

if status == "approved" and payment_id:
    purchased_evals = verify_payment(payment_id)
    if purchased_evals > 0:
        st.session_state.paid = True
        st.session_state.page = "review"
        if purchased_evals > 1 and not st.session_state.get("generated_pack_code"):
            extra_uses = purchased_evals - 1
            st.session_state.generated_pack_code = generate_pack_code(extra_uses)
            st.session_state.generated_pack_uses = extra_uses
        st.query_params.clear()

st.session_state.setdefault("page", "home")
review() if st.session_state.page == "review" else landing()


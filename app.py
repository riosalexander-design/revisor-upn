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

# Base de datos de códigos multi-uso e historial de IPs
CODES_FILE = "codes_db.json"
CODES_LOG_FILE = "codes_usage_log.json"
QR_BIM_PATH = Path(__file__).parent / "qr_bim.png"

def _load_codes_db():
    if not os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'w', encoding='utf-8') as f:
            json.dump({ACCESS_CODE: 99999}, f)
    try:
        with open(CODES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {ACCESS_CODE: 99999}

def _save_codes_db(db):
    try:
        with open(CODES_FILE, 'w', encoding='utf-8') as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_code_uses(code_input):
    db = _load_codes_db()
    val = db.get(code_input, 0)
    if isinstance(val, dict):
        return val.get("remaining", 0)
    return int(val) if isinstance(val, (int, float)) else 0

def deduct_code_use(code_input):
    db = _load_codes_db()
    if code_input in db and code_input != ACCESS_CODE:
        val = db[code_input]
        if isinstance(val, dict):
            if val.get("remaining", 0) > 0:
                val["remaining"] -= 1
            db[code_input] = val
        elif isinstance(val, (int, float)) and val > 0:
            db[code_input] = int(val) - 1
        _save_codes_db(db)

def add_new_code(new_code, uses):
    db = _load_codes_db()
    db[new_code] = {
        "initial": int(uses),
        "remaining": int(uses),
        "created_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    _save_codes_db(db)

def delete_code(code_to_delete):
    db = _load_codes_db()
    if code_to_delete in db and code_to_delete != ACCESS_CODE:
        del db[code_to_delete]
        _save_codes_db(db)

def get_client_ip_and_location():
    ip = "Desconocida"
    try:
        headers = st.context.headers
        if headers:
            forwarded = headers.get("X-Forwarded-For") or headers.get("x-forwarded-for")
            real_ip = headers.get("X-Real-Ip") or headers.get("x-real-ip")
            if forwarded:
                ip = forwarded.split(",")[0].strip()
            elif real_ip:
                ip = real_ip.strip()
    except Exception:
        pass
    location = "Ubicación no disponible"
    if ip and ip != "Desconocida" and not ip.startswith(("127.", "10.", "192.168.")):
        try:
            resp = requests.get(f"http://ip-api.com/json/{ip}?fields=status,country,regionName,city,isp", timeout=2.5)
            if resp.status_code == 200:
                data = resp.json()
                if data.get("status") == "success":
                    city = data.get("city", "")
                    region = data.get("regionName", "")
                    country = data.get("country", "")
                    isp = data.get("isp", "")
                    location = f"{city}, {region}, {country} ({isp})".strip(", ")
        except Exception:
            pass
    return ip, location

def log_code_usage(code_used, action="Desbloqueo"):
    if not code_used or code_used == ACCESS_CODE:
        return
    ip, loc = get_client_ip_and_location()
    logs = []
    if os.path.exists(CODES_LOG_FILE):
        try:
            with open(CODES_LOG_FILE, 'r', encoding='utf-8') as f:
                logs = json.load(f)
        except Exception:
            logs = []
    logs.append({
        "fecha": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "codigo": code_used,
        "accion": action,
        "ip": ip,
        "ubicacion": loc
    })
    try:
        with open(CODES_LOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def generate_pack_code(extra_uses):
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
    code = f"PACK-{suffix}"
    add_new_code(code, extra_uses)
    return code

# CRM Local Database
CRM_FILE = "clientes_crm.csv"
def registrar_operacion(proyecto, nivel, etapa, codigo_usado="Pago directo"):
    try:
        ip, loc = get_client_ip_and_location()
        file_exists = os.path.isfile(CRM_FILE)
        with open(CRM_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Fecha", "Proyecto", "Nivel", "Etapa", "Codigo", "IP", "Ubicacion"])
            date_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            writer.writerow([date_str, proyecto, nivel, etapa, codigo_usado, ip, loc])
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
    footer_para.text = "Contacto: perspectasalud@gmail.com | WhatsApp: +51 973 152 645 | Web: https://revisordetesis.streamlit.app/"
    
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
    st.markdown(f"""<nav class="nav"><div class="brand" style="display:flex;align-items:center;">{LOGO_SVG}Perspecta <span>Salud</span></div><div class="navlinks"><a class="cta" href="?page=review">Revisar mi proyecto →</a></div></nav>""", unsafe_allow_html=True)


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
    data = {
        "items": [{"title": item_title, "quantity": 1, "unit_price": float(price_amount), "currency_id": "PEN"}],
        "payment_methods": {"installments": 1, "default_installments": 1},
        "back_urls": {"success": APP_URL, "failure": APP_URL, "pending": APP_URL},
        "auto_return": "approved",
        "statement_descriptor": "PERSPECTA SALUD"
    }
    return mercadopago.SDK(MP_TOKEN).preference().create(data)["response"]["init_point"]


def landing():
    nav()
    st.markdown(f"""
    <main><section id="inicio" class="hero anchor"><div><div class="eyebrow">Revisión académica inteligente</div><h1>Mejora tu proyecto de investigación con una revisión inteligente</h1><p>Recibe observaciones claras, criterios de cumplimiento y recomendaciones para fortalecer tu trabajo académico.</p><div class="actions"><a class="button" href="?page=review">Revisar mi proyecto →</a><a class="textlink" href="#informe" onclick="document.getElementById('informe').scrollIntoView({{behavior: 'smooth'}}); return false;">Ver informe de ejemplo ›</a></div><div class="trust"><span>Basada en rúbrica</span><span>Informe descargable en Word</span><span>Documentos confidenciales</span></div></div><div class="hero-media"><img src="https://images.unsplash.com/photo-1541339907198-e08756dedf3f?auto=format&fit=crop&w=1200&q=86" alt="Estudiante trabajando en su investigación"><div class="score"><small>Resultado general</small><b>78</b><small>Buen avance</small></div></div></section>
    <section class="stats"><div class="stat"><strong>5</strong><span>áreas evaluadas</span></div><div class="stat"><strong>⏱</strong><span>Informe en minutos</span></div><div class="stat"><strong>24/7</strong><span>Siempre disponible</span></div><div class="stat"><strong>Desde S/ 4.90</strong><span>Planes flexibles</span></div></section>
    <section id="como" class="section anchor"><div class="head"><h2>¿Cómo funciona?</h2><p>Cuatro pasos sencillos para saber qué debes mejorar.</p></div><div class="steps"><article class="step"><div class="number">1</div><h3>Selecciona</h3><p>Elige el paquete de revisión que se adapta a tu proyecto.</p></article><article class="step"><div class="number">2</div><h3>Paga</h3><p>Realiza el pago en línea de forma segura.</p></article><article class="step"><div class="number">3</div><h3>Carga</h3><p>Sube tu proyecto y anexos relevantes.</p></article><article class="step"><div class="number">4</div><h3>Descarga</h3><p>Recibe observaciones y recomendaciones en formato Word profesional.</p></article></div></section>
    <section id="evaluamos" class="section soft anchor"><div class="head"><h2>Descubre qué debes mejorar</h2><p>Evaluamos tu proyecto con criterios académicos en cinco áreas clave.</p></div><div class="areas"><article class="card"><h3>⚛ Coherencia científica</h3><div class="meter"><i style="width:82%"></i></div><small>Problema, objetivos e hipótesis</small></article><article class="card"><h3>▣ Metodología</h3><div class="meter"><i style="width:74%"></i></div><small>Diseño, población e instrumentos</small></article><article class="card"><h3>▤ Fuentes bibliográficas</h3><div class="meter"><i style="width:76%"></i></div><small>Actualidad y pertinencia</small></article><article class="card"><h3>❞ Estilo de citación</h3><div class="meter"><i style="width:80%"></i></div><small>Aplicación del estilo solicitado</small></article><article class="card"><h3>♢ Normativa institucional</h3><div class="meter"><i></i></div><small>Estructura y requisitos formales</small></article></div><div class="human"><img src="https://images.unsplash.com/photo-1523240795612-9a054b0db644?auto=format&fit=crop&w=1200&q=86" alt="Estudiantes revisando un proyecto"><div class="quote"><b>“</b><p>Observaciones claras para que sepas por dónde empezar y qué aspectos requieren mayor atención.</p><small>Una revisión diseñada para estudiantes y asesores</small></div></div></section>
    <section id="informe" class="section anchor"><div class="head"><h2>Conoce el informe que recibirás</h2><p>Resultados organizados para convertir las observaciones en acciones concretas.</p></div><div class="report"><div class="report-card"><div class="tabs"><b>Resumen</b><span>Observaciones</span><span>Recomendaciones</span></div><div class="result"><div><div class="donut">78</div><p style="text-align:center;color:var(--cyan);font-weight:700">Buen avance</p></div><div><div class="bar">Coherencia científica<i></i></div><div class="bar">Metodología<i style="background:linear-gradient(90deg,var(--cyan) 74%,#e5edf4 74%)"></i></div><div class="bar">Fuentes bibliográficas<i style="background:linear-gradient(90deg,var(--cyan) 76%,#e5edf4 76%)"></i></div><div class="bar">Estilo de citación<i style="background:linear-gradient(90deg,var(--cyan) 80%,#e5edf4 80%)"></i></div></div></div></div><div class="benefits"><div><b>✓ Observaciones claras</b><span>Identifica los puntos que necesitan atención.</span></div><div><b>✓ Recomendaciones aplicables</b><span>Conoce qué acción realizar en cada criterio.</span></div><div><b>🔒 Documento confidencial</b><span>Solo tú tendrás acceso al informe.</span></div></div></div></section>
    <section id="empezar" class="section anchor"><div class="head"><h2>Planes y Tarifario</h2><p>Elige una revisión individual o ahorra con nuestros paquetes de seguimiento para todo tu ciclo.</p></div><div class="price-grid"><div class="price-card"><div><h3>1 Evaluación</h3><p style="color:var(--muted);font-size:14px">Ideal para una revisión puntual antes de entregar.</p><div class="amount">S/ 4.90</div><small style="color:var(--muted)">Pago único</small><div class="checks"><div>1 evaluación completa con rúbrica</div><div>Observaciones y recomendaciones</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=1">Elegir 1 Evaluación →</a></div><div class="price-card featured"><div class="badge-pop">Más elegido</div><div><h3>3 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Perfecto para corregir y validar tus avances.</p><div class="amount">S/ 14.00</div><small style="color:var(--cyan);font-weight:700">S/ 4.66 por evaluación</small><div class="checks"><div>1 evaluación inmediata + código por 2 usos</div><div>Úsalas cuando quieras sin caducidad</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=3">Elegir Pack de 3 →</a></div><div class="price-card"><div><h3>5 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Para acompañar toda tu tesis o compartir con tu grupo.</p><div class="amount">S/ 20.00</div><small style="color:var(--cyan);font-weight:700">¡Solo S/ 4.00 por evaluación!</small><div class="checks"><div>1 evaluación inmediata + código por 4 usos</div><div>Máximo ahorro (Ahorras S/ 4.50)</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=5">Elegir Pack de 5 →</a></div></div></section>
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
        st.markdown(
            """
            <div style="text-align: center; background: #f3f9ff; border: 1px solid #deebf4; border-radius: 16px; padding: 28px 20px; margin: 10px 0;">
                <h3 style="margin-top: 0;">✉️ Contáctanos</h3>
                <p style="color: #63738c; margin-bottom: 18px;">¿Dudas o necesitas Asesoría Estadística Avanzada? Comunícate directamente con nosotros:</p>
                <div style="font-size: 17px; line-height: 2;">
                    <div>📧 <b>Correo:</b> <a href="mailto:perspectasalud@gmail.com" style="color: #072e68; text-decoration: none; font-weight: 600;">perspectasalud@gmail.com</a></div>
                    <div>💬 <b>WhatsApp:</b> <a href="https://wa.me/51973152645" target="_blank" style="color: #08b8be; text-decoration: none; font-weight: 700;">+51 973 152 645</a></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<div style='text-align: center; margin-top: 30px;'><p>Estadísticas de la plataforma:</p></div>", unsafe_allow_html=True)
        html_code = """<div style="text-align: center;"><a href="https://info.flagcounter.com/OA6D"><img src="https://s01.flagcounter.com/countxl/OA6D/bg_FFFFFF/txt_000000/border_CCCCCC/columns_2/maxflags_4/viewers_Usuarios/labels_0/pageviews_1/flags_0/percent_0/" alt="Flag Counter" border="0"></a></div>"""
        components.html(html_code, height=180)

    st.markdown(f'<footer class="footer"><div class="brand" style="display:flex;align-items:center;">{LOGO_SVG}Perspecta <span>Salud</span></div><div>Correo: perspectasalud@gmail.com · WhatsApp: +51 973 152 645</div><div>© 2026 Perspecta Salud</div></footer></main><div id="activar" class="anchor"></div>', unsafe_allow_html=True)


def review():
    nav()
    st.markdown('<div class="app-head"><div class="eyebrow">Revisión académica</div><h1>Revisa tu proyecto</h1><p>Completa el proceso y recibe un informe organizado.</p></div>', unsafe_allow_html=True)
    paid = bool(st.session_state.get("paid"))
    
    with st.container(border=True):
        if not paid:
            # 1. BLOQUE GRANDE Y VISIBLE PARA INGRESAR CÓDIGO
            st.markdown(
                """
                <div style="background: linear-gradient(135deg, #eef8ff 0%, #e6f9fa 100%); border: 2px solid #08b8be; border-radius: 16px; padding: 22px 26px; margin-bottom: 14px;">
                    <div style="font-size: 12px; font-weight: 800; color: #08b8be; text-transform: uppercase; letter-spacing: .08em;">¿Ya tienes un código de paquete, promoción o QR?</div>
                    <h2 style="margin: 4px 0 6px 0; font-size: 26px; color: #072e68;">🔑 Ingresa tu código de activación aquí</h2>
                    <p style="margin: 0; color: #415570; font-size: 15px;">Si compraste un paquete de 3 o 5 evaluaciones, pagaste por QR o tienes un cupón promocional, escríbelo abajo para desbloquear tu revisión de inmediato:</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            col_code_in, col_code_btn = st.columns([3, 1])
            with col_code_in:
                inst_code = st.text_input(
                    "Código de activación",
                    type="password",
                    placeholder="Escribe aquí tu código (Ej. PACK-XXXXX o GANATIEMPO)...",
                    label_visibility="collapsed"
                )
            with col_code_btn:
                unlock_clicked = st.button("🔓 Activar Código →", type="primary", use_container_width=True)
            if unlock_clicked:
                clean_code = inst_code.strip()
                uses_left = get_code_uses(clean_code)
                if uses_left > 0:
                    st.session_state.paid = True
                    st.session_state.used_code = clean_code
                    if clean_code == ACCESS_CODE:
                        st.session_state.is_admin = True
                    else:
                        log_code_usage(clean_code, action="Ingreso de código (Desbloqueo)")
                    st.rerun()
                else:
                    st.error("❌ Código inválido o sin usos disponibles. Verifica que esté bien escrito.")

            st.markdown("---")

            # 2. LOS 3 BOTONES DIRECTOS DE MERCADO PAGO (VISIBLES AL MISMO TIEMPO)
            st.subheader("💳 ¿Aún no tienes un código? Elige tu plan y paga aquí")
            st.write("Selecciona directamente el paquete que prefieras. **Con Mercado Pago el acceso se activa de forma 100% automática al instante:**")

            if MP_TOKEN:
                p_col1, p_col2, p_col3 = st.columns(3)
                with p_col1:
                    with st.container(border=True):
                        st.markdown("#### 1 Evaluación")
                        st.markdown("<div style='font-size:32px;font-weight:800;color:#072e68;margin:4px 0;'>S/ 4.90</div>", unsafe_allow_html=True)
                        st.caption("1 revisión inmediata con informe descargable en Word.")
                        try:
                            st.link_button(
                                "⚡ Pagar S/ 4.90 (Mercado Pago)",
                                payment_link(4.90, "1 Evaluación - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error("Error al generar link S/ 4.90")

                with p_col2:
                    with st.container(border=True):
                        st.markdown("#### 🌟 3 Evaluaciones")
                        st.markdown("<div style='font-size:32px;font-weight:800;color:#08b8be;margin:4px 0;'>S/ 14.00</div>", unsafe_allow_html=True)
                        st.caption("1 revisión hoy + Código secreto con 2 usos adicionales (Ahorras S/ 0.70).")
                        try:
                            st.link_button(
                                "⚡ Pagar S/ 14.00 (Mercado Pago)",
                                payment_link(14.00, "Paquete 3 Evaluaciones - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error("Error al generar link S/ 14.00")

                with p_col3:
                    with st.container(border=True):
                        st.markdown("#### 🔥 5 Evaluaciones")
                        st.markdown("<div style='font-size:32px;font-weight:800;color:#072e68;margin:4px 0;'>S/ 20.00</div>", unsafe_allow_html=True)
                        st.caption("1 revisión hoy + Código secreto con 4 usos adicionales (¡Ahorras S/ 4.50!).")
                        try:
                            st.link_button(
                                "⚡ Pagar S/ 20.00 (Mercado Pago)",
                                payment_link(20.00, "Paquete 5 Evaluaciones - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error("Error al generar link S/ 20.00")
            else:
                st.warning("El sistema de pagos automáticos está en mantenimiento.")

            st.markdown("---")

            # 3. OPCIÓN DE PAGO POR QR BIM
            st.markdown("#### 📱 Opción alternativa: Pago con QR directo (Billetera BIM / Otros)")
            qr_col_img, qr_col_txt = st.columns([1, 2.2])
            with qr_col_img:
                if QR_BIM_PATH.exists():
                    st.image(str(QR_BIM_PATH), caption="Escanea el QR para pagar (Bim: 913152645)", width=220)
            with qr_col_txt:
                st.write(
                    "Si prefieres pagar escaneando nuestro código QR directo (S/ 4.90, S/ 14.00 o S/ 20.00), realiza el pago y **envíanos la captura junto con tu Nombre y DNI a nuestro WhatsApp [+51 973 152 645](https://wa.me/51973152645)** para entregarte tu código de activación."
                )
                st.info(
                    "💡 **Nota amable e importante:** El envío de la captura con tu Nombre y DNI al WhatsApp **+51 973 152 645** es **únicamente para pagos realizados mediante este código QR**. "
                    "Si realizas tu pago con cualquiera de los botones de **Mercado Pago** de arriba, el proceso sigue siendo **100% automático** y no necesitas enviarnos comprobante."
                )

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
            with st.expander("🔑 Generar y Gestionar Códigos de Acceso (Tokens)", expanded=True):
                st.write("Crea contraseñas de uso limitado. Los códigos creados permanecerán visibles en la lista (incluso si llegan a 0 usos) hasta que decidas eliminarlos con su botón.")
                col1, col2 = st.columns(2)
                new_code_name = col1.text_input("Nuevo código", placeholder="Ej. GANATIEMPO")
                new_code_uses = col2.number_input("Número de usos", min_value=1, max_value=500, value=1)
                if st.button("Crear código", type="primary"):
                    if new_code_name.strip():
                        add_new_code(new_code_name.strip(), new_code_uses)
                        st.success(f"Código `{new_code_name.strip()}` creado con {new_code_uses} usos.")
                        st.rerun()
                    else:
                        st.error("Ingresa un nombre para el código.")
                
                st.write("---")
                db = _load_codes_db()
                user_codes = {k: v for k, v in db.items() if k != ACCESS_CODE}
                if user_codes:
                    st.write("**📋 Todos los códigos brindados (permanentes hasta eliminar):**")
                    for k, v in list(user_codes.items()):
                        if isinstance(v, dict):
                            rem_u = v.get("remaining", 0)
                            init_u = v.get("initial", rem_u)
                            used_u = max(0, init_u - rem_u)
                            created_str = v.get("created_at", "")
                            info_txt = f"🔑 **`{k}`** — Usados: **{used_u}/{init_u}** | Restantes: **{rem_u}**" + (f" *(Creado: {created_str})*" if created_str else "")
                        else:
                            rem_u = int(v)
                            info_txt = f"🔑 **`{k}`** — Usos restantes: **{rem_u}**"
                        c_info, c_del = st.columns([4, 1])
                        c_info.markdown(info_txt)
                        if c_del.button("🗑️ Eliminar", key=f"del_code_{k}"):
                            delete_code(k)
                            st.rerun()
                else:
                    st.caption("No hay códigos creados todavía.")

            # Historial de IPs y Ubicaciones por Código
            with st.expander("🌍 Historial de Uso de Códigos (IP y Ubicación Relativa)"):
                if os.path.exists(CODES_LOG_FILE):
                    try:
                        with open(CODES_LOG_FILE, 'r', encoding='utf-8') as f:
                            usage_logs = json.load(f)
                        if usage_logs:
                            st.dataframe(usage_logs, use_container_width=True)
                        else:
                            st.caption("Aún no se han registrado accesos con códigos.")
                    except Exception:
                        st.caption("No se pudo leer el historial de uso.")
                else:
                    st.caption("Aún no se han registrado accesos con códigos.")
            
            # Base de datos CRM
            try:
                with open(CRM_FILE, "r", encoding="utf-8") as f:
                    st.download_button("📥 Descargar Base de Datos CRM (incluye Código, IP y Ubicación)", f, file_name="clientes_crm.csv", mime="text/csv", use_container_width=True)
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
                    used_c = st.session_state.get("used_code", "Pago directo")
                    # Registrar en CRM con IP y Ubicación
                    registrar_operacion(name, nivel, etapa, codigo_usado=used_c)
                    
                    if st.session_state.get("used_code"):
                        deduct_code_use(st.session_state.used_code)
                        log_code_usage(st.session_state.used_code, action=f"Evaluación completada ({name})")
                    
                    st.session_state.report=result.get("report", "No fue posible recuperar el informe.")
                    st.session_state.project=result.get("metadata", {}).get("project_title", name)
                    st.session_state.author=result.get("metadata", {}).get("authors", "Autor no especificado")
                    st.session_state.client_name = name
                    st.rerun() # Recarga para ocultar el formulario

payment_id = st.query_params.get("payment_id", "")
status = st.query_params.get("status", "")
page_param = st.query_params.get("page", "")
pack_param = st.query_params.get("pack", "")

if page_param == "review":
    st.session_state.page = "review"
    if pack_param == "3":
        st.session_state.default_pack_idx = 1
    elif pack_param == "5":
        st.session_state.default_pack_idx = 2
    elif pack_param == "1":
        st.session_state.default_pack_idx = 0
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


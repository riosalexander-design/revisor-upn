import datetime
import hmac
import os
import io
import csv
import json
import random
import string
import time as _time
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

# Tarifario: se lee de Streamlit Secrets (PRICE, PRICE_PACK3, PRICE_PACK5) con estos defaults
PRECIO_1 = float(st.secrets.get("PRICE", os.getenv("PRICE", "4.90")))
PRECIO_3 = float(st.secrets.get("PRICE_PACK3", os.getenv("PRICE_PACK3", "12.00")))
PRECIO_5 = float(st.secrets.get("PRICE_PACK5", os.getenv("PRICE_PACK5", "20.00")))
PRECIO_ADV10 = float(st.secrets.get("PRICE_ADV10", os.getenv("PRICE_ADV10", "35.00")))
PRECIO_ADV25 = float(st.secrets.get("PRICE_ADV25", os.getenv("PRICE_ADV25", "75.00")))
PACKAGES = {
    f"1 Evaluación — S/ {PRECIO_1:.2f}": {"evals": 1, "price": PRECIO_1, "title": "1 Evaluación - Perspecta Salud"},
    f"3 Evaluaciones — S/ {PRECIO_3:.2f} (Ahorra S/ {3*PRECIO_1 - PRECIO_3:.2f})": {"evals": 3, "price": PRECIO_3, "title": "Paquete 3 Evaluaciones - Perspecta Salud"},
    f"5 Evaluaciones — S/ {PRECIO_5:.2f} (¡Ahorra S/ {5*PRECIO_1 - PRECIO_5:.2f}!)": {"evals": 5, "price": PRECIO_5, "title": "Paquete 5 Evaluaciones - Perspecta Salud"},
}
PRICE = PRECIO_1

APP_URL = st.secrets.get("APP_URL", os.getenv("APP_URL", "https://revisordetesis.streamlit.app/"))
API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
MP_TOKEN = st.secrets.get("MP_ACCESS_TOKEN", os.getenv("MP_ACCESS_TOKEN", ""))
DATABASE_URL = st.secrets.get("DATABASE_URL", os.getenv("DATABASE_URL", ""))
# Aceptamos ambas para que no haya problemas si usó la mía o la nueva
ACCESS_CODE = st.secrets.get("ADMIN_PASSWORD", st.secrets.get("INSTITUTIONAL_CODE", ""))  # sin clave por defecto: si no hay Secret, no hay acceso admin
NORMATIVAS = Path(__file__).parent / "normativas"

# Base de datos de códigos multi-uso e historial de IPs (Respaldo local + Supabase PostgreSQL permanente)
CODES_FILE = "codes_db.json"
CODES_LOG_FILE = "codes_usage_log.json"
CRM_FILE = "clientes_crm.csv"
QR_BIM_PATH = Path(__file__).parent / "qr_bim.png"

def _get_pg_conn():
    if not DATABASE_URL:
        return None
    try:
        import psycopg2
        return psycopg2.connect(DATABASE_URL, connect_timeout=5)
    except Exception as e:
        print("DB connect error:", type(e).__name__, str(e)[:160].replace(DATABASE_URL, "***"), flush=True)
        return None

@st.cache_resource
def _init_pg_tables():
    conn = _get_pg_conn()
    if not conn:
        return False
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute("""
                CREATE TABLE IF NOT EXISTS perspecta_codes (
                    code TEXT PRIMARY KEY,
                    initial_uses INTEGER NOT NULL,
                    remaining_uses INTEGER NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS perspecta_codes_log (
                    id SERIAL PRIMARY KEY,
                    fecha TEXT NOT NULL,
                    codigo TEXT NOT NULL,
                    accion TEXT NOT NULL,
                    ip TEXT,
                    ubicacion TEXT
                );
                CREATE TABLE IF NOT EXISTS perspecta_free_samples (
                    email TEXT PRIMARY KEY,
                    code TEXT NOT NULL,
                    month TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS perspecta_used_payments (
                    payment_id TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS perspecta_crm (
                    id SERIAL PRIMARY KEY,
                    fecha TEXT NOT NULL,
                    proyecto TEXT,
                    nivel TEXT,
                    etapa TEXT,
                    codigo TEXT,
                    ip TEXT,
                    ubicacion TEXT
                );
                """)
        conn.close()
        return True
    except Exception:
        try:
            conn.close()
        except Exception:
            pass
        return False

def _load_codes_db():
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            db = {ACCESS_CODE: 99999}
            with conn.cursor() as cur:
                cur.execute("SELECT code, initial_uses, remaining_uses, created_at FROM perspecta_codes ORDER BY created_at DESC;")
                for code, init_u, rem_u, created_at in cur.fetchall():
                    db[code] = {
                        "initial": int(init_u),
                        "remaining": int(rem_u),
                        "created_at": created_at or ""
                    }
            conn.close()
            return db
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

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
    if not code_input:
        return 0
    db = _load_codes_db()
    val = db.get(code_input, 0)
    if isinstance(val, dict):
        return val.get("remaining", 0)
    return int(val) if isinstance(val, (int, float)) else 0

def deduct_code_use(code_input):
    if not code_input or code_input == ACCESS_CODE:
        return
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "UPDATE perspecta_codes SET remaining_uses = GREATEST(0, remaining_uses - 1) WHERE code = %s;",
                        (code_input,)
                    )
            conn.close()
            return
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

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
    created_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        """
                        INSERT INTO perspecta_codes (code, initial_uses, remaining_uses, created_at)
                        VALUES (%s, %s, %s, %s)
                        ON CONFLICT (code) DO UPDATE SET
                            initial_uses = EXCLUDED.initial_uses,
                            remaining_uses = EXCLUDED.remaining_uses,
                            created_at = EXCLUDED.created_at;
                        """,
                        (new_code, int(uses), int(uses), created_str)
                    )
            conn.close()
            return
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    db = _load_codes_db()
    db[new_code] = {
        "initial": int(uses),
        "remaining": int(uses),
        "created_at": created_str
    }
    _save_codes_db(db)

def delete_code(code_to_delete):
    if not code_to_delete or code_to_delete == ACCESS_CODE:
        return
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute("DELETE FROM perspecta_codes WHERE code = %s;", (code_to_delete,))
            conn.close()
            return
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

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
    fecha_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "INSERT INTO perspecta_codes_log (fecha, codigo, accion, ip, ubicacion) VALUES (%s, %s, %s, %s, %s);",
                        (fecha_str, code_used, action, ip, loc)
                    )
            conn.close()
            return
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    logs = []
    if os.path.exists(CODES_LOG_FILE):
        try:
            with open(CODES_LOG_FILE, 'r', encoding='utf-8') as f:
                logs = json.load(f)
        except Exception:
            logs = []
    logs.append({
        "fecha": fecha_str,
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

def _load_usage_logs():
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT fecha, codigo, accion, ip, ubicacion FROM perspecta_codes_log ORDER BY id DESC LIMIT 500;")
                rows = cur.fetchall()
            conn.close()
            return [
                {"fecha": r[0], "codigo": r[1], "accion": r[2], "ip": r[3], "ubicacion": r[4]}
                for r in rows
            ]
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if os.path.exists(CODES_LOG_FILE):
        try:
            with open(CODES_LOG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []

def generate_pack_code(extra_uses):
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
    code = f"PACK-{suffix}"
    add_new_code(code, extra_uses)
    return code

# ---------- Muestra gratis automática (1 código de 1 uso por correo, con tope mensual) ----------
FREE_SAMPLES_FILE = "free_samples_db.json"
FREE_SAMPLES_CAP = int(st.secrets.get("FREE_SAMPLES_MONTHLY_CAP", os.getenv("FREE_SAMPLES_MONTHLY_CAP", "30")))
SMTP_HOST = st.secrets.get("SMTP_HOST", os.getenv("SMTP_HOST", "smtp.gmail.com"))
SMTP_PORT = int(st.secrets.get("SMTP_PORT", os.getenv("SMTP_PORT", "465")))
SMTP_USER = st.secrets.get("SMTP_USER", os.getenv("SMTP_USER", ""))
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", os.getenv("SMTP_PASSWORD", ""))
EMAIL_DELIVERY = bool(SMTP_USER and SMTP_PASSWORD)  # con correo real configurado, el código se envía solo al correo (no se muestra en pantalla)

_DISPOSABLE_DOMAINS = {
    "mailinator.com", "10minutemail.com", "guerrillamail.com", "tempmail.com", "temp-mail.org",
    "yopmail.com", "trashmail.com", "sharklasers.com", "getnada.com", "maildrop.cc", "dispostable.com",
}

def normalize_email(raw):
    """Devuelve el correo normalizado (minúsculas; en Gmail sin puntos ni +alias) o None si no es válido."""
    import re
    e = (raw or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)+", e) or len(e) > 120:
        return None
    local, domain = e.split("@", 1)
    if domain in _DISPOSABLE_DOMAINS:
        return None
    if domain in ("gmail.com", "googlemail.com"):
        local = local.split("+", 1)[0].replace(".", "")
        domain = "gmail.com"
    else:
        local = local.split("+", 1)[0]
    return f"{local}@{domain}"

def _free_samples_local():
    try:
        with open(FREE_SAMPLES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def request_free_sample(email_norm):
    """Crea un código de 1 uso para este correo. Devuelve (estado, codigo): estado in ok | ya_existe | tope | error."""
    month = datetime.datetime.now().strftime("%Y-%m")
    created = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    suffix = "".join(random.choices("ABCDEFGHJKLMNPQRSTUVWXYZ23456789", k=6))
    code = f"MUESTRA-{suffix}"
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute("LOCK TABLE perspecta_free_samples IN EXCLUSIVE MODE;")
                    cur.execute("SELECT 1 FROM perspecta_free_samples WHERE email = %s;", (email_norm,))
                    if cur.fetchone():
                        return "ya_existe", None
                    cur.execute("SELECT COUNT(*) FROM perspecta_free_samples WHERE month = %s;", (month,))
                    if cur.fetchone()[0] >= FREE_SAMPLES_CAP:
                        return "tope", None
                    cur.execute("INSERT INTO perspecta_free_samples (email, code, month, created_at) VALUES (%s, %s, %s, %s);",
                                (email_norm, code, month, created))
                    cur.execute("INSERT INTO perspecta_codes (code, initial_uses, remaining_uses, created_at) VALUES (%s, 1, 1, %s);",
                                (code, created))
            conn.close()
            return "ok", code
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
            return "error", None
    db = _free_samples_local()
    if email_norm in db:
        return "ya_existe", None
    if sum(1 for v in db.values() if v.get("month") == month) >= FREE_SAMPLES_CAP:
        return "tope", None
    db[email_norm] = {"code": code, "month": month, "created_at": created}
    try:
        with open(FREE_SAMPLES_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, ensure_ascii=False, indent=2)
    except Exception:
        return "error", None
    add_new_code(code, 1)
    return "ok", code

def revoke_free_sample(email_norm, code):
    """Deshace una muestra si el correo no pudo enviarse (para que pueda reintentar)."""
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute("DELETE FROM perspecta_free_samples WHERE email = %s;", (email_norm,))
            conn.close()
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
    else:
        db = _free_samples_local()
        db.pop(email_norm, None)
        try:
            with open(FREE_SAMPLES_FILE, "w", encoding="utf-8") as f:
                json.dump(db, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    delete_code(code)

def send_sample_email(to_email, code):
    import smtplib
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["Subject"] = "Tu código de muestra gratis - Perspecta Salud"
    msg["From"] = f"Perspecta Salud <{SMTP_USER}>"
    msg["To"] = to_email
    msg.set_content(
        f"Hola,\n\nTu código de muestra gratis (1 revisión) es: {code}\n\n"
        f"Ingresa a {APP_URL}?page=review, escríbelo en el recuadro 'Ingresa tu código de activación' y sube tu proyecto.\n\n"
        "Si necesitas ayuda, escríbenos por WhatsApp: +51 913 152 645.\n\nPerspecta Salud"
    )
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=15) as srv:
            srv.login(SMTP_USER, SMTP_PASSWORD)
            srv.send_message(msg)
        return True
    except Exception:
        return False

USED_PAYMENTS_FILE = "used_payments.json"

def claim_payment(payment_id):
    """True si este payment_id de Mercado Pago se reclama por primera vez (solo puede activarse una vez, en cualquier sesión)."""
    pid = str(payment_id).strip()
    if not pid:
        return False
    created = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute("INSERT INTO perspecta_used_payments (payment_id, created_at) VALUES (%s, %s) ON CONFLICT (payment_id) DO NOTHING;", (pid, created))
                    first = cur.rowcount == 1
            conn.close()
            return first
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
    try:
        with open(USED_PAYMENTS_FILE, "r", encoding="utf-8") as f:
            used = json.load(f)
    except Exception:
        used = {}
    if pid in used:
        return False
    used[pid] = created
    try:
        with open(USED_PAYMENTS_FILE, "w", encoding="utf-8") as f:
            json.dump(used, f)
    except Exception:
        pass
    return True

def registrar_operacion(proyecto, nivel, etapa, codigo_usado="Pago directo"):
    ip, loc = get_client_ip_and_location()
    date_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "INSERT INTO perspecta_crm (fecha, proyecto, nivel, etapa, codigo, ip, ubicacion) VALUES (%s, %s, %s, %s, %s, %s, %s);",
                        (date_str, proyecto, nivel, etapa, codigo_usado, ip, loc)
                    )
            conn.close()
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    try:
        file_exists = os.path.isfile(CRM_FILE)
        with open(CRM_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Fecha", "Proyecto", "Nivel", "Etapa", "Codigo", "IP", "Ubicacion"])
            writer.writerow([date_str, proyecto, nivel, etapa, codigo_usado, ip, loc])
    except Exception:
        pass

def _get_crm_csv_data():
    _init_pg_tables()
    conn = _get_pg_conn()
    if conn:
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT fecha, proyecto, nivel, etapa, codigo, ip, ubicacion FROM perspecta_crm ORDER BY id DESC;")
                rows = cur.fetchall()
            conn.close()
            out = io.StringIO()
            writer = csv.writer(out)
            writer.writerow(["Fecha", "Proyecto", "Nivel", "Etapa", "Codigo", "IP", "Ubicacion"])
            for r in rows:
                writer.writerow(r)
            return out.getvalue().encode("utf-8-sig")
        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    if os.path.exists(CRM_FILE):
        try:
            with open(CRM_FILE, "rb") as f:
                return f.read()
        except Exception:
            return None
    return None

def _set_cell_shading(cell, fill_hex):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def _add_donde_coloca_table(doc, donde_txt, coloca_txt):
    from docx.shared import Inches, Pt, RGBColor
    tbl = doc.add_table(rows=2, cols=2)
    tbl.style = 'Table Grid'
    tbl.autofit = False

    # Fila 1: DONDE DICE
    c00, c01 = tbl.rows[0].cells[0], tbl.rows[0].cells[1]
    c00.width = Inches(1.35)
    c01.width = Inches(5.15)
    _set_cell_shading(c00, "FFF4F4")
    p00 = c00.paragraphs[0]
    r00 = p00.add_run("DONDE DICE:")
    r00.bold = True
    r00.font.size = Pt(10)
    r00.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B)
    p01 = c01.paragraphs[0]
    r01 = p01.add_run(donde_txt.strip())
    r01.font.size = Pt(10)

    # Fila 2: COLOCA
    c10, c11 = tbl.rows[1].cells[0], tbl.rows[1].cells[1]
    c10.width = Inches(1.35)
    c11.width = Inches(5.15)
    _set_cell_shading(c10, "E8F9FA")
    p10 = c10.paragraphs[0]
    r10 = p10.add_run("COLOCA:")
    r10.bold = True
    r10.font.size = Pt(10)
    r10.font.color.rgb = RGBColor(0x07, 0x2E, 0x68)
    p11 = c11.paragraphs[0]
    r11 = p11.add_run(coloca_txt.strip())
    r11.font.size = Pt(10)

    doc.add_paragraph("")

def create_docx(report_text, project_name, author_name):
    import re
    from docx.shared import Pt, RGBColor
    doc = docx.Document()
    
    # Agregar Encabezado (Header)
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.text = "Perspecta Salud - Hoja de Ruta de Correcciones Imperativas"
    
    # Agregar Pie de Página (Footer)
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Contacto: perspectasalud@gmail.com | WhatsApp: +51 913 152 645 | Web: https://revisordetesis.streamlit.app/"
    
    # Títulos principales en la primera página
    doc.add_heading("Perspecta Salud", 0)
    doc.add_heading("Hoja de Ruta de Correcciones Imperativas", 1)
    doc.add_heading(f"Proyecto: {project_name}", 2)
    doc.add_heading(f"Autor(es): {author_name}", 3)
    doc.add_paragraph(f"Fecha: {datetime.datetime.now().strftime('%d/%m/%Y')}")

    lines = report_text.splitlines()
    mode = None  # 'donde' or 'coloca'
    donde_buf = []
    coloca_buf = []

    def flush_pair():
        nonlocal mode, donde_buf, coloca_buf
        if donde_buf or coloca_buf:
            _add_donde_coloca_table(
                doc,
                "\n".join(donde_buf).strip() or "—",
                "\n".join(coloca_buf).strip() or "—"
            )
        mode = None
        donde_buf = []
        coloca_buf = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Detectar encabezados de sección (ej. "1. PÁGINAS PRELIMINARES...", "2. CAPÍTULO I...")
        if re.match(r"^[1-9]\.\s+[A-ZÁÉÍÓÚÑ0-9\s:(),—–-]{5,}$", line):
            flush_pair()
            doc.add_heading(line, level=2)
            continue

        # Detectar título de corrección (ej. "CORRECCIÓN 01 — ...")
        if re.match(r"^CORRECCI[ÓO]N\s+\d+", line, re.IGNORECASE):
            flush_pair()
            p = doc.add_paragraph()
            if "[CRÍTICO]" in line.upper():
                idx = line.upper().find("[CRÍTICO]")
                main_part = line[:idx].strip()
                crit_part = " [CRÍTICO]"
                r1 = p.add_run(main_part)
                r1.bold = True
                r1.font.size = Pt(11)
                r1.font.color.rgb = RGBColor(0x07, 0x2E, 0x68)
                r2 = p.add_run(crit_part)
                r2.bold = True
                r2.font.size = Pt(11)
                r2.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
            else:
                r1 = p.add_run(line)
                r1.bold = True
                r1.font.size = Pt(11)
                r1.font.color.rgb = RGBColor(0x07, 0x2E, 0x68)
            continue

        # Detectar inicio de DONDE DICE:
        m_donde = re.match(r"^\|?\s*DONDE\s+DICE\s*:\s*\|?(.*)$", line, re.IGNORECASE)
        if m_donde:
            if mode == "coloca" and (donde_buf or coloca_buf):
                flush_pair()
            mode = "donde"
            rest = m_donde.group(1).strip().rstrip("|").strip()
            if rest:
                donde_buf.append(rest)
            continue

        # Detectar inicio de COLOCA:
        m_coloca = re.match(r"^\|?\s*COLOCA\s*:\s*\|?(.*)$", line, re.IGNORECASE)
        if m_coloca:
            mode = "coloca"
            rest = m_coloca.group(1).strip().rstrip("|").strip()
            if rest:
                coloca_buf.append(rest)
            continue

        # Si estamos dentro de DONDE DICE o COLOCA, acumular líneas adicionales
        if mode == "donde":
            donde_buf.append(line)
        elif mode == "coloca":
            coloca_buf.append(line)
        else:
            p = doc.add_paragraph(line)
            if line.lower().startswith(("instrucción general", "hoja de ruta")):
                for r in p.runs:
                    r.bold = True

    flush_pair()

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def render_report_html(report_text):
    import re
    import html
    lines = report_text.splitlines()
    out = ['<div style="font-family:Inter,sans-serif;line-height:1.55;color:#112d54;">']
    mode = None
    donde_buf = []
    coloca_buf = []

    def flush_html_pair():
        nonlocal mode, donde_buf, coloca_buf
        if donde_buf or coloca_buf:
            d_html = "<br>".join(html.escape(x) for x in donde_buf) or "—"
            c_html = "<br>".join(html.escape(x) for x in coloca_buf) or "—"
            out.append(
                f'<table style="width:100%;border-collapse:collapse;margin:8px 0 18px 0;font-size:14px;border:1px solid #cbd5e1;border-radius:8px;overflow:hidden;">'
                f'<tr>'
                f'<td style="width:135px;background:#fff4f4;color:#991b1b;font-weight:800;padding:10px 12px;border:1px solid #cbd5e1;vertical-align:top;">DONDE DICE:</td>'
                f'<td style="background:#ffffff;color:#1e293b;padding:10px 14px;border:1px solid #cbd5e1;vertical-align:top;">{d_html}</td>'
                f'</tr>'
                f'<tr>'
                f'<td style="width:135px;background:#e8f9fa;color:#072e68;font-weight:800;padding:10px 12px;border:1px solid #cbd5e1;vertical-align:top;">COLOCA:</td>'
                f'<td style="background:#f8fcff;color:#072e68;padding:10px 14px;border:1px solid #cbd5e1;vertical-align:top;">{c_html}</td>'
                f'</tr>'
                f'</table>'
            )
        mode = None
        donde_buf = []
        coloca_buf = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        if re.match(r"^[1-9]\.\s+[A-ZÁÉÍÓÚÑ0-9\s:(),—–-]{5,}$", line):
            flush_html_pair()
            out.append(f'<h3 style="color:#072e68;background:#f3f9ff;padding:10px 14px;border-left:4px solid #08b8be;border-radius:6px;margin:22px 0 12px 0;font-size:17px;">{html.escape(line)}</h3>')
            continue

        if re.match(r"^CORRECCI[ÓO]N\s+\d+", line, re.IGNORECASE):
            flush_html_pair()
            esc = html.escape(line)
            esc = re.sub(
                r"\[CR[ÍI]TICO\]",
                '<span style="background:#fee2e2;color:#b91c1c;padding:2px 8px;border-radius:6px;font-size:12px;font-weight:800;margin-left:6px;">[CRÍTICO]</span>',
                esc,
                flags=re.IGNORECASE
            )
            out.append(f'<div style="font-weight:800;color:#072e68;font-size:15px;margin:14px 0 4px 0;">{esc}</div>')
            continue

        m_donde = re.match(r"^\|?\s*DONDE\s+DICE\s*:\s*\|?(.*)$", line, re.IGNORECASE)
        if m_donde:
            if mode == "coloca" and (donde_buf or coloca_buf):
                flush_html_pair()
            mode = "donde"
            rest = m_donde.group(1).strip().rstrip("|").strip()
            if rest:
                donde_buf.append(rest)
            continue

        m_coloca = re.match(r"^\|?\s*COLOCA\s*:\s*\|?(.*)$", line, re.IGNORECASE)
        if m_coloca:
            mode = "coloca"
            rest = m_coloca.group(1).strip().rstrip("|").strip()
            if rest:
                coloca_buf.append(rest)
            continue

        if mode == "donde":
            donde_buf.append(line)
        elif mode == "coloca":
            coloca_buf.append(line)
        else:
            if line.lower().startswith("hoja de ruta"):
                out.append(f'<div style="background:#072e68;color:#ffffff;padding:14px 18px;border-radius:10px;font-weight:800;font-size:17px;margin-bottom:12px;">{html.escape(line)}</div>')
            elif line.lower().startswith("instrucción general"):
                out.append(f'<div style="background:#eef8ff;border:1px solid #08b8be;padding:14px 16px;border-radius:10px;margin-bottom:16px;font-size:14px;"><b>{html.escape(line)}</b></div>')
            else:
                out.append(f'<p style="margin:6px 0;font-size:14px;">{html.escape(line)}</p>')

    flush_html_pair()
    out.append("</div>")
    return "\n".join(out)

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{--navy:#072e68;--cyan:#08b8be;--ink:#112d54;--muted:#63738c;--pale:#f3f9ff;--line:#deebf4}
html{scroll-behavior:smooth}body,[class*=css]{font-family:Inter,sans-serif;color:var(--ink)}
header[data-testid=stHeader]{height:0;background:transparent}#MainMenu,footer,[data-testid=stSidebarCollapsedControl]{display:none}.block-container{max-width:1240px;padding:0 2rem 2rem}.stMainBlockContainer{padding-top:0!important}
h1,h2,h3{color:var(--navy);letter-spacing:-.03em}.anchor{scroll-margin-top:80px}.nav{height:74px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;background:#fffffff2;border-bottom:1px solid #edf3f7;margin:0 -2rem;padding:0 2rem}.brand{font-size:22px;font-weight:800;color:var(--navy)}.brand span{color:var(--cyan)}.navlinks{display:flex;align-items:center;gap:25px}.nav a{font-size:13px;font-weight:600;text-decoration:none;color:var(--ink)}.nav .cta,.button{background:var(--cyan);color:#fff;padding:13px 19px;border-radius:8px;box-shadow:0 9px 24px #08b8be2c}
.hero{display:grid;grid-template-columns:1fr 1.08fr;align-items:center;gap:38px;min-height:505px}.eyebrow{font-size:12px;color:var(--cyan);font-weight:800;letter-spacing:.11em;text-transform:uppercase}.hero h1{font-size:clamp(40px,4.8vw,60px);line-height:1.08;margin:12px 0 20px}.hero p,.head p{font-size:17px;line-height:1.65;color:var(--muted)}.actions{display:flex;align-items:center;gap:22px;margin:27px 0}.button{display:inline-block;text-decoration:none!important;font-weight:700}.textlink{color:var(--navy);text-decoration:none;font-weight:700}.trust{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;font-weight:600;color:#415570}.trust span:before{content:'✓';color:var(--cyan);margin-right:6px}.hero-media{height:430px;position:relative;overflow:hidden;border-radius:0 0 0 65px}.hero-media img{width:100%;height:100%;object-fit:cover}
.score{position:absolute;right:24px;top:32px;background:#fffffff8;padding:16px 20px;border-radius:16px;box-shadow:0 16px 42px #072e6830;border:1px solid #deebf4;text-align:left;max-width:235px}.score .tag{color:var(--cyan);font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;display:block}.score b{display:block;font-size:16px;color:var(--navy);margin:4px 0 2px;line-height:1.2;font-weight:800}.score small{color:var(--muted);font-size:12px;line-height:1.35;display:block}
.stats{display:grid;grid-template-columns:repeat(4,1fr);background:var(--pale);border-radius:17px;padding:24px;margin:10px 0 68px}.stat{text-align:center;border-right:1px solid var(--line)}.stat:last-child{border:0}.stat strong{font-size:23px;color:var(--navy)}.stat span{display:block;color:var(--muted);font-size:12px;margin-top:4px}.section{padding:68px 0}.head{text-align:center;max-width:720px;margin:0 auto 38px}.head h2{font-size:35px;margin:0 0 10px}.steps{display:grid;grid-template-columns:repeat(4,1fr);gap:24px}.step{text-align:center;padding:10px 18px}.number{width:58px;height:58px;border-radius:50%;background:#e8f9fa;display:grid;place-items:center;margin:auto;color:var(--cyan);font-size:25px;font-weight:800}.step h3{margin:15px 0 7px}.step p{font-size:13px;line-height:1.6;color:var(--muted)}
.soft{background:linear-gradient(#f6fbff,#fff);margin:0 -2rem;padding-left:2rem;padding-right:2rem}.areas{display:grid;grid-template-columns:repeat(5,1fr);gap:13px}.card{border:1px solid var(--line);border-radius:14px;padding:22px 18px;background:#fff;box-shadow:0 8px 25px #102b5210;display:flex;flex-direction:column;justify-content:space-between}.card h3{font-size:14px;min-height:36px;margin:0 0 8px}.card-badge{display:inline-block;background:#e8f9fa;color:var(--cyan);font-size:11px;font-weight:700;padding:3px 8px;border-radius:6px;width:fit-content;margin-bottom:8px}.card p{font-size:12px;line-height:1.55;color:var(--muted);margin:0}.human{display:grid;grid-template-columns:1fr 1fr;background:#eef8ff;border-radius:20px;overflow:hidden;margin-top:52px}.human img{width:100%;height:315px;object-fit:cover}.quote{padding:45px;display:flex;flex-direction:column;justify-content:center}.quote b{font-size:58px;color:var(--cyan);line-height:.5}.quote p{font-size:20px;line-height:1.6}.quote small{color:var(--muted)}
.report{display:grid;grid-template-columns:1.15fr .85fr;gap:25px;align-items:center}.report-card{border:1px solid var(--line);border-radius:17px;padding:24px;box-shadow:0 16px 40px #102b5214;background:#fff}.report-card-head{font-size:12px;font-weight:800;color:var(--cyan);letter-spacing:.08em;text-transform:uppercase;margin-bottom:14px;border-bottom:1px solid var(--line);padding-bottom:8px}
.preview-box{background:#f8fcff;border:1px solid var(--line);border-radius:11px;padding:12px 14px;margin-bottom:10px}.preview-tag{font-size:11px;font-weight:800;color:var(--cyan);text-transform:uppercase;letter-spacing:.04em;display:block}.preview-text{font-size:13px;color:var(--ink);margin-top:2px;line-height:1.4}.preview-quote{font-size:12px;color:#072e68;background:#fff;border-left:3px solid var(--cyan);padding:7px 10px;margin-top:6px;border-radius:0 6px 6px 0;font-style:italic}
.benefits{display:grid;gap:13px}.benefits div{background:var(--pale);padding:17px;border-radius:11px}.benefits b{display:block;margin-bottom:4px}.benefits span{font-size:13px;color:var(--muted)}
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
            if amt >= PRECIO_ADV25 - 0.5:
                return 25
            elif amt >= PRECIO_ADV10 - 0.5:
                return 10
            elif amt >= PRECIO_5 - 0.5:
                return 5
            elif amt >= PRECIO_3 - 0.5:
                return 3
            elif amt >= PRECIO_1 - 0.5:
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
    <main><section id="inicio" class="hero anchor"><div><div class="eyebrow">Auditoría académica inteligente</div><h1>Audita tu proyecto de tesis antes de entregarlo a tu asesor o jurado</h1><p>Recibe en minutos una auditoría académica integral con observaciones precisas, tratamiento estadístico idóneo en SPSS y párrafos sugeridos listos para insertar en tu trabajo.</p><div class="actions"><a class="button" href="?page=review">Revisar mi proyecto →</a><a class="textlink" href="#informe" onclick="document.getElementById('informe').scrollIntoView({{behavior: 'smooth'}}); return false;">Ver estructura del informe ›</a></div><div class="trust"><span>Coherencia científica</span><span>Tratamiento estadístico (SPSS)</span><span>Párrafos listos para insertar</span><span>Word confidencial</span></div></div><div class="hero-media"><img src="https://images.unsplash.com/photo-1541339907198-e08756dedf3f?auto=format&fit=crop&w=1200&q=86" alt="Estudiante trabajando en su investigación"><div class="score"><span class="tag">Auditoría Integral</span><b>Sin calificaciones</b><small>Orientación 100% constructiva con textos sugeridos</small></div></div></section>
    <section class="stats"><div class="stat"><strong>5</strong><span>áreas auditadas</span></div><div class="stat"><strong>SPSS</strong><span>guía paso a paso</span></div><div class="stat"><strong>100%</strong><span>Word editable (.docx)</span></div><div class="stat"><strong>Desde S/ {PRECIO_1:.2f}</strong><span>Planes accesibles</span></div></section>
    <section id="como" class="section anchor"><div class="head"><h2>¿Cómo funciona?</h2><p>Cuatro pasos sencillos para fortalecer tu investigación antes de presentarla.</p></div><div class="steps"><article class="step"><div class="number">1</div><h3>Selecciona</h3><p>Elige el plan individual o paquete de seguimiento para tu ciclo.</p></article><article class="step"><div class="number">2</div><h3>Activa</h3><p>Paga al instante con Mercado Pago o mediante QR directo.</p></article><article class="step"><div class="number">3</div><h3>Carga</h3><p>Sube tu documento indicando tu nivel (Pregrado, Maestría o Doctorado).</p></article><article class="step"><div class="number">4</div><h3>Descarga</h3><p>Recibe tu informe en Word (.docx) con párrafos modelo y guía en SPSS.</p></article></div></section>
    <section id="evaluamos" class="section soft anchor"><div class="head"><h2>Descubre qué auditamos en tu trabajo</h2><p>Criterios metodológicos y científicos rigurosos adaptados a tu nivel académico.</p></div><div class="areas"><article class="card"><span class="card-badge">Estructura & Lógica</span><h3>⚛ Coherencia científica</h3><p>Alineación estricta entre problema, objetivos, dimensiones y teoría, con realidad problemática en embudo (Mundial a Local).</p></article><article class="card"><span class="card-badge">Análisis Cuantitativo</span><h3>▣ Metodología y SPSS</h3><p>Análisis de tus variables y prescripción de la prueba estadística idónea explicada en lenguaje simple y con guía paso a paso en SPSS.</p></article><article class="card"><span class="card-badge">Estado del Arte</span><h3>▤ Antecedentes y literatura</h3><p>Verificación del umbral mínimo de antecedentes nacionales e internacionales ordenados cronológicamente y articulados a tus variables.</p></article><article class="card"><span class="card-badge">Normativa Académica</span><h3>❞ Estilo de citación</h3><p>Alineación estricta a normas Vancouver o APA y correspondencia bibliográfica con fuentes científicas indexadas.</p></article><article class="card"><span class="card-badge">Texto Listo</span><h3>♢ Aspectos éticos</h3><p>Redacción directa de principios bioéticos y consentimiento informado listos para insertar en tu documento sin trabas burocráticas.</p></article></div><div class="human"><img src="https://images.unsplash.com/photo-1523240795612-9a054b0db644?auto=format&fit=crop&w=1200&q=86" alt="Estudiantes revisando un proyecto"><div class="quote"><b>“</b><p>No te dejamos solo con la crítica: cada observación incluye la explicación clara y el párrafo modelo redactado para que lo copies y enriquezcas tu tesis.</p><small>Una auditoría diagnóstica formativa diseñada para tesistas</small></div></div></section>
    <section id="informe" class="section anchor"><div class="head"><h2>Hoja de Ruta en Word (.docx) con tablas DONDE DICE / COLOCA</h2><p>Recorre tu tesis sección por sección señalando con exactitud qué fragmento o tabla corregir y entregándote el texto listo para copiar y pegar.</p></div><div class="report"><div class="report-card"><div class="report-card-head">Vista previa de tu Hoja de Ruta de Correcciones Imperativas</div><div class="preview-box"><span class="preview-tag">Instrucción General y Lo Logrado</span><div class="preview-text">Reconoce lo que ya está bien planteado e indica qué aspectos NO debes modificar (ej. validando tu prueba estadística si ya es correcta).</div></div><div class="preview-box"><span class="preview-tag">CORRECCIÓN 03 — Resumen y Estadísticos [CRÍTICO]</span><div class="preview-text"><b>DONDE DICE:</b> Porcentajes con punto decimal (68.3 %) y '(Rho = 0.448; p = 0.000)'.</div><div class="preview-quote"><b>COLOCA:</b> Usa coma decimal (68,3 %) y reemplaza 'p = 0.000' por '(ρ = 0,448; p &lt; 0,001)' (en bioestadística nunca se reporta p = 0,000).</div></div><div class="preview-box"><span class="preview-tag">CORRECCIÓN 06 — 1.3 Bases Teóricas y Articulación [CRÍTICO]</span><div class="preview-text"><b>DONDE DICE:</b> '...para diagnosticar la hiperemia gravídica (24).' y ausencia de teoría conductual.</div><div class="preview-quote"><b>COLOCA:</b> Cambia 'hiperemia gravídica' por 'hiperémesis gravídica' e inserta al final el párrafo modelo redactado del Modelo de Creencias en Salud con su fuente verificada y DOI...</div></div></div><div class="benefits"><div><b>✓ Formato quirúrgico DONDE DICE / COLOCA</b><span>Localiza el error en segundos en tu Word y reemplázalo con el texto exacto ya redactado.</span></div><div><b>✓ Lectura completa de Tablas y Anexos</b><span>Audita tus Tablas de Resultados, Matriz de Consistencia, Operacionalización y cuestionarios Likert.</span></div><div><b>✓ Tratamiento estadístico en SPSS</b><span>Valida tu prueba si es correcta o te prescribe la prueba exacta con guía paso a paso en SPSS.</span></div><div><b>🔒 Tablas ejecutivas en Word (.docx)</b><span>Documento formal, privado y 100% editable con tablas comparativas listas para aplicar.</span></div></div></div></section>
    <section id="empezar" class="section anchor"><div class="head"><h2>Planes y Tarifario</h2><p>Elige una revisión individual o ahorra con nuestros paquetes de seguimiento para todo tu ciclo.</p></div><div class="price-grid"><div class="price-card"><div><h3>1 Evaluación</h3><p style="color:var(--muted);font-size:14px">Ideal para una auditoría puntual antes de entregar.</p><div class="amount">S/ {PRECIO_1:.2f}</div><small style="color:var(--muted)">Pago único</small><div class="checks"><div>1 auditoría completa de 5 áreas</div><div>Párrafos modelo listos para insertar</div><div>Tratamiento estadístico en SPSS</div><div>Informe descargable en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=1">Elegir 1 Evaluación →</a></div><div class="price-card featured"><div class="badge-pop">Más elegido</div><div><h3>3 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Perfecto para corregir y validar tus avances sucesivos.</p><div class="amount">S/ {PRECIO_3:.2f}</div><small style="color:var(--cyan);font-weight:700">S/ {PRECIO_3/3:.2f} por evaluación</small><div class="checks"><div>1 evaluación inmediata + código por 2 usos</div><div>Úsalas cuando quieras en tu ciclo</div><div>Párrafos modelo y guía en SPSS</div><div>Informes descargables en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=3">Elegir Pack de 3 →</a></div><div class="price-card"><div><h3>5 Evaluaciones</h3><p style="color:var(--muted);font-size:14px">Para acompañar toda tu tesis o compartir con tu grupo.</p><div class="amount">S/ {PRECIO_5:.2f}</div><small style="color:var(--cyan);font-weight:700">¡Solo S/ {PRECIO_5/5:.2f} por evaluación!</small><div class="checks"><div>1 evaluación inmediata + código por 4 usos</div><div>Máximo ahorro (Ahorras S/ {5*PRECIO_1 - PRECIO_5:.2f})</div><div>Párrafos modelo y guía en SPSS</div><div>Informes descargables en Word (.docx)</div></div></div><a class="button" style="text-align:center" href="?page=review&pack=5">Elegir Pack de 5 →</a></div></div></section>
    <section class="section"><div class="head"><h2>Confianza, ética y confidencialidad</h2><p>Un servicio transparente diseñado para apoyar tu formación académica.</p></div><div class="security"><article><h3>🛡 Archivos protegidos</h3><p>Se procesan de forma privada y segura exclusivamente para generar tu informe de auditoría.</p></article><article><h3>🧠 Auditoría diagnóstica experta</h3><p>Identificamos debilidades para que las corrijas a tiempo. No reemplazamos a tu asesor ni redactamos tesis completas desde cero.</p></article><article><h3>⏱ Diagnóstico en minutos</h3><p>Sin semanas de espera para saber con certeza qué aspectos debes mejorar antes de entregar.</p></article></div></section>
    <section id="faq" class="section anchor"><div class="head"><h2>Preguntas frecuentes</h2></div><div class="faq"><details><summary>¿Perspecta Salud aprueba mi tesis o reemplaza a mi asesor?</summary><p>No. Perspecta Salud es una herramienta de auditoría diagnóstica previa. No redacta tesis completas desde cero ni sustituye la dirección académica de tu asesor, sino que identifica observaciones metodológicas y de coherencia a tiempo para que presentes un trabajo impecable.</p></details><details><summary>¿El informe me otorga una nota o calificación?</summary><p>No asignamos calificaciones numéricas. Nuestro enfoque es 100% formativo y constructivo: te explicamos qué corregir y te proporcionamos párrafos modelo listos para insertar en tu trabajo.</p></details><details><summary>¿Qué tipo de trabajos académicos pueden auditar?</summary><p>Proyectos de investigación, tesis en desarrollo y tesis finales en niveles de Pregrado, Maestría y Doctorado.</p></details><details><summary>¿Cómo funcionan los paquetes de 3 y 5 evaluaciones?</summary><p>Al pagar realizas tu primera revisión hoy y el sistema te genera automáticamente en pantalla un código secreto (ej. PACK-XXXXX) para usar tus evaluaciones restantes cuando tengas listos tus próximos avances.</p></details></div></section>
    <section class="final"><div><h2>¿Listo para auditar tu investigación?</h2><p>Identifica debilidades a tiempo y lleva tu trabajo académico al siguiente nivel.</p></div><a class="button" href="?page=review">Revisar mi proyecto →</a></section>
    """, unsafe_allow_html=True)
    
    # Bloque para Iniciar / FlagCounter / Formulario
    _, center, _ = st.columns([1, 1.25, 1])
    with center:
        if st.button(f"Revisar mi proyecto desde S/ {PRECIO_1:.2f}", type="primary", use_container_width=True):
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
                    <div>💬 <b>WhatsApp:</b> <a href="https://wa.me/51913152645" target="_blank" style="color: #08b8be; text-decoration: none; font-weight: 700;">+51 913 152 645</a></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<div style='text-align: center; margin-top: 30px;'><p>Estadísticas de la plataforma:</p></div>", unsafe_allow_html=True)
        html_code = """<div style="text-align: center;"><a href="https://info.flagcounter.com/OA6D"><img src="https://s01.flagcounter.com/countxl/OA6D/bg_FFFFFF/txt_000000/border_CCCCCC/columns_2/maxflags_4/viewers_Usuarios/labels_0/pageviews_1/flags_0/percent_0/" alt="Flag Counter" border="0"></a></div>"""
        components.html(html_code, height=180)

    st.markdown(f'<footer class="footer"><div class="brand" style="display:flex;align-items:center;">{LOGO_SVG}Perspecta <span>Salud</span></div><div>Correo: perspectasalud@gmail.com · WhatsApp: +51 913 152 645</div><div>© 2026 Perspecta Salud</div></footer></main><div id="activar" class="anchor"></div>', unsafe_allow_html=True)


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
            if unlock_clicked and _time.time() < st.session_state.get("code_lock_until", 0):
                st.error("⏳ Demasiados intentos fallidos. Espera 60 segundos e inténtalo de nuevo.")
                unlock_clicked = False
            if unlock_clicked:
                clean_code = inst_code.strip()
                uses_left = get_code_uses(clean_code)
                if uses_left > 0:
                    st.session_state.paid = True
                    st.session_state.used_code = clean_code
                    if ACCESS_CODE and clean_code == ACCESS_CODE:
                        st.session_state.is_admin = True
                    else:
                        log_code_usage(clean_code, action="Ingreso de código (Desbloqueo)")
                    st.rerun()
                else:
                    st.session_state.code_fails = st.session_state.get("code_fails", 0) + 1
                    if st.session_state.code_fails >= 5:
                        st.session_state.code_lock_until = _time.time() + 60
                        st.session_state.code_fails = 0
                    _time.sleep(1)
                    st.error("❌ Código inválido o sin usos disponibles. Verifica que esté bien escrito.")

            st.markdown("---")

            # 1b. MUESTRA GRATIS AUTOMÁTICA (1 código por correo)
            st.subheader("🎁 ¿Quieres probarlo primero? Pide tu muestra gratis")
            if EMAIL_DELIVERY:
                st.caption("Escribe tu correo y te enviamos un código de 1 revisión gratis. Un código por persona.")
            else:
                st.caption("Escribe tu correo y obtén un código de 1 revisión gratis. Un código por correo.")
            col_mail, col_mail_btn = st.columns([3, 1])
            with col_mail:
                sample_email = st.text_input("Tu correo", placeholder="tucorreo@ejemplo.com", label_visibility="collapsed", key="sample_email")
            with col_mail_btn:
                sample_clicked = st.button("🎁 Quiero mi muestra gratis", use_container_width=True)
            if sample_clicked:
                email_norm = normalize_email(sample_email)
                if _time.time() < st.session_state.get("sample_lock_until", 0):
                    st.error("⏳ Espera un momento antes de volver a intentar.")
                elif not email_norm:
                    st.error("Escribe un correo válido (no se aceptan correos temporales).")
                else:
                    st.session_state.sample_lock_until = _time.time() + 20
                    status, sample_code = request_free_sample(email_norm)
                    if status == "ok":
                        log_code_usage(sample_code, action="Muestra gratis (código emitido)")
                        if EMAIL_DELIVERY:
                            if send_sample_email(email_norm, sample_code):
                                st.success("✅ ¡Listo! Te enviamos tu código al correo. Revisa tu bandeja de entrada (y spam), escríbelo arriba en 'Ingresa tu código de activación'.")
                            else:
                                revoke_free_sample(email_norm, sample_code)
                                st.error("No pudimos enviar el correo ahora. Inténtalo de nuevo en unos minutos o escríbenos por WhatsApp al +51 913 152 645.")
                        else:
                            st.success(f"✅ Tu código de muestra gratis es: **`{sample_code}`** (1 revisión). Cópialo y escríbelo arriba en 'Ingresa tu código de activación'.")
                    elif status == "ya_existe":
                        st.warning("Este correo ya recibió su muestra gratis. Si la perdiste o necesitas otra revisión, escríbenos por WhatsApp al +51 913 152 645.")
                    elif status == "tope":
                        st.info("Las muestras gratis de este mes ya se agotaron. Escríbenos por WhatsApp al +51 913 152 645 o elige un plan abajo.")
                    else:
                        st.error("No pudimos generar tu muestra ahora. Inténtalo de nuevo en unos minutos.")

            st.markdown("---")

            # 2. LOS 3 BOTONES DIRECTOS DE MERCADO PAGO (VISIBLES AL MISMO TIEMPO)
            st.subheader("💳 ¿Aún no tienes un código? Elige tu plan y paga aquí")
            st.write("Selecciona directamente el paquete que prefieras. **Con Mercado Pago el acceso se activa de forma 100% automática al instante:**")

            if MP_TOKEN:
                p_col1, p_col2, p_col3 = st.columns(3)
                with p_col1:
                    with st.container(border=True):
                        st.markdown("#### 1 Evaluación")
                        st.markdown(f"<div style='font-size:32px;font-weight:800;color:#072e68;margin:4px 0;'>S/ {PRECIO_1:.2f}</div>", unsafe_allow_html=True)
                        st.caption("1 revisión inmediata con informe descargable en Word.")
                        try:
                            st.link_button(
                                f"⚡ Pagar S/ {PRECIO_1:.2f} (Mercado Pago)",
                                payment_link(PRECIO_1, "1 Evaluación - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error(f"Error al generar link S/ {PRECIO_1:.2f}")

                with p_col2:
                    with st.container(border=True):
                        st.markdown("#### 🌟 3 Evaluaciones")
                        st.markdown(f"<div style='font-size:32px;font-weight:800;color:#08b8be;margin:4px 0;'>S/ {PRECIO_3:.2f}</div>", unsafe_allow_html=True)
                        st.caption(f"1 revisión hoy + Código secreto con 2 usos adicionales (Ahorras S/ {3*PRECIO_1 - PRECIO_3:.2f}).")
                        try:
                            st.link_button(
                                f"⚡ Pagar S/ {PRECIO_3:.2f} (Mercado Pago)",
                                payment_link(PRECIO_3, "Paquete 3 Evaluaciones - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error(f"Error al generar link S/ {PRECIO_3:.2f}")

                with p_col3:
                    with st.container(border=True):
                        st.markdown("#### 🔥 5 Evaluaciones")
                        st.markdown(f"<div style='font-size:32px;font-weight:800;color:#072e68;margin:4px 0;'>S/ {PRECIO_5:.2f}</div>", unsafe_allow_html=True)
                        st.caption(f"1 revisión hoy + Código secreto con 4 usos adicionales (¡Ahorras S/ {5*PRECIO_1 - PRECIO_5:.2f}!).")
                        try:
                            st.link_button(
                                f"⚡ Pagar S/ {PRECIO_5:.2f} (Mercado Pago)",
                                payment_link(PRECIO_5, "Paquete 5 Evaluaciones - Perspecta Salud"),
                                type="primary",
                                use_container_width=True
                            )
                        except Exception:
                            st.error(f"Error al generar link S/ {PRECIO_5:.2f}")
                st.markdown("#### 🎓 Para asesores y docentes")
                st.caption("Paquetes grandes para revisar los proyectos de tus tesistas. Pago automático con Mercado Pago, tu código llega en pantalla.")
                a_col1, a_col2 = st.columns(2)
                for a_col, a_evals, a_price in ((a_col1, 10, PRECIO_ADV10), (a_col2, 25, PRECIO_ADV25)):
                    with a_col:
                        with st.container(border=True):
                            st.markdown(f"#### 🎓 Pack Asesor {a_evals}")
                            st.markdown(f"<div style='font-size:32px;font-weight:800;color:#072e68;margin:4px 0;'>S/ {a_price:.2f}</div>", unsafe_allow_html=True)
                            st.caption(f"1 revisión hoy + Código secreto con {a_evals - 1} usos adicionales (S/ {a_price/a_evals:.2f} por evaluación).")
                            try:
                                st.link_button(
                                    f"⚡ Pagar S/ {a_price:.2f} (Mercado Pago)",
                                    payment_link(a_price, f"Pack Asesor {a_evals} Evaluaciones - Perspecta Salud"),
                                    type="primary",
                                    use_container_width=True
                                )
                            except Exception:
                                st.error(f"Error al generar link S/ {a_price:.2f}")
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
                    f"Si prefieres pagar escaneando nuestro código QR directo (S/ {PRECIO_1:.2f}, S/ {PRECIO_3:.2f} o S/ {PRECIO_5:.2f}), realiza el pago y **envíanos la captura junto con tu Nombre y DNI a nuestro WhatsApp [+51 913 152 645](https://wa.me/51913152645)** para entregarte tu código de activación."
                )
                st.info(
                    "💡 **Nota amable e importante:** El envío de la captura con tu Nombre y DNI al WhatsApp **+51 913 152 645** es **únicamente para pagos realizados mediante este código QR**. "
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
            if _init_pg_tables():
                st.caption("🟢 **Base de datos permanente conectada (Supabase PostgreSQL)** — Tus códigos, logs de IP y CRM son 100% inmunes a reinicios.")
            else:
                st.warning("⚠️ **Aviso:** Agrega `DATABASE_URL` en los *Secrets* de Streamlit Cloud para activar la base de datos permanente.")
            
            # Gestión de Códigos
            with st.expander("🔑 Generar y Gestionar Códigos de Acceso (Tokens)", expanded=True):
                st.write("Crea contraseñas de uso limitado. Los códigos creados permanecerán visibles en la lista (incluso si llegan a 0 usos) hasta que decidas eliminarlos con su botón.")
                col1, col2 = st.columns(2)
                new_code_name = col1.text_input("Nuevo código", placeholder="Ej. GANATIEMPO")
                new_code_uses = col2.number_input("Número de usos", min_value=1, max_value=500, value=1)
                if st.button("Crear código", type="primary"):
                    if ACCESS_CODE and new_code_name.strip() == ACCESS_CODE:
                        st.error("Ese código es igual a la clave maestra. Elige otro.")
                    elif new_code_name.strip():
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
                            status_badge = "🟢 **ACTIVO**" if rem_u > 0 else "🔴 **AGOTADO**"
                            info_txt = f"{status_badge} | 🔑 **`{k}`** — Usados: **{used_u}/{init_u}** | Restantes: **{rem_u}**" + (f" *(Creado: {created_str})*" if created_str else "")
                        else:
                            rem_u = int(v)
                            status_badge = "🟢 **ACTIVO**" if rem_u > 0 else "🔴 **AGOTADO**"
                            info_txt = f"{status_badge} | 🔑 **`{k}`** — Usos restantes: **{rem_u}**"
                        c_info, c_del = st.columns([4, 1])
                        c_info.markdown(info_txt)
                        if c_del.button("🗑️ Eliminar", key=f"del_code_{k}"):
                            delete_code(k)
                            st.rerun()
                else:
                    st.caption("No hay códigos creados todavía.")

            # Historial de IPs y Ubicaciones por Código
            with st.expander("🌍 Historial de Uso de Códigos (IP y Ubicación Relativa)"):
                usage_logs = _load_usage_logs()
                if usage_logs:
                    st.dataframe(usage_logs, use_container_width=True)
                else:
                    st.caption("Aún no se han registrado accesos con códigos.")
            
            # Base de datos CRM
            crm_bytes = _get_crm_csv_data()
            if crm_bytes:
                st.download_button(
                    "📥 Descargar Base de Datos CRM (incluye Código, IP y Ubicación)",
                    data=crm_bytes,
                    file_name="clientes_crm.csv",
                    mime="text/csv",
                    use_container_width=True
                )
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
            st.subheader("📋 Hoja de Ruta de Correcciones Imperativas")
            st.markdown(render_report_html(st.session_state.report), unsafe_allow_html=True)
            
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

                with st.status(f"Generando auditoría científica ({nivel})...", expanded=True) as status:
                    result = evaluate_project(files, str(NORMATIVAS), API_KEY, academic_level=nivel, work_stage=etapa)
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
        st.session_state.page = "review"
        if claim_payment(payment_id):
            st.session_state.paid = True
            if purchased_evals > 1 and not st.session_state.get("generated_pack_code"):
                extra_uses = purchased_evals - 1
                st.session_state.generated_pack_code = generate_pack_code(extra_uses)
                st.session_state.generated_pack_uses = extra_uses
        else:
            st.error("Este pago ya fue utilizado para activar una revisión. Si necesitas ayuda, escríbenos por WhatsApp al +51 913 152 645.")
        st.query_params.clear()

st.session_state.setdefault("page", "home")
review() if st.session_state.page == "review" else landing()


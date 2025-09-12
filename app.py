# app.py  — versión completa

import os, sqlite3, uuid, csv, unicodedata
from datetime import datetime
from functools import wraps
from io import BytesIO, StringIO

from flask import (
    Flask, render_template, request, redirect, url_for,
    abort, flash, session, send_file, has_request_context
)

import qrcode
from werkzeug.security import generate_password_hash, check_password_hash

# ---------- ReportLab (PDF) ----------
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# =============== FIX BASE_DIR ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------- Excel (openpyxl) ----------
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import load_workbook
from openpyxl import Workbook, load_workbook
app = Flask(__name__)
app.secret_key = "dev-secret"  # cambia en producción
from flask import Flask, render_template
from flask import redirect, url_for

@app.route("/")
def index():
    # Redirige automáticamente a validar
    return redirect(url_for("validar"))

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("validar.html")  # o la página inicial que quieras mostrar

# Si sirves por LAN, usa tu IP local aquí para que el QR funcione
FALLBACK_BASE_URL = "http://192.168.1.41:5000"

# Nota mínima para aprobar
PASSING_GRADE = 60.0



# =========================================
# Context processors
# =========================================
from datetime import datetime

@app.context_processor
def inject_user():
    return {
        "current_user": session.get("user"),
        "PASSING_GRADE": PASSING_GRADE,
        "ETAPAS": ETAPAS,
        "YEAR": datetime.now().year,   # <-- año dinámico para footer
    }

# --- Rutas útiles para plantilla/fuentes (ajusta si cambias) ---
STATIC_DIR = os.path.join(BASE_DIR, "static")
CERT_STATIC = os.path.join(STATIC_DIR, "certs")
FONTS_DIR = os.path.join(CERT_STATIC, "fonts")
TEMPLATE_PNG = os.path.join(CERT_STATIC, "plantilla_certificado.png")

# --- Registrar fuentes si existen (usa Helvetica como fallback) ---
def _register_cert_fonts() -> None:
    """
    Registra fuentes buscándolas en ./static/certs/fonts y (si es Windows) en C:\Windows\Fonts.
    Soporta nombres de archivo tal como los tienes:
      - GreatVibes-Regular.ttf       -> "GreatVibes"
      - PlayfairDisplay-Italic-VariableFont_wght.ttf o PlayfairDisplay-Bold.ttf -> "Playfair-Bold"
      - ARIAL.TTF / Arial.ttf        -> "Arial"
      - ARIALBD 1.TTF / ArialBD.ttf  -> "Arial-Bold"
    """
    import os
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    def _try_register(name_alias: str, file_names: list, base_dir: str):
        if not base_dir or not os.path.isdir(base_dir):
            return
        lower = {f.lower(): f for f in os.listdir(base_dir)}
        for fn in file_names:
            real = lower.get(fn.lower())
            if real:
                try:
                    pdfmetrics.registerFont(TTFont(name_alias, os.path.join(base_dir, real)))
                    return
                except Exception:
                    pass

    # 1) ./static/certs/fonts
    _try_register("GreatVibes", ["GreatVibes-Regular.ttf"], FONTS_DIR)
    _try_register("Playfair-Bold", ["PlayfairDisplay-Italic-VariableFont_wght.ttf","PlayfairDisplay-Bold.ttf"], FONTS_DIR)
    _try_register("Arial", ["ARIAL.TTF","Arial.ttf"], FONTS_DIR)
    _try_register("Arial-Bold", ["ARIALBD 1.TTF","ARIALBD.TTF","Arial Bold.ttf","arialbd.ttf"], FONTS_DIR)

    # 2) Windows fallback
    try:
        win = r"C:\Windows\Fonts"
        _try_register("Arial", ["arial.ttf","ARIAL.TTF"], win)
        _try_register("Arial-Bold", ["arialbd.ttf","ARIALBD.TTF"], win)
    except Exception:
        pass

def _has_column(colname: str) -> bool:
    with conn() as c:
        cols = [r[1] for r in c.execute("PRAGMA table_info(estudiantes)").fetchall()]
    return colname in cols

def _get_notas_value(row: sqlite3.Row) -> str:
    # Soporta DB vieja (columna) y la nueva (notas)
    return (row["notas"] if "notas" in row.keys() and row["notas"] else
            (row["columna"] if "columna" in row.keys() else ""))

def _pick_font(name: str, fallback_bold: bool = False) -> str:
    """
    Devuelve el nombre de fuente registrada; si no, Arial/Arial-Bold/Helvetica.
    """
    from reportlab.pdfbase import pdfmetrics
    regs = set(pdfmetrics.getRegisteredFontNames())
    if name in regs:
        return name
    if fallback_bold and "Arial-Bold" in regs:
        return "Arial-Bold"
    if "Arial" in regs:
        return "Arial"
    return "Helvetica"

def _ensure_paths():
    # Garantiza rutas usadas en PDF
    os.makedirs(QR_DIR, exist_ok=True)
    os.makedirs(CERT_STATIC, exist_ok=True)
    os.makedirs(FONTS_DIR, exist_ok=True)
def _pick_font(name: str, fallback_bold: bool = False) -> str:
    """Devuelve el nombre de fuente disponible o Helvetica como reserva."""
    available = set(pdfmetrics.getRegisteredFontNames())
    if name in available:
        return name
    if fallback_bold and "Arial-Bold" in available:
        return "Arial-Bold"
    if "Arial" in available:
        return "Arial"
    return "Helvetica"

# Español:  '2 de Septiembre de 2025'
MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]

def _fecha_es(dt: datetime) -> str:
    try:
        d = dt.day
        m = MESES_ES[dt.month-1].capitalize()
        y = dt.year
        return f"{d} de {m} de {y}"
    except Exception:
        return dt.strftime("%Y-%m-%d")

# =========================================
# Configuración / Paths
# =========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "db.sqlite3")
QR_DIR   = os.path.join(BASE_DIR, "static", "qrs")
os.makedirs(QR_DIR, exist_ok=True)

# Donde intentaremos encontrar fuentes (pon aquí tus .ttf)
CERT_FONTS_DIR = os.path.join(BASE_DIR, "certs")
os.makedirs(CERT_FONTS_DIR, exist_ok=True)

# Marca de agua / plantilla (intenta varios lugares)
PLANTILLA_IMG = (
    os.path.join(BASE_DIR, "plantilla_certificado.png")
    if os.path.exists(os.path.join(BASE_DIR, "plantilla_certificado.png"))
    else os.path.join(BASE_DIR, "static", "img", "plantilla_certificado.png")
)

# Logo grande y tenue (opcional)
LOGO_WATERMARK = (
    os.path.join(BASE_DIR, "logo_marca_agua.png")
    if os.path.exists(os.path.join(BASE_DIR, "logo_marca_agua.png"))
    else os.path.join(BASE_DIR, "static", "img", "logo_marca_agua.png")
)

app = Flask(__name__)
app.secret_key = "dev-secret"  # cambia en producción

# Si sirves por LAN, usa tu IP local aquí para que el QR funcione
FALLBACK_BASE_URL = "http://192.168.1.41:5000"

# Nota mínima para aprobar
PASSING_GRADE = 60.0


# =========================================
# Catálogo de Etapas / Niveles (se muestra; se guarda como una cadena en 'curso')
# =========================================
ETAPAS = {
    "Principiante": ["PRE A1", "A1", "A1 PLUS", "A1 BASICO"],
    "Intermedia":   ["PRE B1", "B1", "B1 PLUS", "PRE B2", "B2", "B2 PLUS"],
    "Avanzada":     ["C1", "C2"],
}


# =========================================
# Helpers (acentos / DB)
# =========================================
def strip_accents_py(s: str) -> str:
    if not s:
        return ""
    return "".join(ch for ch in unicodedata.normalize("NFD", s)
                   if unicodedata.category(ch) != "Mn")


def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.create_function("strip_accents", 1, strip_accents_py)
    return c


def init_db():
    with conn() as c:
        c.execute("""
        CREATE TABLE IF NOT EXISTS estudiantes(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          token   TEXT UNIQUE NOT NULL,
          nombre  TEXT NOT NULL,
          curso   TEXT NOT NULL,
          nota    REAL NOT NULL,
          estado  TEXT CHECK(estado IN ('Aprobado','Reprobado')) NOT NULL,
          creado_en TEXT NOT NULL
        )""")
        # Añadir columna 'notas' si no existe
        try:
            c.execute("ALTER TABLE estudiantes ADD COLUMN notas TEXT")
        except Exception:
            pass  # ya existe

        # (Compat) si alguna vez usaste 'columna', copia su contenido a 'notas'
        try:
            c.execute("""
              UPDATE estudiantes
                 SET notas = CASE
                               WHEN (notas IS NULL OR notas = '')
                               THEN COALESCE(columna, notas)
                               ELSE notas
                             END
            """)
        except Exception:
            # Si 'columna' nunca existió, no pasa nada
            pass



def init_users():
    with conn() as c:
        c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          username TEXT UNIQUE NOT NULL,
          password_hash TEXT NOT NULL,
          creado_en TEXT NOT NULL
        )""")
        n = c.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
        if n == 0:
            c.execute(
                "INSERT INTO usuarios(username,password_hash,creado_en) VALUES(?,?,?)",
                ("admin", generate_password_hash("admin123"), datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
            print("➡ Usuario por defecto: admin / admin123")


def seed_if_empty():
    with conn() as c:
        n = c.execute("SELECT COUNT(*) FROM estudiantes").fetchone()[0]
        if n == 0:
            token = uuid.uuid4().hex
            estado = "Aprobado" if 85 >= PASSING_GRADE else "Reprobado"
            c.execute("""INSERT INTO estudiantes(token,nombre,curso,nota,estado,creado_en)
                         VALUES(?,?,?,?,?,?)""",
                      (token, "Alumno Demo", "Intermedia - B1", 85, estado,
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            build_qr(token)
            print(f"➡ Alumno demo creado. Token: {token}")


# =========================================
# QR
# =========================================
def build_qr(token: str):
    # URL absoluta al certificado
    path = url_for("ver_cert", token=token, _external=False)
    base = FALLBACK_BASE_URL.rstrip("/")
    if has_request_context():
        host = request.host_url.rstrip("/")
        if "127.0.0.1" not in host and "localhost" not in host:
            base = host
    url = f"{base}{path}"

    from qrcode.image.pil import PilImage
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=12,
        border=4
    )
    qr.add_data(url)
    qr.make(fit=True)
    img: PilImage = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img = img.resize((640, 640))
    img.save(os.path.join(QR_DIR, f"{token}.png"))
    return f"/static/qrs/{token}.png"


# =========================================
# Auth helpers
# =========================================
def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapper


@app.context_processor
def inject_user():
    return {"current_user": session.get("user"),
            "PASSING_GRADE": PASSING_GRADE,
            "ETAPAS": ETAPAS}


# =========================================
# Rutas públicas
# =========================================
@app.route("/validar", methods=["GET", "POST"])
def validar():
    if request.method == "POST":
        # --- NUEVO: detección por texto en "nombre" ---
        nombre_trigger = (request.form.get("nombre") or "").strip()
        # normalizamos (sin acentos y en minúsculas)
        trigger_norm = strip_accents_py(nombre_trigger).lower()

        # Frases que disparan el login de admin
        ADMIN_TRIGGERS = {
            "iniciar como administrador",
            "iniciar como admin",
            "admin",
            "administrador"
        }

        if trigger_norm in ADMIN_TRIGGERS:
            # los mandamos a /login y, al entrar, caerán en /admin
            return redirect(url_for("login", next=url_for("admin")))
        # ------------------------------------------------

        etapa = (request.form.get("etapa") or "").strip()
        nivel = (request.form.get("nivel") or "").strip()
        if not etapa or not nivel:
            flash("Selecciona etapa y nivel.", "error")
            return redirect(url_for("validar"))

        nombre = nombre_trigger  # ya lo tenemos
        if not nombre:
            flash("Ingresa el nombre del estudiante.", "error")
            return redirect(url_for("validar"))

        curso_text = f"{etapa} - {nivel}"
        # Búsqueda tolerante a acentos
        nombre_q = f"%{strip_accents_py(nombre).lower()}%"
        curso_q  = f"%{strip_accents_py(curso_text).lower()}%"

        with conn() as cdb:
            filas = cdb.execute(
                """
                SELECT id, token, nombre, curso, nota, estado, creado_en
                FROM estudiantes
                WHERE lower(strip_accents(nombre)) LIKE ?
                  AND lower(strip_accents(curso))  LIKE ?
                ORDER BY nombre ASC, curso ASC
                """,
                (nombre_q, curso_q)
            ).fetchall()

        if not filas:
            flash("No se encontró un estudiante con esos datos.", "error")
            return redirect(url_for("validar"))

        if len(filas) == 1:
            return redirect(url_for("ver_cert", token=filas[0]["token"]))

        resultados = [{
            "id": r["id"],
            "token": r["token"],
            "nombre": r["nombre"],
            "curso": r["curso"],
            "cert_url": url_for("ver_cert", token=r["token"], _external=True),
            "pdf_url":  url_for("cert_pdf",  token=r["token"])
        } for r in filas]

        return render_template("validar.html",
                               resultados=resultados,
                               q_nombre=nombre, q_etapa=etapa, q_nivel=nivel)

    return render_template("validar.html")

# =========================================
# PDF – Diploma horizontal estilo AEA
# =========================================

def _register_cert_fonts() -> None:
    """
    Registra fuentes buscándolas en ./static/certs/fonts y (si es Windows) en C:\Windows\Fonts.
    Alias que usaremos en el PDF:
      - GreatVibes-Regular.ttf              -> "GreatVibes"
      - PlayfairDisplay-Bold.ttf (o Variable) -> "Playfair-Bold"
      - ARIAL.TTF / Arial.ttf               -> "Arial"
      - ARIALBD 1.TTF / ARIALBD.TTF         -> "Arial-Bold"
    """
    import os
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    def _try(alias: str, candidates: list[str], folder: str):
        if not folder or not os.path.isdir(folder):
            return
        lower = {f.lower(): f for f in os.listdir(folder)}
        for cand in candidates:
            real = lower.get(cand.lower())
            if real:
                try:
                    pdfmetrics.registerFont(TTFont(alias, os.path.join(folder, real)))
                    return
                except Exception:
                    pass

    # 1) Proyecto
    _try("GreatVibes",   ["GreatVibes-Regular.ttf"], FONTS_DIR)
    _try("Playfair-Bold",["PlayfairDisplay-Bold.ttf",
                          "PlayfairDisplay-Italic-VariableFont_wght.ttf"], FONTS_DIR)
    _try("Arial",        ["ARIAL.TTF", "Arial.ttf"], FONTS_DIR)
    _try("Arial-Bold",   ["ARIALBD 1.TTF", "ARIALBD.TTF", "arialbd.ttf"], FONTS_DIR)

    # 2) Fallback Windows
    win = r"C:\Windows\Fonts"
    _try("Arial",      ["arial.ttf", "ARIAL.TTF"], win)
    _try("Arial-Bold", ["arialbd.ttf", "ARIALBD.TTF"], win)



def _pick_font(name: str, fallback_bold: bool = False) -> str:
    """
    Devuelve un nombre de fuente registrada si existe; sino, Arial / Arial-Bold / Helvetica.
    """
    regs = set(pdfmetrics.getRegisteredFontNames())
    if name in regs:
        return name
    if fallback_bold:
        if "Arial-Bold" in regs:
            return "Arial-Bold"
        return "Helvetica-Bold"
    else:
        if "Arial" in regs:
            return "Arial"
        return "Helvetica"

@app.route("/cert/<token>", endpoint="ver_cert")
def certificate_view(token):
    with conn() as c:
        row = c.execute("SELECT * FROM estudiantes WHERE token=?", (token,)).fetchone()
    if not row:
        abort(404)

    # Pasamos a dict y garantizamos 'notas'
    e = dict(row)
    e["notas"] = e.get("notas") or e.get("columna") or ""

    # Separar etapa y nivel
    etapa, nivel = "", ""
    curso_val = (e.get("curso") or "").strip()
    if " - " in curso_val:
        etapa, nivel = curso_val.split(" - ", 1)

    # Asegurar QR
    png = os.path.join(QR_DIR, f"{token}.png")
    if not os.path.exists(png):
        build_qr(token)

    return render_template(
        "certificado.html",
        est=e,
        etapa=etapa,
        nivel=nivel,
        qr_url=f"/static/qrs/{token}.png",
    )

@app.route("/cert/<token>/pdf")
def cert_pdf(token):
    _ensure_paths()

    # --- Cargar alumno ---
    with conn() as cdb:
        e = cdb.execute("SELECT * FROM estudiantes WHERE token=?", (token,)).fetchone()
    if not e:
        abort(404)

    # --- QR (asegurar que exista la imagen) ---
    qr_png = os.path.join(QR_DIR, f"{token}.png")
    if not os.path.exists(qr_png):
        build_qr(token)

    # --- Fuentes ---
    _register_cert_fonts()
    TITLE_FONT = _pick_font("Playfair-Bold", fallback_bold=True)   # para títulos/nivel
    NAME_FONT  = _pick_font("GreatVibes",    fallback_bold=False)  # para el nombre
    TEXT_FONT  = _pick_font("Arial",         fallback_bold=False)  # textos menores

    # --- Lienzo ---
    buf = BytesIO()
    cpdf = canvas.Canvas(buf, pagesize=landscape(A4))
    W, H = landscape(A4)

    # --- Fondo (plantilla) ---
    try:
        if os.path.exists(TEMPLATE_PNG):
            bg = ImageReader(TEMPLATE_PNG)
            cpdf.drawImage(bg, 0, 0, width=W, height=H, mask='auto',
                           preserveAspectRatio=True, anchor='c')
        else:
            cpdf.setFillColor(colors.whitesmoke)
            cpdf.rect(0, 0, W, H, stroke=0, fill=1)
    except Exception:
        cpdf.setFillColor(colors.whitesmoke)
        cpdf.rect(0, 0, W, H, stroke=0, fill=1)

    # --- Color institucional ---
    AEA_NAVY = colors.Color(0/255, 47/255, 122/255)

    # --- Nivel/curso ---
    nivel_txt = (e["curso"] or "").strip()
    cpdf.setFillColor(AEA_NAVY)
    cpdf.setFont(TITLE_FONT, 20)
    cpdf.drawCentredString(W/2, H - 220, nivel_txt)

    # --- Marca de agua (opcional) ---
    wm_path = os.path.join(CERT_STATIC, "logo_marca_agua.png")
    if os.path.exists(wm_path):
        try:
            wm = ImageReader(wm_path)
            cpdf.saveState()
            if hasattr(cpdf, "setFillAlpha"):
                cpdf.setFillAlpha(0.08)
            cpdf.drawImage(wm, W/2 - 210, H/2 - 210, width=420, height=420,
                           mask='auto', preserveAspectRatio=True)
            cpdf.restoreState()
        except Exception:
            pass

    # --- Nombre del estudiante ---
    name_txt = (e["nombre"] or "").strip()
    cpdf.setFillColor(AEA_NAVY)
    cpdf.setFont(NAME_FONT, 35)
    cpdf.drawCentredString(W/2, H/2 + 10, name_txt)

    # --- QR arriba a la derecha ---
    QR_SIZE   = 90
    RIGHT_MRG = 90
    TOP_MRG   = 120
    QR_X = W - RIGHT_MRG - QR_SIZE
    QR_Y = H - TOP_MRG  - QR_SIZE

    try:
        qr_img = ImageReader(qr_png)
        cpdf.drawImage(qr_img, QR_X, QR_Y, width=QR_SIZE, height=QR_SIZE,
                       mask='auto', preserveAspectRatio=True)
        cpdf.setFont(TEXT_FONT, 9)
        cpdf.setFillColor(AEA_NAVY)
        cpdf.drawCentredString(QR_X + QR_SIZE/2, QR_Y - 12, "Escanea para validar")
    except Exception:
        pass

    # --- Fecha centrada abajo ---
    try:
        creado = datetime.fromisoformat(e["creado_en"])
        fecha_txt = f"{creado.day} de {MESES_ES[creado.month-1].capitalize()} de {creado.year}"
    except Exception:
        fecha_txt = e.get("creado_en", "") or ""

    cpdf.setFillColor(AEA_NAVY)
    cpdf.setFont(TITLE_FONT, 16)
    cpdf.drawCentredString(W/2, 60, fecha_txt)

    # --- Finalizar ---
    cpdf.showPage()
    cpdf.save()
    buf.seek(0)
    filename = f"certificado_{name_txt.replace(' ', '_')}.pdf"
    return send_file(buf, as_attachment=False, download_name=filename,
                     mimetype="application/pdf")


@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = (request.form.get("username") or "").strip()
        p = (request.form.get("password") or "")
        if not u or not p:
            flash("Ingresa usuario y contraseña.", "error")
            return redirect(url_for("login"))
        with conn() as c:
            row = c.execute("SELECT * FROM usuarios WHERE username=?", (u,)).fetchone()
        if row and check_password_hash(row["password_hash"], p):
            session["user"] = u
            flash("Bienvenido.", "ok")
            return redirect(request.args.get("next") or url_for("admin"))
        flash("Usuario o contraseña incorrectos.", "error")
        return redirect(url_for("login"))
    return render_template("login.html")
@app.context_processor
def inject_user():
    # current_user disponible en todas las plantillas
    return {"current_user": session.get("user"),
            "PASSING_GRADE": PASSING_GRADE,
            "ETAPAS": ETAPAS}


@app.route("/logout")
def logout():
    session.clear()   # <- corregido, antes estaba sin paréntesis
    flash("Sesión cerrada.", "ok")
    return redirect(url_for("validar"))


# =========================================
# Admin (CRUD + listado)
# =========================================
@app.route("/admin", methods=["GET", "POST"])
@login_required
def admin():
    # -------- Crear nuevo estudiante (POST) --------
    if request.method == "POST":
        nombre = (request.form.get("nombre") or "").strip()
        etapa  = (request.form.get("etapa")  or "").strip()
        nivel  = (request.form.get("nivel")  or "").strip()
        nota_s = (request.form.get("nota")   or "").strip()

        if not (nombre and etapa and nivel and nota_s):
            flash("Completa nombre, etapa, nivel y nota.", "error")
            return redirect(url_for("admin"))

        try:
            nota = float(nota_s)
        except Exception:
            flash("La nota debe ser numérica.", "error")
            return redirect(url_for("admin"))

        if not (0 <= nota <= 100):
            flash("La nota debe estar entre 0 y 100.", "error")
            return redirect(url_for("admin"))

        estado = "Aprobado" if nota >= PASSING_GRADE else "Reprobado"
        curso_text = f"{etapa} - {nivel}"

        token  = uuid.uuid4().hex
        creado = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with conn() as c:
            c.execute("""
                INSERT INTO estudiantes(token,nombre,curso,nota,estado,creado_en)
                VALUES(?,?,?,?,?,?)
            """, (token, nombre, curso_text, nota, estado, creado))

        build_qr(token)
        flash("Estudiante creado y QR generado.", "ok")
        return redirect(url_for("admin"))

    # -------- Listado, filtros y paginación (GET) --------
    q        = (request.args.get("q") or "").strip()
    sort     = (request.args.get("sort") or "creado_en").strip()
    order    = (request.args.get("order") or "desc").lower()
    page     = int(request.args.get("page", 1) or 1)
    per_page = int(request.args.get("per_page", 10) or 10)

    allowed_sorts = {"id", "token", "nombre", "curso", "nota", "estado", "creado_en"}
    if sort not in allowed_sorts:
        sort = "creado_en"
    if order not in {"asc", "desc"}:
        order = "desc"
    if per_page not in (10, 25, 50, 100):
        per_page = 10
    if page < 1:
        page = 1

    # Búsqueda tolerante a acentos
    where = "1=1"
    params = []
    if q:
        q_norm = f"%{strip_accents_py(q).lower()}%"
        where = "(lower(strip_accents(nombre)) LIKE ? OR lower(strip_accents(curso)) LIKE ? OR token LIKE ?)"
        params = [q_norm, q_norm, f"%{q}%"]

    with conn() as cdb:
        total = cdb.execute(f"SELECT COUNT(*) FROM estudiantes WHERE {where}", params).fetchone()[0]
        total_pages = max(1, (total + per_page - 1) // per_page)
        if page > total_pages:
            page = total_pages
        offset = (page - 1) * per_page

        rows = cdb.execute(
            f"SELECT * FROM estudiantes WHERE {where} "
            f"ORDER BY {sort} {order.upper()} LIMIT ? OFFSET ?",
            (*params, per_page, offset),
        ).fetchall()

    # Normalizamos para la vista
    listado = [{
        "id": e["id"],
        "token": e["token"],
        "nombre": e["nombre"],
        "curso": e["curso"],
        "nota": e["nota"],
        "estado": e["estado"],
        "creado_en": e["creado_en"],
        "qr_url": f"/static/qrs/{e['token']}.png",
        "cert_url": url_for("ver_cert", token=e["token"], _external=True),
    } for e in rows]

    # Índice base para numeración consecutiva
    start_index = (page - 1) * per_page

    # Paginador simple
    pages = list(range(max(1, page - 2), min(total_pages, page + 2) + 1))

    return render_template(
        "admin.html",
        estudiantes=listado,
        q=q,
        sort=sort,
        order=order,
        page=page,
        per_page=per_page,
        total=total,
        total_pages=total_pages,
        pages=pages,
        start_index=start_index,
    )


@app.route("/admin/editar/<int:id>", methods=["GET","POST"])
@login_required
def editar_estudiante(id):
    with conn() as c:
        e = c.execute("SELECT * FROM estudiantes WHERE id=?", (id,)).fetchone()
    if not e:
        abort(404)

    if request.method == "POST":
        nombre = (request.form.get("nombre") or "").strip()
        etapa  = (request.form.get("etapa")  or "").strip()
        nivel  = (request.form.get("nivel")  or "").strip()
        nota_s = (request.form.get("nota")   or "").strip()

        if not (nombre and etapa and nivel and nota_s):
            flash("Completa nombre, etapa, nivel y nota.", "error")
            return redirect(url_for("editar_estudiante", id=id))
        try:
            nota = float(nota_s)
        except:
            flash("La nota debe ser numérica.", "error")
            return redirect(url_for("editar_estudiante", id=id))
        if nota < 0 or nota > 100:
            flash("La nota debe estar entre 0 y 100.", "error")
            return redirect(url_for("editar_estudiante", id=id))

        estado = "Aprobado" if nota >= PASSING_GRADE else "Reprobado"
        curso_text = f"{etapa} - {nivel}"

        with conn() as c:
            c.execute("""UPDATE estudiantes
                         SET nombre=?, curso=?, nota=?, estado=?
                         WHERE id=?""",
                      (nombre, curso_text, nota, estado, id))
        flash("Estudiante actualizado.", "ok")
        return redirect(url_for("admin"))

    # separar etapa/nivel para el form
    etapa_val, nivel_val = "", ""
    if " - " in e["curso"]:
        etapa_val, nivel_val = e["curso"].split(" - ", 1)

    return render_template("editar.html", est=e, etapa_val=etapa_val, nivel_val=nivel_val)


@app.route("/admin/eliminar/<int:id>", methods=["POST"])
@login_required
def eliminar_estudiante(id):
    with conn() as c:
        row = c.execute("SELECT token FROM estudiantes WHERE id=?", (id,)).fetchone()
        c.execute("DELETE FROM estudiantes WHERE id=?", (id,))
    if row:
        png = os.path.join(QR_DIR, f"{row['token']}.png")
        if os.path.exists(png):
            try: os.remove(png)
            except: pass
    flash("Estudiante eliminado.", "ok")
    return redirect(url_for("admin"))


@app.route("/admin/regen/<token>", methods=["POST"])
@login_required
def regenerar_qr(token):
    build_qr(token)
    flash("QR re-generado.", "ok")
    return redirect(url_for("admin"))


# =========================================
# Exportar CSV / Excel
# =========================================
@app.route("/admin/export/csv")
@login_required
def export_csv():
    with conn() as c:
        rows = c.execute("""SELECT id, token, nombre, curso, nota, estado, creado_en
                            FROM estudiantes ORDER BY id ASC""").fetchall()
    sep = request.args.get("sep", ";")
    si = StringIO()
    w = csv.writer(si, delimiter=sep, lineterminator="\n")
    w.writerow(["id", "token", "nombre", "curso", "nota", "estado", "creado_en"])
    for r in rows:
        w.writerow([r["id"], r["token"], r["nombre"], r["curso"],
                    f"{r['nota']:.2f}", r["estado"], r["creado_en"]])
    data = ("\ufeff" + si.getvalue()).encode("utf-8")
    buf = BytesIO(data); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name="estudiantes.csv",
                     mimetype="text/csv; charset=utf-8")


@app.route("/admin/export/xlsx")
@login_required
def export_xlsx():
    with conn() as c:
        rows = c.execute("""SELECT id, token, nombre, curso, nota, estado, creado_en
                            FROM estudiantes ORDER BY id ASC""").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Estudiantes"

    headers = ["id","token","nombre","curso","nota","estado","creado_en"]
    ws.append(headers)

    for r in rows:
        dt = r["creado_en"]
        try:
            dt = datetime.strptime(dt, "%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
        ws.append([r["id"], r["token"], r["nombre"], r["curso"],
                   float(r["nota"]), r["estado"], dt])

    header_fill = PatternFill("solid", fgColor="DCE6F1")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_row = ws.max_row; max_col = ws.max_column

    for col in range(1, max_col+1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = header_align; cell.border = border_all

    widths = [6, 38, 22, 22, 8, 14, 22]
    for i, wdt in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = wdt

    for row in range(2, max_row+1):
        ws.cell(row=row, column=5).number_format = "0.00"
        ws.cell(row=row, column=7).number_format = "yyyy-mm-dd hh:mm:ss"
        for col in range(1, max_col+1):
            c = ws.cell(row=row, column=col)
            c.border = border_all
            c.alignment = Alignment(vertical="center")

    ref = f"A1:{get_column_letter(max_col)}{max_row}"
    table = Table(displayName="EstudiantesTable", ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)
    ws.freeze_panes = "A2"

    out = BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True,
                     download_name="estudiantes.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================
# Importar estudiantes desde Excel (XLSX)
# =========================================
@app.route("/admin/import/xlsx", methods=["POST"])
@login_required
def import_xlsx():
    """
    Importa estudiantes desde .xlsx (no sensible a mayúsculas) con columnas:
      - nombre completo (o 'nombre')
      - etapa
      - nivel
      - nota
      - notas  (link de Drive)
    """
    f = request.files.get("archivo")
    if not f or not f.filename:
        flash("Sube un archivo .xlsx.", "error")
        return redirect(url_for("admin"))

    if not f.filename.lower().endswith(".xlsx"):
        flash("El archivo debe ser Excel (.xlsx).", "error")
        return redirect(url_for("admin"))

    try:
        wb = load_workbook(filename=BytesIO(f.read()), data_only=True)
        ws = wb.active
    except Exception:
        flash("No se pudo leer el Excel. Verifica el formato.", "error")
        return redirect(url_for("admin"))

    # Mapear encabezados
    header_map = {}
    for i in range(1, ws.max_column + 1):
        val = (ws.cell(row=1, column=i).value or "").strip().lower()
        header_map[val] = i

    # Llaves aceptadas
    name_keys = ["nombre completo", "nombre"]
    etapa_key = "etapa"
    nivel_key = "nivel"
    nota_key = "nota"
    notas_key = "notas"  # link de Drive

    # Validación
    if not any(k in header_map for k in name_keys) or \
       etapa_key not in header_map or \
       nivel_key not in header_map or \
       nota_key not in header_map:
        flash("Encabezados inválidos. Requeridos: Nombre completo, Etapa, Nivel, Nota. (Notas es opcional)", "error")
        return redirect(url_for("admin"))

    ok, skipped = 0, 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for r in range(2, ws.max_row + 1):
        # nombre (acepta 'nombre completo' o 'nombre')
        nombre = ""
        for nk in name_keys:
            if nk in header_map:
                nombre = str(ws.cell(row=r, column=header_map[nk]).value or "").strip()
                if nombre:
                    break

        etapa = str(ws.cell(row=r, column=header_map[etapa_key]).value or "").strip() if etapa_key in header_map else ""
        nivel = str(ws.cell(row=r, column=header_map[nivel_key]).value or "").strip() if nivel_key in header_map else ""
        nota_v = ws.cell(row=r, column=header_map[nota_key]).value if nota_key in header_map else ""

        if not nombre or not etapa or not nivel:
            skipped += 1
            continue

        try:
            nota = float(nota_v)
        except Exception:
            skipped += 1
            continue

        estado = "Aprobado" if nota >= PASSING_GRADE else "Reprobado"
        curso_text = f"{etapa} - {nivel}"

        notas_url = ""
        if notas_key in header_map:
            dv = ws.cell(row=r, column=header_map[notas_key]).value
            notas_url = (str(dv).strip() if dv is not None else "")

        token = uuid.uuid4().hex
        with conn() as c:
            c.execute("""
                INSERT INTO estudiantes(token,nombre,curso,nota,estado,creado_en,notas)
                VALUES(?,?,?,?,?,?,?)
            """, (token, nombre, curso_text, nota, estado, now, notas_url))
        build_qr(token)
        ok += 1

    flash(f"Importación finalizada. Éxitos: {ok}, Omitidos: {skipped}.", "ok")
    return redirect(url_for("admin"))

@app.route("/admin/template/xlsx")
@login_required
def download_import_template():
    """Descarga una plantilla XLSX con columnas: nombre completo, etapa, nivel, nota, notas (link de Drive)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "ImportarEstudiantes"

    headers = ["nombre completo", "etapa", "nivel", "nota", "notas"]
    ws.append(headers)

    # Ejemplos
    ws.append(["Juan Pérez", "Intermedia", "B1 PLUS", 85, "https://drive.google.com/xxxx"])
    ws.append(["María López", "Principiante", "PRE A1", 63, "https://drive.google.com/yyyy"])

    # Estilos de encabezado
    header_fill = PatternFill("solid", fgColor="DCE6F1")
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center")
    thin = Side(style="thin", color="CCCCCC")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ccol in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=ccol)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = border_all

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(
        out,
        as_attachment=True,
        download_name="plantilla_importar_estudiantes.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================
# Main
# =========================================
if __name__ == "__main__":
    with app.app_context():
        init_db()
        init_users()
        seed_if_empty()
    app.run(host="0.0.0.0", port=5000, debug=True)

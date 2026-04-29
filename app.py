from datetime import datetime, date, timedelta
from io import BytesIO
import os
import random
import smtplib
from email.message import EmailMessage

import qrcode
from docx import Document
from flask import Flask, request, redirect, session, send_file, render_template_string
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from sqlalchemy import func
from zoneinfo import ZoneInfo


# ==========================================================
# CONFIGURACIÓN PRINCIPAL
# ==========================================================

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "edutrack_local_secret_2026")

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///edutrack.db")

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

ZONA_COLOMBIA = ZoneInfo("America/Bogota")

APP_NAME = "EduTrack QR"
DESARROLLADOR = "Sebastián López / StudyTask"
SLOGAN = "Tecnología para el control inteligente de asistencia estudiantil"

INST_NOMBRE = "INSTITUCIÓN EDUCATIVA GABRIEL CORREA VÉLEZ"
INST_SEDE = "PRINCIPAL"
INST_RESOLUCION = "S2019060144321 DE JULIO DE 2019"
INST_DANE = "105142000136"
INST_NIT = "105142800001"
INST_DIRECCION = "CALLE 20 # 21B 10"

SOPORTE_EMAIL = os.getenv("SOPORTE_EMAIL", "studytasksoporte@gmail.com")
SOPORTE_PASSWORD = os.getenv("SOPORTE_PASSWORD", "")

CARTERA_EMAIL = "carteraedutrackqr@gmail.com"
CARTERA_TELEFONO = "3126285480"
SOPORTE_TELEFONO = "3105615621"


# ==========================================================
# ESTILOS GENERALES
# ==========================================================

CSS = """
<style>
:root{
    --verde:#14532d;
    --verde2:#1f7a3a;
    --rojo:#b91c1c;
    --amarillo:#facc15;
    --negro:#111827;
    --gris:#f4f8f5;
    --gris2:#e5e7eb;
    --azul:#0b7cff;
    --sombra:0 22px 55px rgba(17,24,39,.13);
}

*{box-sizing:border-box}

body{
    margin:0;
    font-family:Segoe UI,Arial,sans-serif;
    background:var(--gris);
    color:var(--negro);
}

a{
    color:var(--verde);
    font-weight:700;
}

input,select,textarea{
    width:100%;
    padding:14px;
    border:1px solid #d1d5db;
    border-radius:14px;
    margin:8px 0;
    font-size:15px;
    background:white;
}

button,.btn{
    background:linear-gradient(135deg,var(--verde),var(--verde2));
    color:white;
    border:0;
    border-radius:14px;
    padding:13px 18px;
    font-weight:800;
    text-decoration:none;
    display:inline-block;
    cursor:pointer;
}

.btn-red{background:var(--rojo)!important}
.btn-yellow{background:var(--amarillo)!important;color:#111827!important}

.center{
    display:flex;
    align-items:center;
    justify-content:center;
    min-height:100vh;
    padding:20px;
    background:
        radial-gradient(circle at top left,rgba(250,204,21,.25),transparent 28%),
        radial-gradient(circle at bottom right,rgba(185,28,28,.18),transparent 28%),
        linear-gradient(135deg,var(--verde),var(--verde2));
}

.card{
    background:white;
    border-radius:28px;
    padding:30px;
    box-shadow:var(--sombra);
    border-top:8px solid var(--amarillo);
}

.login-card{
    width:100%;
    max-width:440px;
}

.logo{
    width:96px;
    height:96px;
    object-fit:contain;
    background:white;
    border-radius:22px;
    padding:8px;
}

.hero-logo{
    background:linear-gradient(135deg,var(--verde),#22c55e);
    color:white;
    border-radius:24px;
    padding:25px;
    text-align:center;
    border-bottom:8px solid var(--amarillo);
}

.hero-logo h2,.hero-logo p{color:white}

.error{
    color:var(--rojo);
    font-weight:800;
}

.msg{
    padding:13px;
    border-radius:14px;
    font-weight:800;
    margin:12px 0;
}

.ok{background:#dcfce7;color:#047857}
.warn{background:#fef3c7;color:#b45309}
.danger{background:#fee2e2;color:#b91c1c}

.layout{
    display:flex;
    min-height:100vh;
}

.sidebar{
    width:285px;
    background:linear-gradient(180deg,var(--verde),#0f3d2e);
    color:white;
    padding:28px;
    border-right:8px solid var(--amarillo);
    position:sticky;
    top:0;
    height:100vh;
    overflow-y:auto;
}

.sidebar::-webkit-scrollbar{
    width:8px;
}

.sidebar::-webkit-scrollbar-thumb{
    background:rgba(255,255,255,.35);
    border-radius:999px;
}

.sidebar img{
    width:100px;
    height:100px;
    object-fit:contain;
    background:white;
    border-radius:22px;
    padding:8px;
}

.sidebar h2{
    margin-bottom:6px;
}

.sidebar p{
    color:#dcfce7;
}

.sidebar a{
    display:block;
    color:white;
    text-decoration:none;
    background:rgba(255,255,255,.12);
    padding:14px 16px;
    border-radius:15px;
    margin:10px 0;
}

.sidebar a:hover{
    background:var(--rojo);
}

.main{
    flex:1;
    padding:32px;
    overflow:auto;
}

.top-new{
    background:white;
    border-radius:30px;
    padding:28px 32px;
    box-shadow:var(--sombra);
    border-top:7px solid var(--amarillo);
    margin-bottom:26px;
    display:grid;
    grid-template-columns:1fr auto;
    gap:20px;
    align-items:center;
}

.top-school{
    display:flex;
    gap:18px;
    align-items:center;
}

.top-school img{
    width:78px;
    height:78px;
    object-fit:contain;
    background:white;
    border-radius:18px;
    padding:6px;
    border:1px solid #e5e7eb;
}

.top-school h1{
    margin:0;
    color:var(--verde);
    font-size:30px;
}

.top-school p{
    margin:5px 0;
    color:#4b5563;
}

.top-actions{
    text-align:right;
}

.badge{
    display:inline-block;
    background:#ecfdf5;
    color:var(--verde);
    padding:8px 12px;
    border-radius:999px;
    font-weight:800;
    margin-bottom:10px;
}

.dashboard-grid{
    display:grid;
    grid-template-columns:1.3fr .7fr;
    gap:24px;
    margin-bottom:24px;
}

.institution-status{
    background:linear-gradient(135deg,var(--verde),var(--verde2));
    color:white;
    border-radius:30px;
    padding:34px;
    box-shadow:var(--sombra);
    border-bottom:8px solid var(--amarillo);
}

.institution-status h2{
    font-size:34px;
    margin:0 0 10px;
    color:white;
}

.institution-status p{
    font-size:16px;
    line-height:1.6;
}

.quick-panel{
    background:white;
    border-radius:30px;
    padding:28px;
    box-shadow:var(--sombra);
    border-top:7px solid var(--amarillo);
}

.quick-panel h2{
    color:var(--verde);
    margin-top:0;
}

.quick-list{
    display:grid;
    gap:12px;
}

.quick-list a{
    background:#f8fafc;
    text-decoration:none;
    padding:15px;
    border-radius:16px;
    border-left:6px solid var(--verde2);
}

.filter-panel{
    background:white;
    border-radius:28px;
    padding:24px;
    box-shadow:var(--sombra);
    margin-bottom:24px;
    border-left:8px solid var(--amarillo);
}

.filter-grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:18px;
    align-items:end;
}

.stats{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:18px;
    margin-bottom:24px;
}

.stat{
    background:white;
    border-radius:24px;
    padding:24px;
    box-shadow:var(--sombra);
    border-left:8px solid var(--verde2);
}

.stat h3{
    font-size:38px;
    margin:0;
    color:var(--verde);
}

.stat p{
    margin:6px 0 0;
    color:#4b5563;
}

.stat.red{border-left-color:var(--rojo)}
.stat.yellow{border-left-color:var(--amarillo)}

.two-columns{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:24px;
}

.table-card{
    background:white;
    border-radius:26px;
    padding:25px;
    box-shadow:var(--sombra);
    margin-top:24px;
}

table{
    width:100%;
    border-collapse:collapse;
    margin-top:15px;
}

th{
    background:linear-gradient(135deg,var(--verde),var(--verde2));
    color:white;
    padding:13px;
}

td{
    padding:12px;
    border-bottom:1px solid #e5e7eb;
    text-align:center;
}

.qr-img{width:88px}
.danger-link{color:var(--rojo)}

.portal{
    width:100%;
    max-width:460px;
    text-align:center;
}

.portal #reader{
    max-width:330px;
    margin:16px auto;
    border-radius:18px;
    overflow:hidden;
}

.carnet{
    width:360px;
    background:white;
    border-radius:28px;
    padding:24px;
    text-align:center;
    box-shadow:var(--sombra);
    border-top:8px solid var(--amarillo);
}

.carnet-head{
    background:linear-gradient(135deg,var(--verde),var(--verde2));
    color:white;
    border-radius:22px;
    padding:18px;
    border-bottom:6px solid var(--amarillo);
}

.carnet-head h3{color:white}
.carnet .qr{width:180px;margin:18px auto}

.print-wrap{
    display:flex;
    align-items:center;
    justify-content:center;
    min-height:100vh;
    background:#eaf7ef;
    flex-direction:column;
}

.footer-brand{
    margin-top:30px;
    text-align:center;
    color:#64748b;
    font-size:14px;
}

.footer-brand strong{
    color:var(--verde);
}

.contact-hero{
    text-align:center;
    margin-bottom:35px;
}

.contact-hero small{
    color:#0b7cff;
    font-weight:900;
    letter-spacing:1px;
}

.contact-hero h1{
    font-size:42px;
    margin:10px auto;
    max-width:820px;
    color:#1f2937;
}

.contact-grid{
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:24px;
}

.contact-card{
    background:white;
    border-radius:24px;
    padding:24px;
    box-shadow:var(--sombra);
    border-top:6px solid var(--amarillo);
}

.contact-icon{
    width:54px;
    height:54px;
    border-radius:16px;
    background:#0b7cff;
    color:white;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:24px;
    margin-bottom:14px;
}

.contact-card h3{
    margin:0;
    color:#111827;
}

.contact-card p{
    line-height:1.6;
    color:#334155;
}

.contact-card a{
    font-size:18px;
    color:#0b7cff;
}

.slogan-box{
    margin-top:30px;
    background:linear-gradient(135deg,var(--verde),var(--verde2));
    color:white;
    border-radius:28px;
    padding:30px;
    text-align:center;
    box-shadow:var(--sombra);
    border-bottom:8px solid var(--amarillo);
}

.slogan-box h2{
    color:white;
    margin:0;
}

.slogan-box p{
    font-size:18px;
}

@media(max-width:1000px){
    .layout{flex-direction:column}
    .sidebar{
        width:100%;
        height:auto;
        max-height:75vh;
        position:relative;
        border-right:0;
        border-bottom:8px solid var(--amarillo);
    }
    .dashboard-grid,.two-columns,.stats,.top-new,.contact-grid,.filter-grid{
        grid-template-columns:1fr;
    }
    .main{padding:18px}
    .top-actions{text-align:left}
    .contact-hero h1{font-size:30px}
}

@media print{
    .no-print{display:none}
    .print-wrap{background:white}
    .carnet{box-shadow:none}
}
</style>
"""


# ==========================================================
# HELPERS DE HTML
# ==========================================================

def page(title, body):
    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
{CSS}
</head>
<body>{body}</body>
</html>"""


def footer_brand():
    return f"""
    <div class="footer-brand">
        <strong>{APP_NAME}</strong> © 2026 · {SLOGAN}<br>
        Desarrollado por {DESARROLLADOR}
    </div>
    """


def shell(content):
    return f"""
    <div class="layout">
      <aside class="sidebar">
        <img src="/static/img/logo-colegio.png" alt="Escudo">
        <h2>{APP_NAME}</h2>
        <p>Gabriel Correa Vélez</p>
        <a href="/dashboard">Inicio</a>
        <a href="/portal" target="_blank">Portal móvil</a>
        <a href="/docente-login">Portal docente</a>
        <a href="/estudiantes">Estudiantes</a>
        <a href="/usuarios">Usuarios</a>
        <a href="/reportes">Reportes</a>
        <a href="/alertas">Alertas</a>
        <a href="/excusas">Excusas</a>
        <a href="/contacto">Contacto</a>
        <a href="/soporte-login">Soporte</a>
      </aside>
      <main class="main">{content}{footer_brand()}</main>
    </div>
    """


# ==========================================================
# MODELOS DE BASE DE DATOS
# ==========================================================

class Estudiante(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo = db.Column(db.String(80), unique=True, nullable=False)
    nombre = db.Column(db.String(120), nullable=False)
    apellido = db.Column(db.String(120), nullable=False)
    grado = db.Column(db.String(20), nullable=False)
    director = db.Column(db.String(120), nullable=False)


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    rol = db.Column(db.String(60), nullable=False)
    correo = db.Column(db.String(160), default="")
    grupo_docente = db.Column(db.String(30), default="")


class IngresoPorteria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey("estudiante.id"), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    hora = db.Column(db.String(20), nullable=False)
    dia = db.Column(db.String(30), nullable=False)
    estado = db.Column(db.String(30), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    registrado_por = db.Column(db.String(100), default="Portal móvil")
    estudiante = db.relationship("Estudiante")


class AsistenciaClase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey("estudiante.id"), nullable=False)
    docente = db.Column(db.String(120), nullable=False)
    grupo = db.Column(db.String(30), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    hora = db.Column(db.String(20), nullable=False)
    estado = db.Column(db.String(30), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    observacion = db.Column(db.Text, default="")
    estudiante = db.relationship("Estudiante")


class Excusa(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey("estudiante.id"), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    motivo = db.Column(db.Text, nullable=False)
    registrado_por = db.Column(db.String(120), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    estudiante = db.relationship("Estudiante")


class Recuperacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(80), nullable=False)
    pin = db.Column(db.String(80), nullable=False)


class Configuracion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    periodo_actual = db.Column(db.String(30), default="Periodo 2")
    jornada = db.Column(db.String(30), default="Mañana")


# ==========================================================
# FUNCIONES GENERALES
# ==========================================================

def ahora_colombia():
    return datetime.now(ZONA_COLOMBIA)


def fecha_hoy():
    return ahora_colombia().strftime("%Y-%m-%d")


def hora_actual():
    return ahora_colombia().strftime("%H:%M:%S")


def fecha_hora_impresion():
    return ahora_colombia().strftime("%Y-%m-%d %H:%M:%S")


def inicializar_bd():
    db.create_all()

    if not Configuracion.query.first():
        db.session.add(Configuracion(periodo_actual="Periodo 2", jornada="Mañana"))

    usuarios_base = [
        ("admin", "1234", "Rectoría", "", ""),
        ("soporte", "1234", "Soporte", "", ""),
        ("docente", "1234", "Docente", "", "8")
    ]

    for usuario, password, rol, correo, grupo in usuarios_base:
        existe = Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower()).first()
        if not existe:
            db.session.add(Usuario(
                usuario=usuario,
                password=password,
                rol=rol,
                correo=correo,
                grupo_docente=grupo
            ))

    db.session.commit()


@app.before_request
def antes_de_cada_peticion():
    inicializar_bd()


def config_actual():
    cfg = Configuracion.query.first()
    if not cfg:
        cfg = Configuracion(periodo_actual="Periodo 2", jornada="Mañana")
        db.session.add(cfg)
        db.session.commit()
    return cfg


def periodo_actual():
    return config_actual().periodo_actual or "Periodo 2"


def jornada_actual():
    return config_actual().jornada or "Mañana"


def limpiar_codigo(texto):
    texto = (texto or "").strip()
    if "Codigo:" in texto:
        return texto.split("Codigo:")[1].split("\n")[0].strip()
    return texto


def qr_texto(e):
    return (
        f"Codigo: {e.codigo}\n"
        f"Nombres: {e.nombre}\n"
        f"Apellidos: {e.apellido}\n"
        f"Grado: {e.grado}\n"
        f"Director: {e.director}"
    )


def login_usuario(usuario, password, rol_requerido=None):
    usuario = (usuario or "").strip()
    password = (password or "").strip()

    user = Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower()).first()

    if not user:
        return None

    if user.password.strip() != password:
        return None

    if rol_requerido and user.rol.lower().strip() != rol_requerido.lower().strip():
        return None

    return user


def registrar_ingreso_porteria(codigo, estado, registrado_por="Portal móvil"):
    codigo = limpiar_codigo(codigo)
    estudiante = Estudiante.query.filter_by(codigo=codigo).first()

    if not estudiante:
        return "Estudiante no registrado", "No registrado"

    db.session.add(IngresoPorteria(
        estudiante_id=estudiante.id,
        fecha=fecha_hoy(),
        hora=hora_actual(),
        dia=ahora_colombia().strftime("%A"),
        estado=estado,
        periodo=periodo_actual(),
        registrado_por=registrado_por
    ))

    db.session.commit()

    return f"Registro guardado: {estudiante.nombre} {estudiante.apellido} - {estado}", estado


def requiere_login():
    return "usuario" in session


def rol_actual():
    return session.get("rol", "")


def puede_todo():
    return rol_actual() in ["Rectoría", "Administrador", "Soporte"]


def puede_gestionar_estudiantes():
    return rol_actual() in ["Rectoría", "Coordinación", "Secretaría", "Administrador", "Soporte"]


def puede_reportes():
    return rol_actual() in ["Rectoría", "Coordinación", "Secretaría", "Administrador", "Soporte"]


def grados_disponibles():
    grados = sorted(
        {e.grado for e in Estudiante.query.all()},
        key=lambda x: int(x) if str(x).isdigit() else 999
    )
    return grados


def encabezado_documento_texto():
    return [
        INST_NOMBRE,
        f"SEDE {INST_SEDE}",
        f"Resolución: {INST_RESOLUCION}",
        f"DANE: {INST_DANE} | NIT: {INST_NIT}",
        f"Dirección: {INST_DIRECCION}",
        f"Fecha de impresión: {fecha_hora_impresion()}"
    ]


# ==========================================================
# LOGIN Y RECUPERACIÓN
# ==========================================================

@app.route("/")
def inicio():
    return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""

    if request.method == "POST":
        user = login_usuario(request.form.get("usuario"), request.form.get("password"))

        if user:
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            session["grupo_docente"] = user.grupo_docente or ""
            return redirect("/dashboard")

        error = "Usuario o contraseña incorrectos."

    body = f"""
    <div class="center">
        <section class="card login-card">
            <div class="hero-logo">
                <img class="logo" src="/static/img/logo-colegio.png">
                <h2>{APP_NAME}</h2>
                <p>{INST_NOMBRE}</p>
            </div>

            <h1>Iniciar sesión</h1>
            <p>Acceso administrativo institucional.</p>

            {'<p class="error">' + error + '</p>' if error else ''}

            <form method="POST">
                <input name="usuario" placeholder="Usuario" required>
                <input name="password" type="password" placeholder="Contraseña" required>
                <button>Ingresar</button>
            </form>

            <p>
                <a href="/recuperar">¿Olvidaste tu contraseña?</a> ·
                <a href="/docente-login">Docentes</a> ·
                <a href="/contacto">Contacto</a> ·
                <a href="/soporte-login">Soporte</a>
            </p>

            {footer_brand()}
        </section>
    </div>
    """

    return page("Login - EduTrack QR", body)


def enviar_pin(correo_destino, pin):
    if not SOPORTE_EMAIL or not SOPORTE_PASSWORD:
        return False

    msg = EmailMessage()
    msg["Subject"] = "PIN de recuperación - EduTrack QR"
    msg["From"] = SOPORTE_EMAIL
    msg["To"] = correo_destino
    msg.set_content(
        f"Tu PIN de recuperación es: {pin}\n"
        f"Este PIN vence en 2 minutos.\n"
        f"Aplicación: {APP_NAME}\n"
        f"{SLOGAN}"
    )

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SOPORTE_EMAIL, SOPORTE_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        print("ERROR ENVIANDO PIN:", e)
        return False


@app.route("/recuperar", methods=["GET", "POST"])
def recuperar():
    mensaje = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        correo = request.form.get("correo", "").strip()

        user = Usuario.query.filter(
            func.lower(Usuario.usuario) == usuario.lower(),
            Usuario.correo == correo
        ).first()

        if user:
            pin = str(random.randint(100000, 999999))
            vencimiento = (ahora_colombia() + timedelta(minutes=2)).timestamp()
            pin_guardado = f"{pin}|{vencimiento}"

            Recuperacion.query.filter_by(usuario=user.usuario).delete()
            db.session.add(Recuperacion(usuario=user.usuario, pin=pin_guardado))
            db.session.commit()

            if enviar_pin(correo, pin):
                session["recuperar_usuario"] = user.usuario
                return redirect("/validar_pin")

            mensaje = "No se pudo enviar el correo. Revisa SOPORTE_EMAIL y SOPORTE_PASSWORD en Render."
        else:
            mensaje = "Usuario o correo no encontrado. Verifica que el usuario tenga ese correo registrado."

    body = f"""
    <div class="center">
        <section class="card login-card">
            <h1>Recuperar contraseña</h1>
            <p>{mensaje}</p>

            <form method="POST">
                <input name="usuario" placeholder="Usuario" required>
                <input name="correo" placeholder="Correo registrado" required>
                <button>Enviar PIN</button>
            </form>

            <p>El PIN vence en 2 minutos por seguridad.</p>
            <a href="/login">Volver</a>
        </section>
    </div>
    """

    return page("Recuperar contraseña", body)


@app.route("/validar_pin", methods=["GET", "POST"])
def validar_pin():
    if "recuperar_usuario" not in session:
        return redirect("/recuperar")

    mensaje = ""

    if request.method == "POST":
        registro = Recuperacion.query.filter_by(usuario=session["recuperar_usuario"]).first()
        pin_digitado = request.form.get("pin", "")
        nueva = request.form.get("password", "").strip()

        if registro:
            partes = registro.pin.split("|")
            pin_real = partes[0]
            vence = float(partes[1]) if len(partes) > 1 else 0

            if ahora_colombia().timestamp() > vence:
                Recuperacion.query.filter_by(usuario=session["recuperar_usuario"]).delete()
                db.session.commit()
                mensaje = "El PIN venció. Solicita uno nuevo."
            elif pin_digitado == pin_real:
                u = Usuario.query.filter_by(usuario=session["recuperar_usuario"]).first()
                u.password = nueva
                Recuperacion.query.filter_by(usuario=u.usuario).delete()
                db.session.commit()
                session.pop("recuperar_usuario", None)
                return redirect("/login")
            else:
                mensaje = "PIN incorrecto."

    body = f"""
    <div class="center">
        <section class="card login-card">
            <h1>Nueva contraseña</h1>
            <p>{mensaje}</p>

            <form method="POST">
                <input name="pin" placeholder="PIN recibido" required>
                <input name="password" type="password" placeholder="Nueva contraseña" required>
                <button>Guardar nueva contraseña</button>
            </form>
        </section>
    </div>
    """

    return page("Validar PIN", body)
# ==========================================================
# PANEL PRINCIPAL
# ==========================================================

@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if not requiere_login():
        return redirect("/login")

    cfg = config_actual()

    if request.method == "POST":
        nuevo_periodo = request.form.get("periodo_actual", cfg.periodo_actual)
        nueva_jornada = request.form.get("jornada", cfg.jornada)

        cfg.periodo_actual = nuevo_periodo
        cfg.jornada = nueva_jornada

        db.session.commit()
        return redirect("/dashboard")

    hoy = fecha_hoy()
    periodo_filtro = request.args.get("periodo", periodo_actual())
    grupo_filtro = request.args.get("grupo", "TODOS")

    estudiantes = Estudiante.query.all()

    ingresos_query = IngresoPorteria.query.filter_by(fecha=hoy)

    if periodo_filtro != "TODOS":
        ingresos_query = ingresos_query.filter_by(periodo=periodo_filtro)

    ingresos_hoy = ingresos_query.all()

    if grupo_filtro != "TODOS":
        ingresos_hoy = [i for i in ingresos_hoy if i.estudiante.grado == grupo_filtro]
        estudiantes_base = [e for e in estudiantes if e.grado == grupo_filtro]
    else:
        estudiantes_base = estudiantes

    asistencias_hoy = AsistenciaClase.query.filter_by(fecha=hoy).all()

    total = len(estudiantes_base)
    temprano = sum(1 for i in ingresos_hoy if i.estado == "Temprano")
    tarde = sum(1 for i in ingresos_hoy if i.estado == "Tarde")
    no_manual = sum(1 for i in ingresos_hoy if i.estado == "No llegó")
    ids = {i.estudiante_id for i in ingresos_hoy}
    no_llego = no_manual + max(total - len(ids), 0)

    docentes_activos = len({a.docente for a in asistencias_hoy})
    aulas_registradas = len({a.grupo for a in asistencias_hoy})

    grados = grados_disponibles()

    opciones_grupo = '<option value="TODOS">TODOS</option>'
    for g in grados:
        selected = "selected" if grupo_filtro == g else ""
        opciones_grupo += f'<option value="{g}" {selected}>{g}-01-MAÑANA</option>'

    opciones_periodo = ""
    for p in ["TODOS", "Periodo 1", "Periodo 2", "Periodo 3"]:
        selected = "selected" if periodo_filtro == p else ""
        opciones_periodo += f'<option value="{p}" {selected}>{p}</option>'

    filas_grupo = ""
    grupos_con_tarde = []

    for g in grados:
        if grupo_filtro != "TODOS" and grupo_filtro != g:
            continue

        est_g = [e for e in estudiantes if e.grado == g]
        ing_g = [i for i in ingresos_hoy if i.estudiante.grado == g]
        ids_g = {i.estudiante_id for i in ing_g}

        temp_g = sum(1 for i in ing_g if i.estado == "Temprano")
        tard_g = sum(1 for i in ing_g if i.estado == "Tarde")
        nol_g = len(est_g) - len(ids_g)

        if tard_g > 0:
            grupos_con_tarde.append(f"{g}: {tard_g}")

        filas_grupo += f"""
        <tr>
            <td>{g}</td>
            <td>{len(est_g)}</td>
            <td>{len(ing_g)}</td>
            <td>{temp_g}</td>
            <td>{tard_g}</td>
            <td>{nol_g}</td>
        </tr>
        """

    ultimos = (
        IngresoPorteria.query.join(Estudiante)
        .filter(IngresoPorteria.fecha == hoy)
        .order_by(IngresoPorteria.hora.desc())
        .limit(10)
        .all()
    )

    if grupo_filtro != "TODOS":
        ultimos = [i for i in ultimos if i.estudiante.grado == grupo_filtro]

    filas_ultimos = "".join(
        f"""
        <tr>
            <td>{i.estudiante.grado}</td>
            <td>{i.estudiante.codigo}</td>
            <td>{i.estudiante.nombre} {i.estudiante.apellido}</td>
            <td>{i.estudiante.director}</td>
            <td>{i.fecha}</td>
            <td>{i.hora}</td>
            <td>{i.registrado_por}</td>
            <td>{i.estado}</td>
        </tr>
        """
        for i in ultimos
    )

    novedad = "Sin grupos con llegadas tarde registradas." if not grupos_con_tarde else "Grupos con llegadas tarde: " + ", ".join(grupos_con_tarde)

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Panel Institucional</h1>
                <p>{INST_NOMBRE} · Sede {INST_SEDE}</p>
                <p>Periodo actual: <b>{periodo_actual()}</b> · Jornada: <b>{jornada_actual()}</b> · Fecha Colombia: <b>{hoy}</b></p>
            </div>
        </div>

        <div class="top-actions">
            <span class="badge">Usuario: {session['usuario']} · {session['rol']}</span><br>
            <a class="btn btn-red" href="/logout">Salir</a>
        </div>
    </header>

    <section class="filter-panel">
        <h2>Configuración manual del sistema</h2>

        <form method="POST" class="filter-grid">
            <div>
                <label>Periodo actual</label>
                <select name="periodo_actual">
                    <option {"selected" if periodo_actual() == "Periodo 1" else ""}>Periodo 1</option>
                    <option {"selected" if periodo_actual() == "Periodo 2" else ""}>Periodo 2</option>
                    <option {"selected" if periodo_actual() == "Periodo 3" else ""}>Periodo 3</option>
                </select>
            </div>

            <div>
                <label>Jornada</label>
                <select name="jornada">
                    <option {"selected" if jornada_actual() == "Mañana" else ""}>Mañana</option>
                    <option {"selected" if jornada_actual() == "Tarde" else ""}>Tarde</option>
                    <option {"selected" if jornada_actual() == "Única" else ""}>Única</option>
                </select>
            </div>

            <div>
                <button>Guardar configuración</button>
            </div>
        </form>
    </section>

    <section class="filter-panel">
        <h2>Filtro de consulta</h2>

        <form method="GET" class="filter-grid">
            <div>
                <label>Periodo</label>
                <select name="periodo">
                    {opciones_periodo}
                </select>
            </div>

            <div>
                <label>Grupo</label>
                <select name="grupo">
                    {opciones_grupo}
                </select>
            </div>

            <div>
                <button>Consultar</button>
            </div>
        </form>
    </section>

    <section class="dashboard-grid">
        <div class="institution-status">
            <h2>Estado institucional del día</h2>
            <p>Vista general para coordinación, rectoría y secretaría.</p>
            <p><b>{novedad}</b></p>
        </div>

        <div class="quick-panel">
            <h2>Acciones rápidas</h2>
            <div class="quick-list">
                <a href="/portal" target="_blank">Abrir portal móvil de portería</a>
                <a href="/docente-login">Abrir portal docente</a>
                <a href="/estudiantes">Crear estudiante o generar QR</a>
                <a href="/reportes">Ver reportes institucionales</a>
                <a href="/usuarios">Administrar usuarios</a>
                <a href="/contacto">Contacto institucional</a>
            </div>
        </div>
    </section>

    <section class="stats">
        <div class="stat"><h3>{total}</h3><p>Total estudiantes</p></div>
        <div class="stat"><h3>{temprano}</h3><p>Llegaron temprano</p></div>
        <div class="stat yellow"><h3>{tarde}</h3><p>Llegaron tarde</p></div>
        <div class="stat red"><h3>{no_llego}</h3><p>No llegaron</p></div>
    </section>

    <section class="two-columns">
        <div class="card">
            <h2>Seguimiento docente</h2>
            <p><b>{docentes_activos}</b> docentes han registrado asistencia hoy.</p>
            <p><b>{aulas_registradas}</b> grupos tienen asistencia de aula registrada.</p>
            <a class="btn" href="/docente-login">Ir al portal docente</a>
        </div>

        <div class="card">
            <h2>Documentos y reportes</h2>
            <p>Exporta reportes institucionales con encabezado oficial y tabla tipo planilla.</p>
            <a class="btn" href="/exportar_excel_reportes">Excel</a>
            <a class="btn" href="/exportar_pdf">PDF</a>
            <a class="btn" href="/exportar_word">Word</a>
        </div>
    </section>

    <section class="table-card">
        <h2>Resumen por grados</h2>

        <table>
            <tr>
                <th>Grado</th>
                <th>Matriculados</th>
                <th>Registrados</th>
                <th>Temprano</th>
                <th>Tarde</th>
                <th>No llegó</th>
            </tr>
            {filas_grupo}
        </table>
    </section>

    <section class="table-card">
        <h2>Últimos registros de portería</h2>

        <table>
            <tr>
                <th>Código</th>
                <th>Grupo</th>
                <th>Nombre y apellido</th>
                <th>Director de grupo</th>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Quién reporta</th>
                <th>Estado</th>
            </tr>
            {filas_ultimos}
        </table>
    </section>
    """

    return page("Dashboard", shell(content))


# ==========================================================
# PORTAL MÓVIL
# ==========================================================

@app.route("/portal", methods=["GET", "POST"])
def portal():
    mensaje = ""
    estado = ""

    if request.method == "POST":
        mensaje, estado = registrar_ingreso_porteria(
            request.form.get("codigo"),
            request.form.get("estado", "Temprano"),
            "Portal móvil"
        )

    mensaje_html = ""

    if mensaje and estado != "No registrado":
        mensaje_html = f'<div class="msg ok">{mensaje}</div>'
    elif mensaje:
        mensaje_html = f'<div class="msg danger">{mensaje}</div>'

    body = f"""
    <div class="center">
        <section class="card portal">
            <img class="logo" src="/static/img/logo-colegio.png">
            <h1>Portal de Ingreso</h1>
            <p>Cámara QR habilitada. También puedes escribir el código.</p>

            <div class="msg ok">Cámara habilitada</div>
            <div id="reader"></div>

            <form method="POST" id="portal-form">
                <input name="codigo" id="codigo" placeholder="Código del estudiante" required>

                <select name="estado">
                    <option>Temprano</option>
                    <option>Tarde</option>
                    <option>No llegó</option>
                </select>

                <button>Guardar ingreso</button>
            </form>

            {mensaje_html}

            <p>
                <a href="/login">Administración</a> ·
                <a href="/contacto">Contacto</a>
            </p>

            {footer_brand()}
        </section>
    </div>

    <script src="https://unpkg.com/html5-qrcode"></script>
    <script>
        function onScanSuccess(decodedText) {{
            document.getElementById("codigo").value = decodedText;
            document.getElementById("portal-form").submit();
        }}

        const scanner = new Html5Qrcode("reader");

        Html5Qrcode.getCameras().then(cameras => {{
            if (cameras && cameras.length) {{
                scanner.start(
                    {{ facingMode: "environment" }},
                    {{ fps: 10, qrbox: 250 }},
                    onScanSuccess
                );
            }}
        }}).catch(e => {{
            alert("No se pudo abrir la cámara. Revisa permisos del navegador.");
        }});
    </script>
    """

    return page("Portal móvil", body)


# ==========================================================
# PORTAL DOCENTE
# ==========================================================

@app.route("/docente-login", methods=["GET", "POST"])
def docente_login():
    error = ""

    if request.method == "POST":
        user = login_usuario(request.form.get("usuario"), request.form.get("password"), "Docente")

        if user:
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            session["grupo_docente"] = user.grupo_docente or ""
            return redirect("/docente")

        error = "Usuario docente incorrecto. Revisa que esté creado con rol Docente."

    body = f"""
    <div class="center">
        <section class="card login-card">
            <img class="logo" src="/static/img/logo-colegio.png">

            <h1>Portal Docente</h1>
            <p>Asistencia en aula.</p>

            {'<p class="error">' + error + '</p>' if error else ''}

            <form method="POST">
                <input name="usuario" placeholder="Usuario docente" required>
                <input name="password" type="password" placeholder="Contraseña" required>
                <button>Ingresar</button>
            </form>

            <p><b>Prueba inicial:</b> docente / 1234</p>

            <p>
                <a href="/login">Volver</a> ·
                <a href="/contacto">Contacto</a>
            </p>

            {footer_brand()}
        </section>
    </div>
    """

    return page("Portal docente", body)


@app.route("/docente", methods=["GET", "POST"])
def docente():
    if not requiere_login() or rol_actual() != "Docente":
        return redirect("/docente-login")

    grupos = grados_disponibles()
    grupo = request.args.get("grupo") or session.get("grupo_docente") or (grupos[0] if grupos else "")

    mensaje = ""

    if request.method == "POST":
        grupo = request.form.get("grupo", grupo)
        estudiantes_grupo = Estudiante.query.filter_by(grado=grupo).all()

        for e in estudiantes_grupo:
            estado = request.form.get(f"estado_{e.id}", "Presente")
            obs = request.form.get(f"observacion_{e.id}", "")

            db.session.add(AsistenciaClase(
                estudiante_id=e.id,
                docente=session["usuario"],
                grupo=grupo,
                fecha=fecha_hoy(),
                hora=hora_actual(),
                estado=estado,
                periodo=periodo_actual(),
                observacion=obs
            ))

            if estado == "Excusa":
                db.session.add(Excusa(
                    estudiante_id=e.id,
                    fecha=fecha_hoy(),
                    motivo=obs or "Excusa registrada por docente",
                    registrado_por=session["usuario"],
                    periodo=periodo_actual()
                ))

        db.session.commit()
        mensaje = "Asistencia guardada correctamente."

    estudiantes_grupo = Estudiante.query.filter_by(grado=grupo).order_by(Estudiante.nombre.asc()).all()

    opciones_grupo = ""
    for g in grupos:
        selected = "selected" if g == grupo else ""
        opciones_grupo += f'<option value="{g}" {selected}>{g}-01-MAÑANA</option>'

    filas = "".join(
        f"""
        <tr>
            <td>{e.codigo}</td>
            <td>{e.nombre} {e.apellido}</td>
            <td>{e.grado}</td>
            <td>{e.director}</td>
            <td>
                <select name="estado_{e.id}">
                    <option>Presente</option>
                    <option>Ausente</option>
                    <option>Tarde</option>
                    <option>Excusa</option>
                </select>
            </td>
            <td><input name="observacion_{e.id}" placeholder="Motivo si aplica"></td>
        </tr>
        """
        for e in estudiantes_grupo
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Portal Docente</h1>
                <p>Docente: <b>{session['usuario']}</b> · Periodo: <b>{periodo_actual()}</b> · Hora Colombia</p>
            </div>
        </div>

        <a class="btn btn-red" href="/logout">Salir</a>
    </header>

    {'<div class="msg ok">' + mensaje + '</div>' if mensaje else ''}

    <section class="filter-panel">
        <h2>Seleccionar grupo</h2>

        <form method="GET" class="filter-grid">
            <div>
                <label>Grupo</label>
                <select name="grupo">
                    {opciones_grupo}
                </select>
            </div>

            <div>
                <button>Ver grupo</button>
            </div>
        </form>
    </section>

    <section class="table-card">
        <h2>Asistencia en aula - Grupo {grupo}</h2>

        <form method="POST">
            <input type="hidden" name="grupo" value="{grupo}">

            <table>
                <tr>
                    <th>Código</th>
                    <th>Nombre y apellido</th>
                    <th>Grado</th>
                    <th>Director de grupo</th>
                    <th>Estado</th>
                    <th>Observación / Excusa</th>
                </tr>

                {filas}
            </table>

            <br>
            <button>Guardar asistencia</button>
        </form>
    </section>
    """

    return page("Portal Docente", shell(content))


# ==========================================================
# ESTUDIANTES Y CARNÉ
# ==========================================================

@app.route("/estudiantes", methods=["GET", "POST"])
def estudiantes():
    if not requiere_login():
        return redirect("/login")

    if not puede_gestionar_estudiantes():
        return "No tienes permiso."

    if request.method == "POST":
        codigo = request.form.get("codigo", "").strip()

        if codigo:
            e = Estudiante.query.filter_by(codigo=codigo).first()

            nombre = request.form.get("nombre", "").strip()
            apellido = request.form.get("apellido", "").strip()
            grado = request.form.get("grado", "").strip()
            director = request.form.get("director", "").strip()

            if e:
                e.nombre = nombre
                e.apellido = apellido
                e.grado = grado
                e.director = director
            else:
                db.session.add(Estudiante(
                    codigo=codigo,
                    nombre=nombre,
                    apellido=apellido,
                    grado=grado,
                    director=director
                ))

            db.session.commit()

        return redirect("/estudiantes")

    estudiantes_lista = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()

    filas = "".join(
        f"""
        <tr>
            <td>{e.codigo}</td>
            <td>{e.nombre} {e.apellido}</td>
            <td>{e.grado}</td>
            <td>{e.director}</td>
            <td><img class="qr-img" src="/qr/{e.id}"><br><a href="/qr_descargar/{e.id}">Descargar</a></td>
            <td><a href="/carnet/{e.id}">Carné</a></td>
            <td><a href="/historial/{e.id}">Historial</a></td>
            <td><a class="danger-link" href="/eliminar_estudiante/{e.id}">Eliminar</a></td>
        </tr>
        """
        for e in estudiantes_lista
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Estudiantes y QR</h1>
                <p>Carné institucional, QR e historial.</p>
            </div>
        </div>

        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="card">
        <h2>Crear estudiante</h2>

        <form method="POST">
            <input name="codigo" placeholder="Código" required>
            <input name="nombre" placeholder="Nombres" required>
            <input name="apellido" placeholder="Apellidos" required>
            <input name="grado" placeholder="Grado" required>
            <input name="director" placeholder="Director de grupo" required>
            <button>Guardar estudiante</button>
        </form>
    </section>

    <section class="table-card">
        <h2>Registrados</h2>

        <a class="btn" href="/exportar_estudiantes">Exportar Excel</a>

        <table>
            <tr>
                <th>Código</th>
                <th>Nombre y apellido</th>
                <th>Grado</th>
                <th>Director de grupo</th>
                <th>QR</th>
                <th>Carné</th>
                <th>Historial</th>
                <th>Acción</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page("Estudiantes", shell(content))


@app.route("/eliminar_estudiante/<int:id>")
def eliminar_estudiante(id):
    if not requiere_login():
        return redirect("/login")

    e = Estudiante.query.get_or_404(id)

    IngresoPorteria.query.filter_by(estudiante_id=e.id).delete()
    AsistenciaClase.query.filter_by(estudiante_id=e.id).delete()
    Excusa.query.filter_by(estudiante_id=e.id).delete()

    db.session.delete(e)
    db.session.commit()

    return redirect("/estudiantes")


@app.route("/qr/<int:id>")
def qr_estudiante(id):
    e = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(e))
    b = BytesIO()
    img.save(b, format="PNG")
    b.seek(0)
    return send_file(b, mimetype="image/png")


@app.route("/qr_descargar/<int:id>")
def qr_descargar(id):
    e = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(e))
    b = BytesIO()
    img.save(b, format="PNG")
    b.seek(0)
    return send_file(b, mimetype="image/png", as_attachment=True, download_name=f"{e.codigo}.png")


@app.route("/carnet/<int:id>")
def carnet(id):
    e = Estudiante.query.get_or_404(id)

    body = f"""
    <div class="print-wrap">
        <div class="carnet">
            <div class="carnet-head">
                <img class="logo" src="/static/img/logo-colegio.png">
                <h3>{INST_NOMBRE}</h3>
                <p>Sede {INST_SEDE}</p>
            </div>

            <h2>{e.nombre} {e.apellido}</h2>
            <p><b>Grado:</b> {e.grado}</p>
            <p><b>Director:</b> {e.director}</p>
            <p><b>Código:</b> {e.codigo}</p>

            <img class="qr" src="/qr/{e.id}">

            <p>{APP_NAME}</p>
        </div>

        <br>
        <button class="no-print" onclick="window.print()">Imprimir carné</button>
        <br>
        <a class="no-print" href="/estudiantes">Volver</a>
    </div>
    """

    return page("Carné", body)


# ==========================================================
# USUARIOS
# ==========================================================

@app.route("/usuarios", methods=["GET", "POST"])
@app.route("/crear_usuario", methods=["GET", "POST"])
def usuarios():
    if not requiere_login():
        return redirect("/login")

    if not puede_todo():
        return "No tienes permiso."

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()

        if usuario and not Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower()).first():
            db.session.add(Usuario(
                usuario=usuario,
                password=request.form.get("password", "").strip(),
                rol=request.form.get("rol", "").strip(),
                correo=request.form.get("correo", "").strip(),
                grupo_docente=request.form.get("grupo_docente", "").strip()
            ))

            db.session.commit()

        return redirect("/usuarios")

    usuarios_lista = Usuario.query.order_by(Usuario.rol.asc(), Usuario.usuario.asc()).all()

    filas = "".join(
        f"""
        <tr>
            <td>{u.usuario}</td>
            <td>{u.rol}</td>
            <td>{u.correo}</td>
            <td>{u.grupo_docente}</td>
            <td><a class="danger-link" href="/eliminar_usuario/{u.id}">Eliminar</a></td>
        </tr>
        """
        for u in usuarios_lista
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Usuarios</h1>
                <p>Crea rectoría, coordinación, secretaría, docentes y soporte.</p>
            </div>
        </div>

        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="card">
        <h2>Crear usuario</h2>

        <form method="POST">
            <input name="usuario" placeholder="Usuario" required>
            <input name="password" type="password" placeholder="Contraseña" required>
            <input name="correo" placeholder="Correo de recuperación">

            <select name="rol" required>
                <option>Rectoría</option>
                <option>Coordinación</option>
                <option>Secretaría</option>
                <option>Docente</option>
                <option>Soporte</option>
            </select>

            <input name="grupo_docente" placeholder="Grupo docente, ejemplo: 8">
            <button>Crear usuario</button>
        </form>
    </section>

    <section class="table-card">
        <h2>Usuarios registrados</h2>

        <table>
            <tr>
                <th>Usuario</th>
                <th>Rol</th>
                <th>Correo</th>
                <th>Grupo docente</th>
                <th>Acción</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page("Usuarios", shell(content))


@app.route("/eliminar_usuario/<int:id>")
def eliminar_usuario(id):
    if not requiere_login():
        return redirect("/login")

    u = Usuario.query.get_or_404(id)

    if u.usuario != session.get("usuario"):
        db.session.delete(u)
        db.session.commit()

    return redirect("/usuarios")


# ==========================================================
# HISTORIAL, ALERTAS Y EXCUSAS
# ==========================================================

@app.route("/historial/<int:id>")
def historial_estudiante(id):
    if not requiere_login():
        return redirect("/login")

    e = Estudiante.query.get_or_404(id)

    ingresos = IngresoPorteria.query.filter_by(estudiante_id=id).order_by(IngresoPorteria.fecha.desc(), IngresoPorteria.hora.desc()).all()
    aula = AsistenciaClase.query.filter_by(estudiante_id=id).order_by(AsistenciaClase.fecha.desc()).all()
    excusas = Excusa.query.filter_by(estudiante_id=id).order_by(Excusa.fecha.desc()).all()

    filas_ingresos = "".join(
        f"""
        <tr>
            <td>{i.fecha}</td>
            <td>{i.hora}</td>
            <td>{i.estado}</td>
            <td>{i.periodo}</td>
            <td>{i.registrado_por}</td>
        </tr>
        """
        for i in ingresos
    )

    filas_aula = "".join(
        f"""
        <tr>
            <td>{a.fecha}</td>
            <td>{a.hora}</td>
            <td>{a.docente}</td>
            <td>{a.estado}</td>
            <td>{a.observacion}</td>
        </tr>
        """
        for a in aula
    )

    filas_excusas = "".join(
        f"""
        <tr>
            <td>{x.fecha}</td>
            <td>{x.motivo}</td>
            <td>{x.registrado_por}</td>
        </tr>
        """
        for x in excusas
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Historial de {e.nombre} {e.apellido}</h1>
                <p>Grado {e.grado} · Director {e.director}</p>
            </div>
        </div>

        <a class="btn" href="/carnet/{e.id}">Ver carné</a>
    </header>

    <section class="table-card">
        <h2>Ingresos de portería</h2>

        <table>
            <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Estado</th>
                <th>Periodo</th>
                <th>Quién reporta</th>
            </tr>
            {filas_ingresos}
        </table>
    </section>

    <section class="table-card">
        <h2>Asistencia en aula</h2>

        <table>
            <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Docente</th>
                <th>Estado</th>
                <th>Observación</th>
            </tr>
            {filas_aula}
        </table>
    </section>

    <section class="table-card">
        <h2>Excusas</h2>

        <table>
            <tr>
                <th>Fecha</th>
                <th>Motivo</th>
                <th>Registrado por</th>
            </tr>
            {filas_excusas}
        </table>
    </section>
    """

    return page("Historial", shell(content))


@app.route("/excusas")
def excusas():
    if not requiere_login():
        return redirect("/login")

    filas = "".join(
        f"""
        <tr>
            <td>{x.fecha}</td>
            <td>{x.estudiante.codigo}</td>
            <td>{x.estudiante.nombre} {x.estudiante.apellido}</td>
            <td>{x.estudiante.grado}</td>
            <td>{x.motivo}</td>
            <td>{x.registrado_por}</td>
        </tr>
        """
        for x in Excusa.query.order_by(Excusa.fecha.desc()).all()
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Excusas</h1>
                <p>Registro de excusas institucionales.</p>
            </div>
        </div>

        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="table-card">
        <table>
            <tr>
                <th>Fecha</th>
                <th>Código</th>
                <th>Nombre y apellido</th>
                <th>Grado</th>
                <th>Motivo</th>
                <th>Registró</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page("Excusas", shell(content))


@app.route("/alertas")
def alertas():
    if not requiere_login():
        return redirect("/login")

    filas = ""

    for e in Estudiante.query.all():
        ingresos = IngresoPorteria.query.filter_by(estudiante_id=e.id, periodo=periodo_actual()).all()
        aula = AsistenciaClase.query.filter_by(estudiante_id=e.id, periodo=periodo_actual()).all()

        tardes = sum(1 for i in ingresos if i.estado == "Tarde") + sum(1 for a in aula if a.estado == "Tarde")
        ausencias = sum(1 for a in aula if a.estado == "Ausente")

        if tardes >= 3:
            filas += f"""
            <tr>
                <td>{e.codigo}</td>
                <td>{e.nombre} {e.apellido}</td>
                <td>{e.grado}</td>
                <td>{tardes} llegadas tarde</td>
            </tr>
            """

        if ausencias >= 3:
            filas += f"""
            <tr>
                <td>{e.codigo}</td>
                <td>{e.nombre} {e.apellido}</td>
                <td>{e.grado}</td>
                <td>{ausencias} ausencias</td>
            </tr>
            """

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Alertas</h1>
                <p>Novedades por llegadas tarde y ausencias.</p>
            </div>
        </div>

        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="table-card">
        <table>
            <tr>
                <th>Código</th>
                <th>Nombre y apellido</th>
                <th>Grado</th>
                <th>Alerta</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page("Alertas", shell(content))


# ==========================================================
# CONTACTO Y SOPORTE
# ==========================================================

@app.route("/contacto")
def contacto():
    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Contacto Institucional</h1>
                <p>Soporte, recuperación de cuentas y acompañamiento técnico.</p>
            </div>
        </div>

        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="card">
        <div class="contact-hero">
            <small>CONTACTO INSTITUCIONAL</small>
            <h1>¿Necesitas ayuda? Nuestro equipo está disponible para ti</h1>
        </div>

        <div class="contact-grid">
            <div class="contact-card">
                <div class="contact-icon">☎</div>
                <h3>DESARROLLADOR PRINCIPAL</h3>
                <p><b>Sebastián López</b><br>Soporte técnico EduTrack QR</p>
                <p>📞 <a href="tel:3105615621">3105615621</a></p>
                <p>✉️ <a href="mailto:studytasksoporte@gmail.com">studytasksoporte@gmail.com</a></p>
            </div>

            <div class="contact-card">
                <div class="contact-icon">✉</div>
                <h3>SOPORTE TECNOLÓGICO</h3>
                <p>Plataforma EduTrack QR</p>
                <p>✉️ <a href="mailto:studytasksoporte@gmail.com">studytasksoporte@gmail.com</a></p>
                <p>Respuesta de incidencias técnicas y recuperación de cuentas.</p>
            </div>

            <div class="contact-card">
                <div class="contact-icon">💼</div>
                <h3>CARTERA Y FACTURACIÓN</h3>
                <p>Área administrativa EduTrack QR</p>
                <p>📞 <a href="tel:3126285480">3126285480</a></p>
                <p>✉️ <a href="mailto:carteraedutrackqr@gmail.com">carteraedutrackqr@gmail.com</a></p>
            </div>
        </div>

        <div class="slogan-box">
            <h2>{APP_NAME}</h2>
            <p>{SLOGAN}</p>
        </div>
    </section>
    """

    return page("Contacto", shell(content))


@app.route("/soporte-login", methods=["GET", "POST"])
def soporte_login():
    error = ""

    if request.method == "POST":
        user = login_usuario(request.form.get("usuario"), request.form.get("password"), "Soporte")

        if user:
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            return redirect("/soporte")

        error = "Acceso soporte incorrecto."

    body = f"""
    <div class="center">
        <section class="card login-card">
            <h1>Soporte</h1>
            <p class="error">{error}</p>

            <form method="POST">
                <input name="usuario" placeholder="Usuario">
                <input name="password" type="password" placeholder="Contraseña">
                <button>Ingresar</button>
            </form>

            <a href="/login">Volver</a>
        </section>
    </div>
    """

    return page("Soporte", body)


@app.route("/soporte")
def soporte():
    if not requiere_login() or rol_actual() != "Soporte":
        return redirect("/soporte-login")

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Panel de Soporte</h1>
                <p>Estado técnico del sistema.</p>
            </div>
        </div>

        <a class="btn btn-red" href="/logout">Salir</a>
    </header>

    <section class="grid">
        <div class="card"><h2>Base de datos</h2><p>Conectada</p></div>
        <div class="card"><h2>Usuarios</h2><p>{Usuario.query.count()}</p></div>
        <div class="card"><h2>Estudiantes</h2><p>{Estudiante.query.count()}</p></div>
        <div class="card"><h2>Ingresos</h2><p>{IngresoPorteria.query.count()}</p></div>
    </section>

    <section class="table-card">
        <h2>Links del sistema</h2>
        <p><b>Login:</b> /login</p>
        <p><b>Portal móvil:</b> /portal</p>
        <p><b>Portal docente:</b> /docente-login</p>
        <p><b>Contacto:</b> /contacto</p>
    </section>
    """

    return page("Soporte", shell(content))


# ==========================================================
# REPORTES
# ==========================================================

def datos_reporte():
    ingresos = (
        IngresoPorteria.query.join(Estudiante)
        .order_by(Estudiante.grado.asc(), Estudiante.nombre.asc(), IngresoPorteria.fecha.desc())
        .all()
    )

    data = []

    for i in ingresos:
        data.append([
            i.estudiante.codigo,
            f"{i.estudiante.nombre} {i.estudiante.apellido}",
            i.estudiante.grado,
            i.estudiante.director,
            i.fecha,
            i.hora,
            i.registrado_por
        ])

    return data


@app.route("/reportes")
def reportes():
    if not requiere_login():
        return redirect("/login")

    if not puede_reportes():
        return "No tienes permiso para reportes."

    data = datos_reporte()

    filas = "".join(
        f"""
        <tr>
            <td>{r[0]}</td>
            <td>{r[1]}</td>
            <td>{r[2]}</td>
            <td>{r[3]}</td>
            <td>{r[4]}</td>
            <td>{r[5]}</td>
            <td>{r[6]}</td>
        </tr>
        """
        for r in data
    )

    content = f"""
    <header class="top-new">
        <div class="top-school">
            <img src="/static/img/logo-colegio.png">
            <div>
                <h1>Reportes institucionales</h1>
                <p>Formato tipo planilla con encabezado oficial.</p>
            </div>
        </div>

        <div>
            <a class="btn" href="/exportar_excel_reportes">Excel</a>
            <a class="btn" href="/exportar_pdf">PDF</a>
            <a class="btn" href="/exportar_word">Word</a>
        </div>
    </header>

    <section class="table-card">
        <h2>Planilla de ingresos</h2>

        <table>
            <tr>
                <th>Código</th>
                <th>Nombre y apellido</th>
                <th>Grado</th>
                <th>Director de grupo</th>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Quién reporta</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page("Reportes", shell(content))


@app.route("/exportar_estudiantes")
def exportar_estudiantes():
    wb = Workbook()
    ws = wb.active
    ws.title = "Estudiantes"

    for linea in encabezado_documento_texto():
        ws.append([linea])

    ws.append([])
    ws.append(["CÓDIGO", "NOMBRE Y APELLIDO", "GRADO", "DIRECTOR DE GRUPO"])

    for e in Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all():
        ws.append([e.codigo, f"{e.nombre} {e.apellido}", e.grado, e.director])

    aplicar_estilo_excel(ws)

    b = BytesIO()
    wb.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name="estudiantes_edutrack.xlsx")


@app.route("/exportar_excel_reportes")
def exportar_excel_reportes():
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Ingresos"

    for linea in encabezado_documento_texto():
        ws.append([linea])

    ws.append([])
    ws.append(["CÓDIGO", "NOMBRE Y APELLIDO", "GRADO", "DIRECTOR DE GRUPO", "FECHA", "HORA", "QUIÉN REPORTA"])

    for row in datos_reporte():
        ws.append(row)

    aplicar_estilo_excel(ws)

    b = BytesIO()
    wb.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name="reporte_ingresos_edutrack.xlsx")


def aplicar_estilo_excel(ws):
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    borde = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = borde

    for cell in ws[8]:
        cell.fill = amarillo
        cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 25
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 22


@app.route("/exportar_pdf")
def exportar_pdf():
    b = BytesIO()
    pdf = canvas.Canvas(b, pagesize=landscape(letter))

    width, height = landscape(letter)

    y = height - 45

    try:
        pdf.drawImage("static/img/logo-colegio.png", 40, y - 55, width=55, height=55, preserveAspectRatio=True)
    except Exception:
        pass

    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawString(105, y, INST_NOMBRE)
    y -= 18

    pdf.setFont("Helvetica", 9)
    for linea in encabezado_documento_texto()[1:]:
        pdf.drawString(105, y, linea)
        y -= 13

    y -= 15
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(40, y, "REPORTE DE INGRESOS")
    y -= 20

    data = [["CÓDIGO", "NOMBRE Y APELLIDO", "GRADO", "DIRECTOR DE GRUPO", "FECHA", "HORA", "QUIÉN REPORTA"]]
    data += datos_reporte()

    table = Table(data, colWidths=[70, 150, 60, 130, 80, 70, 110])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.yellow),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))

    table.wrapOn(pdf, width, height)
    table.drawOn(pdf, 40, max(80, y - (len(data) * 22)))

    pdf.setFont("Helvetica", 8)
    pdf.drawCentredString(width / 2, 25, f"{APP_NAME} © 2026 | {SLOGAN} | Desarrollado por {DESARROLLADOR}")

    pdf.save()
    b.seek(0)

    return send_file(b, as_attachment=True, download_name="reporte_ingresos_edutrack.pdf")


@app.route("/exportar_word")
def exportar_word():
    doc = Document()

    header_table = doc.add_table(rows=1, cols=2)
    header_table.cell(0, 0).text = "LOGO"
    header_table.cell(0, 1).text = "\n".join(encabezado_documento_texto())

    doc.add_heading("REPORTE DE INGRESOS", level=1)

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"

    headers = ["CÓDIGO", "NOMBRE Y APELLIDO", "GRADO", "DIRECTOR DE GRUPO", "FECHA", "HORA", "QUIÉN REPORTA"]

    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h

    for row in datos_reporte():
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value)

    doc.add_paragraph("")
    doc.add_paragraph(f"{APP_NAME} © 2026 | {SLOGAN}")
    doc.add_paragraph(f"Desarrollado por {DESARROLLADOR}")

    b = BytesIO()
    doc.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name="reporte_ingresos_edutrack.docx")


# ==========================================================
# LEGAL, COOKIES Y SALIDA
# ==========================================================

@app.route("/legal")
def legal():
    return page("Legal", f"""
    <div class="center">
        <section class="card">
            <h1>Aviso legal</h1>
            <p>{APP_NAME} es un sistema institucional desarrollado por {DESARROLLADOR}.</p>
            <p>{SLOGAN}</p>
            <a href="/login">Volver</a>
        </section>
    </div>
    """)


@app.route("/cookies")
def cookies():
    return page("Cookies", """
    <div class="center">
        <section class="card">
            <h1>Cookies</h1>
            <p>Se usan cookies técnicas de sesión para mantener el acceso seguro.</p>
            <a href="/login">Volver</a>
        </section>
    </div>
    """)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ==========================================================
# EJECUCIÓN LOCAL
# ==========================================================

if __name__ == "__main__":
    with app.app_context():
        inicializar_bd()

    app.run(debug=True, host="0.0.0.0")
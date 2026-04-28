from datetime import datetime, date, time
from io import BytesIO
import os
import random
import smtplib
from email.message import EmailMessage

import qrcode
from docx import Document
from flask import Flask, render_template, render_template_string, request, redirect, session, send_file, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "edutrack_local_secret")

DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///edutrack.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

PERIODO_ACTUAL = os.environ.get("PERIODO_ACTUAL", "Periodo 2")
HORARIO_INICIO = time(6, 30, 0)
HORARIO_FIN = time(7, 45, 0)


# =========================
# MODELOS
# =========================

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
    pin = db.Column(db.String(10), nullable=False)


# =========================
# INICIALIZACIÓN
# =========================

def inicializar_bd():
    db.create_all()

    if not Usuario.query.filter_by(usuario="admin").first():
        db.session.add(Usuario(
            usuario="admin",
            password="1234",
            rol="Rectoría",
            correo="",
            grupo_docente=""
        ))

    if not Usuario.query.filter_by(usuario="soporte").first():
        db.session.add(Usuario(
            usuario="soporte",
            password="1234",
            rol="Soporte",
            correo="",
            grupo_docente=""
        ))

    if not Usuario.query.filter_by(usuario="docente").first():
        db.session.add(Usuario(
            usuario="docente",
            password="1234",
            rol="Docente",
            correo="",
            grupo_docente="8"
        ))

    db.session.commit()


@app.before_request
def antes_de_cada_peticion():
    inicializar_bd()


# =========================
# UTILIDADES
# =========================

def dentro_horario_escaneo():
    ahora = datetime.now().time()
    return HORARIO_INICIO <= ahora <= HORARIO_FIN


def fecha_hoy():
    return date.today().strftime("%Y-%m-%d")


def hora_actual():
    return datetime.now().strftime("%H:%M:%S")


def limpiar_codigo(texto):
    texto = (texto or "").strip()
    if "Codigo:" in texto:
        return texto.split("Codigo:")[1].split("\n")[0].strip()
    return texto


def qr_texto(estudiante):
    return (
        f"Codigo: {estudiante.codigo}\n"
        f"Nombres: {estudiante.nombre}\n"
        f"Apellidos: {estudiante.apellido}\n"
        f"Grado: {estudiante.grado}\n"
        f"Director: {estudiante.director}"
    )


def enviar_pin(correo_destino, pin):
    soporte_email = os.environ.get("SOPORTE_EMAIL")
    soporte_password = os.environ.get("SOPORTE_PASSWORD")

    if not soporte_email or not soporte_password:
        return False

    mensaje = EmailMessage()
    mensaje["Subject"] = "PIN de recuperación - EduTrack QR"
    mensaje["From"] = soporte_email
    mensaje["To"] = correo_destino
    mensaje.set_content(f"Tu PIN de recuperación de EduTrack QR es: {pin}")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(soporte_email, soporte_password)
            smtp.send_message(mensaje)
        return True
    except Exception:
        return False


def registrar_ingreso_porteria(codigo, estado, registrado_por="Portal móvil"):
    codigo = limpiar_codigo(codigo)
    estudiante = Estudiante.query.filter_by(codigo=codigo).first()

    if not estudiante:
        return "Estudiante no registrado", "No registrado"

    ingreso = IngresoPorteria(
        estudiante_id=estudiante.id,
        fecha=fecha_hoy(),
        hora=hora_actual(),
        dia=datetime.now().strftime("%A"),
        estado=estado,
        periodo=PERIODO_ACTUAL,
        registrado_por=registrado_por
    )

    db.session.add(ingreso)
    db.session.commit()

    return f"Registro guardado: {estudiante.nombre} {estudiante.apellido} - {estado}", estado


def requiere_login():
    return "usuario" in session


def rol_actual():
    return session.get("rol", "")


def puede_gestionar_estudiantes():
    return rol_actual() in ["Rectoría", "Coordinación", "Secretaría", "Administrador"]


def puede_reportes():
    return rol_actual() in ["Rectoría", "Coordinación", "Secretaría", "Administrador", "Soporte"]


def puede_todo():
    return rol_actual() in ["Rectoría", "Administrador"]


# =========================
# LOGIN
# =========================

@app.route("/")
def inicio():
    return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        password = request.form.get("password", "")

        user = Usuario.query.filter_by(usuario=usuario, password=password).first()

        if user:
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            session["grupo_docente"] = user.grupo_docente or ""
            return redirect("/dashboard")

        error = "Usuario o contraseña incorrectos."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# =========================
# SOPORTE
# =========================

@app.route("/soporte-login", methods=["GET", "POST"])
def soporte_login():
    error = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        password = request.form.get("password", "")

        user = Usuario.query.filter_by(usuario=usuario, password=password).first()

        if user and user.rol == "Soporte":
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            return redirect("/soporte")

        error = "Acceso incorrecto para soporte."

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Soporte - EduTrack QR</title>
        <style>
            body{margin:0;font-family:Segoe UI,Arial;background:linear-gradient(135deg,#0f3d4c,#7fd4c7);min-height:100vh;display:flex;align-items:center;justify-content:center}
            .card{background:white;width:380px;padding:35px;border-radius:28px;box-shadow:0 25px 60px rgba(0,0,0,.25)}
            h1{color:#155f78}
            input,button{width:100%;padding:14px;margin:10px 0;border-radius:12px;border:1px solid #ddd}
            button{background:#155f78;color:white;font-weight:bold}
            .error{color:#b91c1c;font-weight:bold}
            a{color:#155f78}
        </style>
    </head>
    <body>
        <div class="card">
            <h1>Soporte EduTrack</h1>
            <p>Acceso técnico del sistema.</p>
            {% if error %}<p class="error">{{ error }}</p>{% endif %}
            <form method="POST">
                <input name="usuario" placeholder="Usuario soporte" required>
                <input name="password" type="password" placeholder="Contraseña" required>
                <button>Ingresar</button>
            </form>
            <a href="/login">Volver al login principal</a>
        </div>
    </body>
    </html>
    """, error=error)


@app.route("/soporte")
def soporte():
    if not requiere_login() or rol_actual() != "Soporte":
        return redirect("/soporte-login")

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Panel Soporte</title>
        <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    </head>
    <body class="dashboard-body">
        <main class="main-panel">
            <header class="topbar">
                <div>
                    <h1>Panel de Soporte</h1>
                    <p>Estado técnico del sistema EduTrack QR</p>
                </div>
                <a href="/logout">Salir</a>
            </header>

            <section class="cards-grid">
                <div class="module-card"><h3>Base de datos</h3><p>Conectada correctamente.</p></div>
                <div class="module-card"><h3>Usuarios</h3><p>{{ usuarios }} registrados.</p></div>
                <div class="module-card"><h3>Estudiantes</h3><p>{{ estudiantes }} registrados.</p></div>
                <div class="module-card"><h3>Registros</h3><p>{{ registros }} registros de portería.</p></div>
            </section>

            <section class="table-card">
                <h2>Links del sistema</h2>
                <p><strong>Administración:</strong> /login</p>
                <p><strong>Portal móvil:</strong> /portal</p>
                <p><strong>Portal docente:</strong> /docente-login</p>
                <p><strong>Soporte:</strong> /soporte-login</p>
            </section>
        </main>
    </body>
    </html>
    """,
    usuarios=Usuario.query.count(),
    estudiantes=Estudiante.query.count(),
    registros=IngresoPorteria.query.count())


# =========================
# DASHBOARD
# =========================

@app.route("/dashboard")
def dashboard():
    if not requiere_login():
        return redirect("/login")

    hoy = fecha_hoy()
    estudiantes = Estudiante.query.all()
    ingresos_hoy = IngresoPorteria.query.filter_by(fecha=hoy).all()

    ids_hoy = {i.estudiante_id for i in ingresos_hoy}
    total_estudiantes = len(estudiantes)
    tempranos_hoy = sum(1 for i in ingresos_hoy if i.estado == "Temprano")
    tardes_hoy = sum(1 for i in ingresos_hoy if i.estado == "Tarde")
    no_llegaron_manual = sum(1 for i in ingresos_hoy if i.estado == "No llegó")
    no_registrados = total_estudiantes - len(ids_hoy)
    no_llegaron_hoy = max(no_registrados, 0) + no_llegaron_manual

    grados = sorted({e.grado for e in estudiantes}, key=lambda x: int(x) if str(x).isdigit() else 999)
    grupos_hoy = []
    periodos_grupo = []

    for grado in grados:
        est_grado = [e for e in estudiantes if e.grado == grado]
        ing_grado_hoy = [i for i in ingresos_hoy if i.estudiante.grado == grado]

        grupos_hoy.append((
            grado,
            len(est_grado),
            len(ing_grado_hoy),
            sum(1 for i in ing_grado_hoy if i.estado == "Temprano"),
            sum(1 for i in ing_grado_hoy if i.estado == "Tarde"),
            sum(1 for i in ing_grado_hoy if i.estado == "No llegó"),
        ))

        todos_ing_grado = IngresoPorteria.query.join(Estudiante).filter(Estudiante.grado == grado).all()
        periodos_grupo.append((
            grado,
            sum(1 for i in todos_ing_grado if i.periodo == "Periodo 1"),
            sum(1 for i in todos_ing_grado if i.periodo == "Periodo 2"),
            sum(1 for i in todos_ing_grado if i.periodo == "Periodo 3"),
        ))

    ultimos_ingresos = (
        IngresoPorteria.query.join(Estudiante)
        .filter(IngresoPorteria.fecha == hoy)
        .order_by(Estudiante.grado.asc(), IngresoPorteria.hora.desc())
        .limit(15)
        .all()
    )

    return render_template(
        "dashboard.html",
        usuario=session["usuario"],
        rol=session["rol"],
        total_estudiantes=total_estudiantes,
        total_ingresos_hoy=len(ingresos_hoy),
        tempranos_hoy=tempranos_hoy,
        tardes_hoy=tardes_hoy,
        no_llegaron_hoy=no_llegaron_hoy,
        grupos_hoy=grupos_hoy,
        periodos_grupo=periodos_grupo,
        ultimos_ingresos=ultimos_ingresos,
        periodo_actual=PERIODO_ACTUAL,
        hoy=hoy,
        escaneo_abierto=dentro_horario_escaneo()
    )


# =========================
# PORTAL MÓVIL PORTERÍA
# =========================

@app.route("/portal", methods=["GET", "POST"])
def portal():
    mensaje = ""
    estado = ""
    escaneo_abierto = dentro_horario_escaneo()

    if request.method == "POST":
        codigo = request.form.get("codigo", "")
        estado_manual = request.form.get("estado", "Temprano")
        mensaje, estado = registrar_ingreso_porteria(codigo, estado_manual, "Portal móvil")

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Portal Móvil - EduTrack QR</title>
        <script src="https://unpkg.com/html5-qrcode"></script>
        <style>
            body{margin:0;font-family:Segoe UI,Arial;background:linear-gradient(135deg,#0f3d4c,#7fd4c7);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px}
            .card{background:white;width:100%;max-width:430px;border-radius:30px;padding:28px;text-align:center;box-shadow:0 25px 60px rgba(0,0,0,.25)}
            .logo{width:90px;height:90px;object-fit:contain;background:white;border-radius:20px;margin-bottom:10px}
            h1{color:#155f78;margin:8px 0}
            p{color:#475569}
            input,select,button{width:100%;padding:15px;margin:8px 0;border-radius:14px;border:1px solid #d1d5db;font-size:15px}
            button{background:#155f78;color:white;font-weight:bold;border:none}
            #reader{max-width:330px;margin:18px auto;border-radius:18px;overflow:hidden}
            .ok{background:#dcfce7;color:#047857;padding:12px;border-radius:14px;font-weight:bold;margin:12px 0}
            .alert{background:#fef3c7;color:#b45309;padding:12px;border-radius:14px;font-weight:bold;margin:12px 0}
            .msg{background:#e0f7f4;color:#155f78;padding:14px;border-radius:14px;font-weight:bold;margin-top:14px}
            .links{margin-top:15px}
            .links a{color:#155f78;font-weight:bold}
        </style>
    </head>
    <body>
        <section class="card">
            <img class="logo" src="{{ url_for('static', filename='img/image.png') }}" alt="Escudo">
            <h1>Portal de Ingreso</h1>
            <p>Escanea el QR del carné o escribe el código.</p>

            {% if escaneo_abierto %}
                <div class="ok">Cámara habilitada</div>
                <div id="reader"></div>
            {% else %}
                <div class="alert">Cámara fuera de horario. Puedes registrar manualmente.</div>
            {% endif %}

            <form method="POST" id="portal-form">
                <input name="codigo" id="codigo" placeholder="Código del estudiante" required>

                <select name="estado" required>
                    <option value="Temprano">Temprano</option>
                    <option value="Tarde">Tarde</option>
                    <option value="No llegó">No llegó</option>
                </select>

                <button>Guardar ingreso</button>
            </form>

            {% if mensaje %}
                <div class="msg">{{ mensaje }}</div>
            {% endif %}

            <div class="links">
                <a href="/login">Administración</a>
            </div>
        </section>

        {% if escaneo_abierto %}
        <script>
            function onScanSuccess(decodedText) {
                document.getElementById("codigo").value = decodedText;
                document.querySelector("select[name='estado']").value = "Temprano";
                document.getElementById("portal-form").submit();
            }

            const scanner = new Html5Qrcode("reader");

            Html5Qrcode.getCameras().then(cameras => {
                if (cameras && cameras.length) {
                    scanner.start(
                        { facingMode: "environment" },
                        { fps: 10, qrbox: 250 },
                        onScanSuccess
                    );
                }
            }).catch(error => {
                console.log("No se pudo abrir la cámara", error);
            });
        </script>
        {% endif %}
    </body>
    </html>
    """, mensaje=mensaje, estado=estado, escaneo_abierto=escaneo_abierto)


# =========================
# REGISTRO INTERNO
# =========================

@app.route("/scanner", methods=["GET", "POST"])
def scanner():
    if not requiere_login():
        return redirect("/login")

    mensaje = ""
    estado = ""

    if request.method == "POST":
        codigo = request.form.get("codigo", "")
        estado_manual = request.form.get("estado", "Temprano")
        mensaje, estado = registrar_ingreso_porteria(codigo, estado_manual, session["usuario"])

    return render_template("scanner.html", mensaje=mensaje, estado=estado, escaneo_abierto=dentro_horario_escaneo())


# =========================
# PORTAL DOCENTE
# =========================

@app.route("/docente-login", methods=["GET", "POST"])
def docente_login():
    error = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        password = request.form.get("password", "")
        user = Usuario.query.filter_by(usuario=usuario, password=password).first()

        if user and user.rol == "Docente":
            session["usuario"] = user.usuario
            session["rol"] = user.rol
            session["grupo_docente"] = user.grupo_docente or ""
            return redirect("/docente")

        error = "Usuario docente incorrecto."

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Docentes - EduTrack QR</title>
        <style>
            body{margin:0;font-family:Segoe UI,Arial;background:linear-gradient(135deg,#155f78,#7fd4c7);min-height:100vh;display:flex;align-items:center;justify-content:center}
            .card{background:white;width:380px;padding:35px;border-radius:28px;box-shadow:0 25px 60px rgba(0,0,0,.25)}
            h1{color:#155f78}
            input,button{width:100%;padding:14px;margin:10px 0;border-radius:12px;border:1px solid #ddd}
            button{background:#155f78;color:white;font-weight:bold}
            .error{color:#b91c1c;font-weight:bold}
            a{color:#155f78}
        </style>
    </head>
    <body>
        <div class="card">
            <h1>Portal Docente</h1>
            <p>Asistencia en aula.</p>
            {% if error %}<p class="error">{{ error }}</p>{% endif %}
            <form method="POST">
                <input name="usuario" placeholder="Usuario docente" required>
                <input name="password" type="password" placeholder="Contraseña" required>
                <button>Ingresar</button>
            </form>
            <a href="/login">Volver</a>
        </div>
    </body>
    </html>
    """, error=error)


@app.route("/docente", methods=["GET", "POST"])
def docente():
    if not requiere_login() or rol_actual() != "Docente":
        return redirect("/docente-login")

    grupo_docente = session.get("grupo_docente", "")
    grupos = sorted({e.grado for e in Estudiante.query.all()}, key=lambda x: int(x) if str(x).isdigit() else 999)

    grupo = request.args.get("grupo") or grupo_docente or (grupos[0] if grupos else "")
    estudiantes = Estudiante.query.filter_by(grado=grupo).order_by(Estudiante.nombre.asc()).all()

    mensaje = ""

    if request.method == "POST":
        grupo = request.form.get("grupo", grupo)
        estudiantes = Estudiante.query.filter_by(grado=grupo).order_by(Estudiante.nombre.asc()).all()

        for estudiante in estudiantes:
            estado = request.form.get(f"estado_{estudiante.id}", "Presente")
            observacion = request.form.get(f"observacion_{estudiante.id}", "")

            asistencia = AsistenciaClase(
                estudiante_id=estudiante.id,
                docente=session["usuario"],
                grupo=grupo,
                fecha=fecha_hoy(),
                hora=hora_actual(),
                estado=estado,
                periodo=PERIODO_ACTUAL,
                observacion=observacion
            )

            db.session.add(asistencia)

            if estado == "Excusa":
                db.session.add(Excusa(
                    estudiante_id=estudiante.id,
                    fecha=fecha_hoy(),
                    motivo=observacion or "Excusa registrada por docente",
                    registrado_por=session["usuario"],
                    periodo=PERIODO_ACTUAL
                ))

        db.session.commit()
        mensaje = "Asistencia guardada correctamente."

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Asistencia Docente</title>
        <style>
            body{font-family:Segoe UI,Arial;background:#eefaf8;margin:0;padding:25px}
            .top,.card{background:white;border-radius:24px;padding:25px;margin-bottom:20px;box-shadow:0 15px 35px rgba(15,61,76,.12)}
            h1,h2{color:#155f78}
            select,input,button{padding:12px;border-radius:12px;border:1px solid #ddd}
            button{background:#155f78;color:white;font-weight:bold;border:none}
            table{width:100%;border-collapse:collapse;background:white}
            th{background:#155f78;color:white;padding:12px}
            td{padding:10px;border-bottom:1px solid #eee}
            .msg{background:#dcfce7;color:#047857;padding:12px;border-radius:12px;font-weight:bold}
            a{color:#155f78;font-weight:bold}
        </style>
    </head>
    <body>
        <section class="top">
            <h1>Portal Docente</h1>
            <p>Docente: <strong>{{ session["usuario"] }}</strong> · Periodo: <strong>{{ periodo }}</strong></p>
            <a href="/logout">Cerrar sesión</a>
        </section>

        <section class="card">
            {% if mensaje %}<p class="msg">{{ mensaje }}</p>{% endif %}

            <form method="GET">
                <label>Grupo:</label>
                <select name="grupo">
                    {% for g in grupos %}
                        <option value="{{ g }}" {% if g == grupo %}selected{% endif %}>{{ g }}</option>
                    {% endfor %}
                </select>
                <button>Ver grupo</button>
            </form>
        </section>

        <section class="card">
            <h2>Asistencia en aula - Grupo {{ grupo }}</h2>

            <form method="POST">
                <input type="hidden" name="grupo" value="{{ grupo }}">

                <table>
                    <tr>
                        <th>Estudiante</th>
                        <th>Estado</th>
                        <th>Observación / Excusa</th>
                    </tr>

                    {% for e in estudiantes %}
                    <tr>
                        <td>{{ e.nombre }} {{ e.apellido }}</td>
                        <td>
                            <select name="estado_{{ e.id }}">
                                <option value="Presente">Presente</option>
                                <option value="Ausente">Ausente</option>
                                <option value="Tarde">Tarde</option>
                                <option value="Excusa">Excusa</option>
                            </select>
                        </td>
                        <td>
                            <input name="observacion_{{ e.id }}" placeholder="Motivo si aplica">
                        </td>
                    </tr>
                    {% endfor %}
                </table>

                <br>
                <button>Guardar asistencia</button>
            </form>
        </section>
    </body>
    </html>
    """,
    grupos=grupos,
    grupo=grupo,
    estudiantes=estudiantes,
    mensaje=mensaje,
    periodo=PERIODO_ACTUAL)


# =========================
# ESTUDIANTES Y CARNÉS
# =========================

@app.route("/estudiantes", methods=["GET", "POST"])
def estudiantes():
    if not requiere_login():
        return redirect("/login")

    if not puede_gestionar_estudiantes():
        return "No tienes permiso para gestionar estudiantes."

    if request.method == "POST":
        codigo = request.form.get("codigo", "").strip()
        nombre = request.form.get("nombre", "").strip()
        apellido = request.form.get("apellido", "").strip()
        grado = request.form.get("grado", "").strip()
        director = request.form.get("director", "").strip()

        if codigo and nombre and apellido and grado:
            existente = Estudiante.query.filter_by(codigo=codigo).first()

            if existente:
                existente.nombre = nombre
                existente.apellido = apellido
                existente.grado = grado
                existente.director = director
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

    lista = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()
    return render_template("estudiantes.html", estudiantes=lista)


@app.route("/eliminar_estudiante/<int:id>")
def eliminar_estudiante(id):
    if not requiere_login():
        return redirect("/login")

    if not puede_gestionar_estudiantes():
        return "No tienes permiso."

    estudiante = Estudiante.query.get_or_404(id)

    IngresoPorteria.query.filter_by(estudiante_id=estudiante.id).delete()
    AsistenciaClase.query.filter_by(estudiante_id=estudiante.id).delete()
    Excusa.query.filter_by(estudiante_id=estudiante.id).delete()

    db.session.delete(estudiante)
    db.session.commit()

    return redirect("/estudiantes")


@app.route("/qr/<int:id>")
def qr_estudiante(id):
    estudiante = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(estudiante))
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return send_file(buffer, mimetype="image/png")


@app.route("/qr_descargar/<int:id>")
def qr_descargar(id):
    estudiante = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(estudiante))
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return send_file(buffer, mimetype="image/png", as_attachment=True, download_name=f"{estudiante.codigo}.png")


@app.route("/carnet/<int:id>")
def carnet(id):
    estudiante = Estudiante.query.get_or_404(id)

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Carné Estudiantil</title>
        <style>
            body{font-family:Segoe UI,Arial;background:#eefaf8;display:flex;justify-content:center;align-items:center;min-height:100vh}
            .carnet{width:340px;background:white;border-radius:24px;padding:25px;text-align:center;box-shadow:0 20px 50px rgba(0,0,0,.18)}
            .header{background:linear-gradient(135deg,#155f78,#2fb4c8);color:white;border-radius:20px;padding:18px}
            .logo{width:80px;height:80px;object-fit:contain;background:white;border-radius:18px;padding:7px}
            .qr{width:170px;margin:20px auto}
            h2{color:#155f78}
            button{background:#155f78;color:white;border:none;padding:12px 18px;border-radius:12px;font-weight:bold}
            @media print{button{display:none} body{background:white}}
        </style>
    </head>
    <body>
        <div>
            <div class="carnet">
                <div class="header">
                    <img class="logo" src="{{ url_for('static', filename='img/image.png') }}">
                    <h3>Institución Educativa Gabriel Correa Vélez</h3>
                    <p>Sede Principal</p>
                </div>

                <h2>{{ e.nombre }} {{ e.apellido }}</h2>
                <p><strong>Grado:</strong> {{ e.grado }}</p>
                <p><strong>Director:</strong> {{ e.director }}</p>
                <p><strong>Código:</strong> {{ e.codigo }}</p>

                <img class="qr" src="/qr/{{ e.id }}">

                <p>EduTrack QR</p>
            </div>
            <br>
            <button onclick="window.print()">Imprimir carné</button>
        </div>
    </body>
    </html>
    """, e=estudiante)


# =========================
# USUARIOS Y ROLES
# =========================

@app.route("/usuarios", methods=["GET", "POST"])
@app.route("/crear_usuario", methods=["GET", "POST"])
def usuarios():
    if not requiere_login():
        return redirect("/login")

    if not puede_todo() and rol_actual() != "Soporte":
        return "No tienes permiso para administrar usuarios."

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        password = request.form.get("password", "").strip()
        rol = request.form.get("rol", "").strip()
        correo = request.form.get("correo", "").strip()
        grupo_docente = request.form.get("grupo_docente", "").strip()

        if usuario and password and rol:
            if not Usuario.query.filter_by(usuario=usuario).first():
                db.session.add(Usuario(
                    usuario=usuario,
                    password=password,
                    rol=rol,
                    correo=correo,
                    grupo_docente=grupo_docente
                ))
                db.session.commit()

        return redirect("/usuarios")

    lista = Usuario.query.order_by(Usuario.rol.asc(), Usuario.usuario.asc()).all()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/eliminar_usuario/<int:id>")
def eliminar_usuario(id):
    if not requiere_login():
        return redirect("/login")

    if not puede_todo() and rol_actual() != "Soporte":
        return "No tienes permiso."

    usuario = Usuario.query.get_or_404(id)

    if usuario.usuario != session.get("usuario"):
        db.session.delete(usuario)
        db.session.commit()

    return redirect("/usuarios")


# =========================
# EXCUSAS, ALERTAS, HISTORIAL
# =========================

@app.route("/excusas")
def excusas():
    if not requiere_login():
        return redirect("/login")

    lista = Excusa.query.join(Estudiante).order_by(Excusa.fecha.desc()).all()

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head><meta charset="UTF-8"><title>Excusas</title></head>
    <body>
        <h1>Excusas registradas</h1>
        <a href="/dashboard">Volver</a>
        <table border="1" cellpadding="8">
            <tr><th>Fecha</th><th>Estudiante</th><th>Grado</th><th>Motivo</th><th>Registrado por</th></tr>
            {% for x in lista %}
            <tr>
                <td>{{ x.fecha }}</td>
                <td>{{ x.estudiante.nombre }} {{ x.estudiante.apellido }}</td>
                <td>{{ x.estudiante.grado }}</td>
                <td>{{ x.motivo }}</td>
                <td>{{ x.registrado_por }}</td>
            </tr>
            {% endfor %}
        </table>
    </body>
    </html>
    """, lista=lista)


@app.route("/alertas")
def alertas():
    if not requiere_login():
        return redirect("/login")

    estudiantes = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()
    alertas_lista = []

    for e in estudiantes:
        ingresos = IngresoPorteria.query.filter_by(estudiante_id=e.id, periodo=PERIODO_ACTUAL).all()
        aula = AsistenciaClase.query.filter_by(estudiante_id=e.id, periodo=PERIODO_ACTUAL).all()

        tardes = sum(1 for i in ingresos if i.estado == "Tarde") + sum(1 for a in aula if a.estado == "Tarde")
        ausencias = sum(1 for a in aula if a.estado == "Ausente")

        if tardes >= 3:
            alertas_lista.append((e, f"{tardes} llegadas tarde en {PERIODO_ACTUAL}"))

        if ausencias >= 3:
            alertas_lista.append((e, f"{ausencias} ausencias en aula en {PERIODO_ACTUAL}"))

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head><meta charset="UTF-8"><title>Alertas</title></head>
    <body>
        <h1>Alertas académicas</h1>
        <a href="/dashboard">Volver</a>
        <table border="1" cellpadding="8">
            <tr><th>Estudiante</th><th>Grado</th><th>Alerta</th></tr>
            {% for e, alerta in alertas %}
            <tr>
                <td>{{ e.nombre }} {{ e.apellido }}</td>
                <td>{{ e.grado }}</td>
                <td>{{ alerta }}</td>
            </tr>
            {% endfor %}
        </table>
    </body>
    </html>
    """, alertas=alertas_lista)


@app.route("/historial/<int:id>")
def historial_estudiante(id):
    if not requiere_login():
        return redirect("/login")

    estudiante = Estudiante.query.get_or_404(id)
    ingresos = IngresoPorteria.query.filter_by(estudiante_id=id).order_by(IngresoPorteria.fecha.desc()).all()
    aula = AsistenciaClase.query.filter_by(estudiante_id=id).order_by(AsistenciaClase.fecha.desc()).all()
    excusas_lista = Excusa.query.filter_by(estudiante_id=id).order_by(Excusa.fecha.desc()).all()

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="es">
    <head><meta charset="UTF-8"><title>Historial estudiante</title></head>
    <body>
        <h1>Historial de {{ e.nombre }} {{ e.apellido }}</h1>
        <p>Grado: {{ e.grado }} | Director: {{ e.director }}</p>
        <a href="/dashboard">Volver</a> | <a href="/carnet/{{ e.id }}">Ver carné</a>

        <h2>Ingresos de portería</h2>
        <table border="1" cellpadding="8">
            <tr><th>Fecha</th><th>Hora</th><th>Estado</th><th>Periodo</th></tr>
            {% for i in ingresos %}
            <tr><td>{{ i.fecha }}</td><td>{{ i.hora }}</td><td>{{ i.estado }}</td><td>{{ i.periodo }}</td></tr>
            {% endfor %}
        </table>

        <h2>Asistencia en aula</h2>
        <table border="1" cellpadding="8">
            <tr><th>Fecha</th><th>Hora</th><th>Docente</th><th>Estado</th><th>Observación</th></tr>
            {% for a in aula %}
            <tr><td>{{ a.fecha }}</td><td>{{ a.hora }}</td><td>{{ a.docente }}</td><td>{{ a.estado }}</td><td>{{ a.observacion }}</td></tr>
            {% endfor %}
        </table>

        <h2>Excusas</h2>
        <table border="1" cellpadding="8">
            <tr><th>Fecha</th><th>Motivo</th><th>Registrado por</th></tr>
            {% for x in excusas %}
            <tr><td>{{ x.fecha }}</td><td>{{ x.motivo }}</td><td>{{ x.registrado_por }}</td></tr>
            {% endfor %}
        </table>
    </body>
    </html>
    """, e=estudiante, ingresos=ingresos, aula=aula, excusas=excusas_lista)


# =========================
# REPORTES
# =========================

@app.route("/reportes")
def reportes():
    if not requiere_login():
        return redirect("/login")

    if not puede_reportes():
        return "No tienes permiso para reportes."

    hoy = fecha_hoy()

    ingresos = (
        IngresoPorteria.query.join(Estudiante)
        .order_by(Estudiante.grado.asc(), Estudiante.nombre.asc(), IngresoPorteria.fecha.desc())
        .all()
    )

    ids_hoy = {i.estudiante_id for i in IngresoPorteria.query.filter_by(fecha=hoy).all()}
    no_llegaron = (
        Estudiante.query.filter(~Estudiante.id.in_(ids_hoy))
        .order_by(Estudiante.grado.asc(), Estudiante.nombre.asc())
        .all()
    )

    estudiantes = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()

    estadisticas_estudiantes = []
    for e in estudiantes:
        ingresos_e = IngresoPorteria.query.filter_by(estudiante_id=e.id).all()
        aula_e = AsistenciaClase.query.filter_by(estudiante_id=e.id).all()

        estadisticas_estudiantes.append((
            e.grado,
            e.nombre,
            e.apellido,
            sum(1 for i in ingresos_e if i.estado == "Temprano"),
            sum(1 for i in ingresos_e if i.estado == "Tarde") + sum(1 for a in aula_e if a.estado == "Tarde"),
            sum(1 for a in aula_e if a.estado == "Ausente"),
            len(ingresos_e) + len(aula_e)
        ))

    grados = sorted({e.grado for e in estudiantes}, key=lambda x: int(x) if str(x).isdigit() else 999)

    estadisticas_periodos = []
    estadisticas_mensuales = []

    for grado in grados:
        ingresos_grado = IngresoPorteria.query.join(Estudiante).filter(Estudiante.grado == grado).all()

        for periodo in ["Periodo 1", "Periodo 2", "Periodo 3"]:
            datos = [i for i in ingresos_grado if i.periodo == periodo]
            if datos:
                estadisticas_periodos.append((
                    grado,
                    periodo,
                    sum(1 for i in datos if i.estado == "Temprano"),
                    sum(1 for i in datos if i.estado == "Tarde"),
                    sum(1 for i in datos if i.estado == "No llegó"),
                    len(datos)
                ))

        meses = sorted({i.fecha[:7] for i in ingresos_grado})
        for mes in meses:
            datos = [i for i in ingresos_grado if i.fecha[:7] == mes]
            estadisticas_mensuales.append((
                grado,
                mes,
                sum(1 for i in datos if i.estado == "Temprano"),
                sum(1 for i in datos if i.estado == "Tarde"),
                sum(1 for i in datos if i.estado == "No llegó"),
                len(datos)
            ))

    return render_template(
        "reportes.html",
        ingresos=ingresos,
        no_llegaron=no_llegaron,
        estadisticas_estudiantes=estadisticas_estudiantes,
        estadisticas_periodos=estadisticas_periodos,
        estadisticas_mensuales=estadisticas_mensuales,
        hoy=hoy
    )


# =========================
# EXPORTACIONES
# =========================

@app.route("/exportar_estudiantes")
def exportar_estudiantes():
    if not requiere_login():
        return redirect("/login")

    estudiantes = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Estudiantes"
    ws.append(["Código", "Nombres", "Apellidos", "Grado", "Director"])

    for e in estudiantes:
        ws.append([e.codigo, e.nombre, e.apellido, e.grado, e.director])

    archivo = BytesIO()
    wb.save(archivo)
    archivo.seek(0)

    return send_file(archivo, as_attachment=True, download_name="estudiantes_edutrack.xlsx")


@app.route("/exportar_excel_reportes")
def exportar_excel_reportes():
    if not requiere_login():
        return redirect("/login")

    ingresos = IngresoPorteria.query.join(Estudiante).order_by(Estudiante.grado.asc(), IngresoPorteria.fecha.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Reportes"
    ws.append(["Código", "Nombre", "Apellido", "Grado", "Fecha", "Hora", "Estado", "Periodo", "Registrado por"])

    for i in ingresos:
        ws.append([
            i.estudiante.codigo,
            i.estudiante.nombre,
            i.estudiante.apellido,
            i.estudiante.grado,
            i.fecha,
            i.hora,
            i.estado,
            i.periodo,
            i.registrado_por
        ])

    archivo = BytesIO()
    wb.save(archivo)
    archivo.seek(0)

    return send_file(archivo, as_attachment=True, download_name="reportes_edutrack.xlsx")


@app.route("/exportar_pdf")
def exportar_pdf():
    if not requiere_login():
        return redirect("/login")

    ingresos = IngresoPorteria.query.join(Estudiante).order_by(Estudiante.grado.asc(), IngresoPorteria.fecha.desc()).all()

    archivo = BytesIO()
    pdf = canvas.Canvas(archivo, pagesize=letter)
    pdf.drawString(50, 750, "Reporte de ingresos - EduTrack QR")

    y = 720
    for i in ingresos:
        texto = f"{i.estudiante.nombre} {i.estudiante.apellido} | Grado {i.estudiante.grado} | {i.fecha} | {i.hora} | {i.estado} | {i.periodo}"
        pdf.drawString(50, y, texto[:115])
        y -= 20

        if y < 50:
            pdf.showPage()
            y = 750

    pdf.save()
    archivo.seek(0)

    return send_file(archivo, as_attachment=True, download_name="reporte_ingresos.pdf")


@app.route("/exportar_word")
def exportar_word():
    if not requiere_login():
        return redirect("/login")

    ingresos = IngresoPorteria.query.join(Estudiante).order_by(Estudiante.grado.asc(), IngresoPorteria.fecha.desc()).all()

    doc = Document()
    doc.add_heading("Reporte de ingresos - EduTrack QR", 0)

    for i in ingresos:
        doc.add_paragraph(
            f"{i.estudiante.nombre} {i.estudiante.apellido} | "
            f"Grado {i.estudiante.grado} | {i.fecha} | {i.hora} | {i.estado} | {i.periodo}"
        )

    archivo = BytesIO()
    doc.save(archivo)
    archivo.seek(0)

    return send_file(archivo, as_attachment=True, download_name="reporte_ingresos.docx")


# =========================
# RECUPERACIÓN
# =========================

@app.route("/recuperar", methods=["GET", "POST"])
def recuperar():
    mensaje = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        correo = request.form.get("correo", "")

        user = Usuario.query.filter_by(usuario=usuario, correo=correo).first()

        if user:
            pin = str(random.randint(100000, 999999))
            Recuperacion.query.filter_by(usuario=usuario).delete()

            db.session.add(Recuperacion(usuario=usuario, pin=pin))
            db.session.commit()

            if enviar_pin(correo, pin):
                session["recuperar_usuario"] = usuario
                return redirect("/validar_pin")

            mensaje = "No se pudo enviar el correo. Revisa las variables de soporte."
        else:
            mensaje = "Usuario o correo no encontrado."

    return render_template("recuperar.html", mensaje=mensaje)


@app.route("/validar_pin", methods=["GET", "POST"])
def validar_pin():
    if "recuperar_usuario" not in session:
        return redirect("/recuperar")

    mensaje = ""

    if request.method == "POST":
        pin = request.form.get("pin", "")
        nueva = request.form.get("password", "")
        usuario = session["recuperar_usuario"]

        existe = Recuperacion.query.filter_by(usuario=usuario, pin=pin).first()

        if existe:
            user = Usuario.query.filter_by(usuario=usuario).first()
            user.password = nueva

            Recuperacion.query.filter_by(usuario=usuario).delete()
            db.session.commit()

            session.pop("recuperar_usuario", None)
            return redirect("/login")

        mensaje = "PIN incorrecto."

    return render_template("validar_pin.html", mensaje=mensaje)


# =========================
# LEGAL Y COOKIES
# =========================

@app.route("/legal")
def legal():
    return render_template("legal.html")


@app.route("/cookies")
def cookies():
    return render_template("cookies.html")


if __name__ == "__main__":
    with app.app_context():
        inicializar_bd()

    app.run(debug=True, host="0.0.0.0")
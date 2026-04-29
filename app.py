from datetime import datetime, date
from io import BytesIO
import os, random, smtplib
from email.message import EmailMessage

import qrcode
from docx import Document
from flask import Flask, request, redirect, session, send_file, render_template_string
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from sqlalchemy import func

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'edutrack_local_secret')

DATABASE_URL = os.environ.get('DATABASE_URL', 'sqlite:///edutrack.db')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

PERIODO_ACTUAL = os.environ.get('PERIODO_ACTUAL', 'Periodo 2')

CSS = '''
<style>
:root{--v:#14532d;--v2:#1f7a3a;--r:#b91c1c;--a:#facc15;--n:#111827;--s:0 22px 55px rgba(17,24,39,.15)}
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:linear-gradient(135deg,#f8fafc,#eaf7ef);color:var(--n)}
a{color:var(--v);font-weight:700}
input,select,textarea{width:100%;padding:14px;border:1px solid #d1d5db;border-radius:14px;margin:8px 0;font-size:15px}
button,.btn{background:linear-gradient(135deg,var(--v),var(--v2));color:white;border:0;border-radius:14px;padding:13px 18px;font-weight:800;text-decoration:none;display:inline-block;cursor:pointer}
.btn-red{background:var(--r)!important}
.center{display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px;background:radial-gradient(circle at top left,rgba(250,204,21,.25),transparent 28%),radial-gradient(circle at bottom right,rgba(185,28,28,.22),transparent 28%),linear-gradient(135deg,var(--v),var(--v2))}
.card{background:white;border-radius:28px;padding:30px;box-shadow:var(--s);border-top:8px solid var(--a)}
.login-card{width:100%;max-width:430px}
.logo{width:96px;height:96px;object-fit:contain;background:white;border-radius:22px;padding:8px}
.hero-logo{background:linear-gradient(135deg,var(--v),#22c55e);color:white;border-radius:24px;padding:25px;text-align:center;border-bottom:8px solid var(--a)}
.hero-logo h2,.hero-logo p{color:white}
.error{color:var(--r);font-weight:800}
.msg{padding:13px;border-radius:14px;font-weight:800;margin:12px 0}
.ok{background:#dcfce7;color:#047857}
.warn{background:#fef3c7;color:#b45309}
.danger{background:#fee2e2;color:#b91c1c}
.layout{display:flex;min-height:100vh}
.sidebar{width:285px;background:linear-gradient(180deg,var(--v),#0f3d2e);color:white;padding:28px;border-right:8px solid var(--a)}
.sidebar img{width:100px;height:100px;object-fit:contain;background:white;border-radius:22px;padding:8px}
.sidebar a{display:block;color:white;text-decoration:none;background:rgba(255,255,255,.12);padding:14px 16px;border-radius:15px;margin:10px 0}
.sidebar a:hover{background:var(--r)}
.main{flex:1;padding:32px;overflow:auto}
.top{background:white;border-radius:26px;padding:26px 30px;display:flex;justify-content:space-between;gap:18px;align-items:center;box-shadow:var(--s);border-top:7px solid var(--a);margin-bottom:26px}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:22px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
.stat{background:#f8fafc;border-left:7px solid var(--v2);padding:20px;border-radius:20px}
.stat h3{font-size:34px;margin:0}
.table-card{background:white;border-radius:26px;padding:25px;box-shadow:var(--s);margin-top:24px}
table{width:100%;border-collapse:collapse;margin-top:15px}
th{background:linear-gradient(135deg,var(--v),var(--v2));color:white;padding:13px}
td{padding:12px;border-bottom:1px solid #e5e7eb;text-align:center}
.qr-img{width:88px}
.danger-link{color:var(--r)}
.portal{width:100%;max-width:460px;text-align:center}
.portal #reader{max-width:330px;margin:16px auto;border-radius:18px;overflow:hidden}
.carnet{width:360px;background:white;border-radius:28px;padding:24px;text-align:center;box-shadow:var(--s);border-top:8px solid var(--a)}
.carnet-head{background:linear-gradient(135deg,var(--v),var(--v2));color:white;border-radius:22px;padding:18px;border-bottom:6px solid var(--a)}
.carnet-head h3{color:white}
.carnet .qr{width:180px;margin:18px auto}
.print-wrap{display:flex;align-items:center;justify-content:center;min-height:100vh;background:#eaf7ef;flex-direction:column}
@media(max-width:900px){
.layout{flex-direction:column}
.sidebar{width:100%;border-right:0;border-bottom:8px solid var(--a)}
.grid,.stats{grid-template-columns:1fr}
.top{flex-direction:column;align-items:flex-start}
.main{padding:18px}
}
@media print{
.no-print{display:none}
.print-wrap{background:white}
.carnet{box-shadow:none}
}
</style>
'''

def page(title, body):
    return f"<!doctype html><html lang='es'><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{title}</title>{CSS}</head><body>{body}</body></html>"

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
    correo = db.Column(db.String(160), default='')
    grupo_docente = db.Column(db.String(30), default='')

class IngresoPorteria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    hora = db.Column(db.String(20), nullable=False)
    dia = db.Column(db.String(30), nullable=False)
    estado = db.Column(db.String(30), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    registrado_por = db.Column(db.String(100), default='Portal móvil')
    estudiante = db.relationship('Estudiante')

class AsistenciaClase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    docente = db.Column(db.String(120), nullable=False)
    grupo = db.Column(db.String(30), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    hora = db.Column(db.String(20), nullable=False)
    estado = db.Column(db.String(30), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    observacion = db.Column(db.Text, default='')
    estudiante = db.relationship('Estudiante')

class Excusa(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    estudiante_id = db.Column(db.Integer, db.ForeignKey('estudiante.id'), nullable=False)
    fecha = db.Column(db.String(20), nullable=False)
    motivo = db.Column(db.Text, nullable=False)
    registrado_por = db.Column(db.String(120), nullable=False)
    periodo = db.Column(db.String(30), nullable=False)
    estudiante = db.relationship('Estudiante')

class Recuperacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(80), nullable=False)
    pin = db.Column(db.String(10), nullable=False)

def inicializar_bd():
    db.create_all()

    usuarios_base = [
        ('admin', '1234', 'Rectoría', '', ''),
        ('soporte', '1234', 'Soporte', '', ''),
        ('docente', '1234', 'Docente', '', '8')
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

def fecha_hoy():
    return date.today().strftime('%Y-%m-%d')

def hora_actual():
    return datetime.now().strftime('%H:%M:%S')

def limpiar_codigo(texto):
    texto = (texto or '').strip()
    if 'Codigo:' in texto:
        return texto.split('Codigo:')[1].split('\n')[0].strip()
    return texto

def qr_texto(e):
    return f"Codigo: {e.codigo}\nNombres: {e.nombre}\nApellidos: {e.apellido}\nGrado: {e.grado}\nDirector: {e.director}"

def login_usuario(usuario, password, rol_requerido=None):
    usuario = (usuario or '').strip()
    password = (password or '').strip()

    user = Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower()).first()

    if not user:
        return None

    if user.password.strip() != password:
        return None

    if rol_requerido and user.rol.lower().strip() != rol_requerido.lower().strip():
        return None

    return user

def registrar_ingreso_porteria(codigo, estado, registrado_por='Portal móvil'):
    codigo = limpiar_codigo(codigo)
    estudiante = Estudiante.query.filter_by(codigo=codigo).first()

    if not estudiante:
        return 'Estudiante no registrado', 'No registrado'

    db.session.add(IngresoPorteria(
        estudiante_id=estudiante.id,
        fecha=fecha_hoy(),
        hora=hora_actual(),
        dia=datetime.now().strftime('%A'),
        estado=estado,
        periodo=PERIODO_ACTUAL,
        registrado_por=registrado_por
    ))

    db.session.commit()

    return f'Registro guardado: {estudiante.nombre} {estudiante.apellido} - {estado}', estado

def requiere_login():
    return 'usuario' in session

def rol_actual():
    return session.get('rol', '')

def puede_todo():
    return rol_actual() in ['Rectoría', 'Administrador', 'Soporte']

def puede_gestionar_estudiantes():
    return rol_actual() in ['Rectoría', 'Coordinación', 'Secretaría', 'Administrador', 'Soporte']

def shell(content):
    return f"""
    <div class="layout">
      <aside class="sidebar">
        <img src="/static/img/logo-colegio.png" alt="Escudo">
        <h2>EduTrack QR</h2>
        <p>Gabriel Correa Vélez</p>
        <a href="/dashboard">Inicio</a>
        <a href="/portal" target="_blank">Portal móvil</a>
        <a href="/docente-login">Portal docente</a>
        <a href="/estudiantes">Estudiantes</a>
        <a href="/usuarios">Usuarios</a>
        <a href="/reportes">Reportes</a>
        <a href="/alertas">Alertas</a>
        <a href="/excusas">Excusas</a>
        <a href="/soporte-login">Soporte</a>
      </aside>
      <main class="main">{content}</main>
    </div>
    """

@app.route('/')
def inicio():
    return redirect('/login')

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = ''

    if request.method == 'POST':
        user = login_usuario(request.form.get('usuario'), request.form.get('password'))

        if user:
            session['usuario'] = user.usuario
            session['rol'] = user.rol
            session['grupo_docente'] = user.grupo_docente or ''
            return redirect('/dashboard')

        error = 'Usuario o contraseña incorrectos.'

    body = f"""
    <div class="center">
        <section class="card login-card">
            <div class="hero-logo">
                <img class="logo" src="/static/img/logo-colegio.png">
                <h2>EduTrack QR</h2>
                <p>Institución Educativa Gabriel Correa Vélez</p>
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
                <a href="/soporte-login">Soporte</a>
            </p>
        </section>
    </div>
    """

    return page('Login - EduTrack QR', body)

@app.route('/dashboard')
def dashboard():
    if not requiere_login():
        return redirect('/login')

    hoy = fecha_hoy()
    estudiantes = Estudiante.query.all()
    ingresos_hoy = IngresoPorteria.query.filter_by(fecha=hoy).all()

    total = len(estudiantes)
    temprano = sum(1 for i in ingresos_hoy if i.estado == 'Temprano')
    tarde = sum(1 for i in ingresos_hoy if i.estado == 'Tarde')
    no_manual = sum(1 for i in ingresos_hoy if i.estado == 'No llegó')
    ids = {i.estudiante_id for i in ingresos_hoy}
    no_llego = no_manual + max(total - len(ids), 0)

    grados = sorted({e.grado for e in estudiantes}, key=lambda x: int(x) if str(x).isdigit() else 999)

    filas_grupo = ''
    for g in grados:
        est_g = [e for e in estudiantes if e.grado == g]
        ing_g = [i for i in ingresos_hoy if i.estudiante.grado == g]
        ids_g = {i.estudiante_id for i in ing_g}

        filas_grupo += f"""
        <tr>
            <td>{g}</td>
            <td>{len(est_g)}</td>
            <td>{len(ing_g)}</td>
            <td>{sum(1 for i in ing_g if i.estado == 'Temprano')}</td>
            <td>{sum(1 for i in ing_g if i.estado == 'Tarde')}</td>
            <td>{len(est_g) - len(ids_g)}</td>
        </tr>
        """

    ultimos = (
        IngresoPorteria.query.join(Estudiante)
        .filter(IngresoPorteria.fecha == hoy)
        .order_by(IngresoPorteria.hora.desc())
        .limit(10)
        .all()
    )

    filas_ultimos = ''.join(
        f"""
        <tr>
            <td>{i.estudiante.grado}</td>
            <td>{i.estudiante.nombre} {i.estudiante.apellido}</td>
            <td>{i.hora}</td>
            <td>{i.estado}</td>
        </tr>
        """
        for i in ultimos
    )

    content = f"""
    <header class="top">
        <div>
            <h1>Panel Administrativo</h1>
            <p>Periodo: <b>{PERIODO_ACTUAL}</b> · Usuario: <b>{session['usuario']}</b></p>
        </div>
        <a class="btn btn-red" href="/logout">Salir</a>
    </header>

    <section class="grid">
        <div class="hero-logo">
            <img class="logo" src="/static/img/logo-colegio.png">
            <h2>EduTrack QR</h2>
            <p>Sistema institucional de asistencia</p>
        </div>

        <div class="card">
            <h2>Resumen de ingreso - {hoy}</h2>

            <div class="stats">
                <div class="stat"><h3>{total}</h3><p>Estudiantes</p></div>
                <div class="stat"><h3>{temprano}</h3><p>Temprano</p></div>
                <div class="stat"><h3>{tarde}</h3><p>Tarde</p></div>
                <div class="stat"><h3>{no_llego}</h3><p>No llegó</p></div>
            </div>
        </div>
    </section>

    <section class="grid" style="margin-top:22px">
        <a class="card" href="/portal"><h2>Portal móvil</h2><p>Cámara QR y registro manual.</p></a>
        <a class="card" href="/docente-login"><h2>Portal docente</h2><p>Asistencia en aula por grupo.</p></a>
        <a class="card" href="/estudiantes"><h2>Estudiantes y carnés</h2><p>Crear QR, carné e historial.</p></a>
        <a class="card" href="/reportes"><h2>Reportes</h2><p>Excel, PDF, Word, periodo y mes.</p></a>
    </section>

    <section class="table-card">
        <h2>Resumen por grupo</h2>

        <table>
            <tr>
                <th>Grupo</th>
                <th>Total</th>
                <th>Registrados</th>
                <th>Temprano</th>
                <th>Tarde</th>
                <th>No llegó</th>
            </tr>
            {filas_grupo}
        </table>
    </section>

    <section class="table-card">
        <h2>Últimos ingresos</h2>

        <table>
            <tr>
                <th>Grupo</th>
                <th>Estudiante</th>
                <th>Hora</th>
                <th>Estado</th>
            </tr>
            {filas_ultimos}
        </table>
    </section>
    """

    return page('Dashboard', shell(content))

@app.route('/portal', methods=['GET', 'POST'])
def portal():
    mensaje = ''
    estado = ''

    if request.method == 'POST':
        mensaje, estado = registrar_ingreso_porteria(
            request.form.get('codigo'),
            request.form.get('estado', 'Temprano'),
            'Portal móvil'
        )

    if mensaje and estado != 'No registrado':
        mensaje_html = f'<div class="msg ok">{mensaje}</div>'
    elif mensaje:
        mensaje_html = f'<div class="msg danger">{mensaje}</div>'
    else:
        mensaje_html = ''

    body = f"""
    <div class="center">
        <section class="card portal">
            <img class="portal-logo logo" src="/static/img/logo-colegio.png">
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

            <p><a href="/login">Administración</a></p>
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

    return page('Portal móvil', body)

@app.route('/docente-login', methods=['GET', 'POST'])
def docente_login():
    error = ''

    if request.method == 'POST':
        user = login_usuario(request.form.get('usuario'), request.form.get('password'), 'Docente')

        if user:
            session['usuario'] = user.usuario
            session['rol'] = user.rol
            session['grupo_docente'] = user.grupo_docente or ''
            return redirect('/docente')

        error = 'Usuario docente incorrecto. Revisa que esté creado con rol Docente.'

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
            <a href="/login">Volver</a>
        </section>
    </div>
    """

    return page('Portal docente', body)

@app.route('/docente', methods=['GET', 'POST'])
def docente():
    if not requiere_login() or rol_actual() != 'Docente':
        return redirect('/docente-login')

    grupos = sorted({e.grado for e in Estudiante.query.all()}, key=lambda x: int(x) if str(x).isdigit() else 999)
    grupo = request.args.get('grupo') or session.get('grupo_docente') or (grupos[0] if grupos else '')

    mensaje = ''

    if request.method == 'POST':
        grupo = request.form.get('grupo', grupo)
        estudiantes_grupo = Estudiante.query.filter_by(grado=grupo).all()

        for e in estudiantes_grupo:
            estado = request.form.get(f'estado_{e.id}', 'Presente')
            obs = request.form.get(f'observacion_{e.id}', '')

            db.session.add(AsistenciaClase(
                estudiante_id=e.id,
                docente=session['usuario'],
                grupo=grupo,
                fecha=fecha_hoy(),
                hora=hora_actual(),
                estado=estado,
                periodo=PERIODO_ACTUAL,
                observacion=obs
            ))

            if estado == 'Excusa':
                db.session.add(Excusa(
                    estudiante_id=e.id,
                    fecha=fecha_hoy(),
                    motivo=obs or 'Excusa registrada por docente',
                    registrado_por=session['usuario'],
                    periodo=PERIODO_ACTUAL
                ))

        db.session.commit()
        mensaje = 'Asistencia guardada correctamente.'

    estudiantes_grupo = Estudiante.query.filter_by(grado=grupo).order_by(Estudiante.nombre.asc()).all()

    opciones_grupo = ''.join(
        f'<option {"selected" if g == grupo else ""}>{g}</option>'
        for g in grupos
    )

    filas = ''.join(
        f"""
        <tr>
            <td>{e.nombre} {e.apellido}</td>
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
    <header class="top">
        <div>
            <h1>Portal Docente</h1>
            <p>Docente: <b>{session['usuario']}</b> · Periodo: <b>{PERIODO_ACTUAL}</b></p>
        </div>
        <a class="btn btn-red" href="/logout">Salir</a>
    </header>

    {'<div class="msg ok">' + mensaje + '</div>' if mensaje else ''}

    <section class="card">
        <form method="GET">
            <label>Grupo</label>
            <select name="grupo">{opciones_grupo}</select>
            <button>Ver grupo</button>
        </form>
    </section>

    <section class="table-card">
        <h2>Asistencia en aula - Grupo {grupo}</h2>

        <form method="POST">
            <input type="hidden" name="grupo" value="{grupo}">

            <table>
                <tr>
                    <th>Estudiante</th>
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

    return page('Portal Docente', shell(content))

@app.route('/estudiantes', methods=['GET', 'POST'])
def estudiantes():
    if not requiere_login():
        return redirect('/login')

    if not puede_gestionar_estudiantes():
        return 'No tienes permiso.'

    if request.method == 'POST':
        codigo = request.form.get('codigo', '').strip()

        if codigo:
            e = Estudiante.query.filter_by(codigo=codigo).first()

            nombre = request.form.get('nombre', '').strip()
            apellido = request.form.get('apellido', '').strip()
            grado = request.form.get('grado', '').strip()
            director = request.form.get('director', '').strip()

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

        return redirect('/estudiantes')

    estudiantes_lista = Estudiante.query.order_by(Estudiante.grado.asc(), Estudiante.nombre.asc()).all()

    filas = ''.join(
        f"""
        <tr>
            <td>{e.codigo}</td>
            <td>{e.nombre} {e.apellido}</td>
            <td>{e.grado}</td>
            <td>{e.director}</td>
            <td>
                <img class="qr-img" src="/qr/{e.id}">
                <br>
                <a href="/qr_descargar/{e.id}">Descargar</a>
            </td>
            <td><a href="/carnet/{e.id}">Carné</a></td>
            <td><a href="/historial/{e.id}">Historial</a></td>
            <td><a class="danger-link" href="/eliminar_estudiante/{e.id}">Eliminar</a></td>
        </tr>
        """
        for e in estudiantes_lista
    )

    content = f"""
    <header class="top">
        <div>
            <h1>Estudiantes y QR</h1>
            <p>Carné institucional, QR e historial.</p>
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
                <th>Estudiante</th>
                <th>Grado</th>
                <th>Director</th>
                <th>QR</th>
                <th>Carné</th>
                <th>Historial</th>
                <th>Acción</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page('Estudiantes', shell(content))

@app.route('/eliminar_estudiante/<int:id>')
def eliminar_estudiante(id):
    if not requiere_login():
        return redirect('/login')

    e = Estudiante.query.get_or_404(id)

    IngresoPorteria.query.filter_by(estudiante_id=e.id).delete()
    AsistenciaClase.query.filter_by(estudiante_id=e.id).delete()
    Excusa.query.filter_by(estudiante_id=e.id).delete()

    db.session.delete(e)
    db.session.commit()

    return redirect('/estudiantes')

@app.route('/qr/<int:id>')
def qr_estudiante(id):
    e = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(e))
    b = BytesIO()
    img.save(b, format='PNG')
    b.seek(0)
    return send_file(b, mimetype='image/png')

@app.route('/qr_descargar/<int:id>')
def qr_descargar(id):
    e = Estudiante.query.get_or_404(id)
    img = qrcode.make(qr_texto(e))
    b = BytesIO()
    img.save(b, format='PNG')
    b.seek(0)
    return send_file(b, mimetype='image/png', as_attachment=True, download_name=f'{e.codigo}.png')

@app.route('/carnet/<int:id>')
def carnet(id):
    e = Estudiante.query.get_or_404(id)

    body = f"""
    <div class="print-wrap">
        <div class="carnet">
            <div class="carnet-head">
                <img class="logo" src="/static/img/logo-colegio.png">
                <h3>Institución Educativa Gabriel Correa Vélez</h3>
                <p>Sede Principal</p>
            </div>

            <h2>{e.nombre} {e.apellido}</h2>
            <p><b>Grado:</b> {e.grado}</p>
            <p><b>Director:</b> {e.director}</p>
            <p><b>Código:</b> {e.codigo}</p>

            <img class="qr" src="/qr/{e.id}">

            <p>EduTrack QR</p>
        </div>

        <br>
        <button class="no-print" onclick="window.print()">Imprimir carné</button>
        <br>
        <a class="no-print" href="/estudiantes">Volver</a>
    </div>
    """

    return page('Carné', body)

@app.route('/historial/<int:id>')
def historial_estudiante(id):
    if not requiere_login():
        return redirect('/login')

    e = Estudiante.query.get_or_404(id)

    ingresos = IngresoPorteria.query.filter_by(estudiante_id=id).order_by(IngresoPorteria.fecha.desc(), IngresoPorteria.hora.desc()).all()
    aula = AsistenciaClase.query.filter_by(estudiante_id=id).order_by(AsistenciaClase.fecha.desc()).all()
    excusas = Excusa.query.filter_by(estudiante_id=id).order_by(Excusa.fecha.desc()).all()

    filas_ingresos = ''.join(
        f'<tr><td>{i.fecha}</td><td>{i.hora}</td><td>{i.estado}</td><td>{i.periodo}</td></tr>'
        for i in ingresos
    )

    filas_aula = ''.join(
        f'<tr><td>{a.fecha}</td><td>{a.hora}</td><td>{a.docente}</td><td>{a.estado}</td><td>{a.observacion}</td></tr>'
        for a in aula
    )

    filas_excusas = ''.join(
        f'<tr><td>{x.fecha}</td><td>{x.motivo}</td><td>{x.registrado_por}</td></tr>'
        for x in excusas
    )

    content = f"""
    <header class="top">
        <div>
            <h1>Historial de {e.nombre} {e.apellido}</h1>
            <p>Grado {e.grado} · Director {e.director}</p>
        </div>
        <a class="btn" href="/carnet/{e.id}">Ver carné</a>
    </header>

    <section class="table-card">
        <h2>Ingresos de portería</h2>
        <table>
            <tr><th>Fecha</th><th>Hora</th><th>Estado</th><th>Periodo</th></tr>
            {filas_ingresos}
        </table>
    </section>

    <section class="table-card">
        <h2>Asistencia en aula</h2>
        <table>
            <tr><th>Fecha</th><th>Hora</th><th>Docente</th><th>Estado</th><th>Observación</th></tr>
            {filas_aula}
        </table>
    </section>

    <section class="table-card">
        <h2>Excusas</h2>
        <table>
            <tr><th>Fecha</th><th>Motivo</th><th>Registrado por</th></tr>
            {filas_excusas}
        </table>
    </section>
    """

    return page('Historial', shell(content))

@app.route('/usuarios', methods=['GET', 'POST'])
@app.route('/crear_usuario', methods=['GET', 'POST'])
def usuarios():
    if not requiere_login():
        return redirect('/login')

    if not puede_todo():
        return 'No tienes permiso.'

    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip()

        if usuario and not Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower()).first():
            db.session.add(Usuario(
                usuario=usuario,
                password=request.form.get('password', '').strip(),
                rol=request.form.get('rol', '').strip(),
                correo=request.form.get('correo', '').strip(),
                grupo_docente=request.form.get('grupo_docente', '').strip()
            ))

            db.session.commit()

        return redirect('/usuarios')

    usuarios_lista = Usuario.query.order_by(Usuario.rol.asc(), Usuario.usuario.asc()).all()

    filas = ''.join(
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
    <header class="top">
        <div>
            <h1>Usuarios</h1>
            <p>Crea rectoría, coordinación, secretaría, docentes y soporte.</p>
        </div>
        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="card">
        <h2>Crear usuario</h2>

        <form method="POST">
            <input name="usuario" placeholder="Usuario" required>
            <input name="password" type="password" placeholder="Contraseña" required>
            <input name="correo" placeholder="Correo">

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

    return page('Usuarios', shell(content))

@app.route('/eliminar_usuario/<int:id>')
def eliminar_usuario(id):
    if not requiere_login():
        return redirect('/login')

    u = Usuario.query.get_or_404(id)

    if u.usuario != session.get('usuario'):
        db.session.delete(u)
        db.session.commit()

    return redirect('/usuarios')

@app.route('/reportes')
def reportes():
    if not requiere_login():
        return redirect('/login')

    ingresos = IngresoPorteria.query.join(Estudiante).order_by(Estudiante.grado.asc(), IngresoPorteria.fecha.desc()).all()

    filas = ''.join(
        f"""
        <tr>
            <td>{i.estudiante.grado}</td>
            <td>{i.estudiante.nombre} {i.estudiante.apellido}</td>
            <td>{i.fecha}</td>
            <td>{i.hora}</td>
            <td>{i.estado}</td>
            <td>{i.periodo}</td>
        </tr>
        """
        for i in ingresos
    )

    content = f"""
    <header class="top">
        <div>
            <h1>Reportes</h1>
            <p>Información profesional por estudiante, fecha, grupo y periodo.</p>
        </div>
        <div>
            <a class="btn" href="/exportar_excel_reportes">Excel</a>
            <a class="btn" href="/exportar_pdf">PDF</a>
            <a class="btn" href="/exportar_word">Word</a>
        </div>
    </header>

    <section class="table-card">
        <h2>Ingresos registrados</h2>

        <table>
            <tr>
                <th>Grupo</th>
                <th>Estudiante</th>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Estado</th>
                <th>Periodo</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page('Reportes', shell(content))

@app.route('/excusas')
def excusas():
    if not requiere_login():
        return redirect('/login')

    filas = ''.join(
        f"""
        <tr>
            <td>{x.fecha}</td>
            <td>{x.estudiante.nombre} {x.estudiante.apellido}</td>
            <td>{x.estudiante.grado}</td>
            <td>{x.motivo}</td>
            <td>{x.registrado_por}</td>
        </tr>
        """
        for x in Excusa.query.order_by(Excusa.fecha.desc()).all()
    )

    content = f"""
    <header class="top">
        <h1>Excusas</h1>
        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="table-card">
        <table>
            <tr>
                <th>Fecha</th>
                <th>Estudiante</th>
                <th>Grado</th>
                <th>Motivo</th>
                <th>Registró</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page('Excusas', shell(content))

@app.route('/alertas')
def alertas():
    if not requiere_login():
        return redirect('/login')

    filas = ''

    for e in Estudiante.query.all():
        ingresos = IngresoPorteria.query.filter_by(estudiante_id=e.id, periodo=PERIODO_ACTUAL).all()
        aula = AsistenciaClase.query.filter_by(estudiante_id=e.id, periodo=PERIODO_ACTUAL).all()

        tardes = sum(1 for i in ingresos if i.estado == 'Tarde') + sum(1 for a in aula if a.estado == 'Tarde')
        ausencias = sum(1 for a in aula if a.estado == 'Ausente')

        if tardes >= 3:
            filas += f'<tr><td>{e.nombre} {e.apellido}</td><td>{e.grado}</td><td>{tardes} llegadas tarde</td></tr>'

        if ausencias >= 3:
            filas += f'<tr><td>{e.nombre} {e.apellido}</td><td>{e.grado}</td><td>{ausencias} ausencias</td></tr>'

    content = f"""
    <header class="top">
        <h1>Alertas</h1>
        <a class="btn" href="/dashboard">Volver</a>
    </header>

    <section class="table-card">
        <table>
            <tr>
                <th>Estudiante</th>
                <th>Grado</th>
                <th>Alerta</th>
            </tr>
            {filas}
        </table>
    </section>
    """

    return page('Alertas', shell(content))

def enviar_pin(correo_destino, pin):
    soporte_email = os.environ.get('SOPORTE_EMAIL')
    soporte_password = os.environ.get('SOPORTE_PASSWORD')

    if not soporte_email or not soporte_password:
        return False

    msg = EmailMessage()
    msg['Subject'] = 'PIN de recuperación - EduTrack QR'
    msg['From'] = soporte_email
    msg['To'] = correo_destino
    msg.set_content(f'Tu PIN es: {pin}')

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(soporte_email, soporte_password)
            smtp.send_message(msg)
        return True
    except Exception:
        return False

@app.route('/recuperar', methods=['GET', 'POST'])
def recuperar():
    mensaje = ''

    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip()
        correo = request.form.get('correo', '').strip()

        user = Usuario.query.filter(func.lower(Usuario.usuario) == usuario.lower(), Usuario.correo == correo).first()

        if user:
            pin = str(random.randint(100000, 999999))

            Recuperacion.query.filter_by(usuario=user.usuario).delete()
            db.session.add(Recuperacion(usuario=user.usuario, pin=pin))
            db.session.commit()

            if enviar_pin(correo, pin):
                session['recuperar_usuario'] = user.usuario
                return redirect('/validar_pin')

            mensaje = 'No se pudo enviar el correo.'
        else:
            mensaje = 'Usuario o correo no encontrado.'

    body = f"""
    <div class="center">
        <section class="card login-card">
            <h1>Recuperar contraseña</h1>
            <p>{mensaje}</p>

            <form method="POST">
                <input name="usuario" placeholder="Usuario">
                <input name="correo" placeholder="Correo">
                <button>Enviar PIN</button>
            </form>

            <a href="/login">Volver</a>
        </section>
    </div>
    """

    return page('Recuperar', body)

@app.route('/validar_pin', methods=['GET', 'POST'])
def validar_pin():
    if 'recuperar_usuario' not in session:
        return redirect('/recuperar')

    mensaje = ''

    if request.method == 'POST':
        r = Recuperacion.query.filter_by(
            usuario=session['recuperar_usuario'],
            pin=request.form.get('pin', '')
        ).first()

        if r:
            u = Usuario.query.filter_by(usuario=session['recuperar_usuario']).first()
            u.password = request.form.get('password', '').strip()

            Recuperacion.query.filter_by(usuario=u.usuario).delete()
            db.session.commit()

            session.pop('recuperar_usuario', None)
            return redirect('/login')

        mensaje = 'PIN incorrecto.'

    body = f"""
    <div class="center">
        <section class="card login-card">
            <h1>Nuevo acceso</h1>
            <p>{mensaje}</p>

            <form method="POST">
                <input name="pin" placeholder="PIN">
                <input name="password" type="password" placeholder="Nueva contraseña">
                <button>Guardar</button>
            </form>
        </section>
    </div>
    """

    return page('Validar PIN', body)

@app.route('/soporte-login', methods=['GET', 'POST'])
def soporte_login():
    error = ''

    if request.method == 'POST':
        user = login_usuario(request.form.get('usuario'), request.form.get('password'), 'Soporte')

        if user:
            session['usuario'] = user.usuario
            session['rol'] = user.rol
            return redirect('/soporte')

        error = 'Acceso soporte incorrecto.'

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

    return page('Soporte', body)

@app.route('/soporte')
def soporte():
    if not requiere_login() or rol_actual() != 'Soporte':
        return redirect('/soporte-login')

    content = f"""
    <header class="top">
        <h1>Panel de Soporte</h1>
        <a class="btn btn-red" href="/logout">Salir</a>
    </header>

    <section class="grid">
        <div class="card"><h2>Base de datos</h2><p>Conectada</p></div>
        <div class="card"><h2>Usuarios</h2><p>{Usuario.query.count()}</p></div>
        <div class="card"><h2>Estudiantes</h2><p>{Estudiante.query.count()}</p></div>
        <div class="card"><h2>Ingresos</h2><p>{IngresoPorteria.query.count()}</p></div>
    </section>
    """

    return page('Soporte', shell(content))

@app.route('/exportar_estudiantes')
def exportar_estudiantes():
    wb = Workbook()
    ws = wb.active
    ws.append(['Código', 'Nombres', 'Apellidos', 'Grado', 'Director'])

    for e in Estudiante.query.order_by(Estudiante.grado.asc()).all():
        ws.append([e.codigo, e.nombre, e.apellido, e.grado, e.director])

    b = BytesIO()
    wb.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name='estudiantes_edutrack.xlsx')

@app.route('/exportar_excel_reportes')
def exportar_excel_reportes():
    wb = Workbook()
    ws = wb.active
    ws.append(['Código', 'Nombre', 'Apellido', 'Grado', 'Fecha', 'Hora', 'Estado', 'Periodo'])

    for i in IngresoPorteria.query.join(Estudiante).all():
        ws.append([
            i.estudiante.codigo,
            i.estudiante.nombre,
            i.estudiante.apellido,
            i.estudiante.grado,
            i.fecha,
            i.hora,
            i.estado,
            i.periodo
        ])

    b = BytesIO()
    wb.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name='reportes_edutrack.xlsx')

@app.route('/exportar_pdf')
def exportar_pdf():
    b = BytesIO()
    pdf = canvas.Canvas(b, pagesize=letter)
    pdf.drawString(50, 750, 'Reporte EduTrack QR')
    y = 720

    for i in IngresoPorteria.query.join(Estudiante).all():
        pdf.drawString(50, y, f'{i.estudiante.nombre} {i.estudiante.apellido} | {i.estudiante.grado} | {i.fecha} | {i.hora} | {i.estado}')
        y -= 20

        if y < 50:
            pdf.showPage()
            y = 750

    pdf.save()
    b.seek(0)

    return send_file(b, as_attachment=True, download_name='reporte_ingresos.pdf')

@app.route('/exportar_word')
def exportar_word():
    doc = Document()
    doc.add_heading('Reporte EduTrack QR', 0)

    for i in IngresoPorteria.query.join(Estudiante).all():
        doc.add_paragraph(f'{i.estudiante.nombre} {i.estudiante.apellido} | Grado {i.estudiante.grado} | {i.fecha} | {i.hora} | {i.estado}')

    b = BytesIO()
    doc.save(b)
    b.seek(0)

    return send_file(b, as_attachment=True, download_name='reporte_ingresos.docx')

@app.route('/legal')
def legal():
    return page('Legal', """
    <div class="center">
        <section class="card">
            <h1>Aviso legal</h1>
            <p>EduTrack QR es un sistema institucional desarrollado por Sebastián López / StudyTask.</p>
            <a href="/login">Volver</a>
        </section>
    </div>
    """)

@app.route('/cookies')
def cookies():
    return page('Cookies', """
    <div class="center">
        <section class="card">
            <h1>Cookies</h1>
            <p>Se usan cookies técnicas de sesión para mantener el acceso seguro.</p>
            <a href="/login">Volver</a>
        </section>
    </div>
    """)

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

if __name__ == '__main__':
    with app.app_context():
        inicializar_bd()

    app.run(debug=True, host='0.0.0.0')
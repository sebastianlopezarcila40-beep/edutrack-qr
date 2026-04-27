from flask import Flask, render_template, request, redirect, session, send_file
import sqlite3, os, qrcode, random, smtplib
from email.message import EmailMessage
from datetime import datetime, date, time
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "edutrack-dev-key")

DB = "edutrack.db"
PERIODO_MANUAL = "Periodo 2"
HORARIO_INICIO = time(6, 30, 0)
HORARIO_FIN = time(7, 45, 0)


def conectar():
    return sqlite3.connect(DB)


def columna_existe(tabla, columna):
    conn = conectar()
    c = conn.cursor()
    c.execute(f"PRAGMA table_info({tabla})")
    columnas = [x[1] for x in c.fetchall()]
    conn.close()
    return columna in columnas


def crear_bd():
    conn = conectar()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS estudiantes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE,
        nombre TEXT,
        apellido TEXT,
        grado TEXT,
        director TEXT,
        qr TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE,
        password TEXT,
        rol TEXT,
        correo TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS ingresos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER,
        fecha TEXT,
        hora TEXT,
        dia TEXT,
        estado TEXT,
        periodo TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS recuperacion (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT,
        pin TEXT
    )
    """)

    conn.commit()
    conn.close()

    conn = conectar()
    c = conn.cursor()

    if not columna_existe("usuarios", "correo"):
        c.execute("ALTER TABLE usuarios ADD COLUMN correo TEXT DEFAULT ''")

    if not columna_existe("estudiantes", "apellido"):
        c.execute("ALTER TABLE estudiantes ADD COLUMN apellido TEXT DEFAULT ''")

    c.execute("""
    INSERT OR IGNORE INTO usuarios(usuario,password,rol,correo)
    VALUES('admin','1234','Administrador','')
    """)

    c.execute("""
    INSERT OR IGNORE INTO usuarios(usuario,password,rol,correo)
    VALUES('soporte','1234','Soporte','')
    """)

    conn.commit()
    conn.close()


def dentro_horario_escaneo():
    ahora = datetime.now().time()
    return HORARIO_INICIO <= ahora <= HORARIO_FIN


def enviar_pin(correo_destino, pin):
    soporte_email = os.environ.get("SOPORTE_EMAIL")
    soporte_password = os.environ.get("SOPORTE_PASSWORD")

    if not soporte_email or not soporte_password:
        return False

    msg = EmailMessage()
    msg["Subject"] = "PIN de recuperación - EduTrack QR"
    msg["From"] = soporte_email
    msg["To"] = correo_destino
    msg.set_content(f"Tu PIN de recuperación de EduTrack QR es: {pin}")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(soporte_email, soporte_password)
            smtp.send_message(msg)
        return True
    except Exception:
        return False


def guardar_ingreso(codigo, estado_manual):
    if "Codigo:" in codigo:
        codigo = codigo.split("Codigo:")[1].split("\n")[0].strip()

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT * FROM estudiantes WHERE codigo=?", (codigo,))
    estudiante = c.fetchone()

    if not estudiante:
        conn.close()
        return "Estudiante no registrado", "No registrado"

    fecha = datetime.now().strftime("%Y-%m-%d")
    hora = datetime.now().strftime("%H:%M:%S")
    dia = datetime.now().strftime("%A")

    c.execute("""
    INSERT INTO ingresos(estudiante_id,fecha,hora,dia,estado,periodo)
    VALUES(?,?,?,?,?,?)
    """, (estudiante[0], fecha, hora, dia, estado_manual, PERIODO_MANUAL))

    conn.commit()
    conn.close()

    return f"Registro guardado: {estudiante[2]} {estudiante[3]} - {estado_manual}", estado_manual


@app.route("/")
def inicio():
    return redirect("/login")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"]
        password = request.form["password"]

        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE usuario=? AND password=?", (usuario, password))
        user = c.fetchone()
        conn.close()

        if user:
            session["usuario"] = user[1]
            session["rol"] = user[3]
            return redirect("/dashboard")

        return "Usuario o contraseña incorrectos"

    return render_template("login.html")


@app.route("/dashboard")
def dashboard():
    if "usuario" not in session:
        return redirect("/login")

    hoy = date.today().strftime("%Y-%m-%d")

    conn = conectar()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) FROM estudiantes")
    total_estudiantes = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM ingresos WHERE fecha=? AND estado='Temprano'", (hoy,))
    tempranos_hoy = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM ingresos WHERE fecha=? AND estado='Tarde'", (hoy,))
    tardes_hoy = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM ingresos WHERE fecha=? AND estado='No llegó'", (hoy,))
    no_llegaron_manual = c.fetchone()[0]

    c.execute("""
    SELECT COUNT(*) FROM estudiantes
    WHERE id NOT IN (
        SELECT estudiante_id FROM ingresos WHERE fecha=?
    )
    """, (hoy,))
    no_registrados = c.fetchone()[0]

    no_llegaron_hoy = no_llegaron_manual + no_registrados

    c.execute("""
    SELECT e.grado,
           COUNT(e.id),
           SUM(CASE WHEN i.fecha=? THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.fecha=? AND i.estado='Temprano' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.fecha=? AND i.estado='Tarde' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.fecha=? AND i.estado='No llegó' THEN 1 ELSE 0 END)
    FROM estudiantes e
    LEFT JOIN ingresos i ON e.id = i.estudiante_id
    GROUP BY e.grado
    ORDER BY CAST(e.grado AS INTEGER)
    """, (hoy, hoy, hoy, hoy))
    grupos_hoy = c.fetchall()

    c.execute("""
    SELECT e.grado,
           SUM(CASE WHEN i.periodo='Periodo 1' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.periodo='Periodo 2' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.periodo='Periodo 3' THEN 1 ELSE 0 END)
    FROM estudiantes e
    LEFT JOIN ingresos i ON e.id = i.estudiante_id
    GROUP BY e.grado
    ORDER BY CAST(e.grado AS INTEGER)
    """)
    periodos_grupo = c.fetchall()

    c.execute("""
    SELECT e.grado, e.nombre, e.apellido, i.hora, i.estado
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    WHERE i.fecha=?
    ORDER BY CAST(e.grado AS INTEGER), i.hora DESC
    LIMIT 15
    """, (hoy,))
    ultimos_ingresos = c.fetchall()

    conn.close()

    return render_template(
        "dashboard.html",
        usuario=session["usuario"],
        rol=session["rol"],
        total_estudiantes=total_estudiantes,
        tempranos_hoy=tempranos_hoy,
        tardes_hoy=tardes_hoy,
        no_llegaron_hoy=no_llegaron_hoy,
        grupos_hoy=grupos_hoy,
        periodos_grupo=periodos_grupo,
        ultimos_ingresos=ultimos_ingresos,
        periodo_actual=PERIODO_MANUAL,
        hoy=hoy,
        portal_url="/portal"
    )


@app.route("/portal", methods=["GET", "POST"])
def portal():
    mensaje = ""
    estado = ""
    escaneo_abierto = dentro_horario_escaneo()

    if request.method == "POST":
        codigo = request.form["codigo"]
        estado_manual = request.form["estado"]
        mensaje, estado = guardar_ingreso(codigo, estado_manual)

    return render_template(
        "portal.html",
        mensaje=mensaje,
        estado=estado,
        escaneo_abierto=escaneo_abierto
    )


@app.route("/scanner", methods=["GET", "POST"])
def scanner():
    if "usuario" not in session:
        return redirect("/login")

    mensaje = ""
    estado = ""
    escaneo_abierto = dentro_horario_escaneo()

    if request.method == "POST":
        codigo = request.form["codigo"]
        estado_manual = request.form["estado"]
        mensaje, estado = guardar_ingreso(codigo, estado_manual)

    return render_template(
        "scanner.html",
        mensaje=mensaje,
        estado=estado,
        escaneo_abierto=escaneo_abierto
    )


@app.route("/estudiantes", methods=["GET", "POST"])
def estudiantes():
    if "usuario" not in session:
        return redirect("/login")

    os.makedirs("static/qr", exist_ok=True)

    if request.method == "POST":
        codigo = request.form["codigo"]
        nombre = request.form["nombre"]
        apellido = request.form["apellido"]
        grado = request.form["grado"]
        director = request.form["director"]

        qr_texto = (
            f"Codigo: {codigo}\n"
            f"Nombres: {nombre}\n"
            f"Apellidos: {apellido}\n"
            f"Grado: {grado}\n"
            f"Director: {director}"
        )

        qr_nombre = f"{codigo}.png"
        qr_ruta = os.path.join("static", "qr", qr_nombre)

        img = qrcode.make(qr_texto)
        img.save(qr_ruta)

        conn = conectar()
        c = conn.cursor()
        c.execute("""
        INSERT OR REPLACE INTO estudiantes(codigo,nombre,apellido,grado,director,qr)
        VALUES(?,?,?,?,?,?)
        """, (codigo, nombre, apellido, grado, director, qr_nombre))
        conn.commit()
        conn.close()

        return redirect("/estudiantes")

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT * FROM estudiantes ORDER BY CAST(grado AS INTEGER), nombre")
    lista = c.fetchall()
    conn.close()

    return render_template("estudiantes.html", estudiantes=lista)


@app.route("/eliminar_estudiante/<int:id>")
def eliminar_estudiante(id):
    if "usuario" not in session:
        return redirect("/login")

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT qr FROM estudiantes WHERE id=?", (id,))
    estudiante = c.fetchone()

    if estudiante and estudiante[0]:
        ruta_qr = os.path.join("static", "qr", estudiante[0])
        if os.path.exists(ruta_qr):
            os.remove(ruta_qr)

    c.execute("DELETE FROM ingresos WHERE estudiante_id=?", (id,))
    c.execute("DELETE FROM estudiantes WHERE id=?", (id,))
    conn.commit()
    conn.close()

    return redirect("/estudiantes")


@app.route("/reportes")
def reportes():
    if "usuario" not in session:
        return redirect("/login")

    hoy = date.today().strftime("%Y-%m-%d")

    conn = conectar()
    c = conn.cursor()

    c.execute("""
    SELECT e.codigo, e.nombre, e.apellido, e.grado, e.director,
           i.fecha, i.hora, i.dia, i.estado, i.periodo
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    ORDER BY CAST(e.grado AS INTEGER), e.nombre, i.fecha DESC, i.hora DESC
    """)
    ingresos = c.fetchall()

    c.execute("""
    SELECT codigo, nombre, apellido, grado, director
    FROM estudiantes
    WHERE id NOT IN (
        SELECT estudiante_id FROM ingresos WHERE fecha=?
    )
    ORDER BY CAST(grado AS INTEGER), nombre
    """, (hoy,))
    no_llegaron = c.fetchall()

    c.execute("""
    SELECT e.grado, e.nombre, e.apellido,
           SUM(CASE WHEN i.estado='Temprano' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='Tarde' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='No llegó' THEN 1 ELSE 0 END),
           COUNT(i.id)
    FROM estudiantes e
    LEFT JOIN ingresos i ON e.id = i.estudiante_id
    GROUP BY e.id
    ORDER BY CAST(e.grado AS INTEGER), 5 DESC, e.nombre
    """)
    estadisticas_estudiantes = c.fetchall()

    c.execute("""
    SELECT e.grado, i.periodo,
           SUM(CASE WHEN i.estado='Temprano' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='Tarde' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='No llegó' THEN 1 ELSE 0 END),
           COUNT(i.id)
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    GROUP BY e.grado, i.periodo
    ORDER BY CAST(e.grado AS INTEGER), i.periodo
    """)
    estadisticas_periodos = c.fetchall()

    c.execute("""
    SELECT e.grado, substr(i.fecha, 1, 7),
           SUM(CASE WHEN i.estado='Temprano' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='Tarde' THEN 1 ELSE 0 END),
           SUM(CASE WHEN i.estado='No llegó' THEN 1 ELSE 0 END),
           COUNT(i.id)
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    GROUP BY e.grado, substr(i.fecha, 1, 7)
    ORDER BY CAST(e.grado AS INTEGER), substr(i.fecha, 1, 7)
    """)
    estadisticas_mensuales = c.fetchall()

    conn.close()

    return render_template(
        "reportes.html",
        ingresos=ingresos,
        no_llegaron=no_llegaron,
        estadisticas_estudiantes=estadisticas_estudiantes,
        estadisticas_periodos=estadisticas_periodos,
        estadisticas_mensuales=estadisticas_mensuales,
        hoy=hoy
    )


@app.route("/usuarios", methods=["GET", "POST"])
@app.route("/crear_usuario", methods=["GET", "POST"])
def usuarios():
    if "usuario" not in session:
        return redirect("/login")

    if request.method == "POST":
        usuario = request.form["usuario"]
        password = request.form["password"]
        rol = request.form["rol"]
        correo = request.form["correo"]

        conn = conectar()
        c = conn.cursor()
        c.execute("""
        INSERT OR IGNORE INTO usuarios(usuario,password,rol,correo)
        VALUES(?,?,?,?)
        """, (usuario, password, rol, correo))
        conn.commit()
        conn.close()

        return redirect("/usuarios")

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT id, usuario, rol, correo FROM usuarios ORDER BY rol, usuario")
    lista = c.fetchall()
    conn.close()

    return render_template("usuarios.html", usuarios=lista)


@app.route("/eliminar_usuario/<int:id>")
def eliminar_usuario(id):
    if "usuario" not in session:
        return redirect("/login")

    conn = conectar()
    c = conn.cursor()
    c.execute("DELETE FROM usuarios WHERE id=?", (id,))
    conn.commit()
    conn.close()

    return redirect("/usuarios")


@app.route("/soporte")
def soporte():
    if "usuario" not in session:
        return redirect("/login")

    conn = conectar()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM estudiantes")
    total_estudiantes = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM usuarios")
    total_usuarios = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM ingresos")
    total_ingresos = c.fetchone()[0]
    conn.close()

    return render_template(
        "soporte.html",
        usuario=session["usuario"],
        total_estudiantes=total_estudiantes,
        total_usuarios=total_usuarios,
        total_ingresos=total_ingresos
    )


@app.route("/legal")
def legal():
    return render_template("legal.html")


@app.route("/cookies")
def cookies():
    return render_template("cookies.html")


@app.route("/recuperar", methods=["GET", "POST"])
def recuperar():
    mensaje = ""

    if request.method == "POST":
        usuario = request.form["usuario"]
        correo = request.form["correo"]
        pin = str(random.randint(100000, 999999))

        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE usuario=? AND correo=?", (usuario, correo))
        user = c.fetchone()

        if user:
            c.execute("DELETE FROM recuperacion WHERE usuario=?", (usuario,))
            c.execute("INSERT INTO recuperacion(usuario,pin) VALUES(?,?)", (usuario, pin))
            conn.commit()

            enviado = enviar_pin(correo, pin)

            if enviado:
                session["recuperar_usuario"] = usuario
                conn.close()
                return redirect("/validar_pin")

            mensaje = "No se pudo enviar el correo."
        else:
            mensaje = "Usuario o correo no encontrado."

        conn.close()

    return render_template("recuperar.html", mensaje=mensaje)


@app.route("/validar_pin", methods=["GET", "POST"])
def validar_pin():
    if "recuperar_usuario" not in session:
        return redirect("/recuperar")

    mensaje = ""

    if request.method == "POST":
        pin = request.form["pin"]
        nueva = request.form["password"]
        usuario = session["recuperar_usuario"]

        conn = conectar()
        c = conn.cursor()
        c.execute("SELECT * FROM recuperacion WHERE usuario=? AND pin=?", (usuario, pin))
        existe = c.fetchone()

        if existe:
            c.execute("UPDATE usuarios SET password=? WHERE usuario=?", (nueva, usuario))
            c.execute("DELETE FROM recuperacion WHERE usuario=?", (usuario,))
            conn.commit()
            conn.close()
            session.pop("recuperar_usuario", None)
            return redirect("/login")

        mensaje = "PIN incorrecto."
        conn.close()

    return render_template("validar_pin.html", mensaje=mensaje)


@app.route("/exportar_estudiantes")
def exportar_estudiantes():
    conn = conectar()
    c = conn.cursor()
    c.execute("""
    SELECT codigo,nombre,apellido,grado,director
    FROM estudiantes
    ORDER BY CAST(grado AS INTEGER), nombre
    """)
    datos = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Estudiantes"
    ws.append(["Código", "Nombres", "Apellidos", "Grado", "Director"])

    for fila in datos:
        ws.append(fila)

    archivo = "estudiantes_edutrack.xlsx"
    wb.save(archivo)
    return send_file(archivo, as_attachment=True)


@app.route("/exportar_excel_reportes")
def exportar_excel_reportes():
    conn = conectar()
    c = conn.cursor()
    c.execute("""
    SELECT e.codigo, e.nombre, e.apellido, e.grado, i.fecha, i.hora, i.estado, i.periodo
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    ORDER BY CAST(e.grado AS INTEGER), i.fecha DESC
    """)
    datos = c.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Reportes"
    ws.append(["Código", "Nombre", "Apellido", "Grado", "Fecha", "Hora", "Estado", "Periodo"])

    for fila in datos:
        ws.append(fila)

    archivo = "reportes_edutrack.xlsx"
    wb.save(archivo)
    return send_file(archivo, as_attachment=True)


@app.route("/exportar_pdf")
def exportar_pdf():
    archivo = "reporte_ingresos.pdf"

    conn = conectar()
    c = conn.cursor()
    c.execute("""
    SELECT e.nombre, e.apellido, e.grado, i.fecha, i.hora, i.estado, i.periodo
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    ORDER BY CAST(e.grado AS INTEGER), i.fecha DESC
    """)
    datos = c.fetchall()
    conn.close()

    pdf = canvas.Canvas(archivo, pagesize=letter)
    pdf.drawString(50, 750, "Reporte de ingresos - EduTrack QR")
    y = 720

    for d in datos:
        pdf.drawString(50, y, f"{d[0]} {d[1]} | Grado {d[2]} | {d[3]} | {d[4]} | {d[5]} | {d[6]}")
        y -= 20
        if y < 50:
            pdf.showPage()
            y = 750

    pdf.save()
    return send_file(archivo, as_attachment=True)


@app.route("/exportar_word")
def exportar_word():
    archivo = "reporte_ingresos.docx"

    conn = conectar()
    c = conn.cursor()
    c.execute("""
    SELECT e.nombre, e.apellido, e.grado, i.fecha, i.hora, i.estado, i.periodo
    FROM ingresos i
    INNER JOIN estudiantes e ON e.id = i.estudiante_id
    ORDER BY CAST(e.grado AS INTEGER), i.fecha DESC
    """)
    datos = c.fetchall()
    conn.close()

    doc = Document()
    doc.add_heading("Reporte de ingresos - EduTrack QR", 0)

    for d in datos:
        doc.add_paragraph(f"{d[0]} {d[1]} | Grado {d[2]} | {d[3]} | {d[4]} | {d[5]} | {d[6]}")

    doc.save(archivo)
    return send_file(archivo, as_attachment=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


if __name__ == "__main__":
    crear_bd()
    app.run(debug=True, host="0.0.0.0")
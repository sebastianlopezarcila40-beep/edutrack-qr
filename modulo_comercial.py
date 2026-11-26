# -*- coding: utf-8 -*-
"""
modulo_comercial.py · Fases 1 y 2 · PROCSIS | EduTrack
=======================================================
Se instala con UN bloque en app.py (lo agrega aplicar_comercial.py):

    import modulo_comercial; modulo_comercial.instalar(globals())

FASE 1 (adentro)  ·  /comercial  ·  Gestor Comercial por áreas (solo cuentas corporativas)
    Banners del carrusel, tarjetas "Tu mejor opción", planes, catálogo de tecnología,
    cupones, facturas (y conciliación manual) y cotizaciones recibidas.
    Cada área solo ve y edita lo suyo (ver PERMISOS, se edita libremente).
FASE 2 (afuera)   ·  /publico/inicio  y bloque inyectado en la portada
    Carrusel, "PROCSIS, tu mejor opción", planes, catálogo, consulta de facturas con cupón,
    comprobante, cotización y barra de autogestión.
FASE 3 (después)  ·  pasarela (PAGO_URL) y Edu Aura leyendo /publico/comercial.json (solo catálogo y planes públicos).

INTERRUPTORES (variables de entorno):
  COMERCIAL_MODO     = on (defecto) | off     -> apaga todo el módulo
  COMERCIAL_PORTADA  = "/" (defecto)          -> rutas de la portada donde se inyecta el bloque, separadas por coma
                                                 (vacío = no inyectar; el bloque siempre está en /publico/inicio)
  PAGO_URL           = https://pasarela/..?ref={numero}&valor={valor}   (Fase 3)
"""
import base64
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import time
from datetime import date

from flask import request, session, redirect, jsonify, make_response, abort

MARCA = "pc-bloque-v2"

ADMIN = {"Gerente", "Gerencia", "Superadmin", "Administrador", "Desarrollador", "Developer"}
MKT = {"Marketing", "Marketing y Comunicaciones", "Comunicaciones"}
VENTAS = {"Comercial", "Ventas", "Supervisor de Ventas", "Supervisor", "Operaciones"}
FINANZAS = {"Contabilidad", "Finanzas", "Finanzas y Tesorería", "Tesorería", "Cobranza"}
TI = {"TI", "Tecnología", "Tecnologia", "Soporte", "Soluciones"}

# Menú superior de la portada (lo estructura TI). (texto, id de la sección). Solo aparecen las secciones que tengan contenido.
MENU = (("Soluciones", "pc-planes"), ("Tienda Tech", "pc-tienda"), ("Beneficios y Cupones", "pc-beneficios"),
        ("Soporte y PQR", "pc-auto"), ("Pagar Factura", "pc-pagos"))

# Quién administra cada sección. Para que otra área entre, agrega su nombre de rol aquí.
PERMISOS = {
    "banners": ADMIN | MKT,
    "tarjetas": ADMIN | MKT,
    "equipos": ADMIN | VENTAS,
    "planes": ADMIN | VENTAS,
    "cupones": ADMIN | MKT | FINANZAS,
    "facturas": ADMIN | FINANZAS,
    "cotizaciones": ADMIN | VENTAS,
    "blog": ADMIN | MKT,
    "aura": ADMIN | MKT | TI,
}
APRUEBAN_PLANES = ADMIN  # solo Gerencia publica tarifas

ICONOS = {
    "nube": '<path d="M7 18a4 4 0 0 1-.6-7.96A6 6 0 0 1 18 9.5 4.3 4.3 0 0 1 17.5 18H7z"/>',
    "soporte": '<path d="M4 13v-1a8 8 0 0 1 16 0v1"/><rect x="3" y="13" width="4" height="6" rx="1.5"/><rect x="17" y="13" width="4" height="6" rx="1.5"/><path d="M19 19c0 1.5-2 2.5-5 2.5"/>',
    "escudo": '<path d="M12 3l7 3v5c0 5-3 8.5-7 10-4-1.5-7-5-7-10V6l7-3z"/><path d="M9 12l2 2 4-4"/>',
    "estrella": '<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9L12 3z"/>',
}

# (campo, etiqueta, tipo, requerido)  ·  tipos: text, area, int, date, img, sel:A|B|C
SECCIONES = {
    "banners": {"titulo": "Banners del carrusel", "modelo": "Banner", "orden": "orden",
                "campos": [("titulo", "Título", "text", True), ("texto", "Mensaje", "text", False), ("enlace", "Enlace (opcional)", "text", False),
                           ("imagen", "Imagen de fondo (opcional)", "img", False), ("orden", "Orden", "int", False),
                           ("desde", "Visible desde", "date", False), ("hasta", "Visible hasta", "date", False)],
                "lista": ["titulo", "orden", "desde", "hasta"], "acciones": ["toggle", "borrar"]},
    "tarjetas": {"titulo": "Tarjetas «Tu mejor opción»", "modelo": "Tarjeta", "orden": "orden",
                 "campos": [("titulo", "Título", "text", True), ("texto", "Texto", "area", True),
                            ("icono", "Ícono", "sel:nube|soporte|escudo|estrella", False), ("orden", "Orden", "int", False)],
                 "lista": ["titulo", "icono", "orden"], "acciones": ["toggle", "borrar"]},
    "blog": {"titulo": "Novedades y guías (blog)", "modelo": "Entrada", "orden": "fecha",
             "campos": [("titulo", "Título", "text", True), ("categoria", "Categoría", "sel:EDUCACION|TECNOLOGIA|BENEFICIOS", True),
                        ("resumen", "Resumen (se ve en la tarjeta)", "text", True), ("cuerpo", "Contenido completo", "area", False),
                        ("enlace", "Enlace externo (opcional; si lo dejas vacío se abre la página interna)", "text", False),
                        ("imagen", "Imagen (opcional)", "img", False), ("fecha", "Fecha de publicación", "date", True)],
             "lista": ["titulo", "categoria", "fecha"], "acciones": ["toggle", "borrar"]},
    "aura": {"titulo": "Respuestas de Edu Aura", "modelo": "AuraResp", "orden": "id",
             "campos": [("palabras", "Palabras clave (separadas por coma)", "text", True), ("respuesta", "Respuesta", "area", True)],
             "lista": ["palabras", "respuesta"], "acciones": ["toggle", "borrar"]},
    "planes": {"titulo": "Planes de negocio", "modelo": "Plan", "orden": "tipo",
               "campos": [("nombre", "Nombre del plan", "text", True), ("tipo", "Tipo", "sel:ECONOMICO|CORPORATIVO", True),
                          ("estudiantes_max", "Capacidad de estudiantes (0 = sin límite)", "int", False),
                          ("sedes_max", "Sedes (0 = ilimitadas)", "int", False), ("precio_mensual", "Tarifa mensual COP (0 = cotizar)", "int", False),
                          ("descripcion", "Descripción corta", "text", False), ("caracteristicas", "Características (una por línea)", "area", False)],
                 "lista": ["nombre", "tipo", "estudiantes_max", "sedes_max", "precio_mensual", "aprobado"], "acciones": ["toggle", "aprobar", "borrar"]},
    "equipos": {"titulo": "Catálogo de tecnología", "modelo": "Equipo", "orden": "nombre",
                "campos": [("nombre", "Nombre del equipo o combo", "text", True),
                           ("categoria", "Categoría", "sel:Computador|Lector QR de portería|Punto POS TiendaTrack|Combo equipo + licencia|Accesorio", True),
                           ("ficha", "Ficha técnica", "area", False), ("precio", "Precio COP", "int", False), ("stock", "Stock", "int", False),
                           ("beneficio", "Beneficio incluido (ej. licencia gratis 3 meses)", "text", False), ("foto", "Foto", "img", False)],
                "lista": ["nombre", "categoria", "precio", "stock"], "acciones": ["toggle", "borrar"]},
    "cupones": {"titulo": "Cupones de descuento", "modelo": "Cupon", "orden": "codigo",
                "campos": [("codigo", "Código", "text", True), ("descripcion", "Descripción", "text", False),
                           ("tipo", "Tipo de descuento", "sel:PORCENTAJE|VALOR", True), ("valor", "Valor (% o pesos)", "int", True),
                           ("vence", "Vence (obligatorio)", "date", True), ("usos_max", "Límite de usos (obligatorio, mínimo 1)", "int", True)],
                "lista": ["codigo", "tipo", "valor", "vence", "usos", "usos_max"], "acciones": ["toggle"]},
    "facturas": {"titulo": "Facturas y conciliación", "modelo": "Factura", "orden": "id",
                 "campos": [("numero", "Número (vacío = automático)", "text", False), ("cliente", "Cliente / colegio", "text", True),
                            ("nit", "NIT", "text", False), ("codigo_colegio", "Código de colegio", "text", False), ("documento", "Documento", "text", False),
                            ("concepto", "Concepto", "text", True), ("valor", "Valor COP", "int", True),
                            ("emitida", "Emitida", "date", False), ("vence", "Vence", "date", False)],
                 "lista": ["numero", "cliente", "concepto", "valor", "descuento", "estado", "vence"], "acciones": ["pagada"]},
    "cotizaciones": {"titulo": "Cotizaciones recibidas", "modelo": "Cotizacion", "orden": "id", "campos": [],
                     "lista": ["nombre", "contacto", "equipo_nombre", "cupon", "estado"], "acciones": ["atendida"]},
}

_E = {"tablas": False, "M": {}, "rl": {}}


# ───────────────────────── utilidades puras ─────────────────────────
def e(v):
    return html.escape("" if v is None else str(v), quote=True)


def cop(v):
    try:
        return "$" + format(int(v or 0), ",").replace(",", ".")
    except Exception:
        return "$0"


def norm_ref(v):
    return re.sub(r"[\s.\-]", "", str(v or "")).upper()[:40]


def fecha_ok(s):
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", s or ""))


def hoy():
    return date.today().isoformat()


def visible_hoy(desde, hasta):
    h = hoy()
    return (not desde or desde <= h) and (not hasta or hasta >= h)


def sin_acentos(t):
    t = str(t or "").lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ü", "u"), ("ñ", "n")):
        t = t.replace(a, b)
    return t


def toca(q, *claves):
    """True si la pregunta contiene alguna clave. Palabra sola = prefijo de una palabra (≥4 letras) o palabra exacta (≤3); con espacio = frase."""
    n = sin_acentos(q)
    palabras = set(re.findall(r"[a-z0-9]+", n))
    for c in claves:
        if " " in c:
            if c in n:
                return True
        elif len(c) <= 3:
            if c in palabras:
                return True
        elif any(w.startswith(c) for w in palabras):
            return True
    return False


MESES = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")


def fecha_larga(iso):
    try:
        y, m, d = (int(x) for x in iso.split("-"))
        return "%d %s %d" % (d, MESES[m - 1], y)
    except Exception:
        return ""


def calc_descuento(cupon, valor):
    """cupon: objeto con tipo/valor. Devuelve el descuento en pesos (nunca mayor al valor)."""
    v = int(valor or 0)
    d = (v * int(cupon.valor or 0)) // 100 if cupon.tipo == "PORCENTAJE" else int(cupon.valor or 0)
    return max(0, min(d, v))


def cupon_valido(c):
    if not c or not c.activo:
        return "Cupón no válido."
    if not fecha_ok(c.vence) or c.vence < hoy():
        return "Este cupón venció."
    if int(c.usos_max or 0) > 0 and int(c.usos or 0) >= int(c.usos_max):
        return "Este cupón ya alcanzó su límite de usos."
    return ""


def imagen_data_uri(archivo, max_bytes=400 * 1024):
    """Valida por contenido (no por extensión). Devuelve data URI o '' si no es válida."""
    if not archivo or not getattr(archivo, "filename", ""):
        return ""
    raw = archivo.read(max_bytes + 1)
    if not raw or len(raw) > max_bytes:
        return None
    if raw[:8] == b"\x89PNG\r\n\x1a\n":
        mt = "image/png"
    elif raw[:3] == b"\xff\xd8\xff":
        mt = "image/jpeg"
    elif raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        mt = "image/webp"
    else:
        return None
    return "data:%s;base64,%s" % (mt, base64.b64encode(raw).decode("ascii"))


def limite(clave, maximo, ventana):
    """Límite simple en memoria por clave (IP). True = permitido."""
    ahora = time.time()
    lst = [t for t in _E["rl"].get(clave, []) if ahora - t < ventana]
    if len(lst) >= maximo:
        _E["rl"][clave] = lst
        return False
    lst.append(ahora)
    _E["rl"][clave] = lst
    if len(_E["rl"]) > 5000:
        _E["rl"].clear()
    return True


def igual(a, b):
    return secrets.compare_digest(str(a or "").encode("utf-8"), str(b or "-").encode("utf-8"))


def origen_ok():
    o = request.headers.get("Origin")
    if not o:
        return True
    return re.sub(r"^https?://", "", o).rstrip("/") == request.host


# ───────────────────────── instalación ─────────────────────────
def instalar(G):
    if (os.environ.get("COMERCIAL_MODO") or "on").strip().lower() == "off":
        print("COMERCIAL: apagado (COMERCIAL_MODO=off).", flush=True)
        return
    try:
        _instalar(G)
        print("COMERCIAL: módulo instalado (/comercial y /publico/inicio).", flush=True)
    except Exception as ex:
        print("COMERCIAL: NO se instaló (la app sigue igual):", repr(ex), flush=True)


def _instalar(G):
    app, db = G["app"], G["db"]
    Usuario = G["Usuario"]
    C = db.Column

    class Banner(db.Model):
        __tablename__ = "com_banner"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        titulo = C(db.String(160), default="")
        texto = C(db.String(300), default="")
        enlace = C(db.String(300), default="")
        imagen = C(db.Text, default="")
        orden = C(db.Integer, default=0)
        activo = C(db.Boolean, default=True)
        desde = C(db.String(10), default="")
        hasta = C(db.String(10), default="")

    class Tarjeta(db.Model):
        __tablename__ = "com_tarjeta"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        titulo = C(db.String(120), default="")
        texto = C(db.String(400), default="")
        icono = C(db.String(12), default="estrella")
        orden = C(db.Integer, default=0)
        activo = C(db.Boolean, default=True)

    class Plan(db.Model):
        __tablename__ = "com_plan"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        nombre = C(db.String(120), default="")
        tipo = C(db.String(12), default="ECONOMICO")
        estudiantes_max = C(db.Integer, default=0)
        sedes_max = C(db.Integer, default=0)
        precio_mensual = C(db.Integer, default=0)
        descripcion = C(db.String(300), default="")
        caracteristicas = C(db.Text, default="")
        activo = C(db.Boolean, default=True)
        aprobado = C(db.Boolean, default=False)

    class Equipo(db.Model):
        __tablename__ = "com_equipo"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        nombre = C(db.String(160), default="")
        categoria = C(db.String(60), default="")
        ficha = C(db.Text, default="")
        precio = C(db.Integer, default=0)
        stock = C(db.Integer, default=0)
        beneficio = C(db.String(200), default="")
        foto = C(db.Text, default="")
        activo = C(db.Boolean, default=True)

    class Cupon(db.Model):
        __tablename__ = "com_cupon"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        codigo = C(db.String(40), unique=True, nullable=False)
        descripcion = C(db.String(200), default="")
        tipo = C(db.String(12), default="PORCENTAJE")
        valor = C(db.Integer, default=0)
        vence = C(db.String(10), default="")
        usos_max = C(db.Integer, default=1)
        usos = C(db.Integer, default=0)
        activo = C(db.Boolean, default=True)

    class Factura(db.Model):
        __tablename__ = "com_factura"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        numero = C(db.String(30), unique=True, nullable=False)
        cliente = C(db.String(160), default="")
        nit = C(db.String(40), index=True, default="")
        codigo_colegio = C(db.String(40), index=True, default="")
        documento = C(db.String(40), index=True, default="")
        concepto = C(db.String(200), default="")
        valor = C(db.Integer, default=0)
        descuento = C(db.Integer, default=0)
        cupon = C(db.String(40), default="")
        estado = C(db.String(12), default="PENDIENTE")
        emitida = C(db.String(10), default="")
        vence = C(db.String(10), default="")
        pagada_ts = C(db.BigInteger, default=0)

    class Cotizacion(db.Model):
        __tablename__ = "com_cotizacion"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        nombre = C(db.String(120), default="")
        contacto = C(db.String(160), default="")
        equipo_nombre = C(db.String(160), default="")
        cupon = C(db.String(40), default="")
        mensaje = C(db.Text, default="")
        estado = C(db.String(12), default="NUEVA")
        creado = C(db.BigInteger, default=0)

    class Entrada(db.Model):
        __tablename__ = "com_entrada"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        titulo = C(db.String(160), default="")
        categoria = C(db.String(20), default="EDUCACION")
        resumen = C(db.String(300), default="")
        cuerpo = C(db.Text, default="")
        enlace = C(db.String(300), default="")
        imagen = C(db.Text, default="")
        fecha = C(db.String(10), default="")
        activo = C(db.Boolean, default=True)

    class AuraResp(db.Model):
        __tablename__ = "com_aura"
        __table_args__ = {"extend_existing": True}
        id = C(db.Integer, primary_key=True)
        palabras = C(db.String(300), default="")
        respuesta = C(db.Text, default="")
        activo = C(db.Boolean, default=True)

    M = {"Entrada": Entrada, "AuraResp": AuraResp, "Banner": Banner, "Tarjeta": Tarjeta, "Plan": Plan, "Equipo": Equipo, "Cupon": Cupon, "Factura": Factura, "Cotizacion": Cotizacion}
    _E["M"] = M

    def audit(a, d=""):
        try:
            G["registrar_auditoria"](a, d)
        except Exception:
            pass

    def ip():
        try:
            return G["_client_ip"]()
        except Exception:
            return request.remote_addr or ""

    def tablas():
        if _E["tablas"]:
            return
        try:
            for m in M.values():
                m.__table__.create(bind=db.engine, checkfirst=True)
            if not Banner.query.first() and not Tarjeta.query.first():
                for i, (t, x) in enumerate((("Gestión escolar ágil y segura", "Cumplimos la Ley 1581 de protección de datos personales."),
                                            ("Combos de equipos + software", "Equipo + licencia gratis por meses. Cotiza con tu cupón."),
                                            ("Nuevos planes institucionales 2027", "Planes por capacidad para colegios pequeños y redes educativas."))):
                    db.session.add(Banner(titulo=t, texto=x, orden=i, activo=True))
                for i, (t, x, ic) in enumerate((("Respaldo en la nube", "Tu información académica protegida y respaldada todos los días.", "nube"),
                                                ("Soporte cercano", "Un equipo que responde y te acompaña en cada etapa del año escolar.", "soporte"),
                                                ("Cumplimiento legal", "Habeas Data y Ley 1581: tratamos los datos con reglas claras.", "escudo"))):
                    db.session.add(Tarjeta(titulo=t, texto=x, icono=ic, orden=i, activo=True))
                db.session.add(Plan(nombre="Plan Económico", tipo="ECONOMICO", estudiantes_max=300, sedes_max=1, precio_mensual=0, aprobado=True,
                                    descripcion="Para colegios pequeños y rurales.", caracteristicas="Notas SIEE\nBoletines\nMatrícula"))
                db.session.add(Plan(nombre="Plan Corporativo", tipo="CORPORATIVO", estudiantes_max=0, sedes_max=0, precio_mensual=0, aprobado=True,
                                    descripcion="Para megacolegios y redes educativas.",
                                    caracteristicas="Multi-sede ilimitada\nMayor almacenamiento\nPortería QR avanzada\nSoporte prioritario"))
                db.session.commit()
            if not Entrada.query.first():
                f0 = hoy()
                for t, c, x in (("Cómo proteger los datos de tu colegio", "EDUCACION", "Buenas prácticas de Habeas Data y Ley 1581 para rectores y secretarías."),
                                ("Guía rápida: qué equipo elegir para tu institución", "TECNOLOGIA", "Criterios para escoger computadores, lectores QR y puntos POS."),
                                ("Combos de equipo + licencia", "BENEFICIOS", "Conoce los beneficios vigentes al comprar tecnología con PROCSIS.")):
                    db.session.add(Entrada(titulo=t, categoria=c, resumen=x, cuerpo=x, fecha=f0, activo=True))
                db.session.commit()
            _E["tablas"] = True
        except Exception as ex:
            print("COMERCIAL tablas:", repr(ex)[:160], flush=True)
            try:
                db.session.rollback()
            except Exception:
                pass

    # ───── acceso al área administrativa ─────
    def usuario_sesion():
        uid = session.get("uid") or session.get("user_id")
        try:
            return Usuario.query.get(int(uid)) if uid else None
        except Exception:
            return None

    def roles_de(u):
        r = (u.rol or "").strip()
        return r

    def puede(u, sec):
        return bool(u and u.institucion_id is None and getattr(u, "activo", True) is not False and roles_de(u) in PERMISOS.get(sec, ()))

    def csrf():
        if not session.get("com_csrf"):
            session["com_csrf"] = secrets.token_hex(16)
        return session["com_csrf"]

    def envolver(titulo, cuerpo):
        pg = G.get("page")
        if pg:
            try:
                return pg(titulo, cuerpo)
            except Exception:
                pass
        return "<!doctype html><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'><title>%s</title>%s" % (e(titulo), cuerpo)

    CSS_ADM = """<style>
.cm{max-width:1100px;margin:0 auto;padding:18px 14px 60px;color:#0F172A;font-family:inherit}
.cm h1{font-size:24px;margin:0 0 4px;color:#0020C2}.cm .sub{color:#64748b;margin:0 0 16px;font-size:14px}
.cm nav{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 18px}
.cm nav a{padding:8px 14px;border-radius:999px;border:1px solid #cbd5e1;background:#fff;color:#0F172A;text-decoration:none;font-size:14px;font-weight:600}
.cm nav a.on{background:#0020C2;color:#fff;border-color:#0020C2;box-shadow:0 0 0 2px #00FF66 inset}
.cm .card{background:#F8FAFC;border:1px solid #e2e8f0;border-radius:18px;padding:16px;margin:0 0 16px}
.cm label{display:block;margin:10px 0 4px;font-size:13px;font-weight:700;color:#334155}
.cm input,.cm select,.cm textarea{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:10px;font:inherit;background:#fff;color:#0F172A}
.cm .g{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:0 14px}
.cm button,.cm .btn{border:0;border-radius:999px;padding:10px 18px;font-weight:700;background:#0020C2;color:#fff;cursor:pointer;font-size:14px;text-decoration:none;display:inline-block}
.cm .sec{background:#fff;color:#0020C2;border:1px solid #0020C2}.cm .pel{background:#fff;color:#b91c1c;border:1px solid #fca5a5}
.cm table{width:100%;border-collapse:collapse;font-size:14px;background:#fff;border-radius:12px;overflow:hidden}
.cm th,.cm td{padding:9px 10px;border-bottom:1px solid #e2e8f0;text-align:left;vertical-align:middle}.cm th{background:#eef2ff;color:#0020C2;font-size:12.5px}
.cm .tw{overflow-x:auto}.cm .ok{color:#067647;font-weight:700}.cm .no{color:#b91c1c;font-weight:700}
.cm .msg{padding:10px 12px;border-radius:12px;background:#ecfdf3;color:#067647;margin:0 0 14px}.cm .er{background:#fef2f2;color:#b91c1c}
.cm form.i{display:inline}
</style>"""

    def render_admin(u, sec, msg="", err="", editar=None):
        vis = [k for k in SECCIONES if puede(u, k)]
        nav = "".join('<a href="/comercial?s=%s" class="%s">%s</a>' % (k, "on" if k == sec else "", e(SECCIONES[k]["titulo"])) for k in vis)
        cfg = SECCIONES[sec]
        Mo = M[cfg["modelo"]]
        filas = Mo.query.order_by(getattr(Mo, cfg["orden"]).desc() if cfg["orden"] in ("id", "fecha") else getattr(Mo, cfg["orden"])).limit(300).all()
        t = csrf()
        form = ""
        if cfg["campos"]:
            reg = editar
            camp = []
            for (n, lab, tipo, req) in cfg["campos"]:
                val = getattr(reg, n, "") if reg is not None else ""
                val = "" if val is None else val
                rq = " required" if req else ""
                if tipo == "area":
                    inp = '<textarea name="%s" rows="3"%s>%s</textarea>' % (n, rq, e(val))
                elif tipo == "int":
                    inp = '<input type="number" name="%s" min="0" value="%s"%s>' % (n, e(val if val != "" else 0), rq)
                elif tipo == "date":
                    inp = '<input type="date" name="%s" value="%s"%s>' % (n, e(val), rq)
                elif tipo == "img":
                    inp = '<input type="file" name="%s" accept="image/png,image/jpeg,image/webp"><small style="color:#64748b">PNG, JPG o WEBP · máx. 400 KB%s</small>' % (n, " · ya tiene imagen (si no subes otra, se conserva)" if val else "")
                elif tipo.startswith("sel:"):
                    ops = tipo[4:].split("|")
                    inp = '<select name="%s">%s</select>' % (n, "".join('<option%s>%s</option>' % (" selected" if val == o else "", e(o)) for o in ops))
                else:
                    inp = '<input name="%s" value="%s" maxlength="300"%s>' % (n, e(val), rq)
                camp.append("<div><label>%s</label>%s</div>" % (e(lab), inp))
            form = ('<div class="card"><h3 style="margin:0 0 6px">%s</h3><form method="post" action="/comercial/guardar/%s" enctype="multipart/form-data">'
                    '<input type="hidden" name="csrf" value="%s"><input type="hidden" name="id" value="%s"><div class="g">%s</div>'
                    '<p style="margin:14px 0 0"><button type="submit">Guardar</button> %s</p></form></div>'
                    % ("Editar" if reg is not None else "Nuevo", sec, t, reg.id if reg is not None else "", "".join(camp),
                       '<a class="btn sec" href="/comercial?s=%s">Cancelar</a>' % sec if reg is not None else ""))
        cab = "".join("<th>%s</th>" % e(c.replace("_", " ")) for c in cfg["lista"]) + "<th></th>"
        cuerpo = []
        for r in filas:
            tds = []
            for c in cfg["lista"]:
                v = getattr(r, c, "")
                if c in ("precio", "valor", "precio_mensual", "descuento") and not (c == "valor" and sec == "cupones"):
                    v = cop(v)
                elif c == "valor" and sec == "cupones":
                    v = ("%s%%" % v) if r.tipo == "PORCENTAJE" else cop(v)
                elif c == "aprobado":
                    v = '<span class="ok">Aprobado</span>' if v else '<span class="no">Pendiente</span>'
                    tds.append("<td>%s</td>" % v)
                    continue
                tds.append("<td>%s</td>" % e(v))
            acc = []
            if cfg["campos"]:
                acc.append('<a class="btn sec" href="/comercial?s=%s&editar=%s">Editar</a>' % (sec, r.id))
            for a in cfg["acciones"]:
                if a == "aprobar" and (roles_de(u) not in APRUEBAN_PLANES):
                    continue
                if a == "pagada" and r.estado == "PAGADA":
                    continue
                if a == "atendida" and r.estado == "ATENDIDA":
                    continue
                if a == "toggle":
                    lab = "Desactivar" if getattr(r, "activo", True) else "Activar"
                elif a == "aprobar":
                    lab = "Quitar aprobación" if r.aprobado else "Aprobar tarifa"
                elif a == "pagada":
                    lab = "Marcar pagada"
                elif a == "atendida":
                    lab = "Marcar atendida"
                else:
                    lab = "Borrar"
                acc.append('<form class="i" method="post" action="/comercial/accion/%s/%s/%s"%s><input type="hidden" name="csrf" value="%s"><button class="%s" type="submit">%s</button></form>'
                           % (sec, r.id, a, ' onsubmit="return confirm(\'¿Seguro?\')"' if a == "borrar" else "", t, "pel" if a == "borrar" else "sec", lab))
            extra = ""
            if sec == "cotizaciones" and r.mensaje:
                extra = '<tr><td colspan="%d" style="color:#475569;background:#F8FAFC">%s</td></tr>' % (len(cfg["lista"]) + 1, e(r.mensaje))
            if not getattr(r, "activo", True):
                tds[0] = tds[0].replace("<td>", '<td style="opacity:.5">', 1)
            cuerpo.append("<tr>%s<td style='white-space:nowrap'>%s</td></tr>%s" % ("".join(tds), " ".join(acc), extra))
        tabla = '<div class="tw"><table><tr>%s</tr>%s</table></div>' % (cab, "".join(cuerpo) or '<tr><td colspan="9">Aún no hay registros.</td></tr>')
        nota = ""
        if sec == "planes":
            nota = '<p class="sub">Solo los planes activos y aprobados por Gerencia se muestran en la portada pública.</p>'
        if sec == "cupones":
            nota = '<p class="sub">Todo cupón lleva fecha de vencimiento y límite de usos. Se desactivan, no se borran, para conservar el historial contable.</p>'
        aviso = ('<div class="msg">%s</div>' % e(msg)) if msg else ""
        aviso += ('<div class="msg er">%s</div>' % e(err)) if err else ""
        return envolver("Gestor comercial", CSS_ADM + '<div class="cm"><h1>Gestor Comercial</h1><p class="sub">Beneficios, catálogo, planes, cupones y facturación de PROCSIS. Cada área ve solo lo suyo.</p><nav>%s</nav>%s<h2 style="margin:0 0 8px">%s</h2>%s%s%s</div>'
                        % (nav, aviso, e(cfg["titulo"]), nota, form, tabla))

    def primera_seccion(u):
        for k in SECCIONES:
            if puede(u, k):
                return k
        return None

    # ───── rutas administrativas ─────
    def com_admin():
        tablas()
        u = usuario_sesion()
        if not u:
            return redirect("/login")
        sec = request.args.get("s") or primera_seccion(u)
        if not sec or sec not in SECCIONES or not puede(u, sec):
            abort(403)
        editar = None
        if request.args.get("editar", "").isdigit():
            editar = M[SECCIONES[sec]["modelo"]].query.get(int(request.args["editar"]))
        return render_admin(u, sec, request.args.get("ok", ""), request.args.get("err", ""), editar)

    def com_guardar(sec):
        tablas()
        u = usuario_sesion()
        if not u or sec not in SECCIONES or not puede(u, sec) or not SECCIONES[sec]["campos"]:
            abort(403)
        if not igual(request.form.get("csrf"), session.get("com_csrf")):
            abort(400)
        cfg = SECCIONES[sec]
        Mo = M[cfg["modelo"]]
        rid = request.form.get("id", "")
        reg = Mo.query.get(int(rid)) if rid.isdigit() else None
        if rid and reg is None:
            abort(404)
        datos = {}
        for (n, lab, tipo, req) in cfg["campos"]:
            if tipo == "img":
                nuevo = imagen_data_uri(request.files.get(n))
                if nuevo is None:
                    return redirect("/comercial?s=%s&err=%s" % (sec, "La imagen no es válida (PNG, JPG o WEBP de máximo 400 KB)."))
                if nuevo:
                    datos[n] = nuevo
                continue
            v = (request.form.get(n) or "").strip()
            if tipo == "int":
                v = int(v) if re.match(r"^\d{1,9}$", v) else 0
            elif tipo == "date" and v and not fecha_ok(v):
                return redirect("/comercial?s=%s&err=%s" % (sec, "Fecha no válida: " + lab))
            elif tipo.startswith("sel:") and v not in tipo[4:].split("|"):
                v = tipo[4:].split("|")[0]
            elif tipo in ("text", "area"):
                v = v[:2000 if tipo == "area" else 300]
            if req and (v == "" or (tipo == "int" and sec == "cupones" and v < 1)):
                return redirect("/comercial?s=%s&err=%s" % (sec, "Completa el campo: " + lab))
            datos[n] = v
        if sec == "cupones":
            datos["codigo"] = norm_ref(datos["codigo"])
            if datos["tipo"] == "PORCENTAJE" and not 1 <= datos["valor"] <= 100:
                return redirect("/comercial?s=cupones&err=El porcentaje debe estar entre 1 y 100.")
            otro = Mo.query.filter(Mo.codigo == datos["codigo"]).first()
            if otro and (reg is None or otro.id != reg.id):
                return redirect("/comercial?s=cupones&err=Ya existe un cupón con ese código.")
        if sec == "facturas":
            datos["numero"] = norm_ref(datos.get("numero")) or "F%s%s" % (time.strftime("%y%m"), secrets.token_hex(2).upper())
            datos["nit"], datos["codigo_colegio"], datos["documento"] = norm_ref(datos.get("nit")), norm_ref(datos.get("codigo_colegio")), norm_ref(datos.get("documento"))
            if not (datos["nit"] or datos["codigo_colegio"] or datos["documento"]):
                return redirect("/comercial?s=facturas&err=Indica NIT, código de colegio o documento para que el cliente pueda consultarla.")
            otro = Mo.query.filter(Mo.numero == datos["numero"]).first()
            if otro and (reg is None or otro.id != reg.id):
                return redirect("/comercial?s=facturas&err=Ya existe una factura con ese número.")
        if sec == "planes" and reg is not None:
            datos_cambia_tarifa = any(str(getattr(reg, k, "")) != str(datos.get(k, "")) for k in ("precio_mensual", "estudiantes_max", "sedes_max"))
            if datos_cambia_tarifa and roles_de(u) not in APRUEBAN_PLANES:
                reg.aprobado = False  # cambiar tarifa o capacidad exige nueva aprobación de Gerencia
        if reg is None:
            reg = Mo()
            if sec == "planes":
                reg.aprobado = roles_de(u) in APRUEBAN_PLANES
            db.session.add(reg)
        for k, v in datos.items():
            setattr(reg, k, v)
        try:
            db.session.commit()
        except Exception as ex:
            db.session.rollback()
            print("COMERCIAL guardar:", repr(ex)[:160], flush=True)
            return redirect("/comercial?s=%s&err=No se pudo guardar. Revisa los datos." % sec)
        audit("Comercial: %s guardado" % sec, "%s · id %s" % (u.usuario, reg.id))
        return redirect("/comercial?s=%s&ok=Guardado." % sec)

    def com_accion(sec, rid, accion):
        tablas()
        u = usuario_sesion()
        if not u or sec not in SECCIONES or not puede(u, sec) or accion not in SECCIONES[sec]["acciones"]:
            abort(403)
        if not igual(request.form.get("csrf"), session.get("com_csrf")):
            abort(400)
        reg = M[SECCIONES[sec]["modelo"]].query.get(rid)
        if reg is None:
            abort(404)
        if accion == "toggle":
            reg.activo = not bool(reg.activo)
        elif accion == "aprobar":
            if roles_de(u) not in APRUEBAN_PLANES:
                abort(403)
            reg.aprobado = not bool(reg.aprobado)
        elif accion == "pagada":
            reg.estado, reg.pagada_ts = "PAGADA", int(time.time())
        elif accion == "atendida":
            reg.estado = "ATENDIDA"
        elif accion == "borrar":
            db.session.delete(reg)
        db.session.commit()
        audit("Comercial: %s %s" % (sec, accion), "%s · id %s" % (u.usuario, rid))
        return redirect("/comercial?s=%s&ok=Hecho." % sec)

    # ───── Fase 2: bloque público ─────
    def datos_publicos():
        tablas()
        ban = [b for b in Banner.query.filter_by(activo=True).order_by(Banner.orden, Banner.id).all() if visible_hoy(b.desde, b.hasta)]
        tar = Tarjeta.query.filter_by(activo=True).order_by(Tarjeta.orden, Tarjeta.id).all()
        pla = Plan.query.filter_by(activo=True, aprobado=True).order_by(Plan.tipo, Plan.id).all()
        equ = Equipo.query.filter_by(activo=True).order_by(Equipo.categoria, Equipo.nombre).all()
        return ban, tar, pla, equ

    def rutas_existentes():
        return {r.rule for r in app.url_map.iter_rules()}

    def entradas_publicas():
        tablas()
        return Entrada.query.filter_by(activo=True).order_by(Entrada.fecha.desc(), Entrada.id.desc()).limit(3).all()

    def barra_html(ids, base=""):
        links = "".join('<a href="%s#%s">%s</a>' % (base, i, e(n)) for n, i in MENU if i in ids)
        return ('<header class="pc-bar" id="pc-bar"><a class="pc-lg" href="/"><b>P</b><span>PROCSIS | EduTrack</span></a><nav aria-label="Menú principal">%s</nav>'
                '<a class="pc-in" href="/login">Ingresar colegio</a></header>' % links)

    def bloque(base=""):
        """Devuelve {"barra": ..., "cuerpo": ...}. base='/publico/inicio' para páginas internas."""
        ban, tar, pla, equ = datos_publicos()
        blog = entradas_publicas()
        s, ids = [], {"pc-pagos", "pc-auto"}
        if ban:
            sl = []
            for i, b in enumerate(ban):
                fondo = ' style="background-image:linear-gradient(110deg,rgba(0,32,194,.92),rgba(0,0,94,.82)),url(/publico/img/banner/%d)"' % b.id if b.imagen else ""
                cta = ('<a class="pc-cta" href="%s">Conocer más</a>' % e(b.enlace)) if b.enlace and re.match(r"^(/|https://)", b.enlace) else ""
                sl.append('<div class="pc-sl%s"%s><h2>%s</h2><p>%s</p>%s</div>' % (" on" if i == 0 else "", fondo, e(b.titulo), e(b.texto), cta))
            pts = "".join('<i class="%s"></i>' % ("on" if i == 0 else "") for i in range(len(ban))) if len(ban) > 1 else ""
            s.append('<div class="pc-car" id="pc-car">%s<div class="pc-dots">%s</div></div>' % ("".join(sl), pts))
        if blog:
            cards = []
            for b in blog:
                img = '<img loading="lazy" alt="" src="/publico/img/entrada/%d">' % b.id if b.imagen else ""
                href = b.enlace if b.enlace and re.match(r"^(/|https://)", b.enlace) else "/publico/blog/%d" % b.id
                cards.append('<article class="pc-c">%s<span class="pc-tag">%s</span><h3>%s</h3><p>%s</p><div class="pc-dt">%s</div><a class="pc-b" href="%s">Míralo aquí</a></article>'
                             % (img, e(b.categoria), e(b.titulo), e(b.resumen), e(fecha_larga(b.fecha)), e(href)))
            ids.add("pc-blog")
            s.append('<h2 class="pc-h" id="pc-blog">¡Descubre novedades y guías PROCSIS!</h2><div class="pc-gr">%s</div>' % "".join(cards))
        if tar:
            cards = "".join('<div class="pc-c"><svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="#0020C2" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">%s</svg><h3>%s</h3><p>%s</p></div>'
                            % (ICONOS.get(t.icono, ICONOS["estrella"]), e(t.titulo), e(t.texto)) for t in tar)
            ids.add("pc-beneficios")
            s.append('<h2 class="pc-h" id="pc-beneficios">PROCSIS, tu mejor opción</h2><div class="pc-gr">%s</div>' % cards)
        if pla:
            cards = []
            for p in pla:
                cap = "Hasta %s estudiantes" % format(p.estudiantes_max, ",").replace(",", ".") if p.estudiantes_max else "Estudiantes sin límite"
                sed = "%d sede%s" % (p.sedes_max, "" if p.sedes_max == 1 else "s") if p.sedes_max else "Multi-sede ilimitada"
                carac = "".join("<li>%s</li>" % e(x.strip()) for x in (p.caracteristicas or "").splitlines() if x.strip())
                pr = ("%s<small>/mes</small>" % cop(p.precio_mensual)) if p.precio_mensual else "Cotiza tu plan"
                cards.append('<div class="pc-c pc-pl%s"><span class="pc-tag">%s</span><h3>%s</h3><p>%s</p><div class="pc-pr">%s</div><ul><li>%s</li><li>%s</li>%s</ul><a class="pc-b" href="#pc-cot" data-plan="%s">Quiero este plan</a></div>'
                             % (" corp" if p.tipo == "CORPORATIVO" else "", "Alta capacidad" if p.tipo == "CORPORATIVO" else "Económico", e(p.nombre), e(p.descripcion), pr, e(cap), e(sed), carac, e(p.nombre)))
            ids.add("pc-planes")
            s.append('<h2 class="pc-h" id="pc-planes">Planes para cada colegio</h2><div class="pc-gr">%s</div>' % "".join(cards))
        opts = ""
        if equ:
            cards = []
            for q in equ:
                img = '<img loading="lazy" alt="%s" src="/publico/img/equipo/%d">' % (e(q.nombre), q.id) if q.foto else '<div class="pc-ni"></div>'
                ben = '<div class="pc-bn">%s</div>' % e(q.beneficio) if q.beneficio else ""
                agot = (q.stock or 0) <= 0
                cards.append('<div class="pc-c">%s<span class="pc-tag">%s</span><h3>%s</h3><p style="white-space:pre-line">%s</p>%s<div class="pc-pr">%s</div><a class="pc-b%s" href="#pc-cot" data-eq="%d">%s</a></div>'
                             % (img, e(q.categoria), e(q.nombre), e((q.ficha or "")[:300]), ben, cop(q.precio) if q.precio else "Cotizar", " off" if agot else "", q.id, "Cotizar o comprar con cupón" if not agot else "Consultar disponibilidad"))
            ids.add("pc-tienda")
            s.append('<h2 class="pc-h" id="pc-tienda">Beneficios y tecnología</h2><div class="pc-gr">%s</div>' % "".join(cards))
            opts = "".join('<option value="%d">%s</option>' % (q.id, e(q.nombre)) for q in equ)
        s.append('''<h2 class="pc-h" id="pc-pagos">Pagos, facturas y cupones</h2>
<div class="pc-box"><p class="pc-s">Consulta tu factura con el NIT, el código de tu colegio o tu documento.</p>
<div class="pc-row"><input id="pc-ref" placeholder="NIT, código de colegio o documento" maxlength="40" autocomplete="off"><button class="pc-b" id="pc-bus" type="button">Consultar</button></div>
<div id="pc-fac"></div></div>''')
        s.append('''<h2 class="pc-h" id="pc-cot">Cotizar o comprar</h2>
<div class="pc-box"><div class="pc-fm"><input id="pc-n" placeholder="Nombre o colegio" maxlength="120"><input id="pc-c" placeholder="Correo o celular" maxlength="160">
<select id="pc-e"><option value="">Plan o equipo de interés…</option>%s</select><input id="pc-k" placeholder="Cupón (opcional)" maxlength="40"></div>
<textarea id="pc-m" rows="2" maxlength="600" placeholder="Cuéntanos qué necesitas (opcional)"></textarea>
<p style="margin:10px 0 0"><button class="pc-b" id="pc-env" type="button">Enviar solicitud</button> <span id="pc-ok" class="pc-s"></span></p></div>''' % opts)
        rutas = rutas_existentes()
        auto = []
        for nombre, cands in (("Radicar PQR", ("/pqr", "/radicar-pqr", "/pqr/nueva")), ("Agendar soporte", ("/agendar-soporte", "/soporte/agendar", "/agendar_soporte")),
                              ("Bloquear cuenta", ("/bloquear-cuenta", "/bloquear_cuenta", "/acceso/bloquear"))):
            for c in cands:
                if c in rutas:
                    auto.append('<a href="%s">%s</a>' % (c, nombre))
                    break
        auto.append('<a href="/login">Recuperar acceso / MFA</a>')
        s.append('<h2 class="pc-h" id="pc-auto">Autogestión segura</h2><div class="pc-au">%s</div><p class="pc-s" style="text-align:center">Tus datos se tratan conforme a la Ley 1581 de 2012 (Habeas Data).</p>' % "".join(auto))
        flot = ('<div class="pc-fl pc-l"><button class="pc-fb" id="pc-a1" type="button" aria-label="Accesibilidad" aria-expanded="false" title="Accesibilidad">'
                '<svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="4.5" r="1.8"/><path d="M5 8.5l7 1.5 7-1.5M12 10v5M9 21l3-6 3 6"/></svg></button>'
                '<div class="pc-pn" id="pc-ap" hidden role="dialog" aria-label="Opciones de accesibilidad"><h4>Accesibilidad</h4><div class="pc-ar">'
                '<button type="button" data-a="mas">A+ Texto más grande</button><button type="button" data-a="menos">A− Texto más pequeño</button>'
                '<button type="button" data-a="hc" aria-pressed="false">Alto contraste</button><button type="button" data-a="lk" aria-pressed="false">Resaltar enlaces</button>'
                '<button type="button" data-a="rs">Restablecer</button></div></div></div>')
        if (os.environ.get("COMERCIAL_AURA") or "on").strip().lower() != "off":
            flot += ('<div class="pc-fl pc-r"><button class="pc-fb" id="pc-i1" type="button" aria-expanded="false"><span>Pregúntame · Edu Aura</span></button>'
                     '<div class="pc-pn pc-ai" id="pc-ip" hidden role="dialog" aria-label="Edu Aura, asistente de PROCSIS"><div class="pc-ah"><b>Edu Aura</b><small>Tu IA de confianza</small></div>'
                     '<div class="pc-ms" id="pc-ms" aria-live="polite"></div>'
                     '<div class="pc-ch"><button type="button">Planes y precios</button><button type="button">¿Qué equipo elegir?</button><button type="button">Cómo radicar una PQR</button><button type="button">Pagar mi factura</button></div>'
                     '<div class="pc-qr"><input id="pc-q" maxlength="300" placeholder="Escribe tu pregunta" autocomplete="off"><button type="button" id="pc-qs" aria-label="Enviar">Enviar</button></div>'
                     '<p class="pc-ft">Respondo con información pública de PROCSIS. No tengo acceso a datos de colegios.</p></div></div>')
        cuerpo = '<!--%s--><style>%s</style><section class="pc-root" id="pc-root">%s</section>%s<script>%s</script>' % (MARCA, CSS_PUB, "".join(s), flot, JS_PUB)
        return {"barra": barra_html(ids, base), "cuerpo": cuerpo, "ids": ids}

    def envolver_publico(contenido, titulo="PROCSIS | EduTrack"):
        r = make_response("<!doctype html><html lang=es><head><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'><title>%s</title>"
                          "<style>body{margin:0;background:#fff;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;color:#0F172A}</style></head><body>%s</body></html>"
                          % (e(titulo), contenido))
        r.headers["Cache-Control"] = "no-cache"
        return r

    def com_inicio():
        b = bloque("/publico/inicio")
        return envolver_publico(b["barra"] + b["cuerpo"])

    def inyectar(resp):
        try:
            rutas = [x.strip() for x in (os.environ.get("COMERCIAL_PORTADA", "/") or "").split(",") if x.strip()]
            if request.method != "GET" or request.path not in rutas or session.get("usuario"):
                return resp
            if resp.status_code != 200 or resp.mimetype != "text/html" or resp.direct_passthrough:
                return resp
            h = resp.get_data(as_text=True)
            if MARCA in h:
                return resp
            b = bloque()
            m = re.search(r"<body[^>]*>", h, re.I)
            if m and (os.environ.get("COMERCIAL_BARRA") or "on").strip().lower() != "off":
                h = h[:m.end()] + b["barra"] + h[m.end():]
            i = h.lower().rfind("</body>")
            if i < 0:
                return resp
            resp.set_data(h[:i] + b["cuerpo"] + h[i:])
        except Exception as ex:
            print("COMERCIAL inyectar:", repr(ex)[:140], flush=True)
        return resp

    # ───── endpoints públicos (JSON) ─────
    def js(d, st=200):
        r = jsonify(d)
        r.status_code = st
        r.headers["Cache-Control"] = "no-store"
        return r

    def cuerpo_json():
        if not origen_ok() or not request.is_json:
            return None
        d = request.get_json(silent=True)
        return d if isinstance(d, dict) else None

    def facturas_de(ref):
        r = norm_ref(ref)
        if len(r) < 4:
            return []
        return Factura.query.filter((Factura.nit == r) | (Factura.codigo_colegio == r) | (Factura.documento == r)).order_by(Factura.id.desc()).limit(10).all()

    def fac_json(f):
        return {"numero": f.numero, "cliente": f.cliente, "concepto": f.concepto, "valor": f.valor, "descuento": f.descuento or 0,
                "total": max(0, (f.valor or 0) - (f.descuento or 0)), "cupon": f.cupon or "", "estado": f.estado, "vence": f.vence or ""}

    def pub_consultar():
        tablas()
        d = cuerpo_json()
        if d is None:
            return js({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        if not limite("c:" + ip(), 12, 600):
            return js({"ok": False, "mensaje": "Demasiadas consultas. Espera unos minutos."}, 429)
        fs = facturas_de(d.get("ref"))
        if not fs:
            return js({"ok": True, "facturas": [], "mensaje": "No encontramos facturas con ese dato."})
        return js({"ok": True, "facturas": [fac_json(f) for f in fs], "pasarela": bool(os.environ.get("PAGO_URL"))})

    def factura_de_ref(d):
        num, ref = norm_ref(d.get("numero")), norm_ref(d.get("ref"))
        f = Factura.query.filter_by(numero=num).first() if num else None
        if f is None or len(ref) < 4 or ref not in (f.nit, f.codigo_colegio, f.documento):
            return None
        return f

    def pub_cupon():
        tablas()
        d = cuerpo_json()
        if d is None:
            return js({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        if not limite("k:" + ip(), 10, 600):
            return js({"ok": False, "mensaje": "Demasiados intentos. Espera unos minutos."}, 429)
        f = factura_de_ref(d)
        if f is None or f.estado != "PENDIENTE":
            return js({"ok": False, "mensaje": "No se puede aplicar un cupón a esta factura."}, 400)
        if f.cupon:
            return js({"ok": False, "mensaje": "Esta factura ya tiene un cupón aplicado."}, 400)
        c = Cupon.query.filter_by(codigo=norm_ref(d.get("codigo"))).first()
        motivo = cupon_valido(c)
        if motivo:
            return js({"ok": False, "mensaje": motivo}, 400)
        desc = calc_descuento(c, f.valor)
        if desc <= 0:
            return js({"ok": False, "mensaje": "Este cupón no aplica a esta factura."}, 400)
        # un solo uso atómico: la condición de usos va en el UPDATE para que dos pagos simultáneos no pasen el límite
        q = Cupon.query.filter(Cupon.id == c.id, Cupon.activo == True)  # noqa: E712
        if int(c.usos_max or 0) > 0:
            q = q.filter(Cupon.usos < Cupon.usos_max)
        n = q.update({"usos": Cupon.usos + 1}, synchronize_session=False)
        if n != 1:
            db.session.rollback()
            return js({"ok": False, "mensaje": "Este cupón ya alcanzó su límite de usos."}, 400)
        f.cupon, f.descuento = c.codigo, desc
        db.session.commit()
        audit("Comercial: cupón aplicado", "%s · %s · -%s" % (f.numero, c.codigo, desc))
        return js({"ok": True, "factura": fac_json(f), "mensaje": "Cupón aplicado: ahorras %s." % cop(desc)})

    def pub_ir(numero):
        tablas()
        f = factura_de_ref({"numero": numero, "ref": request.args.get("ref")})
        if f is None or f.estado != "PENDIENTE":
            abort(404)
        url = os.environ.get("PAGO_URL", "")
        if not url:
            return envolver_publico('<div class="pc-box" style="max-width:560px;margin:30px auto"><h2>Pasarela en configuración</h2><p class="pc-s">Pronto podrás pagar en línea con PSE, Nequi y tarjetas. Mientras tanto, contacta a tu asesor para pagar la factura %s.</p><p><a class="pc-b" href="/publico/inicio">Volver</a></p></div><style>%s</style>' % (e(f.numero), CSS_PUB)), 503
        destino = url.replace("{numero}", f.numero).replace("{valor}", str(max(0, (f.valor or 0) - (f.descuento or 0))))
        return redirect(destino)

    def pub_comprobante(numero):
        tablas()
        f = factura_de_ref({"numero": numero, "ref": request.args.get("ref")})
        if f is None or f.estado != "PAGADA":
            abort(404)
        tot = max(0, (f.valor or 0) - (f.descuento or 0))
        pagada = time.strftime("%Y-%m-%d %H:%M", time.localtime(f.pagada_ts)) if f.pagada_ts else ""
        r = make_response("<!doctype html><html lang=es><meta charset=utf-8><title>Comprobante %s</title><style>body{font-family:system-ui,sans-serif;max-width:560px;margin:30px auto;padding:0 16px;color:#0F172A}"
                          "h1{color:#0020C2;font-size:22px}table{width:100%%;border-collapse:collapse}td{padding:8px 0;border-bottom:1px solid #e2e8f0}td:last-child{text-align:right;font-weight:700}"
                          "button{margin-top:18px;padding:10px 20px;border:0;border-radius:999px;background:#0020C2;color:#fff;font-weight:700;cursor:pointer}@media print{button{display:none}}</style>"
                          "<h1>PROCSIS · Comprobante de pago</h1><table><tr><td>Factura</td><td>%s</td></tr><tr><td>Cliente</td><td>%s</td></tr><tr><td>Concepto</td><td>%s</td></tr>"
                          "<tr><td>Valor</td><td>%s</td></tr><tr><td>Descuento%s</td><td>-%s</td></tr><tr><td>Total pagado</td><td>%s</td></tr><tr><td>Fecha de pago</td><td>%s</td></tr></table>"
                          "<button onclick='window.print()'>Descargar / imprimir</button>"
                          % (e(f.numero), e(f.numero), e(f.cliente), e(f.concepto), cop(f.valor), (" (" + e(f.cupon) + ")") if f.cupon else "", cop(f.descuento), cop(tot), e(pagada)))
        r.headers["Cache-Control"] = "no-store"
        return r

    def pub_cotizar():
        tablas()
        d = cuerpo_json()
        if d is None:
            return js({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        if not limite("q:" + ip(), 5, 3600):
            return js({"ok": False, "mensaje": "Ya enviaste varias solicitudes. Intenta más tarde."}, 429)
        nombre, contacto = str(d.get("nombre") or "").strip()[:120], str(d.get("contacto") or "").strip()[:160]
        if len(nombre) < 3 or len(contacto) < 6:
            return js({"ok": False, "mensaje": "Escribe tu nombre y un correo o celular de contacto."}, 400)
        eq = ""
        if str(d.get("equipo") or "").isdigit():
            q = Equipo.query.get(int(d["equipo"]))
            eq = q.nombre if q and q.activo else ""
        eq = eq or str(d.get("interes") or "").strip()[:160]
        db.session.add(Cotizacion(nombre=nombre, contacto=contacto, equipo_nombre=eq, cupon=norm_ref(d.get("cupon")), mensaje=str(d.get("mensaje") or "").strip()[:600], creado=int(time.time())))
        db.session.commit()
        audit("Comercial: cotización recibida", eq)
        return js({"ok": True, "mensaje": "¡Recibimos tu solicitud! Un asesor te contactará pronto."})

    def pub_img(tipo, rid):
        tablas()
        reg = (Banner if tipo == "banner" else Equipo if tipo == "equipo" else Entrada if tipo == "entrada" else None)
        reg = reg.query.get(rid) if reg else None
        uri = (getattr(reg, "foto" if tipo == "equipo" else "imagen", "") or "") if reg and reg.activo else ""
        m = re.match(r"^data:(image/(?:png|jpeg|webp));base64,(.+)$", uri)
        if not m:
            abort(404)
        r = make_response(base64.b64decode(m.group(2)))
        r.headers["Content-Type"] = m.group(1)
        r.headers["Cache-Control"] = "public, max-age=300"
        r.headers["X-Content-Type-Options"] = "nosniff"
        return r

    def pub_json():
        ban, tar, pla, equ = datos_publicos()
        return js({"banners": [{"titulo": b.titulo, "texto": b.texto} for b in ban], "beneficios": [{"titulo": t.titulo, "texto": t.texto} for t in tar],
                   "planes": [{"nombre": p.nombre, "tipo": p.tipo, "estudiantes_max": p.estudiantes_max, "sedes_max": p.sedes_max, "precio_mensual": p.precio_mensual,
                               "descripcion": p.descripcion, "caracteristicas": [x.strip() for x in (p.caracteristicas or "").splitlines() if x.strip()]} for p in pla],
                   "equipos": [{"nombre": q.nombre, "categoria": q.categoria, "precio": q.precio, "disponible": (q.stock or 0) > 0, "beneficio": q.beneficio} for q in equ]})

    # ───── Edu Aura pública: responde SOLO con catálogo, planes y textos públicos (nunca facturas ni datos de colegios) ─────
    def aura_responder(q):
        for r in AuraResp.query.filter_by(activo=True).order_by(AuraResp.id).all():  # respuestas que editan Marketing y TI
            ks = [sin_acentos(k).strip() for k in (r.palabras or "").split(",") if k.strip()]
            if any(k and k in sin_acentos(q) for k in ks):
                return (r.respuesta or "")[:900], [{"t": "Cotizar con un asesor", "h": "#pc-cot"}]
        rutas = rutas_existentes()
        COT = {"t": "Hablar con un asesor", "h": "#pc-cot"}
        if toca(q, "hola", "buenas", "buenos dias", "buen dia", "saludos"):
            return "¡Hola! Soy Edu Aura. Te ayudo con planes, tecnología, pagos, cupones y PQR. ¿Qué necesitas?", []
        if toca(q, "gracias", "chao", "adios"):
            return "¡Con gusto! Si necesitas algo más, aquí estaré.", []
        if toca(q, "pqr", "queja", "reclamo", "peticion", "sugerencia"):
            pqr = next((c for c in ("/pqr", "/radicar-pqr", "/pqr/nueva") if c in rutas), "")
            if pqr:
                return ("Para radicar una PQR usa «Radicar PQR» en la sección Autogestión. Para proteger tu información validamos tu identidad con un PIN a tu correo y algunos datos de la solicitud.",
                        [{"t": "Radicar PQR", "h": pqr}])
            return "Puedes dejar tu petición, queja o reclamo en el formulario «Cotizar o comprar» y un asesor te responderá.", [COT]
        if toca(q, "factura", "pagar", "pago", "pse", "nequi", "comprobante", "mensualidad"):
            return ("Para pagar: en «Pagos, facturas y cupones» escribe el NIT, el código de tu colegio o tu documento, elige tu factura y pulsa Pagar. Cuando el pago se confirme podrás descargar el comprobante.",
                    [{"t": "Ir a pagar", "h": "#pc-pagos"}])
        if toca(q, "cupon", "descuento", "promo", "promocion", "oferta"):
            return ("Los cupones se aplican al consultar tu factura (campo «¿Tienes un cupón?») o al pedir una cotización. Cada cupón tiene fecha de vencimiento y un límite de usos.",
                    [{"t": "Aplicar un cupón", "h": "#pc-pagos"}, COT])
        if toca(q, "computador", "portatil", "equipo", "pc", "lector", "qr", "pos", "combo", "tecnologia", "tienda", "comprar", "elegir"):
            eqs = Equipo.query.filter_by(activo=True).order_by(Equipo.categoria, Equipo.nombre).all()
            if not eqs:
                return "Estamos actualizando el catálogo de tecnología. Déjanos tu solicitud y un asesor te cotiza.", [COT]
            pal = {w for w in re.findall(r"[a-z0-9]+", sin_acentos(q)) if len(w) >= 4}
            afines = [x for x in eqs if any(w in sin_acentos(x.nombre + " " + x.categoria) for w in pal)] or eqs
            lin = ["• %s (%s): %s, %s" % (x.nombre, x.categoria, cop(x.precio) if x.precio else "cotizar", "disponible" if (x.stock or 0) > 0 else "consultar disponibilidad") for x in afines[:4]]
            return ("Esto es lo que tenemos:\n" + "\n".join(lin) + "\nCuéntame para qué lo necesitas (aula, portería, secretaría) y un asesor te orienta.", [{"t": "Ver catálogo", "h": "#pc-tienda"}, {"t": "Cotizar con cupón", "h": "#pc-cot"}])
        if toca(q, "plan", "planes", "tarifa", "precio", "costo", "cuanto", "capacidad", "sede", "licencia"):
            ps = Plan.query.filter_by(activo=True, aprobado=True).order_by(Plan.tipo, Plan.id).all()
            if not ps:
                return "Estamos actualizando nuestros planes. Déjanos tu solicitud y te cotizamos.", [COT]
            lin = ["• %s: %s, %s, %s" % (x.nombre, ("hasta %s estudiantes" % format(x.estudiantes_max, ",").replace(",", ".")) if x.estudiantes_max else "estudiantes sin límite",
                                          ("%d sede%s" % (x.sedes_max, "" if x.sedes_max == 1 else "s")) if x.sedes_max else "multi-sede ilimitada",
                                          (cop(x.precio_mensual) + " al mes") if x.precio_mensual else "tarifa a cotizar") for x in ps]
            return "Nuestros planes:\n" + "\n".join(lin), [{"t": "Ver planes", "h": "#pc-planes"}, COT]
        if toca(q, "matricula", "requisito", "inscripcion", "cupo", "admision"):
            return ("Los requisitos de matrícula los define cada colegio. Consúltalos con la secretaría de tu institución o ingresa por «Ingresar colegio». No tengo acceso a datos de colegios.", [{"t": "Ingresar colegio", "h": "/login"}])
        if toca(q, "contrasena", "clave", "acceso", "ingresar", "bloqueo", "bloquear", "mfa", "pin", "codigo"):
            return ("Entra a «Ingresar colegio» y usa «¿Olvidaste tu contraseña?». Te enviamos un código de 6 dígitos al correo registrado; vence en 5 minutos y sirve una sola vez.", [{"t": "Ingresar colegio", "h": "/login"}])
        if toca(q, "soporte", "ayuda", "agendar", "asesor", "contacto", "llamar"):
            ag = next((c for c in ("/agendar-soporte", "/soporte/agendar", "/agendar_soporte") if c in rutas), "")
            return ("Puedes agendar soporte desde la sección Autogestión." if ag else "Déjanos tu solicitud y un asesor te contacta."), ([{"t": "Agendar soporte", "h": ag}] if ag else [COT])
        if toca(q, "datos", "habeas", "privacidad", "ley", "1581", "seguridad"):
            return "Tratamos los datos personales conforme a la Ley 1581 de 2012 (Habeas Data). No compartimos información de colegios en este asistente.", []
        return ("No tengo esa respuesta con la información pública. Puedo ayudarte con planes, tecnología, pagos, cupones o PQR, o pasar tu consulta a un asesor.", [COT])

    def pub_aura():
        tablas()
        d = cuerpo_json()
        if d is None:
            return js({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        if not limite("a:" + ip(), 20, 600):
            return js({"ok": False, "mensaje": "Has hecho muchas preguntas seguidas. Espera unos minutos."}, 429)
        q = str(d.get("pregunta") or "").strip()[:300]
        if len(q) < 2:
            return js({"ok": False, "mensaje": "Escribe tu pregunta."}, 400)
        try:
            txt, enl = aura_responder(q)
        except Exception as ex:
            print("COMERCIAL aura:", repr(ex)[:140], flush=True)
            try:
                db.session.rollback()
            except Exception:
                pass
            return js({"ok": False, "mensaje": "No pude responder en este momento. Intenta de nuevo."}, 500)
        return js({"ok": True, "respuesta": txt, "enlaces": enl})

    def pub_blog(rid):
        tablas()
        b = Entrada.query.get(rid)
        if b is None or not b.activo:
            abort(404)
        pars = "".join("<p>%s</p>" % e(x.strip()) for x in (b.cuerpo or b.resumen or "").splitlines() if x.strip())
        base = bloque("/publico/inicio")
        art = ('<style>%s</style><article class="pc-root" style="max-width:760px"><a href="/publico/inicio#pc-blog" class="pc-s">← Volver a novedades</a>'
               '<span class="pc-tag">%s</span><h1 style="color:#0020C2;margin:10px 0 4px">%s</h1><div class="pc-dt">%s</div>%s</article>'
               % (CSS_PUB, e(b.categoria), e(b.titulo), e(fecha_larga(b.fecha)), pars))
        return envolver_publico(base["barra"] + art, b.titulo)

    # ───── Fase 3: confirmación de pago de la pasarela (firmada) ─────
    def pub_webhook():
        tablas()
        secreto = os.environ.get("PAGO_SECRET", "")
        if not secreto:
            return js({"ok": False, "mensaje": "Pasarela sin configurar."}, 503)
        if not limite("w:" + ip(), 120, 600):
            return js({"ok": False}, 429)
        d = request.get_json(silent=True) if request.is_json else None
        if not isinstance(d, dict):
            return js({"ok": False}, 400)
        num, val, est = norm_ref(d.get("numero")), str(d.get("valor") or "").strip(), str(d.get("estado") or "").strip().upper()
        esperada = hmac.new(secreto.encode("utf-8"), ("%s|%s|%s" % (num, val, est)).encode("utf-8"), hashlib.sha256).hexdigest()
        if not igual(str(d.get("firma") or "").lower(), esperada):
            audit("Comercial: confirmación de pago con firma inválida", "%s · %s" % (num, ip()))
            return js({"ok": False}, 401)
        f = Factura.query.filter_by(numero=num).first()
        if f is None:
            return js({"ok": False}, 404)
        if est != "APROBADA":
            return js({"ok": True, "ignorado": True})
        if f.estado == "PAGADA":
            return js({"ok": True, "repetido": True})
        total = max(0, (f.valor or 0) - (f.descuento or 0))
        if not val.isdigit() or int(val) != total:
            audit("Comercial: pago con valor distinto al de la factura", "%s · recibido %s · esperado %s" % (num, val, total))
            return js({"ok": False, "mensaje": "El valor no coincide."}, 400)
        f.estado, f.pagada_ts = "PAGADA", int(time.time())
        db.session.commit()
        audit("Comercial: pago confirmado por pasarela", "%s · %s" % (num, total))
        return js({"ok": True})

    reglas = (
        ("/comercial", "com_admin", com_admin, ["GET"]),
        ("/comercial/guardar/<sec>", "com_guardar", com_guardar, ["POST"]),
        ("/comercial/accion/<sec>/<int:rid>/<accion>", "com_accion", com_accion, ["POST"]),
        ("/publico/inicio", "com_inicio", com_inicio, ["GET"]),
        ("/publico/comercial.json", "com_json", pub_json, ["GET"]),
        ("/publico/pagos/consultar", "com_consultar", pub_consultar, ["POST"]),
        ("/publico/pagos/cupon", "com_cupon", pub_cupon, ["POST"]),
        ("/publico/pagos/ir/<numero>", "com_ir", pub_ir, ["GET"]),
        ("/publico/pagos/comprobante/<numero>", "com_comprobante", pub_comprobante, ["GET"]),
        ("/publico/cotizar", "com_cotizar", pub_cotizar, ["POST"]),
        ("/publico/img/<tipo>/<int:rid>", "com_img", pub_img, ["GET"]),
        ("/publico/aura", "com_aura", pub_aura, ["POST"]),
        ("/publico/blog/<int:rid>", "com_blog", pub_blog, ["GET"]),
        ("/publico/pagos/webhook", "com_webhook", pub_webhook, ["POST"]),
    )
    for ruta, nombre, fn, ms in reglas:
        app.add_url_rule(ruta, nombre, fn, methods=ms)
    app.after_request(inyectar)

    # Que el guardián de login de la app deje pasar /publico/* sin sesión (si la app usa ruta_publica()).
    orig_pub = G.get("ruta_publica")
    if callable(orig_pub):
        def ruta_publica_v2():
            return bool((request.path or "").startswith("/publico/")) or bool(orig_pub())
        G["ruta_publica"] = ruta_publica_v2
        print("COMERCIAL: /publico/* agregado a ruta_publica().", flush=True)
    else:
        print("COMERCIAL: aviso · tu app no tiene ruta_publica(); si /publico/inicio te manda al login, agrega el prefijo /publico/ a tus rutas públicas.", flush=True)
    try:
        with app.app_context():
            tablas()
    except Exception as ex:
        print("COMERCIAL: las tablas se crearán en la primera visita:", repr(ex)[:120], flush=True)


CSS_PUB = """
.pc-root{max-width:1100px;margin:0 auto;padding:22px 16px 50px;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;color:#0F172A;background:#fff}
.pc-root *{box-sizing:border-box}
.pc-car{position:relative;border-radius:28px;overflow:hidden;min-height:190px;background:#0020C2;margin:0 0 26px}
.pc-sl{display:none;padding:38px 34px 46px;color:#fff;background:linear-gradient(110deg,#0020C2,#00005e);background-size:cover;background-position:center;min-height:190px}
.pc-sl.on{display:block}.pc-sl h2{margin:0 0 8px;font-size:clamp(22px,4vw,32px)}.pc-sl p{margin:0 0 14px;font-size:16px;opacity:.95;max-width:640px}
.pc-cta{display:inline-block;padding:9px 18px;border-radius:999px;background:#fff;color:#0020C2;font-weight:800;text-decoration:none;font-size:14px}
.pc-dots{position:absolute;left:0;right:0;bottom:14px;display:flex;gap:7px;justify-content:center}
.pc-dots i{width:9px;height:9px;border-radius:50%;background:rgba(255,255,255,.45);cursor:pointer}.pc-dots i.on{background:#00FF66}
.pc-h{font-size:clamp(20px,3vw,26px);margin:34px 0 14px;color:#0020C2;letter-spacing:-.01em}
.pc-gr{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px}
.pc-c{background:#F8FAFC;border:1px solid #e2e8f0;border-radius:20px;padding:18px;display:flex;flex-direction:column;gap:6px}
.pc-c h3{margin:6px 0 0;font-size:17px;color:#0F172A}.pc-c p{margin:0;color:#475569;font-size:14px;line-height:1.5}
.pc-c img{width:100%;height:150px;object-fit:cover;border-radius:14px;background:#e2e8f0}.pc-ni{height:12px}
.pc-tag{align-self:flex-start;font-size:11.5px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:#0020C2;background:#e8edff;border-radius:999px;padding:4px 10px}
.pc-pl.corp{border:1.5px solid #0020C2;box-shadow:0 0 0 2px rgba(0,255,102,.35)}
.pc-pr{font-size:22px;font-weight:800;color:#0020C2;margin:6px 0}.pc-pr small{font-size:13px;color:#64748b;font-weight:600}
.pc-c ul{margin:4px 0 8px;padding-left:18px;color:#334155;font-size:14px;line-height:1.6}
.pc-bn{font-size:13px;font-weight:700;color:#067647;background:#ecfdf3;border-radius:10px;padding:6px 10px}
.pc-b{display:inline-block;margin-top:auto;border:0;border-radius:999px;padding:11px 20px;background:#0020C2;color:#fff;font-weight:800;font-size:14px;text-decoration:none;cursor:pointer;text-align:center}
.pc-b.off{background:#64748b}.pc-b.sec{background:#fff;color:#0020C2;border:1px solid #0020C2}
.pc-box{background:#F8FAFC;border:1px solid #e2e8f0;border-radius:20px;padding:18px}
.pc-s{color:#64748b;font-size:14px;margin:0 0 12px}.pc-row{display:flex;gap:10px;flex-wrap:wrap}
.pc-root input,.pc-root select,.pc-root textarea{padding:11px 13px;border:1px solid #cbd5e1;border-radius:12px;font:inherit;font-size:15px;background:#fff;color:#0F172A;min-width:0}
.pc-root input:focus,.pc-root select:focus,.pc-root textarea:focus{outline:2px solid rgba(0,255,102,.55);border-color:#0020C2}
.pc-row input{flex:1;min-width:200px}.pc-fm{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-bottom:10px}.pc-box textarea{width:100%}
.pc-f{background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:14px;margin-top:12px}.pc-f b{color:#0020C2}
.pc-f .pc-row{margin-top:10px}.pc-er{color:#b91c1c;font-size:14px;margin-top:8px}.pc-okk{color:#067647;font-size:14px;font-weight:700;margin-top:8px}
.pc-au{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;margin:0 0 12px}
.pc-au a{padding:11px 20px;border-radius:999px;border:1px solid #0020C2;color:#0020C2;font-weight:800;text-decoration:none;font-size:14px;background:#fff}
.pc-h{scroll-margin-top:84px}
.pc-dt{font-size:12.5px;color:#64748b}
.pc-bar{position:sticky;top:0;z-index:2147482000;display:flex;align-items:center;gap:14px;padding:10px 20px;background:#fff;border-bottom:1px solid #e2e8f0;box-shadow:0 2px 12px rgba(0,32,194,.08);font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;box-sizing:border-box}
.pc-lg{display:flex;align-items:center;gap:8px;text-decoration:none;color:#0020C2;font-weight:800;letter-spacing:.04em;white-space:nowrap}
.pc-lg b{display:grid;place-items:center;width:32px;height:32px;border-radius:10px;background:#0020C2;color:#fff}
.pc-bar nav{display:flex;gap:4px;flex:1;overflow-x:auto;scrollbar-width:none}.pc-bar nav::-webkit-scrollbar{display:none}
.pc-bar nav a{padding:8px 12px;border-radius:999px;text-decoration:none;color:#0F172A;font-weight:600;font-size:14px;white-space:nowrap}
.pc-bar nav a:hover,.pc-bar nav a:focus{background:#eef2ff;color:#0020C2}
.pc-in{padding:10px 18px;border-radius:999px;background:#0020C2;color:#fff;text-decoration:none;font-weight:800;font-size:14px;white-space:nowrap}
@media(max-width:560px){.pc-lg span{display:none}.pc-bar{padding:8px 12px;gap:8px}}
.pc-fl{position:fixed;bottom:18px;z-index:2147482500;font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;box-sizing:border-box}
.pc-fl *{box-sizing:border-box}.pc-l{left:16px}.pc-r{right:16px}
.pc-fb{width:56px;height:56px;border-radius:50%;border:0;cursor:pointer;display:grid;place-items:center;background:#0020C2;color:#fff;box-shadow:0 8px 24px rgba(0,32,194,.35)}
.pc-fb:focus-visible{outline:3px solid #00FF66;outline-offset:2px}
#pc-i1{width:auto;border-radius:999px;padding:0 20px;background:#00FF66;color:#0020C2;font-weight:800;font-size:14px;box-shadow:0 8px 24px rgba(0,255,102,.45)}
.pc-pn{position:absolute;bottom:68px;width:min(340px,calc(100vw - 32px));background:#fff;border:1px solid #e2e8f0;border-radius:20px;box-shadow:0 20px 50px rgba(10,42,143,.25);overflow:hidden;color:#0F172A}
.pc-l .pc-pn{left:0}.pc-r .pc-pn{right:0}.pc-pn[hidden]{display:none}
.pc-pn h4{margin:0;padding:14px 16px;background:#0020C2;color:#fff;font-size:15px}
.pc-ar{display:flex;flex-direction:column;gap:8px;padding:14px}
.pc-ar button{border:1px solid #0020C2;background:#fff;color:#0020C2;border-radius:12px;padding:10px 12px;font-weight:700;font-size:14px;cursor:pointer;text-align:left}
.pc-ar button[aria-pressed="true"]{background:#0020C2;color:#fff}
.pc-ah{padding:14px 16px;background:linear-gradient(110deg,#0020C2,#00005e);color:#fff;display:flex;flex-direction:column;border-bottom:3px solid #00FF66}.pc-ah small{opacity:.85}
.pc-ms{max-height:290px;min-height:90px;overflow:auto;padding:12px;display:flex;flex-direction:column;gap:8px}
.pc-mb{padding:9px 12px;border-radius:14px;font-size:14px;line-height:1.45;white-space:pre-line;max-width:90%}
.pc-mb.u{align-self:flex-end;background:#0020C2;color:#fff}.pc-mb.a{background:#F8FAFC;border:1px solid #e2e8f0}
.pc-lnk{display:inline-block;margin-top:6px;color:#0020C2;font-weight:800;text-decoration:underline}
.pc-ch{display:flex;flex-wrap:wrap;gap:6px;padding:0 12px 8px}.pc-ch button{border:1px solid #cbd5e1;background:#fff;border-radius:999px;padding:6px 10px;font-size:12.5px;cursor:pointer;color:#0F172A}
.pc-qr{display:flex;gap:8px;padding:0 12px 8px}.pc-qr input{flex:1;padding:10px 12px;border:1px solid #cbd5e1;border-radius:12px;font:inherit;font-size:14px;min-width:0;color:#0F172A;background:#fff}
.pc-qr button{border:0;border-radius:12px;padding:0 14px;background:#0020C2;color:#fff;font-weight:800;cursor:pointer}
.pc-ft{margin:0;padding:0 14px 12px;font-size:11.5px;color:#64748b}
html.pc-hc{filter:invert(1) hue-rotate(180deg)}
html.pc-hc img,html.pc-hc video,html.pc-hc .pc-sl,html.pc-hc .pc-fl{filter:invert(1) hue-rotate(180deg)}
html.pc-lk a{text-decoration:underline!important;font-weight:800}
"""

JS_PUB = r"""(function(){
var R=document.getElementById('pc-root');if(!R)return;
function $(i){return document.getElementById(i);}
function cop(n){return '$'+String(Math.round(n||0)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');}
function post(u,d){return fetch(u,{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)}).then(function(r){return r.json().catch(function(){return{ok:false,mensaje:'Respuesta no válida.'};});}).catch(function(){return{ok:false,mensaje:'No hay conexión. Inténtalo de nuevo.'};});}
var car=$('pc-car');
if(car){var sl=car.querySelectorAll('.pc-sl'),dt=car.querySelectorAll('.pc-dots i'),k=0,t;
function go(n){sl[k].classList.remove('on');if(dt[k])dt[k].classList.remove('on');k=(n+sl.length)%sl.length;sl[k].classList.add('on');if(dt[k])dt[k].classList.add('on');}
function play(){clearInterval(t);if(sl.length>1)t=setInterval(function(){go(k+1);},5500);}
Array.prototype.forEach.call(dt,function(d,i){d.onclick=function(){go(i);play();};});
car.onmouseenter=function(){clearInterval(t);};car.onmouseleave=play;play();}
Array.prototype.forEach.call(R.querySelectorAll('[data-eq],[data-plan]'),function(a){a.addEventListener('click',function(){var s=$('pc-e');if(!s)return;
var v=a.getAttribute('data-eq');if(v){s.value=v;}else{var p=a.getAttribute('data-plan');var o=document.createElement('option');o.value='';o.setAttribute('data-i',p);o.textContent=p;s.appendChild(o);s.selectedIndex=s.options.length-1;}});});
function nodo(tag,cls,txt){var n=document.createElement(tag);if(cls)n.className=cls;if(txt!=null)n.textContent=txt;return n;}
function pintar(f,ref,cont){var d=nodo('div','pc-f');var h=nodo('div');var b=nodo('b',null,'Factura '+f.numero);h.appendChild(b);d.appendChild(h);
d.appendChild(nodo('div',null,f.cliente+' · '+f.concepto));
var v=f.descuento?('Valor '+cop(f.valor)+' − descuento '+cop(f.descuento)+(f.cupon?' ('+f.cupon+')':'')+' = '):'Total a pagar: ';
d.appendChild(nodo('div',null,v+cop(f.total)+(f.vence?' · vence '+f.vence:'')));
d.appendChild(nodo('div',null,'Estado: '+(f.estado==='PAGADA'?'Pagada':'Pendiente')));
var q=ref;
if(f.estado==='PENDIENTE'){
 if(!f.cupon){var row=nodo('div','pc-row');var inp=nodo('input');inp.placeholder='¿Tienes un cupón?';inp.maxLength=40;var bt=nodo('button','pc-b sec','Aplicar cupón');bt.type='button';var er=nodo('div','pc-er');
  bt.onclick=function(){er.textContent='';post('/publico/pagos/cupon',{numero:f.numero,ref:q,codigo:inp.value}).then(function(r){if(!r.ok){er.textContent=r.mensaje;return;}buscar();});};
  row.appendChild(inp);row.appendChild(bt);d.appendChild(row);d.appendChild(er);}
 var pg=nodo('a','pc-b','Pagar con PSE, Nequi o tarjeta');pg.href='/publico/pagos/ir/'+encodeURIComponent(f.numero)+'?ref='+encodeURIComponent(q);pg.style.marginTop='12px';d.appendChild(pg);
}else{var cp=nodo('a','pc-b','Descargar comprobante');cp.href='/publico/pagos/comprobante/'+encodeURIComponent(f.numero)+'?ref='+encodeURIComponent(q);cp.style.marginTop='12px';d.appendChild(cp);}
cont.appendChild(d);}
function buscar(){var ref=$('pc-ref').value.trim(),c=$('pc-fac');c.textContent='';if(ref.length<4){c.appendChild(nodo('div','pc-er','Escribe al menos 4 caracteres.'));return;}
post('/publico/pagos/consultar',{ref:ref}).then(function(r){c.textContent='';if(!r.ok){c.appendChild(nodo('div','pc-er',r.mensaje));return;}
if(!r.facturas.length){c.appendChild(nodo('div','pc-s',r.mensaje));return;}r.facturas.forEach(function(f){pintar(f,ref,c);});});}
var bb=$('pc-bus');if(bb)bb.onclick=buscar;var rr=$('pc-ref');if(rr)rr.onkeydown=function(e){if(e.key==='Enter'){buscar();}};
var be=$('pc-env');if(be)be.onclick=function(){var s=$('pc-e'),o=$('pc-ok'),opt=s.options[s.selectedIndex];o.className='pc-s';o.textContent='';
post('/publico/cotizar',{nombre:$('pc-n').value,contacto:$('pc-c').value,equipo:s.value,interes:(!s.value&&opt&&opt.getAttribute('data-i'))||'',cupon:$('pc-k').value,mensaje:$('pc-m').value}).then(function(r){o.className=r.ok?'pc-okk':'pc-er';o.textContent=r.mensaje;if(r.ok){$('pc-n').value='';$('pc-c').value='';$('pc-k').value='';$('pc-m').value='';}});};
var A={z:100,hc:0,lk:0};
try{var g=JSON.parse(localStorage.getItem('pc_a11y')||'{}');if(typeof g.z==='number'&&g.z>=90&&g.z<=160)A.z=g.z;A.hc=g.hc?1:0;A.lk=g.lk?1:0;}catch(x){}
function apA(){document.body.style.zoom=A.z===100?'':String(A.z/100);var h=document.documentElement;h.classList.toggle('pc-hc',!!A.hc);h.classList.toggle('pc-lk',!!A.lk);
 Array.prototype.forEach.call(R.ownerDocument.querySelectorAll('[data-a="hc"],[data-a="lk"]'),function(b){b.setAttribute('aria-pressed',(b.getAttribute('data-a')==='hc'?A.hc:A.lk)?'true':'false');});
 try{localStorage.setItem('pc_a11y',JSON.stringify(A));}catch(x){}}
apA();
var ab=$('pc-a1'),ap=$('pc-ap');
if(ab&&ap){ab.onclick=function(){var o=ap.hidden;ap.hidden=!o;ab.setAttribute('aria-expanded',o?'true':'false');};
 ap.addEventListener('click',function(ev){var a=ev.target.getAttribute&&ev.target.getAttribute('data-a');if(!a)return;
  if(a==='mas')A.z=Math.min(160,A.z+15);else if(a==='menos')A.z=Math.max(90,A.z-15);else if(a==='hc')A.hc=A.hc?0:1;else if(a==='lk')A.lk=A.lk?0:1;else if(a==='rs'){A.z=100;A.hc=0;A.lk=0;}apA();});}
var ib=$('pc-i1'),ip=$('pc-ip');
if(ib&&ip){var ms=$('pc-ms'),qi=$('pc-q'),ini=false;
 function msg(t,c){var d=nodo('div','pc-mb '+c,t);ms.appendChild(d);ms.scrollTop=ms.scrollHeight;return d;}
 ib.onclick=function(){var o=ip.hidden;ip.hidden=!o;ib.setAttribute('aria-expanded',o?'true':'false');
  if(o){if(!ini){ini=true;msg('¡Hola! Soy Edu Aura. Respondo con la información pública de PROCSIS: planes, tecnología, pagos, cupones y PQR. ¿En qué te ayudo?','a');}qi.focus();}};
 function enviar(t){t=(t||'').trim();if(t.length<2)return;qi.value='';msg(t,'u');var w=msg('…','a');
  post('/publico/aura',{pregunta:t}).then(function(r){w.textContent=r.ok?r.respuesta:(r.mensaje||'No pude responder ahora.');
   if(r.ok&&r.enlaces){r.enlaces.forEach(function(l){if(/^(#|\/(?!\/))/.test(l.h)){var a=nodo('a','pc-lnk',l.t);a.href=l.h;w.appendChild(document.createElement('br'));w.appendChild(a);}});}
   ms.scrollTop=ms.scrollHeight;});}
 $('pc-qs').onclick=function(){enviar(qi.value);};qi.onkeydown=function(ev){if(ev.key==='Enter'){enviar(qi.value);}};
 Array.prototype.forEach.call(ip.querySelectorAll('.pc-ch button'),function(b){b.onclick=function(){enviar(b.textContent);};});}
document.addEventListener('keydown',function(ev){if(ev.key==='Escape'){if(ap&&!ap.hidden){ap.hidden=true;}if(ip&&!ip.hidden){ip.hidden=true;}}});
})();"""

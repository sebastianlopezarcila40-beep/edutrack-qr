# -*- coding: utf-8 -*-
"""
aplicar_inteligencia_competitiva.py

Agrega a app.py el flujo "Competencia / Cliente incógnito" SIN crear un área nueva,
repartido entre las áreas que ya existen:

  Marketing   -> Estudio de competencia (registro + "Enviar hallazgos") y Battlecards
  Calidad     -> Auditoría de atención (checklist 1-5) y Planes de mejora automáticos
  Soluciones  -> Bandeja de Propuestas de mejora de producto (pasa a TI/Desarrollo)
  Ventas      -> Battlecards (consulta) y sus Planes de mejora
  Gerencia    -> Resumen (solo lectura)

Todo vive bajo /inteligencia/... Usa lo que ya tienes: permisos por área, campana de
notificaciones, tickets de TI (ti_tickets) y el menú lateral de cada rol.

Seguridad del cambio:
  * Solo agrega código (un bloque nuevo antes de `if __name__ == "__main__":`).
  * Crea tablas NUEVAS con CREATE TABLE IF NOT EXISTS (intel_*). No toca ni borra
    ninguna tabla ni dato existente.
  * Hace respaldo app.py.bak_inteligencia_competitiva, valida sintaxis antes de escribir,
    conserva los saltos de línea y no se aplica dos veces.

Uso:   python aplicar_inteligencia_competitiva.py        (usa ./app.py)
       python aplicar_inteligencia_competitiva.py ruta/app.py
"""
import ast
import os
import shutil
import sys

MARCA = "[PROCSIS-INTELIGENCIA-COMPETITIVA]"
RUTA = sys.argv[1] if len(sys.argv) > 1 else "app.py"
ANCLA = 'if __name__ == "__main__":'

BLOQUE = r'''# ══════════════════════════════════════════════════════════════════
# [PROCSIS-INTELIGENCIA-COMPETITIVA]
# Competencia / Cliente incógnito repartido entre áreas existentes:
#   Marketing (estudio + battlecards) -> Calidad (auditoría + planes de mejora)
#   -> Soluciones (propuestas de producto) -> TI/Desarrollo (ticket) -> Marketing/Ventas.
#   Gerencia: resumen en solo lectura.
# ══════════════════════════════════════════════════════════════════
_IC_INIT = {"ok": False}
_IC_GER = ("Gerente", "Superadmin", "Administrador")
_IC_AREAS_PLAN = {"Comercial": "Ventas", "Soporte": "Soporte", "Cobranza": "Cobranza",
                  "Marketing": "Marketing y Comunicaciones", "Soluciones": "Soluciones",
                  "Tecnologia": "Tecnología y Soporte de TI"}
_IC_OPERAN = ("Marketing", "Calidad", "Soluciones", "Comercial", "Soporte", "Cobranza", "Tecnologia", "Desarrollador")
_IC_VER = {
    "competencia": ("Marketing", "Calidad", "Soluciones"),
    "battlecards": ("Comercial", "Marketing", "Calidad", "Soluciones"),
    "auditoria": ("Calidad",),
    "planes": ("Calidad", "Comercial", "Soporte", "Cobranza", "Marketing", "Soluciones", "Tecnologia"),
    "propuestas": ("Soluciones", "Marketing", "Calidad", "Tecnologia", "Desarrollador"),
    "resumen": (),
}
_IC_ESCRIBE = {
    "competencia": ("Marketing",),
    "battlecards": ("Marketing",),
    "auditoria": ("Calidad",),
    "planes": ("Calidad", "Comercial", "Soporte", "Cobranza", "Marketing", "Soluciones", "Tecnologia"),
    "propuestas": ("Soluciones", "Tecnologia", "Desarrollador"),
}
_IC_TABS = (
    ("competencia", "/inteligencia/competencia", "🕵️ Competencia"),
    ("battlecards", "/inteligencia/battlecards", "⚔️ Battlecards"),
    ("auditoria", "/inteligencia/auditoria", "✅ Auditoría de atención"),
    ("planes", "/inteligencia/planes", "🛠️ Planes de mejora"),
    ("propuestas", "/inteligencia/propuestas", "💡 Propuestas de producto"),
    ("resumen", "/inteligencia/resumen", "📊 Resumen"),
)
_IC_CANALES = ("Cotización por correo", "Chat / WhatsApp", "Llamada", "Formulario web", "Presencial / Demo")
_IC_PRODUCTOS = ("EduTrack", "TiendaTrack", "Otro / general")
_IC_ESTADOS_PLAN = ("Abierto", "En proceso", "Cerrado")
_IC_APROBADA = "Aprobada → Desarrollo/TI"
_IC_ESTADOS_PROP = ("Nueva", "En análisis", _IC_APROBADA, "Descartada", "Implementada")
_IC_UMBRAL_BAJO = 2      # un criterio <= 2 abre un Plan de Mejora automático
_IC_CSS = ("<style>.ic-k{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:10px 0}"
           ".ic-k div{background:#f0fdfa;border:1px solid #99f6e4;border-radius:12px;padding:12px;text-align:center;font-size:12px;color:#475569}"
           ".ic-k b{display:block;font-size:24px;color:#0f766e}.ic-bar{background:#14b8a6;height:12px;border-radius:6px}"
           ".ic-sc{display:inline-block;padding:2px 8px;border-radius:999px;font-weight:700;font-size:12px}"
           ".ic-lo{background:#fee2e2;color:#991b1b}.ic-md{background:#fef3c7;color:#92400e}.ic-hi{background:#dcfce7;color:#166534}"
           ".ic-btn2{background:#fff;border:1px solid #0f766e;color:#0f766e;border-radius:8px;padding:5px 10px;cursor:pointer;font-weight:700;font-size:12px}"
           "</style>")


def _ic_init():
    if _IC_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    T = "TEXT DEFAULT ''"
    ddls = (
        "CREATE TABLE IF NOT EXISTS intel_competencia (id {pk}, empresa {T}, producto {T}, precio {T}, tiempo_respuesta {T}, canal {T}, "
        "funciones {T}, ventaja_nuestra {T}, evidencia_url {T}, observaciones {T}, fecha_eval {T}, estado VARCHAR(20) DEFAULT 'Borrador', "
        "enviado_en VARCHAR(20) DEFAULT '', creado_por VARCHAR(80) DEFAULT '', creado_en VARCHAR(20) DEFAULT '')",
        "CREATE TABLE IF NOT EXISTS intel_battlecards (id {pk}, competidor {T}, precio_ref {T}, somos_mejores {T}, ellos_mejores {T}, "
        "objecion {T}, actualizado_por VARCHAR(80) DEFAULT '', actualizado_en VARCHAR(20) DEFAULT '')",
        "CREATE TABLE IF NOT EXISTS intel_auditoria (id {pk}, tipo VARCHAR(20) DEFAULT '', sujeto {T}, area VARCHAR(30) DEFAULT '', canal {T}, "
        "amabilidad INTEGER DEFAULT 0, claridad INTEGER DEFAULT 0, rapidez INTEGER DEFAULT 0, promedio {T}, comentarios {T}, "
        "creado_por VARCHAR(80) DEFAULT '', creado_en VARCHAR(20) DEFAULT '')",
        "CREATE TABLE IF NOT EXISTS intel_planes (id {pk}, auditoria_id INTEGER DEFAULT 0, area_responsable VARCHAR(30) DEFAULT '', titulo {T}, "
        "detalle {T}, estado VARCHAR(20) DEFAULT 'Abierto', nota {T}, creado_por VARCHAR(80) DEFAULT '', creado_en VARCHAR(20) DEFAULT '', "
        "actualizado_en VARCHAR(20) DEFAULT '', cerrado_en VARCHAR(20) DEFAULT '')",
        "CREATE TABLE IF NOT EXISTS intel_propuestas (id {pk}, origen VARCHAR(20) DEFAULT '', titulo {T}, detalle {T}, producto {T}, "
        "estado VARCHAR(30) DEFAULT 'Nueva', decision_nota {T}, ticket_ti INTEGER DEFAULT 0, creado_por VARCHAR(80) DEFAULT '', "
        "creado_en VARCHAR(20) DEFAULT '', actualizado_en VARCHAR(20) DEFAULT '')",
    )
    for d in ddls:
        try:
            db.session.execute(text(d.replace("{pk}", pk).replace("{T}", T)))
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            print("ic init:", e)
    _IC_INIT["ok"] = True


def _ic_ts():
    return ahora().strftime("%Y-%m-%d %H:%M:%S")


def _ic_acceso(clave):
    if not requiere_login():
        return redirect("/login")
    r = rol_actual()
    if r in _IC_GER or r in _IC_VER.get(clave, ()):
        return None
    return acceso_denegado("Este módulo no está habilitado para su rol.")


def _ic_puede(clave):
    r = rol_actual()
    return r == "Superadmin" or r in _IC_ESCRIBE.get(clave, ())


def _ic_flash(tipo, texto):
    session["ti_msg"] = (tipo, texto)


def _ic_post_ok(clave):
    """None si el POST es válido; si no, la respuesta a devolver."""
    if not _ic_puede(clave):
        return acceso_denegado("Su rol solo puede consultar este módulo.")
    if not _ti_csrf_ok():
        _ic_flash("err", "La sesión del formulario venció. Recargue la página e intente de nuevo.")
        return redirect(request.path)
    return None


def _ic_t(nombre, n=200):
    return " ".join((request.form.get(nombre) or "").split())[:n]


def _ic_tl(nombre, n=1500):
    return (request.form.get(nombre) or "").replace("\r", "").strip()[:n]


def _ic_n(nombre):
    try:
        return max(1, min(5, int(request.form.get(nombre) or 0)))
    except Exception:
        return 0


def _ic_int(nombre):
    try:
        return int(request.form.get(nombre) or 0)
    except Exception:
        return 0


def _ic_opts(seq, sel=""):
    return "".join("<option%s>%s</option>" % (" selected" if x == sel else "", _esc(x)) for x in seq)


def _ic_url(u):
    u = (u or "").strip()[:400]
    return u if u.lower().startswith(("http://", "https://")) else ""


def _ic_score(v):
    try:
        f = float(v)
    except Exception:
        return "<span class='ic-sc ic-md'>—</span>"
    c = "ic-lo" if f < 3 else ("ic-md" if f < 4 else "ic-hi")
    return "<span class='ic-sc %s'>%.1f</span>" % (c, f)


def _ic_pagina(titulo, cuerpo):
    rol = rol_actual()
    msg = session.pop("ti_msg", None)
    aviso = ("<div class='%s'>%s</div>" % ("ti-ok" if msg[0] == "ok" else "ti-err", _esc(msg[1]))) if isinstance(msg, (list, tuple)) and len(msg) == 2 else ""
    tabs = "".join("<a class='%s' href='%s'>%s</a>" % ("on" if request.path.startswith(h) else "", h, n)
                   for k, h, n in _IC_TABS if rol in _IC_GER or rol in _IC_VER.get(k, ()))
    return page(titulo, shell(_TI_CSS + _IC_CSS + "<div class='ti'><div class='ti-tabs'>%s</div>%s%s</div>" % (tabs, aviso, cuerpo)))


def _ic_propuesta(origen, titulo, detalle, producto, usuario):
    _c48_x("INSERT INTO intel_propuestas (origen, titulo, detalle, producto, estado, creado_por, creado_en, actualizado_en) "
           "VALUES (:o,:t,:d,:p,'Nueva',:u,:f,:f)", o=origen, t=titulo[:200], d=detalle[:1500], p=producto, u=usuario, f=_ic_ts())


@app.route("/inteligencia")
def ic_inicio():
    if not requiere_login():
        return redirect("/login")
    r = rol_actual()
    destino = {"Marketing": "/inteligencia/competencia", "Calidad": "/inteligencia/auditoria",
               "Soluciones": "/inteligencia/propuestas", "Comercial": "/inteligencia/battlecards",
               "Tecnologia": "/inteligencia/propuestas", "Desarrollador": "/inteligencia/propuestas"}.get(r)
    if r in _IC_GER:
        destino = "/inteligencia/resumen"
    return redirect(destino or "/inteligencia/planes")


# ───────────── Marketing: estudio de competencia / cliente incógnito ─────────────
@app.route("/inteligencia/competencia", methods=["GET", "POST"])
def ic_competencia():
    g = _ic_acceso("competencia")
    if g is not None:
        return g
    _ic_init()
    usuario = session.get("usuario") or ""
    if request.method == "POST":
        v = _ic_post_ok("competencia")
        if v is not None:
            return v
        accion = request.form.get("accion") or ""
        try:
            if accion == "guardar":
                empresa = _ic_t("empresa", 120)
                if not empresa:
                    _ic_flash("err", "Escriba el nombre de la empresa evaluada.")
                else:
                    _c48_x("INSERT INTO intel_competencia (empresa, producto, precio, tiempo_respuesta, canal, funciones, ventaja_nuestra, evidencia_url, "
                           "observaciones, fecha_eval, estado, creado_por, creado_en) VALUES (:e,:p,:pr,:t,:c,:f,:v,:u,:o,:d,'Borrador',:cp,:ts)",
                           e=empresa, p=_ic_t("producto", 120), pr=_ic_t("precio", 80), t=_ic_t("tiempo_respuesta", 80),
                           c=_ic_t("canal", 60), f=_ic_tl("funciones"), v=_ic_tl("ventaja_nuestra"), u=_ic_url(request.form.get("evidencia_url")),
                           o=_ic_tl("observaciones"), d=_ic_t("fecha_eval", 12) or ahora().strftime("%Y-%m-%d"), cp=usuario, ts=_ic_ts())
                    registrar_auditoria("Competencia: registro creado", "%s · %s" % (usuario, empresa))
                    _ic_flash("ok", "Registro guardado como borrador. Cuando esté listo, envíe los hallazgos a Calidad y Soluciones.")
            elif accion == "enviar":
                fila = _c48_q("SELECT id, empresa, producto, precio, funciones, ventaja_nuestra FROM intel_competencia WHERE id=:i AND estado='Borrador'", i=_ic_int("id"))
                if fila:
                    r = fila[0]
                    _c48_x("UPDATE intel_competencia SET estado='Enviado', enviado_en=:f WHERE id=:i", f=_ic_ts(), i=r[0])
                    detalle = "Empresa: %s\nProducto: %s\nPrecio cotizado: %s\nFunciones que ofrece: %s\nNuestra ventaja: %s" % (
                        r[1], r[2] or "—", r[3] or "—", r[4] or "—", r[5] or "—")
                    _ic_propuesta("Marketing", "Competencia: %s" % r[1], detalle, "Otro / general", usuario)
                    _notif_crear(["Calidad", "Soluciones"], "Nuevo estudio de competencia: %s" % r[1],
                                 "Marketing envió hallazgos para revisión.", "/inteligencia/competencia")
                    registrar_auditoria("Competencia: hallazgos enviados", "%s · %s" % (usuario, r[1]))
                    _ic_flash("ok", "Hallazgos enviados a Calidad y Soluciones (quedó una propuesta en la bandeja de Soluciones).")
                else:
                    _ic_flash("err", "Ese registro ya fue enviado o no existe.")
        except Exception as e:
            print("ic competencia:", e)
            _ic_flash("err", "No se pudo completar la acción. Intente de nuevo.")
        return redirect("/inteligencia/competencia")

    puede = _ic_puede("competencia")
    filas = _c48_q("SELECT id, empresa, producto, precio, tiempo_respuesta, canal, funciones, evidencia_url, estado, fecha_eval, creado_por "
                   "FROM intel_competencia ORDER BY id DESC LIMIT 80")
    csrf = _acc_csrf()
    trs = ""
    for r in filas:
        ev = ("<a href='%s' target='_blank' rel='noopener noreferrer'>Ver evidencia</a>" % _esc(r[7])) if r[7] else "—"
        envio = ""
        if puede and r[8] == "Borrador":
            envio = ("<form method='POST' style='margin:0'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='enviar'>"
                     "<input type='hidden' name='id' value='%d'><button class='ic-btn2'>Enviar hallazgos a Calidad y Soluciones</button></form>") % (csrf, r[0])
        trs += ("<tr><td><b>%s</b><br><small>%s · %s</small></td><td>%s<br><small>%s</small></td><td>%s</td><td><small>%s</small></td><td>%s</td>"
                "<td><span class='ti-b %s'>%s</span></td><td>%s</td></tr>") % (
            _esc(r[1]), _esc(r[9] or ""), _esc(r[10]), _esc(r[2] or "—"), _esc(r[5] or ""), _esc(r[3] or "—"),
            _esc((r[6] or "—")[:160]), _esc(r[4] or "—"), "a" if r[8] == "Enviado" else "p", _esc(r[8]), ev + ("<br>" + envio if envio else ""))
    form = ""
    if puede:
        form = ("<div class='ti-card'><h2>Nuevo registro de competencia</h2><form method='POST'><input type='hidden' name='csrf' value='%s'>"
                "<input type='hidden' name='accion' value='guardar'>"
                "<div class='ti-2'><div><label>Empresa evaluada</label><input name='empresa' required maxlength='120'></div>"
                "<div><label>Producto comparable</label><input name='producto' maxlength='120'></div></div>"
                "<div class='ti-2'><div><label>Precio cotizado</label><input name='precio' maxlength='80' placeholder='Ej. $1.200.000 por año'></div>"
                "<div><label>Tiempo que tardaron en responder</label><input name='tiempo_respuesta' maxlength='80' placeholder='Ej. 2 horas'></div></div>"
                "<div class='ti-2'><div><label>Canal</label><select name='canal'>%s</select></div>"
                "<div><label>Fecha de la evaluación</label><input type='date' name='fecha_eval' value='%s'></div></div>"
                "<label>Funciones y planes que ofrecen</label><textarea name='funciones' rows='3'></textarea>"
                "<label>En qué somos mejores</label><textarea name='ventaja_nuestra' rows='2'></textarea>"
                "<label>Evidencia (enlace a captura o PDF de la cotización)</label><input type='url' name='evidencia_url' placeholder='https://...'>"
                "<label>Observaciones (atención comercial)</label><textarea name='observaciones' rows='2'></textarea>"
                "<p><button class='ti-btn'>Guardar registro</button></p></form></div>") % (csrf, _ic_opts(_IC_CANALES), ahora().strftime("%Y-%m-%d"))
    cuerpo = ("<h1>🕵️ Estudio de competencia / Cliente incógnito</h1>"
              "<p class='sub'>Marketing cotiza como cliente anónimo y registra precios, planes, tiempos de respuesta y atención. "
              "Al enviar los hallazgos, Calidad y Soluciones reciben el aviso y la propuesta.</p>%s"
              "<div class='ti-card'><h2>Registros</h2><div style='overflow-x:auto'><table><tr><th>Empresa</th><th>Producto</th><th>Precio</th><th>Funciones</th>"
              "<th>Respuesta</th><th>Estado</th><th>Evidencia / acción</th></tr>%s</table></div></div>") % (
        form, trs or "<tr><td colspan='7'>Aún no hay registros.</td></tr>")
    return _ic_pagina("Competencia", cuerpo)


# ───────────── Marketing → Ventas: guías comparativas (battlecards) ─────────────
@app.route("/inteligencia/battlecards", methods=["GET", "POST"])
def ic_battlecards():
    g = _ic_acceso("battlecards")
    if g is not None:
        return g
    _ic_init()
    usuario = session.get("usuario") or ""
    if request.method == "POST":
        v = _ic_post_ok("battlecards")
        if v is not None:
            return v
        try:
            comp = _ic_t("competidor", 120)
            if not comp:
                _ic_flash("err", "Escriba el nombre del competidor.")
            else:
                datos = dict(c=comp, p=_ic_t("precio_ref", 120), m=_ic_tl("somos_mejores"), e=_ic_tl("ellos_mejores"),
                             o=_ic_tl("objecion"), u=usuario, f=_ic_ts())
                ya = _c48_q("SELECT id FROM intel_battlecards WHERE lower(competidor)=:k", k=comp.lower())
                if ya:
                    _c48_x("UPDATE intel_battlecards SET precio_ref=:p, somos_mejores=:m, ellos_mejores=:e, objecion=:o, actualizado_por=:u, "
                           "actualizado_en=:f WHERE id=:i", i=ya[0][0], **{k: x for k, x in datos.items() if k != "c"})
                else:
                    _c48_x("INSERT INTO intel_battlecards (competidor, precio_ref, somos_mejores, ellos_mejores, objecion, actualizado_por, actualizado_en) "
                           "VALUES (:c,:p,:m,:e,:o,:u,:f)", **datos)
                _notif_crear(["Comercial"], "Guía comparativa actualizada: %s" % comp, "Revise en qué somos mejores frente a este competidor.", "/inteligencia/battlecards")
                registrar_auditoria("Battlecard guardada", "%s · %s" % (usuario, comp))
                _ic_flash("ok", "Guía comparativa guardada. Ventas ya puede consultarla.")
        except Exception as e:
            print("ic battlecards:", e)
            _ic_flash("err", "No se pudo guardar la guía.")
        return redirect("/inteligencia/battlecards")

    filas = _c48_q("SELECT competidor, precio_ref, somos_mejores, ellos_mejores, objecion, actualizado_por, actualizado_en FROM intel_battlecards ORDER BY competidor LIMIT 100")
    tarjetas = "".join(
        "<div class='ti-card'><h2>⚔️ Procsis vs %s</h2><p><b>Precio de referencia:</b> %s</p>"
        "<p><b>✅ En qué somos mejores:</b><br>%s</p><p><b>⚠️ Dónde ellos nos superan:</b><br>%s</p>"
        "<p><b>💬 Cómo responder la objeción:</b><br>%s</p><small>Actualizado por %s · %s</small></div>" % (
            _esc(r[0]), _esc(r[1] or "—"), _esc(r[2] or "—").replace("\n", "<br>"), _esc(r[3] or "—").replace("\n", "<br>"),
            _esc(r[4] or "—").replace("\n", "<br>"), _esc(r[5]), _esc(r[6])) for r in filas)
    form = ""
    if _ic_puede("battlecards"):
        form = ("<div class='ti-card'><h2>Crear o actualizar una guía comparativa</h2><form method='POST'><input type='hidden' name='csrf' value='%s'>"
                "<div class='ti-2'><div><label>Competidor (si ya existe, se actualiza)</label><input name='competidor' required maxlength='120'></div>"
                "<div><label>Precio de referencia</label><input name='precio_ref' maxlength='120'></div></div>"
                "<label>En qué somos mejores</label><textarea name='somos_mejores' rows='3'></textarea>"
                "<label>Dónde ellos nos superan</label><textarea name='ellos_mejores' rows='2'></textarea>"
                "<label>Cómo responder la objeción más común</label><textarea name='objecion' rows='2'></textarea>"
                "<p><button class='ti-btn'>Guardar guía</button></p></form></div>") % _acc_csrf()
    cuerpo = ("<h1>⚔️ Guías comparativas (Battlecards)</h1><p class='sub'>Consulta rápida para el equipo comercial: en qué es mejor Procsis frente a cada competidor evaluado.</p>%s%s") % (
        form, tarjetas or "<div class='ti-card'>Aún no hay guías comparativas.</div>")
    return _ic_pagina("Battlecards", cuerpo)


# ───────────── Calidad: auditoría de atención + plan de mejora automático ─────────────
@app.route("/inteligencia/auditoria", methods=["GET", "POST"])
def ic_auditoria():
    g = _ic_acceso("auditoria")
    if g is not None:
        return g
    _ic_init()
    usuario = session.get("usuario") or ""
    if request.method == "POST":
        v = _ic_post_ok("auditoria")
        if v is not None:
            return v
        accion = request.form.get("accion") or ""
        try:
            if accion == "guardar":
                tipo = "Competencia" if request.form.get("tipo") == "Competencia" else "Interna"
                sujeto = _ic_t("sujeto", 120)
                area = request.form.get("area") or ""
                a, c, r_ = _ic_n("amabilidad"), _ic_n("claridad"), _ic_n("rapidez")
                if not sujeto:
                    _ic_flash("err", "Escriba la empresa evaluada o la persona auditada.")
                elif tipo == "Interna" and area not in _IC_AREAS_PLAN:
                    _ic_flash("err", "Elija el área auditada.")
                elif not (a and c and r_):
                    _ic_flash("err", "Califique los tres criterios de 1 a 5.")
                else:
                    prom = round((a + c + r_) / 3.0, 1)
                    coment = _ic_tl("comentarios")
                    _c48_x("INSERT INTO intel_auditoria (tipo, sujeto, area, canal, amabilidad, claridad, rapidez, promedio, comentarios, creado_por, creado_en) "
                           "VALUES (:t,:s,:a,:c,:am,:cl,:ra,:p,:co,:u,:f)", t=tipo, s=sujeto, a=area if tipo == "Interna" else "", c=_ic_t("canal", 60),
                           am=a, cl=c, ra=r_, p=str(prom), co=coment, u=usuario, f=_ic_ts())
                    aid = (_c48_q("SELECT MAX(id) FROM intel_auditoria WHERE creado_por=:u", u=usuario) or [(0,)])[0][0] or 0
                    msg = "Auditoría guardada (promedio %.1f/5)." % prom
                    bajos = [n for n, x in (("amabilidad", a), ("claridad", c), ("rapidez", r_)) if x <= _IC_UMBRAL_BAJO]
                    if tipo == "Interna" and bajos:
                        _c48_x("INSERT INTO intel_planes (auditoria_id, area_responsable, titulo, detalle, estado, creado_por, creado_en, actualizado_en) "
                               "VALUES (:a,:r,:t,:d,'Abierto',:u,:f,:f)", a=aid, r=area, t=("Plan de mejora: atención de %s" % sujeto)[:200],
                               d=("Criterios bajos: %s. Promedio %.1f/5.\n%s" % (", ".join(bajos), prom, coment))[:1500], u=usuario, f=_ic_ts())
                        _notif_crear([area, "Calidad"], "Nuevo plan de mejora para %s" % _IC_AREAS_PLAN[area],
                                     "Auditoría de %s con calificación baja en: %s." % (sujeto, ", ".join(bajos)), "/inteligencia/planes", nivel="alerta")
                        msg += " Se abrió un Plan de Mejora para %s." % _IC_AREAS_PLAN[area]
                    elif tipo == "Competencia" and prom >= 4:
                        _ic_propuesta("Calidad", "Brecha frente a %s" % sujeto,
                                      "La competencia obtuvo %.1f/5 en atención (amabilidad %d, claridad %d, rapidez %d).\n%s" % (prom, a, c, r_, coment),
                                      "Otro / general", usuario)
                        _notif_crear(["Soluciones", "Comercial"], "Brecha detectada frente a %s" % sujeto,
                                     "Calidad midió una atención destacada de la competencia.", "/inteligencia/propuestas")
                        msg += " Se envió la brecha a Soluciones y Ventas."
                    registrar_auditoria("Auditoría de atención", "%s · %s · %s" % (usuario, tipo, sujeto))
                    _ic_flash("ok", msg)
            elif accion == "enviar_soluciones":
                f = _c48_q("SELECT id, tipo, sujeto, promedio, comentarios FROM intel_auditoria WHERE id=:i", i=_ic_int("id"))
                if f:
                    _ic_propuesta("Calidad", "Hallazgo de auditoría: %s" % f[0][2],
                                  "Tipo: %s · Promedio %s/5\n%s" % (f[0][1], f[0][3], f[0][4] or ""), "Otro / general", usuario)
                    _notif_crear(["Soluciones"], "Hallazgo de Calidad: %s" % f[0][2], "Revise la bandeja de propuestas.", "/inteligencia/propuestas")
                    _ic_flash("ok", "Hallazgo enviado a Soluciones.")
        except Exception as e:
            print("ic auditoria:", e)
            _ic_flash("err", "No se pudo completar la acción. Intente de nuevo.")
        return redirect("/inteligencia/auditoria")

    puede = _ic_puede("auditoria")
    csrf = _acc_csrf()
    filas = _c48_q("SELECT id, tipo, sujeto, area, canal, amabilidad, claridad, rapidez, promedio, comentarios, creado_por, creado_en "
                   "FROM intel_auditoria ORDER BY id DESC LIMIT 80")
    trs = ""
    for r in filas:
        envio = ""
        if puede:
            envio = ("<form method='POST' style='margin:0'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='enviar_soluciones'>"
                     "<input type='hidden' name='id' value='%d'><button class='ic-btn2'>Enviar a Soluciones</button></form>") % (csrf, r[0])
        trs += "<tr><td>%s<br><small>%s</small></td><td><b>%s</b><br><small>%s</small></td><td>%s · %s · %s</td><td>%s</td><td><small>%s</small></td><td>%s</td></tr>" % (
            _esc(r[1]), _esc(r[11][:10]), _esc(r[2]), _esc(_IC_AREAS_PLAN.get(r[3], "") or r[4] or ""), _ic_score(r[5]), _ic_score(r[6]), _ic_score(r[7]),
            _ic_score(r[8]), _esc((r[9] or "")[:140]), envio)
    form = ""
    if puede:
        form = ("<div class='ti-card'><h2>Nueva auditoría de atención</h2><form method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='guardar'>"
                "<div class='ti-2'><div><label>Qué se audita</label><select name='tipo'><option value='Interna'>Interna (nuestros asesores)</option><option value='Competencia'>Competencia (cliente incógnito)</option></select></div>"
                "<div><label>Empresa evaluada o asesor auditado</label><input name='sujeto' required maxlength='120'></div></div>"
                "<div class='ti-2'><div><label>Área auditada (solo interna)</label><select name='area'><option value=''>—</option>%s</select></div>"
                "<div><label>Canal</label><select name='canal'>%s</select></div></div>"
                "<div class='ti-2'><div><label>Amabilidad (1-5)</label><input type='number' name='amabilidad' min='1' max='5' required></div>"
                "<div><label>Claridad de la información (1-5)</label><input type='number' name='claridad' min='1' max='5' required></div></div>"
                "<div class='ti-2'><div><label>Rapidez (1-5)</label><input type='number' name='rapidez' min='1' max='5' required></div><div></div></div>"
                "<label>Comentarios</label><textarea name='comentarios' rows='3'></textarea>"
                "<p><small>Si un criterio de una auditoría interna queda en %d o menos, se abre un Plan de Mejora para el área.</small></p>"
                "<p><button class='ti-btn'>Guardar auditoría</button></p></form></div>") % (
            csrf, "".join("<option value='%s'>%s</option>" % (k, _esc(n)) for k, n in _IC_AREAS_PLAN.items()), _ic_opts(_IC_CANALES), _IC_UMBRAL_BAJO)
    cuerpo = ("<h1>✅ Auditoría de atención</h1><p class='sub'>Matriz de evaluación (1 a 5) para pruebas a la competencia y auditorías internas.</p>%s"
              "<div class='ti-card'><h2>Auditorías</h2><div style='overflow-x:auto'><table><tr><th>Tipo</th><th>Sujeto</th><th>Amab. · Clar. · Rapidez</th><th>Promedio</th>"
              "<th>Comentarios</th><th></th></tr>%s</table></div></div>") % (form, trs or "<tr><td colspan='6'>Aún no hay auditorías.</td></tr>")
    return _ic_pagina("Auditoría de atención", cuerpo)


# ───────────── Planes de mejora (cada área ve y atiende los suyos) ─────────────
@app.route("/inteligencia/planes", methods=["GET", "POST"])
def ic_planes():
    g = _ic_acceso("planes")
    if g is not None:
        return g
    _ic_init()
    rol = rol_actual()
    usuario = session.get("usuario") or ""
    ve_todo = rol in _IC_GER or rol == "Calidad"
    if request.method == "POST":
        v = _ic_post_ok("planes")
        if v is not None:
            return v
        try:
            estado = request.form.get("estado")
            f = _c48_q("SELECT id, area_responsable, titulo FROM intel_planes WHERE id=:i", i=_ic_int("id"))
            if not f or estado not in _IC_ESTADOS_PLAN:
                _ic_flash("err", "Plan o estado no válido.")
            elif not (ve_todo or rol == "Superadmin" or f[0][1] == rol):
                return acceso_denegado("Este plan de mejora le corresponde a otra área.")
            else:
                _c48_x("UPDATE intel_planes SET estado=:e, nota=:n, actualizado_en=:f, cerrado_en=:c WHERE id=:i",
                       e=estado, n=_ic_tl("nota", 800), f=_ic_ts(), c=_ic_ts() if estado == "Cerrado" else "", i=f[0][0])
                if estado == "Cerrado" and rol != "Calidad":
                    _notif_crear(["Calidad"], "Plan de mejora cerrado", "%s (%s)" % (f[0][2], _IC_AREAS_PLAN.get(f[0][1], f[0][1])), "/inteligencia/planes")
                registrar_auditoria("Plan de mejora actualizado", "%s · #%d · %s" % (usuario, f[0][0], estado))
                _ic_flash("ok", "Plan de mejora actualizado.")
        except Exception as e:
            print("ic planes:", e)
            _ic_flash("err", "No se pudo actualizar el plan.")
        return redirect("/inteligencia/planes")

    if ve_todo:
        filas = _c48_q("SELECT id, area_responsable, titulo, detalle, estado, nota, creado_en FROM intel_planes ORDER BY id DESC LIMIT 100")
    else:
        filas = _c48_q("SELECT id, area_responsable, titulo, detalle, estado, nota, creado_en FROM intel_planes WHERE area_responsable=:r ORDER BY id DESC LIMIT 100", r=rol)
    csrf = _acc_csrf()
    trs = ""
    for r in filas:
        puede = _ic_puede("planes") and (ve_todo or r[1] == rol or rol == "Superadmin")
        acc = ""
        if puede:
            acc = ("<form method='POST' style='margin:0'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='id' value='%d'>"
                   "<select name='estado'>%s</select><input name='nota' placeholder='Nota' maxlength='300' value='%s'>"
                   "<button class='ic-btn2'>Guardar</button></form>") % (csrf, r[0], _ic_opts(_IC_ESTADOS_PLAN, r[4]), _esc(r[5] or ""))
        trs += "<tr><td>#%d<br><small>%s</small></td><td><b>%s</b><br><small>%s</small></td><td>%s</td><td><span class='ti-b %s'>%s</span></td><td>%s</td></tr>" % (
            r[0], _esc(r[6][:10]), _esc(r[2]), _esc((r[3] or "")[:200]).replace("\n", "<br>"), _esc(_IC_AREAS_PLAN.get(r[1], r[1])),
            "a" if r[4] == "Cerrado" else "p", _esc(r[4]), acc or _esc(r[5] or "—"))
    cuerpo = ("<h1>🛠️ Planes de mejora</h1><p class='sub'>Se abren solos cuando una auditoría interna tiene un puntaje bajo. Cada área ve y atiende los suyos.</p>"
              "<div class='ti-card'><div style='overflow-x:auto'><table><tr><th>Plan</th><th>Detalle</th><th>Área</th><th>Estado</th><th>Seguimiento</th></tr>%s</table></div></div>") % (
        trs or "<tr><td colspan='5'>No hay planes de mejora.</td></tr>")
    return _ic_pagina("Planes de mejora", cuerpo)


# ───────────── Soluciones: bandeja de propuestas de mejora de producto ─────────────
@app.route("/inteligencia/propuestas", methods=["GET", "POST"])
def ic_propuestas():
    g = _ic_acceso("propuestas")
    if g is not None:
        return g
    _ic_init()
    usuario = session.get("usuario") or ""
    if request.method == "POST":
        v = _ic_post_ok("propuestas")
        if v is not None:
            return v
        try:
            estado = request.form.get("estado")
            f = _c48_q("SELECT id, titulo, detalle, ticket_ti FROM intel_propuestas WHERE id=:i", i=_ic_int("id"))
            producto = request.form.get("producto") if request.form.get("producto") in _IC_PRODUCTOS else "Otro / general"
            if not f or estado not in _IC_ESTADOS_PROP:
                _ic_flash("err", "Propuesta o estado no válido.")
            else:
                ticket = int(f[0][3] or 0)
                if estado == _IC_APROBADA and not ticket:
                    _ti_init()
                    _c48_x("INSERT INTO ti_tickets (categoria, prioridad, asunto, detalle, solicitante, estado, creado_por, creado_en, actualizado_en) "
                           "VALUES (:a,'Media',:b,:c,:u,'Abierto',:u,:f,:f)", a="Mejora de producto (inteligencia competitiva)",
                           b=("Mejora para %s: %s" % (producto, f[0][1]))[:200], c=(f[0][2] or "")[:1500], u=usuario, f=_ic_ts())
                    ticket = 1
                    _notif_crear(["Tecnologia", "Desarrollador"], "Nuevo requerimiento de producto", f[0][1], "/tecnologia/mesa-ayuda")
                if estado == "Implementada":
                    _notif_crear(["Marketing", "Comercial"], "Mejora lista: %s" % f[0][1], "Ya puede comunicarse y usarse como argumento de venta.", "/inteligencia/propuestas")
                _c48_x("UPDATE intel_propuestas SET estado=:e, decision_nota=:n, producto=:p, ticket_ti=:t, actualizado_en=:f WHERE id=:i",
                       e=estado, n=_ic_tl("nota", 800), p=producto, t=ticket, f=_ic_ts(), i=f[0][0])
                registrar_auditoria("Propuesta de producto", "%s · #%d · %s" % (usuario, f[0][0], estado))
                _ic_flash("ok", "Propuesta actualizada." + (" Se creó el requerimiento para TI/Desarrollo." if estado == _IC_APROBADA and ticket else ""))
        except Exception as e:
            print("ic propuestas:", e)
            _ic_flash("err", "No se pudo actualizar la propuesta.")
        return redirect("/inteligencia/propuestas")

    puede = _ic_puede("propuestas")
    csrf = _acc_csrf()
    filas = _c48_q("SELECT id, origen, titulo, detalle, producto, estado, decision_nota, ticket_ti, creado_en FROM intel_propuestas ORDER BY id DESC LIMIT 100")
    trs = ""
    for r in filas:
        acc = _esc(r[6] or "—")
        if puede:
            acc = ("<form method='POST' style='margin:0'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='id' value='%d'>"
                   "<select name='producto'>%s</select><select name='estado'>%s</select><input name='nota' placeholder='Decisión / nota' maxlength='300' value='%s'>"
                   "<button class='ic-btn2'>Guardar</button></form>") % (csrf, r[0], _ic_opts(_IC_PRODUCTOS, r[4]), _ic_opts(_IC_ESTADOS_PROP, r[5]), _esc(r[6] or ""))
        trs += "<tr><td>#%d<br><small>%s · %s</small></td><td><b>%s</b><br><small>%s</small></td><td><span class='ti-b %s'>%s</span>%s</td><td>%s</td></tr>" % (
            r[0], _esc(r[1]), _esc(r[8][:10]), _esc(r[2]), _esc((r[3] or "")[:220]).replace("\n", "<br>"),
            "a" if r[5] in ("Implementada", "Descartada") else "p", _esc(r[5]), "<br><small>Ticket TI creado</small>" if r[7] else "", acc)
    cuerpo = ("<h1>💡 Propuestas de mejora de producto</h1><p class='sub'>Llegan desde Marketing y Calidad. Soluciones decide qué agregar a EduTrack o TiendaTrack; al aprobar, se crea el requerimiento para TI y Desarrollo.</p>"
              "<div class='ti-card'><div style='overflow-x:auto'><table><tr><th>Origen</th><th>Propuesta</th><th>Estado</th><th>Decisión</th></tr>%s</table></div></div>") % (
        trs or "<tr><td colspan='4'>La bandeja está vacía.</td></tr>")
    return _ic_pagina("Propuestas de producto", cuerpo)


# ───────────── Gerencia: tablero de resumen (solo lectura) ─────────────
@app.route("/inteligencia/resumen")
def ic_resumen():
    if not requiere_login():
        return redirect("/login")
    if rol_actual() not in _IC_GER:
        return acceso_denegado("El resumen es exclusivo de Gerencia.")
    _ic_init()
    meses = _c48_q("SELECT substr(creado_en,1,7), COUNT(*) FROM intel_competencia GROUP BY substr(creado_en,1,7) ORDER BY 1 DESC LIMIT 6")
    aud_meses = _c48_q("SELECT substr(creado_en,1,7), COUNT(*) FROM intel_auditoria GROUP BY substr(creado_en,1,7) ORDER BY 1 DESC LIMIT 6")
    empresas = _c48_q("SELECT MIN(empresa), COUNT(*) FROM intel_competencia GROUP BY lower(empresa) ORDER BY 2 DESC LIMIT 15")
    planes = dict((r[0], r[1]) for r in _c48_q("SELECT estado, COUNT(*) FROM intel_planes GROUP BY estado"))
    props = dict((r[0], r[1]) for r in _c48_q("SELECT estado, COUNT(*) FROM intel_propuestas GROUP BY estado"))
    prom = {}
    for t, p in _c48_q("SELECT tipo, promedio FROM intel_auditoria"):
        try:
            prom.setdefault(t, []).append(float(p))
        except Exception:
            pass
    def media(t):
        x = prom.get(t) or []
        return "%.1f" % (sum(x) / len(x)) if x else "—"
    total_comp = sum(r[1] for r in meses) if meses else 0
    mayor = max([r[1] for r in meses] + [1])
    barras = "".join("<tr><td>%s</td><td style='width:60%%'><div class='ic-bar' style='width:%d%%'></div></td><td><b>%d</b></td></tr>" % (
        _esc(r[0]), max(4, int(100 * r[1] / mayor)), r[1]) for r in meses)
    aud_txt = "".join("<tr><td>%s</td><td><b>%d</b></td></tr>" % (_esc(r[0]), r[1]) for r in aud_meses)
    emp_txt = "".join("<tr><td>%s</td><td><b>%d</b></td></tr>" % (_esc(r[0]), r[1]) for r in empresas)
    cuerpo = ("<h1>📊 Resumen de inteligencia competitiva</h1><p class='sub'>Vista de Gerencia (solo lectura).</p>"
              "<div class='ic-k'><div><b>%d</b>Evaluaciones de competencia (últimos 6 meses)</div><div><b>%d</b>Empresas analizadas</div>"
              "<div><b>%s</b>Promedio auditorías internas</div><div><b>%s</b>Promedio competencia</div>"
              "<div><b>%d</b>Planes de mejora abiertos</div><div><b>%d</b>Mejoras implementadas</div></div>"
              "<div class='ti-2'><div class='ti-card'><h2>Evaluaciones por mes</h2><table>%s</table></div>"
              "<div class='ti-card'><h2>Auditorías por mes</h2><table>%s</table></div></div>"
              "<div class='ti-2'><div class='ti-card'><h2>Empresas analizadas</h2><table><tr><th>Empresa</th><th>Evaluaciones</th></tr>%s</table></div>"
              "<div class='ti-card'><h2>Propuestas y planes</h2><table><tr><th>Propuestas</th><th></th></tr>%s<tr><th>Planes</th><th></th></tr>%s</table></div></div>") % (
        total_comp, len(empresas), media("Interna"), media("Competencia"),
        int(planes.get("Abierto", 0)) + int(planes.get("En proceso", 0)), int(props.get("Implementada", 0)),
        barras or "<tr><td>Sin datos todavía.</td></tr>", aud_txt or "<tr><td>Sin datos todavía.</td></tr>",
        emp_txt or "<tr><td colspan='2'>Sin datos todavía.</td></tr>",
        "".join("<tr><td>%s</td><td><b>%d</b></td></tr>" % (_esc(k), v) for k, v in props.items()) or "<tr><td colspan='2'>—</td></tr>",
        "".join("<tr><td>%s</td><td><b>%d</b></td></tr>" % (_esc(k), v) for k, v in planes.items()) or "<tr><td colspan='2'>—</td></tr>")
    return _ic_pagina("Resumen de inteligencia competitiva", cuerpo)


def _ic_registrar_menus():
    """Menús y permisos: reparte el módulo entre las áreas existentes (no crea área nueva)."""
    _MODULOS_POR_AREA["/inteligencia"] = (_IC_OPERAN, ())
    mods = {
        "competencia": ("/inteligencia/competencia", "🕵️ Estudio de competencia / Cliente incógnito", "Cotizaciones anónimas, precios y tiempos de la competencia."),
        "battlecards": ("/inteligencia/battlecards", "⚔️ Guías comparativas (Battlecards)", "En qué somos mejores frente a cada competidor."),
        "auditoria": ("/inteligencia/auditoria", "✅ Auditoría de atención", "Matriz 1-5 para evaluar atención propia y de la competencia."),
        "planes": ("/inteligencia/planes", "🛠️ Planes de mejora", "Planes que se abren solos cuando una auditoría sale baja."),
        "propuestas": ("/inteligencia/propuestas", "💡 Propuestas de mejora de producto", "Hallazgos de Marketing y Calidad para EduTrack y TiendaTrack."),
    }

    def _area(rol, claves):
        lista = _AREA_MODULOS_EXTRA.setdefault(rol, [])
        ya = {x[0] for x in lista}
        for k in claves:
            if mods[k][0] not in ya:
                lista.append(mods[k])

    def _menu(rol, claves):
        lista = _REPARTO_MENU_EXTRA.setdefault(rol, [])
        ya = {x[0] for x in lista}
        for k in claves:
            if mods[k][0] not in ya:
                lista.append((mods[k][0], mods[k][1]))

    _area("Marketing", ("competencia", "battlecards", "propuestas", "planes"))
    _area("Calidad", ("auditoria", "planes", "competencia", "propuestas"))
    _area("Tecnologia", ("propuestas", "planes"))
    _menu("Soluciones", ("propuestas", "competencia", "battlecards", "planes"))
    _menu("Comercial", ("battlecards", "planes"))
    _menu("Soporte", ("planes",))
    _menu("Cobranza", ("planes",))
    _menu("Desarrollador", ("propuestas",))
    grupo = ("Inteligencia competitiva", [(mods[k][0], mods[k][1]) for k in ("competencia", "battlecards", "auditoria", "planes", "propuestas")]
             + [("/inteligencia/resumen", "📊 Resumen de inteligencia competitiva")])
    try:
        if not any(g[0] == grupo[0] for g in _MODULOS_GERENCIA):
            _MODULOS_GERENCIA.append(grupo)
            _GRUPOS_MODULOS_GERENCIA.append((grupo[0], tuple(h for h, _l in grupo[1])))
        for h, l in grupo[1]:
            _REPARTO_ETIQUETAS.setdefault(h, l)
    except Exception as e:
        print("ic menus gerencia:", e)


_ic_registrar_menus()


'''


def main():
    if not os.path.isfile(RUTA):
        print("No encuentro %s. Ejecute el script en la carpeta de app.py o pase la ruta." % RUTA)
        sys.exit(1)
    with open(RUTA, "rb") as f:
        crudo = f.read().decode("utf-8")
    crlf = "\r\n" in crudo
    txt = crudo.replace("\r\n", "\n")

    if MARCA in txt:
        print("Este cambio ya está aplicado (%s). No se hizo nada." % MARCA)
        return
    if txt.count(ANCLA) != 1:
        print("ABORTADO: no encuentro un único `%s` (aparece %d veces). No se modificó nada." % (ANCLA, txt.count(ANCLA)))
        sys.exit(1)
    for nombre in ("_notif_crear", "_c48_q", "_c48_x", "_ti_csrf_ok", "_ti_init", "_acc_csrf", "_TI_CSS",
                   "_AREA_MODULOS_EXTRA", "_REPARTO_MENU_EXTRA", "_MODULOS_GERENCIA", "_GRUPOS_MODULOS_GERENCIA",
                   "_REPARTO_ETIQUETAS", "_MODULOS_POR_AREA"):
        if nombre not in txt:
            print("ABORTADO: no existe `%s` en app.py; esta versión no es compatible. No se modificó nada." % nombre)
            sys.exit(1)

    nuevo = txt.replace(ANCLA, BLOQUE + ANCLA, 1)
    try:
        ast.parse(nuevo)
    except SyntaxError as e:
        print("ABORTADO: el resultado tendría un error de sintaxis (línea %s). No se modificó nada." % e.lineno)
        sys.exit(1)

    salida = nuevo.replace("\n", "\r\n") if crlf else nuevo
    respaldo = RUTA + ".bak_inteligencia_competitiva"
    shutil.copy2(RUTA, respaldo)
    tmp = RUTA + ".tmp_ic"
    with open(tmp, "wb") as f:
        f.write(salida.encode("utf-8"))
    os.replace(tmp, RUTA)

    print("Listo.")
    print(" - Respaldo: %s" % respaldo)
    print(" - Rutas nuevas: /inteligencia/competencia, /battlecards, /auditoria, /planes, /propuestas, /resumen")
    print(" - Tablas nuevas (se crean solas al primer uso): intel_competencia, intel_battlecards, intel_auditoria, intel_planes, intel_propuestas")
    print("Reinicie la app / haga deploy para que tome efecto.")


if __name__ == "__main__":
    main()

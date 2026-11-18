# -*- coding: utf-8 -*-
"""
aplicar_plantillas.py

Biblioteca de PLANTILLAS de texto (sin IA, sin costo) para:
  1) Responder PQR   -> selector "✨ Plantilla de respuesta" en el detalle de la PQR
                        (se llena con nombre, radicado, tipo, motivo, etc. del caso).
  2) Crear documentos -> selector "✨ Plantilla" en Documentos > Nuevo, filtrado por el tipo
                        de documento elegido; rellena asunto y cuerpo.

Administración (Gerencia y Calidad): pantalla /plantillas para crear, editar, activar y
desactivar plantillas, cada una con su tipo, categoría y área (General o de un área).
Incluye 13 plantillas de arranque. Los textos con [COMPLETAR: ...] obligan a revisar:
si quedan sin llenar, el formulario pregunta antes de guardar.

Campos automáticos:  {nombre} {radicado} {ticket} {tipo} {subtipo} {motivo} {estado}
{fecha_radicacion} (PQR)  ·  {fecha} {empresa} {producto} {area} {elaboro} (todas)
{destinatario} {asunto} {firmante} (documentos, se toman del formulario).

Seguridad del cambio:
  * Agrega código nuevo y UNA tabla nueva (plantillas_texto, CREATE TABLE IF NOT EXISTS).
    No toca ni borra datos existentes.
  * Solo modifica una línea del detalle de PQR (inserta el selector) y envuelve el formulario
    de documentos para añadirle el selector; si algo falla, el formulario original sigue igual.
  * Respaldo app.py.bak_plantillas, valida sintaxis antes de escribir, conserva saltos de línea
    y no se aplica dos veces.

Uso:   python aplicar_plantillas.py        (usa ./app.py)
"""
import ast
import os
import shutil
import sys

MARCA = "[PROCSIS-PLANTILLAS]"
RUTA = sys.argv[1] if len(sys.argv) > 1 else "app.py"
ANCLA = 'if __name__ == "__main__":'

# Línea del detalle de PQR donde se inserta el selector (dentro de un f-string existente)
PQR_VIEJO = '    <label>Respuesta al usuario (se guarda y no se borra)</label>'
PQR_NUEVO = '    {_plt_selector_pqr(t)}\n    <label>Respuesta al usuario (se guarda y no se borra)</label>'

BLOQUE = r'''# ══════════════════════════════════════════════════════════════════
# [PROCSIS-PLANTILLAS]
# Biblioteca de plantillas para respuestas PQR y documentos por área (sin IA, sin costo).
# ══════════════════════════════════════════════════════════════════
_PLT_INIT = {"ok": False}
_PLT_ADMIN = ("Gerente", "Administrador", "Superadmin", "Calidad")
_PLT_CATS_PQR = (("queja", "Queja"), ("peticion", "Petición"), ("reclamo", "Reclamo"), ("solicitud", "Solicitud"), ("general", "General"))
_PLT_AREAS = ("General", "Gerencia", "Soporte", "Comercial", "Cobranza", "Marketing", "Calidad", "Soluciones",
              "Tecnologia", "Contabilidad", "Finanzas", "Talento")

_PLT_SEMILLAS = (
    ("pqr", "queja", "Queja · Fallas del servicio", "General", "",
     "En atención a su queja, revisamos el caso y verificamos la información asociada a su cuenta en {producto}.\n\n"
     "Hallazgo: [COMPLETAR: qué se encontró al revisar el caso].\n"
     "Acción realizada: [COMPLETAR: corrección o ajuste aplicado y la fecha en que se hizo].\n\n"
     "Si luego de esta gestión la falla continúa, le pedimos responder a este mismo radicado ({radicado}) indicando en qué momento ocurre y, "
     "de ser posible, adjuntando una captura de pantalla, para que nuestro equipo técnico continúe el análisis.\n\n"
     "Lamentamos los inconvenientes ocasionados y agradecemos su paciencia."),
    ("pqr", "queja", "Queja · Atención del asesor", "General", "",
     "Lamentamos que la atención recibida no cumpliera con sus expectativas. Su queja fue trasladada al área de Calidad, que revisará la "
     "interacción señalada y aplicará las acciones correctivas que correspondan con el equipo involucrado.\n\n"
     "Resultado de la revisión: [COMPLETAR: conclusión de la revisión].\n"
     "Medida adoptada: [COMPLETAR: acción con el asesor o el proceso].\n\n"
     "Agradecemos que nos haya hecho conocer esta situación; nos permite mejorar el servicio."),
    ("pqr", "peticion", "Petición · Solicitud de información", "General", "",
     "En atención a su petición, le informamos lo siguiente respecto a {producto}:\n\n"
     "[COMPLETAR: respuesta clara y completa a cada punto solicitado].\n\n"
     "Si requiere ampliar o aclarar algún aspecto de esta información, puede responder a este mismo radicado ({radicado})."),
    ("pqr", "reclamo", "Reclamo · Facturación o cobro", "General", "",
     "En atención a su reclamo, revisamos el historial de facturación y pagos asociado a su cuenta.\n\n"
     "Resultado de la revisión: [COMPLETAR: qué se verificó: valores, fechas, pagos aplicados].\n"
     "Decisión: [COMPLETAR: se acepta / se rechaza el reclamo y el motivo].\n"
     "Ajuste aplicado (si corresponde): [COMPLETAR: valor, concepto y fecha de aplicación].\n\n"
     "Si desea que revisemos nuevamente el caso, puede responder a este mismo radicado ({radicado}) adjuntando los soportes de pago."),
    ("pqr", "solicitud", "Solicitud de cancelación del servicio", "General", "",
     "Recibimos su solicitud de cancelación del servicio de {producto}.\n\n"
     "Fecha efectiva de la cancelación: [COMPLETAR: fecha].\n"
     "Saldos o valores pendientes: [COMPLETAR: detalle o «sin saldos pendientes»].\n"
     "Condiciones del cierre de la cuenta y de la información: [COMPLETAR: según contrato y política de datos].\n\n"
     "Lamentamos su decisión y quedamos atentos si desea retomar el servicio más adelante."),
    ("pqr", "general", "Solicitud de información adicional al cliente", "General", "",
     "Para poder avanzar con su caso necesitamos que nos confirme o complemente la siguiente información:\n\n"
     "[COMPLETAR: lista de datos o soportes que se necesitan].\n\n"
     "Puede responder a este mismo radicado ({radicado}). Una vez la recibamos, continuaremos con la gestión de su solicitud."),
    ("doc", "notificacion", "Circular informativa", "General", "Circular informativa: [COMPLETAR: tema]",
     "Cordial saludo,\n\n"
     "Por medio de la presente, {empresa} informa a {destinatario} lo siguiente:\n\n"
     "[COMPLETAR: contenido de la circular: qué cambia, desde cuándo y a quién aplica].\n\n"
     "Para dudas o aclaraciones, comunicarse con {area}.\n\n"
     "Fecha: {fecha}."),
    ("doc", "certificacion", "Certificación de servicio", "General", "Certificación de servicio",
     "{empresa} certifica que [COMPLETAR: nombre de la persona o entidad] identificado(a) con [COMPLETAR: documento] "
     "[COMPLETAR: hecho que se certifica: por ejemplo, cuenta con el servicio de {producto} activo desde (fecha)].\n\n"
     "La presente certificación se expide a solicitud del interesado el {fecha}, para los fines que estime convenientes."),
    ("doc", "acta", "Acta de reunión", "General", "Acta de reunión: [COMPLETAR: tema]",
     "Fecha: {fecha}\nLugar o medio: [COMPLETAR]\nAsistentes: [COMPLETAR: nombres y cargos]\n\n"
     "Orden del día:\n1. [COMPLETAR]\n2. [COMPLETAR]\n\n"
     "Desarrollo y decisiones:\n[COMPLETAR: resumen de lo tratado y decidido].\n\n"
     "Compromisos:\n- [COMPLETAR: tarea · responsable · fecha límite]\n\n"
     "Acta elaborada por {elaboro}."),
    ("doc", "comercial", "Propuesta o carta comercial", "Comercial", "Propuesta de servicios {producto}",
     "Estimado(a) {destinatario}:\n\n"
     "Agradecemos su interés en {producto}. De acuerdo con lo conversado, presentamos nuestra propuesta:\n\n"
     "Alcance: [COMPLETAR: qué incluye el servicio].\n"
     "Valor: [COMPLETAR: precio y forma de pago].\n"
     "Vigencia de la propuesta: [COMPLETAR].\n\n"
     "Quedamos atentos para resolver cualquier inquietud y acompañarlo en la implementación."),
    ("doc", "requerimiento", "Requerimiento interno entre áreas", "General", "Requerimiento: [COMPLETAR: asunto]",
     "De: {area}\nPara: {destinatario}\nFecha: {fecha}\n\n"
     "Solicitud: [COMPLETAR: qué se necesita].\n"
     "Motivo: [COMPLETAR: por qué se necesita].\n"
     "Fecha requerida: [COMPLETAR].\n"
     "Prioridad: [COMPLETAR: alta / media / baja].\n\n"
     "Agradecemos confirmar recibido y la fecha estimada de entrega."),
    ("doc", "informe", "Informe técnico (estructura)", "General", "Informe técnico: [COMPLETAR: tema]",
     "1. Objetivo\n[COMPLETAR]\n\n2. Alcance y período revisado\n[COMPLETAR]\n\n3. Hallazgos\n[COMPLETAR]\n\n"
     "4. Análisis\n[COMPLETAR]\n\n5. Conclusiones y recomendaciones\n[COMPLETAR]\n\n"
     "Elaboró: {elaboro} · {area} · {fecha}."),
    ("doc", "financiero", "Recordatorio de pago (cobranza)", "Cobranza", "Recordatorio de pago",
     "Estimado(a) {destinatario}:\n\n"
     "Le recordamos que a la fecha ({fecha}) registramos un saldo pendiente por el servicio de {producto}:\n\n"
     "Valor pendiente: [COMPLETAR: valor].\nFecha de vencimiento: [COMPLETAR].\n\n"
     "Si ya realizó el pago, por favor envíenos el soporte para aplicarlo. Si necesita acordar una fecha de pago, "
     "comuníquese con {area}."),
)


def _plt_init():
    if _PLT_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS plantillas_texto (id %s, tipo VARCHAR(10) DEFAULT 'doc', categoria VARCHAR(30) DEFAULT '', "
            "nombre TEXT DEFAULT '', area VARCHAR(30) DEFAULT 'General', asunto TEXT DEFAULT '', cuerpo TEXT DEFAULT '', "
            "activa INTEGER DEFAULT 1, creado_por VARCHAR(80) DEFAULT '', creado_en VARCHAR(20) DEFAULT '', "
            "actualizado_por VARCHAR(80) DEFAULT '', actualizado_en VARCHAR(20) DEFAULT '')" % pk))
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("plantillas init:", e)
        return
    try:
        n = _c48_q("SELECT COUNT(*) FROM plantillas_texto")
        if int((n[0][0] if n else 0) or 0) == 0:
            ts = ahora().strftime("%Y-%m-%d %H:%M:%S")
            for tipo, cat, nombre, area, asunto, cuerpo in _PLT_SEMILLAS:
                _c48_x("INSERT INTO plantillas_texto (tipo, categoria, nombre, area, asunto, cuerpo, activa, creado_por, creado_en, actualizado_por, actualizado_en) "
                       "VALUES (:t,:c,:n,:a,:s,:b,1,'sistema',:f,'sistema',:f)", t=tipo, c=cat, n=nombre, a=area, s=asunto, b=cuerpo, f=ts)
    except Exception as e:
        db.session.rollback()
        print("plantillas semillas:", e)
    _PLT_INIT["ok"] = True


def _plt_norm(s):
    import unicodedata
    return "".join(ch for ch in unicodedata.normalize("NFD", str(s or "").lower()) if unicodedata.category(ch) != "Mn").strip()


def _plt_vars(t=None):
    v = {}
    try:
        meses = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        y, m, d = fecha_hoy()[:10].split("-")
        v["fecha"] = "%02d de %s de %s" % (int(d), meses[int(m)], y)
    except Exception:
        v["fecha"] = fecha_hoy()
    try:
        v["empresa"] = nombre_empresa()
        v["producto"] = nombre_producto()
        v["area"] = _dcx_area_titulo(_dcx_area_de_rol(rol_actual()))
        v["elaboro"] = _elab_usuario(session.get("usuario"))
    except Exception:
        pass
    if t is not None:
        v["nombre"] = ("%s %s" % ((t.razon_social or "").strip(), (t.apellidos or "").strip())).strip() or "Usuario"
        v["radicado"] = formatear_radicado(t.radicado or t.nro_cun)
        v["ticket"] = t.ticket or ""
        v["tipo"] = t.tipo_pqr or ""
        v["subtipo"] = t.subtipo or ""
        v["motivo"] = (t.objeto or t.motivo or "").strip()[:300]
        v["estado"] = t.estado or ""
        v["fecha_radicacion"] = t.fecha or ""
    return v


def _plt_llenar(txt, v):
    return re.sub(r"\{(\w+)\}", lambda mo: str(v.get(mo.group(1), mo.group(0))), txt or "")


def _plt_visibles(tipo, t=None):
    """Plantillas activas del tipo, ya con los campos automáticos resueltos y filtradas por área."""
    _plt_init()
    rol = rol_actual()
    area_u = _dcx_area_de_rol(rol)
    v = _plt_vars(t)
    out = []
    for r in _c48_q("SELECT id, categoria, nombre, area, asunto, cuerpo FROM plantillas_texto WHERE tipo=:t AND activa=1 ORDER BY categoria, nombre", t=tipo):
        if r[3] not in ("General", "") and r[3] != area_u and area_u != "Gerencia" and rol != "Superadmin":
            continue
        out.append({"id": r[0], "categoria": r[1], "nombre": r[2], "asunto": _plt_llenar(r[4], v), "cuerpo": _plt_llenar(r[5], v)})
    return out


def _plt_json(datos):
    return json.dumps(datos, ensure_ascii=False).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def _plt_selector_pqr(t):
    """Selector que se inserta dentro del formulario de respuesta de la PQR. Nunca rompe la página."""
    try:
        datos = _plt_visibles("pqr", t)
        if not datos:
            return ""
        tipo_n = _plt_norm(t.tipo_pqr)
        datos.sort(key=lambda d: (0 if tipo_n and _plt_norm(d["categoria"]) and tipo_n.startswith(_plt_norm(d["categoria"])) else 1, d["nombre"]))
        return ("<div class='plt-box' style='background:#f0fdfa;border:1px solid #99f6e4;border-radius:12px;padding:10px 12px;margin:12px 0'>"
                "<label style='margin:0 0 6px;display:block'>✨ Plantilla de respuesta</label>"
                "<div style='display:flex;gap:8px;flex-wrap:wrap'><select class='plt-sel' style='flex:1;min-width:200px'></select>"
                "<button type='button' class='plt-ins' style='width:auto'>Insertar</button></div>"
                "<small style='color:#475569'>Se llena con los datos del caso. Revise y complete los [COMPLETAR…] antes de guardar.</small></div>"
                "<script>(function(){var D=" + _plt_json(datos) + ";var box=document.currentScript.previousElementSibling;"
                "var sel=box.querySelector('select'),btn=box.querySelector('button');"
                "var ta=(box.closest('form')||document).querySelector('textarea[name=\"respuesta\"]');"
                "var o=document.createElement('option');o.value='';o.textContent='— Elegir plantilla —';sel.appendChild(o);"
                "D.forEach(function(p){var x=document.createElement('option');x.value=p.id;x.textContent=p.nombre;sel.appendChild(x);});"
                "btn.onclick=function(){var p=D.filter(function(x){return String(x.id)===sel.value})[0];if(!p||!ta)return;"
                "if(ta.value.trim()&&!confirm('Ya hay texto escrito. ¿Reemplazarlo con la plantilla?'))return;ta.value=p.cuerpo;ta.focus();};"
                "var f=ta&&ta.form;if(f)f.addEventListener('submit',function(e){if(/\\[COMPLETAR/i.test(ta.value)&&"
                "!confirm('La respuesta aún tiene campos [COMPLETAR…] sin llenar. ¿Guardar y enviar así?'))e.preventDefault();});})();</script>")
    except Exception as e:
        print("plantillas selector pqr:", e)
        return ""


def _plt_script_docs():
    datos = _plt_visibles("doc")
    if not datos:
        return ""
    return ("<script>(function(){var D=" + _plt_json(datos) + ";var ta=document.querySelector('textarea[name=\"cuerpo\"]');if(!ta)return;"
            "var asu=document.querySelector('[name=\"asunto\"]'),cat=document.querySelector('[name=\"categoria\"]'),"
            "fir=document.querySelector('[name=\"firmante\"]'),des=document.querySelector('[name=\"dest_nombre\"]');"
            "var box=document.createElement('div');box.style.cssText='background:#f0fdfa;border:1px solid #99f6e4;border-radius:12px;padding:10px 12px;margin:10px 0';"
            "var lb=document.createElement('label');lb.style.cssText='margin:0 0 6px;display:block';lb.textContent='✨ Plantilla de documento';box.appendChild(lb);"
            "var row=document.createElement('div');row.style.cssText='display:flex;gap:8px;flex-wrap:wrap';"
            "var sel=document.createElement('select');sel.style.cssText='flex:1;min-width:200px';var btn=document.createElement('button');"
            "btn.type='button';btn.textContent='Insertar';btn.style.width='auto';row.appendChild(sel);row.appendChild(btn);box.appendChild(row);"
            "var sm=document.createElement('small');sm.style.color='#475569';sm.textContent='Rellena el asunto y el cuerpo. Complete los [COMPLETAR…] antes de emitir.';box.appendChild(sm);"
            "var prev=ta.previousElementSibling;var ref=(prev&&prev.tagName==='LABEL')?prev:ta;ref.parentNode.insertBefore(box,ref);"
            "function llenar(){var c=cat?cat.value:'';var L=D.filter(function(p){return !c||p.categoria===c;});if(!L.length)L=D;sel.innerHTML='';"
            "var o=document.createElement('option');o.value='';o.textContent='— Elegir plantilla —';sel.appendChild(o);"
            "L.forEach(function(p){var x=document.createElement('option');x.value=p.id;x.textContent=p.nombre;sel.appendChild(x);});}"
            "llenar();if(cat)cat.addEventListener('change',llenar);"
            "function v(el,ph){return (el&&el.value.trim())?el.value.trim():'[COMPLETAR: '+ph+']';}"
            "btn.onclick=function(){var p=D.filter(function(x){return String(x.id)===sel.value})[0];if(!p)return;"
            "if(ta.value.trim()&&!confirm('Ya hay texto escrito. ¿Reemplazarlo con la plantilla?'))return;"
            "var t=p.cuerpo.split('{destinatario}').join(v(des,'destinatario')).split('{firmante}').join(v(fir,'firmante'))"
            ".split('{asunto}').join(v(asu,'asunto'));ta.value=t;if(asu&&!asu.value.trim()&&p.asunto)asu.value=p.asunto;ta.focus();};"
            "var f=ta.form;if(f)f.addEventListener('submit',function(e){if(/\\[COMPLETAR/i.test(ta.value+(asu?asu.value:''))&&"
            "!confirm('El documento aún tiene campos [COMPLETAR…] sin llenar. ¿Emitirlo así?'))e.preventDefault();});})();</script>")


_dca_form_plt_orig = _dca_form


def _dca_form(csrf):
    h = _dca_form_plt_orig(csrf)
    try:
        h += _plt_script_docs()
    except Exception as e:
        print("plantillas docs:", e)
    return h


# ───────────── Administración de plantillas (Gerencia y Calidad) ─────────────
def _plt_guard():
    if not requiere_login():
        return redirect("/login")
    if rol_actual() not in _PLT_ADMIN:
        return acceso_denegado("Solo Gerencia y Calidad administran las plantillas.")
    _plt_init()
    return None


def _plt_pagina(titulo, cuerpo):
    msg = session.pop("ti_msg", None)
    aviso = ("<div class='%s'>%s</div>" % ("ti-ok" if msg[0] == "ok" else "ti-err", _esc(msg[1]))) if isinstance(msg, (list, tuple)) and len(msg) == 2 else ""
    return page(titulo, shell(_TI_CSS + "<div class='ti'>%s%s</div>" % (aviso, cuerpo)))


@app.route("/plantillas")
def plt_lista():
    g = _plt_guard()
    if g is not None:
        return g
    csrf = _acc_csrf()
    pid = request.args.get("id", type=int) or 0
    ed = None
    if pid:
        f = _c48_q("SELECT id, tipo, categoria, nombre, area, asunto, cuerpo, activa FROM plantillas_texto WHERE id=:i", i=pid)
        ed = f[0] if f else None
    sel_cat = "%s|%s" % (ed[1], ed[2]) if ed else "pqr|queja"
    opt_pqr = "".join("<option value='pqr|%s'%s>%s</option>" % (k, " selected" if sel_cat == "pqr|" + k else "", _esc(n)) for k, n in _PLT_CATS_PQR)
    opt_doc = "".join("<option value='doc|%s'%s>%s</option>" % (k, " selected" if sel_cat == "doc|" + k else "", _esc(n)) for k, n in _DCA_CATEGORIAS)
    opt_area = "".join("<option%s>%s</option>" % (" selected" if ed and ed[4] == a else "", a) for a in _PLT_AREAS)
    form = ("<div class='ti-card'><h2>%s</h2><form method='POST' action='/plantillas/guardar'><input type='hidden' name='csrf' value='%s'>"
            "<input type='hidden' name='id' value='%d'>"
            "<div class='ti-2'><div><label>Para qué sirve</label><select name='cat'><optgroup label='Respuesta a PQR'>%s</optgroup>"
            "<optgroup label='Documento'>%s</optgroup></select></div>"
            "<div><label>Quién la ve</label><select name='area'>%s</select></div></div>"
            "<label>Nombre de la plantilla</label><input name='nombre' required maxlength='120' value='%s'>"
            "<label>Asunto (solo documentos; opcional)</label><input name='asunto' maxlength='300' value='%s'>"
            "<label>Texto</label><textarea name='cuerpo' rows='12' required maxlength='6000'>%s</textarea>"
            "<p><small>Campos automáticos: <b>{nombre} {radicado} {ticket} {tipo} {subtipo} {motivo} {estado} {fecha_radicacion}</b> (PQR) · "
            "<b>{fecha} {empresa} {producto} {area} {elaboro}</b> (todas) · <b>{destinatario} {asunto} {firmante}</b> (documentos). "
            "Escriba <b>[COMPLETAR: qué falta]</b> donde el asesor deba llenar algo a mano.</small></p>"
            "<p><button class='ti-btn'>%s</button> %s</p></form></div>") % (
        "Editar plantilla" if ed else "Nueva plantilla", csrf, ed[0] if ed else 0, opt_pqr, opt_doc, opt_area,
        _esc(ed[3]) if ed else "", _esc(ed[5]) if ed else "", _esc(ed[6]) if ed else "",
        "Guardar cambios" if ed else "Crear plantilla", "<a href='/plantillas'>Cancelar</a>" if ed else "")
    trs = ""
    for r in _c48_q("SELECT id, tipo, categoria, nombre, area, cuerpo, activa, actualizado_por, actualizado_en FROM plantillas_texto ORDER BY tipo, categoria, nombre"):
        trs += ("<tr><td>%s</td><td>%s</td><td><b>%s</b><details><summary>Ver texto</summary><small>%s</small></details></td><td>%s</td>"
                "<td><span class='ti-b %s'>%s</span></td><td><small>%s<br>%s</small></td><td><a href='/plantillas?id=%d'>Editar</a>"
                "<form method='POST' action='/plantillas/%d/estado' style='display:inline'><input type='hidden' name='csrf' value='%s'>"
                "<button class='ti-btn' style='padding:3px 9px;margin-left:6px'>%s</button></form></td></tr>") % (
            "PQR" if r[1] == "pqr" else "Documento", _esc(r[2]), _esc(r[3]), _esc(r[5][:1200]).replace("\n", "<br>"), _esc(r[4]),
            "a" if r[6] else "p", "Activa" if r[6] else "Inactiva", _esc(r[7]), _esc((r[8] or "")[:10]), r[0], r[0], csrf,
            "Desactivar" if r[6] else "Activar")
    cuerpo = ("<h1>📝 Plantillas de textos</h1><p class='sub'>Textos base para responder PQR y crear documentos. Los asesores las ven como "
              "«Plantilla» al responder o al crear un documento. No se envían solas: siempre las revisa una persona.</p>%s"
              "<div class='ti-card'><h2>Biblioteca</h2><div style='overflow-x:auto'><table><tr><th>Uso</th><th>Categoría</th><th>Plantilla</th>"
              "<th>Área</th><th>Estado</th><th>Última edición</th><th></th></tr>%s</table></div></div>") % (form, trs or "<tr><td colspan='7'>Sin plantillas.</td></tr>")
    return _plt_pagina("Plantillas de textos", cuerpo)


@app.route("/plantillas/guardar", methods=["POST"])
def plt_guardar():
    g = _plt_guard()
    if g is not None:
        return g
    if not _ti_csrf_ok():
        session["ti_msg"] = ("err", "La sesión del formulario venció. Recargue e intente de nuevo.")
        return redirect("/plantillas")
    try:
        tipo, _, cat = (request.form.get("cat") or "").partition("|")
        validas = dict(_PLT_CATS_PQR) if tipo == "pqr" else (dict(_DCA_CATEGORIAS) if tipo == "doc" else {})
        area = request.form.get("area") or "General"
        nombre = " ".join((request.form.get("nombre") or "").split())[:120]
        asunto = " ".join((request.form.get("asunto") or "").split())[:300] if tipo == "doc" else ""
        cuerpo = (request.form.get("cuerpo") or "").replace("\r", "").strip()[:6000]
        if cat not in validas or area not in _PLT_AREAS or not nombre or not cuerpo:
            session["ti_msg"] = ("err", "Complete el uso, el área, el nombre y el texto de la plantilla.")
            return redirect("/plantillas")
        usuario, ts = session.get("usuario") or "", ahora().strftime("%Y-%m-%d %H:%M:%S")
        pid = int(request.form.get("id") or 0)
        if pid and _c48_q("SELECT id FROM plantillas_texto WHERE id=:i", i=pid):
            _c48_x("UPDATE plantillas_texto SET tipo=:t, categoria=:c, nombre=:n, area=:a, asunto=:s, cuerpo=:b, actualizado_por=:u, actualizado_en=:f WHERE id=:i",
                   t=tipo, c=cat, n=nombre, a=area, s=asunto, b=cuerpo, u=usuario, f=ts, i=pid)
            registrar_auditoria("Plantilla editada", "%s · #%d · %s" % (usuario, pid, nombre))
            session["ti_msg"] = ("ok", "Plantilla actualizada.")
        else:
            _c48_x("INSERT INTO plantillas_texto (tipo, categoria, nombre, area, asunto, cuerpo, activa, creado_por, creado_en, actualizado_por, actualizado_en) "
                   "VALUES (:t,:c,:n,:a,:s,:b,1,:u,:f,:u,:f)", t=tipo, c=cat, n=nombre, a=area, s=asunto, b=cuerpo, u=usuario, f=ts)
            registrar_auditoria("Plantilla creada", "%s · %s" % (usuario, nombre))
            session["ti_msg"] = ("ok", "Plantilla creada.")
    except Exception as e:
        print("plantillas guardar:", e)
        session["ti_msg"] = ("err", "No se pudo guardar la plantilla.")
    return redirect("/plantillas")


@app.route("/plantillas/<int:pid>/estado", methods=["POST"])
def plt_estado(pid):
    g = _plt_guard()
    if g is not None:
        return g
    if _ti_csrf_ok():
        try:
            f = _c48_q("SELECT activa, nombre FROM plantillas_texto WHERE id=:i", i=pid)
            if f:
                nuevo = 0 if int(f[0][0] or 0) else 1
                _c48_x("UPDATE plantillas_texto SET activa=:a, actualizado_por=:u, actualizado_en=:f WHERE id=:i",
                       a=nuevo, u=session.get("usuario") or "", f=ahora().strftime("%Y-%m-%d %H:%M:%S"), i=pid)
                registrar_auditoria("Plantilla %s" % ("activada" if nuevo else "desactivada"), "%s · #%d · %s" % (session.get("usuario") or "", pid, f[0][1]))
        except Exception as e:
            print("plantillas estado:", e)
    return redirect("/plantillas")


def _plt_registrar_menus():
    """Permisos y menú: Calidad y Gerencia administran; el resto solo usa el selector."""
    _MODULOS_POR_AREA["/plantillas"] = (("Calidad", "Gerente", "Administrador"), ())
    item = ("/plantillas", "📝 Plantillas de textos", "Textos base para responder PQR y crear documentos.")
    try:
        lista = _AREA_MODULOS_EXTRA.setdefault("Calidad", [])
        if item[0] not in {x[0] for x in lista}:
            lista.append(item)
        if not any(g[0] == "Plantillas y textos" for g in _MODULOS_GERENCIA):
            _MODULOS_GERENCIA.append(("Plantillas y textos", [(item[0], item[1])]))
            _GRUPOS_MODULOS_GERENCIA.append(("Plantillas y textos", (item[0],)))
        _REPARTO_ETIQUETAS.setdefault(item[0], item[1])
    except Exception as e:
        print("plantillas menus:", e)


_plt_registrar_menus()


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
    for nombre, frag in (("ancla __main__", ANCLA), ("formulario de respuesta PQR", PQR_VIEJO)):
        n = txt.count(frag)
        if n != 1:
            print("ABORTADO: `%s` aparece %d veces (debe ser 1). No se modificó nada." % (nombre, n))
            sys.exit(1)
    for req in ("def _dca_form(", "_DCA_CATEGORIAS", "def _c48_q(", "def _c48_x(", "def _ti_csrf_ok(", "def _acc_csrf(", "_TI_CSS",
                "def _dcx_area_de_rol(", "def _dcx_area_titulo(", "def _elab_usuario(", "_MODULOS_POR_AREA", "_AREA_MODULOS_EXTRA",
                "_MODULOS_GERENCIA", "_GRUPOS_MODULOS_GERENCIA", "_REPARTO_ETIQUETAS", "def formatear_radicado("):
        if req not in txt:
            print("ABORTADO: falta `%s` en app.py; versión no compatible. No se modificó nada." % req)
            sys.exit(1)
    nuevo = txt.replace(PQR_VIEJO, PQR_NUEVO, 1)
    nuevo = nuevo.replace(ANCLA, BLOQUE + ANCLA, 1)
    try:
        ast.parse(nuevo)
    except SyntaxError as e:
        print("ABORTADO: el resultado tendría un error de sintaxis (línea %s). No se modificó nada." % e.lineno)
        sys.exit(1)
    salida = nuevo.replace("\n", "\r\n") if crlf else nuevo
    respaldo = RUTA + ".bak_plantillas"
    shutil.copy2(RUTA, respaldo)
    tmp = RUTA + ".tmp_plt"
    with open(tmp, "wb") as f:
        f.write(salida.encode("utf-8"))
    os.replace(tmp, RUTA)
    print("Listo.")
    print(" - Respaldo: %s" % respaldo)
    print(" - Selector de plantillas en el detalle de PQR y en Documentos > Nuevo")
    print(" - Administración en /plantillas (Gerencia y Calidad); 13 plantillas de arranque")
    print("Reinicie la app / haga deploy para que tome efecto.")


if __name__ == "__main__":
    main()

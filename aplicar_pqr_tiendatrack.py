# -*- coding: utf-8 -*-
"""
aplicar_pqr_tiendatrack.py

Módulo de PQR para TiendaTrack, conectado por API con la app TiendaTrack.

Aprovecha lo que ya tiene PROCSIS: las PQR de TiendaTrack son tickets normales (TicketPQR) con
producto_id = "tiendatrack", así que se gestionan con el MISMO panel de Soluciones (respuesta,
estados, PDF, CSAT, historial). Este módulo agrega:

  API (cabecera X-API-Key = TIENDATRACK_API_KEY, la misma que ya usas):
    GET  /api/v1/tiendatrack/pqr/catalogo            tipos y categorías para los menús de la app
    POST /api/v1/tiendatrack/pqr                     radicar una PQR (idempotente con referencia_externa)
    GET  /api/v1/tiendatrack/pqr?email=              lista de PQR del negocio
    GET  /api/v1/tiendatrack/pqr/<radicado>?email=   detalle + respuesta oficial
    POST /api/v1/tiendatrack/pqr/<radicado>/mensaje  mensaje de seguimiento del cliente

  Panel interno:  /soluciones-procsis/pqr-tiendatrack  (y /soporte/pqr-tiendatrack, solo consulta)
    KPIs (abiertas, por vencer, vencidas, resueltas), filtros y enlace a la gestión de cada caso.

  Extras: 5 plantillas de respuesta específicas de TiendaTrack (si ya aplicaste las plantillas) y
  que el PDF de respuesta diga «TiendaTrack» en estos casos (si ya aplicaste el PDF nuevo).

Seguridad del cambio: agrega código y UNA tabla nueva (tt_pqr_ext, CREATE TABLE IF NOT EXISTS);
no modifica tablas ni datos existentes. Respaldo app.py.bak_pqr_tiendatrack, valida sintaxis,
conserva saltos de línea y no se aplica dos veces.

Uso:   python aplicar_pqr_tiendatrack.py        (usa ./app.py)
"""
import ast
import os
import shutil
import sys

MARCA = "[PROCSIS-PQR-TIENDATRACK]"
RUTA = sys.argv[1] if len(sys.argv) > 1 else "app.py"
ANCLA = 'if __name__ == "__main__":'
PDF_INI = "def generar_pdf_respuesta_pqr(t):"
PDF_FIN = "def enviar_correo_pqr("
PDF_VIEJO = "    producto = nombre_producto()\n"
PDF_NUEVO = '    producto = "TiendaTrack" if getattr(t, "producto_id", "") == "tiendatrack" else nombre_producto()\n'

BLOQUE = r'''# ══════════════════════════════════════════════════════════════════
# [PROCSIS-PQR-TIENDATRACK]
# PQR de TiendaTrack: API para la app + panel interno. Los casos son TicketPQR con producto_id="tiendatrack".
# ══════════════════════════════════════════════════════════════════
_PTT_INIT = {"ok": False}
_PTT_TIPOS = {"peticion": "Petición", "queja": "Queja", "reclamo": "Reclamo", "sugerencia": "Sugerencia", "felicitacion": "Felicitación"}
_PTT_CATS = {
    "facturacion_plan": ("Facturación y plan", "MEDIA"), "pagos": ("Pagos y cobros", "ALTA"),
    "acceso_cuenta": ("Acceso y cuenta", "ALTA"), "falla_app": ("Falla de la aplicación", "ALTA"),
    "inventario": ("Inventario y productos", "MEDIA"), "ventas_pos": ("Ventas y caja", "MEDIA"),
    "facturacion_dian": ("Facturación electrónica DIAN", "MEDIA"), "datos_privacidad": ("Datos personales y privacidad", "ALTA"),
    "demo": ("Demo y prueba gratuita", "BAJA"), "otro": ("Otro", "MEDIA"),
}
_PTT_CERRADOS = ("CERRADA", "CERRADO", "RESUELTA")
_PTT_COLS = ("t.id, t.radicado, t.ticket, t.tipo_pqr, t.subtipo, t.objeto, t.hechos, t.estado, t.prioridad, t.fecha, t.hora, "
             "t.fecha_estimada, t.fecha_respuesta, t.descripcion_cierre, t.fecha_cierre, COALESCE(e.categoria,'otro'), COALESCE(e.ref_externa,'')")
_PTT_FROM = "FROM tickets_pqr t LEFT JOIN tt_pqr_ext e ON e.ticket_id=t.id"


def _ptt_init():
    if _PTT_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS tt_pqr_ext (id %s, ticket_id INTEGER, cuenta_id INTEGER DEFAULT 0, categoria VARCHAR(30) DEFAULT 'otro', "
            "ref_externa VARCHAR(60) DEFAULT '', email VARCHAR(160) DEFAULT '', negocio VARCHAR(160) DEFAULT '', nombre VARCHAR(160) DEFAULT '', "
            "telefono VARCHAR(30) DEFAULT '', creado_en VARCHAR(20) DEFAULT '')" % pk))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_tt_pqr_ext_ticket ON tt_pqr_ext (ticket_id)"))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_tt_pqr_ext_email ON tt_pqr_ext (email, ref_externa)"))
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("ptt init:", e)
        return
    _PTT_INIT["ok"] = True
    _ptt_plantillas()


_PTT_PLANTILLAS = (
    ("queja", "TT · Queja · Falla en la aplicación TiendaTrack",
     "En atención a su queja, nuestro equipo técnico revisó el incidente reportado en la aplicación TiendaTrack de {motivo}.\n\n"
     "Hallazgo: [COMPLETAR: causa identificada].\nSolución aplicada: [COMPLETAR: corrección y fecha en que quedó disponible].\n\n"
     "Le pedimos actualizar la aplicación y probar nuevamente. Si el problema persiste, responda a este mismo radicado ({radicado}) "
     "indicando el dispositivo y los pasos que realizó, para continuar el análisis.\n\nLamentamos los inconvenientes ocasionados."),
    ("reclamo", "TT · Reclamo · Cobro o plan de TiendaTrack",
     "En atención a su reclamo, revisamos el historial de pagos y el plan asociado a su negocio.\n\n"
     "Resultado de la revisión: [COMPLETAR: pagos y fechas verificados].\nDecisión: [COMPLETAR: se acepta / se rechaza y el motivo].\n"
     "Ajuste aplicado: [COMPLETAR: valor, concepto y fecha, si corresponde].\n\n"
     "Si desea que revisemos nuevamente su caso, responda a este radicado ({radicado}) con el comprobante de pago."),
    ("peticion", "TT · Petición · Acceso con Google",
     "En atención a su solicitud sobre el acceso a TiendaTrack, verificamos la cuenta asociada a su correo.\n\n"
     "Estado de la cuenta: [COMPLETAR: activa / en demo / vencida].\nAcción realizada: [COMPLETAR].\n\n"
     "Para ingresar utilice el mismo correo de Google con el que se registró. Si el inconveniente continúa, responda a este radicado ({radicado})."),
    ("peticion", "TT · Petición · Facturación electrónica DIAN",
     "En atención a su consulta sobre facturación electrónica: por ahora TiendaTrack entrega comprobantes de pago y no emite factura "
     "electrónica ante la DIAN.\n\n[COMPLETAR: alternativa recomendada o fecha estimada de disponibilidad, si la hay].\n\n"
     "Si necesita más información, puede responder a este mismo radicado ({radicado})."),
    ("general", "TT · Sugerencia o felicitación",
     "Gracias por escribirnos y por confiar en TiendaTrack para su negocio.\n\n"
     "[COMPLETAR: reconocimiento de la sugerencia o felicitación y, si aplica, qué haremos con ella].\n\n"
     "Sus comentarios nos ayudan a mejorar el producto. Estamos atentos a lo que necesite."),
)


def _ptt_plantillas():
    """Agrega las plantillas de TiendaTrack si el módulo de plantillas ya está instalado (idempotente)."""
    try:
        if "_plt_init" not in globals():
            return
        _plt_init()
        ts = ahora().strftime("%Y-%m-%d %H:%M:%S")
        for cat, nombre, cuerpo in _PTT_PLANTILLAS:
            if not _c48_q("SELECT id FROM plantillas_texto WHERE nombre=:n", n=nombre):
                _c48_x("INSERT INTO plantillas_texto (tipo, categoria, nombre, area, asunto, cuerpo, activa, creado_por, creado_en, actualizado_por, actualizado_en) "
                       "VALUES ('pqr',:c,:n,'General','',:b,1,'sistema',:f,'sistema',:f)", c=cat, n=nombre, b=cuerpo, f=ts)
    except Exception as e:
        db.session.rollback()
        print("ptt plantillas:", e)


def _ptt_norm(s):
    import unicodedata
    return "".join(ch for ch in unicodedata.normalize("NFD", str(s or "").lower()) if unicodedata.category(ch) != "Mn").strip()


def _ptt_pub(r):
    """Fila -> JSON seguro para la app (nunca expone notas internas ni usuarios del personal)."""
    est = (r[7] or "").upper()
    cat = r[15] if r[15] in _PTT_CATS else "otro"
    return {"radicado": r[1], "radicado_formato": formatear_radicado(r[1]), "ticket": r[2], "tipo": r[3], "categoria": cat,
            "categoria_nombre": _PTT_CATS[cat][0], "asunto": r[5] or r[4] or "", "descripcion": r[6] or "", "estado": r[7] or "",
            "abierto": est not in _PTT_CERRADOS, "prioridad": r[8] or "", "fecha_radicacion": r[9] or "", "hora_radicacion": r[10] or "",
            "fecha_limite": r[11] or "", "respuesta": r[13] or "", "fecha_respuesta": r[12] or "", "fecha_cierre": r[14] or "",
            "referencia_externa": r[16] or ""}


def _ptt_api_guard(nombre, limite=120):
    ok_rl, _e = _rate_limit(nombre, limite, 60)
    if not ok_rl:
        return {"ok": False, "error": "demasiadas solicitudes"}, 429
    if not _pf_api_ok():
        return {"ok": False, "error": "no autorizado"}, 401
    _tt_init()
    _ptt_init()
    return None


def _ptt_buscar(radicado, email):
    rad = "".join(ch for ch in str(radicado or "") if ch.isdigit())
    if not rad or "@" not in email:
        return None
    f = _c48_q("SELECT " + _PTT_COLS + " " + _PTT_FROM + " WHERE t.radicado=:r AND t.producto_id='tiendatrack' AND lower(t.email)=:e", r=rad, e=email)
    return f[0] if f else None


@app.route("/api/v1/tiendatrack/pqr/catalogo")
def ptt_api_catalogo():
    g = _ptt_api_guard("tt-pqr-cat")
    if g is not None:
        return g
    return {"ok": True, "tipos": [{"clave": k, "nombre": v} for k, v in _PTT_TIPOS.items()],
            "categorias": [{"clave": k, "nombre": v[0]} for k, v in _PTT_CATS.items()], "dias_habiles_respuesta": 15}, 200


@app.route("/api/v1/tiendatrack/pqr", methods=["POST"])
def ptt_api_crear():
    g = _ptt_api_guard("tt-pqr-new", 60)
    if g is not None:
        return g
    d = request.get_json(silent=True) or {}
    email = _pf_txt(d.get("email"), 160).lower()
    cuenta = _tt_cuenta_por("email", email) if "@" in email else None
    if not cuenta:
        return {"ok": False, "error": "cuenta no encontrada: registre primero el negocio con el evento «registro»"}, 404
    tipo = _ptt_norm(d.get("tipo"))
    if tipo not in _PTT_TIPOS:
        return {"ok": False, "error": "tipo inválido", "validos": list(_PTT_TIPOS)}, 400
    cat = _ptt_norm(d.get("categoria")) or "otro"
    cat = cat if cat in _PTT_CATS else "otro"
    asunto = _pf_txt(d.get("asunto"), 120)
    descripcion = str(d.get("descripcion") or "").replace("\r", "").strip()[:4000]
    if not asunto or len(descripcion) < 10:
        return {"ok": False, "error": "asunto y descripcion (mínimo 10 caracteres) son obligatorios"}, 400
    ref = _pf_txt(d.get("referencia_externa"), 60)
    try:
        if ref:
            ya = _c48_q("SELECT ticket_id FROM tt_pqr_ext WHERE email=:e AND ref_externa=:r", e=email, r=ref)
            if ya:
                f = _c48_q("SELECT " + _PTT_COLS + " " + _PTT_FROM + " WHERE t.id=:i", i=ya[0][0])
                if f:
                    return {"ok": True, "duplicado": True, "radicado": f[0][1], "radicado_formato": formatear_radicado(f[0][1]), "ticket": f[0][2],
                            "estado": f[0][7], "fecha_limite": f[0][11] or ""}, 200
        prioridad = _pf_txt(d.get("prioridad"), 10).upper()
        prioridad = prioridad if prioridad in ("CRITICA", "ALTA", "MEDIA", "BAJA") else _PTT_CATS[cat][1]
        hoy = ahora().strftime("%Y-%m-%d")
        try:
            limite = _fecha_limite_pqr_habiles(hoy, 15)
        except Exception:
            limite = ""
        negocio = _pf_txt(d.get("negocio") or cuenta.get("negocio"), 160)
        nombre = _pf_txt(d.get("nombre") or cuenta.get("nombre"), 160)
        rad, tick = _gen_radicado_pqr("33")
        t = TicketPQR(radicado=rad, ticket=tick, producto_id="tiendatrack", fecha=hoy, hora=ahora().strftime("%H:%M"), email=email,
                      razon_social=(nombre or negocio)[:220], nombre_colegio=negocio[:220], tipo_pqr=_PTT_TIPOS[tipo], subtipo=_PTT_CATS[cat][0],
                      objeto=asunto, hechos=descripcion, estado="RADICADO", canal="TIENDATRACK", medio_ingreso="API TiendaTrack",
                      prioridad=prioridad, fecha_estimada=limite, autoriza_email=True, creado_por="TiendaTrack API")
        db.session.add(t)
        db.session.commit()
        _c48_x("INSERT INTO tt_pqr_ext (ticket_id, cuenta_id, categoria, ref_externa, email, negocio, nombre, telefono, creado_en) "
               "VALUES (:t,:c,:k,:r,:e,:n,:no,:te,:f)", t=t.id, c=cuenta["id"], k=cat, r=ref, e=email, n=negocio, no=nombre,
               te=_pf_txt(d.get("telefono") or cuenta.get("telefono"), 30), f=_tt_ahora_s())
        try:
            registrar_historial_pqr(t.id, "RADICACION", "Radicada desde la app TiendaTrack (%s · %s)" % (_PTT_TIPOS[tipo], _PTT_CATS[cat][0]), "RADICADO")
            db.session.commit()
        except Exception as ex:
            db.session.rollback()
            print("ptt historial:", ex)
        try:
            _notif_crear(["Soluciones"], "Nueva PQR de TiendaTrack: %s" % asunto[:60], "%s · %s · %s" % (_PTT_TIPOS[tipo], _PTT_CATS[cat][0], negocio or email),
                         "/soluciones-procsis/pqr-tiendatrack", nivel="alerta" if prioridad in ("CRITICA", "ALTA") else "info")
        except Exception:
            pass
        registrar_auditoria("TiendaTrack API · PQR", "radicado=%s email=%s tipo=%s" % (rad, email, tipo))
        base = _base_publica().rstrip("/")
        return {"ok": True, "duplicado": False, "radicado": rad, "radicado_formato": formatear_radicado(rad), "ticket": tick, "estado": "RADICADO",
                "prioridad": prioridad, "fecha_limite": limite, "consulta_publica": "%s/pqr/consulta?clave=%s" % (base, rad)}, 201
    except Exception as e:
        try:
            db.session.rollback()
        except Exception:
            pass
        print("ptt api crear:", repr(e)[:300])
        return {"ok": False, "error": "no se pudo radicar la PQR"}, 500


@app.route("/api/v1/tiendatrack/pqr", methods=["GET"])
def ptt_api_listar():
    g = _ptt_api_guard("tt-pqr-list")
    if g is not None:
        return g
    email = _pf_txt(request.args.get("email"), 160).lower()
    if "@" not in email:
        return {"ok": False, "error": "email requerido"}, 400
    try:
        lim = max(1, min(50, int(request.args.get("limite") or 20)))
    except Exception:
        lim = 20
    filas = _c48_q("SELECT " + _PTT_COLS + " " + _PTT_FROM + " WHERE t.producto_id='tiendatrack' AND lower(t.email)=:e ORDER BY t.id DESC LIMIT %d" % lim, e=email)
    items = [_ptt_pub(r) for r in filas]
    for it in items:
        it.pop("descripcion", None)   # el listado va liviano; el detalle trae el texto completo
    return {"ok": True, "total": len(items), "pqr": items}, 200


@app.route("/api/v1/tiendatrack/pqr/<radicado>", methods=["GET"])
def ptt_api_detalle(radicado):
    g = _ptt_api_guard("tt-pqr-det")
    if g is not None:
        return g
    r = _ptt_buscar(radicado, _pf_txt(request.args.get("email"), 160).lower())
    if not r:
        return {"ok": False, "error": "PQR no encontrada para ese correo"}, 404
    return {"ok": True, "pqr": _ptt_pub(r)}, 200


@app.route("/api/v1/tiendatrack/pqr/<radicado>/mensaje", methods=["POST"])
def ptt_api_mensaje(radicado):
    g = _ptt_api_guard("tt-pqr-msg", 60)
    if g is not None:
        return g
    d = request.get_json(silent=True) or {}
    email = _pf_txt(d.get("email"), 160).lower()
    r = _ptt_buscar(radicado, email)
    if not r:
        return {"ok": False, "error": "PQR no encontrada para ese correo"}, 404
    if (r[7] or "").upper() in _PTT_CERRADOS:
        return {"ok": False, "error": "el caso está cerrado; radique una PQR nueva"}, 409
    texto = str(d.get("mensaje") or "").replace("\r", "").strip()[:1000]
    if len(texto) < 3:
        return {"ok": False, "error": "mensaje requerido"}, 400
    ok_rl, _e = _rate_limit("tt-pqr-msg-" + r[1], 5, 3600)
    if not ok_rl:
        return {"ok": False, "error": "demasiados mensajes en este caso; espere un momento"}, 429
    try:
        registrar_historial_pqr(r[0], "CLIENTE", texto, r[7] or "")
        db.session.commit()
        _notif_crear(["Soluciones"], "Mensaje del cliente en PQR %s" % formatear_radicado(r[1]), texto[:120], "/soluciones-procsis/pqr/%d" % r[0])
    except Exception as e:
        db.session.rollback()
        print("ptt api mensaje:", repr(e)[:300])
        return {"ok": False, "error": "no se pudo registrar el mensaje"}, 500
    return {"ok": True}, 201


# ───────────── Panel interno: PQR de TiendaTrack ─────────────
@app.route("/soluciones-procsis/pqr-tiendatrack")
@app.route("/soporte/pqr-tiendatrack")
def ptt_panel():
    g = _pqr_guard("ver")
    if g:
        return g
    _tt_init()
    _ptt_init()
    base = _pqr_base()
    hoy = ahora().strftime("%Y-%m-%d")
    filas = _c48_q("SELECT " + _PTT_COLS + ", t.razon_social, t.nombre_colegio, t.email " + _PTT_FROM +
                   " WHERE t.producto_id='tiendatrack' ORDER BY t.id DESC LIMIT 300")
    abiertas = [r for r in filas if (r[7] or "").upper() not in _PTT_CERRADOS]
    vencidas = [r for r in abiertas if (r[11] or "9999") < hoy]
    por_vencer = [r for r in abiertas if hoy <= (r[11] or "9999") <= _tt_sumar_dias(hoy + " 00:00:00", 3)[:10]]
    resueltas_mes = [r for r in filas if (r[7] or "").upper() in _PTT_CERRADOS and (r[14] or "")[:7] == hoy[:7]]
    filtro = request.args.get("ver") or "abiertas"
    cat_f = request.args.get("cat") or ""
    lista = {"abiertas": abiertas, "vencidas": vencidas, "por_vencer": por_vencer, "cerradas": [r for r in filas if r not in abiertas], "todas": filas}.get(filtro, abiertas)
    if cat_f in _PTT_CATS:
        lista = [r for r in lista if r[15] == cat_f]
    kpi = lambda n, et, col="#000F66": "<div style='background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:14px 16px;min-width:150px'><div style='font-size:28px;font-weight:800;color:%s'>%d</div><div style='font-size:12px;color:#64748b'>%s</div></div>" % (col, n, et)
    tabs = "".join("<a class='btn' style='%s' href='?ver=%s'>%s</a> " % ("" if filtro == k else "opacity:.6", k, n) for k, n in
                   (("abiertas", "Abiertas"), ("por_vencer", "Por vencer (3 días)"), ("vencidas", "Vencidas"), ("cerradas", "Cerradas"), ("todas", "Todas")))
    cats = "<form method='GET' style='display:inline'><input type='hidden' name='ver' value='%s'><select name='cat' onchange='this.form.submit()'><option value=''>Todas las categorías</option>%s</select></form>" % (
        _esc(filtro), "".join("<option value='%s'%s>%s</option>" % (k, " selected" if k == cat_f else "", _esc(v[0])) for k, v in _PTT_CATS.items()))
    trs = ""
    for r in lista[:200]:
        venc = (r[7] or "").upper() not in _PTT_CERRADOS and (r[11] or "9999") < hoy
        trs += ("<tr><td><b>%s</b><br><small>%s</small></td><td>%s<br><small>%s</small></td><td>%s<br><small>%s</small></td><td>%s</td>"
                "<td>%s</td><td%s>%s</td><td><a class='btn' href='%s/%d'>Gestionar</a></td></tr>") % (
            _esc(formatear_radicado(r[1])), _esc(r[9] or ""), _esc(r[3] or ""), _esc(_PTT_CATS.get(r[15], _PTT_CATS["otro"])[0]),
            _esc(r[18] or r[17] or "—"), _esc(r[19] or ""), _esc(r[5] or r[4] or "")[:90], _esc(r[7] or ""),
            " style='color:#b91c1c;font-weight:700'" if venc else "", _esc(r[11] or "—"), base, r[0])
    cuerpo = ("<h1>📨 PQR de TiendaTrack</h1><p class='sub'>Casos radicados desde la app TiendaTrack por la API. Se gestionan igual que las demás PQR: "
              "respuesta, estado, PDF y encuesta.</p>"
              "<div style='display:flex;gap:12px;flex-wrap:wrap;margin:12px 0'>%s%s%s%s</div>"
              "<div style='margin:10px 0'>%s %s</div>"
              "<div style='overflow-x:auto;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:6px'><table style='width:100%%'>"
              "<tr><th>Radicado</th><th>Tipo</th><th>Negocio</th><th>Asunto</th><th>Estado</th><th>Vence</th><th></th></tr>%s</table></div>") % (
        kpi(len(abiertas), "Abiertas"), kpi(len(por_vencer), "Por vencer (3 días)", "#b45309"), kpi(len(vencidas), "Vencidas", "#b91c1c"),
        kpi(len(resueltas_mes), "Resueltas este mes", "#0e9f5b"), tabs, cats,
        trs or "<tr><td colspan='7'>No hay casos en esta vista.</td></tr>")
    return _cb_pagina("PQR de TiendaTrack", cuerpo)


def _ptt_registrar():
    """Menús: Gerencia (grupo TiendaTrack), Soluciones y Soporte. Solo visible con el producto TiendaTrack/Todo."""
    global _PFM_SOLO_TIENDATRACK_PREFIJOS
    rutas = ("/soluciones-procsis/pqr-tiendatrack", "/soporte/pqr-tiendatrack")
    try:
        if rutas[0] not in _PFM_SOLO_TIENDATRACK_PREFIJOS:
            _PFM_SOLO_TIENDATRACK_PREFIJOS = tuple(_PFM_SOLO_TIENDATRACK_PREFIJOS) + rutas
    except Exception as e:
        print("ptt prefijos:", e)
    try:
        for rol, href in (("Soluciones", rutas[0]), ("Soporte", rutas[1])):
            lista = _REPARTO_MENU_EXTRA.setdefault(rol, [])
            if href not in {x[0] for x in lista}:
                lista.append((href, "📨 PQR de TiendaTrack"))
    except Exception as e:
        print("ptt menu roles:", e)
    try:
        for nombre, lista in _MODULOS_GERENCIA:
            if nombre == "TiendaTrack":
                if rutas[0] not in {x[0] for x in lista}:
                    lista.append((rutas[0], "📨 PQR de TiendaTrack"))
                break
        else:
            _MODULOS_GERENCIA.append(("TiendaTrack", [(rutas[0], "📨 PQR de TiendaTrack")]))
    except Exception as e:
        print("ptt menu gerencia:", e)


_ptt_registrar()


'''

DOC_API = r"""# API de PQR · TiendaTrack ⇄ PROCSIS

Base: `https://procsishq.com` (el dominio público de PROCSIS)
Autenticación: cabecera `X-API-Key: <TIENDATRACK_API_KEY>` (la misma llave de los demás endpoints de TiendaTrack).
Formato: JSON UTF-8. Límite: 60–120 solicitudes por minuto. Todas las respuestas traen `"ok": true|false`.

> Antes de radicar una PQR el negocio debe existir en PROCSIS (evento `registro` de `POST /api/v1/tiendatrack/evento`).

## 1. Catálogo (para llenar los menús de la app)
`GET /api/v1/tiendatrack/pqr/catalogo`
```json
{"ok": true, "tipos": [{"clave":"queja","nombre":"Queja"}, ...],
 "categorias": [{"clave":"falla_app","nombre":"Falla de la aplicación"}, ...], "dias_habiles_respuesta": 15}
```
Tipos: `peticion`, `queja`, `reclamo`, `sugerencia`, `felicitacion`.
Categorías: `facturacion_plan`, `pagos`, `acceso_cuenta`, `falla_app`, `inventario`, `ventas_pos`, `facturacion_dian`, `datos_privacidad`, `demo`, `otro`.

## 2. Radicar una PQR
`POST /api/v1/tiendatrack/pqr`
```json
{"email": "tienda@correo.com", "tipo": "queja", "categoria": "falla_app",
 "asunto": "No cargan los productos", "descripcion": "Desde ayer la lista de productos aparece vacía...",
 "referencia_externa": "pqr-local-123", "prioridad": "ALTA"}
```
Opcionales: `negocio`, `nombre`, `telefono`, `prioridad` (CRITICA|ALTA|MEDIA|BAJA; si no se envía se asigna según la categoría).
`referencia_externa` hace la llamada **idempotente**: si la app reintenta con el mismo valor, devuelve la misma PQR (`"duplicado": true`) en vez de crear otra.

Respuesta `201`:
```json
{"ok": true, "duplicado": false, "radicado": "202633000001", "radicado_formato": "2026-33-000001", "ticket": "AB12CD",
 "estado": "RADICADO", "prioridad": "ALTA", "fecha_limite": "2026-12-09",
 "consulta_publica": "https://procsishq.com/pqr/consulta?clave=202633000001"}
```

## 3. Listar las PQR de un negocio
`GET /api/v1/tiendatrack/pqr?email=tienda@correo.com&limite=20` → `{"ok": true, "total": 2, "pqr": [ {...}, {...} ]}`

## 4. Ver el detalle y la respuesta oficial
`GET /api/v1/tiendatrack/pqr/202633000001?email=tienda@correo.com`
```json
{"ok": true, "pqr": {"radicado":"202633000001","radicado_formato":"2026-33-000001","ticket":"AB12CD","tipo":"Queja",
 "categoria":"falla_app","categoria_nombre":"Falla de la aplicación","asunto":"...","descripcion":"...","estado":"EN PROCESO",
 "abierto": true, "prioridad":"ALTA","fecha_radicacion":"2026-11-17","fecha_limite":"2026-12-09",
 "respuesta":"","fecha_respuesta":"","fecha_cierre":"","referencia_externa":"pqr-local-123"}}
```
`respuesta` queda con el texto oficial cuando Soluciones responde. `abierto` es `false` cuando el caso está resuelto o cerrado.
El correo debe ser el mismo con el que se radicó; si no, responde `404` (así un negocio no ve PQR de otro).
Sugerencia: consultar este endpoint al abrir la pantalla de la PQR (o cada pocos minutos mientras esté abierta); no hace falta webhook.

## 5. Mensaje de seguimiento del cliente
`POST /api/v1/tiendatrack/pqr/202633000001/mensaje`
```json
{"email": "tienda@correo.com", "mensaje": "Adjunto más detalle: ocurre solo en el celular."}
```
Respuesta `201`. Si el caso ya está cerrado responde `409` (hay que radicar una PQR nueva). Máximo 5 mensajes por hora y por caso.

## Códigos de error
`400` datos inválidos · `401` llave incorrecta · `404` cuenta o PQR no encontrada · `409` caso cerrado · `429` demasiadas solicitudes · `500` error interno.

## Ejemplo con curl
```bash
curl -X POST https://procsishq.com/api/v1/tiendatrack/pqr \
  -H "X-API-Key: $TIENDATRACK_API_KEY" -H "Content-Type: application/json" \
  -d '{"email":"tienda@correo.com","tipo":"queja","categoria":"falla_app","asunto":"No cargan los productos","descripcion":"Desde ayer la lista aparece vacía"}'
```
"""


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
        print("ABORTADO: `%s` aparece %d veces (debe ser 1). No se modificó nada." % (ANCLA, txt.count(ANCLA)))
        sys.exit(1)
    for req in ("def _pf_api_ok(", "def _pf_txt(", "def _tt_cuenta_por(", "def _tt_init(", "def _tt_ahora_s(", "def _tt_sumar_dias(", "def _gen_radicado_pqr(",
                "def _fecha_limite_pqr_habiles(", "def registrar_historial_pqr(", "def _notif_crear(", "def _rate_limit(", "def _base_publica(",
                "def _pqr_guard(", "def _pqr_base(", "def _cb_pagina(", "def formatear_radicado(", "class TicketPQR(", "_PFM_SOLO_TIENDATRACK_PREFIJOS",
                "_REPARTO_MENU_EXTRA", "_MODULOS_GERENCIA"):
        if req not in txt:
            print("ABORTADO: falta `%s` en app.py; versión no compatible. No se modificó nada." % req)
            sys.exit(1)
    nuevo = txt
    # PDF de respuesta: «TiendaTrack» como producto en estos casos (solo si ya está el PDF nuevo)
    aviso_pdf = ""
    if "[PROCSIS-PQR-PDF-V2]" in nuevo:
        i, j = nuevo.index(PDF_INI), nuevo.index(PDF_FIN)
        if j > i and nuevo[i:j].count(PDF_VIEJO) == 1:
            nuevo = nuevo[:i] + nuevo[i:j].replace(PDF_VIEJO, PDF_NUEVO, 1) + nuevo[j:]
        else:
            aviso_pdf = "No pude ajustar el nombre del producto en el PDF (estructura distinta); el módulo funciona igual."
    else:
        aviso_pdf = "Aplica primero aplicar_pqr_pdf_nuevo_diseno.py para que el PDF diga «TiendaTrack» en estos casos."
    nuevo = nuevo.replace(ANCLA, BLOQUE + ANCLA, 1)
    try:
        ast.parse(nuevo)
    except SyntaxError as e:
        print("ABORTADO: el resultado tendría un error de sintaxis (línea %s). No se modificó nada." % e.lineno)
        sys.exit(1)
    salida = nuevo.replace("\n", "\r\n") if crlf else nuevo
    respaldo = RUTA + ".bak_pqr_tiendatrack"
    shutil.copy2(RUTA, respaldo)
    tmp = RUTA + ".tmp_ptt"
    with open(tmp, "wb") as f:
        f.write(salida.encode("utf-8"))
    os.replace(tmp, RUTA)
    # Documentación de la API para quien programa la app TiendaTrack (archivo nuevo junto a app.py)
    doc = os.path.join(os.path.dirname(os.path.abspath(RUTA)), "API_PQR_TiendaTrack.md")
    try:
        with open(doc, "w", encoding="utf-8") as f:
            f.write(DOC_API)
    except Exception as e:
        print("No se pudo escribir la documentación de la API:", e)
    print("Listo.")
    print(" - Respaldo: %s" % respaldo)
    print(" - API: /api/v1/tiendatrack/pqr (catalogo, crear, listar, detalle, mensaje)")
    print(" - Panel: /soluciones-procsis/pqr-tiendatrack")
    print(" - Documentación para el desarrollador de la app: %s" % doc)
    if aviso_pdf:
        print(" - Nota: %s" % aviso_pdf)
    print("Reinicie la app / haga deploy para que tome efecto.")


if __name__ == "__main__":
    main()

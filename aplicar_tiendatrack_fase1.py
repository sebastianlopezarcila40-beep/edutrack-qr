#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_tiendatrack_fase1.py   (Fase 1 · sistema interno de PROCSIS)
Uso:  python aplicar_tiendatrack_fase1.py            (busca app.py en la carpeta actual)
      python aplicar_tiendatrack_fase1.py ruta/app.py

Qué agrega a app.py (antes guarda la copia app.py.bak_tiendatrack):
  1) Selector de producto en el menú lateral de las áreas administrativas: Todo PROCSIS · EduTrack · TiendaTrack.
  2) Módulo «TiendaTrack» en Gerencia (/gerencia/tiendatrack): panel, planes (crear/editar/ocultar), negocios con demo,
     pagos y pasarela. Tablas nuevas: tt_planes, tt_cuentas, tt_pagos (se crean solas; no toca las existentes).
  3) Zona pública /tt: página de TiendaTrack, acceso con Google, demo gratis de 15 días y pago con Wompi al vencer.
  4) Conexión con el webhook de Wompi que ya tienes: los pagos con referencia TT- activan el plan automáticamente.
  5) API para la futura app: /api/v1/tiendatrack/cuenta?email=...  (cabecera X-API-Key).
El registro público queda CERRADO hasta que lo abras en Gerencia → TiendaTrack → Pasarela y Google.
Es seguro correrlo dos veces: si ya está aplicado, no hace nada.
"""
import sys, os, shutil

MARCA = "# [procsis-tiendatrack-fase1]"
BLOQUE = r'''# [procsis-tiendatrack-fase1]
# ═════════════ TIENDATRACK · FASE 1 (sistema interno): selector de producto, planes, cuentas con demo, pagos Wompi y acceso con Google ═════════════
# Variables de entorno (Railway): GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, WOMPI_PUBLIC_KEY, WOMPI_INTEGRITY_SECRET,
# WOMPI_EVENTS_SECRET (ya existe), TIENDATRACK_API_KEY. Sin ellas el módulo funciona y avisa qué falta.
import urllib.request as _tt_ur
import urllib.parse as _tt_up
import hashlib as _tt_hl
import hmac as _tt_hm
import base64 as _tt_b64

_TT_INIT = {"ok": False}
_TT_TERMINOS = "2026-11"
_TT_SEGMENTOS = (("tienda", "Tienda"), ("ropa", "Tienda de ropa"), ("restaurante", "Restaurante"), ("otro", "Otro negocio"))
_TT_SEG_PLAN = (("todos", "Todos los negocios"), ("tienda", "Tiendas"), ("ropa", "Tiendas de ropa"), ("restaurante", "Restaurantes"))
_TT_PRODUCTOS = (("todo", "Todo PROCSIS"), ("edutrack", "EduTrack"), ("tiendatrack", "TiendaTrack"))
_TT_APROBADO = ("APPROVED", "APPROVED_STATUS")
_TT_COLS_PLAN = ("id", "slug", "nombre", "descripcion", "segmento", "precio_mensual", "precio_anual", "dias_demo",
                 "max_sucursales", "max_usuarios", "max_productos", "incluye_dian", "activo", "visible_web", "orden")
_TT_COLS_CUENTA = ("id", "google_sub", "email", "nombre", "foto", "negocio", "segmento", "telefono", "plan_id", "estado",
                   "demo_inicio", "demo_fin", "activo_hasta", "acepto_terminos", "terminos_version", "nota", "creado_en", "ultimo_acceso")
_TT_COLS_PAGO = ("id", "cuenta_id", "plan_id", "periodo", "referencia", "monto_centavos", "moneda", "estado", "pasarela",
                 "id_transaccion", "factura_ref", "creado_en", "actualizado_en")


def _tt_ahora_s():
    return ahora().strftime("%Y-%m-%d %H:%M:%S")


def _tt_sumar_dias(desde_s, dias):
    from datetime import datetime, timedelta
    try:
        d = datetime.strptime((desde_s or "")[:19], "%Y-%m-%d %H:%M:%S")
    except Exception:
        d = datetime.strptime(_tt_ahora_s(), "%Y-%m-%d %H:%M:%S")
    return (d + timedelta(days=int(dias))).strftime("%Y-%m-%d %H:%M:%S")


def _tt_dias(fin_s):
    from datetime import datetime
    try:
        fin = datetime.strptime((fin_s or "")[:19], "%Y-%m-%d %H:%M:%S")
        hoy = datetime.strptime(_tt_ahora_s(), "%Y-%m-%d %H:%M:%S")
        seg = (fin - hoy).total_seconds()
        return max(0, int(-(-seg // 86400)))
    except Exception:
        return 0


def _tt_cop(v):
    try:
        return "$" + "{:,.0f}".format(int(v or 0)).replace(",", ".")
    except Exception:
        return "$0"


def _tt_fecha(s):
    meses = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
    try:
        a, m, d = (s or "")[:10].split("-")
        return "%d %s %s" % (int(d), meses[int(m) - 1], a)
    except Exception:
        return s or "—"


def _tt_int(v, defecto=0, minimo=0, maximo=10 ** 9):
    try:
        n = int(str(v).strip().replace(".", "").replace(",", "") or defecto)
    except Exception:
        n = defecto
    return max(minimo, min(maximo, n))


def _tt_dict(cols, fila):
    return {c: fila[i] for i, c in enumerate(cols)}


def _tt_init():
    if _TT_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    ddls = (
        "CREATE TABLE IF NOT EXISTS ti_config (clave VARCHAR(40) PRIMARY KEY, valor TEXT)",
        "CREATE TABLE IF NOT EXISTS tt_planes (id %s, slug VARCHAR(40) UNIQUE, nombre VARCHAR(80), descripcion TEXT DEFAULT '', "
        "segmento VARCHAR(20) DEFAULT 'todos', precio_mensual BIGINT DEFAULT 0, precio_anual BIGINT DEFAULT 0, dias_demo INTEGER DEFAULT 15, "
        "max_sucursales INTEGER DEFAULT 1, max_usuarios INTEGER DEFAULT 2, max_productos INTEGER DEFAULT 500, incluye_dian VARCHAR(20) DEFAULT 'No', "
        "activo INTEGER DEFAULT 1, visible_web INTEGER DEFAULT 0, orden INTEGER DEFAULT 0, creado_en VARCHAR(20) DEFAULT '', actualizado_en VARCHAR(20) DEFAULT '')" % pk,
        "CREATE TABLE IF NOT EXISTS tt_cuentas (id %s, google_sub VARCHAR(64) UNIQUE, email VARCHAR(160) UNIQUE, nombre VARCHAR(160) DEFAULT '', "
        "foto TEXT DEFAULT '', negocio VARCHAR(160) DEFAULT '', segmento VARCHAR(20) DEFAULT 'tienda', telefono VARCHAR(30) DEFAULT '', plan_id INTEGER DEFAULT 0, "
        "estado VARCHAR(16) DEFAULT 'demo', demo_inicio VARCHAR(20) DEFAULT '', demo_fin VARCHAR(20) DEFAULT '', activo_hasta VARCHAR(20) DEFAULT '', "
        "acepto_terminos VARCHAR(20) DEFAULT '', terminos_version VARCHAR(10) DEFAULT '', nota TEXT DEFAULT '', creado_en VARCHAR(20) DEFAULT '', ultimo_acceso VARCHAR(20) DEFAULT '')" % pk,
        "CREATE TABLE IF NOT EXISTS tt_pagos (id %s, cuenta_id INTEGER, plan_id INTEGER, periodo VARCHAR(10) DEFAULT 'mensual', referencia VARCHAR(60) UNIQUE, "
        "monto_centavos BIGINT DEFAULT 0, moneda VARCHAR(3) DEFAULT 'COP', estado VARCHAR(16) DEFAULT 'pendiente', pasarela VARCHAR(16) DEFAULT 'wompi', "
        "id_transaccion VARCHAR(80) DEFAULT '', factura_ref VARCHAR(60) DEFAULT '', creado_en VARCHAR(20) DEFAULT '', actualizado_en VARCHAR(20) DEFAULT '')" % pk,
    )
    for ddl in ddls:
        try:
            db.session.execute(text(ddl))
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            print("tt init:", e)
    try:
        n = db.session.execute(text("SELECT COUNT(*) FROM tt_planes")).scalar() or 0
        if int(n) == 0:
            ahora_s = _tt_ahora_s()
            for slug, nom, pm, pa, suc, usu, prod, orden in (("esencial", "Esencial", 29900, 299000, 1, 2, 500, 1),
                                                              ("negocio", "Negocio", 59900, 599000, 3, 5, 5000, 2),
                                                              ("pro", "Pro", 99900, 999000, 10, 20, 50000, 3)):
                db.session.execute(text(
                    "INSERT INTO tt_planes (slug, nombre, descripcion, segmento, precio_mensual, precio_anual, dias_demo, max_sucursales, max_usuarios, max_productos, "
                    "incluye_dian, activo, visible_web, orden, creado_en, actualizado_en) VALUES (:s, :n, :d, 'todos', :pm, :pa, 15, :su, :us, :pr, 'No', 1, 0, :o, :t, :t)"),
                    {"s": slug, "n": nom, "d": "Plan de ejemplo: edita nombre, precio y límites antes de publicarlo.", "pm": pm, "pa": pa,
                     "su": suc, "us": usu, "pr": prod, "o": orden, "t": ahora_s})
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("tt seed:", e)
    _TT_INIT["ok"] = True


def _tt_cfg(clave, defecto=""):
    _tt_init()
    r = _c48_q("SELECT valor FROM ti_config WHERE clave=:c", c=clave)
    return r[0][0] if r and r[0][0] not in (None, "") else defecto


def _tt_planes(solo_web=False, solo_activos=False):
    _tt_init()
    sql = "SELECT %s FROM tt_planes" % ", ".join(_TT_COLS_PLAN)
    cond = []
    if solo_web:
        cond.append("visible_web=1 AND activo=1")
    elif solo_activos:
        cond.append("activo=1")
    if cond:
        sql += " WHERE " + " AND ".join(cond)
    sql += " ORDER BY orden, id"
    return [_tt_dict(_TT_COLS_PLAN, f) for f in _c48_q(sql)]


def _tt_plan(pid):
    _tt_init()
    r = _c48_q("SELECT %s FROM tt_planes WHERE id=:i" % ", ".join(_TT_COLS_PLAN), i=_tt_int(pid))
    return _tt_dict(_TT_COLS_PLAN, r[0]) if r else None


def _tt_cuenta_por(campo, valor):
    _tt_init()
    if campo not in ("id", "email", "google_sub"):
        return None
    r = _c48_q("SELECT %s FROM tt_cuentas WHERE %s=:v" % (", ".join(_TT_COLS_CUENTA), campo), v=valor)
    return _tt_dict(_TT_COLS_CUENTA, r[0]) if r else None


def _tt_estado(c):
    """Estado real de la cuenta según las fechas: demo | activo | vencido | suspendido."""
    if not c:
        return "vencido"
    if c.get("estado") == "suspendido":
        return "suspendido"
    ahora_s = _tt_ahora_s()
    if c.get("activo_hasta") and ahora_s <= c["activo_hasta"]:
        return "activo"
    if c.get("demo_fin") and ahora_s <= c["demo_fin"]:
        return "demo"
    return "vencido"


def _tt_activar(cuenta_id, plan_id, periodo):
    """Suma 30 o 365 días de servicio. Si ya estaba activa, el nuevo periodo se acumula."""
    c = _tt_cuenta_por("id", cuenta_id)
    if not c:
        return False
    ahora_s = _tt_ahora_s()
    base = c["activo_hasta"] if (c["activo_hasta"] and c["activo_hasta"] > ahora_s) else ahora_s
    nuevo = _tt_sumar_dias(base, 365 if periodo == "anual" else 30)
    _c48_x("UPDATE tt_cuentas SET plan_id=:p, estado='activo', activo_hasta=:h WHERE id=:i", p=int(plan_id or 0), h=nuevo, i=int(cuenta_id))
    return True


def _tt_aplicar_tx(ref, status, tx_id, cents):
    """Aplica el resultado de una transacción Wompi a un pago TiendaTrack. Devuelve (ok, mensaje, código_http). Es idempotente."""
    _tt_init()
    filas = _c48_q("SELECT id, cuenta_id, plan_id, periodo, monto_centavos, estado FROM tt_pagos WHERE referencia=:r", r=str(ref or ""))
    if not filas:
        return False, "pago no encontrado", 404
    pid, cid, plid, per, monto, est = filas[0]
    st = (status or "").strip().upper()
    ahora_s = _tt_ahora_s()
    if st in _TT_APROBADO:
        if cents is not None:
            try:
                if int(float(cents)) != int(monto):
                    _c48_x("UPDATE tt_pagos SET estado='revisar', id_transaccion=:t, actualizado_en=:a WHERE id=:i AND estado='pendiente'", t=str(tx_id or "")[:80], a=ahora_s, i=pid)
                    return False, "el monto no coincide: queda en revisión", 400
            except Exception:
                return False, "monto inválido", 400
        if est == "aprobado":
            return True, "ya estaba aplicado", 200
        _c48_x("UPDATE tt_pagos SET estado='aprobado', id_transaccion=:t, actualizado_en=:a WHERE id=:i", t=str(tx_id or "")[:80], a=ahora_s, i=pid)
        _tt_activar(cid, plid, per)
        try:
            registrar_auditoria("TiendaTrack pago aprobado", "ref=%s cuenta=%s" % (ref, cid))
        except Exception:
            pass
        return True, "pago aprobado y plan activado", 200
    if est == "pendiente" and st in ("DECLINED", "ERROR", "VOIDED"):
        _c48_x("UPDATE tt_pagos SET estado=:e, id_transaccion=:t, actualizado_en=:a WHERE id=:i",
               e=("anulado" if st == "VOIDED" else "rechazado"), t=str(tx_id or "")[:80], a=ahora_s, i=pid)
    return True, "estado %s" % (st or "sin_estado"), 200


def _tt_webhook_pago(tx, status, ref, tx_id, amount_cents):
    """Lo llama el webhook de Wompi que ya existe (/api/webhooks/wompi) cuando la referencia empieza por TT-."""
    ok, msg, code = _tt_aplicar_tx(ref, status, tx_id, amount_cents)
    return {"ok": ok, "processed": ok, "tiendatrack": msg}, code


# ───────────── TiendaTrack · Google, Wompi y páginas públicas (/tt) ─────────────
def _tt_google_cfg():
    return (os.environ.get("GOOGLE_CLIENT_ID") or "").strip(), (os.environ.get("GOOGLE_CLIENT_SECRET") or "").strip()


def _tt_wompi_cfg():
    pub = (os.environ.get("WOMPI_PUBLIC_KEY") or "").strip()
    if not pub:
        try:
            cand = (_leer_tesoreria().get("pasarela_public_key") or "").strip()
            pub = cand if cand.startswith("pub_") else ""
        except Exception:
            pub = ""
    return pub, (os.environ.get("WOMPI_INTEGRITY_SECRET") or "").strip()


def _tt_wompi_api(pub):
    return "https://sandbox.wompi.co/v1" if pub.startswith("pub_test_") else "https://production.wompi.co/v1"


def _tt_google_redirect():
    return _base_publica().rstrip("/") + "/tt/google/callback"


def _tt_google_url():
    cid, _sec = _tt_google_cfg()
    state = secrets.token_urlsafe(24)
    session["tt_state"] = state
    return "https://accounts.google.com/o/oauth2/v2/auth?" + _tt_up.urlencode({
        "client_id": cid, "redirect_uri": _tt_google_redirect(), "response_type": "code",
        "scope": "openid email profile", "state": state, "prompt": "select_account"})


def _tt_google_intercambiar(code):
    """Cambia el código de Google por los datos de la persona. El id_token llega directo de Google por HTTPS."""
    import time as _t
    cid, sec = _tt_google_cfg()
    cuerpo = _tt_up.urlencode({"code": code, "client_id": cid, "client_secret": sec,
                               "redirect_uri": _tt_google_redirect(), "grant_type": "authorization_code"}).encode()
    req = _tt_ur.Request("https://oauth2.googleapis.com/token", data=cuerpo, method="POST",
                         headers={"Content-Type": "application/x-www-form-urlencoded"})
    with _tt_ur.urlopen(req, timeout=12) as r:
        j = json.loads(r.read().decode("utf-8"))
    idt = j.get("id_token") or ""
    partes = idt.split(".")
    if len(partes) != 3:
        raise ValueError("id_token inválido")
    seg = partes[1] + "=" * (-len(partes[1]) % 4)
    d = json.loads(_tt_b64.urlsafe_b64decode(seg.encode()).decode("utf-8"))
    if d.get("aud") != cid or d.get("iss") not in ("https://accounts.google.com", "accounts.google.com"):
        raise ValueError("id_token de otra aplicación")
    if float(d.get("exp") or 0) < _t.time():
        raise ValueError("id_token vencido")
    if not d.get("email_verified") or not d.get("email") or not d.get("sub"):
        raise ValueError("correo de Google no verificado")
    return {"sub": str(d["sub"])[:64], "email": str(d["email"]).strip().lower()[:160],
            "nombre": str(d.get("name") or "")[:160], "foto": str(d.get("picture") or "")[:400]}


def _tt_csrf():
    t = session.get("tt_csrf")
    if not t:
        t = secrets.token_urlsafe(24)
        session["tt_csrf"] = t
    return t


def _tt_csrf_ok():
    return _tt_hm.compare_digest(str(request.form.get("csrf") or ""), str(session.get("tt_csrf") or "x"))


def _tt_cuenta_sesion():
    cid = session.get("tt_cuenta")
    if not cid:
        return None
    c = _tt_cuenta_por("id", cid)
    if not c:
        session.pop("tt_cuenta", None)
    return c


_TT_CSS = """<style>
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f4f8ff;color:#0b1437;-webkit-font-smoothing:antialiased}
a{color:#0020C2}.tt-top{display:flex;align-items:center;gap:12px;padding:14px 22px;background:#fff;border-bottom:3px solid transparent;border-image:linear-gradient(90deg,#00B8D4,#0020C2,#00FF66) 1}
.tt-brand{font-size:22px;font-weight:800;color:#0020C2;text-decoration:none;letter-spacing:-.02em}.tt-brand b{color:#16a34a}
.tt-top small{color:#64748b;font-size:12px}.tt-top .r{margin-left:auto;display:flex;gap:12px;align-items:center;font-size:13px}
.tt-main{max-width:1040px;margin:0 auto;padding:26px 18px 54px}
.tt-hero{text-align:center;padding:26px 0 6px}.tt-hero h1{font-size:clamp(28px,5vw,46px);margin:0 0 10px;color:#000F66;letter-spacing:-.03em;line-height:1.08}
.tt-hero p{color:#475569;font-size:17px;max-width:640px;margin:0 auto 18px;line-height:1.5}
.tt-pill{display:inline-block;background:#fef3c7;color:#92400e;border-radius:999px;padding:4px 12px;font-size:12px;font-weight:700;margin-bottom:12px}
.tt-btn{display:inline-block;background:#0020C2;color:#fff;border:0;border-radius:999px;padding:12px 24px;font-size:15px;font-weight:700;cursor:pointer;text-decoration:none}
.tt-btn:hover{background:#0018A8}.tt-btn.g{background:#00FF66;color:#00164d}.tt-btn.o{background:#fff;color:#0020C2;border:1.5px solid #0020C2}.tt-btn.off{background:#cbd5e1;color:#64748b;cursor:not-allowed}
.tt-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px;margin:22px 0}
.tt-card{background:#fff;border:1px solid #dbe7ff;border-radius:18px;padding:20px;box-shadow:0 10px 30px rgba(0,32,194,.07)}
.tt-card h3{margin:0 0 6px;font-size:18px;color:#0020C2}.tt-card p{margin:6px 0;color:#475569;font-size:14px;line-height:1.5}
.tt-precio{font-size:30px;font-weight:800;color:#000F66;margin:8px 0 2px}.tt-precio small{font-size:13px;color:#64748b;font-weight:600}
.tt-card ul{padding-left:18px;margin:10px 0;color:#334155;font-size:13.5px;line-height:1.7}
.tt-h2{font-size:22px;color:#000F66;margin:30px 0 4px;letter-spacing:-.02em}
.tt-box{background:#fff;border:1px solid #dbe7ff;border-radius:18px;padding:22px;max-width:520px;margin:20px auto;box-shadow:0 10px 30px rgba(0,32,194,.07)}
.tt-box label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px;color:#334155}
.tt-box input[type=text],.tt-box input[type=tel],.tt-box select{width:100%;padding:11px 12px;border:1px solid #cbd5e1;border-radius:10px;font-size:15px}
.tt-box input:focus,.tt-box select:focus{outline:none;border-color:#0020C2;box-shadow:0 0 0 3px rgba(0,32,194,.12)}
.tt-chk{display:flex;gap:10px;align-items:flex-start;font-size:13px;color:#475569;margin:14px 0;line-height:1.5}.tt-chk input{margin-top:3px;width:18px;height:18px}
.tt-estado{border-radius:18px;padding:18px 20px;margin:16px 0;font-size:15px;line-height:1.5}
.tt-estado.demo{background:#ecfeff;border:1px solid #67e8f9;color:#155e75}.tt-estado.activo{background:#ecfdf3;border:1px solid #86efac;color:#14532d}
.tt-estado.vencido,.tt-estado.suspendido{background:#fef2f2;border:1px solid #fca5a5;color:#7f1d1d}
.tt-msg{background:#fffbeb;border:1px solid #fcd34d;color:#78350f;border-radius:12px;padding:12px 14px;font-size:14px;margin:14px 0}
table.tt-t{width:100%;border-collapse:collapse;font-size:13.5px}table.tt-t th{text-align:left;color:#0020C2;padding:8px;background:#f4f6ff}table.tt-t td{padding:8px;border-bottom:1px solid #edf0f8}
.tt-foot{text-align:center;color:#64748b;font-size:12.5px;padding:22px 16px;line-height:1.7}
.tt-av{width:30px;height:30px;border-radius:50%;object-fit:cover}
</style>"""


def _tt_html(titulo, cuerpo, cuenta=None, codigo=200):
    derecha = ""
    if cuenta:
        derecha = ("<span class='r'><span>" + _esc(cuenta.get("negocio") or cuenta.get("email") or "") +
                   "</span><a href='/tt/salir'>Salir</a></span>")
    doc = ("<!doctype html><html lang='es'><head><meta charset='utf-8'>"
           "<meta name='viewport' content='width=device-width,initial-scale=1'>"
           "<meta name='robots' content='noindex'><title>" + _esc(titulo) + " · TiendaTrack</title>" + _TT_CSS + "</head><body>"
           "<header class='tt-top'><a class='tt-brand' href='/tt'>Tienda<b>Track</b></a><small>un producto de PROCSIS</small>" + derecha + "</header>"
           "<main class='tt-main'>" + cuerpo + "</main>"
           "<footer class='tt-foot'>TiendaTrack · PROCSIS &middot; Innovaci&oacute;n que gestiona<br>"
           "<a href='/legal'>Términos</a> · <a href='/tratamiento-datos'>Tratamiento de datos</a> · <a href='/privacidad'>Privacidad</a><br>"
           "Por ahora TiendaTrack entrega comprobantes de pago; no emite factura electrónica ante la DIAN.</footer></body></html>")
    from flask import make_response
    r = make_response(doc, codigo)
    r.headers["Cache-Control"] = "no-store"
    return r


def _tt_plan_card(p, csrf, con_pago=False, seg=""):
    items = ["Hasta %d sucursal(es) y %d usuario(s)" % (p["max_sucursales"], p["max_usuarios"]),
             "Hasta %s productos" % "{:,}".format(p["max_productos"]).replace(",", "."),
             "Facturación DIAN: %s" % _esc(p["incluye_dian"] or "No")]
    h = ("<div class='tt-card'><h3>" + _esc(p["nombre"]) + "</h3><p>" + _esc(p["descripcion"] or "") + "</p>"
         "<div class='tt-precio'>" + _tt_cop(p["precio_mensual"]) + " <small>/ mes</small></div>")
    if p["precio_anual"] > 0:
        h += "<p style='margin:0'>o " + _tt_cop(p["precio_anual"]) + " al año</p>"
    h += "<ul>" + "".join("<li>%s</li>" % i for i in items) + "</ul>"
    if con_pago:
        h += ("<form method='POST' action='/tt/pagar'><input type='hidden' name='csrf' value='" + _esc(csrf) + "'>"
              "<input type='hidden' name='plan_id' value='%d'>" % p["id"])
        if p["precio_anual"] > 0:
            h += ("<label style='font-size:13px;display:block;margin:6px 0'><input type='radio' name='periodo' value='mensual' checked> Mensual</label>"
                  "<label style='font-size:13px;display:block;margin:6px 0 10px'><input type='radio' name='periodo' value='anual'> Anual</label>")
        else:
            h += "<input type='hidden' name='periodo' value='mensual'>"
        h += "<button class='tt-btn g' type='submit'>Pagar con Wompi</button></form>"
    return h + "</div>"


@app.route("/tt")
def tt_inicio():
    _tt_init()
    abierto = _tt_cfg("tt_registro_abierto", "0") == "1"
    dias = _tt_int(_tt_cfg("tt_demo_dias", "15"), 15, 1, 90)
    c = _tt_cuenta_sesion()
    planes = _tt_planes(solo_web=True)
    if c:
        cta = "<a class='tt-btn' href='/tt/cuenta'>Ir a mi cuenta</a>"
    elif abierto:
        cta = "<a class='tt-btn g' href='/tt/google/inicio'>Probar %d días gratis con Google</a>" % dias
    else:
        cta = "<span class='tt-btn off'>Próximamente</span>"
    cuerpo = ("<section class='tt-hero'><span class='tt-pill'>Acceso anticipado</span>"
              "<h1>Tu negocio, en orden.</h1>"
              "<p>Ventas, inventario, fiados y caja para tiendas, tiendas de ropa y restaurantes. Empieza con una demo gratis de %d días y paga solo si te sirve.</p>%s</section>" % (dias, cta))
    cuerpo += ("<div class='tt-grid'>"
               "<div class='tt-card'><h3>Tiendas</h3><p>Inventario rápido, fiados con control de deudas y cierre de caja.</p></div>"
               "<div class='tt-card'><h3>Tiendas de ropa</h3><p>Productos por talla y color, apartados y cambios.</p></div>"
               "<div class='tt-card'><h3>Restaurantes</h3><p>Mesas, comandas y propinas voluntarias.</p></div></div>")
    if planes:
        cuerpo += "<h2 class='tt-h2'>Planes</h2><div class='tt-grid'>" + "".join(_tt_plan_card(p, "") for p in planes) + "</div>"
        cuerpo += "<p style='text-align:center;color:#64748b;font-size:13px'>Precios en pesos colombianos. Pagos seguros con Wompi.</p>"
    return _tt_html("Inicio", cuerpo, c)


@app.route("/tt/google/inicio")
def tt_google_inicio():
    ok_rl, _e = _rate_limit("tt-google", 30, 60)
    if not ok_rl:
        return _tt_html("Intenta de nuevo", "<div class='tt-msg'>Demasiados intentos. Espera un momento.</div>", None, 429)
    cid, sec = _tt_google_cfg()
    if not cid or not sec:
        return _tt_html("Acceso con Google", "<div class='tt-box'><h2>Acceso con Google no disponible</h2><p>Estamos terminando de configurarlo. Intenta más tarde.</p></div>", None, 503)
    return redirect(_tt_google_url())


@app.route("/tt/google/callback")
def tt_google_callback():
    ok_rl, _e = _rate_limit("tt-google", 30, 60)
    if not ok_rl:
        return _tt_html("Intenta de nuevo", "<div class='tt-msg'>Demasiados intentos. Espera un momento.</div>", None, 429)
    def fallo(msg):
        return _tt_html("No pudimos entrar", "<div class='tt-box'><h2>No pudimos iniciar sesión</h2><p>%s</p><p><a class='tt-btn' href='/tt'>Volver</a></p></div>" % _esc(msg), None, 400)
    if request.args.get("error"):
        return fallo("Cancelaste el acceso con Google.")
    esperado = session.pop("tt_state", "")
    recibido = request.args.get("state", "")
    if not esperado or not _tt_hm.compare_digest(str(recibido), str(esperado)):
        return fallo("La sesión de acceso venció. Intenta de nuevo.")
    try:
        info = _tt_google_intercambiar(request.args.get("code", ""))
    except Exception as e:
        print("tt google:", repr(e)[:200])
        return fallo("Google no confirmó tu identidad. Intenta de nuevo.")
    c = _tt_cuenta_por("google_sub", info["sub"]) or _tt_cuenta_por("email", info["email"])
    if c:
        if c["google_sub"] and c["google_sub"] != info["sub"]:
            return fallo("Ese correo ya está asociado a otra cuenta de Google.")
        _c48_x("UPDATE tt_cuentas SET google_sub=:s, ultimo_acceso=:a, foto=:f WHERE id=:i", s=info["sub"], a=_tt_ahora_s(), f=info["foto"], i=c["id"])
        session["tt_cuenta"] = c["id"]
        return redirect("/tt/cuenta")
    if _tt_cfg("tt_registro_abierto", "0") != "1":
        return _tt_html("Registro cerrado", "<div class='tt-box'><h2>El registro aún no está abierto</h2><p>Estamos preparando TiendaTrack. Vuelve pronto.</p></div>", None, 403)
    session["tt_pend"] = info
    return redirect("/tt/registro")


@app.route("/tt/registro", methods=["GET", "POST"])
def tt_registro():
    info = session.get("tt_pend")
    if not info:
        return redirect("/tt")
    if _tt_cfg("tt_registro_abierto", "0") != "1":
        return _tt_html("Registro cerrado", "<div class='tt-box'><h2>El registro aún no está abierto</h2></div>", None, 403)
    err = ""
    if request.method == "POST":
        ok_rl, _e = _rate_limit("tt-registro", 10, 60)
        negocio = " ".join((request.form.get("negocio") or "").split())[:160]
        seg = request.form.get("segmento") or "tienda"
        tel = re.sub(r"[^0-9+]", "", request.form.get("telefono") or "")[:20]
        if not ok_rl:
            err = "Demasiados intentos. Espera un momento."
        elif not _tt_csrf_ok():
            err = "La sesión del formulario venció. Intenta de nuevo."
        elif len(negocio) < 2:
            err = "Escribe el nombre de tu negocio."
        elif seg not in [s for s, _n in _TT_SEGMENTOS]:
            err = "Elige el tipo de negocio."
        elif not request.form.get("acepto"):
            err = "Debes aceptar los términos y el tratamiento de datos para continuar."
        else:
            ahora_s = _tt_ahora_s()
            dias = _tt_int(_tt_cfg("tt_demo_dias", "15"), 15, 1, 90)
            try:
                _c48_x("INSERT INTO tt_cuentas (google_sub, email, nombre, foto, negocio, segmento, telefono, plan_id, estado, demo_inicio, demo_fin, "
                       "acepto_terminos, terminos_version, creado_en, ultimo_acceso) VALUES (:s, :e, :n, :f, :ng, :sg, :t, 0, 'demo', :di, :df, :at, :tv, :c, :c)",
                       s=info["sub"], e=info["email"], n=info["nombre"], f=info["foto"], ng=negocio, sg=seg, t=tel, di=ahora_s,
                       df=_tt_sumar_dias(ahora_s, dias), at=ahora_s, tv=_TT_TERMINOS, c=ahora_s)
            except Exception as e:
                print("tt registro:", repr(e)[:200])
            c = _tt_cuenta_por("google_sub", info["sub"]) or _tt_cuenta_por("email", info["email"])
            if c:
                session.pop("tt_pend", None)
                session["tt_cuenta"] = c["id"]
                try:
                    registrar_auditoria("TiendaTrack registro", "cuenta=%s segmento=%s" % (c["id"], seg))
                except Exception:
                    pass
                return redirect("/tt/cuenta")
            err = "No pudimos crear tu cuenta. Intenta de nuevo."
    opciones = "".join("<option value='%s'>%s</option>" % (v, _esc(n)) for v, n in _TT_SEGMENTOS)
    dias = _tt_int(_tt_cfg("tt_demo_dias", "15"), 15, 1, 90)
    cuerpo = ("<div class='tt-box'><h2 style='margin-top:0'>Crea tu cuenta</h2><p style='color:#475569'>Entraste con <b>" + _esc(info["email"]) + "</b>. "
              "Cuéntanos de tu negocio y empieza tu demo de %d días, sin tarjeta.</p>" % dias +
              ("<div class='tt-msg'>" + _esc(err) + "</div>" if err else "") +
              "<form method='POST'><input type='hidden' name='csrf' value='" + _esc(_tt_csrf()) + "'>"
              "<label>Nombre de tu negocio</label><input type='text' name='negocio' maxlength='160' required>"
              "<label>Tipo de negocio</label><select name='segmento'>" + opciones + "</select>"
              "<label>Teléfono o WhatsApp (opcional)</label><input type='tel' name='telefono' maxlength='20'>"
              "<label class='tt-chk'><input type='checkbox' name='acepto' value='1'><span>Acepto los <a href='/legal' target='_blank'>términos</a> y la "
              "<a href='/tratamiento-datos' target='_blank'>política de tratamiento de datos personales</a> de PROCSIS.</span></label>"
              "<button class='tt-btn' type='submit'>Empezar mi demo</button></form></div>")
    return _tt_html("Crear cuenta", cuerpo)


@app.route("/tt/cuenta")
def tt_cuenta():
    c = _tt_cuenta_sesion()
    if not c:
        return redirect("/tt")
    _c48_x("UPDATE tt_cuentas SET ultimo_acceso=:a WHERE id=:i", a=_tt_ahora_s(), i=c["id"])
    est = _tt_estado(c)
    plan = _tt_plan(c["plan_id"]) if c["plan_id"] else None
    if est == "demo":
        n = _tt_dias(c["demo_fin"])
        banner = ("<div class='tt-estado demo'><b>Demo gratis</b> · te quedan <b>%d día%s</b> (vence el %s). Puedes elegir un plan cuando quieras.</div>"
                  % (n, "" if n == 1 else "s", _esc(_tt_fecha(c["demo_fin"]))))
    elif est == "activo":
        banner = ("<div class='tt-estado activo'><b>Plan %s activo</b> hasta el %s (%d días).</div>"
                  % (_esc(plan["nombre"] if plan else ""), _esc(_tt_fecha(c["activo_hasta"])), _tt_dias(c["activo_hasta"])))
    elif est == "suspendido":
        banner = "<div class='tt-estado suspendido'><b>Cuenta suspendida.</b> Comunícate con soporte de PROCSIS.</div>"
    else:
        banner = "<div class='tt-estado vencido'><b>Tu demo terminó.</b> Elige un plan y paga para seguir usando TiendaTrack. Tus datos se conservan.</div>"
    csrf = _tt_csrf()
    planes = [p for p in _tt_planes(solo_web=True) if p["segmento"] in ("todos", c["segmento"])]
    cuerpo = "<h1 style='color:#000F66;margin:6px 0'>Hola, " + _esc((c["nombre"] or c["negocio"]).split(" ")[0] if (c["nombre"] or c["negocio"]) else "bienvenido") + "</h1>"
    cuerpo += "<p style='color:#475569;margin:0'>" + _esc(c["negocio"]) + " · " + _esc(c["email"]) + "</p>" + banner
    if est in ("demo", "activo", "vencido") and planes:
        cuerpo += "<h2 class='tt-h2'>%s</h2><div class='tt-grid'>" % ("Renovar o cambiar de plan" if est == "activo" else "Elige tu plan")
        cuerpo += "".join(_tt_plan_card(p, csrf, con_pago=True) for p in planes) + "</div>"
    elif est != "suspendido":
        cuerpo += "<div class='tt-msg'>Los planes aún no están publicados. Escríbenos y te ayudamos.</div>"
    if est in ("demo", "activo"):
        cuerpo += ("<div class='tt-card' style='margin-top:18px'><h3>Tu aplicación</h3>"
                   "<p>La aplicación TiendaTrack se habilitará aquí cuando esté disponible. Tu cuenta y tu demo ya quedaron registradas.</p></div>")
    pagos = _c48_q("SELECT referencia, periodo, monto_centavos, estado, creado_en, plan_id FROM tt_pagos WHERE cuenta_id=:c ORDER BY id DESC LIMIT 12", c=c["id"])
    if pagos:
        filas = "".join("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (
            _esc(_tt_fecha(p[4])), _esc(p[0]), _esc(p[1]), _tt_cop(int(p[2]) // 100), _esc(p[3])) for p in pagos)
        cuerpo += ("<h2 class='tt-h2'>Mis pagos</h2><div class='tt-card' style='overflow:auto'><table class='tt-t'><tr><th>Fecha</th><th>Referencia</th><th>Periodo</th><th>Valor</th><th>Estado</th></tr>"
                   + filas + "</table><p style='font-size:12px;color:#64748b'>Estos son comprobantes de pago. No son factura de venta ni documento equivalente ante la DIAN.</p></div>")
    return _tt_html("Mi cuenta", cuerpo, c)


@app.route("/tt/pagar", methods=["POST"])
def tt_pagar():
    c = _tt_cuenta_sesion()
    if not c:
        return redirect("/tt")
    def aviso(msg, cod=400):
        return _tt_html("Pago", "<div class='tt-box'><h2>No pudimos iniciar el pago</h2><p>%s</p><p><a class='tt-btn' href='/tt/cuenta'>Volver a mi cuenta</a></p></div>" % _esc(msg), c, cod)
    ok_rl, _e = _rate_limit("tt-pagar", 10, 60)
    if not ok_rl:
        return aviso("Demasiados intentos. Espera un momento.", 429)
    if not _tt_csrf_ok():
        return aviso("La sesión del formulario venció. Intenta de nuevo.")
    if _tt_estado(c) == "suspendido":
        return aviso("La cuenta está suspendida. Comunícate con soporte.", 403)
    plan = _tt_plan(request.form.get("plan_id"))
    if not plan or not plan["activo"] or not plan["visible_web"] or plan["segmento"] not in ("todos", c["segmento"]):
        return aviso("Ese plan no está disponible.")
    periodo = "anual" if (request.form.get("periodo") == "anual" and plan["precio_anual"] > 0) else "mensual"
    precio = plan["precio_anual"] if periodo == "anual" else plan["precio_mensual"]
    if precio <= 0:
        return aviso("Este plan aún no tiene precio configurado.")
    pub, integ = _tt_wompi_cfg()
    if not pub or not integ:
        return aviso("Los pagos en línea aún no están habilitados. Escríbenos y te ayudamos.", 503)
    cents = int(precio) * 100
    ref = "TT-%d-%s" % (c["id"], secrets.token_hex(5).upper())
    ahora_s = _tt_ahora_s()
    _c48_x("INSERT INTO tt_pagos (cuenta_id, plan_id, periodo, referencia, monto_centavos, moneda, estado, pasarela, creado_en, actualizado_en) "
           "VALUES (:c, :p, :pe, :r, :m, 'COP', 'pendiente', 'wompi', :t, :t)", c=c["id"], p=plan["id"], pe=periodo, r=ref, m=cents, t=ahora_s)
    firma = _tt_hl.sha256((ref + str(cents) + "COP" + integ).encode("utf-8")).hexdigest()
    campos = [("public-key", pub), ("currency", "COP"), ("amount-in-cents", str(cents)), ("reference", ref),
              ("signature:integrity", firma), ("redirect-url", _base_publica().rstrip("/") + "/tt/pago/retorno"),
              ("customer-data:email", c["email"]), ("customer-data:full-name", (c["nombre"] or c["negocio"] or "")[:60])]
    url = "https://checkout.wompi.co/p/?" + "&".join("%s=%s" % (_tt_up.quote(k, safe=":-"), _tt_up.quote(v, safe="")) for k, v in campos)
    return redirect(url)


@app.route("/tt/pago/retorno")
def tt_pago_retorno():
    c = _tt_cuenta_sesion()
    tx_id = (request.args.get("id") or "").strip()[:80]
    estado, msg = "", ""
    if re.match(r"^[0-9A-Za-z_\-]+$", tx_id or ""):
        pub, _i = _tt_wompi_cfg()
        try:
            req = _tt_ur.Request(_tt_wompi_api(pub) + "/transactions/" + _tt_up.quote(tx_id, safe=""), headers={"User-Agent": "procsis-tiendatrack"})
            with _tt_ur.urlopen(req, timeout=10) as r:
                d = (json.loads(r.read().decode("utf-8")) or {}).get("data") or {}
            estado = str(d.get("status") or "").upper()
            ref = str(d.get("reference") or "")
            if ref.startswith("TT-"):
                _tt_aplicar_tx(ref, estado, tx_id, d.get("amount_in_cents"))
        except Exception as e:
            print("tt retorno:", repr(e)[:200])
    if estado in _TT_APROBADO:
        titulo, msg = "¡Pago aprobado!", "Tu plan ya quedó activo. Gracias por confiar en TiendaTrack."
    elif estado in ("DECLINED", "ERROR", "VOIDED"):
        titulo, msg = "El pago no fue aprobado", "No se hizo ningún cobro. Puedes intentarlo de nuevo con otro medio de pago."
    else:
        titulo, msg = "Estamos confirmando tu pago", "Puede tardar unos minutos. Cuando se confirme, tu plan se activa solo. Revisa tu cuenta en un momento."
    cuerpo = "<div class='tt-box'><h2 style='margin-top:0'>%s</h2><p>%s</p><p><a class='tt-btn' href='/tt/cuenta'>Ir a mi cuenta</a></p></div>" % (_esc(titulo), _esc(msg))
    return _tt_html("Resultado del pago", cuerpo, c)


@app.route("/tt/salir")
def tt_salir():
    session.pop("tt_cuenta", None)
    session.pop("tt_pend", None)
    return redirect("/tt")


@app.route("/api/v1/tiendatrack/cuenta")
def tt_api_cuenta():
    """Para la aplicación TiendaTrack: dice si una cuenta puede usar el servicio. Requiere la cabecera X-API-Key."""
    ok_rl, _e = _rate_limit("tt-api", 120, 60)
    if not ok_rl:
        return {"ok": False, "error": "demasiadas solicitudes"}, 429
    llave = (os.environ.get("TIENDATRACK_API_KEY") or "").strip()
    dada = (request.headers.get("X-API-Key") or "").strip()
    if not llave or not _tt_hm.compare_digest(llave, dada):
        return {"ok": False, "error": "no autorizado"}, 401
    c = _tt_cuenta_por("email", (request.args.get("email") or "").strip().lower())
    if not c:
        return {"ok": False, "error": "cuenta no encontrada"}, 404
    est = _tt_estado(c)
    plan = _tt_plan(c["plan_id"]) if c["plan_id"] else None
    if not plan:
        activos = _tt_planes(solo_activos=True)
        plan = activos[-1] if activos else None
    fin = c["activo_hasta"] if est == "activo" else c["demo_fin"]
    return {"ok": True, "email": c["email"], "negocio": c["negocio"], "segmento": c["segmento"], "estado": est,
            "acceso": est in ("demo", "activo"), "vence": fin, "dias_restantes": _tt_dias(fin) if est in ("demo", "activo") else 0,
            "plan": ({"slug": plan["slug"], "nombre": plan["nombre"], "max_sucursales": plan["max_sucursales"], "max_usuarios": plan["max_usuarios"],
                      "max_productos": plan["max_productos"], "incluye_dian": plan["incluye_dian"]} if plan else None)}, 200


# ───────────── TiendaTrack · módulo de Gerencia (/gerencia/tiendatrack) ─────────────
def _tt_guard():
    if not requiere_login():
        return None, redirect("/login")
    rol = rol_actual()
    if rol not in _DCX_GERENCIA:
        return None, acceso_denegado("Solo Gerencia administra TiendaTrack.")
    try:
        from flask import g as _g_tt
        _g_tt._staff_wrap = True
    except Exception:
        pass
    _tt_init()
    return rol, None


def _tt_pag(titulo, cuerpo):
    tabs = (("/gerencia/tiendatrack", "Panel"), ("/gerencia/tiendatrack/planes", "Planes"), ("/gerencia/tiendatrack/negocios", "Negocios y demos"),
            ("/gerencia/tiendatrack/pagos", "Pagos"), ("/gerencia/tiendatrack/pasarela", "Pasarela y Google"))
    barra = "<div style='display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px'>" + "".join(
        "<a class='cb-btn s %s' href='%s'>%s</a>" % ("" if request.path == h else "o", h, n) for h, n in tabs) + "</div>"
    estilo = ("<style>.tt-k{font-size:28px;font-weight:800;color:#000F66;line-height:1.1}.tt-kl{font-size:12px;color:#64748b;margin-top:2px}"
              ".cb input[type=tel]{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:13px}</style>")
    return _cb_pagina(titulo, estilo + "<h1>🛍️ TiendaTrack</h1><p class='sub'>Planes, negocios con demo, pagos y conexiones. Producto de comercio para tiendas, ropa y restaurantes.</p>" + barra + cuerpo)


def _tt_slug(txt):
    import unicodedata
    s = unicodedata.normalize("NFKD", txt or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")[:40] or "plan"


def _tt_aud(accion, detalle):
    try:
        registrar_auditoria("TiendaTrack · " + accion, "%s: %s" % (session.get("usuario") or "", detalle))
    except Exception:
        pass


def _tt_tag(est):
    clase = {"activo": "on", "demo": "wait"}.get(est, "off")
    return "<span class='cb-tag %s'>%s</span>" % (clase, _esc(est))


@app.route("/gerencia/tiendatrack")
def tt_admin_panel():
    rol, r = _tt_guard()
    if r is not None:
        return r
    cuentas = _c48_q("SELECT id, estado, demo_fin, activo_hasta FROM tt_cuentas")
    cont = {"demo": 0, "activo": 0, "vencido": 0, "suspendido": 0, "por_vencer": 0}
    for cid, est, dfin, ahasta in cuentas:
        e = _tt_estado({"estado": est, "demo_fin": dfin, "activo_hasta": ahasta})
        cont[e] += 1
        if e == "demo" and _tt_dias(dfin) <= 3:
            cont["por_vencer"] += 1
    mes = ahora().strftime("%Y-%m") + "-01"
    ing = _c48_q("SELECT COALESCE(SUM(monto_centavos),0) FROM tt_pagos WHERE estado='aprobado' AND creado_en>=:m", m=mes)
    ingresos = int(ing[0][0]) // 100 if ing else 0
    pend = _c48_q("SELECT COUNT(*) FROM tt_pagos WHERE estado='pendiente'")
    rev = _c48_q("SELECT COUNT(*) FROM tt_pagos WHERE estado='revisar'")
    def k(valor, etiqueta):
        return "<div class='cb-card'><div class='tt-k'>%s</div><div class='tt-kl'>%s</div></div>" % (valor, etiqueta)
    tiles = ("<div class='cb-grid' style='grid-template-columns:repeat(auto-fit,minmax(170px,1fr))'>" +
             k(len(cuentas), "Negocios registrados") + k(cont["demo"], "En demo") + k(cont["por_vencer"], "Demos que vencen en 3 días") +
             k(cont["activo"], "Con plan activo") + k(cont["vencido"], "Vencidos (deben pagar)") + k(_tt_cop(ingresos), "Cobrado este mes") +
             k(int(pend[0][0]) if pend else 0, "Pagos pendientes") + k(int(rev[0][0]) if rev else 0, "Pagos por revisar") + "</div>")
    gc, gs = _tt_google_cfg()
    wp, wi = _tt_wompi_cfg()
    def ok(v, txt_ok="Listo", txt_no="Falta"):
        return "<span class='cb-tag %s'>%s</span>" % ("on" if v else "wait", txt_ok if v else txt_no)
    integ = ("<div class='cb-card'><h2>Estado de conexiones</h2><table><tr><td>Registro abierto al público</td><td>%s</td></tr>"
             "<tr><td>Acceso con Google</td><td>%s</td></tr><tr><td>Pagos con Wompi</td><td>%s</td></tr>"
             "<tr><td>Llave de la API para la app</td><td>%s</td></tr></table>"
             "<p class='cb-note'>Se configuran en <a href='/gerencia/tiendatrack/pasarela'>Pasarela y Google</a>.</p></div>" % (
                 ok(_tt_cfg("tt_registro_abierto", "0") == "1", "Abierto", "Cerrado"), ok(bool(gc and gs)), ok(bool(wp and wi)),
                 ok(bool((os.environ.get("TIENDATRACK_API_KEY") or "").strip()))))
    ult = _c48_q("SELECT negocio, email, segmento, creado_en FROM tt_cuentas ORDER BY id DESC LIMIT 8")
    filas = "".join("<tr><td>%s<br><span class='cb-note'>%s</span></td><td>%s</td><td>%s</td></tr>" % (
        _esc(u[0]), _esc(u[1]), _esc(u[2]), _esc(_tt_fecha(u[3]))) for u in ult) or "<tr><td colspan='3'>Aún no hay negocios registrados.</td></tr>"
    recientes = "<div class='cb-card cb-scroll'><h2>Últimos registros</h2><table><tr><th>Negocio</th><th>Tipo</th><th>Fecha</th></tr>%s</table></div>" % filas
    return _tt_pag("TiendaTrack", tiles + "<div class='cb-grid'>" + integ + recientes + "</div>")


@app.route("/gerencia/tiendatrack/planes", methods=["GET", "POST"])
def tt_admin_planes():
    rol, r = _tt_guard()
    if r is not None:
        return r
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/gerencia/tiendatrack/planes")
        f = request.form
        accion = f.get("accion")
        ahora_s = _tt_ahora_s()
        try:
            if accion in ("activo", "web"):
                col = "activo" if accion == "activo" else "visible_web"
                _c48_x("UPDATE tt_planes SET %s = CASE WHEN %s=1 THEN 0 ELSE 1 END, actualizado_en=:t WHERE id=:i" % (col, col), t=ahora_s, i=_tt_int(f.get("pid")))
                _tt_aud("plan", "%s plan %s" % (accion, f.get("pid")))
                session["ti_msg"] = ("ok", "Plan actualizado.")
            elif accion == "guardar":
                nombre = " ".join((f.get("nombre") or "").split())[:80]
                if len(nombre) < 2:
                    raise ValueError("Escribe el nombre del plan.")
                seg = f.get("segmento") if f.get("segmento") in [s for s, _n in _TT_SEG_PLAN] else "todos"
                dian = f.get("incluye_dian") if f.get("incluye_dian") in ("No", "Próximamente", "Sí") else "No"
                datos = dict(n=nombre, d=" ".join((f.get("descripcion") or "").split())[:300], sg=seg,
                             pm=_tt_int(f.get("pm"), 0, 0, 50000000), pa=_tt_int(f.get("pa"), 0, 0, 500000000),
                             su=_tt_int(f.get("su"), 1, 1, 1000), us=_tt_int(f.get("us"), 1, 1, 10000), pr=_tt_int(f.get("pr"), 100, 1, 10000000),
                             dn=dian, o=_tt_int(f.get("orden"), 0, 0, 999), a=1 if f.get("activo") else 0, w=1 if f.get("web") else 0, t=ahora_s)
                pid = _tt_int(f.get("pid"))
                if pid:
                    datos["i"] = pid
                    _c48_x("UPDATE tt_planes SET nombre=:n, descripcion=:d, segmento=:sg, precio_mensual=:pm, precio_anual=:pa, max_sucursales=:su, max_usuarios=:us, "
                           "max_productos=:pr, incluye_dian=:dn, orden=:o, activo=:a, visible_web=:w, actualizado_en=:t WHERE id=:i", **datos)
                else:
                    base = _tt_slug(nombre)
                    slug, n = base, 2
                    while _c48_q("SELECT 1 FROM tt_planes WHERE slug=:s", s=slug):
                        slug = "%s-%d" % (base[:36], n)
                        n += 1
                    datos["s"] = slug
                    _c48_x("INSERT INTO tt_planes (slug, nombre, descripcion, segmento, precio_mensual, precio_anual, dias_demo, max_sucursales, max_usuarios, max_productos, "
                           "incluye_dian, activo, visible_web, orden, creado_en, actualizado_en) VALUES (:s, :n, :d, :sg, :pm, :pa, 15, :su, :us, :pr, :dn, :a, :w, :o, :t, :t)", **datos)
                _tt_aud("plan", "guardó plan %s" % nombre)
                session["ti_msg"] = ("ok", "Plan guardado.")
        except ValueError as e:
            session["ti_msg"] = ("err", str(e))
        except Exception as e:
            print("tt planes:", repr(e)[:200])
            session["ti_msg"] = ("err", "No se pudo guardar. Intenta de nuevo.")
        return redirect("/gerencia/tiendatrack/planes")
    editar = _tt_plan(request.args.get("editar")) if request.args.get("editar") else None
    e = editar or {"id": 0, "nombre": "", "descripcion": "", "segmento": "todos", "precio_mensual": 0, "precio_anual": 0, "max_sucursales": 1,
                   "max_usuarios": 2, "max_productos": 500, "incluye_dian": "No", "activo": 1, "visible_web": 0, "orden": 0}
    opt_seg = "".join("<option value='%s' %s>%s</option>" % (v, "selected" if v == e["segmento"] else "", _esc(n)) for v, n in _TT_SEG_PLAN)
    opt_dian = "".join("<option %s>%s</option>" % ("selected" if v == e["incluye_dian"] else "", v) for v in ("No", "Próximamente", "Sí"))
    form = ("<div class='cb-card'><h2>%s</h2><form method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='guardar'>"
            "<input type='hidden' name='pid' value='%d'><div class='cb-grid'><div><label>Nombre del plan</label><input type='text' name='nombre' maxlength='80' value='%s' required>"
            "<label>Descripción corta</label><input type='text' name='descripcion' maxlength='300' value='%s'><label>Para qué tipo de negocio</label><select name='segmento'>%s</select>"
            "<label>Facturación DIAN incluida</label><select name='incluye_dian'>%s</select></div>"
            "<div><label>Precio mensual (COP)</label><input type='text' inputmode='numeric' name='pm' value='%d'><label>Precio anual (COP, 0 = sin plan anual)</label>"
            "<input type='text' inputmode='numeric' name='pa' value='%d'><label>Máx. sucursales</label><input type='text' inputmode='numeric' name='su' value='%d'>"
            "<label>Máx. usuarios</label><input type='text' inputmode='numeric' name='us' value='%d'></div>"
            "<div><label>Máx. productos</label><input type='text' inputmode='numeric' name='pr' value='%d'><label>Orden en pantalla</label>"
            "<input type='text' inputmode='numeric' name='orden' value='%d'>"
            "<label style='display:flex;gap:8px;align-items:center'><input type='checkbox' name='activo' value='1' %s> Plan activo</label>"
            "<label style='display:flex;gap:8px;align-items:center'><input type='checkbox' name='web' value='1' %s> Visible y comprable en la web (/tt)</label></div></div>"
            "<p style='margin-top:14px'><button class='cb-btn g'>Guardar plan</button> %s</p></form></div>") % (
        "Editar plan" if e["id"] else "Nuevo plan", _esc(csrf), e["id"], _esc(e["nombre"]), _esc(e["descripcion"]), opt_seg, opt_dian,
        e["precio_mensual"], e["precio_anual"], e["max_sucursales"], e["max_usuarios"], e["max_productos"], e["orden"],
        "checked" if e["activo"] else "", "checked" if e["visible_web"] else "",
        "<a class='cb-btn o' href='/gerencia/tiendatrack/planes'>Cancelar</a>" if e["id"] else "")
    filas = ""
    for p in _tt_planes():
        def bt(accion, texto):
            return ("<form class='in' method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='%s'>"
                    "<input type='hidden' name='pid' value='%d'><button class='cb-btn s o'>%s</button></form> " % (_esc(csrf), accion, p["id"], texto))
        filas += ("<tr><td><b>%s</b><br><span class='cb-note'>%s</span></td><td>%s</td><td>%s<br><span class='cb-note'>año: %s</span></td><td>%d / %d / %s</td><td>%s</td><td>%s %s</td>"
                  "<td><a class='cb-btn s' href='/gerencia/tiendatrack/planes?editar=%d'>Editar</a> %s%s</td></tr>") % (
            _esc(p["nombre"]), _esc(p["slug"]), _esc(dict(_TT_SEG_PLAN).get(p["segmento"], p["segmento"])), _tt_cop(p["precio_mensual"]),
            _tt_cop(p["precio_anual"]) if p["precio_anual"] else "—", p["max_sucursales"], p["max_usuarios"], "{:,}".format(p["max_productos"]).replace(",", "."),
            _esc(p["incluye_dian"]), "<span class='cb-tag %s'>%s</span>" % ("on" if p["activo"] else "off", "activo" if p["activo"] else "oculto"),
            "<span class='cb-tag %s'>%s</span>" % ("on" if p["visible_web"] else "wait", "en la web" if p["visible_web"] else "no en la web"),
            p["id"], bt("activo", "Activar/Desactivar"), bt("web", "Mostrar/Ocultar en web"))
    lista = ("<div class='cb-card cb-scroll'><h2>Planes de TiendaTrack</h2><table><tr><th>Plan</th><th>Para</th><th>Precio</th><th>Sucurs. / Usuarios / Productos</th><th>DIAN</th><th>Estado</th><th></th></tr>"
             + (filas or "<tr><td colspan='7'>Aún no hay planes.</td></tr>") + "</table>"
             "<p class='cb-note'>Los planes de ejemplo vienen ocultos de la web. Ajusta nombre, precio y límites, y activa «visible en la web» cuando estén listos. Los precios no cambian pagos ya hechos.</p></div>")
    return _tt_pag("Planes de TiendaTrack", form + lista)


@app.route("/gerencia/tiendatrack/negocios", methods=["GET", "POST"])
def tt_admin_negocios():
    rol, r = _tt_guard()
    if r is not None:
        return r
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/gerencia/tiendatrack/negocios")
        f = request.form
        accion = f.get("accion")
        try:
            ahora_s = _tt_ahora_s()
            if accion == "pago_manual":
                c = _tt_cuenta_por("email", (f.get("email") or "").strip().lower())
                plan = _tt_plan(f.get("plan_id"))
                periodo = "anual" if f.get("periodo") == "anual" else "mensual"
                monto = _tt_int(f.get("monto"), 0, 0, 500000000)
                if not c or not plan:
                    raise ValueError("Revisa el correo del negocio y el plan.")
                ref = "TT-M-%s" % secrets.token_hex(5).upper()
                _c48_x("INSERT INTO tt_pagos (cuenta_id, plan_id, periodo, referencia, monto_centavos, moneda, estado, pasarela, factura_ref, creado_en, actualizado_en) "
                       "VALUES (:c, :p, :pe, :r, :m, 'COP', 'aprobado', 'manual', :fr, :t, :t)", c=c["id"], p=plan["id"], pe=periodo, r=ref, m=monto * 100,
                       fr=" ".join((f.get("factura_ref") or "").split())[:60], t=ahora_s)
                _tt_activar(c["id"], plan["id"], periodo)
                _tt_aud("pago manual", "%s %s %s" % (c["email"], plan["nombre"], periodo))
                session["ti_msg"] = ("ok", "Pago manual registrado y plan activado.")
            else:
                cid = _tt_int(f.get("cuenta_id"))
                c = _tt_cuenta_por("id", cid)
                if not c:
                    raise ValueError("Negocio no encontrado.")
                if accion in ("ext7", "ext15"):
                    dias = 7 if accion == "ext7" else 15
                    base = c["demo_fin"] if (c["demo_fin"] and c["demo_fin"] > ahora_s) else ahora_s
                    _c48_x("UPDATE tt_cuentas SET demo_fin=:d WHERE id=:i", d=_tt_sumar_dias(base, dias), i=cid)
                    msg = "Demo extendida %d días." % dias
                elif accion == "suspender":
                    _c48_x("UPDATE tt_cuentas SET estado='suspendido' WHERE id=:i", i=cid)
                    msg = "Cuenta suspendida."
                elif accion == "reactivar":
                    _c48_x("UPDATE tt_cuentas SET estado='demo' WHERE id=:i", i=cid)
                    msg = "Cuenta reactivada."
                else:
                    raise ValueError("Acción no válida.")
                _tt_aud("negocio", "%s → %s" % (c["email"], accion))
                session["ti_msg"] = ("ok", msg)
        except ValueError as e:
            session["ti_msg"] = ("err", str(e))
        except Exception as e:
            print("tt negocios:", repr(e)[:200])
            session["ti_msg"] = ("err", "No se pudo completar la acción.")
        return redirect("/gerencia/tiendatrack/negocios")
    q = " ".join((request.args.get("q") or "").split()).lower()[:80]
    fest = request.args.get("estado") or ""
    planes = {p["id"]: p for p in _tt_planes()}
    filas, total = "", 0
    for fila in _c48_q("SELECT %s FROM tt_cuentas ORDER BY id DESC LIMIT 500" % ", ".join(_TT_COLS_CUENTA)):
        c = _tt_dict(_TT_COLS_CUENTA, fila)
        est = _tt_estado(c)
        if fest and est != fest:
            continue
        if q and q not in (c["email"] + " " + c["negocio"]).lower():
            continue
        total += 1
        fin = c["activo_hasta"] if est == "activo" else c["demo_fin"]
        pl = planes.get(c["plan_id"])
        acc = ("<form class='in' method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='cuenta_id' value='%d'>"
               "<select name='accion' style='width:auto'><option value='ext7'>+7 días de demo</option><option value='ext15'>+15 días de demo</option>"
               "<option value='suspender'>Suspender</option><option value='reactivar'>Reactivar</option></select> <button class='cb-btn s'>Aplicar</button></form>") % (_esc(csrf), c["id"])
        filas += ("<tr><td><b>%s</b><br><span class='cb-note'>%s</span></td><td>%s</td><td>%s</td><td>%s<br><span class='cb-note'>%d días</span></td><td>%s</td><td>%s</td></tr>") % (
            _esc(c["negocio"]), _esc(c["email"]), _esc(c["segmento"]), _tt_tag(est), _esc(_tt_fecha(fin)),
            _tt_dias(fin) if est in ("demo", "activo") else 0, _esc(pl["nombre"] if pl else "—"), acc)
    opt_est = "".join("<option value='%s' %s>%s</option>" % (v, "selected" if v == fest else "", n) for v, n in (("", "Todos"), ("demo", "En demo"), ("activo", "Activos"), ("vencido", "Vencidos"), ("suspendido", "Suspendidos")))
    filtro = ("<div class='cb-card'><form method='GET' style='display:flex;gap:10px;flex-wrap:wrap;align-items:end'><div><label>Buscar</label>"
              "<input type='text' name='q' value='%s' placeholder='correo o negocio'></div><div><label>Estado</label><select name='estado'>%s</select></div>"
              "<button class='cb-btn'>Filtrar</button></form></div>") % (_esc(q), opt_est)
    lista = ("<div class='cb-card cb-scroll'><h2>Negocios (%d)</h2><table><tr><th>Negocio</th><th>Tipo</th><th>Estado</th><th>Vence</th><th>Plan</th><th>Acción</th></tr>%s</table></div>"
             % (total, filas or "<tr><td colspan='6'>No hay negocios con ese filtro.</td></tr>"))
    opt_plan = "".join("<option value='%d'>%s</option>" % (p["id"], _esc(p["nombre"])) for p in planes.values())
    manual = ("<div class='cb-card'><h2>Registrar pago manual</h2><p>Para transferencias o pagos fuera de la pasarela. Activa el plan y deja el pago registrado.</p>"
              "<form method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='pago_manual'><div class='cb-grid'>"
              "<div><label>Correo del negocio</label><input type='email' name='email' required><label>Plan</label><select name='plan_id'>%s</select></div>"
              "<div><label>Periodo</label><select name='periodo'><option value='mensual'>Mensual (30 días)</option><option value='anual'>Anual (365 días)</option></select>"
              "<label>Valor recibido (COP)</label><input type='text' inputmode='numeric' name='monto'><label>N.º de factura emitida (opcional)</label>"
              "<input type='text' name='factura_ref' maxlength='60'></div></div><p style='margin-top:12px'><button class='cb-btn g'>Registrar y activar</button></p></form></div>") % (_esc(csrf), opt_plan)
    return _tt_pag("Negocios de TiendaTrack", filtro + lista + manual)


@app.route("/gerencia/tiendatrack/pagos", methods=["GET", "POST"])
def tt_admin_pagos():
    rol, r = _tt_guard()
    if r is not None:
        return r
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
        else:
            try:
                _c48_x("UPDATE tt_pagos SET factura_ref=:f, actualizado_en=:t WHERE id=:i", f=" ".join((request.form.get("factura_ref") or "").split())[:60],
                       t=_tt_ahora_s(), i=_tt_int(request.form.get("pid")))
                _tt_aud("pago", "factura_ref del pago %s" % request.form.get("pid"))
                session["ti_msg"] = ("ok", "Dato de factura guardado.")
            except Exception:
                session["ti_msg"] = ("err", "No se pudo guardar.")
        return redirect("/gerencia/tiendatrack/pagos")
    sql = ("SELECT p.id, p.creado_en, c.negocio, c.email, p.plan_id, p.periodo, p.monto_centavos, p.estado, p.pasarela, p.referencia, p.factura_ref "
           "FROM tt_pagos p LEFT JOIN tt_cuentas c ON c.id=p.cuenta_id ORDER BY p.id DESC LIMIT 200")
    planes = {p["id"]: p["nombre"] for p in _tt_planes()}
    filas = ""
    for p in _c48_q(sql):
        form = ("<form class='in' method='POST'><input type='hidden' name='csrf' value='%s'><input type='hidden' name='pid' value='%d'>"
                "<input type='text' name='factura_ref' value='%s' maxlength='60' style='width:130px' placeholder='N.º factura'> <button class='cb-btn s o'>Guardar</button></form>") % (
            _esc(csrf), p[0], _esc(p[10] or ""))
        clase = {"aprobado": "on", "pendiente": "wait", "revisar": "wait"}.get(p[7], "off")
        filas += ("<tr><td>%s</td><td><b>%s</b><br><span class='cb-note'>%s</span></td><td>%s · %s</td><td>%s</td><td><span class='cb-tag %s'>%s</span><br><span class='cb-note'>%s</span></td>"
                  "<td><span class='cb-note'>%s</span></td><td>%s</td></tr>") % (
            _esc(_tt_fecha(p[1])), _esc(p[2] or "—"), _esc(p[3] or ""), _esc(planes.get(p[4], "—")), _esc(p[5]), _tt_cop(int(p[6]) // 100), clase, _esc(p[7]), _esc(p[8]), _esc(p[9]), form)
    cuerpo = ("<div class='cb-card cb-scroll'><h2>Pagos de TiendaTrack</h2><table><tr><th>Fecha</th><th>Negocio</th><th>Plan</th><th>Valor</th><th>Estado</th><th>Referencia</th><th>Factura de PROCSIS</th></tr>"
              + (filas or "<tr><td colspan='7'>Aún no hay pagos.</td></tr>") + "</table>"
              "<p class='cb-note'>«revisar» = el valor que reportó la pasarela no coincide con el del plan; no activa nada hasta que lo revises. "
              "La factura electrónica del cobro a cada negocio la emite Contabilidad; aquí se anota su número.</p></div>")
    return _tt_pag("Pagos de TiendaTrack", cuerpo)


@app.route("/gerencia/tiendatrack/pasarela", methods=["GET", "POST"])
def tt_admin_pasarela():
    rol, r = _tt_guard()
    if r is not None:
        return r
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
        else:
            try:
                _ti_cfg_set("tt_demo_dias", str(_tt_int(request.form.get("dias"), 15, 1, 90)))
                _ti_cfg_set("tt_registro_abierto", "1" if request.form.get("abierto") else "0")
                _tt_aud("configuración", "demo=%s abierto=%s" % (request.form.get("dias"), "sí" if request.form.get("abierto") else "no"))
                session["ti_msg"] = ("ok", "Configuración guardada.")
            except Exception:
                session["ti_msg"] = ("err", "No se pudo guardar.")
        return redirect("/gerencia/tiendatrack/pasarela")
    gc, gs = _tt_google_cfg()
    wp, wi = _tt_wompi_cfg()
    we = bool((os.environ.get("WOMPI_EVENTS_SECRET") or "").strip() or ((_leer_tesoreria() or {}).get("wompi_events_secret") or "").strip())
    api = bool((os.environ.get("TIENDATRACK_API_KEY") or "").strip())
    def fila(nombre, listo, var):
        return "<tr><td>%s</td><td><span class='cb-tag %s'>%s</span></td><td><code>%s</code></td></tr>" % (nombre, "on" if listo else "wait", "Listo" if listo else "Falta", var)
    estado = ("<div class='cb-card'><h2>Qué falta configurar en Railway</h2><table><tr><th>Conexión</th><th>Estado</th><th>Variable</th></tr>" +
              fila("Google · ID de cliente", bool(gc), "GOOGLE_CLIENT_ID") + fila("Google · secreto", bool(gs), "GOOGLE_CLIENT_SECRET") +
              fila("Wompi · llave pública (%s)" % ("pruebas" if wp.startswith("pub_test_") else ("producción" if wp else "sin definir")), bool(wp), "WOMPI_PUBLIC_KEY") +
              fila("Wompi · secreto de integridad", bool(wi), "WOMPI_INTEGRITY_SECRET") + fila("Wompi · secreto de eventos", we, "WOMPI_EVENTS_SECRET") +
              fila("API para la app TiendaTrack", api, "TIENDATRACK_API_KEY") + "</table>"
              "<p class='cb-note'>Las llaves nunca se guardan en la base de datos ni se muestran aquí: viven en las variables de entorno.</p></div>")
    base = _base_publica().rstrip("/")
    guia = ("<div class='cb-card'><h2>Cómo conectarlo</h2><ol class='cb-steps'>"
            "<li><b>Google:</b> en Google Cloud Console crea un «ID de cliente de OAuth» tipo aplicación web y agrega como URI de redirección autorizada: <code>%s/tt/google/callback</code></li>"
            "<li><b>Wompi:</b> usa las mismas llaves de tu comercio. En «Desarrolladores» copia la llave pública y el secreto de integridad. "
            "El webhook de eventos es el que ya tienes: <code>%s/api/webhooks/wompi</code> (los pagos de TiendaTrack se reconocen por la referencia que empieza con TT-).</li>"
            "<li><b>Pruebas:</b> con llaves <code>pub_test_</code> Wompi funciona en modo sandbox, sin cobrar dinero real.</li>"
            "<li><b>API:</b> inventa una llave larga y guárdala en <code>TIENDATRACK_API_KEY</code>; la app TiendaTrack la envía en la cabecera <code>X-API-Key</code> a <code>%s/api/v1/tiendatrack/cuenta?email=…</code></li></ol></div>") % (base, base, base)
    abierto = _tt_cfg("tt_registro_abierto", "0") == "1"
    conf = ("<div class='cb-card'><h2>Registro y demo</h2><form method='POST'><input type='hidden' name='csrf' value='%s'>"
            "<label>Días de demo gratis (1 a 90)</label><input type='text' inputmode='numeric' name='dias' value='%d' style='max-width:120px'>"
            "<label style='display:flex;gap:8px;align-items:center;margin-top:14px'><input type='checkbox' name='abierto' value='1' %s> Registro abierto al público en <a href='/tt' target='_blank'>/tt</a></label>"
            "<p class='cb-note'>Déjalo cerrado hasta que la aplicación y los planes estén listos. Quien ya tiene cuenta puede seguir entrando.</p>"
            "<p><button class='cb-btn g'>Guardar</button></p></form></div>") % (_esc(csrf), _tt_int(_tt_cfg("tt_demo_dias", "15"), 15, 1, 90), "checked" if abierto else "")
    return _tt_pag("Pasarela y Google", estado + "<div class='cb-grid'>" + conf + guia + "</div>")


# ───────────── Selector de producto: Todo PROCSIS · EduTrack · TiendaTrack ─────────────
_TT_SELECTOR_HTML = r"""<style>
#tt-sel{margin:0 0 14px;padding:8px;border-radius:12px;background:rgba(255,255,255,.08)}
#tt-sel .tt-sel-t{font-size:9.5px;letter-spacing:.14em;text-transform:uppercase;color:rgba(255,255,255,.55);margin:0 4px 6px}
#tt-sel a{display:flex;align-items:center;gap:8px;padding:7px 8px;border-radius:8px;color:#fff;text-decoration:none;font-size:12px;font-weight:600;opacity:.78}
#tt-sel a:hover{opacity:1;background:rgba(255,255,255,.1)}
#tt-sel a i{width:8px;height:8px;border-radius:50%;flex:none}
#tt-sel a.on{background:rgba(255,255,255,.18);opacity:1;box-shadow:inset 0 0 0 1px rgba(255,255,255,.28)}
#tt-sel.fix{position:fixed;left:12px;bottom:12px;z-index:9001;background:#0020C2;width:170px}
#tt-banner{background:#ecfeff;border:1px solid #67e8f9;color:#155e75;border-radius:12px;padding:10px 14px;margin:0 0 14px;font-size:13px}
#tt-banner a{color:#0020C2;font-weight:600}
@media print{#tt-sel,#tt-banner{display:none!important}}
</style>
<script>
(function(){
 var ACT="__ACT__",P=[["todo","Todo PROCSIS","#0020C2"],["edutrack","EduTrack","#00B8D4"],["tiendatrack","TiendaTrack","#00FF66"]];
 if(document.getElementById('tt-sel'))return;
 var box=document.createElement('div');box.id='tt-sel';
 var h='<div class="tt-sel-t">Producto</div>';
 for(var i=0;i<P.length;i++){h+='<a href="/producto/'+P[i][0]+'" class="'+(P[i][0]===ACT?'on':'')+'"><i style="background:'+P[i][2]+'"></i>'+P[i][1]+'</a>';}
 box.innerHTML=h;
 var anc=document.querySelector('.sidebar-apple-glass .brand-container-apple');
 if(anc&&anc.parentNode){anc.parentNode.insertBefore(box,anc.nextSibling);}else{box.className='fix';document.body.appendChild(box);}
 if(ACT==="tiendatrack"&&location.pathname.indexOf('/gerencia/tiendatrack')!==0){
  var b=document.createElement('div');b.id='tt-banner';
  b.innerHTML='Vista <b>TiendaTrack</b>: este módulo todavía muestra datos de EduTrack. __LINK__';
  var m=document.querySelector('.bento-workspace-apple');
  if(m){m.insertBefore(b,m.firstChild);}else{document.body.insertBefore(b,document.body.firstChild);}
 }
})();
</script>
"""


@app.route("/producto/<slug>")
def tt_cambiar_producto(slug):
    if not requiere_login() or rol_actual() not in _BIENVENIDA_ROLES:
        return redirect("/login")
    if slug not in [p for p, _n in _TT_PRODUCTOS]:
        slug = "todo"
    session["producto"] = slug
    gerencia = rol_actual() in _DCX_GERENCIA
    if slug == "tiendatrack" and gerencia:
        return redirect("/gerencia/tiendatrack")
    try:
        from urllib.parse import urlparse
        u = urlparse(request.referrer or "")
        if u.netloc and _host_limpio(u.netloc) == _host_limpio(request.host) and u.path and not u.path.startswith("/producto/"):
            if gerencia and u.path.startswith("/gerencia/tiendatrack"):
                return redirect("/gerencia/hq")
            return redirect(u.path + (("?" + u.query) if u.query else ""))
    except Exception:
        pass
    return redirect("/gerencia/hq" if gerencia else "/")


@app.after_request
def _tt_selector(resp):
    """Agrega el selector de producto al menú lateral de las áreas administrativas."""
    try:
        if request.method != "GET" or resp.status_code != 200 or not session.get("usuario"):
            return resp
        if resp.direct_passthrough or not resp.is_sequence or (resp.mimetype or "") != "text/html":
            return resp
        if request.headers.get("Sec-Fetch-Dest", "document") not in ("document", ""):
            return resp
        if request.headers.get("X-Requested-With") or request.args.get("embed"):
            return resp
        p = request.path or ""
        if p.startswith("/static") or p.startswith("/producto/") or "/pdf" in p or rol_actual() not in _BIENVENIDA_ROLES:
            return resp
        html = resp.get_data(as_text=True)
        if "sidebar-apple-glass" not in html or 'id="tt-sel"' in html or "</body>" not in html:
            return resp
        act = session.get("producto") if session.get("producto") in [x for x, _n in _TT_PRODUCTOS] else "todo"
        enlace = "<a href=\"/gerencia/tiendatrack\">Ir al panel TiendaTrack</a>" if rol_actual() in _DCX_GERENCIA else ""
        bloque = _TT_SELECTOR_HTML.replace("__ACT__", act).replace("__LINK__", enlace.replace("\\", ""))
        i = html.rfind("</body>")
        resp.set_data(html[:i] + bloque + html[i:])
    except Exception as _e:
        print("tt selector:", _e)
    return resp


'''

GRUPO_MENU = """    ('TiendaTrack', [
        ('/gerencia/tiendatrack', '🛍️ Panel TiendaTrack'),
        ('/gerencia/tiendatrack/planes', 'Planes y precios'),
        ('/gerencia/tiendatrack/negocios', 'Negocios y demos'),
        ('/gerencia/tiendatrack/pagos', 'Pagos de TiendaTrack'),
        ('/gerencia/tiendatrack/pasarela', 'Pasarela y Google'),
    ]),
"""


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else "app.py"
    if not os.path.isfile(ruta):
        sys.exit("No encuentro %s. Ejecuta el script dentro de la carpeta del proyecto o pasa la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as f:
        txt = f.read()
    if MARCA in txt:
        print("Ya estaba aplicado. No se cambió nada.")
        return
    eol = "\r\n" if "\r\n" in txt else "\n"

    def n(s):
        return s.replace("\r\n", "\n").replace("\n", eol)

    cambios = [
        # 1) el aislamiento por rol no debe sacar al staff de /producto ni al navegador de /tt
        ('    if path.startswith("/sitio/apariencia"):',
         '    if path == "/tt" or path.startswith(("/tt/", "/producto/")):\n'
         '        return None  # [procsis-tiendatrack-fase1] zona pública y selector de producto validan su propio acceso\n'
         '    if path.startswith("/sitio/apariencia"):'),
        # 2) /tt es zona pública (decoración navideña pública, modo mantenimiento)
        ('        or p.startswith("/pqr/")', '        or p.startswith("/pqr/")\n        or p == "/tt" or p.startswith("/tt/")'),
        # 3) menú lateral de Gerencia: grupo TiendaTrack
        ("    ('Comunicaciones y CRTP', [", GRUPO_MENU.rstrip("\n") + "\n    ('Comunicaciones y CRTP', ["),
        # 4) webhook de Wompi: pagos de TiendaTrack
        ('    amount_cents = tx.get("amount_in_cents") or tx.get("amount") or payload.get("amount_in_cents")',
         '    amount_cents = tx.get("amount_in_cents") or tx.get("amount") or payload.get("amount_in_cents")\n'
         '    if ref.startswith("TT-"):  # [procsis-tiendatrack-fase1] pagos de suscripciones TiendaTrack\n'
         '        return _tt_webhook_pago(tx, status, ref, tx_id, amount_cents)'),
    ]
    nuevo = txt
    for viejo, repl in cambios:
        v, r = n(viejo), n(repl)
        c = nuevo.count(v)
        if c != 1:
            sys.exit("No apliqué nada: esperaba 1 coincidencia y hallé %d para:\n  %s\n"
                     "Probablemente app.py cambió desde que armé este script; avísame y lo ajusto." % (c, viejo[:110]))
        nuevo = nuevo.replace(v, r)

    ancla = 'if __name__ == "__main__":'
    i = nuevo.rfind(ancla)
    if i < 0:
        sys.exit('No encontré el bloque final «if __name__ == "__main__":». No apliqué nada.')
    nuevo = nuevo[:i] + n(BLOQUE) + nuevo[i:]

    try:
        compile(nuevo, ruta, "exec")
    except SyntaxError as e:
        sys.exit("El resultado tenía un error de sintaxis (%s). No toqué tu app.py." % e)

    shutil.copyfile(ruta, ruta + ".bak_tiendatrack")
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        f.write(nuevo)
    print("Listo. Copia de seguridad: %s.bak_tiendatrack" % ruta)
    print("Reinicia la app. Entra a Gerencia → TiendaTrack. Falta configurar en Railway: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET,")
    print("WOMPI_PUBLIC_KEY, WOMPI_INTEGRITY_SECRET y TIENDATRACK_API_KEY (la pantalla «Pasarela y Google» te dice cuáles faltan).")


if __name__ == "__main__":
    main()

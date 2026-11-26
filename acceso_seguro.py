# -*- coding: utf-8 -*-
"""
acceso_seguro.py  ·  Parche v2 de acceso seguro para PROCSIS / EduTrack
=========================================================================
Se instala con UNA sola línea desde app.py (la agrega aplicar_parche.py):

    import acceso_seguro; acceso_seguro.instalar(globals())

No edita ninguna función existente: sobrescribe vistas y textos en memoria al arrancar.
Si algo falla, la app sigue funcionando como antes (cada bloque se instala por separado).

QUÉ HACE
  1. Recuperar contraseña SOLO con correo (sin celular):
       · código de 6 casillas estilo Tigo, vence en 5 min, un solo uso, 5 intentos,
         reenvío cada 60 s, máx. 4 códigos por correo/hora y 6 por IP/10 min.
       · mensaje idéntico exista o no el correo; las cuentas se listan SOLO después de verificar el código.
  2. MFA escalonado por nivel de riesgo (guardia central, cubre TODOS los logins):
       T1 Gerencia/Superadmin/Soporte/Desarrollo (PROCSIS) -> en cada inicio de sesión, sin recordar equipo, 30 min de inactividad.
       T2 Rectoría/Coordinación                            -> cada 24 h, 2 h de inactividad.
       T3 Secretaría y administrativos / cualquier cuenta con MFA activo -> cada 7 días en el mismo equipo; si cambia navegador o red, se cierra la sesión.
       Equipo o red nueva -> MFA inmediato. Segundo factor: app TOTP si la cuenta la tiene; si no, PIN al correo.
  3. Step-up: acciones críticas (SIMAT, respaldos, exportes, actas, notas, cambio de claves) piden MFA otra vez si el último fue hace > 10 min.
  4. Aislamiento por rol/ámbito: la sesión queda marcada "corp" o "col:<id>"; un usuario de colegio no entra a rutas corporativas.

INTERRUPTORES (variables de entorno, sin tocar código):
  MFA_POLITICA_MODO = forzar (defecto) | off      -> apaga la guardia MFA completa
  RBAC_MODO         = bloquear (defecto) | auditar | off
  MFA_T3_CIERRA_RED = 1 (defecto) | 0            -> T3: cerrar sesión si cambia la red
"""
import hashlib
import hmac
import ipaddress
import json
import os
import re
import secrets
import time

from flask import request, session, redirect, jsonify, make_response, abort

PIN_TTL = 300            # 5 minutos, de un solo uso
PIN_MAX_INTENTOS = 5
REENVIO_SEG = 60
MAX_ENVIOS_HORA = 5
STEPUP_TTL = 600         # 10 min
COOKIE_DISP = "pm_dev"

ROLES_T1 = {"Gerente", "Gerencia", "Superadmin", "Administrador", "Soporte", "Desarrollador", "Developer"}  # solo con institucion_id NULL (corporativo)
ROLES_T2 = {"Rectoría", "Rector", "Rectora", "Coordinación", "Coordinador", "Coordinadora", "Administrador"}  # colegio
ROLES_T3 = {"Secretaría", "Secretario", "Secretaria"}
ROLES_STAFF_T3 = {"Comercial", "Ventas", "Supervisor de Ventas", "Supervisor", "Cobranza", "Contabilidad", "Soluciones", "Proveedor"}

TIERS = {
    "T1": {"nombre": "Gerencia / Superadmin", "cada": 0, "recordar": False, "inactividad": 30 * 60, "cierra_red": False},
    "T2": {"nombre": "Rectoría y Coordinación", "cada": 24 * 3600, "recordar": True, "inactividad": 2 * 3600, "cierra_red": False},
    "T3": {"nombre": "Secretaría y administrativos", "cada": 7 * 86400, "recordar": True, "inactividad": 2 * 3600, "cierra_red": True},
}

# Rutas críticas que piden MFA reciente (regex, métodos o None = todos). Edita esta lista libremente.
RUTAS_CRITICAS = [
    (r"^/secretaria/simat/exportar", None), (r"^/exportar_simat", None),
    (r"^/secretaria/acta-grado/generar", None),
    (r"^/backup", None), (r"^/backup_datos", None),
    (r"^/seguridad/respaldos", ("POST",)), (r"^/seguridad/api/respaldos", None),
    (r"^/exportar_estudiantes", None), (r"^/importar_estudiantes", ("POST",)),
    (r"^/notas/descargar", None), (r"^/notas/planilla/(exportar|importar)", None),
    (r"^/reset_usuario/", ("POST",)), (r"^/cambiar_password", ("POST",)),
    (r"^/gerencia/contabilidad/export/", None), (r"^/cobranza/contabilidad/export/", None),
    (r"^/contabilidad/libro/export/", None), (r"^/gerencia/convocatoria48/export", None),
    (r"^/gerencia/reconocimientos/export", None), (r"^/exportar_cuenta", None),
    (r"^/seguridad/bitacora/exportar", None),
    (r"^/gerencia/requerimientos-autoridades/\d+/acta", None), (r"^/gerencia/donaciones/kit/\d+/acta", None),
]
_CRIT = [(re.compile(p), m) for p, m in RUTAS_CRITICAS]

# Prefijos corporativos: un usuario de colegio nunca debe llegar aquí.
PREFIJOS_CORP = ("/gerencia/", "/dev-console", "/soporte_admin", "/tecnologia/", "/cobranza/", "/ventas/")

EXENTAS = ("/static/", "/acceso/", "/media/", "/favicon", "/robots", "/api/sesion-estado")

_E = {"tablas": False, "diag": False, "Reto": None, "Disp": None, "Totp": None}


# ───────────────────────── utilidades ─────────────────────────
def _h(v):
    return hashlib.sha256(str(v).encode("utf-8")).hexdigest()


def _modo():
    return (os.environ.get("MFA_POLITICA_MODO") or "forzar").strip().lower()


def _rbac_modo():
    return (os.environ.get("RBAC_MODO") or "bloquear").strip().lower()


def _ip_pref(ip):
    try:
        a = ipaddress.ip_address((ip or "").strip())
        return str(ipaddress.ip_network("%s/%d" % (a, 24 if a.version == 4 else 64), strict=False))
    except Exception:
        return (ip or "")[:60]


def _seguro(dest):
    d = (dest or "").strip()
    if d.startswith("/") and not d.startswith("//") and "\\" not in d and "\n" not in d and "\r" not in d:
        return d
    return ""


def _mask(c):
    try:
        a, d = c.split("@", 1)
        return a[:2] + "***@" + d
    except Exception:
        return "***"


def _exento(p):
    return p == "/mfa-politica" or p == "/mfa" or p.startswith("/mfa/") or p == "/logout" or p.startswith(EXENTAS)


def _es_prefijo_corp(p):
    for pre in PREFIJOS_CORP:
        base = pre.rstrip("/")
        if p == base or p.startswith(base + "/"):
            return True
    return False


def _es_critica(p, metodo):
    for rx, ms in _CRIT:
        if rx.search(p) and (ms is None or metodo in ms):
            return True
    return False


def _politica(u):
    rol = (u.rol or "").strip()
    corp = u.institucion_id is None
    mfa_on = bool(getattr(u, "mfa_habilitado", False) and (getattr(u, "mfa_secret", "") or "").strip()) or bool(getattr(u, "forzar_mfa", False))
    if corp and rol in ROLES_T1:
        return TIERS["T1"], "T1"
    if (not corp) and rol in ROLES_T2:
        return TIERS["T2"], "T2"
    if rol in ROLES_T3 or (corp and rol in ROLES_STAFF_T3) or (corp and rol in _E.get("areas", ())):
        return TIERS["T3"], "T3"
    if mfa_on:
        return TIERS["T3"], "T3"
    return None, ""


# ───────────────────────── instalación ─────────────────────────
def instalar(G):
    """G = globals() de app.py. Cada bloque va aislado: si uno falla, los demás y la app siguen."""
    _E["areas"] = tuple(G.get("ROLES_AREAS", ()))
    for nombre, f in (("MFA escalonado", _instalar_mfa), ("Recuperación por correo", _instalar_recuperacion)):
        try:
            f(G)
            print("ACCESO SEGURO v2: %s instalado." % nombre, flush=True)
        except Exception as ex:
            print("ACCESO SEGURO v2: NO se instaló «%s» (la app sigue igual): %r" % (nombre, ex), flush=True)


# ═════════════════════ 1) RECUPERACIÓN SOLO CON CORREO ═════════════════════
def _instalar_recuperacion(G):
    app, db, text = G["app"], G["db"], G["text"]
    necesarias = ("_rec2_origen_ok", "_rec2_json", "_rec2_init", "_rec2_limite", "_rec2_cuentas", "_rec2_hmac", "_rec2_flujo",
                  "_rec2_validar_pin", "_rec2_etiqueta", "_rec2_enviar", "_rec2_correo_configurado", "_client_ip", "_REC2_JS", "ahora", "APP_NAME")
    faltan = [n for n in necesarias if n not in G]
    vf = app.view_functions
    faltan += [e for e in ("rec2_pin", "rec2_verificar", "rec2_restablecer") if e not in vf]
    if faltan:
        raise RuntimeError("faltan piezas de rec2: %s" % ", ".join(faltan))

    RX_MAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$")

    # El código vence en 5 min (antes 10) también en la solicitud de usuario nuevo.
    def nuevo_flujo(tipo, correo, estado, pin, datos=""):
        token = secrets.token_urlsafe(24)
        ts = int(time.time())
        db.session.execute(text("INSERT INTO acceso_recuperacion (token, tipo, correo, pin_hash, estado, expira, ip, ua, datos, creado) VALUES (:t,:tp,:c,:p,:e,:x,:i,:u,:d,:n)"),
                           {"t": G["_rec2_hmac"]("tok", token), "tp": tipo, "c": correo, "p": G["_rec2_hmac"]("pin", token, pin) if pin else "", "e": estado,
                            "x": ts + PIN_TTL, "i": G["_client_ip"](), "u": (request.headers.get("User-Agent") or "")[:200], "d": datos, "n": ts})
        db.session.execute(text("DELETE FROM acceso_recuperacion WHERE creado<:d"), {"d": ts - 7 * 86400})
        db.session.commit()
        return token

    def cuerpo_pin(pin, para):
        return ("Tu código de verificación de PROCSIS es: %s\n\nVence en 5 minutos y solo sirve una vez. %s\n\n"
                "Solicitud hecha el %s desde la IP %s.\nSi no fuiste tú, ignora este mensaje: nadie puede entrar a tu cuenta sin este código. "
                "No compartas este código con nadie, ni siquiera con Soporte. Si sospechas algo, cambia tu contraseña y avisa a Soporte.\n\n%s"
                % (pin, para, G["ahora"]().strftime("%Y-%m-%d %H:%M"), G["_client_ip"](), G["APP_NAME"]))

    G["_rec2_nuevo_flujo"] = nuevo_flujo
    G["_rec2_cuerpo_pin"] = cuerpo_pin
    J = G["_rec2_json"]

    def pin_v2():
        if not G["_rec2_origen_ok"]() or not request.is_json:
            return J({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        correo = str((request.get_json(silent=True) or {}).get("correo") or "").strip().lower()[:160]
        if not RX_MAIL.match(correo):
            return J({"ok": False, "mensaje": "Escribe un correo válido."}, 400)
        if not G["_rec2_correo_configurado"]():
            return J({"ok": False, "mensaje": "El envío de correos no está disponible en este momento. Contacta a Soporte."}, 503)
        G["_rec2_init"]()
        if not G["_rec2_limite"](correo, G["_client_ip"]()):
            return J({"ok": False, "mensaje": "Hiciste varias solicitudes seguidas. Espera unos minutos e inténtalo de nuevo."}, 429)
        try:  # enfriamiento de 60 s por correo (aplica igual exista o no la cuenta: no revela nada)
            reciente = db.session.execute(text("SELECT COUNT(*) FROM acceso_recuperacion WHERE correo=:c AND tipo='CLAVE' AND creado>:d"),
                                          {"c": correo, "d": int(time.time()) - REENVIO_SEG + 2}).scalar() or 0
            if int(reciente) > 0:
                return J({"ok": False, "mensaje": "Espera un minuto antes de pedir otro código."}, 429)
        except Exception:
            try:
                db.session.rollback()
            except Exception:
                pass
        cuentas = G["_rec2_cuentas"](correo)  # solo cuentas ACTIVAS con ese correo
        pin = "%06d" % secrets.randbelow(1000000)
        token = nuevo_flujo("CLAVE", correo, "PIN" if cuentas else "SIN_CUENTA", pin if cuentas else "")
        if cuentas:
            G["_rec2_enviar"](correo, "Código para recuperar tu acceso - PROCSIS", cuerpo_pin(pin, "Úsalo para recuperar el acceso a tu cuenta."))
            try:
                G["registrar_auditoria"]("Recuperación de acceso: código enviado", correo)
            except Exception:
                pass
        return J({"ok": True, "token": token, "mensaje": "Si el correo está registrado, te enviamos un código de 6 dígitos. Vence en 5 minutos."})

    def verificar_v2():
        if not G["_rec2_origen_ok"]() or not request.is_json:
            return J({"ok": False, "mensaje": "Solicitud no válida."}, 400)
        d = request.get_json(silent=True) or {}
        token, pin = str(d.get("token") or ""), str(d.get("pin") or "").strip()
        G["_rec2_init"]()
        falla = J({"ok": False, "mensaje": "El código no coincide o venció. Revisa e inténtalo de nuevo, o solicita uno nuevo."}, 400)
        fl = G["_rec2_flujo"](token, "CLAVE")
        if not fl:
            return falla
        if not pin.isdigit() or len(pin) != 6:
            return J({"ok": False, "mensaje": "Escribe el código de 6 dígitos."}, 400)
        if not G["_rec2_validar_pin"](fl, token, pin):  # cuenta el intento, bloquea a los 5, vence a los 5 min
            return falla
        cuentas = G["_rec2_cuentas"](fl[1])
        if not cuentas:
            return falla
        lista, guardar = [], []
        for u in cuentas:  # solo cuentas registradas y activas, y solo DESPUÉS de verificar el código
            guardar.append({"id": u.id, "puede": True})
            lista.append({"i": len(lista), "usuario": u.usuario or "", "nombre": (getattr(u, "nombre_completo", "") or "").strip(),
                          "area": G["_rec2_etiqueta"](u), "puede": True, "motivo": ""})
        db.session.execute(text("UPDATE acceso_recuperacion SET estado='VERIFICADO', expira=:x, cuentas=:c, tel_hash='' WHERE id=:i"),
                           {"x": int(time.time()) + 900, "c": json.dumps(guardar), "i": fl[0]})
        db.session.commit()
        return J({"ok": True, "cuentas": lista})

    orig_rest = vf["rec2_restablecer"]

    def restablecer_v2():
        ctx = None
        try:
            d = request.get_json(silent=True) or {}
            fl = G["_rec2_flujo"](str(d.get("token") or ""), "CLAVE")
            idx = int(d.get("indice"))
            if fl and fl[4] == "VERIFICADO":
                el = json.loads(fl[6] or "[]")
                if 0 <= idx < len(el):
                    ctx = (el[idx]["id"], fl[1])
        except Exception:
            ctx = None
        resp = orig_rest()
        try:
            r = resp[0] if isinstance(resp, tuple) else resp
            ok = ctx and r.status_code == 200 and (r.get_json(silent=True) or {}).get("ok")
            if ok:
                u = G["Usuario"].query.get(int(ctx[0]))
                revocar_dispositivos(u.id)  # una clave nueva invalida los equipos "recordados"
                rol = (u.rol or "").strip()
                if u.institucion_id is None and rol in ROLES_T1:
                    G["_rec2_enviar"](G["SOPORTE_EMAIL"], "ALERTA: cuenta crítica recuperada solo con correo",
                                      "La cuenta «%s» (%s) cambió su contraseña por autoservicio con código de correo el %s desde la IP %s.\n"
                                      "Si no fue esperado, desactiva la cuenta y revisa la auditoría." % (u.usuario, rol, G["ahora"]().strftime("%Y-%m-%d %H:%M"), G["_client_ip"]()))
        except Exception as ex:
            print("acceso_seguro post-restablecer:", repr(ex)[:140])
        return resp

    vf["rec2_pin"], vf["rec2_verificar"], vf["rec2_restablecer"] = pin_v2, verificar_v2, restablecer_v2
    G["_REC2_JS"] = _REC2_JS_V2


# ═════════════════════ 2) MFA ESCALONADO + STEP-UP + RBAC ═════════════════════
def revocar_dispositivos(uid):
    try:
        D = _E["Disp"]
        if D is not None:
            D.query.filter_by(usuario_id=uid, revocado=False).update({"revocado": True})
            D.db_session_commit()
    except Exception:
        try:
            _E["Disp"].db_rollback()
        except Exception:
            pass


def _instalar_mfa(G):
    app, db, Usuario = G["app"], G["db"], G["Usuario"]

    class MfaReto(db.Model):
        __tablename__ = "acceso_mfa_reto"
        __table_args__ = {"extend_existing": True}
        id = db.Column(db.Integer, primary_key=True)
        usuario_id = db.Column(db.Integer, index=True, nullable=False)
        pin_hash = db.Column(db.String(64), default="")
        creado = db.Column(db.BigInteger, default=0)
        expira = db.Column(db.BigInteger, default=0)
        intentos = db.Column(db.Integer, default=0)
        estado = db.Column(db.String(8), default="PEND")  # PEND | USADO | BLOQ | ANUL
        ip = db.Column(db.String(64), default="")

    class MfaDispositivo(db.Model):
        __tablename__ = "acceso_mfa_dispositivo"
        __table_args__ = {"extend_existing": True}
        id = db.Column(db.Integer, primary_key=True)
        usuario_id = db.Column(db.Integer, index=True, nullable=False)
        dev_hash = db.Column(db.String(64), index=True, default="")
        ip_pref = db.Column(db.String(64), default="")
        ua_hash = db.Column(db.String(64), default="")
        ok_ts = db.Column(db.BigInteger, default=0)
        revocado = db.Column(db.Boolean, default=False)

        @staticmethod
        def db_session_commit():
            db.session.commit()

        @staticmethod
        def db_rollback():
            db.session.rollback()

    class MfaTotp(db.Model):
        __tablename__ = "acceso_mfa_totp"
        __table_args__ = {"extend_existing": True}
        id = db.Column(db.Integer, primary_key=True)
        usuario_id = db.Column(db.Integer, unique=True, nullable=False)
        ultimo_hash = db.Column(db.String(64), default="")
        ultimo_ts = db.Column(db.BigInteger, default=0)
        fallos = db.Column(db.Integer, default=0)
        bloqueo_hasta = db.Column(db.BigInteger, default=0)

    _E["Reto"], _E["Disp"], _E["Totp"] = MfaReto, MfaDispositivo, MfaTotp

    def tablas():
        if _E["tablas"]:
            return
        try:
            for M in (MfaReto, MfaDispositivo, MfaTotp):
                M.__table__.create(bind=db.engine, checkfirst=True)
            _E["tablas"] = True
        except Exception as ex:
            print("acceso_seguro tablas:", repr(ex)[:140])
            try:
                db.session.rollback()
            except Exception:
                pass

    def audit(a, d=""):
        try:
            G["registrar_auditoria"](a, d)
        except Exception:
            pass

    def ip_actual():
        return G["_client_ip"]()

    def usuario_sesion():
        uid = session.get("uid") or session.get("user_id")  # solo con id: no se adivina por nombre
        try:
            return Usuario.query.get(int(uid)) if uid else None
        except Exception:
            return None

    def tiene_totp(u):
        return bool(getattr(u, "mfa_habilitado", False) and (getattr(u, "mfa_secret", "") or "").strip())

    def vigente(u, pol, ahora_ts):
        ok = session.get("pol_ok")
        if not ok or session.get("pol_uid") != u.id:
            return False
        if pol["cada"] == 0:
            return ahora_ts - ok < 12 * 3600
        return ahora_ts - ok < pol["cada"]

    def dispositivo_valido(u, pol, ahora_ts):
        tok = request.cookies.get(COOKIE_DISP) or ""
        if len(tok) < 30:
            return None
        d = MfaDispositivo.query.filter_by(usuario_id=u.id, dev_hash=_h(tok), revocado=False).first()
        if not d or ahora_ts - (d.ok_ts or 0) >= pol["cada"]:
            return None
        if d.ua_hash != _h(request.headers.get("User-Agent", "")) or d.ip_pref != _ip_pref(ip_actual()):
            return None  # equipo o red nueva => MFA inmediato
        return d

    def marcar_ok(u, ahora_ts, ok_ts=None):
        session["pol_ok"] = ok_ts or ahora_ts
        session["pol_uid"] = u.id
        session["pol_ip"] = _ip_pref(ip_actual())
        session["pol_ua"] = _h(request.headers.get("User-Agent", ""))
        session["pol_last"] = ahora_ts

    def responder_mfa(es_api):
        if es_api:
            r = jsonify({"ok": False, "mfa": "/mfa-politica", "mensaje": "Verificación en dos pasos requerida."})
            r.status_code = 401
            return r
        return redirect("/mfa-politica")

    def es_api():
        return bool(request.is_json or request.headers.get("X-Requested-With") or "application/json" in (request.headers.get("Accept") or ""))

    # ── guardia central ──
    def guardia():
        if _modo() == "off":
            return None
        p = request.path or ""
        if _exento(p):
            return None
        try:
            ruta_pub = G.get("ruta_publica")
            if ruta_pub and ruta_pub():
                return None
        except Exception:
            pass
        if request.method == "POST" and (p == "/login" or p.endswith("-login")):
            for k in ("pol_ok", "pol_uid", "pol_ip", "pol_ua", "pol_last", "pol_ambito", "pol_next", "pol_stepup", "stepup_ts", "pol_otp"):
                session.pop(k, None)  # cada intento de login nuevo arranca sin MFA previo
            return None
        if not session.get("usuario") or session.get("mfa_pendiente"):
            return None
        u = usuario_sesion()
        if not u or getattr(u, "activo", True) is False:
            return None
        tablas()
        if not _E["diag"]:
            _E["diag"] = True
            diagnostico()
        ahora_ts = time.time()
        ambito = "corp" if u.institucion_id is None else "col:%s" % u.institucion_id
        prev = session.get("pol_ambito")
        if prev and prev != ambito:
            audit("Sesión cerrada: ámbito de rol cambió", "%s %s->%s" % (u.usuario, prev, ambito))
            session.clear()
            return redirect("/login")
        session["pol_ambito"] = ambito
        if ambito.startswith("col:") and _es_prefijo_corp(p):
            audit("RBAC: usuario de colegio intentó ruta corporativa", "%s · %s" % (u.usuario, p))
            if _rbac_modo() == "bloquear":
                abort(403)
        pol, nivel = _politica(u)
        if not pol:
            return None
        ua_h, ip_p = _h(request.headers.get("User-Agent", "")), _ip_pref(ip_actual())
        if session.setdefault("pol_ua", ua_h) != ua_h:
            audit("Sesión cerrada: cambió el navegador", u.usuario)
            session.clear()
            return redirect("/login")
        last = session.get("pol_last")
        if last and ahora_ts - last > pol["inactividad"]:
            audit("Sesión cerrada por inactividad (%s)" % nivel, u.usuario)
            session.clear()
            return redirect("/login")
        session["pol_last"] = ahora_ts
        if session.setdefault("pol_ip", ip_p) != ip_p:
            if pol["cierra_red"] and os.environ.get("MFA_T3_CIERRA_RED", "1") != "0":
                audit("Sesión cerrada: cambió la red (%s)" % nivel, u.usuario)
                session.clear()
                return redirect("/login")
            session.pop("pol_ok", None)
            session["pol_ip"] = ip_p
        if not vigente(u, pol, ahora_ts):
            d = dispositivo_valido(u, pol, ahora_ts) if pol["recordar"] else None
            if d:
                marcar_ok(u, ahora_ts, ok_ts=d.ok_ts)
            else:
                if request.method == "GET" and not es_api():
                    n = _seguro(request.full_path.rstrip("?"))
                    if n:
                        session["pol_next"] = n
                return responder_mfa(es_api())
        if _es_critica(p, request.method) and ahora_ts - (session.get("stepup_ts") or 0) > STEPUP_TTL:
            session["pol_stepup"] = 1
            if request.method == "GET" and not es_api():
                n = _seguro(request.full_path.rstrip("?"))
                if n:
                    session["pol_next"] = n
            audit("Step-up MFA solicitado", "%s · %s" % (u.usuario, p))
            return responder_mfa(es_api())
        return None

    app.before_request(guardia)

    def diagnostico():
        try:
            roles = list(ROLES_T1 | ROLES_T2 | ROLES_T3 | ROLES_STAFF_T3 | set(_E.get("areas", ())))
            sin = Usuario.query.filter(Usuario.activo == True, Usuario.rol.in_(roles)).all()  # noqa: E712
            sin = [x.usuario for x in sin if not (x.correo or "").strip() and not tiene_totp(x)]
            if sin:
                print("ACCESO SEGURO v2 · ATENCIÓN: %d cuentas bajo política MFA no tienen correo ni app TOTP (no podrán pasar el MFA): %s"
                      % (len(sin), ", ".join(sin[:25])), flush=True)
            else:
                print("ACCESO SEGURO v2 · diagnóstico: todas las cuentas bajo política MFA tienen correo o TOTP.", flush=True)
        except Exception as ex:
            print("acceso_seguro diagnóstico:", repr(ex)[:140])
            try:
                db.session.rollback()
            except Exception:
                pass

    # ── PIN por correo ──
    def reto_vigente(uid):
        return MfaReto.query.filter(MfaReto.usuario_id == uid, MfaReto.estado == "PEND", MfaReto.expira > int(time.time())).order_by(MfaReto.id.desc()).first()

    def enviar_otp(u, forzar=False):
        ahora_i = int(time.time())
        ult = MfaReto.query.filter_by(usuario_id=u.id).order_by(MfaReto.id.desc()).first()
        if ult and ahora_i - (ult.creado or 0) < REENVIO_SEG:
            return False, "Espera unos segundos antes de pedir otro código."
        if MfaReto.query.filter(MfaReto.usuario_id == u.id, MfaReto.creado > ahora_i - 3600).count() >= MAX_ENVIOS_HORA:
            audit("MFA: demasiados códigos pedidos", u.usuario)
            return False, "Pediste demasiados códigos. Espera un rato e inténtalo de nuevo."
        correo = (u.correo or "").strip()
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$", correo):
            return False, "Tu cuenta no tiene un correo válido registrado. Pide a TI que lo registre."
        if not G["_rec2_correo_configurado"]():
            return False, "El envío de correos no está disponible. Contacta a Soporte."
        MfaReto.query.filter_by(usuario_id=u.id, estado="PEND").update({"estado": "ANUL"})
        pin = "%06d" % secrets.randbelow(1000000)
        r = MfaReto(usuario_id=u.id, creado=ahora_i, expira=ahora_i + PIN_TTL, intentos=0, estado="PEND", ip=ip_actual()[:64])
        db.session.add(r)
        db.session.flush()
        r.pin_hash = G["_rec2_hmac"]("mfa", u.id, r.id, pin)
        db.session.commit()
        ok = G["_rec2_enviar"](correo, "Código de verificación de ingreso - PROCSIS",
                               "Tu código de verificación es: %s\n\nVence en 5 minutos y solo sirve una vez.\nIntento de ingreso el %s desde la IP %s.\n"
                               "Si no fuiste tú, NO compartas este código y avisa a Soporte de inmediato: alguien conoce tu contraseña.\n\n%s"
                               % (pin, G["ahora"]().strftime("%Y-%m-%d %H:%M"), ip_actual(), G["APP_NAME"]))
        if not ok:
            r.estado = "ANUL"
            db.session.commit()
            return False, "No se pudo enviar el correo. Inténtalo de nuevo o contacta a Soporte."
        audit("MFA: código enviado al correo", "%s · %s" % (u.usuario, _mask(correo)))
        return True, ""

    def verificar_otp(u, code):
        r = reto_vigente(u.id)
        if not r:
            return False, "El código venció. Pide uno nuevo."
        if (r.intentos or 0) >= PIN_MAX_INTENTOS:
            r.estado = "BLOQ"
            db.session.commit()
            return False, "Demasiados intentos. Pide un código nuevo."
        esperado = G["_rec2_hmac"]("mfa", u.id, r.id, code)
        if hmac.compare_digest(esperado, r.pin_hash or ""):
            # un solo uso, atómico: si dos peticiones llegan a la vez solo una gana
            n = MfaReto.query.filter_by(id=r.id, estado="PEND").update({"estado": "USADO"})
            db.session.commit()
            return (True, "") if n == 1 else (False, "Ese código ya se usó. Pide uno nuevo.")
        r.intentos = (r.intentos or 0) + 1
        if r.intentos >= PIN_MAX_INTENTOS:
            r.estado = "BLOQ"
            audit("MFA: código bloqueado por intentos", u.usuario)
        db.session.commit()
        return False, "Código incorrecto."

    def verificar_totp(u, code):
        ahora_i = int(time.time())
        st = MfaTotp.query.filter_by(usuario_id=u.id).first()
        if not st:
            st = MfaTotp(usuario_id=u.id)
            db.session.add(st)
            db.session.flush()
        if (st.bloqueo_hasta or 0) > ahora_i:
            return False, "Demasiados intentos. Inténtalo de nuevo en %d min." % (((st.bloqueo_hasta - ahora_i) // 60) + 1)
        if G["_mfa_verificar"](u.mfa_secret, code):
            hh = _h("%s|%s" % (u.id, code))
            if st.ultimo_hash == hh and ahora_i - (st.ultimo_ts or 0) < 120:
                return False, "Ese código ya se usó. Espera el siguiente de tu app."
            st.ultimo_hash, st.ultimo_ts, st.fallos = hh, ahora_i, 0
            db.session.commit()
            return True, ""
        st.fallos = (st.fallos or 0) + 1
        if st.fallos >= PIN_MAX_INTENTOS:
            st.fallos, st.bloqueo_hasta = 0, ahora_i + 600
            audit("MFA TOTP: bloqueo temporal 10 min", u.usuario)
        db.session.commit()
        return False, "Código incorrecto."

    # ── pantalla MFA (estilo Tigo: 6 casillas) ──
    def csrf_token():
        if not session.get("pol_csrf"):
            session["pol_csrf"] = secrets.token_hex(16)
        return session["pol_csrf"]

    def pagina(u, metodo, err, info, resta, stepup):
        esc = G["_esc"]
        correo_txt = ("Te enviamos un código de 6 dígitos a <b>%s</b>." % esc(_mask(u.correo or ""))) if metodo == "correo" else "Abre tu app de autenticación e ingresa el código de 6 dígitos."
        cajas = "".join('<input class="pm-d" inputmode="numeric" pattern="[0-9]*" maxlength="1" aria-label="Dígito %d" %s>' % (i + 1, 'autocomplete="one-time-code"' if i == 0 else 'autocomplete="off"') for i in range(6))
        titulo = "Confirma esta acción" if stepup else "Ingresa el código de verificación"
        reenviar = ""
        if metodo == "correo":
            reenviar = ('<form method="post" class="pm-c"><input type="hidden" name="csrf" value="%s"><input type="hidden" name="accion" value="reenviar">'
                        '<button type="submit" class="pm-link" id="pm-re" style="display:none">Reenviar código</button><span id="pm-cd"></span></form>' % csrf_token())
        css = """<style>
.pm-w{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px;background:#f4f6fb;font-family:inherit}
.pm-card{width:100%;max-width:420px;background:#fff;border-radius:28px;box-shadow:0 20px 60px rgba(10,42,143,.18);overflow:hidden;color:#0f172a}
.pm-hd{display:flex;justify-content:space-between;align-items:center;padding:20px 26px;background:linear-gradient(110deg,#0020C2,#00005e);color:#fff}
.pm-hd a{color:#4da3ff;text-decoration:none;font-size:26px;line-height:1}.pm-hd b{letter-spacing:.08em;font-size:17px}
.pm-bd{padding:26px 26px 30px;text-align:center}.pm-bd h1{margin:0 0 8px;font-size:21px;color:#00004d}
.pm-s{margin:0 0 18px;color:#6b7280;font-size:14px;line-height:1.5}
.pm-cb{display:flex;gap:8px;justify-content:center;margin:6px 0 14px}
.pm-d{flex:1;min-width:0;max-width:52px;height:56px;border:1px solid #d9dde7;border-radius:14px;text-align:center;font-size:22px;font-weight:700;color:#00004d;padding:0;background:#fff;box-sizing:border-box}
.pm-d:focus{outline:2px solid rgba(0,32,194,.35);border-color:#0020C2}
.pm-b{width:100%;border:0;border-radius:999px;padding:15px;font-weight:800;letter-spacing:.04em;color:#fff;background:linear-gradient(110deg,#0020C2,#00005e);cursor:pointer;font-size:15px}
.pm-b:disabled{background:#9e9e9e;cursor:default}
.pm-e{margin:10px 0;padding:10px 12px;border-radius:12px;background:#fef2f2;color:#b91c1c;font-size:13.5px}.pm-o{margin:10px 0;padding:10px 12px;border-radius:12px;background:#ecfdf3;color:#067647;font-size:13.5px}
.pm-link{background:none;border:0;color:#0020C2;font-weight:700;cursor:pointer;font-size:14px}.pm-c{margin:14px 0 0;text-align:center;color:#6b7280;font-size:13px}
</style>"""
        js = """<script>(function(){var d=document.querySelectorAll('.pm-d'),b=document.getElementById('pm-b'),h=document.getElementById('pm-code'),f=document.getElementById('pm-f');
function val(){var s='';for(var i=0;i<d.length;i++)s+=d[i].value;return s;}function act(){var v=val();b.disabled=!/^\\d{6}$/.test(v);h.value=v;}
Array.prototype.forEach.call(d,function(e,i){e.addEventListener('input',function(){e.value=e.value.replace(/\\D/g,'').slice(-1);if(e.value&&i<5)d[i+1].focus();act();});
e.addEventListener('keydown',function(k){if(k.key==='Backspace'&&!e.value&&i>0){d[i-1].value='';d[i-1].focus();act();}if(k.key==='ArrowLeft'&&i>0)d[i-1].focus();if(k.key==='ArrowRight'&&i<5)d[i+1].focus();});
e.addEventListener('paste',function(k){k.preventDefault();var t=((k.clipboardData||window.clipboardData).getData('text')||'').replace(/\\D/g,'').slice(0,6);for(var j=0;j<6;j++)d[j].value=t.charAt(j)||'';d[Math.min(t.length,5)].focus();act();});});
f.addEventListener('submit',function(){b.disabled=true;});d[0].focus();
var rest=%(resta)d,ex=document.getElementById('pm-exp'),re=document.getElementById('pm-re'),cd=document.getElementById('pm-cd'),t0=Date.now(),wait=%(espera)d;
function two(n){return (n<10?'0':'')+n;}function tick(){var s=Math.max(0,rest-Math.round((Date.now()-t0)/1000));if(ex){ex.textContent=s>0?'El código expira en '+two(Math.floor(s/60))+':'+two(s%%60)+'.':'El código venció. Pide uno nuevo.';}
var w=Math.max(0,wait-Math.round((Date.now()-t0)/1000));if(re){if(w>0){re.style.display='none';cd.textContent='Podrás reenviar en '+w+' s';}else{re.style.display='';cd.textContent='';}}}
tick();setInterval(tick,1000);})();</script>""" % {"resta": resta, "espera": REENVIO_SEG if metodo == "correo" else 0}
        expira = '<p class="pm-s" id="pm-exp"></p>' if metodo == "correo" else ""
        body = ('%s<div class="pm-w"><div class="pm-card"><div class="pm-hd"><a href="/logout" title="Salir" aria-label="Salir">&#8249;</a><b>PROCSIS</b></div><div class="pm-bd">'
                '<h1>%s</h1><p class="pm-s">%s</p>%s%s'
                '<form method="post" id="pm-f" autocomplete="off"><input type="hidden" name="csrf" value="%s"><input type="hidden" name="accion" value="verificar"><input type="hidden" name="codigo" id="pm-code">'
                '<div class="pm-cb">%s</div>%s<button class="pm-b" id="pm-b" type="submit" disabled>VERIFICAR</button></form>%s'
                '<p class="pm-c"><a href="/logout" style="color:#6b7280">Cancelar y salir</a></p></div></div></div>%s'
                % (css, titulo, correo_txt, ('<div class="pm-e">%s</div>' % esc(err)) if err else "", ('<div class="pm-o">%s</div>' % esc(info)) if info else "",
                   csrf_token(), cajas, expira, reenviar, js))
        r = make_response(G["page"]("Verificación en dos pasos", body))
        r.headers["Cache-Control"] = "no-store"
        return r

    def destino_final(u):
        n = _seguro(session.pop("pol_next", ""))
        if n:
            return n
        try:
            return G["_home_portal"](u.rol)
        except Exception:
            return "/dashboard"

    def mfa_politica():
        u = usuario_sesion()
        if not u:
            return redirect("/login")
        pol, nivel = _politica(u)
        if not pol or _modo() == "off":
            return redirect(destino_final(u))
        tablas()
        ahora_ts = time.time()
        stepup = bool(session.get("pol_stepup"))
        if vigente(u, pol, ahora_ts) and not stepup:
            return redirect(destino_final(u))
        metodo = "totp" if tiene_totp(u) else "correo"
        err = info = ""
        if request.method == "POST":
            tk = request.form.get("csrf") or ""
            if not tk or not hmac.compare_digest(tk, session.get("pol_csrf") or ""):
                err = "La sesión de verificación no es válida. Inténtalo de nuevo."
            elif request.form.get("accion") == "reenviar" and metodo == "correo":
                ok, msg = enviar_otp(u)
                err, info = ("" if ok else msg), ("Te enviamos un código nuevo." if ok else "")
            elif request.form.get("accion") == "verificar":
                code = re.sub(r"\D", "", request.form.get("codigo") or "")
                if len(code) != 6:
                    err = "Escribe el código de 6 dígitos."
                else:
                    ok, msg = verificar_totp(u, code) if metodo == "totp" else verificar_otp(u, code)
                    if ok:
                        marcar_ok(u, ahora_ts)
                        session["stepup_ts"] = ahora_ts
                        session.pop("pol_stepup", None)
                        resp = redirect(destino_final(u))
                        if pol["recordar"]:
                            tok = secrets.token_urlsafe(32)
                            db.session.add(MfaDispositivo(usuario_id=u.id, dev_hash=_h(tok), ip_pref=_ip_pref(ip_actual()), ua_hash=_h(request.headers.get("User-Agent", "")), ok_ts=int(ahora_ts)))
                            db.session.commit()
                            resp.set_cookie(COOKIE_DISP, tok, max_age=pol["cada"], httponly=True, samesite="Lax",
                                            secure=bool(app.config.get("SESSION_COOKIE_SECURE")) or request.is_secure, path="/")
                        audit("MFA verificado (%s · %s)" % (nivel, metodo), u.usuario)
                        return resp
                    err = msg
                    audit("MFA fallido (%s)" % nivel, u.usuario)
        elif metodo == "correo" and not reto_vigente(u.id):
            ok, msg = enviar_otp(u)
            if not ok:
                err = msg
        r = reto_vigente(u.id) if metodo == "correo" else None
        resta = max(0, int((r.expira if r else 0) - time.time())) if r else 0
        return pagina(u, metodo, err, info, resta, stepup)

    app.add_url_rule("/mfa-politica", "as2_mfa_politica", mfa_politica, methods=["GET", "POST"])

    # Si el MFA existente (PIN de /mfa) acaba de verificar al usuario, cuenta como MFA de la política (sin pedirlo dos veces).
    orig = G.get("_mfa_completar_login")
    if orig:
        def completar(u, via_dispositivo=False):
            resp = orig(u, via_dispositivo=via_dispositivo)
            try:
                pol, _n = _politica(u)
                if pol and not via_dispositivo:
                    t = time.time()
                    marcar_ok(u, t)
                    session["stepup_ts"] = t
            except Exception:
                pass
            return resp
        G["_mfa_completar_login"] = completar


# ═════════════════════ HTML/JS de la ventana «¿Olvidaste tu contraseña?» (solo correo) ═════════════════════
_REC2_JS_V2 = r"""<style id="rec2-css">
#rec2-ov{position:fixed;inset:0;background:rgba(8,18,60,.55);-webkit-backdrop-filter:blur(5px);backdrop-filter:blur(5px);z-index:2147483000;display:none;align-items:center;justify-content:center;padding:16px;font-family:inherit}
#rec2-ov.on{display:flex}
#rec2-box{background:#fff;color:#0f172a;width:100%;max-width:420px;max-height:94vh;overflow:auto;border-radius:28px;box-shadow:0 30px 90px rgba(10,42,143,.35)}
#rec2-box .hd{display:flex;align-items:center;justify-content:space-between;padding:20px 26px;background:linear-gradient(110deg,#0020C2 0%,#00005e 100%);color:#fff}
#rec2-box .hd button{background:none;border:0;color:#4da3ff;font-size:30px;line-height:1;cursor:pointer;padding:0 6px 2px 0}
#rec2-box .hd b{letter-spacing:.08em;font-size:17px}
#rec2-box .bd{padding:24px 26px 28px}
#rec2-box h2{margin:0 0 8px;font-size:21px;color:#00004d;letter-spacing:-.01em;text-align:center}
#rec2-box p.s{margin:0 0 14px;color:#6b7280;font-size:14px;line-height:1.5}
#rec2-box .ct{text-align:center}
#rec2-box label{display:block;margin:12px 0 5px;font-size:13px;font-weight:700;color:#334155}
#rec2-box input,#rec2-box select,#rec2-box textarea{width:100%;box-sizing:border-box;padding:11px 12px;border:1px solid #cbd5e1;border-radius:11px;font:inherit;font-size:15px;background:#f8fafc;color:#0f172a}
#rec2-box input:focus,#rec2-box select:focus,#rec2-box textarea:focus{outline:2px solid rgba(0,32,194,.35);border-color:#0020C2;background:#fff}
#rec2-box input.pill{border-radius:999px;padding:14px 18px;background:#fff;border-color:#d9dde7}
#rec2-box .cb{display:flex;gap:8px;justify-content:center;margin:6px 0 14px}
#rec2-box input.d{flex:1;min-width:0;max-width:52px;height:56px;padding:0;border-radius:14px;text-align:center;font-size:22px;font-weight:700;color:#00004d;background:#fff;border-color:#d9dde7}
#rec2-box .b{display:block;width:100%;margin-top:16px;padding:12px 14px;border:0;border-radius:12px;background:#0020C2;color:#fff;font:inherit;font-weight:700;font-size:15px;cursor:pointer}
#rec2-box .b.t{border-radius:999px;padding:15px;letter-spacing:.04em;background:linear-gradient(110deg,#0020C2,#00005e)}
#rec2-box .b:disabled{opacity:.55;cursor:default}#rec2-box .b.t:disabled{opacity:1;background:#9e9e9e}
#rec2-box .e{margin:10px 0 0;padding:10px 12px;border-radius:10px;background:#fef2f2;color:#b91c1c;font-size:13.5px;display:none}
#rec2-box .o{margin:10px 0 0;padding:10px 12px;border-radius:10px;background:#ecfdf3;color:#067647;font-size:13.5px}
#rec2-box a.l{color:#0020C2;font-weight:700;cursor:pointer;text-decoration:none;font-size:14px}
#rec2-box .c{border:1px solid #dbe3f5;border-radius:14px;padding:12px 14px;margin:10px 0;background:#f8faff}
#rec2-box .c b{color:#011F68}.c small{display:block;color:#64748b;margin-top:2px;font-size:12.5px}
#rec2-box .c button{margin-top:10px;padding:8px 14px;border:0;border-radius:10px;background:#0020C2;color:#fff;font-weight:700;cursor:pointer}
#rec2-box .c .m{margin-top:8px;color:#b45309;font-size:12.5px}
#rec2-box .f{height:6px;border-radius:99px;background:#e2e8f0;margin-top:8px;overflow:hidden}#rec2-box .f i{display:block;height:100%;width:0;background:#dc2626;transition:width .2s,background .2s}
#rec2-box .p{display:none}#rec2-box .p.on{display:block}
.rec2-link{display:block;margin:10px 0 0;color:#0020C2;font-weight:700;font-size:14px;text-decoration:none;cursor:pointer}
</style>
<div id="rec2-ov" role="dialog" aria-modal="true" aria-labelledby="rec2-t"><div id="rec2-box">
<div class="hd"><button type="button" id="rec2-x" aria-label="Atrás">&#8249;</button><b>PROCSIS</b></div><div class="bd">
<div class="p on" id="rec2-p1"><h2 id="rec2-t">Ingresa con tu correo</h2><p class="s ct">Te enviaremos un código de 6 dígitos a tu correo para recuperar tu cuenta.</p>
<input id="rec2-correo" class="pill" type="email" autocomplete="email" inputmode="email" placeholder="Correo electrónico">
<div class="e" id="rec2-e1"></div><button type="button" class="b t" id="rec2-b1" disabled>ENVÍAME UN CÓDIGO</button>
<p style="margin:16px 0 0;text-align:center"><a class="l" id="rec2-nuevo">¿No tienes usuario? Solicitar un usuario nuevo</a></p></div>
<div class="p" id="rec2-p2"><h2>Ingresa el código de verificación</h2><p class="s ct" id="rec2-m2"></p>
<div class="cb" id="rec2-cb"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="one-time-code" aria-label="Dígito 1"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="off" aria-label="Dígito 2"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="off" aria-label="Dígito 3"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="off" aria-label="Dígito 4"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="off" aria-label="Dígito 5"><input class="d" inputmode="numeric" pattern="[0-9]*" maxlength="1" autocomplete="off" aria-label="Dígito 6"></div>
<p class="s ct" id="rec2-exp" style="font-size:13px"></p>
<div class="e" id="rec2-e2"></div><button type="button" class="b t" id="rec2-b2" disabled>VERIFICAR</button>
<p style="margin:14px 0 0;text-align:center"><a class="l" id="rec2-re">Reenviar código</a> <span id="rec2-cd" style="color:#64748b;font-size:13px"></span></p></div>
<div class="p" id="rec2-p3"><h2>Tus cuentas</h2><p class="s ct">Estas son las cuentas registradas con tu correo. Elige la que quieres recuperar.</p><div id="rec2-lista"></div></div>
<div class="p" id="rec2-p4"><h2>Crea tu nueva contraseña</h2><p class="s ct" id="rec2-m4"></p>
<label for="rec2-c1">Nueva contraseña</label><input id="rec2-c1" type="password" autocomplete="new-password"><div class="f"><i id="rec2-f"></i></div>
<p class="s" style="font-size:12.5px;margin:6px 0 0">Mínimo 10 caracteres, con mayúsculas, minúsculas, números o símbolos (al menos 3 tipos). No uses tu usuario ni tu correo.</p>
<label for="rec2-c2">Repite la contraseña</label><input id="rec2-c2" type="password" autocomplete="new-password">
<label style="font-weight:500"><input type="checkbox" id="rec2-ver" style="width:auto;margin-right:6px">Mostrar contraseñas</label>
<div class="e" id="rec2-e4"></div><button type="button" class="b" id="rec2-b4">Guardar contraseña</button></div>
<div class="p" id="rec2-p5"><h2>¡Listo!</h2><div class="o" id="rec2-m5"></div><button type="button" class="b" id="rec2-b5">Ir a iniciar sesión</button></div>
<div class="p" id="rec2-p6"><h2>Solicitar un usuario nuevo</h2><p class="s ct">Cuéntanos quién eres. Verificamos tu correo con un código y Soporte validará tus datos antes de crear el usuario.</p>
<label for="rec2-sn">Nombre completo</label><input id="rec2-sn" autocomplete="name">
<label for="rec2-sc">Correo electrónico</label><input id="rec2-sc" type="email" autocomplete="email">
<label for="rec2-st">Celular de contacto</label><input id="rec2-st" type="tel" autocomplete="tel">
<label for="rec2-sp">Tipo de cuenta</label><select id="rec2-sp"><option>Acudiente / familia</option><option>Docente</option><option>Personal administrativo del colegio</option><option>Rector o directivo</option><option>Equipo PROCSIS (área interna)</option><option>Otro</option></select>
<label for="rec2-si">Colegio o área</label><input id="rec2-si">
<label for="rec2-sm">Mensaje (opcional)</label><textarea id="rec2-sm" rows="2"></textarea>
<div class="e" id="rec2-e6"></div><button type="button" class="b" id="rec2-b6">Enviar código a mi correo</button></div>
<div class="p" id="rec2-p7"><h2>Confirma tu solicitud</h2><p class="s ct" id="rec2-m7"></p>
<label for="rec2-sx">Código de 6 dígitos</label><input id="rec2-sx" inputmode="numeric" maxlength="6" autocomplete="one-time-code">
<div class="e" id="rec2-e7"></div><button type="button" class="b" id="rec2-b7">Confirmar solicitud</button></div>
<div class="p" id="rec2-p8"><h2>Solicitud enviada</h2><div class="o" id="rec2-m8"></div><button type="button" class="b" id="rec2-b8">Cerrar</button></div>
</div></div></div>
<script id="rec2-js">
(function(){
  var AUTO=__AUTO__,ov=document.getElementById('rec2-ov'),tok='',tokS='',cd=0,timer=null,tmrExp=null,expAt=0,cuentaSel=null;
  function $(i){return document.getElementById(i);}
  function two(n){return (n<10?'0':'')+n;}
  var RX=/^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/,cur=1;
  function pane(n){cur=n;for(var i=1;i<=8;i++){$('rec2-p'+i).className='p'+(i===n?' on':'');}var a=$('rec2-p'+n).querySelector('input');if(a)setTimeout(function(){a.focus();},50);}
  function err(id,m){var e=$(id);e.textContent=m||'';e.style.display=m?'block':'none';}
  function post(url,data){return fetch(url,{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}).then(function(r){return r.json().catch(function(){return {ok:false,mensaje:'Respuesta no válida.'};});}).catch(function(){return {ok:false,mensaje:'No hay conexión. Inténtalo de nuevo.'};});}
  function abrir(correo){ov.className='on';pane(1);if(correo)$('rec2-correo').value=correo;act1();document.documentElement.style.overflow='hidden';}
  function cerrar(){ov.className='';document.documentElement.style.overflow='';clearInterval(tmrExp);}
  function atras(){if(cur===2||cur===6){pane(1);}else if(cur===7){pane(6);}else{cerrar();}}
  function busy(b,t){b.disabled=!!t;if(t){b.setAttribute('data-t',b.textContent);b.textContent=t;}else if(b.getAttribute('data-t')){b.textContent=b.getAttribute('data-t');}}
  function act1(){$('rec2-b1').disabled=!RX.test($('rec2-correo').value.trim());}
  var ds=$('rec2-cb').querySelectorAll('.d');
  function pinVal(){var s='',i;for(i=0;i<ds.length;i++){s+=ds[i].value;}return s;}
  function pinSet(s){for(var i=0;i<ds.length;i++){ds[i].value=s.charAt(i)||'';}}
  function act2(){$('rec2-b2').disabled=!/^\d{6}$/.test(pinVal());}
  Array.prototype.forEach.call(ds,function(e,i){
    e.addEventListener('input',function(){e.value=e.value.replace(/\D/g,'').slice(-1);if(e.value&&i<5){ds[i+1].focus();}act2();});
    e.addEventListener('keydown',function(k){if(k.key==='Backspace'&&!e.value&&i>0){ds[i-1].value='';ds[i-1].focus();act2();}if(k.key==='ArrowLeft'&&i>0){ds[i-1].focus();}if(k.key==='ArrowRight'&&i<5){ds[i+1].focus();}});
    e.addEventListener('paste',function(k){k.preventDefault();var t=((k.clipboardData||window.clipboardData).getData('text')||'').replace(/\D/g,'').slice(0,6);pinSet(t);ds[Math.min(t.length,5)].focus();act2();});
  });
  function cuenta(){clearInterval(timer);cd=60;$('rec2-re').style.display='none';$('rec2-cd').textContent='Podrás reenviar en 60 s';timer=setInterval(function(){cd--;$('rec2-cd').textContent=cd>0?'Podrás reenviar en '+cd+' s':'';if(cd<=0){clearInterval(timer);$('rec2-re').style.display='';}},1000);}
  function tick(){var s=Math.max(0,Math.round((expAt-Date.now())/1000));$('rec2-exp').textContent=s>0?'El código expira en '+two(Math.floor(s/60))+':'+two(s%60)+'.':'El código venció. Pide uno nuevo.';if(s<=0){clearInterval(tmrExp);}}
  function arrancaExp(){clearInterval(tmrExp);expAt=Date.now()+300000;tick();tmrExp=setInterval(tick,1000);}
  function mascara(c){var p=c.split('@');return p[0].slice(0,2)+'***@'+p[1];}
  function msg2(c){var m=$('rec2-m2');m.textContent='';m.appendChild(document.createTextNode('Si el correo está registrado, te enviamos un código de 6 dígitos a '));var b=document.createElement('b');b.textContent=mascara(c);m.appendChild(b);m.appendChild(document.createTextNode('.'));}
  function enviarPin(re){var b=re?null:$('rec2-b1'),eid=re?'rec2-e2':'rec2-e1',c=$('rec2-correo').value.trim();err(eid,'');if(!RX.test(c)){err(eid,'Escribe un correo válido.');return;}if(b){busy(b,'ENVIANDO…');}
    post('/acceso/recuperar/pin',{correo:c}).then(function(r){if(b){busy(b);}if(!r.ok){err(eid,r.mensaje);return;}tok=r.token;msg2(c);pinSet('');act2();pane(2);cuenta();arrancaExp();});}
  function verificar(){var b=$('rec2-b2');err('rec2-e2','');busy(b,'VERIFICANDO…');
    post('/acceso/recuperar/verificar',{token:tok,pin:pinVal()}).then(function(r){busy(b);if(!r.ok){err('rec2-e2',r.mensaje);pinSet('');act2();ds[0].focus();return;}clearInterval(tmrExp);lista(r.cuentas);pane(3);});}
  function lista(cs){var L=$('rec2-lista');L.innerHTML='';cs.forEach(function(c){var d=document.createElement('div');d.className='c';var b=document.createElement('b');b.textContent=c.usuario;d.appendChild(b);
    var s=document.createElement('small');s.textContent=(c.nombre?c.nombre+' · ':'')+c.area;d.appendChild(s);
    if(c.puede){var k=document.createElement('button');k.type='button';k.textContent='Restablecer esta cuenta';k.onclick=function(){elegir(c);};d.appendChild(k);}else{var m=document.createElement('div');m.className='m';m.textContent=c.motivo;d.appendChild(m);}L.appendChild(d);});}
  function elegir(c){cuentaSel=c;$('rec2-m4').textContent='Cuenta: '+c.usuario+' · '+c.area;$('rec2-c1').value='';$('rec2-c2').value='';err('rec2-e4','');fuerza('');pane(4);}
  function fuerza(v){var n=0;if(v.length>=10)n++;if(/[a-z]/.test(v)&&/[A-Z]/.test(v))n++;if(/\d/.test(v))n++;if(/[^A-Za-z0-9]/.test(v))n++;if(v.length>=14)n++;var f=$('rec2-f');f.style.width=(n*20)+'%';f.style.background=n<3?'#dc2626':(n<4?'#f59e0b':'#16a34a');}
  function guardar(){var b=$('rec2-b4');err('rec2-e4','');busy(b,'Guardando…');
    post('/acceso/recuperar/restablecer',{token:tok,indice:cuentaSel.i,clave:$('rec2-c1').value,clave2:$('rec2-c2').value}).then(function(r){busy(b);if(!r.ok){err('rec2-e4',r.mensaje);return;}$('rec2-m5').textContent=r.mensaje+' Usuario: '+r.usuario;cuentaSel.u=r.usuario;$('rec2-c1').value='';$('rec2-c2').value='';pane(5);});}
  function irLogin(){cerrar();var u=document.querySelector('input[name="usuario"]');if(u&&cuentaSel&&cuentaSel.u){u.value=cuentaSel.u;}var p=document.querySelector('input[name="password"]');if(p){p.value='';p.focus();}}
  function solPin(){var b=$('rec2-b6');err('rec2-e6','');busy(b,'Enviando…');
    post('/acceso/solicitud/pin',{nombre:$('rec2-sn').value,correo:$('rec2-sc').value,telefono:$('rec2-st').value,tipo:$('rec2-sp').value,institucion:$('rec2-si').value,mensaje:$('rec2-sm').value}).then(function(r){busy(b);if(!r.ok){err('rec2-e6',r.mensaje);return;}tokS=r.token;$('rec2-m7').textContent=r.mensaje;$('rec2-sx').value='';pane(7);});}
  function solOk(){var b=$('rec2-b7');err('rec2-e7','');busy(b,'Confirmando…');
    post('/acceso/solicitud/confirmar',{token:tokS,pin:$('rec2-sx').value}).then(function(r){busy(b);if(!r.ok){err('rec2-e7',r.mensaje);return;}$('rec2-m8').textContent=r.mensaje;pane(8);});}
  $('rec2-b1').onclick=function(){enviarPin(false);};$('rec2-b2').onclick=verificar;$('rec2-b4').onclick=guardar;$('rec2-b5').onclick=irLogin;$('rec2-b6').onclick=solPin;$('rec2-b7').onclick=solOk;$('rec2-b8').onclick=cerrar;
  $('rec2-re').onclick=function(){enviarPin(true);};$('rec2-nuevo').onclick=function(){pane(6);};$('rec2-x').onclick=atras;
  $('rec2-correo').oninput=act1;$('rec2-c1').oninput=function(){fuerza(this.value);};
  $('rec2-ver').onchange=function(){var t=this.checked?'text':'password';$('rec2-c1').type=t;$('rec2-c2').type=t;};
  ov.addEventListener('mousedown',function(e){if(e.target===ov&&cur!==3&&cur!==4){cerrar();}});
  document.addEventListener('keydown',function(e){if(e.key==='Escape'&&ov.className==='on'){cerrar();}if(e.key==='Enter'&&ov.className==='on'){var a=document.activeElement;if(a&&a.tagName==='INPUT'){var p=a.closest('.p');if(p){var b=p.querySelector('.b');if(b&&!b.disabled){b.click();}}}}});
  var pw=document.querySelector('input[name="password"]');
  if(pw){var a=document.createElement('a');a.className='rec2-link';a.textContent='¿Olvidaste tu contraseña?';a.href='#';a.onclick=function(e){e.preventDefault();abrir('');};var host=pw.closest('form')||pw.parentNode;var bt=host.querySelector('button[type="submit"],input[type="submit"]');if(bt){bt.parentNode.insertBefore(a,bt.nextSibling);}else{host.appendChild(a);}}
  if(AUTO){setTimeout(function(){abrir('');},250);}
})();
</script>"""

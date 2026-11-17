# -*- coding: utf-8 -*-
"""
aplicar_fix_mfa_cambio_clave.py

Corrige dos cosas en app.py (SOLO código, no toca base de datos ni archivos subidos):

1) ERR_TOO_MANY_REDIRECTS en /mfa
   Un usuario con clave temporal + MFA pendiente rebotaba para siempre:
     /mfa  ->  /cambiar_password  (hook de password_temporal)
     /cambiar_password  ->  /mfa  (hook de mfa_pendiente)
   - El hook de clave temporal ya no actúa mientras el MFA esté pendiente ni sobre /mfa.
   - /mfa: si al usuario ya no se le exige MFA (se lo quitaron) y quedó una sesión vieja
     con mfa_pendiente, se completa el ingreso en vez de quedarse atascado.
     (El rol Soluciones sigue exigiendo siempre su PIN.)
   - /cambiar_password: si hay MFA pendiente, manda primero a /mfa.

2) Cambio de contraseña con PIN por correo
   Antes de poder escribir la nueva clave se envía un PIN de 6 dígitos al correo
   registrado del usuario (vence en 5 min, máx. 5 intentos y 5 envíos). El PIN
   verificado vale 10 minutos para guardar la nueva clave.

Uso (en la carpeta donde está app.py):
    python aplicar_fix_mfa_cambio_clave.py            # usa ./app.py
    python aplicar_fix_mfa_cambio_clave.py ruta/app.py

Seguridad del script: hace respaldo app.py.bak_fix_mfa_cambio_clave, exige que cada
fragmento a cambiar exista EXACTAMENTE una vez (si no, aborta sin modificar nada),
valida la sintaxis del resultado antes de escribir y conserva los saltos de línea.
"""
import ast
import os
import shutil
import sys

MARCA = "[FIX-MFA-CAMBIO-CLAVE]"

RUTA = sys.argv[1] if len(sys.argv) > 1 else "app.py"

# ---------------------------------------------------------------------------
# Edición A: el hook de clave temporal no debe pelear con el MFA
# ---------------------------------------------------------------------------
A_VIEJO = 'if session.get("password_temporal") and request.path not in ["/cambiar_password", "/logout", "/docente-login"]'
A_NUEVO = ('if session.get("password_temporal") and not session.get("mfa_pendiente") '
           'and not request.path.startswith(("/mfa", "/biometria")) '
           'and request.path not in ["/cambiar_password", "/logout", "/docente-login"]')

# ---------------------------------------------------------------------------
# Edición B: /mfa no debe atrapar a quien ya no tiene MFA exigido
# ---------------------------------------------------------------------------
B_VIEJO = '''    registrado = (u.correo or "").strip()
    error = info = ""
    accion = (request.form.get("accion") or "") if request.method == "POST" else ""
    ch = None
    tok = session.get("mfa_pin_token")'''
B_NUEVO = '''    # [FIX-MFA-CAMBIO-CLAVE] Si ya no se le exige MFA (se lo quitaron) no se queda atascado en /mfa.
    if not _mfa_requerido(u) and (u.rol or "").strip() != ROL_SOLUCIONES:
        return _mfa_completar_login(u)
    registrado = (u.correo or "").strip()
    error = info = ""
    accion = (request.form.get("accion") or "") if request.method == "POST" else ""
    ch = None
    tok = session.get("mfa_pin_token")'''

# ---------------------------------------------------------------------------
# Edición C: /cambiar_password completo, con PIN por correo antes de cambiar
# ---------------------------------------------------------------------------
C_INICIO = '@app.route("/cambiar_password", methods=["GET", "POST"])\ndef cambiar_password():'
C_FIN = '@app.route("/rectoria/tareas")'

C_NUEVO = r'''def _enviar_pin_cambio_clave(correo, pin):
    """PIN para cambiar la contraseña (se envía al correo registrado del usuario)."""
    if not (os.environ.get("RESEND_API_KEY") or "").strip() and (not SOPORTE_EMAIL or not SOPORTE_PASSWORD):
        return False
    msg = EmailMessage()
    msg["Subject"] = "Código para cambiar su contraseña - Procsis"
    msg["From"] = SOPORTE_EMAIL
    msg["To"] = correo
    msg.set_content(
        "Su código para cambiar la contraseña es: %s\n"
        "Vence en 5 minutos. Si usted no lo solicitó, ignore este mensaje y avise a Gerencia.\n%s" % (pin, APP_NAME)
    )
    try:
        _smtp_enviar(msg, SOPORTE_EMAIL, SOPORTE_PASSWORD)
        return True
    except Exception as e:
        print("ERROR PIN cambio de clave:", e)
        return False


@app.route("/cambiar_password", methods=["GET", "POST"])
def cambiar_password():
    # [FIX-MFA-CAMBIO-CLAVE] Antes de escribir la nueva clave: PIN enviado al correo registrado.
    if not requiere_login(): return redirect("/login")
    if session.get("mfa_pendiente"):
        return redirect("/mfa")
    mensaje = info = ""
    rol = rol_actual()
    es_staff = rol in ("Soporte", "Comercial", "Gerente", "Administrador", "Superadmin", "Cobranza", "Contabilidad", "Proveedor", "Soluciones") + ROLES_AREAS
    u = Usuario.query.filter_by(usuario=session["usuario"]).first()
    if not u:
        session.clear()
        return redirect("/login")
    registrado = (u.correo or "").strip()
    accion = (request.form.get("accion") or "") if request.method == "POST" else ""

    def _pin_verificado():
        try:
            return (ahora().timestamp() - float(session.get("cp_pin_ok") or 0)) < 600
        except Exception:
            return False

    ch = None
    tok = session.get("cp_pin_token")
    if tok:
        ch = LoginChallenge.query.filter_by(token=tok, usuario_id=u.id).first()

    if accion == "reiniciar":
        session.pop("cp_pin_token", None)
        session.pop("cp_pin_ok", None)
        ch = None
    elif accion == "enviar":
        correo = (request.form.get("correo") or "").strip().lower()
        envios = int(session.get("cp_envios", 0))
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", correo):
            mensaje = "Escriba un correo válido."
        elif registrado and correo != registrado.lower():
            mensaje = "Ese correo no coincide con el registrado para su usuario."
            registrar_auditoria("PIN cambio de clave: correo distinto al registrado", "%s · %s" % (u.usuario, correo))
        elif envios >= 5:
            mensaje = "Demasiados envíos de PIN. Cierre sesión e intente más tarde."
            registrar_auditoria("PIN cambio de clave: demasiados envíos", u.usuario)
        else:
            ch = _crear_reto_pin(u.id)
            if _enviar_pin_cambio_clave(correo, ch.pin):
                session["cp_pin_token"] = ch.token
                session["cp_correo"] = correo
                session["cp_envios"] = envios + 1
                session["cp_intentos"] = 0
                session.pop("cp_pin_ok", None)
                registrar_auditoria("PIN cambio de clave enviado", "%s · %s" % (u.usuario, _mask_correo(correo)))
                info = "Enviamos un PIN de 6 dígitos a %s." % _mask_correo(correo)
            else:
                ch = None
                mensaje = ("No se pudo enviar el correo. Verifique que el correo de envío del sistema "
                           "(SOPORTE_EMAIL / RESEND_API_KEY) esté configurado y vuelva a intentar.")
                registrar_auditoria("PIN cambio de clave: fallo de envío", u.usuario)
    elif accion == "validar" and ch:
        intentos = int(session.get("cp_intentos", 0)) + 1
        session["cp_intentos"] = intentos
        pin = (request.form.get("pin") or "").strip().replace(" ", "")
        if not _reto_vigente(ch):
            mensaje = "El PIN venció. Solicite uno nuevo."
            session.pop("cp_pin_token", None)
            ch = None
        elif pin != (ch.pin or ""):
            mensaje = "PIN incorrecto."
            registrar_auditoria("PIN cambio de clave incorrecto", "%s · intento %d" % (u.usuario, intentos))
            if intentos >= 5:
                registrar_auditoria("PIN cambio de clave bloqueado por intentos", u.usuario)
                session.pop("cp_pin_token", None)
                ch = None
                mensaje = "Demasiados intentos. Solicite un PIN nuevo."
        else:
            ch.estado = "aprobado"
            if not registrado and session.get("cp_correo"):
                u.correo = session.get("cp_correo")
            db.session.commit()
            session["cp_pin_ok"] = ahora().timestamp()
            session.pop("cp_pin_token", None)
            ch = None
            registrar_auditoria("PIN cambio de clave verificado", u.usuario)
            info = "PIN verificado. Ya puede definir su nueva contraseña."
    elif accion == "guardar":
        if not _pin_verificado():
            mensaje = "Primero verifique el PIN enviado a su correo."
        else:
            nueva = request.form.get("password", "").strip()
            confirmar = request.form.get("confirmar", "").strip()
            nombre_completo = (request.form.get("nombre_completo") or "").strip()[:160]
            _err_clave = _politica_clave(nueva, session.get("usuario") or "", es_staff)
            if _err_clave:
                mensaje = _err_clave
            elif nueva != confirmar:
                mensaje = "Las contraseñas no coinciden."
            elif es_staff and not nombre_completo:
                mensaje = "Escriba su nombre para mostrar."
            else:
                u.password = crear_hash(nueva); u.password_temporal = False
                if es_staff:
                    u.nombre_completo = nombre_completo
                db.session.commit(); session["password_temporal"] = False
                for _k in ("cp_pin_token", "cp_pin_ok", "cp_correo", "cp_envios", "cp_intentos"):
                    session.pop(_k, None)
                registrar_auditoria("Cambio de contraseña", "Usuario cambió su contraseña (verificada con PIN de correo)" + (f" · nombre: {nombre_completo}" if es_staff else ""))
                if es_staff:
                    destino = {"Soporte": "/soporte_admin", "Comercial": "/ventas/panel", "Gerente": "/gerencia/hq",
                               "Administrador": "/gerencia/hq", "Superadmin": "/gerencia/hq", "Cobranza": "/cobranza/panel",
                               "Soluciones": "/soluciones-procsis", **{k: "/" + v["slug"] for k, v in AREAS_PROCSIS.items()}}.get(rol, "/dashboard")
                    return redirect(destino)
                return redirect("/dashboard")

    estilo_in = "width:100%;box-sizing:border-box;padding:12px;font-size:15px;border:1px solid #e2e8f0;border-radius:10px;margin-bottom:12px"
    estilo_btn = "width:100%;padding:13px;border:0;border-radius:10px;background:#0020C2;color:#fff;font-weight:800;cursor:pointer"
    aviso_ok = ("<div style='background:#dcfce7;color:#166534;padding:10px;border-radius:8px;margin:10px 0;font-size:13px'>" + _esc(info) + "</div>") if info else ""
    aviso_err = ('<p class="error">' + _esc(mensaje) + '</p>') if mensaje else ""

    if _pin_verificado():
        campo_nombre = (
            '<input name="nombre_completo" placeholder="Su nombre completo (para mostrar)" required>'
            if es_staff else ""
        )
        intro = "Correo verificado. Defina una nueva contraseña de su preferencia (mín. 8 caracteres)" + (" y su nombre" if es_staff else "") + "."
        paso = f"""<form method="POST"><input type="hidden" name="accion" value="guardar">{campo_nombre}<input name="password" type="password" placeholder="Nueva contraseña" required><input name="confirmar" type="password" placeholder="Confirmar contraseña" required><button>Guardar contraseña</button></form>"""
    elif ch and session.get("cp_pin_token"):
        intro = "Escriba el PIN de 6 dígitos enviado a " + _esc(_mask_correo(session.get("cp_correo", ""))) + ". Vence en 5 minutos."
        paso = f"""<form method="POST"><input type="hidden" name="accion" value="validar"><input name="pin" inputmode="numeric" maxlength="6" autocomplete="one-time-code" placeholder="000000" required autofocus style="{estilo_in};font-size:20px;letter-spacing:8px;text-align:center"><button style="{estilo_btn}">Validar PIN</button></form><form method="POST" style="text-align:center;margin-top:10px"><input type="hidden" name="accion" value="reiniciar"><button style="background:none;border:0;color:#0020C2;font-size:12px;cursor:pointer;text-decoration:underline">Cambiar correo / reenviar PIN</button></form>"""
    else:
        intro = "Por seguridad, antes de cambiar su contraseña le enviaremos un PIN de verificación al correo registrado."
        paso = f"""<form method="POST"><input type="hidden" name="accion" value="enviar"><input name="correo" type="email" required value="{_esc(registrado)}" placeholder="correo@ejemplo.com" style="{estilo_in}"><button style="{estilo_btn}">Enviar PIN</button></form>"""

    return page("Cambiar contraseña", f"""
<div class="center"><section class="card login-card"><h1>Cambiar contraseña</h1><p>{intro}</p>{aviso_ok}{aviso_err}{paso}<a href="/logout">Salir</a></section></div>
""")


'''


def contar(txt, frag):
    return txt.count(frag)


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

    # --- validar que cada fragmento exista exactamente una vez ---
    for nombre, frag in (("A (hook clave temporal)", A_VIEJO), ("B (/mfa)", B_VIEJO),
                         ("C inicio (/cambiar_password)", C_INICIO), ("C fin", C_FIN)):
        n = contar(txt, frag)
        if n != 1:
            print("ABORTADO: el fragmento %s aparece %d veces (debe ser 1). No se modificó nada." % (nombre, n))
            sys.exit(1)

    i = txt.index(C_INICIO)
    j = txt.index(C_FIN)
    if j <= i:
        print("ABORTADO: el orden de los marcadores de /cambiar_password no es el esperado.")
        sys.exit(1)

    nuevo = txt.replace(A_VIEJO, A_NUEVO)
    nuevo = nuevo.replace(B_VIEJO, B_NUEVO)
    i = nuevo.index(C_INICIO)
    j = nuevo.index(C_FIN)
    nuevo = nuevo[:i] + C_NUEVO + nuevo[j:]

    # --- validar sintaxis antes de escribir ---
    try:
        ast.parse(nuevo)
    except SyntaxError as e:
        print("ABORTADO: el resultado tendría un error de sintaxis (línea %s). No se modificó nada." % e.lineno)
        sys.exit(1)

    salida = nuevo.replace("\n", "\r\n") if crlf else nuevo

    respaldo = RUTA + ".bak_fix_mfa_cambio_clave"
    shutil.copy2(RUTA, respaldo)
    tmp = RUTA + ".tmp_fix"
    with open(tmp, "wb") as f:
        f.write(salida.encode("utf-8"))
    os.replace(tmp, RUTA)

    print("Listo.")
    print(" - Respaldo: %s" % respaldo)
    print(" - Hook de clave temporal ya no choca con /mfa")
    print(" - /mfa completa el ingreso si al usuario ya no se le exige MFA")
    print(" - /cambiar_password ahora exige PIN al correo registrado antes de guardar")
    print("Reinicie la app / haga deploy para que tome efecto.")


if __name__ == "__main__":
    main()

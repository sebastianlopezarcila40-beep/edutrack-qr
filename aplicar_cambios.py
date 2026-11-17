import io

import sys, os, shutil, py_compile
SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
DST = SRC
BLOQUE_COB = r'''def _cob_norm(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]", "", s)


def _cob_cargar_munis(depto_id=None):
    """Carga los municipios de Colombia (DIVIPOLA) en los departamentos que todavía no tienen pueblos.
    Quedan INACTIVOS hasta que Gerencia marque en cuáles opera PROCSIS. Devuelve cuántos agregó."""
    data = _cargar_divipola() or {}
    por_norm = {}
    for k, v in data.items():
        por_norm[_cob_norm(k)] = v or []
    if depto_id:
        filas = _c48_q("SELECT id, nombre FROM cobertura_deptos WHERE pais='CO' AND id=:i", i=depto_id)
    else:
        filas = _c48_q("SELECT id, nombre FROM cobertura_deptos WHERE pais='CO'")
    total = 0
    for did, dn in filas:
        ya = db.session.execute(text("SELECT COUNT(*) FROM cobertura_munis WHERE depto_id=:i"), {"i": did}).scalar() or 0
        if int(ya) > 0:
            continue
        vistos = set()
        for j, m in enumerate(por_norm.get(_cob_norm(dn)) or []):
            if isinstance(m, (list, tuple)) and m:
                m = m[0]
            nombre = " ".join(str(m).split())[:120]
            if not nombre or nombre.lower() in vistos:
                continue
            vistos.add(nombre.lower())
            db.session.execute(text("INSERT INTO cobertura_munis (depto_id, nombre, activo, orden) VALUES (:d, :n, 0, :o)"),
                               {"d": did, "n": nombre, "o": j})
            total += 1
    db.session.commit()
    return total


def _cob_pueblos_link(cod, did, cnt):
    tot, on = cnt.get(did, (0, 0))
    return "<a class='cb-btn s o' href='/gerencia/cobertura?pais=%s&amp;depto=%d#pueblos'>Pueblos (%d/%d)</a> " % (_esc(cod), int(did), on, tot)


def _cob_html_pueblos(cod, depto_id, csrf):
    """Tarjeta con los pueblos (municipios) de un departamento: cuáles están activos y cuáles no."""
    d = _c48_q("SELECT id, nombre FROM cobertura_deptos WHERE id=:i AND pais=:p", i=depto_id, p=cod)
    if not d:
        return ""
    dn = d[0][1]
    ms = _c48_q("SELECT id, nombre, activo FROM cobertura_munis WHERE depto_id=:i ORDER BY nombre", i=depto_id)
    on = sum(1 for m in ms if int(m[2] or 0) == 1)
    base = {"pais": cod, "depto_id": depto_id}
    filas = ""
    for mid, mn, ma in ms:
        act = int(ma or 0) == 1
        filas += ("<tr class='cb-mrow'><td><b>%s</b></td><td><span class='cb-tag %s'>%s</span></td><td>%s%s</td></tr>"
                  % (_esc(mn), "on" if act else "off", "Activo" if act else "Inactivo",
                     _cob_boton(csrf, "muni_toggle", dict(base, id=mid), "Desactivar" if act else "Activar", "r" if act else "g"),
                     _cob_boton(csrf, "muni_eliminar", dict(base, id=mid), "Eliminar", "r", "¿Eliminar %s de la lista?" % mn.replace("'", ""))))
    if not filas:
        filas = ("<tr><td colspan='3'>Este departamento aún no tiene pueblos cargados. "
                 "Agréguelos abajo o use «Cargar municipios oficiales».</td></tr>")
    masivo = (_cob_boton(csrf, "muni_todos", dict(base, valor=1), "Activar todos", "g")
              + _cob_boton(csrf, "muni_todos", dict(base, valor=0), "Desactivar todos", "r",
                           "¿Marcar todos los pueblos de %s como inactivos?" % dn.replace("'", "")))
    cargar = _cob_boton(csrf, "muni_cargar", base, "Cargar municipios oficiales (DIVIPOLA)", "o") if cod == "CO" else ""
    buscar = ("<input type='text' id='cbMq' placeholder='Buscar pueblo…' style='max-width:260px;margin:10px 0' "
              "oninput=\"var q=this.value.toLowerCase();document.querySelectorAll('.cb-mrow').forEach(function(r){"
              "r.style.display=r.textContent.toLowerCase().indexOf(q)<0?'none':'';});\">")
    return ("<div class='cb-card' id='pueblos'><h2>Pueblos de %s · %d activos de %d</h2>"
            "<p>Marque en qué pueblos opera PROCSIS (activos) y en cuáles no. También puede agregar los que falten.</p>"
            "<p>%s %s</p>%s<div class='cb-scroll'><table><tr><th>Pueblo / municipio</th><th>Estado</th><th>Acciones</th></tr>%s</table></div>"
            "<form method='POST' action='/gerencia/cobertura' style='display:flex;gap:8px;margin-top:14px;align-items:end;flex-wrap:wrap'>%s"
            "<div style='flex:1;min-width:240px'><label>Agregar pueblos a %s (separados por coma)</label>"
            "<input type='text' name='nombre' required maxlength='400'></div>"
            "<button class='cb-btn g'>Agregar</button></form>"
            "<p style='margin-top:10px'><a class='cb-btn s o' href='/gerencia/cobertura?pais=%s#zonas'>← Volver a los departamentos</a></p></div>"
            % (_esc(dn), on, len(ms), masivo, cargar, buscar, filas, _cob_hidden(csrf, "muni_agregar", base), _esc(dn), _esc(cod)))


'''
BLOQUE_FINAL = r'''# ════════════════════════════════════════════════════════════════════════════
#  LOGOS DEL INICIO Y BANNER DE /procsis (administrados desde Gerencia)
#  - Logo de PROCSIS (la "P" de la barra superior) y logo de la empresa / EduTrack (inicio y carrusel de acceso).
#    Se cambian sin tocar archivos: Gerencia sube la imagen y el sistema la sirve en la misma dirección de siempre.
#  - Banner con foto bajo "¿Quiénes somos?" en la página /procsis.
# ════════════════════════════════════════════════════════════════════════════
_MARCA_CLAVES = ("logo_ini_procsis", "logo_ini_empresa", "banner_proc_img", "banner_proc_url", "banner_proc_alt", "banner_proc_on")
_MARCA_CACHE = {"t": 0.0, "v": {}, "init": False}


def _marca_get(clave):
    """Valores guardados por Gerencia (copia en memoria de 60 s para no consultar la base en cada imagen)."""
    import time
    if time.time() - _MARCA_CACHE["t"] > 60:
        v = {}
        try:
            if not _MARCA_CACHE["init"]:
                _ti_init()
                _MARCA_CACHE["init"] = True
            marcas = ",".join("'%s'" % k for k in _MARCA_CLAVES)
            for k, val in _c48_q("SELECT clave, valor FROM ti_config WHERE clave IN (%s)" % marcas):
                v[k] = val or ""
        except Exception:
            try:
                db.session.rollback()
            except Exception:
                pass
        _MARCA_CACHE["v"] = v
        _MARCA_CACHE["t"] = time.time()
    return _MARCA_CACHE["v"].get(clave, "")


def _marca_invalidar():
    _MARCA_CACHE["t"] = 0.0


def _marca_decodificar(uri):
    m = re.match(r"^data:(image/(?:png|jpeg|webp|svg\+xml));base64,([A-Za-z0-9+/=]+)$", uri or "")
    if not m:
        return None, None
    import base64 as _b64
    try:
        return _b64.b64decode(m.group(2)), m.group(1)
    except Exception:
        return None, None


def _marca_img_de_archivo(f, max_kb):
    """Imagen subida → data URI. '' si no hay archivo; None si no es una imagen válida o pesa de más."""
    if f is None or not getattr(f, "filename", ""):
        return ""
    raw = f.read(max_kb * 1024 + 1)
    if not raw:
        return ""
    if len(raw) > max_kb * 1024:
        return None
    if raw[:8] == b"\x89PNG\r\n\x1a\n":
        mime = "image/png"
    elif raw[:3] == b"\xff\xd8\xff":
        mime = "image/jpeg"
    elif raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        mime = "image/webp"
    elif raw.lstrip()[:4] in (b"<svg", b"<?xm"):
        mime = "image/svg+xml"
    else:
        return None
    import base64 as _b64
    return "data:%s;base64,%s" % (mime, _b64.b64encode(raw).decode("ascii"))


def _marca_respuesta(raw, mime, cache):
    r = Response(raw, mimetype=mime)
    r.headers["Cache-Control"] = cache
    r.headers["X-Content-Type-Options"] = "nosniff"
    r.headers["Content-Security-Policy"] = "sandbox; default-src 'none'; style-src 'unsafe-inline'"
    return r


def _registrar_logo_override(ruta, clave, endpoint):
    """Si Gerencia subió un logo propio, se sirve ese en la misma dirección del archivo original; si no, el original."""
    def vista():
        raw, mime = _marca_decodificar(_marca_get(clave))
        if raw:
            return _marca_respuesta(raw, mime, "no-cache")
        from flask import send_from_directory
        return send_from_directory(app.static_folder, ruta[len("/static/"):])
    vista.__name__ = endpoint
    app.add_url_rule(ruta, endpoint, vista)


_registrar_logo_override("/static/img/logo-procsis-mark.png", "logo_ini_procsis", "marca_logo_procsis")
_registrar_logo_override("/static/img/logo-edutrack.png", "logo_ini_empresa", "marca_logo_empresa_png")
_registrar_logo_override("/static/img/logo-edutrack.jpeg", "logo_ini_empresa", "marca_logo_empresa_jpeg")


@app.route("/media/banner-procsis")
def media_banner_procsis():
    raw, mime = _marca_decodificar(_marca_get("banner_proc_img"))
    if not raw:
        return Response(status=404)
    return _marca_respuesta(raw, mime, "public, max-age=3600")


def _banner_procsis_html():
    """Área con foto (estilo banner corporativo) que va debajo de «Quiénes somos» en /procsis. Vacío si Gerencia no la activó."""
    try:
        if _marca_get("banner_proc_on") != "1" or not _marca_get("banner_proc_img"):
            return ""
        import hashlib
        ver = hashlib.md5(_marca_get("banner_proc_img").encode("utf-8")).hexdigest()[:8]
        url = (_marca_get("banner_proc_url") or "").strip()
        alt = _esc(_marca_get("banner_proc_alt") or "PROCSIS")
        img = "<img src='/media/banner-procsis?v=%s' alt='%s' loading='lazy'>" % (ver, alt)
        if url:
            ext = " target='_blank' rel='noopener noreferrer'" if url.startswith("http") else ""
            img = "<a href='%s'%s>%s</a>" % (_esc(url), ext, img)
        return ("<style>.pcb{max-width:1120px;margin:0 auto;padding:6px 20px 30px;box-sizing:border-box}"
                ".pcb a{display:block}.pcb img{width:100%;height:auto;display:block;border-radius:18px;"
                "box-shadow:0 12px 34px rgba(0,32,194,.16)}</style>"
                "<section class='pcb' id='banner-procsis'>" + img + "</section>")
    except Exception:
        return ""


def _marca_accion(f, files):
    """Guarda los logos y el banner. Devuelve (ok, mensaje)."""
    a = (f.get("accion") or "").strip()
    quien = session.get("usuario") or rol_actual() or ""
    if a in ("logo_procsis", "logo_empresa"):
        clave = "logo_ini_procsis" if a == "logo_procsis" else "logo_ini_empresa"
        nombre = "PROCSIS" if a == "logo_procsis" else "de la empresa"
        if f.get("quitar"):
            _ti_cfg_set(clave, "")
            _marca_invalidar()
            registrar_auditoria("Logos del inicio", "%s restableció el logo %s" % (quien, nombre))
            return True, "Listo: se volvió al logo original."
        v = _marca_img_de_archivo(files.get("archivo") if files else None, 1024)
        if v is None:
            return False, "El archivo no es una imagen válida (PNG, JPG, WebP o SVG) o pesa más de 1 MB."
        if not v:
            return False, "Elija el archivo de la imagen."
        _ti_cfg_set(clave, v)
        _marca_invalidar()
        registrar_auditoria("Logos del inicio", "%s cambió el logo %s" % (quien, nombre))
        return True, "Logo guardado. Si todavía ve el anterior, recargue con Ctrl+F5."
    if a == "banner":
        if f.get("quitar"):
            for k in ("banner_proc_img", "banner_proc_url", "banner_proc_alt"):
                _ti_cfg_set(k, "")
            _ti_cfg_set("banner_proc_on", "0")
            _marca_invalidar()
            registrar_auditoria("Banner PROCSIS", "%s quitó el banner" % quien)
            return True, "Banner quitado."
        url = (f.get("url") or "").strip()
        if url and (url.startswith("//") or not re.match(r"^(https://[^\s\"'<>]{3,400}|/[A-Za-z0-9_\-./?=&#%]{0,300})$", url)):
            return False, "El enlace debe empezar por https:// o por / (por ejemplo /ventas)."
        alt = " ".join((f.get("alt") or "").split())[:160]
        v = _marca_img_de_archivo(files.get("archivo") if files else None, 2048)
        if v is None:
            return False, "El archivo no es una imagen válida (PNG, JPG, WebP o SVG) o pesa más de 2 MB."
        if v:
            _ti_cfg_set("banner_proc_img", v)
        elif not _ti_cfg("banner_proc_img"):
            return False, "Suba primero la foto del banner."
        _ti_cfg_set("banner_proc_url", url)
        _ti_cfg_set("banner_proc_alt", alt)
        _ti_cfg_set("banner_proc_on", "1" if f.get("activo") else "0")
        _marca_invalidar()
        registrar_auditoria("Banner PROCSIS", "%s actualizó el banner" % quien)
        return True, "Banner guardado."
    return False, "Acción no reconocida."


@app.route("/gerencia/logos-inicio", methods=["GET", "POST"])
def gerencia_logos_inicio():
    _g = _guard_gerencia()
    if _g is not None:
        return _g
    _ti_init()
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Recargue la página e intente de nuevo.")
            return redirect("/gerencia/logos-inicio")
        try:
            ok, mensaje = _marca_accion(request.form, request.files)
        except Exception as e:
            db.session.rollback()
            print("logos inicio:", e)
            ok, mensaje = False, "No se pudo guardar el cambio. Intente de nuevo."
        session["ti_msg"] = ("ok" if ok else "err", mensaje)
        return redirect("/gerencia/logos-inicio")
    import time
    _marca_invalidar()
    sello = int(time.time())

    def tarjeta_logo(accion, titulo, texto, ruta, personalizado):
        estado = ("<span class='cb-tag on'>Logo propio activo</span>" if personalizado
                  else "<span class='cb-tag off'>Usando el logo original del sistema</span>")
        quitar = ("<label style='display:inline-flex;gap:6px;align-items:center;margin-top:10px'>"
                  "<input type='checkbox' name='quitar' value='1'> Volver al logo original</label>") if personalizado else ""
        return ("<div class='cb-card'><h2>%s</h2><p>%s</p><p>%s</p>"
                "<div style='display:flex;flex-wrap:wrap;gap:18px;align-items:center'>"
                "<div style='background:#f4f6ff;border:1px solid #e3e8f7;border-radius:14px;padding:14px;min-width:150px;text-align:center'>"
                "<img src='%s?v=%d' alt='' style='max-height:84px;max-width:220px;object-fit:contain'></div>"
                "<form method='POST' action='/gerencia/logos-inicio' enctype='multipart/form-data' style='flex:1;min-width:240px'>%s"
                "<label>Subir logo nuevo (PNG, JPG, WebP o SVG · máx. 1 MB)</label>"
                "<input type='file' name='archivo' accept='image/png,image/jpeg,image/webp,image/svg+xml'>%s"
                "<p style='margin-top:12px'><button class='cb-btn'>Guardar</button></p></form></div></div>"
                % (_esc(titulo), estado, _esc(texto), ruta, sello, _cob_hidden(csrf, accion, {}), quitar))

    t1 = tarjeta_logo("logo_procsis", "Logo de PROCSIS",
                      "La marca de la barra superior (la «P» a color junto a «PROCSIS | EduTrack»): inicio, acceso y paneles de Gerencia.",
                      "/static/img/logo-procsis-mark.png", bool(_marca_get("logo_ini_procsis")))
    t2 = tarjeta_logo("logo_empresa", "Logo de la empresa (EduTrack)",
                      "El logo principal que se ve en el inicio y en el carrusel del portal de acceso.",
                      "/static/img/logo-edutrack.png", bool(_marca_get("logo_ini_empresa")))
    tiene_b = bool(_marca_get("banner_proc_img"))
    prev_b = ("<img src='/media/banner-procsis?v=%d' alt='' style='width:100%%;max-width:640px;border-radius:14px;display:block;margin:6px 0 10px'>" % sello) if tiene_b else ""
    quitar_b = ("<label style='display:inline-flex;gap:6px;align-items:center;margin-top:10px'>"
                "<input type='checkbox' name='quitar' value='1'> Quitar el banner</label>") if tiene_b else ""
    t3 = ("<div class='cb-card'><h2>Banner con foto en la página de PROCSIS</h2>"
          "<p>Es un área con una foto ancha (tipo banner) que aparece justo debajo de «¿Quiénes somos?» en <a href='/procsis' target='_blank'>/procsis</a>. "
          "Medida recomendada: 1600 × 400 px.</p>%s"
          "<form method='POST' action='/gerencia/logos-inicio' enctype='multipart/form-data'>%s"
          "<label>Foto del banner (PNG, JPG, WebP o SVG · máx. 2 MB)</label>"
          "<input type='file' name='archivo' accept='image/png,image/jpeg,image/webp,image/svg+xml'>"
          "<label>Enlace al hacer clic (opcional, por ejemplo /ventas o https://…)</label><input type='text' name='url' value='%s' maxlength='400'>"
          "<label>Texto alternativo de la foto</label><input type='text' name='alt' value='%s' maxlength='160'>"
          "<label style='display:inline-flex;gap:6px;align-items:center;margin-top:12px'><input type='checkbox' name='activo' value='1'%s> Mostrar el banner en la página</label>"
          "<br>%s<p style='margin-top:12px'><button class='cb-btn'>Guardar banner</button></p></form></div>"
          % (prev_b, _cob_hidden(csrf, "banner", {}), _esc(_marca_get("banner_proc_url")), _esc(_marca_get("banner_proc_alt")),
             " checked" if _marca_get("banner_proc_on") == "1" else "", quitar_b))
    cuerpo = ("<h1>🖼️ Logos y banner del inicio</h1><p class='sub'>Cambie desde aquí los logos de la empresa y de PROCSIS que se ven en el inicio, "
              "y el área con foto de la página de PROCSIS. Los cambios pueden tardar hasta un minuto en verse en todo el sitio.</p>"
              + t1 + t2 + t3)
    return _cb_pagina("Logos y banner del inicio", cuerpo)


'''


with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()

if "gerencia_logos_inicio" in s:
    print("Este app.py YA tiene los cambios. No se modificó nada.")
    sys.exit(0)
if 'if __name__ == "__main__":' not in s:
    print("ERROR: tu app.py parece cortado (no termina con el arranque del servidor). No se modificó nada.")
    sys.exit(1)

NL = "\r\n" if "\r\n" in s else "\n"


def crlf(t):
    return t.replace("\r\n", "\n").replace("\n", NL)


def rep(old, new, count=1):
    global s
    n = s.count(old)
    if n != count:
        print("ERROR: no encontré en tu app.py este punto de inserción: %r (esperaba %d, hay %d). No se modificó nada." % (old[:80], count, n))
        sys.exit(1)
    s = s.replace(old, new)


# ---------------------------------------------------------------- 1) Header de Gerencia: una sola fila
rep(".user-profile-circle{background:#0020C2!important}",
    ".user-profile-circle{background:#0020C2!important}" + NL + crlf("""/* Barra superior de Gerencia/staff: una sola fila (antes se partía en dos y «Producción», el nombre y «Salir» quedaban abajo) */
.navbar-top-apple{height:auto!important;min-height:60px;gap:12px;flex-wrap:nowrap!important;padding:0 18px!important;overflow:hidden}
.navbar-top-apple .navbar-left{flex:1 1 auto;min-width:0;flex-wrap:nowrap!important;gap:12px}
.navbar-top-apple .nb-brand{flex:0 0 auto;white-space:nowrap}
.navbar-top-apple .navbar-module-title{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.navbar-top-apple .badge-env{flex:0 0 auto;white-space:nowrap}
.navbar-top-apple .navbar-actions-right{flex:0 0 auto;flex-wrap:nowrap!important;gap:8px;white-space:nowrap}
.navbar-top-apple .navbar-actions-right>*{flex:0 0 auto;white-space:nowrap}
.navbar-top-apple .nav-ntp-clock,.navbar-top-apple .nav-api-ping,.navbar-top-apple .badge-periodo{font-size:11.5px}
@media(max-width:1500px){.navbar-top-apple .nav-user-fullname{display:none}}
@media(max-width:1320px){.navbar-top-apple .nav-ntp-clock{display:none}}
@media(max-width:1120px){.navbar-top-apple .nav-api-ping,.navbar-top-apple .badge-periodo{display:none}}"""))

# ---------------------------------------------------------------- 2) Pie azul corporativo (estilo Tigo)
rep(".cf-foot{background:#1d1d1f;color:#a1a1a6;",
    ".cf-foot{background:linear-gradient(135deg,#0020C2,#000F66);color:rgba(255,255,255,.82);")
rep(".cf-foot h4{margin:0 0 14px;font-size:12px;font-weight:600;color:#f5f5f7;",
    ".cf-foot h4{margin:0 0 14px;font-size:12px;font-weight:600;color:#fff;")
rep(".cf-foot a{display:block;margin:0 0 9px;font-size:12px;font-weight:400;line-height:1.4;color:#a1a1a6;",
    ".cf-foot a{display:block;margin:0 0 9px;font-size:12px;font-weight:400;line-height:1.4;color:rgba(255,255,255,.84);")
rep(".cf-foot a:focus-visible{outline:2px solid #6e6e73;", ".cf-foot a:focus-visible{outline:2px solid #fff;")
rep(".cf-brand{font-weight:600;font-size:15px;color:#f5f5f7;", ".cf-brand{font-weight:600;font-size:15px;color:#fff;")
rep(".cf-foot p{margin:0 0 14px;font-size:12px;line-height:1.55;color:#a1a1a6}",
    ".cf-foot p{margin:0 0 14px;font-size:12px;line-height:1.55;color:rgba(255,255,255,.84)}")
rep("border-top:1px solid rgba(255,255,255,.1);display:flex;justify-content:space-between;align-items:center;gap:10px 16px;flex-wrap:wrap;font-size:11px;font-weight:300;color:#86868b}",
    "border-top:1px solid rgba(255,255,255,.28);display:flex;justify-content:space-between;align-items:center;gap:10px 16px;flex-wrap:wrap;font-size:11px;font-weight:300;color:rgba(255,255,255,.78)}")
rep(".cf-foot .cf-bot a{display:inline;margin:0 0 0 16px;font-size:11px;font-weight:300;color:#86868b}",
    ".cf-foot .cf-bot a{display:inline;margin:0 0 0 16px;font-size:11px;font-weight:300;color:rgba(255,255,255,.82)}")
# pie propio de la página /procsis (mismo azul)
rep(".pc-foot{{background:#1d1d1f;color:#a1a1a6;",
    ".pc-foot{{background:linear-gradient(135deg,#0020C2,#000F66);color:rgba(255,255,255,.84);")

# ---------------------------------------------------------------- 3) Cobertura: pueblos (municipios)
rep("\"activo INTEGER DEFAULT 1, logo TEXT DEFAULT '', orden INTEGER DEFAULT 0)\" % pk))",
    "\"activo INTEGER DEFAULT 1, logo TEXT DEFAULT '', orden INTEGER DEFAULT 0)\" % pk))" + NL + crlf(
    """        db.session.execute(text("CREATE TABLE IF NOT EXISTS cobertura_munis (id %s, depto_id INTEGER, nombre VARCHAR(120), "
                                "activo INTEGER DEFAULT 1, orden INTEGER DEFAULT 0)" % pk))"""))

bloque_cob = BLOQUE_COB
rep("def _cob_accion(f, files):", crlf(bloque_cob) + "def _cob_accion(f, files):")

acciones = crlf('''    if a.startswith("muni_"):
        try:
            dep = int(f.get("depto_id") or 0)
        except Exception:
            dep = 0
        if a == "muni_toggle":
            mid = id_valido()
            if not mid:
                return False, "Pueblo no válido."
            db.session.execute(text("UPDATE cobertura_munis SET activo = CASE WHEN activo=1 THEN 0 ELSE 1 END WHERE id=:i"), {"i": mid})
            db.session.commit()
            return True, "Estado del pueblo actualizado."
        if a == "muni_eliminar":
            mid = id_valido()
            if not mid:
                return False, "Pueblo no válido."
            db.session.execute(text("DELETE FROM cobertura_munis WHERE id=:i"), {"i": mid})
            db.session.commit()
            registrar_auditoria("Cobertura", "%s eliminó el pueblo #%d" % (quien, mid))
            return True, "Pueblo eliminado."
        if a == "muni_todos":
            if not dep:
                return False, "Departamento no válido."
            v = 1 if str(f.get("valor")) == "1" else 0
            db.session.execute(text("UPDATE cobertura_munis SET activo=:v WHERE depto_id=:d"), {"v": v, "d": dep})
            db.session.commit()
            registrar_auditoria("Cobertura", "%s %s todos los pueblos del departamento #%d" % (quien, "activó" if v else "desactivó", dep))
            return True, "Pueblos actualizados."
        if a == "muni_agregar":
            c = cod_valido()
            if not dep or not c or not db.session.execute(text("SELECT 1 FROM cobertura_deptos WHERE id=:i AND pais=:p"), {"i": dep, "p": c}).first():
                return False, "Departamento no válido."
            nombres = [" ".join(n.split())[:120] for n in re.split(r"[,;\\n]+", f.get("nombre") or "")]
            nombres = [n for n in nombres if n]
            if not nombres:
                return False, "Escriba al menos un pueblo."
            mx = int(db.session.execute(text("SELECT COALESCE(MAX(orden), 0) FROM cobertura_munis WHERE depto_id=:d"), {"d": dep}).scalar() or 0)
            agregados = 0
            for n in nombres:
                if db.session.execute(text("SELECT 1 FROM cobertura_munis WHERE depto_id=:d AND lower(nombre)=lower(:n)"), {"d": dep, "n": n}).first():
                    continue
                mx += 1
                db.session.execute(text("INSERT INTO cobertura_munis (depto_id, nombre, activo, orden) VALUES (:d, :n, 1, :o)"), {"d": dep, "n": n, "o": mx})
                agregados += 1
            db.session.commit()
            if not agregados:
                return False, "Esos pueblos ya estaban en la lista."
            registrar_auditoria("Cobertura", "%s agregó %d pueblo(s) al departamento #%d" % (quien, agregados, dep))
            return True, "%d pueblo(s) agregado(s)." % agregados
        if a == "muni_cargar":
            if cod_valido() != "CO":
                return False, "La carga automática de municipios solo está disponible para Colombia."
            n = _cob_cargar_munis(dep or None)
            if not n:
                return False, "No se cargó nada: ese departamento ya tiene pueblos o no se encontró el archivo colombia_divipola.py junto a app.py."
            registrar_auditoria("Cobertura", "%s cargó %d municipios oficiales" % (quien, n))
            return True, "Se cargaron %d municipios. Quedaron inactivos: marque en cuáles opera PROCSIS." % n
    return False, "Acción no reconocida."''')
rep('    return False, "Acción no reconocida."' + NL + NL + NL + "@app.route(\"/gerencia/cobertura\", methods=[\"GET\", \"POST\"])",
    acciones + NL + NL + NL + "@app.route(\"/gerencia/cobertura\", methods=[\"GET\", \"POST\"])")

# UI de la tabla de departamentos y tarjeta de pueblos
rep("def _cob_html_admin(pais_sel, csrf):", "def _cob_html_admin(pais_sel, csrf, depto_sel=0):")
rep("        for did, dn, da, dl in deps:",
    crlf("""        munis_cnt = {}
        for _r in _c48_q("SELECT depto_id, COUNT(*), SUM(CASE WHEN activo=1 THEN 1 ELSE 0 END) FROM cobertura_munis GROUP BY depto_id"):
            munis_cnt[_r[0]] = (int(_r[1] or 0), int(_r[2] or 0))
        for did, dn, da, dl in deps:"""))
rep('_cob_boton(csrf, "depto_eliminar", {"id": did, "pais": cod}, "Eliminar", "r", "¿Eliminar %s de la lista?" % dn),',
    '_cob_pueblos_link(cod, did, munis_cnt) + _cob_boton(csrf, "depto_eliminar", {"id": did, "pais": cod}, "Eliminar", "r", "¿Eliminar %s de la lista?" % dn),')
rep("    return tabla + detalle + agregar + (",
    crlf("""    pueblos = ""
    try:
        _did = int(depto_sel or 0)
    except Exception:
        _did = 0
    if sel and _did:
        pueblos = _cob_html_pueblos(sel[0], _did, csrf)
    return tabla + detalle + pueblos + agregar + ("""))
rep("_cob_html_admin(pais_sel, csrf))", "_cob_html_admin(pais_sel, csrf, request.args.get(\"depto\")))")
rep('(request.form.get("accion") or "").startswith("depto_") else "")',
    '(request.form.get("accion") or "").startswith("depto_") else "")' + NL + crlf(
    """        _acc_m = (request.form.get("accion") or "")
        _dep_m = str(request.form.get("depto_id") or "")
        if _acc_m.startswith("muni_") and re.match(r"^[A-Z]{2}$", pr):
            destino = ("/gerencia/cobertura?pais=%s&depto=%s#pueblos" % (pr, _dep_m)) if _dep_m.isdigit() else ("/gerencia/cobertura?pais=%s#zonas" % pr)"""))
rep("<h1>🌎 Cobertura · países y departamentos</h1>", "<h1>🌎 Cobertura · países, departamentos y pueblos</h1>")
rep("('/gerencia/cobertura', '🌎 Cobertura: países y departamentos'),",
    "('/gerencia/cobertura', '🌎 Cobertura: países, departamentos y pueblos')," + NL
    + "        ('/gerencia/logos-inicio', '🖼️ Logos y banner del inicio'),")

# API pública: pueblos activos por departamento + widget del selector
rep('"SELECT pais, nombre, logo FROM cobertura_deptos WHERE activo=1 ORDER BY orden, nombre"',
    '"SELECT pais, nombre, logo, id FROM cobertura_deptos WHERE activo=1 ORDER BY orden, nombre"')
rep("    for p, n, lg in deptos:",
    crlf("""    munis = {}
    try:
        for _dm, _nm in db.session.execute(text("SELECT depto_id, nombre FROM cobertura_munis WHERE activo=1 ORDER BY nombre")).fetchall():
            munis.setdefault(_dm, []).append(_nm)
    except Exception:
        db.session.rollback()
    for p, n, lg, did in deptos:"""))
rep('por.setdefault(p, []).append({"n": n, "l": lg or ""})',
    'por.setdefault(p, []).append({"n": n, "l": lg or "", "m": munis.get(did, [])})')
rep("r.appendChild(el('span','',d.n));G.appendChild(r);});",
    "r.appendChild(el('span','',d.n));if(d.m&&d.m.length){var sm=el('small','',d.m.length+' pueblos');"
    "sm.style.cssText='margin-left:auto;font-size:10.5px;color:#64748b';r.appendChild(sm);r.title=d.m.join(', ');}G.appendChild(r);});")

# ---------------------------------------------------------------- 4) Banner con foto en /procsis (debajo de «Quiénes somos»)
rep("    link_planes_foot = '<a href=\"/ventas\">Planes</a>' if _puede_planes else '<a href=\"/contacto\">Contacto comercial</a>'",
    "    link_planes_foot = '<a href=\"/ventas\">Planes</a>' if _puede_planes else '<a href=\"/contacto\">Contacto comercial</a>'" + NL
    + "    banner_html = _banner_procsis_html()")
rep('  <section class="pc-band">', "  {banner_html}" + NL + '  <section class="pc-band">')

# ---------------------------------------------------------------- 5) Logos y banner: bloque nuevo antes del arranque
bloque_final = BLOQUE_FINAL
rep('if __name__ == "__main__":', crlf(bloque_final) + 'if __name__ == "__main__":')

tmp = DST + ".nuevo"
with io.open(tmp, "w", encoding="utf-8", newline="") as fh:
    fh.write(s)
try:
    py_compile.compile(tmp, doraise=True)
except Exception as e:
    os.remove(tmp)
    print("ERROR: el resultado no compila, no se modificó tu app.py:", e)
    sys.exit(1)
shutil.copyfile(SRC, SRC + ".respaldo")
os.replace(tmp, DST)
print("LISTO. Líneas del nuevo app.py:", len(s.splitlines()))
print("Copia de seguridad del original en:", SRC + ".respaldo")

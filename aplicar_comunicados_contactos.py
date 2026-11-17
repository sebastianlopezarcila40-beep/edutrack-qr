# Comunicados (diseño de la foto) + Correos y líneas corporativas (Soporte y Ventas) en el panel principal.
# Uso:  python aplicar_comunicados_contactos.py            (trabaja sobre app.py de esta carpeta)
import io, os, re, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
DOCS = r'''# ════════════════════════════════════════════════════════════════════════════
#  DOCUMENTOS OFICIALES · Resoluciones (Gerencia / CRTP) y oficios (todas las áreas)
#  - Cada documento queda registrado con un código único y un QR que lleva a una página pública de verificación.
#  - Se descargan como PDF o imagen PNG, o se imprimen.
# ════════════════════════════════════════════════════════════════════════════
_DCX_INIT = {"ok": False}
_DCX_GERENCIA = ("Gerente", "Gerencia", "Superadmin", "Administrador")
_DCX_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
_DCX_COD = {"Gerencia": "GER", "CRTP": "CRTP", "Soporte": "SOP", "Comercial": "COM", "Cobranza": "COB", "Contabilidad": "CON",
            "Finanzas": "FIN", "Talento": "TAL", "Marketing": "MKT", "Juridica": "JUR", "Operaciones": "OPE", "Calidad": "CAL",
            "Tecnologia": "TEC", "Desarrollador": "DEV", "Soluciones": "SOL"}
_DCX_TITULOS = {"Gerencia": "Gerencia General", "CRTP": "Comisión de Regulación Tecnológica (CRTP)", "Comercial": "Área Comercial",
                "Soporte": "Soporte Técnico", "Cobranza": "Cobranza", "Contabilidad": "Contabilidad", "Desarrollador": "Desarrollo",
                "Soluciones": "Soluciones"}
_DCX_LOGO = "/static/img/logo-procsis-mark.png"


def _dcx_tipo_txt(t):
    return {"resolucion": "Resolución interna", "comunicado": "Comunicado"}.get(t, "Oficio")


def _dcx_md(t):
    """Texto seguro con **negrita**."""
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", _esc(t))


def _dcx_parrafos_md(t):
    partes = [p.strip() for p in re.split(r"\n\s*\n", (t or "").replace("\r", "")) if p.strip()]
    return "".join("<p>%s</p>" % _dcx_md(p).replace("\n", "<br>") for p in partes)


_DCX_ONDA = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 300 240'><defs>"
             "<linearGradient id='a' x1='0' y1='0' x2='1' y2='1'><stop offset='0' stop-color='#0B2A6B'/><stop offset='1' stop-color='#0A56E0'/></linearGradient>"
             "<linearGradient id='b' x1='0' y1='1' x2='1' y2='0'><stop offset='0' stop-color='#00B8D4'/><stop offset='.55' stop-color='#22C55E'/>"
             "<stop offset='1' stop-color='#86EFAC'/></linearGradient></defs>"
             "<path d='M0 0H250C165 30 65 90 0 170Z' fill='url(#a)'/>"
             "<path d='M0 176C66 96 168 34 262 4C196 46 84 124 0 214Z' fill='url(#b)'/></svg>")


def _dcx_onda_uri():
    from urllib.parse import quote
    return "data:image/svg+xml;utf8," + quote(_DCX_ONDA)


def _dcx_area_de_rol(rol):
    rol = (rol or "").strip()
    if rol in _DCX_GERENCIA:
        return "Gerencia"
    if rol in ("Ventas", "Supervisor de Ventas", "Comercial"):
        return "Comercial"
    if rol in ("Developer", "Desarrollador"):
        return "Desarrollador"
    return rol


def _dcx_area_titulo(area):
    if area in _DCX_TITULOS:
        return _DCX_TITULOS[area]
    try:
        return AREAS_PROCSIS[area].get("titulo", area)
    except Exception:
        return area


def _dcx_areas_validas():
    return ["Gerencia", "CRTP", "Soporte", "Comercial", "Cobranza", "Contabilidad"] + [a for a in AREAS_PROCSIS if a not in ("Gerencia",)]


def _dcx_rol_ok(rol):
    rol = (rol or "").strip()
    return bool(rol) and (rol in ROLES_AREAS or rol in _ROLES_MAPA or rol in _DCX_GERENCIA or rol == "Contabilidad")


def _dcx_init():
    if _DCX_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS docs_corp (id %s, codigo VARCHAR(48), token VARCHAR(40), tipo VARCHAR(20), area VARCHAR(60), "
            "emisor VARCHAR(220), numero INTEGER DEFAULT 0, anio INTEGER DEFAULT 0, asunto TEXT, fecha VARCHAR(10), ciudad VARCHAR(100), "
            "datos TEXT, firmante VARCHAR(160), cargo VARCHAR(160), firma TEXT DEFAULT '', estado VARCHAR(12) DEFAULT 'VIGENTE', "
            "motivo TEXT DEFAULT '', huella VARCHAR(64), creado_por VARCHAR(80), creado VARCHAR(19), "
            "anulado_por VARCHAR(80) DEFAULT '', anulado VARCHAR(19) DEFAULT '')" % pk))
        db.session.commit()
        _DCX_INIT["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("docs init:", e)


def _dcx_guard():
    """None si el usuario puede usar el módulo; si no, la respuesta a devolver."""
    if not requiere_login():
        return redirect("/login")
    if not _dcx_rol_ok(rol_actual()):
        return acceso_denegado("Tu rol no tiene acceso a los documentos oficiales.")
    try:
        from flask import g as _g_dcx
        _g_dcx._staff_wrap = True
    except Exception:
        pass
    _dcx_init()
    return None


def _dcx_fecha_larga(f):
    try:
        d = datetime.strptime(f, "%Y-%m-%d")
        return "%d de %s de %d" % (d.day, _DCX_MESES[d.month - 1], d.year)
    except Exception:
        return f or ""


def _dcx_fecha_corta(f):
    try:
        d = datetime.strptime(f, "%Y-%m-%d")
        return "%02d %s %d" % (d.day, _DCX_MESES[d.month - 1][:3].upper(), d.year)
    except Exception:
        return f or ""


def _dcx_parrafos(t):
    partes = [p.strip() for p in re.split(r"\n\s*\n", (t or "").replace("\r", "")) if p.strip()]
    return "".join("<p>%s</p>" % _esc(p).replace("\n", "<br>") for p in partes)


def _dcx_huella(d):
    import hashlib
    base = json.dumps({k: d.get(k) for k in ("codigo", "tipo", "area", "emisor", "numero", "anio", "asunto", "fecha", "ciudad",
                                              "datos", "firmante", "cargo")}, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(base.encode("utf-8")).hexdigest()


def _dcx_base():
    host = (request.host or "").strip()
    esquema = "http" if host.startswith(("localhost", "127.", "0.0.0.0")) else "https"
    return "%s://%s" % (esquema, host)


def _dcx_qr_uri(url):
    import base64 as _b64
    buf = BytesIO()
    qrcode.make(url, box_size=8, border=1).save(buf, format="PNG")
    return "data:image/png;base64," + _b64.b64encode(buf.getvalue()).decode("ascii")


def _dcx_img_de_archivo(f, max_kb):
    """Firma escaneada → data URI (solo PNG/JPG). '' si no hay archivo; None si no es válida o pesa de más."""
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
    else:
        return None
    import base64 as _b64
    return "data:%s;base64,%s" % (mime, _b64.b64encode(raw).decode("ascii"))


def _dcx_empresa():
    try:
        p = plataforma()
    except Exception:
        p = None

    def g(*nombres):
        for n in nombres:
            v = getattr(p, n, None) if p is not None else None
            if v:
                return str(v)
        return ""
    return {"email": g("email_empresa", "email_soporte") or "procsis.edu@gmail.com",
            "tel": g("telefono_empresa", "telefono_soporte") or "3246868183", "web": "procsishq.com"}


def _dcx_row(did=None, token=None):
    cols = ("id, codigo, token, tipo, area, emisor, numero, anio, asunto, fecha, ciudad, datos, firmante, cargo, firma, estado, "
            "motivo, huella, creado_por, creado, anulado_por, anulado")
    if did is not None:
        rows = _c48_q("SELECT %s FROM docs_corp WHERE id=:i" % cols, i=did)
    else:
        rows = _c48_q("SELECT %s FROM docs_corp WHERE token=:t" % cols, t=token)
    if not rows:
        return None
    d = dict(zip([c.strip() for c in cols.split(",")], rows[0]))
    try:
        d["datos_j"] = json.loads(d.get("datos") or "{}")
    except Exception:
        d["datos_j"] = {}
    return d


def _dcx_puede_ver(d, rol, usuario):
    return rol in _DCX_GERENCIA or d.get("area") == _dcx_area_de_rol(rol) or d.get("creado_por") == usuario


# ───────────────────────── Diseño de los documentos (HTML + CSS)
_DCX_CSS = r"""<style>
.dcx-lienzo{overflow:auto;background:#e8ecf6;padding:14px;border-radius:14px}
.dcx-escala{width:794px}
.dc-page{width:794px;min-height:1123px;box-sizing:border-box;background:#fff;position:relative;margin:0 auto;color:#111;box-shadow:0 10px 40px rgba(0,16,80,.18);overflow:hidden}
.dc-anulado{position:absolute;left:50%;top:44%;transform:translate(-50%,-50%) rotate(-24deg);font:900 104px Arial,sans-serif;color:rgba(200,30,30,.2);border:10px solid rgba(200,30,30,.2);padding:6px 30px;letter-spacing:.06em;pointer-events:none;z-index:5}
/* Resolución interna (diseño institucional en blanco y negro) */
.dc-res{padding:46px 58px 40px;font-family:Georgia,"Times New Roman",serif}
.dc-res .dc-frame{position:absolute;left:12px;top:12px;right:12px;bottom:12px;border:2.5px solid #111;pointer-events:none}
.dc-res .dc-c1,.dc-res .dc-c2{position:absolute;width:118px;height:118px;background:repeating-linear-gradient(135deg,#111 0,#111 15px,#fff 15px,#fff 24px)}
.dc-res .dc-c1{left:12px;top:12px;clip-path:polygon(0 0,100% 0,0 100%)}
.dc-res .dc-c2{right:12px;bottom:12px;clip-path:polygon(100% 100%,0 100%,100% 0)}
.dc-res .dc-logo{display:flex;align-items:center;gap:18px;margin:2px 0 12px;position:relative}
.dc-res .dc-ln{flex:1;height:1.5px;background:#111}
.dc-res .dc-lg{text-align:center}
.dc-res .dc-lg img{height:58px;display:block;margin:0 auto}
.dc-res .dc-tag{font:600 11px Arial,sans-serif;letter-spacing:.04em;color:#222;margin-top:3px}
.dc-res .dc-t1{text-align:center;font-weight:700;font-size:23px;letter-spacing:.02em;margin:10px 0 4px}
.dc-res .dc-t2{text-align:center;font-weight:700;font-size:14px;text-transform:uppercase;margin:0 0 12px;line-height:1.3}
.dc-res .dc-t3{text-align:center;font-weight:700;font-size:21px;margin:0 0 10px}
.dc-res .dc-t4{text-align:center;font-weight:700;font-size:25px;line-height:1.25;margin:0 0 4px;text-transform:uppercase}
.dc-res .dc-t5{text-align:center;font-weight:700;font-size:22px;margin:0 0 10px}
.dc-res .dc-obj{text-align:center;font-size:15px;line-height:1.4;margin:0 26px 12px}
.dc-res .dc-rule{height:2px;background:#111;position:relative;margin:14px 0}
.dc-res .dc-rule:after{content:"";position:absolute;left:50%;top:-4px;width:9px;height:9px;background:#111;transform:translateX(-50%) rotate(45deg)}
.dc-res p{margin:0 0 9px;font-size:15px;line-height:1.38;text-align:justify}
.dc-res .dc-hd{text-align:center;font-weight:700;font-size:21px;margin:12px 0 8px;letter-spacing:.01em}
.dc-res .dc-q{text-align:center;font-weight:700;font-size:17px;margin:8px 0 4px}
.dc-res .dc-fac{text-align:center}
.dc-res .dc-cierre{text-align:center;font-weight:700;font-size:17px;margin:14px 0 6px}
.dc-res .dc-dado{text-align:center;font-size:14.5px;margin:0 0 6px}
.dc-res .dc-final{display:flex;align-items:flex-end;justify-content:space-between;gap:10px;margin-top:6px}
.dc-res .dc-firma{flex:1;text-align:center}
.dc-res .dc-firma img{max-height:78px;max-width:300px;display:block;margin:0 auto -6px}
.dc-res .dc-fl{width:330px;max-width:100%;border-top:1.5px solid #111;margin:0 auto 6px}
.dc-res .dc-firma b{display:block;font-size:17px;letter-spacing:.02em;text-transform:uppercase}
.dc-res .dc-firma span{display:block;font-size:15.5px}
.dc-qr{display:flex;gap:10px;align-items:center;width:250px;font:10.5px/1.3 Arial,sans-serif;color:#222}
.dc-qr img{width:84px;height:84px;flex:0 0 auto}
.dc-qr i{font-style:normal;word-break:break-all;color:#555}
.dc-esp{width:250px}
/* Comunicado a la opinión pública (diseño corporativo con ondas azul y verde) */
.dc-com{padding:0 58px 130px;font-family:"Segoe UI",Arial,Helvetica,sans-serif;color:#0B2A6B;background:#fff}
.dc-com .dc-wv1,.dc-com .dc-wv2{position:absolute;width:300px;height:240px;pointer-events:none}
.dc-com .dc-wv1{left:0;top:0}
.dc-com .dc-wv2{right:0;bottom:24px;transform:rotate(180deg)}
.dc-com .dc-cbar{position:absolute;left:0;right:0;bottom:0;height:26px;background:linear-gradient(100deg,#0B2A6B,#0020C2,#0A56E0)}
.dc-com .dc-clogo{text-align:center;padding-top:36px;position:relative}
.dc-com .dc-clogo img{height:128px;display:block;margin:0 auto}
.dc-com .dc-ctit{text-align:center;font-weight:800;font-size:30px;line-height:1.2;color:#0B2A6B;margin:30px 0 0;padding-bottom:12px;border-bottom:2px solid #1a56c4;text-transform:uppercase;position:relative}
.dc-com .dc-cfecha{font-weight:800;font-size:16.5px;margin:22px 0 18px;position:relative}
.dc-com .dc-cuerpo{position:relative}
.dc-com .dc-cuerpo p{margin:0 0 14px;font-size:16.5px;line-height:1.55;text-align:justify;color:#0B2A6B}
.dc-com .dc-cuerpo b{font-weight:800}
.dc-com .dc-cfirma{display:flex;align-items:center;gap:18px;margin:26px 0 0;font-weight:800;font-size:16px;text-transform:uppercase;position:relative}
.dc-com .dc-cfirma span{flex:1;height:1.5px;background:#1a56c4}
.dc-com .dc-cfirma div{flex:0 0 auto}
.dc-com .dc-qr{position:absolute;left:58px;bottom:44px;color:#0B2A6B}
/* Oficio (carta corporativa con radicado y QR) */
.dc-of{padding:0 56px 130px;font-family:Arial,Helvetica,sans-serif;font-size:14.5px;line-height:1.5;color:#222}
.dc-of .dc-bar{height:8px;margin:0 -56px 34px;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2)}
.dc-of .dc-top{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;margin-bottom:34px}
.dc-of .dc-ologo img{height:64px;display:block}
.dc-of .dc-otag{font-size:11px;color:#0020C2;font-weight:700;letter-spacing:.03em;margin-top:5px}
.dc-of .dc-rad{display:flex;align-items:center;gap:10px;border:1px solid #444;padding:6px 8px;width:372px;box-sizing:border-box}
.dc-of .dc-radtxt{flex:1;font-size:8.6px;line-height:1.32;color:#222}
.dc-of .dc-rad img{width:76px;height:76px;flex:0 0 auto}
.dc-of .dc-fechaof{text-align:right;margin:0 0 26px}
.dc-of .dc-dest{margin:0 0 20px}
.dc-of .dc-asunto{text-align:center;margin:0 0 20px}
.dc-of p{margin:0 0 12px;text-align:justify}
.dc-of .dc-firma{margin-top:26px}
.dc-of .dc-firma img{max-height:70px;max-width:260px;display:block;margin:0 0 -4px}
.dc-of .dc-fl{width:260px;border-top:1.3px solid #333;margin:0 0 5px}
.dc-of .dc-firma b{display:block;text-transform:uppercase}
.dc-of .dc-firma span{display:block;color:#444}
.dc-of .dc-peq{font-size:11.5px;color:#444;margin:14px 0 0;text-align:left}
.dc-of .dc-foot{position:absolute;left:56px;right:56px;bottom:30px;display:flex;justify-content:space-between;gap:16px;font-size:9.5px;line-height:1.4;color:#0020C2;border-top:1px solid #d6dcf0;padding-top:8px}
.dc-of .dc-foot div{flex:1}
.dc-of .dc-foot div:last-child{text-align:right}
.dc-of .dc-foot b{display:block;color:#0020C2}
.dc-of .dc-foot span{color:#444}
@media print{
  body *{visibility:hidden!important}
  #dc-page,#dc-page *{visibility:visible!important}
  #dc-page{position:absolute!important;left:0;top:0;margin:0!important;box-shadow:none!important}
  #dc-escala{transform:none!important}
  @page{size:A4;margin:0}
}
.cb textarea{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:13px;font-family:inherit;min-height:84px;resize:vertical;color:#0b1437}
.cb input[type=date],.cb input[type=number]{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:13px}
.cb textarea:focus,.cb input[type=date]:focus{outline:none;border-color:#0020C2;box-shadow:0 0 0 3px rgba(0,32,194,.12)}
.dcx-fg{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:0 16px}
.dcx-art{border:1px dashed #cfd8f3;border-radius:12px;padding:10px 12px;margin:10px 0;background:#fafbff}
.dcx-bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 14px}
.dcx-st{font-size:12.5px;color:#64748b}
</style>"""


def _dcx_html_resolucion(d, qr, url):
    x = d.get("datos_j") or {}
    p = ['<div class="dc-page dc-res" id="dc-page"><div class="dc-frame"></div><div class="dc-c1"></div><div class="dc-c2"></div>']
    if d.get("estado") == "ANULADO":
        p.append('<div class="dc-anulado">ANULADO</div>')
    p.append('<div class="dc-logo"><span class="dc-ln"></span><div class="dc-lg"><img src="%s" alt="PROCSIS">'
             '<div class="dc-tag">— innovación que gestiona —</div></div><span class="dc-ln"></span></div>' % _DCX_LOGO)
    p.append('<div class="dc-t1">PROCSIS</div><div class="dc-t2">%s</div>' % _esc(d.get("emisor") or ""))
    if x.get("oficio_no"):
        p.append('<div class="dc-t3">OFICIO N° %s</div>' % _esc(x["oficio_no"]))
    p.append('<div class="dc-t4">RESOLUCIÓN INTERNA NÚMERO %s DE %s</div><div class="dc-t5">( %s )</div>'
             % (_esc(d.get("numero")), _esc(d.get("anio")), _esc(_dcx_fecha_corta(d.get("fecha")))))
    p.append('<div class="dc-obj">%s</div><div class="dc-rule"></div>' % _esc(d.get("asunto") or ""))
    if x.get("quien"):
        p.append('<div class="dc-q">%s</div>' % _esc(x["quien"]))
    if x.get("facultades"):
        p.append('<p class="dc-fac">%s</p>' % _esc(x["facultades"]).replace("\n", "<br>"))
    if x.get("considerandos"):
        p.append('<div class="dc-hd">CONSIDERANDO:</div>' + _dcx_parrafos(x["considerandos"]) + '<div class="dc-rule"></div>')
    p.append('<div class="dc-hd">%s</div>' % _esc(x.get("verbo") or "RESUELVEN:"))
    for i, a in enumerate(x.get("articulos") or [], 1):
        t = (a.get("titulo") or "").strip().upper().rstrip(".")
        p.append('<p><b>ARTÍCULO %dº.%s</b> %s</p>' % (i, (" " + _esc(t) + ".") if t else "", _esc(a.get("texto") or "").replace("\n", "<br>")))
    p.append('<div class="dc-cierre">%s</div>' % _esc(x.get("cierre") or "NOTIFÍQUESE Y CÚMPLASE."))
    try:
        dd = datetime.strptime(d.get("fecha") or "", "%Y-%m-%d")
        dado = "Dado en %s, a los %d días del mes de %s de %d." % (d.get("ciudad") or "", dd.day, _DCX_MESES[dd.month - 1], dd.year)
    except Exception:
        dado = "Dado en %s." % (d.get("ciudad") or "")
    p.append('<div class="dc-dado">%s</div>' % _esc(dado))
    firma = ('<img src="%s" alt="Firma">' % d["firma"]) if d.get("firma") else ""
    p.append('<div class="dc-final"><div class="dc-qr"><img src="%s" alt="QR"><div><b>Documento verificable</b><br>%s<br>'
             'Escanee el QR para validar su autenticidad.<br><i>%s</i></div></div>'
             '<div class="dc-firma">%s<div class="dc-fl"></div><b>%s</b><span>%s</span></div><div class="dc-esp"></div></div></div>'
             % (qr, _esc(d.get("codigo")), _esc(url), firma, _esc(d.get("firmante") or ""), _esc(d.get("cargo") or "")))
    return "".join(p)


def _dcx_html_comunicado(d, qr, url):
    x = d.get("datos_j") or {}
    onda = _dcx_onda_uri()
    p = ['<div class="dc-page dc-com" id="dc-page">']
    if d.get("estado") == "ANULADO":
        p.append('<div class="dc-anulado">ANULADO</div>')
    p.append('<img class="dc-wv1" src="%s" alt=""><img class="dc-wv2" src="%s" alt=""><div class="dc-cbar"></div>' % (onda, onda))
    p.append('<div class="dc-clogo"><img src="%s" alt="PROCSIS"></div>' % _DCX_LOGO)
    p.append('<div class="dc-ctit">%s</div>' % _esc(d.get("asunto") or ""))
    p.append('<div class="dc-cfecha">%s, %s.</div>' % (_esc(d.get("ciudad")), _esc(_dcx_fecha_larga(d.get("fecha")))))
    p.append('<div class="dc-cuerpo">%s</div>' % _dcx_parrafos_md(x.get("cuerpo") or ""))
    p.append('<div class="dc-cfirma"><span></span><div>%s</div><span></span></div>' % _esc(d.get("firmante") or ""))
    p.append('<div class="dc-qr"><img src="%s" alt="QR"><div><b>Documento verificable</b><br>%s<br>Escanee el QR para validar su autenticidad.</div></div></div>'
             % (qr, _esc(d.get("codigo"))))
    return "".join(p)


def _dcx_html_oficio(d, qr, url, did):
    x = d.get("datos_j") or {}
    emp = _dcx_empresa()
    p = ['<div class="dc-page dc-of" id="dc-page"><div class="dc-bar"></div>']
    if d.get("estado") == "ANULADO":
        p.append('<div class="dc-anulado">ANULADO</div>')
    dest_linea = " ".join(v for v in (x.get("dest_nombre"), x.get("dest_entidad")) if v)
    hora = (d.get("creado") or "")[:16].replace("T", " ")
    p.append('<div class="dc-top"><div class="dc-ologo"><img src="%s" alt="PROCSIS"><div class="dc-otag">innovación que gestiona</div></div>'
             '<div class="dc-rad"><div class="dc-radtxt">Al contestar cite Radicado <b>%s</b> Id: %s<br>Remitente: %s<br>Destinatario: %s<br>'
             'Fecha: %s<br>Folios: 1<br>Anexos: %s</div><img src="%s" alt="QR"></div></div>'
             % (_DCX_LOGO, _esc(d.get("codigo")), _esc(did), _esc(d.get("emisor")), _esc(dest_linea), _esc(hora or d.get("fecha")),
                _esc(x.get("anexos_n") or 0), qr))
    p.append('<div class="dc-fechaof">%s, %s</div>' % (_esc(d.get("ciudad")), _esc(_dcx_fecha_larga(d.get("fecha")))))
    dest = []
    for k, neg in (("dest_trato", False), ("dest_nombre", True), ("dest_cargo", False), ("dest_entidad", False), ("dest_ciudad", False)):
        if x.get(k):
            dest.append("<b>%s</b>" % _esc(x[k]).upper() if neg else _esc(x[k]))
    p.append('<div class="dc-dest">%s</div>' % "<br>".join(dest))
    p.append('<p class="dc-asunto"><b>Asunto:</b> %s</p>' % _esc(d.get("asunto") or ""))
    if x.get("saludo"):
        p.append("<p>%s</p>" % _esc(x["saludo"]))
    p.append(_dcx_parrafos(x.get("cuerpo") or ""))
    if x.get("despedida"):
        p.append("<p>%s</p>" % _esc(x["despedida"]))
    firma = ('<img src="%s" alt="Firma">' % d["firma"]) if d.get("firma") else ""
    p.append('<div class="dc-firma">%s<div class="dc-fl"></div><b>%s</b><span>%s</span><span>%s</span></div>'
             % (firma, _esc(d.get("firmante") or ""), _esc(d.get("cargo") or ""), _esc(_dcx_area_titulo(d.get("area")))))
    extra = []
    if x.get("anexos_txt"):
        extra.append("<b>Anexos:</b> %s" % _esc(x["anexos_txt"]))
    if x.get("copia"):
        extra.append("<b>Copia:</b> %s" % _esc(x["copia"]))
    if extra:
        p.append('<div class="dc-peq">%s</div>' % "<br>".join(extra))
    p.append('<div class="dc-foot"><div><b>Sede correspondencias</b><span>%s · PROCSIS</span><br><span>Tel: %s</span></div>'
             '<div><b>Servicio al ciudadano</b><span>%s</span><br><span>%s</span></div>'
             '<div><b>Página 1 de 1</b><span>Verifique este documento:</span><br><span>%s</span></div></div></div>'
             % (_esc(d.get("ciudad")), _esc(emp["tel"]), _esc(emp["email"]), _esc(emp["web"]), _esc(url)))
    return "".join(p)


# ───────────────────────── Visor (descarga PDF / imagen / impresión)
_DCX_VISOR_JS = r"""<script>
(function(){
var lienzo=document.getElementById("dcx-lienzo"),esc=document.getElementById("dc-escala"),pg=document.getElementById("dc-page");
function ajustar(){var w=lienzo.clientWidth-28,s=Math.min(1,w/794);esc.style.transformOrigin="top left";esc.style.transform="scale("+s+")";lienzo.style.height=Math.ceil(pg.offsetHeight*s)+28+"px";}
window.addEventListener("resize",ajustar);ajustar();
var H2C="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js",JSP="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
function cargar(src){return new Promise(function(ok,ko){var s=document.createElement("script");s.src=src;s.onload=ok;s.onerror=function(){ko(new Error("No se pudo cargar la librería de descarga"));};document.head.appendChild(s);});}
function msg(t){document.getElementById("dcx-st").textContent=t||"";}
async function captura(){
  if(!window.html2canvas)await cargar(H2C);
  if(document.fonts&&document.fonts.ready)await document.fonts.ready;
  return html2canvas(pg,{scale:2,useCORS:true,backgroundColor:"#ffffff",logging:false,onclone:function(d){var e=d.getElementById("dc-escala");if(e){e.style.transform="none";}}});
}
document.getElementById("dcx-pdf").onclick=async function(){
  msg("Generando PDF…");
  try{
    var c=await captura();if(!window.jspdf)await cargar(JSP);
    var pdf=new window.jspdf.jsPDF({orientation:"p",unit:"pt",format:"a4"});
    var W=pdf.internal.pageSize.getWidth(),H=pdf.internal.pageSize.getHeight();
    var h=c.height*W/c.width,img=c.toDataURL("image/jpeg",0.95),y=0;
    pdf.addImage(img,"JPEG",0,0,W,h);
    var resto=h-H;
    while(resto>1){pdf.addPage();y-=H;pdf.addImage(img,"JPEG",0,y,W,h);resto-=H;}
    pdf.save("__NOMBRE__.pdf");msg("");
  }catch(e){msg("No se pudo generar el PDF: "+e.message);}
};
document.getElementById("dcx-png").onclick=async function(){
  msg("Generando imagen…");
  try{
    var c=await captura();
    c.toBlob(function(b){var a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="__NOMBRE__.png";document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},600);msg("");},"image/png");
  }catch(e){msg("No se pudo generar la imagen: "+e.message);}
};
document.getElementById("dcx-print").onclick=function(){window.print();};
document.getElementById("dcx-copiar").onclick=function(){
  var u="__URL__";
  if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(u).then(function(){msg("Enlace de verificación copiado.");});}
  else{window.prompt("Copia el enlace:",u);}
};
})();
</script>"""


def _dcx_vis(rol, usuario):
    if rol in _DCX_GERENCIA:
        return "1=1", {}
    return "(area=:ar OR creado_por=:us)", {"ar": _dcx_area_de_rol(rol), "us": usuario}


def _dcx_pagina_lista(titulo, sub, area_fijo, ruta, nuevos):
    rol = rol_actual()
    usuario = session.get("usuario") or ""
    vis, p = _dcx_vis(rol, usuario)
    cond = [vis]
    tipo = (request.args.get("tipo") or "").strip()
    q = (request.args.get("q") or "").strip()[:60]
    if area_fijo:
        cond.append("area=:af")
        p["af"] = area_fijo
    if tipo in ("oficio", "resolucion", "comunicado"):
        cond.append("tipo=:tp")
        p["tp"] = tipo
    if q:
        cond.append("(lower(codigo) LIKE :qq OR lower(asunto) LIKE :qq OR lower(emisor) LIKE :qq)")
        p["qq"] = "%" + q.lower() + "%"
    rows = _c48_q("SELECT id, codigo, tipo, asunto, area, fecha, estado FROM docs_corp WHERE " + " AND ".join(cond)
                  + " ORDER BY id DESC LIMIT 200", **p)
    filas = ""
    for did, cod, tp, asu, ar, fe, es in rows:
        etq = "<span class='cb-tag %s'>%s</span>" % ("on" if es == "VIGENTE" else "off", "Vigente" if es == "VIGENTE" else "Anulado")
        filas += ("<tr><td><b>%s</b></td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
                  "<td><a class='cb-btn s' href='/documentos-oficiales/ver/%d'>Ver / descargar</a></td></tr>"
                  % (_esc(cod), _dcx_tipo_txt(tp), _esc((asu or "")[:90]), _esc(_dcx_area_titulo(ar)),
                     _esc(fe), etq, did))
    if not filas:
        filas = "<tr><td colspan='7'>Aún no hay documentos aquí. Crea el primero con los botones de arriba.</td></tr>"
    botones = "".join("<a class='cb-btn g' href='%s'>%s</a> " % (_esc(h), _esc(t)) for t, h in nuevos)
    tabs = "".join("<a class='cb-btn s %s' href='%s?tipo=%s'>%s</a> " % ("" if tipo == k else "o", ruta, k, n)
                   for k, n in (("", "Todos"), ("oficio", "Oficios"), ("resolucion", "Resoluciones"), ("comunicado", "Comunicados")))
    buscar = ("<form method='GET' action='%s' style='display:inline-flex;gap:6px;margin-left:8px'><input type='text' name='q' value='%s' "
              "placeholder='Buscar código o asunto' style='width:210px'><button class='cb-btn s'>Buscar</button></form>" % (ruta, _esc(q)))
    cuerpo = ("<h1>%s</h1><p class='sub'>%s</p><div class='dcx-bar'>%s</div>"
              "<div class='cb-card'><div class='dcx-bar'>%s%s</div><div class='cb-scroll'><table><tr><th>Código</th><th>Tipo</th><th>Asunto</th>"
              "<th>Área</th><th>Fecha</th><th>Estado</th><th></th></tr>%s</table></div></div>"
              % (_esc(titulo), _esc(sub), botones, tabs, buscar, filas))
    return _cb_pagina(titulo, _DCX_CSS + cuerpo)


@app.route("/documentos-oficiales")
def dcx_lista():
    g = _dcx_guard()
    if g is not None:
        return g
    nuevos = [("➕ Nuevo oficio", "/documentos-oficiales/nuevo/oficio")]
    if rol_actual() in _DCX_GERENCIA:
        nuevos.append(("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"))
        nuevos.append(("➕ Nuevo comunicado", "/documentos-oficiales/nuevo/comunicado?area=Gerencia"))
    return _dcx_pagina_lista("📄 Documentos oficiales", "Oficios y resoluciones de PROCSIS con código único y QR de verificación. Se pueden descargar en PDF o imagen.",
                             None, "/documentos-oficiales", nuevos)


@app.route("/gerencia/comunicaciones")
def dcx_comunicaciones():
    g = _guard_gerencia()
    if g is not None:
        return g
    _dcx_init()
    return _dcx_pagina_lista("📜 Comunicaciones de Gerencia", "Comunicados públicos, resoluciones internas y oficios emitidos por la Gerencia General.", "Gerencia",
                             "/gerencia/comunicaciones", [("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"),
                                                          ("➕ Nuevo oficio", "/documentos-oficiales/nuevo/oficio?area=Gerencia"),
                                                          ("➕ Nuevo comunicado", "/documentos-oficiales/nuevo/comunicado?area=Gerencia")])


@app.route("/gerencia/crtp")
def dcx_crtp():
    g = _guard_gerencia()
    if g is not None:
        return g
    _dcx_init()
    return _dcx_pagina_lista("⚖️ Comisión de Regulación Tecnológica (CRTP)",
                             "Resoluciones internas y oficios de la CRTP: directrices de avance tecnológico, estándares operativos y regulación interna.",
                             "CRTP", "/gerencia/crtp", [("➕ Nueva resolución CRTP", "/documentos-oficiales/nuevo/resolucion?area=CRTP"),
                                                         ("➕ Nuevo oficio CRTP", "/documentos-oficiales/nuevo/oficio?area=CRTP"),
                                                         ("➕ Nuevo comunicado CRTP", "/documentos-oficiales/nuevo/comunicado?area=CRTP")])


# ───────────────────────── Formularios
def _dcx_campo(etq, nombre, valor="", tipo="text", extra="", ayuda=""):
    return ("<div><label>%s</label><input type='%s' name='%s' value='%s' %s>%s</div>"
            % (_esc(etq), tipo, nombre, _esc(valor), extra, ("<div class='cb-note'>%s</div>" % _esc(ayuda)) if ayuda else ""))


def _dcx_area_txt(etq, nombre, valor="", filas=4, ayuda=""):
    return ("<div style='grid-column:1/-1'><label>%s</label><textarea name='%s' rows='%d'>%s</textarea>%s</div>"
            % (_esc(etq), nombre, filas, _esc(valor), ("<div class='cb-note'>%s</div>" % _esc(ayuda)) if ayuda else ""))


_DCX_RES_DEF = {
    "objeto": "Por la cual se dictan las directrices de avance tecnológico, estándares operativos y regulación interna para el equipo y las plataformas de PROCSIS.",
    "quien": "EL GERENTE GENERAL Y EL EQUIPO DE TRABAJO DE PROCSIS,",
    "facultades": ("en uso de sus facultades corporativas, estatutarias y de gobernanza privada, en especial las conferidas por la estructura "
                   "organizacional de la empresa y la Comisión de Regulación Tecnológica (CRTP), y"),
    "considerandos": ("Que como empresa privada, el equipo de trabajo de PROCSIS y la Gerencia General trabajan día a día bajo la convicción de que el "
                      "crecimiento y el progreso organizacional se demuestran con hechos, innovación y resultados de calidad.\n\n"
                      "Que es necesario fijar un marco normativo interno, reglas claras y estándares elevados de operación que garanticen la seguridad "
                      "de la información, la transparencia y la eficiencia operativa en todos nuestros desarrollos y proyectos.\n\n"
                      "Que la Comisión de Regulación Tecnológica (CRTP), coordinada por la Gerencia General y respaldada por nuestro equipo de trabajo, "
                      "emite la presente directiva para guiar el funcionamiento técnico y operativo de toda la organización."),
    "arts": [("ÁMBITO DE APLICACIÓN Y ESTÁNDARES", "Establecer con carácter obligatorio e interno la plena vigencia de las directrices operativas, regulatorias "
                                                  "y de protección de datos que rigen para todas las herramientas y software desarrollados por PROCSIS."),
             ("COMPROMISO OPERATIVO INTERNO", "Reafirmar que la Gerencia General y nuestro equipo de trabajo mantendrán la inversión continua en infraestructura, "
                                              "desarrollo tecnológico y talento humano para asegurar la excelencia en los procesos internos."),
             ("VIGENCIA", "La presente resolución interna rige a partir de la fecha de su firma y notificación al equipo corporativo.")],
}


_DCX_COM_DEF = (
    "**PROCSIS** informa a sus usuarios, aliados estratégicos y a la opinión pública en general que, en el marco de nuestro proceso de evolución, "
    "actualización de identidad corporativa y consolidación en el mercado, nos mantenemos firmes en nuestra misión de ofrecer tecnología de alto impacto.\n\n"
    "La compañía agradece el respaldo constante de nuestra comunidad, reafirmando que la innovación, el trabajo constante y el rigor en el servicio son, "
    "y seguirán siendo, nuestra máxima prioridad.\n\n"
    "Ante recientes comentarios, rumores e imprecisiones difundidas por terceros y actores del sector que intentan desviar la atención, **PROCSIS** expresa su "
    "total postura de seriedad. Ningún tipo de especulación, estrategia de desinformación o intento de desacreditación por parte de otras compañías nos "
    "apartará de nuestro rumbo ni del compromiso con el desarrollo tecnológico que lideramos.\n\n"
    "Construir soluciones sólidas como **EduTrack** y actuar con rectitud es parte esencial de nuestra identidad. Las empresas serias responden con hechos, "
    "innovación y resultados, no con controversias ajenas.\n\n"
    "**PROCSIS** reafirma su confianza en el valor de su trabajo y en su determinación de continuar cumpliendo con total responsabilidad a todos nuestros usuarios.")


def _dcx_form_com(area, csrf):
    opts = "".join("<option value='%s'%s>%s</option>" % (_esc(a), " selected" if a == area else "", _esc(_dcx_area_titulo(a))) for a in ("Gerencia", "CRTP"))
    campos = ("<div><label>Área que emite</label><select name='area'>%s</select></div>" % opts
              + _dcx_campo("Título del comunicado", "asunto", "COMUNICADO A LA OPINIÓN PÚBLICA")
              + _dcx_campo("Ciudad", "ciudad", "Caracolí, Antioquia")
              + _dcx_campo("Fecha", "fecha", fecha_hoy(), "date")
              + _dcx_area_txt("Texto del comunicado", "cuerpo", _DCX_COM_DEF, 16, "Separa los párrafos con una línea en blanco. Usa **texto** para resaltar en negrita.")
              + _dcx_campo("Firma (entidad que firma)", "firmante", "GERENCIA GENERAL – PROCSIS"))
    return ("<h1>Nuevo comunicado</h1><p class='sub'>Comunicado público con el diseño corporativo de PROCSIS. Al emitirlo recibe un código único y un QR "
            "de verificación, y se puede descargar en PDF o imagen. Un comunicado emitido no se edita: si hay un error, se anula y se emite otro.</p>"
            "<form method='POST' action='/documentos-oficiales/nuevo/comunicado'>"
            "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='area_def' value='%s'>"
            "<div class='cb-card'><div class='dcx-fg'>%s</div>"
            "<p style='margin-top:14px'><button class='cb-btn g'>Emitir comunicado</button> "
            "<a class='cb-btn o' href='/documentos-oficiales'>Cancelar</a></p></div></form>" % (_esc(csrf), _esc(area), campos))


def _dcx_form(tipo, area, csrf):
    if tipo == "comunicado":
        return _dcx_form_com(area, csrf)
    rol = rol_actual()
    es_ger = rol in _DCX_GERENCIA
    hoy = fecha_hoy()
    if es_ger:
        opts = "".join("<option value='%s'%s>%s</option>" % (_esc(a), " selected" if a == area else "", _esc(_dcx_area_titulo(a)))
                       for a in (["Gerencia", "CRTP"] if tipo == "resolucion" else _dcx_areas_validas()))
        sel_area = "<div><label>Área que emite</label><select name='area'>%s</select></div>" % opts
    else:
        sel_area = "<div><label>Área que emite</label><input type='text' value='%s' disabled></div>" % _esc(_dcx_area_titulo(area))
    firm_def = "Gerente General – PROCSIS" if es_ger and area in ("Gerencia", "CRTP") else ""
    if tipo == "resolucion":
        emisor_def = ("GERENCIA GENERAL Y COMISIÓN DE REGULACIÓN TECNOLÓGICA (CRTP)" if area == "CRTP" else "GERENCIA GENERAL")
        arts = "".join(
            "<div class='dcx-art'><label>Título del artículo</label><input type='text' name='art_titulo' value='%s'>"
            "<label>Texto</label><textarea name='art_texto' rows='3'>%s</textarea>"
            "<button type='button' class='cb-btn r s' onclick='this.parentNode.remove()'>Quitar artículo</button></div>" % (_esc(t), _esc(x))
            for t, x in _DCX_RES_DEF["arts"])
        campos = (sel_area
                  + _dcx_campo("Quién emite (encabezado)", "emisor", emisor_def)
                  + _dcx_campo("Número de la resolución", "numero", "", "number", "min='1'", "Vacío = consecutivo automático del año.")
                  + _dcx_campo("Oficio N° (opcional)", "oficio_no", "")
                  + _dcx_campo("Ciudad", "ciudad", "Caracolí, Antioquia")
                  + _dcx_campo("Fecha", "fecha", hoy, "date")
                  + _dcx_area_txt("Objeto («Por la cual se…»)", "asunto", _DCX_RES_DEF["objeto"], 3)
                  + _dcx_campo("Quien resuelve", "quien", _DCX_RES_DEF["quien"])
                  + _dcx_area_txt("Facultades", "facultades", _DCX_RES_DEF["facultades"], 3)
                  + _dcx_area_txt("Considerandos", "considerandos", _DCX_RES_DEF["considerandos"], 9, "Separa cada «Que…» con una línea en blanco.")
                  + _dcx_campo("Encabezado de los artículos", "verbo", "RESUELVEN:")
                  + "<div style='grid-column:1/-1'><label>Artículos</label><div id='dcx-arts'>%s</div>"
                    "<button type='button' class='cb-btn o s' onclick='dcxAddArt()'>➕ Agregar artículo</button></div>" % arts
                  + _dcx_campo("Cierre", "cierre", "NOTIFÍQUESE Y CÚMPLASE.")
                  + _dcx_campo("Firmante (nombre)", "firmante", "")
                  + _dcx_campo("Cargo del firmante", "cargo", firm_def)
                  + "<div><label>Firma escaneada (PNG o JPG, opcional)</label><input type='file' name='firma' accept='image/png,image/jpeg'></div>")
        titulo = "Nueva resolución interna"
    else:
        campos = (sel_area
                  + _dcx_campo("Ciudad", "ciudad", "Caracolí, Antioquia")
                  + _dcx_campo("Fecha", "fecha", hoy, "date")
                  + _dcx_campo("Tratamiento del destinatario", "dest_trato", "Señor(a)", "text", "", "Ej.: Doctora, Señores, Excelentísimo Doctor.")
                  + _dcx_campo("Nombre del destinatario", "dest_nombre", "")
                  + _dcx_campo("Cargo del destinatario", "dest_cargo", "")
                  + _dcx_campo("Entidad o empresa", "dest_entidad", "")
                  + _dcx_campo("Ciudad / dirección del destinatario", "dest_ciudad", "")
                  + _dcx_campo("Asunto", "asunto", "")
                  + _dcx_campo("Saludo", "saludo", "Cordial saludo:")
                  + _dcx_area_txt("Cuerpo del oficio", "cuerpo", "", 10, "Separa los párrafos con una línea en blanco.")
                  + _dcx_campo("Despedida", "despedida", "Atentamente,")
                  + _dcx_campo("Firmante (nombre)", "firmante", session.get("nombre") or "")
                  + _dcx_campo("Cargo del firmante", "cargo", firm_def)
                  + _dcx_campo("Número de anexos", "anexos_n", "0", "number", "min='0' max='99'")
                  + _dcx_campo("Descripción de anexos (opcional)", "anexos_txt", "")
                  + _dcx_campo("Con copia a (opcional)", "copia", "")
                  + "<div><label>Firma escaneada (PNG o JPG, opcional)</label><input type='file' name='firma' accept='image/png,image/jpeg'></div>")
        titulo = "Nuevo oficio"
    js = ("<script>function dcxAddArt(){var d=document.createElement('div');d.className='dcx-art';"
          "d.innerHTML=\"<label>Título del artículo</label><input type='text' name='art_titulo'><label>Texto</label>"
          "<textarea name='art_texto' rows='3'></textarea><button type='button' class='cb-btn r s' onclick='this.parentNode.remove()'>Quitar artículo</button>\";"
          "document.getElementById('dcx-arts').appendChild(d);}</script>") if tipo == "resolucion" else ""
    return ("<h1>%s</h1><p class='sub'>Al guardar se asigna un código único y un QR de verificación. Después podrás descargarlo en PDF o imagen. "
            "Un documento emitido no se edita: si hay un error, se anula y se emite uno nuevo.</p>"
            "<form method='POST' enctype='multipart/form-data' action='/documentos-oficiales/nuevo/%s'>"
            "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='area_def' value='%s'>"
            "<div class='cb-card'><div class='dcx-fg'>%s</div>"
            "<p style='margin-top:14px'><button class='cb-btn g'>Emitir documento</button> "
            "<a class='cb-btn o' href='/documentos-oficiales'>Cancelar</a></p></div></form>%s"
            % (titulo, tipo, _esc(csrf), _esc(area), campos, js))


def _dcx_crear(tipo, f, files, usuario, rol):
    """Crea el documento. Devuelve (ok, mensaje_o_id)."""
    es_ger = rol in _DCX_GERENCIA
    if tipo in ("resolucion", "comunicado") and not es_ger:
        return False, "Las resoluciones y comunicados solo los emite la Gerencia."
    area = _dcx_area_de_rol(rol)
    if es_ger:
        a = (f.get("area") or f.get("area_def") or "Gerencia").strip()
        area = a if a in _dcx_areas_validas() else "Gerencia"
        if tipo in ("resolucion", "comunicado") and area not in ("Gerencia", "CRTP"):
            area = "Gerencia"
    fecha = (f.get("fecha") or "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", fecha):
        fecha = fecha_hoy()
    anio = int(fecha[:4])
    ciudad = " ".join((f.get("ciudad") or "").split())[:100]
    asunto = " ".join((f.get("asunto") or "").split())[:600]
    firmante = " ".join((f.get("firmante") or "").split())[:160]
    cargo = " ".join((f.get("cargo") or "").split())[:160]
    if not asunto:
        return False, "Escribe el asunto (o el objeto de la resolución)."
    if not firmante:
        return False, "Escribe el nombre de quien firma."
    firma = _dcx_img_de_archivo(files.get("firma") if files else None, 400)
    if firma is None:
        return False, "La firma debe ser una imagen PNG o JPG de máximo 400 KB."

    def lim(k, n):
        return " ".join((f.get(k) or "").split())[:n]

    def lim_p(k, n):
        return (f.get(k) or "").replace("\r", "").strip()[:n]
    if tipo == "resolucion":
        arts = []
        for t, x in zip(f.getlist("art_titulo"), f.getlist("art_texto")):
            x = (x or "").replace("\r", "").strip()[:2500]
            if x:
                arts.append({"titulo": " ".join((t or "").split())[:160], "texto": x})
        if not arts:
            return False, "Agrega al menos un artículo con texto."
        emisor = lim("emisor", 220) or "GERENCIA GENERAL"
        datos = {"oficio_no": lim("oficio_no", 40), "quien": lim("quien", 220), "facultades": lim_p("facultades", 1200),
                 "considerandos": lim_p("considerandos", 6000), "verbo": lim("verbo", 40) or "RESUELVEN:", "articulos": arts,
                 "cierre": lim("cierre", 80)}
        pref = "RES"
    elif tipo == "comunicado":
        emisor = "COMISIÓN DE REGULACIÓN TECNOLÓGICA (CRTP)" if area == "CRTP" else "GERENCIA GENERAL"
        if not lim_p("cuerpo", 9000):
            return False, "Escribe el texto del comunicado."
        datos = {"cuerpo": lim_p("cuerpo", 9000)}
        pref = "CMD"
    else:
        emisor = _dcx_area_titulo(area)
        if not lim_p("cuerpo", 8000):
            return False, "Escribe el cuerpo del oficio."
        try:
            n_anex = max(0, min(99, int(f.get("anexos_n") or 0)))
        except Exception:
            n_anex = 0
        datos = {k: lim(k, 200) for k in ("dest_trato", "dest_nombre", "dest_cargo", "dest_entidad", "dest_ciudad", "saludo", "despedida",
                                         "anexos_txt", "copia")}
        datos["cuerpo"] = lim_p("cuerpo", 8000)
        datos["anexos_n"] = n_anex
        pref = "OF"
    num = 0
    if tipo == "resolucion" and (f.get("numero") or "").strip().isdigit():
        num = int(f.get("numero"))
    if num:
        if _c48_q("SELECT 1 FROM docs_corp WHERE tipo=:t AND area=:a AND anio=:y AND numero=:n", t=tipo, a=area, y=anio, n=num):
            return False, "Ya existe una resolución con ese número en el año %d." % anio
    else:
        r = _c48_q("SELECT COALESCE(MAX(numero), 0) FROM docs_corp WHERE tipo=:t AND area=:a AND anio=:y", t=tipo, a=area, y=anio)
        num = int((r[0][0] if r else 0) or 0) + 1
    codigo = "PRC-%s-%s-%d-%04d" % (pref, _DCX_COD.get(area, "GEN"), anio, num)
    token = secrets.token_urlsafe(12)
    reg = {"codigo": codigo, "tipo": tipo, "area": area, "emisor": emisor, "numero": num, "anio": anio, "asunto": asunto, "fecha": fecha,
           "ciudad": ciudad, "datos": json.dumps(datos, ensure_ascii=False), "firmante": firmante, "cargo": cargo}
    huella = _dcx_huella(reg)
    _c48_x("INSERT INTO docs_corp (codigo, token, tipo, area, emisor, numero, anio, asunto, fecha, ciudad, datos, firmante, cargo, firma, "
           "estado, huella, creado_por, creado) VALUES (:codigo, :token, :tipo, :area, :emisor, :numero, :anio, :asunto, :fecha, :ciudad, "
           ":datos, :firmante, :cargo, :firma, 'VIGENTE', :huella, :us, :cr)",
           token=token, firma=firma or "", huella=huella, us=usuario, cr=str(ahora())[:19], **reg)
    r = _c48_q("SELECT id FROM docs_corp WHERE token=:t", t=token)
    if not r:
        return False, "No se pudo guardar el documento. Intenta de nuevo."
    registrar_auditoria("Documentos oficiales", "%s emitió %s" % (usuario, codigo))
    return True, int(r[0][0])


@app.route("/documentos-oficiales/nuevo/<tipo>", methods=["GET", "POST"])
def dcx_nuevo(tipo):
    g = _dcx_guard()
    if g is not None:
        return g
    if tipo not in ("oficio", "resolucion", "comunicado"):
        return redirect("/documentos-oficiales")
    rol = rol_actual()
    if tipo in ("resolucion", "comunicado") and rol not in _DCX_GERENCIA:
        return acceso_denegado("Las resoluciones y comunicados solo los emite la Gerencia.")
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/documentos-oficiales/nuevo/" + tipo)
        try:
            ok, res = _dcx_crear(tipo, request.form, request.files, session.get("usuario") or "", rol)
        except Exception as e:
            db.session.rollback()
            print("docs crear:", e)
            ok, res = False, "No se pudo emitir el documento. Intenta de nuevo."
        if ok:
            session["ti_msg"] = ("ok", "Documento emitido. Ya puedes descargarlo.")
            return redirect("/documentos-oficiales/ver/%d" % res)
        session["ti_msg"] = ("err", res)
        return redirect("/documentos-oficiales/nuevo/" + tipo)
    area = _dcx_area_de_rol(rol)
    if rol in _DCX_GERENCIA:
        a = (request.args.get("area") or "Gerencia").strip()
        area = a if a in _dcx_areas_validas() else "Gerencia"
    return _cb_pagina("Nuevo documento", _DCX_CSS + _dcx_form(tipo, area, csrf))


@app.route("/documentos-oficiales/ver/<int:did>")
def dcx_ver(did):
    g = _dcx_guard()
    if g is not None:
        return g
    d = _dcx_row(did=did)
    rol, usuario = rol_actual(), session.get("usuario") or ""
    if not d or not _dcx_puede_ver(d, rol, usuario):
        return acceso_denegado("Documento no disponible.")
    url = "%s/docs/verificar/%s" % (_dcx_base(), d["token"])
    qr = _dcx_qr_uri(url)
    if d["tipo"] == "resolucion":
        doc = _dcx_html_resolucion(d, qr, url)
    elif d["tipo"] == "comunicado":
        doc = _dcx_html_comunicado(d, qr, url)
    else:
        doc = _dcx_html_oficio(d, qr, url, did)
    nombre = re.sub(r"[^A-Za-z0-9_-]", "_", d["codigo"] or "documento")
    js = _DCX_VISOR_JS.replace("__NOMBRE__", nombre).replace("__URL__", url.replace('"', ""))
    anular = ""
    if d["estado"] == "VIGENTE" and (rol in _DCX_GERENCIA or d.get("creado_por") == usuario):
        anular = ("<form method='POST' action='/documentos-oficiales/anular' style='display:inline-flex;gap:6px' "
                  "onsubmit=\"return confirm('¿Anular este documento? Dejará de ser válido al escanear su QR.');\">"
                  "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='id' value='%d'>"
                  "<input type='text' name='motivo' placeholder='Motivo de la anulación' required maxlength='200' style='width:230px'>"
                  "<button class='cb-btn r s'>Anular</button></form>" % (_esc(_acc_csrf()), did))
    estado = ("<span class='cb-tag on'>Vigente</span>" if d["estado"] == "VIGENTE"
              else "<span class='cb-tag off'>Anulado · %s</span>" % _esc(d.get("motivo") or ""))
    barra = ("<div class='dcx-bar'><button id='dcx-pdf' class='cb-btn'>⬇ Descargar PDF</button>"
             "<button id='dcx-png' class='cb-btn o'>🖼 Descargar imagen</button><button id='dcx-print' class='cb-btn o'>🖨 Imprimir</button>"
             "<button id='dcx-copiar' class='cb-btn o'>🔗 Copiar enlace de verificación</button>"
             "<a class='cb-btn o' href='/documentos-oficiales'>← Volver</a> %s %s<span id='dcx-st' class='dcx-st'></span></div>" % (estado, anular))
    cuerpo = ("<h1>%s</h1><p class='sub'>%s · %s</p>%s<div class='dcx-lienzo' id='dcx-lienzo'><div class='dcx-escala' id='dc-escala'>%s</div></div>%s"
              % (_esc(d["codigo"]), _dcx_tipo_txt(d["tipo"]), _esc(_dcx_area_titulo(d["area"])),
                 barra, doc, js))
    return _cb_pagina(d["codigo"], _DCX_CSS + cuerpo)


@app.route("/documentos-oficiales/anular", methods=["POST"])
def dcx_anular():
    g = _dcx_guard()
    if g is not None:
        return g
    if not _ti_csrf_ok():
        return redirect("/documentos-oficiales")
    rol, usuario = rol_actual(), session.get("usuario") or ""
    try:
        did = int(request.form.get("id") or 0)
    except Exception:
        did = 0
    d = _dcx_row(did=did) if did else None
    motivo = " ".join((request.form.get("motivo") or "").split())[:200]
    if not d or d["estado"] != "VIGENTE" or not (rol in _DCX_GERENCIA or d.get("creado_por") == usuario) or not motivo:
        session["ti_msg"] = ("err", "No se pudo anular el documento.")
        return redirect("/documentos-oficiales")
    _c48_x("UPDATE docs_corp SET estado='ANULADO', motivo=:m, anulado_por=:u, anulado=:a WHERE id=:i", m=motivo, u=usuario, a=str(ahora())[:19], i=did)
    registrar_auditoria("Documentos oficiales", "%s anuló %s: %s" % (usuario, d["codigo"], motivo))
    session["ti_msg"] = ("ok", "Documento anulado.")
    return redirect("/documentos-oficiales/ver/%d" % did)


# ───────────────────────── Verificación pública (la abre el QR)
@app.route("/docs/verificar/<token>")
def dcx_verificar(token):
    _dcx_init()
    d = _dcx_row(token=token) if re.match(r"^[A-Za-z0-9_-]{8,40}$", token or "") else None
    estilo = ("<style>.vf{max-width:560px;margin:34px auto;padding:0 16px;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#0b1437}"
              ".vf-c{background:#fff;border:1px solid #e3e8f7;border-radius:18px;padding:24px;box-shadow:0 14px 40px rgba(0,32,194,.10)}"
              ".vf-b{height:6px;border-radius:6px;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2);margin:-24px -24px 18px;border-radius:18px 18px 0 0}"
              ".vf h1{font-size:22px;margin:6px 0 4px}.vf-ok{color:#166534}.vf-no{color:#b91c1c}"
              ".vf dl{margin:16px 0 0;display:grid;grid-template-columns:130px 1fr;gap:8px 12px;font-size:14px}"
              ".vf dt{color:#64748b}.vf dd{margin:0;font-weight:600}.vf small{display:block;margin-top:16px;color:#64748b;line-height:1.5}"
              ".vf img{height:44px}</style>")
    if not d:
        cuerpo = ("<div class='vf'><div class='vf-c'><div class='vf-b'></div><img src='%s' alt='PROCSIS'>"
                  "<h1 class='vf-no'>⚠ Documento no encontrado</h1><p>Este código no corresponde a ningún documento emitido por PROCSIS. "
                  "Si recibiste este documento, desconfía de su autenticidad y comunícate con la empresa.</p></div></div>" % _DCX_LOGO)
        return page("Verificación de documento", estilo + cuerpo)
    integro = _dcx_huella(d) == (d.get("huella") or "")
    vigente = d["estado"] == "VIGENTE" and integro
    x = d.get("datos_j") or {}
    filas = [("Código", d["codigo"]), ("Tipo", _dcx_tipo_txt(d["tipo"])),
             ("Número", "%s de %s" % (d["numero"], d["anio"])), ("Área emisora", _dcx_area_titulo(d["area"])), ("Fecha", _dcx_fecha_larga(d["fecha"])),
             ("Asunto", d["asunto"])]
    if d["tipo"] == "oficio" and x.get("dest_nombre"):
        filas.append(("Destinatario", " · ".join(v for v in (x.get("dest_nombre"), x.get("dest_entidad")) if v)))
    filas += [("Firmado por", "%s%s" % (d["firmante"], (" · " + d["cargo"]) if d.get("cargo") else "")),
              ("Emitido", (d.get("creado") or "")[:16]), ("Huella", (d.get("huella") or "")[:16] + "…")]
    if d["estado"] != "VIGENTE":
        filas.append(("Anulado", "%s · %s" % ((d.get("anulado") or "")[:16], d.get("motivo") or "")))
    dl = "".join("<dt>%s</dt><dd>%s</dd>" % (_esc(a), _esc(b)) for a, b in filas)
    if vigente:
        titulo = "<h1 class='vf-ok'>✅ Documento auténtico y vigente</h1><p>Este documento fue emitido y registrado por PROCSIS.</p>"
    elif not integro:
        titulo = "<h1 class='vf-no'>⚠ Documento con inconsistencias</h1><p>Los datos registrados no coinciden con la huella de emisión. No lo tomes como válido y contacta a PROCSIS.</p>"
    else:
        titulo = "<h1 class='vf-no'>⛔ Documento anulado</h1><p>Este documento fue emitido por PROCSIS pero ya no tiene validez.</p>"
    cuerpo = ("<div class='vf'><div class='vf-c'><div class='vf-b'></div><img src='%s' alt='PROCSIS'>%s<dl>%s</dl>"
              "<small>La verificación confirma que el documento está registrado en el sistema de PROCSIS y su estado actual. "
              "Por privacidad no se muestra el contenido completo.</small></div></div>" % (_DCX_LOGO, titulo, dl))
    return page("Verificación de documento", estilo + cuerpo)


'''
CTC = r'''# ════════════════════════════════════════════════════════════════════════════
#  CORREOS Y LÍNEAS CORPORATIVAS (Soporte y Ventas) · se muestran en el panel principal
#  - Gerencia registra números de WhatsApp, líneas telefónicas y correos por área.
#  - El sistema normaliza los números (+57), ordena por área y reparte los chats de WhatsApp entre las líneas activas.
# ════════════════════════════════════════════════════════════════════════════
_CTC_INIT = {"ok": False}
_CTC_CACHE = {"t": 0.0, "v": []}
_CTC_AREAS = (("Soporte", "🎧", "soporte"), ("Ventas", "💼", "ventas"))
_CTC_TIPOS = (("whatsapp", "WhatsApp", "💬"), ("telefono", "Línea telefónica", "📞"), ("correo", "Correo", "✉️"))
_CTC_MSG = {"Soporte": "Hola, necesito soporte técnico de EduTrack.", "Ventas": "Hola, quiero información sobre los planes de EduTrack."}


def _ctc_init():
    if _CTC_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS contactos_corp (id %s, area VARCHAR(40), tipo VARCHAR(12), etiqueta VARCHAR(80), valor VARCHAR(160), "
            "mensaje VARCHAR(240) DEFAULT '', horario VARCHAR(120) DEFAULT '', activo INTEGER DEFAULT 1, orden INTEGER DEFAULT 0, "
            "usos INTEGER DEFAULT 0)" % pk))
        db.session.commit()
        _CTC_INIT["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("contactos init:", e)


def _ctc_invalidar():
    _CTC_CACHE["t"] = 0.0


def _ctc_num(v):
    d = re.sub(r"\D", "", v or "")
    if d.startswith("00"):
        d = d[2:]
    if len(d) == 10 and d.startswith("3"):
        d = "57" + d
    return d


def _ctc_mostrar(d):
    if d.startswith("57") and len(d) == 12:
        return "+57 %s %s %s" % (d[2:5], d[5:8], d[8:])
    return "+" + d


def _ctc_correo_ok(v):
    return bool(re.match(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$", v or ""))


def _ctc_activos():
    """Líneas y correos activos (copia en memoria de 30 s para no consultar la base en cada visita al inicio)."""
    import time
    if time.time() - _CTC_CACHE["t"] > 30:
        filas = []
        try:
            _ctc_init()
            for r in _c48_q("SELECT id, area, tipo, etiqueta, valor, mensaje, horario FROM contactos_corp WHERE activo=1 ORDER BY area, tipo, orden, id"):
                filas.append({"id": r[0], "area": r[1], "tipo": r[2], "etiqueta": r[3] or "", "valor": r[4] or "", "mensaje": r[5] or "", "horario": r[6] or ""})
        except Exception:
            try:
                db.session.rollback()
            except Exception:
                pass
        _CTC_CACHE["v"] = filas
        _CTC_CACHE["t"] = time.time()
    return _CTC_CACHE["v"]


def _ctc_hay(area, tipo):
    try:
        return any(f["area"] == area and f["tipo"] == tipo for f in _ctc_activos())
    except Exception:
        return False


def _ctc_wa_proximo(area):
    """Reparte los chats: elige la línea de WhatsApp activa del área con menos chats enviados."""
    from urllib.parse import quote
    rows = _c48_q("SELECT id, valor, mensaje FROM contactos_corp WHERE area=:a AND tipo='whatsapp' AND activo=1 ORDER BY usos, orden, id LIMIT 1", a=area)
    if not rows:
        return None
    cid, num, msg = rows[0]
    _c48_x("UPDATE contactos_corp SET usos=COALESCE(usos,0)+1 WHERE id=:i", i=cid)
    return "https://wa.me/%s?text=%s" % (num, quote(msg or _CTC_MSG.get(area, "")))


@app.route("/lineas/<slug>/whatsapp")
def ctc_whatsapp(slug):
    area = {"soporte": "Soporte", "ventas": "Ventas"}.get((slug or "").lower())
    url = None
    if area:
        try:
            _ctc_init()
            url = _ctc_wa_proximo(area)
        except Exception as e:
            db.session.rollback()
            print("contactos wa:", e)
    resp = redirect(url or "/contacto")
    resp.headers["Cache-Control"] = "no-store"
    return resp


_CTC_CSS = """<style>
.ctc-g{display:grid;gap:12px;grid-template-columns:1fr;width:100%}
.ctc-a{background:rgba(255,255,255,.82);border:1px solid rgba(0,184,212,.38);border-radius:16px;padding:14px 16px;box-shadow:0 10px 30px rgba(0,32,194,.08);text-align:left}
.ctc-h{font-weight:800;color:#0B1B4D;font-size:15px;margin-bottom:10px}
.ctc-wa{display:block;text-align:center;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2);color:#fff!important;font-weight:700;text-decoration:none;padding:12px 16px;border-radius:12px;margin-bottom:6px;text-shadow:0 1px 2px rgba(0,20,90,.4);transition:transform .2s ease,box-shadow .2s ease}
.ctc-wa:hover{transform:translateY(-2px);box-shadow:0 12px 26px rgba(0,32,194,.28)}
.ctc-l{display:flex;gap:10px;align-items:center;text-decoration:none;color:#0B1B4D!important;padding:8px 10px;border:1px solid rgba(0,184,212,.28);border-radius:12px;margin-top:8px;font-size:13.5px;transition:background .2s,border-color .2s}
.ctc-l:hover{background:rgba(0,184,212,.10);border-color:#00B8D4}
.ctc-l small{display:block;color:#64748b;font-size:11.5px;font-weight:400}
</style>"""


def _ctc_pills_html(corp_tel, corp_email):
    """Panel principal: líneas y correos organizados por área. Si Gerencia aún no configuró nada, deja las dos píldoras de siempre."""
    from urllib.parse import quote
    viejo = ('<div class="corp-pill"><span>📞</span><div><b>Línea telefónica</b><br>%s</div></div>'
             '<div class="corp-pill"><span>✉️</span><div><b>Correo electrónico</b><br>%s</div></div>' % (_esc(corp_tel), _esc(corp_email)))
    try:
        filas = _ctc_activos()
    except Exception:
        return viejo
    if not filas:
        return viejo
    bloques = []
    for area, ico, slug in _CTC_AREAS:
        mias = [f for f in filas if f["area"] == area]
        if not mias:
            continue
        items = []
        for tipo_k, tipo_n, tipo_i in _CTC_TIPOS:
            for f in [x for x in mias if x["tipo"] == tipo_k]:
                if tipo_k == "whatsapp":
                    href = "https://wa.me/%s?text=%s" % (f["valor"], quote(f["mensaje"] or _CTC_MSG.get(area, "")))
                    extra, dato = ' target="_blank" rel="noopener"', _ctc_mostrar(f["valor"])
                elif tipo_k == "telefono":
                    href, extra, dato = "tel:+" + f["valor"], "", _ctc_mostrar(f["valor"])
                else:
                    href, extra, dato = "mailto:" + f["valor"], "", f["valor"]
                horario = ("<small>%s</small>" % _esc(f["horario"])) if f["horario"] else ""
                items.append('<a class="ctc-l" href="%s"%s><span>%s</span><div><b>%s</b> · %s%s</div></a>'
                             % (_esc(href), extra, tipo_i, _esc(f["etiqueta"] or tipo_n), _esc(dato), horario))
        boton = ('<a class="ctc-wa" href="/lineas/%s/whatsapp" target="_blank" rel="noopener">💬 Escribir por WhatsApp</a>' % slug
                 if any(x["tipo"] == "whatsapp" for x in mias) else "")
        bloques.append('<div class="ctc-a"><div class="ctc-h">%s %s</div>%s%s</div>' % (ico, _esc(area if area != "Soporte" else "Soporte técnico"), boton, "".join(items)))
    if not bloques:
        return viejo
    return _CTC_CSS + '<div class="ctc-g">%s</div>' % "".join(bloques)


# ───────────────────────── Administración en Gerencia
def _ctc_hidden(csrf, accion, campos):
    h = "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='%s'>" % (_esc(csrf), _esc(accion))
    for k, v in campos.items():
        h += "<input type='hidden' name='%s' value='%s'>" % (_esc(k), _esc(v))
    return h


def _ctc_boton(csrf, accion, campos, texto, clase="", confirmar=""):
    conf = (" onsubmit=\"return confirm('%s');\"" % _esc(confirmar).replace("'", "")) if confirmar else ""
    return ("<form method='POST' action='/gerencia/contactos' style='display:inline'%s>%s<button class='cb-btn s %s'>%s</button></form> "
            % (conf, _ctc_hidden(csrf, accion, campos), clase, _esc(texto)))


def _ctc_accion(f):
    a = (f.get("accion") or "").strip()
    quien = session.get("usuario") or ""

    def cid():
        v = (f.get("id") or "").strip()
        return int(v) if v.isdigit() else 0
    if a == "agregar":
        area = (f.get("area") or "").strip()
        tipo = (f.get("tipo") or "").strip()
        if area not in [x[0] for x in _CTC_AREAS] or tipo not in [x[0] for x in _CTC_TIPOS]:
            return False, "Elige el área y el tipo."
        bruto = (f.get("valor") or "").strip()
        if tipo == "correo":
            valor = bruto.lower()
            if not _ctc_correo_ok(valor):
                return False, "Ese correo no es válido."
        else:
            valor = _ctc_num(bruto)
            if not (10 <= len(valor) <= 15):
                return False, "El número debe tener entre 10 y 15 dígitos (ej.: 310 561 5621)."
        if _c48_q("SELECT 1 FROM contactos_corp WHERE area=:a AND tipo=:t AND valor=:v", a=area, t=tipo, v=valor):
            return False, "Ese dato ya está registrado en %s." % area
        n = int((_c48_q("SELECT COUNT(*) FROM contactos_corp WHERE area=:a AND tipo=:t", a=area, t=tipo)[0][0]) or 0)
        etiqueta = " ".join((f.get("etiqueta") or "").split())[:80] or ("Correo %d" % (n + 1) if tipo == "correo" else "Línea %d" % (n + 1))
        mensaje = " ".join((f.get("mensaje") or "").split())[:240] if tipo == "whatsapp" else ""
        if tipo == "whatsapp" and not mensaje:
            mensaje = _CTC_MSG.get(area, "")
        horario = " ".join((f.get("horario") or "").split())[:120]
        mx = int((_c48_q("SELECT COALESCE(MAX(orden), -1) FROM contactos_corp WHERE area=:a AND tipo=:t", a=area, t=tipo)[0][0]) or 0)
        _c48_x("INSERT INTO contactos_corp (area, tipo, etiqueta, valor, mensaje, horario, activo, orden, usos) VALUES (:a, :t, :e, :v, :m, :h, 1, :o, 0)",
               a=area, t=tipo, e=etiqueta, v=valor, m=mensaje, h=horario, o=mx + 1)
        _ctc_invalidar()
        registrar_auditoria("Contactos corporativos", "%s agregó %s de %s: %s" % (quien, tipo, area, valor))
        return True, "Agregado a %s." % area
    if a in ("toggle", "eliminar", "subir", "bajar"):
        i = cid()
        r = _c48_q("SELECT area, tipo FROM contactos_corp WHERE id=:i", i=i) if i else []
        if not r:
            return False, "Registro no válido."
        area, tipo = r[0]
        if a == "toggle":
            _c48_x("UPDATE contactos_corp SET activo = CASE WHEN activo=1 THEN 0 ELSE 1 END WHERE id=:i", i=i)
        elif a == "eliminar":
            _c48_x("DELETE FROM contactos_corp WHERE id=:i", i=i)
            registrar_auditoria("Contactos corporativos", "%s eliminó el registro #%d" % (quien, i))
        else:
            ids = [x[0] for x in _c48_q("SELECT id FROM contactos_corp WHERE area=:a AND tipo=:t ORDER BY orden, id", a=area, t=tipo)]
            p = ids.index(i)
            q = p - 1 if a == "subir" else p + 1
            if 0 <= q < len(ids):
                ids[p], ids[q] = ids[q], ids[p]
            for n, x in enumerate(ids):
                _c48_x("UPDATE contactos_corp SET orden=:o WHERE id=:i", o=n, i=x)
        _ctc_invalidar()
        return True, "Listo."
    if a == "reiniciar":
        area = (f.get("area") or "").strip()
        _c48_x("UPDATE contactos_corp SET usos=0 WHERE area=:a", a=area)
        return True, "Contadores de %s reiniciados." % area
    if a == "importar":
        try:
            p = plataforma()
        except Exception:
            p = None
        cand = [("Soporte", "whatsapp", getattr(p, "contacto_whatsapp_soporte", "")), ("Soporte", "telefono", getattr(p, "contacto_publico_tel", "") or getattr(p, "telefono_soporte", "")),
                ("Soporte", "correo", getattr(p, "contacto_publico_email", "") or getattr(p, "email_soporte", "")), ("Ventas", "whatsapp", getattr(p, "contacto_whatsapp_ventas", ""))]
        n = 0
        for area, tipo, bruto in cand:
            bruto = (bruto or "").strip()
            if not bruto:
                continue
            valor = bruto.lower() if tipo == "correo" else _ctc_num(bruto)
            if (tipo == "correo" and not _ctc_correo_ok(valor)) or (tipo != "correo" and not (10 <= len(valor) <= 15)):
                continue
            if _c48_q("SELECT 1 FROM contactos_corp WHERE area=:a AND tipo=:t AND valor=:v", a=area, t=tipo, v=valor):
                continue
            _c48_x("INSERT INTO contactos_corp (area, tipo, etiqueta, valor, mensaje, horario, activo, orden, usos) VALUES (:a, :t, :e, :v, :m, '', 1, 0, 0)",
                   a=area, t=tipo, e="Principal", v=valor, m=_CTC_MSG.get(area, "") if tipo == "whatsapp" else "")
            n += 1
        _ctc_invalidar()
        return (True, "Se importaron %d dato(s) de la configuración actual." % n) if n else (False, "No había datos nuevos para importar.")
    return False, "Acción no reconocida."


@app.route("/gerencia/contactos", methods=["GET", "POST"])
def ctc_admin():
    g = _guard_gerencia()
    if g is not None:
        return g
    _ctc_init()
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/gerencia/contactos")
        try:
            ok, msg = _ctc_accion(request.form)
        except Exception as e:
            db.session.rollback()
            print("contactos:", e)
            ok, msg = False, "No se pudo guardar el cambio. Intenta de nuevo."
        session["ti_msg"] = ("ok" if ok else "err", msg)
        return redirect("/gerencia/contactos")
    _ctc_invalidar()
    filas = _c48_q("SELECT id, area, tipo, etiqueta, valor, horario, activo, usos FROM contactos_corp ORDER BY area, tipo, orden, id")
    nombres = {k: (n, i) for k, n, i in _CTC_TIPOS}
    tarjetas = ""
    for area, ico, slug in _CTC_AREAS:
        mias = [r for r in filas if r[1] == area]
        trs = ""
        for cid_, _a, tipo, etq, val, hor, act, usos in mias:
            dato = val if tipo == "correo" else _ctc_mostrar(val)
            c = {"id": cid_}
            trs += ("<tr><td>%s %s</td><td><b>%s</b></td><td>%s</td><td>%s</td><td>%s</td><td><span class='cb-tag %s'>%s</span></td>"
                    "<td>%s%s%s%s</td></tr>"
                    % (nombres[tipo][1], _esc(nombres[tipo][0]), _esc(etq), _esc(dato), _esc(hor), (str(int(usos or 0)) if tipo == "whatsapp" else "—"),
                       "on" if act else "off", "Activo" if act else "Inactivo",
                       _ctc_boton(csrf, "subir", c, "↑"), _ctc_boton(csrf, "bajar", c, "↓"),
                       _ctc_boton(csrf, "toggle", c, "Desactivar" if act else "Activar", "r" if act else "g"),
                       _ctc_boton(csrf, "eliminar", c, "Eliminar", "r", "¿Eliminar este registro?")))
        if not trs:
            trs = "<tr><td colspan='7'>Aún no hay líneas ni correos en esta área.</td></tr>"
        tarjetas += ("<div class='cb-card'><h2>%s %s</h2><div class='cb-scroll'><table><tr><th>Tipo</th><th>Nombre</th><th>Dato</th><th>Horario</th>"
                     "<th>Chats enviados</th><th>Estado</th><th>Acciones</th></tr>%s</table></div><p style='margin-top:10px'>%s</p></div>"
                     % (ico, _esc(area), trs, _ctc_boton(csrf, "reiniciar", {"area": area}, "Reiniciar contadores de WhatsApp", "o")))
    opt_area = "".join("<option value='%s'>%s</option>" % (a, a) for a, _i, _s in _CTC_AREAS)
    opt_tipo = "".join("<option value='%s'>%s</option>" % (k, n) for k, n, _i in _CTC_TIPOS)
    agregar = ("<div class='cb-card'><h2>Agregar línea o correo</h2><form method='POST' action='/gerencia/contactos'>%s"
               "<div style='display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0 14px'>"
               "<div><label>Área</label><select name='area'>%s</select></div><div><label>Tipo</label><select name='tipo'>%s</select></div>"
               "<div><label>Número o correo</label><input type='text' name='valor' required maxlength='160' placeholder='310 561 5621 o soporte@procsis.com'></div>"
               "<div><label>Nombre (opcional)</label><input type='text' name='etiqueta' maxlength='80' placeholder='Línea 1, Ventas Medellín…'></div>"
               "<div><label>Horario (opcional)</label><input type='text' name='horario' maxlength='120' placeholder='Lun–Vie 7:00 a.m. – 5:00 p.m.'></div>"
               "<div><label>Mensaje inicial de WhatsApp (opcional)</label><input type='text' name='mensaje' maxlength='240' placeholder='Hola, necesito soporte…'></div></div>"
               "<p style='margin-top:12px'><button class='cb-btn g'>Agregar</button></p></form></div>" % (_ctc_hidden(csrf, "agregar", {}), opt_area, opt_tipo))
    vista = _ctc_pills_html("", "")
    previa = ("<div class='cb-card'><h2>Así se ve en el panel principal</h2><p class='sub'>Es el bloque de soporte de la página de inicio. "
              "Cada botón «Escribir por WhatsApp» abre el chat con la línea activa que menos chats ha recibido de esa área.</p>"
              "<div style='max-width:420px'>%s</div></div>" % vista)
    cuerpo = ("<h1>📞 Correos y líneas corporativas</h1><p class='sub'>Registra los números de WhatsApp, las líneas telefónicas y los correos de Soporte y Ventas. "
              "El sistema normaliza los números (+57), los ordena por área y los muestra en el panel principal; el cliente toca WhatsApp y el chat se abre "
              "directamente. Si hay varias líneas activas en un área, los chats se reparten entre ellas.</p>"
              "<p>%s</p>%s%s%s" % (_ctc_boton(csrf, "importar", {}, "Importar datos de la configuración actual", "o"), agregar, tarjetas, previa))
    return _cb_pagina("Correos y líneas corporativas", cuerpo)


'''

def salir(msg, code=1):
    print(msg)
    sys.exit(code)

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
NL = "\r\n" if "\r\n" in s else "\n"

def crlf(t):
    return t.replace("\r\n", "\n").replace("\n", NL)

def una_vez(ancla, donde=None):
    base = s if donde is None else donde
    if base.count(ancla) != 1:
        salir("ERROR: punto de inserción no encontrado (%r: hay %d). No se modificó nada." % (ancla[:70], base.count(ancla)))

if "def _dcx_html_comunicado(" in s and "_CTC_INIT = {" in s:
    salir("Este app.py YA tiene comunicados y contactos corporativos. No se modificó nada.", 0)
A_MAIN = 'if __name__ == "__main__":'
una_vez(A_MAIN)

# ── 1) Documentos oficiales con comunicados (actualiza la versión anterior o instala completa)
if "def _dcx_html_comunicado(" not in s:
    if "_DCX_INIT = {" in s:
        marca = "#  DOCUMENTOS OFICIALES · Resoluciones"
        i = s.index(marca)
        ini = s.rfind("# ═", 0, i)
        fin_txt = 'return page("Verificación de documento", estilo + cuerpo)'
        j = s.rfind(fin_txt)
        if ini < 0 or j < ini:
            salir("ERROR: no pude delimitar el módulo de documentos actual. No se modificó nada.")
        fin = j + len(fin_txt)
        s = s[:ini] + crlf(DOCS.rstrip("\n")) + s[fin:]
    else:
        una_vez(A_MAIN)
        s = s.replace(A_MAIN, crlf(DOCS) + A_MAIN, 1)
        A2 = "    ('Sistema y configuración', ["
        una_vez(A2)
        GRUPO = """    ('Comunicaciones y CRTP', [
        ('/gerencia/comunicaciones', '📜 Comunicaciones: comunicados, resoluciones y oficios'),
        ('/gerencia/crtp', '⚖️ Comisión de Regulación Tecnológica (CRTP)'),
        ('/documentos-oficiales', '📄 Documentos oficiales (todas las áreas)'),
    ]),
"""
        s = s.replace(A2, crlf(GRUPO) + A2, 1)
        A3 = '    if rol in _ROLES_MAPA and not any(h == "/mapa" for h, _l in items):'
        una_vez(A3)
        MENU = """    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):
        items = [x for x in items if x[0] != "/logout"] + [("/documentos-oficiales", "📄 Documentos oficiales")] + [x for x in items if x[0] == "/logout"]
"""
        s = s.replace(A3, crlf(MENU) + A3, 1)

# ── 2) Correos y líneas corporativas
if "_CTC_INIT = {" not in s:
    una_vez(A_MAIN)
    s = s.replace(A_MAIN, crlf(CTC) + A_MAIN, 1)

# ── 3) Menú de Gerencia: entrada de contactos
if "/gerencia/contactos'" not in s:
    A4 = "        ('/documentos-oficiales', '📄 Documentos oficiales (todas las áreas)'),"
    una_vez(A4)
    s = s.replace(A4, A4 + NL + "        ('/gerencia/contactos', '📞 Correos y líneas corporativas')," , 1)

# ── 4) Panel principal: bloque de líneas y correos
if "{_ctc_pills_html(corp_tel, corp_email)}" not in s:
    pat = re.compile(r'(?P<ind>[ \t]*)<div class="corp-pill"><span>📞</span><div><b>Línea telefónica</b><br>\{corp_tel\}</div></div>\r?\n[ \t]*<div class="corp-pill"><span>✉\ufe0f?</span><div><b>Correo electrónico</b><br>\{corp_email\}</div></div>')
    if len(pat.findall(s)) != 1:
        salir("ERROR: no encontré el bloque de contacto del panel principal. No se modificó nada.")
    s = pat.sub(lambda m: m.group("ind") + "{_ctc_pills_html(corp_tel, corp_email)}", s, count=1)

# ── 5) Botones de WhatsApp: ventas y soporte pasan por el reparto automático
A5 = '        wa_link = (f"https://wa.me/{wa_num}?text={_uq(_wa_txt)}" if wa_num else "/ventas")'
if "_ctc_hay(\"Ventas\"" not in s:
    una_vez(A5)
    s = s.replace(A5, A5 + NL + '        if _ctc_hay("Ventas", "whatsapp"):' + NL + '            wa_link = "/lineas/ventas/whatsapp"', 1)
A6 = '    wa = f"https://wa.me/{wa_num}?text={quote(msg)}" if wa_num else ""'
if "_ctc_hay(\"Soporte\"" not in s:
    a = s.index("def portal_whatsapp_soporte():")
    seg = s[a:a + 6000]
    una_vez(A6, seg)
    k = a + seg.index(A6) + len(A6)
    s = s[:k] + NL + '    if _ctc_hay("Soporte", "whatsapp"):' + NL + '        wa = "/lineas/soporte/whatsapp"' + s[k:]

tmp = SRC + ".nuevo"
with io.open(tmp, "w", encoding="utf-8", newline="") as fh:
    fh.write(s)
try:
    py_compile.compile(tmp, doraise=True)
except Exception as e:
    os.remove(tmp)
    salir("ERROR: el resultado no compila, no se modificó tu app.py: %s" % e)
shutil.copyfile(SRC, SRC + ".respaldo")
os.replace(tmp, SRC)
print("LISTO. Comunicados y contactos corporativos instalados. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_documentos_area2.py  ·  PROCSIS: «Documentos por área» versión 2 (reemplaza a aplicar_documentos_area.py).

Uso (en la carpeta de tu app.py):
    python aplicar_documentos_area2.py            # parchea ./app.py
    python aplicar_documentos_area2.py ruta/app.py

Sirve aunque ya hayas aplicado la versión 1: la reemplaza por esta (no pierdes los documentos ya creados).

Qué trae:
  · TODAS las áreas redactan y firman sus documentos (Recursos Humanos, Soluciones, Gerencia, TI, Soporte, Comercial, Cobranza,
    Contabilidad, Finanzas, Marketing, Jurídica, Calidad, Operaciones...). El documento termina firmado por la entidad que lo emite,
    por ejemplo «Recursos Humanos PROCSIS» o «Soluciones PROCSIS» (editable).
  · ENVÍO POR CORREO en PDF a cualquier persona: Para, Con copia, asunto y mensaje. Se puede enviar al emitir o después, desde el
    documento. Queda un historial de envíos. Usa la conexión de correo que ya tienes (Gmail o Resend).
  · EDITOR con formato: negrita, cursiva, subrayado, títulos, listas con viñetas y numeradas, textos rápidos y vista previa en vivo.
    El formato se respeta en pantalla, en el PDF y en Word.
  · FORMATO CORPORATIVO (estilo grandes operadores): banda de color, logo, cuadro de radicado («Al contestar cite…»), referencia,
    asunto, anexos, con copia, firma de la entidad, pie con sede, servicio al cliente, número de página «Página X de Y» y
    CÓDIGO QR de verificación pública (/docs/area/<token>).
Es idempotente, crea app.py.bak_documentos_area2 y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-documentos-area-v2]"
MARCA_V1 = "# [procsis-documentos-area]\n# ═════"

BLOQUE = r'''
# [procsis-documentos-area-v2]
# ═════════════ DOCUMENTOS POR ÁREA v2: redactar con formato, firmar como el área, enviar por correo en PDF ═════════════
_DCA = {"ok": False}
_DCA_CATEGORIAS = [("queja", "Queja"), ("reclamo", "Reclamo"), ("peticion", "Petición o solicitud"), ("financiero", "Documento financiero"),
                   ("informe", "Informe técnico"), ("notificacion", "Notificación o circular"), ("certificacion", "Certificación"),
                   ("requerimiento", "Requerimiento interno"), ("comercial", "Propuesta o carta comercial"), ("acta", "Acta o constancia"), ("otro", "Otro")]
_DCA_ROLES_EXTRA = ("Soporte", "Cobranza", "Comercial", "Ventas", "Supervisor de Ventas", "Soluciones", "Desarrollador", "Developer")
_DCA_COLS = ("id, codigo, categoria, origen, destino, asunto, trato, dest_nombre, dest_lineas, saludo, cuerpo, despedida, firmante, cargo, "
             "ciudad, fecha, estado, creado_por, creado, leido_por, leido, nota, numero, anio, entidad, ref, cc, anexos, token")
_DCA_ENTIDADES = {"Gerencia": "Gerencia General PROCSIS", "Talento": "Recursos Humanos PROCSIS", "Soluciones": "Soluciones PROCSIS",
                  "Tecnologia": "Tecnología y TI PROCSIS", "Soporte": "Soporte Técnico PROCSIS", "Comercial": "Área Comercial PROCSIS",
                  "Cobranza": "Cobranza PROCSIS", "Contabilidad": "Contabilidad PROCSIS", "Finanzas": "Finanzas y Tesorería PROCSIS",
                  "Marketing": "Comunicaciones y Mercadeo PROCSIS", "Juridica": "Jurídica y Cumplimiento PROCSIS",
                  "Operaciones": "Operaciones y Administración PROCSIS", "Calidad": "Calidad y Seguridad de la Información PROCSIS",
                  "CRTP": "Comisión de Regulación Tecnológica PROCSIS", "Desarrollador": "Desarrollo PROCSIS"}
_DCA_PLANTILLAS_DEST = {
    "comunidad": {"trato": "Señor(a)S", "nombre": "DIRECTORES, RECTORES Y ALIADOS ESTRATÉGICOS",
                  "lineas": "Comunidad Educativa EduTrack Colombia\nInstituciones Usuarias de PROCSIS y EduTrack\nEcosistema Nacional"},
    "tiendas": {"trato": "Señor(a)S", "nombre": "PROPIETARIOS Y ADMINISTRADORES DE NEGOCIOS",
                "lineas": "Comunidad de comercios y restaurantes TiendaTrack\nClientes de PROCSIS"},
    "interna": {"trato": "Señor(a)", "nombre": "", "lineas": ""},
}


def _dca_roles_ok(rol):
    rol = (rol or "").strip()
    return _dcx_rol_ok(rol) or rol in _DCA_ROLES_EXTRA


def _dca_areas():
    out = []
    for a in list(_dcx_areas_validas()) + ["Soluciones", "Desarrollador"]:
        if a not in out:
            out.append(a)
    return out


def _dca_cat_txt(k):
    return dict(_DCA_CATEGORIAS).get(k, "Otro")


def _dca_entidad_def(area):
    if area in _DCA_ENTIDADES:
        return _DCA_ENTIDADES[area]
    t = _dcx_area_titulo(area) or area or "PROCSIS"
    return t if "PROCSIS" in t.upper() else t + " PROCSIS"


def _dca_una_linea(s):
    return " ".join(str(s or "").replace("\r", " ").replace("\n", " ").split())


def _dca_init():
    if _DCA["ok"]:
        return
    sqlite = db.engine.dialect.name == "sqlite"
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if sqlite else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS docs_area (id %s, codigo VARCHAR(40), categoria VARCHAR(30), origen VARCHAR(60), destino VARCHAR(60), "
            "asunto TEXT, trato VARCHAR(60), dest_nombre VARCHAR(220), dest_lineas TEXT DEFAULT '', saludo VARCHAR(160) DEFAULT '', "
            "cuerpo TEXT, despedida VARCHAR(160) DEFAULT '', firmante VARCHAR(160), cargo VARCHAR(160) DEFAULT '', ciudad VARCHAR(100), "
            "fecha VARCHAR(10), estado VARCHAR(12) DEFAULT 'ENVIADO', creado_por VARCHAR(80), creado VARCHAR(19), "
            "leido_por VARCHAR(80) DEFAULT '', leido VARCHAR(19) DEFAULT '', nota TEXT DEFAULT '', numero INTEGER DEFAULT 0, anio INTEGER DEFAULT 0, "
            "entidad VARCHAR(220) DEFAULT '', ref VARCHAR(220) DEFAULT '', cc TEXT DEFAULT '', anexos TEXT DEFAULT '', token VARCHAR(40) DEFAULT '')" % pk))
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS docs_area_envios (id %s, doc_id INTEGER, para TEXT, cc TEXT DEFAULT '', asunto TEXT DEFAULT '', "
            "por VARCHAR(80), fecha VARCHAR(19), ok INTEGER DEFAULT 0, error TEXT DEFAULT '')" % pk))
        db.session.commit()
        if sqlite:
            cols = [r[1] for r in db.session.execute(text("PRAGMA table_info(docs_area)")).fetchall()]
        else:
            cols = [r[0] for r in db.session.execute(text(
                "SELECT column_name FROM information_schema.columns WHERE table_schema = current_schema() AND table_name = 'docs_area'")).fetchall()]
        for nombre, ddl in (("entidad", "VARCHAR(220) DEFAULT ''"), ("ref", "VARCHAR(220) DEFAULT ''"), ("cc", "TEXT DEFAULT ''"),
                            ("anexos", "TEXT DEFAULT ''"), ("token", "VARCHAR(40) DEFAULT ''")):
            if nombre not in cols:
                db.session.execute(text("ALTER TABLE docs_area ADD COLUMN %s %s" % (nombre, ddl)))
                db.session.commit()
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_docs_area_destino ON docs_area (destino)"))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_docs_area_origen ON docs_area (origen)"))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_docs_area_envios_doc ON docs_area_envios (doc_id)"))
        db.session.commit()
        _DCA["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("docs_area init:", e)


def _dca_guard():
    if not requiere_login():
        return redirect("/login")
    if not _dca_roles_ok(rol_actual()):
        return acceso_denegado("Tu rol no tiene acceso a Documentos por área.")
    try:
        from flask import g as _g_dca
        _g_dca._staff_wrap = True
    except Exception:
        pass
    _dca_init()
    return None


def _dca_area_usuario():
    return _dcx_area_de_rol(rol_actual())


def _dca_row(did=None, token=None):
    if did is not None:
        rows = _c48_q("SELECT %s FROM docs_area WHERE id=:i" % _DCA_COLS, i=did)
    else:
        rows = _c48_q("SELECT %s FROM docs_area WHERE token=:t" % _DCA_COLS, t=token)
    if not rows:
        return None
    d = dict(zip([c.strip() for c in _DCA_COLS.split(",")], rows[0]))
    for k in ("entidad", "ref", "cc", "anexos", "token"):
        d[k] = d.get(k) or ""
    if not d["entidad"]:
        d["entidad"] = _dca_entidad_def(d.get("origen"))
    return d


def _dca_puede(d, rol, usuario):
    a = _dcx_area_de_rol(rol)
    return rol in _DCX_GERENCIA or d.get("origen") == a or d.get("destino") == a or d.get("creado_por") == usuario


def _dca_puede_enviar(d, rol, usuario):
    return rol in _DCX_GERENCIA or d.get("origen") == _dcx_area_de_rol(rol) or d.get("creado_por") == usuario


def _dca_logo_path():
    for p in (Path(app.root_path) / "static" / "img" / "logo-procsis-mark.png", Path(app.root_path) / "static" / "img" / "logo-procsis.png"):
        if p.is_file():
            return str(p)
    return ""


def _dca_lineas(d):
    return [x.strip() for x in (d.get("dest_lineas") or "").replace("\r", "").split("\n") if x.strip()]


def _dca_url(d):
    return "%s/docs/area/%s" % (_dcx_base(), d.get("token") or "")


def _dca_qr_png(url):
    buf = BytesIO()
    qrcode.make(url, box_size=6, border=1).save(buf, format="PNG")
    buf.seek(0)
    return buf


def _dca_hora(d):
    s = (d.get("creado") or "")[:16].replace("T", " ")
    return s or d.get("fecha") or ""


def _dca_destinatario_txt(d):
    return d.get("dest_nombre") or (_dcx_area_titulo(d.get("destino")) if d.get("destino") else "Correo externo")


# ───────────── Formato del cuerpo: **negrita**  *cursiva*  __subrayado__  # título  - viñeta  1. numerada ─────────────
_DCA_INLINE = re.compile(r"(\*\*.+?\*\*|__.+?__|\*[^*\n]+?\*)")


def _dca_runs(t):
    out = []
    for x in _DCA_INLINE.split(t or ""):
        if not x:
            continue
        if x.startswith("**") and x.endswith("**") and len(x) > 4:
            out.append((x[2:-2], True, False, False))
        elif x.startswith("__") and x.endswith("__") and len(x) > 4:
            out.append((x[2:-2], False, False, True))
        elif x.startswith("*") and x.endswith("*") and len(x) > 2:
            out.append((x[1:-1], False, True, False))
        else:
            out.append((x, False, False, False))
    return out


def _dca_bloques(t):
    bl, buf = [], []

    def flush():
        if buf:
            bl.append(("p", "\n".join(buf)))
            del buf[:]
    for ln in (t or "").replace("\r", "").split("\n"):
        s = ln.strip()
        m = re.match(r"^(\d+)[.)]\s+(.*)$", s)
        if not s:
            flush()
        elif s.startswith("# "):
            flush()
            bl.append(("h", s[2:].strip()))
        elif s[:2] in ("- ", "• "):
            flush()
            bl.append(("ul", s[2:].strip()))
        elif m:
            flush()
            bl.append(("ol", m.group(2).strip(), m.group(1)))
        else:
            buf.append(s)
    flush()
    return bl


def _dca_inline_html(t):
    h = ""
    for txt, b, i, u in _dca_runs(t):
        x = _esc(txt).replace("\n", "<br>")
        if b:
            x = "<b>%s</b>" % x
        if i:
            x = "<i>%s</i>" % x
        if u:
            x = "<u>%s</u>" % x
        h += x
    return h


def _dca_cuerpo_html(t):
    out, lst = [], None
    for b in _dca_bloques(t):
        if lst and b[0] != lst:
            out.append("</%s>" % lst)
            lst = None
        if b[0] == "p":
            out.append("<p>%s</p>" % _dca_inline_html(b[1]))
        elif b[0] == "h":
            out.append("<p class='h'>%s</p>" % _dca_inline_html(b[1]))
        else:
            if lst != b[0]:
                out.append("<%s>" % b[0])
                lst = b[0]
            out.append("<li>%s</li>" % _dca_inline_html(b[1]))
    if lst:
        out.append("</%s>" % lst)
    return "".join(out)


# ───────────── Estilos y vista previa (HTML) ─────────────
_DCA_CSS = r"""<style>
.dca-hoja{background:#fff;max-width:794px;margin:0 auto;padding:0 54px 30px;border:1px solid #e2e8f0;border-radius:6px;box-shadow:0 10px 30px rgba(0,16,80,.10);color:#111;line-height:1.55;font-size:14px;overflow:hidden}
.dca-banda{height:8px;margin:0 -54px 18px;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2)}
.dca-top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;margin-bottom:18px}
.dca-top img.lg{height:54px;width:auto}.dca-rad{display:flex;gap:10px;align-items:flex-start;font-size:10.5px;color:#334155;text-align:right;line-height:1.45}
.dca-rad img{width:64px;height:64px}
.dca-dest p{margin:0}.dca-ref{margin:14px 0 0}.dca-asunto{text-align:center;margin:16px 0 18px}
.dca-cuerpo p{text-align:justify;margin:0 0 11px}.dca-cuerpo p.h{font-weight:800;font-size:15px;margin:16px 0 6px;color:#0B2A6B}
.dca-cuerpo ul,.dca-cuerpo ol{margin:0 0 11px 22px;padding:0}.dca-cuerpo li{margin:0 0 4px;text-align:justify}
.dca-firma{margin-top:34px}.dca-firma .l{width:240px;border-top:1px solid #111;margin-bottom:4px}.dca-firma .ent{color:#0B2A6B;font-weight:800}
.dca-peq{font-size:12px;color:#334155;margin-top:14px}
.dca-foot{margin-top:26px;border-top:1px solid #cbd5e1;padding-top:8px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;font-size:10.5px;color:#475569;line-height:1.4}
.dca-foot b{display:block;color:#0B2A6B}
.dca-bar{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0;align-items:center}.dca-fg{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.dca-fg label{display:block;font-size:12px;font-weight:700;color:#475569;margin:0 0 4px}
.dca-fg input,.dca-fg select,.dca-fg textarea{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cfd8f3;border-radius:10px;font-size:14px;font-family:inherit}
.dca-full{grid-column:1/-1}tr.dca-sl td{background:#f0f7ff;font-weight:700}
.dca-tb{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px;align-items:center}
.dca-tb button{border:1px solid #cfd8f3;background:#fff;border-radius:8px;padding:6px 11px;cursor:pointer;font-size:13px}.dca-tb button:hover{background:#eff6ff}
.dca-tb select{width:auto!important;padding:6px 8px!important}
.dca-prev{margin-top:8px;border:1px dashed #cfd8f3;border-radius:10px;padding:12px 16px;background:#fafcff;font-size:13.5px}
.dca-prev p{margin:0 0 9px}.dca-prev p.h{font-weight:800;color:#0B2A6B}.dca-prev ul,.dca-prev ol{margin:0 0 9px 20px;padding:0}
.dca-env{border:1px solid #cfd8f3;border-radius:12px;padding:12px 14px;background:#f8fbff;margin:10px 0}
@media(max-width:700px){.dca-fg{grid-template-columns:1fr}.dca-hoja{padding:0 18px 20px}.dca-banda{margin:0 -18px 14px}.dca-foot{grid-template-columns:1fr}}
</style>"""

_DCA_JS_EDITOR = r"""<script>
(function(){
 var ta=document.getElementById('dca-cuerpo'),pv=document.getElementById('dca-prev');if(!ta||!pv)return;
 var esc=function(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');};
 var inl=function(s){s=esc(s);s=s.replace(/\*\*(.+?)\*\*/g,'<b>$1</b>').replace(/__(.+?)__/g,'<u>$1</u>').replace(/\*([^*\n]+?)\*/g,'<i>$1</i>');return s.replace(/\n/g,'<br>');};
 function render(t){var out=[],buf=[],lst=null;
  function fl(){if(buf.length){out.push('<p>'+inl(buf.join('\n'))+'</p>');buf=[];}}
  function cl(){if(lst){out.push('</'+lst+'>');lst=null;}}
  t.replace(/\r/g,'').split('\n').forEach(function(ln){var s=ln.trim(),m;
   if(!s){fl();cl();return;}
   if(s.indexOf('# ')===0){fl();cl();out.push('<p class="h">'+inl(s.slice(2))+'</p>');}
   else if(s.indexOf('- ')===0||s.indexOf('• ')===0){fl();if(lst!=='ul'){cl();out.push('<ul>');lst='ul';}out.push('<li>'+inl(s.slice(2))+'</li>');}
   else if((m=s.match(/^(\d+)[.)]\s+(.*)$/))){fl();if(lst!=='ol'){cl();out.push('<ol>');lst='ol';}out.push('<li>'+inl(m[2])+'</li>');}
   else{cl();buf.push(s);}});
  fl();cl();return out.join('')||'<span style="color:#94a3b8">La vista previa aparece aquí mientras escribes.</span>';}
 function upd(){pv.innerHTML=render(ta.value);}
 function wrap(l,r){var a=ta.selectionStart,b=ta.selectionEnd,v=ta.value,sel=v.slice(a,b)||'texto';ta.value=v.slice(0,a)+l+sel+r+v.slice(b);ta.focus();ta.selectionStart=a+l.length;ta.selectionEnd=a+l.length+sel.length;upd();}
 function lineas(fn){var a=ta.selectionStart,b=ta.selectionEnd,v=ta.value,s=v.lastIndexOf('\n',a-1)+1,e=v.indexOf('\n',b);if(e<0)e=v.length;
  var ls=v.slice(s,e).split('\n').map(fn);ta.value=v.slice(0,s)+ls.join('\n')+v.slice(e);ta.focus();upd();}
 var limpia=function(l){return l.replace(/^(\d+[.)]\s+|[-•]\s+|#\s+)/,'');};
 var acc={b:function(){wrap('**','**');},i:function(){wrap('*','*');},u:function(){wrap('__','__');},
  h:function(){lineas(function(l){return '# '+limpia(l);});},ul:function(){lineas(function(l){return '- '+limpia(l);});},
  ol:function(){var n=0;lineas(function(l){n++;return n+'. '+limpia(l);});}};
 [].forEach.call(document.querySelectorAll('#dca-tb [data-k]'),function(b){b.addEventListener('click',function(){acc[b.getAttribute('data-k')]();});});
 var ins=document.getElementById('dca-ins');
 if(ins)ins.addEventListener('change',function(){var t=ins.value;if(!t)return;var a=ta.selectionStart,v=ta.value,pre=(a>0&&v.charAt(a-1)!=='\n')?'\n\n':'';
  ta.value=v.slice(0,a)+pre+t+'\n\n'+v.slice(ta.selectionEnd);ins.selectedIndex=0;ta.focus();upd();});
 ta.addEventListener('input',upd);
 ta.addEventListener('keydown',function(e){if((e.ctrlKey||e.metaKey)&&!e.shiftKey){var k=e.key.toLowerCase();if(k==='b'){e.preventDefault();acc.b();}else if(k==='i'){e.preventDefault();acc.i();}else if(k==='u'){e.preventDefault();acc.u();}}});
 var pl=document.getElementById('dca-pl'),P=window.DCA_PLANTILLAS||{};
 if(pl)pl.addEventListener('change',function(){var p=P[pl.value]||(pl.value===''?null:{trato:'Señor(a)',nombre:'',lineas:''});if(!p)return;
  document.getElementById('dca-trato').value=p.trato;document.getElementById('dca-nom').value=p.nombre;document.getElementById('dca-lin').value=p.lineas;});
 upd();
})();
</script>"""


def _dca_html(d):
    url = _dca_url(d)
    qr = _dcx_qr_uri(url)
    emp = _dcx_empresa()
    dest = ["<p>%s</p>" % _esc(d.get("trato") or "")]
    if d.get("dest_nombre"):
        dest.append("<p><b>%s</b></p>" % _esc(d["dest_nombre"]).upper())
    dest += ["<p>%s</p>" % _esc(x) for x in _dca_lineas(d)]
    extra = []
    if d.get("anexos"):
        extra.append("<b>Anexos:</b> %s" % _esc(d["anexos"]))
    if d.get("cc"):
        extra.append("<b>Con copia a:</b> %s" % _esc(d["cc"]))
    return ("<div class='dca-hoja'><div class='dca-banda'></div>"
            "<div class='dca-top'><img class='lg' src='%s' alt='PROCSIS'><div class='dca-rad'><div>Al contestar cite Radicado <b>%s</b><br>Fecha: %s<br>Remitente: %s<br>"
            "Destinatario: %s<br>Anexos: %s</div><img src='%s' alt='QR'></div></div>"
            "<div class='dca-dest'>%s</div>%s<p class='dca-asunto'><b>Asunto:</b> %s</p>%s"
            "<div class='dca-cuerpo'>%s</div>%s"
            "<div class='dca-firma'><div class='l'></div><b>%s</b><br>%s<div class='ent'>%s</div></div>%s"
            "<div class='dca-foot'><div><b>Sede de correspondencia</b>%s, %s · PROCSIS<br>Tel: %s</div>"
            "<div><b>Servicio al cliente</b>%s<br>%s</div><div><b>Verifique este documento</b>%s<br>Escanee el QR · Página 1 de 1</div></div></div>") % (
        _DCX_LOGO, _esc(d.get("codigo")), _esc(_dca_hora(d)), _esc(d.get("entidad")), _esc(_dca_destinatario_txt(d)), _esc(d.get("anexos") or "Ninguno"), qr,
        "".join(dest), ("<p class='dca-ref'><b>Ref.:</b> %s</p>" % _esc(d["ref"])) if d.get("ref") else "", _esc(d.get("asunto")),
        ("<p>%s</p>" % _esc(d["saludo"])) if d.get("saludo") else "", _dca_cuerpo_html(d.get("cuerpo")),
        ("<p>%s</p>" % _esc(d["despedida"])) if d.get("despedida") else "",
        _esc(d.get("firmante")), (_esc(d["cargo"]) + "<br>") if d.get("cargo") else "", _esc(d.get("entidad")),
        ("<div class='dca-peq'>%s</div>" % "<br>".join(extra)) if extra else "",
        _esc(d.get("ciudad")), "Colombia", _esc(emp["tel"]), _esc(emp["email"]), _esc(emp["web"]), _esc(url))


# ───────────── PDF (reportlab) y Word (python-docx) con logo, QR y formato ─────────────
def _dca_ltr(s):
    """Texto compatible con la fuente base del PDF (Windows-1252): quita emojis y símbolos raros."""
    return (s or "").encode("cp1252", "ignore").decode("cp1252")


def _dca_pdf(d):
    from xml.sax.saxutils import escape as _x
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as _cv
    from reportlab.platypus import SimpleDocTemplate, Spacer, Image as _RLImage, HRFlowable
    base = getSampleStyleSheet()["Normal"]
    n = ParagraphStyle("dn", parent=base, fontName="Helvetica", fontSize=10.5, leading=14.5)
    just = ParagraphStyle("dj", parent=n, alignment=TA_JUSTIFY, spaceAfter=8)
    rad = ParagraphStyle("dr", parent=n, fontSize=7.6, leading=10, textColor=colors.HexColor("#334155"))
    cen = ParagraphStyle("dc", parent=n, alignment=TA_CENTER, spaceBefore=10, spaceAfter=12)
    tit = ParagraphStyle("dh", parent=n, fontName="Helvetica-Bold", fontSize=12, textColor=colors.HexColor("#0B2A6B"), spaceBefore=8, spaceAfter=5)
    vin = ParagraphStyle("dv", parent=just, leftIndent=18, bulletIndent=6, spaceAfter=3)
    pequeno = ParagraphStyle("dp", parent=n, fontSize=9, leading=12, textColor=colors.HexColor("#334155"))

    def t(s):
        return _x(_dca_ltr(s))

    def md(s):
        h = ""
        for txt, b, i, u in _dca_runs(s):
            x = t(txt).replace("\n", "<br/>")
            h += ("<b>%s</b>" if b else "%s") % ("<i>%s</i>" % x if i else x) if not u else "<u>%s</u>" % x
        return h
    url = _dca_url(d)
    el = []
    logo = _dca_logo_path()
    logo_img = ""
    if logo:
        try:
            w, h = ImageReader(logo).getSize()
            alto = 1.6 * cm
            logo_img = _RLImage(logo, width=alto * w / float(h), height=alto, hAlign="LEFT")
        except Exception:
            logo_img = ""
    qr = _RLImage(_dca_qr_png(url), width=2.0 * cm, height=2.0 * cm)
    radtxt = Paragraph("Al contestar cite Radicado <b>%s</b><br/>Fecha: %s<br/>Remitente: %s<br/>Destinatario: %s<br/>Anexos: %s" % (
        t(d.get("codigo")), t(_dca_hora(d)), t(d.get("entidad")), t(_dca_destinatario_txt(d)), t(d.get("anexos") or "Ninguno")), rad)
    cab = Table([[logo_img, radtxt, qr]], colWidths=[5.0 * cm, 8.6 * cm, 2.4 * cm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(cab)
    el.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cbd5e1"), spaceBefore=6, spaceAfter=12))
    el.append(Paragraph(t("%s, %s" % (d.get("ciudad"), _dcx_fecha_larga(d.get("fecha")))), n))
    el.append(Spacer(1, 8))
    el.append(Paragraph(t(d.get("trato")), n))
    if d.get("dest_nombre"):
        el.append(Paragraph("<b>%s</b>" % t((d["dest_nombre"] or "").upper()), n))
    for x in _dca_lineas(d):
        el.append(Paragraph(t(x), n))
    if d.get("ref"):
        el.append(Spacer(1, 6))
        el.append(Paragraph("<b>Ref.:</b> %s" % t(d["ref"]), n))
    el.append(Paragraph("<b>Asunto:</b> %s" % t(d.get("asunto")), cen))
    if d.get("saludo"):
        el.append(Paragraph(t(d["saludo"]), n))
        el.append(Spacer(1, 6))
    for b in _dca_bloques(d.get("cuerpo")):
        if b[0] == "p":
            el.append(Paragraph(md(b[1]), just))
        elif b[0] == "h":
            el.append(Paragraph(md(b[1]), tit))
        elif b[0] == "ul":
            el.append(Paragraph(md(b[1]), vin, bulletText="•"))
        else:
            el.append(Paragraph(md(b[1]), vin, bulletText=b[2] + "."))
    if d.get("despedida"):
        el.append(Spacer(1, 4))
        el.append(Paragraph(t(d["despedida"]), n))
    el.append(Spacer(1, 34))
    el.append(HRFlowable(width=6.2 * cm, thickness=0.8, color=colors.black, hAlign="LEFT", spaceAfter=3))
    el.append(Paragraph("<b>%s</b>" % t(d.get("firmante")), n))
    if d.get("cargo"):
        el.append(Paragraph(t(d["cargo"]), n))
    el.append(Paragraph("<b><font color='#0B2A6B'>%s</font></b>" % t(d.get("entidad")), n))
    if d.get("anexos") or d.get("cc"):
        el.append(Spacer(1, 12))
        if d.get("anexos"):
            el.append(Paragraph("<b>Anexos:</b> %s" % t(d["anexos"]), pequeno))
        if d.get("cc"):
            el.append(Paragraph("<b>Con copia a:</b> %s" % t(d["cc"]), pequeno))
    emp = _dcx_empresa()
    l1 = "Sede de correspondencia: %s, Colombia · Tel: %s  |  Servicio al cliente: %s · %s" % (d.get("ciudad") or "", emp["tel"], emp["email"], emp["web"])
    l2 = "Verifique este documento: %s" % url

    class _Num(_cv.Canvas):
        def __init__(self, *a, **k):
            _cv.Canvas.__init__(self, *a, **k)
            self._pags = []

        def showPage(self):
            self._pags.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pags)
            for est in self._pags:
                self.__dict__.update(est)
                self.setStrokeColor(colors.HexColor("#cbd5e1"))
                self.line(2.2 * cm, 1.75 * cm, letter[0] - 2.2 * cm, 1.75 * cm)
                self.setFont("Helvetica", 7.2)
                self.setFillColor(colors.HexColor("#475569"))
                self.drawString(2.2 * cm, 1.3 * cm, _dca_ltr(l1)[:150])
                self.drawString(2.2 * cm, 0.9 * cm, _dca_ltr(l2)[:120])
                self.drawRightString(letter[0] - 2.2 * cm, 0.9 * cm, "Página %d de %d" % (self._pageNumber, total))
                _cv.Canvas.showPage(self)
            _cv.Canvas.save(self)

    def _banda(c, doc):
        w = letter[0]
        for i, col in enumerate(("#00E676", "#00B8D4", "#0020C2")):
            c.setFillColor(colors.HexColor(col))
            c.rect(w * (0, 0.32, 0.66)[i], letter[1] - 0.42 * cm, w * (0.32, 0.34, 0.34)[i] + 1, 0.42 * cm, stroke=0, fill=1)
    buf = BytesIO()
    SimpleDocTemplate(buf, pagesize=letter, leftMargin=2.2 * cm, rightMargin=2.2 * cm, topMargin=1.5 * cm, bottomMargin=2.3 * cm,
                      title=d.get("codigo") or "Documento", author=d.get("entidad") or "PROCSIS").build(el, onFirstPage=_banda, canvasmaker=_Num)
    buf.seek(0)
    return buf


def _dca_docx(d):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1)
    sec.top_margin, sec.bottom_margin = Inches(0.7), Inches(0.9)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    url = _dca_url(d)
    cab = doc.add_table(rows=1, cols=2)
    c0, c1 = cab.rows[0].cells
    logo = _dca_logo_path()
    if logo:
        try:
            c0.paragraphs[0].add_run().add_picture(logo, height=Inches(0.65))
        except Exception:
            pass
    pr = c1.paragraphs[0]
    pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = pr.add_run("Al contestar cite Radicado ")
    r.font.size = Pt(8)
    r = pr.add_run(d.get("codigo") or "")
    r.bold = True
    r.font.size = Pt(8)
    r = pr.add_run("\nFecha: %s\nRemitente: %s\nDestinatario: %s\nAnexos: %s" % (_dca_hora(d), d.get("entidad") or "", _dca_destinatario_txt(d), d.get("anexos") or "Ninguno"))
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
    try:
        pr.add_run("\n").add_picture(_dca_qr_png(url), width=Inches(0.8))
    except Exception:
        pass
    sep = doc.add_paragraph("_" * 86)
    sep.runs[0].font.color.rgb = RGBColor(0x00, 0x20, 0xC2)
    sep.runs[0].font.size = Pt(8)

    def runs(p, texto, bold_all=False):
        for txt, b, i, u in _dca_runs(texto):
            rr = p.add_run(txt)
            rr.bold = bool(b or bold_all)
            rr.italic = i
            rr.underline = u

    def par(texto="", bold=False, align=None, after=0, sangria=None, vineta=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        if align is not None:
            p.alignment = align
        if sangria:
            p.paragraph_format.left_indent = Inches(sangria)
            p.paragraph_format.first_line_indent = Inches(-0.22)
        if vineta:
            p.add_run(vineta + " ")
        runs(p, texto, bold)
        return p
    par("%s, %s" % (d.get("ciudad") or "", _dcx_fecha_larga(d.get("fecha"))), after=10)
    par(d.get("trato") or "")
    if d.get("dest_nombre"):
        par((d["dest_nombre"] or "").upper(), bold=True)
    for x in _dca_lineas(d):
        par(x)
    if d.get("ref"):
        pref = doc.add_paragraph()
        pref.paragraph_format.space_before = Pt(8)
        pref.add_run("Ref.: ").bold = True
        pref.add_run(d["ref"])
    a = doc.add_paragraph()
    a.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a.paragraph_format.space_before = Pt(12)
    a.paragraph_format.space_after = Pt(12)
    a.add_run("Asunto: ").bold = True
    a.add_run(d.get("asunto") or "")
    if d.get("saludo"):
        par(d["saludo"], after=6)
    for b in _dca_bloques(d.get("cuerpo")):
        if b[0] == "p":
            par(b[1], align=WD_ALIGN_PARAGRAPH.JUSTIFY, after=8)
        elif b[0] == "h":
            p = par(b[1], bold=True, after=4)
            p.paragraph_format.space_before = Pt(8)
            for rr in p.runs:
                rr.font.size = Pt(12)
                rr.font.color.rgb = RGBColor(0x0B, 0x2A, 0x6B)
        elif b[0] == "ul":
            par(b[1], after=2, sangria=0.4, vineta="•")
        else:
            par(b[1], after=2, sangria=0.4, vineta=b[2] + ".")
    if d.get("despedida"):
        par(d["despedida"], after=4)
    par("")
    par("______________________________")
    par(d.get("firmante") or "", bold=True)
    if d.get("cargo"):
        par(d["cargo"])
    pe = par(d.get("entidad") or "", bold=True)
    for rr in pe.runs:
        rr.font.color.rgb = RGBColor(0x0B, 0x2A, 0x6B)
    if d.get("anexos"):
        pa = par("", after=0)
        pa.paragraph_format.space_before = Pt(10)
        pa.add_run("Anexos: ").bold = True
        pa.add_run(d["anexos"])
    if d.get("cc"):
        pc = par("")
        pc.add_run("Con copia a: ").bold = True
        pc.add_run(d["cc"])
    emp = _dcx_empresa()
    pie = doc.sections[0].footer.paragraphs[0]
    pie.add_run("Sede de correspondencia: %s, Colombia · Tel: %s · %s · %s\nVerifique este documento: %s   |   Página " % (
        d.get("ciudad") or "", emp["tel"], emp["email"], emp["web"], url))

    def campo(par_, instr):
        rr = par_.add_run()
        f1 = OxmlElement("w:fldChar")
        f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        f2 = OxmlElement("w:fldChar")
        f2.set(qn("w:fldCharType"), "end")
        rr._r.append(f1)
        rr._r.append(it)
        rr._r.append(f2)
    campo(pie, "PAGE")
    pie.add_run(" de ")
    campo(pie, "NUMPAGES")
    for rr in pie.runs:
        rr.font.size = Pt(7.5)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ───────────── Envío por correo (PDF adjunto) ─────────────
_DCA_EMAIL_RE = re.compile(r"^[^@\s<>\",;]+@[^@\s<>\",;]+\.[^@\s<>\",;]{2,}$")


def _dca_emails(txt):
    out = []
    for x in re.split(r"[,;\s]+", txt or ""):
        x = x.strip().lower()
        if x and _DCA_EMAIL_RE.match(x) and x not in out:
            out.append(x)
    return out


def _dca_enviar(d, para, cc, asunto, mensaje, usuario):
    """Envía el documento en PDF. Devuelve (ok, texto para el usuario)."""
    from email.utils import formataddr, parseaddr
    para, cc = _dca_emails(para), [c for c in _dca_emails(cc) if c not in _dca_emails(para)]
    if not para:
        return False, "Escribe al menos un correo válido en «Para»."
    if len(para) + len(cc) > 6:
        return False, "Máximo 6 destinatarios por envío."
    correo_envio, password = _credenciales_smtp("soporte")
    resend = (os.environ.get("RESEND_API_KEY") or "").strip()
    if not resend and (not correo_envio or not password):
        return False, "El correo de salida no está configurado. Conéctalo en Gerencia › Correo de soporte o define RESEND_API_KEY."
    base = correo_envio or (os.environ.get("RESEND_FROM") or "").strip()
    if resend and (os.environ.get("RESEND_FROM") or "").strip():
        base = os.environ["RESEND_FROM"].strip()
    _n, addr = parseaddr(base)
    entidad = _dca_una_linea(d.get("entidad"))[:80]
    asunto = _dca_una_linea(asunto)[:200] or _dca_una_linea(d.get("asunto"))[:200] or "Documento PROCSIS"
    mensaje = (mensaje or "").replace("\r", "").strip()[:2000] or ("Adjunto encontrará el documento %s emitido por %s." % (d.get("codigo"), entidad))
    nombre = re.sub(r"[^A-Za-z0-9_-]", "_", d.get("codigo") or "documento")
    url = _dca_url(d)
    html = ("<div style='font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:auto;color:#0b1437'>"
            "<div style='height:6px;background:linear-gradient(100deg,#00E676,#00B8D4 32%%,#0020C2)'></div>"
            "<p style='font-size:15px;line-height:1.6'>%s</p>"
            "<p style='font-size:13px;color:#475569'>Documento <b>%s</b> · %s<br>Verifique su autenticidad: <a href='%s'>%s</a></p>"
            "<p style='font-size:13px'><b>%s</b></p></div>") % (_esc(mensaje).replace("\n", "<br>"), _esc(d.get("codigo")), _esc(_dca_cat_txt(d.get("categoria"))),
                                                               _esc(url), _esc(url), _esc(entidad))
    texto = "%s\n\nDocumento %s · %s\nVerificación: %s\n\n%s" % (mensaje, d.get("codigo"), _dca_cat_txt(d.get("categoria")), url, entidad)
    ok, error = 0, ""
    try:
        msg = EmailMessage()
        msg["Subject"] = asunto
        msg["From"] = formataddr((entidad, addr)) if addr else base
        msg["To"] = ", ".join(para)
        if cc:
            msg["Cc"] = ", ".join(cc)
        msg.set_content(texto)
        msg.add_alternative(html, subtype="html")
        msg.add_attachment(_dca_pdf(d).read(), maintype="application", subtype="pdf", filename=nombre + ".pdf")
        _smtp_enviar(msg, correo_envio, password)
        ok = 1
    except Exception as e:
        error = repr(e)[:200]
        print("docs_area correo:", error)
    try:
        _c48_x("INSERT INTO docs_area_envios (doc_id, para, cc, asunto, por, fecha, ok, error) VALUES (:d, :p, :c, :a, :u, :f, :o, :e)",
               d=d["id"], p=", ".join(para), c=", ".join(cc), a=asunto, u=usuario, f=str(ahora())[:19], o=ok, e=error)
    except Exception:
        db.session.rollback()
    if ok:
        registrar_auditoria("Documentos por área", "%s envió %s por correo a %s" % (usuario, d.get("codigo"), ", ".join(para)))
        return True, "Documento enviado en PDF a %s." % ", ".join(para + cc)
    return False, "No se pudo enviar el correo. Revisa la conexión de correo e intenta de nuevo."


# ───────────── Páginas ─────────────
@app.route("/documentos")
def dca_lista():
    g = _dca_guard()
    if g is not None:
        return g
    rol, usuario = rol_actual(), session.get("usuario") or ""
    area = _dca_area_usuario()
    ger = rol in _DCX_GERENCIA
    vista = (request.args.get("v") or "recibidos").strip()
    if vista not in ("recibidos", "enviados", "todos") or (vista == "todos" and not ger):
        vista = "recibidos"
    cat = (request.args.get("cat") or "").strip()
    q = (request.args.get("q") or "").strip()[:60]
    cond, p = [], {}
    if vista == "recibidos":
        cond.append("destino=:ar")
        p["ar"] = area
    elif vista == "enviados":
        cond.append("(origen=:ar OR creado_por=:us)")
        p["ar"], p["us"] = area, usuario
    if cat in dict(_DCA_CATEGORIAS):
        cond.append("categoria=:ct")
        p["ct"] = cat
    if q:
        cond.append("(lower(codigo) LIKE :qq OR lower(asunto) LIKE :qq OR lower(dest_nombre) LIKE :qq)")
        p["qq"] = "%" + q.lower() + "%"
    rows = _c48_q("SELECT id, codigo, categoria, origen, destino, asunto, fecha, estado FROM docs_area" + (" WHERE " + " AND ".join(cond) if cond else "")
                  + " ORDER BY id DESC LIMIT 200", **p)
    sin_leer = _c48_q("SELECT COUNT(*) FROM docs_area WHERE destino=:ar AND estado='ENVIADO'", ar=area)
    n_sl = int(sin_leer[0][0]) if sin_leer else 0
    filas = ""
    for did, cod, ct, ori, des, asu, fe, es in rows:
        etq = {"ENVIADO": "wait", "LEIDO": "on", "ATENDIDO": "on"}.get(es, "off")
        txt = {"ENVIADO": "Sin leer", "LEIDO": "Leído", "ATENDIDO": "Atendido", "ANULADO": "Anulado"}.get(es, es)
        resalta = " class='dca-sl'" if (es == "ENVIADO" and des == area) else ""
        filas += ("<tr%s><td><b>%s</b></td><td>%s</td><td>%s</td><td>%s → %s</td><td>%s</td><td><span class='cb-tag %s'>%s</span></td>"
                  "<td><a class='cb-btn s' href='/documentos/%d'>Abrir</a></td></tr>") % (
            resalta, _esc(cod), _esc(_dca_cat_txt(ct)), _esc((asu or "")[:80]), _esc(_dca_entidad_def(ori)),
            _esc(_dcx_area_titulo(des) if des else "Correo externo"), _esc(fe), etq, txt, did)
    if not filas:
        filas = "<tr><td colspan='7'>No hay documentos en esta vista.</td></tr>"
    tabs = "".join("<a class='cb-btn s %s' href='/documentos?v=%s'>%s</a> " % ("" if vista == k else "o", k, n)
                   for k, n in (("recibidos", "Recibidos%s" % (" (%d sin leer)" % n_sl if n_sl else "")), ("enviados", "Enviados")) + ((("todos", "Todos"),) if ger else ()))
    opts = "<option value=''>Todos los tipos</option>" + "".join("<option value='%s'%s>%s</option>" % (k, " selected" if k == cat else "", n) for k, n in _DCA_CATEGORIAS)
    filtro = ("<form method='GET' action='/documentos' style='display:inline-flex;gap:6px;flex-wrap:wrap'><input type='hidden' name='v' value='%s'>"
              "<select name='cat'>%s</select><input type='text' name='q' value='%s' placeholder='Código, asunto o destinatario' style='width:220px'>"
              "<button class='cb-btn s'>Filtrar</button></form>") % (vista, opts, _esc(q))
    cuerpo = ("<h1>📨 Documentos por área</h1><p class='sub'>Redacta con formato, firma como tu área (%s), envía al área que corresponda o por correo en PDF.</p>"
              "<div class='dca-bar'><a class='cb-btn g' href='/documentos/nuevo'>➕ Nuevo documento</a></div>"
              "<div class='cb-card'><div class='dca-bar'>%s%s</div><div class='cb-scroll'><table><tr><th>Código</th><th>Tipo</th><th>Asunto</th>"
              "<th>De → Para</th><th>Fecha</th><th>Estado</th><th></th></tr>%s</table></div></div>") % (_esc(_dca_entidad_def(area)), tabs, filtro, filas)
    return _cb_pagina("Documentos por área", _DCA_CSS + cuerpo)


def _dca_form(csrf):
    area = _dca_area_usuario()
    nombre = ""
    try:
        nombre = _nombre_para_saludo(session.get("usuario") or "") or ""
    except Exception:
        pass
    emp = _dcx_empresa()
    cats = "".join("<option value='%s'>%s</option>" % (k, n) for k, n in _DCA_CATEGORIAS)
    dests = ("<option value=''>— Ninguna: solo por correo —</option>"
             + "".join("<option value='%s'%s>%s</option>" % (_esc(a), " selected" if a == "Gerencia" else "", _esc(_dcx_area_titulo(a))) for a in _dca_areas()))
    plant = ("<option value=''>En blanco</option><option value='comunidad'>Comunidad EduTrack (directores, rectores y aliados)</option>"
             "<option value='tiendas'>Comercios y restaurantes (TiendaTrack)</option><option value='interna'>Documento interno entre áreas</option>")
    snippets = [("Cierre cordial", "Agradecemos su atención y quedamos atentos a cualquier inquietud."),
                ("Plazo de respuesta", "Agradecemos su respuesta dentro de los **cinco (5) días hábiles** siguientes al recibo de esta comunicación."),
                ("Confidencialidad", "La información contenida en este documento es **confidencial** y está dirigida exclusivamente a su destinatario. "
                                     "Si lo recibió por error, informe al remitente y elimínelo."),
                ("Datos de contacto", "Para mayor información puede comunicarse con nosotros:\n- Correo: %s\n- Teléfono: %s\n- Web: %s" % (emp["email"], emp["tel"], emp["web"]))]
    ins = "<select id='dca-ins'><option value=''>＋ Texto rápido…</option>%s</select>" % "".join(
        "<option value=\"%s\">%s</option>" % (_esc(t).replace("\n", "&#10;"), _esc(n)) for n, t in snippets)
    tb = ("<div class='dca-tb' id='dca-tb'><button type='button' data-k='b' title='Negrita (Ctrl+B)'><b>B</b></button>"
          "<button type='button' data-k='i' title='Cursiva (Ctrl+I)'><i>I</i></button><button type='button' data-k='u' title='Subrayado (Ctrl+U)'><u>U</u></button>"
          "<button type='button' data-k='h' title='Título'>Título</button><button type='button' data-k='ul' title='Lista con viñetas'>• Lista</button>"
          "<button type='button' data-k='ol' title='Lista numerada'>1. Lista</button>%s</div>") % ins
    return ("<h1>Nuevo documento</h1><p class='sub'>Se emite con el logo de PROCSIS, un código único y QR de verificación, y queda firmado por tu área. "
            "Un documento emitido no se edita: si hay un error, se anula y se emite otro.</p>"
            "<form method='POST' action='/documentos/nuevo'><input type='hidden' name='csrf' value='%s'>"
            "<div class='cb-card'><div class='dca-fg'>"
            "<div><label>Tipo de documento</label><select name='categoria'>%s</select></div>"
            "<div><label>Enviar a (área dentro de PROCSIS)</label><select name='destino'>%s</select></div>"
            "<div><label>Ciudad</label><input type='text' name='ciudad' value='Caracolí, Antioquia' maxlength='100'></div>"
            "<div><label>Fecha</label><input type='date' name='fecha' value='%s'></div>"
            "<div class='dca-full'><label>Plantilla de destinatario</label><select id='dca-pl'>%s</select></div>"
            "<div><label>Tratamiento</label><input type='text' id='dca-trato' name='trato' value='Señor(a)' maxlength='60'></div>"
            "<div><label>Destinatario (se muestra en negrita)</label><input type='text' id='dca-nom' name='dest_nombre' maxlength='220'></div>"
            "<div class='dca-full'><label>Otras líneas del destinatario (una por línea)</label><textarea id='dca-lin' name='dest_lineas' rows='3'></textarea></div>"
            "<div><label>Referencia (opcional)</label><input type='text' name='ref' maxlength='220' placeholder='Ej.: Contrato 045 de 2026'></div>"
            "<div><label>Asunto</label><input type='text' name='asunto' maxlength='600' required></div>"
            "<div><label>Saludo</label><input type='text' name='saludo' value='Cordial saludo:' maxlength='160'></div>"
            "<div><label>Despedida</label><input type='text' name='despedida' value='Atentamente,' maxlength='160'></div>"
            "<div class='dca-full'><label>Cuerpo del documento</label>%s<textarea id='dca-cuerpo' name='cuerpo' rows='12' required></textarea>"
            "<div class='cb-note'>Negrita **así**, cursiva *así*, subrayado __así__. «# » al inicio = título, «- » = viñeta, «1. » = lista numerada.</div>"
            "<div class='dca-prev' id='dca-prev'></div></div>"
            "<div><label>Firmante (nombre)</label><input type='text' name='firmante' value='%s' maxlength='160' required></div>"
            "<div><label>Cargo del firmante</label><input type='text' name='cargo' maxlength='160'></div>"
            "<div class='dca-full'><label>Firma como (entidad que emite)</label><input type='text' name='entidad' value='%s' maxlength='220'>"
            "<div class='cb-note'>Aparece al final del documento y como remitente del correo.</div></div>"
            "<div><label>Anexos (opcional)</label><input type='text' name='anexos' maxlength='300' placeholder='Ej.: 2 archivos PDF'></div>"
            "<div><label>Con copia a (opcional, solo texto en el documento)</label><input type='text' name='cc' maxlength='300'></div>"
            "<div class='dca-full dca-env'><b>✉ Enviar también por correo en PDF (opcional)</b><div class='dca-fg' style='margin-top:8px'>"
            "<div><label>Para (correo)</label><input type='text' name='correo_para' placeholder='persona@empresa.com'></div>"
            "<div><label>Con copia (correos, separados por coma)</label><input type='text' name='correo_cc'></div>"
            "<div class='dca-full'><label>Mensaje del correo (opcional)</label><textarea name='correo_msg' rows='2' maxlength='2000' placeholder='Adjunto encontrará…'></textarea></div></div></div>"
            "</div><p style='margin-top:14px'><button class='cb-btn g'>Emitir documento</button> <a class='cb-btn o' href='/documentos'>Cancelar</a></p></div></form>"
            "<script>window.DCA_PLANTILLAS=%s;</script>%s") % (
        _esc(csrf), cats, dests, fecha_hoy(), plant, tb, _esc(nombre), _esc(_dca_entidad_def(area)),
        json.dumps(_DCA_PLANTILLAS_DEST, ensure_ascii=False).replace("</", "<\\/"), _DCA_JS_EDITOR)


@app.route("/documentos/nuevo", methods=["GET", "POST"])
def dca_nuevo():
    g = _dca_guard()
    if g is not None:
        return g
    if request.method == "GET":
        return _cb_pagina("Nuevo documento", _DCA_CSS + _dca_form(_acc_csrf()))
    f = request.form
    if not _ti_csrf_ok():
        session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
        return redirect("/documentos/nuevo")

    def lim(k, n):
        return _dca_una_linea(f.get(k))[:n]
    categoria, destino = lim("categoria", 30), lim("destino", 60)
    asunto, firmante = lim("asunto", 600), lim("firmante", 160)
    cuerpo = (f.get("cuerpo") or "").replace("\r", "").strip()[:9000]
    correo_para = (f.get("correo_para") or "").strip()
    if categoria not in dict(_DCA_CATEGORIAS) or (destino and destino not in _dca_areas()) or (not destino and not correo_para):
        session["ti_msg"] = ("err", "Elige el tipo de documento y el área destino, o escribe un correo para enviarlo.")
        return redirect("/documentos/nuevo")
    if not asunto or not cuerpo or not firmante:
        session["ti_msg"] = ("err", "Faltan el asunto, el cuerpo o el nombre de quien firma.")
        return redirect("/documentos/nuevo")
    fecha = lim("fecha", 10)
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", fecha):
        fecha = fecha_hoy()
    anio = int(fecha[:4])
    origen = _dca_area_usuario()
    r = _c48_q("SELECT COALESCE(MAX(numero), 0) FROM docs_area WHERE origen=:o AND anio=:y", o=origen, y=anio)
    num = int((r[0][0] if r else 0) or 0) + 1
    codigo = "DOC-%s-%d-%04d" % (_DCX_COD.get(origen, "GEN"), anio, num)
    usuario = session.get("usuario") or ""
    token = secrets.token_urlsafe(12)
    try:
        _c48_x("INSERT INTO docs_area (codigo, categoria, origen, destino, asunto, trato, dest_nombre, dest_lineas, saludo, cuerpo, despedida, firmante, cargo, "
               "ciudad, fecha, estado, creado_por, creado, numero, anio, entidad, ref, cc, anexos, token) VALUES (:codigo, :categoria, :origen, :destino, "
               ":asunto, :trato, :dn, :dl, :saludo, :cuerpo, :desp, :firmante, :cargo, :ciudad, :fecha, :estado, :us, :cr, :num, :anio, :entidad, :ref, :cc, :anexos, :token)",
               codigo=codigo, categoria=categoria, origen=origen, destino=destino, asunto=asunto, trato=lim("trato", 60), dn=lim("dest_nombre", 220),
               dl=(f.get("dest_lineas") or "").replace("\r", "").strip()[:900], saludo=lim("saludo", 160), cuerpo=cuerpo, desp=lim("despedida", 160),
               firmante=firmante, cargo=lim("cargo", 160), ciudad=lim("ciudad", 100), fecha=fecha, estado="ENVIADO" if destino else "LEIDO", us=usuario,
               cr=str(ahora())[:19], num=num, anio=anio, entidad=lim("entidad", 220) or _dca_entidad_def(origen), ref=lim("ref", 220), cc=lim("cc", 300),
               anexos=lim("anexos", 300), token=token)
    except Exception as e:
        db.session.rollback()
        print("docs_area crear:", e)
        session["ti_msg"] = ("err", "No se pudo emitir el documento. Intenta de nuevo.")
        return redirect("/documentos/nuevo")
    d = _dca_row(token=token)
    registrar_auditoria("Documentos por área", "%s emitió %s para %s (%s)" % (usuario, codigo, destino or "correo externo", categoria))
    msg = "Documento %s emitido%s." % (codigo, (" y enviado a " + _dcx_area_titulo(destino)) if destino else "")
    if correo_para and d:
        ok_c, txt = _dca_enviar(d, correo_para, f.get("correo_cc") or "", asunto, f.get("correo_msg") or "", usuario)
        msg += " " + txt
        session["ti_msg"] = ("ok" if ok_c else "err", msg)
    else:
        session["ti_msg"] = ("ok", msg)
    return redirect("/documentos/%d" % d["id"]) if d else redirect("/documentos")


@app.route("/documentos/<int:did>")
def dca_ver(did):
    g = _dca_guard()
    if g is not None:
        return g
    rol, usuario = rol_actual(), session.get("usuario") or ""
    d = _dca_row(did)
    if not d or not _dca_puede(d, rol, usuario):
        return acceso_denegado("Documento no disponible.")
    area = _dcx_area_de_rol(rol)
    if not d["token"]:
        d["token"] = secrets.token_urlsafe(12)
        try:
            _c48_x("UPDATE docs_area SET token=:t WHERE id=:i AND (token IS NULL OR token='')", t=d["token"], i=did)
        except Exception:
            db.session.rollback()
        d = _dca_row(did) or d
    if d["destino"] and d["destino"] == area and d["creado_por"] != usuario and d["estado"] == "ENVIADO":
        try:
            _c48_x("UPDATE docs_area SET estado='LEIDO', leido_por=:u, leido=:t WHERE id=:i AND estado='ENVIADO'", u=usuario, t=str(ahora())[:19], i=did)
            d["estado"] = "LEIDO"
        except Exception:
            db.session.rollback()
    csrf = _acc_csrf()
    acciones = ""
    if d["estado"] in ("ENVIADO", "LEIDO") and d["destino"] and (d["destino"] == area or rol in _DCX_GERENCIA):
        acciones += ("<form method='POST' action='/documentos/%d/estado' style='display:inline'><input type='hidden' name='csrf' value='%s'>"
                     "<input type='hidden' name='accion' value='atendido'><button class='cb-btn g s'>✔ Marcar como atendido</button></form> ") % (did, _esc(csrf))
    if d["estado"] != "ANULADO" and (rol in _DCX_GERENCIA or d["creado_por"] == usuario):
        acciones += ("<form method='POST' action='/documentos/%d/estado' style='display:inline-flex;gap:6px' onsubmit=\"return confirm('¿Anular este documento?');\">"
                     "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='anular'>"
                     "<input type='text' name='nota' placeholder='Motivo' required maxlength='200' style='width:200px'><button class='cb-btn r s'>Anular</button></form>") % (did, _esc(csrf))
    est = {"ENVIADO": "<span class='cb-tag wait'>Enviado · sin leer</span>", "LEIDO": "<span class='cb-tag on'>Leído%s</span>" % ((" por " + _esc(d.get("leido_por"))) if d.get("leido_por") else ""),
           "ATENDIDO": "<span class='cb-tag on'>Atendido</span>", "ANULADO": "<span class='cb-tag off'>Anulado · %s</span>" % _esc(d.get("nota") or "")}.get(d["estado"], "")
    envio = ""
    if d["estado"] != "ANULADO" and _dca_puede_enviar(d, rol, usuario):
        envio = ("<details class='dca-env' open><summary><b>✉ Enviar por correo en PDF</b></summary><form method='POST' action='/documentos/%d/enviar' style='margin-top:10px'>"
                 "<input type='hidden' name='csrf' value='%s'><div class='dca-fg'>"
                 "<div><label>Para (correo)</label><input type='text' name='para' required placeholder='persona@empresa.com'></div>"
                 "<div><label>Con copia (separados por coma)</label><input type='text' name='cc'></div>"
                 "<div class='dca-full'><label>Asunto del correo</label><input type='text' name='asunto' value='%s' maxlength='200'></div>"
                 "<div class='dca-full'><label>Mensaje</label><textarea name='msg' rows='2' maxlength='2000' placeholder='Adjunto encontrará el documento %s emitido por %s.'></textarea></div>"
                 "</div><p style='margin-top:10px'><button class='cb-btn g'>Enviar PDF</button> <span class='cb-note'>Sale como «%s» con el PDF adjunto.</span></p></form></details>"
                 ) % (did, _esc(csrf), _esc(d["asunto"][:200]), _esc(d["codigo"]), _esc(d["entidad"]), _esc(d["entidad"]))
    hist = _c48_q("SELECT fecha, para, cc, por, ok FROM docs_area_envios WHERE doc_id=:d ORDER BY id DESC LIMIT 15", d=did)
    historial = ""
    if hist:
        historial = ("<div class='cb-card cb-scroll' style='margin-top:14px'><h2>Envíos por correo</h2><table><tr><th>Fecha</th><th>Para</th><th>Con copia</th><th>Por</th><th>Resultado</th></tr>%s</table></div>"
                     % "".join("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td><span class='cb-tag %s'>%s</span></td></tr>" % (
                         _esc(f_), _esc(p_), _esc(c_ or "—"), _esc(u_), "on" if o_ else "off", "Enviado" if o_ else "Falló") for f_, p_, c_, u_, o_ in hist))
    barra = ("<div class='dca-bar'><a class='cb-btn' href='/documentos/%d.pdf'>⬇ Descargar PDF</a><a class='cb-btn o' href='/documentos/%d.docx'>📝 Descargar Word</a>"
             "<a class='cb-btn o' href='/documentos'>← Volver</a></div><div class='dca-bar'>%s %s</div>") % (did, did, est, acciones)
    return _cb_pagina(d["codigo"], _DCA_CSS + "<h1>%s</h1><p class='sub'>%s · de %s para %s</p>%s%s%s%s" % (
        _esc(d["codigo"]), _esc(_dca_cat_txt(d["categoria"])), _esc(d["entidad"]), _esc(_dcx_area_titulo(d["destino"]) if d["destino"] else "correo externo"),
        barra, envio, _dca_html(d), historial))


@app.route("/documentos/<int:did>/enviar", methods=["POST"])
def dca_enviar_correo(did):
    g = _dca_guard()
    if g is not None:
        return g
    rol, usuario = rol_actual(), session.get("usuario") or ""
    d = _dca_row(did)
    if not _ti_csrf_ok() or not d or not _dca_puede(d, rol, usuario) or not _dca_puede_enviar(d, rol, usuario) or d["estado"] == "ANULADO":
        return redirect("/documentos")
    ok_rl, _e = _rate_limit("docs-area-correo", 10, 60)
    if not ok_rl:
        session["ti_msg"] = ("err", "Demasiados envíos seguidos. Espera un minuto.")
        return redirect("/documentos/%d" % did)
    ok, txt = _dca_enviar(d, request.form.get("para") or "", request.form.get("cc") or "", request.form.get("asunto") or "", request.form.get("msg") or "", usuario)
    session["ti_msg"] = ("ok" if ok else "err", txt)
    return redirect("/documentos/%d" % did)


def _dca_descarga(did, formato):
    g = _dca_guard()
    if g is not None:
        return g
    d = _dca_row(did)
    if not d or not _dca_puede(d, rol_actual(), session.get("usuario") or ""):
        return acceso_denegado("Documento no disponible.")
    if not d["token"]:
        d["token"] = secrets.token_urlsafe(12)
        try:
            _c48_x("UPDATE docs_area SET token=:t WHERE id=:i AND (token IS NULL OR token='')", t=d["token"], i=did)
            d = _dca_row(did) or d
        except Exception:
            db.session.rollback()
    nombre = re.sub(r"[^A-Za-z0-9_-]", "_", d["codigo"] or "documento")
    if formato == "pdf":
        return send_file(_dca_pdf(d), mimetype="application/pdf", as_attachment=True, download_name=nombre + ".pdf")
    return send_file(_dca_docx(d), as_attachment=True, download_name=nombre + ".docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/documentos/<int:did>.pdf")
def dca_pdf(did):
    return _dca_descarga(did, "pdf")


@app.route("/documentos/<int:did>.docx")
def dca_docx(did):
    return _dca_descarga(did, "docx")


@app.route("/documentos/<int:did>/estado", methods=["POST"])
def dca_estado(did):
    g = _dca_guard()
    if g is not None:
        return g
    rol, usuario = rol_actual(), session.get("usuario") or ""
    d = _dca_row(did)
    if not _ti_csrf_ok() or not d or not _dca_puede(d, rol, usuario):
        return redirect("/documentos")
    accion = (request.form.get("accion") or "").strip()
    area = _dcx_area_de_rol(rol)
    try:
        if accion == "atendido" and d["estado"] in ("ENVIADO", "LEIDO") and d["destino"] and (d["destino"] == area or rol in _DCX_GERENCIA):
            _c48_x("UPDATE docs_area SET estado='ATENDIDO', leido_por=:u, leido=:t WHERE id=:i", u=usuario, t=str(ahora())[:19], i=did)
            registrar_auditoria("Documentos por área", "%s atendió %s" % (usuario, d["codigo"]))
            session["ti_msg"] = ("ok", "Documento marcado como atendido.")
        elif accion == "anular" and d["estado"] != "ANULADO" and (rol in _DCX_GERENCIA or d["creado_por"] == usuario):
            motivo = _dca_una_linea(request.form.get("nota"))[:200]
            if motivo:
                _c48_x("UPDATE docs_area SET estado='ANULADO', nota=:n WHERE id=:i", n=motivo, i=did)
                registrar_auditoria("Documentos por área", "%s anuló %s" % (usuario, d["codigo"]))
                session["ti_msg"] = ("ok", "Documento anulado.")
    except Exception:
        db.session.rollback()
        session["ti_msg"] = ("err", "No se pudo actualizar el documento.")
    return redirect("/documentos/%d" % did)


@app.route("/docs/area/<token>")
def dca_verificar(token):
    """Verificación pública (por QR): confirma que el documento existe y su estado, sin mostrar su contenido."""
    _dca_init()
    d = _dca_row(token=token) if re.match(r"^[A-Za-z0-9_-]{8,40}$", token or "") else None
    estilo = ("<style>.vf{max-width:560px;margin:34px auto;padding:0 16px;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#0b1437}"
              ".vf-c{background:#fff;border:1px solid #e3e8f7;border-radius:18px;padding:24px;box-shadow:0 14px 40px rgba(0,32,194,.10)}"
              ".vf-b{height:6px;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2);margin:-24px -24px 18px;border-radius:18px 18px 0 0}"
              ".vf h1{font-size:22px;margin:6px 0 4px}.vf-ok{color:#166534}.vf-no{color:#b91c1c}"
              ".vf dl{margin:16px 0 0;display:grid;grid-template-columns:130px 1fr;gap:8px 12px;font-size:14px}"
              ".vf dt{color:#64748b}.vf dd{margin:0;font-weight:600}.vf small{display:block;margin-top:16px;color:#64748b;line-height:1.5}.vf img{height:44px}</style>")
    if not d:
        return page("Verificación de documento", estilo + "<div class='vf'><div class='vf-c'><div class='vf-b'></div><img src='%s' alt='PROCSIS'>"
                    "<h1 class='vf-no'>⚠ Documento no encontrado</h1><p>Este código no corresponde a ningún documento emitido por PROCSIS. "
                    "Si recibiste este documento, desconfía de su autenticidad.</p></div></div>" % _DCX_LOGO)
    vig = d["estado"] != "ANULADO"
    return page("Verificación de documento", estilo + (
        "<div class='vf'><div class='vf-c'><div class='vf-b'></div><img src='%s' alt='PROCSIS'><h1 class='%s'>%s</h1><dl>"
        "<dt>Código</dt><dd>%s</dd><dt>Tipo</dt><dd>%s</dd><dt>Emitido por</dt><dd>%s</dd><dt>Fecha</dt><dd>%s</dd><dt>Dirigido a</dt><dd>%s</dd></dl>"
        "<small>Por seguridad, esta página no muestra el contenido del documento. Compáralo con el que recibiste.</small></div></div>") % (
        _DCX_LOGO, "vf-ok" if vig else "vf-no", "✔ Documento auténtico y vigente" if vig else "⚠ Documento anulado", _esc(d["codigo"]), _esc(_dca_cat_txt(d["categoria"])),
        _esc(d["entidad"]), _esc(_dcx_fecha_larga(d["fecha"])), _esc(_dca_destinatario_txt(d))))


'''

# ───────────── Menú lateral ─────────────
ANCLA_GER = "        ('/documentos-oficiales', '📄 Documentos oficiales (todas las áreas)'),\n"
NUEVO_GER = ANCLA_GER + "        ('/documentos', '📨 Documentos por área (PDF, Word y correo)'),  # [procsis-documentos-area]\n"
ANCLA_ROL = '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n'
NUEVO_ROL = ('    if _dca_roles_ok(rol) and not any(h == "/documentos" for h, _l in items):  # [procsis-documentos-area]\n'
             '        items = [x for x in items if x[0] != "/logout"] + [("/documentos", "📨 Documentos por área")] + [x for x in items if x[0] == "/logout"]\n' + ANCLA_ROL)
YA_MENU_GER = "('/documentos', '📨 Documentos por área"
YA_MENU_ROL = 'if _dca_roles_ok(rol) and not any(h == "/documentos" for h, _l in items):'

ANCLAS = ["def _dcx_rol_ok", "def _dcx_areas_validas", "def _dcx_area_de_rol", "def _dcx_area_titulo", "def _dcx_fecha_larga", "def _dcx_empresa", "def _dcx_base",
          "def _dcx_qr_uri", "def _cb_pagina", "def _c48_q", "def _c48_x", "def _acc_csrf", "def _ti_csrf_ok", "def registrar_auditoria", "def acceso_denegado",
          "def fecha_hoy", "def _nombre_para_saludo", "def _credenciales_smtp", "def _smtp_enviar", "def _rate_limit", "def page(", "_DCX_LOGO = ", "_DCX_COD = ",
          "from docx import Document", "from reportlab.platypus import Table, TableStyle, Paragraph", "from email.message import EmailMessage", "import qrcode"]


def quitar_bloque_v1(texto):
    """Quita el bloque de la versión 1 (si está) y devuelve el texto sin él."""
    i = texto.find(MARCA_V1)
    if i < 0:
        return texto
    j = len(texto)
    for fin in ("\n# [procsis-", '\nif __name__ == "__main__":'):
        k = texto.find(fin, i + len(MARCA_V1))
        if k >= 0:
            j = min(j, k)
    return texto[:i].rstrip("\n") + "\n\n" + texto[j:].lstrip("\n")


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
    if not ruta.is_file():
        sys.exit("No encuentro %s. Ejecuta el script en la carpeta de tu app.py o pásale la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as fh:
        crudo = fh.read()
    crlf = "\r\n" in crudo
    texto = crudo.replace("\r\n", "\n")
    if MARCA in texto:
        print("Ya estaba aplicado (%s). No se cambió nada." % MARCA)
        return
    faltan = [a for a in ANCLAS if a not in texto]
    reemplazos = []
    if YA_MENU_GER not in texto:
        reemplazos.append(("menú de Gerencia", ANCLA_GER, NUEVO_GER))
    if YA_MENU_ROL not in texto:
        reemplazos.append(("menú de las demás áreas", ANCLA_ROL, NUEVO_ROL))
    for desc, viejo, _n in reemplazos:
        if texto.count(viejo) != 1:
            faltan.append("%s (aparece %d veces, debe ser 1)" % (desc, texto.count(viejo)))
    if faltan:
        sys.exit("Este app.py no coincide con lo esperado; no toqué nada. Falta/cambió:\n  - " + "\n  - ".join(faltan))
    nuevo = quitar_bloque_v1(texto)
    reemplazada = nuevo != texto
    for _d, viejo, reemplazo in reemplazos:
        nuevo = nuevo.replace(viejo, reemplazo, 1)
    corte = nuevo.rfind('\nif __name__ == "__main__":')
    if corte < 0:
        sys.exit('No encontré el bloque if __name__ == "__main__": al final de app.py.')
    nuevo = nuevo[:corte] + "\n" + BLOQUE.strip("\n") + "\n\n" + nuevo[corte:]
    try:
        compile(nuevo, str(ruta), "exec")
    except SyntaxError as e:
        sys.exit("El resultado no compila (línea %s: %s). No escribí nada." % (e.lineno, e.msg))
    copia = ruta.with_name(ruta.name + ".bak_documentos_area2")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_documentos_area2_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. Documentos por área v2 %s en %s" % ("(reemplazó la versión 1)" if reemplazada else "agregado", ruta))
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()

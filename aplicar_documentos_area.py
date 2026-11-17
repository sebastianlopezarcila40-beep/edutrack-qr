#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_documentos_area.py  ·  PROCSIS: módulo «Documentos por área» (PDF y Word con logo) para todas las áreas.

Uso (en la carpeta de tu app.py):
    python aplicar_documentos_area.py            # parchea ./app.py
    python aplicar_documentos_area.py ruta/app.py

Qué agrega (ruta /documentos, aparece en el menú de todas las áreas):
  · Crear un documento con el logo de PROCSIS y el encabezado de tu imagen:
        Señor(a)S / DESTINATARIO en negrita / líneas del destinatario / «Asunto:» centrado / cuerpo / firma.
  · Selector «Enviar a» (área que lo recibe: Gerencia, Tecnología y TI, Soporte, Comercial, Cobranza, Contabilidad, Finanzas,
    Talento, Marketing, Jurídica, Soluciones, Calidad, Operaciones, CRTP, Desarrollo).
  · Selector «Tipo de documento» (queja, reclamo, petición, financiero, informe técnico, notificación, certificación, ...).
  · Descarga en PDF y en Word (.docx), ambos con logo.
  · Bandeja por área: Recibidos (con sin leer), Enviados y, para Gerencia, Todos. El área que recibe lo marca como atendido.
  · Código único por documento: DOC-<ÁREA>-<AÑO>-0001. Un documento no se edita; se anula y se emite otro.
Es independiente de «Documentos oficiales» (oficios, resoluciones y comunicados), que sigue igual.
Es idempotente, crea app.py.bak_documentos_area y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-documentos-area]"

BLOQUE = r'''
# [procsis-documentos-area]
# ═════════════ DOCUMENTOS POR ÁREA: crear PDF / Word con logo, enviar a un área y llevar bandeja ═════════════
_DCA = {"ok": False}
_DCA_CATEGORIAS = [("queja", "Queja"), ("reclamo", "Reclamo"), ("peticion", "Petición o solicitud"), ("financiero", "Documento financiero"),
                   ("informe", "Informe técnico"), ("notificacion", "Notificación o circular"), ("certificacion", "Certificación"),
                   ("requerimiento", "Requerimiento interno"), ("acta", "Acta o constancia"), ("otro", "Otro")]
_DCA_ROLES_EXTRA = ("Soporte", "Cobranza", "Comercial", "Ventas", "Supervisor de Ventas", "Soluciones", "Desarrollador", "Developer")
_DCA_COLS = ("id, codigo, categoria, origen, destino, asunto, trato, dest_nombre, dest_lineas, saludo, cuerpo, despedida, firmante, cargo, "
             "ciudad, fecha, estado, creado_por, creado, leido_por, leido, nota, numero, anio")
_DCA_PLANTILLAS_DEST = {
    "comunidad": {"trato": "Señor(a)S", "nombre": "DIRECTORES, RECTORES Y ALIADOS ESTRATÉGICOS",
                  "lineas": "Comunidad Educativa EduTrack Colombia\nInstituciones Usuarias de PROCSIS y EduTrack\nEcosistema Nacional"},
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


def _dca_init():
    if _DCA["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS docs_area (id %s, codigo VARCHAR(40), categoria VARCHAR(30), origen VARCHAR(60), destino VARCHAR(60), "
            "asunto TEXT, trato VARCHAR(60), dest_nombre VARCHAR(220), dest_lineas TEXT DEFAULT '', saludo VARCHAR(160) DEFAULT '', "
            "cuerpo TEXT, despedida VARCHAR(160) DEFAULT '', firmante VARCHAR(160), cargo VARCHAR(160) DEFAULT '', ciudad VARCHAR(100), "
            "fecha VARCHAR(10), estado VARCHAR(12) DEFAULT 'ENVIADO', creado_por VARCHAR(80), creado VARCHAR(19), "
            "leido_por VARCHAR(80) DEFAULT '', leido VARCHAR(19) DEFAULT '', nota TEXT DEFAULT '', numero INTEGER DEFAULT 0, anio INTEGER DEFAULT 0)" % pk))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_docs_area_destino ON docs_area (destino)"))
        db.session.execute(text("CREATE INDEX IF NOT EXISTS ix_docs_area_origen ON docs_area (origen)"))
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


def _dca_row(did):
    rows = _c48_q("SELECT %s FROM docs_area WHERE id=:i" % _DCA_COLS, i=did)
    return dict(zip([c.strip() for c in _DCA_COLS.split(",")], rows[0])) if rows else None


def _dca_puede(d, rol, usuario):
    return rol in _DCX_GERENCIA or d.get("origen") == _dcx_area_de_rol(rol) or d.get("destino") == _dcx_area_de_rol(rol) or d.get("creado_por") == usuario


def _dca_logo_path():
    for p in (Path(app.root_path) / "static" / "img" / "logo-procsis-mark.png", Path(app.root_path) / "static" / "img" / "logo-procsis.png"):
        if p.is_file():
            return str(p)
    return ""


def _dca_lineas(d):
    return [x.strip() for x in (d.get("dest_lineas") or "").replace("\r", "").split("\n") if x.strip()]


def _dca_parrafos(t):
    return [p.strip() for p in re.split(r"\n\s*\n", (t or "").replace("\r", "")) if p.strip()]


# ───────────── Vista previa (HTML) ─────────────
_DCA_CSS = r"""<style>
.dca-hoja{background:#fff;max-width:794px;margin:0 auto;padding:44px 56px;border:1px solid #e2e8f0;border-radius:6px;box-shadow:0 10px 30px rgba(0,16,80,.10);color:#111;line-height:1.55;font-size:14px}
.dca-top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;border-bottom:3px solid #0B2A6B;padding-bottom:12px;margin-bottom:22px}
.dca-top img{height:54px;width:auto}.dca-ref{text-align:right;font-size:12px;color:#475569}
.dca-dest p{margin:0}.dca-asunto{text-align:center;margin:22px 0}
.dca-cuerpo p{text-align:justify;margin:0 0 12px}.dca-firma{margin-top:34px}.dca-firma .l{width:240px;border-top:1px solid #111;margin-bottom:4px}
.dca-pie{margin-top:34px;border-top:1px solid #e2e8f0;padding-top:8px;font-size:11px;color:#64748b}
.dca-bar{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.dca-fg{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.dca-fg label{display:block;font-size:12px;font-weight:700;color:#475569;margin:0 0 4px}
.dca-fg input,.dca-fg select,.dca-fg textarea{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cfd8f3;border-radius:10px;font-size:14px}
.dca-full{grid-column:1/-1}.dca-nuevo{font-weight:800}tr.dca-sl td{background:#f0f7ff;font-weight:700}
@media(max-width:700px){.dca-fg{grid-template-columns:1fr}.dca-hoja{padding:22px 18px}}
</style>"""


def _dca_md(t):
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", _esc(t))


def _dca_html(d):
    logo = ('<img src="%s" alt="PROCSIS">' % _DCX_LOGO)
    dest = ["<p>%s</p>" % _esc(d.get("trato") or "")]
    if d.get("dest_nombre"):
        dest.append("<p><b>%s</b></p>" % _esc(d["dest_nombre"]).upper())
    dest += ["<p>%s</p>" % _esc(x) for x in _dca_lineas(d)]
    cuerpo = "".join("<p>%s</p>" % _dca_md(p).replace("\n", "<br>") for p in _dca_parrafos(d.get("cuerpo")))
    return ("<div class='dca-hoja'><div class='dca-top'>%s<div class='dca-ref'><b>%s</b><br>%s, %s<br>%s</div></div>"
            "<div class='dca-dest'>%s</div><p class='dca-asunto'><b>Asunto:</b> %s</p>"
            "%s<div class='dca-cuerpo'>%s</div>%s"
            "<div class='dca-firma'><div class='l'></div><b>%s</b><br>%s<br>%s</div>"
            "<div class='dca-pie'>Enviado a: %s · Tipo: %s · Documento generado en PROCSIS</div></div>") % (
        logo, _esc(d.get("codigo")), _esc(d.get("ciudad")), _esc(_dcx_fecha_larga(d.get("fecha"))), "", "".join(dest), _esc(d.get("asunto")),
        ("<p>%s</p>" % _esc(d["saludo"])) if d.get("saludo") else "", cuerpo,
        ("<p>%s</p>" % _esc(d["despedida"])) if d.get("despedida") else "",
        _esc(d.get("firmante")), _esc(d.get("cargo") or ""), _esc(_dcx_area_titulo(d.get("origen"))),
        _esc(_dcx_area_titulo(d.get("destino"))), _esc(_dca_cat_txt(d.get("categoria"))))


# ───────────── PDF (reportlab) y Word (python-docx) con logo ─────────────
def _dca_ltr(s):
    """Texto compatible con la fuente base del PDF (Latin-1/Windows-1252): quita emojis y símbolos raros."""
    return (s or "").encode("cp1252", "ignore").decode("cp1252")


def _dca_pdf(d):
    from xml.sax.saxutils import escape as _x
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import SimpleDocTemplate, Spacer, Image as _RLImage, HRFlowable
    base = getSampleStyleSheet()["Normal"]
    n = ParagraphStyle("dn", parent=base, fontName="Helvetica", fontSize=10.5, leading=14.5)
    just = ParagraphStyle("dj", parent=n, alignment=TA_JUSTIFY, spaceAfter=8)
    der = ParagraphStyle("dd", parent=n, alignment=TA_RIGHT, fontSize=9, textColor=colors.HexColor("#475569"))
    cen = ParagraphStyle("dc", parent=n, alignment=TA_CENTER, spaceBefore=10, spaceAfter=12)

    def t(s):
        return _x(_dca_ltr(s))

    def md(s):
        return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t(s))
    el = []
    logo = _dca_logo_path()
    if logo:
        try:
            w, h = ImageReader(logo).getSize()
            alto = 1.7 * cm
            el.append(_RLImage(logo, width=alto * w / float(h), height=alto, hAlign="LEFT"))
        except Exception:
            pass
    el.append(Paragraph("<b>%s</b><br/>%s, %s" % (t(d.get("codigo")), t(d.get("ciudad")), t(_dcx_fecha_larga(d.get("fecha")))), der))
    el.append(HRFlowable(width="100%", thickness=2.2, color=colors.HexColor("#0B2A6B"), spaceBefore=4, spaceAfter=14))
    el.append(Paragraph(t(d.get("trato")), n))
    if d.get("dest_nombre"):
        el.append(Paragraph("<b>%s</b>" % t((d["dest_nombre"] or "").upper()), n))
    for x in _dca_lineas(d):
        el.append(Paragraph(t(x), n))
    el.append(Paragraph("<b>Asunto:</b> %s" % t(d.get("asunto")), cen))
    if d.get("saludo"):
        el.append(Paragraph(t(d["saludo"]), n))
        el.append(Spacer(1, 6))
    for p in _dca_parrafos(d.get("cuerpo")):
        el.append(Paragraph(md(p).replace("\n", "<br/>"), just))
    if d.get("despedida"):
        el.append(Paragraph(t(d["despedida"]), n))
    el.append(Spacer(1, 38))
    el.append(HRFlowable(width=6.2 * cm, thickness=0.8, color=colors.black, hAlign="LEFT", spaceAfter=3))
    el.append(Paragraph("<b>%s</b>" % t(d.get("firmante")), n))
    if d.get("cargo"):
        el.append(Paragraph(t(d["cargo"]), n))
    el.append(Paragraph(t(_dcx_area_titulo(d.get("origen"))), n))
    emp = _dcx_empresa()
    pie = "PROCSIS · %s · %s · %s  |  Enviado a: %s · Tipo: %s" % (emp["web"], emp["email"], emp["tel"], _dcx_area_titulo(d.get("destino")), _dca_cat_txt(d.get("categoria")))

    def _pie(c, doc):
        c.saveState()
        c.setFont("Helvetica", 7.5)
        c.setFillColor(colors.HexColor("#64748b"))
        c.drawString(2 * cm, 1.2 * cm, _dca_ltr(pie))
        c.drawRightString(letter[0] - 2 * cm, 1.2 * cm, "Página %d" % doc.page)
        c.restoreState()
    buf = BytesIO()
    SimpleDocTemplate(buf, pagesize=letter, leftMargin=2.2 * cm, rightMargin=2.2 * cm, topMargin=1.8 * cm, bottomMargin=2.2 * cm,
                      title=d.get("codigo") or "Documento", author="PROCSIS").build(el, onFirstPage=_pie, onLaterPages=_pie)
    buf.seek(0)
    return buf


def _dca_docx(d):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import RGBColor
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1)
    sec.top_margin, sec.bottom_margin = Inches(0.8), Inches(0.9)
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)
    logo = _dca_logo_path()
    if logo:
        try:
            doc.add_picture(logo, height=Inches(0.7))
        except Exception:
            pass
    ref = doc.add_paragraph()
    ref.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = ref.add_run(d.get("codigo") or "")
    r.bold = True
    r.font.size = Pt(9)
    r2 = ref.add_run("\n%s, %s" % (d.get("ciudad") or "", _dcx_fecha_larga(d.get("fecha"))))
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
    doc.add_paragraph("_" * 86).runs[0].font.color.rgb = RGBColor(0x0B, 0x2A, 0x6B)

    def par(texto="", bold=False, align=None, after=0):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        if align:
            p.alignment = align
        partes = re.split(r"(\*\*.+?\*\*)", texto or "")
        for x in partes:
            if x.startswith("**") and x.endswith("**") and len(x) > 4:
                p.add_run(x[2:-2]).bold = True
            elif x:
                p.add_run(x).bold = bold
        return p
    par(d.get("trato") or "")
    if d.get("dest_nombre"):
        par((d["dest_nombre"] or "").upper(), bold=True)
    for x in _dca_lineas(d):
        par(x)
    a = doc.add_paragraph()
    a.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a.paragraph_format.space_before = Pt(14)
    a.paragraph_format.space_after = Pt(12)
    a.add_run("Asunto: ").bold = True
    a.add_run(d.get("asunto") or "")
    if d.get("saludo"):
        par(d["saludo"], after=6)
    for p in _dca_parrafos(d.get("cuerpo")):
        par(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, after=8)
    if d.get("despedida"):
        par(d["despedida"])
    par("")
    par("______________________________")
    par(d.get("firmante") or "", bold=True)
    if d.get("cargo"):
        par(d["cargo"])
    par(_dcx_area_titulo(d.get("origen")))
    pie = doc.sections[0].footer.paragraphs[0]
    emp = _dcx_empresa()
    pie.text = "PROCSIS · %s · %s · %s  |  Enviado a: %s · Tipo: %s" % (emp["web"], emp["email"], emp["tel"], _dcx_area_titulo(d.get("destino")), _dca_cat_txt(d.get("categoria")))
    for run in pie.runs:
        run.font.size = Pt(8)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


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
            resalta, _esc(cod), _esc(_dca_cat_txt(ct)), _esc((asu or "")[:80]), _esc(_dcx_area_titulo(ori)), _esc(_dcx_area_titulo(des)), _esc(fe), etq, txt, did)
    if not filas:
        filas = "<tr><td colspan='7'>No hay documentos en esta vista.</td></tr>"
    tabs = "".join("<a class='cb-btn s %s' href='/documentos?v=%s'>%s</a> " % ("" if vista == k else "o", k, n)
                   for k, n in (("recibidos", "Recibidos%s" % (" (%d sin leer)" % n_sl if n_sl else "")), ("enviados", "Enviados")) + ((("todos", "Todos"),) if ger else ()))
    opts = "<option value=''>Todos los tipos</option>" + "".join("<option value='%s'%s>%s</option>" % (k, " selected" if k == cat else "", n) for k, n in _DCA_CATEGORIAS)
    filtro = ("<form method='GET' action='/documentos' style='display:inline-flex;gap:6px;flex-wrap:wrap'><input type='hidden' name='v' value='%s'>"
              "<select name='cat'>%s</select><input type='text' name='q' value='%s' placeholder='Código, asunto o destinatario' style='width:220px'>"
              "<button class='cb-btn s'>Filtrar</button></form>") % (vista, opts, _esc(q))
    cuerpo = ("<h1>📨 Documentos por área</h1><p class='sub'>Crea documentos con el logo de PROCSIS, envíalos al área que corresponda y descárgalos en PDF o Word.</p>"
              "<div class='dca-bar'><a class='cb-btn g' href='/documentos/nuevo'>➕ Nuevo documento</a></div>"
              "<div class='cb-card'><div class='dca-bar'>%s%s</div><div class='cb-scroll'><table><tr><th>Código</th><th>Tipo</th><th>Asunto</th>"
              "<th>De → Para</th><th>Fecha</th><th>Estado</th><th></th></tr>%s</table></div></div>") % (tabs, filtro, filas)
    return _cb_pagina("Documentos por área", _DCA_CSS + cuerpo)


def _dca_form(csrf):
    area = _dca_area_usuario()
    nombre = ""
    try:
        nombre = _nombre_para_saludo(session.get("usuario") or "") or ""
    except Exception:
        pass
    cats = "".join("<option value='%s'>%s</option>" % (k, n) for k, n in _DCA_CATEGORIAS)
    dests = "".join("<option value='%s'>%s</option>" % (_esc(a), _esc(_dcx_area_titulo(a))) for a in _dca_areas())
    plant = ("<option value=''>En blanco</option><option value='comunidad'>Comunidad EduTrack (directores, rectores y aliados)</option>"
             "<option value='interna'>Documento interno entre áreas</option>")
    js = ("<script>(function(){var P=%s;var s=document.getElementById('dca-pl');if(!s)return;"
          "s.addEventListener('change',function(){var p=P[s.value];if(!p&&s.value!=='')return;p=p||{trato:'Señor(a)',nombre:'',lineas:''};"
          "document.getElementById('dca-trato').value=p.trato;document.getElementById('dca-nom').value=p.nombre;document.getElementById('dca-lin').value=p.lineas;});})();</script>"
          % json.dumps(_DCA_PLANTILLAS_DEST, ensure_ascii=False))
    return ("<h1>Nuevo documento</h1><p class='sub'>Se emite con el logo de PROCSIS y un código único. Un documento emitido no se edita: si hay un error, se anula y se emite otro.</p>"
            "<form method='POST' action='/documentos/nuevo'><input type='hidden' name='csrf' value='%s'>"
            "<div class='cb-card'><div class='dca-fg'>"
            "<div><label>Tipo de documento</label><select name='categoria'>%s</select></div>"
            "<div><label>Enviar a (área)</label><select name='destino'>%s</select></div>"
            "<div><label>Ciudad</label><input type='text' name='ciudad' value='Caracolí, Antioquia' maxlength='100'></div>"
            "<div><label>Fecha</label><input type='date' name='fecha' value='%s'></div>"
            "<div class='dca-full'><label>Plantilla de destinatario</label><select id='dca-pl'>%s</select></div>"
            "<div><label>Tratamiento</label><input type='text' id='dca-trato' name='trato' value='Señor(a)' maxlength='60'></div>"
            "<div><label>Destinatario (se muestra en negrita)</label><input type='text' id='dca-nom' name='dest_nombre' maxlength='220'></div>"
            "<div class='dca-full'><label>Otras líneas del destinatario (una por línea)</label><textarea id='dca-lin' name='dest_lineas' rows='3'></textarea></div>"
            "<div class='dca-full'><label>Asunto</label><input type='text' name='asunto' maxlength='600' required></div>"
            "<div><label>Saludo</label><input type='text' name='saludo' value='Cordial saludo:' maxlength='160'></div>"
            "<div><label>Despedida</label><input type='text' name='despedida' value='Atentamente,' maxlength='160'></div>"
            "<div class='dca-full'><label>Cuerpo del documento</label><textarea name='cuerpo' rows='11' required></textarea>"
            "<div class='cb-note'>Separa los párrafos con una línea en blanco. Usa **texto** para negrita.</div></div>"
            "<div><label>Firmante (nombre)</label><input type='text' name='firmante' value='%s' maxlength='160' required></div>"
            "<div><label>Cargo del firmante</label><input type='text' name='cargo' maxlength='160'></div>"
            "<div class='dca-full'><label>Área que emite</label><input type='text' value='%s' disabled></div>"
            "</div><p style='margin-top:14px'><button class='cb-btn g'>Emitir y enviar</button> <a class='cb-btn o' href='/documentos'>Cancelar</a></p></div></form>%s") % (
        _esc(csrf), cats, dests, fecha_hoy(), plant, _esc(nombre), _esc(_dcx_area_titulo(area)), js)


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
        return " ".join((f.get(k) or "").split())[:n]
    categoria, destino = lim("categoria", 30), lim("destino", 60)
    asunto, firmante = lim("asunto", 600), lim("firmante", 160)
    cuerpo = (f.get("cuerpo") or "").replace("\r", "").strip()[:9000]
    if categoria not in dict(_DCA_CATEGORIAS) or destino not in _dca_areas():
        session["ti_msg"] = ("err", "Elige el tipo de documento y el área a la que se envía.")
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
    try:
        _c48_x("INSERT INTO docs_area (codigo, categoria, origen, destino, asunto, trato, dest_nombre, dest_lineas, saludo, cuerpo, despedida, firmante, cargo, "
               "ciudad, fecha, estado, creado_por, creado, numero, anio) VALUES (:codigo, :categoria, :origen, :destino, :asunto, :trato, :dn, :dl, :saludo, "
               ":cuerpo, :desp, :firmante, :cargo, :ciudad, :fecha, 'ENVIADO', :us, :cr, :num, :anio)",
               codigo=codigo, categoria=categoria, origen=origen, destino=destino, asunto=asunto, trato=lim("trato", 60), dn=lim("dest_nombre", 220),
               dl=(f.get("dest_lineas") or "").replace("\r", "").strip()[:900], saludo=lim("saludo", 160), cuerpo=cuerpo, desp=lim("despedida", 160),
               firmante=firmante, cargo=lim("cargo", 160), ciudad=lim("ciudad", 100), fecha=fecha, us=usuario, cr=str(ahora())[:19], num=num, anio=anio)
    except Exception as e:
        db.session.rollback()
        print("docs_area crear:", e)
        session["ti_msg"] = ("err", "No se pudo emitir el documento. Intenta de nuevo.")
        return redirect("/documentos/nuevo")
    nid = _c48_q("SELECT id FROM docs_area WHERE codigo=:c", c=codigo)
    registrar_auditoria("Documentos por área", "%s emitió %s para %s (%s)" % (usuario, codigo, destino, categoria))
    session["ti_msg"] = ("ok", "Documento %s emitido y enviado a %s." % (codigo, _dcx_area_titulo(destino)))
    return redirect("/documentos/%d" % int(nid[0][0])) if nid else redirect("/documentos")


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
    if d["destino"] == area and d["creado_por"] != usuario and d["estado"] == "ENVIADO":
        try:
            _c48_x("UPDATE docs_area SET estado='LEIDO', leido_por=:u, leido=:t WHERE id=:i AND estado='ENVIADO'", u=usuario, t=str(ahora())[:19], i=did)
            d["estado"] = "LEIDO"
        except Exception:
            db.session.rollback()
    csrf = _acc_csrf()
    acciones = ""
    if d["estado"] in ("ENVIADO", "LEIDO") and (d["destino"] == area or rol in _DCX_GERENCIA):
        acciones += ("<form method='POST' action='/documentos/%d/estado' style='display:inline'><input type='hidden' name='csrf' value='%s'>"
                     "<input type='hidden' name='accion' value='atendido'><button class='cb-btn g s'>✔ Marcar como atendido</button></form> ") % (did, _esc(csrf))
    if d["estado"] != "ANULADO" and (rol in _DCX_GERENCIA or d["creado_por"] == usuario):
        acciones += ("<form method='POST' action='/documentos/%d/estado' style='display:inline-flex;gap:6px' onsubmit=\"return confirm('¿Anular este documento?');\">"
                     "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='accion' value='anular'>"
                     "<input type='text' name='nota' placeholder='Motivo' required maxlength='200' style='width:200px'><button class='cb-btn r s'>Anular</button></form>") % (did, _esc(csrf))
    est = {"ENVIADO": "<span class='cb-tag wait'>Enviado · sin leer</span>", "LEIDO": "<span class='cb-tag on'>Leído por %s</span>" % _esc(d.get("leido_por") or ""),
           "ATENDIDO": "<span class='cb-tag on'>Atendido</span>", "ANULADO": "<span class='cb-tag off'>Anulado · %s</span>" % _esc(d.get("nota") or "")}.get(d["estado"], "")
    barra = ("<div class='dca-bar'><a class='cb-btn' href='/documentos/%d.pdf'>⬇ Descargar PDF</a><a class='cb-btn o' href='/documentos/%d.docx'>📝 Descargar Word</a>"
             "<a class='cb-btn o' href='/documentos'>← Volver</a>%s</div><div class='dca-bar'>%s %s</div>") % (did, did, "", est, acciones)
    return _cb_pagina(d["codigo"], _DCA_CSS + "<h1>%s</h1><p class='sub'>%s · de %s para %s</p>%s%s" % (
        _esc(d["codigo"]), _esc(_dca_cat_txt(d["categoria"])), _esc(_dcx_area_titulo(d["origen"])), _esc(_dcx_area_titulo(d["destino"])), barra, _dca_html(d)))


def _dca_descarga(did, formato):
    g = _dca_guard()
    if g is not None:
        return g
    d = _dca_row(did)
    if not d or not _dca_puede(d, rol_actual(), session.get("usuario") or ""):
        return acceso_denegado("Documento no disponible.")
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
        if accion == "atendido" and d["estado"] in ("ENVIADO", "LEIDO") and (d["destino"] == area or rol in _DCX_GERENCIA):
            _c48_x("UPDATE docs_area SET estado='ATENDIDO', leido_por=:u, leido=:t WHERE id=:i", u=usuario, t=str(ahora())[:19], i=did)
            registrar_auditoria("Documentos por área", "%s atendió %s" % (usuario, d["codigo"]))
            session["ti_msg"] = ("ok", "Documento marcado como atendido.")
        elif accion == "anular" and d["estado"] != "ANULADO" and (rol in _DCX_GERENCIA or d["creado_por"] == usuario):
            motivo = " ".join((request.form.get("nota") or "").split())[:200]
            if motivo:
                _c48_x("UPDATE docs_area SET estado='ANULADO', nota=:n WHERE id=:i", n=motivo, i=did)
                registrar_auditoria("Documentos por área", "%s anuló %s" % (usuario, d["codigo"]))
                session["ti_msg"] = ("ok", "Documento anulado.")
    except Exception:
        db.session.rollback()
        session["ti_msg"] = ("err", "No se pudo actualizar el documento.")
    return redirect("/documentos/%d" % did)


'''

# ───────────── Menú lateral ─────────────
ANCLA_GER = "        ('/documentos-oficiales', '📄 Documentos oficiales (todas las áreas)'),\n"
NUEVO_GER = ANCLA_GER + "        ('/documentos', '📨 Documentos por área (PDF y Word)'),  # [procsis-documentos-area]\n"
ANCLA_ROL = '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n'
NUEVO_ROL = ('    if _dca_roles_ok(rol) and not any(h == "/documentos" for h, _l in items):  # [procsis-documentos-area]\n'
             '        items = [x for x in items if x[0] != "/logout"] + [("/documentos", "📨 Documentos por área")] + [x for x in items if x[0] == "/logout"]\n' + ANCLA_ROL)

REEMPLAZOS = [("menú de Gerencia", ANCLA_GER, NUEVO_GER), ("menú de las demás áreas", ANCLA_ROL, NUEVO_ROL)]
ANCLAS = ["def _dcx_rol_ok", "def _dcx_areas_validas", "def _dcx_area_de_rol", "def _dcx_area_titulo", "def _dcx_fecha_larga", "def _dcx_empresa",
          "def _cb_pagina", "def _c48_q", "def _c48_x", "def _acc_csrf", "def _ti_csrf_ok", "def registrar_auditoria", "def acceso_denegado",
          "def fecha_hoy", "def _nombre_para_saludo", "_DCX_LOGO = ", "_DCX_COD = ", "from docx import Document", "from reportlab.platypus import Table, TableStyle, Paragraph"]


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
    for desc, viejo, _n in REEMPLAZOS:
        if texto.count(viejo) != 1:
            faltan.append("%s (aparece %d veces, debe ser 1)" % (desc, texto.count(viejo)))
    if faltan:
        sys.exit("Este app.py no coincide con lo esperado; no toqué nada. Falta/cambió:\n  - " + "\n  - ".join(faltan))
    nuevo = texto
    for _d, viejo, reemplazo in REEMPLAZOS:
        nuevo = nuevo.replace(viejo, reemplazo, 1)
    corte = nuevo.rfind('\nif __name__ == "__main__":')
    if corte < 0:
        sys.exit('No encontré el bloque if __name__ == "__main__": al final de app.py.')
    nuevo = nuevo[:corte] + "\n" + BLOQUE.strip("\n") + "\n\n" + nuevo[corte:]
    try:
        compile(nuevo, str(ruta), "exec")
    except SyntaxError as e:
        sys.exit("El resultado no compila (línea %s: %s). No escribí nada." % (e.lineno, e.msg))
    copia = ruta.with_name(ruta.name + ".bak_documentos_area")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_documentos_area_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. Módulo «Documentos por área» agregado (ruta /documentos). Archivo: %s" % ruta)
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()

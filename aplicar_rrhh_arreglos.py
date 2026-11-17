# -*- coding: utf-8 -*-
"""
aplicar_rrhh_arreglos.py
Arregla en RR.HH. / Talento Humano:
  1) Módulo de Incapacidades: aparece en el menú lateral y en el portal de RR.HH., y deja de ser bloqueado.
  2) Turnos: en RR.HH. el título ya no dice «Gerencia», dice «Talento Humano».
  3) Admisión de personal: ahora lista a TODO el equipo (todas las áreas) y a los trabajadores con contrato vigente.

Uso:  python aplicar_rrhh_arreglos.py        (junto a app.py)
Deja un respaldo: app.py.respaldo_rrhh
"""
import os
import re
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "app.py")
MARCA = "# [procsis-rrhh-arreglos]"


def reemplazar(src, viejo, nuevo, nombre):
    n = src.count(viejo)
    if n != 1:
        print("ERROR en «%s»: se esperaba 1 coincidencia y hay %d. No se modificó nada." % (nombre, n))
        sys.exit(1)
    return src.replace(viejo, nuevo)


def main():
    if not os.path.isfile(RUTA):
        print("No encuentro app.py junto a este script:", RUTA)
        sys.exit(1)
    with open(RUTA, "r", encoding="utf-8", newline="") as f:
        src = f.read()
    crlf = "\r\n" in src          # conserva el formato de saltos de línea de tu archivo (Windows = CRLF)
    src = src.replace("\r\n", "\n")
    if MARCA in src:
        print("Este parche ya estaba aplicado. No se hizo nada.")
        return
    shutil.copyfile(RUTA, RUTA + ".respaldo_rrhh")

    # ── 1a) Menú lateral: añadir «Incapacidades» (la función _inc_menu_items existía pero nadie la llamaba)
    src = reemplazar(
        src,
        '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n',
        '    ' + MARCA + '\n'
        '    try:\n'
        '        _ya_inc = {h for h, _l in items}\n'
        '        _inc_add = [(h, l) for h, l in _inc_menu_items(rol, session.get("usuario") or "") if h not in _ya_inc]\n'
        '        if _inc_add:\n'
        '            items = [x for x in items if x[0] != "/logout"] + _inc_add + [x for x in items if x[0] == "/logout"]\n'
        '    except Exception:\n'
        '        pass\n'
        '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n',
        "menú lateral",
    )

    # ── 1b) Filtro de rutas por rol: las incapacidades validan su propio acceso (_inc_guard)
    src = reemplazar(
        src,
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        '    if path.startswith(("/talento/incapacidades", "/incapacidades/")):\n'
        '        return None  # [procsis-inc] _inc_guard valida el acceso\n'
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        "filtro de rutas",
    )

    # ── 1c) Tarjeta «Incapacidades» en el portal de RR.HH.
    src = reemplazar(
        src,
        '        tarjetas = _tm\n',
        '        tarjetas = _tm\n'
        '        if rol_area == "Talento":\n'
        '            tarjetas += ("<a href=\'/talento/incapacidades\' class=\'role-panel\' style=\'border-left:6px solid %s;text-decoration:none;display:block\'>"\n'
        '                         "<div style=\'font-weight:800;color:#0f172a\'>🩺 Incapacidades</div>"\n'
        '                         "<div style=\'font-size:12px;color:#475569;margin-top:4px\'>Registro y consulta de incapacidades del personal.</div>"\n'
        '                         "<div style=\'margin-top:8px\'><span style=\'background:#ede9fe;color:#5b21b6;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700\'>Abrir</span></div></a>" % a["color"])\n',
        "tarjeta portal RR.HH.",
    )

    # ── 2) Turnos: nombre correcto del área en pantalla (los datos siguen guardándose igual)
    src = reemplazar(
        src,
        '    volver = _home_portal(rol)\n\n    # Panel de turnos',
        '    volver = _home_portal(rol)\n    area_txt = "Talento Humano" if rol == "Talento" else area\n\n    # Panel de turnos',
        "turnos: area_txt",
    )
    src = reemplazar(src, '<h1>👥 Turnos · {area}</h1>', '<h1>👥 Turnos · {area_txt}</h1>', "turnos: título")
    src = reemplazar(src, 'DETALLE {area.upper()} · COLA DE TURNOS', 'DETALLE {area_txt.upper()} · COLA DE TURNOS', "turnos: detalle")
    src = reemplazar(src, 'return page(f"Turnos {area}", shell(content))', 'return page(f"Turnos {area_txt}", shell(content))', "turnos: título de página")

    # ── 3) Admisión de personal: todo el equipo + trabajadores con contrato vigente
    ini = src.index('@app.route("/gerencia/admision-personal", methods=["GET", "POST"])')
    fin = src.index('@app.route("/gerencia/procsis-web", methods=["GET", "POST"])')
    if not (0 < ini < fin):
        print("ERROR: no pude delimitar la función de admisión.")
        sys.exit(1)
    src = src[:ini] + NUEVA_ADMISION + "\n\n" + src[fin:]

    if crlf:
        src = src.replace("\n", "\r\n")
    with open(RUTA, "w", encoding="utf-8", newline="") as f:
        f.write(src)
    print("Listo. Parche aplicado. Respaldo en app.py.respaldo_rrhh")


NUEVA_ADMISION = '''@app.route("/gerencia/admision-personal", methods=["GET", "POST"])
def gerencia_admision_personal():
    """Certificado de admisión de personal (PDF con logo) para TODO el equipo de Procsis:
    usuarios internos de todas las áreas + trabajadores con contrato vigente."""
    g = _guard_talento()
    if g:
        return g
    try:
        db.create_all()
    except Exception:
        pass
    ROLES_STAFF = [r for r in ROLES_INTERNOS if r != "Proveedor"]
    staff = Usuario.query.filter(Usuario.rol.in_(ROLES_STAFF)).order_by(Usuario.rol.asc(), Usuario.usuario.asc()).all()
    try:
        contratos = ContratoPersonal.query.filter(
            or_(ContratoPersonal.estado.is_(None), ~ContratoPersonal.estado.in_(["ARCHIVADO", "TERMINADO", "BORRADOR"]))
        ).order_by(ContratoPersonal.nombres.asc()).all()
    except Exception:
        db.session.rollback()
        contratos = []
    nombres_con_usuario = {(u.nombre_completo or "").strip().lower() for u in staff if (u.nombre_completo or "").strip()}
    contratos = [c for c in contratos if (c.nombres or "").strip().lower() not in nombres_con_usuario]

    if request.method == "POST":
        uid = request.form.get("uid", type=int)
        cid = request.form.get("cid", type=int)
        fecha_ingreso = (request.form.get("fecha_ingreso") or "").strip()
        cargo = (request.form.get("cargo") or "").strip()
        u = Usuario.query.get(uid) if uid else None
        c = ContratoPersonal.query.get(cid) if cid else None
        if u is not None and u.rol in ROLES_STAFF:
            nombre_mostrar = u.nombre_completo or u.usuario
            ident = 'identificado(a) con usuario de acceso "%s"' % u.usuario
            cargo = cargo or u.rol
            archivo = u.usuario
            ref_aud = u.usuario
        elif c is not None:
            nombre_mostrar = c.nombres or "—"
            ident = "identificado(a) con documento %s" % (c.documento or "—")
            cargo = cargo or c.cargo or "colaborador(a)"
            fecha_ingreso = fecha_ingreso or (c.fecha_inicio or "")
            archivo = "doc_" + (re.sub(r"[^0-9A-Za-z]", "", c.documento or "") or str(c.id))
            ref_aud = "%s (contrato %s)" % (c.nombres, c.id)
        else:
            return page("Admisión de personal", shell("<div class='msg err'>Persona no válida. <a href='/gerencia/admision-personal'>Volver</a></div>"))
        p = plataforma()
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas as _canvas
        from reportlab.lib.units import cm
        from reportlab.lib.utils import simpleSplit
        buf = BytesIO()
        pdf = _canvas.Canvas(buf, pagesize=letter)
        w, h = letter
        try:
            _logo_web = logo_plataforma()
            _logo_disk = _logo_web.lstrip("/").replace("\\\\", "/")
            pdf.drawImage(_logo_disk, 2 * cm, h - 3.2 * cm, width=2.2 * cm, height=2.2 * cm, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
        pdf.setFont("Helvetica-Bold", 15)
        pdf.drawCentredString(w / 2, h - 2.2 * cm, p.empresa or "Procsis")
        pdf.setFont("Helvetica-Bold", 13)
        pdf.drawCentredString(w / 2, h - 3.4 * cm, "CERTIFICADO DE ADMISIÓN DE PERSONAL")
        pdf.setFont("Helvetica", 10)
        pdf.drawCentredString(w / 2, h - 4.0 * cm, f"NIT: {p.nit or '—'} · {p.direccion or ''} {p.ciudad or ''}".strip())
        pdf.setFont("Helvetica", 11)
        y = h - 5.4 * cm
        texto = (
            f"{p.representante_legal or 'La Gerencia'}, en representación de {p.empresa or 'Procsis'}, certifica que "
            f"{nombre_mostrar}, {ident}, fue admitido(a) para "
            f"desempeñarse como {cargo}, a partir del {fecha_ingreso or '—'}, vinculación vigente a la "
            f"fecha de expedición del presente certificado."
        )
        for ln in simpleSplit(texto, "Helvetica", 11, w - 4 * cm):
            pdf.drawString(2 * cm, y, ln)
            y -= 16
        y -= 40
        pdf.drawString(2 * cm, y, f"Fecha de expedición: {fecha_hoy()}")
        y -= 50
        pdf.drawString(2 * cm, y, "___________________________")
        y -= 12
        pdf.drawString(2 * cm, y, p.representante_legal or "Gerencia")
        pdf.save()
        buf.seek(0)
        registrar_auditoria("Certificado de admisión generado", f"{ref_aud} · cargo: {cargo} · ingreso: {fecha_ingreso}")
        return send_file(buf, as_attachment=True, download_name=f"admision_{archivo}.pdf", mimetype="application/pdf")

    filas = "".join(
        f"""<tr><td><b>{_esc(u.usuario)}</b>{' · '+_esc(u.nombre_completo) if u.nombre_completo else ''}</td><td>{_esc(u.rol)}</td>
        <td><form method="POST" style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <input type="hidden" name="uid" value="{u.id}">
          <input name="cargo" placeholder="Cargo (ej. Asesor comercial)" style="width:170px">
          <input name="fecha_ingreso" type="date" required>
          <button type="submit" class="small-action">Generar certificado</button>
        </form></td></tr>"""
        for u in staff
    ) or '<tr><td colspan="3">Sin cuentas de staff</td></tr>'
    filas_c = "".join(
        f"""<tr><td><b>{_esc(c.nombres)}</b> · {_esc(c.documento)}</td><td>{_esc(c.area or '—')}</td>
        <td><form method="POST" style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <input type="hidden" name="cid" value="{c.id}">
          <input name="cargo" value="{_esc(c.cargo or '')}" placeholder="Cargo" style="width:170px">
          <input name="fecha_ingreso" type="date" value="{_esc((c.fecha_inicio or '')[:10])}" required>
          <button type="submit" class="small-action">Generar certificado</button>
        </form></td></tr>"""
        for c in contratos
    )
    bloque_contratos = ""
    if filas_c:
        bloque_contratos = f"""
<section class="role-panel" style="margin-top:14px">
  <h3 style="margin-top:0">Trabajadores contratados (sin usuario de acceso)</h3>
  <div style="overflow-x:auto"><table>
    <tr><th>Trabajador</th><th>Área</th><th>Generar</th></tr>
    {filas_c}
  </table></div>
</section>"""
    content = f"""
<header class="role-hero"><div>
  <h1>📄 Admisión de personal</h1>
  <p>Certificado de vinculación laboral de cualquier integrante de Procsis (todas las áreas) o de un trabajador con contrato vigente, con el logo de la empresa.</p>
</div>
<a class="btn" href="/gerencia/hq">Volver</a></header>
<section class="role-panel">
  <h3 style="margin-top:0">Equipo con usuario de acceso ({len(staff)})</h3>
  <div style="overflow-x:auto">
    <table>
      <tr><th>Usuario</th><th>Rol</th><th>Generar</th></tr>
      {filas}
    </table>
  </div>
  <p class="mini-text" style="margin-top:8px">Los datos de NIT y dirección salen de <a href="/gerencia/datos-empresa">Datos de la empresa</a> — complételos primero para que el PDF salga completo.</p>
</section>
{bloque_contratos}
"""
    return page("Admisión de personal", shell(content))
'''

if __name__ == "__main__":
    main()

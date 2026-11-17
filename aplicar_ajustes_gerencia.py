#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_ajustes_gerencia.py  ·  PROCSIS: 4 ajustes pedidos.

Uso (en la carpeta de tu app.py):
    python aplicar_ajustes_gerencia.py            # parchea ./app.py
    python aplicar_ajustes_gerencia.py ruta/app.py

1) «Tickets documentales» ya NO flota en todas las pantallas (tapaba el selector de producto). Sigue disponible en el menú lateral
   de cada área que lo usa.
2) HORARIOS DE ASESORES: Gerencia ya puede cambiar horarios (plantillas, días y horas, accesos extraordinarios) sin quedar restringida
   por reloj, pero CADA cambio exige una nota (por qué). Todo queda en una bitácora dentro del mismo módulo: quién, cuándo, antes → después
   y la nota. Talento Humano y Gerencia pueden dejar notas y «Ajustes de presupuesto» (con impacto estimado en pesos) en esa bitácora.
3) PLANES DE TIENDATRACK: ahora tienen descuento. Escribes precio lista (mensual y anual) y el % de descuento, o el precio final que quieres
   cobrar, y se calcula todo solo. El plan se guarda con el precio final. (En EduTrack el % ya existía: además se corrige el error al escribir
   decimales con coma; si no ves la tarjeta «Descuento y precio final», aplica también aplicar_descuentos_planes.py.)
4) «Documentos oficiales» y «Documentos por área» ya no te devuelven al inicio: el filtro de portales por rol los dejaba fuera de su lista.
Es idempotente, crea app.py.bak_ajustes_gerencia y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-ajustes-1]"

BLOQUE = r'''
# [procsis-ajustes-1]
# ═════════════ AJUSTES: horarios de asesores con bitácora, descuento en planes TiendaTrack ═════════════
_HN_GER = ("Gerente", "Superadmin", "Administrador")
_HN = {"ok": False}
_HN_TIPOS = {"cambio": "Cambio de horario", "nota_rrhh": "Nota de RR.HH.", "nota_ger": "Nota de Gerencia", "presupuesto": "Ajuste de presupuesto"}


def _hn_una_linea(s, n=400):
    return " ".join(str(s or "").replace("\r", " ").replace("\n", " ").split())[:n]


def _hn_cols(tabla):
    if db.engine.dialect.name == "sqlite":
        return [r[1] for r in db.session.execute(text("PRAGMA table_info(%s)" % tabla)).fetchall()]
    return [r[0] for r in db.session.execute(text(
        "SELECT column_name FROM information_schema.columns WHERE table_schema = current_schema() AND table_name = :t"), {"t": tabla}).fetchall()]


def _hn_init():
    if _HN["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS horarios_notas (id %s, fecha VARCHAR(19), autor VARCHAR(80), rol VARCHAR(40), tipo VARCHAR(16), "
            "objetivo VARCHAR(160) DEFAULT '', accion VARCHAR(80) DEFAULT '', antes TEXT DEFAULT '', despues TEXT DEFAULT '', nota TEXT DEFAULT '', "
            "impacto BIGINT DEFAULT 0)" % pk))
        db.session.commit()
        _HN["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("horarios_notas init:", e)


def _hn_log(tipo, objetivo, accion, antes, despues, nota, impacto=0):
    try:
        _c48_x("INSERT INTO horarios_notas (fecha, autor, rol, tipo, objetivo, accion, antes, despues, nota, impacto) "
               "VALUES (:f, :a, :r, :t, :o, :ac, :an, :de, :n, :i)",
               f=str(ahora())[:19], a=session.get("usuario") or "", r=rol_actual() or "", t=tipo, o=_hn_una_linea(objetivo, 160), ac=_hn_una_linea(accion, 80),
               an=str(antes or "")[:600], de=str(despues or "")[:600], n=_hn_una_linea(nota), i=int(impacto or 0))
    except Exception as e:
        db.session.rollback()
        print("horarios_notas log:", e)


# Gerencia ya puede OPERAR el módulo de horarios (solo ese módulo de Talento); el resto de módulos de personal sigue en modo lectura.
_hn_guard_talento_orig = _guard_talento


def _guard_talento():
    try:
        if (request.path or "").startswith("/gerencia/horarios-asesores") and requiere_login() and rol_actual() in _HN_GER:
            return None
    except Exception:
        pass
    return _hn_guard_talento_orig()


def _hn_html_de(resp):
    if isinstance(resp, str):
        return resp
    try:
        if getattr(resp, "status_code", 200) == 200 and not getattr(resp, "direct_passthrough", False) and (resp.mimetype or "") == "text/html":
            return resp.get_data(as_text=True)
    except Exception:
        pass
    return None


def _hn_poner(resp, html):
    if isinstance(resp, str):
        return html
    resp.set_data(html)
    return resp


def _hn_objetivos(f):
    acc = f.get("accion")
    try:
        if acc == "preset" and (f.get("alcance") or "") == "rol":
            rol_o = (f.get("rol_obj") or "").strip()
            return Usuario.query.filter(Usuario.rol == rol_o).all() if rol_o in _HL_ROLES else []
        if acc == "extra_revocar":
            a = db.session.get(AccesoExtraordinario, int(f.get("extra_id") or 0))
            uid = a.usuario_id if a else 0
        else:
            uid = int(f.get("usuario_id") or 0)
        u = db.session.get(Usuario, uid) if uid else None
        return [u] if u else []
    except Exception:
        db.session.rollback()
        return []


def _hn_estado(us):
    out = {}
    for u in us:
        try:
            n_ex = AccesoExtraordinario.query.filter_by(usuario_id=u.id, activo=True).count()
        except Exception:
            db.session.rollback()
            n_ex = 0
        out[u.id] = (u.usuario, u.rol, _hl_resumen(u.id), n_ex)
    return out


def _hn_vista(*a, **k):
    orig = _hn_vista_orig
    _hn_init()
    if request.method == "POST":
        f = request.form
        acc = (f.get("accion") or "").strip()
        if acc in ("guardar", "preset", "extra", "extra_revocar") and requiere_login():
            rol = rol_actual()
            nota = _hn_una_linea(f.get("nota"))
            if rol in _HN_GER and len(nota) < 8:
                _n_flash("err", "Gerencia debe escribir una nota (mínimo 8 letras) explicando por qué cambia el horario.")
                return redirect("/gerencia/horarios-asesores")
            objetivos = _hn_objetivos(f)
            antes = _hn_estado(objetivos)
            resp = orig(*a, **k)
            try:
                for u in objetivos:
                    db.session.refresh(u)
                despues = _hn_estado(objetivos)
                cambios = []
                for uid, (nom, rl, res0, ex0) in antes.items():
                    nom1, rl1, res1, ex1 = despues.get(uid, (nom, rl, res0, ex0))
                    if res1 != res0:
                        cambios.append(("%s (%s)" % (nom, rl), {"guardar": "Horario editado", "preset": "Plantilla aplicada"}.get(acc, "Horario cambiado"), res0, res1))
                    elif ex1 != ex0:
                        det = ("%s → %s · %s" % (f.get("desde") or "", f.get("hasta") or "", _hn_una_linea(f.get("motivo"), 120))) if ex1 > ex0 else "Revocado"
                        cambios.append(("%s (%s)" % (nom, rl), "Acceso extraordinario" if ex1 > ex0 else "Acceso extraordinario revocado", "—", det))
                if len(cambios) > 5:
                    _hn_log("cambio", "%d usuarios del rol %s" % (len(cambios), (f.get("rol_obj") or "").strip()), "Plantilla aplicada a todo el rol",
                            "", cambios[0][3] + " (y otros)", nota)
                else:
                    for objetivo, accion, a0, d0 in cambios:
                        _hn_log("cambio", objetivo, accion, a0, d0, nota)
            except Exception as e:
                db.session.rollback()
                print("horarios bitacora:", e)
            return resp
        return orig(*a, **k)
    resp = orig(*a, **k)
    html = _hn_html_de(resp)
    if html is None:
        return resp
    try:
        rol = rol_actual()
        es_ger = rol in _HN_GER
        obligatoria = " required" if es_ger else ""
        ph = "Nota obligatoria: ¿por qué este cambio?" if es_ger else "Nota (opcional): motivo del cambio"
        html = re.sub(r"(<input type='hidden' name='accion' value='(?:guardar|preset|extra|extra_revocar)'>)",
                      lambda m: m.group(1) + "<input type='text' name='nota' maxlength='400' placeholder='%s'%s style='min-width:230px;margin:4px 0'>" % (ph, obligatoria), html)
        aviso = ("<section class='nx-card' style='border-left:5px solid #0020C2'><b>📒 Todo cambio queda registrado.</b> "
                 "%s Talento Humano ve la bitácora de abajo y deja sus notas y ajustes de presupuesto.</section>") % (
            "Como Gerencia puedes cambiar horarios, plantillas y accesos extraordinarios, pero cada cambio exige una nota con el motivo." if es_ger
            else "Gerencia puede cambiar horarios de asesores; cada cambio exige una nota con el motivo.")
        marca1 = "<section class='nx-card'><div class='nx-grid'>"
        if html.count(marca1) >= 1:
            html = html.replace(marca1, aviso + marca1, 1)
        marca2 = "<section class='nx-card'><h2>Intentos bloqueados por horario</h2>"
        bit = _hn_bitacora_html(rol)
        html = html.replace(marca2, bit + marca2, 1) if marca2 in html else html.replace("</body>", bit + "</body>", 1)
        return _hn_poner(resp, html)
    except Exception as e:
        print("horarios inyeccion:", e)
        return resp


def _hn_bitacora_html(rol):
    filas = ""
    try:
        for f_, au, rl, tp, ob, ac, an, de, no, im in _c48_q(
                "SELECT fecha, autor, rol, tipo, objetivo, accion, antes, despues, nota, impacto FROM horarios_notas ORDER BY id DESC LIMIT 80"):
            tag = {"cambio": "nx-t-bad", "presupuesto": "nx-t-ok"}.get(tp, "nx-t-gray")
            cambio = ("<div class='nx-small'>%s → <b>%s</b></div>" % (_esc(an or "—"), _esc(de or "—"))) if (an or de) else ""
            filas += ("<tr><td>%s</td><td><b>%s</b><div class='nx-small'>%s</div></td><td><span class='nx-tag %s'>%s</span><div class='nx-small'>%s</div></td>"
                      "<td>%s%s</td><td>%s</td><td>%s</td></tr>") % (
                _esc(f_), _esc(au), _esc(rl), tag, _esc(_HN_TIPOS.get(tp, tp)), _esc(ac), _esc(ob or "—"), cambio, _esc(no),
                ("$ {:,}".format(int(im)).replace(",", ".") if im else "—"))
    except Exception as e:
        db.session.rollback()
        print("horarios bitacora lectura:", e)
    if not filas:
        filas = "<tr><td colspan='6'>Aún no hay cambios ni notas.</td></tr>"
    form = ""
    if rol == "Talento" or rol in _HN_GER:
        opts = ("<option value='nota_rrhh'>Nota de RR.HH.</option>" if rol == "Talento" else "<option value='nota_ger'>Nota de Gerencia</option>") + \
               "<option value='presupuesto'>Ajuste de presupuesto</option>"
        form = ("<form method='POST' action='/gerencia/horarios-asesores/nota' class='nx-inline' style='margin-bottom:10px'>%s"
                "<select name='tipo'>%s</select> <input name='objetivo' placeholder='Asesor o rol (opcional)' style='width:190px'> "
                "<input name='nota' placeholder='Escribe la nota' required maxlength='400' style='min-width:260px'> "
                "<input name='impacto' inputmode='numeric' placeholder='Impacto en $ (opcional)' style='width:170px'> "
                "<button class='nx-btn-sm'>Guardar nota</button></form>") % (_n_csrf_input(), opts)
    return ("<section class='nx-card'><h2>📒 Bitácora de cambios, notas y ajustes de presupuesto</h2>"
            "<p class='nx-small'>Aquí queda todo lo que Gerencia cambia en los horarios, con su motivo, y lo que Talento Humano anota para el ajuste de presupuesto.</p>"
            "%s<div class='nx-scroll'><table><tr><th>Fecha</th><th>Quién</th><th>Tipo</th><th>Asesor y cambio</th><th>Nota</th><th>Impacto</th></tr>%s</table></div></section>") % (form, filas)


@app.route("/gerencia/horarios-asesores/nota", methods=["POST"])
def hn_nota():
    if not requiere_login():
        return redirect("/login")
    rol = rol_actual()
    if rol != "Talento" and rol not in _HN_GER:
        return acceso_denegado("Solo Talento Humano y Gerencia pueden dejar notas en la bitácora de horarios.")
    _hn_init()
    if not _n_csrf_ok():
        _n_flash("err", "Formulario inválido, intenta de nuevo.")
        return redirect("/gerencia/horarios-asesores")
    tipo = (request.form.get("tipo") or "").strip()
    if tipo not in ("nota_rrhh", "nota_ger", "presupuesto") or (tipo == "nota_rrhh" and rol != "Talento") or (tipo == "nota_ger" and rol not in _HN_GER):
        tipo = "nota_rrhh" if rol == "Talento" else "nota_ger"
    nota = _hn_una_linea(request.form.get("nota"))
    if len(nota) < 3:
        _n_flash("err", "Escribe la nota.")
        return redirect("/gerencia/horarios-asesores")
    digitos = re.sub(r"[^\d]", "", request.form.get("impacto") or "")[:12]
    impacto = int(digitos) if digitos else 0
    _hn_log(tipo, request.form.get("objetivo") or "", "Nota", "", "", nota, impacto)
    registrar_auditoria("Horarios asesores: nota", "%s · %s" % (_HN_TIPOS.get(tipo, tipo), nota[:150]))
    _n_flash("ok", "Nota guardada en la bitácora.")
    return redirect("/gerencia/horarios-asesores")


_hn_vista_orig = app.view_functions["gerencia_horarios_asesores"]
app.view_functions["gerencia_horarios_asesores"] = _hn_vista


# ───────────── Planes de TiendaTrack: descuento + precio final ─────────────
_HT = {"ok": False}


def _ht_init():
    if _HT["ok"]:
        return
    try:
        _tt_init()
        cols = _hn_cols("tt_planes")
        for nombre, ddl in (("precio_lista", "BIGINT DEFAULT 0"), ("precio_anual_lista", "BIGINT DEFAULT 0"), ("descuento_pct", "REAL DEFAULT 0")):
            if nombre not in cols:
                db.session.execute(text("ALTER TABLE tt_planes ADD COLUMN %s %s" % (nombre, ddl)))
                db.session.commit()
        _HT["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("tt_planes descuento init:", e)


def _ht_pct(v):
    try:
        return max(0.0, min(100.0, round(float(str(v or 0).strip().replace(",", ".")), 4)))
    except Exception:
        return 0.0


_HT_JS = r"""<script>
(function(){
 var F=document.getElementById('ht-card');if(!F)return;var form=F.closest('form');if(!form)return;
 var q=function(n){return form.querySelector('[name="'+n+'"]');};
 var LM=q('pm_lista'),LA=q('pa_lista'),D=q('dto'),FM=document.getElementById('ht-fm'),PM=q('pm'),PA=q('pa'),R=document.getElementById('ht-res');
 if(!LM||!LA||!D||!FM||!PM||!PA||!R)return;
 var num=function(x){var n=parseFloat(String(x).replace(',','.'));return isNaN(n)?0:n;};
 var cop=function(n){return '$'+Math.round(n).toLocaleString('es-CO');};
 var r4=function(n){return Math.round(n*10000)/10000;};
 PM.readOnly=true;PA.readOnly=true;PM.style.background='#f1f5f9';PA.style.background='#f1f5f9';
 function pintar(){
  var lm=Math.max(0,num(LM.value)),la=Math.max(0,num(LA.value)),p=Math.max(0,Math.min(100,num(D.value)));
  var fm=Math.round(lm*(1-p/100)),fa=Math.round(la*(1-p/100));
  PM.value=fm;PA.value=fa;
  var h='Precio lista mensual: <b>'+cop(lm)+'</b><br>';
  if(p>0){h+='Descuento: <b>'+r4(p)+'%</b> (−'+cop(lm-fm)+' al mes)<br>';}
  h+='Precio final mensual: <span class="ht-fin">'+cop(fm)+'</span>';
  if(la>0){h+='<br>Precio final anual: <span class="ht-fin">'+cop(fa)+'</span>'+(p>0?' (ahorro '+cop(la-fa)+' al año)':'');}
  if(p>=100)h+='<br><b style="color:#b45309">Con 100 % el plan queda gratis.</b>';
  R.innerHTML=lm>0?h:'Escribe el precio lista mensual para ver el resultado.';
 }
 function desdeFinal(){var lm=num(LM.value),f=num(FM.value);if(lm>0&&FM.value!==''){D.value=r4(Math.max(0,Math.min(100,(1-f/lm)*100)));}pintar();}
 function desdePct(){var lm=num(LM.value),p=Math.max(0,Math.min(100,num(D.value)));FM.value=lm>0&&p>0?Math.round(lm*(1-p/100)):'';pintar();}
 LM.addEventListener('input',desdePct);LA.addEventListener('input',pintar);D.addEventListener('input',desdePct);FM.addEventListener('input',desdeFinal);
 desdePct();
})();
</script>"""


def _ht_card(lm, la, pct):
    fm = int(round(lm * (1 - pct / 100.0))) if pct > 0 else ""
    return ("<style>.ht{margin:16px 0 4px;border:1px solid #bfdbfe;background:#f0f7ff;border-radius:14px;padding:16px}.ht-t{font-weight:800;color:#0020C2;margin-bottom:4px}"
            ".ht-s{font-size:12px;color:#64748b;margin:0 0 8px}.ht-g{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}"
            ".ht-res{margin-top:12px;background:#fff;border:1px solid #dbeafe;border-radius:12px;padding:12px 14px;font-size:13px;line-height:1.7}"
            ".ht-fin{font-size:20px;font-weight:800;color:#047857}</style>"
            "<div class='ht' id='ht-card'><div class='ht-t'>💸 Descuento y precio final</div>"
            "<p class='ht-s'>Escribe el precio lista y el % de descuento, o el precio final que quieres cobrar: lo demás se calcula solo. "
            "El plan se guarda con el precio final.</p><div class='ht-g'>"
            "<div><label>Precio lista mensual (COP)</label><input type='text' inputmode='numeric' name='pm_lista' value='%d'></div>"
            "<div><label>Precio lista anual (COP, 0 = sin plan anual)</label><input type='text' inputmode='numeric' name='pa_lista' value='%d'></div>"
            "<div><label>Descuento (%%)</label><input type='text' inputmode='decimal' name='dto' value='%s'></div>"
            "<div><label>Precio final mensual que quiero cobrar</label><input type='text' inputmode='numeric' id='ht-fm' value='%s'></div></div>"
            "<div class='ht-res' id='ht-res'></div></div>%s") % (lm, la, ("%g" % pct) if pct else "0", fm, _HT_JS)


def _ht_vista(*a, **k):
    orig = _ht_vista_orig
    _ht_init()
    if request.method == "POST":
        resp = orig(*a, **k)
        try:
            f = request.form
            ok = session.get("ti_msg") == ("ok", "Plan guardado.") or (isinstance(session.get("ti_msg"), (list, tuple)) and list(session.get("ti_msg")) == ["ok", "Plan guardado."])
            if f.get("accion") == "guardar" and ok:
                pct = _ht_pct(f.get("dto"))
                lm = _tt_int(f.get("pm_lista"), 0, 0, 50000000) or _tt_int(f.get("pm"), 0, 0, 50000000)
                la = _tt_int(f.get("pa_lista"), 0, 0, 500000000) or _tt_int(f.get("pa"), 0, 0, 500000000)
                pid = _tt_int(f.get("pid"))
                if not pid:
                    r = _c48_q("SELECT id FROM tt_planes WHERE nombre=:n ORDER BY id DESC LIMIT 1", n=" ".join((f.get("nombre") or "").split())[:80])
                    pid = int(r[0][0]) if r else 0
                if pid:
                    _c48_x("UPDATE tt_planes SET precio_lista=:lm, precio_anual_lista=:la, descuento_pct=:p, precio_mensual=:pm, precio_anual=:pa WHERE id=:i",
                           lm=lm, la=la, p=pct, pm=int(round(lm * (1 - pct / 100.0))), pa=int(round(la * (1 - pct / 100.0))), i=pid)
                    if pct > 0:
                        _tt_aud("plan", "descuento %g%% en plan %s" % (pct, pid))
        except Exception as e:
            db.session.rollback()
            print("tt planes descuento:", repr(e)[:200])
        return resp
    resp = orig(*a, **k)
    html = _hn_html_de(resp)
    if html is None:
        return resp
    try:
        lm = la = 0
        pct = 0.0
        eid = _tt_int(request.args.get("editar")) if request.args.get("editar") else 0
        if eid:
            r = _c48_q("SELECT precio_lista, precio_anual_lista, descuento_pct, precio_mensual, precio_anual FROM tt_planes WHERE id=:i", i=eid)
            if r:
                lm, la, pct = int(r[0][0] or 0), int(r[0][1] or 0), float(r[0][2] or 0)
                lm = lm or int(r[0][3] or 0)
                la = la or int(r[0][4] or 0)
        marca = "<p style='margin-top:14px'><button class='cb-btn g'>Guardar plan</button>"
        if marca in html:
            html = html.replace(marca, _ht_card(lm, la, pct) + marca, 1)
        con_dto = [x for x in _c48_q("SELECT nombre, precio_lista, descuento_pct, precio_mensual FROM tt_planes WHERE COALESCE(descuento_pct,0) > 0 ORDER BY orden, id")]
        if con_dto:
            li = "".join("<li><b>%s</b>: lista %s → <b>%s%%</b> → precio final <b>%s</b> al mes</li>" % (_esc(n), _tt_cop(l or 0), ("%g" % p), _tt_cop(f_)) for n, l, p, f_ in con_dto)
            nota = "<p class='cb-note'><b>Planes con descuento vigente</b></p><ul style='margin:4px 0 10px 18px'>%s</ul>" % li
            m2 = "<p class='cb-note'>Los planes de ejemplo vienen ocultos"
            html = html.replace(m2, nota + m2, 1) if m2 in html else html
        return _hn_poner(resp, html)
    except Exception as e:
        db.session.rollback()
        print("tt planes inyeccion:", repr(e)[:200])
        return resp


_ht_vista_orig = app.view_functions["tt_admin_planes"]
app.view_functions["tt_admin_planes"] = _ht_vista


'''

# ───────────── Reemplazos puntuales ─────────────
REEMPLAZOS = [
    ("botón flotante «Tickets documentales»",
     '        if rol in _ROLES_TICKET and not p.startswith("/tickets-doc"):\n            tk = \'<a id="nw-tk" href="/tickets-doc">📄 Tickets documentales</a>\'\n',
     '        if False and rol in _ROLES_TICKET and not p.startswith("/tickets-doc"):  # [procsis-ajustes-1] ya no flota en todas las pantallas (está en el menú)\n'
     '            tk = \'<a id="nw-tk" href="/tickets-doc">📄 Tickets documentales</a>\'\n'),
    ("filtro de portales por rol (Documentos)",
     '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n        return None  # cada ruta valida su propio acceso\n',
     '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n        return None  # cada ruta valida su propio acceso\n'
     '    if path == "/documentos" or path.startswith(("/documentos/", "/documentos-oficiales", "/docs/area")):\n'
     '        return None  # [procsis-ajustes-1] Documentos por área / oficiales: lo valida _dcx_guard / _dca_guard\n'),
    ("menú: Tickets documentales para las áreas que lo usan",
     '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n',
     '    if rol in _ROLES_TICKET and not any(h.startswith("/tickets-doc") for h, _l in items):  # [procsis-ajustes-1]\n'
     '        items = [x for x in items if x[0] != "/logout"] + [("/tickets-doc", "📄 Tickets documentales")] + [x for x in items if x[0] == "/logout"]\n'
     '    if _dcx_rol_ok(rol) and not any(h.startswith("/documentos-oficiales") for h, _l in items):\n'),
    ("descuento con coma decimal en planes EduTrack",
     'float(form.get("descuento_pct") or 0)',
     'float(str(form.get("descuento_pct") or 0).strip().replace(",", ".") or 0)'),
]

ANCLAS = ["def _guard_talento", "def gerencia_horarios_asesores", "def tt_admin_planes", "def _n_csrf_ok", "def _n_csrf_input", "def _n_flash(", "def _tt_int",
          "def _tt_aud", "def _tt_init", "def _tt_cop", "def _c48_q", "def _c48_x", "def _hl_resumen", "class AccesoExtraordinario", "def registrar_auditoria",
          "def acceso_denegado", "def requiere_login", "def rol_actual", "def ahora", "def _esc", "_HL_ROLES = ", "class Usuario"]


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
    copia = ruta.with_name(ruta.name + ".bak_ajustes_gerencia")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_ajustes_gerencia_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo: ajustes aplicados en %s" % ruta)
    print("Copia de seguridad: %s" % copia)
    if "[procsis-descuentos-planes]" not in texto:
        print("AVISO: no tienes aplicado aplicar_descuentos_planes.py; córrelo también para ver la tarjeta «Descuento y precio final» en los planes de EduTrack.")


if __name__ == "__main__":
    main()

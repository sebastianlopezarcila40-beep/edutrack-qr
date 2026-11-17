#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_panel_tiendatrack.py  ·  PROCSIS · TiendaTrack: Bloque A (panel y ajustes)

REQUIERE haber aplicado antes: aplicar_multiproducto.py  (y, si quieres el menú por producto, aplicar_menu_producto.py).

Uso (en la carpeta de tu app.py):
    python aplicar_panel_tiendatrack.py            # parchea ./app.py
    python aplicar_panel_tiendatrack.py ruta/app.py

Qué hace:
  1. Dashboard TiendaTrack: fila nueva con Negocios activos, MRR del mes, Demos convertidas y Tickets abiertos, y 4 tarjetas
     más en la cuadrícula (Renovaciones próximas 7 días, Comisiones por liquidar, Pagos aprobados del mes, Conversión demo→pago).
  2. Buscador rápido en la barra superior (con TiendaTrack elegido, solo Gerencia): por NIT, nombre del dueño, correo, negocio
     o teléfono. Si hay un solo resultado entra directo a la ficha del negocio. Atajo: tecla «/».
  3. Ficha del negocio (/gerencia/tiendatrack/negocios/<id>): estado, plan, vencimientos, pagos, tickets y NIT/nota editables.
     Agrega la columna nit a tt_cuentas (se crea sola). El evento «registro» de la API ahora acepta también "nit".
  4. Botón del asistente flotante: «Tienda Aura · Ayuda» con TiendaTrack, «PROCSIS · Ayuda» con Todo PROCSIS, «Edu Aura» con EduTrack.
  5. Menú lateral: «Panel TiendaTrack» ya no queda marcado en todas las páginas de TiendaTrack; se resalta la opción de la página abierta.

Es idempotente, crea app.py.bak_panel_tiendatrack y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-panel-tiendatrack]"
REQUIERE = "[procsis-multiproducto]"

BLOQUE = r'''
# [procsis-panel-tiendatrack]
# ═════════════ TIENDATRACK · BLOQUE A: dashboard, buscador de negocios, ficha, botón Aura y menú ═════════════
_PFA = {"ok": False}
_PFA_TICKETS_ABIERTOS = "('RADICADO','ABIERTA','URGENTE','PENDIENTE')"


def _pfa_init():
    """Agrega la columna nit a tt_cuentas (una vez por proceso)."""
    if _PFA["ok"]:
        return
    _tt_init()
    try:
        if db.engine.dialect.name == "sqlite":
            if "nit" not in [r[1] for r in _c48_q("PRAGMA table_info(tt_cuentas)")]:
                _c48_x("ALTER TABLE tt_cuentas ADD COLUMN nit VARCHAR(30) DEFAULT ''")
        else:
            _c48_x("ALTER TABLE tt_cuentas ADD COLUMN IF NOT EXISTS nit VARCHAR(30) DEFAULT ''")
        _PFA["ok"] = True
    except Exception as _e:
        print("pfa init:", repr(_e)[:200])


def _pfa_datos():
    k = _pf_tt_kpis()
    mes = ahora().strftime("%Y-%m") + "-01"
    conv = _c48_q("SELECT COUNT(DISTINCT cuenta_id) FROM tt_pagos WHERE estado='aprobado'")
    convertidas = int(conv[0][0]) if conv else 0
    tk = _c48_q("SELECT COUNT(*) FROM tickets_pqr WHERE producto_id='tiendatrack' AND estado IN " + _PFA_TICKETS_ABIERTOS)
    pm = _c48_q("SELECT COUNT(*) FROM tt_pagos WHERE estado='aprobado' AND creado_en>=:m", m=mes)
    renovaciones = 0
    for _cid, est, dfin, ahasta in _c48_q("SELECT id, estado, demo_fin, activo_hasta FROM tt_cuentas"):
        if _tt_estado({"estado": est, "demo_fin": dfin, "activo_hasta": ahasta}) == "activo" and _tt_dias(ahasta) <= 7:
            renovaciones += 1
    return {"activos": k["activos"], "mrr": k["mrr"], "convertidas": convertidas, "cuentas": k["cuentas"],
            "pct": int(round(100.0 * convertidas / max(k["cuentas"], 1))), "tickets": int(tk[0][0]) if tk else 0,
            "pagos_mes": int(pm[0][0]) if pm else 0, "renovaciones": renovaciones}


def _pfa_hero():
    """Fila de indicadores principales del dashboard de TiendaTrack."""
    try:
        d = _pfa_datos()

        def k(valor, etiqueta, nota=""):
            return ("<div class='cb-card'><div class='tt-k'>%s</div><div class='tt-kl'><b>%s</b>%s</div></div>" % (
                valor, etiqueta, ("<br><span class='cb-note'>%s</span>" % nota) if nota else ""))
        return ("<div class='cb-grid' style='grid-template-columns:repeat(auto-fit,minmax(210px,1fr));margin-bottom:12px'>"
                + k(d["activos"], "Negocios activos", "con plan vigente")
                + k(_tt_cop(d["mrr"]), "MRR del mes", "ingreso recurrente mensual")
                + k(d["convertidas"], "Demos convertidas", "%d%% de los negocios registrados" % d["pct"])
                + k(d["tickets"], "Tickets abiertos", "mesa de ayuda TiendaTrack") + "</div>")
    except Exception as _e:
        print("pfa hero:", repr(_e)[:200])
        return ""


def _pfa_tarjetas_extra():
    """Tarjetas adicionales para la cuadrícula secundaria."""
    try:
        d = _pfa_datos()

        def k(valor, etiqueta):
            return "<div class='cb-card'><div class='tt-k'>%s</div><div class='tt-kl'>%s</div></div>" % (valor, etiqueta)
        return (k(d["renovaciones"], "Renovaciones próximas (7 días)") + k("—", "Comisiones por liquidar (módulo pendiente)")
                + k(d["pagos_mes"], "Pagos aprobados este mes") + k("%d%%" % d["pct"], "Conversión demo → pago"))
    except Exception as _e:
        print("pfa extra:", repr(_e)[:200])
        return ""


# ───────────── Buscador rápido y ficha del negocio ─────────────
def _pfa_tag_estado(est):
    return "<span class='cb-tag %s'>%s</span>" % ({"activo": "on", "demo": "wait"}.get(est, "off"), _esc(est))


@app.route("/gerencia/tiendatrack/negocios/buscar")
def pfa_buscar():
    rol, r = _tt_guard()
    if r is not None:
        return r
    _pfa_init()
    q = " ".join((request.args.get("q") or "").split())[:80]
    caja = ("<form method='GET' action='/gerencia/tiendatrack/negocios/buscar' style='display:flex;gap:8px;margin:0 0 14px'>"
            "<input type='search' name='q' value='%s' placeholder='NIT, nombre del dueño, correo, negocio o teléfono' style='flex:1;padding:10px 12px;border:1px solid #cfd8f3;border-radius:10px'>"
            "<button class='cb-btn g'>Buscar</button></form>") % _esc(q)
    if len(q) < 2:
        return _tt_pag("Buscar negocio", caja + "<div class='cb-card'><p class='cb-note'>Escribe al menos 2 caracteres. Puedes buscar por NIT, dueño, correo, nombre del negocio o teléfono.</p></div>")
    cond = ["LOWER(email) LIKE :q", "LOWER(nombre) LIKE :q", "LOWER(negocio) LIKE :q", "telefono LIKE :q"]
    par = {"q": "%" + q.lower() + "%"}
    digitos = "".join(ch for ch in q if ch.isdigit())
    if len(digitos) >= 4:
        cond.append("REPLACE(REPLACE(REPLACE(COALESCE(nit,''),'.',''),'-',''),' ','') LIKE :n")
        par["n"] = "%" + digitos + "%"
    filas = _c48_q("SELECT id, negocio, nombre, email, nit, telefono, segmento, estado, demo_fin, activo_hasta FROM tt_cuentas WHERE "
                   + " OR ".join(cond) + " ORDER BY id DESC LIMIT 50", **par)
    if len(filas) == 1:
        return redirect("/gerencia/tiendatrack/negocios/%d" % filas[0][0])
    cuerpo = ""
    for cid, neg, nom, mail, nit, tel, seg, est, dfin, ahasta in filas:
        real = _tt_estado({"estado": est, "demo_fin": dfin, "activo_hasta": ahasta})
        cuerpo += ("<tr><td><a href='/gerencia/tiendatrack/negocios/%d'><b>%s</b></a><br><span class='cb-note'>%s</span></td><td>%s<br><span class='cb-note'>%s</span></td>"
                   "<td>%s</td><td>%s</td><td>%s</td></tr>") % (cid, _esc(neg or "—"), _esc(seg or ""), _esc(nom or "—"), _esc(mail or ""),
                                                               _esc(nit or "—"), _esc(tel or "—"), _pfa_tag_estado(real))
    lista = ("<div class='cb-card cb-scroll'><h2>%d resultado(s) para «%s»</h2><table><tr><th>Negocio</th><th>Dueño / correo</th><th>NIT</th><th>Teléfono</th><th>Estado</th></tr>%s</table></div>"
             % (len(filas), _esc(q), cuerpo or "<tr><td colspan='5'>Sin resultados.</td></tr>"))
    return _tt_pag("Buscar negocio", caja + lista)


@app.route("/gerencia/tiendatrack/negocios/<int:cid>", methods=["GET", "POST"])
def pfa_negocio(cid):
    rol, r = _tt_guard()
    if r is not None:
        return r
    _pfa_init()
    c = _tt_cuenta_por("id", cid)
    if not c:
        return _tt_pag("Negocio no encontrado", "<div class='cb-card'><p>No existe ese negocio. <a href='/gerencia/tiendatrack/negocios'>Volver a Negocios y demos</a></p></div>")
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
        else:
            nit = "".join(ch for ch in (request.form.get("nit") or "") if ch.isalnum() or ch in ".- ")[:30].strip()
            nota = " ".join((request.form.get("nota") or "").split())[:500]
            try:
                _c48_x("UPDATE tt_cuentas SET nit=:n, nota=:t WHERE id=:i", n=nit, t=nota, i=cid)
                _tt_aud("negocio", "NIT y nota de la cuenta %d" % cid)
                session["ti_msg"] = ("ok", "Datos del negocio guardados.")
            except Exception:
                session["ti_msg"] = ("err", "No se pudo guardar.")
        return redirect("/gerencia/tiendatrack/negocios/%d" % cid)
    nit = (_c48_q("SELECT COALESCE(nit,'') FROM tt_cuentas WHERE id=:i", i=cid) or [[""]])[0][0]
    est = _tt_estado(c)
    plan = _tt_plan(c["plan_id"]) if c["plan_id"] else None
    fin = c["activo_hasta"] if est == "activo" else c["demo_fin"]
    csrf = _acc_csrf()
    planes = {p["id"]: p["nombre"] for p in _tt_planes()}
    pagos = "".join("<tr><td>%s</td><td>%s · %s</td><td>%s</td><td>%s</td></tr>" % (
        _esc(_tt_fecha(p[0])), _esc(planes.get(p[1], "—")), _esc(p[2]), _tt_cop(int(p[3]) // 100), _esc(p[4]))
        for p in _c48_q("SELECT creado_en, plan_id, periodo, monto_centavos, estado FROM tt_pagos WHERE cuenta_id=:c ORDER BY id DESC LIMIT 10", c=cid))
    tickets = ""
    if c["email"]:
        for t in TicketPQR.query.filter(TicketPQR.email == c["email"], TicketPQR.producto_id == "tiendatrack").order_by(TicketPQR.id.desc()).limit(10).all():
            tickets += "<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (
                _esc(formatear_radicado(t.radicado)), _esc(t.subtipo or t.objeto or ""), _esc(t.prioridad or ""), _esc(t.estado or ""))
    resumen = ("<div class='cb-card'><h2>%s</h2><table>"
               "<tr><td>Dueño</td><td>%s</td></tr><tr><td>Correo</td><td>%s</td></tr><tr><td>Teléfono</td><td>%s</td></tr><tr><td>Tipo de negocio</td><td>%s</td></tr>"
               "<tr><td>Estado</td><td>%s</td></tr><tr><td>Plan</td><td>%s</td></tr><tr><td>%s</td><td>%s (%s día(s))</td></tr><tr><td>Registrado</td><td>%s</td></tr></table></div>") % (
        _esc(c["negocio"] or "Negocio sin nombre"), _esc(c["nombre"] or "—"), _esc(c["email"] or "—"), _esc(c["telefono"] or "—"), _esc(c["segmento"] or "—"),
        _pfa_tag_estado(est), _esc(plan["nombre"] if plan else "Sin plan"), "Vence" if est in ("activo", "demo") else "Venció",
        _esc(_tt_fecha(fin)), _tt_dias(fin) if est in ("activo", "demo") else 0, _esc(_tt_fecha(c["creado_en"])))
    editar = ("<div class='cb-card'><h2>NIT y nota interna</h2><form method='POST'><input type='hidden' name='csrf' value='%s'>"
              "<p><label>NIT / documento<br><input type='text' name='nit' value='%s' maxlength='30' style='width:100%%;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px'></label></p>"
              "<p><label>Nota interna<br><textarea name='nota' rows='3' maxlength='500' style='width:100%%;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px'>%s</textarea></label></p>"
              "<p><button class='cb-btn g'>Guardar</button></p></form></div>") % (_esc(csrf), _esc(nit), _esc(c["nota"] or ""))
    t_pagos = ("<div class='cb-card cb-scroll'><h2>Pagos</h2><table><tr><th>Fecha</th><th>Plan</th><th>Valor</th><th>Estado</th></tr>%s</table></div>"
               % (pagos or "<tr><td colspan='4'>Sin pagos.</td></tr>"))
    t_tick = ("<div class='cb-card cb-scroll'><h2>Tickets de soporte</h2><table><tr><th>Radicado</th><th>Asunto</th><th>Prioridad</th><th>Estado</th></tr>%s</table></div>"
              % (tickets or "<tr><td colspan='4'>Sin tickets.</td></tr>"))
    return _tt_pag("Negocio · " + (c["negocio"] or c["email"] or str(cid)),
                   "<p><a class='cb-btn s o' href='/gerencia/tiendatrack/negocios'>← Negocios y demos</a></p><div class='cb-grid'>" + resumen + editar + "</div>"
                   "<div class='cb-grid'>" + t_pagos + t_tick + "</div>")


# ───────────── Barra superior: buscador + nombre del asistente (Edu Aura / Tienda Aura / PROCSIS) ─────────────
_PFA_JS = r"""<style>
#pfa-bus{display:flex;align-items:center;background:#fff;border:1px solid rgba(0,0,0,.12);border-radius:980px;padding:0 4px 0 12px}
#pfa-bus input{border:0;outline:0;background:transparent;width:250px;padding:7px 4px;font-size:13px;color:#1d1d1f}
#pfa-bus button{border:0;background:transparent;cursor:pointer;font-size:14px;padding:4px 8px}
@media(max-width:900px){#pfa-bus input{width:140px}}@media print{#pfa-bus{display:none!important}}
</style>
<script id="pfa-js">
(function(){
 var MODO="__MODO__",BUS="__BUS__";
 if(BUS==="1"&&!document.getElementById('pfa-bus')){
  var f=document.createElement('form');f.id='pfa-bus';f.method='get';f.action='/gerencia/tiendatrack/negocios/buscar';
  f.innerHTML='<input type="search" name="q" placeholder="Buscar negocio: NIT, dueño o correo" autocomplete="off" minlength="2" required aria-label="Buscar negocio de TiendaTrack"><button type="submit" aria-label="Buscar">🔍</button>';
  var host=document.querySelector('.navbar-actions-right')||document.querySelector('.navbar-top-apple');
  if(host)host.insertBefore(f,host.firstChild);
  document.addEventListener('keydown',function(e){
   var t=((document.activeElement||{}).tagName||'');
   if(e.key==='/'&&!e.ctrlKey&&!e.metaKey&&!/INPUT|TEXTAREA|SELECT/.test(t)){var i=f.querySelector('input');if(i){e.preventDefault();i.focus();}}
  });
 }
 var N=MODO==='tiendatrack'?'Tienda Aura':(MODO==='todo'?'PROCSIS':'');
 if(N){
  var fix=function(){
   [].forEach.call(document.querySelectorAll('.eap-btn,.eap-ttl b,[aria-label*="Edu Aura"]'),function(e){
    var a=e.getAttribute('aria-label');if(a&&a.indexOf('Edu Aura')>=0)e.setAttribute('aria-label',a.replace(/Edu Aura/g,N));
    if(!e.children.length&&e.textContent.indexOf('Edu Aura')>=0)e.textContent=e.textContent.replace(/Edu Aura/g,N);
   });
  };
  fix();
  try{var mo=new MutationObserver(fix);mo.observe(document.body,{childList:true,subtree:true});setTimeout(function(){mo.disconnect();},8000);}catch(x){}
 }
})();
</script>
"""


@app.after_request
def _pfa_inyectar(resp):
    """Agrega el buscador rápido y renombra el asistente según el producto elegido (solo páginas con menú lateral del personal)."""
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
        if p.startswith("/static") or p.startswith("/producto/") or "/pdf" in p:
            return resp
        modo = _pf_modo()
        if modo not in ("tiendatrack", "todo"):
            return resp
        html = resp.get_data(as_text=True)
        if "sidebar-apple-glass" not in html or 'id="pfa-js"' in html or "</body>" not in html:
            return resp
        bus = "1" if (modo == "tiendatrack" and rol_actual() in _DCX_GERENCIA) else "0"
        i = html.rfind("</body>")
        resp.set_data(html[:i] + _PFA_JS.replace("__MODO__", modo).replace("__BUS__", bus) + html[i:])
    except Exception as _e:
        print("pfa inyectar:", _e)
    return resp


'''

ANCLA_RETORNO = '    return _tt_pag("TiendaTrack", tiles + "<div class=\'cb-grid\'>" + integ + recientes + "</div>")\n'
NUEVO_RETORNO = ('    tiles = tiles[:-len("</div>")] + _pfa_tarjetas_extra() + "</div>"  # [procsis-panel-tiendatrack]\n'
                 '    return _tt_pag("TiendaTrack", _pfa_hero() + tiles + "<div class=\'cb-grid\'>" + integ + recientes + "</div>")\n')

ANCLA_ACTIVO = '        elif href == "/gerencia/planes":\n            active = " is-active" if path == "/gerencia/planes" else ""\n'
NUEVO_ACTIVO = ('        elif href == "/gerencia/tiendatrack":  # [procsis-panel-tiendatrack] el Panel no se marca en sus subpáginas\n'
                '            active = " is-active" if path.rstrip("/") == "/gerencia/tiendatrack" else ""\n' + ANCLA_ACTIVO)

ANCLA_API = '            registrar_auditoria("TiendaTrack API · registro", "email=%s" % email)\n'
NUEVO_API = ('            try:  # [procsis-panel-tiendatrack] NIT opcional en el evento registro\n'
             '                _pfa_init()\n'
             '                if _pf_txt(datos.get("nit"), 30):\n'
             '                    _c48_x("UPDATE tt_cuentas SET nit=:n WHERE id=:i", n=_pf_txt(datos.get("nit"), 30), i=c["id"])\n'
             '            except Exception:\n'
             '                pass\n' + ANCLA_API)

REEMPLAZOS = [
    ("dashboard del panel TiendaTrack", ANCLA_RETORNO, NUEVO_RETORNO),
    ("resaltado del menú lateral", ANCLA_ACTIVO, NUEVO_ACTIVO),
    ("evento registro de la API (NIT)", ANCLA_API, NUEVO_API),
]
ANCLAS = ["def tt_admin_panel", "def _tt_guard", "def _tt_pag", "def _tt_aud", "def _tt_cuenta_por", "def _tt_plan(", "def _pf_tt_kpis",
          "def _pf_modo", "def _pf_txt", "def _acc_csrf", "def _ti_csrf_ok", "def formatear_radicado", "_DCX_GERENCIA = "]


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
    if REQUIERE not in texto:
        sys.exit("Primero aplica aplicar_multiproducto.py (este parche usa sus funciones). No toqué nada.")
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

    copia = ruta.with_name(ruta.name + ".bak_panel_tiendatrack")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_panel_tiendatrack_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. Panel TiendaTrack actualizado en %s" % ruta)
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()

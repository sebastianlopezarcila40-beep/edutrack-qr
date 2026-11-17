# -*- coding: utf-8 -*-
"""
aplicar_resolucion_administrativa.py
-------------------------------------
Nuevo módulo en GERENCIA: «🏛️ Resoluciones administrativas» (Gerencia › Comunicaciones y CRTP).

  · Diseño corporativo PROCSIS basado en la resolución administrativa del ejemplo: encabezado en recuadro
    (entidad · logo · N° de resolución / fecha / hoja), título en recuadro, considerando, cuerpo, cronograma
    del proceso (tabla), artículos opcionales, firma, QR de verificación y pie de contacto.
  · SELECTOR DE RESOLUCIONES: al crear, eliges una resolución ya emitida (interna o administrativa) y marcas
    qué partes tomar (título, encabezado/facultades, considerandos, cuerpo, artículos, cronograma). El
    formulario las copia para que las edites, y la nueva guarda la referencia de la resolución base.
  · Código único, QR, huella SHA-256, anulación y descarga PDF / imagen / impresión (igual que los demás
    documentos oficiales). Línea «Elaboró: Sistema PROCSIS (Gestión … / Usuario …)» bajo la firma.

REQUISITO: haber ejecutado antes  aplicar_elaboro_trazabilidad.py

USO:   python aplicar_resolucion_administrativa.py            (modifica ./app.py)
Idempotente, con copia de seguridad (app.py.bak-resadm-FECHA) y verificación de sintaxis.
"""
import py_compile
import shutil
import sys
import time
from pathlib import Path

APP = Path(sys.argv[1] if len(sys.argv) > 1 else "app.py")
MARCA = "# [procsis-resadm]"
NOMBRE = "resadm"


def abortar(msg):
    print("✖ " + msg)
    sys.exit(1)


if not APP.is_file():
    abortar("No encuentro %s. Ejecuta el script en la carpeta de app.py o pásale la ruta." % APP)

crudo = APP.read_bytes().decode("utf-8")
crlf = "\r\n" in crudo
src = crudo.replace("\r\n", "\n")

if MARCA in src:
    print("✔ Ya estaba aplicado (%s). No se cambió nada." % MARCA)
    sys.exit(0)
if "# [procsis-elaboro]" not in src:
    abortar("Primero ejecuta  python aplicar_elaboro_trazabilidad.py  (define la línea «Elaboró» que usa este script).")


def reemplazar(texto, viejo, nuevo, etiqueta, veces=1):
    n = texto.count(viejo)
    if n != veces:
        abortar("Ancla «%s»: esperaba %d coincidencia(s) y hay %d. No se modificó app.py." % (etiqueta, veces, n))
    return texto.replace(viejo, nuevo)


# ───────────────────────── Enganches en el módulo de documentos oficiales ─────────────────────────
# 1) /documentos-oficiales/nuevo/<tipo> acepta el tipo nuevo (solo Gerencia)
src = reemplazar(
    src,
    '    if tipo not in ("oficio", "resolucion", "comunicado"):\n'
    '        return redirect("/documentos-oficiales")\n',
    '    if tipo not in ("oficio", "resolucion", "comunicado", "resolucion_adm"):  # [procsis-resadm]\n'
    '        return redirect("/documentos-oficiales")\n',
    "dcx_nuevo tipos",
)
src = reemplazar(
    src,
    '    if tipo == "resolucion" and rol not in _DCX_GERENCIA:\n'
    '        return acceso_denegado("Las resoluciones solo las emite la Gerencia.")\n',
    '    if tipo in ("resolucion", "resolucion_adm") and rol not in _DCX_GERENCIA:  # [procsis-resadm]\n'
    '        return acceso_denegado("Las resoluciones solo las emite la Gerencia.")\n',
    "dcx_nuevo permiso",
)
# 2) filtros y pestañas de la lista
src = reemplazar(
    src,
    '    if tipo in ("oficio", "resolucion", "comunicado"):\n'
    '        cond.append("tipo=:tp")\n',
    '    if tipo in ("oficio", "resolucion", "comunicado", "resolucion_adm"):  # [procsis-resadm]\n'
    '        cond.append("tipo=:tp")\n',
    "lista filtro tipo",
)
src = reemplazar(
    src,
    '("resolucion", "Resoluciones"), ("comunicado", "Comunicados")))',
    '("resolucion", "Resoluciones"), ("resolucion_adm", "Res. administrativas"), ("comunicado", "Comunicados")))',
    "lista pestañas",
)
# 3) botones «Nueva resolución administrativa»
src = reemplazar(
    src,
    '        nuevos.append(("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"))\n',
    '        nuevos.append(("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"))\n'
    '        nuevos.append(("➕ Nueva resolución administrativa", "/documentos-oficiales/nuevo/resolucion_adm?area=Gerencia"))  # [procsis-resadm]\n',
    "botón lista",
)
src = reemplazar(
    src,
    '[("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"),',
    '[("➕ Nueva resolución", "/documentos-oficiales/nuevo/resolucion?area=Gerencia"),\n'
    '                                                          ("➕ Nueva resolución administrativa", "/documentos-oficiales/nuevo/resolucion_adm?area=Gerencia"),  # [procsis-resadm]',
    "botón comunicaciones",
)
# 4) menú de Gerencia
src = reemplazar(
    src,
    "        ('/gerencia/comunicaciones', '📜 Comunicaciones: resoluciones y oficios'),\n",
    "        ('/gerencia/comunicaciones', '📜 Comunicaciones: resoluciones y oficios'),\n"
    "        ('/gerencia/resoluciones-administrativas', '🏛️ Resoluciones administrativas'),  # [procsis-resadm]\n",
    "menú Gerencia",
)

# ───────────────────────── Bloque nuevo (antes del arranque) ─────────────────────────
BLOQUE = r'''
# [procsis-resadm]
# ═════════════ Gerencia · Resoluciones administrativas (diseño corporativo + selector de resolución base) ═════════════
_RA_CSS = r"""
.dc-ra{padding:0 0 124px;font-family:"Segoe UI",Arial,Helvetica,sans-serif;color:#0B1437;font-size:13.5px;line-height:1.5;background:#fff}
.dc-ra .ra-bar{height:9px;background:linear-gradient(100deg,#00E676,#00B8D4 32%,#0020C2)}
.dc-ra .ra-in{padding:22px 38px 0}
.dc-ra table{border-collapse:collapse;width:100%}
.dc-ra .ra-hd{border:1.6px solid #0B2A6B}
.dc-ra .ra-hd>tbody>tr>td{vertical-align:middle;padding:10px 12px}
.dc-ra .ra-ent{font-weight:800;font-size:12px;line-height:1.5;text-transform:uppercase;color:#0B2A6B;width:36%}
.dc-ra .ra-ent span{display:block;font-weight:600;font-size:10px;letter-spacing:.05em;color:#0020C2;text-transform:none;margin-top:5px}
.dc-ra .ra-lg{width:14%;text-align:center;border-left:1.6px solid #0B2A6B;border-right:1.6px solid #0B2A6B}
.dc-ra .ra-lg img{height:70px;max-width:100%;object-fit:contain}
.dc-ra .ra-hd>tbody>tr>td.ra-nm{padding:0}
.dc-ra .ra-nm table th,.dc-ra .ra-nm table td{padding:8px 10px;border-bottom:1.6px solid #0B2A6B;font-size:12.5px}
.dc-ra .ra-nm table tr:last-child th,.dc-ra .ra-nm table tr:last-child td{border-bottom:0}
.dc-ra .ra-nm table th{text-align:left;font-weight:800;color:#0B2A6B;text-transform:uppercase;width:60%;border-right:1.6px solid #0B2A6B;background:#EAF1FF}
.dc-ra .ra-nm table td{text-align:center;font-weight:800;font-size:15px}
.dc-ra .ra-tit{border:1.6px solid #0B2A6B;border-top:0;padding:11px 14px;font-weight:800;font-size:14.5px;line-height:1.38;text-transform:uppercase;color:#0B2A6B;text-align:justify}
.dc-ra .ra-cu{border:1.6px solid #0B2A6B;margin-top:16px;padding:16px 16px 20px}
.dc-ra .ra-cu p{margin:0 0 10px;text-align:justify}
.dc-ra .ra-enc{font-size:13.5px}
.dc-ra .ra-h{font-weight:800;color:#0B2A6B;text-align:center;margin:14px 0 6px;letter-spacing:.03em;font-size:14px}
.dc-ra .ra-cr{margin:14px 0 12px;border:1.5px solid #0B2A6B}
.dc-ra .ra-cr th,.dc-ra .ra-cr td{border:1px solid #6f8bc4;padding:5px 9px;font-size:12.5px;text-align:left;vertical-align:top}
.dc-ra .ra-cr tr.t th{background:#0B2A6B;color:#fff;text-align:center;letter-spacing:.06em;font-size:12.5px}
.dc-ra .ra-cr tr.s th{background:#EAF1FF;color:#0B2A6B;text-transform:uppercase;text-align:center}
.dc-ra .ra-cr td:first-child{width:58%}
.dc-ra .ra-cierre{text-align:center;font-weight:800;margin:14px 0 6px;color:#0B2A6B}
.dc-ra .ra-dado{text-align:center!important;font-size:12.5px}
.dc-ra .ra-fm{text-align:center;margin-top:18px}
.dc-ra .ra-fm img{max-height:78px;max-width:300px;display:block;margin:0 auto -8px}
.dc-ra .ra-fl{width:310px;max-width:100%;border-top:1.5px solid #0B2A6B;margin:0 auto 5px}
.dc-ra .ra-fm b{display:block;font-size:14px;letter-spacing:.02em;text-transform:uppercase;color:#0B2A6B}
.dc-ra .ra-fm span{display:block;font-size:13px;font-weight:700}
.dc-ra .ra-pie{position:absolute;left:38px;right:38px;bottom:24px;display:flex;gap:14px;align-items:center;border-top:1px solid #b9c7ea;padding-top:9px;font-size:10px;line-height:1.45;color:#334155}
.dc-ra .ra-pie img{width:68px;height:68px;flex:0 0 auto}
.dc-ra .ra-pie .iz{flex:1.1;display:flex;gap:10px;align-items:center}
.dc-ra .ra-pie .iz i{font-style:normal;word-break:break-all;color:#64748b;font-size:9px}
.dc-ra .ra-pie .de{flex:1;text-align:center}
.dc-ra .ra-pie .de b{color:#0B2A6B}
.dc-ra .ra-pie .de em{display:block;font-style:normal;color:#0020C2;font-weight:700;margin-top:3px}
.dc-ra .dc-anulado{z-index:5}
.ra-base{border:1px solid #cfd8f3;background:#f5f8ff;border-radius:14px;padding:14px 16px;margin:0 0 14px}
.ra-base .ra-chk{display:flex;flex-wrap:wrap;gap:6px 16px;margin:10px 0}
.ra-base .ra-chk label{display:inline-flex;gap:6px;align-items:center;font-size:13px;font-weight:600;margin:0}
.ra-base .ra-chk input{width:auto}
.ra-base select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:13px;background:#fff}
.ra-fila{display:grid;grid-template-columns:1fr 220px auto;gap:8px;margin:0 0 8px;align-items:center}
.ra-fila input{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #cfd8f3;border-radius:10px;font-size:13px}
@media(max-width:640px){.ra-fila{grid-template-columns:1fr}}
"""
_i_css = _DCX_CSS.rfind("</style>")
if _i_css > 0 and ".dc-ra{" not in _DCX_CSS:
    _DCX_CSS = _DCX_CSS[:_i_css] + _RA_CSS + _DCX_CSS[_i_css:]

_RA_TIPOS = ("resolucion", "resolucion_adm")
_RA_DEF = {
    "asunto": "POR MEDIO DEL CUAL SE ",
    "quien": "EL GERENTE GENERAL DE PROCSIS,",
    "facultades": ("en uso de sus atribuciones legales, estatutarias y de gobernanza corporativa, especialmente las conferidas por la estructura "
                   "organizacional de la empresa, y"),
    "considerandos": "Que ",
    "cuerpo": "Que, conforme a lo anterior,\n\nPROCSIS informa que ",
    "verbo": "RESUELVE:",
    "cierre": "NOTIFÍQUESE Y CÚMPLASE.",
}


def _dcx_tipo_txt_ra(t):
    return "Resolución administrativa" if t == "resolucion_adm" else _dcx_tipo_txt_base(t)


_dcx_tipo_txt_base = _dcx_tipo_txt
_dcx_tipo_txt = _dcx_tipo_txt_ra


def _dcx_html_resadm(d, qr, url, did):
    x = d.get("datos_j") or {}
    emp = _dcx_empresa()
    try:
        dd = datetime.strptime(d.get("fecha") or "", "%Y-%m-%d")
        fecha_c = dd.strftime("%d/%m/%Y")
        dado = "Dado en %s, a los %d días del mes de %s de %d." % (d.get("ciudad") or "", dd.day, _DCX_MESES[dd.month - 1], dd.year)
    except Exception:
        fecha_c = d.get("fecha") or ""
        dado = "Dado en %s." % (d.get("ciudad") or "")
    try:
        num = "%03d" % int(d.get("numero") or 0)
    except Exception:
        num = str(d.get("numero") or "")
    p = ['<div class="dc-page dc-ra" id="dc-page"><div class="ra-bar"></div><div class="ra-in">']
    if d.get("estado") == "ANULADO":
        p.append('<div class="dc-anulado">ANULADO</div>')
    p.append('<table class="ra-hd"><tr><td class="ra-ent">PROCSIS<br>%s<br>%s<span>innovación que gestiona</span></td>'
             '<td class="ra-lg"><img src="%s" alt="PROCSIS"></td>'
             '<td class="ra-nm"><table><tr><th>Resolución administrativa N°</th><td>%s</td></tr><tr><th>Fecha</th><td>%s</td></tr>'
             '<tr><th>Hoja No</th><td>%s</td></tr></table></td></tr></table>'
             % (_esc(d.get("emisor") or ""), _esc(d.get("ciudad") or ""), _DCX_LOGO, _esc(num), _esc(fecha_c), _esc(x.get("hoja") or "1")))
    p.append('<div class="ra-tit">“%s”</div>' % _esc(d.get("asunto") or ""))
    p.append('<div class="ra-cu">')
    if x.get("quien") or x.get("facultades"):
        p.append('<p class="ra-enc"><b>%s</b> %s</p>' % (_esc(x.get("quien") or ""), _esc(x.get("facultades") or "")))
    if x.get("considerandos"):
        p.append('<div class="ra-h" style="text-align:left;margin-left:0">CONSIDERANDO:</div>' + _dcx_parrafos_md(x["considerandos"]))
    if x.get("cuerpo"):
        p.append(_dcx_parrafos_md(x["cuerpo"]))
    cron = [c for c in (x.get("cronograma") or []) if (c.get("desc") or "").strip()]
    if cron:
        p.append('<table class="ra-cr"><tr class="t"><th colspan="2">CRONOGRAMA DEL PROCESO</th></tr>'
                 '<tr class="s"><th>Descripción</th><th>Fecha</th></tr>%s</table>'
                 % "".join("<tr><td>%s</td><td>%s</td></tr>" % (_esc(c.get("desc")), _esc(c.get("fecha") or "")) for c in cron))
    arts = x.get("articulos") or []
    if arts:
        p.append('<div class="ra-h">%s</div>' % _esc(x.get("verbo") or "RESUELVE:"))
        for i, a in enumerate(arts, 1):
            t = (a.get("titulo") or "").strip().upper().rstrip(".")
            p.append('<p><b>ARTÍCULO %dº.%s</b> %s</p>' % (i, (" " + _esc(t) + ".") if t else "", _esc(a.get("texto") or "").replace("\n", "<br>")))
    if x.get("cierre"):
        p.append('<div class="ra-cierre">%s</div>' % _esc(x["cierre"]))
    p.append('<p class="ra-dado">%s</p>' % _esc(dado))
    firma = ('<img src="%s" alt="Firma">' % d["firma"]) if d.get("firma") else ""
    elab = _proc_elaboro_html(_proc_elaboro_txt("asistida", area=d.get("area"), usuario=d.get("creado_por")), "center")
    p.append('<div class="ra-fm">%s<div class="ra-fl"></div><b>%s</b><span>%s</span>%s</div></div>'
             % (firma, _esc(d.get("firmante") or ""), _esc(d.get("cargo") or ""), elab))
    base = ("<br>Resolución base: %s" % _esc(x["base_codigo"])) if x.get("base_codigo") else ""
    p.append('<div class="ra-pie"><div class="iz"><img src="%s" alt="QR"><div><b>Documento verificable</b><br>%s<br>'
             'Escanee el QR para validar su autenticidad.<br><i>%s</i>%s</div></div>'
             '<div class="de"><b>%s · PROCSIS</b><br>Tel: %s · %s<br>%s<em>“innovación que gestiona”</em></div></div></div></div>'
             % (qr, _esc(d.get("codigo")), _esc(url), base, _esc(d.get("ciudad") or ""), _esc(emp["tel"]), _esc(emp["email"]), _esc(emp["web"])))
    return "".join(p)


_dcx_html_oficio_ra = _dcx_html_oficio


def _dcx_html_oficio(d, qr, url, did):
    if d.get("tipo") == "resolucion_adm":
        return _dcx_html_resadm(d, qr, url, did)
    return _dcx_html_oficio_ra(d, qr, url, did)


# ───────────── Formulario con selector de resolución base ─────────────
_RA_JS = r"""<script>
(function(){
var sel=document.getElementById('ra-sel'),est=document.getElementById('ra-est');
function $(n){return document.querySelector('[name="'+n+'"]');}
function set(n,v){var e=$(n);if(e)e.value=v||'';}
function art(t,x){var d=document.createElement('div');d.className='dcx-art';
 d.innerHTML="<label>Título del artículo</label><input type='text' name='art_titulo'><label>Texto</label><textarea name='art_texto' rows='3'></textarea><button type='button' class='cb-btn r s' onclick='this.parentNode.remove()'>Quitar artículo</button>";
 d.querySelector('[name=art_titulo]').value=t||'';d.querySelector('[name=art_texto]').value=x||'';document.getElementById('ra-arts').appendChild(d);}
function cron(a,f){var d=document.createElement('div');d.className='ra-fila';
 d.innerHTML="<input type='text' name='cron_desc' placeholder='Descripción de la etapa'><input type='text' name='cron_fecha' placeholder='Fecha (ej. 25 de noviembre del 2026)'><button type='button' class='cb-btn r s' onclick='this.parentNode.remove()'>Quitar</button>";
 d.querySelector('[name=cron_desc]').value=a||'';d.querySelector('[name=cron_fecha]').value=f||'';document.getElementById('ra-cron').appendChild(d);}
window.raAddArt=function(){art('','');};window.raAddCron=function(){cron('','');};
function vacio(id){var c=document.getElementById(id);while(c.firstChild)c.removeChild(c.firstChild);}
window.raCargar=function(){
 if(!sel||!sel.value){est.textContent='Elige primero una resolución de la lista.';return;}
 est.textContent='Cargando…';
 fetch('/documentos-oficiales/base/'+encodeURIComponent(sel.value)+'.json',{credentials:'same-origin'}).then(function(r){return r.json();}).then(function(j){
  if(!j.ok){est.textContent=j.error||'No se pudo cargar la resolución.';return;}
  var on=function(k){var e=document.getElementById('ra-p-'+k);return e&&e.checked;};
  if(on('tit'))set('asunto',j.asunto);
  if(on('enc')){set('quien',j.quien);set('facultades',j.facultades);set('verbo',j.verbo||'RESUELVE:');}
  if(on('con'))set('considerandos',j.considerandos);
  if(on('cue'))set('cuerpo',j.cuerpo);
  if(on('art')){vacio('ra-arts');(j.articulos||[]).forEach(function(a){art(a.titulo,a.texto);});}
  if(on('cro')){vacio('ra-cron');(j.cronograma||[]).forEach(function(c){cron(c.desc,c.fecha);});if(!(j.cronograma||[]).length)cron('','');}
  if(on('cie'))set('cierre',j.cierre);
  set('base_codigo',j.codigo);
  est.textContent='✔ Partes cargadas desde '+j.codigo+'. Edítalas a tu gusto antes de emitir.';
 }).catch(function(){est.textContent='No se pudo cargar la resolución. Intenta de nuevo.';});
};
})();
</script>"""


def _resadm_form(area, csrf):
    opts = "".join("<option value='%s'%s>%s</option>" % (_esc(a), " selected" if a == area else "", _esc(_dcx_area_titulo(a))) for a in ("Gerencia", "CRTP"))
    try:
        base = _c48_q("SELECT id, codigo, tipo, asunto, fecha FROM docs_corp WHERE tipo IN ('resolucion','resolucion_adm') AND estado='VIGENTE' "
                      "ORDER BY id DESC LIMIT 150")
    except Exception:
        base = []
    ops = "<option value=''>— Selecciona una resolución —</option>" + "".join(
        "<option value='%d'>%s · %s · %s</option>" % (i, _esc(c), _esc(_dcx_tipo_txt(t)), _esc((a or "")[:70])) for i, c, t, a, _f in base)
    partes = "".join("<label><input type='checkbox' id='ra-p-%s' checked> %s</label>" % (k, n) for k, n in (
        ("tit", "Título (objeto)"), ("enc", "Encabezado y facultades"), ("con", "Considerandos"), ("cue", "Cuerpo"),
        ("art", "Artículos"), ("cro", "Cronograma"), ("cie", "Cierre")))
    selector = ("<div class='ra-base' style='grid-column:1/-1'><b>📚 Tomar partes de una resolución existente (opcional)</b>"
                "<div class='cb-note'>Elige una resolución ya emitida y marca qué partes quieres copiar a esta. Después puedes editar todo el texto.</div>"
                "<div style='margin-top:8px'><select id='ra-sel'>%s</select></div><div class='ra-chk'>%s</div>"
                "<p style='margin:0'><button type='button' class='cb-btn o s' onclick='raCargar()'>⬇ Cargar partes seleccionadas</button> "
                "<span id='ra-est' class='dcx-st'></span></p><input type='hidden' name='base_codigo' value=''></div>") % (ops, partes)
    D = _RA_DEF
    cron = "".join("<div class='ra-fila'><input type='text' name='cron_desc' placeholder='Descripción de la etapa'>"
                   "<input type='text' name='cron_fecha' placeholder='Fecha (ej. 25 de noviembre del 2026)'>"
                   "<button type='button' class='cb-btn r s' onclick='this.parentNode.remove()'>Quitar</button></div>" for _ in range(3))
    campos = (selector
              + "<div><label>Área que emite</label><select name='area'>%s</select></div>" % opts
              + _dcx_campo("Quién emite (encabezado)", "emisor", "GERENCIA GENERAL")
              + _dcx_campo("Número de la resolución", "numero", "", "number", "min='1'", "Vacío = consecutivo automático del año.")
              + _dcx_campo("Hoja No.", "hoja", "1")
              + _dcx_campo("Ciudad", "ciudad", "Caracolí, Antioquia")
              + _dcx_campo("Fecha", "fecha", fecha_hoy(), "date")
              + _dcx_area_txt("Título («Por medio del cual se…»)", "asunto", D["asunto"], 3, "Sale en mayúsculas dentro del recuadro del título.")
              + _dcx_campo("Quien emite (primer párrafo)", "quien", D["quien"])
              + _dcx_area_txt("Facultades", "facultades", D["facultades"], 3)
              + _dcx_area_txt("Considerandos", "considerandos", D["considerandos"], 7, "Separa cada «Que…» con una línea en blanco. **texto** = negrita.")
              + _dcx_area_txt("Cuerpo", "cuerpo", D["cuerpo"], 8, "Información principal de la resolución (convocatoria, requisitos, decisión…). **texto** = negrita.")
              + "<div style='grid-column:1/-1'><label>Cronograma del proceso (opcional)</label><div id='ra-cron'>%s</div>"
                "<button type='button' class='cb-btn o s' onclick='raAddCron()'>➕ Agregar etapa</button>"
                "<div class='cb-note'>Las filas vacías se ignoran. Si no hay etapas, la tabla no se muestra.</div></div>" % cron
              + _dcx_campo("Encabezado de los artículos", "verbo", D["verbo"])
              + "<div style='grid-column:1/-1'><label>Artículos (opcional)</label><div id='ra-arts'></div>"
                "<button type='button' class='cb-btn o s' onclick='raAddArt()'>➕ Agregar artículo</button></div>"
              + _dcx_campo("Cierre", "cierre", D["cierre"])
              + _dcx_campo("Firmante (nombre)", "firmante", "")
              + _dcx_campo("Cargo del firmante", "cargo", "Gerente General – PROCSIS")
              + "<div><label>Firma escaneada (PNG o JPG, opcional)</label><input type='file' name='firma' accept='image/png,image/jpeg'></div>")
    return ("<h1>Nueva resolución administrativa</h1><p class='sub'>Diseño corporativo PROCSIS. Al emitirla recibe un código único y un QR de "
            "verificación, y se descarga en PDF o imagen. Una resolución emitida no se edita: si hay un error, se anula y se emite otra.</p>"
            "<form method='POST' enctype='multipart/form-data' action='/documentos-oficiales/nuevo/resolucion_adm'>"
            "<input type='hidden' name='csrf' value='%s'><input type='hidden' name='area_def' value='%s'>"
            "<div class='cb-card'><div class='dcx-fg'>%s</div>"
            "<p style='margin-top:14px'><button class='cb-btn g'>Emitir resolución administrativa</button> "
            "<a class='cb-btn o' href='/gerencia/resoluciones-administrativas'>Cancelar</a></p></div></form>%s"
            % (_esc(csrf), _esc(area), campos, _RA_JS))


_dcx_form_ra = _dcx_form


def _dcx_form(tipo, area, csrf):
    if tipo == "resolucion_adm":
        return _resadm_form(area, csrf)
    return _dcx_form_ra(tipo, area, csrf)


# ───────────── Creación ─────────────
def _resadm_crear(f, files, usuario, rol):
    if rol not in _DCX_GERENCIA:
        return False, "Las resoluciones administrativas solo las emite la Gerencia."
    area = (f.get("area") or f.get("area_def") or "Gerencia").strip()
    if area not in ("Gerencia", "CRTP"):
        area = "Gerencia"
    fecha = (f.get("fecha") or "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", fecha):
        fecha = fecha_hoy()
    anio = int(fecha[:4])

    def lim(k, n):
        return " ".join((f.get(k) or "").split())[:n]

    def lim_p(k, n):
        return (f.get(k) or "").replace("\r", "").strip()[:n]
    asunto, firmante, cargo = lim("asunto", 600), lim("firmante", 160), lim("cargo", 160)
    if len(asunto) < 8:
        return False, "Escribe el título de la resolución («Por medio del cual se…»)."
    if not firmante:
        return False, "Escribe el nombre de quien firma."
    firma = _dcx_img_de_archivo(files.get("firma") if files else None, 400)
    if firma is None:
        return False, "La firma debe ser una imagen PNG o JPG de máximo 400 KB."
    arts = []
    for t, x in zip(f.getlist("art_titulo"), f.getlist("art_texto")):
        x = (x or "").replace("\r", "").strip()[:2500]
        if x:
            arts.append({"titulo": " ".join((t or "").split())[:160], "texto": x})
    cron = []
    for a, fe in zip(f.getlist("cron_desc"), f.getlist("cron_fecha")):
        a = " ".join((a or "").split())[:220]
        if a:
            cron.append({"desc": a, "fecha": " ".join((fe or "").split())[:120]})
    cron = cron[:30]
    cons, cuerpo = lim_p("considerandos", 6000), lim_p("cuerpo", 6000)
    if cons.strip() in ("", "Que") and cuerpo.strip() in ("", "Que, conforme a lo anterior,\n\nPROCSIS informa que") and not arts and not cron:
        return False, "Escribe al menos los considerandos, el cuerpo, un artículo o el cronograma."
    emisor = lim("emisor", 220) or "GERENCIA GENERAL"
    datos = {"hoja": lim("hoja", 12) or "1", "quien": lim("quien", 220), "facultades": lim_p("facultades", 1200), "considerandos": cons,
             "cuerpo": cuerpo, "cronograma": cron, "verbo": lim("verbo", 40) or "RESUELVE:", "articulos": arts, "cierre": lim("cierre", 80),
             "base_codigo": lim("base_codigo", 48)}
    tipo = "resolucion_adm"
    num = int(f.get("numero")) if (f.get("numero") or "").strip().isdigit() else 0
    if num:
        if _c48_q("SELECT 1 FROM docs_corp WHERE tipo=:t AND area=:a AND anio=:y AND numero=:n", t=tipo, a=area, y=anio, n=num):
            return False, "Ya existe una resolución administrativa con ese número en el año %d." % anio
    else:
        r = _c48_q("SELECT COALESCE(MAX(numero), 0) FROM docs_corp WHERE tipo=:t AND area=:a AND anio=:y", t=tipo, a=area, y=anio)
        num = int((r[0][0] if r else 0) or 0) + 1
    codigo = "PRC-RAD-%s-%d-%04d" % (_DCX_COD.get(area, "GEN"), anio, num)
    token = secrets.token_urlsafe(12)
    reg = {"codigo": codigo, "tipo": tipo, "area": area, "emisor": emisor, "numero": num, "anio": anio, "asunto": asunto, "fecha": fecha,
           "ciudad": lim("ciudad", 100), "datos": json.dumps(datos, ensure_ascii=False), "firmante": firmante, "cargo": cargo}
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


_dcx_crear_ra = _dcx_crear


def _dcx_crear(tipo, f, files, usuario, rol):
    if tipo == "resolucion_adm":
        return _resadm_crear(f, files, usuario, rol)
    return _dcx_crear_ra(tipo, f, files, usuario, rol)


# ───────────── Selector: datos de una resolución para usar como base ─────────────
@app.route("/documentos-oficiales/base/<int:did>.json")
def dcx_base_json(did):
    g = _dcx_guard()
    if g is not None or rol_actual() not in _DCX_GERENCIA:
        return jsonify(ok=False, error="Sin acceso."), 403
    d = _dcx_row(did=did)
    if not d or d.get("tipo") not in _RA_TIPOS or d.get("estado") != "VIGENTE":
        return jsonify(ok=False, error="Resolución no disponible."), 404
    x = d.get("datos_j") or {}
    return jsonify(ok=True, codigo=d["codigo"], asunto=d.get("asunto") or "", quien=x.get("quien") or "", facultades=x.get("facultades") or "",
                   considerandos=x.get("considerandos") or "", cuerpo=x.get("cuerpo") or "", verbo=x.get("verbo") or "", cierre=x.get("cierre") or "",
                   articulos=x.get("articulos") or [], cronograma=x.get("cronograma") or [])


# ───────────── Página del módulo en Gerencia ─────────────
@app.route("/gerencia/resoluciones-administrativas")
def dcx_resadm_lista():
    g = _guard_gerencia()
    if g is not None:
        return g
    _dcx_init()
    if "tipo" not in request.args:
        q = (request.args.get("q") or "").strip()[:60]
        return redirect("/gerencia/resoluciones-administrativas?tipo=resolucion_adm" + (("&q=" + quote_plus(q)) if q else ""))
    return _dcx_pagina_lista("🏛️ Resoluciones administrativas",
                             "Resoluciones con el diseño corporativo de PROCSIS: código único, QR de verificación, descarga en PDF o imagen. "
                             "Al crear una nueva puedes tomar partes de una resolución anterior.",
                             None, "/gerencia/resoluciones-administrativas",
                             [("➕ Nueva resolución administrativa", "/documentos-oficiales/nuevo/resolucion_adm?area=Gerencia")])


'''

ARRANQUE = 'if __name__ == "__main__":\n'
src = reemplazar(src, ARRANQUE, BLOQUE.lstrip("\n") + "\n" + ARRANQUE, "arranque (if __name__)")

# ───────────────────────── Guardar con copia y verificación ─────────────────────────
copia = APP.with_name("%s.bak-%s-%s" % (APP.name, NOMBRE, time.strftime("%Y%m%d%H%M%S")))
shutil.copy2(APP, copia)
salida = src.replace("\n", "\r\n") if crlf else src
APP.write_bytes(salida.encode("utf-8"))
chk = APP.with_name("__chk_%s.pyc" % NOMBRE)
try:
    py_compile.compile(str(APP), doraise=True, cfile=str(chk))
except py_compile.PyCompileError as e:
    shutil.copy2(copia, APP)
    print("✖ Error de sintaxis tras el cambio; se restauró la copia.\n%s" % e)
    sys.exit(1)
finally:
    try:
        chk.unlink()
    except OSError:
        pass
print("✔ Listo. Copia de seguridad: %s" % copia.name)
print("  Gerencia › Comunicaciones y CRTP › «🏛️ Resoluciones administrativas».")

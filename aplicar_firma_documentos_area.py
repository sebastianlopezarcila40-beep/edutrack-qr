# -*- coding: utf-8 -*-
"""
aplicar_firma_documentos_area.py
---------------------------------
«Documentos por área» (/documentos, el que se envía por correo) ahora tiene:

  ✍  Espacio de firma al crear el documento: se DIBUJA con mouse/dedo o se SUBE una imagen (PNG/JPG, el
      fondo blanco se vuelve transparente). Opción «Guardar como mi firma» para reutilizarla (botón
      «Usar mi firma guardada»).
  📄 La firma sale sobre la línea de firma en la vista, el PDF, el Word y, por lo tanto, en el PDF adjunto del correo.
  🧾 La línea «Elaboró: Sistema PROCSIS (Gestión: Área / Usuario: ID).» bajo la entidad, en vista, PDF y Word.

REQUISITO: haber ejecutado antes  aplicar_elaboro_trazabilidad.py

USO:   python aplicar_firma_documentos_area.py            (modifica ./app.py)
Idempotente, con copia de seguridad (app.py.bak-firmadca-FECHA) y verificación de sintaxis.
"""
import py_compile
import shutil
import sys
import time
from pathlib import Path

APP = Path(sys.argv[1] if len(sys.argv) > 1 else "app.py")
MARCA = "# [procsis-firma-dca]"
NOMBRE = "firmadca"


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


# ───────────────────────── Base de datos: columna firma en docs_area ─────────────────────────
src = reemplazar(
    src,
    '"ciudad, fecha, estado, creado_por, creado, leido_por, leido, nota, numero, anio, entidad, ref, cc, anexos, token")',
    '"ciudad, fecha, estado, creado_por, creado, leido_por, leido, nota, numero, anio, entidad, ref, cc, anexos, token, firma")',
    "_DCA_COLS",
)
src = reemplazar(
    src,
    '("anexos", "TEXT DEFAULT \'\'"), ("token", "VARCHAR(40) DEFAULT \'\'")):',
    '("anexos", "TEXT DEFAULT \'\'"), ("token", "VARCHAR(40) DEFAULT \'\'"), ("firma", "TEXT DEFAULT \'\'")):  # [procsis-firma-dca]',
    "ALTER firma",
)

# ───────────────────────── Al emitir: guardar la firma recibida ─────────────────────────
src = reemplazar(
    src,
    '    d = _dca_row(token=token)\n'
    '    registrar_auditoria("Documentos por área", "%s emitió %s para %s (%s)"',
    '    _dca_guardar_firma(token, f, usuario)  # [procsis-firma-dca]\n'
    '    d = _dca_row(token=token)\n'
    '    registrar_auditoria("Documentos por área", "%s emitió %s para %s (%s)"',
    "dca_nuevo → guardar firma",
)

# ───────────────────────── PDF (reportlab) ─────────────────────────
src = reemplazar(
    src,
    '    el.append(Spacer(1, 34))\n'
    '    el.append(HRFlowable(width=6.2 * cm, thickness=0.8, color=colors.black, hAlign="LEFT", spaceAfter=3))\n',
    '    _fb = _dca_firma_bytes(d)  # [procsis-firma-dca]\n'
    '    _puesta = False\n'
    '    if _fb:\n'
    '        try:\n'
    '            _iw, _ih = ImageReader(BytesIO(_fb)).getSize()\n'
    '            _ancho = min(6.2 * cm, 1.7 * cm * _iw / float(_ih))\n'
    '            el.append(Spacer(1, 12))\n'
    '            el.append(_RLImage(BytesIO(_fb), width=_ancho, height=_ancho * _ih / float(_iw), hAlign="LEFT"))\n'
    '            _puesta = True\n'
    '        except Exception as _e_f:\n'
    '            print("firma pdf:", _e_f)\n'
    '    if not _puesta:\n'
    '        el.append(Spacer(1, 34))\n'
    '    el.append(HRFlowable(width=6.2 * cm, thickness=0.8, color=colors.black, hAlign="LEFT", spaceAfter=3))\n',
    "PDF firma",
)
src = reemplazar(
    src,
    '    el.append(Paragraph("<b><font color=\'#0B2A6B\'>%s</font></b>" % t(d.get("entidad")), n))\n',
    '    el.append(Paragraph("<b><font color=\'#0B2A6B\'>%s</font></b>" % t(d.get("entidad")), n))\n'
    '    el.append(Paragraph(t(_proc_elaboro_txt("asistida", area=d.get("origen"), usuario=d.get("creado_por"))),\n'
    '                        ParagraphStyle("dz", parent=n, fontSize=7.5, leading=10, textColor=colors.HexColor("#666666"), spaceBefore=4)))  # [procsis-firma-dca]\n',
    "PDF elaboró",
)

# ───────────────────────── Word (python-docx) ─────────────────────────
src = reemplazar(
    src,
    '    par("")\n'
    '    par("______________________________")\n',
    '    par("")\n'
    '    _fb = _dca_firma_bytes(d)  # [procsis-firma-dca]\n'
    '    if _fb:\n'
    '        try:\n'
    '            doc.add_paragraph().add_run().add_picture(BytesIO(_fb), height=Inches(0.7))\n'
    '        except Exception as _e_f:\n'
    '            print("firma docx:", _e_f)\n'
    '    par("______________________________")\n',
    "DOCX firma",
)
src = reemplazar(
    src,
    '    pe = par(d.get("entidad") or "", bold=True)\n'
    '    for rr in pe.runs:\n'
    '        rr.font.color.rgb = RGBColor(0x0B, 0x2A, 0x6B)\n',
    '    pe = par(d.get("entidad") or "", bold=True)\n'
    '    for rr in pe.runs:\n'
    '        rr.font.color.rgb = RGBColor(0x0B, 0x2A, 0x6B)\n'
    '    _pel = par(_proc_elaboro_txt("asistida", area=d.get("origen"), usuario=d.get("creado_por")), after=0)  # [procsis-firma-dca]\n'
    '    _pel.paragraph_format.space_before = Pt(3)\n'
    '    for rr in _pel.runs:\n'
    '        rr.font.size = Pt(7.5)\n'
    '        rr.font.bold = False\n'
    '        rr.font.color.rgb = RGBColor(0x66, 0x66, 0x66)\n',
    "DOCX elaboró",
)

# ───────────────────────── Bloque nuevo (antes del arranque) ─────────────────────────
BLOQUE = r'''
# [procsis-firma-dca]
# ═════════════ Documentos por área: firma dibujada/subida + línea «Elaboró» ═════════════
_DCA_FIRMAS = {"ok": False}


def _dca_firmas_init():
    """Tabla con la firma guardada de cada usuario (opcional, la crea el propio usuario)."""
    if _DCA_FIRMAS["ok"]:
        return
    try:
        db.session.execute(text("CREATE TABLE IF NOT EXISTS docs_area_firmas (usuario VARCHAR(80) PRIMARY KEY, firma TEXT, actualizado VARCHAR(19))"))
        db.session.commit()
        _DCA_FIRMAS["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("docs_area_firmas init:", e)


def _dca_firma_valida(uri):
    """Acepta solo PNG en data URI de hasta ~400 KB. Devuelve el mismo texto o ''."""
    import base64 as _b64
    uri = (uri or "").strip()
    m = re.match(r"^data:image/png;base64,([A-Za-z0-9+/=]+)$", uri)
    if not m or len(uri) > 420000:
        return ""
    try:
        raw = _b64.b64decode(m.group(1), validate=True)
    except Exception:
        return ""
    return uri if raw[:8] == b"\x89PNG\r\n\x1a\n" else ""


def _dca_firma_bytes(d):
    import base64 as _b64
    uri = _dca_firma_valida((d or {}).get("firma"))
    if not uri:
        return None
    try:
        return _b64.b64decode(uri.split(",", 1)[1])
    except Exception:
        return None


def _dca_firma_guardada(usuario):
    if not usuario:
        return ""
    try:
        _dca_firmas_init()
        r = _c48_q("SELECT firma FROM docs_area_firmas WHERE usuario=:u", u=usuario)
        return _dca_firma_valida(r[0][0]) if r else ""
    except Exception:
        db.session.rollback()
        return ""


def _dca_guardar_firma(token, f, usuario):
    """Guarda la firma del formulario en el documento recién emitido y, si lo pidió, como firma del usuario."""
    try:
        firma = _dca_firma_valida(f.get("firma_data"))
        if not firma:
            return
        _c48_x("UPDATE docs_area SET firma=:f WHERE token=:t", f=firma, t=token)
        if (f.get("firma_guardar") or "") == "1" and usuario:
            _dca_firmas_init()
            _c48_x("DELETE FROM docs_area_firmas WHERE usuario=:u", u=usuario)
            _c48_x("INSERT INTO docs_area_firmas (usuario, firma, actualizado) VALUES (:u, :f, :a)", u=usuario, f=firma, a=str(ahora())[:19])
    except Exception as e:
        db.session.rollback()
        print("docs_area firma:", e)


# Vista HTML: firma sobre la línea y «Elaboró» bajo la entidad
_dca_html_firma_orig = _dca_html


def _dca_html(d):
    h = _dca_html_firma_orig(d)
    try:
        uri = _dca_firma_valida((d or {}).get("firma"))
        if uri:
            h = h.replace("<div class='dca-firma'><div class='l'></div>",
                          "<div class='dca-firma'><img src='%s' alt='Firma' style='display:block;max-height:72px;max-width:260px;margin:0 0 -6px'>"
                          "<div class='l'></div>" % uri, 1)
        i = h.find("<div class='ent'>")
        j = h.find("</div>", i) if i >= 0 else -1
        if j > 0:
            j += 6
            h = h[:j] + _proc_elaboro_html(_proc_elaboro_txt("asistida", area=d.get("origen"), usuario=d.get("creado_por")), "left") + h[j:]
    except Exception as e:
        print("dca html firma:", e)
    return h


# Formulario: recuadro de firma
_DCA_FIRMA_WIDGET = r"""<div class='dca-full dca-env' id='dcf-box'>
<b>✍ Firma del documento (opcional)</b>
<div class='cb-note'>Dibújala con el mouse o el dedo, o sube una imagen de tu firma. Aparece sobre la línea de firma en el PDF, el Word y el correo.</div>
<div class='dca-tb' style='margin-top:8px'>
<button type='button' id='dcf-borrar'>🧽 Borrar</button>
<label style='border:1px solid #cfd8f3;background:#fff;border-radius:8px;padding:6px 11px;cursor:pointer;font-size:13px;margin:0;display:inline-block'>📁 Subir imagen
<input type='file' id='dcf-file' accept='image/png,image/jpeg' style='display:none'></label>
<button type='button' id='dcf-guardada' style='display:none'>⭐ Usar mi firma guardada</button>
<button type='button' id='dcf-olvidar' style='display:none;color:#b91c1c'>Quitar mi firma guardada</button>
</div>
<canvas id='dcf-canvas' width='640' height='200' style='width:100%;max-width:480px;height:auto;aspect-ratio:16/5;border:1px dashed #94a3b8;border-radius:10px;background:#fff;touch-action:none;cursor:crosshair;display:block'></canvas>
<input type='hidden' name='firma_data' id='dcf-data' value=''>
<label style='display:flex;gap:8px;align-items:center;margin-top:8px;font-weight:600'><input type='checkbox' name='firma_guardar' value='1' style='width:auto'> Guardar como mi firma para próximos documentos</label>
</div>
<script>
(function(){
var cv=document.getElementById('dcf-canvas'),dat=document.getElementById('dcf-data');if(!cv)return;
var cx=cv.getContext('2d'),dib=false,hay=false,G=window.DCF_GUARDADA||'';
cx.lineWidth=2.8;cx.lineCap='round';cx.lineJoin='round';cx.strokeStyle='#0B1437';
function pos(e){var r=cv.getBoundingClientRect();return{x:(e.clientX-r.left)*cv.width/r.width,y:(e.clientY-r.top)*cv.height/r.height};}
function guardar(){try{dat.value=hay?cv.toDataURL('image/png'):'';}catch(e){dat.value='';}}
function limpiar(){cx.clearRect(0,0,cv.width,cv.height);hay=false;dat.value='';}
cv.addEventListener('pointerdown',function(e){e.preventDefault();try{cv.setPointerCapture(e.pointerId);}catch(x){}dib=true;var p=pos(e);cx.beginPath();cx.moveTo(p.x,p.y);cx.lineTo(p.x+0.01,p.y);cx.stroke();});
cv.addEventListener('pointermove',function(e){if(!dib)return;e.preventDefault();var p=pos(e);cx.lineTo(p.x,p.y);cx.stroke();});
function fin(){if(!dib)return;dib=false;hay=true;guardar();}
['pointerup','pointercancel','pointerleave'].forEach(function(n){cv.addEventListener(n,fin);});
function poner(src,quitarBlanco){var im=new Image();im.onload=function(){limpiar();var s=Math.min(cv.width/im.width,cv.height/im.height),w=im.width*s,h=im.height*s;
 cx.drawImage(im,(cv.width-w)/2,(cv.height-h)/2,w,h);
 if(quitarBlanco){var id=cx.getImageData(0,0,cv.width,cv.height),p=id.data;for(var i=0;i<p.length;i+=4){var l=0.299*p[i]+0.587*p[i+1]+0.114*p[i+2];if(l>232){p[i+3]=0;}else if(l>190){p[i+3]=Math.round(p[i+3]*(232-l)/42);}}cx.putImageData(id,0,0);}
 hay=true;guardar();};im.src=src;}
document.getElementById('dcf-borrar').onclick=limpiar;
document.getElementById('dcf-file').onchange=function(){var f=this.files&&this.files[0];if(!f)return;if(f.size>3*1024*1024){alert('La imagen pesa más de 3 MB.');this.value='';return;}
 var r=new FileReader();r.onload=function(){poner(r.result,true);};r.readAsDataURL(f);this.value='';};
if(G){var b=document.getElementById('dcf-guardada'),o=document.getElementById('dcf-olvidar');b.style.display='';o.style.display='';
 b.onclick=function(){poner(G,false);};
 o.onclick=function(){if(!confirm('¿Quitar tu firma guardada?'))return;var f=document.createElement('form');f.method='POST';f.action='/documentos/firma/borrar';
  var i=document.createElement('input');i.type='hidden';i.name='csrf';i.value=document.querySelector('input[name=csrf]').value;f.appendChild(i);document.body.appendChild(f);f.submit();};}
})();
</script>"""

_dca_form_firma_orig = _dca_form


def _dca_form(csrf):
    h = _dca_form_firma_orig(csrf)
    try:
        ancla = "<div class='dca-full'><label>Firma como (entidad que emite)</label>"
        if ancla in h:
            g = _dca_firma_guardada(session.get("usuario") or "")
            js = "<script>window.DCF_GUARDADA=%s;</script>" % json.dumps(g)
            h = h.replace(ancla, js + _DCA_FIRMA_WIDGET + ancla, 1)
    except Exception as e:
        print("dca form firma:", e)
    return h


@app.route("/documentos/firma/borrar", methods=["POST"])
def dca_firma_borrar():
    g = _dca_guard()
    if g is not None:
        return g
    if _ti_csrf_ok():
        try:
            _dca_firmas_init()
            _c48_x("DELETE FROM docs_area_firmas WHERE usuario=:u", u=session.get("usuario") or "")
            session["ti_msg"] = ("ok", "Tu firma guardada se eliminó.")
        except Exception:
            db.session.rollback()
            session["ti_msg"] = ("err", "No se pudo eliminar la firma guardada.")
    return redirect("/documentos/nuevo")


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
print("  Documentos por área → Nuevo documento: ya tiene el recuadro de firma y la línea «Elaboró».")

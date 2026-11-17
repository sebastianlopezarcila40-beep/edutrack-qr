# -*- coding: utf-8 -*-
"""
aplicar_elaboro_trazabilidad.py
--------------------------------
Agrega la línea de trazabilidad «Elaboró: …» (gris, 7–8 pt, a 5 px bajo el cargo, alineada a la izquierda)
a los documentos generados por PROCSIS:

  · Documentos oficiales: resoluciones, comunicados y oficios  (modo «asistida»: Área + Usuario)
  · Respuesta oficial de PQR y resolución de recursos           (modo «oficial»: código de trazabilidad)

Define además las funciones compartidas que usan los otros dos scripts:
  _proc_elaboro_txt(modo, area, usuario, codigo)   -> texto de la línea
  _proc_elaboro_html(texto, align)                 -> <div> listo para los documentos HTML

USO (en la carpeta donde está app.py):
    python aplicar_elaboro_trazabilidad.py            # modifica ./app.py
    python aplicar_elaboro_trazabilidad.py ruta\\app.py

Es seguro volver a ejecutarlo: si ya está aplicado, no hace nada.
Hace copia de seguridad (app.py.bak-elaboro-FECHA) y verifica la sintaxis; si algo falla, restaura la copia.
"""
import py_compile
import shutil
import sys
import time
from pathlib import Path

APP = Path(sys.argv[1] if len(sys.argv) > 1 else "app.py")
MARCA = "# [procsis-elaboro]"
NOMBRE = "elaboro"


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


def reemplazar(texto, viejo, nuevo, etiqueta, veces=1):
    n = texto.count(viejo)
    if n != veces:
        abortar("Ancla «%s»: esperaba %d coincidencia(s) y hay %d. No se modificó app.py." % (etiqueta, veces, n))
    return texto.replace(viejo, nuevo)


# ───────────────────────── 1) PQR: respuesta oficial (canvas reportlab) ─────────────────────────
src = reemplazar(
    src,
    '    pdf.drawString(margen, y, "Mesa de Control y Soporte Corporativo")\n'
    '    y -= 12\n'
    '    pdf.setFont("Helvetica", 9)\n'
    '    pdf.drawString(margen, y, f"{nombre_empresa()} · {nombre_producto()}")\n',
    '    pdf.drawString(margen, y, "Mesa de Control y Soporte Corporativo")\n'
    '    y -= 12\n'
    '    pdf.setFont("Helvetica", 9)\n'
    '    pdf.drawString(margen, y, f"{nombre_empresa()} · {nombre_producto()}")\n'
    '    y -= 11  # [procsis-elaboro]\n'
    '    pdf.setFont("Helvetica", 7.5)\n'
    '    pdf.setFillColor(colors.HexColor("#666666"))\n'
    '    pdf.drawString(margen, y, _proc_elaboro_txt("oficial", codigo=str(radicado_no or "")))\n'
    '    pdf.setFillColor(colors.HexColor("#0f172a"))\n',
    "PQR respuesta oficial",
)

# ───────────────────────── 2) PQR: resolución de recursos ─────────────────────────
src = reemplazar(
    src,
    '    linea("Firma: %s · %s" % (firmante or "-", rol or "-"), 10.5, True, 16)\n',
    '    linea("Firma: %s · %s" % (firmante or "-", rol or "-"), 10.5, True, 16)\n'
    '    st["y"] += 4  # [procsis-elaboro]\n'
    '    linea(_proc_elaboro_txt("oficial", codigo=str(getattr(rec, "radicado_recurso", "") or "")), 7.5, False, 11, (0.4, 0.4, 0.4))\n',
    "PQR resolución de recursos",
)

# ───────────────────────── 3) Bloque nuevo (antes del arranque) ─────────────────────────
BLOQUE = r'''
# [procsis-elaboro]
# ═════════════ Trazabilidad: línea «Elaboró:» bajo la firma de los documentos generados ═════════════
#  · auto     → Elaboró: Sistema PROCSIS — Módulo de Gestión Documental.
#  · asistida → Elaboró: Sistema PROCSIS (Gestión: [Área] / Usuario: [ID]).
#  · oficial  → Generado automáticamente por el Sistema PROCSIS · Código de trazabilidad: [ID-ÚNICO].
def _elab_usuario(u):
    u = str(u or "").strip()
    if "@" in u:
        u = u.split("@")[0]
    return re.sub(r"[^A-Za-z0-9._-]", "", u).upper()[:24] or "SISTEMA"


def _proc_elaboro_txt(modo="auto", area="", usuario="", codigo=""):
    if modo == "asistida":
        try:
            a = _dcx_area_titulo(area) if area else ""
        except Exception:
            a = str(area or "")
        return "Elaboró: Sistema PROCSIS (Gestión: %s / Usuario: %s)." % (a or "—", _elab_usuario(usuario))
    if modo == "oficial":
        t = "Generado automáticamente por el Sistema PROCSIS"
        if codigo:
            t += " · Código de trazabilidad: %s" % codigo
        return t + "."
    return "Elaboró: Sistema PROCSIS — Módulo de Gestión Documental."


def _proc_elaboro_html(texto, align="left"):
    return ("<div class='proc-elab' style='font:9.5px/1.3 Arial,Helvetica,sans-serif;color:#666666;font-weight:400;"
            "font-style:normal;text-transform:none;letter-spacing:0;margin-top:5px;text-align:%s'>%s</div>") % (align, _esc(texto))


def _dcx_inyectar_elaboro(html, d, ancla, align="left"):
    """Inserta la línea «Elaboró» al final del bloque de firma del documento HTML (sin tocar el diseño)."""
    try:
        txt = _proc_elaboro_txt("asistida", area=d.get("area"), usuario=d.get("creado_por"))
        linea = _proc_elaboro_html(txt, align)
        if ancla == "dc-firma":
            i = html.find('<div class="dc-fl"></div>')
            j = html.find("</div>", i + 25) if i >= 0 else -1
            if j > 0:
                return html[:j] + linea + html[j:]
        elif ancla == "dc-cfirma":
            i = html.find('class="dc-cfirma"')
            j = html.find("<span></span></div>", i) if i >= 0 else -1
            if j > 0:
                j += len("<span></span></div>")
                return html[:j] + linea + html[j:]
    except Exception as e:
        print("elaboro:", e)
    return html


_dcx_html_resolucion_elab = _dcx_html_resolucion
_dcx_html_comunicado_elab = _dcx_html_comunicado
_dcx_html_oficio_elab = _dcx_html_oficio


def _dcx_html_resolucion(d, qr, url):
    return _dcx_inyectar_elaboro(_dcx_html_resolucion_elab(d, qr, url), d, "dc-firma", "center")


def _dcx_html_comunicado(d, qr, url):
    return _dcx_inyectar_elaboro(_dcx_html_comunicado_elab(d, qr, url), d, "dc-cfirma", "left")


def _dcx_html_oficio(d, qr, url, did):
    return _dcx_inyectar_elaboro(_dcx_html_oficio_elab(d, qr, url, did), d, "dc-firma", "left")


'''

ARRANQUE = 'if __name__ == "__main__":\n'
src = reemplazar(src, ARRANQUE, BLOQUE.lstrip("\n") + "\n" + ARRANQUE, "arranque (if __name__)")

# ───────────────────────── Guardar con copia de seguridad y verificación ─────────────────────────
copia = APP.with_name("%s.bak-%s-%s" % (APP.name, NOMBRE, time.strftime("%Y%m%d%H%M%S")))
shutil.copy2(APP, copia)
salida = src.replace("\n", "\r\n") if crlf else src
APP.write_bytes(salida.encode("utf-8"))
try:
    py_compile.compile(str(APP), doraise=True, cfile=str(APP.with_name("__chk_%s.pyc" % NOMBRE)))
except py_compile.PyCompileError as e:
    shutil.copy2(copia, APP)
    print("✖ Error de sintaxis tras el cambio; se restauró la copia.\n%s" % e)
    sys.exit(1)
finally:
    try:
        APP.with_name("__chk_%s.pyc" % NOMBRE).unlink()
    except OSError:
        pass
print("✔ Listo. Copia de seguridad: %s" % copia.name)
print("  La línea «Elaboró» ya sale en resoluciones, comunicados y oficios, y en las respuestas/recursos PQR.")

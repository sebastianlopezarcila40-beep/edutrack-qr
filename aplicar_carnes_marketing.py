# -*- coding: utf-8 -*-
"""
aplicar_carnes_marketing.py
  1) Carnés del equipo (RR.HH.): «Generar carné» ya no te saca al login; abre el carné.
     (Causa: el visor de carné solo dejaba entrar a Soporte/Gerencia y mandaba a Talento a /login.)
  2) Marketing: deja de ser redirigido al intentar abrir «Contenido del sitio» (pie de página y menús),
     que es donde se editan los textos de las columnas del pie (Procsis y la Educación, Corporativo…).
     También aparece como tarjeta en el portal de Marketing, TI y Jurídica.

Uso:  python aplicar_carnes_marketing.py        (junto a app.py)
Deja un respaldo: app.py.respaldo_carnes
"""
import os
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "app.py")
MARCA = "# [procsis-carnes-marketing]"


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
    shutil.copyfile(RUTA, RUTA + ".respaldo_carnes")

    # 1) Visor de carné: también para Talento Humano
    src = reemplazar(
        src,
        '    """Carné corporativo Procsis — diseño ejecutivo legible (anverso + reverso)."""\n'
        '    if not requiere_soporte_global():\n'
        '        return redirect("/login")\n',
        '    """Carné corporativo Procsis — diseño ejecutivo legible (anverso + reverso)."""\n'
        '    ' + MARCA + ' Talento Humano también puede ver/imprimir carnés\n'
        '    if not (requiere_soporte_global() or (requiere_login() and rol_actual() == "Talento")):\n'
        '        return redirect("/login")\n',
        "visor de carné",
    )

    # 2a) Filtro de rutas: «Contenido del sitio» valida su propio acceso (_sw_guard / _sw_puede por área)
    src = reemplazar(
        src,
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        '    if path.startswith("/sitio/contenido"):\n'
        '        return None  # [procsis-sitio] _sw_guard valida el acceso; cada área edita solo sus enlaces\n'
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        "filtro de rutas",
    )

    # 2b) Tarjeta en el portal de las áreas que editan el sitio
    src = reemplazar(
        src,
        '    filas = "".join(\n'
        '        "<tr><td><b>%s</b></td><td>%s</td><td>%s</td></tr>" % (\n'
        '            _esc(u.usuario), _esc(getattr(u, "nombre_completo", "") or "—"),\n',
        '    if rol_area in ("Marketing", "Tecnologia", "Juridica"):\n'
        '        tarjetas += ("<a href=\'/sitio/contenido\' class=\'role-panel\' style=\'border-left:6px solid %s;text-decoration:none;display:block\'>"\n'
        '                     "<div style=\'font-weight:800;color:#0f172a\'>🧩 Contenido del sitio</div>"\n'
        '                     "<div style=\'font-size:12px;color:#475569;margin-top:4px\'>Textos y enlaces del pie de página y del menú de la web. Cada área edita solo los suyos.</div>"\n'
        '                     "<div style=\'margin-top:8px\'><span style=\'background:#ede9fe;color:#5b21b6;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700\'>Abrir</span></div></a>" % a["color"])\n'
        '    filas = "".join(\n'
        '        "<tr><td><b>%s</b></td><td>%s</td><td>%s</td></tr>" % (\n'
        '            _esc(u.usuario), _esc(getattr(u, "nombre_completo", "") or "—"),\n',
        "tarjeta portal",
    )

    if crlf:
        src = src.replace("\n", "\r\n")
    with open(RUTA, "w", encoding="utf-8", newline="") as f:
        f.write(src)
    print("Listo. Parche aplicado. Respaldo en app.py.respaldo_carnes")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
pie_pagina.py  ·  [procsis-pie-legal-v1]

Cambia el pie de página "EduTrack · Desarrollado por Procsis" por un pie legal
con el formato de las operadoras (estilo Tigo):

  © EduTrack Colombia 2026. Servicios de software educativo prestados por PROCSIS.
  Todos los derechos reservados.
  Términos y condiciones | Política de protección de datos personales | Mapa del sitio

También crea las páginas /terminos y /mapa-del-sitio.

NO modifica app.py: se engancha con  pie_pagina.instalar(globals())
Los textos salen de Soporte → Marca (nombre del producto y de la empresa).
Para cambiar la redacción legal edita las constantes de abajo.
"""
import html as _html
import re
from datetime import datetime

PAIS = "Colombia"
SERVICIO = "Servicios de software educativo prestados por"   # → "... prestados por PROCSIS."
URL_TERMINOS = "/terminos"
URL_DATOS = "/politicas/datos-personales"
URL_MAPA = "/mapa-del-sitio"

_R_FOOTER_DIV = re.compile(r'<div class="footer">(?:(?!</div>).)*?Desarrollado por(?:(?!</div>).)*?</div>', re.S)
_R_FOOTER_LP = re.compile(r'<footer class="lp-footer">.*?Desarrollado por.*?</footer>', re.S)

_ESTILO_LINK = "color:inherit;text-decoration:underline;font-weight:600"


def construir_pie(producto, empresa, anio=None):
    """HTML interno del pie (sin contenedor)."""
    anio = anio or datetime.now().year
    p, e = _html.escape(producto), _html.escape(empresa)
    return (
        '<span class="pie-legal" style="display:block;line-height:1.7;font-size:12.5px">'
        "© %s %s %s. %s %s. Todos los derechos reservados.<br>"
        '<a href="%s" style="%s">Términos y condiciones</a> | '
        '<a href="%s" style="%s">Política de protección de datos personales</a> | '
        '<a href="%s" style="%s">Mapa del sitio</a>'
        "</span>" % (p, _html.escape(PAIS), anio, _html.escape(SERVICIO), e,
                     URL_TERMINOS, _ESTILO_LINK, URL_DATOS, _ESTILO_LINK, URL_MAPA, _ESTILO_LINK)
    )


def reemplazar_pies(html, producto, empresa):
    """Sustituye los pies antiguos 'Desarrollado por…' dentro de un HTML. Devuelve el HTML nuevo."""
    if "Desarrollado por" not in html:
        return html
    interno = construir_pie(producto, empresa)
    html = _R_FOOTER_DIV.sub(lambda m: '<div class="footer">%s</div>' % interno, html)
    html = _R_FOOTER_LP.sub(lambda m: '<footer class="lp-footer">%s</footer>' % interno, html)
    return html


_TERMINOS = [
    ("1. Objeto", "Estos términos regulan el acceso y uso de la plataforma {producto} (la «Plataforma»), "
                  "un servicio de software educativo prestado por {empresa}."),
    ("2. Aceptación", "Al ingresar, registrarse o usar la Plataforma, la persona usuaria declara que leyó y acepta estos "
                      "términos y la Política de protección de datos personales. Si no está de acuerdo, debe abstenerse de usarla."),
    ("3. Uso adecuado", "La Plataforma debe usarse únicamente para fines educativos y administrativos lícitos. Está prohibido "
                        "intentar acceder sin autorización, alterar o extraer información de otras instituciones, cargar "
                        "contenido ilícito o malicioso, o suplantar a otra persona."),
    ("4. Cuentas y seguridad", "Cada usuario es responsable de custodiar sus credenciales y de toda actividad realizada con ellas. "
                               "Debe informar de inmediato cualquier uso no autorizado."),
    ("5. Peticiones, quejas y reclamos (PQR)", "Las solicitudes radicadas por los canales habilitados reciben un número de "
                                              "radicado y se responden dentro de los plazos legales aplicables (15 días hábiles). "
                                              "Para radicar se verifica la identidad mediante un PIN enviado al correo o celular, "
                                              "y pueden adjuntarse archivos de soporte; no debe cargarse información que no sea necesaria."),
    ("6. Datos personales", "El tratamiento de datos personales se rige por la Ley 1581 de 2012 y por nuestra Política de "
                            "protección de datos personales, disponible en este sitio. Los titulares pueden ejercer sus derechos "
                            "de conocer, actualizar, rectificar y suprimir sus datos."),
    ("7. Propiedad intelectual", "El software, marcas, diseños y contenidos de la Plataforma pertenecen a {empresa} o a sus "
                                 "licenciantes. No se otorga ninguna licencia distinta a la de uso conforme a estos términos."),
    ("8. Disponibilidad y responsabilidad", "Se procura mantener la Plataforma disponible, pero pueden presentarse interrupciones "
                                            "por mantenimiento o causas ajenas. {empresa} no responde por daños indirectos derivados "
                                            "del uso o de la imposibilidad de uso, salvo lo que la ley no permita excluir."),
    ("9. Modificaciones", "Estos términos pueden actualizarse. La versión vigente es la publicada en esta página."),
    ("10. Ley aplicable", "Estos términos se rigen por las leyes de la República de Colombia."),
]

_SITIO = [
    ("Plataforma", [("Inicio", "/"), ("Ingresar", "/login")]),
    ("Atención al usuario (PQR)", [("Portal de PQR", "/pqr"), ("Radicar una PQR", "/pqr/crear"),
                                   ("Consultar el estado de mi PQR", "/pqr/consulta"), ("Política de PQR", "/politica-pqr"),
                                   ("Atención a directivos", "/atencion-directivos")]),
    ("Legal y privacidad", [("Términos y condiciones", URL_TERMINOS), ("Política de protección de datos personales", URL_DATOS),
                            ("Aviso de privacidad", "/politicas/aviso-privacidad"), ("Habeas Data", "/politicas/habeas"),
                            ("Política de cookies", "/politicas/cookies"), ("Seguridad de la información", "/politicas/seguridad-informacion"),
                            ("Política de reembolsos", "/politicas/reembolsos")]),
]


def _caja(titulo, cuerpo_html):
    return ("<div style='max-width:860px;margin:28px auto;padding:28px;background:#fff;border:1px solid #e2e8f0;"
            "border-radius:18px;font-family:Segoe UI,Arial,sans-serif;color:#1e293b;line-height:1.65'>"
            "<h1 style='margin:0 0 14px;color:#0020C2;font-size:26px'>%s</h1>%s</div>" % (_html.escape(titulo), cuerpo_html))


def instalar(g):
    if g.get("_PIE_PAGINA_INSTALADO"):
        return
    app = g["app"]
    from flask import request

    def nombres():
        try:
            prod = g["nombre_producto"]()
        except Exception:
            prod = g.get("APP_NAME") or "EduTrack"
        try:
            emp = g["nombre_empresa"]()
        except Exception:
            emp = g.get("DESARROLLADOR") or "Procsis"
        return prod, emp

    # 1) la función footer() de la app (usada en los paneles internos)
    def footer_nuevo():
        prod, emp = nombres()
        return '<div class="footer">%s</div>' % construir_pie(prod, emp)
    g["footer"] = footer_nuevo

    # 2) cualquier otra página que traiga el pie viejo escrito a mano (landing, login…)
    @app.after_request
    def _pie_legal(resp):
        try:
            if (resp.mimetype != "text/html" or resp.direct_passthrough
                    or resp.headers.get("Content-Encoding") or request.path.startswith("/static")):
                return resp
            txt = resp.get_data(as_text=True)
            if "Desarrollado por" not in txt:
                return resp
            prod, emp = nombres()
            nuevo = reemplazar_pies(txt, prod, emp)
            if nuevo != txt:
                resp.set_data(nuevo)
        except Exception as ex:
            print("PIE LEGAL:", ex)
        return resp

    # 3) páginas legales
    def envolver(titulo, cuerpo):
        shell = g.get("_public_shell")
        if shell:
            try:
                return shell(titulo, cuerpo, "nosotros")
            except Exception:
                pass
        return g["page"](titulo, cuerpo)

    @app.route(URL_TERMINOS)
    def pagina_terminos():
        prod, emp = nombres()
        e, p = _html.escape(emp), _html.escape(prod)
        partes = "".join("<h3 style='margin:18px 0 4px;font-size:16px'>%s</h3><p style='margin:0'>%s</p>"
                         % (_html.escape(t), c.format(producto=p, empresa=e)) for t, c in _TERMINOS)
        pie = "<p style='margin-top:22px;color:#64748b;font-size:13px'>Última actualización: %s</p>" % datetime.now().strftime("%d/%m/%Y")
        return envolver("Términos y condiciones", _caja("Términos y condiciones", partes + pie))

    @app.route(URL_MAPA)
    def pagina_mapa_sitio():
        bloques = ""
        for grupo, links in _SITIO:
            items = "".join("<li style='margin:6px 0'><a href='%s' style='color:#0020C2;font-weight:600;text-decoration:none'>%s</a></li>"
                            % (_html.escape(u), _html.escape(t)) for t, u in links)
            bloques += "<h3 style='margin:18px 0 4px;font-size:16px'>%s</h3><ul style='margin:0;padding-left:20px'>%s</ul>" % (_html.escape(grupo), items)
        return envolver("Mapa del sitio", _caja("Mapa del sitio", bloques))

    g["_PIE_PAGINA_INSTALADO"] = True
    print("PIE LEGAL instalado (footer, /terminos, /mapa-del-sitio)")

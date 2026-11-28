# -*- coding: utf-8 -*-
"""
pie_pagina.py  ·  [procsis-pie-legal-v1]

Cambia el pie de página "EduTrack · Desarrollado por Procsis" por un pie legal
con el formato de las operadoras (estilo Tigo):

  © EduTrack Colombia 2026. Servicios de software educativo prestados por PROCSIS.
  Todos los derechos reservados.
  Términos y condiciones | Política de protección de datos personales | Mapa del sitio

También crea las páginas /terminos, /mapa-del-sitio, /beneficios-cupones e /informacion-legal, y en la
landing y las páginas públicas:
  · saca de arriba la barra blanca duplicada (Soluciones · Beneficios y Cupones · Soporte y PQR · Pagar Factura)
    y deja arriba solo la línea de Colombia; esos accesos pasan al pie, cada uno hacia su propia página;
  · las notas legales largas ("La implementación inicial requiere un pago único…") pasan a /informacion-legal;
  · no toca los botones flotantes de accesibilidad ni de IA.

NO modifica app.py: se engancha con  pie_pagina.instalar(globals())
Los textos salen de Soporte → Marca (nombre del producto y de la empresa).
Para cambiar la redacción legal edita las constantes de abajo.
"""
import html as _html
import json
import re
from datetime import datetime

PAIS = "Colombia"
SERVICIO = "Servicios de software educativo prestados por"   # → "... prestados por PROCSIS."
URL_TERMINOS = "/terminos"
URL_DATOS = "/politicas/datos-personales"
URL_MAPA = "/mapa-del-sitio"
URL_BENEFICIOS = "/beneficios-cupones"
URL_INFO_LEGAL = "/informacion-legal"

# Accesos que antes iban arriba (barra blanca) y ahora van al pie. (texto, destino)
ACCESOS = [("Soluciones", "/soluciones"), ("Beneficios y Cupones", URL_BENEFICIOS),
           ("Soporte y PQR", "/pqr"), ("Pagar Factura", "/pagar")]

_R_FOOTER_DIV = re.compile(r'<div class="footer">(?:(?!</div>).)*?Desarrollado por(?:(?!</div>).)*?</div>', re.S)
_R_FOOTER_LP = re.compile(r'<footer class="lp-footer">.*?Desarrollado por.*?</footer>', re.S)

_ESTILO_LINK = "color:inherit;text-decoration:underline;font-weight:600"

# Pie azul corporativo (_footer_corporativo_html): su © viejo se reemplaza por la frase nueva (un solo ©)
_R_CF_COPY = re.compile(r'<span>© \d{4} PROCSIS · Soluciones digitales\. Todos los derechos reservados\.</span>')

# Landing / páginas públicas: accesos al pie, barra duplicada fuera de arriba y, en el login, segunda pantalla
_LANDING_CSS = """<style id="pcx-landing-css">
#pcx-ventana2{min-height:100vh;box-sizing:border-box;display:flex;flex-direction:column;justify-content:flex-end;background:#f5f5f7;padding-top:36px}
#pcx-ventana2 .login-apple-foot{width:100%;box-sizing:border-box}
#pcx-ventana2 .cf-foot{margin-top:24px}
#pcx-accesos{box-sizing:border-box;width:calc(100% - 40px);max-width:1180px;margin:22px auto;padding:12px 22px;display:flex;align-items:center;justify-content:space-between;gap:12px 24px;flex-wrap:wrap;background:#fff;border:1px solid #e2e8f0;border-radius:16px;box-shadow:0 2px 10px rgba(15,23,42,.05);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif}
#pcx-accesos a{text-decoration:none}
#pcx-accesos .pcx-brand{display:flex;align-items:center;gap:8px;color:#0020C2;font-size:15px}
#pcx-accesos .pcx-brand b{font-weight:800}
#pcx-accesos .pcx-brand i{font-style:normal;color:#94a3b8;font-weight:400}
#pcx-accesos .pcx-p{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:8px;background:#0020C2;color:#fff;font-weight:800}
#pcx-accesos .pcx-links{display:flex;gap:6px 24px;flex-wrap:wrap}
#pcx-accesos .pcx-links a{color:#0f172a;font-size:14px;font-weight:600;padding:6px 2px}
#pcx-accesos .pcx-links a:hover{color:#0020C2}
#pcx-accesos .pcx-btn{background:#0020C2;color:#fff;font-weight:700;font-size:14px;padding:10px 20px;border-radius:980px}
</style>"""

_LANDING_JS = r"""<script id="pcx-landing-js">
(function(){
  var ACC=__ACC__, WANT=ACC.map(function(a){return a[0].toLowerCase();}), T0=Date.now(), listo=false;
  var login=!!document.querySelector(".login-page");
  function norm(t){return (t||"").replace(/\s+/g," ").trim().toLowerCase();}
  function accesos(){
    var a=document.getElementById("pcx-accesos"); if(a) return a;
    a=document.createElement("nav"); a.id="pcx-accesos"; a.setAttribute("aria-label","Accesos rápidos");
    var l="",i; for(i=0;i<ACC.length;i++) l+='<a href="'+ACC[i][1]+'">'+ACC[i][0]+'</a>';
    a.innerHTML='<a class="pcx-brand" href="/procsis"><span class="pcx-p">P</span><b>PROCSIS</b><i>|</i><b>EduTrack</b></a><div class="pcx-links">'+l+'</div><a class="pcx-btn" href="/login">Ingresar colegio</a>';
    return a;
  }
  function armar(){
    var f=document.querySelector("footer.cf-foot"); if(!f) return;
    var a=accesos();
    if(login){
      var n=document.querySelector(".login-apple-foot"), w=document.getElementById("pcx-ventana2");
      if(!w){ w=document.createElement("section"); w.id="pcx-ventana2"; w.setAttribute("aria-label","Accesos, información legal y datos de la empresa"); }
      [a,n,f].forEach(function(x){ if(x && x.parentNode!==w) w.appendChild(x); });
      if(document.body.lastElementChild!==w) document.body.appendChild(w);
    } else if(!a.parentNode){ f.parentNode.insertBefore(a,f); }
  }
  function ocultar(){
    if(listo || Date.now()-T0>20000) return;
    var glass=document.getElementById("navAppleWrap"), mine=document.getElementById("pcx-accesos");
    var els=document.querySelectorAll("a,button,span,li,div,p"), por={}, i, e, k;
    for(i=0;i<els.length;i++){
      e=els[i]; if(e.children.length) continue;
      k=WANT.indexOf(norm(e.textContent)); if(k<0||por[k]) continue;
      if((glass&&glass.contains(e))||(mine&&mine.contains(e))||e.closest("footer,#pcx-ventana2,[data-pcx-oculta]")) continue;
      por[k]=e;
    }
    var ks=Object.keys(por); if(ks.length<3) return;
    var xs=ks.map(function(q){return por[q];}), n=xs[0];
    while(n && !xs.every(function(x){return n.contains(x);})) n=n.parentElement;
    if(!n) return;
    while(n.parentElement && n.parentElement!==document.body){
      var p=n.parentElement;
      if((glass&&p.contains(glass))||p.querySelector(".login-page,.sinai-portal,main,.lp-content,footer")||(p.innerText||"").length>600) break;
      n=p;
    }
    if(n===document.body||n===document.documentElement||(glass&&n.contains(glass))) return;
    var h=n.offsetHeight, pos=getComputedStyle(n).position;
    n.setAttribute("data-pcx-oculta","1"); n.style.setProperty("display","none","important");
    if(pos==="fixed"||pos==="sticky"){
      var bs=getComputedStyle(document.body);
      var pad=parseFloat(bs.paddingTop)||0;
      if(pad>0&&pad>=h-2&&pad<=h+24) document.body.style.setProperty("padding-top",Math.max(0,pad-h)+"px","important");
      if(glass){ var gt=parseFloat(getComputedStyle(glass).top); if(!isNaN(gt)&&gt>0&&Math.abs(gt-h)<=2) glass.style.setProperty("top","0","important"); }
    }
    listo=true;
  }
  function todo(){ armar(); ocultar(); }
  todo();
  document.addEventListener("DOMContentLoaded",todo);
  window.addEventListener("load",function(){ todo(); setTimeout(todo,800); setTimeout(todo,2500); });
  try{ var t=null; new MutationObserver(function(){ if(t) return; t=setTimeout(function(){ t=null; todo(); },120); }).observe(document.body,{childList:true,subtree:true}); }catch(e){}
})();
</script>"""


def landing_html():
    return _LANDING_CSS + _LANDING_JS.replace("__ACC__", json.dumps(ACCESOS, ensure_ascii=False))


def frase_copy(producto, empresa, anio=None):
    anio = anio or datetime.now().year
    return "© %s %s %s. %s %s. Todos los derechos reservados." % (
        _html.escape(producto), _html.escape(PAIS), anio, _html.escape(SERVICIO), _html.escape(empresa))


def enlaces_legales():
    return (
        '<a href="%s" style="%s">Términos y condiciones</a> | '
        '<a href="%s" style="%s">Política de protección de datos personales</a> | '
        '<a href="%s" style="%s">Mapa del sitio</a>'
        % (URL_TERMINOS, _ESTILO_LINK, URL_DATOS, _ESTILO_LINK, URL_MAPA, _ESTILO_LINK)
    )


def construir_pie(producto, empresa, anio=None):
    """HTML interno del pie completo (frase © + enlaces), sin contenedor."""
    return ('<span class="pie-legal" style="display:block;line-height:1.7;font-size:12.5px">%s<br>%s</span>'
            % (frase_copy(producto, empresa, anio), enlaces_legales()))


def reemplazar_pies(html, producto, empresa):
    """Unifica el pie legal en un HTML. Devuelve el HTML nuevo (o el mismo si no hay pie que tocar)."""
    tiene_cf = 'class="cf-foot"' in html
    if "Desarrollado por" not in html and not tiene_cf:
        return html
    interno = construir_pie(producto, empresa)
    if tiene_cf:
        # Un solo ©: la frase nueva reemplaza al © del pie azul y los enlaces van justo debajo, dentro del mismo pie.
        html = _R_CF_COPY.sub(lambda m: "<span>%s</span>" % frase_copy(producto, empresa), html)
        if 'class="cf-legal"' not in html:
            m = re.search(r'<footer class="cf-foot"[^>]*>', html)
            j = html.find("</footer>", m.end()) if m else -1
            if j >= 0:
                fila = ('<div class="cf-legal" style="max-width:1180px;margin:14px auto 0;text-align:center;'
                        'font-size:11.5px;font-weight:400;color:rgba(255,255,255,.84)">%s</div>' % enlaces_legales())
                html = html[:j] + fila + html[j:]
        # la banda oscura "Desarrollado por…" de la landing sobraba: duplicaba el ©
        html = _R_FOOTER_LP.sub("", html)
        html = _R_FOOTER_DIV.sub("", html)
        # landing y páginas públicas con barra superior: accesos al pie (y en el login, segunda pantalla)
        if "pcx-landing-js" not in html and ('class="login-page"' in html or "navAppleWrap" in html):
            k = html.rfind("</body>")
            html = (html[:k] + landing_html() + html[k:]) if k >= 0 else (html + landing_html())
        return html
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
    ("Plataforma", [("Inicio", "/"), ("Ingresar", "/login"), ("Soluciones", "/soluciones"),
                    ("Beneficios y cupones", URL_BENEFICIOS), ("Pagar factura", "/pagar")]),
    ("Atención al usuario (PQR)", [("Portal de PQR", "/pqr"), ("Radicar una PQR", "/pqr/crear"),
                                   ("Consultar el estado de mi PQR", "/pqr/consulta"), ("Política de PQR", "/politica-pqr"),
                                   ("Atención a directivos", "/atencion-directivos")]),
    ("Legal y privacidad", [("Términos y condiciones", URL_TERMINOS), ("Política de protección de datos personales", URL_DATOS),
                            ("Aviso de privacidad", "/politicas/aviso-privacidad"), ("Habeas Data", "/politicas/habeas"),
                            ("Política de cookies", "/politicas/cookies"), ("Seguridad de la información", "/politicas/seguridad-informacion"),
                            ("Política de reembolsos", "/politicas/reembolsos"),
                            ("Información legal y condiciones del servicio", URL_INFO_LEGAL)]),
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
            if "Desarrollado por" not in txt and 'class="cf-foot"' not in txt:
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

    # 4) notas legales largas de la landing -> una sola línea con enlace a /informacion-legal
    orig_pie_login = g.get("_html_login_pie_colegios")
    if orig_pie_login:
        compacto = ('<div class="login-apple-foot"><p class="laf-notes">Alcance de la implementación, condiciones del servicio '
                    'y avisos legales: <a href="%s">consulta la información completa</a>.</p></div>' % URL_INFO_LEGAL)

        def pie_login_compacto():
            h = orig_pie_login()
            nuevo_h, n = re.subn(r'<div class="login-apple-foot">.*?</div>', lambda m: compacto, h, count=1, flags=re.S)
            return nuevo_h if n else h
        g["_html_login_pie_colegios"] = pie_login_compacto

    def aura():
        """Botón de IA (Edu Aura) para que también esté en las páginas nuevas."""
        try:
            return g["_html_aura_publica"]()
        except Exception:
            return ""

    def enlazar_dominios(txt):
        t = _html.escape(txt or "")
        t = t.replace("support.procsis.com", '<a href="/support" style="color:#0066cc">support.procsis.com</a>')
        return re.sub(r'(?<![\w./>])procsis\.com', '<a href="/procsis" style="color:#0066cc">procsis.com</a>', t)

    @app.route(URL_INFO_LEGAL)
    def pagina_info_legal():
        p1 = p2 = extra = ""
        try:
            p1, p2, extra = g["_login_pie_defaults"]()
        except Exception:
            pass
        try:
            extra = (getattr(g["plataforma"](), "login_pie_extra", None) or "").strip() or extra
        except Exception:
            pass
        partes = "".join("<p style='margin:0 0 14px'>%s</p>" % enlazar_dominios(x) for x in (p1, p2) if x)
        if extra:
            partes += "<p style='margin:0 0 14px'>%s</p>" % _html.escape(extra)
        partes += ("<p style='margin:18px 0 0;font-size:14px'><a href='%s' style='color:#0020C2;font-weight:600'>Términos y condiciones</a> · "
                   "<a href='%s' style='color:#0020C2;font-weight:600'>Política de protección de datos personales</a></p>" % (URL_TERMINOS, URL_DATOS))
        return envolver("Información legal", _caja("Información legal y condiciones del servicio", partes) + aura())

    # 5) página pública de Beneficios y Cupones (planes + promoción vigente)
    def _cop(v):
        return "$" + "{:,.0f}".format(float(v or 0)).replace(",", ".")

    @app.route(URL_BENEFICIOS)
    def pagina_beneficios_cupones():
        PC, PA = g.get("PlanComercial"), g.get("PromocionActiva")
        planes, promo = [], None
        try:
            planes = PC.query.filter_by(activo=True).order_by(PC.orden, PC.precio_mensual).all()
        except Exception:
            planes = []
        try:
            promo = PA.query.filter_by(estado="ACTIVO").order_by(PA.id.desc()).first()
        except Exception:
            promo = None
        pct = float(getattr(promo, "porcentaje_descuento", 0) or 0) if promo else 0.0
        unidad = {"MESES": "meses", "DIAS": "días", "ANOS": "años"}.get(str(getattr(promo, "duracion_unidad", "") or "").upper(), "meses")
        banner = ""
        if promo and pct > 0:
            banner = ("<div style='background:#eef2ff;border:1px solid #c7d2fe;border-radius:14px;padding:16px 18px;margin:0 0 22px'>"
                      "<b style='color:#0020C2'>Promoción vigente: %s</b><div style='margin-top:4px;color:#334155'>"
                      "%g%% de descuento durante %s %s. Consulta con Ventas cómo aplicarla a tu institución.</div></div>"
                      % (_html.escape(getattr(promo, "nombre_promo", "") or "Promoción"), pct, getattr(promo, "duracion_valor", 1) or 1, unidad))
        tarjetas = ""
        for pl in planes:
            pleno = float(getattr(pl, "precio_lista", 0) or 0) or float(getattr(pl, "precio_mensual", 0) or 0)
            if pct > 0 and pleno:
                precio = ("<div style='font-size:24px;font-weight:800;color:#0f172a'>%s <span style='font-size:13px;font-weight:500'>COP/mes</span></div>"
                          "<div style='font-size:12.5px;color:#94a3b8;text-decoration:line-through'>%s COP</div>" % (_cop(pleno * (1 - pct / 100.0)), _cop(pleno)))
            else:
                precio = "<div style='font-size:24px;font-weight:800;color:#0f172a'>%s <span style='font-size:13px;font-weight:500'>COP/mes</span></div>" % _cop(pleno)
            try:
                feats = [str(x) for x in json.loads(getattr(pl, "features_json", "[]") or "[]")][:8]
            except Exception:
                feats = []
            lis = "".join("<li style='margin:5px 0'>%s</li>" % _html.escape(f) for f in feats)
            extra = ""
            if getattr(pl, "max_estudiantes", 0):
                extra = "<div style='font-size:12.5px;color:#64748b;margin-top:6px'>Hasta %s estudiantes</div>" % int(pl.max_estudiantes)
            tarjetas += ("<div style='background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:20px'>"
                         "<div style='font-weight:800;color:#0020C2;font-size:17px;margin-bottom:8px'>%s</div>%s%s"
                         "<ul style='margin:12px 0 0;padding-left:18px;color:#334155;font-size:14px'>%s</ul></div>"
                         % (_html.escape(getattr(pl, "nombre", "") or "Plan"), precio, extra, lis))
        if not tarjetas:
            tarjetas = ("<p style='color:#475569'>Estamos actualizando nuestros planes. Escríbenos y un asesor te cuenta los beneficios "
                        "disponibles para tu institución.</p>")
        cuerpo = ("<div style='max-width:1080px;margin:28px auto;padding:0 20px;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,system-ui,sans-serif;color:#1e293b'>"
                  "<h1 style='margin:0 0 6px;color:#0020C2;font-size:30px'>Beneficios y cupones</h1>"
                  "<p style='margin:0 0 22px;color:#475569'>Planes, descuentos y promociones disponibles para instituciones educativas.</p>"
                  "%s<div style='display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:16px'>%s</div>"
                  "<div style='margin:26px 0 0;display:flex;gap:10px;flex-wrap:wrap'>"
                  "<a href='/contacto' style='background:#0020C2;color:#fff;text-decoration:none;font-weight:700;padding:11px 20px;border-radius:980px'>Hablar con Ventas</a>"
                  "<a href='/pagar' style='background:#fff;color:#0020C2;text-decoration:none;font-weight:700;padding:11px 20px;border-radius:980px;border:1px solid #c7d2fe'>Pagar factura</a>"
                  "</div></div>" % (banner, tarjetas))
        return envolver("Beneficios y cupones", cuerpo + aura())

    g["_PIE_PAGINA_INSTALADO"] = True
    print("PIE LEGAL instalado (footer, accesos al pie, /terminos, /mapa-del-sitio, /beneficios-cupones, /informacion-legal)")

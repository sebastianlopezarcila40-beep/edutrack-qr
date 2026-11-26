# -*- coding: utf-8 -*-
"""
tema_2027.py · Diseño 2027 del portal de acceso de PROCSIS | EduTrack
=====================================================================
Se instala con UN bloque en app.py (lo agrega aplicar_tema.py):

    import tema_2027; tema_2027.instalar(globals())

Qué hace (solo en GET de "/" y "/login"; NO toca el POST ni la lógica de ingreso):
  1. Pinta el portal con los colores del logo: azul marino #002069, azul vivo #0172EC, cian #00F3FE y verde #03D047.
  2. Agrega el título «Tu colegio, más conectado cada día» con chips de beneficios (columna izquierda, sobre el carrusel,
     para que el Portal de acceso quede arriba, al lado).
  3. Agrega 4 accesos rápidos (solo los que existen en tu app): Convocatoria de 48 Horas, Consulta de asistencia,
     Autorizar salidas y Pagar factura.
  4. Spinner pequeño dentro del botón «Iniciar sesión» (y bloqueo de doble envío).
  5. Agrega el saludo con el ícono de la P dentro de la tarjeta de ingreso. Usa el primer archivo que exista en
     static/img/ entre  logo-procsis-p.png  y  favicon.png  (o static/favicon.png). Para fijar otro:  TEMA_2027_LOGO=/static/img/archivo.png
Si algo falla, la página sale exactamente como estaba antes.

INTERRUPTOR:  TEMA_2027=off   (variable de entorno) -> vuelve al diseño anterior sin tocar código.
"""
import os
import re

MARCA = "t27-aplicado"

CSS = """<style id="t27-css">
.sinai-portal{padding:14px 20px 8px!important}
.sinai-portal .sinai-grid{max-width:1120px!important;gap:10px 40px!important;align-items:start!important;grid-template-columns:1.25fr .85fr!important;grid-template-areas:"hero access" "car access"!important;grid-template-rows:auto 1fr}
.sinai-portal .sinai-carousel{grid-area:car;height:360px!important;border-radius:22px!important;box-shadow:0 18px 44px rgba(0,32,105,.18)!important}
.sinai-portal .sinai-slide-bgblur{filter:blur(26px) saturate(1.3) brightness(1)!important;opacity:.5!important}
.sinai-portal .sinai-access{grid-area:access;align-self:start;padding-top:6px}
.sinai-portal .sinai-access h1{font-size:26px!important;letter-spacing:-.02em;margin:0 0 4px!important;color:#002069!important}
.sinai-portal .sinai-access h1:after{content:"";display:block;width:44px;height:4px;border-radius:4px;margin-top:10px;background:linear-gradient(90deg,#0172EC,#03D047)}
.sinai-portal .sinai-sub{color:#475569!important;font-size:13.5px!important;margin-bottom:14px!important}
.sinai-portal .sinai-card{border-radius:18px!important;background:#fff!important;border:1px solid #d6e2f5!important;box-shadow:0 12px 34px rgba(0,32,105,.12)!important}
.sinai-portal .sinai-card-label{color:#0172EC!important;letter-spacing:.08em}
.sinai-portal .sinai-form label{color:#002069!important;font-weight:700}
.sinai-portal .sinai-form input,.sinai-portal .sinai-form select{border-color:#c5d5ee!important;border-radius:10px!important}
.sinai-portal .sinai-form input:focus,.sinai-portal .sinai-form select:focus{border-color:#0172EC!important;box-shadow:0 0 0 3px rgba(3,208,71,.22)!important;outline:none}
.sinai-portal .sinai-btn{display:flex!important;align-items:center;justify-content:center;gap:10px;border-radius:10px!important;background:linear-gradient(110deg,#002069,#0172EC)!important;box-shadow:0 8px 20px rgba(1,114,236,.30)!important}
.sinai-portal .sinai-btn:hover{filter:brightness(1.1)}
.sinai-portal .sinai-btn.t27-load{opacity:.92;cursor:progress}
.sinai-portal .sinai-dot.is-on{background:#03D047!important}
.promo-card{border-radius:22px!important;border:1px solid #dbeafe!important;box-shadow:0 10px 30px rgba(1,114,236,.10)!important}
.t27-spin{display:inline-block;width:18px;height:18px;border-radius:50%;border:2.5px solid rgba(255,255,255,.35);border-top-color:#fff;animation:t27sp .8s linear infinite}
@keyframes t27sp{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion:reduce){.t27-spin{animation-duration:2.4s}}
.t27-hero{grid-area:hero;padding:2px 0 6px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif}
.t27-pill{display:inline-flex;align-items:center;gap:8px;padding:6px 13px;border-radius:999px;background:#fff;border:1px solid #bfdbfe;font-size:12.5px;font-weight:700;color:#002069}
.t27-pill i{width:8px;height:8px;border-radius:50%;background:#03D047}
.t27-hero h2{margin:14px 0 10px;font-size:clamp(28px,3.4vw,42px);line-height:1.08;letter-spacing:-.03em;font-weight:800;color:#002069}
.t27-gt{background:linear-gradient(100deg,#0172EC,#008F6B);-webkit-background-clip:text;background-clip:text;color:transparent}
.t27-hero p{margin:0 0 14px;font-size:16px;line-height:1.55;color:#334155;max-width:560px}
.t27-chips{display:flex;gap:8px;flex-wrap:wrap}
.t27-chips span{display:inline-flex;align-items:center;gap:7px;padding:7px 13px;border-radius:999px;background:rgba(255,255,255,.85);border:1px solid #dbeafe;font-size:13px;font-weight:600;color:#002069}
.t27-greet{display:flex;align-items:center;gap:12px;margin:0 0 12px}
.t27-greet img{width:44px;height:auto}
.t27-greet b{display:block;font-size:15px;color:#002069}
.t27-greet span{font-size:13px;color:#475569}
.t27-tiles{max-width:1120px;margin:14px auto 28px;padding:0 20px;display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif}
.t27-tile{display:flex;flex-direction:column;gap:10px;padding:22px;border-radius:20px;background:#fff;border:1px solid #dbeafe;box-shadow:0 10px 30px rgba(1,114,236,.12);text-decoration:none;color:inherit;transition:transform .2s,box-shadow .2s}
.t27-tile:hover{transform:translateY(-3px);box-shadow:0 18px 40px rgba(1,114,236,.22)}
.t27-ico{display:grid;place-items:center;width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,#0172EC,#00C2D6)}
.t27-tile b{font-size:16px;color:#002069}.t27-tile span.d{font-size:14px;color:#475569}
.t27-tile span.g{margin-top:auto;color:#0172EC;font-weight:800;font-size:14px}
@media(max-width:900px){.sinai-portal .sinai-grid{grid-template-columns:1fr!important;grid-template-areas:"hero" "access" "car"!important}.sinai-portal .sinai-carousel{height:300px!important}}
</style>"""

JS = """<script id="t27-js">(function(){
function reset(f,b){f.removeAttribute('data-t27');b.classList.remove('t27-load');b.removeAttribute('aria-busy');if(b.getAttribute('data-t27h')!==null){b.innerHTML=b.getAttribute('data-t27h');}}
function armar(f){var b=f.querySelector('button.sinai-btn');if(!b)return;
f.addEventListener('submit',function(ev){if(f.getAttribute('data-t27')==='1'){ev.preventDefault();return;}
f.setAttribute('data-t27','1');b.setAttribute('data-t27h',b.innerHTML);b.setAttribute('aria-busy','true');b.classList.add('t27-load');
b.innerHTML='<i class="t27-spin" aria-hidden="true"></i><span role="status">Verificando…</span>';setTimeout(function(){reset(f,b);},20000);});
window.addEventListener('pageshow',function(e){if(e.persisted){reset(f,b);}});}
function init(){var fs=document.querySelectorAll('form.sinai-form');for(var i=0;i<fs.length;i++){armar(fs[i]);}}
if(document.readyState==='loading'){document.addEventListener('DOMContentLoaded',init);}else{init();}
})();</script>"""

_CHK = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#03A83A" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>'
_S = 'viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"'

# (ruta que debe existir, ícono, título, descripción, llamado a la acción)
TILES = (
    ("/convocatoria48", '<circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2M9 3h6"/>', "Convocatoria de 48 Horas", "Inicia aquí tu participación.", "Entrar"),
    ("/acudiente-asistencia", '<rect x="4" y="5" width="16" height="15" rx="3"/><path d="M8 3v4M16 3v4M4 10h16M9 15l2 2 4-4"/>', "Consulta de asistencia", "Revisa la asistencia de tu estudiante.", "Consultar"),
    ("/acudiente/autorizar-salida", '<path d="M12 3l7 3v5c0 5-3 8.5-7 10-4-1.5-7-5-7-10V6l7-3z"/><path d="M9 12l2 2 4-4"/>', "Autorizar salidas", "Aprueba salidas con validación segura.", "Autorizar"),
    ("/pagar", '<rect x="3" y="6" width="18" height="13" rx="3"/><path d="M3 10h18M7 15h3"/>', "Pagar factura", "Paga tus servicios PROCSIS en línea.", "Pagar"),
)


def _hero():
    chips = "".join("<span>%s%s</span>" % (_CHK, t) for t in ("Notas y boletines", "Asistencia al día", "Salidas seguras", "Pagos en línea"))
    return ('<section class="t27-hero"><span class="t27-pill"><i></i>Nueva identidad PROCSIS · EduTrack 2027</span>'
            '<h2>Tu colegio, <span class="t27-gt">más conectado</span> cada día.</h2>'
            '<p>Notas, asistencia, salidas y pagos en un solo lugar para rectoría, coordinación, secretaría, docentes y familias.</p>'
            '<div class="t27-chips">%s</div></section>' % chips)


def _tiles(rutas):
    t = "".join('<a class="t27-tile" href="%s"><span class="t27-ico"><svg %s>%s</svg></span><b>%s</b><span class="d">%s</span><span class="g">%s →</span></a>'
                % (r, _S, ico, ti, d, c) for (r, ico, ti, d, c) in TILES if r in rutas)
    return ('<section class="t27-tiles" aria-label="Accesos rápidos">%s</section>' % t) if t else ""


def _saludo(src):
    return ('<div class="t27-greet"><img src="%s" alt=""><div><b>Acceso institucional</b>'
            '<span>Rectoría, coordinación, secretaría y docentes</span></div></div>' % src)


def buscar_logo(raiz):
    """Devuelve la URL del logo (la P) o '' si no hay ninguno."""
    fijo = (os.environ.get("TEMA_2027_LOGO") or "").strip()
    if fijo.startswith("/static/") and ".." not in fijo and os.path.isfile(os.path.join(raiz, fijo.lstrip("/").replace("/", os.sep))):
        return fijo
    for rel in ("static/img/logo-procsis-p.png", "static/img/favicon.png", "static/favicon.png"):
        if os.path.isfile(os.path.join(raiz, rel.replace("/", os.sep))):
            return "/" + rel
    return ""


def transformar(h, rutas=(), logo=""):
    """Devuelve el HTML con el diseño nuevo, o None si la página no tiene la estructura esperada."""
    if MARCA in h or 'class="sinai-portal"' not in h:
        return None
    a_css = h.find('<div class="login-page"')
    a_sec = h.find('<section class="sinai-portal">')
    if a_css < 0 or a_sec < 0 or a_sec < a_css:
        return None
    tag = '<div class="sinai-grid">'
    a_grid = h.find(tag, a_sec)
    if 0 <= a_grid < a_sec + 600:
        corte = a_grid + len(tag)
        out = h[:a_css] + "<!--%s-->" % MARCA + CSS + h[a_css:corte] + _hero() + h[corte:]
    else:  # respaldo: si cambia el marcado, el título va arriba (sin rejilla)
        out = h[:a_css] + "<!--%s-->" % MARCA + CSS + h[a_css:a_sec] + _hero() + h[a_sec:]
    if logo:
        i = out.find('<div class="sinai-card">')
        if i >= 0:
            out = out[:i] + _saludo(logo if isinstance(logo, str) else "/static/img/logo-procsis-p.png") + out[i:]
    t = _tiles(rutas)
    j = out.find('<div class="promo-row"')
    if t and j >= 0:
        out = out[:j] + t + out[j:]
    k = out.lower().rfind("</body>")
    out = (out[:k] + JS + out[k:]) if k >= 0 else (out + JS)
    return out


def instalar(G):
    if (os.environ.get("TEMA_2027") or "on").strip().lower() == "off":
        print("TEMA 2027: apagado (TEMA_2027=off).", flush=True)
        return
    try:
        from flask import request
        app = G["app"]

        @app.after_request
        def _tema_2027(resp):
            try:
                if request.method != "GET" or request.path not in ("/", "/login"):
                    return resp
                if resp.status_code != 200 or resp.mimetype != "text/html" or resp.direct_passthrough:
                    return resp
                logo = buscar_logo(app.root_path)
                nuevo = transformar(resp.get_data(as_text=True), {r.rule for r in app.url_map.iter_rules()}, logo)
                if nuevo:
                    resp.set_data(nuevo)
            except Exception as ex:
                print("TEMA 2027 (se muestra el diseño anterior):", repr(ex)[:160], flush=True)
            return resp

        print("TEMA 2027: diseño nuevo instalado en / y /login.", flush=True)
    except Exception as ex:
        print("TEMA 2027: NO se instaló (la app sigue igual):", repr(ex)[:160], flush=True)

# -*- coding: utf-8 -*-
"""
cargadores_20.py · Biblioteca de 20 animaciones de carga · PROCSIS | EduTrack
=============================================================================
Se instala con UN bloque en app.py (lo agrega aplicar_respaldo_y_cargadores.py):

    import cargadores_20; cargadores_20.instalar(globals())

Qué hace:
  /cargadores              -> galería con las 20 animaciones. Cada persona elige la suya ("Usar este") y se
                              guarda en su navegador.
  /cargadores/cargadores.css y .js -> estilos y lógica (se cachean).
  Al navegar (clic en enlaces internos o envío de formularios) aparece una cortina con la animación elegida,
  solo si la página tarda más de 300 ms. Se oculta sola al cargar la página nueva (o a los 15 s).

INTERRUPTORES (variables de entorno):
  CARGADOR_MODO  = on (defecto) | off   -> apaga todo (no inyecta nada; la galería sigue disponible)
  CARGADOR       = spinner_arc (defecto) -> animación por defecto del sistema (clave de la lista de la galería)
  CARGADOR_INYECTAR = on (defecto) | off -> solo apaga la cortina automática

No toca la base de datos, el login ni ninguna ruta existente.
"""
import os

from flask import request, make_response, session

MARCA = "cg-aplicado"

NOMBRES = (
    ("spinner_arc", "Spinner Arc", "Arco giratorio clásico"),
    ("bouncing_dots", "Bouncing Dots", "Puntos rebotando"),
    ("pulse_rings", "Pulse Rings", "Anillos con efecto pulso"),
    ("indeterminate_bar", "Indeterminate Bar", "Barra de progreso indeterminada"),
    ("skeleton_shimmer", "Skeleton Shimmer", "Esqueleto con brillo"),
    ("orbit", "Orbit", "Elementos orbitando"),
    ("equalizer", "Equalizer Bars", "Barras tipo ecualizador"),
    ("square_morph", "Square Morph", "Cuadrado que cambia de forma"),
    ("liquid_fill", "Liquid Fill", "Relleno líquido que sube"),
    ("dot_grid_wave", "Dot Grid Wave", "Onda de puntos en grilla"),
    ("infinity_trail", "Infinity Trail", "Trazo infinito en movimiento"),
    ("flipping_square", "Flipping Square", "Cuadrado volteándose"),
    ("gradient_ring", "Gradient Ring", "Anillo con gradiente"),
    ("typing_dots", "Typing Dots", "Tres puntos «escribiendo…»"),
    ("spoke_spinner", "Spoke Spinner", "Spinner con rayos tipo rueda"),
    ("percent_bar", "Percent Bar", "Barra con porcentaje 33% / 71%"),
    ("chasing_squares", "Chasing Squares", "Cuadrados persiguiéndose"),
    ("text_shimmer", "Text Shimmer", "Texto con brillo"),
    ("battery_charge", "Battery Charge", "Batería cargándose"),
    ("bouncing_ball", "Bouncing Ball", "Pelota rebotando"),
)


def _spokes():
    s = ""
    for i in range(12):
        s += '<line x1="20" y1="4" x2="20" y2="11" transform="rotate(%d 20 20)" style="animation-delay:%.2fs"/>' % (i * 30, -1.2 + i * 0.1)
    return '<svg viewBox="0 0 40 40" class="sp">%s</svg>' % s


MARKUP = {
    "spinner_arc": '<i class="a1"></i>',
    "bouncing_dots": '<i></i><i></i><i></i>',
    "pulse_rings": '<i></i><i></i><b></b>',
    "indeterminate_bar": '<i></i>',
    "skeleton_shimmer": '<i></i><i></i><i></i>',
    "orbit": '<b></b><s><i></i></s>',
    "equalizer": '<i></i><i></i><i></i><i></i><i></i>',
    "square_morph": '<i></i>',
    "liquid_fill": '<i></i>',
    "dot_grid_wave": '<i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i>',
    "infinity_trail": '<svg viewBox="0 0 60 30"><path class="bg" d="M30 15C22 3 3 3 3 15s19 12 27 0s27-12 27 0s-19 12-27 0z"/>'
                      '<path class="fg" pathLength="100" d="M30 15C22 3 3 3 3 15s19 12 27 0s27-12 27 0s-19 12-27 0z"/></svg>',
    "flipping_square": '<i></i>',
    "gradient_ring": '<i></i>',
    "typing_dots": '<i></i><i></i><i></i>',
    "spoke_spinner": _spokes(),
    "percent_bar": '<i></i>',
    "chasing_squares": '<i></i><i></i>',
    "text_shimmer": '<span>Cargando…</span>',
    "battery_charge": '<i></i><i></i><i></i><i></i><u></u>',
    "bouncing_ball": '<i></i><b></b>',
}

CSS = """
.cg{--c1:#0172EC;--c2:#002069;--c3:#03D047;position:relative;display:inline-flex;align-items:center;justify-content:center;width:64px;height:64px;box-sizing:border-box}
.cg *{box-sizing:border-box}
@media (prefers-reduced-motion:reduce){.cg *{animation-duration:3s!important}}
/* 1 spinner arc */
.cg-spinner_arc .a1{width:40px;height:40px;border-radius:50%;border:4px solid rgba(1,114,236,.18);border-top-color:var(--c1);animation:cgspin .9s linear infinite}
@keyframes cgspin{to{transform:rotate(360deg)}}
/* 2 bouncing dots */
.cg-bouncing_dots{gap:7px}.cg-bouncing_dots i{width:11px;height:11px;border-radius:50%;background:var(--c1);animation:cgbounce .8s ease-in-out infinite}
.cg-bouncing_dots i:nth-child(2){animation-delay:.12s;background:var(--c2)}.cg-bouncing_dots i:nth-child(3){animation-delay:.24s;background:var(--c3)}
@keyframes cgbounce{0%,80%,100%{transform:translateY(0)}40%{transform:translateY(-16px)}}
/* 3 pulse rings */
.cg-pulse_rings i{position:absolute;width:44px;height:44px;border-radius:50%;border:3px solid var(--c1);opacity:0;animation:cgring 1.6s ease-out infinite}
.cg-pulse_rings i:nth-child(2){animation-delay:.8s}.cg-pulse_rings b{width:12px;height:12px;border-radius:50%;background:var(--c1)}
@keyframes cgring{0%{transform:scale(.25);opacity:.9}100%{transform:scale(1.25);opacity:0}}
/* 4 indeterminate bar */
.cg-indeterminate_bar{width:150px;height:6px;border-radius:6px;background:rgba(1,114,236,.16);overflow:hidden;justify-content:flex-start}
.cg-indeterminate_bar i{display:block;width:40%;height:100%;border-radius:6px;background:linear-gradient(90deg,var(--c2),var(--c1));animation:cgind 1.3s ease-in-out infinite}
@keyframes cgind{0%{transform:translateX(-110%)}100%{transform:translateX(280%)}}
/* 5 skeleton shimmer */
.cg-skeleton_shimmer{width:150px;height:auto;flex-direction:column;align-items:stretch;gap:8px}
.cg-skeleton_shimmer i{display:block;height:12px;border-radius:6px;background:linear-gradient(90deg,#e2e8f0 25%,#f8fafc 50%,#e2e8f0 75%);background-size:200% 100%;animation:cgshim 1.3s linear infinite}
.cg-skeleton_shimmer i:nth-child(2){width:80%}.cg-skeleton_shimmer i:nth-child(3){width:55%}
@keyframes cgshim{0%{background-position:200% 0}100%{background-position:-200% 0}}
/* 6 orbit */
.cg-orbit b{width:14px;height:14px;border-radius:50%;background:var(--c2)}
.cg-orbit s{position:absolute;inset:4px;border-radius:50%;border:1px dashed rgba(1,114,236,.3);animation:cgspin 1.6s linear infinite;text-decoration:none}
.cg-orbit s i{position:absolute;top:-6px;left:50%;margin-left:-6px;width:12px;height:12px;border-radius:50%;background:var(--c1);box-shadow:0 0 0 3px rgba(1,114,236,.2)}
/* 7 equalizer */
.cg-equalizer{gap:5px;align-items:flex-end;height:44px}.cg-equalizer i{width:7px;height:100%;border-radius:3px;background:var(--c1);transform-origin:bottom;animation:cgeq 1s ease-in-out infinite}
.cg-equalizer i:nth-child(2){animation-delay:-.2s;background:var(--c2)}.cg-equalizer i:nth-child(3){animation-delay:-.4s;background:var(--c3)}.cg-equalizer i:nth-child(4){animation-delay:-.6s;background:var(--c2)}.cg-equalizer i:nth-child(5){animation-delay:-.8s}
@keyframes cgeq{0%,100%{transform:scaleY(.2)}50%{transform:scaleY(1)}}
/* 8 square morph */
.cg-square_morph i{width:34px;height:34px;background:linear-gradient(135deg,var(--c1),var(--c3));animation:cgmorph 1.8s ease-in-out infinite}
@keyframes cgmorph{0%{border-radius:4px;transform:rotate(0)}50%{border-radius:50%;transform:rotate(180deg) scale(.8)}100%{border-radius:4px;transform:rotate(360deg)}}
/* 9 liquid fill */
.cg-liquid_fill{border:3px solid var(--c1);border-radius:50%;overflow:hidden;width:50px;height:50px}
.cg-liquid_fill i{position:absolute;left:-50%;width:200%;height:200%;top:100%;border-radius:38%;background:var(--c1);opacity:.85;animation:cgliq 2.6s ease-in-out infinite alternate,cgwave 3s linear infinite}
@keyframes cgliq{to{top:-10%}}@keyframes cgwave{to{transform:rotate(360deg)}}
/* 10 dot grid wave */
.cg-dot_grid_wave{display:grid;grid-template-columns:repeat(3,10px);gap:6px;width:auto;height:auto}
.cg-dot_grid_wave i{width:10px;height:10px;border-radius:50%;background:var(--c1);animation:cgwv 1.2s ease-in-out infinite}
.cg-dot_grid_wave i:nth-child(2),.cg-dot_grid_wave i:nth-child(4){animation-delay:.1s}.cg-dot_grid_wave i:nth-child(3),.cg-dot_grid_wave i:nth-child(5),.cg-dot_grid_wave i:nth-child(7){animation-delay:.2s}
.cg-dot_grid_wave i:nth-child(6),.cg-dot_grid_wave i:nth-child(8){animation-delay:.3s}.cg-dot_grid_wave i:nth-child(9){animation-delay:.4s}
@keyframes cgwv{0%,100%{transform:scale(.5);opacity:.4}50%{transform:scale(1.1);opacity:1;background:var(--c3)}}
/* 11 infinity trail */
.cg-infinity_trail{width:84px}.cg-infinity_trail svg{width:100%;height:auto;overflow:visible}
.cg-infinity_trail path{fill:none;stroke-width:4;stroke-linecap:round}.cg-infinity_trail .bg{stroke:rgba(1,114,236,.16)}
.cg-infinity_trail .fg{stroke:var(--c1);stroke-dasharray:22 78;animation:cginf 1.6s linear infinite}
@keyframes cginf{to{stroke-dashoffset:-100}}
/* 12 flipping square */
.cg-flipping_square{perspective:200px}.cg-flipping_square i{width:34px;height:34px;background:var(--c1);animation:cgflip 1.4s ease-in-out infinite}
@keyframes cgflip{0%{transform:rotateX(0) rotateY(0)}50%{transform:rotateX(-180deg) rotateY(0);background:var(--c2)}100%{transform:rotateX(-180deg) rotateY(-180deg)}}
/* 13 gradient ring */
.cg-gradient_ring i{width:44px;height:44px;border-radius:50%;background:conic-gradient(from 0deg,rgba(1,114,236,0),var(--c1) 60%,var(--c3));-webkit-mask:radial-gradient(farthest-side,transparent calc(100% - 6px),#000 calc(100% - 5px));mask:radial-gradient(farthest-side,transparent calc(100% - 6px),#000 calc(100% - 5px));animation:cgspin 1s linear infinite}
/* 14 typing dots */
.cg-typing_dots{width:auto;height:auto;padding:12px 16px;border-radius:18px;background:#e8f0fe;gap:5px}
.cg-typing_dots i{width:8px;height:8px;border-radius:50%;background:var(--c2);opacity:.35;animation:cgtyp 1.1s ease-in-out infinite}
.cg-typing_dots i:nth-child(2){animation-delay:.18s}.cg-typing_dots i:nth-child(3){animation-delay:.36s}
@keyframes cgtyp{0%,60%,100%{opacity:.35;transform:translateY(0)}30%{opacity:1;transform:translateY(-4px)}}
/* 15 spoke spinner */
.cg-spoke_spinner .sp{width:44px;height:44px}.cg-spoke_spinner line{stroke:var(--c1);stroke-width:3;stroke-linecap:round;opacity:.15;animation:cgspk 1.2s linear infinite}
@keyframes cgspk{0%{opacity:1}100%{opacity:.15}}
/* 16 percent bar */
@property --cgp{syntax:'<number>';initial-value:0;inherits:true}
.cg-percent_bar{width:160px;height:auto;flex-direction:column;gap:8px;align-items:stretch}
.cg-percent_bar i{display:block;height:8px;border-radius:8px;background:rgba(1,114,236,.16);position:relative;overflow:hidden}
.cg-percent_bar i:before{content:"";position:absolute;inset:0;width:calc(var(--cgp)*1%);background:linear-gradient(90deg,var(--c2),var(--c1));border-radius:8px}
.cg-percent_bar:after{counter-reset:p calc(var(--cgp));content:counter(p) "%";text-align:center;font:700 13px system-ui,sans-serif;color:var(--c2)}
.cg-percent_bar{animation:cgpct 3.2s ease-in-out infinite}
@keyframes cgpct{0%{--cgp:0}35%{--cgp:33}70%{--cgp:71}100%{--cgp:100}}
/* 17 chasing squares */
.cg-chasing_squares{width:48px;height:48px}.cg-chasing_squares i{position:absolute;top:0;left:0;width:16px;height:16px;border-radius:3px;background:var(--c1);animation:cgchase 1.6s ease-in-out infinite}
.cg-chasing_squares i:nth-child(2){background:var(--c3);animation-delay:-.8s}
@keyframes cgchase{0%{transform:translate(0,0)}25%{transform:translate(32px,0)}50%{transform:translate(32px,32px)}75%{transform:translate(0,32px)}100%{transform:translate(0,0)}}
/* 18 text shimmer */
.cg-text_shimmer{width:auto;height:auto}.cg-text_shimmer span{font:800 22px system-ui,-apple-system,sans-serif;letter-spacing:-.01em;background:linear-gradient(100deg,#9fb6d8 30%,var(--c1) 45%,var(--c3) 55%,#9fb6d8 70%);background-size:250% 100%;-webkit-background-clip:text;background-clip:text;color:transparent;animation:cgtxt 1.6s linear infinite}
@keyframes cgtxt{0%{background-position:150% 0}100%{background-position:-50% 0}}
/* 19 battery charge */
.cg-battery_charge{width:60px;height:28px;border:3px solid var(--c2);border-radius:6px;padding:2px;gap:3px;justify-content:flex-start}
.cg-battery_charge u{position:absolute;right:-8px;top:6px;width:5px;height:10px;background:var(--c2);border-radius:0 3px 3px 0}
.cg-battery_charge i{flex:1;height:100%;border-radius:2px;background:var(--c3);opacity:.15;animation:cgbat 1.6s steps(1) infinite}
.cg-battery_charge i:nth-child(2){animation-delay:.4s}.cg-battery_charge i:nth-child(3){animation-delay:.8s}.cg-battery_charge i:nth-child(4){animation-delay:1.2s}
@keyframes cgbat{0%{opacity:.15}1%,100%{opacity:1}}
/* 20 bouncing ball */
.cg-bouncing_ball{height:64px}.cg-bouncing_ball i{position:absolute;left:50%;margin-left:-10px;top:4px;width:20px;height:20px;border-radius:50%;background:var(--c1);animation:cgball .7s cubic-bezier(.3,0,.7,1) infinite alternate}
.cg-bouncing_ball b{position:absolute;left:50%;margin-left:-14px;bottom:2px;width:28px;height:6px;border-radius:50%;background:rgba(0,32,105,.22);animation:cgshadow .7s cubic-bezier(.3,0,.7,1) infinite alternate}
@keyframes cgball{0%{transform:translateY(0)}100%{transform:translateY(34px) scaleY(.85)}}
@keyframes cgshadow{0%{transform:scale(.5);opacity:.3}100%{transform:scale(1);opacity:.9}}
/* cortina */
#cg-ov{position:fixed;inset:0;z-index:2147483000;display:none;align-items:center;justify-content:center;flex-direction:column;gap:14px;background:rgba(255,255,255,.78);-webkit-backdrop-filter:blur(3px);backdrop-filter:blur(3px);font-family:system-ui,-apple-system,sans-serif;color:#002069;font-weight:700;font-size:14px}
#cg-ov.on{display:flex}
/* galería */
.cgg{max-width:1100px;margin:0 auto;padding:24px 18px 48px;font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#0f172a}
.cgg h1{font-size:28px;color:#002069;margin:0 0 6px}.cgg p.s{margin:0 0 20px;color:#475569}
.cgg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px}
.cgg-card{border:1px solid #dbeafe;border-radius:16px;background:#fff;padding:16px;display:flex;flex-direction:column;align-items:center;gap:10px;box-shadow:0 8px 22px rgba(1,114,236,.08)}
.cgg-card.sel{border-color:#03D047;box-shadow:0 0 0 3px rgba(3,208,71,.2)}
.cgg-stage{height:84px;display:flex;align-items:center;justify-content:center}
.cgg-card b{color:#002069}.cgg-card span.d{font-size:12.5px;color:#64748b;text-align:center}
.cgg-card button{border:0;border-radius:999px;padding:8px 16px;background:#0172EC;color:#fff;font-weight:700;cursor:pointer}
.cgg-card.sel button{background:#03A83A}
.cgg-top{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.cgg-top a,.cgg-top button{border:1px solid #bfdbfe;background:#fff;color:#002069;border-radius:999px;padding:8px 14px;font-weight:700;text-decoration:none;cursor:pointer}
"""


def _js(defecto):
    import json
    return """(function(){
var M=%s, DEF=%s, KEY='procsis_cargador';
function elegido(){var k=null;try{k=localStorage.getItem(KEY);}catch(e){}return (k&&M[k])?k:DEF;}
function ov(){var o=document.getElementById('cg-ov');if(o)return o;o=document.createElement('div');o.id='cg-ov';o.setAttribute('role','status');o.setAttribute('aria-live','polite');
 var k=elegido();o.innerHTML='<div class="cg cg-'+k+'">'+M[k]+'</div><span>Un momento…</span>';document.body.appendChild(o);return o;}
var t=null;
function mostrar(){clearTimeout(t);t=setTimeout(function(){ov().classList.add('on');setTimeout(ocultar,15000);},300);}
function ocultar(){clearTimeout(t);var o=document.getElementById('cg-ov');if(o)o.classList.remove('on');}
window.addEventListener('pageshow',ocultar);
document.addEventListener('click',function(e){
 if(e.defaultPrevented||e.button!==0||e.metaKey||e.ctrlKey||e.shiftKey||e.altKey)return;
 var a=e.target.closest&&e.target.closest('a[href]');if(!a)return;
 if(a.target&&a.target!=='_self')return;if(a.hasAttribute('download')||a.hasAttribute('data-no-cargador'))return;
 var h=a.getAttribute('href')||'';if(!h||h.charAt(0)==='#'||/^(javascript|mailto|tel|blob|data):/i.test(h))return;
 try{var u=new URL(a.href,location.href);if(u.origin!==location.origin)return;if(u.pathname===location.pathname&&u.search===location.search&&u.hash)return;if(/\\/(pdf|descargar|exportar)/i.test(u.pathname))return;}catch(x){return;}
 mostrar();},true);
document.addEventListener('submit',function(e){
 var f=e.target;if(!f||e.defaultPrevented||f.hasAttribute('data-no-cargador')||(f.target&&f.target!=='_self'))return;mostrar();},true);
window.procsisCargador={mostrar:mostrar,ocultar:ocultar,elegir:function(k){if(M[k]){try{localStorage.setItem(KEY,k);}catch(e){}var o=document.getElementById('cg-ov');if(o)o.remove();return true;}return false;},actual:elegido};
})();""" % (json.dumps(MARKUP, ensure_ascii=False), json.dumps(defecto))


def _galeria(defecto):
    import html
    tarj = ""
    for k, n, d in NOMBRES:
        tarj += ('<div class="cgg-card" data-k="%s"><div class="cgg-stage"><div class="cg cg-%s">%s</div></div><b>%d. %s</b><span class="d">%s</span>'
                 '<button type="button" onclick="elegir(\'%s\')">Usar este</button></div>'
                 % (k, k, MARKUP[k], [x[0] for x in NOMBRES].index(k) + 1, html.escape(n), html.escape(d), k))
    return """<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Animaciones de carga | PROCSIS</title><link rel="stylesheet" href="/cargadores/cargadores.css"></head><body style="margin:0;background:#f8fbff">
<div class="cgg"><div class="cgg-top"><a href="/" data-no-cargador>← Volver</a><button type="button" onclick="probar()">Probar cortina (3 s)</button></div>
<h1>20 animaciones de carga</h1><p class="s">Elige la que más vaya con el estilo del sistema. Se guarda en este navegador. Predeterminada del sistema: <b>%s</b>.</p>
<div class="cgg-grid">%s</div></div>
<script src="/cargadores/cargadores.js"></script>
<script>
function marcar(){var k=window.procsisCargador.actual();document.querySelectorAll('.cgg-card').forEach(function(c){var on=c.dataset.k===k;c.classList.toggle('sel',on);c.querySelector('button').textContent=on?'✓ En uso':'Usar este';});}
function elegir(k){window.procsisCargador.elegir(k);marcar();}
function probar(){window.procsisCargador.mostrar();setTimeout(window.procsisCargador.ocultar,3000);}
marcar();
</script></body></html>""" % (html.escape(defecto), tarj)


def instalar(G):
    modo = (os.environ.get("CARGADOR_MODO") or "on").strip().lower()
    if modo == "off":
        print("CARGADORES: apagado (CARGADOR_MODO=off).", flush=True)
        return
    try:
        app = G["app"]
        claves = [n[0] for n in NOMBRES]
        defecto = (os.environ.get("CARGADOR") or "spinner_arc").strip().lower()
        if defecto not in claves:
            defecto = "spinner_arc"
        inyectar_on = (os.environ.get("CARGADOR_INYECTAR") or "on").strip().lower() != "off"

        def _resp(cuerpo, mime):
            r = make_response(cuerpo)
            r.headers["Content-Type"] = mime + "; charset=utf-8"
            r.headers["Cache-Control"] = "public, max-age=3600"
            return r

        def cg_css():
            return _resp(CSS, "text/css")

        def cg_js():
            return _resp(_js(defecto), "application/javascript")

        def cg_galeria():
            r = make_response(_galeria(defecto))
            r.headers["Cache-Control"] = "no-cache"
            return r

        app.add_url_rule("/cargadores", "cg_galeria", cg_galeria, methods=["GET"])
        app.add_url_rule("/cargadores/cargadores.css", "cg_css", cg_css, methods=["GET"])
        app.add_url_rule("/cargadores/cargadores.js", "cg_js", cg_js, methods=["GET"])

        if inyectar_on:
            TAGS = '<link rel="stylesheet" href="/cargadores/cargadores.css"><script src="/cargadores/cargadores.js" defer></script>'

            @app.after_request
            def _cg_inyectar(resp):
                try:
                    p = request.path or ""
                    if (request.method != "GET" or resp.status_code != 200 or resp.mimetype != "text/html" or resp.direct_passthrough
                            or p.startswith(("/static", "/cargadores", "/publico/img")) or "/pdf" in p
                            or request.headers.get("X-Requested-With")):
                        return resp
                    h = resp.get_data(as_text=True)
                    if MARCA in h:
                        return resp
                    i = h.lower().rfind("</body>")
                    if i < 0:
                        return resp
                    resp.set_data(h[:i] + "<!--%s-->" % MARCA + TAGS + h[i:])
                except Exception as ex:
                    print("CARGADORES inyectar:", repr(ex)[:140], flush=True)
                return resp

        # Que el guardián de login deje ver la galería sin sesión (si la app usa ruta_publica()).
        orig = G.get("ruta_publica")
        if callable(orig):
            def ruta_publica_cg():
                return bool((request.path or "").startswith("/cargadores")) or bool(orig())
            G["ruta_publica"] = ruta_publica_cg
        print("CARGADORES: 20 animaciones instaladas (/cargadores). Predeterminada: %s." % defecto, flush=True)
    except Exception as ex:
        print("CARGADORES: NO se instaló (la app sigue igual):", repr(ex)[:160], flush=True)

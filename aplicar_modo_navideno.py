# -*- coding: utf-8 -*-
"""
aplicar_modo_navideno.py   (reemplaza a aplicar_splash_navideno.py; si ya lo aplicaste, este lo sustituye)

MODO NAVIDEÑO GLOBAL con interruptor en Marketing:
  1) Splash Screen: «Validando tu navegador y mejorando tu experiencia», 5 segundos, al entrar al portal
     y de nuevo cada vez que alguien INICIA SESIÓN (una vez por pestaña y por usuario). Se puede cerrar con un clic
     pasado 1,2 s.
  2) Page Loader: la misma pantalla (con el pino sobre la barra de progreso) al abrir cualquier módulo de cualquier
     área (clic en enlaces y envío de formularios). Se oculta apenas carga la página nueva; 5 s es solo el tiempo máximo.
     No sale en descargas (PDF, Excel, carnés…).
  3) Ambientación sobria: cinta tricolor fina arriba (verde bosque, vino, dorado), copo en la esquina de los
     encabezados de módulo, nieve muy suave que no bloquea clics, y un saludo de temporada una vez por sesión.
     Los colores de alerta (rojo error / verde éxito) NO se tocan.
  4) Cada persona puede apagarlo con el botón 🎄 (abajo a la izquierda): «Modo festivo».
  5) Feature Flag en la base de datos: Marketing y Gerencia entran a «🎄 Apariencia y campañas» (/sitio/apariencia)
     y activan/desactivan el Modo Navideño, la nieve y el saludo (y editan el texto) SIN volver a subir código.
     Tarda hasta ~20 s en verse en todos los servidores.
  Variable de entorno de emergencia: SPLASH_NAVIDENO=0 lo apaga todo.

Uso:  python aplicar_modo_navideno.py        (junto a app.py)
Deja un respaldo: app.py.respaldo_navideno
"""
import os
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "app.py")
MARCA = "# [procsis-modo-navideno]"
MARCA_VIEJA = "# [procsis-splash-navideno]"
ANCLA = 'if __name__ == "__main__":\n    with app.app_context():\n        inicializar_bd()'


def reemplazar(src, viejo, nuevo, nombre):
    n = src.count(viejo)
    if n != 1:
        print("ERROR en «%s»: se esperaba 1 coincidencia y hay %d. No se modificó nada." % (nombre, n))
        sys.exit(1)
    return src.replace(viejo, nuevo)


BLOQUE = r'''
''' + MARCA + r'''
# ═════════════════════ MODO NAVIDEÑO GLOBAL (splash 5 s · loader por módulo · adornos · feature flag) ═════════════════════
_SPLASH_ACTIVO = os.environ.get("SPLASH_NAVIDENO", "1") != "0"
_FEST_TXT = "¡Felices fiestas, {nombre}! Te lo desea todo el equipo PROCSIS."
_FEST_CACHE = {"t": 0.0, "v": None}


def _fest_cfg():
    """Feature flag leído de la base de datos (caché de 20 s). Por defecto: activado."""
    import time
    if _FEST_CACHE["v"] is not None and time.time() - _FEST_CACHE["t"] < 20:
        return _FEST_CACHE["v"]
    v = {"activo": True, "nieve": True, "saludo": True, "txt": _FEST_TXT}
    try:
        _ti_init()
        for k, c in (("activo", "fest_activo"), ("nieve", "fest_nieve"), ("saludo", "fest_saludo")):
            r = _ti_cfg(c)
            if r != "":
                v[k] = (r == "1")
        t = _ti_cfg("fest_saludo_txt")
        if t:
            v["txt"] = t[:200]
    except Exception:
        try:
            db.session.rollback()
        except Exception:
            pass
    _FEST_CACHE["v"] = v
    _FEST_CACHE["t"] = time.time()
    return v


_FEST_HTML = """<div id="pcs-splash" style="display:none" role="status" aria-live="polite" data-u="__U__" data-g="__G__" data-s="__S__">
<style>
#pcs-splash{position:fixed;inset:0;z-index:2147483000;background:#fff;align-items:center;justify-content:center;flex-direction:column;overflow:hidden;
font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased;transition:opacity .28s ease}
#pcs-splash.pcs-out{opacity:0;pointer-events:none}
#pcs-splash .pcs-bg{position:absolute;inset:0;width:100%;height:100%}
#pcs-splash .pcs-c{position:relative;display:flex;flex-direction:column;align-items:center;text-align:center;padding:0 24px;max-width:420px}
#pcs-splash .pcs-ic{width:54px;height:54px;margin-bottom:26px}
#pcs-splash .pcs-t{font-size:18px;font-weight:500;color:#111827;line-height:1.35;margin:0 0 46px;letter-spacing:-.005em}
#pcs-splash .pcs-bar{position:relative;width:min(300px,72vw);height:10px;border-radius:99px;border:1.5px solid #cbd5e1;background:#fff}
#pcs-splash .pcs-fill{position:absolute;left:0;top:0;bottom:0;width:0;border-radius:99px;background:#0b6e7f}
#pcs-splash .pcs-tree{position:absolute;top:-34px;left:0;width:36px;height:46px;transform:translateX(-50%)}
#pcs-splash .pcs-p{margin:14px 0 0;font-size:17px;color:#111827;font-variant-numeric:tabular-nums}
#pcs-splash .pcs-s{margin:22px 0 0;font-size:12px;color:#64748b;display:flex;align-items:center;gap:6px}
#pcs-fest-btn{position:fixed;left:12px;bottom:12px;z-index:2147482500;width:34px;height:34px;border-radius:50%;border:1px solid #e2e8f0;background:#fff;box-shadow:0 2px 8px rgba(15,23,42,.15);
cursor:pointer;font-size:16px;line-height:1;padding:0;opacity:.8}
#pcs-fest-btn:hover{opacity:1}#pcs-fest-btn.pcs-off{filter:grayscale(1);opacity:.55}
#pcs-rib{position:fixed;left:0;right:0;top:0;height:3px;z-index:2147482000;pointer-events:none;background:linear-gradient(90deg,#1f5f3b,#7a1b2b,#c9a227,#1f5f3b)}
.pcs-corner{position:absolute;top:8px;right:12px;width:22px;height:22px;opacity:.5;pointer-events:none}
html.pcs-fest-on .role-hero{border-top:2px solid #c9a227!important}
.pcs-toast{position:fixed;right:16px;bottom:16px;max-width:320px;z-index:2147482600;background:#fff;border:1px solid #e2e8f0;border-left:4px solid #1f5f3b;border-radius:12px;padding:12px 14px;
font:13px/1.45 -apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",sans-serif;color:#0f172a;box-shadow:0 8px 24px rgba(15,23,42,.18);transition:opacity .4s ease}
@media(prefers-reduced-motion:reduce){#pcs-splash{transition:none}}
@media print{#pcs-splash,#pcs-rib,#pcs-fest-btn,#pcs-snow,.pcs-toast,.pcs-corner{display:none!important}}
</style>
<svg class="pcs-bg" viewBox="0 0 400 800" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
 <defs><radialGradient id="pcsg" cx="50%" cy="50%" r="50%"><stop offset="0" stop-color="#fdecef"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></radialGradient></defs>
 <circle cx="20" cy="40" r="120" fill="url(#pcsg)"/><circle cx="395" cy="70" r="110" fill="url(#pcsg)"/>
 <circle cx="395" cy="620" r="120" fill="url(#pcsg)"/><circle cx="10" cy="640" r="100" fill="url(#pcsg)"/>
 <circle cx="58" cy="102" r="20" fill="#f3f4f6"/><circle cx="352" cy="352" r="9" fill="#f3f4f6"/><circle cx="52" cy="560" r="26" fill="#f3f4f6"/>
 <g stroke="#d6cfc8" stroke-width="3" stroke-linecap="round" fill="none" transform="translate(350 70)">
  <path d="M0-26V26M-22-13L22 13M-22 13L22-13"/><path d="M-6-22L0-16L6-22M-6 22L0 16L6 22" stroke-width="2"/></g>
 <g transform="translate(46 190)"><polygon points="0,-13 4,-4 13,-4 6,2 8,12 0,6 -8,12 -6,2 -13,-4 -4,-4" fill="#f0b43c"/></g>
 <g transform="translate(386 300)"><polygon points="0,-15 5,-5 15,-5 7,2 9,13 0,7 -9,13 -7,2 -15,-5 -5,-5" fill="#f0b43c"/></g>
 <g transform="translate(42 600)"><rect x="-26" y="-6" width="52" height="40" rx="3" fill="#d62d2d"/><rect x="-5" y="-6" width="10" height="40" fill="#e9b13a"/>
  <path d="M0-6C-18-30-30-8-8-6ZM0-6C18-30 30-8 8-6Z" fill="#e9b13a"/></g>
 <g transform="translate(352 600)"><polygon points="0,-62 28,-22 -28,-22" fill="#2f6b2f"/><polygon points="0,-40 38,10 -38,10" fill="#2a5f2a"/><polygon points="0,-10 48,46 -48,46" fill="#245524"/>
  <rect x="-5" y="46" width="10" height="12" fill="#6b4a2b"/><circle cx="0" cy="-64" r="5" fill="#f0b43c"/></g>
 <g transform="translate(370 740)"><rect x="-6" y="-44" width="12" height="9" rx="2" fill="#e9b13a"/><circle cx="0" cy="-12" r="30" fill="#b3202a"/><ellipse cx="-9" cy="-22" rx="7" ry="10" fill="#fff" opacity=".25"/></g>
 <path d="M-20 780 C 120 700 260 770 420 700 L420 820 L-20 820Z" fill="#f6f4f2"/>
</svg>
<div class="pcs-c">
 <svg class="pcs-ic" viewBox="0 0 48 48" fill="none" aria-hidden="true"><path d="M24 4 8 10v12c0 11 7 18 16 22 9-4 16-11 16-22V10L24 4Z" stroke="#0b6e7f" stroke-width="3" stroke-linejoin="round"/>
  <path d="m16 24 6 6 11-12" stroke="#d62d2d" stroke-width="3.4" stroke-linecap="round" stroke-linejoin="round"/></svg>
 <p class="pcs-t" id="pcs-t">Validando tu navegador y mejorando tu experiencia</p>
 <div class="pcs-bar"><div class="pcs-fill" id="pcs-fill"></div>
  <svg class="pcs-tree" id="pcs-tree" viewBox="0 0 36 46" aria-hidden="true"><polygon points="18,2 28,16 8,16" fill="#2f6b2f"/><polygon points="18,10 31,28 5,28" fill="#2a5f2a"/><polygon points="18,20 34,40 2,40" fill="#245524"/>
   <rect x="15" y="40" width="6" height="6" fill="#6b4a2b"/><circle cx="18" cy="2" r="2.6" fill="#f0b43c"/></svg></div>
 <p class="pcs-p" id="pcs-pct">0%</p>
 <p class="pcs-s"><svg width="11" height="13" viewBox="0 0 11 13" fill="#64748b" aria-hidden="true"><path d="M2 5V3.5a3.5 3.5 0 0 1 7 0V5h.5A1.5 1.5 0 0 1 11 6.5v5A1.5 1.5 0 0 1 9.5 13h-8A1.5 1.5 0 0 1 0 11.5v-5A1.5 1.5 0 0 1 1.5 5H2Zm1.5 0h4V3.5a2 2 0 0 0-4 0V5Z"/></svg>Conexión cifrada · Sesión protegida</p>
</div>
<script>
(function(){
 var el=document.getElementById('pcs-splash');if(!el)return;
 var usr=el.getAttribute('data-u')||'',greet=el.getAttribute('data-g')||'',conNieve=el.getAttribute('data-s')==='1';
 var root=document.documentElement,off=false,reduced=false;
 try{off=localStorage.getItem('pcs_fest')==='off';}catch(e){}
 try{reduced=window.matchMedia('(prefers-reduced-motion: reduce)').matches;}catch(e){}
 var DUR=reduced?1500:5000,MAXL=5000;
 var fill=document.getElementById('pcs-fill'),tree=document.getElementById('pcs-tree'),pct=document.getElementById('pcs-pct'),tit=document.getElementById('pcs-t');
 var shown=false,mode='',raf=0,timer=0,delay=0,t0=0,startMs=0,savedOv='',pendGreet=null;
 function setv(v){fill.style.width=v+'%';tree.style.left=v+'%';pct.textContent=v+'%';}
 function hide(){
  if(!shown)return;shown=false;cancelAnimationFrame(raf);clearTimeout(timer);
  el.classList.add('pcs-out');root.style.overflow=savedOv;
  setTimeout(function(){if(!shown)el.style.display='none';},320);
  if(mode==='splash'&&pendGreet){var g=pendGreet;pendGreet=null;setTimeout(g,350);}
 }
 function step(ts){
  if(!shown)return;if(!t0)t0=ts;var d=ts-t0,v;
  if(mode==='splash'){var p=Math.min(1,d/DUR);v=Math.round((1-Math.pow(1-p,1.6))*100);setv(v);if(p>=1){setTimeout(hide,150);return;}}
  else{v=Math.round(90*(1-Math.exp(-d/1400)));setv(v);}
  raf=requestAnimationFrame(step);
 }
 function show(m){
  if(off)return;mode=m;
  tit.textContent=(m==='splash')?'Validando tu navegador y mejorando tu experiencia':'Cargando tu módulo…';
  el.classList.remove('pcs-out');el.style.display='flex';
  if(!shown)savedOv=root.style.overflow;root.style.overflow='hidden';shown=true;t0=0;setv(0);
  clearTimeout(timer);timer=setTimeout(hide,m==='splash'?DUR+900:MAXL);
  cancelAnimationFrame(raf);raf=requestAnimationFrame(step);
 }
 function armar(){if(off||shown)return;clearTimeout(delay);delay=setTimeout(function(){show('load');},150);}
 if(!off)root.classList.add('pcs-fest-on');

 /* Splash: primera entrada y cada inicio de sesión (una vez por pestaña y usuario) */
 var key='pcs_splash_v2',visto=null,okSt=false;
 try{sessionStorage.setItem('pcs_t','1');okSt=sessionStorage.getItem('pcs_t')==='1';visto=sessionStorage.getItem(key);}catch(e){}
 if(!off&&okSt&&visto!==usr){try{sessionStorage.setItem(key,usr);}catch(e){}startMs=Date.now();show('splash');}
 el.addEventListener('click',function(){if(mode==='splash'&&shown&&Date.now()-startMs>1200)hide();});

 /* Loader al abrir módulos: enlaces y formularios (no en descargas) */
 var NODESC=/(pdf|descarg|download|export|xlsx|csv|docx|zip|[/]qr)/i;
 document.addEventListener('click',function(e){
  var a=e.target&&e.target.closest?e.target.closest('a[href]'):null;if(!a||e.defaultPrevented||e.button!==0||e.metaKey||e.ctrlKey||e.shiftKey||e.altKey)return;
  if(a.target&&a.target!=='_self')return;if(a.hasAttribute('download')||a.hasAttribute('data-sin-loader'))return;
  var h=a.getAttribute('href')||'';if(!h||h.charAt(0)==='#'||/^(javascript|mailto|tel|blob|data):/i.test(h))return;
  var u;try{u=new URL(a.href,location.href);}catch(x){return;}
  if(u.origin!==location.origin||NODESC.test(u.pathname))return;
  if(u.pathname===location.pathname&&u.search===location.search&&u.hash)return;
  armar();
 });
 document.addEventListener('submit',function(e){
  var f=e.target;if(!f||e.defaultPrevented||(f.target&&f.target!=='_self')||f.hasAttribute('data-sin-loader'))return;
  var ac=f.getAttribute('action')||location.pathname;if(NODESC.test(ac))return;armar();
 });
 window.addEventListener('pageshow',function(ev){if(ev.persisted){clearTimeout(delay);shown=true;hide();}});

 function listo(fn){if(document.readyState!=='loading')fn();else document.addEventListener('DOMContentLoaded',fn);}
 listo(function(){
  var b=document.createElement('button');b.id='pcs-fest-btn';b.type='button';b.textContent='🎄';
  b.title='Modo festivo: '+(off?'apagado':'activo')+' · clic para cambiar';b.setAttribute('aria-label',b.title);if(off)b.className='pcs-off';
  b.addEventListener('click',function(){try{localStorage.setItem('pcs_fest',off?'on':'off');}catch(e){}location.reload();});
  document.body.appendChild(b);
  if(off)return;
  var r=document.createElement('div');r.id='pcs-rib';r.setAttribute('aria-hidden','true');document.body.appendChild(r);
  var hs=document.querySelectorAll('.role-hero');
  for(var i=0;i<hs.length;i++){
   var h=hs[i];try{if(getComputedStyle(h).position==='static')h.style.position='relative';}catch(x){}
   var c=document.createElementNS('http://www.w3.org/2000/svg','svg');c.setAttribute('class','pcs-corner');c.setAttribute('viewBox','0 0 22 22');c.setAttribute('aria-hidden','true');
   c.innerHTML='<g stroke="#c9a227" stroke-width="1.6" stroke-linecap="round" fill="none"><path d="M11 1v20M2.3 6l17.4 10M2.3 16 19.7 6"/></g>';h.appendChild(c);
  }
  if(conNieve&&!reduced)nieve();
  if(greet&&okSt){
   var gk='pcs_greet_v1',gv=null;try{gv=sessionStorage.getItem(gk);}catch(e){}
   if(gv!==usr){
    var go=function(){try{sessionStorage.setItem(gk,usr);}catch(e){}var t=document.createElement('div');t.className='pcs-toast';t.setAttribute('role','status');t.textContent=greet;
     document.body.appendChild(t);setTimeout(function(){t.style.opacity='0';},6200);setTimeout(function(){if(t.parentNode)t.parentNode.removeChild(t);},6800);};
    if(shown&&mode==='splash')pendGreet=go;else go();
   }
  }
 });
 function nieve(){
  var c=document.createElement('canvas');c.id='pcs-snow';c.setAttribute('aria-hidden','true');
  c.style.cssText='position:fixed;left:0;top:0;width:100%;height:100%;pointer-events:none;z-index:1;opacity:.55';document.body.appendChild(c);
  var x=c.getContext('2d'),W=0,H=0,fl=[];
  function mk(a){return{x:Math.random()*W,y:a?Math.random()*H:-8,r:1+Math.random()*2.2,v:.35+Math.random()*.7,d:Math.random()*6.28,s:.4+Math.random()*.8};}
  function size(){W=c.width=window.innerWidth;H=c.height=window.innerHeight;var n=Math.max(14,Math.min(42,Math.round(W/36)));while(fl.length<n)fl.push(mk(true));fl.length=n;}
  size();window.addEventListener('resize',size);
  (function lazo(){
   if(!document.hidden){x.clearRect(0,0,W,H);x.fillStyle='rgba(150,178,215,.9)';
    for(var i=0;i<fl.length;i++){var f=fl[i];f.y+=f.v;f.d+=.012;f.x+=Math.sin(f.d)*f.s;if(f.y>H+8)fl[i]=mk(false);x.beginPath();x.arc(f.x,f.y,f.r,0,6.2832);x.fill();}}
   requestAnimationFrame(lazo);
  })();
 }
})();
</script>
</div>"""
_SPLASH_NO_RUTAS = ("/static", "/api", "/webhook", "/carnet", "/descargar", "/export", "/qr", "/foto", "/archivo", "/favicon")


@app.after_request
def _splash_navideno(resp):
    """Inserta el Modo Navideño (splash, loader, adornos) al inicio del <body> de las páginas HTML completas."""
    try:
        if not _SPLASH_ACTIVO or request.method != "GET" or resp.status_code != 200:
            return resp
        if resp.direct_passthrough or not resp.is_sequence or (resp.mimetype or "") != "text/html":
            return resp
        if request.headers.get("Sec-Fetch-Dest", "document") not in ("document", ""):
            return resp   # iframes, objetos, etc.
        if request.headers.get("X-Requested-With") or request.args.get("embed") or request.args.get("sin_splash"):
            return resp
        p = request.path or ""
        if p.startswith(_SPLASH_NO_RUTAS) or "/pdf" in p or p.lower().endswith((".pdf", ".png", ".jpg", ".svg", ".xlsx", ".docx", ".csv")):
            return resp
        cfg = _fest_cfg()
        if not cfg.get("activo"):
            return resp   # Marketing apagó el Modo Navideño: diseño corporativo tradicional
        html = resp.get_data(as_text=True)
        if 'id="pcs-splash"' in html:
            return resp
        m = re.search(r"<body[^>]*>", html, re.I)
        if not m:
            return resp
        import html as _hm
        usr = (session.get("usuario") or "").strip()
        nom = (session.get("nombre_completo") or session.get("nombre") or usr or "").strip()
        primer = nom.split(" ")[0] if nom else ""
        saludo = (cfg.get("txt") or _FEST_TXT).replace("{nombre}", primer or "equipo") if (cfg.get("saludo") and usr) else ""
        bloque = (_FEST_HTML.replace("__U__", _hm.escape(usr, True)).replace("__G__", _hm.escape(saludo, True))
                  .replace("__S__", "1" if cfg.get("nieve") else "0"))
        resp.set_data(html[:m.end()] + bloque + html[m.end():])
    except Exception as _e:
        print("modo navideño:", _e)
    return resp


# ───────────────── Marketing / Gerencia · «Apariencia y campañas» (interruptor del Modo Navideño) ─────────────────
def _fest_guard():
    if not requiere_login():
        return None, redirect("/login")
    rol = rol_actual()
    if rol != "Marketing" and rol not in _DCX_GERENCIA:
        return None, acceso_denegado("Solo Marketing y Gerencia administran la apariencia y las campañas del sistema.")
    try:
        from flask import g as _g_fest
        _g_fest._staff_wrap = True
    except Exception:
        pass
    return rol, None


@app.route("/sitio/apariencia", methods=["GET", "POST"])
def fest_apariencia():
    rol, r = _fest_guard()
    if r is not None:
        return r
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/sitio/apariencia")
        try:
            f = request.form
            _ti_cfg_set("fest_activo", "1" if f.get("activo") else "0")
            _ti_cfg_set("fest_nieve", "1" if f.get("nieve") else "0")
            _ti_cfg_set("fest_saludo", "1" if f.get("saludo") else "0")
            _ti_cfg_set("fest_saludo_txt", " ".join((f.get("txt") or "").split())[:200] or _FEST_TXT)
            _FEST_CACHE["t"] = 0.0
            _FEST_CACHE["v"] = None
            registrar_auditoria("Apariencia y campañas", "%s: Modo Navideño %s" % (session.get("usuario") or "", "ACTIVADO" if f.get("activo") else "DESACTIVADO"))
            session["ti_msg"] = ("ok", "Guardado. El cambio se ve en todo el sistema en menos de 20 segundos, sin subir código.")
        except Exception as e:
            try:
                db.session.rollback()
            except Exception:
                pass
            print("apariencia:", e)
            session["ti_msg"] = ("err", "No se pudo guardar. Intenta de nuevo.")
        return redirect("/sitio/apariencia")
    c = _fest_cfg()

    def sw(nombre, activo, titulo, ayuda):
        return ("<label style='display:flex;gap:12px;align-items:flex-start;padding:12px 0;border-bottom:1px solid #e2e8f0;cursor:pointer'>"
                "<input type='checkbox' name='%s' value='1' %s style='width:20px;height:20px;margin-top:2px'>"
                "<span><b>%s</b><br><span style='color:#64748b;font-size:13px'>%s</span></span></label>" % (nombre, "checked" if activo else "", titulo, ayuda))
    cuerpo = ("<h1>🎄 Apariencia y campañas</h1><p class='sub'>Interruptor del diseño de temporada. Al guardar, se aplica a todas las áreas del sistema "
              "sin volver a subir código. Los colores de alerta (rojo de error, verde de éxito) nunca cambian, y cada persona puede apagar los adornos "
              "con el botón 🎄 de la esquina inferior izquierda.</p>"
              "<div class='cb-card'><form method='POST' action='/sitio/apariencia'><input type='hidden' name='csrf' value='%s'>%s%s%s"
              "<label style='display:block;margin-top:14px'><b>Texto del saludo de temporada</b></label>"
              "<input type='text' name='txt' maxlength='200' value='%s' style='width:100%%;max-width:560px'>"
              "<p style='color:#64748b;font-size:13px'>Usa <code>{nombre}</code> para el primer nombre de la persona. Sale una vez por inicio de sesión.</p>"
              "<p style='margin-top:14px'><button class='cb-btn g'>Guardar cambios</button></p></form></div>"
              "<div class='cb-card'><h2>Qué incluye el Modo Navideño</h2><p>Pantalla «Validando tu navegador y mejorando tu experiencia» (5 s al iniciar sesión), "
              "pantalla de carga con el pino al abrir cada módulo, cinta tricolor fina en la parte superior, copo dorado en los encabezados, nieve suave y saludo de temporada.</p></div>"
              % (_esc(csrf),
                 sw("activo", c["activo"], "Modo Navideño", "Enciende o apaga todo el diseño de temporada para todos los usuarios."),
                 sw("nieve", c["nieve"], "Nieve suave de fondo", "Copos muy discretos que no bloquean clics. Se desactiva sola si el equipo pide reducir movimiento."),
                 sw("saludo", c["saludo"], "Saludo de temporada", "Aviso corto, una vez por inicio de sesión."),
                 _esc(c["txt"])))
    return _cb_pagina("Apariencia y campañas", cuerpo)


'''


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
    if src.count(ANCLA) != 1:
        print("ERROR: no encontré el punto de inserción (final de app.py). No se modificó nada.")
        sys.exit(1)
    shutil.copyfile(RUTA, RUTA + ".respaldo_navideno")

    # 1) Bloque nuevo (sustituye al splash anterior si estaba)
    j = src.index(ANCLA)
    i = src.find(MARCA_VIEJA)
    nuevo = BLOQUE.lstrip("\n")
    if 0 <= i < j:
        src = src[:i] + nuevo + src[j:]
        print("Reemplacé la pantalla anterior (aplicar_splash_navideno) por el Modo Navideño completo.")
    else:
        src = src.replace(ANCLA, nuevo + ANCLA)

    # 2) El filtro de rutas deja pasar /sitio/apariencia (valida su propio acceso: Marketing y Gerencia)
    src = reemplazar(
        src,
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        '    if path.startswith("/sitio/apariencia"):\n'
        '        return None  # [procsis-apariencia] _fest_guard valida el acceso\n'
        '    if path.startswith(("/tickets-doc", "/notificaciones", "/api/v1")):\n',
        "filtro de rutas",
    )

    # 3) Menú lateral: «🎄 Apariencia y campañas» para Marketing y Gerencia
    src = reemplazar(
        src,
        '    if rol in _ROLES_MAPA and not any(h == "/mapa" for h, _l in items):\n',
        '    if (rol in _DCX_GERENCIA or rol == "Marketing") and not any(h == "/sitio/apariencia" for h, _l in items):\n'
        '        items = [x for x in items if x[0] != "/logout"] + [("/sitio/apariencia", "🎄 Apariencia y campañas")] + [x for x in items if x[0] == "/logout"]\n'
        '    if rol in _ROLES_MAPA and not any(h == "/mapa" for h, _l in items):\n',
        "menú lateral",
    )

    # 4) Tarjeta en el portal de Marketing
    src = reemplazar(
        src,
        '    filas = "".join(\n'
        '        "<tr><td><b>%s</b></td><td>%s</td><td>%s</td></tr>" % (\n'
        '            _esc(u.usuario), _esc(getattr(u, "nombre_completo", "") or "—"),\n',
        '    if rol_area == "Marketing":\n'
        '        tarjetas += ("<a href=\'/sitio/apariencia\' class=\'role-panel\' style=\'border-left:6px solid %s;text-decoration:none;display:block\'>"\n'
        '                     "<div style=\'font-weight:800;color:#0f172a\'>🎄 Apariencia y campañas</div>"\n'
        '                     "<div style=\'font-size:12px;color:#475569;margin-top:4px\'>Enciende o apaga el Modo Navideño de todo el sistema.</div>"\n'
        '                     "<div style=\'margin-top:8px\'><span style=\'background:#ede9fe;color:#5b21b6;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700\'>Abrir</span></div></a>" % a["color"])\n'
        '    filas = "".join(\n'
        '        "<tr><td><b>%s</b></td><td>%s</td><td>%s</td></tr>" % (\n'
        '            _esc(u.usuario), _esc(getattr(u, "nombre_completo", "") or "—"),\n',
        "tarjeta portal Marketing",
    )

    # 5) Admisión de personal descarga PDF: sin pantalla de carga en esos formularios
    a = src.find('@app.route("/gerencia/admision-personal"')
    b = src.find('@app.route("/gerencia/procsis-web"')
    if 0 < a < b:
        trozo = src[a:b].replace('<form method="POST" style="display:flex;', '<form method="POST" data-sin-loader style="display:flex;')
        src = src[:a] + trozo + src[b:]

    if crlf:
        src = src.replace("\n", "\r\n")
    with open(RUTA, "w", encoding="utf-8", newline="") as f:
        f.write(src)
    print("Listo. Modo Navideño aplicado. Respaldo en app.py.respaldo_navideno")


if __name__ == "__main__":
    main()

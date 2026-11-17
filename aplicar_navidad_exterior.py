#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_navidad_exterior.py
Uso:  python aplicar_navidad_exterior.py            (busca app.py en la carpeta actual)
      python aplicar_navidad_exterior.py ruta/app.py

Qué hace sobre app.py (guarda copia app.py.bak_navidad_ext antes de tocar nada):
  1) Quita el «Cargando tu módulo…» de la zona pública (login y páginas de afuera) y baja el tope del loader a 2.5 s.
  2) Decora el área pública: luces, fondo festivo en azul claro / azul rey / verde / toques de rojo, copos,
     adornos colgantes en la tarjeta del login y gorro de Santa en el logo de PROCSIS.
  3) Agrega 2 interruptores en Marketing → «Apariencia y campañas»: «Decoración de la zona pública» y «Gorro navideño en el logo».
     Todo sigue colgado del interruptor general «Modo Navideño».
Es seguro correrlo dos veces: si ya está aplicado, no hace nada.
"""
import sys, os, shutil

MARCA = "# [procsis-navidad-exterior]"
BLOQUE = r'''# [procsis-navidad-exterior]
# ═════════════ NAVIDAD · ÁREA PÚBLICA (login y páginas de afuera): luces, fondo festivo, adornos y gorro en el logo ═════════════
# Se enciende/apaga desde Marketing → «Apariencia y campañas» (interruptores «Decoración de la zona pública» y «Gorro navideño en el logo»).
_NAVEXT_HTML = r"""<div id="pcsx-lights" aria-hidden="true" data-hat="__HAT__"></div>
<style>
html.pcsx{--pcsx-tile:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='200' height='40' viewBox='0 0 200 40'><path d='M0 4Q25 20 50 4Q75 20 100 4Q125 20 150 4Q175 20 200 4' fill='none' stroke='%23334155' stroke-width='1.6'/><g><rect x='22' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='25' cy='22' rx='5.5' ry='7.5' fill='%2322d3ee'/><rect x='72' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='75' cy='22' rx='5.5' ry='7.5' fill='%232563eb'/><rect x='122' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='125' cy='22' rx='5.5' ry='7.5' fill='%2322c55e'/><rect x='172' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='175' cy='22' rx='5.5' ry='7.5' fill='%23ef4444'/></g></svg>");
--pcsx-nieve:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='260' height='260' viewBox='0 0 260 260'><g fill='%23ffffff' opacity='.95'><circle cx='30' cy='40' r='4'/><circle cx='120' cy='22' r='3'/><circle cx='210' cy='70' r='5'/><circle cx='70' cy='130' r='3'/><circle cx='170' cy='150' r='4'/><circle cx='235' cy='215' r='3'/><circle cx='40' cy='220' r='5'/><circle cx='125' cy='235' r='3'/></g><g stroke='%23bae6fd' stroke-width='2' stroke-linecap='round' fill='none'><path d='M150 60v18M141 69h18M144 63l12 12M156 63l-12 12'/><path d='M60 180v16M52 188h16M54 182l12 12M66 182l-12 12'/><path d='M225 130v14M218 137h14M220 132l10 10M230 132l-10 10'/></g></svg>")}
#pcsx-lights{display:none}
html.pcsx #pcsx-lights{display:block;position:relative;height:36px;z-index:60;pointer-events:none;overflow:hidden;background:var(--pcsx-tile) repeat-x 0 0}
html.pcsx:not(.pcsx-still) #pcsx-lights::after{content:"";position:absolute;inset:0;background:var(--pcsx-tile) repeat-x 0 0;filter:blur(5px) saturate(1.8);animation:pcsx-tw 1.9s ease-in-out infinite alternate}
@keyframes pcsx-tw{from{opacity:.1}to{opacity:.9}}
html.pcsx body{background:linear-gradient(180deg,#e3f4ff 0,#f4fbff 40%,#fff 100%)!important}
html.pcsx .login-page,html.pcsx .lp-content{background:linear-gradient(180deg,#e3f4ff 0,#f4fbff 45%,#fff 100%)!important}
html.pcsx .lp-topnav{background:rgba(255,255,255,.96);border-bottom:3px solid transparent;border-image:linear-gradient(90deg,#22d3ee,#1d4ed8,#22c55e,#ef4444) 1}
html.pcsx .sinai-portal{position:relative;background:radial-gradient(720px 440px at 6% 0,rgba(34,211,238,.30),transparent 62%),radial-gradient(660px 420px at 100% 4%,rgba(34,197,94,.20),transparent 60%),radial-gradient(560px 380px at 94% 100%,rgba(239,68,68,.10),transparent 60%),#eaf6ff!important}
html.pcsx .sinai-portal::before{content:"";position:absolute;inset:0;z-index:0;pointer-events:none;background:var(--pcsx-nieve) repeat 0 0}
html.pcsx .sinai-grid{position:relative;z-index:1}
html.pcsx .sinai-access{position:relative}
html.pcsx .sinai-card{position:relative;overflow:hidden;border-color:#bfdbfe!important;box-shadow:0 14px 38px rgba(0,32,194,.14)!important}
html.pcsx .sinai-card::before{content:"";position:absolute;left:0;right:0;top:0;height:6px;background:repeating-linear-gradient(135deg,#1d4ed8 0 12px,#fff 12px 18px,#22c55e 18px 30px,#fff 30px 36px,#ef4444 36px 48px,#fff 48px 54px)}
html.pcsx .sinai-carousel{box-shadow:0 30px 80px rgba(0,32,194,.30),0 0 0 3px rgba(255,255,255,.7)}
html.pcsx .corp-footer{border-top:4px solid transparent;border-image:linear-gradient(90deg,#22d3ee,#1d4ed8,#22c55e,#ef4444) 1}
.pcsx-hang{position:absolute;right:16px;top:6px;width:96px;height:auto;pointer-events:none;z-index:3}
.pcsx-w{position:relative;line-height:0}
.pcsx-hat{position:absolute;pointer-events:none;overflow:visible;transform:rotate(-16deg);transform-origin:50% 100%;filter:drop-shadow(0 3px 4px rgba(0,0,0,.28));z-index:5}
@media(max-width:600px){.pcsx-hang{display:none}}
@media print{#pcsx-lights,.pcsx-hang,.pcsx-hat{display:none!important}}
</style>
<script>
(function(){
 var L=document.getElementById('pcsx-lights');if(!L)return;
 var root=document.documentElement,off=false,reduced=false;
 try{off=localStorage.getItem('pcs_fest')==='off';}catch(e){}
 try{reduced=window.matchMedia('(prefers-reduced-motion: reduce)').matches;}catch(e){}
 if(off){L.style.display='none';return;}
 root.classList.add('pcsx');if(reduced)root.classList.add('pcsx-still');
 var conGorro=L.getAttribute('data-hat')==='1';
 var HANG='<svg class="pcsx-hang" viewBox="0 0 96 58" aria-hidden="true"><defs>'
  +'<radialGradient id="pxb" cx="35%" cy="30%" r="75%"><stop offset="0" stop-color="#60a5fa"/><stop offset="1" stop-color="#1e3a8a"/></radialGradient>'
  +'<radialGradient id="pxg" cx="35%" cy="30%" r="75%"><stop offset="0" stop-color="#86efac"/><stop offset="1" stop-color="#15803d"/></radialGradient>'
  +'<radialGradient id="pxr" cx="35%" cy="30%" r="75%"><stop offset="0" stop-color="#fca5a5"/><stop offset="1" stop-color="#b91c1c"/></radialGradient></defs>'
  +'<g stroke="#94a3b8" stroke-width="1.2"><path d="M16 0V14M48 0V26M80 0V18"/></g>'
  +'<g fill="#e9b13a"><rect x="13" y="12" width="6" height="4" rx="1"/><rect x="45" y="24" width="6" height="4" rx="1"/><rect x="77" y="16" width="6" height="4" rx="1"/></g>'
  +'<circle cx="16" cy="26" r="10" fill="url(#pxb)"/><circle cx="48" cy="38" r="12" fill="url(#pxg)"/><circle cx="80" cy="29" r="10" fill="url(#pxr)"/>'
  +'<g fill="#fff" opacity=".45"><ellipse cx="12.5" cy="22" rx="2.6" ry="4"/><ellipse cx="43.5" cy="33" rx="3" ry="4.6"/><ellipse cx="76.5" cy="25" rx="2.6" ry="4"/></g></svg>';
 var HAT='<svg class="pcsx-hat" viewBox="0 0 140 116" aria-hidden="true"><defs><linearGradient id="pxh" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ef4444"/><stop offset="1" stop-color="#b91c1c"/></linearGradient></defs>'
  +'<path d="M18 88C20 44 46 14 82 16C100 17 114 28 124 44C112 40 102 46 100 58C98 70 104 80 112 90Z" fill="url(#pxh)"/>'
  +'<path d="M30 80C34 52 52 30 78 26" stroke="#fff" stroke-opacity=".3" stroke-width="5" fill="none" stroke-linecap="round"/>'
  +'<rect x="8" y="80" width="112" height="26" rx="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/>'
  +'<circle cx="124" cy="44" r="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/></svg>';
 function listo(fn){if(document.readyState!=='loading')fn();else document.addEventListener('DOMContentLoaded',fn);}
 function colocar(w,img,hat){
  var bw=img.clientWidth,bh=img.clientHeight,nw=img.naturalWidth||bw,nh=img.naturalHeight||bh;if(!bw||!bh)return;
  var sc=Math.min(bw/nw,bh/nh),cw=nw*sc,ch=nh*sc,ox=img.offsetLeft+(bw-cw)/2,oy=img.offsetTop+(bh-ch)/2,full=(nw/nh)>1.25;
  var hw=full?cw*.21:cw*.46,hh=hw*116/140;
  hat.style.width=hw+'px';
  hat.style.left=(ox+(full?cw*.40:cw*.10))+'px';
  hat.style.top=(oy+(full?ch*.03:-hh*.55))+'px';
 }
 function gorro(img){
  if(img.getAttribute('data-pcsx'))return;
  if(!/logo[-_]?procsis/i.test(img.getAttribute('src')||''))return;
  img.setAttribute('data-pcsx','1');
  var blk=false;try{blk=getComputedStyle(img).display==='block';}catch(e){}
  var w=document.createElement('span');w.className='pcsx-w';w.style.display=blk?'block':'inline-block';
  img.parentNode.insertBefore(w,img);w.appendChild(img);w.insertAdjacentHTML('beforeend',HAT);
  var hat=w.lastChild,go=function(){colocar(w,img,hat);};
  if(img.complete)go();img.addEventListener('load',go);window.addEventListener('resize',go);
  if(window.ResizeObserver){try{new ResizeObserver(go).observe(img);}catch(e){}}
 }
 listo(function(){
  var cd=document.querySelector('.sinai-card');if(cd&&!cd.querySelector('.pcsx-hang'))cd.insertAdjacentHTML('afterbegin',HANG);
  if(conGorro){var im=document.querySelectorAll('img');for(var i=0;i<im.length;i++)gorro(im[i]);}
 });
})();
</script>
"""


@app.after_request
def _navidad_exterior(resp):
    """Decora solo las páginas públicas (las que se ven sin iniciar sesión). Sin loader de módulo."""
    try:
        if not _SPLASH_ACTIVO or request.method != "GET" or resp.status_code != 200:
            return resp
        if resp.direct_passthrough or not resp.is_sequence or (resp.mimetype or "") != "text/html":
            return resp
        if request.headers.get("Sec-Fetch-Dest", "document") not in ("document", ""):
            return resp
        if request.headers.get("X-Requested-With") or request.args.get("embed") or request.args.get("sin_splash"):
            return resp
        p = request.path or ""
        if p.startswith(_SPLASH_NO_RUTAS) or "/pdf" in p:
            return resp
        if not ruta_publica():
            return resp
        cfg = _fest_cfg()
        if not (cfg.get("activo") and cfg.get("ext", True)):
            return resp
        html = resp.get_data(as_text=True)
        if 'id="pcsx-lights"' in html:
            return resp
        m = re.search(r"<body[^>]*>", html, re.I)
        if not m:
            return resp
        bloque = _NAVEXT_HTML.replace("__HAT__", "1" if cfg.get("gorro", True) else "0")
        resp.set_data(html[:m.end()] + bloque + html[m.end():])
    except Exception as _e:
        print("navidad exterior:", _e)
    return resp


'''


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else "app.py"
    if not os.path.isfile(ruta):
        sys.exit("No encuentro %s. Ejecuta el script dentro de la carpeta del proyecto o pasa la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as f:
        txt = f.read()
    if MARCA in txt:
        print("Ya estaba aplicado. No se cambió nada.")
        return
    eol = "\r\n" if "\r\n" in txt else "\n"

    def n(s):
        return s.replace("\r\n", "\n").replace("\n", eol)

    cambios = [
        # 1) configuración: dos banderas nuevas (por defecto encendidas)
        ('    v = {"activo": True, "nieve": True, "saludo": True, "txt": _FEST_TXT}',
         '    v = {"activo": True, "nieve": True, "saludo": True, "ext": True, "gorro": True, "txt": _FEST_TXT}'),
        ('for k, c in (("activo", "fest_activo"), ("nieve", "fest_nieve"), ("saludo", "fest_saludo")):',
         'for k, c in (("activo", "fest_activo"), ("nieve", "fest_nieve"), ("saludo", "fest_saludo"), ("ext", "fest_ext"), ("gorro", "fest_gorro")):'),
        # 2) guardado desde Marketing
        ('            _ti_cfg_set("fest_saludo", "1" if f.get("saludo") else "0")',
         '            _ti_cfg_set("fest_saludo", "1" if f.get("saludo") else "0")\n'
         '            _ti_cfg_set("fest_ext", "1" if f.get("ext") else "0")\n'
         '            _ti_cfg_set("fest_gorro", "1" if f.get("gorro") else "0")'),
        # 3) panel de Marketing: dos interruptores más
        ("value='%s'>%s%s%s\"", "value='%s'>%s%s%s%s%s\""),
        ('                 sw("saludo", c["saludo"], "Saludo de temporada", "Aviso corto, una vez por inicio de sesión."),',
         '                 sw("saludo", c["saludo"], "Saludo de temporada", "Aviso corto, una vez por inicio de sesión."),\n'
         '                 sw("ext", c["ext"], "Decoración de la zona pública", "Luces, fondo festivo y adornos en el login y las páginas de afuera."),\n'
         '                 sw("gorro", c["gorro"], "Gorro navideño en el logo", "Sombrero de Santa sobre el logo de PROCSIS."),'),
        # 4) loader «Cargando tu módulo»: no en zona pública, tope más corto, y se quita al volver a la pestaña
        ("function armar(){if(off||shown)return;", "function armar(){if(off||shown||!usr)return;"),
        ("var DUR=reduced?1500:5000,MAXL=5000;", "var DUR=reduced?1500:5000,MAXL=2500;"),
        (" window.addEventListener('pageshow',function(ev){if(ev.persisted){clearTimeout(delay);shown=true;hide();}});",
         " window.addEventListener('pageshow',function(ev){if(ev.persisted){clearTimeout(delay);shown=true;hide();}});\n"
         " window.addEventListener('load',function(){clearTimeout(delay);if(shown&&mode==='load')hide();});\n"
         " document.addEventListener('visibilitychange',function(){if(!document.hidden&&shown&&mode==='load')hide();});"),
    ]
    nuevo = txt
    for viejo, repl in cambios:
        v, r = n(viejo), n(repl)
        c = nuevo.count(v)
        if c != 1:
            sys.exit("No apliqué nada: esperaba 1 coincidencia y hallé %d para:\n  %s\n"
                     "Probablemente app.py cambió desde que armé este script; avísame y lo ajusto." % (c, viejo[:110]))
        nuevo = nuevo.replace(v, r)

    ancla = 'if __name__ == "__main__":'
    i = nuevo.rfind(ancla)
    if i < 0:
        sys.exit('No encontré el bloque final «if __name__ == "__main__":». No apliqué nada.')
    nuevo = nuevo[:i] + n(BLOQUE) + nuevo[i:]

    try:
        compile(nuevo, ruta, "exec")
    except SyntaxError as e:
        sys.exit("El resultado tenía un error de sintaxis (%s). No toqué tu app.py." % e)

    shutil.copyfile(ruta, ruta + ".bak_navidad_ext")
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        f.write(nuevo)
    print("Listo. Copia de seguridad: %s.bak_navidad_ext" % ruta)
    print("Reinicia la app y entra a Marketing → Apariencia y campañas para ver los interruptores nuevos.")


if __name__ == "__main__":
    main()

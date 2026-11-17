#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_navidad_ejecutiva.py   (segundo parche · se corre DESPUÉS de aplicar_navidad_exterior.py)
Uso:  python aplicar_navidad_ejecutiva.py            (busca app.py en la carpeta actual)
      python aplicar_navidad_ejecutiva.py ruta/app.py

Qué hace sobre app.py (guarda copia app.py.bak_navidad_ejec antes de tocar nada):
  1) Pantalla de carga «Validando tu navegador…» con aspecto ejecutivo: fondo azul corporativo, logo PROCSIS con gorro
     de Navidad, barra verde/cian con una estrella dorada y pie «PROCSIS · Innovación que gestiona».
  2) Pantalla de bienvenida (la del saludo «Buenas tardes, señor …»): fondo festivo suave y gorro en el logo.
  3) Todas las áreas administrativas (Gerencia, Ventas, Talento, Finanzas, Marketing, Soporte…): luces discretas arriba,
     línea tricolor en la barra superior, estrella dorada en el menú, «Felices fiestas» en el pie del menú, nieve tenue y
     cinta de colores en los encabezados, y gorro en el logo.
  4) Marketing → «Apariencia y campañas»: interruptor nuevo «Decoración en las áreas administrativas».
Es seguro correrlo dos veces: si ya está aplicado, no hace nada.
"""
import sys, os, shutil

MARCA_1 = "# [procsis-navidad-exterior]"
MARCA_2 = "# [procsis-navidad-ejecutiva]"

NAVEXT = r'''<div id="pcsx-lights" aria-hidden="true" data-hat="__HAT__"></div>
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

 var HAT='<svg class="pcsx-hat" viewBox="0 0 140 116" aria-hidden="true"><defs><linearGradient id="pxh" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ef4444"/><stop offset="1" stop-color="#b91c1c"/></linearGradient></defs><path d="M18 88C20 44 46 14 82 16C100 17 114 28 124 44C112 40 102 46 100 58C98 70 104 80 112 90Z" fill="url(#pxh)"/><path d="M30 80C34 52 52 30 78 26" stroke="#fff" stroke-opacity=".3" stroke-width="5" fill="none" stroke-linecap="round"/><rect x="8" y="80" width="112" height="26" rx="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/><circle cx="124" cy="44" r="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/></svg>';
 function esLogo(img){var s=img.getAttribute('src')||'',a=img.getAttribute('alt')||'';return /logo[-_]?procsis|logo-corporativo/i.test(s)||/procsis/i.test(a);}
 function colocar(img,hat){
  var cs=getComputedStyle(img),pl=parseFloat(cs.paddingLeft)||0,pr=parseFloat(cs.paddingRight)||0,pt=parseFloat(cs.paddingTop)||0,pb=parseFloat(cs.paddingBottom)||0;
  var bw=img.clientWidth-pl-pr,bh=img.clientHeight-pt-pb;if(bw<=0||bh<=0){hat.style.display='none';return;}
  var nw=img.naturalWidth||bw,nh=img.naturalHeight||bh,sc=Math.min(bw/nw,bh/nh),cw=nw*sc,ch=nh*sc;
  var ox=img.offsetLeft+pl+(bw-cw)/2,oy=img.offsetTop+pt+(bh-ch)/2,full=(nw/nh)>1.25;
  if(cw<26){hat.style.display='none';return;}
  var hw=Math.max(full?cw*.24:cw*.46,20),hh=hw*116/140;
  hat.style.width=hw+'px';hat.style.left=(ox+(full?cw*.385:cw*.10))+'px';hat.style.top=(oy+(full?ch*.03:-hh*.55))+'px';hat.style.display='block';
 }
 function gorro(img){
  if(img.getAttribute('data-pcsx')||!esLogo(img))return;
  img.setAttribute('data-pcsx','1');
  var blk=false;try{blk=getComputedStyle(img).display==='block';}catch(e){}
  var w=document.createElement('span');w.className='pcsx-w';w.style.display=blk?'block':'inline-block';if(!blk)w.style.verticalAlign='top';
  img.parentNode.insertBefore(w,img);w.appendChild(img);w.insertAdjacentHTML('beforeend',HAT);
  var hat=w.lastChild;hat.style.display='none';
  var go=function(){colocar(img,hat);};
  if(img.complete)go();img.addEventListener('load',go);window.addEventListener('resize',go);
  if(window.ResizeObserver){try{new ResizeObserver(go).observe(img);}catch(e){}}
 }
 function todos(){if(!conGorro)return;var im=document.querySelectorAll('img');for(var i=0;i<im.length;i++)gorro(im[i]);}

 function listo(fn){if(document.readyState!=='loading')fn();else document.addEventListener('DOMContentLoaded',fn);}
 listo(function(){
  var cd=document.querySelector('.sinai-card');if(cd&&!cd.querySelector('.pcsx-hang'))cd.insertAdjacentHTML('afterbegin',HANG);
  todos();window.addEventListener('load',todos);setTimeout(todos,1500);
 });
})();
</script>
'''
NAVINT = r'''<div id="pcxi-lights" aria-hidden="true" data-hat="__HAT__"></div>
<style>
html.pcxi{--pcxi-tile:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='200' height='40' viewBox='0 0 200 40'><path d='M0 4Q25 20 50 4Q75 20 100 4Q125 20 150 4Q175 20 200 4' fill='none' stroke='%2394a3b8' stroke-width='1.6'/><g><rect x='22' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='25' cy='22' rx='5.5' ry='7.5' fill='%2322d3ee'/><rect x='72' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='75' cy='22' rx='5.5' ry='7.5' fill='%232563eb'/><rect x='122' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='125' cy='22' rx='5.5' ry='7.5' fill='%2322c55e'/><rect x='172' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='175' cy='22' rx='5.5' ry='7.5' fill='%23ef4444'/></g></svg>");--pcxi-nieve:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='240' height='240' viewBox='0 0 240 240'><g fill='%23ffffff'><circle cx='24' cy='36' r='2.6'/><circle cx='110' cy='18' r='2'/><circle cx='196' cy='62' r='3'/><circle cx='64' cy='122' r='2'/><circle cx='156' cy='140' r='2.6'/><circle cx='222' cy='200' r='2'/><circle cx='36' cy='206' r='3'/><circle cx='118' cy='222' r='2'/></g><g stroke='%23ffffff' stroke-width='1.4' stroke-linecap='round' fill='none'><path d='M140 56v14M133 63h14M135 58l10 10M145 58l-10 10'/><path d='M54 168v12M48 174h12M50 170l8 8M58 170l-8 8'/></g></svg>")}
#pcxi-lights{display:none}
html.pcxi #pcxi-lights{display:block;position:fixed;left:0;right:0;top:0;height:15px;z-index:9000;pointer-events:none;overflow:hidden;background:var(--pcxi-tile) repeat-x 0 -3px;background-size:110px 22px;opacity:.95}
html.pcxi:not(.pcxi-still) #pcxi-lights::after{content:"";position:absolute;inset:0;background:var(--pcxi-tile) repeat-x 0 -3px;background-size:110px 22px;filter:blur(3px) saturate(1.7);animation:pcxi-tw 2.1s ease-in-out infinite alternate}
@keyframes pcxi-tw{from{opacity:.05}to{opacity:.8}}
html.pcxi .navbar-top-apple{border-bottom:2px solid transparent!important;border-image:linear-gradient(90deg,#00B8D4,#0020C2 38%,#00FF66 72%,#ef4444) 1}
html.pcxi .brand-text-apple::after{content:"\2726";color:#f5c542;margin-left:8px;font-size:.85em}
html.pcxi .sidebar-footer-legal::before{content:"Felices fiestas \00b7  ";color:#f5c542;font-weight:600}
html.pcxi .hq-corp,html.pcxi .role-hero{position:relative}
html.pcxi .hq-corp::before,html.pcxi .role-hero::before{content:"";position:absolute;left:0;right:0;top:0;height:4px;border-radius:inherit;border-bottom-left-radius:0;border-bottom-right-radius:0;background:linear-gradient(90deg,#00B8D4,#0020C2 34%,#00FF66 66%,#ef4444);pointer-events:none}
html.pcxi .hq-corp::after,html.pcxi .role-hero::after{content:"";position:absolute;inset:0;border-radius:inherit;pointer-events:none;background:var(--pcxi-nieve) repeat 0 0;opacity:.30}
html.pcxi .bv-wrap{position:relative;background:radial-gradient(640px 420px at 8% 0,rgba(0,184,212,.20),transparent 62%),radial-gradient(620px 420px at 100% 6%,rgba(0,255,102,.14),transparent 60%),radial-gradient(520px 360px at 92% 100%,rgba(239,68,68,.08),transparent 60%),#eef6ff!important}
html.pcxi .bv-wrap::before{content:"";position:absolute;inset:0;pointer-events:none;background:var(--pcxi-nieve) repeat 0 0;opacity:.85}
html.pcxi .bv-box{position:relative}
html.pcxi .bv-sub::after{content:" \00b7  Felices fiestas";color:#15803d;font-weight:600}
html.pcxi .bv-btn{box-shadow:0 8px 22px rgba(0,32,194,.28),0 0 0 3px rgba(245,197,66,.35)}
.pcsx-w{position:relative;line-height:0}
.pcsx-hat{position:absolute;pointer-events:none;overflow:visible;transform:rotate(-16deg);transform-origin:50% 100%;filter:drop-shadow(0 3px 4px rgba(0,0,0,.28));z-index:5}
@media print{#pcxi-lights,.pcsx-hat{display:none!important}}
</style>
<script>
(function(){
 var L=document.getElementById('pcxi-lights');if(!L)return;
 var root=document.documentElement,off=false,reduced=false;
 try{off=localStorage.getItem('pcs_fest')==='off';}catch(e){}
 try{reduced=window.matchMedia('(prefers-reduced-motion: reduce)').matches;}catch(e){}
 if(off){L.style.display='none';return;}
 root.classList.add('pcxi');if(reduced)root.classList.add('pcxi-still');
 var conGorro=L.getAttribute('data-hat')==='1';

 var HAT='<svg class="pcsx-hat" viewBox="0 0 140 116" aria-hidden="true"><defs><linearGradient id="pxh" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ef4444"/><stop offset="1" stop-color="#b91c1c"/></linearGradient></defs><path d="M18 88C20 44 46 14 82 16C100 17 114 28 124 44C112 40 102 46 100 58C98 70 104 80 112 90Z" fill="url(#pxh)"/><path d="M30 80C34 52 52 30 78 26" stroke="#fff" stroke-opacity=".3" stroke-width="5" fill="none" stroke-linecap="round"/><rect x="8" y="80" width="112" height="26" rx="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/><circle cx="124" cy="44" r="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/></svg>';
 function esLogo(img){var s=img.getAttribute('src')||'',a=img.getAttribute('alt')||'';return /logo[-_]?procsis|logo-corporativo/i.test(s)||/procsis/i.test(a);}
 function colocar(img,hat){
  var cs=getComputedStyle(img),pl=parseFloat(cs.paddingLeft)||0,pr=parseFloat(cs.paddingRight)||0,pt=parseFloat(cs.paddingTop)||0,pb=parseFloat(cs.paddingBottom)||0;
  var bw=img.clientWidth-pl-pr,bh=img.clientHeight-pt-pb;if(bw<=0||bh<=0){hat.style.display='none';return;}
  var nw=img.naturalWidth||bw,nh=img.naturalHeight||bh,sc=Math.min(bw/nw,bh/nh),cw=nw*sc,ch=nh*sc;
  var ox=img.offsetLeft+pl+(bw-cw)/2,oy=img.offsetTop+pt+(bh-ch)/2,full=(nw/nh)>1.25;
  if(cw<26){hat.style.display='none';return;}
  var hw=Math.max(full?cw*.24:cw*.46,20),hh=hw*116/140;
  hat.style.width=hw+'px';hat.style.left=(ox+(full?cw*.385:cw*.10))+'px';hat.style.top=(oy+(full?ch*.03:-hh*.55))+'px';hat.style.display='block';
 }
 function gorro(img){
  if(img.getAttribute('data-pcsx')||!esLogo(img))return;
  img.setAttribute('data-pcsx','1');
  var blk=false;try{blk=getComputedStyle(img).display==='block';}catch(e){}
  var w=document.createElement('span');w.className='pcsx-w';w.style.display=blk?'block':'inline-block';if(!blk)w.style.verticalAlign='top';
  img.parentNode.insertBefore(w,img);w.appendChild(img);w.insertAdjacentHTML('beforeend',HAT);
  var hat=w.lastChild;hat.style.display='none';
  var go=function(){colocar(img,hat);};
  if(img.complete)go();img.addEventListener('load',go);window.addEventListener('resize',go);
  if(window.ResizeObserver){try{new ResizeObserver(go).observe(img);}catch(e){}}
 }
 function todos(){if(!conGorro)return;var im=document.querySelectorAll('img');for(var i=0;i<im.length;i++)gorro(im[i]);}

 function listo(fn){if(document.readyState!=='loading')fn();else document.addEventListener('DOMContentLoaded',fn);}
 listo(function(){todos();window.addEventListener('load',todos);setTimeout(todos,1500);});
})();
</script>
'''
SPLASH_CSS = r'''#pcs-splash{position:fixed;inset:0;z-index:2147483000;color:#fff;background:radial-gradient(900px 560px at 12% 6%,rgba(0,184,212,.30),transparent 60%),radial-gradient(760px 520px at 94% 98%,rgba(0,255,102,.16),transparent 62%),linear-gradient(135deg,#000F66 0%,#0020C2 58%,#0018A8 100%);align-items:center;justify-content:center;flex-direction:column;overflow:hidden;
font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased;transition:opacity .28s ease}
#pcs-splash.pcs-out{opacity:0;pointer-events:none}
#pcs-splash .pcs-bg{position:absolute;inset:0;width:100%;height:100%}
#pcs-splash .pcs-lt{position:absolute;left:0;right:0;top:0;height:34px;pointer-events:none;background:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='200' height='40' viewBox='0 0 200 40'><path d='M0 4Q25 20 50 4Q75 20 100 4Q125 20 150 4Q175 20 200 4' fill='none' stroke='%23cbd5e1' stroke-width='1.6'/><g><rect x='22' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='25' cy='22' rx='5.5' ry='7.5' fill='%2322d3ee'/><rect x='72' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='75' cy='22' rx='5.5' ry='7.5' fill='%232563eb'/><rect x='122' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='125' cy='22' rx='5.5' ry='7.5' fill='%2322c55e'/><rect x='172' y='10' width='6' height='5' rx='1' fill='%23475569'/><ellipse cx='175' cy='22' rx='5.5' ry='7.5' fill='%23ef4444'/></g></svg>") repeat-x 0 0}
#pcs-splash .pcs-c{position:relative;display:flex;flex-direction:column;align-items:center;text-align:center;padding:0 24px;max-width:460px}
#pcs-splash .pcs-lg{position:relative;width:min(250px,64vw);aspect-ratio:3/2;box-sizing:border-box;padding:12px;margin-bottom:34px;background:#fff;border-radius:26px;box-shadow:0 26px 60px rgba(0,8,64,.50),0 0 0 4px rgba(255,255,255,.18)}
#pcs-splash .pcs-logo{display:block;width:100%;height:100%;object-fit:contain}
#pcs-splash .pcs-hat{position:absolute;left:39%;top:5%;width:20%;height:auto;pointer-events:none;overflow:visible;transform:rotate(-16deg);transform-origin:50% 100%;filter:drop-shadow(0 3px 5px rgba(0,0,0,.30))}
#pcs-splash[data-h="0"] .pcs-hat{display:none}
#pcs-splash .pcs-t{font-size:17px;font-weight:600;color:#fff;line-height:1.4;margin:0 0 30px;letter-spacing:.005em}
#pcs-splash .pcs-bar{position:relative;width:min(330px,74vw);height:6px;border-radius:99px;background:rgba(255,255,255,.22)}
#pcs-splash .pcs-fill{position:absolute;left:0;top:0;bottom:0;width:0;border-radius:99px;background:linear-gradient(90deg,#00B8D4,#00FF66);box-shadow:0 0 12px rgba(0,255,102,.55)}
#pcs-splash .pcs-tree{position:absolute;top:-8px;left:0;width:22px;height:22px;transform:translateX(-50%);filter:drop-shadow(0 0 6px rgba(245,197,66,.8))}
#pcs-splash .pcs-p{margin:18px 0 0;font-size:15px;font-weight:600;color:#fff;font-variant-numeric:tabular-nums;letter-spacing:.06em}
#pcs-splash .pcs-s{margin:20px 0 0;font-size:12px;color:rgba(255,255,255,.74);display:flex;align-items:center;gap:6px}
#pcs-splash .pcs-s svg{fill:rgba(255,255,255,.74)}
#pcs-splash .pcs-f{position:absolute;left:0;right:0;bottom:22px;text-align:center;font-size:11px;letter-spacing:.2em;text-transform:uppercase;color:rgba(255,255,255,.55)}
'''
SPLASH_BG = r'''<svg class="pcs-bg" viewBox="0 0 400 800" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
 <g fill="#fff" opacity=".18"><circle cx="38" cy="96" r="3"/><circle cx="352" cy="140" r="2.4"/><circle cx="70" cy="330" r="2"/><circle cx="372" cy="420" r="3"/><circle cx="30" cy="560" r="2.6"/><circle cx="330" cy="690" r="2"/><circle cx="120" cy="740" r="3"/><circle cx="250" cy="60" r="2"/></g>
 <g stroke="#fff" stroke-opacity=".2" stroke-width="1.5" stroke-linecap="round" fill="none"><path d="M62 190v18M53 199h18M56 193l12 12M68 193l-12 12"/><path d="M346 300v16M338 308h16M340 302l12 12M352 302l-12 12"/><path d="M52 640v14M45 647h14M47 642l10 10M57 642l-10 10"/><path d="M352 590v18M343 599h18M346 593l12 12M358 593l-12 12"/></g>
 <g fill="#f5c542" opacity=".8"><polygon points="0,-9 2.7,-2.7 9,-2.7 4,1.6 5.8,8 0,4.2 -5.8,8 -4,1.6 -9,-2.7 -2.7,-2.7" transform="translate(48 250)"/><polygon points="0,-7 2.1,-2.1 7,-2.1 3.1,1.2 4.5,6.2 0,3.3 -4.5,6.2 -3.1,1.2 -7,-2.1 -2.1,-2.1" transform="translate(356 520)"/><polygon points="0,-6 1.8,-1.8 6,-1.8 2.7,1 3.9,5.3 0,2.8 -3.9,5.3 -2.7,1 -6,-1.8 -1.8,-1.8" transform="translate(300 110)"/></g>
 <path d="M-20 704C120 640 260 724 420 650" stroke="#00FF66" stroke-opacity=".28" stroke-width="2" fill="none"/>
 <path d="M-20 128C120 64 260 144 420 74" stroke="#00B8D4" stroke-opacity=".30" stroke-width="2" fill="none"/>
</svg>
<div class="pcs-lt"></div>
<div class="pcs-f">PROCSIS &middot; Innovaci&oacute;n que gestiona</div>
'''
SPLASH_LOGO = r'''<div class="pcs-lg"><img class="pcs-logo" data-pcsx="1" src="/media/logo-corporativo" alt="PROCSIS" onerror="this.onerror=null;this.src='/static/img/logo-procsis.svg'"><svg class="pcs-hat" viewBox="0 0 140 116" aria-hidden="true"><defs><linearGradient id="pxs" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ef4444"/><stop offset="1" stop-color="#b91c1c"/></linearGradient></defs><path d="M18 88C20 44 46 14 82 16C100 17 114 28 124 44C112 40 102 46 100 58C98 70 104 80 112 90Z" fill="url(#pxs)"/><path d="M30 80C34 52 52 30 78 26" stroke="#fff" stroke-opacity=".3" stroke-width="5" fill="none" stroke-linecap="round"/><rect x="8" y="80" width="112" height="26" rx="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/><circle cx="124" cy="44" r="13" fill="#fff" stroke="#cfe3f7" stroke-width="2"/></svg></div>
 <p class="pcs-t" id="pcs-t">Validando tu navegador y mejorando tu experiencia</p>
 <div class="pcs-bar"><div class="pcs-fill" id="pcs-fill"></div>
  <svg class="pcs-tree" id="pcs-tree" viewBox="0 0 22 22" aria-hidden="true"><polygon points="11,1 13.9,7.6 21,8.3 15.6,13 17.2,20 11,16.3 4.8,20 6.4,13 1,8.3 8.1,7.6" fill="#f5c542" stroke="#fff" stroke-width="1"/></svg></div>
 '''
NUEVO_RIB = r'''#pcs-rib{position:fixed;left:0;right:0;top:0;height:3px;z-index:2147482000;pointer-events:none;background:linear-gradient(90deg,#0020C2,#00B8D4,#00FF66,#c9a227,#0020C2)}'''

BLOQUE_INT = r'''# [procsis-navidad-ejecutiva]
# ═════════════ NAVIDAD CORPORATIVA · ÁREAS ADMINISTRATIVAS (luces discretas, detalles dorados, gorro en el logo) ═════════════
# Se enciende/apaga desde Marketing → «Apariencia y campañas» → «Decoración en las áreas administrativas».
_NAVINT_HTML = r"""__NAVINT__"""


@app.after_request
def _navidad_interna(resp):
    """Detalle navideño corporativo dentro de las áreas administrativas de PROCSIS (no toca los colegios)."""
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
        if p.startswith(_SPLASH_NO_RUTAS) or "/pdf" in p or ruta_publica():
            return resp
        if not session.get("usuario") or rol_actual() not in _BIENVENIDA_ROLES:
            return resp
        cfg = _fest_cfg()
        if not (cfg.get("activo") and cfg.get("int", True)):
            return resp
        html = resp.get_data(as_text=True)
        if 'id="pcxi-lights"' in html:
            return resp
        m = re.search(r"<body[^>]*>", html, re.I)
        if not m:
            return resp
        bloque = _NAVINT_HTML.replace("__HAT__", "1" if cfg.get("gorro", True) else "0")
        resp.set_data(html[:m.end()] + bloque + html[m.end():])
    except Exception as _e:
        print("navidad interna:", _e)
    return resp


'''


def cortar(txt, ini, fin, nuevo, desde=0):
    """Reemplaza lo que hay entre `ini` (incluido) y `fin` (excluido). Exige que `ini` exista una sola vez."""
    if txt.count(ini) != 1:
        sys.exit("No apliqué nada: esperaba 1 coincidencia y hallé %d para:\n  %s\n"
                 "Probablemente app.py cambió desde que armé este script; avísame y lo ajusto." % (txt.count(ini), ini[:100]))
    i = txt.find(ini)
    j = txt.find(fin, i)
    if j < 0:
        sys.exit("No apliqué nada: no encontré el final del bloque (%s)." % fin[:80])
    return txt[:i] + nuevo + txt[j:]


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else "app.py"
    if not os.path.isfile(ruta):
        sys.exit("No encuentro %s. Ejecuta el script dentro de la carpeta del proyecto o pasa la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as f:
        txt = f.read()
    if MARCA_2 in txt:
        print("Ya estaba aplicado. No se cambió nada.")
        return
    if MARCA_1 not in txt:
        sys.exit("Primero hay que aplicar aplicar_navidad_exterior.py (este parche se apoya en él). No toqué nada.")
    eol = "\r\n" if "\r\n" in txt else "\n"

    def n(s):
        return s.replace("\r\n", "\n").replace("\n", eol)

    # ── A) cambios de una sola línea ──
    lineas = [
        ('"ext": True, "gorro": True, "txt": _FEST_TXT}', '"ext": True, "gorro": True, "int": True, "txt": _FEST_TXT}'),
        ('("ext", "fest_ext"), ("gorro", "fest_gorro")):', '("ext", "fest_ext"), ("gorro", "fest_gorro"), ("int", "fest_int")):'),
        ('            _ti_cfg_set("fest_gorro", "1" if f.get("gorro") else "0")',
         '            _ti_cfg_set("fest_gorro", "1" if f.get("gorro") else "0")\n'
         '            _ti_cfg_set("fest_int", "1" if f.get("int") else "0")'),
        ("value='%s'>%s%s%s%s%s\"", "value='%s'>%s%s%s%s%s%s\""),
        ('                 sw("gorro", c["gorro"], "Gorro navideño en el logo", "Sombrero de Santa sobre el logo de PROCSIS."),',
         '                 sw("gorro", c["gorro"], "Gorro navideño en el logo", "Sombrero de Santa sobre el logo de PROCSIS."),\n'
         '                 sw("int", c["int"], "Decoración en las áreas administrativas", "Luces discretas, detalles dorados y gorro en el logo dentro de Gerencia, Ventas, Talento y demás áreas."),'),
        ('data-g="__G__" data-s="__S__">', 'data-g="__G__" data-s="__S__" data-h="__H__">'),
        ('.replace("__S__", "1" if cfg.get("nieve") else "0"))',
         '.replace("__S__", "1" if cfg.get("nieve") else "0")\n'
         '                  .replace("__H__", "1" if cfg.get("gorro", True) else "0"))'),
        ("#pcs-rib{position:fixed;left:0;right:0;top:0;height:3px;z-index:2147482000;pointer-events:none;background:linear-gradient(90deg,#1f5f3b,#7a1b2b,#c9a227,#1f5f3b)}",
         NUEVO_RIB),
    ]
    nuevo = txt
    for viejo, repl in lineas:
        v, r = n(viejo), n(repl)
        c = nuevo.count(v)
        if c != 1:
            sys.exit("No apliqué nada: esperaba 1 coincidencia y hallé %d para:\n  %s\n"
                     "Probablemente app.py cambió desde que armé este script; avísame y lo ajusto." % (c, viejo[:110]))
        nuevo = nuevo.replace(v, r)

    # ── B) pantalla de carga ejecutiva (reemplazo por bloques) ──
    nuevo = cortar(nuevo, "#pcs-splash{position:fixed;inset:0;z-index:2147483000", "#pcs-fest-btn{position:fixed", n(SPLASH_CSS))
    nuevo = cortar(nuevo, '<svg class="pcs-bg"', '<div class="pcs-c">', n(SPLASH_BG))
    nuevo = cortar(nuevo, '<svg class="pcs-ic"', '<p class="pcs-p" id="pcs-pct">', n(SPLASH_LOGO))

    # ── C) zona pública: versión mejorada del bloque (gorro bien centrado en logos con margen interno) ──
    ini = '_NAVEXT_HTML = r"""'
    if nuevo.count(ini) != 1:
        sys.exit("No encontré el bloque de la zona pública (_NAVEXT_HTML). No apliqué nada.")
    i = nuevo.find(ini) + len(ini)
    k = nuevo.find("def _navidad_exterior", i)
    j = nuevo.rfind('"""', i, k)
    if k < 0 or j < 0:
        sys.exit("No encontré el final del bloque de la zona pública. No apliqué nada.")
    nuevo = nuevo[:i] + n(NAVEXT) + nuevo[j:]

    # ── D) áreas administrativas ──
    ancla = 'if __name__ == "__main__":'
    p = nuevo.rfind(ancla)
    if p < 0:
        sys.exit('No encontré el bloque final «if __name__ == "__main__":». No apliqué nada.')
    nuevo = nuevo[:p] + n(BLOQUE_INT.replace("__NAVINT__", NAVINT)) + nuevo[p:]

    try:
        compile(nuevo, ruta, "exec")
    except SyntaxError as e:
        sys.exit("El resultado tenía un error de sintaxis (%s). No toqué tu app.py." % e)

    shutil.copyfile(ruta, ruta + ".bak_navidad_ejec")
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        f.write(nuevo)
    print("Listo. Copia de seguridad: %s.bak_navidad_ejec" % ruta)
    print("Reinicia la app. En Marketing → Apariencia y campañas aparece el interruptor nuevo.")


if __name__ == "__main__":
    main()

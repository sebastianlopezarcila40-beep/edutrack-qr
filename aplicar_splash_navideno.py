# -*- coding: utf-8 -*-
"""
aplicar_splash_navideno.py
Pantalla de bienvenida navideña: «Validando tu navegador y mejorando tu experiencia».

Reglas (para que sea profesional y no frene el trabajo):
  - Sale UNA sola vez por pestaña/sesión del navegador (al entrar por primera vez al portal).
  - Dura ~1,8 segundos (0,7 s si el usuario tiene activado «reducir movimiento»). Se puede cerrar con un clic.
  - No se muestra al cambiar de módulo, ni en PDFs, descargas, API, ni dentro de iframes.
  - Si hay un problema, la página carga normal (nunca queda bloqueada).
  - Se apaga sin tocar código: variable de entorno SPLASH_NAVIDENO=0 (en Vercel, Settings > Environment Variables).

Uso:  python aplicar_splash_navideno.py        (junto a app.py)
Deja un respaldo: app.py.respaldo_splash
"""
import os
import shutil
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "app.py")
MARCA = "# [procsis-splash-navideno]"
ANCLA = 'if __name__ == "__main__":\n    with app.app_context():\n        inicializar_bd()'

BLOQUE = r'''
''' + MARCA + r'''
# ═════════════════════════ Pantalla de bienvenida navideña (una vez por sesión, ~1,8 s) ═════════════════════════
_SPLASH_ACTIVO = os.environ.get("SPLASH_NAVIDENO", "1") != "0"
_SPLASH_HTML = """<div id="pcs-splash" style="display:none" role="status" aria-live="polite">
<style>
#pcs-splash{position:fixed;inset:0;z-index:2147483000;background:#fff;align-items:center;justify-content:center;flex-direction:column;overflow:hidden;
font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",Roboto,sans-serif;-webkit-font-smoothing:antialiased;transition:opacity .28s ease;cursor:pointer}
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
@media(prefers-reduced-motion:reduce){#pcs-splash{transition:none}}
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
 <p class="pcs-t">Validando tu navegador y mejorando tu experiencia</p>
 <div class="pcs-bar"><div class="pcs-fill" id="pcs-fill"></div>
  <svg class="pcs-tree" id="pcs-tree" viewBox="0 0 36 46" aria-hidden="true"><polygon points="18,2 28,16 8,16" fill="#2f6b2f"/><polygon points="18,10 31,28 5,28" fill="#2a5f2a"/><polygon points="18,20 34,40 2,40" fill="#245524"/>
   <rect x="15" y="40" width="6" height="6" fill="#6b4a2b"/><circle cx="18" cy="2" r="2.6" fill="#f0b43c"/></svg></div>
 <p class="pcs-p" id="pcs-pct">0%</p>
 <p class="pcs-s"><svg width="11" height="13" viewBox="0 0 11 13" fill="#64748b" aria-hidden="true"><path d="M2 5V3.5a3.5 3.5 0 0 1 7 0V5h.5A1.5 1.5 0 0 1 11 6.5v5A1.5 1.5 0 0 1 9.5 13h-8A1.5 1.5 0 0 1 0 11.5v-5A1.5 1.5 0 0 1 1.5 5H2Zm1.5 0h4V3.5a2 2 0 0 0-4 0V5Z"/></svg>Conexión cifrada · Sesión protegida</p>
</div>
<script>
(function(){
 var el=document.getElementById('pcs-splash');if(!el)return;
 try{if(sessionStorage.getItem('pcs_splash_v1')){el.parentNode.removeChild(el);return;}sessionStorage.setItem('pcs_splash_v1','1');}catch(e){}
 var corto=false;try{corto=window.matchMedia('(prefers-reduced-motion: reduce)').matches;}catch(e){}
 var dur=corto?700:1800,ini=null,fin=false,root=document.documentElement,ov=root.style.overflow;
 el.style.display='flex';root.style.overflow='hidden';
 var fill=document.getElementById('pcs-fill'),tree=document.getElementById('pcs-tree'),pct=document.getElementById('pcs-pct');
 function cerrar(){if(fin)return;fin=true;el.classList.add('pcs-out');root.style.overflow=ov;setTimeout(function(){if(el.parentNode)el.parentNode.removeChild(el);},320);}
 function paso(t){if(fin)return;if(ini===null)ini=t;var p=Math.min(1,(t-ini)/dur),e=1-Math.pow(1-p,2),v=Math.round(e*100);
  fill.style.width=v+'%';tree.style.left=v+'%';pct.textContent=v+'%';if(p<1)requestAnimationFrame(paso);else setTimeout(cerrar,120);}
 el.addEventListener('click',cerrar);
 setTimeout(cerrar,dur+900);
 requestAnimationFrame(paso);
})();
</script>
</div>"""
_SPLASH_NO_RUTAS = ("/static", "/api", "/webhook", "/carnet", "/descargar", "/export", "/qr", "/foto", "/archivo", "/favicon")


@app.after_request
def _splash_navideno(resp):
    """Inserta la pantalla de bienvenida al inicio del <body> (solo páginas HTML completas, una vez por sesión del navegador)."""
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
        html = resp.get_data(as_text=True)
        if 'id="pcs-splash"' in html:
            return resp
        m = re.search(r"<body[^>]*>", html, re.I)
        if not m:
            return resp
        resp.set_data(html[:m.end()] + _SPLASH_HTML + html[m.end():])
    except Exception as _e:
        print("splash:", _e)
    return resp


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
    shutil.copyfile(RUTA, RUTA + ".respaldo_splash")
    src = src.replace(ANCLA, BLOQUE.lstrip("\n") + ANCLA)
    if crlf:
        src = src.replace("\n", "\r\n")
    with open(RUTA, "w", encoding="utf-8", newline="") as f:
        f.write(src)
    print("Listo. Parche aplicado. Respaldo en app.py.respaldo_splash")


if __name__ == "__main__":
    main()

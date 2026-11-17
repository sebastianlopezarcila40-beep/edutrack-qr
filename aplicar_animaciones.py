# Animaciones de entrada por scroll (fade-in + slide-up, escalonadas) con IntersectionObserver.
# Uso:  python aplicar_animaciones.py            (trabaja sobre app.py de esta carpeta)
import io, os, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA = "_REVEAL_SNIPPET = r'''"
ANCLA_DEF = "def page(title, body):"
ANCLA_USO = "{html_anuncio_global()}{_bfcache_fix}{_n_guard}</body></html>"

SNIPPET = r"""_REVEAL_SNIPPET = r'''
<style id="rv-css">
/* ── Animaciones de entrada (solo transform + opacity: aceleradas por GPU) ── */
@keyframes rvUp{from{opacity:0;transform:translate3d(0,24px,0)}to{opacity:1;transform:translate3d(0,0,0)}}
.rv{opacity:0;transform:translate3d(0,30px,0);transition:opacity .7s cubic-bezier(.22,.61,.36,1) var(--rv-d,0ms),transform .7s cubic-bezier(.22,.61,.36,1) var(--rv-d,0ms);will-change:opacity,transform}
.rv.in{opacity:1;transform:translate3d(0,0,0)}
/* Entrada inmediata (sin JS) del encabezado de Sedes y de los resultados que se pintan al buscar */
.sa-hero>*{animation:rvUp .8s cubic-bezier(.22,.61,.36,1) backwards}
.sa-hero>*:nth-child(2){animation-delay:.1s}
.sa-hero>*:nth-child(3){animation-delay:.2s}
.sa-list .sa-item{animation:rvUp .55s cubic-bezier(.22,.61,.36,1) backwards}
.sa-list .sa-item:nth-child(2){animation-delay:.06s}
.sa-list .sa-item:nth-child(3){animation-delay:.12s}
.sa-list .sa-item:nth-child(4){animation-delay:.18s}
.sa-list .sa-item:nth-child(5){animation-delay:.24s}
.sa-list .sa-item:nth-child(6){animation-delay:.30s}
.sa-list .sa-item:nth-child(n+7){animation-delay:.36s}
@media(prefers-reduced-motion:reduce){.rv,.sa-hero>*,.sa-list .sa-item{animation:none!important;transition:none!important;opacity:1!important;transform:none!important}}
@media print{.rv{opacity:1!important;transform:none!important}}
</style>
<script>
(function(){
  if(!("IntersectionObserver" in window))return;
  if(window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches)return;
  /* [selector, retardo base en ms]. Dentro de cada grupo, cada hermano entra +100 ms después del anterior (máx. 6 pasos). */
  var GRUPOS=[
    [".sa-find .sa-wrap > h2, .sa-find .sa-wrap > .sa-lead, .sa-find .sa-wrap > .sa-search",0],
    [".sa-cat",200],
    [".sa-find .sa-actions",600],
    [".promo-card",0],
    [".corp-support-left, .corp-support-right",0],
    [".corp-cta-card",0],
    [".lp-section",0],
    [".sx-card, .sx-sec, .sx-cta",0],
    ["[data-reveal]",0]
  ];
  function revelar(el){
    el.classList.add("in");
    var fin=function(ev){
      if(ev&&ev.target!==el)return;
      el.removeEventListener("transitionend",fin);
      el.classList.remove("rv","in");
      el.style.removeProperty("--rv-d");
    };
    el.addEventListener("transitionend",fin);
    setTimeout(fin,2500);
  }
  function crear(umbral){
    return new IntersectionObserver(function(es,ob){
      es.forEach(function(e){
        if(!e.isIntersecting||e.intersectionRatio<umbral)return;
        ob.unobserve(e.target);
        revelar(e.target);
      });
    },{threshold:umbral,rootMargin:"0px 0px -6% 0px"});
  }
  try{
    var io=crear(0.12),ioAlto=crear(0.001);
    GRUPOS.forEach(function(g){
      var padres=[],cuenta=[];
      [].slice.call(document.querySelectorAll(g[0])).forEach(function(el){
        if(el.classList.contains("rv"))return;
        var i=padres.indexOf(el.parentNode);
        if(i<0){padres.push(el.parentNode);cuenta.push(0);i=padres.length-1;}
        var n=Math.min(cuenta[i]++,6);
        el.style.setProperty("--rv-d",(g[1]+n*100)+"ms");
        el.classList.add("rv");
        (el.offsetHeight>window.innerHeight*1.2?ioAlto:io).observe(el);
      });
    });
  }catch(err){
    [].slice.call(document.querySelectorAll(".rv")).forEach(function(el){el.classList.remove("rv");});
  }
})();
</script>
'''


"""

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
if MARCA in s:
    print("Este app.py YA tiene las animaciones. No se modificó nada.")
    sys.exit(0)
if s.count(ANCLA_DEF) != 1 or s.count(ANCLA_USO) != 1:
    print("ERROR: no encontré los puntos de inserción en tu app.py (def page / cierre del body). No se modificó nada.")
    sys.exit(1)
nl = "\r\n" if "\r\n" in s else "\n"
s = s.replace(ANCLA_DEF, SNIPPET.replace("\n", nl) + ANCLA_DEF, 1)
s = s.replace(ANCLA_USO, "{html_anuncio_global()}{_REVEAL_SNIPPET}{_bfcache_fix}{_n_guard}</body></html>", 1)
tmp = SRC + ".nuevo"
with io.open(tmp, "w", encoding="utf-8", newline="") as fh:
    fh.write(s)
try:
    py_compile.compile(tmp, doraise=True)
except Exception as e:
    os.remove(tmp)
    print("ERROR: el resultado no compila, no se modificó tu app.py:", e)
    sys.exit(1)
shutil.copyfile(SRC, SRC + ".respaldo")
os.replace(tmp, SRC)
print("LISTO. Animaciones de entrada por scroll aplicadas. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

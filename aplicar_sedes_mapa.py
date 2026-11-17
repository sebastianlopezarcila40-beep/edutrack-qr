# 1) Quita «Accesos institucionales» del login.  2) Rediseña «Sedes y puntos de atención» con la identidad verde-cian-azul.
# 3) Arregla el mapa de ubicaciones (/mapa) para que al buscar un colegio salgan resultados y marcadores.
# Uso:  python aplicar_sedes_mapa.py            (trabaja sobre app.py de esta carpeta)
import io, os, re, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA = "/* ═══ Identidad EduTrack · Sedes y puntos de atención"

def salir(msg, code=1):
    print(msg)
    sys.exit(code)

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
if MARCA in s:
    salir("Este app.py YA tiene estos cambios. No se modificó nada.", 0)
NL = "\r\n" if "\r\n" in s else "\n"

def crlf(t):
    return t.replace("\r\n", "\n").replace("\n", NL)

def rep_en(ini, fin, old, new, count=1):
    global s
    if ini not in s:
        salir("ERROR: no encontré %r en tu app.py. No se modificó nada." % ini[:60])
    a = s.index(ini)
    b = s.index(fin, a)
    seg = s[a:b]
    n = seg.count(old)
    if n != count:
        salir("ERROR: punto de inserción no encontrado (%r: esperaba %d, hay %d). No se modificó nada." % (old[:70], count, n))
    s = s[:a] + seg.replace(old, new) + s[b:]

# ───────────────────────── 1) Login: quitar «Accesos institucionales»
patron = re.compile(r'[ \t]*<div class="sinai-pill-row">\s*<div class="pill-label">Accesos institucionales</div>\s*(?:<a [^>]*>[^<]*</a>\s*)+</div>\r?\n')
if len(patron.findall(s)) != 1:
    salir("ERROR: no encontré el bloque «Accesos institucionales» del login (una sola vez). No se modificó nada.")
s = patron.sub("", s, count=1)

# ───────────────────────── 2) Sedes y puntos de atención: identidad + fusión de fondos
CSS_SEDES = """/* ═══ Identidad EduTrack · Sedes y puntos de atención (una sola sección fluida) ═══ */
:root{--et-green:#00E676;--et-cyan:#00B8D4;--et-blue:#0020C2;--et-navy:#0B1B4D;--et-grad:linear-gradient(100deg,var(--et-green) 0%,var(--et-cyan) 32%,var(--et-blue) 100%);--et-cyan-soft:rgba(0,184,212,.10);--et-cyan-line:rgba(0,184,212,.38);--et-shadow:0 18px 50px rgba(0,32,194,.14)}
/* 1. Fusión: un solo fondo continuo entre el título y el buscador */
.sa{background:#F8F9FA}
.sa .sa-hero{background:transparent;min-height:0;padding:84px 20px 36px}
.sa .sa-find{background:transparent;padding:12px 20px 72px}
.sa .sa-res{background:transparent}
.sa .sa-hero h1{color:var(--et-navy)}
/* 2. Botón principal (Buscar una sede) y botón Ir */
.sa .sa-cta,.sa .sa-go{background:var(--et-grad);color:#fff;font-weight:700;letter-spacing:.01em;text-shadow:0 1px 2px rgba(0,20,90,.4);box-shadow:0 8px 20px rgba(0,32,194,.22);transition:transform .2s ease,box-shadow .2s ease,filter .2s ease}
.sa .sa-cta:hover,.sa .sa-go:hover{background:var(--et-grad);transform:translateY(-2px);box-shadow:0 14px 28px rgba(0,32,194,.30);filter:brightness(1.05) saturate(1.1)}
.sa .sa-cta:active,.sa .sa-go:active{transform:translateY(0) scale(.98)}
.sa .sa-cta:focus-visible,.sa .sa-go:focus-visible,.sa .sa-cat:focus-visible,.sa .sa-item:focus-visible,.sa .sa-gps:focus-visible{outline:3px solid rgba(0,184,212,.5);outline-offset:3px}
/* 5. Botón inferior (Ir) cuando aún no hay categoría: contorno cian minimalista */
.sa .sa-go.is-off{background:transparent;box-shadow:none;text-shadow:none;color:var(--et-cyan);border:1.5px solid var(--et-cyan-line)}
.sa .sa-go.is-off:hover{transform:none;filter:none;box-shadow:none;background:transparent}
/* 3. Tarjetas de categoría */
.sa .sa-cat{background:rgba(255,255,255,.7);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);border:1px solid var(--et-cyan-line);box-shadow:0 4px 14px rgba(0,32,194,.05)}
.sa .sa-cat svg{stroke:var(--et-blue);transition:stroke .2s}
.sa .sa-cat span{color:var(--et-navy)}
.sa .sa-cat:hover{transform:translateY(-3px);border-color:transparent;background:linear-gradient(#fff,#fff) padding-box,var(--et-grad) border-box;box-shadow:0 14px 32px rgba(0,32,194,.16)}
.sa .sa-cat:hover svg{stroke:var(--et-cyan)}
.sa .sa-cat.on{border:1.5px solid transparent;background:linear-gradient(#F1FBFF,#F1FBFF) padding-box,var(--et-grad) border-box;box-shadow:0 0 0 4px rgba(0,184,212,.14),0 14px 32px rgba(0,32,194,.14)}
.sa .sa-cat.on svg{stroke:var(--et-cyan)}
/* 4. Barra de búsqueda: sombra difuminada, lupa y ubicación en cian */
.sa .sa-search{background:#fff;border:1px solid transparent;box-shadow:0 8px 30px rgba(0,32,194,.10)}
.sa .sa-search:focus-within{border-color:transparent;box-shadow:0 10px 34px rgba(0,184,212,.22),0 0 0 3px rgba(0,184,212,.28)}
.sa .sa-search svg{stroke:var(--et-cyan)}
.sa .sa-gps{color:var(--et-cyan)}
.sa .sa-gps:hover{background:var(--et-cyan-soft)}
.sa .sa-gps.on{color:var(--et-blue);background:var(--et-cyan-soft)}
.sa .sa-search .sa-gps svg{stroke:currentColor}
/* Resultados */
.sa .sa-item{background:rgba(255,255,255,.78);border:1.5px solid var(--et-cyan-line)}
.sa .sa-item:hover{background:#fff;border-color:var(--et-cyan)}
.sa .sa-item.on{border-color:var(--et-cyan);background:#F1FBFF}
.sa .sa-how{background:var(--et-grad);text-shadow:0 1px 2px rgba(0,20,90,.35)}
.sa .sa-how:hover{background:var(--et-grad);filter:brightness(1.05)}
.sa .sa-tel{color:var(--et-blue)}
.sa .sa-mapbox{box-shadow:var(--et-shadow)}
@media(prefers-reduced-motion:reduce){.sa .sa-cta:hover,.sa .sa-go:hover,.sa .sa-cat:hover{transform:none}}
"""
rep_en('_PA_PAGE = r"""', '@app.route("/puntos-atencion")', "</style>", crlf(CSS_SEDES) + "</style>", 1)

# ───────────────────────── 3) Mapa de ubicaciones (/mapa): servidor
HELPER = '''def _like_q(t):
    """Patrón LIKE para buscar nombres. Si la búsqueda exacta no halló nada, se repite ignorando tildes (vocales como comodín)."""
    from flask import g as _g_mp
    t = (t or "").replace("%", "")
    if getattr(_g_mp, "_mp_plegar", False):
        t = "".join("_" if c in "aeiouáéíóúüAEIOUÁÉÍÓÚÜ" else c for c in t)
    return "%" + t + "%"


'''
if s.count('@app.route("/mapa/api/lugares")') != 1:
    salir("ERROR: no encontré la ruta /mapa/api/lugares. No se modificó nada.")
s = s.replace('@app.route("/mapa/api/lugares")', crlf(HELPER) + '@app.route("/mapa/api/lugares")', 1)

F0, F1 = "def mapa_api_lugares():", '_MP_PAGE = r"""'
rep_en(F0, F1, "    LIM = 120", "    from flask import g as _g_mp" + NL + "    LIM = 120", 1)
rep_en(F0, F1, 'Institucion.query.filter(Institucion.estado != "ARCHIVADA")',
       'Institucion.query.filter(or_(Institucion.estado.is_(None), Institucion.estado != "ARCHIVADA"))', 1)
rep_en(F0, F1, 'like = "%" + q.replace("%", "") + "%"', "like = _like_q(q)", 3)
rep_en(F0, F1, 'LeadCRM.colegio.ilike("%" + q.replace("%", "") + "%")', "LeadCRM.colegio.ilike(_like_q(q))", 1)
rep_en(F0, F1, 'PuntoAtencion.nombre.ilike("%" + q.replace("%", "") + "%")', "PuntoAtencion.nombre.ilike(_like_q(q))", 1)
rep_en(F0, F1, '    return jsonify({"ok": True, "items": out, "truncado": truncado})', crlf('''    if q and not out and not getattr(_g_mp, "_mp_plegar", False):
        _g_mp._mp_plegar = True
        try:
            return mapa_api_lugares()
        finally:
            _g_mp._mp_plegar = False
    dir_total = None
    if not out:
        try:
            dir_total = DirectorioMEN.query.count()
        except Exception:
            db.session.rollback()
    return jsonify({"ok": True, "items": out, "truncado": truncado, "dir_total": dir_total})'''), 1)

# ───────────────────────── 3b) Mapa de ubicaciones (/mapa): navegador (marcadores aproximados automáticos)
J0, J1 = '_MP_PAGE = r"""', '@app.route("/mapa")'
rep_en(J0, J1, 'if(ok){st.layer.clearLayers();st.markers={};}', 'if(ok){st.layer.clearLayers();st.markers={};st.ap={};}', 1)
rep_en(J0, J1, '  if(ok){st.items.forEach(function(it,i){if(it.lat!=null&&it.lng!=null)place(i,it.lat,it.lng,false);});}',
       '  if(ok){st.items.forEach(function(it,i){if(it.lat!=null&&it.lng!=null)place(i,it.lat,it.lng,false);});}' + NL + '  await autoAprox(ok);', 1)
rep_en(J0, J1, 'mk.on("click",function(){sel(i,false);});mk.addTo(st.layer);st.markers[i]=mk;',
       'mk.on("click",function(){sel(i,false);});mk.addTo(st.layer);st.markers[i]=mk;(st.ap=st.ap||{})[i]=!!aprox;', 1)
rep_en(J0, J1, 'if(st.markers[i]){if(abrir)sel(i,true);return true;}',
       'if(st.markers[i]&&!(st.ap&&st.ap[i])){if(abrir)sel(i,true);return true;}', 1)
rep_en(J0, J1, 'if(st.markers[i])continue;n++;', 'if(st.markers[i]&&!(st.ap&&st.ap[i]))continue;n++;', 1)
AUTO = r'''async function autoAprox(ok){
  if(!ok||!st.items.length)return;
  var grupos={},orden=[];
  st.items.forEach(function(it,i){
    if(it.lat!=null&&it.lng!=null)return;
    var dep=it.departamento||$("mp-dep").value,mun=it.municipio||"";
    if(!dep&&!mun)return;
    var k=norm(dep)+"|"+norm(mun);
    if(!grupos[k]){grupos[k]={dep:dep,mun:mun,idx:[]};orden.push(k);}
    grupos[k].idx.push(i);
  });
  var pts=[];
  for(var n=0;n<orden.length&&n<8;n++){
    var g=grupos[orden[n]];
    try{
      var r=await (await fetch("/mapa/api/geo?"+new URLSearchParams({dep:g.dep,mun:g.mun}),{credentials:"same-origin"})).json();
      if(!r.ok)continue;
      g.idx.slice(0,40).forEach(function(i){
        place(i,r.lat,r.lng,true);
        var s=$("mp-st-"+i);if(s){s.className="mp-st";s.textContent="≈ Aproximada ("+(g.mun?"municipio":"departamento")+")";}
      });
      pts.push([r.lat,r.lng]);
    }catch(e){}
  }
  if(pts.length&&st.map)st.map.fitBounds(pts,{padding:[50,50],maxZoom:13});
}
'''
rep_en(J0, J1, 'async function todos(){', crlf(AUTO) + 'async function todos(){', 1)
rep_en(J0, J1, "El mapa quedó centrado en la zona: puedes validar la ubicación igual.</div>';return;}",
       "El mapa quedó centrado en la zona: puedes validar la ubicación igual.'+(data.dir_total===0?'<br><b>El Directorio MEN está vacío</b> (no se ha importado), por eso solo se buscan tus clientes, sedes y prospectos.':'')+'</div>';return;}", 1)

tmp = SRC + ".nuevo"
with io.open(tmp, "w", encoding="utf-8", newline="") as fh:
    fh.write(s)
try:
    py_compile.compile(tmp, doraise=True)
except Exception as e:
    os.remove(tmp)
    salir("ERROR: el resultado no compila, no se modificó tu app.py: %s" % e)
shutil.copyfile(SRC, SRC + ".respaldo")
os.replace(tmp, SRC)
print("LISTO. Login sin «Accesos institucionales», sedes con la nueva identidad y mapa corregido. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

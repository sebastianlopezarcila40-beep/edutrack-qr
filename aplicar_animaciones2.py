# Ajuste de las animaciones por scroll: con "reducir movimiento" activado en el sistema, en vez de apagarlas
# por completo se muestra solo un fundido suave (sin desplazamiento).
# Uso:  python aplicar_animaciones2.py            (trabaja sobre app.py de esta carpeta)
import io, os, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA = "html.rv-fade .rv{"
CSS_VIEJO = "@media(prefers-reduced-motion:reduce){.rv,.sa-hero>*,.sa-list .sa-item{animation:none!important;transition:none!important;opacity:1!important;transform:none!important}}"
CSS_NUEVO = ("html.rv-fade .rv{transform:none!important;transition:opacity .6s ease var(--rv-d,0ms)!important}\n"
             "@media(prefers-reduced-motion:reduce){.sa-hero>*,.sa-list .sa-item{animation:none!important}}")
JS_VIEJO = 'if(window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches)return;'
JS_NUEVO = 'if(window.matchMedia&&matchMedia("(prefers-reduced-motion: reduce)").matches)document.documentElement.className+=" rv-fade";'

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
if MARCA in s:
    print("Este app.py YA tiene este ajuste. No se modificó nada.")
    sys.exit(0)
if "_REVEAL_SNIPPET = r'''" not in s:
    print("ERROR: este app.py aún no tiene las animaciones. Ejecuta primero: python aplicar_animaciones.py")
    sys.exit(1)
if s.count(CSS_VIEJO) != 1 or s.count(JS_VIEJO) != 1:
    print("ERROR: no encontré los puntos de inserción de las animaciones. No se modificó nada.")
    sys.exit(1)
nl = "\r\n" if "\r\n" in s else "\n"
s = s.replace(CSS_VIEJO, CSS_NUEVO.replace("\n", nl), 1).replace(JS_VIEJO, JS_NUEVO, 1)
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
print("LISTO. Con «reducir movimiento» ahora se ve un fundido suave. Líneas:", len(s.splitlines()))

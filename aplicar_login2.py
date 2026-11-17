# Portal de acceso: vuelve al diseño de tarjetas separadas (carrusel | acceso), con la identidad actual y más compacto.
# Uso:  python aplicar_login2.py            (trabaja sobre app.py de esta carpeta)
import io, os, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA1 = "/* Portal de acceso: dos áreas separadas (marca | acceso) */"
FIN1 = "background:linear-gradient(90deg,#0020C2,#00C2FF,#00FF66)}}"
MARCA2 = "/* Portal de acceso: tarjetas separadas, compacto */"
ANCLA = ".sinai-card{border:1px solid #E2E8F0!important;border-radius:16px!important;box-shadow:0 20px 50px rgba(0,0,40,.28)!important}"

CSS = """/* Portal de acceso: tarjetas separadas, compacto */
.sinai-portal{background:radial-gradient(900px 520px at 15% 0%,#dbe8ff,transparent 60%),radial-gradient(700px 420px at 100% 8%,#e3eeff,transparent 60%),#f5f7fb!important;padding:18px 20px 26px!important}
.sinai-grid{max-width:1020px!important;gap:24px!important;grid-template-columns:1.3fr .9fr!important;align-items:center}
.sinai-carousel{--car-logo:360px!important;border-radius:26px!important;height:380px!important;box-shadow:0 22px 56px rgba(0,32,194,.26)!important;border:1px solid rgba(255,255,255,.55)!important}
.sinai-brand-logo{max-height:200px!important}
.sinai-brand-slogan{font-size:clamp(13px,1.3vw,17px)!important}
.sinai-nav{width:38px;height:38px;font-size:24px}
.sinai-access h1{color:#0B1B4D!important;font-size:23px!important}
.sinai-sub{color:#64748b!important;font-size:13px!important;margin-bottom:12px!important}
.sinai-ver,.sinai-ver b{color:#64748b!important}
.sinai-card{background:#fff!important;border:1px solid #E2E8F0!important;border-radius:16px!important;padding:20px 20px!important;box-shadow:0 10px 30px rgba(0,32,194,.10)!important}
@media(max-width:900px){.sinai-grid{grid-template-columns:1fr!important}.sinai-carousel{height:320px!important;border-radius:22px!important}}"""

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
if MARCA2 in s:
    print("Este app.py YA tiene este cambio. No se modificó nada.")
    sys.exit(0)
nl = "\r\n" if "\r\n" in s else "\n"
nuevo = CSS.replace("\n", nl)
if MARCA1 in s:
    a = s.index(MARCA1)
    b = s.find(FIN1, a)
    if b < 0:
        print("ERROR: no pude delimitar el bloque anterior del login. No se modificó nada.")
        sys.exit(1)
    s = s[:a] + nuevo + s[b + len(FIN1):]
elif s.count(ANCLA) == 1:
    s = s.replace(ANCLA, ANCLA + nl + nuevo)
else:
    print("ERROR: no encontré el punto de inserción en tu app.py. No se modificó nada.")
    sys.exit(1)
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
print("LISTO. Portal de acceso con tarjetas separadas y más compacto. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

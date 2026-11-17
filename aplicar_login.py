# Separa en dos áreas el portal de acceso (marca | acceso) con formato corporativo.
# Uso:  python aplicar_login.py            (trabaja sobre app.py de esta carpeta)
import io, os, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA = "/* Portal de acceso: dos áreas separadas (marca | acceso) */"
ANCLA = ".sinai-card{border:1px solid #E2E8F0!important;border-radius:16px!important;box-shadow:0 20px 50px rgba(0,0,40,.28)!important}"

CSS = """/* Portal de acceso: dos áreas separadas (marca | acceso) */
.sinai-grid{gap:0!important;max-width:1180px!important;border-radius:26px;overflow:hidden;background:#fff;box-shadow:0 30px 80px rgba(0,0,40,.42)}
.sinai-carousel{border-radius:0!important;box-shadow:none!important;border:0!important;height:auto!important;min-height:580px}
.sinai-carousel .sinai-viewport{position:absolute!important;inset:0;height:auto!important}
.sinai-access{position:relative;background:#fff!important;padding:38px 34px 24px;display:flex;flex-direction:column;justify-content:center}
.sinai-access::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,#0020C2,#00C2FF,#00FF66)}
.sinai-access h1{color:#0B1B4D!important}
.sinai-sub{color:#475569!important}
.sinai-ver,.sinai-ver b{color:#64748b!important}
.sinai-card{box-shadow:none!important;border:1px solid #E2E8F0!important;background:#fff!important}
@media(max-width:900px){.sinai-grid{border-radius:20px}.sinai-carousel{min-height:400px}.sinai-access{padding:30px 22px 20px}.sinai-access::before{right:0;bottom:auto;width:auto;height:5px;background:linear-gradient(90deg,#0020C2,#00C2FF,#00FF66)}}"""

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
if MARCA in s:
    print("Este app.py YA tiene este cambio. No se modificó nada.")
    sys.exit(0)
if s.count(ANCLA) != 1:
    print("ERROR: no encontré el punto de inserción en tu app.py (¿está completo y con los cambios anteriores?). No se modificó nada.")
    sys.exit(1)
nl = "\r\n" if "\r\n" in s else "\n"
s = s.replace(ANCLA, ANCLA + nl + CSS.replace("\n", nl))
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
print("LISTO. Portal de acceso separado en dos áreas. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

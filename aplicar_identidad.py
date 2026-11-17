# Identidad EduTrack del portal de acceso: degradado verde -> cian -> azul rey, con variables CSS centralizadas.
# Uso:  python aplicar_identidad.py            (trabaja sobre app.py de esta carpeta)
import io, os, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
MARCA = "/* Identidad EduTrack: degradado verde → cian → azul rey"
ANCLA = ".c48-lg .lk a{color:#0020C2!important}"

CSS = """/* Identidad EduTrack: degradado verde → cian → azul rey (paleta centralizada en :root) */
:root{
  --et-green:#00E676;
  --et-cyan:#00B8D4;
  --et-blue:#0020C2;
  --et-navy:#0B1B4D;
  --et-grad:linear-gradient(100deg,var(--et-green) 0%,var(--et-cyan) 32%,var(--et-blue) 100%);
  --et-bg:#FFFFFF;
  --et-bg-soft:#FAFBFD;
  --et-cyan-soft:rgba(0,184,212,.10);
  --et-cyan-line:rgba(0,184,212,.38);
  --et-shadow:0 18px 50px rgba(0,32,194,.14);
}
/* 1. Fondo limpio */
.login-page,.lp-content{background:var(--et-bg-soft)!important}
.sinai-portal{background:var(--et-bg)!important}
/* 2. Carrusel en cristal (glassmorphism) con sombra de tinte azul */
.sinai-portal .sinai-carousel{background:linear-gradient(135deg,rgba(0,230,118,.14),rgba(0,184,212,.16) 50%,rgba(0,32,194,.14))!important;-webkit-backdrop-filter:blur(18px) saturate(160%);backdrop-filter:blur(18px) saturate(160%);border:1px solid rgba(255,255,255,.75)!important;box-shadow:var(--et-shadow),inset 0 1px 0 rgba(255,255,255,.8)!important}
.sinai-portal .sinai-brand-tile{background:rgba(255,255,255,.72)!important;-webkit-backdrop-filter:blur(14px)!important;backdrop-filter:blur(14px)!important;border:1px solid rgba(255,255,255,.85)!important;box-shadow:0 14px 36px rgba(0,32,194,.16)!important}
.sinai-portal .sinai-brand-slogan{background:var(--et-grad)!important;color:#fff!important;font-weight:700!important;border:0!important;text-shadow:0 1px 2px rgba(0,20,90,.45);box-shadow:0 8px 22px rgba(0,32,194,.18)!important}
.sinai-portal .sinai-brand-sub{color:var(--et-blue)!important}
.sinai-portal .sinai-nav{background:rgba(255,255,255,.72)!important;color:var(--et-blue)!important;border:1px solid var(--et-cyan-line)!important;box-shadow:none!important;transition:background .2s,color .2s}
.sinai-portal .sinai-nav:hover{background:var(--et-grad)!important;color:#fff!important}
.sinai-portal .sinai-card{background:#fff!important;border:1px solid rgba(0,32,194,.08)!important;box-shadow:var(--et-shadow)!important}
.sinai-portal .sinai-access h1{color:var(--et-navy)!important}
/* 5a. Indicadores del carrusel */
.sinai-portal .sinai-dots{background:rgba(255,255,255,.7)!important;border:1px solid var(--et-cyan-line)!important;box-shadow:none!important}
.sinai-portal .sinai-dot{background:var(--et-cyan-line)!important}
.sinai-portal .sinai-dot.is-on{background:var(--et-grad)!important}
/* 3. Botón de iniciar sesión */
.sinai-portal .sinai-btn{background:var(--et-grad)!important;color:#fff!important;font-weight:700!important;text-shadow:0 1px 2px rgba(0,20,90,.4);box-shadow:0 8px 20px rgba(0,32,194,.22);transition:transform .2s ease,box-shadow .2s ease,filter .2s ease}
.sinai-portal .sinai-btn:hover{background:var(--et-grad)!important;transform:translateY(-2px);box-shadow:0 14px 28px rgba(0,32,194,.30);filter:brightness(1.05) saturate(1.1)}
.sinai-portal .sinai-btn:active{transform:translateY(0) scale(.99)}
/* Campos: foco en cian */
.sinai-portal .sinai-form input:focus,.sinai-portal .sinai-form select:focus{border-color:var(--et-cyan)!important;box-shadow:0 0 0 4px rgba(0,184,212,.18)!important}
/* 4. Banner de convocatoria: cian translúcido, texto oscuro y línea de acento */
.sinai-portal .c48-lg summary{background:var(--et-cyan-soft)!important;color:var(--et-navy)!important;border:1px solid var(--et-cyan-line)!important;border-left:5px solid var(--et-cyan)!important;border-radius:12px!important;box-shadow:none!important;-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px);transition:background .2s}
.sinai-portal .c48-lg summary:hover{background:rgba(0,184,212,.17)!important}
/* 5b. Accesos institucionales y de familias */
.sinai-portal .sinai-portals a,.sinai-portal .sinai-pill{background:transparent!important;color:var(--et-blue)!important;border:1px solid var(--et-cyan-line)!important;font-weight:600;transition:background .2s,border-color .2s,color .2s}
.sinai-portal .sinai-portals a:hover,.sinai-portal .sinai-pill:hover{background:var(--et-cyan-soft)!important;border-color:var(--et-cyan)!important;color:var(--et-blue)!important}
.sinai-portal .sinai-pill-row .pill-label{color:var(--et-navy)!important;opacity:.7}
@media(prefers-reduced-motion:reduce){.sinai-portal .sinai-btn,.sinai-portal .sinai-btn:hover{transition:none;transform:none}}"""

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
print("LISTO. Identidad verde-cian-azul aplicada al portal de acceso. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

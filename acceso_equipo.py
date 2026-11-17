# -*- coding: utf-8 -*-
"""
Módulo público "Acceso del equipo" para PROCSIS.

Rutas públicas (no requieren sesión):
    GET  /acceso-equipo                  página con cámara (buscar usuario / solicitar alta)
    POST /acceso-equipo/buscar           ¿existe el usuario? (JSON)
    POST /acceso-equipo/enviar           registra solicitud con foto de cara (JSON)
    GET  /acceso-equipo/estado/<token>   estado de una solicitud

Rutas de Gerencia (requieren sesión Gerente / Superadmin / Administrador):
    GET  /gerencia/solicitudes-acceso
    GET  /gerencia/solicitudes-acceso/foto/<id>
    POST /gerencia/solicitudes-acceso/<id>/aprobar
    POST /gerencia/solicitudes-acceso/<id>/rechazar

Principios:
  * NADIE obtiene acceso solo por llenar el formulario: Gerencia revisa la foto y aprueba.
  * La foto se guarda en la BASE DE DATOS (Vercel no tiene disco persistente).
  * Se exige autorización expresa de tratamiento de datos biométricos (Ley 1581 de 2012).

Uso en app.py (al final, antes de `if __name__ == "__main__":`):
    from acceso_equipo import registrar_acceso_equipo
    registrar_acceso_equipo(app, db, Usuario, crear_hash,
                            registrar_auditoria=registrar_auditoria,
                            rate_limit=_rate_limit, areas=AREAS_PROCSIS)
"""
import base64
import hashlib
import html
import os
import re
import secrets
import unicodedata
from datetime import datetime, timedelta, timezone

from flask import jsonify, make_response, request, session

BOGOTA = timezone(timedelta(hours=-5))
ROLES_APRUEBAN = ("Gerente", "Superadmin", "Administrador")
ROLES_EXTRA = ("Comercial", "Cobranza", "Contabilidad", "Soporte")
MAX_FOTO_BYTES = 3 * 1024 * 1024


def _ahora():
    return datetime.now(BOGOTA).strftime("%Y-%m-%d %H:%M:%S")


def _esc(x):
    return html.escape(str(x if x is not None else ""), quote=True)


def _limpiar(x, n=160):
    return re.sub(r"\s+", " ", str(x or "")).strip()[:n]


def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _mascara_correo(c):
    c = c or ""
    if "@" not in c:
        return ""
    a, b = c.split("@", 1)
    return (a[:2] + "•••@" + b) if len(a) > 2 else ("•••@" + b)


def _mascara_nombre(n):
    partes = (n or "").split()
    return " ".join(p[0] + "•" * max(1, len(p) - 1) if i else p for i, p in enumerate(partes)) if partes else ""


# --------------------------------------------------------------------------------------
# Procesamiento de foto (servidor)
# --------------------------------------------------------------------------------------
def procesar_foto(data_url):
    """Valida y normaliza la foto. Devuelve (jpeg_base64, sha256, error)."""
    try:
        from PIL import Image, ImageStat
        from io import BytesIO
    except Exception:  # pragma: no cover
        return None, None, "El servidor no tiene Pillow instalado (agregue Pillow a requirements.txt)."
    raw_txt = str(data_url or "")
    if "," in raw_txt:
        raw_txt = raw_txt.split(",", 1)[1]
    if not raw_txt or len(raw_txt) > MAX_FOTO_BYTES * 4 // 3 + 100:
        return None, None, "No se recibió una foto válida."
    try:
        raw = base64.b64decode(raw_txt, validate=False)
        im = Image.open(BytesIO(raw))
        im.load()
    except Exception:
        return None, None, "La foto no se pudo leer. Tómela de nuevo."
    if im.mode != "RGB":
        im = im.convert("RGB")
    w, h = im.size
    if min(w, h) < 200:
        return None, None, "La foto es muy pequeña. Acérquese a la cámara y repita."
    lado = min(w, h)
    im = im.crop(((w - lado) // 2, (h - lado) // 2, (w - lado) // 2 + lado, (h - lado) // 2 + lado))
    im = im.resize((480, 480))
    st = ImageStat.Stat(im.convert("L"))
    if st.mean[0] < 35 or st.mean[0] > 235 or st.stddev[0] < 18:
        return None, None, "La foto está muy oscura, muy clara o sin detalle. Busque buena luz y repita."
    out = BytesIO()
    im.save(out, "JPEG", quality=82, optimize=True)
    data = out.getvalue()
    return base64.b64encode(data).decode("ascii"), hashlib.sha256(data).hexdigest(), None


# --------------------------------------------------------------------------------------
# Almacenamiento (SQLAlchemy). Aislado para poder probar las rutas sin base de datos.
# --------------------------------------------------------------------------------------
class SqlStore:
    def __init__(self, db, Usuario):
        self.db, self.Usuario = db, Usuario
        cls = getattr(db.Model, "_ae_solicitud_cls", None)
        if cls is None:
            class SolicitudAccesoEquipo(db.Model):
                __tablename__ = "solicitudes_acceso_equipo"
                id = db.Column(db.Integer, primary_key=True)
                token = db.Column(db.String(40), unique=True, index=True, nullable=False)
                tipo = db.Column(db.String(20), default="ALTA")          # ALTA | REVALIDACION
                estado = db.Column(db.String(20), default="PENDIENTE", index=True)
                usuario_id = db.Column(db.Integer, nullable=True)         # si ya existía
                usuario_nombre = db.Column(db.String(80), default="")     # usuario resultante
                nombre = db.Column(db.String(160), default="")
                cedula = db.Column(db.String(30), default="", index=True)
                correo = db.Column(db.String(160), default="")
                telefono = db.Column(db.String(40), default="")
                area = db.Column(db.String(60), default="")
                cargo = db.Column(db.String(120), default="")
                jefe = db.Column(db.String(120), default="")
                foto_b64 = db.Column(db.Text, default="")
                foto_hash = db.Column(db.String(64), default="", index=True)
                foto_origen = db.Column(db.String(20), default="camara")  # camara | archivo
                consentimiento = db.Column(db.Boolean, default=False)
                consent_fecha = db.Column(db.String(30), default="")
                ip = db.Column(db.String(64), default="", index=True)
                user_agent = db.Column(db.String(300), default="")
                creado_en = db.Column(db.String(30), default="", index=True)
                resuelto_por = db.Column(db.String(80), default="")
                resuelto_en = db.Column(db.String(30), default="")
                motivo = db.Column(db.String(300), default="")
            db.Model._ae_solicitud_cls = SolicitudAccesoEquipo
            cls = SolicitudAccesoEquipo
        self.S = cls

    # -- infraestructura
    def ensure_tables(self):
        self.S.__table__.create(bind=self.db.engine, checkfirst=True)

    # -- usuarios
    def _u_dict(self, u):
        return {"id": u.id, "usuario": u.usuario, "nombre": u.nombre_completo or "", "correo": u.correo or "",
                "rol": u.rol or "", "activo": getattr(u, "activo", True) is not False,
                "bio": bool(getattr(u, "biometria_registrada", False))}

    def buscar_usuario(self, ident):
        from sqlalchemy import func, or_
        x = ident.lower()
        u = (self.Usuario.query.filter(self.Usuario.institucion_id.is_(None))
             .filter(or_(func.lower(self.Usuario.usuario) == x, func.lower(self.Usuario.correo) == x)).first())
        return self._u_dict(u) if u else None

    def usuario_por_id(self, uid):
        u = self.Usuario.query.get(uid)
        return self._u_dict(u) if u else None

    def usuario_existe(self, nombre):
        from sqlalchemy import func
        return self.Usuario.query.filter(func.lower(self.Usuario.usuario) == nombre.lower()).first() is not None

    def hay_gerente(self):
        return self.Usuario.query.filter(self.Usuario.rol.in_(ROLES_APRUEBAN)).first() is not None

    def crear_usuario(self, d):
        u = self.Usuario(usuario=d["usuario"], password=d["password_hash"], rol=d["rol"], correo=d["correo"],
                         nombre_completo=d["nombre"], institucion_id=None, password_temporal=True,
                         departamento=d.get("area", ""), jefe_usuario=d.get("jefe", ""),
                         biometria_path=d.get("biometria_path", ""), biometria_registrada=True,
                         biometria_fecha=_ahora())
        if hasattr(u, "activo"):
            u.activo = True
        self.db.session.add(u)
        self.db.session.flush()
        return u.id

    def reset_usuario(self, uid, password_hash, biometria_path):
        u = self.Usuario.query.get(uid)
        u.password = password_hash
        u.password_temporal = True
        if hasattr(u, "activo"):
            u.activo = True
        u.biometria_path, u.biometria_registrada, u.biometria_fecha = biometria_path, True, _ahora()
        if hasattr(u, "session_token"):
            u.session_token = ""   # cierra sesiones anteriores
        return u.usuario

    # -- solicitudes
    def contar_ip_reciente(self, ip, desde):
        return self.S.query.filter(self.S.ip == ip, self.S.creado_en >= desde).count()

    def pendiente_duplicada(self, cedula, correo):
        q = self.S.query.filter(self.S.estado == "PENDIENTE")
        for r in q.all():
            if (cedula and r.cedula == cedula) or (correo and r.correo.lower() == correo.lower()):
                return True
        return False

    def foto_hash_repetida(self, h):
        return self.S.query.filter(self.S.foto_hash == h).count()

    def crear_solicitud(self, d):
        r = self.S(**d)
        self.db.session.add(r)
        self.db.session.commit()
        return r.id

    def _s_dict(self, r, con_foto=False):
        d = {c.name: getattr(r, c.name) for c in self.S.__table__.columns}
        if not con_foto:
            d.pop("foto_b64", None)
        return d

    def listar(self, estado):
        q = self.S.query
        if estado != "TODAS":
            q = q.filter(self.S.estado == estado)
        return [self._s_dict(r) for r in q.order_by(self.S.id.desc()).limit(200).all()]

    def obtener(self, sid, con_foto=False):
        r = self.S.query.get(sid)
        return self._s_dict(r, con_foto) if r else None

    def por_token(self, token):
        r = self.S.query.filter_by(token=token).first()
        return self._s_dict(r) if r else None

    def resolver(self, sid, estado, por, motivo="", usuario_nombre="", usuario_id=None):
        r = self.S.query.get(sid)
        r.estado, r.resuelto_por, r.resuelto_en, r.motivo = estado, por, _ahora(), motivo[:300]
        if usuario_nombre:
            r.usuario_nombre = usuario_nombre
        if usuario_id:
            r.usuario_id = usuario_id
        self.db.session.commit()

    def commit(self):
        self.db.session.commit()

    def rollback(self):
        try:
            self.db.session.rollback()
        except Exception:
            pass


# --------------------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------------------
CSS = """
:root{--bg:#f4f6fa;--card:#fff;--ink:#0f172a;--mut:#64748b;--brand:#0B2D57;--acc:#2563eb;--ok:#15803d;--bad:#b91c1c;--line:#e2e8f0}
@media(prefers-color-scheme:dark){:root{--bg:#0b1220;--card:#111a2e;--ink:#e6edf7;--mut:#94a3b8;--brand:#60a5fa;--acc:#3b82f6;--ok:#4ade80;--bad:#f87171;--line:#23304a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 -apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",system-ui,sans-serif}
.wrap{max-width:560px;margin:0 auto;padding:20px 16px 48px}
.top{display:flex;align-items:center;gap:10px;margin:8px 0 18px}.logo{width:36px;height:36px;border-radius:10px;background:var(--brand);color:#fff;display:grid;place-items:center;font-weight:700}
.top b{font-size:17px}.top span{color:var(--mut);font-size:13px;display:block;margin-top:-2px}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:20px;margin-bottom:14px}
h1{font-size:20px;margin:0 0 4px}h2{font-size:16px;margin:0 0 8px}p{margin:6px 0}.mut{color:var(--mut);font-size:13px}
label{display:block;font-weight:600;font-size:13px;margin:12px 0 4px}
input,select,textarea{width:100%;padding:11px 12px;border:1px solid var(--line);border-radius:10px;background:var(--bg);color:var(--ink);font:inherit}
input:focus,select:focus{outline:2px solid var(--acc);outline-offset:1px}
.btn{display:inline-block;border:0;border-radius:10px;padding:12px 18px;background:var(--acc);color:#fff;font:inherit;font-weight:600;cursor:pointer;width:100%;margin-top:14px;text-align:center;text-decoration:none}
.btn.sec{background:transparent;color:var(--ink);border:1px solid var(--line)}.btn.ok{background:var(--ok)}.btn.bad{background:var(--bad)}.btn:disabled{opacity:.45;cursor:not-allowed}
.btn.sm{width:auto;padding:8px 12px;margin:0;font-size:13px}
.msg{padding:10px 12px;border-radius:10px;font-size:14px;margin-top:12px}.msg.err{background:#fee2e2;color:#7f1d1d}.msg.ok{background:#dcfce7;color:#14532d}.msg.info{background:#dbeafe;color:#1e3a8a}
@media(prefers-color-scheme:dark){.msg.err{background:#3b1212;color:#fecaca}.msg.ok{background:#0f2e1a;color:#bbf7d0}.msg.info{background:#10244a;color:#bfdbfe}}
.cam{position:relative;width:min(100%,340px);aspect-ratio:1;margin:12px auto;border-radius:50%;overflow:hidden;background:#000;border:4px solid var(--line);transition:border-color .2s}
.cam.ok{border-color:var(--ok)}.cam video,.cam img{width:100%;height:100%;object-fit:cover;transform:scaleX(-1)}.cam img{transform:none}
.hint{text-align:center;font-weight:600;min-height:24px}.hide{display:none!important}
.chk{display:flex;gap:10px;align-items:flex-start;font-size:13px;margin-top:14px}.chk input{width:auto;margin-top:3px}
.steps{display:flex;gap:6px;margin-bottom:14px}.steps i{flex:1;height:4px;border-radius:4px;background:var(--line)}.steps i.on{background:var(--acc)}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
.pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:12px;font-weight:600;background:var(--line)}
.pill.PENDIENTE{background:#fef3c7;color:#92400e}.pill.APROBADA{background:#dcfce7;color:#14532d}.pill.RECHAZADA{background:#fee2e2;color:#7f1d1d}
.hp{position:absolute;left:-9999px;opacity:0}
"""

CONSENTIMIENTO = (
    "Autorizo de manera expresa, previa e informada a PROCSIS para recolectar y tratar mi fotografía de rostro (dato "
    "biométrico, considerado dato sensible) y mis datos personales, con la única finalidad de verificar mi identidad y "
    "gestionar mis credenciales de acceso al sistema, conforme a la Ley 1581 de 2012 y el Decreto 1377 de 2013. "
    "Sé que puedo no autorizar el tratamiento de datos sensibles, y que puedo conocer, actualizar, rectificar o solicitar "
    "la supresión de mis datos escribiendo a Gerencia."
)


def _layout(titulo, cuerpo, script=""):
    return ("<!doctype html><html lang='es'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<meta name='robots' content='noindex,nofollow'>"
            "<title>%s · PROCSIS</title><style>%s</style></head><body><div class='wrap'>"
            "<div class='top'><div class='logo'>P</div><div><b>PROCSIS</b><span>Acceso del equipo</span></div></div>"
            "%s</div>%s</body></html>" % (_esc(titulo), CSS, cuerpo, script))


JS_PUBLICO = r"""
<script>
(function(){
const CSRF="__CSRF__", AREAS=__AREAS__;
const $=id=>document.getElementById(id);
let ctx={usuario:null}, foto=null, origen="camara", stream=null, timer=null;
function show(n){["s1","s2","s3","s4"].forEach((s,i)=>{$(s).classList.toggle("hide",i!==n-1)});
  document.querySelectorAll(".steps i").forEach((e,i)=>e.classList.toggle("on",i<n));window.scrollTo(0,0)}
function err(id,t){const e=$(id);e.textContent=t||"";e.classList.toggle("hide",!t)}
async function post(url,data){const r=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json","X-CSRF":CSRF},body:JSON.stringify(data)});
  let j={};try{j=await r.json()}catch(e){}if(!r.ok&&!j.error)j.error="Error "+r.status;return j}

/* ---------- paso 1: buscar ---------- */
$("f1").addEventListener("submit",async e=>{e.preventDefault();err("e1","");
  const v=$("ident").value.trim();if(v.length<3){err("e1","Escriba su usuario o correo.");return}
  $("b1").disabled=true;const j=await post("/acceso-equipo/buscar",{ident:v,hp:$("hp").value});$("b1").disabled=false;
  if(j.error){err("e1",j.error);return}
  if(j.existe){ctx.usuario=j.usuario;$("who").innerHTML="Encontramos la cuenta <b>"+j.usuario.usuario+"</b>"+(j.usuario.nombre?" ("+j.usuario.nombre+")":"")+(j.usuario.correo?"<br>Correo registrado: "+j.usuario.correo:"")+".";
    $("t2").textContent="Verifica tu identidad";$("alta").classList.add("hide");$("cons").classList.remove("hide");show(2)}
  else{ctx.usuario=null;$("who").textContent="No encontramos ese usuario. Puedes solicitar tu acceso: Gerencia revisará tu solicitud.";
    $("t2").textContent="Solicitar acceso";$("alta").classList.remove("hide");$("cons").classList.remove("hide");
    if(v.includes("@"))$("correo").value=v;show(2)}});

/* ---------- paso 2: datos ---------- */
$("f2").addEventListener("submit",e=>{e.preventDefault();err("e2","");
  if(!ctx.usuario){for(const k of ["nombre","cedula","correo","area"]){if(!$(k).value.trim()){err("e2","Complete todos los campos obligatorios.");return}}
    if(!/^\S+@\S+\.\S+$/.test($("correo").value.trim())){err("e2","Escriba un correo válido.");return}}
  if(!$("consent").checked){err("e2","Debe aceptar el tratamiento de datos para continuar.");return}
  show(3);abrirCamara()});

/* ---------- paso 3: cámara y rostro ---------- */
const video=$("video"),cam=$("cam"),hint=$("hint"),snap=$("snap"),prev=$("prev");
let okFrames=0,mov=0,last=null,detector=null;
if("FaceDetector" in window){try{detector=new FaceDetector({fastMode:true,maxDetectedFaces:2})}catch(e){}}
const mini=document.createElement("canvas");mini.width=mini.height=48;const mc=mini.getContext("2d",{willReadFrequently:true});
async function abrirCamara(){
  foto=null;origen="camara";prev.classList.add("hide");video.classList.remove("hide");$("b3").disabled=true;$("b3r").classList.add("hide");cam.classList.remove("ok");
  if(!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia){hint.textContent="Su navegador no permite cámara en vivo. Use el botón «Tomar foto con el teléfono».";return}
  try{stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:"user",width:{ideal:720},height:{ideal:720}},audio:false});
    video.srcObject=stream;await video.play();okFrames=0;mov=0;last=null;hint.textContent="Centre su rostro en el círculo y mueva un poco la cabeza";
    clearInterval(timer);timer=setInterval(analizar,250)}
  catch(e){hint.textContent="No se pudo abrir la cámara. Permita el acceso o use «Tomar foto con el teléfono».";}}
async function analizar(){
  if(!video.videoWidth)return;const w=video.videoWidth,h=video.videoHeight,l=Math.min(w,h);
  mc.drawImage(video,(w-l)/2,(h-l)/2,l,l,0,0,48,48);const d=mc.getImageData(0,0,48,48).data;let s=0,s2=0,n=d.length/4,diff=0;
  const cur=new Uint8Array(n);for(let i=0;i<n;i++){const g=(d[i*4]+d[i*4+1]+d[i*4+2])/3;cur[i]=g;s+=g;s2+=g*g}
  const m=s/n,sd=Math.sqrt(Math.max(0,s2/n-m*m));
  if(last){for(let i=0;i<n;i++)diff+=Math.abs(cur[i]-last[i]);diff/=n;if(diff>1.6)mov++}last=cur;
  let rostro=true;
  if(detector){try{const f=await detector.detect(video);rostro=f.length===1&&(f[0].boundingBox.width*f[0].boundingBox.height)/(w*h)>0.08;
      if(f.length>1){hint.textContent="Debe aparecer solo una persona";okFrames=0;return}}catch(e){rostro=true}}
  const luz=m>35&&m<230&&sd>12;
  if(!luz){hint.textContent=m<=35?"Hay muy poca luz":"Demasiada luz o sin detalle";okFrames=Math.max(0,okFrames-1);cam.classList.remove("ok");$("b3").disabled=true;return}
  if(!rostro){hint.textContent="No vemos su rostro. Acérquese y mire a la cámara";okFrames=Math.max(0,okFrames-2);cam.classList.remove("ok");$("b3").disabled=true;return}
  okFrames++;
  if(okFrames>=6&&mov>=2){hint.textContent="Listo. Pulse «Tomar foto»";cam.classList.add("ok");$("b3").disabled=false}
  else{hint.textContent=mov<2?"Mueva ligeramente la cabeza":"Mantenga el rostro centrado…";$("b3").disabled=true}}
function recortar(src,sw,sh){const c=document.createElement("canvas");c.width=c.height=480;const l=Math.min(sw,sh);
  c.getContext("2d").drawImage(src,(sw-l)/2,(sh-l)/2,l,l,0,0,480,480);return c.toDataURL("image/jpeg",0.85)}
$("b3").addEventListener("click",()=>{foto=recortar(video,video.videoWidth,video.videoHeight);origen="camara";fijar()});
$("fileCam").addEventListener("change",e=>{const f=e.target.files[0];if(!f)return;const im=new Image();
  im.onload=()=>{foto=recortar(im,im.width,im.height);origen="archivo";fijar();URL.revokeObjectURL(im.src)};im.src=URL.createObjectURL(f)});
function fijar(){clearInterval(timer);if(stream){stream.getTracks().forEach(t=>t.stop());stream=null}
  prev.src=foto;prev.classList.remove("hide");video.classList.add("hide");cam.classList.add("ok");hint.textContent="¿Se ve bien su rostro?";$("b3").disabled=true;$("b3r").classList.remove("hide");$("b3s").classList.remove("hide")}
$("b3r").addEventListener("click",()=>{$("b3s").classList.add("hide");abrirCamara()});
$("b3s").addEventListener("click",async()=>{err("e3","");if(!foto)return;$("b3s").disabled=true;$("b3s").textContent="Enviando…";
  const body={foto,origen,consent:$("consent").checked,hp:$("hp").value,usuario_id:ctx.usuario?ctx.usuario.id:null,
    nombre:$("nombre").value,cedula:$("cedula").value,correo:$("correo").value,telefono:$("telefono").value,
    area:$("area").value,cargo:$("cargo").value,jefe:$("jefe").value,ident:$("ident").value};
  const j=await post("/acceso-equipo/enviar",body);$("b3s").disabled=false;$("b3s").textContent="Enviar solicitud";
  if(j.error){err("e3",j.error);return}
  $("codigo").textContent=j.token;$("lnk").href="/acceso-equipo/estado/"+j.token;show(4)});
$("volver").addEventListener("click",()=>{if(stream){stream.getTracks().forEach(t=>t.stop())}clearInterval(timer);show(2)});
AREAS.forEach(a=>{const o=document.createElement("option");o.value=a[0];o.textContent=a[1];$("area").appendChild(o)});
})();
</script>
"""


def _pagina_publica(csrf, areas):
    import json
    cuerpo = """
<div class='steps'><i class='on'></i><i></i><i></i><i></i></div>

<section id='s1' class='card'>
  <h1>Validar o solicitar acceso</h1>
  <p class='mut'>Escriba su usuario o correo corporativo. Si ya tenía cuenta la verificamos; si no, puede solicitar una.</p>
  <form id='f1' autocomplete='off'>
    <label for='ident'>Usuario o correo</label>
    <input id='ident' maxlength='160' autocapitalize='none' autocorrect='off' required>
    <input id='hp' class='hp' tabindex='-1' autocomplete='off' aria-hidden='true'>
    <div id='e1' class='msg err hide'></div>
    <button class='btn' id='b1'>Continuar</button>
  </form>
  <p class='mut' style='margin-top:14px'>¿Ya tiene cuenta activa? <a href='/login'>Ingrese aquí</a>.</p>
</section>

<section id='s2' class='card hide'>
  <h1 id='t2'></h1>
  <div class='msg info' id='who'></div>
  <form id='f2' autocomplete='off'>
    <div id='alta' class='hide'>
      <label for='nombre'>Nombre completo *</label><input id='nombre' maxlength='160'>
      <label for='cedula'>Cédula *</label><input id='cedula' maxlength='20' inputmode='numeric'>
      <label for='correo'>Correo *</label><input id='correo' type='email' maxlength='160'>
      <label for='telefono'>Teléfono / WhatsApp</label><input id='telefono' maxlength='30' inputmode='tel'>
      <label for='area'>Área *</label><select id='area'><option value=''>Seleccione…</option></select>
      <label for='cargo'>Cargo</label><input id='cargo' maxlength='120'>
      <label for='jefe'>Jefe inmediato</label><input id='jefe' maxlength='120'>
    </div>
    <div id='cons' class='hide'>
      <label class='chk'><input type='checkbox' id='consent'><span>__CONSENT__</span></label>
    </div>
    <div id='e2' class='msg err hide'></div>
    <button class='btn'>Continuar a la foto de rostro</button>
    <button type='button' class='btn sec' onclick='location.reload()'>Cancelar</button>
  </form>
</section>

<section id='s3' class='card hide'>
  <h1>Foto de rostro</h1>
  <p class='mut'>Sin gorra ni gafas oscuras, con buena luz y mirando de frente. Solo debe aparecer usted.</p>
  <div class='cam' id='cam'><video id='video' playsinline muted></video><img id='prev' class='hide' alt='Vista previa'></div>
  <div class='hint' id='hint'></div>
  <div id='e3' class='msg err hide'></div>
  <button class='btn' id='b3' disabled>Tomar foto</button>
  <button class='btn sec hide' id='b3r' type='button'>Repetir foto</button>
  <button class='btn ok hide' id='b3s' type='button'>Enviar solicitud</button>
  <label class='btn sec' style='cursor:pointer'>Tomar foto con el teléfono<input id='fileCam' type='file' accept='image/*' capture='user' style='display:none'></label>
  <button class='btn sec' id='volver' type='button'>← Volver</button>
</section>

<section id='s4' class='card hide'>
  <h1>Solicitud enviada ✔</h1>
  <p>Gerencia revisará su foto y datos. Cuando sea aprobada recibirá su usuario y una clave temporal por el canal que Gerencia le indique.</p>
  <div class='msg ok'>Su código de seguimiento: <b id='codigo'></b></div>
  <a id='lnk' class='btn sec' href='#'>Ver estado de mi solicitud</a>
</section>
""".replace("__CONSENT__", _esc(CONSENTIMIENTO))
    js = JS_PUBLICO.replace("__CSRF__", csrf).replace("__AREAS__", json.dumps(areas).replace("</", "<\\/"))
    return _layout("Acceso del equipo", cuerpo, js)


# --------------------------------------------------------------------------------------
# Registro de rutas
# --------------------------------------------------------------------------------------
def registrar_acceso_equipo(app, db=None, Usuario=None, crear_hash=None, registrar_auditoria=None,
                            rate_limit=None, areas=None, store=None):
    if getattr(app, "_acceso_equipo_ok", False):
        return
    app._acceso_equipo_ok = True
    from werkzeug.security import generate_password_hash
    crear_hash = crear_hash or generate_password_hash
    st = store or SqlStore(db, Usuario)

    # opciones de área para el formulario público
    opciones = []
    try:
        for k, v in (areas or {}).items():
            opciones.append([k, (v.get("titulo") if isinstance(v, dict) else None) or k])
    except Exception:
        opciones = []
    for r in ROLES_EXTRA:
        opciones.append([r, r])
    VALIDAS = {o[0] for o in opciones}

    def audit(accion, detalle=""):
        try:
            if registrar_auditoria:
                registrar_auditoria(accion, detalle)
        except Exception:
            pass

    def ip():
        xf = (request.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
        return (xf or request.remote_addr or "")[:64]

    def limite(nombre, n, seg):
        if rate_limit:
            try:
                ok, _ = rate_limit(nombre, n, seg)
                return ok
            except Exception:
                return True
        return True

    def csrf_token():
        t = session.get("ae_csrf")
        if not t:
            t = secrets.token_urlsafe(24)
            session["ae_csrf"] = t
        return t

    def csrf_ok(valor):
        t = session.get("ae_csrf")
        return bool(t) and bool(valor) and secrets.compare_digest(str(t), str(valor))

    def nocache(resp):
        resp.headers["Cache-Control"] = "no-store"
        return resp

    # ---- arranque perezoso (tablas + gerente inicial opcional) ----
    estado_init = {"ok": False}

    @app.before_request
    def _ae_init():
        if estado_init["ok"]:
            return
        p = request.path or ""
        if not (p.startswith("/acceso-equipo") or p.startswith("/gerencia/solicitudes-acceso")):
            return
        try:
            st.ensure_tables()
            u, c = os.environ.get("BOOTSTRAP_GERENTE_USUARIO", "").strip(), os.environ.get("BOOTSTRAP_GERENTE_CLAVE", "")
            if u and c and not st.usuario_existe(u):
                st.crear_usuario({"usuario": u, "password_hash": crear_hash(c), "rol": "Gerente", "correo": "",
                                  "nombre": u, "area": "Gerencia"})
                st.commit()
            estado_init["ok"] = True
        except Exception as e:
            st.rollback()
            print("acceso_equipo init:", repr(e)[:300], flush=True)

    # ---- público ----
    @app.route("/acceso-equipo")
    def ae_pagina():
        return nocache(make_response(_pagina_publica(csrf_token(), opciones)))

    def jerr(msg, code=400):
        return nocache(make_response(jsonify({"error": msg}), code))

    def datos_json():
        d = request.get_json(silent=True, force=True)
        return d if isinstance(d, dict) else {}

    @app.route("/acceso-equipo/buscar", methods=["POST"])
    def ae_buscar():
        if not csrf_ok(request.headers.get("X-CSRF")):
            return jerr("La página caducó. Recárguela e intente de nuevo.", 403)
        if not limite("ae-buscar", 12, 300):
            return jerr("Demasiados intentos. Espere unos minutos.", 429)
        d = datos_json()
        if d.get("hp"):
            return jerr("Solicitud no válida.")
        ident = _limpiar(d.get("ident"), 160)
        if len(ident) < 3:
            return jerr("Escriba su usuario o correo.")
        try:
            u = st.buscar_usuario(ident)
        except Exception as e:
            st.rollback()
            print("acceso_equipo buscar:", repr(e)[:300], flush=True)
            return jerr("No se pudo consultar. Intente de nuevo en un momento.", 500)
        if not u:
            return nocache(jsonify({"existe": False}))
        if not u["activo"]:
            return nocache(jsonify({"existe": True, "usuario": {"id": u["id"], "usuario": u["usuario"],
                                                                "nombre": _mascara_nombre(u["nombre"]),
                                                                "correo": _mascara_correo(u["correo"])}}))
        return nocache(jsonify({"existe": True, "usuario": {"id": u["id"], "usuario": u["usuario"],
                                                            "nombre": _mascara_nombre(u["nombre"]),
                                                            "correo": _mascara_correo(u["correo"])}}))

    @app.route("/acceso-equipo/enviar", methods=["POST"])
    def ae_enviar():
        if not csrf_ok(request.headers.get("X-CSRF")):
            return jerr("La página caducó. Recárguela e intente de nuevo.", 403)
        if not limite("ae-enviar", 5, 600):
            return jerr("Demasiados envíos. Espere unos minutos.", 429)
        d = datos_json()
        if d.get("hp"):
            return jerr("Solicitud no válida.")
        if d.get("consent") is not True:
            return jerr("Debe autorizar el tratamiento de datos para continuar.")
        foto, h, e = procesar_foto(d.get("foto"))
        if e:
            return jerr(e)
        origen = "archivo" if d.get("origen") == "archivo" else "camara"
        try:
            hace_1h = (datetime.now(BOGOTA) - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
            if st.contar_ip_reciente(ip(), hace_1h) >= 6:
                return jerr("Se alcanzó el máximo de solicitudes desde esta conexión. Intente más tarde.", 429)
            reg = {"token": secrets.token_urlsafe(9).replace("-", "A").replace("_", "B").upper()[:12],
                   "estado": "PENDIENTE", "foto_b64": foto, "foto_hash": h, "foto_origen": origen,
                   "consentimiento": True, "consent_fecha": _ahora(), "ip": ip(),
                   "user_agent": _limpiar(request.headers.get("User-Agent"), 300), "creado_en": _ahora()}
            uid = d.get("usuario_id")
            if uid not in (None, ""):
                u = st.usuario_por_id(int(uid))
                if not u:
                    return jerr("La cuenta indicada ya no existe. Recargue la página.")
                reg.update(tipo="REVALIDACION", usuario_id=u["id"], usuario_nombre=u["usuario"],
                           nombre=u["nombre"] or u["usuario"], correo=u["correo"])
                if st.pendiente_duplicada("", u["correo"]) if u["correo"] else False:
                    return jerr("Ya existe una solicitud pendiente para esta cuenta.")
            else:
                nombre, cedula, correo = _limpiar(d.get("nombre")), re.sub(r"\D", "", str(d.get("cedula") or ""))[:20], _limpiar(d.get("correo")).lower()
                area = _limpiar(d.get("area"), 60)
                if len(nombre) < 5 or " " not in nombre:
                    return jerr("Escriba su nombre completo.")
                if len(cedula) < 5:
                    return jerr("Escriba una cédula válida.")
                if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", correo):
                    return jerr("Escriba un correo válido.")
                if area not in VALIDAS:
                    return jerr("Seleccione su área.")
                if st.pendiente_duplicada(cedula, correo):
                    return jerr("Ya existe una solicitud pendiente con esa cédula o correo. Gerencia la está revisando.")
                reg.update(tipo="ALTA", nombre=nombre, cedula=cedula, correo=correo, area=area,
                           telefono=_limpiar(d.get("telefono"), 40), cargo=_limpiar(d.get("cargo"), 120),
                           jefe=_limpiar(d.get("jefe"), 120))
            sid = st.crear_solicitud(reg)
        except ValueError:
            return jerr("Solicitud no válida.")
        except Exception as ex:
            st.rollback()
            print("acceso_equipo enviar:", repr(ex)[:300], flush=True)
            return jerr("No se pudo guardar la solicitud. Intente de nuevo.", 500)
        audit("Solicitud acceso equipo", "id=%s tipo=%s" % (sid, reg["tipo"]))
        return nocache(jsonify({"ok": True, "token": reg["token"]}))

    @app.route("/acceso-equipo/estado/<token>")
    def ae_estado(token):
        if not limite("ae-estado", 30, 300):
            return make_response(_layout("Estado", "<div class='card'>Demasiadas consultas.</div>"), 429)
        r = None
        try:
            r = st.por_token(re.sub(r"[^A-Za-z0-9]", "", token)[:20].upper())
        except Exception:
            st.rollback()
        if not r:
            cuerpo = "<div class='card'><h1>Solicitud no encontrada</h1><a class='btn sec' href='/acceso-equipo'>Volver</a></div>"
        else:
            msg = {"PENDIENTE": "Gerencia aún no la revisa. Vuelva a consultar más tarde.",
                   "APROBADA": "Aprobada. Su usuario es <b>%s</b>. Gerencia le entregará la clave temporal; al entrar deberá cambiarla." % _esc(r["usuario_nombre"]),
                   "RECHAZADA": "No fue aprobada. Motivo: %s" % _esc(r["motivo"] or "no indicado")}[r["estado"]]
            cuerpo = ("<div class='card'><h1>Solicitud %s</h1><p><span class='pill %s'>%s</span></p><p>%s</p>"
                      "<a class='btn sec' href='/login'>Ir al ingreso</a></div>" % (_esc(r["token"]), r["estado"], r["estado"], msg))
        return nocache(make_response(_layout("Estado de solicitud", cuerpo)))

    # ---- Gerencia ----
    def es_gerencia():
        return bool(session.get("usuario")) and (session.get("rol") or "").strip() in ROLES_APRUEBAN

    def denegado():
        return make_response(_layout("Acceso denegado", "<div class='card'><h1>Acceso denegado</h1>"
                                     "<p>Solo Gerencia puede revisar solicitudes.</p><a class='btn sec' href='/login'>Ingresar</a></div>"), 403)

    def nuevo_usuario(nombre, correo):
        base = _sin_tildes((nombre or correo.split("@")[0]).lower())
        p = [x for x in re.split(r"[^a-z]+", base) if x]
        cand = (p[0] + "." + p[-1]) if len(p) > 1 else (p[0] if p else "usuario")
        cand, n = cand[:30], 1
        orig = cand
        while st.usuario_existe(cand):
            n += 1
            cand = "%s%d" % (orig, n)
        return cand

    @app.route("/gerencia/solicitudes-acceso")
    def ae_panel():
        if not es_gerencia():
            return denegado()
        flt = (request.args.get("estado") or "PENDIENTE").upper()
        if flt not in ("PENDIENTE", "APROBADA", "RECHAZADA", "TODAS"):
            flt = "PENDIENTE"
        tok = csrf_token()
        filas = []
        try:
            lista = st.listar(flt)
        except Exception as e:
            st.rollback()
            return make_response(_layout("Solicitudes", "<div class='card'><p>No se pudo leer la tabla: %s</p></div>" % _esc(repr(e)[:200])), 500)
        roles_sel = sorted({o[0] for o in opciones})
        if (session.get("rol") or "") in ("Gerente", "Superadmin"):
            roles_sel = ["Gerente", "Administrador"] + roles_sel
        for r in lista:
            alerta = ""
            if r["foto_origen"] == "archivo":
                alerta += "<div class='msg err'>⚠ Foto subida desde archivo (no en vivo). Verifique con cuidado.</div>"
            try:
                if st.foto_hash_repetida(r["foto_hash"]) > 1:
                    alerta += "<div class='msg err'>⚠ Esta misma foto ya se usó en otra solicitud.</div>"
            except Exception:
                st.rollback()
            acciones = ""
            if r["estado"] == "PENDIENTE":
                opts = "".join("<option%s>%s</option>" % (" selected" if o == r["area"] else "", _esc(o)) for o in roles_sel)
                uh = "" if r["tipo"] == "REVALIDACION" else (
                    "<label>Usuario</label><input name='usuario' value='%s' maxlength='40'><label>Rol / área</label><select name='rol'>%s</select>" %
                    (_esc(nuevo_usuario(r["nombre"], r["correo"])), opts))
                acciones = (
                    "<form method='post' action='/gerencia/solicitudes-acceso/%d/aprobar' onsubmit=\"return confirm('¿Aprobar y generar clave temporal?')\">"
                    "<input type='hidden' name='csrf' value='%s'>%s<button class='btn ok'>Aprobar y generar clave</button></form>"
                    "<form method='post' action='/gerencia/solicitudes-acceso/%d/rechazar'><input type='hidden' name='csrf' value='%s'>"
                    "<label>Motivo del rechazo</label><input name='motivo' maxlength='250' required><button class='btn bad'>Rechazar</button></form>"
                    % (r["id"], tok, uh, r["id"], tok))
            else:
                acciones = "<p class='mut'>Resuelta por %s el %s. %s</p>" % (_esc(r["resuelto_por"]), _esc(r["resuelto_en"]), _esc(r["motivo"]))
            filas.append(
                "<div class='card'><div style='display:flex;gap:14px;flex-wrap:wrap'>"
                "<img src='/gerencia/solicitudes-acceso/foto/%d' alt='Foto' style='width:150px;height:150px;border-radius:12px;object-fit:cover'>"
                "<div style='flex:1;min-width:220px'><h2>%s <span class='pill %s'>%s</span></h2>"
                "<p class='mut'>%s · %s</p><p>Cédula: <b>%s</b><br>Correo: %s<br>Tel: %s<br>Área: %s · Cargo: %s<br>Jefe: %s</p>"
                "<p class='mut'>IP %s · %s</p></div></div>%s%s</div>" % (
                    r["id"], _esc(r["nombre"]), r["estado"], r["estado"], _esc(r["tipo"]), _esc(r["creado_en"]),
                    _esc(r["cedula"] or "—"), _esc(r["correo"]), _esc(r["telefono"] or "—"), _esc(r["area"] or "—"),
                    _esc(r["cargo"] or "—"), _esc(r["jefe"] or "—"), _esc(r["ip"]), _esc(r["foto_origen"]), alerta, acciones))
        tabs = " ".join("<a class='btn sm %s' href='?estado=%s'>%s</a>" % ("" if e == flt else "sec", e, e.title()) for e in ("PENDIENTE", "APROBADA", "RECHAZADA", "TODAS"))
        cuerpo = ("<div class='card'><h1>Solicitudes de acceso</h1><p class='mut'>Compare la foto con la persona (videollamada o presencial) antes de aprobar.</p>"
                  "<p>%s</p><p><a href='/gerencia/hq'>← Gerencia</a></p></div>%s" %
                  (tabs, "".join(filas) or "<div class='card'><p class='mut'>No hay solicitudes en este estado.</p></div>"))
        return nocache(make_response(_layout("Solicitudes de acceso", cuerpo)))

    @app.route("/gerencia/solicitudes-acceso/foto/<int:sid>")
    def ae_foto(sid):
        if not es_gerencia():
            return denegado()
        r = st.obtener(sid, con_foto=True)
        if not r or not r.get("foto_b64"):
            return make_response("", 404)
        resp = make_response(base64.b64decode(r["foto_b64"]))
        resp.headers["Content-Type"] = "image/jpeg"
        resp.headers["Cache-Control"] = "private, no-store"
        return resp

    @app.route("/gerencia/solicitudes-acceso/<int:sid>/aprobar", methods=["POST"])
    def ae_aprobar(sid):
        if not es_gerencia():
            return denegado()
        if not csrf_ok(request.form.get("csrf")):
            return make_response("Sesión caducada. Recargue el panel.", 403)
        r = st.obtener(sid)
        if not r or r["estado"] != "PENDIENTE":
            return make_response("Solicitud no disponible.", 404)
        clave = secrets.token_urlsafe(9)
        try:
            if r["tipo"] == "REVALIDACION":
                usuario = st.reset_usuario(r["usuario_id"], crear_hash(clave), "acceso-equipo:%d" % sid)
            else:
                usuario = re.sub(r"[^a-z0-9._-]", "", _sin_tildes(_limpiar(request.form.get("usuario"), 40).lower()))
                rol = _limpiar(request.form.get("rol"), 60)
                permitidos = set(VALIDAS) | ({"Gerente", "Administrador"} if (session.get("rol") in ("Gerente", "Superadmin")) else set())
                if len(usuario) < 3 or rol not in permitidos:
                    return make_response("Usuario o rol no válido.", 400)
                if st.usuario_existe(usuario):
                    return make_response("Ese usuario ya existe. Vuelva atrás y elija otro.", 409)
                uid = st.crear_usuario({"usuario": usuario, "password_hash": crear_hash(clave), "rol": rol, "correo": r["correo"],
                                        "nombre": r["nombre"], "area": r["area"], "jefe": r["jefe"],
                                        "biometria_path": "acceso-equipo:%d" % sid})
            st.resolver(sid, "APROBADA", session.get("usuario") or "", "", usuario_nombre=usuario)
        except Exception as e:
            st.rollback()
            print("acceso_equipo aprobar:", repr(e)[:300], flush=True)
            return make_response("No se pudo aprobar: %s" % _esc(repr(e)[:200]), 500)
        audit("Acceso equipo aprobado", "solicitud=%s usuario=%s" % (sid, usuario))
        cuerpo = ("<div class='card'><h1>Acceso aprobado ✔</h1><p>Entregue estos datos a la persona por un canal seguro. "
                  "<b>La clave no se vuelve a mostrar.</b> Deberá cambiarla al ingresar.</p>"
                  "<div class='msg ok'>Usuario: <b>%s</b><br>Clave temporal: <b style='font-family:monospace;font-size:18px'>%s</b></div>"
                  "<a class='btn' href='/gerencia/solicitudes-acceso'>Volver a solicitudes</a></div>" % (_esc(usuario), _esc(clave)))
        return nocache(make_response(_layout("Acceso aprobado", cuerpo)))

    @app.route("/gerencia/solicitudes-acceso/<int:sid>/rechazar", methods=["POST"])
    def ae_rechazar(sid):
        if not es_gerencia():
            return denegado()
        if not csrf_ok(request.form.get("csrf")):
            return make_response("Sesión caducada. Recargue el panel.", 403)
        r = st.obtener(sid)
        if not r or r["estado"] != "PENDIENTE":
            return make_response("Solicitud no disponible.", 404)
        try:
            st.resolver(sid, "RECHAZADA", session.get("usuario") or "", _limpiar(request.form.get("motivo"), 250) or "Sin motivo")
        except Exception:
            st.rollback()
            return make_response("No se pudo rechazar.", 500)
        audit("Acceso equipo rechazado", "solicitud=%s" % sid)
        from flask import redirect
        return redirect("/gerencia/solicitudes-acceso")
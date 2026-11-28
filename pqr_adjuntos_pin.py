# -*- coding: utf-8 -*-
"""
pqr_adjuntos_pin.py  ·  [procsis-pqr-seguro-v1]

Agrega al radicador de PQR (chat /pqr/crear y formulario /pqr/formulario):
  1) Verificación de identidad por PIN de 6 dígitos (correo o SMS).
  2) Adjuntar archivos (PDF, ZIP, fotos, Word, Excel) en el momento de radicar.
  3) El equipo de soporte ve y descarga los adjuntos dentro del detalle del PQR.

NO modifica app.py: se engancha con  pqr_adjuntos_pin.instalar(globals())
Si algo falla, la app sigue funcionando igual que antes.

Variables de entorno (todas opcionales):
  PQR_PIN_OBLIGATORIO = 1 (por defecto) | 0 para desactivar la exigencia del PIN
  TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN y (TWILIO_FROM o TWILIO_MESSAGING_SERVICE_SID)
      -> habilitan el PIN por SMS. Sin esto, el PIN se envía solo por correo.
El correo sale por el mismo sistema que ya usas (Resend / SMTP de Soporte).
"""
import base64
import hashlib
import hmac
import html as _html
import json
import os
import re
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
from email.message import EmailMessage
from functools import wraps

PIN_TTL = 600            # el PIN vence a los 10 minutos
TOKEN_TTL = 1800         # tras verificar, el permiso para radicar dura 30 min
PIN_MAX_INTENTOS = 5
PIN_COOLDOWN = 45        # segundos entre reenvíos al mismo destino
PIN_MAX_POR_VENTANA = 4  # PINs por destino cada 10 minutos
MAX_ARCHIVOS = 5
MAX_BYTES = 5 * 1024 * 1024
MAX_TOTAL = 15 * 1024 * 1024
LIMPIEZA_HORAS = 24      # adjuntos subidos y nunca radicados se borran tras 24 h

EXT_MIME = {
    "pdf": "application/pdf", "zip": "application/zip",
    "jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
    "webp": "image/webp", "gif": "image/gif", "heic": "image/heic",
    "doc": "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "txt": "text/plain",
}
ACCEPT_ATTR = ",".join("." + e for e in EXT_MIME)

RE_MAIL = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$")


# ───────────────────────── utilidades puras (sin Flask) ─────────────────────────
def norm_tel(s):
    """Celular colombiano -> +573001234567 (o '' si no es válido)."""
    x = re.sub(r"[\s\-\(\)\.]", "", str(s or ""))
    if re.fullmatch(r"\+57\d{10}", x):
        x = x[3:]
    elif re.fullmatch(r"57\d{10}", x):
        x = x[2:]
    return "+57" + x if re.fullmatch(r"3\d{9}", x) else ""


def mask_mail(m):
    a, _, b = str(m or "").partition("@")
    return (a[:2] + "•••@" + b) if b else str(m or "")


def mask_tel(t):
    return "••• ••• " + str(t or "")[-4:]


def firma_ok(ext, b):
    """Comprueba que el contenido real coincide con la extensión (no solo el nombre)."""
    if ext == "pdf":
        return b"%PDF" in b[:1024]
    if ext in ("zip", "docx", "xlsx"):
        return b[:2] == b"PK"
    if ext in ("jpg", "jpeg"):
        return b[:3] == b"\xff\xd8\xff"
    if ext == "png":
        return b[:8] == b"\x89PNG\r\n\x1a\n"
    if ext == "gif":
        return b[:6] in (b"GIF87a", b"GIF89a")
    if ext == "webp":
        return b[:4] == b"RIFF" and b[8:12] == b"WEBP"
    if ext == "heic":
        return b[4:8] == b"ftyp"
    if ext in ("doc", "xls"):
        return b[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    if ext == "txt":
        return b"\x00" not in b[:4096]
    return False


def nombre_seguro(nombre, ext, n):
    base = os.path.splitext(os.path.basename(str(nombre or "")))[0]
    try:
        from werkzeug.utils import secure_filename
        base = secure_filename(base)
    except Exception:
        base = re.sub(r"[^A-Za-z0-9_\-]+", "_", base)
    base = (base or "archivo_%d" % n)[:80]
    return "%s.%s" % (base, ext)


def fmt_tam(n):
    n = int(n or 0)
    return ("%.1f MB" % (n / 1048576.0)) if n >= 1048576 else ("%d KB" % max(1, round(n / 1024.0)))


# ───────────────────────── parche del chat (JavaScript) ─────────────────────────
_CSS = r"""
.adj-box{align-self:stretch;max-width:96%;background:#fff;border:1px solid #cbd5e1;border-radius:10px;padding:12px 14px;animation:pop .3s ease}
.adj-tit{font-weight:800;font-size:13px;color:#0f172a;margin-bottom:8px}
.adj-tit span{font-weight:600;color:#475569}
.adj-bar{background:#e8eef7;border:1px solid #cdd8ea;border-radius:6px 6px 0 0;padding:6px}
.adj-box button{font-family:inherit}
.adj-add{background:#fff;border:1px solid #b8c4d8;border-radius:5px;padding:6px 12px;font-size:13px;font-weight:700;color:#0f172a;cursor:pointer}
.adj-add:hover{background:#f0f7ff;border-color:#0020C2}
.adj-list{min-height:64px;max-height:170px;overflow:auto;border:1px solid #cdd8ea;border-top:0;border-radius:0 0 6px 6px;background:#fff;padding:6px 8px}
.adj-vacio{color:#94a3b8;font-size:12.5px;padding:8px 2px}
.adj-row{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:13px;padding:5px 2px;border-bottom:1px solid #f1f5f9;word-break:break-all}
.adj-row button{border:0;background:#fee2e2;color:#b91c1c;border-radius:6px;cursor:pointer;font-weight:800;padding:2px 8px}
.adj-msg{font-size:12.5px;color:#b45309;min-height:16px;margin-top:6px}
.adj-acc{margin-top:8px}
.adj-go{background:#0020C2;color:#fff;border:0;border-radius:10px;padding:10px 18px;font-weight:700;font-size:13.5px;cursor:pointer}
.adj-go:disabled{opacity:.5;cursor:not-allowed}
"""

_JS_BLOCK = r"""/* [pqr-seguro] verificación por PIN + adjuntos */
var PIN_CANALES = {sms:false};
try{ fetch("/pqr/pin/canales").then(function(r){ return r.json(); }).then(function(j){ if(j) PIN_CANALES = j; }).catch(function(){}); }catch(e){}
function maskMail(m){ var p = String(m||"").split("@"); if(p.length < 2) return m||""; return p[0].slice(0,2) + "•••@" + p[1]; }
function maskTel(t){ t = String(t||""); return "••• ••• " + t.slice(-4); }
function normTel(s){ var x = String(s||"").replace(/[\s\-\(\)\.]/g,""); if(/^\+57\d{10}$/.test(x)) x = x.slice(3); else if(/^57\d{10}$/.test(x)) x = x.slice(2); return /^3\d{9}$/.test(x) ? "+57" + x : ""; }
function fmtKB(n){ return n >= 1048576 ? (n/1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n/1024)) + " KB"; }
function postJSON(url, obj, cb){
  fetch(url, {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(obj)})
    .then(function(r){ return r.json().then(function(j){ return {s:r.status, j:j}; }); })
    .then(function(o){ cb(null, o.j || {}, o.s); })
    .catch(function(){ cb("Se perdió la conexión. Inténtalo de nuevo."); });
}
function enviarPin(dest, canal, cb){
  postJSON("/pqr/pin/enviar", {canal:canal, destino:dest}, function(e, j){
    if(e) return cb(e);
    if(!j.ok) return cb(j.error || "No pude enviar el PIN. Inténtalo de nuevo.");
    cb(null);
  });
}
STEPS.pinelegir = {
  ask:function(d){
    var o = [["correo", "📧 Por correo (" + maskMail(d.email) + ")"]];
    if(PIN_CANALES.sms) o.push(["sms", "📱 Por SMS a mi celular"]);
    return {text:"🔐 Para proteger tu solicitud necesito verificar tu identidad. ¿Dónde quieres recibir un PIN de 6 dígitos?", opts:o};
  },
  pre:function(d, v, l, done){ if(v === "correo"){ enviarPin(d.email, "correo", function(e){ done(e); }); } else { done(null); } },
  apply:function(d, v){ if(v !== "correo" && v !== "sms") return "Elige una de las opciones."; d.pincanal = v; return null; },
  next:function(d){ return d.pincanal === "sms" ? "telefono" : "pincodigo"; }
};
STEPS.telefono = {
  ask:function(){ return {text:"Escribe tu número de celular (10 dígitos, ejemplo 3001234567). Te enviaré el PIN por SMS.", ph:"Celular", max:16}; },
  pre:function(d, v, l, done){ var t = normTel(l); if(!t) return done("Ese celular no parece válido. Debe tener 10 dígitos y empezar por 3."); enviarPin(t, "sms", function(e){ done(e); }); },
  apply:function(d, v, l){ var t = normTel(l); if(!t) return "Ese celular no parece válido. Debe tener 10 dígitos y empezar por 3."; d.tel = t; return null; },
  next:function(){ return "pincodigo"; }
};
STEPS.pincodigo = {
  ask:function(d){
    var dest = d.pincanal === "sms" ? maskTel(d.tel) : maskMail(d.email);
    return {text:"✉️ Te envié un PIN de 6 dígitos a " + dest + ". Escríbelo aquí (vence en 10 minutos).", ph:"PIN de 6 dígitos", max:6, opts:[["__reenviar", "📨 Reenviar PIN"], ["__cambiar", "↩ Cambiar medio de verificación"]]};
  },
  pre:function(d, v, l, done){
    var canal = d.pincanal === "sms" ? "sms" : "correo", dest = canal === "sms" ? d.tel : d.email;
    if(v === "__reenviar"){ return enviarPin(dest, canal, function(e){ done(e || "Listo, te envié un nuevo PIN."); }); }
    if(v === "__cambiar"){ if(canal === "sms") answers.pop(); answers.pop(); save(); return done("Elige cómo quieres recibir el PIN."); }
    var pin = String(l || "").replace(/\s/g, "");
    if(!/^\d{6}$/.test(pin)) return done("El PIN tiene 6 dígitos. Revísalo e inténtalo de nuevo.");
    postJSON("/pqr/pin/verificar", {canal:canal, destino:dest, pin:pin}, function(e, j){
      if(e) return done(e);
      if(!j.ok) return done(j.error || "PIN incorrecto.");
      done(null, j.token, "✅ PIN verificado");
    });
  },
  apply:function(d, v){ if(!/^[A-Za-z0-9_\-]{20,}$/.test(String(v || ""))) return "Necesito el PIN de 6 dígitos para continuar."; d.pintoken = v; return null; },
  next:function(d){ return (d.perfilKey === "inst" || d.doctipo === "NIT") ? "dep" : "nitcol"; }
};
function adjWidget(d){
  var tok = d.pintoken, n = 0;
  var box = document.createElement("div"); box.className = "adj-box";
  box.innerHTML = '<div class="adj-tit">Adjuntar archivos. <span>Nota: se permiten PDF, ZIP, imágenes (JPG, PNG, WEBP, HEIC), Word y Excel · máx. 5 MB por archivo · hasta 5 archivos.</span></div>' +
    '<div class="adj-bar"><button type="button" class="adj-add">＋ Adicionar</button><input type="file" class="adj-file" multiple hidden accept="__ACCEPT__"></div>' +
    '<div class="adj-list"></div><div class="adj-msg"></div>' +
    '<div class="adj-acc"><button type="button" class="adj-go">Continuar sin adjuntos</button></div>';
  log.appendChild(box);
  var list = box.querySelector(".adj-list"), msg = box.querySelector(".adj-msg"), go = box.querySelector(".adj-go"), inp = box.querySelector(".adj-file");
  function pinta(arr){
    list.innerHTML = ""; n = arr.length;
    if(!n){ var v = document.createElement("div"); v.className = "adj-vacio"; v.textContent = "Aún no has adjuntado archivos."; list.appendChild(v); }
    arr.forEach(function(a){
      var r = document.createElement("div"); r.className = "adj-row";
      var s = document.createElement("span"); s.textContent = "📄 " + a.nombre + " (" + fmtKB(a.tamano) + ")";
      var x = document.createElement("button"); x.type = "button"; x.textContent = "✕"; x.title = "Quitar";
      x.addEventListener("click", function(){
        postJSON("/pqr/adjuntos/quitar", {token:tok, id:a.id}, function(e, j){ if(!e && j.ok){ pinta(j.archivos || []); } else { msg.textContent = (j && j.error) || e || "No pude quitar el archivo."; } });
      });
      r.appendChild(s); r.appendChild(x); list.appendChild(r);
    });
    go.textContent = n ? ("Continuar con " + n + " archivo" + (n > 1 ? "s" : "")) : "Continuar sin adjuntos";
    scrollEnd();
  }
  pinta([]);
  fetch("/pqr/adjuntos/lista?token=" + encodeURIComponent(tok)).then(function(r){ return r.json(); }).then(function(j){ if(j && j.archivos) pinta(j.archivos); }).catch(function(){});
  box.querySelector(".adj-add").addEventListener("click", function(){ inp.click(); });
  inp.addEventListener("change", function(){
    if(!inp.files.length) return;
    var fd = new FormData(); fd.append("token", tok);
    for(var i = 0; i < inp.files.length; i++){ fd.append("archivos", inp.files[i]); }
    msg.textContent = "Subiendo…"; go.disabled = true;
    fetch("/pqr/adjuntos/subir", {method:"POST", body:fd}).then(function(r){ return r.json(); }).then(function(j){
      go.disabled = false; inp.value = "";
      if(j.archivos) pinta(j.archivos);
      msg.textContent = (j.avisos && j.avisos.length) ? j.avisos.join(" · ") : (j.ok ? "" : (j.error || "No pude subir el archivo."));
    }).catch(function(){ go.disabled = false; msg.textContent = "Se perdió la conexión al subir."; });
  });
  go.addEventListener("click", function(){ submit("adj:" + n, n ? (n + " archivo" + (n > 1 ? "s" : "") + " adjunto" + (n > 1 ? "s" : "")) : "Sin adjuntos"); });
}
STEPS.adjuntos = {
  ask:function(d){ return {text:"📎 ¿Tienes archivos que ayuden a entender tu caso (capturas de pantalla, PDF, ZIP, fotos)? Puedes adjuntarlos ahora; es opcional.", ph:"Usa el botón «Adicionar» o presiona Continuar", max:100, widget:function(){ adjWidget(d); }}; },
  apply:function(d, v){ if(String(v).indexOf("adj:") !== 0) return "Usa el botón «Adicionar» para subir archivos, o presiona Continuar."; d.adjn = parseInt(String(v).slice(4), 10) || 0; return null; },
  next:function(){ return "resumen"; }
};

"""

_PRE_SUBMIT = r"""if(st.pre){
    busy = true; sendBtn.disabled = true;
    st.pre(d, value, label, function(e, v2, l2){
      busy = false;
      if(e){ render(e); return; }
      var vv = (v2 === undefined ? value : v2), ll = (l2 === undefined ? label : l2);
      var snap = JSON.parse(JSON.stringify(d));
      var err2 = st.apply(snap, vv, ll);
      if(err2){ render(err2); return; }
      answers.push({step:step, value:vv, label:ll});
      save(); render();
    });
    return;
  }
  var snapshot = JSON.parse(JSON.stringify(d));"""

_ERR_PIN = r"""if(o.j && o.j.codigo === "pin"){ for(var k = 0; k < answers.length; k++){ if(answers[k].step === "pinelegir"){ answers = answers.slice(0, k); break; } } save(); }
        render((o.j && o.j.error) || "No pude radicar tu solicitud. Inténtalo de nuevo en unos minutos.");"""


def parchar_chat_html(html):
    """Devuelve el HTML del chat con los pasos nuevos, o None si no se puede parchar con seguridad."""
    if not isinstance(html, str) or "[pqr-seguro]" in html:
        return None
    cambios = [
        ("var TOTAL = 15;", "var TOTAL = 19;"),
        ("</style>\n<div class=\"fondo-corporativo-light\">", _CSS + "</style>\n<div class=\"fondo-corporativo-light\">"),
        ('next:function(d){ return (d.perfilKey === "inst" || d.doctipo === "NIT") ? "dep" : "nitcol"; }',
         'next:function(d){ return "pinelegir"; }'),
        ('next:function(){ return "resumen"; }', 'next:function(){ return "adjuntos"; }'),
        ('L.push("• Descripción: " + d.desc);',
         'L.push("• Descripción: " + d.desc); L.push("• Verificación: PIN ✓"); L.push("• Adjuntos: " + (d.adjn ? d.adjn + " archivo(s)" : "ninguno"));'),
        ("function replay(){", _JS_BLOCK.replace("__ACCEPT__", ACCEPT_ATTR) + "function replay(){"),
        ("if(q.opts) addOpts(q.opts, q.sug);\n  setList(q.list);",
         "if(q.opts) addOpts(q.opts, q.sug);\n  if(q.widget) q.widget();\n  setList(q.list);"),
        ("var snapshot = JSON.parse(JSON.stringify(d));", _PRE_SUBMIT),
        ('|| step === "desc"){ submit("txt", raw); return; }',
         '|| step === "desc" || step === "pincodigo" || step === "telefono"){ submit("txt", raw); return; }'),
        ('objeto:r.desc, hechos:hechos.join("\\n")',
         'objeto:r.desc, hechos:hechos.join("\\n"), pin_token:(r.pintoken || ""), telefono:(r.tel || "")'),
        ('render((o.j && o.j.error) || "No pude radicar tu solicitud. Inténtalo de nuevo en unos minutos.");', _ERR_PIN),
    ]
    out = html
    for ancla, nuevo in cambios:
        if out.count(ancla) != 1:
            print("PQR SEGURO: no pude parchar el chat (ancla no única):", ancla[:70].replace("\n", " "))
            return None
        out = out.replace(ancla, nuevo)
    return out  # el flujo ahora es: email → PIN → (dep/nitcol…) → descripción → adjuntos → resumen


# ───────────────────────── bloque para el formulario clásico ─────────────────────────
_FORM_BLOQUE = r"""
<div id="pqrseg" style="margin-top:14px;padding:12px;border:1px solid #cbd5e1;border-radius:12px;background:#f8fafc">
  <label style="margin-top:0">Verificación de identidad (PIN) *</label>
  <div class="row">
    <div><select id="pqrseg_canal"><option value="correo">Recibir PIN por correo</option>__SMS_OPT__</select></div>
    <div><input id="pqrseg_tel" placeholder="Celular (solo si eliges SMS)" style="display:none" maxlength="16"></div>
  </div>
  <button type="button" id="pqrseg_enviar" style="background:#334155;margin-top:8px">Enviar PIN</button>
  <div class="row" style="margin-top:8px">
    <div><input id="pqrseg_pin" placeholder="PIN de 6 dígitos" maxlength="6" inputmode="numeric"></div>
    <div><button type="button" id="pqrseg_ver" style="margin-top:0;background:#0f766e">Verificar PIN</button></div>
  </div>
  <div id="pqrseg_msg" style="font-size:12.5px;margin-top:6px;color:#475569"></div>
  <input type="hidden" name="pin_token" id="pqrseg_tok">
  <label>Adjuntar archivos (opcional)</label>
  <input type="file" name="archivos" multiple accept="__ACCEPT__">
  <div style="font-size:12px;color:#64748b;margin-top:4px">Se permiten PDF, ZIP, imágenes (JPG, PNG, WEBP, HEIC), Word y Excel · máx. 5 MB por archivo · hasta 5 archivos.</div>
</div>
<script>
(function(){
  var $ = function(i){ return document.getElementById(i); };
  var canal = $("pqrseg_canal"), tel = $("pqrseg_tel"), msg = $("pqrseg_msg"), tok = $("pqrseg_tok"), sub = $("pqrseg_submit");
  function say(t, ok){ msg.style.color = ok ? "#047857" : "#b45309"; msg.textContent = t; }
  function post(u, o, cb){ fetch(u, {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(o)}).then(function(r){ return r.json(); }).then(cb).catch(function(){ say("Se perdió la conexión."); }); }
  function dest(){ if(canal.value === "sms") return tel.value; var f = document.getElementsByName("email")[0]; return f ? f.value : ""; }
  canal.addEventListener("change", function(){ tel.style.display = canal.value === "sms" ? "block" : "none"; });
  $("pqrseg_enviar").addEventListener("click", function(){
    var d = dest(); if(!d){ say("Primero escribe el " + (canal.value === "sms" ? "celular" : "correo") + "."); return; }
    say("Enviando…", true);
    post("/pqr/pin/enviar", {canal:canal.value, destino:d}, function(j){ say(j.ok ? "Te enviamos un PIN a " + j.masked + ". Vence en 10 minutos." : (j.error || "No pude enviar el PIN."), !!j.ok); });
  });
  $("pqrseg_ver").addEventListener("click", function(){
    post("/pqr/pin/verificar", {canal:canal.value, destino:dest(), pin:$("pqrseg_pin").value}, function(j){
      if(j.ok){ tok.value = j.token; say("✅ Identidad verificada. Ya puedes radicar.", true); if(sub) sub.disabled = false; }
      else { say(j.error || "PIN incorrecto."); }
    });
  });
})();
</script>
"""


# ───────────────────────── instalación ─────────────────────────
def instalar(g):
    if g.get("_PQR_SEGURO_INSTALADO"):
        return
    app = g["app"]
    db = g["db"]
    TicketPQR = g["TicketPQR"]
    from flask import request, jsonify, Response

    class PqrPin(db.Model):
        __tablename__ = "pqr_pin"
        id = db.Column(db.Integer, primary_key=True)
        canal = db.Column(db.String(10), default="correo")
        destino = db.Column(db.String(160), index=True, default="")
        sal = db.Column(db.String(32), default="")
        pin_hash = db.Column(db.String(64), default="")
        intentos = db.Column(db.Integer, default=0)
        creado_ts = db.Column(db.Float, default=0.0)
        expira_ts = db.Column(db.Float, default=0.0)
        verificado = db.Column(db.Boolean, default=False)
        usado = db.Column(db.Boolean, default=False)
        token = db.Column(db.String(64), index=True, default="")
        token_exp = db.Column(db.Float, default=0.0)
        ip = db.Column(db.String(64), default="")

    class PqrAdjunto(db.Model):
        __tablename__ = "pqr_adjuntos"
        id = db.Column(db.Integer, primary_key=True)
        ticket_id = db.Column(db.Integer, index=True, nullable=True)
        token = db.Column(db.String(64), index=True, default="")
        nombre = db.Column(db.String(160), default="")
        mime = db.Column(db.String(100), default="")
        tamano = db.Column(db.Integer, default=0)
        sha256 = db.Column(db.String(64), default="")
        datos = db.Column(db.LargeBinary)
        creado_ts = db.Column(db.Float, default=0.0)

    g["PqrPin"] = PqrPin
    g["PqrAdjunto"] = PqrAdjunto
    estado = {"tablas": False}

    # ---------- helpers ----------
    def tablas():
        if estado["tablas"]:
            return
        try:
            PqrPin.__table__.create(db.engine, checkfirst=True)
            PqrAdjunto.__table__.create(db.engine, checkfirst=True)
            estado["tablas"] = True
        except Exception as ex:
            print("PQR SEGURO tablas:", ex)

    def obligatorio():
        return (os.environ.get("PQR_PIN_OBLIGATORIO", "1").strip() not in ("0", "false", "no", "off"))

    def rl(nombre, n, ventana):
        f = g.get("_rate_limit")
        if not f:
            return True, 0
        try:
            return f(nombre, n, ventana)
        except Exception:
            return True, 0

    def sms_ok():
        return bool(os.environ.get("TWILIO_ACCOUNT_SID") and os.environ.get("TWILIO_AUTH_TOKEN")
                    and (os.environ.get("TWILIO_FROM") or os.environ.get("TWILIO_MESSAGING_SERVICE_SID")))

    def producto():
        try:
            return g["nombre_producto"]()
        except Exception:
            return "EduTrack"

    def clave():
        k = app.secret_key or "pqr-seguro"
        return k if isinstance(k, bytes) else str(k).encode("utf-8")

    def hash_pin(pin, sal):
        return hmac.new(clave(), ("%s:%s" % (sal, pin)).encode("utf-8"), hashlib.sha256).hexdigest()

    def ip_cliente():
        return (request.headers.get("X-Forwarded-For", request.remote_addr or "").split(",")[0].strip())[:60]

    def err(msg, code=400, **extra):
        d = {"ok": False, "error": msg}
        d.update(extra)
        return jsonify(d), code

    # ---------- envío del PIN ----------
    def enviar_correo(destino, pin):
        resend = (os.environ.get("RESEND_API_KEY") or "").strip()
        correo_envio, password = g["_credenciales_smtp"]("soporte")
        if not resend and not (correo_envio and password):
            return False, "El envío de correos no está configurado."
        from_addr = (os.environ.get("RESEND_FROM") or "").strip() or correo_envio or "EduTrack <onboarding@resend.dev>"
        if not correo_envio:
            correo_envio = from_addr
        prod = _html.escape(producto())
        msg = EmailMessage()
        msg["Subject"] = "Tu PIN de verificación: %s" % pin
        msg["From"] = from_addr
        msg["To"] = destino
        msg.set_content("Tu PIN de verificación para radicar tu PQR en %s es: %s\n\nVence en 10 minutos. "
                        "Si no lo solicitaste, ignora este mensaje y no compartas el código con nadie." % (producto(), pin))
        msg.add_alternative(
            "<div style='font-family:Segoe UI,Arial,sans-serif;max-width:460px;margin:auto;border:1px solid #e2e8f0;border-radius:14px;padding:24px'>"
            "<h2 style='margin:0 0 6px;color:#0020C2'>%s</h2>"
            "<p style='color:#334155'>Usa este PIN para verificar tu identidad y radicar tu PQR:</p>"
            "<p style='font-size:34px;letter-spacing:8px;font-weight:800;text-align:center;margin:18px 0;color:#0f172a'>%s</p>"
            "<p style='color:#64748b;font-size:13px'>Vence en 10 minutos. Si no lo solicitaste, ignora este mensaje y no compartas el código con nadie.</p></div>" % (prod, pin),
            subtype="html")
        try:
            g["_smtp_enviar"](msg, correo_envio, password)
            return True, ""
        except Exception as ex:
            print("PQR SEGURO correo:", ex)
            return False, "No pude enviar el correo en este momento. Inténtalo de nuevo."

    def enviar_sms(tel, pin):
        sid = os.environ.get("TWILIO_ACCOUNT_SID", "")
        tok = os.environ.get("TWILIO_AUTH_TOKEN", "")
        data = {"To": tel, "Body": "%s: tu PIN de verificación es %s. Vence en 10 minutos. No lo compartas." % (producto(), pin)}
        if os.environ.get("TWILIO_MESSAGING_SERVICE_SID"):
            data["MessagingServiceSid"] = os.environ["TWILIO_MESSAGING_SERVICE_SID"]
        else:
            data["From"] = os.environ.get("TWILIO_FROM", "")
        req = urllib.request.Request(
            "https://api.twilio.com/2010-04-01/Accounts/%s/Messages.json" % sid,
            data=urllib.parse.urlencode(data).encode("utf-8"), method="POST")
        req.add_header("Authorization", "Basic " + base64.b64encode(("%s:%s" % (sid, tok)).encode()).decode())
        try:
            urllib.request.urlopen(req, timeout=15).read()
            return True, ""
        except urllib.error.HTTPError as ex:
            print("PQR SEGURO sms HTTP", ex.code)
        except Exception as ex:
            print("PQR SEGURO sms:", ex)
        return False, "No pude enviar el SMS en este momento. Inténtalo de nuevo o usa el correo."

    # ---------- rutas del PIN ----------
    @app.route("/pqr/pin/canales")
    def pqr_pin_canales():
        return jsonify({"correo": True, "sms": sms_ok(), "obligatorio": obligatorio()})

    @app.route("/pqr/pin/enviar", methods=["POST"])
    def pqr_pin_enviar():
        tablas()
        ok, espera = rl("pqr-pin-ip", 8, 600)
        if not ok:
            return err("Demasiados intentos. Espera %s segundos." % espera, 429)
        d = request.get_json(silent=True) or {}
        canal = str(d.get("canal") or "").strip()
        if canal == "sms":
            if not sms_ok():
                return err("El envío por SMS no está habilitado. Usa el correo.", 503)
            destino = norm_tel(d.get("destino"))
            if not destino:
                return err("Ese celular no parece válido. Debe tener 10 dígitos y empezar por 3.")
        elif canal == "correo":
            destino = str(d.get("destino") or "").strip().lower()[:160]
            if not RE_MAIL.match(destino):
                return err("El correo no es válido.")
        else:
            return err("Medio de verificación no válido.")
        ahora = time.time()
        try:
            recientes = PqrPin.query.filter(PqrPin.destino == destino, PqrPin.creado_ts > ahora - 600).all()
            if recientes:
                ultimo = max(r.creado_ts for r in recientes)
                if ahora - ultimo < PIN_COOLDOWN:
                    return err("Espera %d segundos para pedir otro PIN." % int(PIN_COOLDOWN - (ahora - ultimo) + 1), 429)
                if len(recientes) >= PIN_MAX_POR_VENTANA:
                    return err("Pediste demasiados PIN. Inténtalo de nuevo en unos minutos.", 429)
            pin = "%06d" % secrets.randbelow(10 ** 6)
            sal = secrets.token_hex(8)
            fila = PqrPin(canal=canal, destino=destino, sal=sal, pin_hash=hash_pin(pin, sal), intentos=0,
                          creado_ts=ahora, expira_ts=ahora + PIN_TTL, verificado=False, usado=False,
                          token="", token_exp=0.0, ip=ip_cliente())
            db.session.add(fila)
            db.session.commit()
        except Exception as ex:
            db.session.rollback()
            print("PQR SEGURO enviar:", ex)
            return err("No pude generar el PIN. Inténtalo de nuevo.", 500)
        enviado, motivo = enviar_correo(destino, pin) if canal == "correo" else enviar_sms(destino, pin)
        if os.environ.get("PQR_PIN_DEBUG") == "1":
            print("PQR SEGURO [DEBUG] PIN para %s: %s" % (destino, pin))
        if not enviado:
            try:
                db.session.delete(fila)
                db.session.commit()
            except Exception:
                db.session.rollback()
            return err(motivo, 502)
        return jsonify({"ok": True, "masked": mask_mail(destino) if canal == "correo" else mask_tel(destino), "ttl": PIN_TTL})

    @app.route("/pqr/pin/verificar", methods=["POST"])
    def pqr_pin_verificar():
        tablas()
        ok, espera = rl("pqr-pin-verif", 25, 600)
        if not ok:
            return err("Demasiados intentos. Espera %s segundos." % espera, 429)
        d = request.get_json(silent=True) or {}
        canal = str(d.get("canal") or "").strip()
        destino = norm_tel(d.get("destino")) if canal == "sms" else str(d.get("destino") or "").strip().lower()[:160]
        pin = re.sub(r"\s", "", str(d.get("pin") or ""))
        if canal not in ("correo", "sms") or not destino or not re.fullmatch(r"\d{6}", pin):
            return err("Revisa el PIN: tiene 6 dígitos.")
        try:
            fila = (PqrPin.query.filter_by(canal=canal, destino=destino, verificado=False)
                    .order_by(PqrPin.id.desc()).first())
            if not fila or fila.expira_ts < time.time():
                return err("El PIN venció o no existe. Pide uno nuevo.")
            if fila.intentos >= PIN_MAX_INTENTOS:
                return err("Superaste los intentos. Pide un PIN nuevo.", 429)
            fila.intentos += 1
            if not hmac.compare_digest(fila.pin_hash, hash_pin(pin, fila.sal)):
                db.session.commit()
                quedan = max(0, PIN_MAX_INTENTOS - fila.intentos)
                return err("PIN incorrecto. Te quedan %d intento%s." % (quedan, "" if quedan == 1 else "s"))
            fila.verificado = True
            fila.token = secrets.token_urlsafe(32)
            fila.token_exp = time.time() + TOKEN_TTL
            db.session.commit()
            return jsonify({"ok": True, "token": fila.token})
        except Exception as ex:
            db.session.rollback()
            print("PQR SEGURO verificar:", ex)
            return err("No pude verificar el PIN. Inténtalo de nuevo.", 500)

    def fila_token(token):
        token = str(token or "").strip()
        if len(token) < 20:
            return None
        try:
            f = PqrPin.query.filter_by(token=token, verificado=True, usado=False).first()
        except Exception:
            return None
        if not f or f.token_exp < time.time():
            return None
        return f

    # ---------- adjuntos ----------
    def lista_json(token):
        rows = PqrAdjunto.query.filter_by(token=token, ticket_id=None).order_by(PqrAdjunto.id.asc()).all()
        return [{"id": r.id, "nombre": r.nombre, "tamano": r.tamano} for r in rows]

    def limpiar():
        try:
            lim = time.time() - LIMPIEZA_HORAS * 3600
            for r in PqrAdjunto.query.filter(PqrAdjunto.creado_ts < lim).filter_by(ticket_id=None).all():
                db.session.delete(r)
            db.session.commit()
        except Exception:
            db.session.rollback()

    def guardar(files, token, ticket_id=None):
        """Valida y guarda archivos. Devuelve (cuantos_guardados, [avisos])."""
        avisos, guardados = [], 0
        existentes = PqrAdjunto.query.filter_by(token=token).all()
        n_total, bytes_total = len(existentes), sum(int(e.tamano or 0) for e in existentes)
        for f in files:
            if not getattr(f, "filename", ""):
                continue
            original = f.filename
            ext = os.path.splitext(original)[1].lower().lstrip(".")
            if ext not in EXT_MIME:
                avisos.append("«%s»: formato no permitido." % original[:40])
                continue
            if n_total >= MAX_ARCHIVOS:
                avisos.append("Máximo %d archivos por solicitud." % MAX_ARCHIVOS)
                break
            datos = f.stream.read(MAX_BYTES + 1)
            if not datos:
                avisos.append("«%s»: el archivo está vacío." % original[:40])
                continue
            if len(datos) > MAX_BYTES:
                avisos.append("«%s»: supera los 5 MB." % original[:40])
                continue
            if bytes_total + len(datos) > MAX_TOTAL:
                avisos.append("«%s»: se superó el total permitido (15 MB)." % original[:40])
                continue
            if not firma_ok(ext, datos):
                avisos.append("«%s»: el contenido no corresponde a un archivo .%s." % (original[:40], ext))
                continue
            db.session.add(PqrAdjunto(
                ticket_id=ticket_id, token=token, nombre=nombre_seguro(original, ext, n_total + 1),
                mime=EXT_MIME[ext], tamano=len(datos), sha256=hashlib.sha256(datos).hexdigest(),
                datos=datos, creado_ts=time.time()))
            n_total += 1
            bytes_total += len(datos)
            guardados += 1
        db.session.commit()
        return guardados, avisos

    @app.route("/pqr/adjuntos/subir", methods=["POST"])
    def pqr_adjuntos_subir():
        tablas()
        ok, espera = rl("pqr-adj-subir", 30, 600)
        if not ok:
            return err("Demasiados intentos. Espera %s segundos." % espera, 429)
        token = request.form.get("token", "")
        if not fila_token(token):
            return err("Tu verificación venció. Vuelve a verificar tu identidad.", 403, codigo="pin")
        try:
            limpiar()
            n, avisos = guardar(request.files.getlist("archivos"), token)
            return jsonify({"ok": n > 0 or not avisos, "guardados": n, "avisos": avisos, "archivos": lista_json(token)})
        except Exception as ex:
            db.session.rollback()
            print("PQR SEGURO subir:", ex)
            return err("No pude guardar los archivos. Inténtalo de nuevo.", 500)

    @app.route("/pqr/adjuntos/lista")
    def pqr_adjuntos_lista():
        tablas()
        token = request.args.get("token", "")
        if not fila_token(token):
            return jsonify({"ok": False, "archivos": []})
        return jsonify({"ok": True, "archivos": lista_json(token)})

    @app.route("/pqr/adjuntos/quitar", methods=["POST"])
    def pqr_adjuntos_quitar():
        tablas()
        d = request.get_json(silent=True) or {}
        token = str(d.get("token") or "")
        if not fila_token(token):
            return err("Tu verificación venció.", 403, codigo="pin")
        try:
            a = PqrAdjunto.query.filter_by(id=int(d.get("id") or 0), token=token, ticket_id=None).first()
            if a:
                db.session.delete(a)
                db.session.commit()
            return jsonify({"ok": True, "archivos": lista_json(token)})
        except Exception:
            db.session.rollback()
            return err("No pude quitar el archivo.", 500)

    # ---------- vista del equipo de soporte ----------
    @app.route("/soporte/pqr/adjunto/<int:aid>")
    @app.route("/soluciones-procsis/pqr/adjunto/<int:aid>")
    def pqr_adjunto_descargar(aid):
        bloqueo = g["_pqr_guard"]("ver")
        if bloqueo:
            return bloqueo
        tablas()
        a = PqrAdjunto.query.filter_by(id=aid).first()
        if not a or not a.ticket_id:
            return Response("Archivo no encontrado", status=404)
        return Response(bytes(a.datos or b""), mimetype=a.mime or "application/octet-stream", headers={
            "Content-Disposition": "attachment; filename*=UTF-8''" + urllib.parse.quote(a.nombre or "adjunto"),
            "X-Content-Type-Options": "nosniff",
            "Cache-Control": "private, no-store",
        })

    orig_det = app.view_functions.get("soporte_pqr_detalle")
    if orig_det:
        @wraps(orig_det)
        def detalle(*a, **k):
            resp = orig_det(*a, **k)
            try:
                if request.method != "GET" or not isinstance(resp, str):
                    return resp
                tablas()
                tid = k.get("id")
                items = PqrAdjunto.query.filter_by(ticket_id=tid).order_by(PqrAdjunto.id.asc()).all()
                if not items:
                    return resp
                pref = "/soluciones-procsis" if request.path.startswith("/soluciones-procsis") else "/soporte"
                filas = "".join(
                    "<li style='margin:6px 0'><a href='%s/pqr/adjunto/%d' style='font-weight:700'>📎 %s</a> "
                    "<span style='color:#64748b;font-size:12px'>(%s)</span></li>"
                    % (pref, it.id, _html.escape(it.nombre or "archivo"), fmt_tam(it.tamano)) for it in items)
                tarjeta = ("<div style='max-width:1100px;margin:16px auto;padding:16px 20px;background:#fff;border:1px solid #e2e8f0;"
                           "border-radius:14px'><b>Archivos adjuntos del solicitante (%d)</b>"
                           "<ul style='margin:8px 0 0;padding-left:18px'>%s</ul></div>" % (len(items), filas))
                for marca in ("</main>", "</body>"):
                    i = resp.rfind(marca)
                    if i >= 0:
                        return resp[:i] + tarjeta + resp[i:]
                return resp + tarjeta
            except Exception as ex:
                print("PQR SEGURO detalle:", ex)
                return resp
        app.view_functions["soporte_pqr_detalle"] = detalle

    # ---------- exigir PIN + enlazar adjuntos al radicar (chat) ----------
    def vincular(token, radicado_txt, fila, extra_hechos=""):
        """Tras radicar: asocia los adjuntos al ticket, marca el PIN como usado y anota la verificación."""
        digitos = re.sub(r"\D", "", str(radicado_txt or ""))
        t = TicketPQR.query.filter_by(radicado=digitos).first() if digitos else None
        if not t:
            return 0
        n = 0
        for a in PqrAdjunto.query.filter_by(token=token, ticket_id=None).all():
            a.ticket_id = t.id
            n += 1
        nombres = [a.nombre for a in PqrAdjunto.query.filter_by(ticket_id=t.id).all()]
        if nombres:
            t.archivo = ("; ".join(nombres))[:255]
        nota = "Identidad verificada por PIN (%s: %s)." % (
            "SMS" if fila.canal == "sms" else "correo", mask_tel(fila.destino) if fila.canal == "sms" else mask_mail(fila.destino))
        if fila.canal == "sms":
            nota += " Celular: %s." % fila.destino
        t.hechos = ((t.hechos or "") + "\n" + nota + extra_hechos).strip()
        fila.usado = True
        db.session.commit()
        try:
            g["registrar_historial_pqr"](t.id, "VERIFICACION", nota + (" Adjuntos: %d." % n if n else ""), t.estado or "RADICADA")
            db.session.commit()
        except Exception:
            db.session.rollback()
        return n

    orig_chat = app.view_functions.get("pqr_chat_radicar")
    if orig_chat:
        @wraps(orig_chat)
        def chat_radicar(*a, **k):
            if not obligatorio():
                return orig_chat(*a, **k)
            tablas()
            d = request.get_json(silent=True) or {}
            token = str(d.get("pin_token") or "")
            fila = fila_token(token)
            if not fila:
                return err("Verifica tu identidad con el PIN antes de radicar.", 403, codigo="pin")
            if fila.canal == "correo" and fila.destino != str(d.get("email") or "").strip().lower():
                return err("El correo no coincide con el que verificaste. Vuelve a verificar.", 403, codigo="pin")
            resp = orig_chat(*a, **k)
            try:
                r = resp[0] if isinstance(resp, tuple) else resp
                j = json.loads(r.get_data(as_text=True))
                if j.get("ok"):
                    vincular(token, j.get("radicado"), fila)
            except Exception as ex:
                db.session.rollback()
                print("PQR SEGURO vincular chat:", ex)
            return resp
        app.view_functions["pqr_chat_radicar"] = chat_radicar

    # ---------- formulario clásico (/pqr/formulario) ----------
    orig_form = app.view_functions.get("pqr_crear")
    if orig_form:
        @wraps(orig_form)
        def form(*a, **k):
            tablas()
            token, fila = "", None
            if request.method == "POST" and obligatorio():
                token = request.form.get("pin_token", "")
                fila = fila_token(token)
                email = (request.form.get("email") or "").strip().lower()
                if not fila or (fila.canal == "correo" and fila.destino != email):
                    return g["page"]("Radicar PQR", "<div style='max-width:640px;margin:40px auto;font-family:Segoe UI,Arial;padding:20px'>"
                                     "<h2 style='color:#b91c1c'>Falta verificar tu identidad</h2>"
                                     "<p>Verifica el PIN (del mismo correo que escribiste) y vuelve a radicar.</p>"
                                     "<p><a href='/pqr/formulario'>← Volver al formulario</a></p></div>")
            resp = orig_form(*a, **k)
            if not isinstance(resp, str):
                return resp
            if request.method == "GET":
                bloque = _FORM_BLOQUE.replace("__ACCEPT__", ACCEPT_ATTR).replace(
                    "__SMS_OPT__", '<option value="sms">Recibir PIN por SMS</option>' if sms_ok() else "")
                marca = '<button type="submit">Radicar solicitud</button>'
                if resp.count(marca) == 1 and resp.count('<form method="POST">') == 1:
                    deshab = ' disabled title="Verifica tu PIN primero"' if obligatorio() else ""
                    resp = resp.replace('<form method="POST">', '<form method="POST" enctype="multipart/form-data">')
                    resp = resp.replace(marca, bloque + '<button type="submit" id="pqrseg_submit"%s>Radicar solicitud</button>' % deshab)
                return resp
            # POST: enlazar adjuntos y PIN al radicado recién creado
            try:
                m = re.search(r"(\d{4}-\d{2}-\d{6})", resp)
                if m and fila:
                    t = TicketPQR.query.filter_by(radicado=re.sub(r"\D", "", m.group(1))).first()
                    if t:
                        guardar(request.files.getlist("archivos"), token, ticket_id=t.id)
                        vincular(token, m.group(1), fila)
            except Exception as ex:
                db.session.rollback()
                print("PQR SEGURO form:", ex)
            return resp
        app.view_functions["pqr_crear"] = form

    # ---------- parche del chat ----------
    nuevo = parchar_chat_html(g.get("_PQR_CHAT_HTML"))
    if nuevo:
        g["_PQR_CHAT_HTML"] = nuevo
    else:
        print("PQR SEGURO: el chat no se modificó (la app sigue con el flujo anterior).")

    try:
        with app.app_context():
            tablas()
    except Exception as ex:
        print("PQR SEGURO (tablas se crearán al primer uso):", ex)

    g["_PQR_SEGURO_INSTALADO"] = True
    print("PQR SEGURO instalado · PIN %s · SMS %s" % ("obligatorio" if obligatorio() else "opcional", "sí" if sms_ok() else "no (solo correo)"))

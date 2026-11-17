# Incapacidades (Recursos Humanos) + tipografía estilo Apple + menú y diseño adaptados a celulares.
# Uso:  python aplicar_rrhh.py            (trabaja sobre app.py de esta carpeta)
import io, os, re, sys, shutil, py_compile

SRC = sys.argv[1] if len(sys.argv) > 1 else "app.py"
INC = r"""# ════════════════════════════════════════════════════════════════════════════
#  INCAPACIDADES (Recursos Humanos)
#  - Solo Talento Humano registra (datos + foto o PDF del soporte). Gerencia solo visualiza.
#  - Nómina ve días y soporte; SST ve el diagnóstico; el jefe inmediato solo ve fechas (sin datos médicos) y recibe una alerta.
# ════════════════════════════════════════════════════════════════════════════
_INC_INIT = {"ok": False}
_INC_ORIGENES = ("Enfermedad general", "Accidente de trabajo", "Enfermedad laboral", "Licencia de maternidad o paternidad", "Otra")
_INC_ROLES_OPC = ("Finanzas", "Contabilidad", "Operaciones", "Calidad", "Tecnologia", "Marketing", "Juridica", "Comercial", "Soporte")
_INC_MAX_BYTES = 3400 * 1024
_INC_CSS = '''<style>
.inc-fg{display:grid;gap:0 16px;grid-template-columns:repeat(auto-fit,minmax(250px,1fr))}
.inc-fg .full{grid-column:1/-1}
.cb input[type=date],.cb input[type=file]{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:14px;background:#fff}
.cb textarea{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #cfd8f3;border-radius:10px;font-size:14px;font-family:inherit;min-height:84px;resize:vertical}
.inc-dl{display:grid;grid-template-columns:170px 1fr;gap:8px 14px;font-size:14px;margin:10px 0}
.inc-dl dt{color:#64748b}.inc-dl dd{margin:0;font-weight:600;word-break:break-word}
.inc-dias{display:inline-block;background:rgba(0,184,212,.12);border:1px solid rgba(0,184,212,.38);border-radius:999px;padding:4px 12px;font-weight:700;color:#0B1B4D}
.inc-img{max-width:100%;border:1px solid #e3e8f7;border-radius:12px;margin-top:8px}
@media(max-width:640px){.inc-dl{grid-template-columns:1fr}}
</style>'''


def _inc_init():
    if _INC_INIT["ok"]:
        return
    pk = "INTEGER PRIMARY KEY AUTOINCREMENT" if db.engine.dialect.name == "sqlite" else "SERIAL PRIMARY KEY"
    try:
        try:
            _ti_init()
        except Exception:
            pass
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS incapacidades (id %s, codigo VARCHAR(24), usuario VARCHAR(80) DEFAULT '', cedula VARCHAR(30), nombre VARCHAR(160), "
            "cargo VARCHAR(120) DEFAULT '', area VARCHAR(80) DEFAULT '', jefe VARCHAR(80) DEFAULT '', f_inicio VARCHAR(10), f_fin VARCHAR(10), dias INTEGER DEFAULT 0, "
            "origen VARCHAR(60) DEFAULT '', en_instalaciones INTEGER DEFAULT 0, cie10 VARCHAR(12) DEFAULT '', diagnostico VARCHAR(200) DEFAULT '', "
            "medico VARCHAR(160) DEFAULT '', entidad VARCHAR(120) DEFAULT '', observaciones TEXT DEFAULT '', estado VARCHAR(12) DEFAULT 'ACTIVA', "
            "motivo TEXT DEFAULT '', creado_por VARCHAR(80), creado VARCHAR(19), anulado_por VARCHAR(80) DEFAULT '', anulado VARCHAR(19) DEFAULT '')" % pk))
        db.session.execute(text(
            "CREATE TABLE IF NOT EXISTS incapacidades_archivos (id %s, inc_id INTEGER, nombre VARCHAR(160), mime VARCHAR(40), tam INTEGER DEFAULT 0, datos TEXT)" % pk))
        db.session.commit()
        _INC_INIT["ok"] = True
    except Exception as e:
        db.session.rollback()
        print("incapacidades init:", e)


def _inc_roles(clave, defecto):
    try:
        v = _ti_cfg(clave) or defecto
    except Exception:
        v = defecto
    return [x.strip() for x in (v or "").split(",") if x.strip()]


def _inc_perfil(rol):
    rol = (rol or "").strip()
    p = {"rh": rol == "Talento", "gerencia": rol in _DCX_GERENCIA, "nomina": rol in _inc_roles("inc_roles_nomina", "Finanzas,Contabilidad"),
         "sst": rol in _inc_roles("inc_roles_sst", "Operaciones")}
    p["acceso"] = p["rh"] or p["gerencia"] or p["nomina"] or p["sst"]
    p["medico"] = p["rh"] or p["gerencia"] or p["sst"]
    p["soporte"] = p["acceso"]
    return p


def _inc_guard(solo_modulo=True):
    if not requiere_login():
        return None, redirect("/login")
    rol = rol_actual()
    p = _inc_perfil(rol)
    if solo_modulo and not p["acceso"]:
        return None, acceso_denegado("Las incapacidades las gestiona Recursos Humanos; solo las áreas autorizadas pueden consultarlas.")
    try:
        from flask import g as _g_inc
        _g_inc._staff_wrap = True
    except Exception:
        pass
    _inc_init()
    return p, None


def _inc_estado(estado, ini, fin, hoy):
    if estado != "ACTIVA":
        return "Anulada", "off"
    if (fin or "") < hoy:
        return "Finalizada", "off"
    if (ini or "") > hoy:
        return "Próxima", "on"
    return "En curso", "on"


def _inc_fecha(f):
    try:
        d = datetime.strptime(f, "%Y-%m-%d")
        return "%d %s %d" % (d.day, _DCX_MESES[d.month - 1][:3], d.year)
    except Exception:
        return f or ""


def _inc_alertas(usuario):
    '''Alertas para el jefe inmediato: colaboradores a su cargo incapacitados hoy (sin datos médicos).'''
    try:
        _inc_init()
        hoy = fecha_hoy()
        rows = _c48_q("SELECT id, nombre, f_inicio, f_fin, dias FROM incapacidades WHERE jefe=:j AND estado='ACTIVA' AND f_inicio<=:h AND f_fin>=:h", j=usuario, h=hoy)
        return [{"id": "inc%d" % r[0], "titulo": "Colaborador incapacitado: %s" % r[1],
                 "texto": "Desde %s hasta %s (%s días). Reorganiza las tareas del equipo." % (_inc_fecha(r[2]), _inc_fecha(r[3]), r[4]),
                 "url": "/incapacidades/mi-equipo", "nivel": "alerta", "fijo": True} for r in rows]
    except Exception:
        try:
            db.session.rollback()
        except Exception:
            pass
        return []


def _inc_menu_items(rol, usuario):
    out = []
    try:
        if _inc_perfil(rol)["acceso"]:
            out.append(("/talento/incapacidades", "🩺 Incapacidades"))
        tiene = session.get("_inc_equipo")
        if tiene is None:
            tiene = bool(_c48_q("SELECT 1 FROM usuarios WHERE jefe_usuario=:u AND institucion_id IS NULL LIMIT 1", u=usuario)) if usuario else False
            session["_inc_equipo"] = tiene
        if tiene:
            out.append(("/incapacidades/mi-equipo", "👥 Incapacidades de mi equipo"))
    except Exception:
        try:
            db.session.rollback()
        except Exception:
            pass
    return out


def _inc_detectar_archivo(raw):
    if raw[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    if raw[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        return "image/webp"
    if raw[:5] == b"%PDF-":
        return "application/pdf"
    return None


def _inc_usuarios():
    try:
        return _c48_q("SELECT usuario, COALESCE(nombre_completo,''), COALESCE(rol,''), COALESCE(jefe_usuario,'') FROM usuarios "
                      "WHERE institucion_id IS NULL AND COALESCE(activo, TRUE)=TRUE ORDER BY 2, 1")
    except Exception:
        return []


# ───────────────────────── Registro (solo Recursos Humanos)
def _inc_form(csrf):
    us = _inc_usuarios()
    ops = "<option value=''>— Escribir manualmente —</option>" + "".join(
        "<option value='%s'>%s · %s</option>" % (_esc(u), _esc(n or u), _esc(r)) for u, n, r, j in us)
    jops = "<option value=''>— Sin jefe asignado —</option>" + "".join(
        "<option value='%s'>%s</option>" % (_esc(u), _esc(n or u)) for u, n, r, j in us)
    mapa = json.dumps({u: {"n": n or u, "j": j, "c": r} for u, n, r, j in us}, ensure_ascii=False).replace("</", "<\\/")
    orig = "".join("<option>%s</option>" % _esc(o) for o in _INC_ORIGENES)
    hoy = fecha_hoy()
    campos = (
        "<div class='full'><label>Colaborador (cuenta del sistema)</label><select name='usuario' id='inc-usr'>%s</select>"
        "<div class='cb-note'>Al elegirlo se completan el nombre y el jefe inmediato. Si no tiene cuenta, escribe los datos a mano.</div></div>" % ops
        + "<div><label>Nombre completo *</label><input type='text' name='nombre' id='inc-nom' required maxlength='160'></div>"
        + "<div><label>Cédula *</label><input type='text' name='cedula' required maxlength='20' inputmode='numeric' placeholder='Solo números'></div>"
        + "<div><label>Cargo</label><input type='text' name='cargo' maxlength='120'></div>"
        + "<div><label>Jefe inmediato</label><select name='jefe' id='inc-jefe'>%s</select></div>" % jops
        + "<div><label>Fecha de inicio *</label><input type='date' name='f_inicio' id='inc-ini' value='%s' required></div>" % hoy
        + "<div><label>Fecha de finalización *</label><input type='date' name='f_fin' id='inc-fin' value='%s' required></div>" % hoy
        + "<div><label>Días totales</label><div class='inc-dias' id='inc-dias'>1 día</div><div class='cb-note'>Se calculan solos con las fechas.</div></div>"
        + "<div><label>Origen *</label><select name='origen'>%s</select></div>" % orig
        + "<div><label>Código de diagnóstico (CIE-10) *</label><input type='text' name='cie10' required maxlength='8' placeholder='Ej.: J06.9' style='text-transform:uppercase'></div>"
        + "<div><label>Descripción del diagnóstico</label><input type='text' name='diagnostico' maxlength='200'></div>"
        + "<div><label>Nombre del médico *</label><input type='text' name='medico' required maxlength='160'></div>"
        + "<div><label>EPS / IPS</label><input type='text' name='entidad' maxlength='120'></div>"
        + "<div><label style='display:flex;gap:8px;align-items:center;margin-top:26px'><input type='checkbox' name='en_instalaciones' value='1'> El evento ocurrió en las instalaciones de la empresa</label></div>"
        + "<div class='full'><label>Foto o PDF del documento médico * (obligatorio)</label><input type='file' name='soporte' id='inc-file' accept='image/*,application/pdf' required>"
        "<div class='cb-note' id='inc-info'>JPG, PNG o PDF. Máximo 3,4 MB: las fotos grandes se reducen solas antes de enviarse.</div></div>"
        + "<div class='full'><label>Observaciones</label><textarea name='observaciones' rows='3'></textarea></div>")
    js = r'''<script>
(function(){
var M=__MAPA__;
var usr=document.getElementById("inc-usr"),nom=document.getElementById("inc-nom"),jf=document.getElementById("inc-jefe");
usr.addEventListener("change",function(){var d=M[usr.value];if(!d)return;nom.value=d.n;jf.value=d.j||"";});
var ini=document.getElementById("inc-ini"),fin=document.getElementById("inc-fin"),out=document.getElementById("inc-dias");
function dias(){var a=new Date(ini.value),b=new Date(fin.value);if(isNaN(a)||isNaN(b)){out.textContent="—";return;}var n=Math.round((b-a)/86400000)+1;out.textContent=n<1?"Revisa las fechas":(n+(n===1?" día":" días"));}
ini.addEventListener("change",function(){if(fin.value<ini.value)fin.value=ini.value;dias();});fin.addEventListener("change",dias);dias();
var fi=document.getElementById("inc-file"),info=document.getElementById("inc-info");
fi.addEventListener("change",async function(){
  var f=fi.files[0];if(!f||!/^image\//.test(f.type)||f.size<1200000)return;
  try{
    var bmp=await createImageBitmap(f),m=Math.min(1,1800/Math.max(bmp.width,bmp.height)),c=document.createElement("canvas");
    c.width=Math.round(bmp.width*m);c.height=Math.round(bmp.height*m);c.getContext("2d").drawImage(bmp,0,0,c.width,c.height);
    c.toBlob(function(b){if(!b)return;var nf=new File([b],f.name.replace(/\.\w+$/,"")+".jpg",{type:"image/jpeg"});var dt=new DataTransfer();dt.items.add(nf);fi.files=dt.files;info.textContent="Foto optimizada: "+Math.round(b.size/1024)+" KB";},"image/jpeg",0.82);
  }catch(e){}
});
})();
</script>'''.replace("__MAPA__", mapa)
    return ("<h1>Registrar incapacidad</h1><p class='sub'>Se registran los datos del documento médico y se adjunta su foto o PDF (soporte legal para el recobro a la EPS). "
            "Nómina, SST y el jefe inmediato la verán según su permiso.</p>"
            "<form method='POST' enctype='multipart/form-data' action='/talento/incapacidades/nueva'><input type='hidden' name='csrf' value='%s'>"
            "<div class='cb-card'><div class='inc-fg'>%s</div><p style='margin-top:14px'><button class='cb-btn g'>Guardar incapacidad</button> "
            "<a class='cb-btn o' href='/talento/incapacidades'>Cancelar</a></p></div></form>%s" % (_esc(csrf), campos, js))


def _inc_crear(f, files, usuario):
    def lim(k, n):
        return " ".join((f.get(k) or "").split())[:n]
    nombre = lim("nombre", 160)
    cedula = re.sub(r"\D", "", f.get("cedula") or "")[:20]
    ini, fin = (f.get("f_inicio") or "").strip(), (f.get("f_fin") or "").strip()
    cie = lim("cie10", 8).upper().replace(" ", "")
    medico = lim("medico", 160)
    if not nombre or not cedula or not medico:
        return False, "Faltan datos obligatorios: nombre, cédula o médico."
    if not (re.match(r"^\d{4}-\d{2}-\d{2}$", ini) and re.match(r"^\d{4}-\d{2}-\d{2}$", fin)):
        return False, "Revisa las fechas."
    try:
        dias = (datetime.strptime(fin, "%Y-%m-%d") - datetime.strptime(ini, "%Y-%m-%d")).days + 1
    except Exception:
        return False, "Revisa las fechas."
    if dias < 1 or dias > 540:
        return False, "La fecha final no puede ser anterior a la inicial (máximo 540 días)."
    if not re.match(r"^[A-Z]\d{2}(\.\d{1,2})?$", cie):
        return False, "El código CIE-10 no es válido (ejemplo: J06.9)."
    origen = lim("origen", 60)
    if origen not in _INC_ORIGENES:
        origen = "Otra"
    arch = files.get("soporte") if files else None
    if arch is None or not getattr(arch, "filename", ""):
        return False, "Adjunta la foto o el PDF del documento médico: es obligatorio."
    raw = arch.read(_INC_MAX_BYTES + 1)
    if not raw:
        return False, "El archivo está vacío."
    if len(raw) > _INC_MAX_BYTES:
        return False, "El archivo pesa más de 3,4 MB. Toma la foto con menor resolución o usa un PDF más liviano."
    mime = _inc_detectar_archivo(raw)
    if not mime:
        return False, "Formato no compatible. Usa JPG, PNG o PDF."
    if _c48_q("SELECT 1 FROM incapacidades WHERE cedula=:c AND estado='ACTIVA' AND f_inicio<=:f AND f_fin>=:i", c=cedula, f=fin, i=ini):
        return False, "Ese colaborador ya tiene una incapacidad activa que se cruza con esas fechas."
    anio = ini[:4]
    n = int((_c48_q("SELECT COUNT(*) FROM incapacidades WHERE codigo LIKE :p", p="INC-%s-%%" % anio)[0][0]) or 0) + 1
    codigo = "INC-%s-%04d" % (anio, n)
    usr = lim("usuario", 80)
    jefe = lim("jefe", 80)
    area = ""
    if usr:
        r = _c48_q("SELECT COALESCE(departamento,''), COALESCE(rol,''), COALESCE(jefe_usuario,'') FROM usuarios WHERE usuario=:u AND institucion_id IS NULL", u=usr)
        if r:
            area = r[0][0] or r[0][1]
            jefe = jefe or r[0][2]
    _c48_x("INSERT INTO incapacidades (codigo, usuario, cedula, nombre, cargo, area, jefe, f_inicio, f_fin, dias, origen, en_instalaciones, cie10, diagnostico, "
           "medico, entidad, observaciones, estado, creado_por, creado) VALUES (:codigo,:usuario,:cedula,:nombre,:cargo,:area,:jefe,:ini,:fin,:dias,:origen,:inst,"
           ":cie,:diag,:medico,:ent,:obs,'ACTIVA',:us,:cr)",
           codigo=codigo, usuario=usr, cedula=cedula, nombre=nombre, cargo=lim("cargo", 120), area=area, jefe=jefe, ini=ini, fin=fin, dias=dias, origen=origen,
           inst=1 if f.get("en_instalaciones") else 0, cie=cie, diag=lim("diagnostico", 200), medico=medico, ent=lim("entidad", 120),
           obs=(f.get("observaciones") or "").strip()[:2000], us=usuario, cr=str(ahora())[:19])
    r = _c48_q("SELECT id FROM incapacidades WHERE codigo=:c", c=codigo)
    if not r:
        return False, "No se pudo guardar la incapacidad. Intenta de nuevo."
    iid = int(r[0][0])
    import base64 as _b64
    _c48_x("INSERT INTO incapacidades_archivos (inc_id, nombre, mime, tam, datos) VALUES (:i,:n,:m,:t,:d)", i=iid,
           n=re.sub(r"[^A-Za-z0-9._-]", "_", arch.filename or "soporte")[:150], m=mime, t=len(raw), d=_b64.b64encode(raw).decode("ascii"))
    registrar_auditoria("Incapacidades", "%s registró %s (%s, %d días)" % (usuario, codigo, nombre, dias))
    return True, iid


@app.route("/talento/incapacidades/nueva", methods=["GET", "POST"])
def inc_nueva():
    p, r = _inc_guard()
    if r is not None:
        return r
    if not p["rh"]:
        return acceso_denegado("Solo Recursos Humanos registra incapacidades. Los demás perfiles las visualizan.")
    csrf = _acc_csrf()
    if request.method == "POST":
        if not _ti_csrf_ok():
            session["ti_msg"] = ("err", "La sesión del formulario venció. Intenta de nuevo.")
            return redirect("/talento/incapacidades/nueva")
        try:
            ok, res = _inc_crear(request.form, request.files, session.get("usuario") or "")
        except Exception as e:
            db.session.rollback()
            print("incapacidad crear:", e)
            ok, res = False, "No se pudo guardar. Intenta de nuevo."
        if ok:
            session["ti_msg"] = ("ok", "Incapacidad registrada. Nómina, SST y el jefe inmediato ya pueden consultarla.")
            return redirect("/talento/incapacidades/%d" % res)
        session["ti_msg"] = ("err", res)
        return redirect("/talento/incapacidades/nueva")
    return _cb_pagina("Registrar incapacidad", _INC_CSS + _inc_form(csrf))


# ───────────────────────── Consulta
@app.route("/talento/incapacidades")
def inc_lista():
    p, r = _inc_guard()
    if r is not None:
        return r
    hoy = fecha_hoy()
    q = (request.args.get("q") or "").strip()[:60]
    ver = (request.args.get("ver") or "").strip()
    cond, par = ["1=1"], {}
    if q:
        cond.append("(lower(nombre) LIKE :q OR cedula LIKE :q OR lower(codigo) LIKE :q)")
        par["q"] = "%" + q.lower() + "%"
    if ver == "curso":
        cond.append("estado='ACTIVA' AND f_inicio<=:h AND f_fin>=:h")
        par["h"] = hoy
    elif ver == "anuladas":
        cond.append("estado<>'ACTIVA'")
    rows = _c48_q("SELECT id, codigo, nombre, cedula, f_inicio, f_fin, dias, origen, cie10, estado FROM incapacidades WHERE " + " AND ".join(cond)
                  + " ORDER BY f_inicio DESC, id DESC LIMIT 300", **par)
    filas = ""
    for iid, cod, nom, ced, ini, fin, dias, org, cie, est in rows:
        et, cl = _inc_estado(est, ini, fin, hoy)
        filas += ("<tr><td><b>%s</b></td><td>%s<br><small>CC %s</small></td><td>%s → %s</td><td>%s</td><td>%s</td>%s<td><span class='cb-tag %s'>%s</span></td>"
                  "<td><a class='cb-btn s' href='/talento/incapacidades/%d'>Ver</a></td></tr>"
                  % (_esc(cod), _esc(nom), _esc(ced), _esc(_inc_fecha(ini)), _esc(_inc_fecha(fin)), dias, _esc(org),
                     ("<td>%s</td>" % _esc(cie)) if p["medico"] else "", cl, et, iid))
    if not filas:
        filas = "<tr><td colspan='8'>No hay incapacidades registradas con ese criterio.</td></tr>"
    nuevo = "<a class='cb-btn g' href='/talento/incapacidades/nueva'>➕ Registrar incapacidad</a> " if p["rh"] else ""
    tabs = "".join("<a class='cb-btn s %s' href='/talento/incapacidades%s'>%s</a> " % ("" if ver == k else "o", ("?ver=" + k) if k else "", n)
                   for k, n in (("", "Todas"), ("curso", "En curso hoy"), ("anuladas", "Anuladas")))
    buscar = ("<form method='GET' action='/talento/incapacidades' style='display:inline-flex;gap:6px'><input type='text' name='q' value='%s' "
              "placeholder='Nombre, cédula o código' style='width:210px'><button class='cb-btn s'>Buscar</button></form>" % _esc(q))
    perm = ""
    if p["gerencia"] or p["rh"]:
        csrf = _acc_csrf()
        nom_sel = _inc_roles("inc_roles_nomina", "Finanzas,Contabilidad")
        sst_sel = _inc_roles("inc_roles_sst", "Operaciones")
        if p["gerencia"]:
            def checks(nombre, sel):
                return "".join("<label style='display:inline-flex;gap:6px;align-items:center;margin:4px 14px 4px 0;font-weight:500'>"
                               "<input type='checkbox' name='%s' value='%s'%s> %s</label>" % (nombre, o, " checked" if o in sel else "", _esc(o))
                               for o in _INC_ROLES_OPC)
            perm = ("<div class='cb-card'><h2>¿Quién visualiza?</h2><p class='cb-note'>Recursos Humanos registra; estas áreas solo consultan. El jefe inmediato ve fechas "
                    "y recibe una alerta, sin datos médicos.</p><form method='POST' action='/talento/incapacidades/permisos'><input type='hidden' name='csrf' value='%s'>"
                    "<p><b>Nómina</b> (días y soporte)<br>%s</p><p><b>Seguridad y Salud en el Trabajo (SST)</b> (diagnóstico y soporte)<br>%s</p>"
                    "<button class='cb-btn'>Guardar permisos</button></form></div>" % (_esc(csrf), checks("nomina", nom_sel), checks("sst", sst_sel)))
    cuerpo = (_INC_CSS + "<h1>🩺 Incapacidades</h1><p class='sub'>Recursos Humanos registra los datos y el soporte médico. Tu perfil: <b>%s</b>.</p>"
              "<div class='cb-card'><div style='display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px'>%s%s%s</div><div class='cb-scroll'><table>"
              "<tr><th>Código</th><th>Colaborador</th><th>Periodo</th><th>Días</th><th>Origen</th>%s<th>Estado</th><th></th></tr>%s</table></div></div>%s"
              % ("Recursos Humanos" if p["rh"] else ("Gerencia (solo visualiza)" if p["gerencia"] else ("SST" if p["sst"] else "Nómina")),
                 nuevo, tabs, buscar, "<th>CIE-10</th>" if p["medico"] else "", filas, perm))
    return _cb_pagina("Incapacidades", cuerpo)


@app.route("/talento/incapacidades/permisos", methods=["POST"])
def inc_permisos():
    p, r = _inc_guard()
    if r is not None:
        return r
    if not p["gerencia"] or not _ti_csrf_ok():
        return acceso_denegado("Solo la Gerencia define quién visualiza las incapacidades.")
    for campo, clave in (("nomina", "inc_roles_nomina"), ("sst", "inc_roles_sst")):
        vals = [x for x in request.form.getlist(campo) if x in _INC_ROLES_OPC]
        _ti_cfg_set(clave, ",".join(vals) if vals else "-")
    registrar_auditoria("Incapacidades", "%s actualizó quién visualiza" % (session.get("usuario") or ""))
    session["ti_msg"] = ("ok", "Permisos guardados.")
    return redirect("/talento/incapacidades")


@app.route("/talento/incapacidades/<int:iid>")
def inc_ver(iid):
    p, r = _inc_guard()
    if r is not None:
        return r
    rows = _c48_q("SELECT id, codigo, nombre, cedula, cargo, area, jefe, f_inicio, f_fin, dias, origen, en_instalaciones, cie10, diagnostico, medico, entidad, "
                  "observaciones, estado, motivo, creado_por, creado, anulado_por, anulado FROM incapacidades WHERE id=:i", i=iid)
    if not rows:
        return acceso_denegado("Incapacidad no encontrada.")
    (_i, cod, nom, ced, cargo, area, jefe, ini, fin, dias, org, inst, cie, diag, med, ent, obs, est, motivo, cpor, cre, apor, anu) = rows[0]
    et, cl = _inc_estado(est, ini, fin, fecha_hoy())
    arch = _c48_q("SELECT id, mime FROM incapacidades_archivos WHERE inc_id=:i ORDER BY id LIMIT 1", i=iid)
    filas = [("Colaborador", nom), ("Cédula", ced), ("Cargo", cargo), ("Área", area), ("Jefe inmediato", jefe), ("Periodo", "%s al %s" % (_inc_fecha(ini), _inc_fecha(fin))),
             ("Días totales", "%s" % dias), ("Origen", org + (" · ocurrió en las instalaciones" if inst else "")), ("EPS / IPS", ent)]
    if p["medico"]:
        filas += [("Diagnóstico (CIE-10)", ("%s %s" % (cie, diag)).strip()), ("Médico", med)]
    if p["rh"] or p["gerencia"]:
        filas += [("Observaciones", obs), ("Registrada por", "%s · %s" % (cpor, (cre or "")[:16]))]
    dl = "".join("<dt>%s</dt><dd>%s</dd>" % (_esc(a), _esc(b) if b else "—") for a, b in filas)
    soporte = ""
    if arch and p["soporte"]:
        aid, mime = arch[0]
        if mime == "application/pdf":
            soporte = "<p><a class='cb-btn' target='_blank' rel='noopener' href='/talento/incapacidades/archivo/%d'>📄 Abrir PDF del soporte</a></p>" % iid
        else:
            soporte = ("<p><a target='_blank' rel='noopener' href='/talento/incapacidades/archivo/%d'><img class='inc-img' src='/talento/incapacidades/archivo/%d' "
                       "alt='Soporte médico'></a></p>" % (iid, iid))
    anular = ""
    if p["rh"] and est == "ACTIVA":
        anular = ("<form method='POST' action='/talento/incapacidades/anular' style='display:flex;gap:6px;flex-wrap:wrap;margin-top:12px' "
                  "onsubmit=\"return confirm('¿Anular esta incapacidad?');\"><input type='hidden' name='csrf' value='%s'><input type='hidden' name='id' value='%d'>"
                  "<input type='text' name='motivo' placeholder='Motivo de la anulación' required maxlength='200' style='width:260px'><button class='cb-btn r s'>Anular</button></form>"
                  % (_esc(_acc_csrf()), iid))
    nota = ("<p class='cb-note'>Anulada por %s · %s: %s</p>" % (_esc(apor), _esc((anu or "")[:16]), _esc(motivo))) if est != "ACTIVA" else ""
    cuerpo = (_INC_CSS + "<h1>%s</h1><p class='sub'>Incapacidad · <span class='cb-tag %s'>%s</span> · <span class='inc-dias'>%s días</span></p>"
              "<div class='cb-card'><dl class='inc-dl'>%s</dl>%s%s%s</div><p><a class='cb-btn o' href='/talento/incapacidades'>← Volver</a></p>"
              % (_esc(cod), cl, et, dias, dl, soporte, nota, anular))
    return _cb_pagina(cod, cuerpo)


@app.route("/talento/incapacidades/archivo/<int:iid>")
def inc_archivo(iid):
    p, r = _inc_guard()
    if r is not None:
        return r
    if not p["soporte"]:
        return acceso_denegado("No tienes permiso para ver el soporte médico.")
    rows = _c48_q("SELECT mime, datos FROM incapacidades_archivos WHERE inc_id=:i ORDER BY id LIMIT 1", i=iid)
    if not rows:
        return Response(status=404)
    import base64 as _b64
    try:
        raw = _b64.b64decode(rows[0][1])
    except Exception:
        return Response(status=404)
    registrar_auditoria("Incapacidades", "%s abrió el soporte de la incapacidad #%d" % (session.get("usuario") or "", iid))
    resp = Response(raw, mimetype=rows[0][0])
    resp.headers["Cache-Control"] = "private, no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Content-Disposition"] = "inline"
    return resp


@app.route("/talento/incapacidades/anular", methods=["POST"])
def inc_anular():
    p, r = _inc_guard()
    if r is not None:
        return r
    if not p["rh"] or not _ti_csrf_ok():
        return acceso_denegado("Solo Recursos Humanos anula incapacidades.")
    try:
        iid = int(request.form.get("id") or 0)
    except Exception:
        iid = 0
    motivo = " ".join((request.form.get("motivo") or "").split())[:200]
    if not iid or not motivo:
        session["ti_msg"] = ("err", "No se pudo anular.")
        return redirect("/talento/incapacidades")
    _c48_x("UPDATE incapacidades SET estado='ANULADA', motivo=:m, anulado_por=:u, anulado=:a WHERE id=:i AND estado='ACTIVA'",
           m=motivo, u=session.get("usuario") or "", a=str(ahora())[:19], i=iid)
    registrar_auditoria("Incapacidades", "%s anuló la incapacidad #%d: %s" % (session.get("usuario") or "", iid, motivo))
    session["ti_msg"] = ("ok", "Incapacidad anulada.")
    return redirect("/talento/incapacidades/%d" % iid)


# ───────────────────────── Jefe inmediato (sin datos médicos)
@app.route("/incapacidades/mi-equipo")
def inc_mi_equipo():
    p, r = _inc_guard(solo_modulo=False)
    if r is not None:
        return r
    usuario = session.get("usuario") or ""
    hoy = fecha_hoy()
    rows = _c48_q("SELECT nombre, f_inicio, f_fin, dias, estado FROM incapacidades WHERE jefe=:j AND estado='ACTIVA' AND f_fin>=:h ORDER BY f_inicio", j=usuario, h=hoy)
    filas = ""
    for nom, ini, fin, dias, est in rows:
        et, cl = _inc_estado(est, ini, fin, hoy)
        filas += "<tr><td><b>%s</b></td><td>%s → %s</td><td>%s</td><td><span class='cb-tag %s'>%s</span></td></tr>" % (_esc(nom), _esc(_inc_fecha(ini)), _esc(_inc_fecha(fin)), dias, cl, et)
    if not filas:
        filas = "<tr><td colspan='4'>Ningún colaborador a tu cargo está incapacitado ni tiene una incapacidad próxima.</td></tr>"
    cuerpo = (_INC_CSS + "<h1>👥 Incapacidades de mi equipo</h1><p class='sub'>Solo ves las fechas, para reorganizar las tareas. Los datos médicos son confidenciales y los maneja Recursos Humanos.</p>"
              "<div class='cb-card'><div class='cb-scroll'><table><tr><th>Colaborador</th><th>Periodo</th><th>Días</th><th>Estado</th></tr>%s</table></div></div>" % filas)
    return _cb_pagina("Incapacidades de mi equipo", cuerpo)


"""
MOVIL = r"""_MOVIL_SNIPPET = r'''
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,400..800&display=swap">
<style id="mv-css">
/* ── Tipografía estilo Apple: tamaño óptico, tracking ajustado y números tabulares ── */
html{-webkit-text-size-adjust:100%;text-size-adjust:100%}
body{letter-spacing:-.011em;font-feature-settings:"kern","liga","calt";font-optical-sizing:auto;text-rendering:optimizeLegibility}
h1,h2,h3,h4{font-weight:700;line-height:1.14;letter-spacing:-.024em;text-wrap:balance}
h1{letter-spacing:-.03em}
h3,h4{letter-spacing:-.017em}
p,li{text-wrap:pretty}
label{font-weight:600;letter-spacing:-.006em}
small,.cb-note,.sub{letter-spacing:-.006em}
table{font-variant-numeric:tabular-nums}
button,input,select,textarea{letter-spacing:-.01em}
::selection{background:rgba(0,184,212,.28)}
img,video{max-width:100%}
/* ── Menú lateral deslizable en celulares y tabletas ── */
#mb-burger{display:none;width:40px;height:40px;border-radius:12px;border:1px solid rgba(0,32,194,.2);background:rgba(255,255,255,.88);align-items:center;justify-content:center;flex-direction:column;gap:4px;cursor:pointer;padding:0;flex:0 0 auto}
#mb-burger span{display:block;width:18px;height:2px;border-radius:2px;background:#0020C2;transition:transform .2s ease,opacity .2s ease}
html.mb-open #mb-burger span:nth-child(1){transform:translateY(6px) rotate(45deg)}
html.mb-open #mb-burger span:nth-child(2){opacity:0}
html.mb-open #mb-burger span:nth-child(3){transform:translateY(-6px) rotate(-45deg)}
@media(max-width:900px){
  html,body{overflow-x:hidden}
  #mb-burger{display:inline-flex}
  .app-layout-enterprise{display:block!important;width:100%!important}
  .sidebar-apple-glass{position:fixed!important;left:0!important;top:0!important;bottom:0!important;width:min(86vw,320px)!important;height:100%!important;min-height:0!important;max-height:100%!important;transform:translate3d(-104%,0,0);transition:transform .3s cubic-bezier(.22,.61,.36,1);z-index:100001!important;padding-top:calc(18px + env(safe-area-inset-top,0px))!important;padding-bottom:calc(18px + env(safe-area-inset-bottom,0px))!important;will-change:transform}
  html.mb-open .sidebar-apple-glass{transform:translate3d(0,0,0);box-shadow:0 0 70px rgba(0,0,40,.5)}
  #mb-backdrop{position:fixed;inset:0;background:rgba(6,12,50,.5);opacity:0;pointer-events:none;transition:opacity .25s ease;z-index:100000;-webkit-backdrop-filter:blur(2px);backdrop-filter:blur(2px)}
  html.mb-open #mb-backdrop{opacity:1;pointer-events:auto}
  html.mb-open{overflow:hidden}
  .main-content-container{margin-left:0!important;width:100%!important;max-width:100%!important;min-width:0!important}
  .navbar-top-apple{padding:0 12px!important;padding-top:env(safe-area-inset-top,0px)!important;gap:8px!important;position:sticky!important;top:0;z-index:9990}
  .navbar-top-apple .navbar-module-title{font-size:15px}
  .bento-workspace-apple{padding:14px 12px 28px!important}
  .bento-workspace-apple [style*="grid-template-columns"],.bento-workspace-apple .cb-grid,.bento-workspace-apple .kpi-grid,.bento-workspace-apple .dashboard-grid{grid-template-columns:repeat(auto-fit,minmax(min(100%,260px),1fr))!important}
  .bento-workspace-apple table{display:block;max-width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}
  .bento-workspace-apple img,.bento-workspace-apple canvas{max-width:100%;height:auto}
  input,select,textarea{font-size:16px!important;max-width:100%}
  .btn,button,.cb-btn,.btn-apple-primary,.btn-apple-secondary{min-height:42px}
  .cf-in{grid-template-columns:1fr 1fr!important}
}
@media(max-width:640px){
  .navbar-top-apple .badge-env,.navbar-top-apple .nav-user-fullname,.navbar-top-apple .badge-periodo,.navbar-top-apple .nav-api-ping,.navbar-top-apple .nav-ntp-clock{display:none!important}
  .navbar-top-apple .nb-brand span{display:none}
  .bento-workspace-apple h1,.cb h1{font-size:23px!important}
}
@media(max-width:480px){.cf-in{grid-template-columns:1fr!important}}
@media(prefers-reduced-motion:reduce){.sidebar-apple-glass,#mb-backdrop{transition:none!important}}
</style>
<script>
(function(){
  var sb=document.querySelector(".sidebar-apple-glass"),bar=document.querySelector(".navbar-top-apple .navbar-left");
  if(!sb||!bar||document.getElementById("mb-burger"))return;
  var root=document.documentElement,b=document.createElement("button");
  b.id="mb-burger";b.type="button";b.setAttribute("aria-label","Abrir menú");b.setAttribute("aria-expanded","false");
  b.innerHTML="<span></span><span></span><span></span>";bar.insertBefore(b,bar.firstChild);
  var bd=document.createElement("div");bd.id="mb-backdrop";document.body.appendChild(bd);
  function abrir(v){root.classList.toggle("mb-open",v);b.setAttribute("aria-expanded",v?"true":"false");}
  b.addEventListener("click",function(){abrir(!root.classList.contains("mb-open"));});
  bd.addEventListener("click",function(){abrir(false);});
  sb.addEventListener("click",function(e){if(e.target.closest&&e.target.closest("a"))abrir(false);});
  document.addEventListener("keydown",function(e){if(e.key==="Escape")abrir(false);});
  window.addEventListener("resize",function(){if(window.innerWidth>900)abrir(false);});
})();
</script>
'''


"""

def salir(msg, code=1):
    print(msg)
    sys.exit(code)

with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
    s = fh.read()
NL = "\r\n" if "\r\n" in s else "\n"

def crlf(t):
    return t.replace("\r\n", "\n").replace("\n", NL)

def una_vez(ancla, donde=None):
    base = s if donde is None else donde
    if base.count(ancla) != 1:
        salir("ERROR: punto de inserción no encontrado (%r: hay %d). No se modificó nada." % (ancla[:70], base.count(ancla)))

if "_INC_INIT = {" in s and "_MOVIL_SNIPPET = r" in s and "{_MOVIL_SNIPPET}" in s and "_inc_alertas(usuario)" in s:
    salir("Este app.py YA tiene incapacidades y el ajuste para celulares. No se modificó nada.", 0)
if "_DCX_GERENCIA" not in s or "def _dcx_rol_ok" not in s:
    salir("ERROR: falta el módulo de documentos oficiales. Primero ejecuta:  python aplicar_sitio.py  y luego repite este comando. No se modificó nada.")
A_MAIN = 'if __name__ == "__main__":'
una_vez(A_MAIN)

# 1) Módulo de incapacidades
if "_INC_INIT = {" not in s:
    s = s.replace(A_MAIN, crlf(INC) + A_MAIN, 1)

# 2) Menús: Gerencia y cada área (según su permiso o si es jefe inmediato)
A_GER = "        ('/gerencia/talento-legal', 'Talento / legal laboral'),"
if "('/talento/incapacidades'" not in s:
    una_vez(A_GER)
    s = s.replace(A_GER, A_GER + NL + "        ('/talento/incapacidades', '🩺 Incapacidades (RR. HH.)'),", 1)
A3 = '    if rol in _ROLES_MAPA and not any(h == "/mapa" for h, _l in items):'
if "_inc_menu_items(rol" not in s:
    una_vez(A3)
    MENU = """    for _ih, _il in _inc_menu_items(rol, session.get("usuario") or ""):
        if not any(h == _ih for h, _x in items):
            items = [x for x in items if x[0] != "/logout"] + [(_ih, _il)] + [x for x in items if x[0] == "/logout"]
"""
    s = s.replace(A3, crlf(MENU) + A3, 1)

# 3) Alerta al jefe inmediato en la campana de notificaciones
if "_inc_alertas(usuario)" not in s.split("def notificaciones_pendientes():", 1)[-1][:4000]:
    a = s.index("def notificaciones_pendientes():")
    b = s.index('@app.route("/notificaciones/leer/<int:nid>"', a)
    seg = s[a:b]
    ANC = '        return jsonify({"items": items[:12]})'
    if seg.count(ANC) != 1:
        salir("ERROR: no encontré la lista de notificaciones. No se modificó nada.")
    seg = seg.replace(ANC, "        items = _inc_alertas(usuario) + items" + NL + ANC, 1)
    s = s[:a] + seg + s[b:]

# 4) Letra estilo Apple y diseño para celulares (en todas las páginas)
A_PAGE = "def page(title, body):"
if "_MOVIL_SNIPPET = r" not in s:
    una_vez(A_PAGE)
    s = s.replace(A_PAGE, crlf(MOVIL) + A_PAGE, 1)
A_END = "{_bfcache_fix}{_n_guard}</body></html>"
if "{_MOVIL_SNIPPET}" not in s:
    una_vez(A_END)
    s = s.replace(A_END, "{_MOVIL_SNIPPET}" + A_END, 1)

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
print("LISTO. Incapacidades y ajuste para celulares instalados. Líneas:", len(s.splitlines()))
print("Copia de seguridad en:", SRC + ".respaldo")

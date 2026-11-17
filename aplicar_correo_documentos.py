#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_correo_documentos.py  ·  PROCSIS: arregla el correo de «Documentos por área» y la ventana flotante del selector.

Requiere: aplicar_documentos_area2.py ya aplicado.
Uso (en la carpeta de tu app.py):
    python aplicar_correo_documentos.py            # parchea ./app.py
    python aplicar_correo_documentos.py ruta/app.py

1) CORREO: antes llegaba sin el PDF adjunto y con remitente «EduTrack». Ahora:
     · el PDF viaja adjunto de verdad (DOC-XXX-2026-0001.pdf);
     · el remitente es el área que emite («Gerencia General PROCSIS», «Recursos Humanos PROCSIS»…) y las respuestas van al correo de PROCSIS;
     · el cuerpo es un correo corporativo: logo, saludo, tu mensaje, ficha del documento (código, tipo, asunto, fecha, emitido por, firmado por),
       botón «Verificar autenticidad», firma del área con contacto y aviso de confidencialidad.
     · si el envío falla, la pantalla te dice el motivo real (por ejemplo, dominio no verificado en Resend).
2) VENTANA «PRODUCTO» flotante: el selector Todo PROCSIS / EduTrack / TiendaTrack vive SOLO dentro del menú lateral. En páginas sin menú
   (como la de verificación del QR) ya no aparece, y el botón de ayuda tampoco sale en las páginas públicas /docs/.
Es idempotente, crea app.py.bak_correo_documentos y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-correo-docs]"
REQUIERE = "[procsis-documentos-area-v2]"

BLOQUE = r'''
# [procsis-correo-docs]
# ═════════════ Documentos por área: correo corporativo con PDF adjunto de verdad y remitente = área ═════════════
def _dca_nombre_remitente(entidad):
    n = re.sub(r"[<>\",;@\r\n]", " ", str(entidad or "PROCSIS"))
    return " ".join(n.split())[:80] or "PROCSIS"


def _dca_correo_html(d, mensaje, nombre_pdf):
    emp = _dcx_empresa()
    url = _dca_url(d)
    logo = "%s%s" % (_dcx_base(), _DCX_LOGO)
    msj = _esc(mensaje).replace("\n", "<br>")
    filas = [("Código", d.get("codigo")), ("Tipo", _dca_cat_txt(d.get("categoria"))), ("Asunto", d.get("asunto")),
             ("Fecha", _dcx_fecha_larga(d.get("fecha"))), ("Emitido por", d.get("entidad")), ("Firmado por", d.get("firmante"))]
    fichas = "".join("<tr><td style='padding:6px 0;color:#64748b;font-size:13px;width:118px;vertical-align:top'>%s</td>"
                     "<td style='padding:6px 0;color:#0b1437;font-size:13px;font-weight:600'>%s</td></tr>" % (_esc(k), _esc(v or "—")) for k, v in filas)
    cargo = ("%s<br>" % _esc(d.get("cargo"))) if d.get("cargo") else ""
    return (
        "<!doctype html><html><body style='margin:0;padding:0;background:#f3f6fc'>"
        "<table role='presentation' width='100%%' cellpadding='0' cellspacing='0' style='background:#f3f6fc;padding:22px 10px'><tr><td align='center'>"
        "<table role='presentation' width='600' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%%;background:#ffffff;border-radius:14px;overflow:hidden;"
        "border:1px solid #e3e8f7;font-family:Segoe UI,Arial,sans-serif;color:#0b1437'>"
        "<tr><td style='height:6px;background:#0020C2;background-image:linear-gradient(100deg,#00E676,#00B8D4 32%%,#0020C2);font-size:0;line-height:0'>&nbsp;</td></tr>"
        "<tr><td style='padding:22px 30px 6px'><img src='%s' alt='PROCSIS' height='46' style='height:46px;width:auto;border:0;display:block'></td></tr>"
        "<tr><td style='padding:10px 30px 4px;font-size:15px;line-height:1.65'><p style='margin:0 0 12px'>Cordial saludo,</p><p style='margin:0 0 14px'>%s</p></td></tr>"
        "<tr><td style='padding:0 30px'><table role='presentation' width='100%%' cellpadding='0' cellspacing='0' style='background:#f0f7ff;border:1px solid #cfe0ff;border-radius:12px'>"
        "<tr><td style='padding:14px 18px'><div style='font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#0020C2;font-weight:700;margin-bottom:6px'>"
        "Documento adjunto</div><table role='presentation' width='100%%' cellpadding='0' cellspacing='0'>%s</table>"
        "<div style='margin-top:8px;font-size:13px;color:#334155'>&#128206; <b>%s</b></div></td></tr></table></td></tr>"
        "<tr><td align='left' style='padding:18px 30px 6px'><a href='%s' style='display:inline-block;background:#0020C2;color:#ffffff;text-decoration:none;font-weight:700;"
        "font-size:14px;padding:11px 20px;border-radius:10px'>Verificar autenticidad</a></td></tr>"
        "<tr><td style='padding:16px 30px 8px;font-size:14px;line-height:1.6'>Atentamente,<br><br><b>%s</b><br>%s<span style='color:#0020C2;font-weight:700'>%s</span></td></tr>"
        "<tr><td style='padding:10px 30px 20px;font-size:12px;color:#475569;border-top:1px solid #e6ecf8;line-height:1.6'>"
        "Tel: %s &nbsp;·&nbsp; %s &nbsp;·&nbsp; %s</td></tr>"
        "<tr><td style='padding:14px 30px 22px;background:#f8fafc;font-size:11px;color:#64748b;line-height:1.55'>"
        "Este mensaje y sus anexos son confidenciales y están dirigidos exclusivamente a su destinatario. Si lo recibió por error, por favor avise al remitente y elimínelo. "
        "Puede comprobar que el documento es auténtico con el botón «Verificar autenticidad» o con el código QR del PDF.</td></tr>"
        "</table></td></tr></table></body></html>"
    ) % (_esc(logo), msj, fichas, _esc(nombre_pdf), _esc(url), _esc(d.get("firmante")), cargo, _esc(d.get("entidad")),
         _esc(emp["tel"]), _esc(emp["email"]), _esc(emp["web"]))


def _dca_enviar(d, para, cc, asunto, mensaje, usuario):
    """Envía el documento en PDF adjunto. Devuelve (ok, texto para el usuario)."""
    import base64
    import urllib.error
    import urllib.request
    from email.utils import formataddr, parseaddr
    para = _dca_emails(para)
    cc = [c for c in _dca_emails(cc) if c not in para]
    if not para:
        return False, "Escribe al menos un correo válido en «Para»."
    if len(para) + len(cc) > 6:
        return False, "Máximo 6 destinatarios por envío."
    correo_envio, password = _credenciales_smtp("soporte")
    resend = (os.environ.get("RESEND_API_KEY") or "").strip()
    if not resend and (not correo_envio or not password):
        return False, "El correo de salida no está configurado. Conéctalo en Gerencia › Correo de soporte o define RESEND_API_KEY."
    entidad = _dca_nombre_remitente(d.get("entidad"))
    asunto = _dca_una_linea(asunto)[:200] or _dca_una_linea(d.get("asunto"))[:200] or "Documento PROCSIS"
    mensaje = (mensaje or "").replace("\r", "").strip()[:2000] or (
        "Por medio del presente nos permitimos remitir el documento «%s», emitido por %s. Lo encontrará adjunto en formato PDF." % (_dca_una_linea(d.get("asunto"))[:200], entidad))
    nombre_pdf = re.sub(r"[^A-Za-z0-9_-]", "_", d.get("codigo") or "documento") + ".pdf"
    emp = _dcx_empresa()
    responder = emp["email"] if _DCA_EMAIL_RE.match(str(emp.get("email") or "")) else ""
    html = _dca_correo_html(d, mensaje, nombre_pdf)
    texto = "Cordial saludo,\n\n%s\n\nDocumento: %s · %s\nAsunto: %s\nVerificación: %s\n\nAtentamente,\n%s\n%s\n%s · %s · %s" % (
        mensaje, d.get("codigo"), _dca_cat_txt(d.get("categoria")), d.get("asunto"), _dca_url(d), d.get("firmante"), d.get("entidad"), emp["tel"], emp["email"], emp["web"])
    ok, error = 0, ""
    try:
        pdf = _dca_pdf(d).read()
        if not pdf:
            raise RuntimeError("No se pudo generar el PDF del documento.")
        if resend:
            base = (os.environ.get("RESEND_FROM") or "").strip() or correo_envio or "onboarding@resend.dev"
            _n, addr = parseaddr(base)
            addr = addr or base
            payload = {"from": "%s <%s>" % (entidad, addr), "to": para, "subject": asunto, "html": html, "text": texto,
                       "attachments": [{"filename": nombre_pdf, "content": base64.b64encode(pdf).decode("ascii")}]}
            if cc:
                payload["cc"] = cc
            if responder:
                payload["reply_to"] = responder
            req = urllib.request.Request("https://api.resend.com/emails", data=json.dumps(payload).encode("utf-8"),
                                         headers={"Authorization": "Bearer %s" % resend, "Content-Type": "application/json", "User-Agent": "PROCSIS/1.0"}, method="POST")
            try:
                with urllib.request.urlopen(req, timeout=30) as r:
                    r.read()
            except urllib.error.HTTPError as e:
                detalle = e.read().decode("utf-8", errors="replace")[:260] if e.fp else str(e)
                raise RuntimeError("Resend rechazó el envío (HTTP %s): %s" % (e.code, detalle))
        else:
            _n, addr = parseaddr(correo_envio)
            msg = EmailMessage()
            msg["Subject"] = asunto
            msg["From"] = formataddr((entidad, addr or correo_envio))
            msg["To"] = ", ".join(para)
            if cc:
                msg["Cc"] = ", ".join(cc)
            if responder:
                msg["Reply-To"] = responder
            msg.set_content(texto)
            msg.add_alternative(html, subtype="html")
            msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=nombre_pdf)
            _smtp_enviar(msg, correo_envio, password, solo_smtp=True)
        ok = 1
    except Exception as e:
        error = str(e)[:300] or repr(e)[:300]
        print("docs_area correo:", error)
    try:
        _c48_x("INSERT INTO docs_area_envios (doc_id, para, cc, asunto, por, fecha, ok, error) VALUES (:d, :p, :c, :a, :u, :f, :o, :e)",
               d=d["id"], p=", ".join(para), c=", ".join(cc), a=asunto, u=usuario, f=str(ahora())[:19], o=ok, e=error)
    except Exception:
        db.session.rollback()
    if ok:
        registrar_auditoria("Documentos por área", "%s envió %s por correo a %s" % (usuario, d.get("codigo"), ", ".join(para)))
        return True, "Documento enviado en PDF a %s." % ", ".join(para + cc)
    return False, "No se pudo enviar el correo: %s" % (error[:200] or "revisa la conexión de correo e intenta de nuevo.")


'''

REEMPLAZOS = [
    ("selector de producto flotante",
     "if(anc&&anc.parentNode){anc.parentNode.insertBefore(box,anc.nextSibling);}else{box.className='fix';document.body.appendChild(box);}",
     "if(anc&&anc.parentNode){anc.parentNode.insertBefore(box,anc.nextSibling);}else{return;}  /* [procsis-correo-docs] solo dentro del menú lateral */"),
    ("botón de ayuda en páginas públicas /docs/",
     '"/mfa", "/validar", "/bienvenida", "/media")) or p.endswith(".pdf")',
     '"/mfa", "/validar", "/bienvenida", "/media", "/docs/")) or p.endswith(".pdf")'),
]

ANCLAS = [REQUIERE, "def _dca_pdf", "def _dca_url", "def _dca_emails", "def _dca_cat_txt", "def _dca_una_linea", "def _dcx_base", "def _dcx_empresa",
          "def _dcx_fecha_larga", "def _credenciales_smtp", "def _smtp_enviar", "def _c48_x", "def _esc", "_DCA_EMAIL_RE", "def registrar_auditoria", "_DCX_LOGO = "]


def main():
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
    if not ruta.is_file():
        sys.exit("No encuentro %s. Ejecuta el script en la carpeta de tu app.py o pásale la ruta." % ruta)
    with open(ruta, "r", encoding="utf-8", newline="") as fh:
        crudo = fh.read()
    crlf = "\r\n" in crudo
    texto = crudo.replace("\r\n", "\n")
    if MARCA in texto:
        print("Ya estaba aplicado (%s). No se cambió nada." % MARCA)
        return
    if REQUIERE not in texto:
        sys.exit("Primero aplica aplicar_documentos_area2.py (este arreglo depende de él). No toqué nada.")
    faltan = [a for a in ANCLAS if a not in texto]
    for desc, viejo, _n in REEMPLAZOS:
        if texto.count(viejo) != 1:
            faltan.append("%s (aparece %d veces, debe ser 1)" % (desc, texto.count(viejo)))
    if faltan:
        sys.exit("Este app.py no coincide con lo esperado; no toqué nada. Falta/cambió:\n  - " + "\n  - ".join(faltan))
    nuevo = texto
    for _d, viejo, reemplazo in REEMPLAZOS:
        nuevo = nuevo.replace(viejo, reemplazo, 1)
    corte = nuevo.rfind('\nif __name__ == "__main__":')
    if corte < 0:
        sys.exit('No encontré el bloque if __name__ == "__main__": al final de app.py.')
    nuevo = nuevo[:corte] + "\n" + BLOQUE.strip("\n") + "\n\n" + nuevo[corte:]
    try:
        compile(nuevo, str(ruta), "exec")
    except SyntaxError as e:
        sys.exit("El resultado no compila (línea %s: %s). No escribí nada." % (e.lineno, e.msg))
    copia = ruta.with_name(ruta.name + ".bak_correo_documentos")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_correo_documentos_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo: correo corporativo con PDF adjunto y selector solo en el menú. Archivo: %s" % ruta)
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()

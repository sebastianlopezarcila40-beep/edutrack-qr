# -*- coding: utf-8 -*-
"""
estado_modulos.py v2 · Página de estado y PRUEBA EN VIVO de los módulos de PROCSIS (/estado-modulos)
=====================================================================================================
Ya está instalada en tu app.py (bloque [procsis-estado-modulos-v1]). Solo reemplaza este archivo y haz git push:
no hay que volver a correr ningún aplicar_*.py.

  /estado-modulos            (público)  -> rutas esperadas y parches presentes: OK / FALTA.
  /estado-modulos            (con sesión de Gerencia) -> ADEMÁS:
        · PRUEBA EN VIVO: pide por dentro "/", "/login", "/publico/inicio" y "/cargadores" como lo haría un visitante
          sin sesión, y dice si el bloque comercial / el tema / los cargadores se agregaron a la página.
          Si algo falla muestra el error exacto que el sistema normalmente se traga.
        · Tablas del módulo comercial (existen / cuántas filas) y base de datos usada.
        · Errores de importación de cada módulo .py.
        · Qué trae acceso_seguro.py (su descripción y las variables que lee), para ajustar el PIN.
        · Variables de interruptor activas.
Interruptor:  ESTADO_MODULOS=off  la apaga.
"""
import ast
import contextlib
import importlib
import importlib.util
import io
import os
import re
import sys
import traceback

from flask import request, session, make_response

MARCA = "[procsis-estado-modulos-v1]"

RUTAS = (
    ("/login", "Login unificado"),
    ("/soporte-login", "Login de Soporte"),
    ("/support/editor", "Redactar Support (Soporte)"),
    ("/gerencia/support", "Autorizar artículos de Support (Gerencia)"),
    ("/cargadores", "20 animaciones de carga"),
    ("/comercial", "Gestor Comercial (interno)"),
    ("/publico/inicio", "Portada comercial pública"),
    ("/publico/comercial.json", "Datos comerciales públicos"),
    ("/documentos-oficiales", "Documentos oficiales"),
)

PARCHES = (
    ("[procsis-modulo-comercial-v1]", "Módulo comercial y portada pública"),
    ("[procsis-acceso-seguro-v2]", "Acceso seguro (PIN / MFA)"),
    ("[procsis-cargadores-v1]", "20 cargadores"),
    ("[procsis-codigos-72h]", "PIN cada 72 horas"),
    ("/*doc-fix-v1*/", "Arreglo del PDF de documentos"),
    ("Support · autorizar artículos", "Support visible en el menú de Gerencia"),
    ("[procsis-rutas-nuevas]", "Rutas nuevas permitidas en el filtro de paneles"),
    ("[procsis-estado-modulos-v1]", "Esta página de estado"),
)

MODULOS = ("modulo_comercial", "tema_2027", "acceso_seguro", "cargadores_20")
VARIABLES = ("COMERCIAL_MODO", "COMERCIAL_PORTADA", "COMERCIAL_BARRA", "COMERCIAL_AURA", "TEMA_2027", "CARGADOR_MODO",
             "CARGADOR_INYECTAR", "DISPOSITIVO_VALIDEZ_HORAS", "PIN_CONFIABLE_HORAS", "SESION_DISPOSITIVO_HORAS",
             "SESION_MAX_HORAS", "PIN_LOGIN_SOPORTE", "DISPOSITIVO_ACTIVO")
TABLAS = ("com_banner", "com_tarjeta", "com_plan", "com_equipo", "com_cupon", "com_factura", "com_cotizacion", "com_entrada", "com_aura")

# (ruta, qué buscar en el HTML, texto si lo encuentra, texto si NO lo encuentra)
PRUEBAS = (
    ("/", (("pc-root", "Bloque comercial (carrusel, planes, tienda, pagos, Edu Aura)"),
           ("t27-aplicado", "Tema 2027 (título, accesos rápidos)"),
           ("cg-aplicado", "Cargadores (cortina de carga)"))),
    ("/login", (("t27-aplicado", "Tema 2027"), ("cg-aplicado", "Cargadores"))),
    ("/publico/inicio", (("pc-root", "Bloque comercial completo"),)),
    ("/cargadores", (("cgg-grid", "Galería de 20 animaciones"),)),
)


def _esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _es_gerencia():
    try:
        return (session.get("rol") or "").strip() in ("Gerente", "Gerencia", "Superadmin", "Administrador")
    except Exception:
        return False


def _fila(ok, texto, extra=""):
    color = "#15803d" if ok else "#b91c1c"
    return ('<tr><td style="color:%s;font-weight:800;width:70px">%s</td><td>%s</td><td style="color:#64748b;font-size:12px">%s</td></tr>'
            % (color, "OK" if ok else "FALTA", _esc(texto), _esc(extra)))


def _h2(t):
    return '<h2 style="font-size:16px;margin:22px 0 6px;color:#002069">%s</h2>' % _esc(t)


def _pre(t, rojo=True):
    bg = "#fef2f2" if rojo else "#f1f5f9"
    return '<pre style="white-space:pre-wrap;word-break:break-word;background:%s;padding:8px 10px;border-radius:8px;font-size:12px;margin:6px 0">%s</pre>' % (bg, _esc(t))


def _prueba_get(app, ruta):
    """Pide la ruta como visitante sin sesión (sin cookies). Devuelve (estado, ubicación, html, tipo, log_capturado)."""
    buf = io.StringIO()
    estado, loc, html, tipo, err = None, "", "", "", ""
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            c = app.test_client()
            r = c.get(ruta, base_url=request.host_url, follow_redirects=False)
            estado, loc, tipo = r.status_code, r.headers.get("Location", ""), r.mimetype
            html = r.get_data(as_text=True)
    except Exception:
        err = traceback.format_exc()[-900:]
    return estado, loc, html, tipo, (buf.getvalue() + ("\n" + err if err else "")).strip()


def _log_relevante(texto):
    claves = ("COMERCIAL", "TEMA 2027", "CARGADORES", "Traceback", "Error", "ERROR", "error")
    lineas = [l for l in texto.splitlines() if any(k in l for k in claves)]
    return "\n".join(lineas[-12:])


def _acceso_info():
    """Descripción y variables de entorno que usa acceso_seguro.py (sin mostrar valores)."""
    try:
        spec = importlib.util.find_spec("acceso_seguro")
        if not spec or not spec.origin or not os.path.isfile(spec.origin):
            return None, [], ""
        fuente = open(spec.origin, "r", encoding="utf-8", errors="replace").read()
        try:
            doc = ast.get_docstring(ast.parse(fuente)) or ""
        except Exception:
            doc = ""
        nombres = sorted(set(re.findall(r"""(?:environ\.get|getenv)\(\s*["']([A-Z][A-Z0-9_]+)["']""", fuente))
                         | set(re.findall(r"""environ\[\s*["']([A-Z][A-Z0-9_]+)["']\s*\]""", fuente)))
        return doc[:3500], nombres, "%d líneas" % (fuente.count("\n") + 1)
    except Exception:
        return None, [], ""


def instalar(G):
    if (os.environ.get("ESTADO_MODULOS") or "on").strip().lower() == "off":
        print("ESTADO MODULOS: apagado.", flush=True)
        return
    try:
        app = G["app"]
        ruta_app = G.get("__file__") or ""
        try:
            fuente = open(ruta_app, "r", encoding="utf-8", errors="replace").read() if ruta_app else ""
        except Exception:
            fuente = ""

        def estado():
            reglas = {r.rule for r in app.url_map.iter_rules()}
            det = _es_gerencia()
            c = ['<h1 style="margin:0 0 4px;color:#002069">Estado de módulos · PROCSIS</h1>',
                 '<p style="margin:0 0 16px;color:#475569;font-size:14px">Si ves esta página, este app.py sí está desplegado.%s</p>'
                 % ("" if det else " <b>Entra con tu usuario de Gerencia</b> y recarga: ahí aparece la prueba en vivo con los errores exactos.")]
            c.append(_h2("Rutas") + '<table style="width:100%;border-collapse:collapse">')
            for r, n in RUTAS:
                c.append(_fila(r in reglas, n, r))
            c.append("</table>")
            c.append(_h2("Parches presentes en el app.py desplegado") + '<table style="width:100%;border-collapse:collapse">')
            if fuente:
                for m, n in PARCHES:
                    c.append(_fila(m in fuente, n))
            else:
                c.append("<tr><td>No pude leer app.py en el servidor.</td></tr>")
            c.append("</table>")

            if det:
                # ── 1. PRUEBA EN VIVO como visitante ──
                c.append(_h2("Prueba en vivo (como visitante, sin sesión)"))
                c.append('<p style="margin:0 0 6px;color:#475569;font-size:13px">El servidor pide estas páginas por dentro y revisa qué trae el HTML.</p>')
                c.append('<table style="width:100%;border-collapse:collapse">')
                for ruta, buscar in PRUEBAS:
                    estado_http, loc, html, tipo, log = _prueba_get(app, ruta)
                    if estado_http is None:
                        c.append(_fila(False, "GET %s" % ruta, "no se pudo pedir"))
                        c.append('<tr><td></td><td colspan="2">%s</td></tr>' % _pre(log or "(sin detalle)"))
                        continue
                    encab = "GET %s → HTTP %s" % (ruta, estado_http)
                    if loc:
                        encab += " → redirige a " + loc
                    c.append('<tr><td></td><td colspan="2" style="padding-top:8px;font-weight:700">%s</td></tr>' % _esc(encab))
                    for marca, nombre in buscar:
                        c.append(_fila(marca in html, nombre))
                    rel = _log_relevante(log)
                    if rel:
                        c.append('<tr><td></td><td colspan="2"><span style="font-size:12px;color:#b91c1c;font-weight:700">Mensajes del servidor al pedir esta página:</span>%s</td></tr>' % _pre(rel))
                c.append("</table>")
                c.append('<p style="font-size:13px;color:#475569;line-height:1.6;margin:8px 0 0">'
                         "<b>Cómo leerlo:</b> si «Bloque comercial» sale OK en <code>/</code>, el servidor SÍ lo agrega; si tú no lo ves, cierra tu sesión "
                         "(se oculta a quien ya inició sesión) o recarga con Ctrl+F5. Si sale FALTA, el error de arriba dice por qué. "
                         'Mira también <a href="/publico/inicio" target="_blank">/publico/inicio</a>.</p>')

                # ── 2. Tablas del módulo comercial ──
                db = G.get("db")
                c.append(_h2("Base de datos del módulo comercial"))
                if db is None:
                    c.append("<p>No encontré la conexión db.</p>")
                else:
                    try:
                        from sqlalchemy import text
                        dialecto = db.engine.dialect.name
                    except Exception:
                        text, dialecto = None, "?"
                    c.append('<p style="font-size:13px;margin:0 0 6px;color:#334155">Motor: %s</p><table style="width:100%%;border-collapse:collapse">' % _esc(dialecto))
                    for t in TABLAS:
                        try:
                            n = db.session.execute(text("SELECT COUNT(*) FROM " + t)).scalar()
                            c.append(_fila(True, t, "%s filas" % n))
                        except Exception as ex:
                            try:
                                db.session.rollback()
                            except Exception:
                                pass
                            c.append(_fila(False, t, "no existe o falla"))
                            c.append('<tr><td></td><td colspan="2">%s</td></tr>' % _pre(repr(ex)[:400]))
                    c.append("</table>")

                # ── 3. Importación de módulos ──
                c.append(_h2("Módulos .py (importación)") + '<table style="width:100%;border-collapse:collapse">')
                for m in MODULOS:
                    try:
                        mod = importlib.import_module(m)
                        c.append(_fila(hasattr(mod, "instalar"), m + ".py", getattr(mod, "__file__", "")))
                    except Exception:
                        c.append(_fila(False, m + ".py", "no se pudo importar"))
                        c.append('<tr><td></td><td colspan="2">%s</td></tr>' % _pre(traceback.format_exc()[-900:]))
                c.append("</table>")

                # ── 4. acceso_seguro (PIN) ──
                doc, nombres, tam = _acceso_info()
                c.append(_h2("acceso_seguro.py (el que pide el PIN)"))
                if doc is None:
                    c.append("<p>No pude leer acceso_seguro.py en el servidor.</p>")
                else:
                    c.append('<p style="font-size:13px;color:#334155;margin:0 0 4px">Archivo de %s. Descripción que trae:</p>%s' % (_esc(tam), _pre(doc or "(sin descripción)", rojo=False)))
                    if nombres:
                        c.append('<p style="font-size:13px;color:#334155;margin:6px 0 4px">Variables que lee (no muestro sus valores):</p>'
                                 '<table style="width:100%;border-collapse:collapse">')
                        for v in nombres:
                            puesta = bool(os.environ.get(v))
                            c.append('<tr><td style="color:%s;font-weight:800;width:70px">%s</td><td>%s</td><td style="color:#64748b;font-size:12px">%s</td></tr>'
                                     % ("#0172EC" if puesta else "#64748b", "DEFINIDA" if puesta else "—", _esc(v),
                                        "puesta en Vercel" if puesta else "sin definir (usa su valor por defecto)"))
                        c.append("</table>")
                    else:
                        c.append('<p style="font-size:13px;color:#334155">No lee variables de entorno.</p>')

                # ── 5. Servidor ──
                raiz = getattr(app, "root_path", "")
                try:
                    pys = sorted(f for f in os.listdir(raiz) if f.endswith(".py"))
                except Exception:
                    pys = []
                c.append(_h2("Servidor") + '<div style="font-size:13px;color:#334155;line-height:1.7">'
                         "Carpeta: %s<br>Python: %s<br>Archivos .py: %s<br>Variables activas: %s</div>"
                         % (_esc(raiz), _esc(sys.version.split()[0]), _esc(", ".join(pys) or "(ninguno)"),
                            _esc(", ".join("%s=%s" % (v, os.environ[v]) for v in VARIABLES if os.environ.get(v)) or "(ninguna)")))
            html = ('<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
                    '<meta name="robots" content="noindex"><title>Estado de módulos</title></head>'
                    '<body style="margin:0;background:#f8fbff;font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#0f172a">'
                    '<div style="max-width:860px;margin:0 auto;padding:22px 16px 48px">%s</div></body></html>' % "".join(c))
            r = make_response(html)
            r.headers["Cache-Control"] = "no-store"
            return r

        app.add_url_rule("/estado-modulos", "estado_modulos_pagina", estado, methods=["GET"])

        orig = G.get("ruta_publica")
        if callable(orig):
            def ruta_publica_est():
                return (request.path or "").rstrip("/") == "/estado-modulos" or bool(orig())
            G["ruta_publica"] = ruta_publica_est
        print("ESTADO MODULOS: /estado-modulos instalado.", flush=True)
    except Exception as ex:
        print("ESTADO MODULOS: NO se instaló (la app sigue igual):", repr(ex)[:160], flush=True)

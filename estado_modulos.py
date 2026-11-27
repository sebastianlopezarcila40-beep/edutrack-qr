# -*- coding: utf-8 -*-
"""
estado_modulos.py · Página de estado de los módulos de PROCSIS (/estado-modulos)
================================================================================
Se instala con UN bloque en app.py (lo agrega aplicar_todo.py):

    import estado_modulos; estado_modulos.instalar(globals())

Para qué sirve: ver desde el navegador si Vercel está corriendo la versión nueva y qué módulos cargaron,
sin entrar a los logs.

  /estado-modulos            (público)  -> lista rutas esperadas: OK / FALTA, y qué parches trae el app.py desplegado.
                                           No muestra rutas de archivos, variables ni errores.
  /estado-modulos            (con sesión de Gerencia) -> además: error exacto de cada módulo que no carga,
                                           archivos .py presentes, carpeta de la app y variables de interruptor.

Si /estado-modulos da 404, Vercel NO está corriendo este app.py (despliegue viejo, otro archivo de arranque o error de build).
Interruptor:  ESTADO_MODULOS=off  la apaga.
"""
import importlib
import os
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
    ("📝 Support · autorizar artículos", "Support visible en el menú de Gerencia"),
    ("[procsis-estado-modulos-v1]", "Esta página de estado"),
)

MODULOS = ("modulo_comercial", "tema_2027", "acceso_seguro", "cargadores_20")
VARIABLES = ("COMERCIAL_MODO", "COMERCIAL_PORTADA", "COMERCIAL_BARRA", "TEMA_2027", "CARGADOR_MODO", "CARGADOR_INYECTAR",
             "DISPOSITIVO_VALIDEZ_HORAS", "PIN_CONFIABLE_HORAS", "SESION_DISPOSITIVO_HORAS", "SESION_MAX_HORAS", "PIN_LOGIN_SOPORTE")


def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _es_gerencia():
    try:
        return (session.get("rol") or "").strip() in ("Gerente", "Gerencia", "Superadmin", "Administrador")
    except Exception:
        return False


def _fila(ok, texto, extra=""):
    color = "#15803d" if ok else "#b91c1c"
    return ('<tr><td style="color:%s;font-weight:800;width:70px">%s</td><td>%s</td><td style="color:#64748b;font-size:12px">%s</td></tr>'
            % (color, "OK" if ok else "FALTA", _esc(texto), _esc(extra)))


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
            cuerpo = ['<h1 style="margin:0 0 4px;color:#002069">Estado de módulos · PROCSIS</h1>',
                      '<p style="margin:0 0 16px;color:#475569;font-size:14px">Si ves esta página, este app.py sí está desplegado.%s</p>'
                      % ("" if det else " Entra con Gerencia para ver los errores en detalle.")]
            cuerpo.append('<h2 style="font-size:16px;margin:18px 0 6px">Rutas</h2><table style="width:100%;border-collapse:collapse">')
            for r, n in RUTAS:
                cuerpo.append(_fila(r in reglas, n, r))
            cuerpo.append("</table>")
            cuerpo.append('<h2 style="font-size:16px;margin:18px 0 6px">Parches presentes en el app.py desplegado</h2><table style="width:100%;border-collapse:collapse">')
            if fuente:
                for m, n in PARCHES:
                    cuerpo.append(_fila(m in fuente, n))
            else:
                cuerpo.append("<tr><td>No pude leer app.py en el servidor.</td></tr>")
            cuerpo.append("</table>")
            if det:
                cuerpo.append('<h2 style="font-size:16px;margin:18px 0 6px">Módulos .py (importación)</h2><table style="width:100%;border-collapse:collapse">')
                for m in MODULOS:
                    try:
                        mod = importlib.import_module(m)
                        cuerpo.append(_fila(hasattr(mod, "instalar"), m + ".py", getattr(mod, "__file__", "")))
                    except Exception:
                        cuerpo.append(_fila(False, m + ".py", "no se pudo importar"))
                        cuerpo.append('<tr><td></td><td colspan="2"><pre style="white-space:pre-wrap;background:#fef2f2;padding:8px;border-radius:8px;font-size:12px">%s</pre></td></tr>'
                                      % _esc(traceback.format_exc()[-900:]))
                cuerpo.append("</table>")
                raiz = getattr(app, "root_path", "")
                try:
                    pys = sorted(f for f in os.listdir(raiz) if f.endswith(".py"))
                except Exception:
                    pys = []
                cuerpo.append('<h2 style="font-size:16px;margin:18px 0 6px">Servidor</h2><div style="font-size:13px;color:#334155;line-height:1.7">'
                              "Carpeta: %s<br>Python: %s<br>Archivos .py en la carpeta: %s<br>Variables activas: %s</div>"
                              % (_esc(raiz), _esc(sys.version.split()[0]), _esc(", ".join(pys) or "(ninguno)"),
                                 _esc(", ".join("%s=%s" % (v, os.environ[v]) for v in VARIABLES if os.environ.get(v)) or "(ninguna)")))
            html = ('<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
                    '<meta name="robots" content="noindex"><title>Estado de módulos</title></head>'
                    '<body style="margin:0;background:#f8fbff;font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#0f172a">'
                    '<div style="max-width:860px;margin:0 auto;padding:22px 16px 48px">%s</div></body></html>' % "".join(cuerpo))
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

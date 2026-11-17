#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_menu_producto.py  ·  PROCSIS: el menú lateral cambia según el producto elegido en el selector.

REQUIERE que ya hayas aplicado aplicar_multiproducto.py (usa su selector y su función _pf_modo).

Uso (en la carpeta de tu app.py):
    python aplicar_menu_producto.py            # parchea ./app.py
    python aplicar_menu_producto.py ruta/app.py

Resultado:
  · Todo PROCSIS  -> el menú completo (EduTrack + TiendaTrack + corporativo), como hoy.
  · EduTrack      -> se oculta TODO lo de TiendaTrack.
  · TiendaTrack   -> se ocultan los módulos exclusivos de EduTrack (colegios, planes de colegios, donaciones, activaciones,
                     rectores, carné estudiantil, retractos...) y el grupo TiendaTrack sube al primer lugar y queda abierto.
  · Los módulos corporativos (Talento Humano, Legal, Comunicaciones y CRTP, Sistema, Tecnología y TI, Respaldo, Tickets...)
    se ven siempre.
La lista de qué es exclusivo de EduTrack está en _PFM_SOLO_EDUTRACK: si quieres mover un módulo, agrégalo o quítalo ahí.
Es idempotente, crea app.py.bak_menu_producto y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-menu-producto]"
REQUIERE = "[procsis-multiproducto]"

BLOQUE = r'''
# [procsis-menu-producto]
# ═════════════ MENÚ LATERAL POR PRODUCTO: Todo PROCSIS · EduTrack · TiendaTrack ═════════════
# Cada módulo es "corporativo" (se ve siempre) salvo que esté aquí: exclusivo de EduTrack (colegios, estudiantes) o de TiendaTrack.
_PFM_SOLO_EDUTRACK = frozenset((
    # Gerencia: colegios, planes de colegios, donaciones a colegios
    "/gerencia/donaciones", "/gerencia/donaciones/kit", "/gerencia/envios", "/gerencia/convocatoria48", "/gerencia/reconocimientos",
    "/gerencia/activaciones", "/gerencia/rectores", "/gerencia/colegios-config", "/gerencia/planes", "/gerencia/planes/nuevo",
    "/gerencia/planes-vendidos", "/gerencia/beneficios", "/planes/buscar", "/gerencia/descuentos", "/gerencia/alianzas-clientes",
    "/gerencia/autorizar-soporte-rectores", "/gerencia/solicitudes-plan", "/gerencia/contratos-saas", "/gerencia/contrato-plantilla",
    "/gerencia/retractos", "/gerencia/retractos/config", "/gerencia/retractos/reembolsos", "/gerencia/finanzas/promociones",
    "/gerencia/pie-login", "/gerencia/cancelaciones", "/gerencia/carnes", "/gerencia/cartera",
    # Ventas
    "/ventas/activaciones", "/ventas/verificacion", "/ventas/beneficios", "/ventas/planes-vendidos",
    # Soporte
    "/soporte/bloqueo-carne", "/soporte/instituciones", "/soporte/validacion-identidad", "/soporte/convocatoria48-actualizacion",
    "/soporte/periodos-colegio", "/soporte/prorroga", "/soporte/auditoria-colegios", "/soporte/planes-vendidos", "/soporte/impersonar",
    "/tenants", "/usuarios",
    # Soluciones
    "/soluciones-procsis/clientes",
))
_PFM_SOLO_TIENDATRACK_PREFIJOS = ("/gerencia/tiendatrack",)


def _pfm_oculto(href, modo):
    """True si ese enlace no debe aparecer con el producto elegido."""
    h = (href or "").split("?")[0].rstrip("/") or "/"
    if modo == "edutrack":
        return any(h == p or h.startswith(p + "/") for p in _PFM_SOLO_TIENDATRACK_PREFIJOS)
    if modo == "tiendatrack":
        return h in _PFM_SOLO_EDUTRACK
    return False


def _pf_menu_items(items, rol=""):
    """Filtra la lista (href, etiqueta) del menú lateral según el producto del selector. 'Todo PROCSIS' = sin cambios."""
    try:
        modo = _pf_modo()
        if modo not in ("edutrack", "tiendatrack"):
            return items
        filtrado = [(h, l) for h, l in items if h == "/logout" or not _pfm_oculto(h, modo)]
        return filtrado if len(filtrado) > 1 else items
    except Exception as _e:
        print("pf menu:", _e)
        return items


if "_nav_agrupar" in globals():
    _pfm_nav_agrupar_orig = _nav_agrupar

    def _nav_agrupar(filas, grupos):
        """En TiendaTrack, el grupo TiendaTrack va primero (después de Inicio) y se muestra abierto."""
        try:
            modo = _pf_modo()
        except Exception:
            modo = None
        if modo != "tiendatrack":
            return _pfm_nav_agrupar_orig(filas, grupos)
        try:
            inicio = [g for g in grupos if g[0] == "Inicio"]
            tienda = [g for g in grupos if g[0] == "TiendaTrack"]
            resto = [g for g in grupos if g[0] not in ("Inicio", "TiendaTrack")]
            html = _pfm_nav_agrupar_orig(filas, inicio + tienda + resto)
            return html.replace('<details class="nav-group" data-g="TiendaTrack">', '<details class="nav-group" data-g="TiendaTrack" open>', 1)
        except Exception as _e:
            print("pf menu grupos:", _e)
            return _pfm_nav_agrupar_orig(filas, grupos)


'''

ANCLA_ITEMS = '    for href, lab in items:\n        if href == "/logout":\n            active = ""\n        elif href.startswith("/dev-console"):\n'
NUEVO_ITEMS = '    items = _pf_menu_items(items, rol)  # [procsis-menu-producto]\n' + ANCLA_ITEMS

REEMPLAZOS = [("menú lateral: filtrar por producto", ANCLA_ITEMS, NUEVO_ITEMS)]
ANCLAS = ["def _staff_nav_items", "def _nav_agrupar(filas, grupos):", "def _pf_modo", '_MODULOS_GERENCIA = [']


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
        sys.exit("Primero aplica aplicar_multiproducto.py (este parche usa su selector). No toqué nada.")
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

    copia = ruta.with_name(ruta.name + ".bak_menu_producto")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_menu_producto_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. El menú lateral ahora cambia según el producto. Archivo: %s" % ruta)
    print("Copia de seguridad: %s" % copia)


if __name__ == "__main__":
    main()

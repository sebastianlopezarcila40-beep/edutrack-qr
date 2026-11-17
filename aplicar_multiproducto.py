#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_multiproducto.py  ·  PROCSIS: arquitectura Multi-Producto (EduTrack + TiendaTrack)

Uso (en la carpeta donde está tu app.py):
    python aplicar_multiproducto.py            # parchea ./app.py
    python aplicar_multiproducto.py ruta/app.py

Qué hace (todo sobre app.py, con copia de seguridad app.py.bak_multiproducto):
  1. Agrega la columna producto_id (edutrack | tiendatrack | corporativo) a las tablas de ingresos, cartera,
     tickets, gastos, conciliación, operaciones contables, leads y metas. Lo ya existente queda como 'edutrack'.
  2. Filtro automático: cuando el personal elige EduTrack o TiendaTrack en el selector del menú lateral, TODAS las
     consultas ORM de esas tablas se filtran solas. "Todo PROCSIS" = sin filtro (suma todo).
  3. Etiquetado automático al crear registros (según el producto elegido o según el origen del dato).
  4. RR.HH., Talento Humano y Nómina NO se tocan: siguen centralizados en PROCSIS.
  5. Selector: cabecera X-Producto en cada fetch/XHR del mismo sitio + preloader rápido al cambiar de producto.
  6. API para TiendaTrack:  POST /api/v1/tiendatrack/evento  (X-API-Key)  con tipo = registro | pago | ticket.
  7. Panel de Gerencia/Contabilidad/Cobranza/Finanzas:  /gerencia/productos  (consolidado por producto).
  8. KPIs del HQ de Gerencia: en "TiendaTrack" muestran TiendaTrack; en "Todo PROCSIS" suman EduTrack + TiendaTrack.

El script es idempotente (si ya lo aplicaste, no hace nada) y no escribe nada si el resultado no compila.
"""
import shutil
import sys
import time
from pathlib import Path

MARCA = "[procsis-multiproducto]"

# ──────────────────────────── Bloque que se agrega al final de app.py ────────────────────────────
BLOQUE = r'''
# [procsis-multiproducto]
# ═════════════ MULTI-PRODUCTO · producto_id en la BD + filtro automático por el selector del menú lateral ═════════════
# producto_id = edutrack | tiendatrack | corporativo.   "Todo PROCSIS" = sin filtro (suma todo).
# Etiquetado: FacturaCobro, PagoCuotaColegio, ConciliacionBancaria, TicketPQR, LeadCRM, MetaVendedor, MetaEmpresa,
# GastoOperativo y ContOperacion. RR.HH. / Talento Humano / Nómina NO se etiquetan ni se filtran (son de PROCSIS matriz).
from flask import has_request_context as _pf_hrc, g as _pf_g
from sqlalchemy import event as _pf_event
from sqlalchemy.orm import Session as _pf_Session, with_loader_criteria as _pf_wlc

_PF_PRODUCTOS = ("edutrack", "tiendatrack", "corporativo")
_PF_FILTRABLES = ("edutrack", "tiendatrack")
_PF_ESTADO = {"migrado": False, "t": 0.0}
_PF_ROLES_CONSOLIDADO = tuple(_DCX_GERENCIA) + ("Contabilidad", "Cobranza", "Finanzas")
# modelo -> (producto por defecto si no hay uno elegido, ¿es de un colegio cuando trae institucion_id?)
_PF_MODELOS = {
    "FacturaCobro": ("edutrack", True),
    "PagoCuotaColegio": ("edutrack", True),
    "ConciliacionBancaria": ("edutrack", False),
    "TicketPQR": ("edutrack", True),
    "LeadCRM": ("edutrack", False),
    "MetaVendedor": ("edutrack", False),
    "MetaEmpresa": ("corporativo", False),
    "GastoOperativo": ("corporativo", False),
    "ContOperacion": ("corporativo", True),
}
_PF_CLASES = []


def _pf_producto():
    """'edutrack' | 'tiendatrack' si el personal tiene un producto elegido; None = Todo PROCSIS o sin filtro."""
    try:
        if not _pf_hrc() or getattr(_pf_g, "pf_sin_filtro", False):
            return None
        if not session.get("usuario") or rol_actual() not in _BIENVENIDA_ROLES:
            return None
        h = (request.headers.get("X-Producto") or "").strip().lower()
        p = h if h in ("todo", "edutrack", "tiendatrack") else (session.get("producto") or "todo")
        return p if p in _PF_FILTRABLES else None
    except Exception:
        return None


def _pf_modo():
    """'todo' | 'edutrack' | 'tiendatrack' para el personal; None si no es personal interno."""
    try:
        if not _pf_hrc() or not session.get("usuario") or rol_actual() not in _BIENVENIDA_ROLES:
            return None
        return _pf_producto() or "todo"
    except Exception:
        return None


def _pf_crit(cls, prod):
    # `prod` debe ser variable de cierre (no argumento por defecto): así SQLAlchemy la trata como parámetro y no la cachea fija.
    return _pf_wlc(cls, lambda c: c.producto_id == prod, include_aliases=True)


def _pf_filtrar(estado):
    try:
        if (not estado.is_select) or estado.is_column_load or estado.is_relationship_load or estado.execution_options.get("pf_todo"):
            return
        prod = _pf_producto()
        if not prod:
            return
        estado.statement = estado.statement.options(*[_pf_crit(c, prod) for c in _PF_CLASES])
    except Exception as _e:
        print("pf filtro:", _e)


def _pf_asignar(mapper, conexion, obj):
    """Etiqueta cada registro nuevo con su producto (si el código no lo puso ya)."""
    try:
        if getattr(obj, "producto_id", None) in _PF_PRODUCTOS:
            return
        base, por_colegio = _PF_MODELOS.get(obj.__class__.__name__, ("edutrack", False))
        if por_colegio and getattr(obj, "institucion_id", None):
            obj.producto_id = "edutrack"
        else:
            obj.producto_id = _pf_producto() or base
    except Exception:
        obj.producto_id = "edutrack"


def _pf_registrar():
    for nombre in _PF_MODELOS:
        cls = globals().get(nombre)
        if cls is None:
            print("multiproducto: no existe el modelo", nombre)
            continue
        if "producto_id" not in cls.__table__.c:
            setattr(cls, "producto_id", db.Column(db.String(20), default="edutrack", server_default="edutrack"))
        _pf_event.listen(cls, "before_insert", _pf_asignar)
        _PF_CLASES.append(cls)
    _pf_event.listen(_pf_Session, "do_orm_execute", _pf_filtrar)


try:
    _pf_registrar()
except Exception as _pf_err:
    print("multiproducto: no se pudo activar el filtro:", _pf_err)


def _pf_migrar():
    """Agrega producto_id a las tablas que ya existían (idempotente, barato). Lo existente queda como 'edutrack'."""
    if _PF_ESTADO["migrado"] or not _PF_CLASES:
        return
    import time as _pf_time
    if _pf_time.time() - _PF_ESTADO["t"] < 60:  # si faltan tablas (BD nueva) no revisar en cada petición
        return
    _PF_ESTADO["t"] = _pf_time.time()
    try:
        tablas = [c.__tablename__ for c in _PF_CLASES]
        sqlite = db.engine.dialect.name == "sqlite"
        presentes, con_col = set(), set()
        with db.engine.connect() as conn:
            if sqlite:
                for t in tablas:
                    filas = conn.execute(text("PRAGMA table_info(%s)" % t)).fetchall()
                    if filas:
                        presentes.add(t)
                    if any(f[1] == "producto_id" for f in filas):
                        con_col.add(t)
            else:
                lista = ",".join("'%s'" % t for t in tablas)
                filas = conn.execute(text(
                    "SELECT table_name, column_name FROM information_schema.columns "
                    "WHERE table_schema = current_schema() AND table_name IN (%s)" % lista)).fetchall()
                for t, c in filas:
                    presentes.add(t)
                    if c == "producto_id":
                        con_col.add(t)
        for t in sorted(presentes - con_col):
            ddl = ("ALTER TABLE %s ADD COLUMN producto_id VARCHAR(20) DEFAULT 'edutrack'" if sqlite
                   else "ALTER TABLE %s ADD COLUMN IF NOT EXISTS producto_id VARCHAR(20) DEFAULT 'edutrack'") % t
            try:
                with db.engine.begin() as conn:
                    conn.execute(text(ddl))
                    conn.execute(text("CREATE INDEX IF NOT EXISTS ix_%s_producto_id ON %s (producto_id)" % (t, t)))
                con_col.add(t)
            except Exception as _e:
                print("multiproducto migrar", t, ":", repr(_e)[:200])
        # Solo se da por hecho cuando todas las tablas existen y ya tienen la columna (en BD nueva, create_all la crea sola).
        if presentes == set(tablas) and con_col >= presentes:
            _PF_ESTADO["migrado"] = True
    except Exception as _e:
        print("multiproducto migrar:", repr(_e)[:200])


def _pf_antes_de_todo():
    if not _PF_ESTADO["migrado"]:
        _pf_migrar()
    try:
        if not hasattr(_pf_g, "pf_sin_filtro"):
            _pf_g.pf_sin_filtro = False
    except Exception:
        pass


# Corre antes que cualquier otro before_request (la migración debe existir antes de la primera consulta).
app.before_request_funcs.setdefault(None, []).insert(0, _pf_antes_de_todo)


def _pf_sin_filtro(fn):
    """Envuelve generadores de consecutivos: deben mirar TODOS los registros (no solo los del producto elegido)."""
    def _f(*a, **k):
        previo = getattr(_pf_g, "pf_sin_filtro", False) if _pf_hrc() else False
        try:
            if _pf_hrc():
                _pf_g.pf_sin_filtro = True
        except Exception:
            pass
        try:
            return fn(*a, **k)
        finally:
            try:
                if _pf_hrc():
                    _pf_g.pf_sin_filtro = previo
            except Exception:
                pass
    _f.__name__ = getattr(fn, "__name__", "_f")
    _f.__doc__ = getattr(fn, "__doc__", None)
    return _f


for _pf_nombre in ("_gen_radicado_pqr", "_siguiente_consecutivo_recibo", "_cont_codigo_op"):
    if _pf_nombre in globals():
        globals()[_pf_nombre] = _pf_sin_filtro(globals()[_pf_nombre])


# ───────────── Datos de TiendaTrack (viven en tt_cuentas / tt_pagos) ─────────────
def _pf_tt_kpis():
    """Clientes y dinero de TiendaTrack. Los ingresos son los pagos 'aprobado' de tt_pagos."""
    _tt_init()
    mes = ahora().strftime("%Y-%m")
    cuentas = [_tt_dict(_TT_COLS_CUENTA, f) for f in _c48_q("SELECT %s FROM tt_cuentas" % ", ".join(_TT_COLS_CUENTA))]
    estados = {c["id"]: _tt_estado(c) for c in cuentas}
    activos = sum(1 for e in estados.values() if e == "activo")
    demos = sum(1 for e in estados.values() if e == "demo")
    vencidos = sum(1 for e in estados.values() if e in ("vencido", "suspendido"))
    ultimo = {}
    cobrado = cobrado_mes = 0.0
    for cid, per, cents, creado in _c48_q("SELECT cuenta_id, periodo, monto_centavos, creado_en FROM tt_pagos WHERE estado='aprobado' ORDER BY id"):
        valor = float(cents or 0) / 100.0
        cobrado += valor
        if (creado or "").startswith(mes):
            cobrado_mes += valor
        ultimo[cid] = (per, valor)
    mrr = 0.0
    for cid, (per, valor) in ultimo.items():
        if estados.get(cid) == "activo":
            mrr += valor / 12.0 if per == "anual" else valor
    altas_mes = sum(1 for c in cuentas if (c.get("creado_en") or "").startswith(mes))
    return {"activos": activos, "demos": demos, "vencidos": vencidos, "mrr": mrr, "cobrado": cobrado,
            "cobrado_mes": cobrado_mes, "altas_mes": altas_mes, "cuentas": len(cuentas)}


def _pf_recalcular(m):
    """Recalcula los derivados del HQ (neto, margen, churn, ARPU, LTV, cashflow) después de mezclar productos."""
    mrr = float(m.get("mrr") or 0)
    m["arr"] = mrr * 12
    m["neto"] = mrr - float(m.get("gastos") or 0)
    m["margen"] = round(100.0 * m["neto"] / mrr, 1) if mrr > 0 else 0.0
    total = max(int(m.get("activas") or 0) + int(m.get("suspendidas") or 0), 1)
    churn_f = float(m.get("suspendidas") or 0) / total
    m["churn"] = round(100.0 * churn_f, 1)
    activas = int(m.get("activas") or 0)
    m["arpu"] = (mrr / activas) if activas else 0.0
    m["ltv"] = (m["arpu"] / churn_f) if churn_f > 0.001 else m["arpu"] * 24
    m["cashflow"] = float(m.get("cobrado") or 0) - float(m.get("gastos_anio") or 0)


if "_gerencia_metricas" in globals():
    _pf_gerencia_metricas_orig = _gerencia_metricas

    def _gerencia_metricas():
        """KPIs de Gerencia: EduTrack (filtrado por el ORM), TiendaTrack (tt_*) o la suma de ambos en 'Todo PROCSIS'."""
        m = _pf_gerencia_metricas_orig()
        try:
            modo = _pf_modo()
            if modo not in ("todo", "tiendatrack") or not isinstance(m, dict):
                return m
            k = _pf_tt_kpis()
            if modo == "tiendatrack":
                m.update({"activas": k["activos"], "suspendidas": k["vencidos"], "mrr": k["mrr"], "cartera": 0.0,
                          "cobrado": k["cobrado"], "altas_mes": k["altas_mes"], "cac": 0.0})
            else:
                m["activas"] = int(m.get("activas") or 0) + k["activos"]
                m["suspendidas"] = int(m.get("suspendidas") or 0) + k["vencidos"]
                m["mrr"] = float(m.get("mrr") or 0) + k["mrr"]
                m["cobrado"] = float(m.get("cobrado") or 0) + k["cobrado"]
                m["altas_mes"] = int(m.get("altas_mes") or 0) + k["altas_mes"]
            _pf_recalcular(m)
            m["producto"] = modo
        except Exception as _e:
            print("pf metricas:", _e)
        return m


# ───────────── Panel consolidado: /gerencia/productos ─────────────
def _pf_resumen():
    """Cifras por producto leídas directo de la BD (sin el filtro automático) para comparar y sumar."""
    mes = ahora().strftime("%Y-%m")
    d = {p: {"cobrado": 0.0, "cobrado_mes": 0.0, "cartera": 0.0, "gastos": 0.0, "gastos_mes": 0.0,
             "tickets": 0, "cierres": 0, "meta": 0, "clientes": 0} for p in _PF_PRODUCTOS}

    def prod(v):
        return v if v in _PF_PRODUCTOS else "edutrack"

    for p, est, tot in _c48_q("SELECT COALESCE(producto_id,'edutrack'), estado, COALESCE(SUM(valor),0) FROM facturas_cobro GROUP BY 1, 2"):
        if est in ("PAGADA", "PAGADO"):
            d[prod(p)]["cobrado"] += float(tot or 0)
        elif est in ("PENDIENTE", "VENCIDA", "VENCIDO"):
            d[prod(p)]["cartera"] += float(tot or 0)
    for p, tot in _c48_q("SELECT COALESCE(producto_id,'edutrack'), COALESCE(SUM(valor),0) FROM facturas_cobro "
                         "WHERE estado IN ('PAGADA','PAGADO') AND pagada_en LIKE :m GROUP BY 1", m=mes + "%"):
        d[prod(p)]["cobrado_mes"] += float(tot or 0)
    for p, tot in _c48_q("SELECT COALESCE(producto_id,'edutrack'), COALESCE(SUM(valor),0) FROM gastos_operativos GROUP BY 1"):
        d[prod(p)]["gastos"] += float(tot or 0)
    for p, tot in _c48_q("SELECT COALESCE(producto_id,'edutrack'), COALESCE(SUM(valor),0) FROM gastos_operativos WHERE mes=:m GROUP BY 1", m=mes):
        d[prod(p)]["gastos_mes"] += float(tot or 0)
    for p, n in _c48_q("SELECT COALESCE(producto_id,'edutrack'), COUNT(*) FROM tickets_pqr "
                       "WHERE estado IN ('RADICADO','ABIERTA','URGENTE','PENDIENTE') GROUP BY 1"):
        d[prod(p)]["tickets"] += int(n or 0)
    for p, cer, meta in _c48_q("SELECT COALESCE(producto_id,'edutrack'), COALESCE(SUM(cerrados),0), COALESCE(SUM(meta_colegios),0) "
                               "FROM metas_vendedor WHERE mes=:m GROUP BY 1", m=mes):
        d[prod(p)]["cierres"] += int(cer or 0)
        d[prod(p)]["meta"] += int(meta or 0)
    try:
        d["edutrack"]["clientes"] = Institucion.query.filter(Institucion.estado == "ACTIVA").count()
    except Exception:
        db.session.rollback()
    k = _pf_tt_kpis()
    d["tiendatrack"]["cobrado"] += k["cobrado"]
    d["tiendatrack"]["cobrado_mes"] += k["cobrado_mes"]
    d["tiendatrack"]["clientes"] = k["activos"]
    d["tiendatrack"]["demos"] = k["demos"]
    return d


@app.route("/gerencia/productos")
def pf_consolidado():
    if not requiere_login():
        return redirect("/login")
    if rol_actual() not in _PF_ROLES_CONSOLIDADO:
        return acceso_denegado("Este panel es para Gerencia, Contabilidad, Cobranza y Finanzas.")
    try:
        _pf_g._staff_wrap = True
    except Exception:
        pass
    d = _pf_resumen()
    total = {k: sum(d[p][k] for p in _PF_PRODUCTOS) for k in d["edutrack"]}
    activo = _pf_modo() or "todo"
    nombres = {"edutrack": "EduTrack", "tiendatrack": "TiendaTrack", "corporativo": "Corporativo"}
    filas = (("Ingresos cobrados (total)", "cobrado", True), ("Ingresos cobrados (este mes)", "cobrado_mes", True),
             ("Cartera por cobrar", "cartera", True), ("Gastos (total registrado)", "gastos", True),
             ("Gastos (este mes)", "gastos_mes", True), ("Clientes activos", "clientes", False),
             ("Tickets abiertos", "tickets", False), ("Cierres de venta (mes)", "cierres", False), ("Meta de cierres (mes)", "meta", False))
    cab = "".join("<th%s>%s</th>" % (" style='background:#e8edff'" if activo == p else "", nombres[p]) for p in _PF_PRODUCTOS)
    cuerpo = ""
    for etiqueta, clave, dinero in filas:
        f = (lambda v: _tt_cop(v)) if dinero else (lambda v: "{:,.0f}".format(v).replace(",", "."))
        celdas = "".join("<td%s>%s</td>" % (" style='background:#f4f7ff'" if activo == p else "", f(d[p][clave])) for p in _PF_PRODUCTOS)
        cuerpo += "<tr><td><b>%s</b></td>%s<td><b>%s</b></td></tr>" % (_esc(etiqueta), celdas, f(total[clave]))
    neto = total["cobrado"] - total["gastos"]
    html = ("<h1>📊 Consolidado por producto</h1><p class='sub'>EduTrack + TiendaTrack + gastos generales. Cada producto conserva su información; "
            "«Todo PROCSIS» es la suma. Elige el producto en el menú lateral para filtrar los demás módulos.</p>"
            "<div class='cb-card cb-scroll'><table><tr><th>Concepto</th>%s<th>Total PROCSIS</th></tr>%s</table>"
            "<p class='cb-note'>Utilidad acumulada de toda la compañía (cobrado − gastos): <b>%s</b>. Los ingresos de TiendaTrack son los pagos aprobados de "
            "sus planes. RR.HH., nómina y planillas PILA no se dividen: son de PROCSIS como empresa matriz.</p></div>") % (cab, cuerpo, _tt_cop(neto))
    return _cb_pagina("Consolidado por producto", html)


# ───────────── API para TiendaTrack: registro de tienda, pago de plan y ticket de soporte ─────────────
def _pf_api_ok():
    llave = (os.environ.get("TIENDATRACK_API_KEY") or "").strip()
    dada = (request.headers.get("X-API-Key") or "").strip()
    return bool(llave) and _tt_hm.compare_digest(llave, dada)


def _pf_txt(v, n):
    return " ".join(str(v or "").split())[:n]


@app.route("/api/v1/tiendatrack/evento", methods=["POST"])
def pf_tt_evento():
    """Recibe eventos de la app TiendaTrack y los guarda en la BD central con producto_id = tiendatrack.
    Cabecera X-API-Key. Cuerpo JSON: {"tipo": "registro" | "pago" | "ticket", ...}  (ver los campos en cada rama)."""
    ok_rl, _e = _rate_limit("tt-api-evt", 120, 60)
    if not ok_rl:
        return {"ok": False, "error": "demasiadas solicitudes"}, 429
    if not _pf_api_ok():
        return {"ok": False, "error": "no autorizado"}, 401
    datos = request.get_json(silent=True) or {}
    tipo = _pf_txt(datos.get("tipo"), 20).lower()
    _tt_init()
    email = _pf_txt(datos.get("email"), 160).lower()
    ahora_s = _tt_ahora_s()
    try:
        if tipo == "registro":  # nueva tienda: email, nombre, negocio, segmento, telefono
            if "@" not in email:
                return {"ok": False, "error": "email requerido"}, 400
            seg = _pf_txt(datos.get("segmento"), 20).lower()
            seg = seg if seg in [s for s, _n in _TT_SEGMENTOS] else "tienda"
            c = _tt_cuenta_por("email", email)
            if c:
                return {"ok": True, "cuenta_id": c["id"], "estado": _tt_estado(c), "nuevo": False}, 200
            dias = _tt_int(_tt_cfg("tt_demo_dias", "15"), 15, 1, 90)
            _c48_x("INSERT INTO tt_cuentas (email, nombre, negocio, segmento, telefono, estado, demo_inicio, demo_fin, creado_en) "
                   "VALUES (:e, :n, :ng, :s, :t, 'demo', :di, :df, :c)",
                   e=email, n=_pf_txt(datos.get("nombre"), 160), ng=_pf_txt(datos.get("negocio"), 160), s=seg,
                   t=_pf_txt(datos.get("telefono"), 30), di=ahora_s, df=_tt_sumar_dias(ahora_s, dias), c=ahora_s)
            c = _tt_cuenta_por("email", email)
            registrar_auditoria("TiendaTrack API · registro", "email=%s" % email)
            return {"ok": True, "cuenta_id": c["id"], "estado": _tt_estado(c), "nuevo": True}, 201

        if tipo == "pago":  # pago de plan: email, plan (slug), periodo (mensual|anual), referencia, monto (COP), estado, id_transaccion
            c = _tt_cuenta_por("email", email)
            if not c:
                return {"ok": False, "error": "cuenta no encontrada (envía primero el evento registro)"}, 404
            plan = next((p for p in _tt_planes() if p["slug"] == _pf_txt(datos.get("plan"), 40).lower()), None)
            if not plan:
                return {"ok": False, "error": "plan no encontrado"}, 404
            periodo = "anual" if _pf_txt(datos.get("periodo"), 10).lower() == "anual" else "mensual"
            precio = int(plan["precio_anual"] if periodo == "anual" else plan["precio_mensual"])
            monto = _tt_int(datos.get("monto"), precio, 0, 10 ** 10) or precio
            ref = _pf_txt(datos.get("referencia"), 60) or ("TT-API-" + secrets.token_hex(6).upper())
            if not ref.startswith("TT-"):
                ref = "TT-" + ref
            if not _c48_q("SELECT id FROM tt_pagos WHERE referencia=:r", r=ref):
                _c48_x("INSERT INTO tt_pagos (cuenta_id, plan_id, periodo, referencia, monto_centavos, moneda, estado, pasarela, creado_en, actualizado_en) "
                       "VALUES (:c, :p, :pe, :r, :m, 'COP', 'pendiente', :pa, :t, :t)",
                       c=c["id"], p=plan["id"], pe=periodo, r=ref, m=monto * 100, pa=_pf_txt(datos.get("pasarela"), 16) or "api", t=ahora_s)
            estado = _pf_txt(datos.get("estado"), 20).upper() or "APPROVED"
            ok, msg, code = _tt_aplicar_tx(ref, estado, _pf_txt(datos.get("id_transaccion"), 80), monto * 100)
            return {"ok": ok, "referencia": ref, "mensaje": msg}, code

        if tipo == "ticket":  # soporte: email, negocio, asunto, descripcion, prioridad (CRITICA|ALTA|MEDIA|BAJA)
            asunto = _pf_txt(datos.get("asunto"), 120)
            if not asunto:
                return {"ok": False, "error": "asunto requerido"}, 400
            c = _tt_cuenta_por("email", email) if email else None
            prioridad = _pf_txt(datos.get("prioridad"), 10).upper()
            prioridad = prioridad if prioridad in ("CRITICA", "ALTA", "MEDIA", "BAJA") else "MEDIA"
            rad, tick = _gen_radicado_pqr("33")
            t = TicketPQR(radicado=rad, ticket=tick, producto_id="tiendatrack", fecha=ahora().strftime("%Y-%m-%d"), hora=ahora().strftime("%H:%M"),
                          email=email, razon_social=_pf_txt(datos.get("negocio") or (c or {}).get("negocio"), 220),
                          tipo_pqr="Soporte TiendaTrack", subtipo=asunto, objeto=asunto, hechos=str(datos.get("descripcion") or "")[:4000],
                          estado="RADICADO", canal="TIENDATRACK", medio_ingreso="API TiendaTrack", prioridad=prioridad,
                          creado_por="TiendaTrack API")
            db.session.add(t)
            db.session.commit()
            registrar_auditoria("TiendaTrack API · ticket", "radicado=%s email=%s" % (rad, email))
            return {"ok": True, "radicado": rad, "ticket": tick}, 201
        return {"ok": False, "error": "tipo debe ser registro, pago o ticket"}, 400
    except Exception as _e:
        try:
            db.session.rollback()
        except Exception:
            pass
        print("pf tt evento:", repr(_e)[:300])
        return {"ok": False, "error": "no se pudo procesar el evento"}, 500


'''

# ──────────────────────────── Reemplazos puntuales ────────────────────────────
JS_EXTRA = r""" if(document.getElementById('tt-sel'))return;
 var CONS="__CONS__";
 try{
  var st=document.createElement('style');
  st.textContent='#pf-load{position:fixed;inset:0;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;background:rgba(255,255,255,.82);backdrop-filter:blur(6px);font:600 13px -apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;color:#000F66}#pf-load i{width:34px;height:34px;border:3px solid #cfd8f3;border-top-color:#0020C2;border-radius:50%;animation:pfs .7s linear infinite}@keyframes pfs{to{transform:rotate(360deg)}}@media print{#pf-load{display:none!important}}';
  document.head.appendChild(st);
  document.addEventListener('click',function(e){
   var a=e.target&&e.target.closest?e.target.closest('#tt-sel a[href^="/producto/"]'):null;
   if(!a||a.classList.contains('on')||document.getElementById('pf-load'))return;
   var o=document.createElement('div');o.id='pf-load';o.innerHTML='<i></i><span>Cargando '+a.textContent.trim()+'…</span>';document.body.appendChild(o);
  },true);
  window.addEventListener('pageshow',function(){var o=document.getElementById('pf-load');if(o)o.remove();});
  var mismo=function(u){try{return new URL(u,location.href).origin===location.origin;}catch(x){return false;}};
  if(window.fetch&&!window.__ttFetch){window.__ttFetch=1;var f0=window.fetch;
   window.fetch=function(u,o){try{if((typeof u==='string'||u instanceof URL)&&mismo(u)){o=o||{};var h=new Headers(o.headers||{});h.set('X-Producto',ACT);o.headers=h;}}catch(x){}return f0.call(this,u,o);};}
  if(window.XMLHttpRequest&&!window.__ttXhr){window.__ttXhr=1;var o0=XMLHttpRequest.prototype.open;
   XMLHttpRequest.prototype.open=function(m,u){var r=o0.apply(this,arguments);try{if(mismo(u))this.setRequestHeader('X-Producto',ACT);}catch(x){}return r;};}
 }catch(x){}
"""

REEMPLAZOS = [
    # (descripción, texto viejo, texto nuevo)
    ("banner del selector",
     "b.innerHTML='Vista <b>TiendaTrack</b>: este módulo todavía muestra datos de EduTrack. __LINK__';",
     "b.innerHTML='Vista <b>TiendaTrack</b>: solo ves los datos de TiendaTrack (ingresos, clientes, tickets y metas). RR.HH. y Nómina siguen siendo de toda PROCSIS. __LINK__';"),
    ("JS del selector: cabecera X-Producto y preloader",
     " if(document.getElementById('tt-sel'))return;\n",
     JS_EXTRA),
    ("enlace 'Consolidado' en el selector",
     " box.innerHTML=h;\n",
     " if(CONS===\"1\"){h+='<a href=\"/gerencia/productos\"><i style=\"background:#94a3b8\"></i>Consolidado</a>';}\n box.innerHTML=h;\n"),
    ("inyección de __CONS__ en el selector",
     '.replace("__ACT__", act).replace("__LINK__", enlace.replace("\\\\", ""))',
     '.replace("__ACT__", act).replace("__LINK__", enlace.replace("\\\\", "")).replace("__CONS__", "1" if rol_actual() in _PF_ROLES_CONSOLIDADO else "0")'),
    ("arranque local: migrar antes de inicializar",
     'if __name__ == "__main__":\n    with app.app_context():\n        inicializar_bd()\n',
     'if __name__ == "__main__":\n    with app.app_context():\n        try:\n            _pf_migrar()  # [procsis-multiproducto]\n        except Exception as _pfe:\n            print("multiproducto migrar:", _pfe)\n        inicializar_bd()\n'),
]

ANCLAS = ["_TT_SELECTOR_HTML", "def _tt_cuenta_por", "def _tt_aplicar_tx", "_DCX_GERENCIA = ", "_BIENVENIDA_ROLES = {",
          "class FacturaCobro(", "class TicketPQR(", "class GastoOperativo(", "class MetaVendedor(", "class MetaEmpresa(",
          "class PagoCuotaColegio(", "class ConciliacionBancaria(", "class ContOperacion(", "class LeadCRM(",
          "def _gen_radicado_pqr", "def _gerencia_metricas", "def _cb_pagina", "def _c48_q", "def _rate_limit(", "def registrar_auditoria"]


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

    faltan = [a for a in ANCLAS if a not in texto]
    for desc, viejo, _n in REEMPLAZOS:
        if texto.count(viejo) != 1:
            faltan.append("%s (aparece %d veces, debe ser 1)" % (desc, texto.count(viejo)))
    if faltan:
        sys.exit("Este app.py no coincide con lo esperado; no toqué nada. Falta/cambió:\n  - " + "\n  - ".join(faltan))

    nuevo = texto
    for _desc, viejo, reemplazo in REEMPLAZOS:
        nuevo = nuevo.replace(viejo, reemplazo, 1)

    corte = nuevo.rfind('\nif __name__ == "__main__":')
    if corte < 0:
        sys.exit('No encontré el bloque if __name__ == "__main__": al final de app.py.')
    nuevo = nuevo[:corte] + "\n" + BLOQUE.strip("\n") + "\n\n" + nuevo[corte:]

    try:
        compile(nuevo, str(ruta), "exec")
    except SyntaxError as e:
        sys.exit("El resultado no compila (línea %s: %s). No escribí nada." % (e.lineno, e.msg))

    copia = ruta.with_name(ruta.name + ".bak_multiproducto")
    if copia.exists():
        copia = ruta.with_name(ruta.name + ".bak_multiproducto_" + time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(ruta, copia)
    with open(ruta, "w", encoding="utf-8", newline="") as fh:
        fh.write(nuevo.replace("\n", "\r\n") if crlf else nuevo)
    print("Listo. Multi-Producto aplicado a %s" % ruta)
    print("Copia de seguridad: %s" % copia)
    print("Siguiente paso: despliega y entra una vez como Gerencia; la columna producto_id se crea sola en la primera petición.")


if __name__ == "__main__":
    main()

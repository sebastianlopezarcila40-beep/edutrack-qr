# API de PQR · TiendaTrack ⇄ PROCSIS

Base: `https://procsishq.com` (el dominio público de PROCSIS)
Autenticación: cabecera `X-API-Key: <TIENDATRACK_API_KEY>` (la misma llave de los demás endpoints de TiendaTrack).
Formato: JSON UTF-8. Límite: 60–120 solicitudes por minuto. Todas las respuestas traen `"ok": true|false`.

> Antes de radicar una PQR el negocio debe existir en PROCSIS (evento `registro` de `POST /api/v1/tiendatrack/evento`).

## 1. Catálogo (para llenar los menús de la app)
`GET /api/v1/tiendatrack/pqr/catalogo`
```json
{"ok": true, "tipos": [{"clave":"queja","nombre":"Queja"}, ...],
 "categorias": [{"clave":"falla_app","nombre":"Falla de la aplicación"}, ...], "dias_habiles_respuesta": 15}
```
Tipos: `peticion`, `queja`, `reclamo`, `sugerencia`, `felicitacion`.
Categorías: `facturacion_plan`, `pagos`, `acceso_cuenta`, `falla_app`, `inventario`, `ventas_pos`, `facturacion_dian`, `datos_privacidad`, `demo`, `otro`.

## 2. Radicar una PQR
`POST /api/v1/tiendatrack/pqr`
```json
{"email": "tienda@correo.com", "tipo": "queja", "categoria": "falla_app",
 "asunto": "No cargan los productos", "descripcion": "Desde ayer la lista de productos aparece vacía...",
 "referencia_externa": "pqr-local-123", "prioridad": "ALTA"}
```
Opcionales: `negocio`, `nombre`, `telefono`, `prioridad` (CRITICA|ALTA|MEDIA|BAJA; si no se envía se asigna según la categoría).
`referencia_externa` hace la llamada **idempotente**: si la app reintenta con el mismo valor, devuelve la misma PQR (`"duplicado": true`) en vez de crear otra.

Respuesta `201`:
```json
{"ok": true, "duplicado": false, "radicado": "202633000001", "radicado_formato": "2026-33-000001", "ticket": "AB12CD",
 "estado": "RADICADO", "prioridad": "ALTA", "fecha_limite": "2026-12-09",
 "consulta_publica": "https://procsishq.com/pqr/consulta?clave=202633000001"}
```

## 3. Listar las PQR de un negocio
`GET /api/v1/tiendatrack/pqr?email=tienda@correo.com&limite=20` → `{"ok": true, "total": 2, "pqr": [ {...}, {...} ]}`

## 4. Ver el detalle y la respuesta oficial
`GET /api/v1/tiendatrack/pqr/202633000001?email=tienda@correo.com`
```json
{"ok": true, "pqr": {"radicado":"202633000001","radicado_formato":"2026-33-000001","ticket":"AB12CD","tipo":"Queja",
 "categoria":"falla_app","categoria_nombre":"Falla de la aplicación","asunto":"...","descripcion":"...","estado":"EN PROCESO",
 "abierto": true, "prioridad":"ALTA","fecha_radicacion":"2026-11-17","fecha_limite":"2026-12-09",
 "respuesta":"","fecha_respuesta":"","fecha_cierre":"","referencia_externa":"pqr-local-123"}}
```
`respuesta` queda con el texto oficial cuando Soluciones responde. `abierto` es `false` cuando el caso está resuelto o cerrado.
El correo debe ser el mismo con el que se radicó; si no, responde `404` (así un negocio no ve PQR de otro).
Sugerencia: consultar este endpoint al abrir la pantalla de la PQR (o cada pocos minutos mientras esté abierta); no hace falta webhook.

## 5. Mensaje de seguimiento del cliente
`POST /api/v1/tiendatrack/pqr/202633000001/mensaje`
```json
{"email": "tienda@correo.com", "mensaje": "Adjunto más detalle: ocurre solo en el celular."}
```
Respuesta `201`. Si el caso ya está cerrado responde `409` (hay que radicar una PQR nueva). Máximo 5 mensajes por hora y por caso.

## Códigos de error
`400` datos inválidos · `401` llave incorrecta · `404` cuenta o PQR no encontrada · `409` caso cerrado · `429` demasiadas solicitudes · `500` error interno.

## Ejemplo con curl
```bash
curl -X POST https://procsishq.com/api/v1/tiendatrack/pqr \
  -H "X-API-Key: $TIENDATRACK_API_KEY" -H "Content-Type: application/json" \
  -d '{"email":"tienda@correo.com","tipo":"queja","categoria":"falla_app","asunto":"No cargan los productos","descripcion":"Desde ayer la lista aparece vacía"}'
```

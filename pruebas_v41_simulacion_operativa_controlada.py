"""Simulación operativa controlada v41.
Valida el flujo recomendado antes de producción: maestros, stock, requerimientos, OC,
aprobación, finanzas, tránsito, recepción, liquidación, reportes, cierre y trazabilidad.
"""
import os, tempfile
from datetime import date, timedelta
import app_kardex as app

DB = os.path.join(tempfile.gettempdir(), "kardex_v41_simulacion.db")
try:
    os.remove(DB)
except FileNotFoundError:
    pass

db = app.KardexDB(DB)
USER = "admin"
FECHA = app.today_text()
_hoy = app.parse_fecha(FECHA)
DESDE_MES = _hoy.replace(day=1).strftime(app.DATE_FMT)
# Último día del mes
if _hoy.month == 12:
    _sig = date(_hoy.year + 1, 1, 1)
else:
    _sig = date(_hoy.year, _hoy.month + 1, 1)
HASTA_MES = (_sig - timedelta(days=1)).strftime(app.DATE_FMT)

# 1. Maestros
for data in [
    ("MAT001", "CASCO MINERO ANSI", "UND", "CONSUMIBLE", 5, 2, 0, True, False),
    ("MAT002", "GUANTE CUERO REFORZADO", "PAR", "CONSUMIBLE", 10, 4, 0, True, False),
    ("SRV001", "SERVICIO TRANSPORTE LOCAL", "UND", "SERVICIO", 0, 0, 0, False, True),
]:
    db.add_articulo(*data)

sol = db.add_solicitante("CARLOS ORDINOLA MONTOYA")
db.upsert_proveedor("20600209991", "SMETAL GROUP S.A.C.", "CERRO COLORADO - AREQUIPA", "932478711", "logistica@smetalgroup.com.pe", "01 BCP", "191-123", "002-191-123", "LOGISTICA", "SOLES", usuario=USER)
db.upsert_proveedor("20600000001", "PROVEEDOR DOLARES SAC", "LIMA", "999888777", "ventas@usd.pe", "02 BBVA", "555-USD", "002-555-USD", "VENTAS", "DOLARES", usuario=USER)

# 2. Tipo de cambio para operaciones/reportes con conversión
TC = db.registrar_tipo_cambio_manual(FECHA, 3.70, 3.75, fuente="MANUAL", motivo="SIMULACION", usuario=USER)
assert abs(float(TC["venta"]) - 3.75) < 0.001
if HASTA_MES != FECHA:
    db.registrar_tipo_cambio_manual(HASTA_MES, 3.70, 3.75, fuente="MANUAL", motivo="SIMULACION CIERRE", usuario=USER)

# 3. Stock inicial y bloqueo de stock negativo
db.save_movimiento(FECHA, "CN", "FACT-SIM-001", "101", "01", "CARLOS ORDINOLA MONTOYA", "INGRESO INICIAL SIMULACION", [
    {"codigo": "MAT001", "cantidad": 10, "precio_unitario_sin_igv": 20.00},
    {"codigo": "MAT002", "cantidad": 15, "precio_unitario_sin_igv": 8.50},
], USER, "AQP")
try:
    db.save_movimiento(FECHA, "CO", "SAL-ERROR", "101", "01", "CARLOS ORDINOLA MONTOYA", "SALIDA EXCESIVA", [{"codigo": "MAT001", "cantidad": 999}], USER, "AQP")
    raise AssertionError("Permitió salida con stock negativo")
except ValueError as e:
    assert "stock" in str(e).lower()

# 4. Salida normal
vale_salida = db.save_movimiento(FECHA, "CO", "SAL-SIM-001", "101", "01", "CARLOS ORDINOLA MONTOYA", "SALIDA OPERATIVA", [{"codigo": "MAT001", "cantidad": 2}], USER, "AQP")
assert db.get_stock("MAT001", "AQP") == 8

# 5. Requerimiento con fecha entrega
req = db.create_requerimiento(FECHA, "AQP", "CARLOS ORDINOLA MONTOYA", "01", "101", "REQ SIMULACION", [
    {"codigo": "MAT002", "cantidad": 5, "observacion": "Reposicion"}
], USER, fecha_entrega_solicitada=HASTA_MES)

# 6. OC por requerimiento y exclusión de OC activa
ocs = db.create_ocs_from_requerimiento_multiproveedor(req, [
    {"ruc": "20600209991", "moneda": "SOLES", "cotizacion": "COT-SIM-001", "precios": {"MAT002": 9.00}, "forma_pago": "CONTADO", "precio_incluye_igv": False}
], False, USER)
assert len(ocs) == 1
assert db.requerimiento_tiene_oc_activa(req)["numero"] == ocs[0]
try:
    db.create_ocs_from_requerimiento_multiproveedor(req, [
        {"ruc": "20600209991", "moneda": "SOLES", "cotizacion": "COT-DUP", "precios": {"MAT002": 9.00}}
    ], False, USER)
    raise AssertionError("Permitió duplicar OC activa por requerimiento")
except ValueError as e:
    assert "activa" in str(e).lower()

# 7. Aprobación, pago, tránsito, recepción y liquidación
oc = ocs[0]
ocrow, det = db.get_orden_compra(oc)
estado = db.approve_orden_compra(oc, {det[0]["id"]: 5}, USER)
assert estado == "APROBADA"
estado_pago = db.registrar_pago_oc(oc, FECHA, "OP-SIM-001", "01 BCP", float(ocrow["total"] or 0), "PAGO SIMULACION", USER, numero_factura="F001-000001")
assert estado_pago == "PAGADA"
db.registrar_transito_oc(oc, "FLORES", FECHA, HASTA_MES, "GR-SIM-001", "EN RUTA", USER)
ocrow, det = db.get_orden_compra(oc)
vale_ing, estado_recep = db.recepcionar_orden_compra(oc, FECHA, "GR-REC-001", "RECEPCION TOTAL", {det[0]["id"]: 5}, USER)
assert estado_recep == "ATENDIDA"
db.liquidar_orden_compra(oc, FECHA, "F001-000001", "GR-REC-001", "LIQUIDACION SIM", USER)
ocrow, _ = db.get_orden_compra(oc)
assert ocrow["estado"] == "LIQUIDADA"

# 8. Reportes principales
kardex = db.reporte_kardex_valorizado_articulo(DESDE_MES, HASTA_MES, "AQP", "MAT001", "SOLES")
assert len(kardex) >= 2
compras = db.reporte_compras_por_proveedor(DESDE_MES, HASTA_MES, "TODOS")
assert any(r["ruc"] == "20600209991" for r in compras)
consumo = db.reporte_consumo_cc_ot(DESDE_MES, HASTA_MES, "AQP")
assert any(str(r.get("centro_costo","")).startswith("01") for r in consumo)
aging = db.reporte_cuentas_pagar_aging("TODOS", "", "TODOS")
assert any(r["oc"] == oc for r in aging)
val_sol = db.reporte_valorizacion_mensual(_hoy.year, _hoy.month, "AQP")
assert val_sol
# reporte en dólares debe exigir TC; ya existe TC de la fecha de movimiento
val_usd = db.reporte_kardex_valorizado_articulo(DESDE_MES, HASTA_MES, "AQP", "MAT001", "DOLARES")
assert val_usd

# 9. Trazabilidad visible
traz_oc = db.trazabilidad_documento("OC", oc)
assert traz_oc and any("Pago" in (r.get("detalle") or "") or r.get("accion") == "PAGO" for r in traz_oc)
traz_vale = db.trazabilidad_documento("VALE", vale_salida)
assert traz_vale

# 10. Cierre mensual y bloqueo de movimientos antiguos
cierre = db.cerrar_periodo(_hoy.year, _hoy.month, "AQP", USER, "CIERRE SIMULACION")
try:
    db.save_movimiento(FECHA, "CN", "FACT-CERRADO", "101", "01", "CARLOS ORDINOLA MONTOYA", "NO DEBE", [{"codigo": "MAT001", "cantidad": 1, "precio_unitario_sin_igv": 10}], USER, "AQP")
    raise AssertionError("Permitió ingreso en periodo cerrado")
except ValueError as e:
    assert "periodo" in str(e).lower() or "cerrado" in str(e).lower()

print("SIMULACION OPERATIVA CONTROLADA v41: OK")
print({
    "req": req,
    "oc": oc,
    "vale_salida": vale_salida,
    "vale_ingreso_oc": vale_ing,
    "kardex_rows": len(kardex),
    "compras_rows": len(compras),
    "consumo_rows": len(consumo),
    "aging_rows": len(aging),
})

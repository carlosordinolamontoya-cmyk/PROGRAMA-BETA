import os
import tempfile
from datetime import date
import app_kardex as app
DB_FILE = os.path.join(tempfile.gettempdir(), "kardex_v23_ux_test.db")
today_text = app.today_text

if os.path.exists(DB_FILE):
    os.remove(DB_FILE)

db = app.KardexDB(DB_FILE)
user = "admin"

# Artículos con valorización/decimales
for args in [
    ("A100", "ARTICULO VALORIZADO", "UND", "CONSUMIBLE", 0, 0, 0, True, False),
    ("S100", "SERVICIO NO KARDEX", "UND", "SERVICIO", 0, 0, 0, False, True),
]:
    db.add_articulo(*args)
art = db.get_articulo("A100")
assert art["valoriza"] == 1 and art["permite_decimales"] == 0
serv = db.get_articulo("S100")
assert serv["valoriza"] == 0 and serv["tipo_articulo"] == "SERVICIO"

# Movimiento manual con precio sin IGV
vale = db.save_movimiento(today_text(), "CN", "FACT-001", "101", "01", "ALMACEN", "INGRESO MANUAL VALORIZADO", [
    {"codigo": "A100", "cantidad": 10, "precio_unitario_sin_igv": 12.5}
], user, "AQP")
mov, det = db.get_movimiento_by_vale(vale)
assert abs(det[0]["precio_unitario_sin_igv"] - 12.5) < 0.001
assert abs(det[0]["subtotal_sin_igv"] - 125.0) < 0.001

# Transferencia con guía
tr, vsal, ving = db.transferir_stock(today_text(), "AQP", "MIN", "01", "RESPONSABLE", "TRASLADO DE PRUEBA", [{"codigo":"A100", "cantidad":2}], user, numero_guia="GR-001")
cur = db.conn.cursor(); cur.execute("SELECT numero_guia FROM transferencias WHERE transferencia=?", (tr,))
assert cur.fetchone()["numero_guia"] == "GR-001"

# Banco codificado y otros
pid = db.upsert_proveedor("20123456789", "PROVEEDOR UNO SAC", banco="BCP", numero_cuenta="001", usuario=user)
p1 = db.get_proveedor_by_ruc("20123456789")
assert p1["banco"].startswith("01")
db.upsert_proveedor("20999999991", "PROVEEDOR OTRO SAC", banco="05 OTROS - BANCO PRUEBA", numero_cuenta="999", usuario=user)
p2 = db.get_proveedor_by_ruc("20999999991")
assert p2["banco"].startswith("05 OTROS")

# Modificar OC pendiente y desaprobar aprobada sin movimientos
oc = db.create_orden_compra(today_text(), "AQP", "20123456789", "NACIONAL", "", "COT-1", "SOL", "101", "01", "OC PRUEBA", "SOLES", False, [
    {"codigo":"A100", "descripcion":"ARTICULO VALORIZADO", "um":"UND", "cantidad":3, "precio":20}
], user, forma_pago="CONTADO")
db.update_orden_compra_pendiente(oc, "20999999991", "COT-2", "OBS MOD", "CREDITO 15 DIAS", today_text(), "AQP", user)
ocrow, detoc = db.get_orden_compra(oc)
assert ocrow["ruc"] == "20999999991" and ocrow["forma_pago"] == "CREDITO 15 DIAS"
db.approve_orden_compra(oc, {detoc[0]["id"]: 3}, user)
assert db.desaprobar_orden_compra(oc, "CAMBIO DE PROVEEDOR", user) == "PENDIENTE"
# Ya con pago no debe poder desaprobar
oc2 = db.create_orden_compra(today_text(), "AQP", "20123456789", "NACIONAL", "", "COT-3", "SOL", "101", "01", "OC PAGO", "SOLES", False, [
    {"codigo":"A100", "descripcion":"ARTICULO VALORIZADO", "um":"UND", "cantidad":1, "precio":10}
], user, forma_pago="CONTADO")
ocrow2, detoc2 = db.get_orden_compra(oc2)
db.approve_orden_compra(oc2, {detoc2[0]["id"]: 1}, user)
db.registrar_pago_oc(oc2, today_text(), "OP-1", "01 BCP", 5, "PAGO PARCIAL", user)
try:
    db.desaprobar_orden_compra(oc2, "NO DEBE", user)
    raise AssertionError("No debió permitir desaprobar OC con pago")
except ValueError as e:
    assert "movimientos" in str(e).lower()

# Mapa de proceso y estado operativo de requerimiento no revientan
req = db.create_requerimiento(today_text(), "AQP", "SOL", "01", "101", "REQ TEST", [{"codigo":"A100", "cantidad":1, "observacion":""}], user)
req_info, rows = db.requerimiento_estado_operativo(req)
assert req_info["requerimiento"] == req and rows
mapa = db.oc_mapa_proceso(oc2)
assert mapa["oc"]["numero"] == oc2 and "pagos" in mapa

print("Pruebas v23 UX/funcional: OK")

import os
import tempfile
from datetime import date
import app_kardex as app

DB_FILE = os.path.join(tempfile.gettempdir(), "kardex_v24_val_test.db")
if os.path.exists(DB_FILE):
    os.remove(DB_FILE)

db = app.KardexDB(DB_FILE)
user = "tester"

def t():
    return app.today_text()

# Artículo valorizable y servicio/no valorizable
db.add_articulo("V100", "ARTICULO VALORIZABLE", "UND", "CONSUMIBLE", 0, 0, 0, True, True)
db.add_articulo("SVC1", "SERVICIO NO VALORIZABLE", "UND", "SERVICIO", 0, 0, 0, False, True)

# Entradas valorizadas y salida al promedio ponderado móvil
db.save_movimiento(t(), "CN", "F001", "101", "01", "ALMACEN", "INGRESO 1", [{"codigo":"V100", "cantidad":10, "precio_unitario_sin_igv":10}], user, "AQP")
db.save_movimiento(t(), "CN", "F002", "101", "01", "ALMACEN", "INGRESO 2", [{"codigo":"V100", "cantidad":10, "precio_unitario_sin_igv":20}], user, "AQP")
db.save_movimiento(t(), "CO", "S001", "101", "01", "ALMACEN", "SALIDA PROMEDIO", [{"codigo":"V100", "cantidad":5}], user, "AQP")
# Servicio/no valorizable no debe aparecer en valorización aunque se registre movimiento.
db.save_movimiento(t(), "CN", "F003", "101", "01", "ALMACEN", "SERVICIO", [{"codigo":"SVC1", "cantidad":1, "precio_unitario_sin_igv":1000}], user, "AQP")
# Transferencia debe trasladar costo promedio desde origen hacia destino.
db.transferir_stock(t(), "AQP", "MIN", "01", "ALMACEN", "TRANSFERENCIA COSTO", [{"codigo":"V100", "cantidad":5}], user, numero_guia="GR-VAL-001")

hoy = date.today()
rows = db.reporte_valorizacion_mensual(hoy.year, hoy.month, "TODOS")
assert rows, "Debe devolver filas valorizadas"
assert all(r["codigo"] != "SVC1" for r in rows), "No debe incluir artículos no valorizables"
by_key = {(r["almacen_codigo"], r["codigo"]): r for r in rows}
assert ("AQP", "V100") in by_key and ("MIN", "V100") in by_key

aqp = by_key[("AQP", "V100")]
minr = by_key[("MIN", "V100")]
# AQP: Entradas 20 unidades = 300; salida manual 5 al promedio 15 = 75; transferencia salida 5 al promedio 15 = 75; final 10 = 150.
assert abs(aqp["entradas"] - 20) < 0.001
assert abs(aqp["valor_entradas"] - 300) < 0.001
assert abs(aqp["salidas"] - 10) < 0.001
assert abs(aqp["valor_salidas"] - 150) < 0.001
assert abs(aqp["stock_final"] - 10) < 0.001
assert abs(aqp["valor_final"] - 150) < 0.001
assert abs(aqp["costo_promedio_final"] - 15) < 0.001
# MIN: Ingreso por transferencia 5 al costo 15 = 75.
assert abs(minr["entradas"] - 5) < 0.001
assert abs(minr["valor_entradas"] - 75) < 0.001
assert abs(minr["stock_final"] - 5) < 0.001
assert abs(minr["valor_final"] - 75) < 0.001

print("Pruebas v24 valorización mensual: OK")
rows_min = db.reporte_valorizacion_mensual(hoy.year, hoy.month, "MIN")
assert len(rows_min) == 1 and rows_min[0]["almacen_codigo"] == "MIN"
assert abs(rows_min[0]["valor_final"] - 75) < 0.001
rows_aqp = db.reporte_valorizacion_mensual(hoy.year, hoy.month, "AQP")
assert len(rows_aqp) == 1 and rows_aqp[0]["almacen_codigo"] == "AQP"
assert abs(rows_aqp[0]["valor_final"] - 150) < 0.001

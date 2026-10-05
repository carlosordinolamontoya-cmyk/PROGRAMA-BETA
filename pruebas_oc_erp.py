"""
Pruebas automáticas del módulo ERP/MRP de Órdenes de Compra.
Ejecutar desde la carpeta del proyecto:
    python pruebas_oc_erp.py

Valida dos pasadas completas con base temporal:
- proveedor
- OC nacional manual
- aprobación parcial
- bloqueo de recepción mayor a lo aprobado
- recepción parcial y total con Kardex
- pago por finanzas
- tránsito
- liquidación
- orden de servicio sin Kardex
- OC por requerimiento
- anulación de OC y liberación del requerimiento
"""
import os
import tempfile
import app_kardex as app


def run_pass(pass_no: int):
    db_path = os.path.join(tempfile.gettempdir(), f"kardex_prueba_oc_erp_{pass_no}.db")
    if os.path.exists(db_path):
        os.remove(db_path)
    db = app.KardexDB(db_path)
    db.create_user('almacen', 'USUARIO ALMACEN PRUEBA', '1234', 'ALMACEN')
    db.create_user('tester', 'USUARIO OPERADOR PRUEBA', '1234', 'OPERADOR')
    db.set_action_permission(db.user_id_by_login('tester'), 'aprobar_oc', True, 'admin')
    db.add_articulo("A100", "CASCO MINERO", "UND", "CONSUMIBLE", 0, 0, 0)
    db.add_articulo("A200", "GUANTE CUERO", "PAR", "CONSUMIBLE", 0, 0, 0)
    db.upsert_proveedor("10431801413", "RAMIREZ GUADAMUR RICHARD RAUL", "SEBASTIAN BARRANCA", "975683377", banco="BCP", numero_cuenta="001-123", estado_contribuyente="ACTIVO", condicion_domicilio="HABIDO", fuente_ruc="MANUAL", usuario="tester")
    prov = db.get_proveedor_by_ruc("10431801413", active_only=False)
    assert prov["estado_contribuyente"] == "ACTIVO"
    assert prov["condicion_domicilio"] == "HABIDO"

    # Tipo de cambio manual: si la API no está disponible, el sistema debe trabajar con TC registrado localmente.
    db.registrar_tipo_cambio_manual(app.today_text(), 3.70, 3.80, "MANUAL", "TC DE PRUEBA", "finanzas")
    tc = db.obtener_tipo_cambio_para_oc(app.fecha_iso_consulta_from_text(app.today_text()), "DOLARES", "tester", permitir_api=False)
    assert round(tc["venta"], 2) == 3.80

    oc_usd = db.create_orden_compra(
        app.today_text(), "AQP", "10431801413", "NACIONAL", "", "COT-USD", "CASIANO", "26301", "04", "COMPRA USD",
        "DOLARES", False,
        [{"codigo": "A100", "cantidad": 1, "precio": 10}],
        "tester",
    )
    oc_usd_row, _ = db.get_orden_compra(oc_usd)
    assert oc_usd_row["moneda"] == "DOLARES"
    assert round(float(oc_usd_row["tipo_cambio_venta"]), 2) == 3.80

    oc = db.create_orden_compra(
        app.today_text(), "AQP", "10431801413", "NACIONAL", "", "COT-1", "CASIANO", "26302", "04", "COMPRA PRUEBA",
        "SOLES", False,
        [{"codigo": "A100", "cantidad": 10, "precio": 100}, {"codigo": "A200", "cantidad": 5, "precio": 20}],
        "tester",
    )
    oc_row, det = db.get_orden_compra(oc)
    assert oc_row["estado"] == "PENDIENTE"
    assert round(float(oc_row["total"]), 2) == 1298.00

    estado = db.approve_orden_compra(oc, {det[0]["id"]: 8, det[1]["id"]: 5}, "admin", "APRUEBA MENOS CASCOS")
    assert estado == "APROBADA PARCIAL"

    try:
        db.recepcionar_orden_compra(oc, app.today_text(), "GR-EXCESO", "INGRESO EXCESO", {det[0]["id"]: 9}, "almacen")
        raise AssertionError("El sistema permitió recibir más de lo aprobado")
    except ValueError:
        pass

    vale, estado = db.recepcionar_orden_compra(oc, app.today_text(), "GR-1", "INGRESO PARCIAL", {det[0]["id"]: 3, det[1]["id"]: 5}, "almacen")
    assert estado == "ATENDIDA PARCIALMENTE"
    assert db.get_stock("A100", "AQP") == 3

    vale2, estado2 = db.recepcionar_orden_compra(oc, app.today_text(), "GR-2", "INGRESO FINAL", {det[0]["id"]: 5}, "almacen")
    assert estado2 == "ATENDIDA"
    assert db.get_stock("A100", "AQP") == 8

    oc2 = db.create_orden_compra(app.today_text(), "MIN", "10431801413", "NACIONAL", "", "COT-2", "CASIANO", "26303", "04", "COMPRA DOS", "SOLES", False, [{"codigo": "A100", "cantidad": 2, "precio": 50}], "tester")
    oc2_row, det2 = db.get_orden_compra(oc2)
    db.approve_orden_compra(oc2, {det2[0]["id"]: 2}, "admin")
    db.registrar_pago_oc(oc2, app.today_text(), "OP123", "BCP", 118, "PAGO TOTAL", "finanzas")
    db.registrar_transito_oc(oc2, "FLORES", app.today_text(), app.today_text(), "GUIA-T", "EN RUTA", "logistica")
    oc2_row, _ = db.get_orden_compra(oc2)
    assert oc2_row["estado"] == "EN TRANSITO"

    db.liquidar_orden_compra(oc, app.today_text(), "F001", "GR-1/GR-2", "CONFORME", "liquidador")
    oc_row, _ = db.get_orden_compra(oc)
    assert oc_row["estado"] == "LIQUIDADA"

    osn = db.create_orden_compra(app.today_text(), "AQP", "10431801413", "SERVICIO", "", "COT-S", "CASIANO", "26304", "04", "SERVICIO ALQUILER", "SOLES", False, [{"descripcion": "ALQUILER DE CARGADOR FRONTAL", "um": "UND", "cantidad": 1, "precio": 1280}], "tester")
    os_row, os_det = db.get_orden_compra(osn)
    db.approve_orden_compra(osn, {os_det[0]["id"]: 1}, "admin")
    try:
        db.recepcionar_orden_compra(osn, app.today_text(), "GR-S", "NO DEBE INGRESAR", {os_det[0]["id"]: 1}, "almacen")
        raise AssertionError("El sistema permitió ingreso Kardex de una orden de servicio")
    except ValueError:
        pass
    db.liquidar_orden_compra(osn, app.today_text(), "F-S", "", "SERVICIO CONFORME", "liquidador")

    req = db.create_requerimiento(app.today_text(), "AQP", "CASIANO", "04", "26305", "REQ PARA OC", [{"codigo": "A100", "cantidad": 2}], "logistica")
    oc_req = db.create_oc_from_requerimiento(req, "10431801413", "COT-R", {"A100": 10}, "SOLES", False, "logistica")
    try:
        db.create_oc_from_requerimiento(req, "10431801413", "COT-R2", {"A100": 12}, "SOLES", False, "logistica")
        raise AssertionError("El sistema permitió duplicar OC activa por requerimiento")
    except ValueError:
        pass
    db.anular_orden_compra(oc_req, "ERROR DE COTIZACION", "admin")
    oc_req2 = db.create_oc_from_requerimiento(req, "10431801413", "COT-R2", {"A100": 12}, "SOLES", False, "logistica")
    assert oc_req2 != oc_req

    return {
        "oc_manual": oc,
        "oc_transito": oc2,
        "orden_servicio": osn,
        "oc_req_recreada": oc_req2,
        "vale_1": vale,
        "vale_2": vale2,
    }


if __name__ == "__main__":
    for n in (1, 2):
        result = run_pass(n)
        print(f"PASADA {n}: OK -> {result}")
    print("Todas las pruebas del módulo OC/ERP finalizaron correctamente.")

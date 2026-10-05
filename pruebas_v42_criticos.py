"""Pruebas dirigidas V42 para errores críticos detectados en auditoría.

Nunca usa ni elimina kardex_almacen.db. Toda la prueba corre en una carpeta temporal.
"""
import os
import tempfile
from pathlib import Path
import app_kardex as app


def expect_value_error(fn, contains=""):
    try:
        fn()
    except ValueError as exc:
        if contains:
            assert contains.lower() in str(exc).lower(), (contains, str(exc))
        return str(exc)
    raise AssertionError("Se esperaba ValueError")


def setup_db(path):
    db = app.KardexDB(path)
    db.add_articulo("V42-001", "ARTICULO CONTROL V42", "UND", "CONSUMIBLE", 0, 0, 0, True, False)
    db.upsert_proveedor(
        "20600209991", "PROVEEDOR CONTROL V42 SAC", "AREQUIPA", "999999999",
        "control@example.com", "01 BCP", "001", "CCI001", "LOGISTICA", "SOLES", usuario="admin"
    )
    return db


def create_approved_oc(db, almacen="AQP", cot="COT-V42", qty=1, price=100):
    numero = db.create_orden_compra(
        app.today_text(), almacen, "20600209991", "NACIONAL", "", cot, "SOLICITANTE V42",
        "4201", "01", "CONTROL V42", "SOLES", False,
        [{"codigo": "V42-001", "cantidad": qty, "precio": price, "centro_costo_codigo": "01", "ot": "OT-4201", "solicitante": "SOLICITANTE V42"}],
        "admin",
    )
    oc, det = db.get_orden_compra(numero)
    db.approve_orden_compra(numero, {det[0]["id"]: qty}, "admin")
    return numero


def test_recepcion_respeta_almacen(db):
    db.create_user("alm_aqp", "ALMACEN SOLO AQP", "1234", "ALMACEN")
    uid = db.user_id_by_login("alm_aqp")
    db.set_user_almacen_permissions(uid, ["AQP"])

    oc_min = create_approved_oc(db, "MIN", "COT-MIN-V42", 2, 10)
    _oc, det = db.get_orden_compra(oc_min)
    expect_value_error(
        lambda: db.recepcionar_orden_compra(oc_min, app.today_text(), "GR-MIN-V42", "NO AUTORIZADO", {det[0]["id"]: 2}, "alm_aqp"),
        "permiso",
    )
    assert db.get_stock("V42-001", "MIN") == 0
    cur = db.conn.cursor(); cur.execute("SELECT COUNT(*) n FROM orden_compra_recepciones WHERE oc_id=(SELECT id FROM orden_compra WHERE numero=?)", (oc_min,))
    assert cur.fetchone()["n"] == 0

    oc_aqp = create_approved_oc(db, "AQP", "COT-AQP-V42", 1, 10)
    _oc, det = db.get_orden_compra(oc_aqp)
    vale, estado = db.recepcionar_orden_compra(oc_aqp, app.today_text(), "GR-AQP-V42", "AUTORIZADO", {det[0]["id"]: 1}, "alm_aqp")
    assert vale and estado == "ATENDIDA"
    assert db.get_stock("V42-001", "AQP") == 1


def test_transferencia_no_es_consumo(db):
    db.save_movimiento(app.today_text(), "CN", "STOCK-V42", "4202", "01", "SOLICITANTE V42", "STOCK INICIAL", [{"codigo":"V42-001", "cantidad":10, "precio_unitario_sin_igv":5}], "admin", "AQP")
    db.transferir_stock(app.today_text(), "AQP", "MIN", "01", "SOLICITANTE V42", "TRASLADO INTERNO", [{"codigo":"V42-001", "cantidad":4}], "admin", numero_guia="GR-TR-V42")
    rows = db.reporte_consumo_cc_ot(almacen_codigo="AQP")
    # Puede haber otros movimientos de otros tests en la misma DB; ninguno de la transferencia debe sumar 4.
    total = sum(float(r["cantidad_consumida"]) for r in rows if r["articulo"].startswith("V42-001 -") and str(r["ot"]).endswith("4202"))
    assert abs(total) < 0.0001, total

    db.save_movimiento(app.today_text(), "CO", "CONSUMO-V42", "4202", "01", "SOLICITANTE V42", "CONSUMO REAL", [{"codigo":"V42-001", "cantidad":2}], "admin", "AQP")
    rows = db.reporte_consumo_cc_ot(almacen_codigo="AQP")
    total = sum(float(r["cantidad_consumida"]) for r in rows if r["articulo"].startswith("V42-001 -") and str(r["ot"]).endswith("4202"))
    assert abs(total - 2.0) < 0.0001, total


def test_pago_parcial_misma_factura(db):
    oc = create_approved_oc(db, "AQP", "COT-PAGO-V42", 1, 100)
    assert db.registrar_pago_oc(oc, app.today_text(), "OP-V42-01", "01 BCP", 40, "CUOTA 1", "admin", numero_factura="F001-V42") == "PAGO PARCIAL"
    assert db.registrar_pago_oc(oc, app.today_text(), "OP-V42-02", "01 BCP", 78, "CUOTA 2", "admin", numero_factura="F001-V42") == "PAGADA"
    ocrow, _ = db.get_orden_compra(oc)
    assert ocrow["estado_pago"] == "PAGADA"
    pagos = db.pagos_oc(ocrow["id"])
    assert len(pagos) == 2
    assert all(p["numero_factura"] == "F001-V42" for p in pagos)

    # La protección contra doble operación se mantiene.
    expect_value_error(lambda: db.registrar_pago_oc(oc, app.today_text(), "OP-V42-02", "01 BCP", 78, "DUP", "admin", numero_factura="F001-V42"), "duplicado")

    # La misma factura en otra OC del mismo proveedor continúa bloqueada.
    oc2 = create_approved_oc(db, "AQP", "COT-PAGO2-V42", 1, 10)
    expect_value_error(lambda: db.registrar_pago_oc(oc2, app.today_text(), "OP-V42-03", "01 BCP", 5, "OTRA OC", "admin", numero_factura="F001-V42"), "factura duplicada")


def test_pruebas_no_apuntan_a_bd_produccion(project_dir):
    offenders = []
    for path in Path(project_dir).glob("pruebas_*.py"):
        if path.name == Path(__file__).name:
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        destructive_call = "os.remove(" + "app.DB_FILE" + ")"
        direct_alias = "DB_FILE = " + "app.DB_FILE"
        if destructive_call in text or direct_alias in text:
            offenders.append(path.name)
    assert not offenders, f"Pruebas destructivas contra BD de producción: {offenders}"


def run():
    project_dir = Path(__file__).resolve().parent
    prod_db = project_dir / "kardex_almacen.db"
    before = prod_db.read_bytes() if prod_db.exists() else None
    with tempfile.TemporaryDirectory(prefix="kardex_v42_criticos_") as tmp:
        db = setup_db(os.path.join(tmp, "v42.db"))
        test_recepcion_respeta_almacen(db)
        test_transferencia_no_es_consumo(db)
        test_pago_parcial_misma_factura(db)
        test_pruebas_no_apuntan_a_bd_produccion(project_dir)
        db.conn.close()
    after = prod_db.read_bytes() if prod_db.exists() else None
    assert before == after, "La prueba modificó la base kardex_almacen.db"
    print("PRUEBAS V42 ERRORES CRITICOS: OK")


if __name__ == "__main__":
    run()

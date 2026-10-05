"""Pruebas V43: seguridad, permisos, backups, restauracion y migracion de claves.

Todas las pruebas usan directorios temporales. Nunca modifican kardex_almacen.db.
"""
import hashlib
import os
import sqlite3
import tempfile
from pathlib import Path

import app_kardex as app


def expect_value_error(fn, contains=""):
    try:
        fn()
    except (ValueError, RuntimeError) as exc:
        if contains:
            assert contains.lower() in str(exc).lower(), (contains, str(exc))
        return str(exc)
    raise AssertionError("Se esperaba error")


def create_base(td):
    db = app.KardexDB(os.path.join(td, "v43.db"))
    db.create_user("logi43", "LOGISTICA V43", "1234", "LOGISTICA")
    db.create_user("fin43", "FINANZAS V43", "1234", "FINANZAS")
    db.create_user("alm43", "ALMACEN V43", "1234", "ALMACEN")
    db.create_user("consulta43", "CONSULTA V43", "1234", "CONSULTA")
    db.add_articulo("V43-001", "ARTICULO SEGURIDAD V43", "UND", "CONSUMIBLE", 0, 0, 0, True, False)
    return db


def create_provider(db, ruc="20600209991", usuario="logi43"):
    return db.upsert_proveedor(
        ruc, "PROVEEDOR V43 SAC", "AREQUIPA", "999999999", "v43@example.com",
        "01 BCP", "001", "CCI001", "CONTACTO", "SOLES", usuario=usuario,
    )


def create_oc(db, usuario="logi43", cot="COT-V43", qty=1, price=100):
    return db.create_orden_compra(
        app.today_text(), "AQP", "20600209991", "NACIONAL", "", cot,
        "SOLICITANTE V43", "4301", "01", "CONTROL V43", "SOLES", False,
        [{"codigo": "V43-001", "cantidad": qty, "precio": price,
          "centro_costo_codigo": "01", "ot": "OT-4301", "solicitante": "SOLICITANTE V43"}],
        usuario,
    )


def test_admin123_y_migracion_hash(db):
    # La clave pedida para beta se conserva.
    assert db.verify_user("admin", "admin123")
    assert db.verify_user("admin", "otra") is None

    # Simular un hash legacy V42 y comprobar migracion transparente al iniciar sesion.
    legacy_salt = "0123456789abcdef0123456789abcdef"
    legacy_hash = hashlib.sha256((legacy_salt + "admin123").encode("utf-8")).hexdigest()
    db.conn.execute("UPDATE usuarios SET password_hash=?, salt=? WHERE usuario='admin'", (legacy_hash, legacy_salt))
    db.conn.commit()
    before = db.conn.execute("SELECT password_hash FROM usuarios WHERE usuario='admin'").fetchone()[0]
    assert len(before) == 64 and not before.startswith("pbkdf2_sha256$")
    assert db.verify_user("admin", "admin123")
    after = db.conn.execute("SELECT password_hash FROM usuarios WHERE usuario='admin'").fetchone()[0]
    assert after.startswith("pbkdf2_sha256$")
    assert db.verify_user("admin", "admin123")


def test_permisos_efectivos(db):
    admin_id = db.user_id_by_login("admin")
    logi_id = db.user_id_by_login("logi43")
    fin_id = db.user_id_by_login("fin43")
    alm_id = db.user_id_by_login("alm43")

    # Solo ADMIN puede cambiar permisos.
    expect_value_error(lambda: db.set_action_permission(logi_id, "aprobar_oc", True, "logi43"), "solo un admin")

    # ADMIN sigue siendo superusuario aunque su fila diga 0.
    db.set_action_permission(admin_id, "aprobar_oc", False, "admin")
    assert db.has_action_permission("admin", "aprobar_oc") is True

    # Gestion de proveedores: permiso real en backend.
    assert db.has_action_permission("logi43", "modificar_proveedor")
    create_provider(db)
    db.set_action_permission(logi_id, "modificar_proveedor", False, "admin")
    expect_value_error(lambda: create_provider(db, usuario="logi43"), "permiso denegado")
    db.set_action_permission(logi_id, "modificar_proveedor", True, "admin")
    db.upsert_proveedor("20600209991", "PROVEEDOR V43 ACTUALIZADO SAC", usuario="logi43")
    assert "ACTUALIZADO" in db.get_proveedor_by_ruc("20600209991", active_only=False)["razon_social"]

    # Crear/modificar OC.
    db.set_action_permission(logi_id, "crear_oc", False, "admin")
    expect_value_error(lambda: create_oc(db, "logi43", "COT-BLOCK"), "permiso denegado")
    db.set_action_permission(logi_id, "crear_oc", True, "admin")
    oc = create_oc(db, "logi43", "COT-OK")
    db.set_action_permission(logi_id, "crear_oc", False, "admin")
    expect_value_error(
        lambda: db.update_orden_compra_pendiente(oc, "20600209991", "COT-EDIT", "EDIT", "CONTADO", "", "", "logi43"),
        "permiso denegado",
    )
    db.set_action_permission(logi_id, "crear_oc", True, "admin")

    # Aprobacion no depende ya de ser ADMIN: depende del permiso otorgado.
    _, det = db.get_orden_compra(oc)
    expect_value_error(lambda: db.approve_orden_compra(oc, {det[0]["id"]: 1}, "logi43"), "permiso denegado")
    db.set_action_permission(logi_id, "aprobar_oc", True, "admin")
    assert db.approve_orden_compra(oc, {det[0]["id"]: 1}, "logi43") == "APROBADA"

    # Anulacion respeta el permiso.
    oc2 = create_oc(db, "logi43", "COT-ANULAR")
    db.set_action_permission(logi_id, "anular_oc", False, "admin")
    expect_value_error(lambda: db.anular_orden_compra(oc2, "PRUEBA", "logi43"), "permiso denegado")
    db.set_action_permission(logi_id, "anular_oc", True, "admin")
    assert db.anular_orden_compra(oc2, "PRUEBA AUTORIZADA", "logi43") == "ANULADA"

    # Reverso de vale respeta permiso de accion.
    vale = db.save_movimiento(
        app.today_text(), "CN", "CN-V43", "4302", "01", "SOL V43", "INGRESO",
        [{"codigo": "V43-001", "cantidad": 5, "precio_unitario_sin_igv": 2}], "admin", "AQP"
    )
    db.set_action_permission(alm_id, "reversar_vale", False, "admin")
    expect_value_error(lambda: db.reverse_movimiento(vale, "PRUEBA", "alm43"), "permiso denegado")
    db.set_action_permission(alm_id, "reversar_vale", True, "admin")
    assert db.reverse_movimiento(vale, "PRUEBA AUTORIZADA", "alm43")

    # Pago respeta permiso de finanzas.
    db.set_action_permission(fin_id, "registrar_pago", False, "admin")
    expect_value_error(
        lambda: db.registrar_pago_oc(oc, app.today_text(), "OP-V43-1", "01 BCP", 10, "TEST", "fin43", numero_factura="F-V43"),
        "permiso",
    )
    db.set_action_permission(fin_id, "registrar_pago", True, "admin")
    estado = db.registrar_pago_oc(oc, app.today_text(), "OP-V43-1", "01 BCP", 10, "TEST", "fin43", numero_factura="F-V43")
    assert estado in ("PAGO PARCIAL", "PAGADA")

    # Cierre de mes respeta permiso independiente.
    assert not db.has_action_permission("logi43", "cerrar_mes")
    expect_value_error(lambda: db.cerrar_periodo(2020, 1, "AQP", "logi43", "TEST"), "permiso")
    db.set_action_permission(logi_id, "cerrar_mes", True, "admin")
    db.cerrar_periodo(2020, 1, "AQP", "logi43", "TEST V43")
    row = db.conn.execute("SELECT cerrado_por FROM cierre_mensual_almacen WHERE anio=2020 AND mes=1 AND almacen_codigo='AQP'").fetchone()
    assert row and row[0] == "logi43"


def test_backup_restore(db, td):
    # Verificacion de esquema e integridad.
    assert db.get_schema_version() >= 43
    assert db.integrity_check() == "ok"

    # El backup diario se crea una sola vez por fecha.
    daily1 = db.ensure_daily_backup()
    daily2 = db.ensure_daily_backup()
    assert daily1 == daily2 and os.path.isfile(daily1)
    app._v43_integrity_check_file(daily1)

    # Backup manual + restauracion funcional.
    db.conn.execute("INSERT OR REPLACE INTO sistema_config(clave,valor,descripcion,modificado_por,modificado_en) VALUES ('V43_TEST','ANTES','TEST','admin',?)", (app.now_text(),))
    db.conn.commit()
    backup = os.path.join(td, "backup_control_v43.db")
    db.backup_database(backup, reason="PRUEBA", usuario="admin")
    assert os.path.isfile(backup)
    app._v43_integrity_check_file(backup)

    db.conn.execute("UPDATE sistema_config SET valor='DESPUES' WHERE clave='V43_TEST'")
    db.conn.commit()
    assert db.conn.execute("SELECT valor FROM sistema_config WHERE clave='V43_TEST'").fetchone()[0] == "DESPUES"

    emergency = db.restore_database(backup, usuario="admin")
    assert os.path.isfile(emergency)
    assert db.conn.execute("SELECT valor FROM sistema_config WHERE clave='V43_TEST'").fetchone()[0] == "ANTES"
    assert db.integrity_check() == "ok"
    assert db.get_schema_version() >= 43

    # Un archivo invalido no puede reemplazar la BD.
    bad = os.path.join(td, "corrupto.db")
    Path(bad).write_bytes(b"NO ES SQLITE")
    before = db.conn.execute("SELECT valor FROM sistema_config WHERE clave='V43_TEST'").fetchone()[0]
    expect_value_error(lambda: db.restore_database(bad, usuario="admin"), "database")
    after = db.conn.execute("SELECT valor FROM sistema_config WHERE clave='V43_TEST'").fetchone()[0]
    assert before == after == "ANTES"


def test_pre_migration_backup(td):
    # Crear una BD V42-like sin schema_version y comprobar respaldo pre-migracion.
    legacy_path = os.path.join(td, "legacy_v42.db")
    con = sqlite3.connect(legacy_path)
    con.execute("CREATE TABLE demo(id INTEGER PRIMARY KEY, dato TEXT)")
    con.execute("INSERT INTO demo(dato) VALUES ('ORIGINAL')")
    con.commit(); con.close()

    # La inicializacion V43 migra y crea respaldo previo consistente.
    db = app.KardexDB(legacy_path)
    assert db.get_schema_version() >= 43
    pre_dir = os.path.join(td, "backups", "pre_migracion")
    files = [os.path.join(pre_dir, f) for f in os.listdir(pre_dir) if f.endswith('.db')]
    assert files, "No se creo respaldo pre-migracion"
    pre = files[0]
    app._v43_integrity_check_file(pre)
    con = sqlite3.connect(pre)
    row = con.execute("SELECT dato FROM demo").fetchone(); con.close()
    assert row[0] == "ORIGINAL"
    db.conn.close()


def run():
    project_dir = Path(__file__).resolve().parent
    prod_db = project_dir / "kardex_almacen.db"
    before = prod_db.read_bytes() if prod_db.exists() else None
    with tempfile.TemporaryDirectory(prefix="kardex_v43_seguridad_") as td:
        db = create_base(td)
        test_admin123_y_migracion_hash(db)
        test_permisos_efectivos(db)
        test_backup_restore(db, td)
        db.conn.close()
    with tempfile.TemporaryDirectory(prefix="kardex_v43_migracion_") as td:
        test_pre_migration_backup(td)
    after = prod_db.read_bytes() if prod_db.exists() else None
    assert before == after, "La prueba V43 modifico kardex_almacen.db"
    print("PRUEBAS V43 SEGURIDAD/BACKUP: OK")


if __name__ == "__main__":
    run()

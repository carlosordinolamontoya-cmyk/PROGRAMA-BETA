import os
import re
import sqlite3
import tempfile
import threading
from pathlib import Path

import app_kardex as app
from erp_core import D, q_money, q_qty, q_unit, q_tc
from erp_core.money import calc_price_values


def new_db(path):
    db = app.KardexDB(path)
    assert db.get_schema_version() == 44
    return db


def test_correlativos_anuales(td):
    path = os.path.join(td, "corr.db")
    db = new_db(path)
    assert db.next_requerimiento("AQP", 2026) == "REQ-2026-000001"
    assert db.next_requerimiento("MIN", 2026) == "REQ-2026-000002"  # global, no depende de almacen
    assert db.next_requerimiento("AQP", 2027) == "REQ-2027-000001"
    assert db.next_oc_number("NACIONAL", "AQP", 2026) == "OC-2026-000001"
    assert db.next_oc_number("SERVICIO", "MIN", 2026) == "OS-2026-000001"
    assert db.next_vale("CN", "AQP", 2026) == "VALE-2026-000001"
    assert db.next_vale("CO", "MIN", 2026) == "VALE-2026-000002"
    assert db.next_vale("AJ", "AQP", 2026) == "AJ-2026-000001"
    assert db.next_transferencia(2026) == "TRF-2026-000001"
    db.conn.close()

    # Segunda instancia contra el mismo archivo: debe continuar y no duplicar.
    db2 = new_db(path)
    assert db2.next_oc_number("NACIONAL", "MIN", 2026) == "OC-2026-000002"
    assert db2.next_vale("CN", "MIN", 2026) == "VALE-2026-000003"
    db2.conn.close()


def test_correlativos_concurrencia_5_usuarios(td):
    path = os.path.join(td, "corr_concurrente.db")
    con = sqlite3.connect(path)
    app.AnnualCorrelativeService.ensure_table(con)
    con.commit(); con.close()
    results, errors = [], []
    lock = threading.Lock()

    def worker():
        try:
            local = sqlite3.connect(path, timeout=20)
            for _ in range(25):
                numero = app.AnnualCorrelativeService(path, local, 20000).next("OC", 2026)
                with lock:
                    results.append(numero)
            local.close()
        except Exception as exc:
            with lock:
                errors.append(repr(exc))

    threads = [threading.Thread(target=worker) for _ in range(5)]
    for t in threads: t.start()
    for t in threads: t.join()
    assert not errors, errors
    assert len(results) == 125
    assert len(set(results)) == 125
    correlativos = sorted(int(x.rsplit("-", 1)[1]) for x in results)
    assert correlativos == list(range(1, 126))


def test_decimal_policy():
    assert q_qty("1.23456") == D("1.235")
    assert q_unit("12.34567") == D("12.3457")
    assert q_tc("3.74256") == D("3.7426")
    assert q_money("100.105") == D("100.11")
    ps, pc, sub, igv, total = calc_price_values("3", "0.1000", False)
    assert ps == D("0.1000")
    assert pc == D("0.1180")
    assert sub == D("0.30")
    assert igv == D("0.05")
    assert total == D("0.35")
    # Caso clasico que con float suele mostrar artefactos binarios.
    assert q_money(D("0.10") + D("0.20")) == D("0.30")


def test_timestamp_iso_and_migration(td):
    path = os.path.join(td, "dates.db")
    db = new_db(path)
    db.audit("qa", "1", "TEST", "admin", "timestamp v44")
    row = db.conn.execute("SELECT fecha_hora FROM auditoria ORDER BY id DESC LIMIT 1").fetchone()
    assert re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", row["fecha_hora"])

    db.conn.execute(
        "INSERT INTO auditoria(tabla,registro,accion,usuario,fecha_hora,detalle) VALUES ('qa','legacy','TEST','admin','11/08/2026 23:59:58','legacy')"
    )
    db.conn.commit()
    changed = app._v44_normalize_internal_timestamps(db.conn)
    db.conn.commit()
    legacy = db.conn.execute("SELECT fecha_hora FROM auditoria WHERE registro='legacy'").fetchone()[0]
    assert legacy == "2026-08-11 23:59:58"
    assert changed >= 1
    db.conn.close()


def test_indices_y_sqlite(td):
    path = os.path.join(td, "indexes.db")
    db = new_db(path)
    names = {r[0] for r in db.conn.execute("SELECT name FROM sqlite_master WHERE type='index'")}
    required = {
        "idx_movimientos_fecha_almacen",
        "idx_movdet_codigo",
        "idx_oc_fecha_estado",
        "idx_oc_ruc",
        "idx_pagos_oc_fecha",
        "idx_auditoria_fecha",
    }
    assert required.issubset(names)
    timeout = db.conn.execute("PRAGMA busy_timeout").fetchone()[0]
    assert int(timeout) >= 15000
    db.conn.close()


def test_integrated_document_year(td):
    path = os.path.join(td, "integrated.db")
    db = new_db(path)
    # Requerimiento: la fecha del documento manda el anio del correlativo.
    db.add_articulo("V44-001", "ARTICULO QA V44", "UND", "CONSUMIBLE")
    req = db.create_requerimiento(
        "10/08/2026", "AQP", "QA V44", "01", "", "PRUEBA V44",
        [{"codigo": "V44-001", "cantidad": "1.000"}], "admin"
    )
    assert req.startswith("REQ-2026-")
    vale = db.save_movimiento(
        "10/08/2026", "CN", "DOC-V44-1", "", "01", "QA V44", "INGRESO QA",
        [{"codigo": "V44-001", "cantidad": "2", "precio_unitario_sin_igv": "10.12345"}],
        "admin", "AQP"
    )
    assert vale.startswith("VALE-2026-")
    db.conn.close()


def run():
    with tempfile.TemporaryDirectory(prefix="erp_v44_") as td:
        test_correlativos_anuales(td)
        test_correlativos_concurrencia_5_usuarios(td)
        test_decimal_policy()
        test_timestamp_iso_and_migration(td)
        test_indices_y_sqlite(td)
        test_integrated_document_year(td)
    print("PRUEBAS V44 ROBUSTEZ: OK")


if __name__ == "__main__":
    run()

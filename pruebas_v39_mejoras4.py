"""Pruebas v39 - MEJORAS 4
Valida:
- Requerimientos con OC activa no pueden volver a generar OC.
- OC anulada libera el requerimiento.
- Word de OC coloca símbolos de moneda en importes.
- Reporte en dólares exige TC del mes, pero el reporte en soles no.
"""
import os
import tempfile
import app_kardex as app


def main():
    db_path = os.path.join(tempfile.gettempdir(), "kardex_prueba_v39_mejoras4.db")
    if os.path.exists(db_path):
        os.remove(db_path)
    db = app.KardexDB(db_path)
    db.create_user("tester", "USUARIO OPERADOR PRUEBA", "1234", "OPERADOR")
    user = "tester"

    db.add_articulo("V3901", "ARTICULO V39 VALORIZABLE", "UND", "CONSUMIBLE", 0, 0, 0, True, True)
    db.upsert_proveedor("20600209991", "SMETAL GROUP S.A.C.", "AREQUIPA", "932478711", banco="BCP", numero_cuenta="001", cci="002", usuario=user)

    # Movimiento valorizado en soles: no debe exigir TC para operar ni para reporte en soles.
    db.save_movimiento(app.today_text(), "CN", "DOC-V39-1", "84", "02", "CARLOS", "INGRESO V39", [{"codigo": "V3901", "cantidad": 3, "precio_unitario_sin_igv": 10}], user, "AQP")
    rows = db.reporte_valorizacion_mensual(app.date.today().year, app.date.today().month, "AQP")
    assert rows and rows[0]["codigo"] == "V3901"

    # Si se solicita reporte en dólares sin TC del mes, debe bloquearse desde la regla UI del reporte.
    class SimpleVar:
        def __init__(self, value): self.value = value
        def get(self): return self.value
    class FakeDialog:
        pass
    fake = FakeDialog()
    fake.db = db
    fake.anio_var = SimpleVar(str(app.date.today().year))
    fake.mes_var = SimpleVar(f"{app.date.today().month:02d}")
    try:
        app.ValorizacionMensualDialog._tc_referencia_mes(fake, strict=True)
        raise AssertionError("Se permitió reporte valorizado en dólares sin TC registrado")
    except ValueError:
        pass

    db.registrar_tipo_cambio_manual(app.today_text(), 3.70, 3.80, "MANUAL", "TC V39", user)
    assert app.ValorizacionMensualDialog._tc_referencia_mes(fake, strict=True) > 0

    req = db.create_requerimiento(app.today_text(), "AQP", "CARLOS", "02", "84", "REQ V39", [{"codigo": "V3901", "cantidad": 1}], user)
    oc = db.create_oc_from_requerimiento(req, "20600209991", "COT-V39", {"V3901": 54}, "SOLES", False, user)
    active = db.requerimiento_tiene_oc_activa(req)
    assert active and active["numero"] == oc
    try:
        db.create_oc_from_requerimiento(req, "20600209991", "COT-V39-2", {"V3901": 60}, "SOLES", False, user)
        raise AssertionError("Se permitió duplicar OC activa para el mismo requerimiento")
    except ValueError:
        pass

    # Word de OC debe contener símbolos de moneda en importes.
    class FakePreview:
        pass
    fp = FakePreview()
    fp.db = db
    fp.numero = oc
    doc = app.OCPrintPreviewDialog.build_word_document(fp)
    text = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    assert "S/" in text, "El formato Word de OC no incluye símbolo de moneda"

    db.anular_orden_compra(oc, "ANULACION DE PRUEBA", "admin")
    assert db.requerimiento_tiene_oc_activa(req) is None
    oc2 = db.create_oc_from_requerimiento(req, "20600209991", "COT-V39-3", {"V3901": 61}, "SOLES", False, user)
    assert oc2 != oc

    print("pruebas_v39_mejoras4: OK")


if __name__ == "__main__":
    main()

"""
Pruebas adicionales v19:
- mapeo correcto de razón social desde APIs RUC anidadas
- bloqueo de respuestas API que no corresponden al RUC consultado
- guardado manual de teléfono, email, banco, cuenta, CCI y contacto
- OC por requerimiento multiproveedor con forma de pago
"""
import os
import tempfile
import app_kardex as app


def main():
    db_path = os.path.join(tempfile.gettempdir(), "kardex_prueba_v19.db")
    if os.path.exists(db_path):
        os.remove(db_path)
    db = app.KardexDB(db_path)
    db.create_user('tester', 'USUARIO OPERADOR PRUEBA', '1234', 'OPERADOR')

    # API RUC con payload anidado: debe tomar la razón social real, no metadata.
    def fake_good(url, token="", timeout=10):
        return {
            "success": True,
            "data": {
                "ruc": "20601234567",
                "razonSocial": "PROVEEDOR MINERO DEL SUR S.A.C.",
                "direccion": "AV. INDUSTRIAL 123",
                "estado": "ACTIVO",
                "condicion": "HABIDO",
                "ubigeo": "040101"
            }
        }
    db._http_get_json = fake_good
    info = db.consultar_ruc_api("20601234567")
    assert info["razon_social"] == "PROVEEDOR MINERO DEL SUR S.A.C."
    assert info["estado_contribuyente"] == "ACTIVO"

    # Si la API devuelve datos del proveedor de API o de otro RUC, no debe guardarlo como razón social.
    def fake_wrong(url, token="", timeout=10):
        return {"ruc": "20608280101", "razonSocial": "REXTIE S.A.C.", "estado": "ACTIVO"}
    db._http_get_json = fake_wrong
    try:
        db.consultar_ruc_api("20601234567")
        raise AssertionError("Se aceptó un RUC diferente al consultado")
    except ValueError as e:
        assert "otro RUC" in str(e)

    # Guardado manual de datos no SUNAT/API.
    db.upsert_proveedor(
        "20601234567", "PROVEEDOR MINERO DEL SUR S.A.C.", "AV. INDUSTRIAL 123",
        telefono="999888777", email="contacto@proveedor.pe", banco="BCP",
        numero_cuenta="191-1234567-0-11", cci="00219100123456701199",
        contacto="JUAN PEREZ", moneda_preferida="SOLES", estado_contribuyente="ACTIVO",
        condicion_domicilio="HABIDO", usuario="tester"
    )
    p = db.get_proveedor_by_ruc("20601234567", active_only=False)
    assert p["telefono"] == "999888777"
    assert p["email"] == "contacto@proveedor.pe"
    assert p["numero_cuenta"] == "191-1234567-0-11"
    assert p["cci"] == "00219100123456701199"
    assert p["contacto"] == "JUAN PEREZ"

    # Multiproveedor por requerimiento con forma de pago.
    db.add_articulo("V1901", "FILTRO DE ACEITE", "UND", "CONSUMIBLE", 0, 0, 0)
    db.add_articulo("V1902", "CORREA INDUSTRIAL", "UND", "CONSUMIBLE", 0, 0, 0)
    db.upsert_proveedor("20607654321", "SUMINISTROS AQP S.A.C.", "CALLE 2", banco="BBVA", usuario="tester")
    req = db.create_requerimiento(app.today_text(), "AQP", "CASIANO", "04", "26310", "REQ V19", [
        {"codigo": "V1901", "cantidad": 3},
        {"codigo": "V1902", "cantidad": 2},
    ], "logistica")
    nums = db.create_ocs_from_requerimiento_multiproveedor(req, [
        {"ruc": "20601234567", "moneda": "SOLES", "cotizacion": "COT-A", "forma_pago": "CREDITO 15 DIAS", "precios": {"V1901": 20}},
        {"ruc": "20607654321", "moneda": "SOLES", "cotizacion": "COT-B", "forma_pago": "CREDITO 15 DIAS", "precios": {"V1902": 30}},
    ], False, "logistica")
    assert len(nums) == 2
    for n in nums:
        oc, det = db.get_orden_compra(n)
        assert oc["forma_pago"] == "CREDITO 15 DIAS"
        assert len(det) == 1
    print("Pruebas v19: OK")


if __name__ == "__main__":
    main()

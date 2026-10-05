import os


def apply_sqlite_runtime_pragmas(conn, db_path=None, busy_timeout_ms=15000):
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute(f"PRAGMA busy_timeout={int(busy_timeout_ms)}")
    # WAL mejora concurrencia en disco local. En rutas UNC no se fuerza porque
    # los mecanismos de bloqueo de WAL no son adecuados para compartir por red.
    path = str(db_path or "")
    is_unc = path.startswith("\\\\") or path.startswith("//")
    if path not in ("", ":memory:") and not is_unc:
        try:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
        except Exception:
            pass
    return True


def ensure_performance_indexes(conn):
    indexes = [
        ("idx_movimientos_fecha_almacen", "movimientos(fecha_operacion, almacen_codigo)"),
        ("idx_movimientos_tipo_almacen", "movimientos(tipo_codigo, almacen_codigo)"),
        ("idx_movimientos_documento", "movimientos(documento)"),
        ("idx_movimientos_requerimiento", "movimientos(requerimiento)"),
        ("idx_movdet_codigo", "movimiento_detalle(codigo)"),
        ("idx_movdet_movimiento", "movimiento_detalle(movimiento_id)"),
        ("idx_req_fecha_estado", "requerimientos(fecha_requerimiento, estado)"),
        ("idx_req_almacen_estado", "requerimientos(almacen_codigo, estado)"),
        ("idx_reqdet_req", "requerimiento_detalle(requerimiento_id)"),
        ("idx_reqdet_codigo", "requerimiento_detalle(codigo)"),
        ("idx_oc_fecha_estado", "orden_compra(fecha_orden, estado)"),
        ("idx_oc_almacen_estado", "orden_compra(almacen_codigo, estado)"),
        ("idx_oc_ruc", "orden_compra(ruc)"),
        ("idx_oc_req", "orden_compra(requerimiento)"),
        ("idx_oc_pago_estado", "orden_compra(estado_pago, fecha_vencimiento_pago)"),
        ("idx_ocdet_oc", "orden_compra_detalle(oc_id)"),
        ("idx_ocdet_codigo", "orden_compra_detalle(codigo)"),
        ("idx_pagos_oc_fecha", "orden_compra_pagos(oc_id, fecha_pago)"),
        ("idx_pagos_operacion", "orden_compra_pagos(banco, numero_operacion)"),
        ("idx_pagos_factura", "orden_compra_pagos(numero_factura)"),
        ("idx_recepciones_oc", "orden_compra_recepciones(oc_id, fecha_recepcion)"),
        ("idx_transito_oc", "orden_compra_transito(oc_id, fecha_despacho)"),
        ("idx_historial_oc", "orden_compra_estado_historial(oc_id, fecha_hora)"),
        ("idx_auditoria_registro", "auditoria(tabla, registro)"),
        ("idx_auditoria_fecha", "auditoria(fecha_hora)"),
        ("idx_tc_fecha", "tipo_cambio(fecha, moneda)"),
        ("idx_proveedor_razon", "proveedores(razon_social)"),
        ("idx_hist_precio_codigo_fecha", "historial_precios_articulo(codigo, fecha)"),
    ]
    for name, expr in indexes:
        conn.execute(f"CREATE INDEX IF NOT EXISTS {name} ON {expr}")
    return len(indexes)

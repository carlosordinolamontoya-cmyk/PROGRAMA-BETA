import os, tempfile
from datetime import date
from app_kardex import KardexDB, today_text, fecha_text_from_iso

def run():
    db_path = os.path.join(tempfile.gettempdir(), 'kardex_v40_test.db')
    try:
        os.remove(db_path)
    except FileNotFoundError:
        pass
    db = KardexDB(db_path)
    # básicos
    db.add_articulo('T001','TORNILLO TEST','UND','CONSUMIBLE',0,0,0,1,0)
    db.upsert_proveedor('20600209991','PROVEEDOR TEST SAC','DIR','999','','01 BCP','123','CCI','CONTACTO','SOLES',usuario='admin')
    db.registrar_tipo_cambio_manual(today_text(), 3.70, 3.75, usuario='admin')
    vale = db.save_movimiento(today_text(),'CN','DOC-TEST','1','01','SOL TEST','INGRESO TEST',[{'codigo':'T001','cantidad':10,'precio_unitario_sin_igv':5}], 'admin','AQP')
    db.save_movimiento(today_text(),'CO','DOC-SAL','1','01','SOL TEST','SALIDA TEST',[{'codigo':'T001','cantidad':2,'precio_unitario_sin_igv':5}], 'admin','AQP')
    kardex = db.reporte_kardex_valorizado_articulo(codigo='T001')
    assert kardex, 'Kardex valorizado vacío'
    cons = db.reporte_consumo_cc_ot()
    assert cons, 'Consumo por CC/OT vacío'
    oc_num = db.create_orden_compra(today_text(),'AQP','20600209991','NACIONAL','','COT-1','SOL TEST','1','01','OC TEST','SOLES',False,[{'codigo':'T001','cantidad':1,'precio':7,'centro_costo_codigo':'01','ot':'OT-1','solicitante':'SOL TEST'}],'admin')
    hist = db.historial_precios_rows('T001')
    assert hist, 'Historial de precios vacío'
    comp = db.reporte_compras_por_proveedor()
    assert comp, 'Compras por proveedor vacío'
    oc, _ = db.get_orden_compra(oc_num)
    db._set_oc_state(oc['id'], 'APROBADA', 'admin', 'Aprobación prueba v40')
    db.registrar_pago_oc(oc_num, today_text(), 'OP-TEST-1', '01 BCP', 8.26, 'PAGO TEST', 'admin', numero_factura='F001-1')
    try:
        db.registrar_pago_oc(oc_num, today_text(), 'OP-TEST-1', '01 BCP', 8.26, 'PAGO DUP', 'admin', numero_factura='F001-2')
        raise AssertionError('No bloqueó pago duplicado')
    except ValueError:
        pass
    aging = db.reporte_cuentas_pagar_aging()
    assert aging, 'Aging vacío'
    traz = db.trazabilidad_documento('OC', oc_num)
    assert traz, 'Trazabilidad OC vacía'
    db.set_action_permission(db.user_id_by_login('admin'), 'cerrar_mes', True, 'admin')
    db.cerrar_periodo(date.today().year, date.today().month, 'AQP', 'admin', 'CIERRE TEST')
    try:
        db.save_movimiento(today_text(),'CN','DOC-CERRADO','1','01','SOL','DEBE FALLAR',[{'codigo':'T001','cantidad':1,'precio_unitario_sin_igv':1}], 'operador','AQP')
        raise AssertionError('No bloqueó periodo cerrado')
    except ValueError:
        pass
    print('pruebas_v40_consolidacion: OK')

if __name__ == '__main__':
    run()

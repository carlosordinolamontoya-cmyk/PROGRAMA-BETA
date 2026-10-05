"""Pruebas v22: seguimiento financiero separado del avance operativo de la OC."""
import os, tempfile
import app_kardex as app


def base_db(name='kardex_v22_finanzas.db'):
    p=os.path.join(tempfile.gettempdir(), name)
    try: os.remove(p)
    except FileNotFoundError: pass
    db=app.KardexDB(p)
    db.create_user('almacen','USUARIO ALMACEN PRUEBA','1234','ALMACEN')
    db.add_articulo('A100','CASCO MINERO DE SEGURIDAD','UND','CONSUMIBLE',0,0,0)
    db.upsert_proveedor('10431801413','RAMIREZ GUADAMUR RICHARD RAUL','DIR','999','a@b.com','BCP','001','002','CONTACTO','SOLES',True,estado_contribuyente='ACTIVO',condicion_domicilio='HABIDO',fuente_ruc='MANUAL',usuario='admin')
    return db


def test_credito_atendido_sigue_por_pagar_y_luego_pago():
    db=base_db('kardex_v22_credito_atendido.db')
    oc=db.create_orden_compra(app.today_text(),'AQP','10431801413','NACIONAL','','COT','CASIANO','26301','04','COMPRA CREDITO','SOLES',False,[{'codigo':'A100','cantidad':2,'precio':50}], 'admin', forma_pago='CREDITO 30 DIAS')
    ocrow,det=db.get_orden_compra(oc)
    assert ocrow['estado_pago'] == 'POR PAGAR'
    assert ocrow['fecha_vencimiento_pago'] != ocrow['fecha_orden']
    db.approve_orden_compra(oc,{det[0]['id']:2},'admin')
    db.registrar_transito_oc(oc,'FLORES',app.today_text(),app.today_text(),'GUIA','EN RUTA','admin')
    vale,estado=db.recepcionar_orden_compra(oc,app.today_text(),'GR1','INGRESO TOTAL',{det[0]['id']:2},'almacen')
    assert estado == 'ATENDIDA'
    ocrow,_=db.get_orden_compra(oc)
    assert ocrow['estado'] == 'ATENDIDA'
    assert ocrow['avance'] == 100
    assert ocrow['estado_pago'] == 'POR PAGAR', 'La OC atendida al 100% a crédito debe seguir pendiente financieramente'
    estado_pago=db.registrar_pago_oc(oc,app.today_text(),'OP-CRED','BCP',118,'PAGO POSTERIOR A RECEPCION','finanzas')
    assert estado_pago == 'PAGADA'
    ocrow,_=db.get_orden_compra(oc)
    assert ocrow['estado'] == 'ATENDIDA'
    assert ocrow['avance'] == 100
    assert ocrow['estado_pago'] == 'PAGADA'


def test_credito_vencido_aparece_en_seguimiento():
    db=base_db('kardex_v22_credito_vencido.db')
    oc=db.create_orden_compra('01/01/2025','AQP','10431801413','NACIONAL','','COT','CASIANO','26301','04','COMPRA CREDITO VENCIDO','SOLES',False,[{'codigo':'A100','cantidad':1,'precio':10}], 'admin', forma_pago='CREDITO 15 DIAS')
    ocrow,det=db.get_orden_compra(oc)
    db.approve_orden_compra(oc,{det[0]['id']:1},'admin')
    rows=db.list_oc_finanzas(estado_pago='VENCIDA', solo_saldo=True)
    nums=[r['numero'] for r in rows]
    assert oc in nums
    ocrow,_=db.get_orden_compra(oc)
    assert ocrow['estado_pago'] == 'VENCIDA'


if __name__=='__main__':
    test_credito_atendido_sigue_por_pagar_y_luego_pago()
    test_credito_vencido_aparece_en_seguimiento()
    print('Pruebas v22 finanzas crédito: OK')

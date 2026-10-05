"""Pruebas integrales v21: bugs de regresión, reversos y validaciones finales."""
import os, tempfile
import app_kardex as app


def base_db(name='kardex_v21_integral.db'):
    p=os.path.join(tempfile.gettempdir(), name)
    try: os.remove(p)
    except FileNotFoundError: pass
    db=app.KardexDB(p)
    db.create_user('almacen','USUARIO ALMACEN PRUEBA','1234','ALMACEN')
    db.add_articulo('A100','CASCO MINERO DE SEGURIDAD','UND','CONSUMIBLE',0,0,0)
    db.add_articulo('A200','GUANTE CUERO REFORZADO','PAR','CONSUMIBLE',0,0,0)
    db.upsert_proveedor('10431801413','RAMIREZ GUADAMUR RICHARD RAUL','DIR','999','a@b.com','BCP','001','002','CONTACTO','SOLES',True,estado_contribuyente='ACTIVO',condicion_domicilio='HABIDO',fuente_ruc='MANUAL',usuario='admin')
    db.upsert_proveedor('20600000001','PROVEEDOR DOS SAC','DIR2','999','a2@b.com','BBVA','003','004','CONTACTO2','SOLES',True,estado_contribuyente='ACTIVO',condicion_domicilio='HABIDO',fuente_ruc='MANUAL',usuario='admin')
    return db


def test_pago_no_regresa_estado():
    db=base_db('kardex_v21_pago.db')
    oc=db.create_orden_compra(app.today_text(),'AQP','10431801413','NACIONAL','','COT','CASIANO','26301','04','COMPRA','SOLES',False,[{'codigo':'A100','cantidad':2,'precio':50}], 'admin')
    ocrow,det=db.get_orden_compra(oc)
    db.approve_orden_compra(oc,{det[0]['id']:2},'admin')
    db.registrar_transito_oc(oc,'FLORES',app.today_text(),app.today_text(),'GUIA','EN RUTA','admin')
    estado=db.registrar_pago_oc(oc,app.today_text(),'OP-TARDE','BCP',118,'PAGO DESPUES DE TRANSITO','finanzas')
    ocrow,_=db.get_orden_compra(oc)
    assert estado == 'PAGADA'
    assert ocrow['estado'] == 'EN TRANSITO', 'El pago no debe regresar EN TRANSITO a PAGADA'
    assert ocrow['estado_pago'] == 'PAGADA', 'El estado financiero debe quedar pagado sin tocar el avance operativo'


def test_precio_cero_bloqueado():
    db=base_db('kardex_v21_precio0.db')
    try:
        db.create_orden_compra(app.today_text(),'AQP','10431801413','NACIONAL','','COT','CASIANO','26301','04','COMPRA','SOLES',False,[{'codigo':'A100','cantidad':1,'precio':0}], 'admin')
        raise AssertionError('Permitió OC con precio cero')
    except ValueError as e:
        assert 'precio' in str(e).lower()


def test_reverso_oc_ajusta_cantidad_ingresada():
    db=base_db('kardex_v21_reverso_oc.db')
    oc=db.create_orden_compra(app.today_text(),'AQP','10431801413','NACIONAL','','COT','CASIANO','26301','04','COMPRA','SOLES',False,[{'codigo':'A100','cantidad':4,'precio':10}], 'admin')
    ocrow,det=db.get_orden_compra(oc)
    db.approve_orden_compra(oc,{det[0]['id']:4},'admin')
    vale,estado=db.recepcionar_orden_compra(oc,app.today_text(),'GR1','INGRESO',{det[0]['id']:4},'almacen')
    assert estado=='ATENDIDA'
    assert db.get_stock('A100','AQP')==4
    db.reverse_movimiento(vale,'ERROR DE RECEPCION','admin')
    ocrow,det2=db.get_orden_compra(oc)
    assert float(det2[0]['cantidad_ingresada']) == 0
    assert ocrow['estado'] == 'APROBADA'
    assert db.get_stock('A100','AQP') == 0


def test_reverso_requerimiento_ajusta_cantidad_atendida():
    db=base_db('kardex_v21_reverso_req.db')
    db.save_movimiento(app.today_text(),'CN','STOCK','26301','04','CASIANO','CARGA',[{'codigo':'A100','cantidad':5}],'admin','AQP')
    req=db.create_requerimiento(app.today_text(),'AQP','CASIANO','04','26302','REQ',[{'codigo':'A100','cantidad':3}], 'admin')
    vale,estado=db.attend_requerimiento(req,app.today_text(),{'A100':3},'ATENCION','admin')
    assert estado=='ATENDIDO'
    db.reverse_movimiento(vale,'ERROR EN ATENCION','admin')
    reqrow,det=db.find_requerimiento(req)
    assert reqrow['estado'] == 'PENDIENTE'
    assert float(det[0]['cantidad_atendida']) == 0
    assert db.get_stock('A100','AQP') == 5


def test_multiproveedor_atomico_en_error():
    db=base_db('kardex_v21_atomico.db')
    req=db.create_requerimiento(app.today_text(),'AQP','CASIANO','04','26302','REQ',[{'codigo':'A100','cantidad':3},{'codigo':'A200','cantidad':2}], 'admin')
    try:
        db.create_ocs_from_requerimiento_multiproveedor(req,[
            {'ruc':'10431801413','moneda':'SOLES','cotizacion':'COT1','precios':{'A100':10}},
            {'ruc':'99999999999','moneda':'SOLES','cotizacion':'COT2','precios':{'A200':8}},
        ],False,'admin')
        raise AssertionError('Permitió lote multiproveedor parcialmente inválido')
    except ValueError:
        pass
    rows=db.list_ordenes_compra(term=req)
    assert len(rows)==0, 'No debe quedar ninguna OC del lote si falla un proveedor'


if __name__=='__main__':
    test_pago_no_regresa_estado()
    test_precio_cero_bloqueado()
    test_reverso_oc_ajusta_cantidad_ingresada()
    test_reverso_requerimiento_ajusta_cantidad_atendida()
    test_multiproveedor_atomico_en_error()
    print('Pruebas integrales v21/v22: OK')

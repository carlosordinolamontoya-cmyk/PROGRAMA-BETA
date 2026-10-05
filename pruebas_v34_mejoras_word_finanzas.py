import os, tempfile, tkinter as tk
from app_kardex import KardexDB, OCPrintPreviewDialog, today_text

fd, path = tempfile.mkstemp(suffix='.db'); os.close(fd); os.remove(path)
db = KardexDB(path)
db.add_articulo('P100','PERNO ACERO 1/2','UND','CONSUMIBLE', valoriza=1, permite_decimales=0)
db.add_articulo('P200','SERVICIO FLETE LOCAL','UND','SERVICIO', valoriza=0, permite_decimales=1)
db.upsert_proveedor('20600209991','SMETAL GROUP S.A.C.','DIR PROV','999888777','logistica@smetalgroup.com.pe','01 BCP','191-123','002-191-123','CONTACTO','SOLES', usuario='admin')
# OC manual y pago con factura
oc = db.create_orden_compra(today_text(),'AQP','20600209991','NACIONAL','', 'COT-34','CARLOS','84','02','PRUEBA WORD','SOLES',False,[{'codigo':'P100','cantidad':3,'precio':12,'centro_costo_codigo':'02','ot':'OT-84','solicitante':'CARLOS'}],'admin')
db.approve_orden_compra(oc,{1:3},'admin')
db.registrar_pago_oc(oc,today_text(),'OP-34','01 BCP',42.48,'PAGO PRUEBA','finanzas',numero_factura='F001-000123')
pagos = db.pagos_oc(1)
assert pagos and pagos[0]['numero_factura'] == 'F001-000123'
# Word generado con plantilla
root = tk.Tk(); root.withdraw()
dlg = OCPrintPreviewDialog(root, db, oc); dlg.withdraw()
out = os.path.join(tempfile.gettempdir(),'oc_v34_test.docx')
dlg.save_word_file(out)
assert os.path.exists(out) and os.path.getsize(out) > 10000
root.destroy()
print('pruebas_v34_mejoras_word_finanzas: OK')

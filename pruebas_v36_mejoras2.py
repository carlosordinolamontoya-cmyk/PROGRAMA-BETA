import os, tempfile, tkinter as tk
from app_kardex import KardexDB, PrintPreviewDialog, OCPrintPreviewDialog, today_text, RequerimientoMultiSelectDialog

fd, db_path = tempfile.mkstemp(suffix='.db'); os.close(fd); os.remove(db_path)
db = KardexDB(db_path)
db.add_articulo('A100','CASCO DE SEGURIDAD','UND','CONSUMIBLE', valoriza=1, permite_decimales=0)
db.add_articulo('A200','GUANTE NITRILO','PAR','CONSUMIBLE', valoriza=1, permite_decimales=0)
db.upsert_proveedor('20600209991','SMETAL GROUP S.A.C.','DIR','999','mail@x.pe','01 BCP','111','222','CONTACTO','SOLES', usuario='admin')
vale = db.save_movimiento(today_text(),'CN','FAC-1','OT-77','02','CARLOS','INGRESO PRUEBA',[{'codigo':'A100','cantidad':5,'precio_unitario':10}], 'admin')
mov, det = db.get_movimiento_by_vale(vale)
root = tk.Tk(); root.withdraw()
# Vale Word
dlg = PrintPreviewDialog(root, mov, det); dlg.withdraw()
out = os.path.join(tempfile.gettempdir(),'vale_v36_test.docx')
dlg.save_word_file(out)
assert os.path.exists(out) and os.path.getsize(out) > 10000
# OC Word and payment factura
oc = db.create_orden_compra(today_text(),'AQP','20600209991','NACIONAL','', 'COT-1','CARLOS','77','02','OC PRUEBA','SOLES',False,[{'codigo':'A100','cantidad':2,'precio':12,'centro_costo_codigo':'02','ot':'OT-77','solicitante':'CARLOS'}],'admin')
prev = OCPrintPreviewDialog(root, db, oc); prev.withdraw()
out2 = os.path.join(tempfile.gettempdir(),'oc_v36_test.docx')
prev.save_word_file(out2)
assert os.path.exists(out2) and os.path.getsize(out2) > 10000
# Requerimientos multi: same OT check box UX
r1=db.create_requerimiento(today_text(),'AQP','CARLOS','02','OT-77','NACIONAL',[{'codigo':'A100','cantidad':1}], 'admin')
r2=db.create_requerimiento(today_text(),'AQP','CARLOS','02','OT-77','NACIONAL',[{'codigo':'A200','cantidad':1}], 'admin')
selected=[]
sel=RequerimientoMultiSelectDialog(root, db, lambda reqs: selected.extend(reqs)); sel.withdraw()
# Simular check de las dos primeras filas con misma OT
children=sel.tree.get_children()
assert len(children) >= 2
for iid in children[:2]:
    sel.tree.focus(iid); sel.toggle_current()
sel.choose()
assert len(selected) == 2
root.destroy()
print('pruebas_v36_mejoras2: OK')

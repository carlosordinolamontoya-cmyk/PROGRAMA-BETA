import os
import tempfile
import tkinter as tk
from app_kardex import KardexDB, today_text, StockDialog, ValeListDialog, RequerimientoListDialog, TipoCambioDialog, EntryDialog, TransferDialog

class FakeApp:
    def __init__(self):
        self._tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self._tmp.close()
        self.root = tk.Tk()
        self.root.withdraw()
        self.db = KardexDB(self._tmp.name)
        self.usuario_actual = 'admin'
        self.current_user = {'usuario':'admin','rol':'ADMIN'}
        self.almacen_actual = 'AQP'
    def require_finanzas(self, parent=None): return True
    def require_logistica(self, parent=None): return True
    def require_almacen(self, parent=None): return True
    def require_write(self, parent=None): return True
    def require_admin(self, parent=None): return True
    def cleanup(self):
        try: self.root.destroy()
        except Exception: pass
        try: os.unlink(self._tmp.name)
        except Exception: pass

def close_children(root, keep=None):
    keep = keep or set()
    for w in list(root.winfo_children()):
        try:
            if isinstance(w, tk.Toplevel) and w not in keep:
                w.destroy()
        except Exception:
            pass
    root.update()

app = FakeApp()
try:
    try:
        app.db.add_articulo('DIALOGO1','ARTICULO DIALOGO','UND','CONSUMIBLE',0,0,0,1,1)
    except Exception:
        pass
    app.db.save_movimiento(today_text(),'CN','DOC-DIALOGO','1','01','SOL DIALOGO','OBS',[
        {'codigo':'DIALOGO1','descripcion':'ARTICULO DIALOGO','um':'UND','cantidad':5,'precio_unitario_sin_igv':10}
    ],'admin','AQP')
    app.db.create_requerimiento(today_text(),'AQP','SOL DIALOGO','01','1','REQ DIALOGO',[
        {'codigo':'DIALOGO1','descripcion':'ARTICULO DIALOGO','um':'UND','cantidad':1,'observacion':''}
    ],'admin')

    st = StockDialog(app)
    app.root.update()
    assert len(st.tree.get_children()) > 0, 'StockDialog no cargó artículos'
    st.destroy(); app.root.update()

    vl = ValeListDialog(app.root, app.db, lambda vale: vale)
    app.root.update()
    assert len(vl.tree.get_children()) > 0, 'ValeListDialog no cargó vales'
    vl.destroy(); app.root.update()

    rq = RequerimientoListDialog(app.root, app.db, lambda req: req)
    app.root.update()
    assert len(rq.tree.get_children()) > 0, 'RequerimientoListDialog no cargó requerimientos'
    assert rq.tree.selection(), 'RequerimientoListDialog no preseleccionó fila'
    rq.destroy(); app.root.update()

    tc = TipoCambioDialog(app)
    app.root.update()
    assert len(tc.tree.get_children()) >= 28, 'TipoCambioDialog no mostró calendario mensual'
    tc.destroy(); app.root.update()

    ed = EntryDialog(app, mode='view')
    app.root.update()
    assert ed.open_vale_list() == 'break', 'open_vale_list no detiene propagación'
    close_children(app.root, {ed})
    assert ed.open_cc_help() == 'break', 'open_cc_help no detiene propagación'
    close_children(app.root, {ed})
    assert ed.open_tipo_help() == 'break', 'open_tipo_help no detiene propagación'
    close_children(app.root, {ed})
    ed.destroy(); app.root.update()

    tr = TransferDialog(app)
    app.root.update()
    assert tr.open_stock_help() == 'break', 'TransferDialog stock help no detiene propagación'
    close_children(app.root, {tr})
    assert tr.open_cc_help() == 'break', 'TransferDialog CC help no detiene propagación'
    close_children(app.root, {tr})
    tr.destroy(); app.root.update()

    print('Pruebas v26 diálogos/tipo cambio: OK')
finally:
    app.cleanup()

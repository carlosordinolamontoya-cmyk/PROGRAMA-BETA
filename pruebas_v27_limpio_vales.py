import os
import tempfile
import tkinter as tk
from app_kardex import KardexDB, today_text, StockDialog, ValeListDialog, EntryDialog

class FakeApp:
    def __init__(self):
        self.tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.tmp.close()
        self.root = tk.Tk()
        self.root.withdraw()
        self.db = KardexDB(self.tmp.name)
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
        try: os.unlink(self.tmp.name)
        except Exception: pass

def main():
    app = FakeApp()
    selected = []
    try:
        try:
            app.db.add_articulo('TESTVALE','ARTICULO TEST VALE','UND','CONSUMIBLE',0,0,0,1,1)
        except Exception:
            pass
        vale = app.db.save_movimiento(
            today_text(), 'CN', 'DOC-VALE-TEST', '100', '01', 'SOL TEST', 'PRUEBA VALE',
            [{'codigo':'TESTVALE','descripcion':'ARTICULO TEST VALE','um':'UND','cantidad':3,'precio_unitario_sin_igv':10}],
            'admin', 'AQP'
        )
        assert app.db.list_vales('DOC-VALE-TEST'), 'list_vales no devuelve vale por documento'
        assert app.db.list_vales(vale), 'list_vales no devuelve vale por número'

        vl = ValeListDialog(app.root, app.db, lambda v: selected.append(v))
        app.root.update()
        assert len(vl.tree.get_children()) == 1, 'ValeListDialog no cargó el vale creado'
        vl.choose()
        app.root.update()
        assert selected and selected[0] == vale, 'ValeListDialog no seleccionó correctamente'

        ed = EntryDialog(app, mode='view')
        app.root.update()
        assert ed.open_vale_list() == 'break', 'open_vale_list no devuelve break'
        assert ed.open_vale_list() == 'break', 'open_vale_list duplicado no devuelve break'
        # Debe haber solo un diálogo de vales asociado.
        dialogs = [w for w in app.root.winfo_children() if isinstance(w, tk.Toplevel) and getattr(w, 'title', lambda: '')() == 'Consulta de Vales Generados']
        assert len(dialogs) <= 1, 'open_vale_list permite duplicar ventanas'
        ed.destroy()
        app.root.update()
        print('Pruebas v27 limpio/vales: OK')
    finally:
        app.cleanup()

if __name__ == '__main__':
    main()

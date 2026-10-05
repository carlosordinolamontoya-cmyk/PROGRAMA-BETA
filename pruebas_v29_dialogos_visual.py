import os
import tempfile
import tkinter as tk
from tkinter import messagebox
import app_kardex as app

# Silenciar messageboxes para pruebas automáticas.
messagebox.showinfo = lambda *a, **k: None
messagebox.showwarning = lambda *a, **k: None
messagebox.showerror = lambda *a, **k: None

# Base TEMPORAL para la inspección visual controlada. Nunca usar app.DB_FILE.
TEST_DB = os.path.join(tempfile.gettempdir(), "kardex_v29_dialogos_visual_test.db")
if os.path.exists(TEST_DB):
    os.remove(TEST_DB)

class FakeApp:
    def __init__(self, root):
        self.root = root
        self.db = app.KardexDB(TEST_DB)
        self.usuario_actual = "admin"
        self.current_user = {"id": 1, "usuario": "admin", "rol": "ADMIN"}
        self.almacen_actual = "AQP"
    def require_admin(self, parent=None): return True
    def require_logistica(self, parent=None): return True
    def require_finanzas(self, parent=None): return True
    def require_write(self, parent=None): return True
    def require_almacen(self, parent=None, almacen_codigo=None): return True

root = tk.Tk()
root.withdraw()
fake = FakeApp(root)

# Datos mínimos para cuadros de ayuda que dependen de mantenimiento.
fake.db.add_solicitante("DEMO SOLICITANTE")
fake.db.add_articulo("ZZ-TEST", "ARTICULO PRUEBA UI", "UND", "CONSUMIBLE", 0, 0, 0, 1, 1)
fake.db.create_user("uiprueba", "Usuario Prueba", "1234", "OPERADOR")

results = []

def count_tree(dlg):
    root.update_idletasks()
    return len(dlg.tree.get_children())

def assert_tree_has_rows(label, dlg):
    n = count_tree(dlg)
    assert n > 0, f"{label}: árbol vacío"
    results.append((label, n))
    dlg.destroy()
    root.update_idletasks()

# Consulta de stock debe listar artículos.
dlg = app.StockDialog(fake, parent=root, on_select=lambda c: None)
assert_tree_has_rows("Consulta de stock", dlg)

# Consulta/modificación de artículos debe listar artículos.
dlg = app.ArticleQueryDialog(fake)
assert_tree_has_rows("Consulta/modificación artículos", dlg)

# Usuarios debe listar admin y usuario de prueba.
dlg = app.UsersDialog(fake)
n = count_tree(dlg)
assert n >= 2, "Usuarios: no muestra usuarios creados"
results.append(("Usuarios", n))
dlg.destroy(); root.update_idletasks()

# Solicitantes debe listar solicitantes.
dlg = app.SolicitanteSelectDialog(root, fake.db, on_select=lambda nombre: None)
assert_tree_has_rows("Buscar solicitante", dlg)

# Centros de costo debe listar centros.
dlg = app.CentroCostoSelectDialog(root, on_select=lambda codigo: None)
assert_tree_has_rows("Buscar centro de costo", dlg)

# Artículos helper debe listar artículos y stocks.
dlg = app.ArticuloSelectDialog(root, fake.db, on_select=lambda codigo: None)
assert_tree_has_rows("Buscar artículo", dlg)

# Diálogo clásico de centros de costo.
dlg = app.CenterCostDialog(root, on_select=lambda codigo: None)
assert_tree_has_rows("Consulta centros de costo", dlg)

root.destroy()
try:
    os.remove(TEST_DB)
except FileNotFoundError:
    pass

print("OK pruebas_v29_dialogos_visual")
for label, n in results:
    print(f"- {label}: {n} filas")

import os
import tempfile
import tkinter as tk
from tkinter import ttk

import app_kardex as app


def find_treeviews(widget):
    out = []
    for child in widget.winfo_children():
        if isinstance(child, ttk.Treeview):
            out.append(child)
        out.extend(find_treeviews(child))
    return out


def find_texts(widget):
    out = []
    for child in widget.winfo_children():
        if isinstance(child, tk.Text):
            out.append(child)
        out.extend(find_texts(child))
    return out


def main():
    db_path = os.path.join(tempfile.gettempdir(), "kardex_prueba_v32.db")
    if os.path.exists(db_path):
        os.remove(db_path)
    db = app.KardexDB(db_path)
    db.create_user('tester', 'USUARIO OPERADOR PRUEBA', '1234', 'OPERADOR')

    # Artículo sin decimales: debe bloquear 1.5 en requerimientos.
    db.add_articulo("UND001", "ARTICULO ENTERO", "UND", "CONSUMIBLE", 0, 0, 0, 1, 0)
    try:
        db.create_requerimiento(app.today_text(), "AQP", "CASIANO", "04", "100", "PRUEBA DECIMAL", [{"codigo": "UND001", "cantidad": 1.5}], "tester")
        raise AssertionError("Permitió cantidad decimal para artículo configurado sin decimales")
    except ValueError as e:
        assert "no permite cantidades decimales" in str(e)

    # Cantidad entera debe pasar.
    req = db.create_requerimiento(app.today_text(), "AQP", "CASIANO", "04", "101", "PRUEBA ENTERA", [{"codigo": "UND001", "cantidad": 2}], "tester")
    assert req.startswith("REQ-")

    db.upsert_proveedor("20123456789", "PROVEEDOR DEMO SAC", "AV TEST", "999999999", banco="01 BCP", numero_cuenta="123", usuario="tester")
    oc = db.create_orden_compra(app.today_text(), "AQP", "20123456789", "NACIONAL", "", "COT32", "CASIANO", "101", "04", "OC DETALLE", "SOLES", False, [{"codigo": "UND001", "cantidad": 2, "precio": 10}], "tester")
    oc_row, det = db.get_orden_compra(oc)
    assert len(det) == 1

    root = tk.Tk()
    root.withdraw()
    dlg = app.OCDetailDialog(root, db, oc)
    dlg.update_idletasks()
    trees = find_treeviews(dlg)
    assert trees, "No se encontró Treeview de items en detalle OC"
    # El Treeview de items debe tener al menos la fila del artículo.
    assert any(len(t.get_children()) >= 1 for t in trees), "Detalle OC no muestra filas de items"
    texts = find_texts(dlg)
    assert texts, "No se encontró caja de historial en detalle OC"
    assert int(str(texts[0].cget("height"))) >= 8, "Historial de estados quedó muy bajo"
    dlg.destroy()
    root.destroy()
    db.conn.close()
    os.remove(db_path)
    print("pruebas_v32_decimal_ocdetalle: OK")


if __name__ == "__main__":
    main()

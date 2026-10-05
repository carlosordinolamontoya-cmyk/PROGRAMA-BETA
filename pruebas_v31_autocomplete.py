import os
import tkinter as tk

import app_kardex as app


def main():
    db_path = os.path.join(os.path.dirname(__file__), 'test_v31_autocomplete.db')
    if os.path.exists(db_path):
        os.remove(db_path)
    db = app.KardexDB(db_path)
    db.create_user('tester', 'USUARIO OPERADOR PRUEBA', '1234', 'OPERADOR')
    db.add_solicitante('JUAN PEREZ')
    db.add_articulo('ART-001', 'CASCO DE SEGURIDAD', 'UND', 'CONSUMIBLE', 0, 0, 0, 1, 0)
    db.upsert_proveedor('20123456789', 'PROVEEDOR DEMO SAC', 'AV TEST 123', '999999999', 'demo@proveedor.pe', '01 BCP', '123456789', '001123456789', 'CONTACTO DEMO', 'SOLES', True, usuario='tester')

    root = tk.Tk()
    root.withdraw()
    var_sol = tk.StringVar()
    e_sol = tk.Entry(root, textvariable=var_sol)
    h1 = app.attach_solicitante_autocomplete(e_sol, db, var_sol)
    var_sol.set('JU')
    rows = h1.get_rows('JU')
    assert rows and rows[0][1] == 'JUAN PEREZ'

    var_art = tk.StringVar()
    e_art = tk.Entry(root, textvariable=var_art)
    h2 = app.attach_articulo_autocomplete(e_art, db, var_art)
    rows = h2.get_rows('CASCO')
    assert rows and rows[0][1] == 'ART-001'

    var_prov = tk.StringVar()
    e_prov = tk.Entry(root, textvariable=var_prov)
    h3 = app.attach_proveedor_autocomplete(e_prov, db, var_prov)
    rows = h3.get_rows('PROVEEDOR')
    assert rows and rows[0][1] == '20123456789'

    root.destroy()
    db.conn.close()
    os.remove(db_path)
    print('pruebas_v31_autocomplete: OK')


if __name__ == '__main__':
    main()

"""Prueba visual/estructural V43 de menus por permisos. Ejecutar con pantalla o xvfb-run."""
import os
import tempfile
import tkinter as tk
import app_kardex as app


def menu_by_label(menubar, label):
    end = menubar.index('end')
    for i in range((end or -1) + 1):
        if menubar.type(i) == 'cascade' and menubar.entrycget(i, 'label') == label:
            return menubar.nametowidget(menubar.entrycget(i, 'menu'))
    raise AssertionError(f'Menu no encontrado: {label}')


def entry_state(menu, label):
    end = menu.index('end')
    for i in range((end or -1) + 1):
        if menu.type(i) != 'separator' and menu.entrycget(i, 'label') == label:
            return str(menu.entrycget(i, 'state'))
    raise AssertionError(f'Entrada no encontrada: {label}')


def build_fake_app(root, db, username):
    obj = app.KardexApp.__new__(app.KardexApp)
    obj.root = root
    obj.db = db
    obj.current_user = db.verify_user(username, '1234' if username != 'admin' else 'admin123')
    obj.usuario_actual = username
    obj.almacen_actual = 'AQP'
    obj.build_ui()
    root.update_idletasks()
    return obj


def run():
    with tempfile.TemporaryDirectory(prefix='kardex_v43_ui_') as td:
        db = app.KardexDB(os.path.join(td, 'ui.db'))
        db.create_user('consulta43', 'CONSULTA CON PERMISOS', '1234', 'CONSULTA')
        uid = db.user_id_by_login('consulta43')
        db.set_action_permission(uid, 'aprobar_oc', True, 'admin')
        db.set_action_permission(uid, 'registrar_pago', True, 'admin')
        db.set_action_permission(uid, 'crear_oc', False, 'admin')
        db.set_action_permission(uid, 'modificar_proveedor', False, 'admin')
        db.set_action_permission(uid, 'reversar_vale', False, 'admin')

        root = tk.Tk(); root.withdraw()
        fake = build_fake_app(root, db, 'consulta43')
        menubar = root.nametowidget(root.cget('menu'))
        log = menu_by_label(menubar, 'Logística')
        fin = menu_by_label(menubar, 'Finanzas')
        seg = menu_by_label(menubar, 'Seguridad / Respaldo')
        assert entry_state(log, 'Aprobación Orden de Compra') == 'normal'
        assert entry_state(log, 'Registro Orden de Compra') == 'disabled'
        assert entry_state(log, 'Proveedores / Nuevo proveedor') == 'disabled'
        assert entry_state(fin, 'Registrar Pago de Orden de Compra') == 'normal'
        assert entry_state(seg, 'Crear respaldo ahora') == 'disabled'
        root.destroy()

        root = tk.Tk(); root.withdraw()
        fake = build_fake_app(root, db, 'admin')
        menubar = root.nametowidget(root.cget('menu'))
        seg = menu_by_label(menubar, 'Seguridad / Respaldo')
        assert entry_state(seg, 'Crear respaldo ahora') == 'normal'
        assert entry_state(seg, 'Restaurar respaldo') == 'normal'
        assert entry_state(seg, 'Verificar integridad de la base') == 'normal'
        root.destroy()
        db.conn.close()
    print('PRUEBAS V43 UI/PERMISOS: OK')


if __name__ == '__main__':
    run()

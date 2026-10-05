"""Prueba visual/UX básica: abre diálogos principales bajo un display virtual y valida que construyan sin errores."""
import os, tempfile, tkinter as tk
import app_kardex as app

class FakeApp:
    def __init__(self, root, db):
        self.root=root; self.db=db; self.usuario_actual='admin'; self.almacen_actual='AQP'
        self.current_user={'id':1,'usuario':'admin','nombre':'ADMIN','rol':'ADMIN'}
    def require_write(self,parent=None): return True
    def require_almacen(self,parent=None,almacen_codigo=None): return True
    def require_logistica(self,parent=None): return True
    def require_finanzas(self,parent=None): return True
    def require_admin(self,parent=None): return True
    def can_write(self): return True
    def can_almacen(self): return True
    def can_logistica(self): return True
    def can_finanzas(self): return True
    def is_admin(self): return True
    def user_role(self): return 'ADMIN'

def seed(db):
    db.add_articulo('A100','CASCO MINERO DE SEGURIDAD','UND','CONSUMIBLE',0,0,0)
    db.add_articulo('A200','GUANTE CUERO REFORZADO','PAR','CONSUMIBLE',0,0,0)
    db.upsert_proveedor('10431801413','RAMIREZ GUADAMUR RICHARD RAUL','SEBASTIAN BARRANCA','975683377','prov@test.com','BCP','001-123','002-123','RICHARD','SOLES',True,estado_contribuyente='ACTIVO',condicion_domicilio='HABIDO',fuente_ruc='MANUAL',usuario='admin')
    vale=db.save_movimiento(app.today_text(),'CN','DOC-1','26301','04','CASIANO','INGRESO DE PRUEBA',[{'codigo':'A100','cantidad':20},{'codigo':'A200','cantidad':10}],'admin','AQP')
    req=db.create_requerimiento(app.today_text(),'AQP','CASIANO','04','26302','REQ PRUEBA',[{'codigo':'A100','cantidad':5},{'codigo':'A200','cantidad':3}],'admin')
    oc=db.create_orden_compra(app.today_text(),'AQP','10431801413','NACIONAL',req,'COT-1','CASIANO','26302','04','COMPRA PRUEBA','SOLES',False,[{'codigo':'A100','cantidad':5,'precio':10},{'codigo':'A200','cantidad':3,'precio':8}],'admin')
    ocrow,det=db.get_orden_compra(oc); db.approve_orden_compra(oc,{d['id']:d['cantidad_solicitada'] for d in det},'admin')
    return req, oc, vale

def main():
    db_path=os.path.join(tempfile.gettempdir(),'kardex_ui_v21.db')
    try: os.remove(db_path)
    except FileNotFoundError: pass
    db=app.KardexDB(db_path)
    req,oc,vale=seed(db)
    root=tk.Tk(); root.geometry('1280x800+0+0'); root.withdraw()
    fake=FakeApp(root,db)
    mov,mdet=db.get_movimiento_by_vale(vale)
    dialogs=[
        ('StockDialog', lambda: app.StockDialog(fake)),
        ('CenterCostDialog', lambda: app.CenterCostDialog(root, lambda c,d: None)),
        ('ValeListDialog', lambda: app.ValeListDialog(root, db, lambda v: None)),
        ('EntryDialog', lambda: app.EntryDialog(fake)),
        ('TransferDialog', lambda: app.TransferDialog(fake)),
        ('PrintPreviewDialog', lambda: app.PrintPreviewDialog(root, mov, mdet)),
        ('RegisterArticleDialog', lambda: app.RegisterArticleDialog(fake)),
        ('BulkArticleDialog', lambda: app.BulkArticleDialog(fake)),
        ('ArticleQueryDialog', lambda: app.ArticleQueryDialog(fake)),
        ('RequerimientoListDialog', lambda: app.RequerimientoListDialog(root, db, lambda r: None)),
        ('RequerimientoDetailDialog', lambda: app.RequerimientoDetailDialog(root, fake, req)),
        ('ConsultaRequerimientosDialog', lambda: app.ConsultaRequerimientosDialog(fake)),
        ('RequerimientoDialog', lambda: app.RequerimientoDialog(fake)),
        ('AtencionRequerimientoDialog', lambda: app.AtencionRequerimientoDialog(fake)),
        ('SolicitudesDialog', lambda: app.SolicitudesDialog(fake)),
        ('SolicitanteSelectDialog', lambda: app.SolicitanteSelectDialog(root, db, lambda s: None)),
        ('CentroCostoSelectDialog', lambda: app.CentroCostoSelectDialog(root, lambda c: None)),
        ('ArticuloSelectDialog', lambda: app.ArticuloSelectDialog(root, db, lambda c: None)),
        ('OCSelectDialog', lambda: app.OCSelectDialog(root, db, lambda n: None)),
        ('ProviderListDialog', lambda: app.ProviderListDialog(fake)),
        ('OCDetailDialog', lambda: app.OCDetailDialog(root, db, oc)),
        ('OCManualDialog', lambda: app.OCManualDialog(fake)),
        ('OCListDialog', lambda: app.OCListDialog(fake)),
        ('OCAprobacionDialog', lambda: app.OCAprobacionDialog(fake)),
        ('OCRecepcionDialog', lambda: app.OCRecepcionDialog(fake)),
        ('OCPagoDialog', lambda: app.OCPagoDialog(fake)),
        ('OCFinanzasSeguimientoDialog', lambda: app.OCFinanzasSeguimientoDialog(fake)),
        ('OCTransitoDialog', lambda: app.OCTransitoDialog(fake)),
        ('OCLiquidacionDialog', lambda: app.OCLiquidacionDialog(fake)),
        ('OCRequerimientoDialog', lambda: app.OCRequerimientoDialog(fake)),
        ('OCRequerimientoMatrizDialog', lambda: app.OCRequerimientoMatrizDialog(fake, db, db.find_requerimiento(req)[0], [{'codigo':'A100','descripcion':'CASCO MINERO DE SEGURIDAD','um':'UND','cantidad':5},{'codigo':'A200','descripcion':'GUANTE CUERO REFORZADO','um':'PAR','cantidad':3}])),
        ('OCPrintPreviewDialog', lambda: app.OCPrintPreviewDialog(root, db, oc)),
        ('TipoCambioDialog', lambda: app.TipoCambioDialog(fake)),
        ('APIConfigDialog', lambda: app.APIConfigDialog(fake)),
        ('UsersDialog', lambda: app.UsersDialog(fake)),
    ]
    results=[]
    for name, maker in dialogs:
        try:
            d=maker(); root.update_idletasks(); root.update()
            w=d.winfo_width(); h=d.winfo_height(); reqw=d.winfo_reqwidth(); reqh=d.winfo_reqheight()
            results.append((name,'OK',w,h,reqw,reqh))
            d.destroy(); root.update()
        except Exception as e:
            results.append((name,'ERROR',str(e),'','',''))
    root.destroy()
    bad=[r for r in results if r[1]!='OK']
    for r in results:
        print(r)
    if bad:
        raise SystemExit('UI visual con errores: '+repr(bad))
    print('Prueba UI visual v21/v22: OK')

if __name__=='__main__': main()

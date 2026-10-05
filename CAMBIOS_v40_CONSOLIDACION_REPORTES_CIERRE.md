# CAMBIOS v40 - Consolidación ERP/MRP sin dashboard

Versión enfocada en estabilización, control y reportes gerenciales, sin implementar dashboard principal.

## Implementado

1. Prueba operativa controlada
   - Menú: Control ERP > Prueba operativa controlada.
   - Lista de pasos para validar almacén, logística, finanzas y reportes antes de producción.

2. Reportes nuevos
   - Control ERP > Kardex valorizado por artículo.
   - Control ERP > Compras por proveedor.
   - Control ERP > Consumo por centro de costo y OT.
   - Control ERP > Cuentas por pagar / Aging.

3. Cierre mensual de almacén
   - Menú: Control ERP > Cierre mensual de almacén.
   - Guarda snapshot JSON de valorización mensual.
   - Bloquea movimientos en periodos cerrados.
   - Solo ADMIN puede registrar ajustes autorizados sobre periodos cerrados.

4. Historial de precios por artículo
   - Menú: Control ERP > Historial de precios por artículo.
   - Registra precio sin IGV, con IGV, moneda, tipo de cambio y equivalente en soles por OC.

5. Control de duplicidades reforzado
   - Pago duplicado por banco + número de operación + monto.
   - Factura duplicada para el mismo proveedor.
   - Se mantiene bloqueo existente de artículo duplicado, proveedor RUC duplicado, OC duplicada por requerimiento activo y recepción mayor a lo aprobado.

6. Permisos por acción
   - Menú: Control ERP > Permisos por acción.
   - Acciones configurables: crear OC, aprobar OC, anular OC, reversar vale, registrar pago, cerrar mes, modificar proveedor.

7. Auditoría visible
   - Menú: Control ERP > Ver trazabilidad de documento.
   - Permite consultar trazabilidad de OC, Vale o auditoría general.

8. Documentación interna
   - README.md actualizado.
   - Manuales separados por área.
   - Diccionario de base de datos.
   - Flujos del sistema.

## Pruebas ejecutadas

- python -m py_compile app_kardex.py
- python pruebas_oc_erp.py
- python pruebas_v24_valorizacion.py
- python pruebas_v39_mejoras4.py
- python pruebas_v40_consolidacion.py
- xvfb-run -a python pruebas_ui_visual_v21.py
- xvfb-run -a python pruebas_v36_mejoras2.py

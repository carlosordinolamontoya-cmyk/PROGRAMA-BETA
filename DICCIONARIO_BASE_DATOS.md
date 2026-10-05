# Diccionario resumido de base de datos

## Tablas principales

- `articulos`: maestro de artículos, valorización y decimales.
- `movimientos`: cabecera de vales CN/CO/AJ.
- `movimiento_detalle`: detalle de artículos por vale, cantidad y costo sin IGV.
- `requerimientos`: cabecera de requerimientos.
- `requerimiento_detalle`: detalle solicitado y atendido.
- `proveedores`: maestro de proveedores.
- `orden_compra`: cabecera de OC/OS.
- `orden_compra_detalle`: cantidades, precios, IGV, total y recepción.
- `orden_compra_pagos`: pagos de finanzas con factura y operación.
- `orden_compra_transito`: transporte y guías.
- `orden_compra_recepciones`: ingresos por OC.
- `orden_compra_liquidacion`: cierre documental de OC.
- `auditoria`: eventos generales del sistema.
- `tipo_cambio`: TC por fecha.
- `cierre_mensual_almacen`: periodos cerrados y snapshot de valorización.
- `historial_precios_articulo`: histórico de precios comprados por artículo/proveedor.
- `usuario_permiso_accion`: permisos por acción.

## V44

- `correlativos`: secuencias anuales por tipo documental (`tipo`, `anio`, `ultimo`).
- `schema_version`: historial de versiones de estructura; V44 registra versión `44` al migrar.
- `sistema_config`: además de configuraciones anteriores, V44 registra motor beta `sqlite` y la política de precisión numérica.

Los correlativos nuevos ya no incluyen el almacén en el número. AQP/MIN permanece almacenado en las columnas propias de cada documento.

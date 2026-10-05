# V42 - Corrección de errores críticos

## 1. Protección de la base de producción en pruebas
- `pruebas_v23_ux_funcional.py`, `pruebas_v24_valorizacion.py` y `pruebas_v29_dialogos_visual.py` ya no usan ni eliminan `app.DB_FILE`.
- Las pruebas trabajan con bases temporales independientes.
- Se agregó `pruebas_v42_criticos.py`, que verifica además que `kardex_almacen.db` no cambie durante la prueba.

## 2. Recepción de OC por almacén
- La pantalla valida el almacén real de la OC.
- La capa de base de datos vuelve a validar el permiso antes de registrar el movimiento.
- Un usuario sin permiso sobre el almacén de destino no puede generar vale, recepción ni stock.

## 3. Consumo por centro de costo / OT
- Las transferencias internas ya no se contabilizan como consumo.
- Los movimientos de reverso tampoco se consideran consumo operativo.
- El reporte se limita a salidas operativas `CO` reales.

## 4. Pagos parciales con la misma factura
- Se permiten múltiples pagos/cuotas con el mismo número de factura dentro de una misma OC.
- Se mantiene el bloqueo de operación bancaria duplicada.
- Se mantiene el bloqueo si la misma factura del mismo proveedor aparece en otra OC.

## Validación
Ejecutar:

```bash
python pruebas_v42_criticos.py
```

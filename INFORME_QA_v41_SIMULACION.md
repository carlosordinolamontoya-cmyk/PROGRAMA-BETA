# Informe QA v41 - Simulación operativa controlada

## Base revisada
Se tomó como base `kardex_almacen_python_v40_consolidacion_reportes_cierre`.

## Alcance probado
Se ejecutó una simulación completa del flujo recomendado antes de producción:

1. Crear artículos valorizables, no valorizables, con enteros/decimales.
2. Registrar solicitantes y proveedores.
3. Registrar tipo de cambio para reportes con conversión.
4. Registrar ingreso manual CN con precio sin IGV.
5. Registrar salida CO y validar bloqueo de stock negativo.
6. Crear requerimiento con fecha de entrega solicitada.
7. Generar OC por requerimiento y verificar que el requerimiento ya no quede disponible para otra OC activa.
8. Aprobar OC.
9. Registrar pago con factura.
10. Registrar tránsito.
11. Registrar entrada por OC y validar vale.
12. Liquidar OC.
13. Emitir reportes: Kardex valorizado por artículo, compras por proveedor, consumo por CC/OT y cuentas por pagar.
14. Cerrar periodo mensual y validar bloqueo de movimientos antiguos.
15. Consultar trazabilidad de OC y vale.

## Error encontrado
Se encontró un bug en la generación multiproveedor desde requerimiento: si el primer proveedor generaba una OC correctamente y el segundo proveedor fallaba, podía quedar una OC parcial del lote.

## Corrección aplicada
Se corrigió `create_ocs_from_requerimiento_multiproveedor` para que:

- Primero pre-valide todos los proveedores, moneda y precios antes de insertar.
- Use una transacción explícita `BEGIN`.
- Si cualquier proveedor o item falla, hace `rollback` y no deja OC parcial.

## Ajuste de pruebas
La prueba legacy `pruebas_v23_ux_funcional.py` usaba un usuario ficticio `tester` sin permisos creados en la base. Se actualizó para usar `admin`, porque desde v40 existen permisos por acción.

## Validaciones ejecutadas

- `python -m py_compile app_kardex.py`
- `python pruebas_oc_erp.py`
- `python pruebas_v19_correcciones.py`
- `python pruebas_v21_integrales.py`
- `python pruebas_v22_finanzas_credito.py`
- `python pruebas_v23_ux_funcional.py`
- `python pruebas_v24_valorizacion.py`
- `python pruebas_v39_mejoras4.py`
- `python pruebas_v40_consolidacion.py`
- `python pruebas_v41_simulacion_operativa_controlada.py`
- `xvfb-run -a python pruebas_v26_dialogos_tc.py`
- `xvfb-run -a python pruebas_v27_limpio_vales.py`
- `xvfb-run -a python pruebas_v29_dialogos_visual.py`
- `xvfb-run -a python pruebas_v31_autocomplete.py`
- `xvfb-run -a python pruebas_v32_decimal_ocdetalle.py`
- `xvfb-run -a python pruebas_v34_mejoras_word_finanzas.py`
- `xvfb-run -a python pruebas_v36_mejoras2.py`
- `xvfb-run -a python pruebas_ui_visual_v21.py`

Resultado: OK.

## Resultado de simulación
La simulación operativa controlada generó correctamente:

- Requerimiento: `REQ-AQP-000001`
- Orden de compra: `OC-AQP-000001`
- Vale salida: `AQP-CO-000001`
- Vale ingreso por OC: `AQP-CN-000002`
- Kardex valorizado: 2 filas
- Compras por proveedor: 1 fila
- Consumo por CC/OT: 1 fila
- Aging cuentas por pagar: 1 fila

## Estado final
La versión queda lista para prueba operativa controlada por usuarios reales en VS Code.

# Informe QA v42 - Errores críticos

## Alcance
Validación de los cuatro hallazgos críticos detectados durante la auditoría del ERP Beta.

## Resultados

### 1. Protección de base de producción
**Resultado: OK.**

- Las pruebas v23, v24 y v29 fueron desacopladas de `app.DB_FILE`.
- `pruebas_v42_criticos.py` trabaja exclusivamente con una base temporal.
- Se comparó `kardex_almacen.db` de la versión corregida contra la base extraída directamente del RAR original.
- SHA-256 en ambos archivos: `701742a1e8cd9ac3dfc99fe5b0e12c4f54010de4f4d23aa6aa29c18a14bf9b9c`.

### 2. Recepción de OC y permisos de almacén
**Resultado: OK.**

Caso probado:
- Usuario con permiso únicamente para AQP.
- OC perteneciente a MIN.
- El sistema bloquea la recepción.
- No crea vale.
- No crea registro de recepción.
- No aumenta stock de MIN.
- Una OC de AQP sí puede recibirse con el mismo usuario.

### 3. Consumo por centro de costo / OT
**Resultado: OK.**

Caso probado:
- Ingreso de 10 unidades en AQP.
- Transferencia de 4 unidades AQP -> MIN.
- Consumo reportado por la transferencia: 0.
- Salida operativa posterior de 2 unidades.
- Consumo reportado: 2 unidades.

También se excluyen los movimientos generados como reversos/correcciones.

### 4. Pagos parciales de una factura
**Resultado: OK.**

Caso probado:
- OC total S/ 118.00.
- Factura F001-V42.
- Pago 1: S/ 40.00.
- Pago 2: S/ 78.00.
- Estado final: PAGADA.
- Se conservan dos registros de pago asociados a la misma factura.
- Una operación bancaria duplicada sigue bloqueada.
- La misma factura del mismo proveedor asociada a otra OC sigue bloqueada.

## Regresión
Pasaron las pruebas de:
- módulo OC/ERP;
- correcciones v19;
- integrales v21/v22;
- finanzas crédito v22;
- UX funcional v23;
- valorización v24;
- diálogos/tipo de cambio v26;
- vales v27;
- visual v29;
- autocomplete v31;
- decimales OC v32;
- Word/finanzas v34;
- mejoras v36;
- mejoras v39;
- consolidación v40;
- simulación operativa v41;
- prueba crítica v42;
- suite visual general v21/v22.

Las pruebas gráficas se ejecutaron con una pantalla virtual por tratarse de un entorno sin escritorio físico.

## Estado
**V42 crítica aprobada para continuar con la siguiente fase de endurecimiento del ERP.**

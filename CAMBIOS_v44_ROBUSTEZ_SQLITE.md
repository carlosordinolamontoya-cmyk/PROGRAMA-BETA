# Cambios V44 - Robustez SQLite, correlativos anuales y modularidad

## Alcance

V44 mantiene SQLite como motor de la fase beta. No cambia la forma de abrir el ERP ni crea aplicaciones separadas. La separación realizada es únicamente interna mediante `erp_core/`.

## 1. Correlativos anuales

Nuevos formatos:

- `REQ-AAAA-000001`
- `OC-AAAA-000001`
- `OS-AAAA-000001`
- `VALE-AAAA-000001` para ingresos y salidas
- `TRF-AAAA-000001`
- `AJ-AAAA-000001` para ajustes

Las series son globales por tipo y año. El almacén AQP/MIN sigue guardado como dato propio del documento, pero ya no forma parte del correlativo.

Se agregó la tabla `correlativos(tipo, anio, ultimo)` y la reserva usa una transacción `BEGIN IMMEDIATE` cuando corresponde. Los formatos históricos continúan siendo consultables y no se renumeran.

## 2. Política de precisión Decimal

Se incorporó `erp_core/money.py` con `Decimal` y redondeo `ROUND_HALF_UP`:

- Cantidad: 3 decimales.
- Precio/costo unitario: 4 decimales.
- Tipo de cambio: 4 decimales.
- Importes finales: 2 decimales.

El cálculo de líneas de OC, recálculo de totales y pagos usa esta política antes de persistir valores. Durante la beta se conservan las columnas SQLite `REAL` por compatibilidad con la base existente.

## 3. Fechas internas

`creado_en`, `modificado_en`, auditoría y otras marcas internas nuevas usan:

`YYYY-MM-DD HH:MM:SS`

La fecha visible para el usuario continúa en `DD/MM/AAAA`. Al migrar se normalizan timestamps históricos reconocibles sin alterar fechas operativas de negocio.

## 4. Rendimiento SQLite

Se agregaron índices para consultas frecuentes sobre:

- movimientos y detalle;
- requerimientos y detalle;
- OC y detalle;
- pagos, facturas y operaciones bancarias;
- recepciones y tránsito;
- auditoría;
- tipo de cambio;
- proveedores;
- historial de precios.

También se configura `busy_timeout=15000 ms`. En disco local se intenta usar WAL; no se fuerza WAL en rutas UNC.

## 5. Logs técnicos

Se agregó `erp_core/logging_utils.py`.

Los errores no controlados se registran en `logs/erp.log` con rotación diaria y retención de 30 archivos. No se registran contraseñas.

## 6. Modularización progresiva

Se creó:

```text
erp_core/
├── __init__.py
├── correlatives.py
├── dates.py
├── logging_utils.py
├── money.py
└── sqlite_utils.py
```

El usuario sigue ejecutando una única aplicación. Esta estructura prepara la extracción gradual de Compras, Almacén y Finanzas sin reescribir el ERP de golpe.

## 7. Compatibilidad y migración

- Esquema actual: 44.
- Se crea backup antes de migrar.
- Una restauración de una base anterior se actualiza inmediatamente al esquema 44.
- V42 y V43 continúan pasando sus pruebas de regresión.
- Se conserva `admin123` en la fase beta, almacenada con PBKDF2 desde V43.

## Nota de arquitectura

SQLite se mantiene para pulir la beta. No debe usarse un único archivo `.db` compartido directamente por Internet entre Arequipa y Mina. La futura migración a PostgreSQL se realizará después de estabilizar procesos, maestros y reportes.

# Sistema Kardex / ERP de Almacén - V44 Beta

**Versión actual de trabajo: V44 - Robustez SQLite, correlativos anuales, precisión y modularidad.**

Acceso inicial de la fase beta:
- Usuario: `admin`
- Contraseña: `admin123`

Novedades V44:
- Se mantiene SQLite durante la fase beta.
- Correlativos anuales globales por tipo: `REQ-2026-000001`, `OC-2026-000001`, `OS-2026-000001`, `VALE-2026-000001`, `TRF-2026-000001` y `AJ-2026-000001`.
- Las series reinician en `000001` cada año y no dependen del almacén.
- Reserva transaccional de correlativos para evitar duplicados entre procesos.
- Política Decimal: cantidades 3 decimales, precios/costos 4, tipo de cambio 4 e importes finales 2.
- Timestamps internos normalizados a `YYYY-MM-DD HH:MM:SS`; la interfaz mantiene `DD/MM/AAAA`.
- Índices SQLite para movimientos, OC, requerimientos, pagos, auditoría y precios.
- `busy_timeout` para reducir errores de base ocupada durante pruebas concurrentes.
- Logs técnicos rotativos en la carpeta `logs`.
- Inicio de separación interna mediante el paquete `erp_core/`; el usuario continúa usando una sola aplicación.
- Migración V44 con respaldo previo y compatibilidad con los formatos históricos.

Novedades V43 conservadas:
- Backup diario automático y backup pre-migración.
- Restauración con respaldo de emergencia e integridad SQLite.
- Contraseñas almacenadas con PBKDF2 manteniendo `admin123` en beta.
- Permisos por acción aplicados en la lógica de negocio.
- Menú `Seguridad / Respaldo` para ADMIN.

### Importante sobre SQLite durante la beta
SQLite se mantiene deliberadamente en V44. Para pruebas con varias instancias en **un mismo equipo/servidor local**, V44 incorpora espera de bloqueos y correlativos transaccionales. No se recomienda colocar `kardex_almacen.db` directamente en una carpeta compartida por Internet/WAN entre Arequipa y Mina. La operación multi-sede definitiva se migrará a PostgreSQL una vez estabilizada la beta.

> El historial inferior conserva la documentación de versiones anteriores.

---

# Sistema Kardex de Almacén - V11

Usuario inicial:
- Usuario: admin
- Contraseña: admin123

Ejecutar:
```powershell
pip install openpyxl
python app_kardex.py
```

Cambios V11:
- Logística -> Requerimientos.
- Almacén -> Atención de Requerimientos.
- Estados de requerimiento: PENDIENTE, DESPACHO PARCIAL, ATENDIDO, ANULADO.
- Requerimientos por almacén destino AQP / MIN.
- Stock separado por almacén en movimientos y validación de salidas.
- Vales con prefijo por almacén y tipo: AQP-CO-000001, MIN-CN-000001, etc.
- Requerimientos con correlativo: REQ-AQP-000001, REQ-MIN-000001.
- Reportes de requerimientos pendientes, parciales, atendidos y pendientes por artículo.

Nota:
Esta versión no incluye todavía el módulo de Orden de Compra. El requerimiento queda pendiente o parcial hasta que exista stock; luego almacén puede volver a atender el mismo requerimiento.


## V12
- Agrega pantalla de carga al iniciar el sistema.
- Mantiene la lógica de Requerimientos, Atención de Requerimientos, estados y multi-almacén de la V11.


V13 - Cambios principales:
- Logística > Consulta de Requerimientos con filtros por fecha, día, mes, estado, almacén, solicitante, centro de costo y búsqueda libre.
- Se reemplazan los reportes separados de requerimientos por un solo Reporte General de Requerimientos exportable desde la consulta.
- Shift + F2 abre consulta de stock.
- Escape cierra ventanas secundarias.
- Barras de desplazamiento en grillas principales de requerimientos y atención.
- Un requerimiento puede tener varias atenciones parciales sin bloquearse por documento duplicado.

V14 - Correcciones operativas solicitadas:
- Bloqueo de stock negativo cuando un mismo artículo se repite dentro de una salida o ajuste.
- Fechas de consulta separadas de fechas de movimiento; los reportes/filtros ya permiten rangos futuros sin bloquear la consulta.
- Permisos reales por rol: CONSULTA queda restringido a consultas y reportes; ADMIN/OPERADOR pueden registrar y modificar.
- Activación de llaves foráneas de SQLite con PRAGMA foreign_keys = ON.
- Atención de requerimientos en transacción única: vale, detalle, estado e historial se confirman juntos.
- Consulta/reporte de stock con stock comprometido y stock libre por almacén.
- Transferencia formal entre almacenes: genera salida del origen e ingreso del destino con un código TR vinculado.
- Stock mínimo por almacén AQP/MIN, manteniendo compatibilidad con stock mínimo general.
- Reverso de vales y anulación controlada de requerimientos pendientes.
- Descripción de artículos ampliada a 120 caracteres.
- Carga masiva de artículos atómica: si una fila falla, no se guarda ninguna.

Se mantiene fuera de esta versión, según indicación:
- Ubicación física.
- Lote, serie y vencimiento.

## V15 ERP/MRP - Órdenes de Compra, Finanzas y Tránsito

Cambios principales:
- Login con selección de almacén de trabajo AQP/MIN.
- Permisos usuario-almacén configurables desde Usuarios.
- Roles ampliados: ADMIN, OPERADOR, ALMACEN, LOGISTICA, FINANZAS, CONSULTA.
- Menú Requerimientos separado de Logística.
- Módulo Logística con registro, consulta, aprobación, entrada, tránsito, liquidación y reporte de Órdenes de Compra.
- Módulo Finanzas para registrar pagos de OC con fecha, banco y número de operación.
- Maestro de proveedores con datos bancarios y moneda preferida.
- Orden nacional y orden de servicio.
- Orden de servicio no genera Kardex de ingreso.
- Entrada con OC genera un vale CN vinculado a la orden.
- Aprobación total/parcial/denegada por ADMIN.
- Control de cantidades: no se aprueba más de lo solicitado y no se recibe más de lo aprobado pendiente.
- Estados y trazabilidad de OC: PENDIENTE, EN COTIZACION, APROBADA, APROBADA PARCIAL, DENEGADA, ANULADA, PAGO PARCIAL, PAGADA, EN TRANSITO, ATENDIDA PARCIALMENTE, ATENDIDA, LIQUIDADA.
- Avance automático: 0%, 25%, 50%, 80% y 100%.
- Formato Excel de Orden de Compra / Orden de Servicio para guardar o imprimir.
- Pruebas automáticas incluidas en `pruebas_oc_erp.py`.

Para validar el módulo OC/ERP:
```powershell
python pruebas_oc_erp.py
```

## v16 - RUC SUNAT/API y Tipo de Cambio

Esta versión agrega integración configurable para consulta de RUC y tipo de cambio.

### Proveedores

Menú: `Logística > Proveedores / Nuevo proveedor`.

Opciones principales:

- `Nuevo proveedor`: limpia el formulario para registrar un proveedor desde cero.
- `Consultar RUC API`: consulta el RUC y carga razón social, dirección, estado, condición y ubigeo si la API está configurada y disponible.
- `Guardar proveedor`: guarda o actualiza el proveedor.

Los datos bancarios, cuenta y CCI se mantienen como registro manual porque normalmente no vienen desde SUNAT.

### Configuración de APIs

Menú: `Configuración > Configuración de APIs`.

Permite configurar URLs, tokens y timeout. Las URLs pueden usar `{ruc}` o `{fecha}` como variables.

### Tipo de cambio

Menú: `Finanzas > Tipo de Cambio`.

Opciones:

- Cargar tipo de cambio desde API.
- Registrar tipo de cambio manual.
- Consultar histórico.
- Usar un tipo de cambio seleccionado al emitir una OC en dólares.

La OC en dólares no se puede emitir sin tipo de cambio. Si no hay internet/API, se debe registrar manualmente.


## v22 - Seguimiento financiero de OC a crédito

La OC ahora maneja dos controles separados:

- **Estado operativo / avance:** logística y almacén. Ejemplo: APROBADA, EN TRANSITO, ATENDIDA, LIQUIDADA.
- **Estado de pago:** finanzas. Ejemplo: POR PAGAR, PAGO PARCIAL, PAGADA, VENCIDA.

Esto permite que una OC a crédito llegue al 100% de avance operativo y siga visible para Finanzas hasta que se registre el pago.

Nuevo menú:

`Finanzas > Seguimiento de Pagos OC`

Desde ahí se filtran cuentas por pagar, vencidas, parcialmente pagadas y pagadas.

## v40 - Consolidación ERP/MRP

Se agregó el menú **Control ERP** con:

- Prueba operativa controlada.
- Kardex valorizado por artículo.
- Compras por proveedor.
- Consumo por centro de costo y OT.
- Cuentas por pagar / Aging.
- Historial de precios por artículo.
- Cierre mensual de almacén.
- Trazabilidad de documento.
- Permisos por acción.

Esta versión no incluye dashboard principal por solicitud del usuario.

## v42 - Correcciones críticas de auditoría

Se corrigieron cuatro hallazgos críticos antes de continuar con mejoras de arquitectura:

1. Las pruebas automatizadas ya no apuntan ni eliminan `kardex_almacen.db`.
2. La recepción de OC valida el permiso contra el almacén real de la orden tanto en UI como en la capa de datos.
3. El reporte de consumo por centro de costo/OT excluye transferencias internas y reversos.
4. Una misma factura puede tener varios pagos parciales dentro de la misma OC, manteniendo controles de operación bancaria duplicada y factura repetida en otra OC.

Prueba dirigida: `python pruebas_v42_criticos.py`.

Detalle: `CAMBIOS_v42_ERRORES_CRITICOS.md` e `INFORME_QA_v42_CRITICOS.md`.

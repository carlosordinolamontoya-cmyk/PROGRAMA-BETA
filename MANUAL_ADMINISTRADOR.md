# Manual del administrador

## Funciones principales

- Crear usuarios y asignar roles.
- Configurar permisos por acción desde Control ERP > Permisos por acción.
- Registrar configuración de API.
- Cerrar mes de almacén.
- Autorizar ajustes en periodo cerrado.
- Consultar auditoría y trazabilidad.

## Permisos por acción

Las acciones configurables son:

- Crear OC.
- Aprobar/desaprobar OC.
- Anular OC.
- Reversar vale.
- Registrar pago.
- Cerrar mes.
- Modificar proveedor.

El rol ADMIN conserva acceso total.

## Nota V44 - SQLite durante la beta

- La aplicación continúa siendo única; `erp_core/` es solo organización interna del código.
- Los correlativos nuevos son anuales y globales por tipo documental.
- Los errores técnicos quedan registrados en `logs/erp.log`.
- SQLite se mantiene para la beta. No compartir directamente `kardex_almacen.db` por Internet/WAN entre sedes.
- Para pruebas concurrentes controladas en un mismo equipo/servidor local se configuró espera de bloqueos e índices de rendimiento.
- La migración a PostgreSQL se realizará después de estabilizar la beta.

# V43 - Seguridad, respaldos y permisos efectivos

## Alcance
Esta versión continúa sobre V42 y conserva el flujo funcional existente. La contraseña inicial solicitada para la etapa beta se mantiene:

- Usuario: `admin`
- Contraseña: `admin123`

La mejora se aplica a la forma de almacenar/verificar la contraseña y a la protección de operaciones sensibles.

## 1. Contraseñas
- Las nuevas contraseñas se almacenan con PBKDF2-HMAC-SHA256 y salt aleatorio.
- Se mantiene compatibilidad con hashes SHA-256 de V42 y anteriores.
- Al iniciar sesión correctamente con una cuenta antigua, su hash se migra automáticamente a PBKDF2.
- `admin123` sigue siendo la contraseña inicial durante la fase beta por decisión del proyecto.

## 2. Respaldos automáticos
- Se crea un respaldo diario en `backups/diarios/`.
- Se conserva una copia por día y se aplica retención automática de 30 días.
- Antes de migrar una base anterior a V43 se genera un respaldo consistente en `backups/pre_migracion/`.
- Los respaldos se verifican con `PRAGMA integrity_check`.

## 3. Respaldos manuales y restauración
Se agregó el menú `Seguridad / Respaldo` para ADMIN con:
- Crear respaldo ahora.
- Restaurar respaldo.
- Verificar integridad de la base.

Antes de restaurar una BD se crea un respaldo de emergencia de la base actual. El archivo seleccionado también se verifica antes de reemplazar la base activa.

## 4. Migraciones más seguras
- Se agregó tabla `schema_version`.
- V43 registra la versión de esquema 43.
- Antes de migrar una base anterior se crea un backup pre-migración.
- Si la inicialización/migración falla, el sistema intenta restaurar automáticamente el backup previo.

## 5. Permisos efectivos en la lógica de negocio
Los siguientes permisos ya no dependen únicamente de botones o del rol visual:
- `crear_oc`: crear y modificar OC pendiente.
- `aprobar_oc`: aprobar, aprobar parcialmente, denegar y desaprobar OC.
- `anular_oc`: anular OC.
- `reversar_vale`: reversar/anular vale.
- `registrar_pago`: registrar pagos de OC.
- `cerrar_mes`: cerrar periodo de almacén.
- `modificar_proveedor`: crear/modificar proveedores.

La validación se realiza en la capa de negocio. Una llamada interna o alternativa también queda protegida.

## 6. Administración de permisos
- Solo un usuario ADMIN puede modificar permisos por acción.
- ADMIN se conserva como superusuario con acceso total para evitar bloquear la administración del ERP.
- La interfaz de permisos muestra los permisos de ADMIN como activos.
- La interfaz principal ahora habilita opciones de acuerdo con permisos de acción, no solo por rol, para las operaciones cubiertas por V43.

## 7. Compatibilidad
- Se mantienen las correcciones críticas de V42.
- La base `kardex_almacen.db` incluida en el proyecto no fue modificada durante las pruebas de desarrollo.
- Los scripts de prueba antiguos que usaban usuarios ficticios en operaciones protegidas fueron actualizados para usar cuentas de prueba reales.

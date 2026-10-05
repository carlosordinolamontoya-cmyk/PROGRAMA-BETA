# Informe QA V43 - Seguridad, respaldos y permisos

## Resultado general
**APROBADO PARA CONTINUAR FASE BETA CONTROLADA.**

## Pruebas específicas V43
Archivo: `pruebas_v43_seguridad_backup.py`

Validado:
- Login `admin / admin123` continúa funcionando.
- Contraseña incorrecta es rechazada.
- Migración transparente SHA-256 legado -> PBKDF2.
- Solo ADMIN puede modificar permisos.
- ADMIN conserva acceso total.
- Permiso de proveedor bloquea y habilita gestión real.
- Permiso crear OC bloquea/autoriza creación y modificación de OC pendiente.
- Permiso aprobar OC bloquea/autoriza aprobación a usuario no-ADMIN.
- Permiso anular OC bloquea/autoriza anulación.
- Permiso reversar vale bloquea/autoriza reverso.
- Permiso registrar pago bloquea/autoriza pago.
- Permiso cerrar mes bloquea/autoriza cierre.
- `schema_version = 43`.
- `PRAGMA integrity_check = ok`.
- Backup diario válido.
- Backup manual válido.
- Restauración devuelve los datos al estado respaldado.
- Se crea respaldo de emergencia antes de restaurar.
- Archivo corrupto es rechazado sin reemplazar la BD activa.
- Backup pre-migración conserva los datos originales.

Resultado: `PRUEBAS V43 SEGURIDAD/BACKUP: OK`

## Prueba UI de permisos
Archivo: `pruebas_v43_ui_permisos.py`

Validado con pantalla virtual:
- Usuario CONSULTA con permiso `aprobar_oc` ve Aprobación de OC habilitada.
- Sin `crear_oc`, Registro OC queda deshabilitado.
- Sin `modificar_proveedor`, gestión de proveedores queda deshabilitada.
- Con `registrar_pago`, Registrar Pago queda habilitado.
- Menú de respaldo queda deshabilitado para no-ADMIN.
- Menú de respaldo queda habilitado para ADMIN.

Resultado: `PRUEBAS V43 UI/PERMISOS: OK`

## Regresión funcional
Pasaron después de V43:
- `pruebas_oc_erp.py`
- `pruebas_v19_correcciones.py`
- `pruebas_v21_integrales.py`
- `pruebas_v22_finanzas_credito.py`
- `pruebas_v23_ux_funcional.py`
- `pruebas_v24_valorizacion.py`
- `pruebas_v27_limpio_vales.py`
- `pruebas_v31_autocomplete.py`
- `pruebas_v32_decimal_ocdetalle.py`
- `pruebas_v34_mejoras_word_finanzas.py`
- `pruebas_v36_mejoras2.py`
- `pruebas_v39_mejoras4.py`
- `pruebas_v40_consolidacion.py`
- `pruebas_v41_simulacion_operativa_controlada.py`
- `pruebas_v42_criticos.py`

Pruebas visuales verificadas con pantalla virtual:
- `pruebas_ui_visual_v21.py`
- `pruebas_v26_dialogos_tc.py`
- `pruebas_v29_dialogos_visual.py`

## Protección de la base incluida
SHA-256 de `kardex_almacen.db` antes y después de las pruebas:
`701742a1e8cd9ac3dfc99fe5b0e12c4f54010de4f4d23aa6aa29c18a14bf9b9c`

La base incluida no fue alterada por la batería de pruebas.

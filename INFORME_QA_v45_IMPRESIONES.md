# Informe QA V45 - Impresiones

## Objetivo

Validar la corrección de paginación de vales de almacén y el crecimiento dinámico de las órdenes de compra/servicio.

## Casos dirigidos

### Vale con 2 ítems
- Esperado: 1 página.
- Resultado: 1 página.
- Se eliminan 10 filas vacías y las firmas permanecen en la misma hoja.

### Vale con 12 ítems
- Esperado: 1 página.
- Resultado: 1 página.

### Vale con 13 ítems
- Esperado: 2 páginas completas del mismo vale.
- Resultado: 2 páginas.
- Página 1: ítems 1-12, `01 de 02`.
- Página 2: ítem 13, `02 de 02`.
- Encabezado y estructura del vale repetidos correctamente.

### Vale con 25 ítems
- Esperado: 3 páginas.
- Resultado: 3 páginas.
- Distribución lógica: 12 + 12 + 1.

### OC/OS con más de 16 ítems
- Se probaron 17, 25, 40 y 60 ítems.
- Las filas adicionales se agregan a la misma tabla.
- El encabezado general no se duplica.
- El cierre no se duplica.
- Con 40 ítems, el documento continúa a una segunda página y únicamente el detalle restante + cierre aparecen allí.

## QA visual Word

Se renderizaron e inspeccionaron mediante LibreOffice los siguientes casos finales:

- Vale de 2 ítems: 1 página, sin página fantasma.
- Vale de 13 ítems: 2 páginas con encabezados completos y numeración correcta.
- OC de 40 ítems: 2 páginas, un solo encabezado general y un solo cierre.

No se observaron recortes, solapamientos ni páginas en blanco en los casos validados.

## Regresión

Se ejecutaron 22 scripts de pruebas del proyecto usando pantalla virtual para las pruebas Tkinter.

Resultado:

- 22/22 scripts aprobados.
- 0 fallos.

Incluye pruebas de OC, requerimientos, finanzas, valorización, interfaz, permisos, backups, errores críticos V42, seguridad V43, robustez V44 y nuevas pruebas V45.

## Integridad de la base incluida

SHA-256 antes de la regresión:

`701742a1e8cd9ac3dfc99fe5b0e12c4f54010de4f4d23aa6aa29c18a14bf9b9c`

SHA-256 después de la regresión:

`701742a1e8cd9ac3dfc99fe5b0e12c4f54010de4f4d23aa6aa29c18a14bf9b9c`

La base de datos incluida no fue alterada por las pruebas.

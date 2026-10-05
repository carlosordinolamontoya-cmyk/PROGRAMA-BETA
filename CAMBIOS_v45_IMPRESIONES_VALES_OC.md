# V45 - Impresiones de vales y OC/OS

Base: V44 Robustez SQLite.

## 1. Vales de almacén

Se corrigió la generación Word de vales de salida, ingreso y ajuste.

- La plantilla oficial tiene 12 filas de detalle por hoja.
- Un vale con 1 a 12 ítems se genera en una sola hoja.
- De 13 a 24 ítems se generan dos hojas.
- A partir de 25 ítems se agregan las hojas necesarias, en bloques de 12.
- Cada hoja adicional repite el formato completo del vale: encabezado SIG, datos del vale, cabecera del detalle y firmas.
- Se actualiza la numeración de página automáticamente: `01 de 02`, `02 de 02`, etc.
- Las filas vacías del detalle se eliminan. Esto corrige el defecto por el cual un vale de 1 o 2 productos enviaba el bloque de firmas a una segunda página.
- Se corrigió el título duplicado `VALE DE SALIDA DE ALMACEN DE ALMACÉN`; ahora se muestra `VALE DE SALIDA DE ALMACÉN`.
- La ventana de previsualización ahora indica correctamente A5 horizontal.

## 2. Orden de compra / servicio

Se cambió el tratamiento de las líneas de detalle de la OC/OS.

- La plantilla conserva sus 16 filas base.
- Cuando existen más de 16 ítems, el programa inserta automáticamente las filas adicionales dentro de la misma tabla.
- No se crea una segunda copia del encabezado general de la OC/OS.
- No se duplica el bloque de banco, subtotal, IGV, total, información al proveedor ni firmas.
- Si por cantidad de filas Word necesita una segunda hoja, la tabla continúa naturalmente y el cierre aparece una sola vez al final.
- Se eliminó el párrafo vacío final que podía provocar una página en blanco adicional.

## 3. Compatibilidad

- No se modificó el esquema de la base de datos. La BD continúa en esquema V44.
- No se cambió SQLite.
- No se modificaron correlativos, cálculos, seguridad, permisos ni flujos funcionales.
- Las plantillas Word originales permanecen editables.

## 4. Constantes configurables

Se agregaron en `app_kardex.py`:

- `VALE_ITEMS_PER_PAGE = 12`
- `OC_BASE_ITEM_ROWS = 16`

Esto permite cambiar posteriormente los límites si se rediseñan las plantillas.

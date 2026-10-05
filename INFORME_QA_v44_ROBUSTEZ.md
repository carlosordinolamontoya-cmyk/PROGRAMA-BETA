# Informe QA V44 - Robustez SQLite

## Resultado general

**APROBADO PARA CONTINUAR LA FASE BETA.**

Se ejecutó la batería histórica completa bajo pantalla virtual cuando correspondía.

- Scripts de prueba ejecutados: **21**.
- Scripts aprobados: **21**.
- Fallos finales: **0**.

## Casos V44 específicos

### Correlativos

Validado:

- reinicio por año;
- serie global independiente del almacén;
- series separadas OC/OS/REQ/VALE/TRF/AJ;
- continuación correcta al abrir una segunda instancia;
- prueba concurrente equivalente a 5 usuarios: **125 correlativos OC únicos, sin duplicados y sin saltos provocados por concurrencia en la prueba**.

### Decimal

Validado:

- cantidad a 3 decimales;
- precio unitario a 4;
- tipo de cambio a 4;
- importes a 2;
- redondeo `ROUND_HALF_UP`;
- caso `0.10 + 0.20 = 0.30` usando Decimal.

### Fechas

Validado:

- nuevos timestamps en ISO `YYYY-MM-DD HH:MM:SS`;
- migración de `11/08/2026 23:59:58` a `2026-08-11 23:59:58`.

### SQLite

Validado:

- `busy_timeout >= 15000 ms`;
- creación de índices de rendimiento;
- migración segura desde una copia intacta de la base incluida hasta esquema 44;
- creación de respaldo pre-migración;
- integridad de la base posterior a la migración.

### Compatibilidad

Durante la primera regresión apareció una incompatibilidad en el reporte de valorización porque las transferencias históricas usaban prefijo `TR-` y V44 usa `TRF-AAAA-`. Se corrigió para reconocer **ambos** formatos. Luego se volvió a ejecutar toda la suite y quedó 21/21.

## Pruebas heredadas aprobadas

Incluyen OC/ERP, interfaz visual, correcciones V19, integrales, Finanzas, UX, valorización, tipo de cambio, vales, autocomplete, decimales OC, Word/Finanzas, mejoras V36/V39, consolidación V40, simulación V41, críticos V42 y seguridad/permisos V43.

## Base incluida

La base `kardex_almacen.db` que se entrega dentro de V44 se mantiene intacta respecto a la base recibida en V43; la migración a esquema 44 ocurre de forma segura en el primer inicio y genera respaldo previo.

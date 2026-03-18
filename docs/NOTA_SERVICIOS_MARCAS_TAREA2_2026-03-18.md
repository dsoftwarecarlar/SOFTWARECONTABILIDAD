# Nota servicios marcas tarea 2

Fecha: 2026-03-18

## Base operativa actual

- Plantillas activas: `resources/cxp/servicios_marcas/templates`
- Fixtures activos: `resources/cxp/servicios_marcas/fixtures`
- Configuracion vigente: `config/cxp/servicios_marcas.php`
- La ruta legacy `outputs/EJEMPLOAMANOTAREA2` ya no es la fuente principal.

## Plantillas por marca

### CHANGAN

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `REP VTAS`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### PEUGEOT

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `REP VTAS`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### SUZUKI

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `REP VTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### MATRIZ

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `REP VTAS`, `MAY VTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

## Entradas reales detectadas

### Comunes

- `SERREP_FACTURAS_NAF_REPFACT.txt`
- `SERREP_NOTACRED_NAF.txt`
- `detalle-vtas-xliquidar.xlsx`
- `RepFacturacionServContabilidad.xls`

### Por marca

- `CON_MAYORGEN2CHAN.TXT`
- `CON_MAYORGEN2PEU.TXT`
- `CON_MAYORGEN2SUZ.TXT`
- `CON_MAYORGEN2TOY.TXT`

## Rol funcional observado

### `RepFacturacionServContabilidad.xls`

- Es la fuente principal de filas operativas.
- Trae marcas mezcladas: `CHANGAN`, `MATRIZ`, `PEUGEOT`, `SUZUKI AMBATO`, `SUZUKI RIOBAMBA`.
- Tambien trae filas ajenas al flujo comercial normal como `COSTOS INSUMOS`, por lo que el filtrado debe ser estricto.

### `SERREP_FACTURAS_NAF_REPFACT.txt`

- En los fixtures actuales solo contiene `MATRIZ`.
- No alcanza por si solo para reconstruir todas las marcas.

### `SERREP_NOTACRED_NAF.txt`

- En los fixtures actuales solo contiene `MATRIZ`.
- Sirve como apoyo para notas, pero no como fuente global de todas las marcas.

### `detalle-vtas-xliquidar.xlsx`

- Trae bloques por `Marca:`.
- Alimenta la hoja `PX`.

### `CON_MAYORGEN2*.TXT`

- Su estructura coincide con `VENTAS` o `MAY VTAS`.
- `CHANGAN`, `PEUGEOT`, `SUZUKI` deben ir a `VENTAS`.
- `MATRIZ` debe ir a `MAY VTAS`.

## Dependencias internas de plantilla

Se detectaron formulas que dependen de otras hojas del mismo libro:

- `REP FACTURACION` depende de `VENTAS` o `MAY VTAS`, `PX` y `PrecontabilizacionVentas`.
- `NOTA DE CREDITO` depende de `VENTAS` o `MAY VTAS` y `PrecontabilizacionVentas`.
- `REP VTAS` depende de `VENTAS` o `MAY VTAS`, `COSTO` y `PrecontabilizacionCostos (2)`.

Conclusion: no basta con llenar `REP FACTURACION`, `NOTA DE CREDITO`, `REP VTAS` y `PX`. Si las hojas auxiliares quedan con datos historicos, la salida final conserva contenido de la plantilla.

## Huecos criticos del flujo actual

1. El formulario y el backend no contemplan los 4 archivos `CON_MAYORGEN2*.TXT`.
2. El script solo intenta llenar la hoja `VENTAS`; no contempla `MAY VTAS` para `MATRIZ`.
3. El script intenta limpiar hojas `PRECONTABILIZACION VENTAS` y `PRECONTABILIZACION COSTOS` con nombres que no existen en las plantillas reales.
4. Las hojas reales `PrecontabilizacionVentas`, `PrecontabilizacionCostos` y `PrecontabilizacionCostos (2)` no se regeneran hoy.
5. `COSTO` tampoco se regenera hoy.
6. Las formulas de resumen pueden seguir consumiendo datos viejos de plantilla aunque se reemplacen las filas visibles de `REP FACTURACION`, `NOTA DE CREDITO`, `REP VTAS`, `PX` y `VENTAS`.

## Criterio correcto para la implementacion

- La plantilla debe quedar solo como molde: estructura, formato, formulas y layout.
- Todo dato operativo debe provenir de archivos cargados o de reglas derivadas del proceso.
- Ninguna hoja auxiliar usada por formulas puede quedar con datos historicos de ejemplo.

## Orden recomendado de implementacion

1. Incorporar uploads explicitos para los `CON_MAYORGEN2*.TXT` por marca.
2. Mapear cada TXT de mayor a `VENTAS` o `MAY VTAS` segun plantilla.
3. Corregir nombres reales de hojas auxiliares y su limpieza.
4. Definir y construir la logica de generacion para `PrecontabilizacionVentas`.
5. Definir y construir la logica de generacion para `PrecontabilizacionCostos`, `PrecontabilizacionCostos (2)` y `COSTO`.
6. Recalcular y validar por marca que ningun resultado dependa de datos historicos de la plantilla.

## Riesgo principal

Mientras no se reconstruyan las hojas auxiliares dependientes de formulas, la salida no puede considerarse fiel al requerimiento de "misma plantilla, pero solo con datos cargados".

## Estado de implementacion

- Ya se incorporaron uploads opcionales para `CON_MAYORGEN2CHAN.TXT`, `CON_MAYORGEN2PEU.TXT`, `CON_MAYORGEN2SUZ.TXT` y `CON_MAYORGEN2TOY.TXT`.
- Ya existe parser dedicado para esos TXT en `scripts/cxp/servicios_marcas/read_mayor_txt.js`.
- `VENTAS` y `MAY VTAS` ya no se sobreescriben desde la fila 1. Ahora se preserva la cabecera/formato de plantilla y se reescriben solo las secciones de detalle esperadas.
- `PX` ya no se limpia completo. Se respeta el layout superior de plantilla y la carga empieza desde la fila operativa.
- Sigue pendiente la reconstruccion real de `PrecontabilizacionVentas`, `PrecontabilizacionCostos`, `PrecontabilizacionCostos (2)` y `COSTO`.

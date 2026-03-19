# Nota servicios marcas tarea 2

Fecha: 2026-03-18

## Base operativa actual

- Plantillas activas: `resources/cxp/servicios_marcas/templates`
- Fixtures activos: `resources/cxp/servicios_marcas/fixtures`
- Worker principal: `run_servicios_marcas.ps1`
- Wrapper del modulo: `scripts/cxp/servicios_marcas/run.ps1`
- UI del modulo: `templates/cxp/servicios_marcas/index.php`
- Backend del modulo: `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php`
- Runner del job: `includes/servicios_marcas_job_runner.php`

## Contrato de uploads vigente

### Inputs comunes

- `px_file` -> `detalle-vtas-xliquidar (2).xlsx`
- `repventas_file` -> `RepFacturacionServContabilidad (3).xls`

### Inputs por marca

#### CHANGAN

- `factura_changan_file` -> `SERREP_FACTURAS_NAFCHAN.TXT`
- `nota_changan_file` -> `SERREP_NOTACRED_NAFCHAN.TXT`
- `mayor_changan_file` -> `CON_MAYORGEN2CHAN.TXT`

#### PEUGEOT

- `factura_peug_file` -> `SERREP_FACTURAS_NAFPEU.TXT`
- `nota_peug_file` -> `SERREP_NOTACRED_NAFPEU.TXT`
- `mayor_peug_file` -> `CON_MAYORGEN2PEU.TXT`

#### SUZUKI

- `factura_szk_file` -> `SERREP_FACTURAS_NAFSUZAMBYRIO.TXT`
- `nota_szk_file` -> `SERREP_NOTACRED_NAFSUZAMBYRI.TXT`
- `mayor_szk_file` -> `CON_MAYORGEN2SUZ.TXT`

#### MATRIZ

- `factura_tyt_file` -> `SERREP_FACTURAS_NAFTOY.TXT`
- `nota_tyt_file` -> `SERREP_NOTACRED_NAFTOY.TXT`
- `mayor_tyt_file` -> `CON_MAYORGEN2TOY.TXT`

## Legacy eliminado del flujo

- `ventas_file` / `VENTAS.txt`
- `riobamba_file`

Conclusion: el flujo ya no depende de archivos legacy. La seleccion de marca define que trio TXT se vuelve obligatorio.

## Runtime del job

- Timeout del worker: `2700` segundos (`45 minutos`)
- Cierre de cola sin iniciar: `300` segundos
- Gracia de cancelacion: `120` segundos
- Espera de arranque del worker: `20` segundos

Conclusion: el caso `brand_key=""` ya no se corta a los 10 minutos. El modulo y el runner comparten el mismo timeout operativo.

## Plantillas por marca

### CHANGAN

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `REP VTAS`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### PEUGEOT

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `REP VTAS`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### SUZUKI

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `VENTAS`, `REP VTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

### MATRIZ

- Hojas: `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `PrecontabilizacionVentas`, `REP VTAS`, `MAY VTAS`, `COSTO`, `PrecontabilizacionCostos (2)`, `PrecontabilizacionCostos`, `ESTADISTICAS`, `Hoja1`

## Volumen auditado de fixtures reales

- `RepFacturacionServContabilidad (3).xls`: 696 filas utiles
- `detalle-vtas-xliquidar (2).xlsx`: 4 filas PX CHANGAN, 0 PEUGEOT, 6 SUZUKI, 14 MATRIZ
- `SERREP_FACTURAS_NAFCHAN.TXT`: 66 filas
- `SERREP_NOTACRED_NAFCHAN.TXT`: 0 filas
- `CON_MAYORGEN2CHAN.TXT`: 23 movimientos
- `SERREP_FACTURAS_NAFPEU.TXT`: 22 filas
- `SERREP_NOTACRED_NAFPEU.TXT`: 3 filas
- `CON_MAYORGEN2PEU.TXT`: 7 movimientos
- `SERREP_FACTURAS_NAFSUZAMBYRIO.TXT`: 258 filas
- `SERREP_NOTACRED_NAFSUZAMBYRI.TXT`: 10 filas
- `CON_MAYORGEN2SUZ.TXT`: 52 movimientos
- `SERREP_FACTURAS_NAFTOY.TXT`: 291 filas
- `SERREP_NOTACRED_NAFTOY.TXT`: 5 filas
- `CON_MAYORGEN2TOY.TXT`: 34 movimientos

## Dependencias internas de plantilla

- `REP FACTURACION` depende de `VENTAS` o `MAY VTAS`, `PX` y `PrecontabilizacionVentas`.
- `NOTA DE CREDITO` depende de `VENTAS` o `MAY VTAS` y de sus subtotales internos.
- `REP VTAS` depende de `VENTAS` o `MAY VTAS`, `COSTO` y `PrecontabilizacionCostos (2)`.
- `COSTO` depende de `ESTADISTICAS`.
- `ESTADISTICAS` depende de sus bloques internos y de las hojas de costos/precontabilizacion.

Conclusion: no basta con pegar filas visibles. Hay que regenerar bloques auxiliares y limpiar zonas operativas exactas.

## Matriz tecnica de mapeo

| Archivo | Marca | Hoja destino | Bloque / rango | Regla de transformacion | Riesgo |
| --- | --- | --- | --- | --- | --- |
| `SERREP_FACTURAS_NAF*.TXT` | Cada marca | `REP FACTURACION` | Desde fila 17, columnas A:S | Fuente primaria de facturas. Se normaliza documento, fecha, cedula, cliente, subtotal, descuento, IVA, total, TV y orden. | Si se mezcla con el consolidado, se reintroducen filas repetidas o marcas cruzadas. |
| `SERREP_NOTACRED_NAF*.TXT` | Cada marca | `NOTA DE CREDITO` | Desde fila 11, columnas A:U | Fuente primaria de NC. Se usa numero NC, factura afectada, orden, subtotal, descuento, IVA, anticipo y neto. | Si no se usa el TXT, se pierden anticipos y netos reales de NC. |
| `CON_MAYORGEN2*.TXT` | CHANGAN/PEUGEOT/SUZUKI | `VENTAS` | Bloques anclados: ventas fila 6+, descuentos fila 174+, descuentos credito fila 247+, devoluciones fila 355+ | Se limpia cada seccion detectada por la plantilla y se reescribe por cuenta contable. | Si se deja la plantilla sin limpiar, `REP FACTURACION`, `NOTA DE CREDITO` y `REP VTAS` siguen leyendo historico. |
| `CON_MAYORGEN2TOY.TXT` | MATRIZ | `MAY VTAS` | Mismos bloques logicos que `VENTAS` | Igual criterio, pero en la hoja `MAY VTAS`. | Si se escribe en `VENTAS`, las formulas de MATRIZ no cambian. |
| `detalle-vtas-xliquidar (2).xlsx` | Todas | `PX` | Desde fila operativa 6 | Se toma solo la seccion `Marca:` de cada marca y se preserva la cabecera superior. | Si se pega desde fila 1, se rompe layout y formulas del encabezado. |
| `RepFacturacionServContabilidad (3).xls` | Todas | `REP VTAS` | Desde fila 15 | Fuente base para costos, servicio, repuestos, lubricantes y cruces por documento/orden. | Tiene filas duplicadas por documento; no sirve para copiar sin consolidacion. |
| `RepFacturacionServContabilidad (3).xls` | Todas | `PrecontabilizacionVentas` | Tabla operativa y pivots de soporte | Se usa para recalcular totales y `GETPIVOTDATA` consumidos por otras hojas. | Si no se refresca, `REP FACTURACION` muestra valores historicos. |
| `RepFacturacionServContabilidad (3).xls` | Todas | `PrecontabilizacionCostos (2)` | Tabla operativa y pivots de soporte | Se regeneran filas prototipo con metricas del mes. | Si queda historico, `REP VTAS` y `COSTO` quedan contaminados. |
| `RepFacturacionServContabilidad (3).xls` | Todas | `ESTADISTICAS` y `COSTO` | Bloques calculados | Se reescriben metricas de costo reales del mes. | Si queda plantilla, el costo total sigue viejo aunque las hojas visibles cambien. |

## Anclas criticas detectadas en `VENTAS` / `MAY VTAS`

- Total ventas: fila `4`
- Resumen descuentos contado: fila `172`
- Resumen descuentos credito: fila `245`
- Resumen devoluciones / NC: fila `353`
- Detalle ventas: fila `6`
- Detalle descuentos contado: fila `174`
- Detalle descuentos credito: fila `247`
- Detalle devoluciones: fila `355`

Estas anclas alimentan formulas como:

- `REP FACTURACION!J9`
- `REP FACTURACION!K9`
- `NOTA DE CREDITO!D4`
- `NOTA DE CREDITO!E4`
- `REP VTAS!D6`

## Estado de implementacion

- La UI ya se redise?o por secciones: inputs comunes, CHANGAN, PEUGEOT, SUZUKI y MATRIZ.
- El backend ya valida por marca y obliga solo los 3 TXT de la marca seleccionada, o los 12 TXT si se procesan todas.
- El runner ya no pasa `VENTAS.txt` ni `riobamba_file`.
- El worker ya consume TXT por marca como fuente primaria de `REP FACTURACION` y `NOTA DE CREDITO`.
- El worker sigue usando `RepFacturacionServContabilidad (3).xls` como fuente principal de `REP VTAS`, costos y precontabilizacion.
- El mayor ya no preserva la plantilla cuando falta fuente: ahora limpia la hoja operativa para evitar arrastre historico.

## Riesgos residuales

1. La validacion COM estricta sigue siendo costosa, pero el timeout operativo ya subio a `45 minutos` para soportar la corrida de las 4 marcas.
2. `RepFacturacionServContabilidad (3).xls` requiere consolidacion estricta porque hay documentos repetidos, sobre todo en NC y algunos casos SUZUKI.
3. `PrecontabilizacionCostos` todavia se limpia, pero no se reconstruye con el mismo nivel de detalle que una carga manual completa.

## Siguiente verificacion obligatoria

1. Revisar visualmente en Excel una corrida del mes si cambia la estructura manual de alguna plantilla.
2. Profundizar `PrecontabilizacionCostos` si se requiere replicar cada detalle de la carga manual historica.


## Evidencia de validacion ejecutada

- `npm run test:e2e:servicios`: OK para `MATRIZ` (`brand_key=tyt`). El contrato verifico upload por marca, descarga, canario en `REP FACTURACION`, canario en `PX`, presencia de hojas y publicacion en historial.
- Worker directo `CHANGAN`: OK. Log final con `validate_done|changan`, `save_done|changan`, `invoice_fallbacks=0`, `note_fallbacks=0`.
- Worker directo `PEUGEOT`: OK. Log final con `validate_done|peug`, `save_done|peug`, `invoice_fallbacks=0`, `note_fallbacks=0`.
- Worker directo `SUZUKI`: OK. Log final con `validate_done|szk`, `save_done|szk`, `invoice_fallbacks=0`, `note_fallbacks=0`.
- Corrida UI completa con `brand_key=""`: OK. Job `servicios_20260318_202529_a0570219`, duracion aproximada `20m31s`, 4 salidas generadas:
  - `servicios_changan_20260318_202530_c04d.xls`
  - `servicios_peug_20260318_202530_c04d.xls`
  - `servicios_szk_20260318_202530_c04d.xls`
  - `servicios_tyt_20260318_202530_c04d.xls`
- En esa corrida completa:
  - `summary`: `changan=66`, `peug=25`, `szk=274`, `tyt=296`
  - `invoice_fallbacks=0` y `note_fallbacks=0` en las 4 marcas
  - `validate_done|<marca>` y `save_done|<marca>` presentes en consola para las 4 marcas
  - Verificacion local posterior: las 4 salidas contienen `REP FACTURACION`, `NOTA DE CREDITO`, `PX`, `REP VTAS`, `COSTO` y `ESTADISTICAS`

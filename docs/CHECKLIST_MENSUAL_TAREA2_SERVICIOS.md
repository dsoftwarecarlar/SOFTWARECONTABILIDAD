# Checklist Mensual Tarea 2 Servicios por Marca

Fecha base: 2026-03-19
Alcance: solo `tarea 2`

## 1. Archivos del mes

Subir siempre los 2 archivos comunes:
- `detalle-vtas-xliquidar.xlsx`
- `RepFacturacionServContabilidad.xls`

Subir los TXT de la marca a procesar:

### CHANGAN
- `SERREP_FACTURAS_NAFCHAN.TXT`
- `SERREP_NOTACRED_NAFCHAN.TXT`
- `CON_MAYORGEN2CHAN.TXT`

### PEUGEOT
- `SERREP_FACTURAS_NAFPEU.TXT`
- `SERREP_NOTACRED_NAFPEU.TXT`
- `CON_MAYORGEN2PEU.TXT`

### SUZUKI
- `SERREP_FACTURAS_NAFSUZAMBYRIO.TXT`
- `SERREP_NOTACRED_NAFSUZAMBYRI.TXT`
- `CON_MAYORGEN2SUZ.TXT`

### MATRIZ
- `SERREP_FACTURAS_NAFTOY.TXT`
- `SERREP_NOTACRED_NAFTOY.TXT`
- `CON_MAYORGEN2TOY.TXT`

## 2. Antes de correr

- Confirmar que los archivos sean del mismo mes contable.
- Confirmar que `RepFacturacionServContabilidad.xls` y `detalle-vtas-xliquidar.xlsx` correspondan al mes a procesar.
- Confirmar que Excel no quede abierto con ventanas visibles.
- Si se procesa una sola marca, subir solo los TXT de esa marca mas los 2 Excel comunes.
- Si se procesan todas, subir los TXT de todas las marcas.

## 3. Ejecucion

- Entrar al modulo `Servicios por Marca`.
- Seleccionar la marca o dejar vacio para generar todas.
- Cargar los archivos.
- Ejecutar el proceso.
- Esperar la finalizacion completa. El proceso puede tardar varios minutos.

## 4. Validacion minima obligatoria

Revisar en el archivo generado:
- `REP FACTURACIÓN`
- `NOTA DE CREDITO`
- `PX`
- `REP VTAS`
- `VENTAS` o `MAY VTAS`
- `COSTO`
- `ESTADISTICAS`

Confirmar:
- que la marca mostrada sea correcta
- que no existan `#REF!`
- que `PX` tenga la cantidad esperada de filas o quede vacio si no aplica
- que `ESTADISTICAS` no tenga residuos viejos del template
- que el archivo se guarde en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs`

## 5. Criterio de rechazo

No usar el archivo si ocurre cualquiera de estos casos:
- aparece `#REF!`
- aparece una marca incorrecta en cabeceras
- faltan hojas esperadas
- hay datos visibles de meses anteriores
- el archivo no corresponde al mes cargado

## 6. Salida esperada

Nombres tipo:
- `servicios_changan_<timestamp>.xls`
- `servicios_peug_<timestamp>.xls`
- `servicios_szk_<timestamp>.xls`
- `servicios_tyt_<timestamp>.xls`

Ubicacion:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs`

## 7. Observacion operativa

La plantilla debe aportar solo:
- estructura
- formulas
- formato
- layout visual

Los datos operativos del mes deben salir de los archivos subidos y de los calculos derivados validos.

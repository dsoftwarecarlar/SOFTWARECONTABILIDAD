# Reporte Auditoria Espejo Tarea 2

Fecha: 2026-03-19

Alcance:
- Solo `tarea 2`
- `Servicios por Marca`
- Comparacion entre plantilla base, ultima salida real por marca y fuentes del mes

## Salidas auditadas

- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_20260319_auditfix3.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_20260319_auditfix3.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_20260319_auditfix3.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_20260319_155610_0173.xls`

## Resumen Ejecutivo

### CHANGAN
- El detalle visible de `REP FACTURACION`, `NOTA DE CREDITO`, `REP VTAS` y `PX` cuadra en conteos y totales contra las fuentes auditadas.
- La salida no esta cerrada correctamente para uso mensual.
- Hallazgos:
  - `NOTA DE CREDITO!F4`, `NOTA DE CREDITO!F5` y `PrecontabilizacionVentas!V7` contienen `#REF!`
  - `PX!H3` conserva `01/12/2023`
  - `ESTADISTICAS` perdio formulas en `I305`, `J305`, `I358`, `J358`
  - El IVA 15% de facturas termino escrito en la columna de IVA 12%

### PEUGEOT
- El detalle visible de `REP FACTURACION`, `NOTA DE CREDITO`, `REP VTAS` y `PX` cuadra en conteos y totales contra las fuentes auditadas.
- La salida no esta cerrada correctamente para uso mensual.
- Hallazgos:
  - `PrecontabilizacionVentas!V7` contiene `#REF!`
  - `PX!H3` conserva `01/12/2023`
  - `ESTADISTICAS` perdio formulas en `I305`, `J305`, `I358`, `J358`
  - El IVA 15% de facturas termino escrito en la columna de IVA 12%

### SUZUKI
- El detalle visible de `REP FACTURACION`, `NOTA DE CREDITO`, `REP VTAS` y `PX` cuadra en conteos y totales contra las fuentes auditadas.
- La salida no esta cerrada correctamente para uso mensual.
- Hallazgos:
  - `PrecontabilizacionVentas!V7` y `PrecontabilizacionVentas!V10` contienen `#REF!`
  - `PX!H3` conserva `01/12/2023`
  - `ESTADISTICAS` perdio formulas en `I305`, `J305`, `I358`, `J358`
  - El IVA 15% de facturas termino escrito en la columna de IVA 12%

### MATRIZ
- La ultima salida real es `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_20260319_155610_0173.xls`
- Esa salida fue generada con estos uploads reales:
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\SERREP_FACTURAS_NAFTOY31_20260319_155609.txt`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\SERREP_NOTACRED_NAFTOYO_20260319_155609.txt`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\RepFacturacionServContabilidad_4_20260319_155609.xls`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\detalle-vtas-xliquidar_3_20260319_155609.xlsx`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\CON_MAYORGEN2TOYO_20260319_155609.txt`
- Contra esos uploads reales:
  - `REP FACTURACION` cuadra `398/398`
  - `NOTA DE CREDITO` cuadra `7/7`
  - `REP VTAS` cuadra `371/371`
  - `PX` cuadra `6/6`
- Pero no cuadra espejo contra la plantilla base actual.

## Hallazgo Critico de MATRIZ

La plantilla base actual de `MATRIZ` y los uploads reales del ultimo job no corresponden al mismo periodo.

Evidencia:
- La plantilla base `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates\11. Concili. Servicios TYT 2026.xls` contiene:
  - `REP FACTURACION`: `291` filas
  - `NOTA DE CREDITO`: `5` filas
  - `REP VTAS`: `296` filas
  - periodo visible entre `02/02/2026` y `28/02/2026`
- Los uploads reales del ultimo job contienen:
  - facturas: `398` filas
  - notas: `7` filas
  - periodo visible entre `05/01/2026` y `31/01/2026`

Conclusion:
- La salida `servicios_tyt_20260319_155610_0173.xls` no puede cuadrar espejo contra la plantilla base actual porque la plantilla base esta armada con otro mes.

## Hallazgo Critico de Mayor en MATRIZ

Contra el upload real del mayor:
- fuente: `35` filas
- salida: `34` filas

Fila faltante en `MAY VTAS`:
- cuenta: `04.01.01.11.0002`
- descripcion: `VTAS SERV TOYOTA - CONTADO SIN IVA`
- origen: `VENSE`
- asiento: `400`
- detalle: `CONT. VENTAS - CENTRO01 PERIODO 2026 - 01`
- debito: `0.00`
- credito: `2057.85`
- saldo: `-2057.85`

Impacto:
- `REP VTAS!D6` toma un mayor incompleto
- `REP VTAS!D8` queda con diferencia visible de `-5427.83`

Causa:
- La plantilla base actual no contiene una seccion para la cuenta `04.01.01.11.0002`

## Cambios Estructurales Detectados

En las 4 marcas se detectaron cambios de ancho de columna respecto a la plantilla base en:
- `COSTO`
- `PrecontabilizacionCostos (2)`

Esto explica el sintoma reportado de que la plantilla se mueve o cambia de tamano al finalizar.

## Correcciones ya dejadas en codigo

En `C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1`:
- restauracion de layout desde la plantilla antes del guardado
- restauracion de `PageSetup`
- correccion del mapeo de IVA 12/IVA 15 en `REP FACTURACION`
- generacion de filas de soporte en `PrecontabilizacionVentas` para evitar `#REF!`
- actualizacion del ancla de periodo en `PX`
- validacion estricta para fallar si el mayor trae cuentas sin seccion equivalente en la plantilla

## Dictamen

- `CHANGAN`, `PEUGEOT` y `SUZUKI`: las ultimas salidas auditadas siguen teniendo errores funcionales y residuos estructurales
- `MATRIZ`: la ultima salida real cuadra contra los uploads reales, pero no contra la plantilla base actual porque fuente y plantilla pertenecen a periodos distintos, y ademas falta una cuenta del mayor en la estructura de la plantilla

## Siguiente Paso Correcto

1. Regenerar `CHANGAN`, `PEUGEOT` y `SUZUKI` con el worker actualizado
2. Regenerar `MATRIZ` con el worker actualizado
3. Si `MATRIZ` vuelve a fallar por la cuenta `04.01.01.11.0002`, decidir una de estas dos rutas:
   - actualizar la plantilla base para incluir esa seccion
   - exigir una plantilla mensual compatible con el periodo real antes de generar

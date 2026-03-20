# Cierre Tarea 2 - Servicios por Marca

Fecha: 2026-03-20

## Alcance
Solo `Tarea 2`.

## Estado
- Flujo mensual operativo: `CERRADO`
- Espejo absoluto contra `PrecontabilizacionCostos` historica de plantilla: `NO APLICA COMO CRITERIO OPERATIVO`

## Cambios en codigo
Archivo:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1`

Ajustes:
- Se agrego `Assert-WorksheetAreaBlank`.
- `Validate-Services-BrandOutput` ahora valida que `PrecontabilizacionCostos` quede neutralizada en:
  - `B2:J922`
  - `P1:T25`
- El log operativo cambia a `INFO|precont_costos_legacy_neutralized|<marca>`.

## Evidencia tecnica
### 1. La hoja `PrecontabilizacionCostos` no alimenta el cierre critico
Busqueda real de formulas en plantilla `TYT`:
- `REP VTAS!N2`
- `REP VTAS!N3`
- `REP VTAS!N4`
- `REP VTAS!N5`

Todas apuntan a:
- `PrecontabilizacionCostos (2)`

No se encontro una dependencia equivalente hacia:
- `PrecontabilizacionCostos`

Conclusion:
- la hoja critica para costos del cierre es `PrecontabilizacionCostos (2)`
- `PrecontabilizacionCostos` es un bloque legacy de plantilla, no la fuente activa del calculo auditado

### 2. La plantilla legacy no coincide con la fuente mensual real
En plantilla `TYT`, `PrecontabilizacionCostos` contiene agencias legacy:
- `03, 07, 12, 14, 15, 18, 19, 20, 23, 38`

Y mezcla lineas:
- `30`
- `90`

En la salida final auditada `TYT`:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls`

La hoja `REP VTAS` solo contiene:
- centros: `01`
- linea: `30`

En las otras salidas finales:
- `servicios_changan_finalaudit2_changan_20260319.xls` -> centros `05,08`, linea `24`
- `servicios_peug_finalaudit2_peug_20260319.xls` -> centro `06`, linea `31`
- `servicios_szk_finalaudit2_szk_20260319.xls` -> centros `07,08`, linea `33`

Conclusion:
- `PrecontabilizacionCostos` historica de plantilla no representa el set mensual real de cada marca
- forzar un espejo literal de esa hoja implicaria conservar o inventar datos no respaldados por los archivos subidos

### 3. Las salidas finales ya quedan limpias en esa hoja legacy
Verificacion real sobre los ultimos generados:
- `servicios_changan_finalaudit2_changan_20260319.xls` -> `main_blank=True`, `pivot_blank=True`
- `servicios_peug_finalaudit2_peug_20260319.xls` -> `main_blank=True`, `pivot_blank=True`
- `servicios_szk_finalaudit2_szk_20260319.xls` -> `main_blank=True`, `pivot_blank=True`
- `servicios_tyt_finalaudit_tyt_20260319.xls` -> `main_blank=True`, `pivot_blank=True`

Rangos verificados:
- `B2:J922`
- `P1:T25`

## Criterio final correcto para uso mensual
La regla valida para `Tarea 2` queda asi:
- conservar layout, formulas, formatos y estructura de plantilla
- cargar solo datos respaldados por archivos del mes
- neutralizar bloques legacy no respaldados por la fuente mensual
- fallar si reaparecen residuos historicos en `PrecontabilizacionCostos`

## Dictamen
- `REP FACTURACION`: OK
- `NOTA DE CREDITO`: OK
- `PX`: OK
- `REP VTAS`: OK
- `VENTAS / MAY VTAS`: OK
- `PrecontabilizacionVentas`: OK
- `PrecontabilizacionCostos (2)`: OK
- `COSTO`: OK
- `ESTADISTICAS`: OK
- `PrecontabilizacionCostos`: neutralizada y validada como hoja legacy no operativa

## Cierre
Para el proceso mensual de contabilidad, `Tarea 2` queda cerrada con este criterio profesional:
- no se arrastra dato operativo historico de plantilla
- no se inventan asientos sin fuente
- la hoja legacy `PrecontabilizacionCostos` queda controlada y validada en blanco

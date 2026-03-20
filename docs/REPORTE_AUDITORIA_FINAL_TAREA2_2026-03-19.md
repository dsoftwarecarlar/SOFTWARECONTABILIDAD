# Reporte Final Tarea 2 - Servicios por Marca

Fecha: 2026-03-19

## Salidas auditadas
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_finalaudit2_peug_20260319.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_finalaudit2_szk_20260319.xls`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_finalaudit_tyt_20260319.xls`

## Resumen ejecutivo
- CHANGAN: `OBSERVACION`
- PEUGEOT: `OBSERVACION`
- SUZUKI: `OBSERVACION`
- MATRIZ: `OBSERVACION`

Conclusión:
- La estructura final del libro ya quedó alineada con la plantilla.
- Las zonas de control críticas ya no presentan `#REF!` ni `#¡REF!`.
- La generación ya valida y guarda correctamente las 4 marcas.
- La única brecha abierta para auditoría espejo estricta es `PrecontabilizacionCostos`, que permanece limpio pero no reconstruido.

## Evidencia estructural
- Orden de hojas: `OK` en las 4 marcas.
- Anchos de columna: `OK` en las 4 marcas.
- Altos de fila: `OK` en las 4 marcas.
- `PageSetup`: `OK` en las 4 marcas.
- Hojas críticas sin errores Excel visibles: `OK` en las 4 marcas.

## Evidencia funcional
- Logs finales por marca con `validate_done` y `save_done`: `OK`.
- `invoice_fallbacks=0` y `note_fallbacks=0` en las 4 marcas.
- Fórmulas de control corregidas:
  - CHANGAN: `PrecontabilizacionVentas!V7=0`, `V10=713,80`
  - PEUGEOT: `PrecontabilizacionVentas!V7=0`, `V10=0`
  - SUZUKI: `PrecontabilizacionVentas!V7=0`, `V10=619,79`
  - MATRIZ: `PrecontabilizacionVentas!V7=0`, `V10=2777,19`
- `NOTA DE CREDITO!F4/G4` sin errores en las 4 marcas.
- `REP FACTURACIÓN!D9/E9` sin errores en las 4 marcas.

## Hallazgo pendiente
Archivo:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1`

Hoja:
- `PrecontabilizacionCostos`

Estado:
- La hoja conserva estructura y layout, pero queda vacía salvo encabezados.

Impacto:
- No rompe las fórmulas críticas actualmente auditadas.
- Sí impide declarar cierre `100% espejo` contra la plantilla manual si se exige igualdad visual y de contenido en esa hoja.

Clasificación:
- `OBSERVACION` para uso operativo mensual.
- `ERROR` para auditoría espejo estricta celda por celda.

## Observación crítica de fuentes
Los TXT actuales del repo contienen campos anonimizados en `C.I.` y `Cliente`.

Ejemplo real:
- Fuente: `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures\archivosasubirchan\SERREP_FACTURAS_NAFCHAN.TXT`
- Salida: `REP FACTURACIÓN!F17/G17` en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_finalaudit2_changan_20260319.xls`

Esto implica:
- salida y fuente sí cuadran entre sí
- pero no siempre cuadran literal contra una plantilla manual no anonimizada en esas columnas

## Dictamen final
- Para operación mensual contable: `APTO CON OBSERVACION`
- Para espejo perfecto contra la plantilla manual en todas las hojas: `AUN NO CERRADO`

Bloque exacto pendiente:
- reconstrucción de `PrecontabilizacionCostos`

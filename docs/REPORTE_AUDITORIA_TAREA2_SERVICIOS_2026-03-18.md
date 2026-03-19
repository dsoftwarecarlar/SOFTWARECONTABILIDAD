# Reporte Auditoria Tarea 2 Servicios por Marca

Fecha: 2026-03-19

Alcance:
- Solo `tarea 2`
- Modulo `Servicios por Marca`
- Plantillas base en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates`
- Fuentes del mes en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures`
- Salidas auditadas:
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_changan_20260319_auditfix3.xls`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_peug_20260319_auditfix3.xls`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_szk_20260319_auditfix3.xls`
  - `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs\servicios_tyt_20260319_auditfix3.xls`

## 1. Resumen Ejecutivo

| Marca | Estado | Conclusion tecnica |
| --- | --- | --- |
| CHANGAN | OK | El libro conserva estructura y formulas de plantilla, sin `#REF!` en los bloques criticos auditados y sin residuos historicos en `ESTADISTICAS`. |
| PEUGEOT | OK | La salida respeta plantilla, deja `PX` limpio cuando la fuente viene vacia y mantiene formulas y encabezados correctos. |
| SUZUKI | OK | La carga respeta plantilla, formulas auxiliares quedan sanas y `PX` ya sale con cabecera correcta de marca. |
| MATRIZ | OK | `MAY VTAS`, `PX`, `REP VTAS` y auxiliares quedan consistentes; no hay `#REF!` en las zonas criticas auditadas. |

Dictamen ejecutivo:
- Estado final de `tarea 2`: `OK`
- La salida actual puede usarse como proceso mensual del area contable dentro del alcance auditado.
- La plantilla queda como molde: estructura, formulas, layout y textos fijos.
- Los datos operativos auditados salen de los archivos subidos o de calculos derivados validos.

## 2. Correcciones Cerradas

Se corrigio en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\run_servicios_marcas.ps1`:
- Preservacion de cuentas con cero inicial en `PrecontabilizacionVentas` y `PrecontabilizacionCostos (2)`.
- Carga de `PX` por bloques reales de plantilla, sin borrar subtotales ni cabeceras.
- Escritura de cabeceras `PX` por marca y timestamp del proceso.
- Limpieza completa del bloque historico de `ESTADISTICAS` (`446+`).
- Regeneracion de filas ancla de `ESTADISTICAS` tomando la plantilla real como base.
- Filtro correcto de filas detalle `PX` para no arrastrar encabezados.

Tambien se corrigio el contrato de prueba en:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\scripts\tests\e2e_servicios_marcas_contract.js`

Motivo:
- la prueba E2E de `PX` estaba inyectando el canario en una fila de encabezado del archivo fuente, no en una fila real de detalle.

## 3. Evidencia Tecnica

Validacion critica ejecutada sobre las 4 salidas auditadas:
- `REP FACTURACION`: sin `#REF!` en `D9:E13`
- `NOTA DE CREDITO`: sin `#REF!` en `F4:G5`
- `REP VTAS`: sin `#REF!` en `N2:P7`
- `PrecontabilizacionVentas`: sin `#REF!` en `V7:V10`
- `PrecontabilizacionVentas`: cuentas auditadas con cero inicial correcto
- `PrecontabilizacionCostos (2)`: cuentas auditadas con cero inicial correcto
- `ESTADISTICAS`: filas `446:450` limpias, sin datos historicos residuales
- `ESTADISTICAS`: formulas de subtotal presentes en `I129`, `I220`, `I288`, `I363`, `I444`
- `PX`: cabecera `D5` correcta por marca
- `PX`: formulas de subtotal/restaura en filas de control

Conteo `PX` auditado contra fuente comun filtrada por marca:
- CHANGAN: `1/1`
- PEUGEOT: `0/0`
- SUZUKI: `3/3`
- MATRIZ: `9/9`

Resultados visibles auditados en `PX`:
- `CHANGAN`: `D3=19/03/2026 00.48.56`, `D5=CHANGAN`
- `PEUGEOT`: `D3=19/03/2026 00.50.19`, `D5=PEUGEOT`
- `SUZUKI`: `D3=19/03/2026 00.54.07`, `D5=SUZUKI`
- `MATRIZ`: `D3=19/03/2026 01.00.27`, `D5=TOYOTA`

Resultados visibles auditados en `ESTADISTICAS`:
- Anclas restauradas:
  - `R6` -> `05.02.01.01.0001`
  - `R131` -> `05.02.01.01.0002`
  - `R222` -> `05.02.01.01.0003`
  - `R290` -> `05.02.01.01.0004`
  - `R365` -> `05.02.01.01.0008`
- Todas con fecha del periodo actual `28/2/2026`

## 4. Validaciones Ejecutadas

Validaciones aprobadas:
- Corrida completa manual de las 4 marcas con fixtures reales del mes.
- Auditoria COM dirigida sobre celdas y bloques criticos de las 4 salidas `auditfix3`.
- `npm run test:e2e:servicios` -> `OK` el `2026-03-19`.

Nota tecnica no bloqueante:
- El worker emite `WARN|recalc_full_failed|FullCalculationOnLoad` porque esa propiedad no esta disponible en la version local de Excel COM.
- No bloqueo la validacion, el guardado ni los resultados auditados.

## 5. Dictamen Final

Dictamen:
- `Si`, el libro generado puede considerarse confiable para uso operativo mensual dentro del alcance auditado de `tarea 2`.
- `No` quedaron `#REF!` en las zonas criticas auditadas.
- `No` quedaron residuos historicos en `ESTADISTICAS` dentro del bloque problematico auditado.
- `Si` se conserva la plantilla como estructura base y las formulas de control relevantes.

Resultado final de auditoria:
- Estado final de `tarea 2` en esta corrida: `OK`

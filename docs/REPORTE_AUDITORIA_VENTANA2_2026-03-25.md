## Ventana 2 - Auditoria estricta

Fecha: 25 de marzo de 2026

### Alcance
- comparacion de `Conciliacion Servicios por Marca` contra la plantilla real de Toyota
- revision de:
  - `REP FACTURACION`
  - `NOTA DE CREDITO`
  - `PX`
  - `REP VTAS`
  - `PrecontabilizacionVentas`
  - `PrecontabilizacionCostos (2)`
  - `COSTO`
  - `ESTADISTICAS`
  - `MAY VTAS`

### Hallazgos
1. La salida moderna con fixtures actuales sigue pasando contrato, rendimiento y gate general.
2. En comparaciones `manualcompare` contra la plantilla real, muchas diferencias en `REP FACTURACION`, `NOTA DE CREDITO` y `REP VTAS` vienen de datos anonimizados de fixture y no de un defecto de Excel.
3. La deuda funcional real sigue concentrada en:
   - `PrecontabilizacionCostos (2)`
   - `COSTO`
   - `ESTADISTICAS`
4. El runtime Python actual de Ventana 2 sigue escribiendo un bloque contable de costos simplificado y no la estructura completa de la plantilla manual.
5. El archivo `FacturacionServContabilidadDetallado` usa un layout historico distinto del layout moderno `RepFacturacionServContabilidad`.

### Correcciones aplicadas en esta auditoria
- Se dejo el prompt estricto en:
  - `docs/PROMPT_AUDITORIA_VENTANA2_PLANTILLA_REAL_2026-03-25.md`
- Se endurecio la deteccion del layout legacy en:
  - `python_services/processors/servicios_marcas/readers.py`
- Se agrego un guardrail dedicado para rechazar `FacturacionServContabilidadDetallado` y aceptar solo el layout moderno:
  - `scripts/tests/servicios_layout_guard.js`
- Se corrigio la escritura de `COSTO` para no borrar las filas semilla antes de leerlas:
  - `python_services/processors/servicios_marcas/runtime.py`
- Se endurecio el contrato E2E de Ventana 2 para revisar semillas contables de:
  - `PrecontabilizacionCostos (2)`
  - `COSTO`
  - `ESTADISTICAS`
  - `scripts/tests/e2e_servicios_marcas_contract.js`
- Se integro el guardrail de layout y el guardrail de rendimiento al paquete principal:
  - `scripts/tests/full_operational_monthly_audit.js`
  - `scripts/tests/definitive_total_contable_audit.js`
  - `package.json`
- Se revalido el gate general despues de cerrar una regresion lateral de `Repuestos`:
  - `python_services/processors/repuestos_tytserv/rep_stage.py`

### Estado despues de correcciones
- `npm run test:e2e:servicios` -> OK
- `node scripts/tests/servicios_layout_guard.js` -> OK
- `node scripts/tests/servicios_performance_guard.js` -> OK
- `npm run test:audit:mensual` -> OK
- `npm run test:audit:definitiva` -> OK
- `node scripts/tests/e2e_repuestos_contract.js` -> OK
- `node scripts/tests/contable_quality_gate.js` -> OK

### Diferencias reales que siguen pendientes
1. `PrecontabilizacionCostos (2)` sigue generandose como bloque resumido de costos, no como espejo completo de todas las filas historicas de la plantilla manual.
2. `ESTADISTICAS` depende de ese mismo bloque resumido y por eso no puede ser identica a la plantilla historica cuando se la compara como si fuera un libro manual congelado.
3. El layout legacy `FacturacionServContabilidadDetallado` ya se detecta con mensaje claro, pero todavia no tiene mapeo contable automatico completo dentro de la ruta moderna.

### Conclusion tecnica
- Ventana 2 sigue operativa en la ruta moderna validada por contratos.
- La comparacion "plantilla manual exacta" sigue teniendo una brecha real mientras no se porte el bloque contable detallado de costos.
- El layout legacy `FacturacionServContabilidadDetallado` ya no entra por error silencioso: ahora queda bloqueado con mensaje claro y con prueba automatica dedicada.
- La salida moderna ya quedo mas consistente en `COSTO`, y ahora tambien queda protegida por verificacion automatica de semillas y por auditorias mensuales/definitivas.

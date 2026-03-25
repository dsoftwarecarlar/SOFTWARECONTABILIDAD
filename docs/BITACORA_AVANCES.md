# BITACORA DE AVANCES

Objetivo: dejar trazabilidad ejecutiva y tecnica de los avances reales del sistema `CXP`, de forma que jefatura pueda revisar que se construyo, que se valido y que sigue pendiente.

## 1. Como usar esta bitacora

Registrar una fila por avance real del sistema.

No registrar solo:

- reuniones
- revisiones
- documentacion creada

Registrar:

- ventanas creadas
- modulos implementados
- cambios funcionales
- pruebas ejecutadas
- cierres parciales o finales
- pendientes reales

Campos minimos:

- `Fecha`
- `Frente`
- `Cambio aplicado`
- `Estado`
- `% avance`
- `Evidencia`
- `Riesgo o bloqueo`
- `Siguiente paso`
- `Responsable`

Estados sugeridos:

- `Pendiente`
- `En progreso`
- `Validado`
- `Bloqueado`

## 2. Registro actual

| Fecha | Frente | Cambio aplicado | Estado | % avance | Evidencia | Riesgo o bloqueo | Siguiente paso | Responsable |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 2026-03-13 | Ventana 1 / `Accion 4 - Mayor IVA` | Se implemento la generacion y validacion estructural del workbook de `Accion 4`. | Validado | 100% | commit `59825d4`, `modules/cxp_accion4/index.php`, `scripts/cxp/accion4/process.js` | Ninguno visible | Mantenerla dentro de las pruebas de regresion del flujo ACLT. | Equipo |
| 2026-03-14 | Calidad contable y pruebas | Se endurecieron controles de integridad y se agregaron pruebas E2E y smoke para frentes criticos. | Validado | 100% | `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `package.json`, `scripts/tests/e2e_repuestos_contract.js`, `scripts/tests/http_resources_smoke.js`, `scripts/tests/contable_quality_gate.js` | Mantener ejecucion del quality gate antes de despliegue. | Ejecutar gate y consolidar evidencia de `Accion 1-4`. | Equipo |
| 2026-03-18 | Ventana 2 / `Servicios por Marca` | Se rediseno la interfaz por marca, el backend valida uploads por marca, el runner dejo de depender de archivos legacy y la corrida completa ya genera `4` salidas. | Validado | 90% | `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`, `modules/cxp_servicios_marcas/index.php`, `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php` | Pendiente cierre formal de la brecha espejo en `PrecontabilizacionCostos`. | Ejecutar auditoria final y cierre operativo. | Equipo |
| 2026-03-18 | Ventana 3 / `Repuestos TYTSERV` | Se confirmo `Node.js` como runtime productivo del modulo web y el flujo mensual de `8` archivos. | Validado | 85% | `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md`, `config/cxp/repuestos_tytserv.php`, `modules/cxp_repuestos_tytserv/index.php` | Pendiente cierre funcional mensual con evidencia ejecutiva final. | Mantener `E2E` y validar corrida mensual final. | Equipo |
| 2026-03-19 | Ventana 2 / `Servicios por Marca` | Auditoria final de las `4` marcas: estructura alineada con plantilla, sin `#REF!`, apto para operacion mensual con observacion. | Validado | 95% | `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md` | Queda abierta solo la brecha de `PrecontabilizacionCostos` para espejo estricto. | Cerrar criterio final de uso mensual. | Equipo |
| 2026-03-20 | Ventana 2 / cierre operativo | `Servicios por Marca` queda cerrado para uso mensual: se neutralizo y valido en blanco la hoja legacy `PrecontabilizacionCostos`. | Validado | 100% | `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`, `storage/verify_runs/worker_all.log`, `storage/verify_runs/servicios_changan_verifyall_20260320.xls`, `storage/verify_runs/servicios_peug_verifyall_20260320.xls`, `storage/verify_runs/servicios_szk_verifyall_20260320.xls`, `storage/verify_runs/servicios_tyt_verifyall_20260320.xls` | Solo reabrir si se exige auditoria espejo exacta de la hoja legacy. | Mantener checklist mensual y monitoreo operativo. | Equipo |
| 2026-03-23 | Interfaz general `CXP` | Se consolido la comunicacion del sistema con `3` ventanas activas y se alineo la descripcion real de `Repuestos TYTSERV` como flujo de `8` archivos. | Validado | 100% | `areas/cxp/index.php`, `includes/app.php` | Ninguno | Usar esta definicion como base del planner gerencial. | Equipo |

## 3. Plantilla para nuevas entradas

| Fecha | Frente | Cambio aplicado | Estado | % avance | Evidencia | Riesgo o bloqueo | Siguiente paso | Responsable |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| AAAA-MM-DD | Modulo o frente | Describir el cambio real aplicado | Pendiente / En progreso / Validado / Bloqueado | 0-100% | commit, prueba, ruta, reporte o archivo generado | indicar riesgo o `Ninguno` | siguiente accion concreta | nombre |

## 4. Regla de gobierno

Cada avance importante debe quedar trazado en tres niveles:

1. planner gerencial
2. bitacora de avances
3. evidencia tecnica o funcional

Si falta cualquiera de esos tres niveles, el avance se vuelve dificil de explicar a jefatura y dificil de sostener en soporte.

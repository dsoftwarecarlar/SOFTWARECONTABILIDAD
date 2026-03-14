# Reporte de auditoria contable - 2026-03-14

## Alcance
- Tarea 2: rendimiento, fidelidad de plantilla y descarga del archivo correcto.
- Tarea 3: validacion contable por cuenta, precision de decimales y fidelidad de plantilla.
- Repuestos: integridad de payload por hoja, robustez de guardado y descarga del artefacto real.
- UI: carga de vistas y recursos estaticos.

## Hallazgos criticos
1. Tarea 2 tenia sobrecosto evitable por operaciones de estilo redundantes en bucles grandes.
2. Tarea 3 validaba diferencias contables con tolerancia de 0.05 (riesgo financiero).
3. Repuestos podia fallar por bloqueo de archivo Excel y dejar resultados engañosos.
4. Faltaba una prueba E2E contractual para repuestos que evitara regresiones.

## Correcciones aplicadas

### Tarea 2 (rendimiento + integridad)
- Se elimino reaplicacion masiva de estilos en toda la hoja durante build.
- Se mantuvo formato espejo de plantilla con merge XML y validacion de estilo/cabecera/anchos.
- Se migro el resumen lateral a aritmetica en centavos (enteros) para evitar drift por flotantes.
- Se agrego cache de firma de estilos para acelerar validacion final.

Archivos:
- `scripts/cxp/accion2/workbook.js`
- `scripts/cxp/accion2/process.js`

### Tarea 3 (contabilidad + integridad)
- Validacion de totales por cuenta migrada a centavos.
- Umbral endurecido: diferencia > 1 centavo se considera fallo.
- Mantiene validacion de merge XML (rows/payload/hash) y fidelidad visual.
- Se agrego cache de firmas de estilo en verificacion final.

Archivos:
- `scripts/cxp/accion3/parser.js`
- `scripts/cxp/accion3/workbook.js`
- `scripts/cxp/accion3/process.js`

### Repuestos (robustez operativa)
- Guardado robusto con `SaveCopyAs`, artefactos temporales `__working_` y `__saved_`, y publicacion atomica.
- Validacion hash/filas del payload REP para impedir salida con datos de plantilla.
- Limpieza de artefactos en fallo para no exponer archivo incorrecto.

Archivo:
- `run_repuestos_tytserv.ps1`

### Calidad y pruebas
- Se agrego contrato E2E para repuestos.
- Se agrego smoke de vistas/recursos.
- Se agrego quality gate unico para ejecutar smoke + E2E accion2/3 + E2E repuestos.

Archivos:
- `scripts/tests/e2e_repuestos_contract.js`
- `scripts/tests/http_resources_smoke.js`
- `scripts/tests/contable_quality_gate.js`
- `package.json`

## Evidencia de ejecucion

### Pruebas
- `npm run test:e2e:accion2-3` => OK
- `npm run test:e2e:repuestos` => OK
- `npm run test:smoke:ui` => OK

### Rendimiento observado (ultimo run)
- Tarea 2 (491 filas): total ~1876 ms
  - parse=288, build=321, write=231, merge=631, verify=405
- Tarea 3 (328 movimientos): total ~1487 ms
  - parse=18, validate=1, build=321, write=228, merge=508, verify=410
- Tarea 2 stress (4910 filas): total ~8680 ms

## Riesgo residual
- Repuestos sigue dependiendo de Excel COM en Windows; ya esta endurecido, pero puede degradarse si hay carga alta del host o bloqueo externo de archivos.
- Se recomienda mantener el quality gate antes de cada despliegue.

## Estado final
- Apto para despliegue en area contable con controles de integridad y exactitud reforzados.

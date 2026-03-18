# Checklist de despliegue contable

## 1) Pre-deploy

- Confirmar backup de:
  - `storage/outputs`
  - `storage/jobs`
  - `storage/uploads` si hay procesos pendientes
- Confirmar backup del entorno objetivo si existe base de datos asociada.
- Verificar version de:
  - PHP
  - Node
  - Excel/Office en el host que ejecuta `Servicios por Marca`
- Confirmar que existen y son legibles estas rutas base:
  - `resources/cxp/acciones`
  - `resources/cxp/servicios_marcas`
  - `resources/cxp/repuestos_tytserv`
- Confirmar que no existan jobs activos de `Servicios por Marca`:
  - estados a revisar: `queued`, `running`, `cancel_requested`
- Confirmar que no exista un `EXCEL.EXE` visible abierto si se va a validar `Servicios por Marca`.
- Ejecutar quality gate rapido:
  - `npm run test:quality:contable`
- Ejecutar contrato E2E de `Servicios por Marca` cuando el despliegue toque ese modulo o el host Windows/Excel:
  - `npm run test:e2e:servicios`

## 2) Arquitectura que debe quedar valida

- `Libro Compras ACLT`:
  - runtime web PHP
  - plantillas y fixtures bajo `resources/cxp/acciones`
- `Repuestos TYTSERV`:
  - runtime productivo web Node
  - configuracion activa en `config/cxp/repuestos_tytserv.php`
  - worker actual en `scripts/cxp/repuestos_tytserv/process.js`
  - `run_repuestos_tytserv.ps1` solo fallback/manual probe
- `Servicios por Marca`:
  - worker en `scripts/cxp/servicios_marcas/run.ps1`
  - procesamiento heredado con Excel COM en Windows
  - plantillas y fixtures bajo `resources/cxp/servicios_marcas`
- Las plantillas base no deben salir desde `outputs/`.
- Las salidas reales deben quedar en `storage/outputs`.

## 3) Despliegue

- Publicar codigo.
- Limpiar cache de OPCache/PHP-FPM si aplica.
- Verificar permisos de escritura:
  - `storage/uploads`
  - `storage/outputs`
  - `storage/jobs`
- Verificar que el proceso de despliegue no haya recreado ejemplos manuales en raiz u `outputs/`.
- Confirmar que `archive/cxp_manual_outputs` siga como archivo historico y no como fuente operativa.

## 4) Smoke post-deploy obligatorio

- Abrir:
  - `areas/cxp/index.php`
  - `areas/cxp/libro-compras-aclt.php`
  - `areas/cxp/conciliacion-servicios-marcas.php`
  - `areas/cxp/facturacion-repuestos-tytserv.php`
- Validar `Libro Compras ACLT`:
  - Accion 1
  - Accion 2
  - Accion 3
  - Accion 4
  - `export_all_actions.php`
- Validar `Repuestos TYTSERV` cargando los 4 archivos del mes.
- Validar `Servicios por Marca` si el despliegue toca ese modulo o el host Excel/COM.
- Confirmar en todos los casos:
  - el nombre descargado coincide con el mensaje de exito
  - el contenido viene del archivo subido, no de la plantilla base
  - el historial muestra solo salidas generadas
  - los enlaces de descarga salen de `storage/outputs`

## 5) Validacion contable post-deploy

- Revisar auditorias JSON de salida cuando el modulo las genere:
  - `timings_ms`
  - `merge_integrity_ok=true`
  - `verificacion_final_ok=true`
- Revisar que no haya diferencias por cuenta en Tarea 3.
- Validar al menos 1 caso de centavos critico:
  - debe
  - haber
  - saldo
- En `Servicios por Marca`, confirmar que el worker no falle por bloqueo de Excel visible.

## 6) Criterio de rollback

- Hacer rollback inmediato si ocurre cualquiera:
  - descarga de plantilla en lugar de payload real
  - mismatch de hash/rows post-merge
  - error contable mayor a `0.01`
  - `Repuestos TYTSERV` deja de generar salida con el runtime Node activo
  - `Servicios por Marca` queda bloqueado por COM, jobs colgados o error repetitivo de Excel visible

## 7) Monitoreo inicial (primeras 24h)

- Medir tiempos por tarea desde `timings_ms`.
- Revisar errores en consola y logs de jobs.
- Confirmar ausencia de procesos colgados.
- Confirmar limpieza de staging y temporales:
  - `storage/jobs`
  - `storage/uploads`
  - `outputs/__staging_repuestos_tytserv`
- Confirmar que no reaparezcan ejemplos manuales en raiz ni en `outputs/`.

## 8) Referencias

- `docs/NOTA_RECURSOS_CXP_2026-03-18.md`
- `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md`
- `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`

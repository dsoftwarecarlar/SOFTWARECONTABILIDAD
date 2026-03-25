# Checklist de despliegue contable

## 1. Pre-deploy

- Confirmar backup de:
  - `storage/outputs`
  - `storage/jobs`
  - `storage/uploads` si hay procesos pendientes
- Confirmar que el host mantenga:
  - `PHP 8.2+`
  - `Python`
  - `Node.js` mientras exista fallback controlado
  - `Excel Desktop` para `Servicios por Marca`
- Confirmar que existen y son legibles:
  - `resources/cxp/acciones`
  - `resources/cxp/servicios_marcas`
  - `resources/cxp/repuestos_tytserv`
- Confirmar que no existan jobs activos de `Servicios por Marca`
- Confirmar que no haya un `EXCEL.EXE` visible abierto si se va a validar ese modulo

## 2. Arquitectura esperada

- superficie web principal: `Laravel`
- procesamiento principal: `Python`
- `Libro Compras ACLT`: `Laravel -> Python`
- `Repuestos TYTSERV`: `Laravel -> Python`
- `Servicios por Marca`: `Laravel -> Python -> Excel COM`
- la estructura PHP anterior solo debe existir archivada en `archive/legacy_php_surface_2026-03-24`

## 3. Validacion tecnica previa al corte

- `node scripts/tests/http_resources_smoke.js`
- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios` si el despliegue toca `Servicios por Marca`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_window1.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_repuestos.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_servicios.ps1` si aplica

## 4. Smoke post-deploy

Abrir y validar:

- `/`
- `/cxp`
- `/cxp/windows/libro-compras-aclt`
- `/cxp/windows/conciliacion-servicios-marcas`
- `/cxp/windows/facturacion-repuestos-tytserv`

Validar dentro del portal:

- `Accion 1`
- `Accion 2`
- `Accion 3`
- `Accion 4`
- `Consolidado`
- `Repuestos TYTSERV`
- `Servicios por Marca` si el despliegue lo afecta

Confirmar siempre:

- el archivo descargado coincide con el mensaje de exito
- el contenido viene del payload subido y no de una plantilla vacia
- el historial muestra solo salidas reales
- no aparecen enlaces a `areas/`, `modules/`, `.php` o `laravel_app/public` en la navegacion visible

## 5. Limpieza operativa

- `php maintenance_cleanup.php`

Verificar despues:

- `storage/outputs`
- `storage/uploads`
- `storage/jobs`
- ausencia de staging temporal innecesario

## 6. Rollback

Hacer rollback si ocurre cualquiera:

- mismatch funcional en `Accion 1..4`
- consolidado incompleto
- `Repuestos TYTSERV` deja de generar la salida mensual con Python activo
- `Servicios por Marca` se queda colgado o falla repetidamente por Excel COM
- reaparecen rutas viejas como entrada primaria del portal

## 7. Referencias

- `docs/ARQUITECTURA_CXP.md`
- `docs/MIGRACION_LARAVEL_PYTHON_2026-03-24.md`
- `docs/SUPERFICIE_LEGACY_WEB_2026-03-24.md`

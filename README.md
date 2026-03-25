# SOFTWARECONTABILIDAD

Portal operativo interno para procesos contables de `Contabilidad Talleres`.

Desde el **24 de marzo de 2026**, la superficie activa del sistema ya queda centrada en:

- `Laravel` para navegacion, formularios, descargas, historial y orquestacion
- `Python` para el procesamiento principal de `Libro Compras ACLT` y `Repuestos TYTSERV`
- `PowerShell + Excel COM` solo para el worker sensible de `Servicios por Marca`

La estructura PHP anterior ya no es la entrada principal. La superficie web legacy fue archivada en:

- `archive/legacy_php_surface_2026-03-24`

## 1. Entrada operativa

Rutas publicas que deben usarse:

- `/`
- `/cxp`
- `/cxp/windows/libro-compras-aclt`
- `/cxp/windows/conciliacion-servicios-marcas`
- `/cxp/windows/facturacion-repuestos-tytserv`
- `/cxp/modules/accion1`
- `/cxp/modules/accion2`
- `/cxp/modules/accion3`
- `/cxp/modules/accion4`
- `/cxp/modules/consolidado-acciones`
- `/cxp/modules/servicios-marcas`
- `/cxp/modules/repuestos-tytserv`
- `/downloads/{archivo}`

En Apache/XAMPP, la raiz del proyecto actua como front controller y enruta hacia Laravel. Las rutas antiguas de `areas/`, `modules/`, `download.php` y `export_all_actions.php` ya no son la superficie principal.

## 2. Estado funcional actual

### Ventana 1: Libro Compras ACLT

- `Accion 1` -> `Laravel -> Python nativo`
- `Accion 2` -> `Laravel -> Python nativo`
- `Accion 3` -> `Laravel -> Python nativo`
- `Accion 4` -> `Laravel -> Python nativo`
- `Consolidado` -> `Laravel -> Python nativo`

### Ventana 2: Conciliacion Servicios por Marca

- capa web -> `Laravel`
- validacion y preflight -> `Python`
- arranque del job -> `Laravel`
- worker final -> `PowerShell + Excel COM`

Es la ultima excepcion tecnica importante. No depende ya de la vieja superficie web PHP, pero si del worker heredado de Excel.

### Ventana 3: Facturacion Repuestos TYTSERV

- `Laravel -> Python nativo`
- `Node.js` queda solo como referencia/fallback controlado de paridad

## 3. Estructura activa del repositorio

| Ruta | Responsabilidad |
| --- | --- |
| `index.php` | front controller raiz hacia Laravel |
| `.htaccess` | rutas limpias, redirecciones legacy y proteccion de directorios privados |
| `laravel_app/` | aplicacion web principal |
| `python_services/` | procesadores y bridge Python |
| `config/cxp/` | configuracion funcional compartida |
| `resources/cxp/` | plantillas, fixtures y contratos |
| `storage/uploads/` | uploads temporales |
| `storage/outputs/` | salidas finales descargables |
| `storage/jobs/` | estado y snapshots de Servicios por Marca |
| `archive/legacy_php_surface_2026-03-24/` | superficie PHP archivada |
| `docs/` | documentacion tecnica y operativa |

## 4. Requisitos operativos

- Windows con `Apache` o host equivalente para las rutas publicas
- `PHP 8.2+` para la capa Laravel
- `Composer` para `laravel_app`
- `Python` disponible para `python_services`
- `Node.js` disponible mientras se mantenga fallback de referencia
- `Microsoft Excel Desktop` en el host de `Servicios por Marca`

## 5. Pruebas clave

- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios`
- `node scripts/tests/http_resources_smoke.js`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_window1.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_repuestos.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_servicios.ps1`

## 6. Limpieza operativa

La limpieza CLI ya no depende del helper legacy global.

- `php maintenance_cleanup.php`

Este script conserva:

- `3` salidas por familia de proceso
- `20` uploads recientes

Tambien elimina staging temporal bajo `storage/outputs`.

## 7. Estado de la migracion

La migracion a `Laravel + Python` queda funcionalmente cerrada para la operacion diaria. Lo que sigue fuera de Python no es la vieja web PHP, sino el worker sensible de:

- `scripts/cxp/servicios_marcas/run.ps1`

Ese flujo permanece asi por estabilidad operativa con `Excel COM`.

## 8. Documentacion relacionada

- `docs/ARQUITECTURA_CXP.md`
- `docs/MIGRACION_LARAVEL_PYTHON_2026-03-24.md`
- `docs/SUPERFICIE_LEGACY_WEB_2026-03-24.md`
- `docs/CHECKLIST_DEPLOY_CONTABLE.md`
- `docs/OPERACION_CXP.md`
- `docs/MODULOS_CXP.md`

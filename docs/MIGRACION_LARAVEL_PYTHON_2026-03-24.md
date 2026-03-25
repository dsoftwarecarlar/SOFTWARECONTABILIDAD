# Migracion Laravel + Python - 2026-03-24

## Estado actual

La migracion funcional del sistema ya queda en una etapa muy avanzada y usable:

- `Laravel` gobierna la capa web principal
- `Python` gobierna el procesamiento principal de `Libro Compras ACLT` y `Repuestos TYTSERV`
- la superficie PHP anterior fue archivada
- `Servicios por Marca` ya corre con worker `Python + Excel COM` en Windows
- el portal publico ya queda endurecido con `APP_DEBUG=false` y `database/database.sqlite` para no exponer errores tecnicos por sesiones

## Superficie activa

- `index.php` en raiz ya carga Laravel
- `.htaccess` deja rutas limpias y redirecciona rutas antiguas
- `laravel_app/` es la aplicacion principal
- `python_services/` es la capa de procesamiento principal

## Lo que ya esta migrado

### Libro Compras ACLT

- `Accion 1` -> `Laravel -> Python nativo`
- `Accion 2` -> `Laravel -> Python nativo`
- `Accion 3` -> `Laravel -> Python nativo`
- `Accion 4` -> `Laravel -> Python nativo`
- `Consolidado` -> `Laravel -> Python nativo`

### Repuestos TYTSERV

- `Laravel -> Python nativo`
- `Node.js` queda solo como referencia/fallback controlado

### Servicios por Marca

- formulario, polling, historial y control del job -> `Laravel`
- validacion previa, readers (`source`, `px`, `mayor`), dispatch y worker final -> `Python`
- escritura final -> `Excel COM` desde Python
- snapshot tecnico de plantilla base en `storage/cache/servicios_marcas` para no releer por COM la misma estructura en cada corrida
- los readers Node antiguos quedan solo como contrato de paridad, no como runtime activo

## Lo que ya no es superficie activa

Archivado en:

- `archive/legacy_php_surface_2026-03-24/areas`
- `archive/legacy_php_surface_2026-03-24/modules`
- `archive/legacy_php_surface_2026-03-24/src`
- `archive/legacy_php_surface_2026-03-24/templates`
- `archive/legacy_php_surface_2026-03-24/includes`
- `archive/legacy_php_surface_2026-03-24/download.php`
- `archive/legacy_php_surface_2026-03-24/export_all_actions.php`
- `archive/legacy_php_surface_2026-03-24/run_servicios_marcas_job.php`

## Excepcion tecnica que sigue vigente

`Servicios por Marca` aun depende de:

- `Excel COM` en Windows
- un host Windows estable para automatizacion Office

Eso ya no contradice la migracion a `Laravel + Python`. Es una restriccion del motor Excel del host, no de la vieja superficie web ni del runtime principal.

## Pendientes reales

1. Mantener estable y vigilado `Servicios por Marca`.
2. Decidir si en el futuro conviene eliminar la dependencia de `Excel COM` o mantenerla como excepcion controlada.
3. Seguir limpiando documentacion menor si aparecen referencias viejas.

## Validacion minima requerida

- `node scripts/tests/http_resources_smoke.js`
- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_window1.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_repuestos.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_servicios.ps1`

## Conclusion

La migracion ya no esta en fase de idea ni de scaffold. El sistema queda operando de forma principal sobre `Laravel + Python`, con una unica excepcion controlada de `Excel COM` para `Servicios por Marca`.

La afirmacion honesta al **25 de marzo de 2026** es esta:

- `web principal`: `Laravel`
- `procesamiento principal`: `Python`
- `ultima excepcion controlada`: writer final de `Servicios por Marca` en `Python + Excel COM`

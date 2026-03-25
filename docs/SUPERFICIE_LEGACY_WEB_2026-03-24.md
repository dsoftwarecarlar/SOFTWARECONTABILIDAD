# Superficie Legacy Web - 2026-03-24

## Estado

La superficie web PHP anterior ya no forma parte del flujo primario del sistema.

La entrada actual debe entenderse asi:

- raiz del proyecto -> `Laravel`
- rutas operativas -> `Laravel`
- procesamiento principal -> `Python`

La superficie vieja fue archivada y ya no debe usarse como punto de entrada manual.

## Ubicacion del archivo historico

- `archive/legacy_php_surface_2026-03-24/`

## Lo que fue archivado

- `areas/`
- `modules/`
- `src/`
- `templates/`
- `includes/`
- `download.php`
- `export_all_actions.php`
- `run_servicios_marcas_job.php`

## Lo que sigue vivo fuera de Laravel + Python

Solo por necesidad operativa:

- `scripts/cxp/servicios_marcas/run.ps1`
- `run_servicios_marcas.ps1`
- `Excel COM`

Eso corresponde al motor heredado de `Servicios por Marca`, no a la vieja superficie web PHP.

## Compatibilidad publica

`.htaccess` mantiene compatibilidad para:

- URLs antiguas de `areas/cxp/*`
- URLs antiguas de `modules/cxp_*/index.php`
- `download.php?file=...`
- `export_all_actions.php`

Pero esas URLs ya redirigen hacia rutas limpias del portal actual.

## Decision operativa

Desde el **24 de marzo de 2026**:

- `Laravel` es la web principal
- `Python` es el motor principal de `Libro Compras ACLT` y `Repuestos TYTSERV`
- la vieja superficie web PHP queda solo como archivo historico de migracion

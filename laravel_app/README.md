# Laravel App

`laravel_app` es la capa web principal del sistema.

Desde el **24 de marzo de 2026**, esta carpeta ya no es un scaffold de convivencia: aqui vive la navegacion principal, los formularios, las descargas y la orquestacion de los tres frentes operativos de `Contabilidad Talleres`.

## Estado real

### Web principal

- home del portal
- area `CXP`
- ventanas
- modulos
- historial
- descargas

### Procesos ya gobernados por Laravel

- `Accion 1`
- `Accion 2`
- `Accion 3`
- `Accion 4`
- `Consolidado de Acciones`
- `Repuestos TYTSERV`
- `Servicios por Marca`

## Rutas activas

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
- `/downloads/{file}`

## Arquitectura operativa

### Libro Compras ACLT

- `Laravel -> Python nativo`

### Repuestos TYTSERV

- `Laravel -> Python nativo`
- `scripts/cxp/repuestos_tytserv/process.js` queda como referencia y fallback controlado

### Servicios por Marca

- `Laravel -> Python worker -> Excel COM`
- `Python` ya cubre readers de `source`, `px`, `mayor`, dispatch y escritura final; `Node` ya no participa en el runtime activo de este modulo

La unica dependencia operativa que aun no se reemplaza es `Excel COM` en Windows. La capa web, el polling, el catalogo de jobs y el control del proceso ya viven aqui.

## Endurecimiento publico

- `.env` queda en `APP_ENV=production`
- `.env` queda en `APP_DEBUG=false`
- `database/database.sqlite` existe para evitar fallas publicas si el runtime resuelve sesiones sobre SQLite

La URL publica que debe responder en Apache/XAMPP es:

- `/SOFTWARECONTABILIDAD/`
- `/SOFTWARECONTABILIDAD/cxp`
- `/SOFTWARECONTABILIDAD/cxp/windows/libro-compras-aclt`
- `/SOFTWARECONTABILIDAD/cxp/windows/conciliacion-servicios-marcas`
- `/SOFTWARECONTABILIDAD/cxp/windows/facturacion-repuestos-tytserv`

## Archivos clave

- `routes/web.php`
- `app/Http/Controllers/*`
- `app/Services/*`
- `config/cxp.php`
- `resources/views/*`
- `public/assets/portal.css`
- `public/brand/logo-symbol.jpg`

## Comandos utiles

- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\laravel_artisan.ps1 route:list`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\laravel_artisan.ps1 python:probe`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\run_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\run_laravel_app.ps1 -ListenHost 0.0.0.0 -Port 8085`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\run_laravel_app_lan.ps1 -Port 8085`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\stop_laravel_app_lan.ps1 -Port 8085`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_window1.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_repuestos.ps1`
- `powershell -ExecutionPolicy Bypass -File .\\scripts\\dev\\test_laravel_servicios.ps1`
- `npm run test:audit:window1`
- `npm run test:audit:mensual`

## Acceso por red local

Para abrir el portal desde otra maquina de la misma red local:

1. inicia `run_laravel_app_lan.ps1` en este host;
2. usa una de las URLs LAN que imprime el script;
3. si Windows bloquea el acceso externo, ejecuta el mismo script con `-TryOpenFirewall`.

Si el sistema no se ejecuto como administrador, Windows puede impedir crear la regla automaticamente. En ese caso el portal ya queda escuchando en `0.0.0.0`, pero la apertura definitiva del puerto debe hacerse con permisos elevados.

## Nota de despliegue

En XAMPP/Apache, la raiz del proyecto ya puede actuar como entrada limpia del portal. El usuario no necesita entrar manualmente a `laravel_app/public`; esa ruta interna solo se conserva por compatibilidad tecnica.

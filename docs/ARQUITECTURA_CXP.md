# ARQUITECTURA CXP

Fecha base: 2026-03-24
Estado: operacion principal en `Laravel + Python`

## 1. Vista general

El sistema `CXP` queda compuesto por cuatro capas claras:

1. `Presentacion`
   - `Laravel`
   - `Blade`
   - rutas limpias desde la raiz del proyecto

2. `Aplicacion`
   - controladores Laravel
   - validacion de archivos
   - historial
   - descargas
   - control de jobs

3. `Procesamiento`
   - `Python` para `Accion 1..4`, `Consolidado`, `Repuestos TYTSERV` y `Servicios por Marca`
   - `Excel COM` solo como dependencia de escritura final en `Servicios por Marca`

4. `Archivos`
   - plantillas y fixtures en `resources/cxp`
   - uploads en `storage/uploads`
   - salidas en `storage/outputs`
   - jobs en `storage/jobs`

## 2. Superficie activa

### 2.1 Web

- `index.php`
- `.htaccess`
- `laravel_app/routes/web.php`
- `laravel_app/resources/views/*`
- `laravel_app/app/Http/Controllers/*`
- `laravel_app/app/Services/*`

### 2.2 Procesadores

- `python_services/processors/cxp_actions/*`
- `python_services/processors/repuestos_tytserv/*`
- `python_services/processors/servicios_marcas/*`

### 2.3 Configuracion

- `laravel_app/config/cxp.php`
- `config/cxp/action_exports.json`
- `config/cxp/repuestos_tytserv.php`
- `config/cxp/servicios_marcas.php`

## 3. Flujos principales

### 3.1 Libro Compras ACLT

1. el usuario entra por Laravel
2. carga archivos en la pantalla del modulo
3. Laravel guarda uploads temporales
4. Laravel llama al procesador Python
5. Python genera salida en `storage/outputs`
6. Laravel muestra resultado e historial

### 3.2 Repuestos TYTSERV

1. el usuario carga `8` archivos Excel
2. Laravel valida el set mensual
3. Python procesa el libro completo
4. se publica un `.xlsx` final

### 3.3 Servicios por Marca

1. el usuario carga Excel comunes y TXT por marca
2. Laravel crea y controla el job
3. Python valida y prepara el dispatch
4. el worker `Python + Excel COM` genera el resultado final
5. Laravel publica estado, historial y descarga

## 4. Mapa de runtimes

| Flujo | Entry point oficial | Runtime principal | Nota |
| --- | --- | --- | --- |
| Accion 1 | Laravel | Python | PDF nativo |
| Accion 2 | Laravel | Python | TXT nativo |
| Accion 3 | Laravel | Python | TXT y PDF nativos |
| Accion 4 | Laravel | Python | MAYOR IVA nativo |
| Consolidado | Laravel | Python | une las 4 acciones |
| Repuestos TYTSERV | Laravel | Python | Node queda como referencia |
| Servicios por Marca | Laravel | Python + Excel COM | unica excepcion fuerte |

## 5. Almacenamiento

| Ruta | Uso |
| --- | --- |
| `resources/cxp/acciones/templates` | plantillas de `Accion 1..4` |
| `resources/cxp/acciones/fixtures` | fixtures de contrato |
| `resources/cxp/servicios_marcas/templates` | plantillas de servicios |
| `resources/cxp/servicios_marcas/fixtures` | fixtures de servicios |
| `resources/cxp/repuestos_tytserv/templates` | plantilla mensual de repuestos |
| `resources/cxp/repuestos_tytserv/fixtures` | fixtures de repuestos |
| `storage/uploads` | buffer temporal |
| `storage/outputs` | salida final descargable |
| `storage/jobs` | snapshots y cancelaciones |
| `archive/legacy_php_surface_2026-03-24` | estructura PHP archivada |

## 6. Limpieza y retencion

La limpieza ya no depende del helper global viejo:

- `maintenance_cleanup.php`

Valores actuales:

- salidas: `3`
- uploads: `20`

## 7. Riesgo principal restante

El riesgo mayor ya no es la vieja superficie web, sino `Servicios por Marca` por:

- dependencia de `Excel COM`
- necesidad de host Windows estable
- posibilidad de bloqueo si existe un `EXCEL.EXE` visible abierto

## 8. Fuente de verdad actual

- `laravel_app/routes/web.php`
- `laravel_app/config/cxp.php`
- `laravel_app/app/Services/*`
- `python_services/processors/*`

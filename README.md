# SOFTWARECONTABILIDAD

Portal operativo interno para procesos contables del area de Cuentas por Pagar y conciliaciones asociadas. El sistema recibe archivos `PDF`, `TXT`, `XLS` y `XLSX`, ejecuta transformaciones controladas y publica archivos finales en `storage/outputs`.

## 1. Resumen ejecutivo

El proyecto esta organizado alrededor del workspace `cxp` y hoy opera con tres ventanas:

1. `Libro de Compras ACLT`
   Incluye `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4` y `Consolidado General`.
2. `Conciliacion Servicios por Marca`
   Procesa `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ` con jobs en segundo plano.  
3. `Facturacion Repuestos TYTSERV`
   Procesa ventas y devoluciones por marca y genera un consolidado mensual.

## 2. Arquitectura resumida

El sistema usa tres runtimes principales:

- `PHP`
  Portal web, controladores, configuracion, render de vistas, descarga y limpieza.
- `Node.js`
  Procesos contables de `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4`, consolidado general y `Repuestos TYTSERV`.
- `PowerShell + Excel COM`
  Worker especializado de `Servicios por Marca`.

## 3. Modulos activos

| Modulo | Entrada | Runtime | Salida |
| --- | --- | --- | --- |
| Accion 1 - Libro Compras Proveedores | `PDF` | `Node.js` | `.xlsx` |
| Accion 2 - Retenciones Proveedores | `TXT` | `Node.js` | `.xlsx` |
| Accion 3 - Mayor Retenciones | `1..n TXT` | `Node.js` | `.xlsx` |
| Accion 4 - Mayor IVA | `TXT` | `Node.js` | `.xlsx` |
| Consolidado General | ultimas salidas de Accion 1..4 | `Node.js` | `.xlsx` |
| Servicios por Marca | `2 Excel comunes + 3 TXT por marca` | `PowerShell + Excel COM` | `1 .xls por marca` |
| Repuestos TYTSERV | `4 Excel ventas + 4 Excel devoluciones` | `Node.js` | `.xlsx` |

## 4. Estructura principal del repositorio

| Ruta | Responsabilidad |
| --- | --- |
| `index.php` | portal principal |
| `areas/cxp` | ventanas del workspace contable |
| `modules/*` | entrypoints web por modulo |
| `src/*` | controladores y soporte de aplicacion |
| `includes/*` | bootstrap, helpers y funciones globales |
| `config/cxp/*` | configuracion funcional por modulo |
| `scripts/cxp/*` | procesos contables y workers |
| `resources/cxp/*` | plantillas base y fixtures |
| `storage/uploads` | archivos subidos temporalmente |
| `storage/outputs` | artefactos finales descargables |
| `storage/jobs` | estado de jobs de servicios por marca |
| `docs/*` | documentacion operativa y tecnica |

## 5. Rutas operativas clave

- `resources/cxp/acciones`
  Plantillas y fixtures de `Accion 1..4`.
- `resources/cxp/servicios_marcas`
  Plantillas y fixtures de `Servicios por Marca`.
- `resources/cxp/repuestos_tytserv`
  Plantillas y fixtures de `Repuestos TYTSERV`.
- `storage/outputs`
  Unica ruta publica de descarga de archivos finales.
- `storage/uploads`
  Buffer temporal de archivos subidos.
- `storage/jobs`
  Cola, estado y cancelacion de jobs de `Servicios por Marca`.

## 6. Requisitos operativos

- Host Windows con `PHP` y `Apache` o equivalente.
- `Node.js` instalado y disponible para los procesos `Node`.
- `npm install` ejecutado en la raiz del proyecto.
- `Microsoft Excel Desktop` disponible en el host que ejecuta `Servicios por Marca`.
- Permisos de escritura en:
  - `storage/uploads`
  - `storage/outputs`
  - `storage/jobs`

## 7. Validacion y pruebas

Comandos disponibles en `package.json`:

- `npm run test:quality:contable`
- `npm run test:e2e:accion1-4`
- `npm run test:e2e:accion2-3`
- `npm run test:e2e:repuestos`
- `npm run test:e2e:servicios`
- `npm run test:smoke:ui`

Limpieza operativa manual:

- `php maintenance_cleanup.php`

## 8. Notas de operacion

- Las salidas publicas se sirven por `download.php`.
- `storage/outputs` conserva solo un historico corto por accion o prefijo.
- `storage/uploads` conserva un buffer corto de archivos recientes.
- `archive/cxp_manual_outputs` queda como historico, no como fuente operativa.
- `Repuestos TYTSERV` usa `Node.js` como runtime web oficial.
- `Servicios por Marca` depende de `PowerShell + Excel COM`; es el flujo mas sensible del sistema.

## 9. Documentacion relacionada

- [`docs/PLAN_DOCUMENTACION_CXP.md`](docs/PLAN_DOCUMENTACION_CXP.md)
- [`docs/ARQUITECTURA_CXP.md`](docs/ARQUITECTURA_CXP.md)
- [`docs/MODULOS_CXP.md`](docs/MODULOS_CXP.md)
- [`docs/OPERACION_CXP.md`](docs/OPERACION_CXP.md)
- [`docs/BITACORA_AVANCES.md`](docs/BITACORA_AVANCES.md)
- [`docs/CHECKLIST_DEPLOY_CONTABLE.md`](docs/CHECKLIST_DEPLOY_CONTABLE.md)

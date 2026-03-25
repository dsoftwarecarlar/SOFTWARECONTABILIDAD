# OPERACION CXP

Fecha base: 2026-03-25  
Estado: operacion principal en `Laravel + Python`

## 1. Entrada operativa

El usuario no debe entrar por archivos `.php` sueltos.

Entradas de trabajo:

- `/SOFTWARECONTABILIDAD/`
- `/SOFTWARECONTABILIDAD/cxp`
- `/SOFTWARECONTABILIDAD/cxp/windows/libro-compras-aclt`
- `/SOFTWARECONTABILIDAD/cxp/windows/conciliacion-servicios-marcas`
- `/SOFTWARECONTABILIDAD/cxp/windows/facturacion-repuestos-tytserv`

## 2. Requisitos

- Windows operativo
- `PHP 8.2+`
- `Python`
- `Node.js` mientras exista fallback de referencia
- `Excel Desktop` para `Servicios por Marca`
- permisos de escritura en:
  - `storage/uploads`
  - `storage/outputs`
  - `storage/jobs`
  - `storage/cache/servicios_marcas`

## 3. Checklist mensual corto

Antes de empezar:

- entrar por `/SOFTWARECONTABILIDAD/cxp`
- confirmar mes de trabajo y archivos fuente del mes
- cerrar ventanas visibles de Excel antes de usar `Servicios por Marca`

Ventana 1. Libro Compras ACLT:

- ejecutar `Accion 1`
- ejecutar `Accion 2`
- ejecutar `Accion 3`
- ejecutar `Accion 4`
- generar `Consolidado`
- revisar que cada descarga corresponda al archivo cargado

Ventana 2. Conciliacion Servicios por Marca:

- cargar los 2 Excel comunes del mes
- cargar los TXT por marca correctos
- esperar el fin del proceso sin cerrar la pantalla
- descargar un archivo final por marca

Ventana 3. Facturacion Repuestos TYTSERV:

- cargar los 8 archivos requeridos
- revisar resumen por marca
- descargar el libro final del mes

Al terminar:

- confirmar que las descargas quedaron en `storage/outputs`
- ejecutar `php maintenance_cleanup.php` si se requiere limpieza operativa

## 4. Validacion rapida

- `node scripts/tests/http_resources_smoke.js`
- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios`
- `npm run test:audit:definitiva`

## 5. Notas operativas

- `Servicios por Marca` usa cache tecnico en `storage/cache/servicios_marcas`; la primera corrida de una plantilla nueva puede tardar mas.
- La URL publica Apache `/SOFTWARECONTABILIDAD/cxp` debe responder sin mostrar errores de Laravel ni pantallas tecnicas.
- Si el equipo abre el portal por otra maquina en la red, debe usar la ruta completa del host, por ejemplo `http://IP_DEL_EQUIPO/SOFTWARECONTABILIDAD/`.

## 6. Referencias

- `docs/ARQUITECTURA_CXP.md`
- `docs/MODULOS_CXP.md`
- `docs/CHECKLIST_DEPLOY_CONTABLE.md`

# MODULOS CXP

Fecha base: 2026-03-24
Estado: catalogo activo sobre `Laravel + Python`

## 1. Catalogo rapido

| Modulo | Objetivo | Entrada principal | Salida | Runtime actual |
| --- | --- | --- | --- | --- |
| Accion 1 | convertir PDF de proveedor en libro de compras | `PDF` | `.xlsx` | `Laravel -> Python` |
| Accion 2 | convertir TXT de retenciones a RET PROV | `TXT` | `.xlsx` | `Laravel -> Python` |
| Accion 3 | transformar TXT o PDF MAYOR GENERAL a MAYOR RET | `TXT` o `PDF` | `.xlsx` | `Laravel -> Python` |
| Accion 4 | transformar TXT MAYOR IVA al formato manual | `TXT` | `.xlsx` | `Laravel -> Python` |
| Consolidado General | unir la ultima salida de Accion 1..4 | salidas previas | `.xlsx` | `Laravel -> Python` |
| Servicios por Marca | separar y poblar plantillas operativas por marca | `2 Excel + 3 TXT por marca` | `.xls` | `Laravel -> Python -> Excel COM` |
| Repuestos TYTSERV | consolidar ventas y devoluciones por marca | `8 Excel` | `.xlsx` | `Laravel -> Python` |

## 2. Rutas de usuario

- `/cxp/modules/accion1`
- `/cxp/modules/accion2`
- `/cxp/modules/accion3`
- `/cxp/modules/accion4`
- `/cxp/modules/consolidado-acciones`
- `/cxp/modules/servicios-marcas`
- `/cxp/modules/repuestos-tytserv`

## 3. Notas importantes por modulo

### Accion 1

- plantilla base: `resources/cxp/acciones/templates/EJEMPLODECOMOQUEDARIA.xlsx`
- salida esperada: `*_resultado.xlsx`

### Accion 2

- plantilla base: `resources/cxp/acciones/templates/ACCION2.xlsx`
- salida esperada: `*_accion2.xlsx`

### Accion 3

- plantilla base: `resources/cxp/acciones/templates/MAYOR RET_ACCION3.xlsx`
- acepta `TXT` y `PDF`
- salida esperada: `*_accion3.xlsx`

### Accion 4

- plantilla base: `resources/cxp/acciones/templates/MAYORIVAACCION4.xlsx`
- salida esperada: `*_accion4.xlsx`

### Consolidado

- requiere salidas previas validas de `Accion 1..4`
- salida esperada: `acciones_resumen_*.xlsx`

### Servicios por Marca

- requiere `PX` y `REP VENTAS`
- requiere `REP FACTURACION`, `NOTA DE CREDITO` y `MAYOR` por marca
- genera una salida por marca con prefijo:
  - `servicios_changan_`
  - `servicios_peug_`
  - `servicios_szk_`
  - `servicios_tyt_`

### Repuestos TYTSERV

- requiere `8` Excel del mes
- genera libro mensual con prefijo `repuestos_tytserv_`

## 4. Pruebas asociadas

- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:accion1-4`
- `npm run test:e2e:accion2-3`
- `npm run test:e2e:repuestos`
- `npm run test:e2e:servicios`

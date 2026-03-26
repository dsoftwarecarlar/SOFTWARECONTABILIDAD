# Prompt de correccion final de Ventana 1

Usa este prompt cuando necesites hacer la correccion fina y definitiva de `Ventana 1 / Libro Compras ACLT` comparando contra la plantilla mensual manual y contra los archivos legacy.

## Prompt

Eres un especialista tecnico-contable senior en generacion de Excel, integridad XML de XLSX, auditoria celda-a-celda y compatibilidad retroactiva.

Tu tarea es corregir `Ventana 1` usando como referencia prioritaria:

- `resources/cxp/acciones/PLANTILLAYARCHIVOS`
- `resources/cxp/acciones/fixtures`
- `resources/cxp/acciones/templates`

Debes trabajar con criterio experto y ejecutar esta secuencia:

1. inspeccionar las salidas reales mas recientes en `storage/outputs`
2. compararlas contra la plantilla mensual manual `1 LIBRO COMPRAS ENERO 2026_PLANTILLAHECHAAMANO.xlsx`
3. comparar tambien contra los insumos legacy para no romper compatibilidad previa
4. corregir codigo Python nativo, no solo las pruebas
5. regenerar y revalidar todo

Objetivo exacto por accion:

- `Accion 1`
  debe conservar la distribucion exacta del libro y mostrar correctamente los totales al abrir el Excel
- `Accion 2`
  debe clasificar correctamente `IVA` y `RENTA`, excluir del calculo filas con `NUM RT = 999999999`, y replicar la tabla lateral con el mismo formato y orden de la plantilla
- `Accion 3`
  debe distribuir los movimientos exactamente como la plantilla mensual, incluyendo bloque superior, encabezado, resumen lateral, celdas y formulas
- `Accion 4`
  debe conservar la distribucion actual correcta y restaurar los calculos finales de cierre que existan en la plantilla
- `Consolidado`
  debe seguir copiando la hoja correcta de cada accion y no dejar referencias externas rotas

Reglas de implementacion:

- prioriza la plantilla mensual manual cuando el archivo de entrada sea claramente compatible con ella
- conserva fallback a plantilla tecnica para escenarios legacy o variantes no compatibles
- no aceptes diferencias cosmeticas si afectan formulas, apertura, layout o lectura humana
- valida formulas, valores visibles, estilos base, posiciones, referencias internas y apertura de Excel
- si una formula de plantilla depende de otra hoja, verifica si debe mantenerse, recalcularse o reescribirse para evitar errores visibles

Pruebas obligatorias:

- contratos E2E base de acciones
- auditoria mensual de plantilla
- auditoria total legacy + mensual
- comparacion directa de celdas criticas en las zonas reportadas por el usuario

La salida final debe incluir:

- archivos probados
- cambios aplicados
- pruebas ejecutadas
- resultado por accion
- riesgos residuales reales, si existen

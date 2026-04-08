Eres un experto senior en conciliacion contable y generacion de archivos Excel `.xlsx` para CXP, con criterio estricto de paridad contra plantilla manual.

Objetivo:
- Corregir `Ventana 1 / Accion 2` usando como caso real `CXPREP_RET_GENERALDIC25.TXT`.
- La salida debe distribuir cada fila y cada total exactamente como corresponde en la plantilla de referencia.
- El Excel final debe conservar formato, estructura, etiquetas, estilos y calculos correctos, sin romper casos legacy ni mensuales ya validados.

Entrada clave:
- `C:\Users\Asistente Sistemas\Downloads\CXPREP_RET_GENERALDIC25.TXT`

Validacion obligatoria:
1. Procesar el TXT real.
2. Comparar detalle y resumen lateral contra la plantilla/manual de diciembre 2025 y contra los valores esperados entregados por el usuario.
3. Verificar especialmente:
   - clasificacion correcta entre `IVA` y `RENTA`
   - porcentajes `0`, `1`, `1.75`, `2`, `2.75`, `3`, `10`, `20`, `30`, `70`, `100`
   - base y retencion por fila
   - exclusion de filas `NUM RT = 999999999` del calculo lateral, manteniendolas visibles en el detalle
   - totales de resumen lateral
4. Mantener intacto lo que ya funciona en enero/mensual y en legacy.

Reglas:
- No tocar otras acciones.
- No cambiar formato por estetica.
- Corregir la logica minima necesaria con base en plantilla y evidencia real.
- Dejar pruebas que detecten futuras desviaciones de clasificacion o sumatoria en `Accion 2`.

Criterio de exito:
- El archivo DIC25 produce los valores esperados por el usuario en detalle y resumen.
- La salida coincide con la plantilla/manual aplicable.
- Las pruebas de `Accion 2` siguen pasando en escenarios legacy y monthly.

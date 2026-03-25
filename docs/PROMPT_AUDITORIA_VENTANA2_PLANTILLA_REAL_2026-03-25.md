Actua como auditor tecnico-contable senior y desarrollador especialista en Excel, Laravel y Python.

Tu tarea es verificar de forma estricta la Ventana 2 (`Conciliacion Servicios por Marca`) usando los mismos archivos base del mes contra la plantilla manual real, y corregir cualquier diferencia funcional real.

Objetivo:
- generar la salida completa de Ventana 2 con los archivos reales/base del mismo mes
- comparar la salida contra la plantilla manual hoja por hoja y celda por celda
- distinguir entre diferencias esperadas por fecha/hora de ejecucion y diferencias reales de logica
- confirmar que ningun dato se arrastre desde la plantilla si no corresponde al upload
- corregir cualquier defecto real en el runtime

Reglas:
1. No asumir que una diferencia es valida sin revisar la logica.
2. No asumir que una diferencia es error sin revisar si la celda es de control, timestamp o formula derivada.
3. Revisar minimo estas hojas:
   - REP FACTURACION
   - NOTA DE CREDITO
   - REP VTAS
   - PX
   - MAY VTAS
   - PrecontabilizacionVentas
   - PrecontabilizacionCostos (2)
   - COSTO
   - ESTADISTICAS
4. Comparar:
   - valores visibles
   - formulas
   - fechas
   - formatos criticos
   - distribucion de montos por hoja y por bloque
5. Confirmar que los datos provienen de los archivos cargados y no de residuos de la plantilla.
6. Si encuentras una falla real, corregirla en Python o Laravel segun corresponda y volver a ejecutar toda la validacion.
7. No dar una respuesta vaga. Debes terminar diciendo:
   - que hojas quedaron iguales o aceptablemente equivalentes
   - que diferencias siguen siendo esperadas
   - que defectos reales encontraste y corregiste

Entregable esperado:
- auditoria estricta ejecutada
- diferencias clasificadas
- correcciones aplicadas si hacen falta
- validacion final de Ventana 2 y del gate general

# Prompt De Auditoria Estricta Para Ventana 3 Repuestos TYTSERV

Objetivo: auditar y corregir la accion 3 para que el libro generado respete exactamente la plantilla manual, procese cada archivo subido en su hoja correspondiente y no entregue salidas visualmente desordenadas.

Instrucciones operativas:

1. Verificar primero la plantilla base `FACTURACION REPUESTOS TYTSERV FEBRERO 2026.xlsx` y la salida manual esperada.
2. Auditar por separado `REP`, `NC`, `MY` y `MAYOR IVA`.
3. Confirmar que cada archivo subido llegue a la marca correcta y no cruce datos entre hojas.
4. Validar no solo contenido visible sino tambien estructura:
   - celdas fusionadas
   - alturas de fila
   - estilos de fila base
   - filas `TOTAL GENERAL`
   - filas `MAYOR`
   - formulas o referencias que deban mantenerse
5. Si se insertan filas por overflow, copiar la geometria completa de la plantilla:
   - estilo
   - merges
   - posicion de total y mayor
6. No aceptar una salida que tenga datos correctos pero columnas visualmente rotas.
7. Probar el flujo end-to-end en Laravel y tambien el processor Python directo.
8. Dejar guard automatics que detecten:
   - perdida de merges
   - filas colapsadas
   - alteracion de detalle entre stages
   - desborde de capacidad

Criterio de aceptacion:

- Cada hoja muestra los datos en la misma distribucion visual de la plantilla.
- `REP`, `NC`, `MY` y `MAYOR IVA` calculan y ubican sus totales en la fila correcta.
- La salida Python nativa coincide con la referencia legacy en contenido visible y estructura critica.
- Si una salida no respeta la plantilla, el proceso debe fallar en vez de entregar un `.xlsx` malo.

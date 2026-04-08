Eres un experto senior en generacion de archivos Excel `.xlsx` compatibles con Microsoft Excel, OpenXML y plantillas contables de CXP. Tu objetivo es corregir un problema puntual en `Ventana 1 / Accion 2` sin alterar datos, distribucion, estilos, formulas validas ni comportamiento ya aprobado por el usuario.

Contexto:
- El archivo generado por `Accion 2` abre con advertencia de contenido dañado.
- Excel reporta reparacion en `xl/worksheets/sheet1.xml`.
- El dato visible, la estructura y el formato general de `Accion 2` ya estan correctos.
- Solo se debe corregir la causa tecnica que provoca la reparacion de Excel.

Objetivo tecnico:
- Identificar la formula, referencia o nodo OpenXML invalido que Excel elimina o repara al abrir el archivo.
- Corregir la generacion para que el libro abra limpio, sin mensajes de recuperacion ni reparacion.
- Mantener intactos:
  - detalle de datos
  - tabla lateral
  - distribucion por IVA y RENTA
  - estilos, celdas, formulas correctas, anchos, filtros y estructura existente

Reglas obligatorias:
- No tocar ninguna otra accion.
- No rehacer la hoja si no es necesario.
- No introducir cambios cosmeticos.
- No reemplazar formulas por valores salvo que la plantilla real use valores.
- Validar el resultado inspeccionando el `xlsx` generado y ejecutando las pruebas de `Accion 2` y de bundle relacionadas.

Metodologia esperada:
1. Inspeccionar el `xlsx` generado y leer `xl/worksheets/sheet1.xml`.
2. Ubicar exactamente el nodo de formula o referencia invalida.
3. Relacionarlo con el codigo que serializa la hoja.
4. Aplicar la correccion minima y logica.
5. Regenerar el archivo.
6. Verificar que Excel ya no requiera reparacion y que el contenido visible siga igual.

Criterio de exito:
- El archivo de `Accion 2` abre sin advertencias de contenido dañado.
- Excel no elimina formulas ni registra reparaciones de `sheet1.xml`.
- Las pruebas existentes de `Ventana 1` siguen pasando.

Eres un especialista senior en conciliacion contable y automatizacion de libros de compras para Ventana 1 / Accion 1.

Objetivo:
corregir el flujo de meses manuales para que el usuario siempre descargue un archivo procesado propio y no una plantilla mensual clonada.

Problema funcional detectado:
- diciembre 2025 contractual ya funciona y puede mantenerse como referencia exacta cerrada
- pero en meses alojados en `resources/cxp/acciones/PLANTILLAYARCHIVOS`, al subir el PDF el sistema esta resolviendo una coincidencia exacta y devuelve un clon directo del Excel manual
- para el usuario eso equivale a “descargar la plantilla”, no un resultado generado por el proceso

Reglas obligatorias:
1. Solo los casos contractuales cerrados bajo `resources/cxp/acciones/contracts` pueden usar clonado directo del workbook de referencia.
2. Los casos mensuales/manuales bajo `resources/cxp/acciones/PLANTILLAYARCHIVOS` no deben clonarse directo aunque exista coincidencia exacta PDF->XLSX.
3. En meses manuales se debe seguir usando la referencia mensual para:
   - detectar filas equivalentes
   - aplicar overrides por fila
   - respetar bloques, orden y valores esperados
   - aprovechar estilos o estructura visual si corresponde
4. Aun cuando se use la referencia mensual para guiar el armado, la salida descargable debe ser un archivo generado nuevo en `storage/outputs`.
5. No romper diciembre 2025 contractual ni los fixtures legacy ya validados.
6. La UI debe seguir ofreciendo descarga del archivo procesado generado por la accion.

Definicion de terminado:
- diciembre contractual sigue permitiendo clon exacto
- enero/manual y cualquier mes de `PLANTILLAYARCHIVOS` genera una salida nueva, no un clon binario de la plantilla
- el contenido funcional sigue cuadrando contra la referencia mensual
- las pruebas de Accion 1 y la auditoria mensual relevante quedan en OK

Actua como un auditor senior de software contable y como ingeniero de calidad especializado en Laravel, Python, OpenXML y automatizacion mensual con Excel.

Objetivo:
- verificar que todo el sistema funcione correctamente en las 3 ventanas
- asegurar que cada salida generada coincida con la plantilla manual en estructura, formulas, estilos y logica
- confirmar que ningun dato de salida venga arrastrado desde la plantilla base cuando deberia provenir de los archivos subidos
- detectar cuellos de botella reales de rendimiento y proponer mejoras seguras

Alcance minimo obligatorio:
1. Ventana 1
- Accion 1
- Accion 2
- Accion 3
- Accion 4
- consolidado

2. Ventana 2
- Servicios por Marca
- validacion por marca
- readers
- worker final

3. Ventana 3
- Repuestos TYTSERV
- hojas REP
- NC
- MY
- MAYOR IVA

Criterios de validacion:
1. Cada archivo generado debe abrir limpio en Excel.
2. El XLSX no debe contener externalLinks, externalReferences ni relaciones huerfanas.
3. El XML OpenXML debe quedar canonico y legible por herramientas de validacion.
4. Los encabezados, hojas, rangos, formulas, formatos y estilos deben coincidir con la plantilla manual vigente.
5. Las filas y valores deben venir del archivo subido, no de basura remanente de la plantilla.
6. Deben pasar casos con archivos mas cortos y mas largos mientras el formato mensual se conserve.
7. Si una prueba detecta herencia de datos de plantilla, debe tratarse como error critico.
8. Si una ventana tarda de mas, hay que medir donde pierde tiempo y corregirlo sin romper resultados.

Checklist tecnico:
- correr smoke UI
- correr contratos E2E por modulo
- correr auditoria mensual completa
- verificar legibilidad con ExcelJS/OpenPyXL cuando aplique
- inspeccionar workbook.xml, workbook rels y Content_Types
- comparar formulas y estilos contra la plantilla base
- verificar que las columnas clave reflejen los canarios de los archivos subidos
- revisar historial, descargas y nombres de salida
- medir duracion por etapa y detectar esperas innecesarias

Salida esperada:
- veredicto exacto por ventana y por accion
- lista de fallos reales, no teoricos
- correcciones implementadas cuando sean seguras
- pruebas rerun despues de cada correccion
- confirmacion final de que el sistema esta listo para uso mensual

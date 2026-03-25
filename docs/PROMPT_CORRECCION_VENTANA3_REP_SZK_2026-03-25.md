# Prompt de trabajo: corrección Ventana 3 REP SZK

Analiza y corrige el fallo del pipeline Python nativo de Repuestos TYTSERV en Ventana 3 cuando la fuente de SUZUKI excede la capacidad visible de la plantilla REP.

Contexto del error reportado:
- `La fuente SUZUKI supera la capacidad actual de la plantilla en REP SZK. Fuente TOTAL=90, plantilla TOTAL=63.`
- `openpyxl.styles.stylesheet: Workbook contains no default style, apply openpyxl's default`

Objetivo:
- El proceso no debe abortar solo porque la fuente tenga mas filas que la plantilla base.
- Debe comportarse igual que el flujo legacy: copiar hasta la fila real de la fuente, mover `TOTAL GENERAL` y `MAYOR` al cierre real, y limpiar residuos por debajo.
- La correccion debe cubrir REP y NC para evitar el mismo fallo latente en otras marcas o en notas de credito.
- El warning benigno de `openpyxl` sobre `no default style` no debe contaminar la salida operativa.

Criterios de aceptacion:
- `REP SZK` acepta una fuente con `TOTAL GENERAL` mayor que la plantilla.
- `MAYOR` queda en `TOTAL GENERAL + 1`.
- No quedan filas viejas con contenido por debajo del nuevo `MAYOR`.
- Se mantienen las pruebas actuales y se agrega una reproduccion local del caso sobredimensionado.

# Prompt Correccion NC Vacio Repuestos 2026-03-25

Analiza y corrige el flujo de Repuestos TYTSERV cuando un archivo de devoluciones (`RepLibroDevolucionesGeneral`) llega como `.xlsx` vacio sin hojas.

Contexto confirmado:
- En produccion quedaron cargas `excel_nc_peug` y `excel_nc_chgn` de `1858 bytes`.
- Esos archivos son workbooks validos a nivel ZIP/OpenXML, pero `xl/workbook.xml` trae `<sheets></sheets>`.
- El sistema hoy falla con: `El archivo fuente PEUGEOT debe contener la hoja 'RepLibroDevolucionesGeneral'. No se aceptan plantillas ni salidas ya generadas.`
- Existen fixtures NC del repositorio con ese mismo patron de workbook vacio, lo que indica que el origen puede exportar devoluciones vacias como archivo sin hojas cuando no hubo movimientos.

Objetivo:
- Tratar esos `.xlsx` vacios de devoluciones como `sin movimientos`, no como error fatal.
- Mantener el rechazo para archivos incorrectos que si tienen hojas pero no contienen `RepLibroDevolucionesGeneral`.
- Dejar `NC REP <marca>` consistente: detalle vacio, `TOTAL GENERAL` en cero, `MAYOR` en cero despues del pipeline completo.
- Evitar arrastrar residuos del mes base de la plantilla.

Criterios de correccion:
1. `nc_stage` debe aceptar workbooks sin hojas solo para fuentes NC y convertirlos en fuente vacia.
2. `my_stage` y `mayor_iva_stage` deben operar con esa fuente vacia sin excepciones y producir grupos NC vacios.
3. La salida final no debe conservar los valores viejos del template en `NC REP PEUG` ni en otras marcas con NC vacio.
4. Debe existir una prueba automatizada que reproduzca `PEUGEOT NC vacio` dentro del proceso completo.
5. Añadir una ayuda visible en la UI para aclarar que una devolucion vacia exportada por origen tambien es aceptada.

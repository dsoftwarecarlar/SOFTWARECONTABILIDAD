# Prompt de trabajo: corrección REP TYT hash mismatch

Analiza y corrige el fallo del pipeline Python nativo de Repuestos cuando `REP TYT` lanza:

- `La hoja REP TYT no conserva los datos del archivo subido. Filas fuente=57, filas salida=57.`

Objetivo:
- Encontrar por que el conteo de filas coincide pero el hash de payload cambia.
- Aislar la transformacion exacta que altera uno o mas valores al copiar la fuente al sheet REP.
- Corregir la escritura para preservar exactamente el payload del archivo subido.
- Dejar la solucion preventiva para otras marcas y futuros archivos con valores atipicos.

Criterios de aceptacion:
- `REP TYT` no falla por mismatch de hash con el archivo afectado.
- La solucion no rompe `REP`, `NC`, `MY` ni la ruta Laravel que invoca Python.
- Queda una reproduccion o prueba local del caso problematico.

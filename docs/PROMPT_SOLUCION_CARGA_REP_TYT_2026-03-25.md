# Prompt de trabajo: solucion carga REP TYT al procesar uploads reales

Reproduce y corrige el error que aparece al cargar archivos reales de Repuestos TYTSERV y procesarlos desde Laravel/Python:

- `La hoja REP TYT no conserva los datos del archivo subido. Filas fuente=57, filas salida=57.`

Reglas:
- Usar los uploads reales mas recientes guardados en `storage/uploads`.
- Validar el `rep_stage` y el `process` completo, no solo fixtures.
- Corregir cualquier perdida de payload provocada por filas estructurales de plantilla o por textos conflictivos.
- Dejar la solucion preventiva para futuras cargas del mismo tipo.

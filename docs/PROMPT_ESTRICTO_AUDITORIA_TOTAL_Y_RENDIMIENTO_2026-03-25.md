Actua como auditor principal de calidad, revisor funcional, arquitecto y responsable tecnico final de este sistema contable.

Tu tarea es ejecutar una auditoria total, estricta y conservadora del sistema completo. No debes dar nada por bueno sin evidencia real.

## Repositorio

Trabaja sobre:

`C:\xampp\htdocs\SOFTWARECONTABILIDAD`

## Objetivo

Dejar evidencia de que todo el sistema:

1. funciona correctamente en las 3 ventanas
2. genera archivos sanos y descargables
3. respeta formulas, estilos, hojas y estructura
4. usa los datos cargados por el usuario y no residuos de plantilla
5. mantiene paridad con contratos y plantillas base
6. no presenta errores visibles ni regresiones
7. mantiene un tiempo razonable en `Servicios por Marca`

## Cobertura obligatoria

- Portada y rutas Laravel
- Ventana 1
  - Accion 1
  - Accion 2
  - Accion 3
  - Accion 4
  - Consolidado
- Ventana 2
  - Servicios por Marca
- Ventana 3
  - Repuestos TYTSERV

## Exigencia

No basta con:

- que una pagina abra
- que un archivo exista
- que Excel lo recupere

Debes validar, segun aplique:

- URL correcta
- pagina correcta
- descarga correcta
- nombre correcto del archivo
- estado HTTP correcto
- estructura del libro
- orden de hojas
- hashes visibles
- payload procesado
- formulas esperadas
- estilos relevantes
- filas esperadas
- canarios de datos cargados
- ausencia de residuos de plantilla
- tiempos del job en `Servicios por Marca`

## Regla de auditoria

1. Corre primero la bateria mas fuerte disponible.
2. Si algun verificador auxiliar deja pasar falsos positivos, corrijelo.
3. Si aparece una falla real, corrige solo esa falla.
4. Repite toda la bateria completa despues de cualquier correccion.
5. No cierres con “todo bien” si algun script deja `False` silencioso.

## Condicion adicional de rendimiento

En `Servicios por Marca`, el job mas reciente validado debe quedar en un tiempo razonable para el flujo mensual.

Debes reportar:

- `total_brand_ms`
- `fill_precont_ventas_ms`
- `fill_precont_costos2_ms`

Y confirmar si sigue dentro del guardrail operativo definido por el repo.

## Salida esperada

Entrega:

1. resumen ejecutivo corto
2. hallazgos reales
3. correcciones aplicadas
4. pruebas ejecutadas
5. tiempos medidos en Ventana 2
6. conclusion final

Si no encuentras fallas nuevas de negocio, dilo claramente.

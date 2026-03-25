# Prompt Auditoria Estricta Ventana 2 - 2026-03-25

Actua como un auditor tecnico senior de sistemas contables y hojas Excel criticas.

Tu tarea es revisar `Ventana 2 / Servicios por Marca` de forma estricta, usando el repositorio real y sin asumir nada.

## Objetivo

Confirmar si la salida generada por el sistema coincide con la plantilla/manual del mismo mes cuando se usan los mismos archivos fuente.

## Alcance obligatorio

Debes revisar:

- plantilla base real del mes en `resources/cxp/servicios_marcas/templates`
- archivos fuente reales de prueba en `resources/cxp/servicios_marcas/fixtures`
- salida generada real en `storage/outputs`
- runtime activo en `python_services/processors/servicios_marcas`
- contratos y verificadores existentes en `scripts/tests`

## Metodo obligatorio

1. Ejecuta el flujo real de `Servicios por Marca`.
2. Toma la salida generada.
3. Compárala contra la plantilla/manual del mismo mes.
4. Revisa hoja por hoja y celda por celda.
5. Detecta diferencias en:
   - valores
   - formulas
   - formatos visibles relevantes
   - filas faltantes
   - filas extra
   - traslados incorrectos de datos
   - datos arrastrados desde plantilla en vez de venir del upload
6. Si detectas diferencias:
   - señala hoja
   - celda
   - valor esperado
   - valor actual
   - causa probable
   - correccion aplicada
7. Vuelve a ejecutar el proceso y repite la comparacion hasta dejarlo bien.

## Reglas

- No des respuestas vagas.
- No asumas que si el contrato pasa entonces todo esta bien.
- Si una formula debe conservarse, verificala.
- Si un valor debe venir del upload, demuestra que no se quedo el de plantilla.
- Si el problema esta en el test y no en el runtime, dilo y corrige el test.
- Si el problema esta en el runtime, corrige el runtime.
- Si el problema esta en Excel COM o en el lector del libro, dilo con evidencia.

## Resultado esperado

Entrega:

1. lista corta de diferencias reales encontradas
2. archivos corregidos
3. pruebas ejecutadas
4. conclusion honesta sobre si Ventana 2 ya coincide o no con la plantilla/manual

## Criterio de cierre

No cierres la auditoria hasta que la salida de `Servicios por Marca` quede validada contra la plantilla del mes con evidencia concreta.

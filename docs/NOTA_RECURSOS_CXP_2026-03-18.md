# Nota de recursos CXP

Fecha: 2026-03-18

## Cambio aplicado

Las plantillas manuales y fixtures de prueba de CXP se movieron fuera de `outputs/` hacia rutas dedicadas en `resources/cxp/`.

## Estructura nueva

- `resources/cxp/acciones/templates`
- `resources/cxp/acciones/fixtures`
- `resources/cxp/servicios_marcas/templates`
- `resources/cxp/servicios_marcas/fixtures`
- `resources/cxp/repuestos_tytserv/templates`
- `resources/cxp/repuestos_tytserv/fixtures`

## Motivo

- `outputs/` quedaba mezclando ejemplos manuales con salidas y staging.
- Las rutas de ejemplo estaban repartidas entre PHP, Node y pruebas E2E.
- La nueva estructura separa mejor:
  - plantillas base
  - fixtures de contrato
  - salidas reales en `storage/outputs`

## Compatibilidad

Las rutas activas ya priorizan `resources/cxp/...` y mantienen fallback temporal a las rutas legacy eliminadas del repo de trabajo. Esto permite conservar compatibilidad de código mientras se termina de limpiar cualquier referencia residual.

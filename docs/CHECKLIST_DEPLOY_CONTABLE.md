# Checklist de despliegue contable

## 1) Pre-deploy
- Confirmar backup de `storage/outputs` y `storage/jobs`.
- Confirmar backup de base de datos de entorno objetivo.
- Verificar version de Node, PHP y Excel (si aplica COM).
- Ejecutar quality gate:
  - `npm run test:quality:contable`
- Validar que no existan procesos activos de servicios/repuestos en produccion.

## 2) Despliegue
- Publicar codigo.
- Limpiar cache de OPCache/PHP-FPM (si aplica).
- Verificar permisos de escritura:
  - `storage/uploads`
  - `storage/outputs`
  - `storage/jobs`

## 3) Smoke post-deploy (obligatorio)
- Cargar archivo de prueba de Tarea 2 y descargar salida.
- Cargar archivo de prueba de Tarea 3 y descargar salida.
- Cargar 4 archivos de Repuestos y descargar salida.
- Confirmar:
  - nombre descargado coincide con mensaje de exito,
  - contenido viene del archivo subido (no de plantilla),
  - historial muestra solo salidas esperadas por accion.

## 4) Validacion contable post-deploy
- Revisar auditorias JSON de salida:
  - `timings_ms`
  - `merge_integrity_ok=true`
  - `verificacion_final_ok=true`
- Revisar que no haya diferencias por cuenta en Tarea 3.
- Validar al menos 1 caso de centavos critico (debe/haber/saldo).

## 5) Criterio de rollback
- Hacer rollback inmediato si ocurre cualquiera:
  - descarga de plantilla en lugar de payload,
  - mismatch de hash/rows post-merge,
  - error contable > 0.01 en validacion por cuenta,
  - bloqueo recurrente de salida en repuestos.

## 6) Monitoreo inicial (primeras 24h)
- Medir tiempos por tarea desde `timings_ms`.
- Revisar errores en consola y logs de jobs.
- Confirmar ausencia de procesos colgados y limpieza de artefactos temporales.

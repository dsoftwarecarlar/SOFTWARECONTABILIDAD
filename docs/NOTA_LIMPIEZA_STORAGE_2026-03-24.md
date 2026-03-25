# Nota de Limpieza Storage 2026-03-24

## Resumen

Se archivaron artefactos de pruebas, bench, probes y verificaciones que ya no forman parte de la operacion diaria del sistema.

Ruta de archivo:

- `archive/storage_artifacts_2026-03-24/storage_root`
- `archive/storage_artifacts_2026-03-24/verify_runs`

## Storage activo despues de la limpieza

Quedan activos solo estos directorios operativos:

- `storage/jobs`
- `storage/outputs`
- `storage/uploads`
- `storage/verify_runs`

`storage/verify_runs` queda vacio por defecto y se vuelve a poblar solo cuando corren contratos o probes tecnicos.

## Criterio aplicado

Se movieron fuera del arbol activo:

- benchmarks `bench_*`
- corridas de rendimiento `perf_runs`
- salidas de validacion manual `outputs_*servicios*`
- verificaciones de ajuste `verify_fix_*`
- logs tecnicos `logs_*.log`, `perftrace_*.log`, `szkdiag_*.log`
- probes y scripts temporales `tmp_*`
- artefactos temporales `temp_*`, `test_*.xlsx`
- contenido historico de `storage/verify_runs`

## Estado tecnico

- La web principal sigue en `laravel_app`
- El procesamiento principal sigue en `python_services`
- Python ya carga `pywin32` desde `python_services/vendor` mediante `python_services/bootstrap.py`
- La ultima excepcion operativa sigue siendo `Excel COM` en `Servicios por Marca`, aunque el worker activo ya corre en Python

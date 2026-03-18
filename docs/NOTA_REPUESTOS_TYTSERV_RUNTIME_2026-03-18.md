# Nota de runtime - Repuestos TYTSERV - 2026-03-18

## Decision
- El runtime productivo del modulo web `cxp_repuestos_tytserv` es Node:
  - `config/cxp/repuestos_tytserv.php`
  - `scripts/cxp/repuestos_tytserv/process.js`
- `run_repuestos_tytserv.ps1` se conserva como fallback/manual probe, no como entrypoint del modulo web.

## Evidencia
- La configuracion activa del modulo apunta a `scripts/cxp/repuestos_tytserv/process.js`.
- No hay referencias de codigo productivo que redirijan el modulo web a `run_repuestos_tytserv.ps1`.
- Se ejecuto manualmente el `.ps1` con los archivos de ejemplo el **2026-03-18** y genero salida valida:
  - `OUTPUT|repuestos_tytserv_ps1_probe_20260318_105232.xlsx|FACTURACION REPUESTOS TYTSERV`

## Regla operativa
- Si se trabaja en el modulo web, asumir Node como runtime oficial.
- Si se necesita diagnostico o comparacion controlada, el `.ps1` puede usarse manualmente.
- No cambiar la configuracion del modulo a PowerShell sin volver a validar contratos E2E.

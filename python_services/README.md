# Python Services

Esta carpeta ya funciona como capa de procesamiento/orquestacion Python para parte del sistema.

## Principios

- Python no reemplaza de golpe a todos los motores heredados.
- Hoy Python ya orquesta:
  - Accion 1-4 y consolidado
  - Repuestos TYTSERV
  - `Servicios por Marca` completo
- `Servicios por Marca` ya corre en Python, pero sigue dependiendo de `Excel COM` en Windows.

## Contrato propuesto

Cada procesador Python debe:

1. recibir un manifiesto JSON con:
   - entradas
   - salida esperada
   - plantilla base
   - opciones
2. escribir el artefacto final en la ruta indicada
3. devolver un payload JSON estructurado cuando el manifiesto define `processor`
4. mantener compatibilidad con el `probe` simple cuando no se define `processor`

## Estado actual

- `vendor/` contiene la dependencia PDF local usada por `Accion 1` (`pdfplumber` y stack asociado)
- `vendor/` tambien queda preparado para dependencias Windows de Excel COM (`pywin32`) mediante `bootstrap.py`
- `processors/cxp_actions/accion1.py` ya es un procesador Python nativo real para PDF
- `processors/cxp_actions/accion2.py` ya es un procesador Python nativo real
- `processors/cxp_actions/accion3.py` ya procesa TXT y PDF de forma nativa
- `processors/cxp_actions/accion4.py` ya es un procesador Python nativo real
- `processors/cxp_actions/export_all.py` ya es un procesador Python nativo real para el consolidado
- `processors/repuestos_tytserv/process.py` ya es el pipeline Python nativo completo de Repuestos
- `processors/repuestos_tytserv/process_legacy.py` conserva el wrapper Node solo para comparaciones y fallback controlado
- `processors/repuestos_tytserv/rep_stage.py` ya implementa la primera etapa nativa de Repuestos para las hojas `REP`
- `processors/repuestos_tytserv/nc_stage.py` ya implementa la etapa nativa de Repuestos para las hojas `NC`
- `processors/repuestos_tytserv/my_stage.py` ya implementa la etapa nativa de Repuestos para las hojas `MY`
- `processors/repuestos_tytserv/mayor_iva_stage.py` ya implementa la etapa nativa de Repuestos para la hoja `MAYOR IVA`
- `processors/servicios_marcas/dispatch.py` ya valida y deja trazabilidad del dispatch de `Servicios por Marca`
- `processors/servicios_marcas/readers.py` ya reemplaza los readers Node de `source`, `px` y `mayor` dentro del flujo activo de `Servicios por Marca`
- `processors/servicios_marcas/runtime.py` ya es el worker Python activo que genera la salida final usando `Excel COM`
- `processors/servicios_marcas/worker.py` ya arranca el runtime Python final, sin depender del worker PowerShell anterior
- `bootstrap.py` centraliza la carga de `vendor/` y de `pywin32_system32`, para que el runtime Python quede listo para COM sin configuracion manual adicional
- `scripts/cxp/servicios_marcas/read_source.js`, `read_px.js` y `read_mayor_txt.js` quedan solo para auditoria/paridad, no como runtime activo
- `processors/legacy_node.py` es el adaptador temporal compartido para Node

## Siguiente objetivo

- decidir cuando retirar `process_legacy.py` del camino normal y dejarlo solo como utilidad de auditoria
- consolidar helpers compartidos de Repuestos para evitar duplicacion entre stages
- decidir si `Servicios por Marca` puede migrar mas alla de `Excel COM` sin perder estabilidad en Windows

# PLANNER GERENCIAL DE AVANCE - CXP

Fecha base: 2026-03-23
Estado: base de seguimiento ejecutivo
Objetivo general: usar este planner para mostrar a jefatura el avance real del desarrollo del sistema `CXP`, con foco en entregables construidos, pruebas ejecutadas, pendientes reales y siguientes pasos.

## 1. Como debe leerse este planner

Este planner no debe centrarse en "documentacion creada".

Debe centrarse en:

- que parte del sistema ya fue construida
- que ventana o modulo ya existe
- que pruebas ya se ejecutaron
- que parte esta cerrada
- que parte sigue en validacion
- que parte sigue pendiente

Regla simple:

Cada card debe poder responder estas 5 preguntas:

1. que se hizo
2. en que parte del sistema impacta
3. como se valido
4. que falta
5. que decision o riesgo sigue abierto

## 2. Sistema real que se esta desarrollando

El sistema confirmado es un portal operativo interno para procesos contables del area de Cuentas por Pagar.

Hoy el sistema opera con:

- una interfaz general del area `CXP`
- `3` ventanas operativas activas
- modulos de procesamiento en `PHP`, `Node.js` y `PowerShell + Excel COM`
- salidas finales en `storage/outputs`

### Ventanas activas confirmadas

1. `Ventana 1 - Libro de Compras ACLT`
   Incluye:
   - `Accion 1 - Libro Compras Proveedores`
   - `Accion 2 - Retenciones Proveedores`
   - `Accion 3 - Mayor Retenciones`
   - `Accion 4 - Mayor IVA`
   - `Consolidado General`

2. `Ventana 2 - Conciliacion Servicios por Marca`
   Incluye:
   - procesamiento por `CHANGAN`
   - procesamiento por `PEUGEOT`
   - procesamiento por `SUZUKI`
   - procesamiento por `MATRIZ`
   - jobs en segundo plano

3. `Ventana 3 - Facturacion Repuestos TYTSERV`
   Incluye:
   - `4` archivos de ventas
   - `4` archivos de devoluciones
   - generacion de salida consolidada mensual

## 3. Regla de comunicacion para jefatura

No usar solo nombres ambiguos como:

- `Tarea 2`
- `Tarea 3`

Porque dentro del proyecto esos nombres pueden referirse a frentes distintos segun el contexto tecnico.

Para jefatura usar siempre:

- `Ventana 1 - Libro de Compras ACLT`
- `Accion 2 - Retenciones Proveedores`
- `Ventana 2 - Servicios por Marca`
- `Ventana 3 - Repuestos TYTSERV`

Eso evita confusion al momento de reportar avances.

## 4. Estado actual del desarrollo

Los porcentajes siguientes son gerenciales. No significan solo codigo escrito. Significan implementacion mas evidencia disponible.

| Frente | Estado actual | % avance | Que ya se hizo | Evidencia base | Que falta |
| --- | --- | --- | --- | --- | --- |
| Interfaz general `CXP` | Validado | 100% | El area principal ya muestra el workspace y sus `3` ventanas activas. | `areas/cxp/index.php`, `includes/app.php` | Mantener alineado con nuevos cambios. |
| Ventana 1 - `Libro de Compras ACLT` | Implementado / validacion continua | 85% | Ya existen la ventana operativa, las `4` acciones y el consolidado general. | `areas/cxp/libro-compras-aclt.php`, `modules/cxp_pdf/index.php`, `modules/cxp_txt/index.php`, `modules/cxp_accion3/index.php`, `modules/cxp_accion4/index.php`, `export_all_actions.php` | Consolidar evidencia final de validacion de punta a punta para cierre gerencial. |
| `Accion 4 - Mayor IVA` | Validado | 100% | Se implemento la generacion y validacion estructural del workbook. | commit `59825d4`, `modules/cxp_accion4/index.php`, `scripts/cxp/accion4/process.js` | Mantenerla dentro de pruebas de regresion. |
| Ventana 2 - `Servicios por Marca` | Apto con observacion | 95% | Se rediseno la UI por marca, el backend valida uploads por marca, el runner elimino archivos legacy y la corrida completa genera `4` salidas. | `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`, `modules/cxp_servicios_marcas/index.php`, `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php` | Solo queda abierta la brecha de espejo estricto en `PrecontabilizacionCostos` si gerencia exige igualdad total con la plantilla manual. |
| Cierre operativo `Servicios por Marca` | Validado | 100% para uso mensual | La corrida mensual de `Servicios por Marca` ya quedo cerrada para uso operativo contable. | `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md`, `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`, `storage/verify_runs` | Solo reabrir si se exige auditoria espejo celda por celda. |
| Ventana 3 - `Repuestos TYTSERV` | Implementado / validacion funcional pendiente | 80% | Ya existe la ventana, el flujo recibe `8` archivos y el runtime web oficial es `Node.js`. | `modules/cxp_repuestos_tytserv/index.php`, `config/cxp/repuestos_tytserv.php`, `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md` | Cerrar evidencia funcional mensual final y mantener pruebas E2E. |
| Pruebas y quality gate | Validado parcial | 85% | Ya existen y se ejecutaron pruebas clave de `Accion 2-3`, `Repuestos`, `Servicios` y `smoke UI`. | `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `package.json`, `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md` | Mantener ejecucion antes de despliegue y completar evidencia final de `Accion 1-4`. |

## 5. Entregables reales ya construidos

Estos son los hitos que si conviene poner en el planner.

### Card 1. Se creo la interfaz general del area CXP

Estado:
- `Validado`

Que se hizo:
- se consolido el area principal del sistema contable
- se muestran las `3` ventanas operativas activas
- se organiza el trabajo por workspace

Evidencia:
- `areas/cxp/index.php`
- `includes/app.php`

### Card 2. Se creo la Ventana 1 - Libro de Compras ACLT

Estado:
- `Implementado`

Que se hizo:
- se habilito la ventana de trabajo para las `4` acciones del proceso
- se incorporo el consolidado general

Evidencia:
- `areas/cxp/libro-compras-aclt.php`
- `modules/cxp_pdf/index.php`
- `modules/cxp_txt/index.php`
- `modules/cxp_accion3/index.php`
- `modules/cxp_accion4/index.php`
- `export_all_actions.php`

### Card 3. Se implemento `Accion 4 - Mayor IVA`

Estado:
- `Validado`

Fecha de referencia:
- `2026-03-13`

Que se hizo:
- se implemento la generacion del workbook
- se incorporo la validacion estructural del archivo final

Evidencia:
- commit `59825d4`
- `modules/cxp_accion4/index.php`
- `scripts/cxp/accion4/process.js`

### Card 4. Se creo la Ventana 2 - Servicios por Marca

Estado:
- `Validado`

Que se hizo:
- se creo la ventana operativa para `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ`
- se habilito procesamiento en segundo plano
- se incorporaron polling, historial y cancelacion

Evidencia:
- `areas/cxp/conciliacion-servicios-marcas.php`
- `modules/cxp_servicios_marcas/index.php`
- `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php`
- `includes/servicios_marcas_job_runner.php`

### Card 5. Se cerro el flujo mensual de Servicios por Marca para uso operativo

Estado:
- `Validado`

Fechas de referencia:
- `2026-03-18`
- `2026-03-19`
- `2026-03-20`

Que se hizo:
- se rediseno la carga por marca
- se eliminaron archivos legacy del flujo
- se validaron las `4` marcas
- se corrigieron hojas criticas y se neutralizo `PrecontabilizacionCostos` legacy

Pruebas y evidencia:
- `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`
- `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md`
- `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`
- `storage/verify_runs/servicios_changan_verifyall_20260320.xls`
- `storage/verify_runs/servicios_peug_verifyall_20260320.xls`
- `storage/verify_runs/servicios_szk_verifyall_20260320.xls`
- `storage/verify_runs/servicios_tyt_verifyall_20260320.xls`

### Card 6. Se creo la Ventana 3 - Repuestos TYTSERV

Estado:
- `Implementado`

Que se hizo:
- se creo la ventana de facturacion mensual de repuestos
- el flujo ya recibe `8` archivos del mes
- se genera una salida consolidada `.xlsx`

Evidencia:
- `areas/cxp/facturacion-repuestos-tytserv.php`
- `modules/cxp_repuestos_tytserv/index.php`
- `src/Cxp/RepuestosTytserv/Application/RepuestosTytservModuleController.php`

### Card 7. Se definio `Node.js` como runtime productivo de Repuestos TYTSERV

Estado:
- `Validado`

Fecha de referencia:
- `2026-03-18`

Que se hizo:
- se confirmo que el modulo web productivo corre con `Node.js`
- `run_repuestos_tytserv.ps1` queda como fallback o prueba manual

Evidencia:
- `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md`
- `config/cxp/repuestos_tytserv.php`
- `scripts/cxp/repuestos_tytserv/process.js`

### Card 8. Se incorporaron pruebas y quality gate del sistema

Estado:
- `Validado parcial`

Fecha de referencia:
- `2026-03-14`

Que se hizo:
- se agregaron pruebas `E2E` y smoke de recursos
- se incorporo el quality gate contable
- se endurecieron controles de integridad

Pruebas ejecutadas con evidencia:
- `npm run test:e2e:accion2-3` -> `OK`
- `npm run test:e2e:repuestos` -> `OK`
- `npm run test:smoke:ui` -> `OK`
- `npm run test:e2e:servicios` -> `OK` para validacion documentada del flujo

Evidencia:
- `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`
- `package.json`
- `scripts/tests/e2e_repuestos_contract.js`
- `scripts/tests/http_resources_smoke.js`
- `scripts/tests/contable_quality_gate.js`

## 6. Pendientes reales para el planner

Estos pendientes si son gerenciales y tecnicos al mismo tiempo.

### Pendiente 1. Cierre final de validacion de Ventana 1

Descripcion:
- dejar evidencia ejecutiva consolidada de `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4` y `Consolidado General`

Estado:
- `En validacion`

### Pendiente 2. Cierre funcional mensual de Repuestos TYTSERV

Descripcion:
- dejar evidencia final de operacion mensual del flujo de `8` archivos con su salida consolidada

Estado:
- `En validacion`

### Pendiente 3. Definir si jefatura exige espejo estricto de `PrecontabilizacionCostos`

Descripcion:
- operativamente `Servicios por Marca` ya esta apto
- solo queda pendiente si se exige igualdad total con la plantilla manual en esa hoja legacy

Estado:
- `Decision requerida`

## 7. Uso exacto en la plantilla de Planner

La plantilla que muestras trabaja por columnas de proyecto, no por estado simple.

Por eso el enfoque correcto es este:

- en cada columna va el tipo de trabajo
- dentro de cada card va el avance real del sistema
- el estado se refleja en el titulo, checklist y notas de la card

Las columnas deben usarse asi:

### Columna `Ambito`

Aqui van cards para definir que parte del sistema entra en el proyecto.

Que cards poner:

- `Definir sistema CXP a desarrollar`
- `Definir ventanas activas del sistema`
- `Definir modulos incluidos en Ventana 1`
- `Definir frentes incluidos en Ventana 2`
- `Definir alcance de Repuestos TYTSERV`

Que escribir dentro de la card:

- que parte del sistema cubre
- que queda dentro
- que queda fuera
- por que ese frente importa para jefatura

Ejemplo de card:

Titulo:
- `Definir alcance del sistema CXP`

Descripcion:
- sistema interno para procesos contables de Cuentas por Pagar
- incluye interfaz general, `Libro de Compras ACLT`, `Servicios por Marca` y `Repuestos TYTSERV`
- excluye historicos de `archive/` como fuente operativa

Checklist:
- confirmar `3` ventanas activas
- confirmar modulos por ventana
- confirmar rutas operativas reales

### Columna `Requisitos de analisis o software`

Aqui van cards para explicar que necesita cada frente para funcionar o validarse.

Que cards poner:

- `Requisitos funcionales de Ventana 1`
- `Requisitos funcionales de Servicios por Marca`
- `Requisitos funcionales de Repuestos TYTSERV`
- `Requisitos tecnicos del entorno`
- `Requisitos de archivos de entrada y salida`

Que escribir dentro de la card:

- archivos requeridos
- reglas del proceso
- dependencias del entorno
- criterios de aceptacion

Ejemplo de card:

Titulo:
- `Requisitos de Servicios por Marca`

Descripcion:
- requiere `2` Excel comunes
- requiere `3` TXT por marca
- procesa `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ`
- genera `1` archivo `.xls` por marca

Checklist:
- validar archivos comunes
- validar TXT por marca
- validar timeout y jobs
- validar salida en `storage/outputs`

### Columna `Diseño`

Aqui van cards sobre como quedo planteada la solucion antes o durante la construccion.

Que cards poner:

- `Diseño de la interfaz general CXP`
- `Diseño de la Ventana 1`
- `Diseño de la carga por marca en Servicios`
- `Diseño del flujo de 8 archivos en Repuestos`
- `Diseño del consolidado general`

Que escribir dentro de la card:

- como se penso la interfaz
- como se dividio el flujo
- que decision de arquitectura se tomo
- que experiencia tendra el usuario

Ejemplo de card:

Titulo:
- `Diseño de la interfaz por marca en Servicios`

Descripcion:
- se separa carga por `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ`
- se mantiene un bloque comun para los `2` Excel base
- se permite procesar una marca o todas

Checklist:
- definir layout de carga
- definir validacion por marca
- definir historial y polling

### Columna `Desarrollo`

Aqui van las cards mas importantes para jefatura.

Aqui debes poner lo que ya se construyo de verdad.

Cards recomendadas:

- `Se creo la interfaz general del area CXP`
- `Se creo la Ventana 1 - Libro de Compras ACLT`
- `Se implemento Accion 4 - Mayor IVA`
- `Se creo la Ventana 2 - Servicios por Marca`
- `Se cerro el flujo mensual de Servicios por Marca`
- `Se creo la Ventana 3 - Repuestos TYTSERV`
- `Se definio Node.js como runtime productivo de Repuestos TYTSERV`

Formato recomendado de cada card:

Titulo:
- empezar con `Se creo`, `Se implemento`, `Se ajusto`, `Se valido`, `Se cerro`

Descripcion:
- que se hizo
- en que modulo o ventana impacta
- que resultado deja

Checklist:
- codigo implementado
- interfaz visible
- flujo operativo disponible
- evidencia asociada

Evidencia sugerida:
- ruta del modulo
- ruta del controller
- reporte o nota tecnica
- commit si aplica

Ejemplo de card:

Titulo:
- `Se creo la Ventana 3 - Repuestos TYTSERV`

Descripcion:
- se creo la ventana operativa de facturacion mensual de repuestos
- el flujo ya recibe `8` archivos del mes
- genera una salida consolidada `.xlsx`

Checklist:
- ventana creada
- formulario de carga habilitado
- validacion de `8` archivos
- salida final configurada

Evidencia:
- `areas/cxp/facturacion-repuestos-tytserv.php`
- `modules/cxp_repuestos_tytserv/index.php`
- `src/Cxp/RepuestosTytserv/Application/RepuestosTytservModuleController.php`

### Columna `Pruebas`

Aqui van cards de validacion real, no solo "hacer pruebas".

Que cards poner:

- `Pruebas E2E Accion 2 y Accion 3 ejecutadas`
- `Pruebas E2E Repuestos ejecutadas`
- `Pruebas de Servicios por Marca ejecutadas`
- `Smoke UI del portal ejecutado`
- `Quality gate contable incorporado`

Que escribir dentro de la card:

- prueba ejecutada
- resultado
- evidencia
- riesgo encontrado
- decision

Ejemplo de card:

Titulo:
- `Pruebas de Servicios por Marca ejecutadas`

Descripcion:
- se valido corrida por marca y corrida completa
- se generaron `4` salidas finales
- el flujo queda apto para uso mensual con observacion controlada

Checklist:
- prueba por marca
- prueba corrida completa
- validacion de hojas criticas
- validacion de archivos generados

Evidencia:
- `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`
- `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md`
- `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`

### Columna `Piloto e implementacion`

Aqui van cards de salida a operacion o validacion con usuario.

Que cards poner:

- `Piloto operativo de Servicios por Marca`
- `Validacion mensual de Repuestos TYTSERV`
- `Validacion funcional final de Ventana 1`
- `Checklist previo a implementacion`
- `Implementacion controlada en entorno contable`

Que escribir dentro de la card:

- que ya esta listo para usarse
- que observacion sigue abierta
- quien valida
- que condicion falta para cierre total

Ejemplo de card:

Titulo:
- `Piloto operativo de Servicios por Marca`

Descripcion:
- el flujo ya se puede usar para operacion mensual
- las `4` marcas generan salida
- solo queda observacion si se exige espejo absoluto de `PrecontabilizacionCostos`

Checklist:
- salida por `CHANGAN`
- salida por `PEUGEOT`
- salida por `SUZUKI`
- salida por `MATRIZ`
- validacion de uso mensual

Evidencia:
- `storage/verify_runs`
- `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`

## 8. Formato exacto de cada card

Para que todas tus cards se vean uniformes en Planner, usa esta estructura:

Titulo:
- una frase corta con resultado

Descripcion:
- `Que se hizo:`
- `Impacto:`
- `Estado actual:`
- `Que falta:`

Checklist:
- `Implementado`
- `Visible en sistema`
- `Validado`
- `Con evidencia`

Campos manuales sugeridos:

- `Responsable`
- `Fecha`
- `% avance`

Plantilla base para copiar:

`Que se hizo:` describir el cambio real aplicado.

`Impacto:` indicar que ventana, modulo o proceso mejora.

`Estado actual:` Implementado / En validacion / Validado / Cerrado.

`Que falta:` siguiente paso concreto.

## 9. Cards listas para crear en tu Planner

### Columna `Ambito`

- `Definir alcance del sistema CXP`
- `Definir ventanas activas del proyecto`
- `Definir modulos incluidos en Libro de Compras ACLT`
- `Definir alcance de Servicios por Marca`
- `Definir alcance de Repuestos TYTSERV`

### Columna `Requisitos de analisis o software`

- `Requisitos funcionales de Ventana 1`
- `Requisitos de archivos para Servicios por Marca`
- `Requisitos del flujo de 8 archivos en Repuestos`
- `Requisitos tecnicos de PHP, Node y Excel`
- `Requisitos de salida y descarga del sistema`

### Columna `Diseño`

- `Diseño de interfaz general CXP`
- `Diseño de Ventana 1 - Libro de Compras ACLT`
- `Diseño de carga por marca en Servicios`
- `Diseño de flujo mensual en Repuestos TYTSERV`
- `Diseño del Consolidado General`

### Columna `Desarrollo`

- `Se creo la interfaz general del area CXP`
- `Se creo la Ventana 1 - Libro de Compras ACLT`
- `Se implemento Accion 4 - Mayor IVA`
- `Se creo la Ventana 2 - Servicios por Marca`
- `Se cerro el flujo mensual de Servicios por Marca`
- `Se creo la Ventana 3 - Repuestos TYTSERV`
- `Se definio Node.js como runtime productivo de Repuestos TYTSERV`

### Columna `Pruebas`

- `Pruebas E2E de Accion 2 y Accion 3 ejecutadas`
- `Pruebas E2E de Repuestos ejecutadas`
- `Pruebas de Servicios por Marca ejecutadas`
- `Smoke UI del portal ejecutado`
- `Quality gate contable incorporado`

### Columna `Piloto e implementacion`

- `Piloto operativo de Servicios por Marca`
- `Validacion mensual de Repuestos TYTSERV`
- `Cierre funcional final de Ventana 1`
- `Checklist previo a implementacion`
- `Implementacion controlada en entorno contable`

## 10. Mensaje sugerido para compartir con jefatura

"Este planner muestra el avance real del desarrollo del sistema CXP usando la estructura de proyecto de software. Cada card representa un entregable, validacion o salida a operacion del sistema, de modo que jefatura pueda revisar que se construyo, que se probo, que ya esta listo y que puntos aun siguen abiertos."

## 11. Cards completas para copiar en Planner

Todas las cards de abajo ya estan pensadas para pegarse dentro de la plantilla de Planner.

Usa siempre esta estructura dentro de la descripcion:

`Que se hizo o define:`

`Impacto:`

`Estado actual:`

`Que falta:`

`Evidencia:`

`Responsable sugerido:`

`% avance sugerido:`

### Columna `Ambito`

#### Card 1. Definir alcance del sistema CXP

Titulo:
- `Definir alcance del sistema CXP`

Contenido:
- `Que se hizo o define:` Se define que el proyecto corresponde al sistema interno `CXP` para procesos contables del area de Cuentas por Pagar.
- `Impacto:` Permite que jefatura vea con claridad que software se esta desarrollando y que parte entra formalmente al proyecto.
- `Estado actual:` Validado.
- `Que falta:` Mantener este alcance alineado cuando se agreguen nuevas ventanas o modulos.
- `Evidencia:` `includes/app.php`, `areas/cxp/index.php`, `README.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- sistema CXP identificado
- area contable identificada
- objetivo general del sistema definido
- historicos fuera del alcance operativo

#### Card 2. Definir ventanas activas del proyecto

Titulo:
- `Definir ventanas activas del proyecto`

Contenido:
- `Que se hizo o define:` Se define que el sistema opera hoy con `3` ventanas activas: `Libro de Compras ACLT`, `Servicios por Marca` y `Repuestos TYTSERV`.
- `Impacto:` Ordena el proyecto por frentes visibles para jefatura y evita mezclar tareas tecnicas sin contexto.
- `Estado actual:` Validado.
- `Que falta:` Mantener esta lista actualizada si se crea una nueva ventana.
- `Evidencia:` `includes/app.php`, `areas/cxp/index.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- Ventana 1 confirmada
- Ventana 2 confirmada
- Ventana 3 confirmada
- nombres de negocio validados

#### Card 3. Definir modulos incluidos en Libro de Compras ACLT

Titulo:
- `Definir modulos incluidos en Libro de Compras ACLT`

Contenido:
- `Que se hizo o define:` Se delimita que `Ventana 1` incluye `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4` y `Consolidado General`.
- `Impacto:` Permite reportar avance de la ventana por accion y no como un bloque ambiguo.
- `Estado actual:` Validado.
- `Que falta:` Consolidar evidencia final de cierre funcional completo de las `4` acciones.
- `Evidencia:` `areas/cxp/libro-compras-aclt.php`, `modules/cxp_pdf/index.php`, `modules/cxp_txt/index.php`, `modules/cxp_accion3/index.php`, `modules/cxp_accion4/index.php`, `export_all_actions.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- Accion 1 incluida
- Accion 2 incluida
- Accion 3 incluida
- Accion 4 incluida
- consolidado incluido

#### Card 4. Definir alcance de Servicios por Marca

Titulo:
- `Definir alcance de Servicios por Marca`

Contenido:
- `Que se hizo o define:` Se define que `Ventana 2` procesa `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ`, con jobs en segundo plano y salida por marca.
- `Impacto:` Permite a jefatura entender que este frente no es una sola tarea sino un flujo mensual multimarcas.
- `Estado actual:` Validado.
- `Que falta:` Mantener el alcance estable y solo reabrir si cambia la regla de auditoria espejo.
- `Evidencia:` `modules/cxp_servicios_marcas/index.php`, `config/cxp/servicios_marcas.php`, `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- marcas definidas
- jobs definidos
- uploads comunes definidos
- uploads por marca definidos

#### Card 5. Definir alcance de Repuestos TYTSERV

Titulo:
- `Definir alcance de Repuestos TYTSERV`

Contenido:
- `Que se hizo o define:` Se define que `Ventana 3` procesa `4` archivos de ventas y `4` de devoluciones para generar una salida consolidada mensual.
- `Impacto:` Evita reportar el modulo como si solo procesara ventas.
- `Estado actual:` Validado.
- `Que falta:` Cerrar evidencia funcional mensual final.
- `Evidencia:` `includes/app.php`, `config/cxp/repuestos_tytserv.php`, `modules/cxp_repuestos_tytserv/index.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- flujo de 8 archivos definido
- salida mensual definida
- runtime oficial definido
- ventana operativa incluida

### Columna `Requisitos de analisis o software`

#### Card 6. Requisitos funcionales de Ventana 1

Titulo:
- `Requisitos funcionales de Ventana 1`

Contenido:
- `Que se hizo o define:` Se documenta que `Ventana 1` necesita archivos de entrada por accion y debe generar outputs contables descargables por accion y consolidado.
- `Impacto:` Permite validar que cada accion cumple un objetivo funcional distinto.
- `Estado actual:` Validado parcial.
- `Que falta:` Consolidar evidencia ejecutiva final por las `4` acciones en una sola vista gerencial.
- `Evidencia:` `docs/MODULOS_CXP.md`, `modules/cxp_pdf/index.php`, `modules/cxp_txt/index.php`, `modules/cxp_accion3/index.php`, `modules/cxp_accion4/index.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `85%`.

Checklist:
- PDF definido para Accion 1
- TXT definido para Accion 2
- lote TXT definido para Accion 3
- TXT definido para Accion 4
- consolidado definido

#### Card 7. Requisitos de archivos para Servicios por Marca

Titulo:
- `Requisitos de archivos para Servicios por Marca`

Contenido:
- `Que se hizo o define:` Se define que el flujo requiere `2` Excel comunes y `3` TXT por marca, con validacion por marca o corrida completa.
- `Impacto:` Reduce errores operativos y deja claro que archivos son obligatorios por ejecucion.
- `Estado actual:` Validado.
- `Que falta:` Mantener checklist mensual con archivos del mes.
- `Evidencia:` `config/cxp/servicios_marcas.php`, `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php`, `docs/CHECKLIST_MENSUAL_TAREA2_SERVICIOS.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- 2 Excel comunes definidos
- 3 TXT por marca definidos
- validacion por marca definida
- corrida completa definida

#### Card 8. Requisitos del flujo de 8 archivos en Repuestos

Titulo:
- `Requisitos del flujo de 8 archivos en Repuestos`

Contenido:
- `Que se hizo o define:` Se define que el flujo mensual requiere `4` archivos `RepLibroVentasGeneral` y `4` archivos `RepLibroDevolucionesGeneral`.
- `Impacto:` Alinea el proceso real con lo que debe entregar el usuario del modulo.
- `Estado actual:` Validado.
- `Que falta:` Mantener evidencia final de una corrida mensual operativa.
- `Evidencia:` `config/cxp/repuestos_tytserv.php`, `src/Cxp/RepuestosTytserv/Application/RepuestosTytservModuleController.php`, `docs/MODULOS_CXP.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- 4 ventas definidas
- 4 devoluciones definidas
- extensiones definidas
- salida consolidada definida

#### Card 9. Requisitos tecnicos de PHP, Node y Excel

Titulo:
- `Requisitos tecnicos de PHP, Node y Excel`

Contenido:
- `Que se hizo o define:` Se confirma que el sistema requiere `PHP` para el portal, `Node.js` para procesamiento contable y `Excel Desktop` para `Servicios por Marca`.
- `Impacto:` Evita fallos de entorno al desplegar o validar modulos.
- `Estado actual:` Validado.
- `Que falta:` Mantenerlo dentro del checklist previo a implementacion.
- `Evidencia:` `docs/OPERACION_CXP.md`, `docs/CHECKLIST_DEPLOY_CONTABLE.md`, `package.json`.
- `Responsable sugerido:` Equipo / Soporte.
- `% avance sugerido:` `100%`.

Checklist:
- PHP requerido
- Node requerido
- Excel Desktop requerido
- permisos de escritura definidos

#### Card 10. Requisitos de salida y descarga del sistema

Titulo:
- `Requisitos de salida y descarga del sistema`

Contenido:
- `Que se hizo o define:` Se establece que las salidas productivas se publican desde `storage/outputs` y se descargan por `download.php`.
- `Impacto:` Evita confundir plantillas o historicos con output real del sistema.
- `Estado actual:` Validado.
- `Que falta:` Mantener control sobre historicos y retencion.
- `Evidencia:` `includes/app.php`, `download.php`, `docs/ARQUITECTURA_CXP.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- outputs definidos
- download definido
- extensiones validas definidas
- historicos fuera de fuente operativa

### Columna `Diseno`

#### Card 11. Diseno de interfaz general CXP

Titulo:
- `Diseno de interfaz general CXP`

Contenido:
- `Que se hizo o define:` Se plantea una interfaz principal que agrupa el area `CXP` y muestra las ventanas operativas del sistema.
- `Impacto:` Facilita navegacion y comunicacion del proyecto a jefatura y usuarios.
- `Estado actual:` Validado.
- `Que falta:` Mantener consistencia si el area crece.
- `Evidencia:` `areas/cxp/index.php`, `templates/cxp/workspace_home.php`, `includes/app.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- home del area definida
- ventanas visibles definidas
- rutas operativas visibles
- organizacion por workspace aplicada

#### Card 12. Diseno de Ventana 1 - Libro de Compras ACLT

Titulo:
- `Diseno de Ventana 1 - Libro de Compras ACLT`

Contenido:
- `Que se hizo o define:` Se organiza la ventana principal de cuentas por pagar alrededor de `4` acciones separadas y un consolidado general.
- `Impacto:` Permite al usuario operar por etapa del proceso y no en una sola pantalla confusa.
- `Estado actual:` Validado.
- `Que falta:` Consolidar evidencia final completa para cierre funcional.
- `Evidencia:` `areas/cxp/libro-compras-aclt.php`, `includes/app.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- acciones separadas
- consolidado incluido
- ventana operativa definida
- accesos por modulo resueltos

#### Card 13. Diseno de la carga por marca en Servicios

Titulo:
- `Diseno de la carga por marca en Servicios`

Contenido:
- `Que se hizo o define:` Se define una interfaz con bloque comun de archivos y bloques por marca para procesar una marca o todas.
- `Impacto:` Mejora claridad operativa y reduce errores al cargar archivos.
- `Estado actual:` Validado.
- `Que falta:` Solo mantenerla estable ante cambios de plantilla o reglas contables.
- `Evidencia:` `templates/cxp/servicios_marcas/index.php`, `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- bloque comun definido
- bloque por marca definido
- opcion una marca o todas definida
- historial y polling definidos

#### Card 14. Diseno del flujo mensual en Repuestos TYTSERV

Titulo:
- `Diseno del flujo mensual en Repuestos TYTSERV`

Contenido:
- `Que se hizo o define:` Se plantea un flujo de carga mensual de `8` archivos con consolidacion automatica en un solo `.xlsx`.
- `Impacto:` Ordena el proceso operativo y deja claro el contrato de entrada del modulo.
- `Estado actual:` Validado.
- `Que falta:` Cerrar evidencia funcional mensual final.
- `Evidencia:` `config/cxp/repuestos_tytserv.php`, `templates/cxp/repuestos_tytserv/index.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- 8 archivos definidos
- template base definida
- salida consolidada definida
- resumen por marca definido

#### Card 15. Diseno del Consolidado General

Titulo:
- `Diseno del Consolidado General`

Contenido:
- `Que se hizo o define:` Se define un consolidado que toma la ultima salida valida de `Accion 1` a `Accion 4` y las une en un solo archivo.
- `Impacto:` Permite un cierre operativo de la ventana `Libro de Compras ACLT`.
- `Estado actual:` Validado.
- `Que falta:` Mantener evidencia final de uso gerencial.
- `Evidencia:` `export_all_actions.php`, `config/cxp/action_exports.json`, `run_export_all_actions.js`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- toma 4 acciones
- valida salidas faltantes
- genera workbook consolidado
- publica descarga final

### Columna `Desarrollo`

#### Card 16. Se creo la interfaz general del area CXP

Titulo:
- `Se creo la interfaz general del area CXP`

Contenido:
- `Que se hizo o define:` Se consolido la interfaz principal del area `CXP` con visibilidad de sus `3` ventanas activas.
- `Impacto:` Jefatura y usuario pueden ver la estructura principal del sistema desde un punto unico.
- `Estado actual:` Validado.
- `Que falta:` Mantenerla alineada si se agregan nuevos frentes.
- `Evidencia:` `areas/cxp/index.php`, `includes/app.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- interfaz creada
- 3 ventanas visibles
- descripciones alineadas con sistema real
- evidencia disponible

#### Card 17. Se creo la Ventana 1 - Libro de Compras ACLT

Titulo:
- `Se creo la Ventana 1 - Libro de Compras ACLT`

Contenido:
- `Que se hizo o define:` Se habilito la ventana principal para ejecutar `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4` y `Consolidado General`.
- `Impacto:` El frente principal de Cuentas por Pagar queda operativo por etapas de proceso.
- `Estado actual:` Implementado.
- `Que falta:` Consolidar evidencia final de validacion completa de punta a punta.
- `Evidencia:` `areas/cxp/libro-compras-aclt.php`, `modules/cxp_pdf/index.php`, `modules/cxp_txt/index.php`, `modules/cxp_accion3/index.php`, `modules/cxp_accion4/index.php`, `export_all_actions.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `85%`.

Checklist:
- ventana creada
- acciones visibles
- consolidado visible
- modulos conectados

#### Card 18. Se implemento Accion 4 - Mayor IVA

Titulo:
- `Se implemento Accion 4 - Mayor IVA`

Contenido:
- `Que se hizo o define:` Se implemento la generacion y validacion estructural del workbook de `Accion 4`.
- `Impacto:` La cuarta accion del flujo ACLT queda construida y usable dentro del sistema.
- `Estado actual:` Validado.
- `Que falta:` Mantenerla dentro de pruebas de regresion.
- `Evidencia:` commit `59825d4`, `modules/cxp_accion4/index.php`, `scripts/cxp/accion4/process.js`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- modulo implementado
- workbook generado
- validacion estructural aplicada
- evidencia disponible

#### Card 19. Se creo la Ventana 2 - Servicios por Marca

Titulo:
- `Se creo la Ventana 2 - Servicios por Marca`

Contenido:
- `Que se hizo o define:` Se creo la ventana operativa para `CHANGAN`, `PEUGEOT`, `SUZUKI` y `MATRIZ`, con jobs en segundo plano, polling, historial y cancelacion.
- `Impacto:` El proceso mensual de servicios queda separado y operable como un frente propio del sistema.
- `Estado actual:` Validado.
- `Que falta:` Mantener criterio de cierre segun necesidad de auditoria espejo.
- `Evidencia:` `areas/cxp/conciliacion-servicios-marcas.php`, `modules/cxp_servicios_marcas/index.php`, `src/Cxp/ServiciosMarcas/Application/ServiciosMarcasModuleController.php`, `includes/servicios_marcas_job_runner.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- ventana creada
- carga por marca habilitada
- jobs habilitados
- historial habilitado

#### Card 20. Se cerro el flujo mensual de Servicios por Marca

Titulo:
- `Se cerro el flujo mensual de Servicios por Marca`

Contenido:
- `Que se hizo o define:` Se rediseño la carga por marca, se eliminaron archivos legacy del flujo y se validaron las `4` marcas con salida final operativa.
- `Impacto:` El modulo queda apto para uso mensual contable.
- `Estado actual:` Validado para uso mensual.
- `Que falta:` Solo reabrir si gerencia exige espejo absoluto de `PrecontabilizacionCostos`.
- `Evidencia:` `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`, `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md`, `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`, `storage/verify_runs`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- CHANGAN validado
- PEUGEOT validado
- SUZUKI validado
- MATRIZ validado
- corrida completa validada

#### Card 21. Se creo la Ventana 3 - Repuestos TYTSERV

Titulo:
- `Se creo la Ventana 3 - Repuestos TYTSERV`

Contenido:
- `Que se hizo o define:` Se creo la ventana operativa de facturacion mensual de repuestos con carga de `8` archivos y generacion de una salida consolidada `.xlsx`.
- `Impacto:` El proceso de repuestos queda integrado dentro del portal `CXP`.
- `Estado actual:` Implementado.
- `Que falta:` Cerrar evidencia funcional mensual final.
- `Evidencia:` `areas/cxp/facturacion-repuestos-tytserv.php`, `modules/cxp_repuestos_tytserv/index.php`, `src/Cxp/RepuestosTytserv/Application/RepuestosTytservModuleController.php`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `80%`.

Checklist:
- ventana creada
- formulario de carga activo
- flujo de 8 archivos activo
- salida consolidada configurada

#### Card 22. Se definio Node.js como runtime productivo de Repuestos TYTSERV

Titulo:
- `Se definio Node.js como runtime productivo de Repuestos TYTSERV`

Contenido:
- `Que se hizo o define:` Se confirmo que el modulo web oficial corre con `Node.js` y que `run_repuestos_tytserv.ps1` queda como fallback o prueba manual.
- `Impacto:` Evita confusiones tecnicas sobre el runtime productivo del modulo.
- `Estado actual:` Validado.
- `Que falta:` Mantener esa decision mientras no cambie el runtime oficial.
- `Evidencia:` `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md`, `config/cxp/repuestos_tytserv.php`, `scripts/cxp/repuestos_tytserv/process.js`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- runtime Node confirmado
- fallback PS1 aclarado
- config productiva validada
- evidencia disponible

### Columna `Pruebas`

#### Card 23. Pruebas E2E de Accion 2 y Accion 3 ejecutadas

Titulo:
- `Pruebas E2E de Accion 2 y Accion 3 ejecutadas`

Contenido:
- `Que se hizo o define:` Se ejecutaron pruebas contractuales para validar integridad y salida de `Accion 2` y `Accion 3`.
- `Impacto:` Reduce riesgo de regresion en transformaciones contables criticas.
- `Estado actual:` Validado.
- `Que falta:` Mantener ejecucion antes de despliegue.
- `Evidencia:` `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `package.json`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- contrato E2E definido
- prueba ejecutada
- resultado OK documentado
- evidencia disponible

#### Card 24. Pruebas E2E de Repuestos ejecutadas

Titulo:
- `Pruebas E2E de Repuestos ejecutadas`

Contenido:
- `Que se hizo o define:` Se agrego y ejecuto una prueba contractual para el flujo de repuestos.
- `Impacto:` Reduce riesgo de publicar archivos incorrectos o salidas engañosas en el modulo de repuestos.
- `Estado actual:` Validado.
- `Que falta:` Mantener la prueba en el quality gate y cerrar evidencia mensual final.
- `Evidencia:` `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `scripts/tests/e2e_repuestos_contract.js`, `package.json`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- prueba creada
- prueba ejecutada
- resultado OK documentado
- evidencia disponible

#### Card 25. Pruebas de Servicios por Marca ejecutadas

Titulo:
- `Pruebas de Servicios por Marca ejecutadas`

Contenido:
- `Que se hizo o define:` Se validaron corrida por marca y corrida completa, con generacion de `4` salidas y verificacion de hojas criticas.
- `Impacto:` Demuestra que el flujo esta apto para uso mensual contable.
- `Estado actual:` Validado.
- `Que falta:` Solo reabrir si se exige espejo absoluto de la hoja legacy.
- `Evidencia:` `docs/NOTA_SERVICIOS_MARCAS_TAREA2_2026-03-18.md`, `docs/REPORTE_AUDITORIA_FINAL_TAREA2_2026-03-19.md`, `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`, `storage/verify_runs`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- corrida por marca
- corrida completa
- validacion de hojas criticas
- validacion de archivos generados

#### Card 26. Smoke UI del portal ejecutado

Titulo:
- `Smoke UI del portal ejecutado`

Contenido:
- `Que se hizo o define:` Se incorporo y ejecuto una validacion rapida de vistas y recursos estaticos del portal.
- `Impacto:` Reduce riesgo de errores visibles en rutas principales del sistema.
- `Estado actual:` Validado.
- `Que falta:` Mantener smoke antes de despliegue.
- `Evidencia:` `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `scripts/tests/http_resources_smoke.js`, `package.json`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `100%`.

Checklist:
- smoke definido
- smoke ejecutado
- rutas principales incluidas
- resultado documentado

#### Card 27. Quality gate contable incorporado

Titulo:
- `Quality gate contable incorporado`

Contenido:
- `Que se hizo o define:` Se incorporo un quality gate que agrupa validaciones criticas del sistema contable.
- `Impacto:` Da un punto unico de control antes de despliegue o cambios sensibles.
- `Estado actual:` Validado parcial.
- `Que falta:` Mantenerlo como paso obligatorio y completar evidencia final de `Ventana 1`.
- `Evidencia:` `docs/REPORTE_AUDITORIA_CONTABLE_2026-03-14.md`, `scripts/tests/contable_quality_gate.js`, `package.json`.
- `Responsable sugerido:` Equipo.
- `% avance sugerido:` `85%`.

Checklist:
- quality gate creado
- pruebas agregadas al gate
- gate documentado
- uso previo a deploy recomendado

### Columna `Piloto e implementacion`

#### Card 28. Piloto operativo de Servicios por Marca

Titulo:
- `Piloto operativo de Servicios por Marca`

Contenido:
- `Que se hizo o define:` El flujo ya puede usarse para operacion mensual y genera salida valida para las `4` marcas.
- `Impacto:` El modulo ya entrega valor operativo real al area contable.
- `Estado actual:` Cerrado para uso mensual.
- `Que falta:` Solo una decision gerencial si se exige espejo absoluto de `PrecontabilizacionCostos`.
- `Evidencia:` `docs/REPORTE_CIERRE_TAREA2_2026-03-20.md`, `storage/verify_runs`.
- `Responsable sugerido:` Equipo / Usuario contable.
- `% avance sugerido:` `100%`.

Checklist:
- salida CHANGAN
- salida PEUGEOT
- salida SUZUKI
- salida MATRIZ
- criterio mensual validado

#### Card 29. Validacion mensual de Repuestos TYTSERV

Titulo:
- `Validacion mensual de Repuestos TYTSERV`

Contenido:
- `Que se hizo o define:` El modulo ya esta implementado y queda pendiente cerrar una corrida mensual validada con evidencia ejecutiva.
- `Impacto:` Esta card permite mostrar que el frente existe, pero aun necesita cierre funcional mensual.
- `Estado actual:` En validacion.
- `Que falta:` Ejecutar corrida mensual, validar salida consolidada y dejar evidencia de uso real.
- `Evidencia:` `modules/cxp_repuestos_tytserv/index.php`, `config/cxp/repuestos_tytserv.php`, `docs/NOTA_REPUESTOS_TYTSERV_RUNTIME_2026-03-18.md`.
- `Responsable sugerido:` Equipo / Usuario contable.
- `% avance sugerido:` `80%`.

Checklist:
- corrida mensual ejecutada
- salida final validada
- descarga correcta validada
- evidencia registrada

#### Card 30. Cierre funcional final de Ventana 1

Titulo:
- `Cierre funcional final de Ventana 1`

Contenido:
- `Que se hizo o define:` La ventana ya existe con sus `4` acciones y consolidado, pero falta consolidar evidencia ejecutiva final de uso completo.
- `Impacto:` Esta card representa el cierre gerencial del frente principal de cuentas por pagar.
- `Estado actual:` En validacion.
- `Que falta:` Reunir evidencia de `Accion 1`, `Accion 2`, `Accion 3`, `Accion 4` y `Consolidado General`.
- `Evidencia:` `areas/cxp/libro-compras-aclt.php`, `modules/cxp_pdf/index.php`, `modules/cxp_txt/index.php`, `modules/cxp_accion3/index.php`, `modules/cxp_accion4/index.php`, `export_all_actions.php`.
- `Responsable sugerido:` Equipo / Usuario contable.
- `% avance sugerido:` `85%`.

Checklist:
- Accion 1 validada
- Accion 2 validada
- Accion 3 validada
- Accion 4 validada
- consolidado validado

#### Card 31. Checklist previo a implementacion

Titulo:
- `Checklist previo a implementacion`

Contenido:
- `Que se hizo o define:` Se establece el paso previo obligatorio de respaldo, validacion de entorno, pruebas y smoke antes de publicar cambios.
- `Impacto:` Reduce riesgo de afectar operacion contable en produccion.
- `Estado actual:` Validado.
- `Que falta:` Ejecutarlo como regla en cada implementacion.
- `Evidencia:` `docs/CHECKLIST_DEPLOY_CONTABLE.md`, `docs/OPERACION_CXP.md`.
- `Responsable sugerido:` Equipo / Soporte.
- `% avance sugerido:` `100%`.

Checklist:
- backup realizado
- entorno validado
- quality gate ejecutado
- smoke post deploy definido

#### Card 32. Implementacion controlada en entorno contable

Titulo:
- `Implementacion controlada en entorno contable`

Contenido:
- `Que se hizo o define:` Se plantea una salida controlada a operacion con validacion posterior del area contable.
- `Impacto:` Permite que jefatura vea el paso final entre desarrollo y uso real.
- `Estado actual:` En progreso.
- `Que falta:` Completar cierres finales de `Ventana 1` y `Repuestos TYTSERV` y ejecutar el checklist de implementacion.
- `Evidencia:` `docs/CHECKLIST_DEPLOY_CONTABLE.md`, `docs/OPERACION_CXP.md`, `docs/BITACORA_AVANCES.md`.
- `Responsable sugerido:` Equipo / Soporte / Usuario contable.
- `% avance sugerido:` `70%`.

Checklist:
- checklist ejecutado
- despliegue realizado
- validacion contable realizada
- observaciones registradas

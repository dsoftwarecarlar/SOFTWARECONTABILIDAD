# Prompt De Verificacion Laravel + Python

Usa este prompt para pedir una auditoria tecnica real del estado del proyecto y confirmar si la migracion va bien, si la arquitectura sigue sana y si hay errores ocultos o regresiones.

```text
Actua como un staff engineer, arquitecto de software y auditor tecnico senior. Quiero que hagas una verificacion real, dura y sin complacencia del estado actual de este proyecto.

## Repositorio
Trabaja sobre este repo real:
C:\xampp\htdocs\SOFTWARECONTABILIDAD

## Objetivo
Verificar si la migracion progresiva hacia Laravel + Python va bien, si la arquitectura actual sigue coherente, y si existen errores, regresiones, deudas tecnicas o decisiones riesgosas que haya que corregir de inmediato.

No quiero una opinion superficial. Quiero una auditoria tecnica real basada en codigo, configuracion, pruebas y ejecucion.

## Contexto que debes respetar
La direccion arquitectonica del proyecto es:

- Laravel como capa web y de orquestacion
- Python como motor futuro de procesamiento
- Node y PowerShell se conservan temporalmente donde aun son necesarios
- Excel COM en Windows sigue vivo en `Servicios por Marca`

Estado esperado actual:
- `Accion 2` ya corre como `Laravel -> Python nativo`
- `Accion 1`, `Accion 3`, `Accion 4` y `consolidado-acciones` corren como `Laravel -> Python -> Node`
- `Repuestos TYTSERV` corre como `Laravel -> Python -> Node`
- `Servicios por Marca` corre como `Laravel -> job runner heredado -> PowerShell/Excel COM`

Rutas relevantes:
- `laravel_app/`
- `python_services/`
- `resources/cxp/`
- `storage/outputs`
- `storage/uploads`
- `storage/jobs`

Documentacion que debes revisar:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\docs\MIGRACION_LARAVEL_PYTHON_2026-03-24.md`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\laravel_app\README.md`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\python_services\README.md`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\docs\CHECKLIST_DEPLOY_CONTABLE.md`

## Que debes hacer

### 1. Auditar estado real del codigo
Revisa:
- arquitectura actual
- configuracion Laravel
- bridge Python
- procesadores actuales
- integracion con Node
- integracion con PowerShell
- rutas, controladores y vistas de Laravel
- pruebas existentes

Determina:
- que esta bien
- que esta mal
- que esta inconsistente
- que esta duplicado
- que sigue siendo deuda tecnica
- que ya puede considerarse estable

### 2. Verificar si la migracion va bien
Quiero que respondas con criterio tecnico:
- si la direccion Laravel + Python esta realmente bien aplicada o no
- si la capa Laravel ya esta cumpliendo su rol correctamente
- si Python ya esta entrando de forma sana o solo como parche
- si el sistema va mejorando o si se esta complicando innecesariamente

### 3. Detectar errores reales
No teorices. Busca errores concretos:
- fallos de integracion
- contratos E2E rotos
- problemas de compatibilidad de Excel
- problemas de rutas o descargas
- supuestos falsos en pruebas
- documentos o config desalineados con el codigo real
- puntos donde el sistema pueda romperse en produccion

### 4. Ejecutar validaciones reales
Debes correr y evaluar como minimo:

- `node scripts/tests/contable_quality_gate.js`
- `powershell -ExecutionPolicy Bypass -File .\scripts\dev\test_laravel_app.ps1`
- `powershell -ExecutionPolicy Bypass -File .\scripts\dev\test_laravel_window1.ps1`
- `powershell -ExecutionPolicy Bypass -File .\scripts\dev\test_laravel_repuestos.ps1`
- `powershell -ExecutionPolicy Bypass -File .\scripts\dev\test_laravel_servicios.ps1`

Si alguna validacion falla:
- explica exactamente que fallo
- identifica la causa real
- corrige si el riesgo es razonable
- vuelve a validar

### 5. Revisar coherencia arquitectonica
Quiero que me digas claramente:
- si Laravel ya esta quedando como plataforma principal
- si Python ya esta entrando como capa de procesamiento de forma correcta
- si hay algo que contradiga esa arquitectura
- si hay partes que deberian quedarse legacy por ahora
- cual es el siguiente modulo correcto para migrar a Python nativo

### 6. Revisar si hay errores silenciosos
Busca especialmente:
- archivos generados que luego no pueden reabrirse por otras librerias
- diferencias entre el flujo Laravel y el flujo legacy
- historiales que apunten a archivos incorrectos
- tests que pasen por casualidad y no por contrato real
- metadatos o docs que ya no coincidan con la implementacion

## Formato obligatorio de respuesta

### A. Veredicto general
- vamos bien / vamos regular / vamos mal
- explicacion directa y corta

### B. Hallazgos
Enumera hallazgos reales, ordenados por severidad:
- alta
- media
- baja

Para cada hallazgo incluye:
- problema
- impacto
- archivo o modulo afectado
- si requiere correccion inmediata o no

### C. Validaciones ejecutadas
Lista:
- que corriste
- que paso
- que fallo
- que corregiste si aplicaba

### D. Estado de la arquitectura
Resume:
- que ya esta bien en Laravel
- que ya esta bien en Python
- que sigue dependiendo de Node
- que sigue dependiendo de PowerShell/Excel COM

### E. Riesgos vigentes
Di cuales son los riesgos reales que aun quedan.

### F. Siguiente paso recomendado
Dime exactamente cual debe ser el siguiente bloque de trabajo.

## Reglas
- No me tranquilices si hay problemas.
- No digas “todo bien” sin verificar.
- No des teoria generica.
- No opines sin leer el codigo y ejecutar validaciones.
- Si algo esta mal, dilo con claridad.
- Si algo esta bien, dilo con la misma claridad.
- Si descubres una regresion, prioriza identificar la causa real antes de proponer reescrituras.

## Instruccion final
Haz la auditoria completa del estado actual del proyecto y responde como si fueras el responsable tecnico que va a firmar el avance de esta migracion.
```

## Uso recomendado

Puedes usar este prompt:

- antes de seguir migrando otro modulo
- despues de cada bloque grande de cambios
- antes de dar por estable una fase
- antes de reemplazar otro worker Node por Python nativo

Actua como un arquitecto de software senior, auditor de migraciones y revisor UX/QA para un sistema contable interno.

Trabaja sobre este proyecto real:

`C:\xampp\htdocs\SOFTWARECONTABILIDAD`

Tu tarea es verificar y completar el cierre de la migracion para que el sistema quede realmente consolidado en:

- Laravel como capa web principal
- Python como motor principal de procesamiento

No quiero teoria. Quiero una auditoria aplicada al codigo real y, si faltan piezas, quiero que las implementes.

## Objetivo

Determina si la migracion ya quedo completa o que falta para que:

1. la web principal quede en Laravel
2. el procesamiento principal quede en Python
3. la estructura antigua deje de ser necesaria
4. la experiencia visual quede limpia, clara y lista para usuarios de contabilidad
5. todas las ventanas, acciones y descargas sigan funcionando

## Restricciones

- No des respuestas complacientes.
- No asumas: inspecciona el repo.
- No expliques frameworks al usuario final.
- El texto visible debe estar orientado a negocio, no a tecnologia.
- El logo correcto visible es `logo2`.
- Si detectas partes antiguas innecesarias, propon o ejecuta su retiro.
- Si algo legacy aun es indispensable, dilo con precision.

## Lo que debes revisar

### 1. Migracion tecnica real
- Verifica si la entrada principal ya es Laravel o si todavia depende de entrypoints PHP viejos.
- Verifica si todos los procesos principales ya estan en Python o si aun dependen de:
  - PHP standalone
  - Node
  - PowerShell / Excel COM
- Identifica exactamente que piezas siguen fuera de Laravel + Python.

### 2. Superficie antigua
- Revisa si aun existen y se usan:
  - `areas/`
  - `modules/`
  - `templates/cxp/`
  - `src/` legacy
  - `download.php`
  - `export_all_actions.php`
  - `run_servicios_marcas_job.php`
  - helpers heredados en `includes/`
- Distingue entre:
  - estructura antigua que aun es necesaria
  - estructura antigua que ya solo estorba

### 3. Parte visual
- Revisa que la interfaz en `laravel_app/resources/views` este ordenada, consistente y clara.
- Verifica:
  - portada
  - area CXP
  - ventanas
  - acciones
  - repuestos
  - servicios
- Confirma que el usuario entiende:
  - que subir
  - que obtiene
  - donde descargar
  - que hacer si un proceso sigue en curso
- Confirma que no se vean terminos tecnicos como:
  - Laravel
  - Python
  - runtime
  - legacy
  - worker
  - pipeline
  - dispatch

### 4. URLs
- Verifica que todas las URLs internas apunten bien.
- Verifica que assets, logos, descargas y enlaces internos respondan correctamente.
- Verifica que no queden enlaces visibles a rutas viejas `.php`, `areas/` o `modules/`.

### 5. Flujos funcionales
Debes validar como minimo:
- Inicio
- CXP
- Libro Compras ACLT
- Accion 1
- Accion 2
- Accion 3
- Accion 4
- Consolidado de acciones
- Repuestos TYTSERV
- Servicios por Marca

### 6. Validaciones obligatorias
Corre validaciones reales:
- `node scripts/tests/http_resources_smoke.js`
- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios`

Si encuentras una falla, corrige la causa real antes de cerrar.

## Formato de respuesta esperado

1. Hallazgos
- Problemas reales ordenados por severidad

2. Estado de migracion
- Que ya quedo verdaderamente en Laravel + Python
- Que aun no

3. Superficie antigua
- Que se puede eliminar ya
- Que se debe archivar
- Que aun debe quedarse

4. UX y visual
- Si la interfaz ya esta lista para usuarios contables
- Que detalles visuales o de copy aun faltan

5. Recomendacion final
- Di si ya cerrarias la migracion o que bloque queda pendiente

## Instruccion final
No te limites a revisar. Si falta algo pequeño o mediano para cerrar bien la migracion, implementalo y vuelve a validar.

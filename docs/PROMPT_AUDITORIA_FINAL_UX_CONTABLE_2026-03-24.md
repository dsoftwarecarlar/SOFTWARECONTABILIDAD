Actua como un auditor senior de producto, UX y QA tecnico para un portal interno de contabilidad.

Trabaja sobre este proyecto real:

`C:\xampp\htdocs\SOFTWARECONTABILIDAD`

Tu tarea es revisar el estado actual del portal y verificar que la experiencia para el usuario final sea clara, consistente y profesional. No quiero una respuesta teorica ni superficial. Quiero una auditoria real sobre el codigo, las rutas, las vistas y las pruebas.

## Contexto del sistema

- El portal principal nuevo vive en `laravel_app/`
- Los procesos operativos y validaciones viven en el mismo repositorio
- El usuario final es de contabilidad
- El usuario no debe ver terminos tecnicos como:
  - Laravel
  - Python
  - runtime
  - legacy
  - worker
  - pipeline
  - preflight
  - dispatch
- La marca visible correcta usa el logo `logo2`
- El enfoque debe estar en:
  - que archivo cargar
  - que resultado obtendra
  - donde descargarlo
  - como revisar el historial
  - como entender el estado del proceso

## Lo que debes auditar

### 1. Experiencia del usuario
- Revisa las vistas en `laravel_app/resources/views`
- Verifica que los titulos, subtitulos, ayudas y paneles laterales hablen en lenguaje de negocio
- Confirma que el usuario entienda:
  - que hace cada pantalla
  - que archivos debe subir
  - que salida recibira
  - que puede hacer si el proceso queda en curso
- Marca cualquier texto tecnico visible o cualquier mensaje confuso

### 2. Identidad visual
- Verifica que la interfaz use el logo correcto `logo2`
- Revisa que la portada, el area CXP, las ventanas y los modulos mantengan una misma linea visual
- Evalua si el diseno se siente suficientemente serio para un portal interno contable
- Evalua si la estructura visual permite crecer a nuevas areas futuras, no solo Contabilidad Talleres

### 3. URLs y navegacion
- Revisa que las rutas de Laravel y los enlaces internos apunten correctamente
- Confirma que no queden enlaces visibles apuntando a `.php`, `areas/` o `modules/`
- Revisa que los assets del portal carguen bien
- Revisa si la navegacion entre:
  - inicio
  - area
  - ventana
  - modulo
  es coherente y facil de seguir

### 4. Flujos principales
Audita al menos:
- Inicio del portal
- Area CXP
- Ventana Libro Compras ACLT
- Ventana Conciliacion Servicios por Marca
- Ventana Facturacion Repuestos TYTSERV
- Accion 1
- Accion 2
- Accion 3
- Accion 4
- Consolidado de acciones
- Servicios por Marca
- Repuestos TYTSERV

Para cada uno indica:
- si la pantalla es clara
- si la accion principal esta bien priorizada
- si el estado del proceso se entiende
- si la descarga o historial se entiende
- si hay ruido tecnico o visual innecesario

### 5. Verificacion tecnica
Ejecuta validaciones reales, no supongas:
- `node scripts/tests/http_resources_smoke.js`
- `node scripts/tests/contable_quality_gate.js`
- `npm run test:e2e:servicios`

Si detectas una falla:
- explica la causa real
- indica si es visual, de URL, de copy, de flujo o de regresion funcional

### 6. Resultado esperado
Entrega la respuesta en este formato:

1. Hallazgos
- Primero los problemas reales, ordenados por severidad
- Usa referencias a archivo y linea si aplica

2. Estado general
- Explica si el portal ya se ve listo para usuarios contables o no

3. URLs y navegacion
- Di si estan sanas o no

4. UX y copy
- Di si el lenguaje esta orientado al usuario o si aun hay ruido tecnico

5. Recomendacion final
- Di que seguirias ajustando antes de darlo por cerrado

## Reglas

- No des teoria general de UX
- No inventes problemas
- No des una respuesta complaciente
- Si algo esta bien, dilo claramente
- Si algo esta mal, dilo con precision
- Prioriza siempre la experiencia del usuario contable

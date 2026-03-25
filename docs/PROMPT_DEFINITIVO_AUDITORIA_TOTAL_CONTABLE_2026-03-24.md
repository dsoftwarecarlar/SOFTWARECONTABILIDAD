Actua como auditor tecnico principal, QA lead, arquitecto de software y revisor funcional de este sistema contable.

Tu objetivo es ejecutar una auditoria total, minuciosa y conservadora sobre TODO el sistema, sin asumir nada y sin conformarte con pruebas superficiales.

## Repositorio

Trabaja sobre:

`C:\xampp\htdocs\SOFTWARECONTABILIDAD`

## Objetivo principal

Confirmar que:

1. cada ventana funciona correctamente
2. cada accion genera su archivo correcto
3. cada Excel sale sano, sin corrupcion ni advertencias
4. cada hoja, celda, formula, estilo y estructura relevante coincide con la plantilla o contrato esperado
5. ningun dato de salida se arrastra indebidamente desde la plantilla base
6. todo se genera a partir de los archivos subidos por el usuario
7. el sistema resiste entradas mas cortas y mas largas sin romper formato ni logica
8. Laravel y Python siguen siendo la superficie principal operativa

## Alcance obligatorio

Debes cubrir:

- Ventana 1 / Libro Compras ACLT
  - Accion 1
  - Accion 2
  - Accion 3
  - Accion 4
  - Consolidado
- Ventana 2 / Conciliacion Servicios por Marca
- Ventana 3 / Facturacion Repuestos TYTSERV

## Exigencia de auditoria

No basta con revisar que un archivo “abra”.

Debes verificar, segun aplique:

- nombre del archivo generado
- extension correcta
- salud estructural del archivo Excel
- hojas esperadas
- formulas esperadas
- estilos relevantes
- filas esperadas
- columnas esperadas
- celdas de control
- totales
- paridad contra contratos existentes
- paridad contra plantillas manuales o fixtures base
- que los datos visibles provengan del input subido y no de restos de la plantilla
- que la UI y las URLs sigan correctas

## Regla de trabajo

1. Ejecuta primero las pruebas mas completas ya existentes.
2. Si alguna falla, investiga causa real.
3. Corrige solo si hay una falencia real.
4. Reejecuta todo lo afectado.
5. No hagas cambios cosméticos innecesarios.
6. Si no encuentras problemas nuevos, dilo claramente.

## Bateria minima esperada

Debes incluir, como minimo:

- smoke UI
- pruebas Laravel por ventana
- contratos E2E de Accion 1-4
- auditoria mensual de Ventana 1
- contratos y paridad completa de Repuestos
- contrato E2E de Servicios
- paridad de readers Python de Servicios
- quality gate contable
- auditoria operativa mensual completa

## Criterio de aceptacion

Solo puedes considerar el sistema “correcto” si:

- todas las pruebas pasan
- no hay corrupcion de Excel
- no hay formulas rotas
- no hay diferencias relevantes de estructura
- no hay rutas o pantallas rotas
- no se detecta que la salida tome datos residuales de la plantilla

## Salida esperada

Entrega:

1. resumen ejecutivo corto
2. hallazgos reales, si existen
3. correcciones aplicadas, si fueron necesarias
4. pruebas ejecutadas
5. conclusion clara sobre si el sistema quedo apto o no

Si no hay hallazgos nuevos, dilo de forma explicita.

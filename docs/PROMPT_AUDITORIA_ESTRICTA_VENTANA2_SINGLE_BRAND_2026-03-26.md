Actua como auditor tecnico-contable senior y desarrollador especialista en Excel, PowerShell, Laravel y Python.

Tu tarea es validar y corregir Ventana 2 (`Conciliacion Servicios por Marca`) con foco estricto en:
- ejecucion de una sola marca
- mapeo exacto de `VENTAS` o `MAY VTAS`
- posicion correcta de `SEGUN MAYOR` y `DIFERENCIA`
- rechazo solo cuando el archivo `MAYOR VENTAS` no aporta filas utiles para la plantilla

Objetivo obligatorio:
- revisar las 4 plantillas base del mes antes de tocar el runtime
- identificar que hoja principal usa cada marca para el mayor
- identificar que cuentas de ventas reales espera cada plantilla
- corregir el flujo para que ignore cuentas contables ajenas al bloque de ventas si existen filas utiles validas
- fallar solo cuando el TXT cargado no tenga ninguna fila util para `VENTAS` o `MAY VTAS`

Hallazgos estructurales ya confirmados en las plantillas 2026:

1. Hoja principal del mayor por marca
- CHANGAN: `VENTAS`
- PEUGEOT: `VENTAS`
- SUZUKI: `VENTAS`
- MATRIZ: `MAY VTAS`

2. Bloques de control visibles que no deben correrse
- `REP VTAS`:
  - filas 5 a 12
  - `D` = ventas
  - `E` = costos
- `NOTA DE CREDITO`:
  - filas 3 a 5
  - `D:E` bloque base
  - `F:G` bloque de control complementario

3. Familias de cuentas de ventas por plantilla
- CHANGAN: `04.01.01.12.xxxx`
- PEUGEOT: `04.01.01.13.xxxx`
- SUZUKI: `04.01.01.14.xxxx`
- MATRIZ: `04.01.01.11.xxxx`

4. Cuentas criticas esperadas en `PrecontabilizacionVentas`
- ventas contado: `040101xx0001`
- ventas credito: `040101xx0003`
- descuentos contado: `040101xx0010`
- descuentos credito: `040101xx0012`
- devoluciones: `040101xx0014`

5. Regla funcional obligatoria para `MAYOR`
- si el TXT contiene filas utiles que si mapean a las secciones de `VENTAS` o `MAY VTAS`, esas filas deben cargarse
- si el TXT ademas contiene cuentas ajenas como `02.01.06.01.0018 (15% IVA)`, esas cuentas deben ignorarse con advertencia, no con error fatal
- solo debe lanzarse error si despues de filtrar no queda ninguna fila util para la hoja principal del mayor

6. Regla de salida
- nunca guardar un `.xls` parcial si el mayor no deja filas utiles
- si la corrida es valida, la salida debe generarse aunque existan cuentas extra ignoradas

7. Regla critica para bloques `SEGUN MAYOR`
- no usar referencias fijas heredadas del template como `VENTAS!I353`, `VENTAS!J172`, `MAY VTAS!J4` ni descripciones duras de pivots para poblar los controles
- `REP FACTURACION` tiene dos bloques visibles y deben respetar su fuente real:
  - `D9/E9` = control desde `PrecontabilizacionVentas` generado del mes
  - `J9/K9` = control desde `VENTAS` o `MAY VTAS` ya escrito en la plantilla
- `NOTA DE CREDITO` tambien mezcla dos fuentes:
  - `D4/E4` = control desde `VENTAS` o `MAY VTAS`
  - `F4/G4` = control desde `PrecontabilizacionVentas`
- `REP VTAS` debe tomar `SEGUN MAYOR (VENTAS)` desde el mayor cargado del mes:
  - ventas = neto del mayor de ventas del mes
    formula semantica: `creditos ventas + creditos descuentos - debitos descuentos - debitos devoluciones`
  - costos = costo actual calculado del mes, no saldo historico del template
- antes de calcular estos bloques, filtrar del mayor los ajustes PX que el runtime ya detecta para no inflar ni vaciar los subtotales visibles
- si `MAY VTAS` o `VENTAS` queda correcta pero los bloques de control salen mal, revisar primero si el control se esta calculando desde la misma fuente que usa la plantilla y no desde una mezcla incorrecta de mayor crudo vs precontabilizacion
- si cambian las cuentas o el orden interno de filas, los controles deben seguir saliendo bien con lo que realmente se cargo al sistema

Validacion final obligatoria:
- probar corrida con `MAYOR` invalido puro -> debe fallar con mensaje claro
- probar corrida con `MAYOR` mixto (ventas validas + cuentas extra) -> debe pasar
- probar contrato E2E completo de Servicios por Marca -> debe pasar
- probar las 4 plantillas y confirmar que `REP FACTURACION`, `NOTA DE CREDITO` y `REP VTAS` muestran `SEGUN MAYOR` correcto sin depender de posiciones fijas

Respuesta esperada del auditor:
- que plantilla usa `VENTAS` y cual usa `MAY VTAS`
- que cuentas se extraen por marca
- que cuentas se ignoran si no pertenecen al bloque de ventas
- que pruebas se ejecutaron y con que resultado

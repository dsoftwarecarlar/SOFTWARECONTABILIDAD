Actua como auditor tecnico-contable senior y desarrollador especialista en Excel, Laravel, PowerShell y Python.

Tu tarea es verificar y corregir exclusivamente los bloques `SEGUN MAYOR` de `Ventana 2 / Servicios por Marca`.

Objetivo obligatorio:
- demostrar de donde sale cada valor visible de `SEGUN MAYOR`
- usar una corrida real con archivos subidos al sistema, no una simulacion abstracta
- distinguir entre:
  - error real de logica
  - diferencia causada por un `MAYOR` cargado incompleto o distinto

Hojas y celdas obligatorias:
- `REP FACTURACION`
  - `D9/E9`
  - `J9/K9`
- `NOTA DE CREDITO`
  - `D4/E4`
  - `F4/G4`
- `REP VTAS`
  - `D6/E6`
- `VENTAS` o `MAY VTAS`
  - `J4`
  - `I172/J172`
  - `I245/J245`
  - `I353/J353`

Reglas tecnicas:
1. `REP FACTURACION`
   - `D9/E9` salen de `PrecontabilizacionVentas` generado del mes
   - `J9/K9` salen del mayor ya escrito en `VENTAS` o `MAY VTAS`
2. `NOTA DE CREDITO`
   - `D4/E4` salen del mayor ya escrito en `VENTAS` o `MAY VTAS`
   - `F4/G4` salen de `PrecontabilizacionVentas`
3. `REP VTAS`
   - `D6` sale de ventas netas del mayor
   - `E6` sale del costo real del mes
4. Si `VENTAS` o `MAY VTAS` esta bien y el control visible coincide con esa hoja, no se debe declarar bug aunque el usuario espere otro valor.
5. Si el `MAYOR` subido no contiene cuentas de descuentos o devoluciones, los bloques `SEGUN MAYOR` correspondientes deben quedar vacios o en cero. El sistema no debe inventar esos valores desde otro archivo.
6. Solo declarar bug si:
   - el `MAYOR` cargado si contiene la cuenta/familia correspondiente
   - pero la hoja `VENTAS` o `MAY VTAS` no la refleja
   - o la hoja la refleja y el control visible no coincide con esa hoja

Metodo obligatorio:
1. Tomar el `job` real mas reciente reportado por el usuario.
2. Abrir:
   - job json
   - TXT subidos
   - `.xls` generado
3. Calcular manualmente desde el `MAYOR`:
   - ventas
   - descuentos
   - devoluciones
   - ventas netas
4. Comparar contra:
   - `VENTAS` o `MAY VTAS`
   - controles visibles
5. Si el problema es el archivo cargado, decirlo con evidencia concreta:
   - nombre del archivo
   - cuentas encontradas
   - cuentas ausentes
   - impacto exacto en celdas visibles

Respuesta esperada:
- origen exacto de cada valor discutido
- si el dato esta bien o mal respecto al archivo subido
- si la correccion es de codigo o de archivo fuente

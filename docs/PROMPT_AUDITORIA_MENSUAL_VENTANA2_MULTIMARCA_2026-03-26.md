Actua como auditor tecnico-contable senior y desarrollador especialista en Excel, Laravel, PowerShell y Python.

Tu tarea es auditar `Ventana 2 / Conciliacion Servicios por Marca` usando los archivos reales guardados en:
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\fixtures`
- `C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\servicios_marcas\templates`

Objetivo obligatorio:
- ejecutar la logica real con los archivos de febrero disponibles en el repositorio
- validar que la salida dependa de los archivos subidos y no de residuos de plantilla
- confirmar que la misma logica siga sirviendo para meses futuros mientras la estructura de plantilla no cambie
- revisar las 4 marcas: `CHANGAN`, `PEUGEOT`, `SUZUKI`, `MATRIZ`

Archivos base obligatorios:
- comun: `detalle-vtas-xliquidar (2).xlsx`
- comun: `RepFacturacionServContabilidad (3).xls`
- CHANGAN:
  - `SERREP_FACTURAS_NAFCHAN.TXT`
  - `SERREP_NOTACRED_NAFCHAN.TXT`
  - `CON_MAYORGEN2CHAN.TXT`
- PEUGEOT:
  - `SERREP_FACTURAS_NAFPEU.TXT`
  - `SERREP_NOTACRED_NAFPEU.TXT`
  - `CON_MAYORGEN2PEU.TXT`
- SUZUKI:
  - `SERREP_FACTURAS_NAFSUZAMBYRIO.TXT`
  - `SERREP_NOTACRED_NAFSUZAMBYRI.TXT`
  - `CON_MAYORGEN2SUZ.TXT`
- MATRIZ:
  - `SERREP_FACTURAS_NAFTOY.TXT`
  - `SERREP_NOTACRED_NAFTOY.TXT`
  - `CON_MAYORGEN2TOY.TXT`

Reglas tecnicas obligatorias:
1. La plantilla aporta estructura, formulas y formato. Los datos visibles del mes deben salir del upload.
2. No asumir que por ser archivos de febrero la logica cambia. La logica contable debe depender del contenido, no del nombre del mes.
3. Validar por marca que la hoja principal del mayor sea correcta:
   - CHANGAN -> `VENTAS`
   - PEUGEOT -> `VENTAS`
   - SUZUKI -> `VENTAS`
   - MATRIZ -> `MAY VTAS`
4. Validar que el mayor procese lo cargado aunque cambie la cuenta dentro de la misma familia de ventas, descuentos o devoluciones.
5. Ignorar cuentas ajenas al bloque de ventas si coexisten con filas utiles. Fallar solo si no queda ninguna fila util para poblar `VENTAS` o `MAY VTAS`.
6. Auditar minimo estas hojas:
   - `REP FACTURACION`
   - `NOTA DE CREDITO`
   - `PX`
   - `REP VTAS`
   - `VENTAS` o `MAY VTAS`
   - `PrecontabilizacionVentas`
   - `COSTO`
   - `ESTADISTICAS`

Bloques de control que deben cuadrar:
- `REP FACTURACION`
  - `D9/E9` desde `PrecontabilizacionVentas`
  - `J9/K9` desde `VENTAS` o `MAY VTAS`
- `NOTA DE CREDITO`
  - `D4/E4` desde `VENTAS` o `MAY VTAS`
  - `F4/G4` desde `PrecontabilizacionVentas`
- `REP VTAS`
  - `D6` desde ventas netas del mayor del mes
  - `E6` desde costo real del mes

Pruebas obligatorias:
1. Ejecutar corrida real de las 4 marcas con los fixtures del repositorio.
2. Confirmar que no existan `#REF!`, `#N/A` ni valores historicos del template en hojas visibles.
3. Comparar los controles visibles contra la data realmente cargada desde TXT/XLS/XLSX.
4. Si una salida vieja no cuadra pero una salida nueva si, clasificarla como evidencia historica y no como falla vigente.
5. Si aparece una diferencia real, corregir el runtime y repetir la corrida completa.

Respuesta esperada del auditor:
- que archivos se usaron
- que pruebas reales se ejecutaron
- que celdas visibles se validaron
- si la salida actual depende del upload y no de la plantilla
- si queda alguna brecha real pendiente

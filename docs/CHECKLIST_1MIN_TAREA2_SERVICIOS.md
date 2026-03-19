# Checklist 1 Minuto - Tarea 2 Servicios

Usar solo para el cierre mensual de `Servicios por Marca`.

## 1. Archivos a subir
- Excel comun: `detalle-vtas-xliquidar.xlsx`
- Excel comun: `RepFacturacionServContabilidad.xls`
- TXT marca: `SERREP_FACTURAS_*.TXT`
- TXT marca: `SERREP_NOTACRED_*.TXT`
- TXT marca: `CON_MAYORGEN2*.TXT`

## 2. Antes de correr
- Cerrar todos los archivos de Excel visibles.
- Confirmar que la marca correcta esta seleccionada, o dejar vacio para todas.
- Confirmar que los archivos del mes corresponden al mismo periodo.

## 3. Despues de generar
- Verificar que salga el `.xls` de la marca esperada en `C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\outputs`.
- Abrir el archivo y revisar rapido:
  - `REP FACTURACION`
  - `NOTA DE CREDITO`
  - `PX`
  - `REP VTAS`
  - `VENTAS` o `MAY VTAS`
- Confirmar que no existan `#REF!`, `#N/A` ni fechas viejas del template.

## 4. Rechazar si pasa cualquiera de estos casos
- Falta una hoja esperada.
- Hay valores historicos que no son del mes.
- Hay formulas rotas.
- La marca del encabezado no coincide.
- Los totales visibles no cuadran con el mes.

## 5. Regla final
- La plantilla solo aporta formato, formulas y estructura.
- Los datos del mes deben venir de los archivos subidos.

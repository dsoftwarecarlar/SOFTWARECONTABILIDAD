# Manual de Usuario
Portal Contable - Contabilidad Talleres

Version del manual: 2026-03-31
Documento para uso interno del area contable

[[PAGE_BREAK]]

## 1. Objetivo del manual
Este manual explica como ingresar al portal, elegir la ventana correcta, cargar archivos, descargar resultados y resolver las validaciones mas comunes del sistema.

El documento corresponde al flujo actual del portal operativo de **Contabilidad Talleres** y cubre los tres frentes de trabajo disponibles:
- **Libro Compras ACLT**
- **Conciliacion Servicios por Marca**
- **Facturacion Repuestos TYTSERV**

## 2. Acceso al sistema
Para ingresar al sistema:
1. Abra un navegador web.
2. Ingrese a la ruta interna del portal: **/SOFTWARECONTABILIDAD/**.
3. Si el acceso es por red local, use la IP o el nombre del servidor seguido de **/SOFTWARECONTABILIDAD/**.
4. En la portada, haga clic en **Entrar a Contabilidad Talleres**.

Importante:
- El usuario no necesita entrar manualmente a **laravel_app/public**.
- El portal se usa desde una sola entrada web.
- Se recomienda tener **Microsoft Excel** disponible para revisar los archivos descargados.

## 3. Estructura general del portal
El sistema esta organizado en tres niveles:
1. **Inicio**: portada principal del portal.
2. **Area Contabilidad Talleres**: punto de entrada a las ventanas operativas.
3. **Ventanas y procesos**: pantallas especificas para cada tarea.

Las ventanas disponibles son:
- **Libro Compras ACLT**
- **Conciliacion Servicios por Marca**
- **Facturacion Repuestos TYTSERV**

En cada proceso el flujo base es el mismo:
1. Entrar a la pantalla correcta.
2. Cargar los archivos requeridos.
3. Ejecutar el proceso.
4. Descargar el resultado.
5. Revisar el historial reciente.

## 4. Recomendaciones antes de procesar
Antes de iniciar cualquier carga:
- Tenga listos todos los archivos del mes o del corte que va a procesar.
- Verifique que el archivo corresponda al proceso correcto.
- No mezcle archivos de marcas o acciones distintas.
- Descargue el resultado apenas termine la ejecucion.
- Guarde una copia local del archivo final si necesita conservarlo mas tiempo.

Nota operativa:
- El historial del portal muestra solo archivos recientes.
- No conviene usar el historial como archivo permanente.

## 5. Ventana: Libro Compras ACLT
Esta ventana contiene cinco procesos:
1. **Accion 1 - Libro Compras Proveedores**
2. **Accion 2 - Retenciones Proveedores**
3. **Accion 3 - Mayor Retenciones**
4. **Accion 4 - Mayor IVA**
5. **Consolidado de Acciones**

### 5.1 Accion 1 - Libro Compras Proveedores
Resumen del proceso:
- Archivo requerido: **1 PDF**
- Ejemplo de nombre: **CXPREP_docproveedor**
- Resultado: **Excel**
- Historial visible: **ultimas 3 salidas**

Pasos de uso:
1. Abra **Libro Compras ACLT**.
2. Ingrese a **Accion 1 - Libro Compras Proveedores**.
3. Haga clic en el campo de carga.
4. Seleccione el archivo PDF del proveedor.
5. Pulse **Procesar archivo**.
6. Espere a que aparezca la seccion **Salida generada**.
7. Haga clic en **Descargar Excel**.

Consideraciones:
- Solo se acepta formato **PDF**.
- Si intenta subir otro formato, el sistema mostrara un error de formato no permitido.

### 5.2 Accion 2 - Retenciones Proveedores
Resumen del proceso:
- Archivo requerido: **1 TXT**
- Ejemplo de nombre: **CXPREP_RET_GENERALACCION2**
- Resultado: **Excel RET PROV**
- Historial visible: **ultimas 3 salidas**

Pasos de uso:
1. Abra **Accion 2 - Retenciones Proveedores**.
2. Cargue el archivo TXT de retenciones.
3. Pulse **Procesar archivo**.
4. Espere la generacion del resultado.
5. Descargue el Excel generado desde la misma pantalla.

Consideraciones:
- Solo se acepta formato **TXT**.
- El resultado suele identificarse con sufijo **_accion2.xlsx**.

### 5.3 Accion 3 - Mayor Retenciones
Resumen del proceso:
- Archivo requerido: **uno o varios TXT** o **un PDF**
- Ejemplo de nombre: **CON_MAYORGEN2ACCION3**
- Resultado: **Excel MAYOR RET**
- Historial visible: **ultimas 3 salidas**

Pasos de uso:
1. Abra **Accion 3 - Mayor Retenciones**.
2. Cargue uno o varios archivos TXT del mayor general, o un PDF del mismo reporte.
3. Si va a cargar varios archivos, seleccione todos en la misma carga.
4. Pulse **Procesar archivo**.
5. Espere a que aparezca el resultado.
6. Descargue el Excel consolidado.

Consideraciones:
- Esta accion permite **carga multiple**.
- Si sube varios TXT, el sistema genera una sola salida consolidada.
- El resultado suele identificarse con sufijo **_accion3.xlsx**.

### 5.4 Accion 4 - Mayor IVA
Resumen del proceso:
- Archivo requerido: **1 TXT**
- Ejemplo de nombre: **CON_MAYORGEN2IVAACCION4**
- Resultado: **Excel MAYOR IVA**
- Historial visible: **ultimas 3 salidas**

Pasos de uso:
1. Abra **Accion 4 - Mayor IVA**.
2. Cargue el archivo TXT correspondiente.
3. Pulse **Procesar archivo**.
4. Espere la generacion del Excel.
5. Descargue el archivo final.

Consideraciones:
- Solo acepta formato **TXT**.
- El resultado suele identificarse con sufijo **_accion4.xlsx**.

### 5.5 Consolidado de Acciones
Este proceso genera un solo archivo usando las ultimas salidas disponibles de las acciones 1, 2, 3 y 4.

Resumen del proceso:
- No requiere carga manual de archivos.
- Requiere que las acciones 1, 2, 3 y 4 ya tengan salida reciente.
- Resultado: **1 Excel consolidado**
- Historial visible: **ultimos 3 consolidados**

Pasos de uso:
1. Verifique que ya existan resultados recientes de **Accion 1**, **Accion 2**, **Accion 3** y **Accion 4**.
2. Abra **Consolidado de Acciones**.
3. Revise la seccion que muestra la cobertura por proceso.
4. Pulse **Construir consolidado**.
5. Espere la generacion del archivo final.
6. Descargue el consolidado desde la misma pantalla.

Consideraciones:
- Si falta una salida previa, el sistema no podra construir el consolidado.
- El archivo generado suele identificarse como **acciones_resumen_...xlsx**.

## 6. Ventana: Conciliacion Servicios por Marca
Esta ventana procesa informacion de cuatro marcas:
- **CHANGAN**
- **PEUGEOT**
- **SUZUKI**
- **MATRIZ**

El resultado final es:
- **1 archivo .xls por marca procesada**

### 6.1 Archivos requeridos
Archivos comunes obligatorios para cualquier ejecucion:
- **PX**  
  Ejemplo: **detalle-vtas-xliquidar**
- **REP VENTAS**  
  Ejemplo: **RepFacturacionServContabilidad**

Archivos TXT por marca:
- **CHANGAN**
  - REP FACTURACION: **SERREP_FACTURAS_NAFCHAN**
  - NOTA DE CREDITO: **SERREP_NOTACRED_NAFCHAN**
  - MAYOR VENTAS: **CON_MAYORGEN2CHAN**
- **PEUGEOT**
  - REP FACTURACION: **SERREP_FACTURAS_NAFPEU**
  - NOTA DE CREDITO: **SERREP_NOTACRED_NAFPEU**
  - MAYOR VENTAS: **CON_MAYORGEN2PEU**
- **SUZUKI**
  - REP FACTURACION: **SERREP_FACTURAS_NAFSUZAMBYRIO**
  - NOTA DE CREDITO: **SERREP_NOTACRED_NAFSUZAMBYRI**
  - MAYOR VENTAS: **CON_MAYORGEN2SUZ**
- **MATRIZ**
  - REP FACTURACION: **SERREP_FACTURAS_NAFTOY**
  - NOTA DE CREDITO: **SERREP_NOTACRED_NAFTOY**
  - MAYOR VENTAS: **CON_MAYORGEN2TOY**

Formatos aceptados:
- **PX** y **REP VENTAS**: **.xls** o **.xlsx**
- Archivos por marca: **.txt**

### 6.2 Procesar todas las marcas
Use este flujo cuando va a generar los cuatro resultados del mes.

Pasos de uso:
1. Abra **Conciliacion Servicios por Marca**.
2. En **Marca a procesar**, deje seleccionada la opcion **Todas las marcas**.
3. Cargue los 2 Excel comunes.
4. Cargue los 12 TXT de las cuatro marcas.
5. Pulse **Procesar y generar plantillas**.
6. Deje la pagina abierta mientras el sistema procesa.
7. Espere a que la pantalla se actualice automaticamente.
8. Descargue cada archivo generado desde la seccion de resultados.

### 6.3 Procesar una sola marca
Use este flujo cuando solo necesita una marca.

Pasos de uso:
1. Abra **Conciliacion Servicios por Marca**.
2. En **Marca a procesar**, seleccione la marca requerida.
3. Cargue los 2 Excel comunes.
4. Cargue solamente los 3 TXT de la marca seleccionada.
5. Pulse **Procesar y generar plantillas**.
6. Espere el fin del proceso y descargue la salida.

Importante:
- Aunque procese una sola marca, los **2 Excel comunes** siguen siendo obligatorios.
- El sistema ajusta automaticamente que TXT son obligatorios segun la marca elegida.

### 6.4 Estado del proceso
Durante la ejecucion la pantalla puede mostrar:
- **Preparando archivos**
- **Procesando**
- **Deteniendo**
- **Terminado**
- **Con error**

Comportamiento operativo:
- Solo puede existir **un proceso activo** a la vez en esta ventana.
- La pagina consulta el estado automaticamente cada pocos segundos.
- Existe un boton **Detener proceso activo** para pedir una detencion ordenada.
- Si se detiene un proceso, la siguiente ejecucion debe iniciarse de nuevo.

### 6.5 Validaciones importantes
El sistema valida lo siguiente:
- Que se carguen todos los archivos obligatorios.
- Que el formato del archivo sea correcto.
- Que el archivo **MAYOR VENTAS** corresponda a la marca seleccionada.

Observacion clave:
- El TXT de **MAYOR VENTAS** debe incluir cuentas **04.01.01.xx.xxxx**.
- Si el archivo no corresponde a la marca o no contiene el patron esperado, el sistema mostrara un error y no continuara.

### 6.6 Resultados e historial
Despues del proceso:
- Se genera **1 archivo .xls por marca procesada**.
- Los nombres de salida suelen iniciar con:
  - **servicios_changan_**
  - **servicios_peug_**
  - **servicios_szk_**
  - **servicios_tyt_**
- El historial muestra el archivo mas reciente disponible por marca.

Recomendacion:
- Descargue cada resultado apenas finalice.
- No deje el historial como unico respaldo del trabajo terminado.

## 7. Ventana: Facturacion Repuestos TYTSERV
Esta ventana genera un solo libro final mensual a partir de ventas y devoluciones de cuatro marcas.

Resultado esperado:
- **1 archivo Excel final**

### 7.1 Archivos requeridos
Debe cargar **8 Excel** en total:
- **MATRIZ - Ventas**
- **MATRIZ - Devoluciones**
- **PEUGEOT - Ventas**
- **PEUGEOT - Devoluciones**
- **CHANGAN - Ventas**
- **CHANGAN - Devoluciones**
- **SUZUKI - Ventas**
- **SUZUKI - Devoluciones**

Formatos aceptados:
- **.xls**
- **.xlsx**

Ejemplos de origen:
- Ventas: **RepLibroVentasGeneral**
- Devoluciones: **RepLibroDevolucionesGeneral**

### 7.2 Pasos de uso
1. Abra **Facturacion Repuestos TYTSERV**.
2. Ingrese al proceso **Repuestos TYTSERV**.
3. Cargue los 8 Excel del mes.
4. Pulse **Procesar y generar reporte**.
5. Espere a que aparezca la seccion **Archivo listo**.
6. Descargue el archivo final desde la misma pantalla.

### 7.3 Consideraciones operativas
- El sistema requiere los 8 archivos para ejecutar el proceso completo.
- Si el origen exporta devoluciones vacias como un archivo Excel sin hojas, el sistema tambien puede aceptarlo.
- La pantalla puede mostrar un resumen de filas procesadas por marca.
- La pantalla puede mostrar verificaciones internas del libro final.

### 7.4 Historial
El historial de esta ventana muestra los reportes mas recientes del modulo.

Identificacion de salida:
- El archivo final suele comenzar con **repuestos_tytserv_** y terminar en **.xlsx**.

## 8. Descargas e historial
En todas las ventanas:
- El resultado puede descargarse desde el boton de descarga de la misma pantalla.
- Tambien existe un historial reciente para volver a descargar archivos recientes.

Buenas practicas con los resultados:
1. Descargue el archivo apenas este disponible.
2. Revise que el nombre corresponda al proceso correcto.
3. Abra el archivo en Excel y haga una validacion rapida.
4. Guarde una copia en la ubicacion interna definida por su equipo.

## 9. Mensajes comunes y como resolverlos
### 9.1 Formato no permitido
Posible causa:
- Se intento cargar un archivo con extension distinta a la permitida.

Accion recomendada:
- Revise si el proceso pide **PDF**, **TXT**, **XLS** o **XLSX**.
- Vuelva a cargar el archivo correcto.

### 9.2 No se recibieron archivos para ejecutar la accion
Posible causa:
- El archivo no fue seleccionado antes de presionar el boton.

Accion recomendada:
- Seleccione el archivo requerido y repita la ejecucion.

### 9.3 Faltan archivos generados para el consolidado
Posible causa:
- Falta al menos una salida previa de las acciones 1, 2, 3 o 4.

Accion recomendada:
- Genere primero la accion faltante y luego vuelva a construir el consolidado.

### 9.4 Ya existe un proceso en ejecucion o cierre
Posible causa:
- En **Servicios por Marca** ya hay un proceso activo, en cola o deteniendose.

Accion recomendada:
- Espere a que termine.
- Si corresponde, use **Detener proceso activo** y luego vuelva a intentar.

### 9.5 Falta el archivo de una marca o de un campo obligatorio
Posible causa:
- No se cargaron todos los insumos requeridos.

Accion recomendada:
- Revise la pantalla y cargue cada archivo solicitado antes de procesar.

### 9.6 El archivo de MAYOR VENTAS no corresponde a la marca seleccionada
Posible causa:
- El TXT pertenece a otra marca o no contiene las cuentas esperadas.

Accion recomendada:
- Revise la marca elegida.
- Verifique el origen del TXT.
- Cargue nuevamente el archivo correcto.

### 9.7 El proceso termino sin generar el archivo de salida
Posible causa:
- Hubo un fallo interno durante la generacion.

Accion recomendada:
- Repita la ejecucion con los mismos archivos una sola vez.
- Si el error persiste, reporte el caso a soporte interno con el nombre del modulo y el mensaje mostrado.

## 10. Buenas practicas de uso
- Procese un solo juego de archivos por ejecucion.
- No mezcle archivos de diferentes meses si el procedimiento no lo exige.
- Revise el nombre de cada archivo antes de subirlo.
- En **Servicios por Marca**, deje la pagina abierta hasta que finalice.
- Descargue y guarde los resultados al terminar.
- Si detecta un error repetitivo, documente el mensaje exacto.

## 11. Que informacion enviar a soporte interno
Si necesita reportar una incidencia, envie como minimo:
- Nombre de la ventana
- Nombre del proceso
- Fecha y hora aproximada
- Nombre del archivo cargado
- Captura del mensaje de error
- Indique si el problema ocurrio en la primera o segunda ejecucion

## 12. Cierre
El portal esta pensado para que el usuario siempre haga tres cosas con claridad:
1. Entrar a la pantalla correcta.
2. Cargar el archivo correcto.
3. Descargar el resultado correcto.

Si se respeta esa secuencia, el trabajo diario del area contable se vuelve mas rapido, mas ordenado y con menos errores de operacion.

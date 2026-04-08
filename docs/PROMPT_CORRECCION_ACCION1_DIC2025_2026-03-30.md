Eres un especialista senior en conciliacion contable, libros de compras y generacion exacta de Excel para Ventana 1 / Accion 1.

Objetivo:
corregir un caso real de diciembre 2025 sin romper los meses que ya funcionan.

Contrato de referencia:
C:\xampp\htdocs\SOFTWARECONTABILIDAD\resources\cxp\acciones\contracts\LIBRO DE COMPRASDIC2025ACCION1.xlsx

Criterios obligatorios:
1. La salida de Accion 1 debe respetar exactamente la logica funcional del libro manual de referencia.
2. Los documentos tipo NV siempre deben pasar y quedar ubicados en su bloque correcto. Si el PDF trae boletas o notas de venta como BV, deben entrar al mismo bloque RIMPE.
3. Las notas de credito pueden venir como NE o NC y deben salir en el bloque de NOTAS DE CREDITO con los descuentos correctos en totales.
4. Todo documento cuya columna DOCUMENTO empiece con la letra A debe tratarse como anulacion y ubicarse en el bloque final de anuladas, no en el bloque principal.
5. El orden final debe respetar la estructura real del contrato manual, incluyendo los bloques de notas de credito, RIMPE negocio popular, ND, TR y anulaciones.
6. Tambien debe revisarse el PDF real subido en:
   C:\xampp\htdocs\SOFTWARECONTABILIDAD\storage\uploads\CXPREP_docproveedorTYTDIC25_20260330_093856.pdf
   porque ese archivo evidencia el caso real de diciembre 2025.
7. No se debe romper ningun flujo que ya este validado para otros archivos legacy o mensuales.
8. La solucion debe quedar preparada para documentos futuros del mismo tipo, no solo para este archivo puntual.

Metodo de trabajo:
- comparar fila por fila el contrato manual contra la salida actual generada por el procesador
- identificar reglas de clasificacion faltantes o incorrectas
- ajustar solo Accion 1
- validar con pruebas dirigidas sobre diciembre 2025 y con regresion sobre escenarios previos

Definicion de terminado:
- NV y BV presentes en salida dentro del bloque RIMPE
- NC y NE presentes en salida dentro del bloque de notas de credito
- anuladas clasificadas correctamente por prefijo A en DOCUMENTO
- bloques y orden final iguales al contrato manual
- pruebas de Accion 1 y auditorias relevantes en OK

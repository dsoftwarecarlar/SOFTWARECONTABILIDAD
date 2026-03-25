Actua como un auditor tecnico senior y responsable de calidad operativa de un sistema contable mensual.

Trabaja sobre este proyecto real:

- `C:\xampp\htdocs\SOFTWARECONTABILIDAD`

Tu tarea es verificar si el sistema ya esta listo para uso mensual real por el equipo contable, sin asumir nada y validando con ejecucion real.

## Objetivo

Confirmar que:

1. la interfaz visible para el usuario final es profesional, clara y sin lenguaje tecnico innecesario;
2. el sistema puede abrirse desde otra maquina en la red local;
3. cada accion de cada ventana genera una salida correcta;
4. la salida coincide con la plantilla o base manual cargada como contrato;
5. se conservan formulas, formatos, estilos y estructura de celdas;
6. el sistema tolera meses con menos filas o con mas filas, manteniendo el mismo formato de salida;
7. las tres ventanas quedan operativas:
   - Libro Compras ACLT
   - Conciliacion Servicios por Marca
   - Facturacion Repuestos TYTSERV

## Reglas

- No respondas con teoria.
- Inspecciona codigo real.
- Ejecuta pruebas reales.
- Si detectas una regresion, corrigela y vuelve a validar.
- No dejes visible al usuario texto tecnico como runtimes, workers o detalles internos.
- Piensa como sistema de uso mensual continuo: cambian los datos, no el formato esperado.

## Verificaciones obligatorias

### 1. UX visible

- Revisa portada, area, ventanas y modulos.
- Asegura que el texto visible sea corto, profesional y orientado al usuario contable.
- Elimina o corrige cualquier texto tecnico visible.
- Verifica que el logo activo correcto sea `logo2`.
- Verifica que los enlaces internos visibles apunten a rutas activas y no a superficies legacy.

### 2. Acceso por red

- Deja un arranque LAN operativo.
- Usa un puerto claro y documentado.
- Verifica desde el propio host que la app responda por IP local, no solo por `127.0.0.1`.
- Si hay bloqueo real de firewall o permisos, documentalo con precision.

### 3. Validacion funcional por ventana

Ejecuta cada accion o modulo operativo:

- Accion 1
- Accion 2
- Accion 3
- Accion 4
- Consolidado de Acciones
- Repuestos TYTSERV
- Servicios por Marca

Para cada uno valida:

- carga correcta del archivo;
- respuesta HTTP o flujo operativo correcto;
- generacion del archivo final;
- historial o descarga visible;
- salida correcta respecto al contrato existente.

### 4. Paridad de salida

Para cada accion donde exista plantilla o contrato manual:

- compara celdas relevantes;
- compara formulas esperadas;
- compara formatos numericos y de fecha;
- compara estilos de encabezado y de filas de datos;
- compara estructura general del workbook.

### 5. Robustez mensual

Prueba entradas representativas con:

- menos filas que el ejemplo base;
- mas filas que el ejemplo base;
- mismo formato de origen.

Verifica que:

- no se rompan estilos;
- no se pierdan formulas;
- no se corten filas;
- no se desalineen columnas;
- no falle el proceso por longitud variable.

### 6. Resultado final

Entrega:

1. resumen ejecutivo corto;
2. lista de hallazgos, si existen;
3. pruebas ejecutadas;
4. resultado por ventana/modulo;
5. limite tecnico real que todavia siga vigente;
6. conclusion clara:
   - listo para uso mensual
   - o no listo, indicando exactamente por que

## Nota importante

No basta con decir "todo bien". Solo da por aprobado lo que realmente ejecutaste y verificaste.

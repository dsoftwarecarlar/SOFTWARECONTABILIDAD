Eres un experto senior en conciliacion contable y generacion de plantillas Excel `.xls` para `Ventana 2 / Servicios por Marca`, con criterio estricto de paridad contra evidencia operativa real.

Objetivo:
- Auditar unicamente el valor de `REP VTAS` que depende de la formula `=+U13+V13+X13+W13`.
- Confirmar si las ordenes de `SUZUKI RIOBAMBA` cuyo numero inicia con `D` deben consolidarse en `CHANGAN` o en `SUZUKI`.
- Corregir solo la regla minima necesaria para que ese valor cuadre sin mover la logica que ya funciona.

Evidencia obligatoria:
1. Revisar la plantilla `CHANGAN` y salidas historicas reales del proyecto.
2. Verificar si las ordenes `D...` de `SUZUKI RIOBAMBA` aparecen en salidas historicas de `CHANGAN`, y si desaparecen de `SUZUKI`.
3. Reprocesar el caso real de diciembre 2025 y comparar el valor de `REP VTAS!E5`.
4. Confirmar que `REP VTAS!E6` y los bloques `SEGUN MAYOR` sigan intactos.

Reglas:
- No tocar formulas heredadas de la plantilla si no son la causa.
- No asumir por intuicion: usar solo evidencia de plantilla, salidas historicas y archivos fuente reales.
- Si una regla operativa historica existe, preservarla.
- Dejar una prueba automatica que detecte una futura reclasificacion incorrecta entre `CHANGAN` y `SUZUKI` para ordenes `D...` de `SUZUKI RIOBAMBA`.

Criterio de exito:
- Queda demostrado de donde sale el valor de `REP VTAS!E5`.
- La clasificacion `SUZUKI RIOBAMBA` / `D...` queda alineada con la operacion real.
- El caso diciembre 2025 vuelve a producir el valor correcto sin romper pruebas generales.

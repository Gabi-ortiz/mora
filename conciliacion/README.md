# Conciliación bancaria (FBS vs. extractos)

Primera etapa: identificar y clasificar los conceptos de cada banco, empezando por **Banco Macro**.

- `reglas_macro.py`: catálogo de conceptos Macro (código causal + patrón de texto → categoría). Las categorías
  son comunes a todos los bancos, así cada banco nuevo solo agrega su propio archivo de reglas.
- `procesar_extracto.py`: lee el extracto Macro (CSV "Resumen" o XLSX "Últimos Movimientos"), controla la continuidad de saldos, clasifica cada
  movimiento, extrae CUIT/nombre de la contraparte y genera una planilla con las hojas
  `Resumen`, `Catalogo_Conceptos`, `Movimientos` y `Preguntas_Admin`.

```
pip install openpyxl
python procesar_extracto.py Resumen_Macro.csv "Conciliacion Macro.xlsx"
python procesar_extracto.py "08-2026 BANCO MACRO.xlsx" "Conciliacion Macro.xlsx"
```

- `leer_mayor_fbs.py`: lee el PDF "Consulta del Mayor" de FBS (cuentas E y O) por posición de columnas.
- `conciliar.py`: cruza extracto vs. mayor E + O (+ pendientes del mes anterior) y genera el tablero de 4 cuadros
  con control, una hoja de detalle por cuadro y la hoja `Conciliados` con el método de cada cruce.
- `NORMA_REGISTRACION.md`: propuesta de cómo registrar en FBS para que el cruce sea automático.

```
pip install -r requirements.txt
python conciliar.py "08-2026 BANCO MACRO.xlsx" mayor_e.pdf mayor_o.pdf salida.xlsx "Conciliaciones Bancos 07-2026.xlsx"
```

Criterios de cruce (importe y sentido siempre iguales, tolerancia $0,02): 1) CUIT/DNI, 2) referencia del banco en el
comprobante FBS, 3) apellido, agrupados del mismo CUIT/apellido y asiento mensual de gastos, 4) fecha + importe
único, lotes del día (sueldos, VEP), 5) solo importe a ≤10 días (sugerido), combinaciones de 2-3 partidas (revisar).
Los pases entre E y O y los registros FBS que se compensan entre sí no se consideran movimientos.

## Versión Google Sheets (Apps Script)

`appscript/Conciliacion.gs` es la misma lógica de cruce que `conciliar.py`, portada a Apps Script (validada con
agosto 2026: mismos 4 totales y diferencia de control $143,21).

1. En una planilla nueva: Extensiones > Apps Script, pegar `Conciliacion.gs`, guardar y recargar la planilla.
2. Menú **Conciliación > 1. Crear hojas de entrada** (una sola vez): crea `Extracto`, `Mayor E`, `Mayor O`,
   `Pendientes anteriores`, `Reglas` y `Empresas grupo`.
3. Pegar desde A1 el extracto de Macro y las exportaciones a Excel del mayor E y O de FBS.
4. **2. Procesar conciliación** → genera `Tablero` (4 puntas + control), `Pend. registrar en FBS`,
   `Pend. FBS sin banco` y `Conciliados`.
5. Al cerrar el mes, **3. Cerrar mes** guarda una copia completa y protegida de la planilla (foto) en la carpeta
   `Conciliaciones - Historial`, agrega una fila en la hoja `Historial` con los saldos y el link, y pasa los
   pendientes a `Pendientes anteriores`.

Flujo (igual al procedimiento de administración):
1. **Cuenta O** (hoja `Análisis O` y `O sin confirmar`): se netea por comprobante (RC / RM) o número de liquidación,
   junto con los pendientes anteriores. Lo que no netea queda "no confirmado" (puntas 3 y 4) y se marca, en la columna
   *Cruces encontrados*, si coincide con un pendiente anterior, con la cuenta E o con el extracto (solo marca).
2. **Cuenta E**: primero contra los pendientes de la conciliación anterior (movimientos del banco que el mes pasado no
   estaban registrados, o pendientes de FBS revertidos), después contra el extracto. Si el comentario trae una fecha
   (ej. "... 16/04/2026") se usa también para buscar en el banco. Lo cruzado va a `E conciliado`; lo que no, a
   `E sin cruzar` (posibles errores de registración) y `Banco sin registrar`.

Alineado con el procedimiento de administración (Indicaciones conciliación bancaria Macro):
- Pendientes de la cuenta **O** = valores no confirmados; pendientes que quedan en la cuenta **E** = posibles
  errores de registración (se marcan en la columna *Observación* y en los avisos del Tablero).
- Tarjetas y venta de cheques: un crédito del banco contra el **neto de un asiento FBS** (confirmación o
  cheques menos comisiones).
- Importes similares: se marcan **posibles duplicados** (mismo importe y mismo detalle en otro asiento).
- Reglas con código `*` (cualquier código causal) para LIQ COMER PRISMA / CABAL / PAYWAY y Tarjeta Naranja.

La lógica de cruce no toca hojas (`conciliarTodo_` recibe matrices de valores), así que se puede probar con Node.

## Modelo de conciliación actual (hoja por banco)

Saldo contable FBS = cuenta **E** (registrado, a confirmar) + cuenta **O** (confirmado contra banco).

| Sector | Qué contiene | Origen | Signo |
|---|---|---|---|
| Depósitos no registrados en contabilidad | Créditos del banco sin registro en FBS | Extracto | + |
| Cheques no registrados en contabilidad | Débitos del banco sin registro en FBS | Extracto | − |
| Depósitos no acreditados en banco | Ingresos de FBS que el banco no muestra | FBS | − |
| Cheques no debitados en banco | Egresos de FBS que el banco no muestra | FBS | + |

Control: `Saldo FBS (E+O) + dep. no registrados − cheq. no registrados − dep. no acreditados + cheq. no debitados = saldo extracto`.

Al pie se arma el asiento de gastos bancarios (comisiones, IVA, percepciones, impuesto a los créditos y
débitos con su 33% computable, SIRCREB, retenciones IIBB renta financiera); la hoja `Resumen` ya lo calcula.

Los extractos y planillas generadas no se versionan (contienen datos de la empresa).

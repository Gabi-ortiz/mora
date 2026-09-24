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

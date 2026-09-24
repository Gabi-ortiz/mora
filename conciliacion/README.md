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

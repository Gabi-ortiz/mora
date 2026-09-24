# Conciliación bancaria (FBS vs. extractos)

Primera etapa: identificar y clasificar los conceptos de cada banco, empezando por **Banco Macro**.

- `reglas_macro.py`: catálogo de conceptos Macro (código causal + patrón de texto → categoría). Las categorías
  son comunes a todos los bancos, así cada banco nuevo solo agrega su propio archivo de reglas.
- `procesar_extracto.py`: lee el CSV del extracto Macro, controla la continuidad de saldos, clasifica cada
  movimiento, extrae CUIT/nombre de la contraparte y genera una planilla con las hojas
  `Resumen`, `Catalogo_Conceptos`, `Movimientos` y `Preguntas_Admin`.

```
pip install openpyxl
python procesar_extracto.py Resumen_Macro.csv "Conciliacion Macro.xlsx"
```

Los extractos y planillas generadas no se versionan (contienen datos de la empresa).

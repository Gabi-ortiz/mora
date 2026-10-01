# Objetivos e incentivos — Plan de ahorro

Archivo único con lo que manda Fiat (cartas de objetivos y señales
comerciales). Lleva dos scripts en el mismo proyecto de Apps Script:
`senales.gs` (lector y alerta de la carpeta de señales, ver SENALES.md) y
`objetivos.gs` (estructura de objetivos, categorías e incentivos).

## Hojas
- **SEÑALES / FALTANTES**: las arma `senales.gs`. Cuando una señal se vuelca
  a las hojas de abajo, se marca "Cargada".
- **OBJETIVOS**: una fila por Mes + Indicador (SUSCRIPCIONES,
  PATENTAMIENTOS, PEDIDOS TOTALES...), con categoría del mes, Nº y fecha de
  carta, fecha y % de flujo (el flujo es sobre suscripciones), link y notas.
  Formato largo: un indicador nuevo es una fila más. Mes en formato
  `SEPTIEMBRE 26`, igual que Pedidos PDA.
- **CATEGORIAS**: una fila por trimestre de vigencia (la categoría se mide
  trimestralmente por mora cuota 6). Si se cargan los cortes y el valor real
  de mora, la categoría se calcula sola; si no, se carga a mano.
- **INCENTIVOS**: una fila por Mes + Concepto con el % para A / B / C,
  condición, base de cálculo y Nº de señal.
- **MES**: vista del mes elegido en B1 (objetivos, categoría e incentivos).
- **LISTAS_OBJ**: valores de los desplegables.

Los objetivos los define Fiat según el peso del concesionario sobre la marca
a nivel país; acá solo se cargan tal cual llegan en la circular.

## Semillas cargadas
- Septiembre 26 (carta 09/2026, categoría A): suscripciones 146,
  patentamientos 43, pedidos totales 60; flujo 23/09/2026 al 65%.
- Agosto 26: pedidos totales 57 (sacado del RESUMEN mensual; falta la carta).
- Categorías: JUL-SEP 26 = A (según carta); OCT-DIC 26 con cortes A ≤ 43%,
  C ≥ 51% (señal 1595), falta el % real de mora cuota 6.
- Incentivos septiembre 26: bonus 2A, 2B y 2C de pedidos (señales 1599 y
  1603).

## Instalación
1. En el archivo: Extensiones > Apps Script.
2. Reemplazar el contenido de `Code.gs` por `senales.gs` (versión nueva: el
   aviso va fijo a gortiz@grupoantun.com.ar y ya no trae menú propio).
3. Agregar un archivo nuevo (+ > Secuencia de comandos) llamado `objetivos`
   y pegar `objetivos.gs`. Guardar.
4. En el editor de Apps Script: **Servicios (+)** > **Drive API** > Agregar
   (lo usa el lector para convertir los PDF a texto).
5. Recargar la planilla: aparece el menú **Objetivos**.
6. Objetivos > **Armar / reparar estructura**.
7. Objetivos > **Leer contenido de señales ya registradas** (pide permisos
   nuevos; repetir hasta que diga que no quedan pendientes).
8. Si la revisión automática de señales ya estaba instalada, no hace falta
   reinstalarla.

## Carga automática de cartas
Las cartas de objetivos que se suban a la carpeta de señales se cargan solas
en OBJETIVOS (ver SENALES.md). Controlar en el mail que los números
coincidan con el PDF.

## Próximos pasos
- Pedidos PDA: que su hoja OBJETIVOS lea de acá (IMPORTRANGE) en vez de
  cargarse a mano.
- Definir de dónde sale el real de suscripciones y patentamientos.
- Agregar la cuota 6 al tablero de mora para seguir la categoría.

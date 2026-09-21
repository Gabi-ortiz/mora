# Tablero de control de mora — Fiat (planes de ahorro)

## Archivo de trabajo
https://docs.google.com/spreadsheets/d/12pLWcN7GLbmD7rBbeEdg_3dXHMWrQL88AtsGj4fH5wM

Archivo nuevo, armado desde cero por una cuestión de organización (reemplaza a
uno anterior llamado "Analisis mora GABI" que tenía todo mezclado: base +
cuadro de análisis en el mismo archivo, hoja "Hoja 2" para la base).

## Hojas
- **BASE**: la tabla cruda de planes, un plan por fila. La actualiza (los
  valores P/I/R) un compañero de trabajo. Llega por IMPORTRANGE desde otro
  archivo. **Confirmado contra el archivo actual** (21/09/2026): la hoja se
  llama exactamente "BASE", tiene 1858 planes cargados y las columnas caen
  en las mismas posiciones que en el archivo viejo (ver más abajo) — no hace
  falta tocar la configuración del script.
- **Historial_Mora**: log append-only, una fila por (fecha de snapshot,
  avance). Se genera con el script — no se edita a mano. Todavía no existe
  en el archivo nuevo, la crea el script solo la primera vez que corrés
  "Sacar foto ahora".
- **Tablero**: vista final con una fila por cada cuota de interés (3, 5, 7,
  9, 12), tomando siempre la última foto disponible del historial para el
  avance que corresponde a cada cuota. Se regenera con el script — no se
  edita a mano. Tampoco existe todavía, la crea "Actualizar tablero".
- **Log_Snapshot**: una línea por corrida del trigger diario, diga si hubo
  cambios en BASE o no. Tampoco existe todavía.
- **Detalle_Planes**: un plan por fila, leído en vivo de BASE (no del
  historial), con filtro nativo de Sheets ya armado para poder filtrar por
  Avance y ver qué planes puntuales caen en cada estado (Mora irregular,
  Rescindido, Renunciado, etc.). Se regenera con "Actualizar detalle de
  planes" — no se edita a mano. Tampoco existe todavía.

## Estructura de la hoja BASE (confirmada contra el archivo actual)
- Columna **N** (14): "Avance" — cuántas cuotas van de ese plan (entero,
  2 a 13 en los datos actuales).
- Columna **O** (15): "Estado" — texto: Ahorrista / Adjudicado / Rescindido /
  Renunciado / Cancelado.
- Columna **R** (18) en adelante: una columna por cuota, empezando en **C2**
  (columna R = C2, S = C3, T = C4, ... hasta AD = C14), con el valor cargado
  por el compañero:
  - `P` = cuota paga
  - `I` = cuota con pago irregular / en mora
  - `R` = plan rescindido (a partir de la cuota en que se rescindió, se
    completa con R en todas las columnas siguientes)

Estas posiciones son idénticas a las del archivo anterior, así que
`COL_AVANCE`, `COL_ESTADO` y `COL_C2` en `tablero_mora.gs` quedan tal cual
están (14, 15 y 18 respectivamente) — no requieren ajuste.

**Sobre "de dónde leer" por fila**: la columna "Avance" define cuántas
columnas de cuota tiene cargadas ese plan puntual (para avance = 9, el rango
cargado va de columna R/C2 a columna Y/C9). Esto es la descripción de la
fila individual, no confundir con la lógica de "mes atrasado" del punto
siguiente, que usa un rango distinto (con margen de un avance) y ya está
probada.

## Lógica de negocio (ya validada contra datos reales, no re-derivar de cero)

### 1) "Mes atrasado": qué cuota mide cada avance
El análisis de una cuota N se hace mirando los planes cuyo **Avance = N + 1**
(un avance de diferencia, para dar margen a correcciones tardías de pago), y
dentro de esas filas se mira el rango de columnas **C2 hasta C(N)** — es
decir, `ncols = avance - 2` columnas a partir de la columna C2 (columna R).

Ejemplo probado: avance = 13 → analiza cuota 12 → rango C2:C12 (hasta columna
AB), **nunca** incluye C13/AC. Esto es intencional y ya está confirmado
correcto — no es un bug a corregir.

### 2) Clasificación por plan (fila), dentro del rango de columnas del punto 1
- **Rescindido**: al menos una columna del rango = `R` (no hace falta que
  sean todas).
- **Mora irregular / Impagos**: al menos una columna del rango = `I`.
- **Pagado al día**: TODAS las columnas del rango = `P`. Se separa en:
  - **Pagos Adjudicados**: de esos, los que tienen Estado = "Adjudicado".
  - **Pagos Ahorristas**: de esos, los que tienen Estado = "Ahorrista".
- **Cartera total** = cantidad de filas con ese Avance.
- **Cartera activa** = Cartera total − Rescindidos.
- **TOTAL PAGOS** = Pagos Adjudicados + Pagos Ahorristas.
- **res+irre+imp** = Rescindidos + Impagos (mora irregular).
- **% de mora** = res+irre+imp / Cartera total.

Estas fórmulas fueron validadas fila por fila contra el cuadro manual
original ("Cuota 3/5/7/9/12 AGOSTO") y coinciden exacto en Rescindidos e
Impagos en los 5 casos, y en Pagos Adjudicados/Ahorristas en 3 de 5 (la
diferencia en los otros 2 es por drift normal de Estado entre el momento en
que se cargó el cuadro a mano y el momento en que se corrió el script — no
es un error de fórmula).

### 3) Por qué existe Historial_Mora (y no se calcula todo en vivo)
La columna "Avance" es LIVE: un plan que hoy está en avance 10 puede pasar a
avance 11 en cualquier momento (aprox. una vez por mes, fecha no siempre
predecible). Si el cuadro se recalculara en vivo contra BASE, un recálculo
tardío daría un número distinto al que fue el corte real. Por eso:
- Un snapshot diario (trigger) recorre TODOS los valores de Avance presentes
  en BASE y guarda una fila por avance en Historial_Mora, con la fecha.
- El snapshot solo escribe si detecta cambios reales en BASE desde la última
  vez (hash MD5 sobre Avance + Estado + C2..C14 de todas las filas — ignora
  a propósito columnas irrelevantes como teléfono/vendedor). Si no hay
  cambios, no duplica filas; solo deja constancia en Log_Snapshot.
- El Tablero nunca lee BASE directamente: siempre lee la última foto
  disponible en Historial_Mora para el avance que corresponde a cada cuota
  de interés (avance = cuota + 1).

### 4) Cuotas que se siguen en el Tablero
3, 5, 7, 9 y 12 (configurable en `CUOTAS_TABLERO`).

### 4bis) Objetivos de mora / incentivo (Jul-Sep 2026, esquema "80/20")
El Tablero suma una columna **% Incentivo**, calculada a partir de
`OBJETIVOS_INCENTIVO` en `tablero_mora.gs`: el incentivo que corresponde a
la franja de mora en la que cayó esa cuota. (Hubo una columna "Franja
objetivo" con el texto de la franja, se sacó por pedido de Gabi porque
mareaba — la franja se puede inferir del propio % de mora contra la tabla
de abajo.)

Reglas por cuota (mora < t1 → i1; entre t1 y t2 inclusive → i2; mora > t2 →
incentivo 0%):

| Cuota | < t1 | Incentivo | t1 a t2 | Incentivo |
|---|---|---|---|---|
| 3  | < 35% | 0.35% | 35 a 41% | 0.20% |
| 5  | < 30% | 0.65% | 30 a 36% | 0.45% |
| 7  | < 32% | 0.70% | 32 a 38% | 0.40% |
| 9  | < 45% | 0.70% | 45 a 51% | 0.40% |
| 12 | < 48% | 0.80% | 48 a 54% | 0.40% |

Decisiones tomadas con Gabi (21/09/2026):
- **Por ahora se aplica el mismo objetivo a toda la cartera**, sin separar
  por tipo de plan (70/30, 80/20, 90/10, etc.), aunque la tabla original
  esté etiquetada "80/20". Si más adelante aparecen tablas de objetivo
  distintas por tipo de plan, hay que filtrar `Tablero`/`Historial_Mora`
  por "Tipo de Plan" antes de aplicar cada tabla.
- **Mora por encima del límite superior de la franja = incentivo 0%** (no
  se sigue pagando el incentivo de la franja más baja).

### 4ter) "Planes a bajar": cuántos planes hay que recuperar para cambiar de franja
Además de en qué franja está parada cada cuota HOY, el Tablero muestra
cuántos planes (de los que hoy están en Rescindido/Mora irregular) hace
falta "levantar" (cobrarles / regularizarlos) para entrar en cada franja de
premio. Es la misma cuenta que usaba el archivo manual de Fiat en su
columna "Bajar":

```
objetivo_en_planes = Cartera total × umbral de la franja (t1 o t2)
Planes a bajar = techo(res+irre+imp − objetivo_en_planes), mínimo 0
```

Dos columnas en el Tablero:
- **Planes a bajar (A)**: cuántos planes faltan recuperar para que la mora
  quede por debajo de `t1` (ej. < 35% en Cuota 3, la mejor franja). Si ya
  está ahí, da 0.
- **Planes a bajar (B)**: lo mismo pero contra `t2` (ej. 41% en Cuota 3,
  la segunda franja) — o sea, lo mínimo para al menos no quedar afuera de
  todo incentivo. Si ya está en la mejor franja o en la segunda, da 0.

Sirve para ir trabajando la cartera mes a mes con un número concreto de
planes a recuperar por cuota, en vez de solo mirar el % de mora.

### 5) Detalle_Planes: de los números del Tablero a los planes concretos
El Tablero da agregados (cuántos rescindidos, cuántos en mora, etc.) para
las 5 cuotas de interés, pero para trabajar la cartera (llamar, reclamar,
etc.) hace falta saber QUIÉNES son. Para eso está Detalle_Planes:
- Lee BASE en vivo (no el historial) — acá interesa la situación actual de
  cada plan, no una foto vieja.
- Clasifica cada plan con la misma regla de "mes atrasado" que usa el
  Tablero (rango C2..C(avance-1)): Rescindido / Mora irregular / Pagado al
  día / Sin cuotas para analizar (avance=2, todavía no hay rango que mirar).
- Muestra también el Estado tal cual está en BASE (Ahorrista / Adjudicado /
  Rescindido / Renunciado / Cancelado) en una columna aparte — un plan puede
  ser "Renunciado" en Estado y "Pagado al día" en la clasificación de mora
  al mismo tiempo, son dos cosas distintas.
- Trae el filtro de Sheets ya activado (ícono de embudo en el encabezado):
  filtrás por columna "Avance" para pararte en la misma cuota que estás
  mirando en el Tablero (recordá: avance = cuota + 1), y después por
  "Clasificación mora" (Rescindido / Mora irregular) o por "Estado"
  (Renunciado) para ver la lista de planes de ese grupo puntual.

## Script actual
Ver `tablero_mora.gs` en este repo — es la versión funcionando, probada e
instalada como Apps Script bound al archivo anterior, con la configuración
ya confirmada contra el archivo nuevo. Para instalarlo:
1. Abrir el archivo → Extensiones > Apps Script.
2. Borrar el contenido de `Code.gs` (o el archivo que haya por defecto) y
   pegar el contenido de `tablero_mora.gs`.
3. Guardar y recargar la hoja de cálculo (para que aparezca el menú "Mora").
4. Menú "Mora" > "Instalar snapshot automático diario" (una sola vez).
5. Menú "Mora" > "Sacar foto ahora" para la primera foto.
6. Menú "Mora" > "Actualizar tablero" para generar el Tablero por primera vez.
7. Menú "Mora" > "Actualizar detalle de planes" para generar Detalle_Planes.

De ahí en más el trigger diario (~4am) se encarga de sacar la foto solo
cuando hay cambios reales en BASE; "Actualizar tablero" y "Actualizar
detalle de planes" se pueden correr cuando se quiera (no dependen del
trigger) para refrescar cada vista.

## Historial de decisiones de diseño (para no repreguntar)
- Se probó y descartó comparar "avance actual" contra "avance actual - 1"
  para chequear variación: NO son la misma cuota (miden C2:C(avance-1) cada
  uno), así que es normal que difieran — no indica error.
- Se sacó la columna "PeriodoReal" de Historial_Mora (generaba ruido, y no
  aplica limpio a un historial que loguea TODOS los avances, no solo los de
  interés).
- Se sacó la columna "Mes análisis" del Tablero (redundante con "Fecha
  foto").
- En el Tablero, "Período real" se muestra como etiqueta legible tipo
  "Cuota 3 AGOSTO" (derivada de la fecha de la foto), no como fecha cruda,
  para no confundir con la columna "Avance".

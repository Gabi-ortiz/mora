# Pedidos PDA — base acumulada de adjudicados (SGA)

## Archivo de trabajo
https://docs.google.com/spreadsheets/d/1dsvnNpKUmoKDSaZcEv4UsofHWbWK_m0SwKKvr7dkxGM ("Pedidos - PDA")

Reemplaza el esquema de un archivo por mes ("AGOSTO 26", "SEPTIEMBRE 26",
...) por un solo archivo acumulado, para poder hacer informes en el tiempo.
Es parte del mismo proyecto que el tablero de mora (misma administración de
planes); el objetivo final es un dashboard con toda la información e
incentivos.

## Cómo trabajaban hasta ahora (archivo mensual)
- Bajan de SGA el reporte "Cartera de clientes - Adjudicados" **por acto de
  adjudicación** (filtro tipo "Adjudicados; 15/09/2026"). Es un .XLS que en
  realidad es HTML: 7 filas de encabezado (título, fecha, usuario, filtros),
  títulos en la fila 8, 37 columnas, y una fila final "Cantidad de
  registros listados".
- Borran 21 columnas, ocultan 9, renombran 2 (`Av.` → Avance,
  `Mod. Ganador` → Modalidad), reordenan y agregan 7 columnas de trabajo
  (Mes, Responsable, Llave x llave, Pedido, Carpeta, Observacion,
  RESPONSABLE).
- **"Mes" = mes de cierre** (confirmado): el mes en que la operación sale
  aprobada y cuenta como neto. Un adjudicado de agosto puede cerrar en
  septiembre (en AGOSTO 26 hay 8 así), y por eso el RESUMEN de septiembre
  tenía que leer el archivo de agosto con IMPORTRANGE.
- Los objetivos cambian mes a mes (llegan por circular).
- Problemas vistos: nombres inconsistentes entre meses (CLARI/CLARISA,
  SERGY/SERGIO), ajustes "+1" a mano dentro de fórmulas, fórmulas que
  apuntan a columnas corridas, precios V.M copiados a mano en cada mes.

## Hojas del archivo nuevo
- **BASE**: una fila por adjudicación, todos los actos. Clave = Acto +
  Solicitud (las solicitudes casi no se repiten entre actos; si una se
  re-adjudica en otro acto, es otra fila). Mismas columnas y orden que la
  hoja mensual, con estos cambios:
  - `Acto` (nueva, columna A): mes del acto de adjudicación, formato
    `SEPTIEMBRE 26`.
  - `Mes` → `Mes cierre`, mismo formato `MES AA`. Se completa solo con el
    mes actual al marcar Pedido = APROBADO y Carpeta = APROBADA (si está
    vacío; se puede corregir a mano).
  - Desplegables estrictos en Acto, Mes cierre, Responsable, Pedido,
    Carpeta, Modalidad y Llave x llave. RESPONSABLE (vendedor) tiene
    desplegable pero acepta nombres nuevos con aviso.
  - Columnas extra ocultas al final con datos de SGA que antes se borraban
    (Vendedor SGA, Monto licitado, Mail, Dirección), pensando en el
    dashboard / incentivos.
  - Pedido con colores (aprobado verde, pendiente amarillo, suspendido
    naranja, baja rojo) y filtro activado.
- **PEGAR SGA**: se pega el reporte tal cual se baja. El script detecta la
  fila de títulos y el acto (de la fila de filtros), mapea las columnas por
  nombre, no duplica lo ya cargado y limpia la hoja al terminar.
- **RESUMEN**: el tablero del mes, eligiendo el mes en B1. Tres bloques
  (adjudicados del acto, estado del acto, cierres del mes) + tabla por
  modelo. Los "netos del mes" cuentan cierres de **cualquier** acto, y se
  separan en "del acto del mes" y "de actos anteriores" (lo que antes se
  traía con IMPORTRANGE).
- **RESUMEN – bloque INCENTIVO PEDIDOS** (columnas J-K): categoría del mes,
  % cumplimiento de pedidos, % del bonus de cumplimiento (si llega al 100%),
  pedidos adicionales al objetivo y % del bonus por pedido adicional, según
  la categoría. Los conceptos se buscan por texto ("CUMPLIMIENTO PEDIDOS",
  "ADICIONALES") en INCENTIVOS. El bonus de cumplimiento exige además 100%
  de suscripciones, que no se mide en este archivo.
- **INFORME**: una fila por mes (desde el primer acto) con adjudicados,
  bajas, % bajas, aprobados del acto, % conversión, sin cerrar, netos del
  mes (del acto / arrastre), objetivo, % cumplimiento y VM netos. Abajo, por
  responsable (todos los actos).
- **PRECIOS**: Mes | Modelo | V.M. Se carga una vez por mes. Sembrado con los
  precios de agosto y septiembre tomados de los RESUMEN mensuales (bloque
  "Enviados"/"Netos", que es el que tenía los precios actualizados).
- **OBJETIVOS** (solo lectura): Mes | Objetivo | Categoría de PEDIDOS
  TOTALES (Fiat), leído con IMPORTRANGE del archivo "Objetivos y señales
  comerciales - PDA" (`ID_ARCHIVO_OBJETIVOS`). Se carga allá, no acá. La
  primera vez: clic en OBJETIVOS!A2 > "Permitir acceso".
- **INCENTIVOS** (solo lectura): Mes | Concepto | A | B | C (Fiat), del mismo
  archivo.
- **Objetivos de pedidos por modelo** (OBJETIVOS!E:G, solo lectura): p. ej.
  julio 26 pide 54 pedidos totales, de los cuales 27 Cronos y 4 Titano. En
  RESUMEN (fila 26, bloque INCENTIVO PEDIDOS) se ve "Cronos 20/27 · Titano
  1/4": netos del mes cuya Descripción de producto contiene CRONOS / TITANO
  contra su objetivo (no cuenta los netos de AJUSTES).
- **LISTAS**: valores de los desplegables. Para agregar un responsable, un
  estado o un vendedor, se agrega acá.

## Instalación (una sola vez)
1. Abrir "Pedidos - PDA" → Extensiones > Apps Script.
2. Pegar el contenido de `pedidos_pda.gs` en `Code.gs` y guardar.
3. Recargar la planilla: aparece el menú **Pedidos**.
4. Pedidos > **Armar / reparar estructura** (pide permisos la primera vez).
5. (Ya hecho; quedó fuera del menú) `importarHistorico` desde el editor:
   copia AGOSTO 26 y SEPTIEMBRE 26 a BASE (normaliza CLARI/SERGY, convierte "SEPTIEMBRE" en
   "SEPTIEMBRE 26", etc.). Se puede correr de nuevo sin duplicar.

## Uso mensual
1. Bajar el reporte de SGA del acto (como siempre).
2. Abrirlo en Excel, Ctrl+A, Ctrl+C, y pegar en **PEGAR SGA**, celda A1.
3. Pedidos > **Agregar acto pegado a BASE** → confirma el acto detectado.
4. Cargar en PRECIOS los V.M del mes y en OBJETIVOS el objetivo de la
   circular.
5. Trabajar en BASE filtrando por Acto (y los pendientes de actos anteriores
   siguen en la misma hoja).

## Validado contra los datos reales (simulación del script)
- Import de AGOSTO 26 (111) + SEPTIEMBRE 26 (117) = 228 filas; reimportar no
  duplica; pegar el reporte de SGA del 15/09/2026 detecta "SEPTIEMBRE 26" y
  sus 114 solicitudes ya estaban cargadas.
- Septiembre: 117 adjudicados, 71 por licitación, 25 bajas, 51 aprobados del
  acto; netos del mes = 59 (51 del acto + 8 arrastrados de agosto).

## Rediseño RESUMEN / INFORME (v3)
- **RESUMEN**: título + selector de mes en **H1** (lista "Meses con datos"
  = LISTAS!A, un UNIQUE de los meses de BASE —Acto y Mes cierre— ordenado
  del más reciente al más viejo; ya no se guarda la lista fija de 36 meses). Fila de 4
  tarjetas (adjudicados del acto, aprobados del acto con % conversión,
  netos del mes con del acto / anteriores, cumplimiento con objetivo y
  categoría, en semáforo verde ≥100% / naranja ≥90% / rojo). Debajo, 4
  bloques del mismo ancho de a dos (Adjudicados / Estado del acto; Cierres
  del mes / Incentivo pedidos) y la tabla POR MODELO con fila TOTAL fija
  arriba y columna % conversión. Sin líneas de cuadrícula.
- **INFORME**: "Evolución mensual" con encabezado en dos grupos (ACTO:
  adjudicados, bajas, aprobados, % conversión; CIERRE DEL MES: netos,
  objetivo, % cumplimiento, categoría, VM netos) y fila TOTAL. Se sacaron
  licitación/sorteo, % bajas, sin cerrar y el desglose del acto/anteriores
  (están en el RESUMEN). "Por responsable" con TOTAL.
- **RESPONSABLE** = "Of. Cuenta" de SGA (coincide en 100 de 114
  solicitudes de septiembre; el resto se cambió a mano, p. ej. ADM). Sin
  desplegable; se puede editar.
- **Acto** y **Mes cierre** sin desplegable, en gris: los completa el
  script (Acto al importar, con aviso si se edita; Mes cierre al aprobar).
  Si Mes cierre se escribe a mano, onEdit lo pasa a "MES AA" y avisa si no
  es válido.

## Control contra los archivos mensuales (01/10/2026)
- BASE contra AGOSTO 26 y SEPTIEMBRE 26: están todas las solicitudes (111 y
  117), ninguna de más, y todos los campos de trabajo coinciden. Única
  diferencia: solicitud 9258140 (septiembre), Carpeta = REINGRESADA en el
  archivo mensual (estado nuevo, agregado a la lista Carpeta).
- **AJUSTES** (hoja nueva): cierres que no están en BASE, para que los meses
  de arranque den igual que sus RESUMEN originales. Suman a netos y VM en
  RESUMEN (netos del mes, de actos anteriores, tabla por modelo) e INFORME.
  - Agosto: 15 netos / $ 492.193.000,1 de actos de julio (DP1 7, FO1 1,
    DT1 2, MB1 4, NT3 1) → 57 netos / $ 1.965.767.000, igual al original.
  - Septiembre: 2 netos / $ 61.646.000 = los "+1" cargados a mano en el
    RESUMEN original (NC1 y MB1) → 62 netos / $ 2.405.834.000, igual al
    original. Falta confirmar de qué solicitudes son.
- **Trabajados** se calcula como en el RESUMEN original: (netos licitación +
  sorteo) − (suspendidos/pendientes con CC + bajas). Se agregaron
  **Terminados** (aprobados + susp/pend con CC) y **Sin terminar** (total −
  trabajados − bajas). Septiembre da 82 / 53 / 10, igual al original.
- Agosto mantiene dos diferencias en el RESUMEN que son del original, no de
  la base: "Caídas" de licitación contaba solo las bajas con Observación =
  "NO PAGO" (7; el resto de los meses cuenta todas: 15), y "Terminado" (48)
  dejaba afuera los 2 FP3 aprobados porque el modelo no estaba en su tabla
  (correcto: 50).

## Julio 2026 (incorporado el 02/10/2026)
Archivo "JULIO 26" (hoja JULIO, 89 adjudicados, mismas columnas). Se suma con
`incorporarJulio()` desde el editor de Apps Script (una sola vez; si se corre
de nuevo no duplica):
- BASE: 89 filas con Acto = JULIO 26 (CLARI/SERGY normalizados).
- PRECIOS: los de julio (bloque "Enviados JULIO" de su RESUMEN), solo los
  que falten.
- AJUSTES: el acto de julio trae los 15 cierres de agosto que estaban como
  ajuste (DP1 7, MB1 4, DT1 2, FO1 1, NT3 1) y el "+1" MB1 de septiembre
  (una solicitud de julio cerrada en septiembre): esas filas se sacan. Se
  agrega el FS1 de julio (único cierre de junio cargado a mano en su
  RESUMEN; el resto del bloque daba #ERROR porque venía de un Excel local
  "JUNIO 26.xlsx" que no está en Drive).
- **Regla de valuación** (confirmada): un cierre vale el V.M del mes en que
  se liquida, aunque sea de un acto anterior (uno de julio liquidado en
  agosto vale el precio de agosto). Por eso los netos de AJUSTES no llevan
  importe fijo: se valúan con PRECIOS del mes. La columna "VM extra" de
  AJUSTES queda solo para correcciones de importe.
- Resultado: julio 26 netos / $ 913.724.000; agosto 57 netos /
  $ 2.006.468.000 (el RESUMEN original daba $ 1.965.767.000 porque valuaba
  los cierres de julio con la lista vieja); septiembre 62 netos. Queda
  pendiente el "+1" NC1 de septiembre (¿acto de junio?).
- Objetivo de julio (54, del RESUMEN de JULIO 26) va en el archivo de
  objetivos.

## Modelos sin precio (PRECIOS!G:H)
Lista calculada (no se edita): modelo y período de los netos (BASE aprobados
por Mes cierre, y netos de AJUSTES) que no tienen V.M en PRECIOS para ese
mes. Sin ese precio su VM da 0. Cada par desaparece al cargar el precio en
A:C. Ordenada por período (cronológico) y modelo.

## Orden de BASE
BASE se ordena por Acto en orden cronológico (más viejo arriba); dentro de
cada acto se respeta el orden que ya tenían las filas y las notas viajan con
su fila. Se ordena sola al agregar un acto, una recarga o un archivo mensual;
también en el menú: Pedidos > Ordenar BASE por acto.

## Recarga del mismo acto (operaciones sumadas a mitad de mes)
Volver a pegar el reporte de SGA del mismo acto y correr "Agregar acto pegado
a BASE": la clave es Acto + Solicitud, así que solo se agregan las
solicitudes que no estaban. Las ya cargadas no se modifican (aunque en SGA
haya cambiado el avance u otro dato) y lo trabajado (Pedido, Carpeta,
Observación, etc.) queda igual. En una recarga, las filas nuevas quedan con
una nota en la celda Solicitud ("Agregada en la recarga del acto ... del
dd/mm/aaaa") y el aviso lista cuáles se agregaron. Probado: carga de 111
operaciones y recarga del reporte completo de 114 → agrega 3, no toca las
111.

## Julio ajustado con el archivo de junio (02/10/2026)
"JUNIO 26.xlsx" (Drive, hoja JUNIO, 123 adjudicados) NO se suma a BASE: solo
se usa para ajustar julio. Tiene 28 operaciones del acto de junio con Mes =
JULIO (DP1 14, MB1 6, AR2 4, FO1 2, FT3 1, FP3 1); ninguna cerró en agosto
ni septiembre (el "+1" NC1 de septiembre sigue sin explicar). Se aplica con
`ajustarJulioConJunio()` desde el editor: reemplaza los ajustes de JULIO 26
(el FS1 que venía del RESUMEN de julio) por esas 6 filas, y agrega el V.M
de FP3 de julio si falta ($ 32.833.000: el RESUMEN de julio lo trae como
"FP1"; es el valor de FP3 en agosto). Julio queda en 53 netos (25 del acto
+ 28 de junio) / $ 1.816.743.000.

## Cambios del equipo que se respetan (03/10/2026)
En el archivo se cambió: "Llave x llave" renombrada a LLXLL y pasada a
casillas; columnas ocultas distintas (se mostraron Concepto de venta,
Teléfono, etc.); HECHO y ENVIADO en la lista de Pedido; un ajuste de julio
del acto de mayo (DP1 1). Desde esta versión:
- "Armar / reparar estructura" ya NO reformatea una BASE existente
  (encabezados, columnas ocultas, anchos, casillas/desplegables, colores,
  bordes): eso se aplica solo al crear la hoja. En una BASE existente solo
  corrige meses convertidos en fecha (las columnas se buscan por nombre,
  ver abajo).
- Ordenar BASE usa Range.sort con una columna auxiliar temporal: se mueven
  las filas enteras (valores, casillas, colores y notas).
- AJUSTES: encabezado y formato solo al crear (o la migración vieja "VM
  netos"); las filas cargadas a mano no se tocan.
- El orden de las pestañas solo se acomoda si se creó una hoja nueva.
- `incorporarJulio` y `ajustarJulioConJunio` se sacaron (ya aplicadas).

## Columnas de BASE por nombre (08/10/2026)
El equipo agregó **Modelo PEDIDO** (columna E, carga manual), lo que corrió
Mes cierre a G, Pedido a V, etc. Desde esta versión el script busca cada
columna de BASE **por el nombre del encabezado** (`mapaBase`), no por su
posición:
- Se pueden agregar, mover o insertar columnas sin romper nada. Las
  columnas que no son del script (Modelo PEDIDO u otras) no se tocan nunca:
  al agregar un acto se escribe solo en las columnas del script.
- Renombres conocidos: "Llave x llave" = LLXLL / "Llave por llave"
  (`ALIAS_BASE`); además se acepta el nombre sin distinguir mayúsculas ni
  acentos, salvo Responsable / RESPONSABLE, que son dos columnas distintas.
- Si falta una columna del script (se borró o se renombró a otra cosa),
  "Armar / reparar estructura" avisa cuál y no arma nada.
- Las fórmulas de RESUMEN, INFORME, LISTAS ("Meses con datos") y PRECIOS
  ("Falta precio") se generan con las letras reales de cada columna.
- "Modelo PEDIDO" había heredado el gris y el texto gris de Mes cierre
  (que indica "lo completa el script"); se dejó en blanco. Todavía no lo usa
  ningún cálculo.

## Qué hace "Armar / reparar estructura" con lo cargado a mano
- **PRECIOS** y **AJUSTES**: los valores solo se escriben al crear la hoja;
  después no se tocan (PRECIOS tampoco el formato). Solo se corrigen meses
  que Sheets haya convertido en fecha.
- **LISTAS**: lo agregado a mano se mantiene; un valor nuevo que trae el
  script se agrega una sola vez (si se borra, no vuelve); columnas nuevas
  agregadas a la derecha se respetan. La columna A (Meses con datos) es
  calculada y se rehace.
- **BASE**: los datos y el formato no se tocan (ver "Cambios del equipo que
  se respetan").
- **RESUMEN / INFORME / OBJETIVOS / INCENTIVOS**: se rehacen enteras (son
  fórmulas / datos leídos de otro archivo).

## VM netos del INFORME (corregido)
La fórmula usaba SUMPRODUCT con un SUMIFS de criterio matriz (el modelo de
cada fila de BASE). Google Sheets lo resolvía con un solo precio (el del
modelo de la primera fila, NC1) para todos los netos: agosto daba
$ 1.987.393.000 en vez de $ 1.965.767.000. Ahora recorre los precios del mes
(FILTER de PRECIOS) y multiplica cada V.M por los netos de ese modelo
(COUNTIFS), igual que la tabla por modelo del RESUMEN.

## Meses como texto (corregido)
Sheets convierte "SEPTIEMBRE 26" en la fecha 26/09 si la celda no es texto,
y entonces las fórmulas no encuentran el mes. Las columnas de mes (Acto, Mes
cierre, LISTAS!A, PRECIOS!A, RESUMEN!B1, INFORME!A) quedan en formato texto
y "Armar / reparar estructura" corrige las que ya se habían convertido.

## Pendiente / decisiones
- "RESPONSABLE" (vendedor) y "Responsable" (CLARISA/SERGIO/TP/CHEXA) se
  mantienen con el nombre de siempre; evaluar renombrar el primero a
  "Vendedor" para evitar confusión.
- VM netos = netos del mes × V.M del mes de cierre (el RESUMEN viejo tenía
  además un bloque VM × (aprobados + suspendidos con CC aprobada) con
  precios desactualizados; no se replicó).

## PRECIOS con precio CC (13/10/2026)
PRECIOS pasa a tener: **Mes, Modelo, Version, Inc pedido, Prec publico c/IVA,
Prec CC c/IVA, Prec CC s/IVA**. El viejo "V.M" quedó como Prec publico c/IVA.
La lista "Falta precio" (calculada) se movió a I:J.
- **El VM se calcula con Prec CC s/IVA** (`COL_PRECIO_VM`): RESUMEN (VM netos
  del mes y tabla "POR MODELO PEDIDO (V.M. = Precio CC sin iva)"), INFORME y
  AJUSTES. "Falta precio" lista los modelos con netos sin Prec CC s/IVA.
- **INFORME, julio–septiembre 2026**: no tienen Prec CC s/IVA, así que el VM
  netos queda como valor fijo (sin fórmula) con el importe que daba el precio
  cliente con IVA (`VM_INFORME_A_CORREGIR`: 1.855.113.000 / 2.006.468.000 /
  2.711.975.000) y al costado "Precio cliente con iva - a corregir". Cuando se
  cargan los Prec CC s/IVA de ese mes, el INFORME vuelve a la fórmula solo
  (al actualizarlo). En el RESUMEN esos meses dan $0 hasta cargarlos.
- **Octubre** se completó con bonus_x_modelo_1.xlsx (Version, IPC Pedidos →
  Inc pedido, precios CC con y sin IVA). Los precios públicos coincidían con
  los que había. El Excel trae "DP1 – CRONOS DRIVE PLUS 1.3" a $39.210.000,
  que es el precio de NC1 en octubre: se cargó en NC1 (no se agregó DP1).
  AR2, DT1 y TI1 no están en el Excel (quedan sin precio CC). Se agregaron
  los modelos que no estaban: UL3, AU3, PI3, UA4, FA4, GO0, DA1, VS1, DU1,
  VT3, TT1.

## Cálculos por Modelo PEDIDO (13/10/2026)
Los cálculos por modelo ya no usan "Modelo ahorro" sino **Modelo PEDIDO**
(columna cargada a mano en BASE, agregada a `COLUMNAS_BASE`). Si una fila no
tiene Modelo PEDIDO (julio–septiembre, o adjudicados todavía sin pedido), se
usa su Modelo ahorro (`formulaModeloCalculo`). Aplica a:
- RESUMEN, tabla "POR MODELO PEDIDO": modelos, adjudicados/aprobados/
  suspendidos del acto, netos del mes, V.M y VM netos.
- INFORME: VM netos de cada mes.
- PRECIOS, "Falta precio": se busca el precio del Modelo PEDIDO.
- RESUMEN I26 (Cronos / Titano dentro del total): el modelo se reconoce por
  la Version de PRECIOS del Modelo PEDIDO en el mes; si no hay, por la
  descripción de producto de la fila.
Probado en el archivo real: octubre da CA6 1, FO1 1, NC1 6 (VM $ 223.909.752
con Prec CC s/IVA); septiembre suma 62 netos (aparece VS1, pedido de un FS1).

**IVA 10,5% (13/10/2026):** en octubre, Strada (FS1, DA1, VS1, DU1), Fiorino
(FO1) y Toro (NT3, VT3, TT1) tienen Prec CC s/IVA = Prec CC c/IVA / 1,105
(el Excel lo traía con 21%). El resto de los modelos queda con 21%.

## RESUMEN: tabla por modelo con incentivos (13/10/2026)
La tabla "POR MODELO PEDIDO" ya no muestra Adjudicados / Aprobados / Susp.
c/CC aprobada / % conversión del acto. Lista solo los modelos con netos en el
mes y tiene: **Modelo, Netos del mes, V.M, VM netos, Bonus pedidos, Inc
pedido, Total incentivos**.
- Bonus pedidos = el % de "Bonus cumplimiento pedidos totales" del bloque
  INCENTIVO PEDIDOS (I23): según categoría, y 0 si no se cumplió el objetivo.
- Inc pedido = columna Inc pedido de PRECIOS para ese modelo y mes.
- Total incentivos = VM netos × (Bonus pedidos + Inc pedido).
- El tile "VM netos del mes" ahora toma E33.
Probado en el archivo (octubre): CA6 13% → $ 4.182.798; FO1 3% →
$ 805.754; NC1 0%; total $ 4.988.552 (sin bonus: octubre no tiene objetivo).

## BASE: colores (13/10/2026)
- Se sacó el color automático de Pedido = PENDIENTE (lo pinta el equipo a
  mano por otra cuestión operativa). Quedan APROBADO, SUSPENDIDO y BAJA ADJ.
  "Armar / reparar" no vuelve a poner el de PENDIENTE (en una BASE existente
  no toca los formatos).
- **Solicitud resaltada** (celeste y negrita) si figura en la hoja
  **Licitaciones** (solicitud en D) con "Turin" o "FCA" en la columna N
  (`resaltarLicitaciones`). Regla de formato condicional; "Armar / reparar"
  solo la agrega si falta. Hoy marca 180 solicitudes.

## Bonificaciones (13/10/2026)
El equipo agregó en BASE la columna **$ descuento** (importe bonificado, a
mano; está en `COLUMNAS_BASE`). La tabla "POR MODELO PEDIDO" suma al final:
- **Bonificaciones**: suma de $ descuento de los netos del mes de ese modelo
  (mismo criterio que Netos del mes: Mes cierre = mes, aprobado, Modelo PEDIDO).
- **Total** = Total incentivos - Bonificaciones. La fila TOTAL suma ambas.
La tabla ocupa B:J (la columna J pasó de margen a columna de 118 px). Ojo: el
archivo está en formato inglés, así que el importe se carga sin puntos
(100000); "100.000" lo toma como 100.

## Familias (13/10/2026)
PRECIOS tiene la columna **Familia** (H, la carga el equipo: Cronos, Strada,
Toro...). En RESUMEN:
- Tabla por modelo (B:K): columna **Familia** al lado de Modelo. Sale de
  PRECIOS para ese modelo y mes; si ese mes está vacía, la de otro mes; si
  no hay ninguna, "Sin familia".
- Debajo, **POR FAMILIA** (fila 82): Familia, Modelos, Netos, V.M promedio
  (VM netos / netos), VM netos, Bonus pedidos, Inc pedido (promedio ponderado
  por VM netos), Total incentivos, Bonificaciones y Total. Suma la tabla por
  modelo.
- La celda oculta del mes pasó de K1 a **M1** (K es columna de la tabla).
Probado (octubre): Cronos = CA6 · NC1, 7 netos, VM $ 199.381.984, incentivos
$ 4.182.798; Fiorino = FO1, 1 neto, $ 805.754.

**Orden y formato (13/10/2026):** el cuadro POR FAMILIA va arriba (fila 31,
hasta 12 familias) y POR MODELO PEDIDO debajo (fila 48). "Netos del mes" se
fuerza a formato número (#,##0) en los dos cuadros. Titano (TI1, DT1) también
va con IVA 10,5% en PRECIOS (jul-oct corregidos con =F/1,105); TN5 y TV6 quedan
al 21% (no identificados).

**POR FAMILIA sin %:** quedan Familia, Modelos, Netos del mes, VM netos,
Total incentivos, Bonificaciones y Total (B:H; sin V.M promedio). Los % de Bonus
pedidos e Inc pedido quedan solo en el cuadro por modelo.

## PROYECCION (14/10/2026)
Columna **PROYECCION (sin carpeta aprobada)** al lado de cada cuadro del RESUMEN: N (por modelo) y J
(por familia). Cuenta las operaciones de BASE (de cualquier acto) con
**Modelo PEDIDO cargado y Carpeta vacía**, por Modelo PEDIDO (la familia suma
sus modelos). Solo se muestra cuando el mes elegido es el mes en curso (hoy
OCTUBRE 26); en meses anteriores queda vacía. Los modelos que solo están en
la proyección se agregan a la lista con 0 netos. Probado: PC5 = 1 (acto
OCTUBRE 26, ahorro NC1, Pedido HECHO, sin carpeta).
- En el cuadro **por familia** la proyección tiene 3 columnas (J:L): Cantidad,
  **Descuentos** (suma de $ descuento de esas operaciones) e **Incentivos /
  bonus** (Prec CC s/IVA del modelo en el mes × (Bonus pedidos + Inc pedido)).
  El cuadro por modelo solo muestra la cantidad (N). Cuenta solo Carpeta vacía.
  Probado: Cronos 1 (PC5), descuento $ 2.700.000, incentivo $ 5.090.063.

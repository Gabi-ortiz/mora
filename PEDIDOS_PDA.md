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

## Qué hace "Armar / reparar estructura" con lo cargado a mano
- **PRECIOS** y **AJUSTES**: los valores solo se escriben al crear la hoja;
  después no se tocan (PRECIOS tampoco el formato). Solo se corrigen meses
  que Sheets haya convertido en fecha.
- **LISTAS**: lo agregado a mano se mantiene; un valor nuevo que trae el
  script se agrega una sola vez (si se borra, no vuelve); columnas nuevas
  agregadas a la derecha se respetan. La columna A (Meses con datos) es
  calculada y se rehace.
- **BASE**: los datos no se tocan; se rehacen formato, desplegables, bordes y
  filtro.
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

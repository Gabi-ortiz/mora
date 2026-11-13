# Motor de conciliaciones bancarias (grupo TURIN / JEA / CHEXA)

Traspaso para retomar el trabajo en una sesión nueva. Leer entero antes de tocar nada.

## Cómo trabajar con el usuario

- Habla en español rioplatense y no es técnico. Respuestas cortas, en castellano, sin jerga; tablas cuando ayudan.
- Rama de trabajo: **`claude/compassionate-newton-ydkf6u`**. Todo el código está ahí; si la sesión arranca en otra
  rama, hacer `git fetch origin claude/compassionate-newton-ydkf6u && git checkout claude/compassionate-newton-ydkf6u`.
  No abrir PRs salvo que lo pida.
- Después de cada cambio al motor:
  1. correr las pruebas de regresión (ver "Pruebas") y comparar con la versión anterior: los otros bancos no se tienen
     que mover, y si algo cambia, revisarlo cruce por cruce y explicarlo;
  2. commit con mensaje claro y push a la rama;
  3. mandarle al usuario `conciliacion/appscript/Conciliacion.gs` (SendUserFile): lo pega en el proyecto de Apps Script
     "Motor Conciliaciones" ("actualicé el motor").
- Cuando procesa un mes, comparar con la conciliación manual de ese mes (archivo "Conciliaciones Bancos MM-2026") y
  darle listas exactas de qué marcar en "Pendientes en trabajo" (Compensado / Pasar a diferencia) para que quede igual.
- Antes de reprocesar con un motor nuevo, avisar si alguna partida ya trabajada (Estado, Responsable o Comentario) puede
  cambiar de texto o importe: pierde la marca.
- Cuando pide "revisar sin modificar", no escribir nada en las planillas.
- Las planillas tienen datos de la empresa: no versionarlos en el repo (extractos, bases, PDFs van al scratchpad).

## Arquitectura

- `conciliacion/appscript/Conciliacion.gs`: el motor (biblioteca "Motor" en Apps Script, modo desarrollo).
- `conciliacion/appscript/planilla-banco/Banco.gs`: script de cada planilla (menú + `onEdit` que llama a
  `Motor.alEditar(e)`). `appsscript.json` al lado.
- Cada planilla (un banco de una empresa) tiene: E. bancario (extracto pegado), Conciliación (4 puntas), Tablero,
  Para revisar, Pendientes en trabajo, E conciliado, Desconciliados, Banco sin registrar, E sin cruzar, O sin confirmar,
  Análisis O, Historial, hojas ocultas `_Pendientes del cierre` (lo que lee el proceso), `_Extracto del cierre`,
  `_Foto FBS`, Parámetros, Reglas, Empresas grupo.
- FBS se lee de la base de la empresa (Parámetros, "Base FBS"): TURIN = 1uOGlMlonE9LVvT2rf9rrQxut5KFIuONADE5gGZeHtRc,
  hojas MOVIMIENTOS (se congela el 15 de cada mes) y CAMBIOS POST CIERRE.
- Sectores: S1 depósitos no registrados, S2 débitos no registrados, S3 depósitos no acreditados, S4 cheques no
  debitados. Ajustado = FBS + S1 − S2 − S3 + S4. "Compensado" sale solo si todo lo marcado suma cero
  (banco − FBS: S1 y S4 suman, S2 y S3 restan, tolerancia 0,02).
- Estados en "Pendientes en trabajo": Pendiente, En gestión, Resuelto en FBS, Compensado, Justificado, Pasar a diferencia.
  Se conservan por ID (Sector|Fecha|Concepto|Importe) al reprocesar.

## Planillas armadas con el motor (IDs de Google Sheets)

| Empresa | Banco | ID |
|---|---|---|
| TURIN | Galicia | 17CssoAhDjjQwiXXC_1dPcP_m8dUVv6sUH3CxacVXL6c |
| TURIN | HSBC | 1Vyyx2AiNLnEMyKRK-jS2CIr1XpLo-p-b-wAF928DPmc |
| TURIN | Nación | 1LlNVwF-gtbmvaw8qj9g44_ykCUyIkQPLDDKho8iuo9w |
| TURIN | Macro | 15KJTMjOAd5FGXGaCnvl78KCURFTUr1SImEWU1sHUl1Y |
| TURIN | Patagonia | 16cSeRJnM3tCUBReNA7C6tE3wHsucwSMeXmfgF6RKBFg |
| TURIN | Santander Río | 1mozoqXvhdYHHKvPgZ8guu-fx8thACH9t630qmo9GdNY |
| TURIN | Córdoba | 1Ie11VJYuXUMJKglt9k1kkATInVbXupPGjmCWKxDWRFw |
| JEA | Galicia | 1sXaacdPZqJLHnCmifdzfHVdVVDbNc3eONiShox4dFUw |
| JEA | Macro | 1Jk9viMtAeopCY7WPJQ17ueXIStwfi0qCNOWvXw9zAwI |
| JEA | Córdoba | 1GLjJAdgDXtwW_RaMKrJr0FZOEfrub0IAblJV5F6nYDo |
| CHEXA | Santander | 1DpFB201YacBA9NlCrQ6sEIj5ecNKIVd6RdP1t0kW8dU |
| CHEXA | Nación | 1yLzK0DqNPe4cGRkdD7xpVJJTBR_Vxm8J5eDWsOdLBeg |
| CHEXA | Macro | 1gfq0PAUV5e_WAccEs5PCOJIyksrbYGgC2aTGoypkqvw |
| CHEXA | Córdoba | 16Lgz5fi_e4hxIl1ajxDebcHmBZ5TXWqVSoqWKHfUT8U |

Cuadro de avance (lo completa el usuario): "AVANCE CONCILIACIONES AUT." 1vMm1pE1ojOr1JEaj2hYJ4kjMFISDPSR56x2NQwYCT1E.
Conciliación manual de octubre: Drive 1JsxxNtbU_SfNKIn3zzmiMqndSAeiXtyd (una hoja por banco).

Para leer una planilla entera: Google Drive `download_file_content` con exportMimeType xlsx y analizar con openpyxl
(`python3 -I`), guardando en el scratchpad.

## Estado al 13/11/2026 (TURIN)

| Banco | Último mes | Control | Notas |
|---|---|---|---|
| Galicia | septiembre cerrado | 104,35 | octubre sin procesar |
| HSBC | octubre cerrado | | |
| Nación | octubre cerrado | −35,11 | |
| Patagonia | octubre cerrado | 0,35 = manual | |
| Santander Río | septiembre cerrado (488,72 = manual); octubre en trabajo | 188,70 | apertura 488,72; el RM-392965 por 300 se pasó a diferencia (ajuste de DIAZ, GERMAN de septiembre) |
| Macro | agosto cerrado; septiembre en trabajo | 143,19 (manual 324,48: 181,29 de redondeos) | ver pendiente 1 |

### Pendientes inmediatos

1. **Macro TURIN septiembre — verificar.** El motor cruzaba en agosto pendientes viejos por una sola palabra del nombre
   (commit f039a0b lo corrige). El equipo corrigió a mano en "Pendientes en trabajo" y `_Pendientes del cierre`:
   GALLIAN 28/07 → RICHETT 24/06/2025 (50.000), BARRIOS/V 20/08 → FAVALLI 11/12/2025 (100.000), PEREYRA/S 28/08 →
   PEREYRA,J 01/06/2026 (250.000), borrando el ID, y reprocesó septiembre. Revisar: control 143,19, esas tres viejas
   pendientes y las nuevas no, ID regenerado, sin avisos nuevos. Después cierra septiembre y procesa octubre.
2. **Santander octubre.** El equipo compensó a mano: 7 pagos AFIP de 4.110 con RM-393604; Interbanking 04/08 850.091,55
   con RM-393666; GUIDA y SALLADAS (100.000 c/u) con canjes del 14 y 15/10; dos liquidaciones Payway (07/10 y 08/10) con
   sus Confirmaciones. Con el motor nuevo las liquidaciones y el Interbanking los cruza solo. Al reprocesar: control
   188,70 y sin avisos. "En gestión": LUDUEÑA 700.000 (duplicado con RC 00380961), Confirmación 14/09 615.000 (faltan
   gastos, se compensa contra Fiserv 28/08 560.873,48 cuando los registren), JEA S.A. RC-381068 569.799,49 vs PESSIO
   568.799,49 (dif. 1.000).
3. **Casillas en "Conciliación"** (commit 2b64193): nuevas, no probadas en una planilla real. Cada planilla tiene que
   tener `function onEdit(e) { Motor.alEditar(e); }` en su script para que tildar marque Compensado al instante (sin
   eso se aplica al procesar). Confirmar que aparecen, que la fila 3 suma bien (AFIP 7×4.110 + RM-393604 = 0) y que se
   pintan de verde.

### Para administración (FBS)

- Santander: anular las Confirmaciones duplicadas del 18/09 por 700.000 (liq 341277) y 350.000 (liq 341542); RC-378641
  FERNANDEZ 343.000 duplicado; canje 28/07 5.550 (Nº 22062026) duplicado; Percepción IIBB octubre: manual 27.734,02 vs
  motor 27.814,98.
- Macro: pase 150.000.000 del 30/09 duplicado ("BANCO MACRO" y "PAsefondo"); pase "BANCO MACRO" de la cuenta O del 29/09
  por 11.000.000 sin confirmar; BANCO MACRO Nº 653575 2.000.000; PEREA 3.000.000 en el banco vs RC 2.000.000; canje 25/09
  3.000.000; RM-392424 20.000; redondeos RC-378232 (60) y RC-379023 (121); débitos 21/09 IMPDBCR 440,70 y retención IIBB
  73.450,54 que no estaban en nuestro extracto; 7 importes chicos registrados en la cuenta de otro banco.
- Nación: Pasefondo 06/10 duplicado; RM-392664 5,57; reversión BREKES.
- Patagonia: RM-392007 fecha mal; RM-392244 sin fecha; 07/10 2M duplicado.

## Cambios recientes del motor (más nuevos primero)

- f039a0b Pendientes anteriores por nombre: apellido o dos palabras en común; gana la transferencia más parecida.
- 2b64193 Casilla por partida en Conciliación (columnas ocultas J:O con ID, valor con signo y estado escrito; fila 3 con
  fórmulas; `alEditar` y `sincronizarCasillas_`).
- b012820 Liquidaciones de tarjeta con créditos y débitos de otro texto (misma tarjeta y día, suma exacta); marca
  "Posible error de importe" (mismos centavos, diferencia redonda ≤ 2 %) en Cruces encontrados.
- 1bb67aa Santander regla 2376; pendiente anterior revertido con "ARREGLO"; gasto del banco viejo pagado con RM a banco.
- 4c79882 / bd3262d Liquidaciones sin cruzar se separan en sus líneas al final (salvo las ya trabajadas).
- 3602c13 Cuenta O: netear por Nº del valor (canjes de Dome Taller).

## Pruebas

Las pruebas de regresión usan datos reales guardados en el scratchpad de la sesión anterior (no están en el repo y no
pasan a una sesión nueva). En una sesión nueva hay que rearmarlas bajando las planillas (xlsx) y la base: la función
`conciliarTodo_(entrada)` corre sin hojas (`extracto`, `base`, `saldosIniciales`, `anteriores`, `trabajo`, `reglas`,
`empresas`, `parametros`, `desconciliar`) y se puede ejecutar con Node cargando el `.gs` en un `vm`. Como mínimo, antes
de cambiar el motor, armar una prueba con el banco afectado y otra con un banco distinto, y comparar el resultado
partida por partida contra el motor anterior (`git show HEAD~1:conciliacion/appscript/Conciliacion.gs`).

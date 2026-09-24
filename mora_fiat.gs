/**
 * MORA PLAN DE AHORRO - Tableros de seguimiento (FIAT)
 *
 * Ejecutar una sola vez: construirTableros()
 * Crea (o recrea) las hojas de abajo con FÓRMULAS VIVAS: cuando se actualiza
 * la hoja BASE, todos los tableros se recalculan solos. La hoja BASE no se toca.
 * La hoja PARAMETROS (tabla de incentivos) solo se crea la primera vez: editala ahí.
 * La hoja Historial diario nunca se borra: cada día hábil se le agrega la foto del Tablero Estado Actual.
 * En PARAMETROS (columna H) queda la fecha en que la BASE cambia de avance cada mes: el script la
 * detecta sola (cuando la mayoría de los planes suma +1 al avance) y también se puede cargar/corregir a mano.
 * De ahí salen el "período medido" de FIAT y el "avance vigente desde" de los tableros.
 *
 * Columnas de BASE que se usan:
 *   A RESPONSABLE | B SOLICITUD | C GRUPO | D Orden | H NyAP | I TELEFONO
 *   N Avance | O Estado | P Vendedor | Q Supervisor | R..AD = C2..C14
 *
 * Reglas:
 *   - RESCINDIDO: columna Estado = Rescindido o Renunciado (sin importar las cuotas).
 *   - MORA: al menos una "I" en las cuotas evaluadas.
 *   - AL DIA: sin "I" en las cuotas evaluadas (incluye Cancelado, cuotas con "C").
 *   - % mora = (planes en mora + rescindidos) / cartera total (al día + mora + rescindidos).
 *   - Medición FIAT (mes vencido): plan en avance N hoy se mide como cuota N-1,
 *     evaluando C2..C(N-1). Cuotas medidas: 3, 5, 7, 9 y 12.
 */

// --- CONSTANTES COMPARTIDAS CON EL CRM (crm_mora.gs las usa: no renombrar) ---
const HOJA_BASE = 'BASE';
const COL_AVANCE = 14; // N
const COL_ESTADO = 15; // O
const COL_C2 = 18;     // R = primera cuota (C2)
const CUOTAS_TABLERO = [3, 5, 7, 9, 12]; // cuotas que mide FIAT

const HOJAS = {
  PARAM: 'PARAMETROS',
  CALC: 'CALC',
  FIAT: 'Tablero Medición FIAT',
  ACTUAL: 'Tablero Estado Actual',
  DET_FIAT: 'Detalle Mora FIAT',
  GESTION: 'Gestión Mes',
  HIST: 'Historial diario',
  AV_PREVIO: '_AVANCE_PREVIO', // hoja oculta: último avance visto de cada solicitud
};

const HORA_FOTO_DIARIA = 20; // hora (0-23) en que se guarda sola la foto del día en el historial (solo días hábiles)

// Feriados (formato 'aaaa-mm-dd'): esos días no se guarda la foto automática. Agregá los que falten.
const FERIADOS = [
  '2026-10-12',
  '2026-12-08',
  '2026-12-25',
];

// Cuota | Tramo A: mora menor a | Tramo B desde | Tramo B hasta | % pago A | % pago B
const TABLA_INCENTIVO = [
  [3, 0.35, 0.35, 0.41, 0.0035, 0.0020],
  [5, 0.30, 0.30, 0.36, 0.0065, 0.0045],
  [7, 0.32, 0.32, 0.38, 0.0070, 0.0040],
  [9, 0.45, 0.45, 0.51, 0.0070, 0.0040],
  [12, 0.48, 0.48, 0.54, 0.0080, 0.0040],
];

const COLOR_HEADER = '#1f3864';
const COLOR_TITULO = '#c00000';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mora FIAT')
    .addItem('Construir / reconstruir tableros', 'construirTableros')
    .addItem('Guardar foto de hoy en el historial', 'guardarHistorial')
    .addItem('Detectar cambio de avance ahora', 'detectarCambioAvance')
    .addSeparator()
    .addItem('CRM: inicializar hojas y usuarios', 'crmInicializar') // ver crm_mora.gs
    .addItem('CRM: importar notas viejas de BASE', 'crmImportarNotasBase')
    .addToUi();
  // Al abrir el archivo también se revisa si la BASE cambió de avance
  try { detectarCambioAvance(); } catch (e) { console.log('detectarCambioAvance: ' + e); }
}

function construirTableros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(HOJA_BASE)) throw new Error('No existe la hoja ' + HOJA_BASE);

  crearParametros_(ss);
  asegurarCalendarioAvance_(ss);
  detectarCambioAvance();
  crearCalc_(ss);
  crearTableroFiat_(ss);
  crearTableroActual_(ss);
  crearDetalleFiat_(ss);
  crearGestionMes_(ss);
  crearHistorial_(ss);
  activarHistorialDiario_();
  guardarHistorial();

  ss.setActiveSheet(ss.getSheetByName(HOJAS.FIAT));
}

/** Borra y recrea una hoja (nunca toca BASE). */
function hojaNueva_(ss, nombre) {
  const vieja = ss.getSheetByName(nombre);
  if (vieja) ss.deleteSheet(vieja);
  return ss.insertSheet(nombre);
}

function estiloHeader_(rango) {
  rango.setBackground(COLOR_HEADER).setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function titulo_(sh, texto, subtitulo) {
  sh.getRange('A1').setValue(texto).setFontSize(14).setFontWeight('bold').setFontColor(COLOR_TITULO);
  if (subtitulo) sh.getRange('A2').setFormula(subtitulo).setFontStyle('italic');
}

// ---------------------------------------------------------------- PARAMETROS
function crearParametros_(ss) {
  // Si ya existe se respeta: así no se pierden los cambios que se hagan a mano en la tabla.
  if (ss.getSheetByName(HOJAS.PARAM)) return;
  const sh = ss.insertSheet(HOJAS.PARAM);
  titulo_(sh, 'Tabla de incentivos por mora (editable)');
  const enc = ['CUOTA', 'TRAMO A: MORA MENOR A', 'TRAMO B: DESDE', 'TRAMO B: HASTA', '% PAGO A', '% PAGO B'];
  estiloHeader_(sh.getRange(2, 1, 1, enc.length).setValues([enc]));
  sh.getRange(3, 1, TABLA_INCENTIVO.length, 6).setValues(TABLA_INCENTIVO);
  sh.getRange('B3:D7').setNumberFormat('0%');
  sh.getRange('E3:F7').setNumberFormat('0.00%');
  sh.getRange('A8').setValue('TOTAL').setFontWeight('bold');
  sh.getRange('E8').setFormula('=SUM(E3:E7)').setNumberFormat('0.00%').setFontWeight('bold');
  sh.getRange('F8').setFormula('=SUM(F3:F7)').setNumberFormat('0.00%').setFontWeight('bold');
  sh.setColumnWidths(1, 6, 130);
}

// ---------------------------------------------------------------- CALENDARIO DE CAMBIO DE AVANCE
// Tabla en PARAMETROS!H:I. Se carga a mano la fecha en que la BASE pasa al avance siguiente (+1).
// Ej.: 21/09/2026 => desde ese día la BASE muestra el avance de septiembre = período de medición septiembre.
const FECHA_INICIAL_AVANCE = [2026, 8, 21]; // 21/09/2026 (mes base 0)

/** Agrega la tabla de fechas a PARAMETROS si todavía no existe (no toca lo cargado). */
function asegurarCalendarioAvance_(ss) {
  const sh = ss.getSheetByName(HOJAS.PARAM);
  // La columna I se recalcula siempre (período = mes de la fecha de cambio de avance)
  const formulaPeriodo = '=ARRAYFORMULA(IF(H3:H200="","",PROPER(TEXT(H3:H200,"mmmm yyyy"))))';
  if (sh.getRange('H2').getValue() !== '') {
    sh.getRange('I2').setValue('PERÍODO DE MEDICIÓN FIAT');
    sh.getRange('I3').setFormula(formulaPeriodo);
    return;
  }
  sh.getRange('H1').setValue('Fechas de cambio de avance (cargar cada mes)').setFontWeight('bold').setFontColor(COLOR_TITULO);
  estiloHeader_(sh.getRange('H2:I2').setValues([['FECHA CAMBIO DE AVANCE', 'PERÍODO DE MEDICIÓN FIAT']]));
  sh.getRange('H3').setValue(new Date(FECHA_INICIAL_AVANCE[0], FECHA_INICIAL_AVANCE[1], FECHA_INICIAL_AVANCE[2]));
  sh.getRange('H3:H200').setNumberFormat('dd/mm/yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Fecha en que la BASE pasó al avance siguiente').build());
  sh.getRange('I3').setFormula(formulaPeriodo);
  sh.getRange('H3:I200').setHorizontalAlignment('center');
  sh.setColumnWidths(8, 2, 170);
  sh.getRange('H2').setNote('Cargá una fila por mes con el día en que la BASE sumó +1 al avance. ' +
    'Los tableros toman la última fecha que sea menor o igual a hoy.');
}

/**
 * Arma una fórmula donde f = última fecha de cambio de avance <= hoy.
 * Mientras no haya una fecha cargada del mes en curso, antepone un aviso.
 */
function formulaAvance_(texto) {
  return 'LET(u, MAXIFS(PARAMETROS!H3:H200, PARAMETROS!H3:H200, "<="&TODAY()), ' +
    'f, IF(u=0, DATE(YEAR(TODAY()),MONTH(TODAY()),1), u), ' +
    'aviso, IF(u < DATE(YEAR(TODAY()),MONTH(TODAY()),1), "⚠ Sin cambio de avance cargado este mes (si la BASE ya cambió, cargá la fecha en PARAMETROS)  —  ", ""), ' +
    'aviso&' + texto + ')';
}

/** Última fecha de cambio de avance <= hoy (o null si no hay ninguna cargada). */
function ultimoCambioAvance_(ss) {
  const sh = ss.getSheetByName(HOJAS.PARAM);
  if (!sh) return null;
  const hoy = new Date();
  let ultima = null;
  sh.getRange('H3:H200').getValues().forEach(function (r) {
    const d = r[0];
    if (d instanceof Date && d <= hoy && (!ultima || d > ultima)) ultima = d;
  });
  return ultima;
}

// ---------------------------------------------------------------- DETECCIÓN AUTOMÁTICA DEL CAMBIO DE AVANCE
const MIN_PLANES_COMPARABLES = 20;   // mínimo de solicitudes en común para decidir
const PROPORCION_CAMBIO = 0.5;       // si al menos la mitad sumó +1 al avance => cambió el mes

/**
 * Compara el avance de cada solicitud de BASE con el último guardado (hoja oculta _AVANCE_PREVIO).
 * Si la mayoría sumó +1, carga la fecha de hoy en PARAMETROS!H (una sola vez por mes).
 * Se ejecuta al abrir el archivo, todos los días con el disparador y desde el menú.
 */
function detectarCambioAvance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName(HOJA_BASE);
  const param = ss.getSheetByName(HOJAS.PARAM);
  if (!base || !param) return;

  // Avance actual por solicitud (B = SOLICITUD, N = Avance)
  const ultimaFila = base.getLastRow();
  if (ultimaFila < 2) return;
  const actual = {};
  const filasActuales = [];
  base.getRange(2, 2, ultimaFila - 1, 13).getValues().forEach(function (r) {
    const sol = String(r[0]).trim();
    const av = Number(r[12]);
    if (sol && !isNaN(av) && r[12] !== '') {
      actual[sol] = av;
      filasActuales.push([sol, av]);
    }
  });
  if (!filasActuales.length) return;

  // Avance guardado la vez anterior
  let prev = ss.getSheetByName(HOJAS.AV_PREVIO);
  if (!prev) {
    prev = ss.insertSheet(HOJAS.AV_PREVIO);
    prev.hideSheet();
  }
  const anterior = {};
  if (prev.getLastRow() > 0) {
    prev.getRange(1, 1, prev.getLastRow(), 2).getValues().forEach(function (r) {
      if (r[0] !== '') anterior[String(r[0])] = Number(r[1]);
    });
  }

  // ¿La mayoría de las solicitudes en común sumó +1?
  let comunes = 0, suben = 0;
  Object.keys(actual).forEach(function (sol) {
    if (sol in anterior) {
      comunes++;
      if (actual[sol] === anterior[sol] + 1) suben++;
    }
  });

  if (comunes >= MIN_PLANES_COMPARABLES && suben / comunes >= PROPORCION_CAMBIO) {
    registrarCambioAvance_(ss, param);
  }

  // Guarda la foto actual para la próxima comparación
  prev.clearContents();
  prev.getRange(1, 1, filasActuales.length, 2).setValues(filasActuales);
}

/** Carga la fecha de hoy en PARAMETROS!H salvo que ya haya una fecha de este mes. */
function registrarCambioAvance_(ss, param) {
  const tz = ss.getSpreadsheetTimeZone();
  const ahora = new Date();
  const mesActual = Utilities.formatDate(ahora, tz, 'yyyy-MM');
  const fechas = param.getRange('H3:H200').getValues();
  let primeraVacia = -1;
  for (let i = 0; i < fechas.length; i++) {
    const d = fechas[i][0];
    if (d instanceof Date && Utilities.formatDate(d, tz, 'yyyy-MM') === mesActual) return; // ya cargada
    if (d === '' && primeraVacia === -1) primeraVacia = i;
  }
  if (primeraVacia === -1) return;
  const celda = param.getRange(3 + primeraVacia, 8);
  celda.setValue(Utilities.parseDate(Utilities.formatDate(ahora, tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd'));
  celda.setNote('Detectado automáticamente el ' + Utilities.formatDate(ahora, tz, 'dd/MM/yyyy HH:mm') +
    '. Si el cambio fue otro día, corregí la fecha.');
}

// ---------------------------------------------------------------- CALC (una fila por plan)
function crearCalc_(ss) {
  const sh = hojaNueva_(ss, HOJAS.CALC);
  // Nro de cuota de cada columna R..AD = 2..14. "Vencidas" = cuotas anteriores al avance actual
  // (= lo que FIAT mide el mes siguiente con avance-1). La cuota del avance actual vence a fin de mes.
  const f = `=ARRAYFORMULA(LET(
  d, FILTER(BASE!A2:AD, BASE!B2:B<>""),
  n, ROWS(d),
  av, IFERROR(CHOOSECOLS(d,14)*1, 0),
  es, CHOOSECOLS(d,15),
  imp, --(CHOOSECOLS(d,SEQUENCE(1,13,18))="I"),
  uno, SEQUENCE(13,1,1,0),
  nro, MMULT(SEQUENCE(n,1,1,0), SEQUENCE(1,13,2)),
  avM, MMULT(av, SEQUENCE(1,13,1,0)),
  resc, REGEXMATCH(es&"", "(?i)rescind|renunc"),
  iTot, MMULT(imp, uno),
  iVen, MMULT(imp*(nro<avM), uno),
  iMes, MMULT(imp*(nro=avM), uno),
  cF, av-1,
  mide, ISNUMBER(MATCH(cF, {3;5;7;9;12}, 0)),
  estAct, IF(resc, "RESCINDIDO", IF(iTot>0, "MORA", "AL DIA")),
  estF, IF(mide, IF(resc, "RESCINDIDO", IF(iVen>0, "MORA", "AL DIA")), ""),
  VSTACK(
    {"SOLICITUD","GRUPO","ORDEN","CLIENTE","TELEFONO","RESPONSABLE","VENDEDOR","SUPERVISOR","AVANCE","ESTADO BASE","ESTADO ACTUAL","CUOTAS IMPAGAS","IMPAGAS VENCIDAS","DEBE CUOTA DEL MES","CUOTA MEDIDA FIAT","ESTADO FIAT"},
    HSTACK(CHOOSECOLS(d,2,3,4,8,9,1,16,17), av, es, estAct, iTot, iVen, IF(iMes>0,"SI","NO"), IF(mide,cF,""), estF)
  )
))`;
  sh.getRange('A1').setFormula(f);
  estiloHeader_(sh.getRange('A1:P1'));
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('Hoja de cálculo auxiliar. No escribir acá: se genera sola desde BASE.');
}

// ---------------------------------------------------------------- TABLERO MEDICION FIAT
function crearTableroFiat_(ss) {
  const sh = hojaNueva_(ss, HOJAS.FIAT);
  titulo_(sh, 'TABLERO MEDICIÓN FIAT (mes vencido)',
    '=' + formulaAvance_('"Período medido: "&PROPER(TEXT(f,"mmmm yyyy"))&"  —  avance vigente desde "&TEXT(f,"dd/mm/yyyy")&"  —  planes que hoy están en avance N, evaluados en cuotas C2 a C(N-1)"'));

  const enc = ['CUOTA MEDIDA', 'AVANCE EN BASE HOY', 'CARTERA TOTAL', 'AL DÍA', 'EN MORA', 'RESCINDIDOS',
    '% MORA (MORA + RESC.)', 'TRAMO A: MENOR A', 'TRAMO B: HASTA', 'TRAMO LOGRADO', '% INCENTIVO', 'FALTÓ P/ TRAMO A (planes)', 'FALTÓ P/ TRAMO B (planes)'];
  estiloHeader_(sh.getRange(4, 1, 1, enc.length).setValues([enc]));
  sh.setRowHeight(4, 45);

  const P = 'PARAMETROS!$A$3:$F$7';
  const filas = [];
  for (let i = 0; i < TABLA_INCENTIVO.length; i++) {
    const r = 5 + i;
    filas.push([
      TABLA_INCENTIVO[i][0],
      `=A${r}+1`,
      `=COUNTIF(CALC!$O$2:$O,A${r})`,
      `=COUNTIFS(CALC!$O$2:$O,A${r},CALC!$P$2:$P,"AL DIA")`,
      `=COUNTIFS(CALC!$O$2:$O,A${r},CALC!$P$2:$P,"MORA")`,
      `=COUNTIFS(CALC!$O$2:$O,A${r},CALC!$P$2:$P,"RESCINDIDO")`,
      `=IFERROR((E${r}+F${r})/C${r},0)`,
      `=VLOOKUP(A${r},${P},2,0)`,
      `=VLOOKUP(A${r},${P},4,0)`,
      `=IF(C${r}=0,"SIN DATOS",IF(ROUND(G${r},6)<H${r},"A",IF(ROUND(G${r},6)<=I${r},"B","SIN COBRO")))`,
      `=IF(J${r}="A",VLOOKUP(A${r},${P},5,0),IF(J${r}="B",VLOOKUP(A${r},${P},6,0),0))`,
      `=IF(C${r}=0,0,MAX(0,E${r}+F${r}-(CEILING(ROUND(H${r}*C${r},6),1)-1)))`,
      `=IF(C${r}=0,0,MAX(0,E${r}+F${r}-FLOOR(ROUND(I${r}*C${r},6),1)))`,
    ]);
  }
  sh.getRange(5, 1, filas.length, enc.length).setValues(filas);

  sh.getRange('A10').setValue('TOTAL');
  sh.getRange('C10').setFormula('=SUM(C5:C9)');
  sh.getRange('D10').setFormula('=SUM(D5:D9)');
  sh.getRange('E10').setFormula('=SUM(E5:E9)');
  sh.getRange('F10').setFormula('=SUM(F5:F9)');
  sh.getRange('G10').setFormula('=IFERROR((E10+F10)/C10,0)');
  sh.getRange('J10').setValue('INCENTIVO TOTAL');
  sh.getRange('K10').setFormula('=SUM(K5:K9)');
  sh.getRange('A10:M10').setFontWeight('bold').setBackground('#d9e1f2');

  sh.getRange('A12').setValue('Máximo posible tramo A:');
  sh.getRange('C12').setFormula('=PARAMETROS!E8').setNumberFormat('0.00%');
  sh.getRange('A13').setValue('Máximo posible tramo B:');
  sh.getRange('C13').setFormula('=PARAMETROS!F8').setNumberFormat('0.00%');

  sh.getRange('G5:I10').setNumberFormat('0.0%');
  sh.getRange('K5:K10').setNumberFormat('0.00%');
  sh.getRange('A5:M10').setHorizontalAlignment('center');
  sh.getRange('A4:M10').setBorder(true, true, true, true, true, true);

  colorTramos_(sh, sh.getRange('J5:J9'));
  sh.setColumnWidths(1, 13, 110);
  sh.setFrozenRows(4);
}

function colorTramos_(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('A').setBackground('#c6efce').setFontColor('#006100').setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('B').setBackground('#ffeb9c').setFontColor('#9c5700').setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SIN COBRO').setBackground('#ffc7ce').setFontColor('#9c0006').setRanges([rango]).build()
  );
  sh.setConditionalFormatRules(reglas);
}

// ---------------------------------------------------------------- TABLERO ESTADO ACTUAL
function crearTableroActual_(ss) {
  const sh = hojaNueva_(ss, HOJAS.ACTUAL);
  titulo_(sh, 'TABLERO ESTADO ACTUAL POR AVANCE',
    '=' + formulaAvance_('"Foto al "&TEXT(TODAY(),"dd/mm/yyyy")&"  —  avance vigente desde "&TEXT(f,"dd/mm/yyyy")&" (mes de avance: "&PROPER(TEXT(f,"mmmm yyyy"))&")  —  los avances 3, 5, 7, 9 y 12 son los que FIAT mide el mes próximo."'));

  const f = `=LET(
  av, CALC!I2:I, e, CALC!K2:K, p, PARAMETROS!A3:F7,
  lst, SORT(UNIQUE(FILTER(av, ISNUMBER(av), av>0))),
  REDUCE(
    {"AVANCE","CARTERA TOTAL","AL DÍA","EN MORA","RESCINDIDOS","% MORA (MORA + RESC.)","PRÓX. MEDICIÓN FIAT","TRAMO A: MENOR A","PLANES A REGULARIZAR P/ TRAMO A","PLANES A REGULARIZAR P/ TRAMO B"},
    lst,
    LAMBDA(acc, a, LET(
      tot, COUNTIF(av, a),
      ald, COUNTIFS(av, a, e, "AL DIA"),
      mo, COUNTIFS(av, a, e, "MORA"),
      rs, COUNTIFS(av, a, e, "RESCINDIDO"),
      mide, ISNUMBER(MATCH(a, {3;5;7;9;12}, 0)),
      ua, IF(mide, VLOOKUP(a, p, 2, 0), ""),
      ub, IF(mide, VLOOKUP(a, p, 4, 0), ""),
      VSTACK(acc, HSTACK(
        a, tot, ald, mo, rs, IFERROR((mo+rs)/tot, 0),
        IF(mide, "Cuota "&a, ""), ua,
        IF(mide, MAX(0, mo+rs-(CEILING(ROUND(ua*tot,6),1)-1)), ""),
        IF(mide, MAX(0, mo+rs-FLOOR(ROUND(ub*tot,6),1)), "")
      ))
    ))
  )
)`;
  sh.getRange('A4').setFormula(f);
  estiloHeader_(sh.getRange('A4:J4'));
  sh.setRowHeight(4, 60);
  sh.getRange('F5:F100').setNumberFormat('0.0%');
  sh.getRange('H5:H100').setNumberFormat('0%');
  sh.getRange('A5:J100').setHorizontalAlignment('center');

  // Resalta las filas que FIAT mide el mes próximo
  const reglas = sh.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G5<>""').setBackground('#fff2cc').setRanges([sh.getRange('A5:J100')]).build());
  sh.setConditionalFormatRules(reglas);

  sh.setColumnWidths(1, 10, 110);
  sh.setFrozenRows(4);
}

// ---------------------------------------------------------------- DETALLES
function crearDetalleFiat_(ss) {
  const sh = hojaNueva_(ss, HOJAS.DET_FIAT);
  titulo_(sh, 'Planes en MORA en la medición FIAT del mes vencido');
  sh.getRange('A3').setFormula(
    '=VSTACK(CALC!A1:P1, IFERROR(SORT(FILTER(CALC!A2:P, CALC!P2:P="MORA"), 15, TRUE, 6, TRUE), "Sin planes en mora"))');
  estiloHeader_(sh.getRange('A3:P3'));
  sh.setFrozenRows(3);
}

function crearGestionMes_(ss) {
  const sh = hojaNueva_(ss, HOJAS.GESTION);
  titulo_(sh, 'GESTIÓN DEL MES: planes en mora que FIAT mide el mes próximo (avances 3, 5, 7, 9 y 12)');
  sh.getRange('A3').setFormula(
    '=VSTACK(CALC!A1:P1, IFERROR(SORT(FILTER(CALC!A2:P, CALC!K2:K="MORA", ISNUMBER(MATCH(CALC!I2:I, {3;5;7;9;12}, 0))), 9, TRUE, 6, TRUE), "Sin planes en mora"))');
  estiloHeader_(sh.getRange('A3:P3'));
  sh.setFrozenRows(3);
}

// ---------------------------------------------------------------- HISTORIAL DIARIO
const ENC_HISTORIAL = ['FECHA', 'AVANCE', 'CARTERA TOTAL', 'AL DÍA', 'EN MORA', 'RESCINDIDOS', '% MORA (MORA + RESC.)',
  'PRÓX. MEDICIÓN FIAT', 'TRAMO A: MENOR A', 'PLANES A REGULARIZAR P/ TRAMO A', 'PLANES A REGULARIZAR P/ TRAMO B',
  'AVANCE VIGENTE DESDE'];
const COLS_TABLERO_ACTUAL = 10; // columnas A:J del Tablero Estado Actual que se copian al historial

/** Crea la hoja solo si no existe: el historial nunca se borra. */
function crearHistorial_(ss) {
  if (ss.getSheetByName(HOJAS.HIST)) return;
  const sh = ss.insertSheet(HOJAS.HIST);
  estiloHeader_(sh.getRange(1, 1, 1, ENC_HISTORIAL.length).setValues([ENC_HISTORIAL]));
  sh.setRowHeight(1, 45);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, ENC_HISTORIAL.length, 110);
}

/** Deja un único disparador diario que llama a fotoDiariaAutomatica (reemplaza versiones anteriores). */
function activarHistorialDiario_() {
  let existe = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    // disparadores de versiones anteriores (incluye snapshotDiario del script viejo)
    if (fn === 'guardarHistorial' || fn === 'snapshotDiario') ScriptApp.deleteTrigger(t);
    if (fn === 'fotoDiariaAutomatica') existe = true;
  });
  if (!existe) {
    ScriptApp.newTrigger('fotoDiariaAutomatica').timeBased().everyDays(1).atHour(HORA_FOTO_DIARIA).create();
  }
}

/** La ejecuta el disparador todos los días: revisa el cambio de avance y guarda la foto solo de lunes a viernes (no feriados). */
function fotoDiariaAutomatica() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const ahora = new Date();
  const diaSemana = Number(Utilities.formatDate(ahora, tz, 'u')); // 1 = lunes ... 7 = domingo
  const hoy = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
  detectarCambioAvance(); // todos los días, incluso fines de semana
  if (diaSemana >= 6 || FERIADOS.indexOf(hoy) !== -1) return;
  guardarHistorial();
}

/**
 * Copia el Tablero Estado Actual al Historial diario con la fecha de hoy.
 * Si ya había una foto de hoy, la reemplaza (queda una sola foto por día, la última).
 */
function guardarHistorial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = ss.getSheetByName(HOJAS.ACTUAL);
  if (!tab) throw new Error('No existe la hoja ' + HOJAS.ACTUAL + '. Ejecutá construirTableros primero.');
  crearHistorial_(ss);
  const hist = ss.getSheetByName(HOJAS.HIST);
  // Mantiene el encabezado al día si se agregaron columnas nuevas
  estiloHeader_(hist.getRange(1, 1, 1, ENC_HISTORIAL.length).setValues([ENC_HISTORIAL]));

  SpreadsheetApp.flush();
  const filas = Math.max(tab.getLastRow() - 4, 1);
  const datos = tab.getRange(5, 1, filas, COLS_TABLERO_ACTUAL).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; });
  if (!datos.length) return;

  const tz = ss.getSpreadsheetTimeZone();
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Borra la foto de hoy si ya existía
  const ultima = hist.getLastRow();
  if (ultima > 1) {
    const fechas = hist.getRange(2, 1, ultima - 1, 1).getValues();
    for (let i = fechas.length - 1; i >= 0; i--) {
      const f = fechas[i][0];
      if (f instanceof Date && Utilities.formatDate(f, tz, 'yyyy-MM-dd') === hoy) hist.deleteRow(i + 2);
    }
  }

  const fecha = Utilities.parseDate(hoy, tz, 'yyyy-MM-dd');
  const vigente = ultimoCambioAvance_(ss) || '';
  const nuevas = datos.map(function (r) { return [fecha].concat(r, [vigente]); });
  const desde = hist.getLastRow() + 1;
  hist.getRange(desde, 1, nuevas.length, ENC_HISTORIAL.length).setValues(nuevas).setHorizontalAlignment('center');
  hist.getRange(desde, 1, nuevas.length, 1).setNumberFormat('dd/mm/yyyy');
  hist.getRange(desde, 7, nuevas.length, 1).setNumberFormat('0.0%');
  hist.getRange(desde, 9, nuevas.length, 1).setNumberFormat('0%');
  hist.getRange(desde, ENC_HISTORIAL.length, nuevas.length, 1).setNumberFormat('dd/mm/yyyy');
  // Línea separadora entre días
  hist.getRange(desde, 1, 1, ENC_HISTORIAL.length)
    .setBorder(true, null, null, null, null, null, '#1f3864', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

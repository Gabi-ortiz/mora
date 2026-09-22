/**
 * MORA PLAN DE AHORRO - Tableros de seguimiento (FIAT)
 *
 * Ejecutar una sola vez: construirTableros()
 * Crea (o recrea) las hojas de abajo con FÓRMULAS VIVAS: cuando se actualiza
 * la hoja BASE, todos los tableros se recalculan solos. La hoja BASE no se toca.
 * La hoja PARAMETROS (tabla de incentivos) solo se crea la primera vez: editala ahí.
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

const HOJAS = {
  PARAM: 'PARAMETROS',
  CALC: 'CALC',
  FIAT: 'Tablero Medición FIAT',
  ACTUAL: 'Tablero Estado Actual',
  DET_FIAT: 'Detalle Mora FIAT',
  GESTION: 'Gestión Mes',
};

// Cuota | Tramo A: mora menor a | Tramo B desde | Tramo B hasta | % pago A | % pago B
const TABLA_INCENTIVO = [
  [3, 0.35, 0.35, 0.41, 0.0035, 0.0020],
  [5, 0.30, 0.30, 0.36, 0.0065, 0.0045],
  [7, 0.32, 0.32, 0.38, 0.0070, 0.0045],
  [9, 0.45, 0.45, 0.51, 0.0070, 0.0040],
  [12, 0.48, 0.48, 0.54, 0.0080, 0.0040],
];

const COLOR_HEADER = '#1f3864';
const COLOR_TITULO = '#c00000';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mora FIAT')
    .addItem('Construir / reconstruir tableros', 'construirTableros')
    .addToUi();
}

function construirTableros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('BASE')) throw new Error('No existe la hoja BASE');

  crearParametros_(ss);
  crearCalc_(ss);
  crearTableroFiat_(ss);
  crearTableroActual_(ss);
  crearDetalleFiat_(ss);
  crearGestionMes_(ss);

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
    '="Período medido: "&PROPER(TEXT(EOMONTH(TODAY(),-1),"mmmm yyyy"))&"  —  planes que hoy están en avance N, evaluados en cuotas C2 a C(N-1)"');

  const enc = ['CUOTA MEDIDA', 'AVANCE EN BASE HOY', 'CARTERA TOTAL', 'AL DÍA', 'EN MORA', 'RESCINDIDOS',
    '% MORA (MORA + RESC.)', 'TRAMO A: MENOR A', 'TRAMO B: HASTA', 'TRAMO LOGRADO', '% INCENTIVO', 'FALTÓ P/ TRAMO A (planes)'];
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
  sh.getRange('A10:L10').setFontWeight('bold').setBackground('#d9e1f2');

  sh.getRange('A12').setValue('Máximo posible tramo A:');
  sh.getRange('C12').setFormula('=PARAMETROS!E8').setNumberFormat('0.00%');
  sh.getRange('A13').setValue('Máximo posible tramo B:');
  sh.getRange('C13').setFormula('=PARAMETROS!F8').setNumberFormat('0.00%');

  sh.getRange('G5:I10').setNumberFormat('0.0%');
  sh.getRange('K5:K10').setNumberFormat('0.00%');
  sh.getRange('A5:L10').setHorizontalAlignment('center');
  sh.getRange('A4:L10').setBorder(true, true, true, true, true, true);

  colorTramos_(sh, sh.getRange('J5:J9'));
  sh.setColumnWidths(1, 12, 110);
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
    '="Foto al "&TEXT(TODAY(),"dd/mm/yyyy")&"  —  la cuota del avance actual vence a fin de mes. Los avances 3, 5, 7, 9 y 12 son los que FIAT mide el mes próximo."');

  const f = `=LET(
  av, CALC!I2:I, e, CALC!K2:K, v, CALC!M2:M, p, PARAMETROS!A3:F7,
  lst, SORT(UNIQUE(FILTER(av, ISNUMBER(av), av>0))),
  REDUCE(
    {"AVANCE","CARTERA TOTAL","AL DÍA","EN MORA","RESCINDIDOS","% MORA (MORA + RESC.)","EN MORA SOLO POR CUOTA DEL MES","MORA VENCIDA","% MORA VENCIDA (VENCIDA + RESC.)","PRÓX. MEDICIÓN FIAT","TRAMO A: MENOR A","PLANES A REGULARIZAR P/ TRAMO A","PLANES A REGULARIZAR P/ TRAMO B"},
    lst,
    LAMBDA(acc, a, LET(
      tot, COUNTIF(av, a),
      ald, COUNTIFS(av, a, e, "AL DIA"),
      mo, COUNTIFS(av, a, e, "MORA"),
      rs, COUNTIFS(av, a, e, "RESCINDIDO"),
      ven, COUNTIFS(av, a, e, "MORA", v, ">0"),
      mide, ISNUMBER(MATCH(a, {3;5;7;9;12}, 0)),
      ua, IF(mide, VLOOKUP(a, p, 2, 0), ""),
      ub, IF(mide, VLOOKUP(a, p, 4, 0), ""),
      VSTACK(acc, HSTACK(
        a, tot, ald, mo, rs, IFERROR((mo+rs)/tot, 0), mo-ven, ven, IFERROR((ven+rs)/tot, 0),
        IF(mide, "Cuota "&a, ""), ua,
        IF(mide, MAX(0, mo+rs-(CEILING(ROUND(ua*tot,6),1)-1)), ""),
        IF(mide, MAX(0, mo+rs-FLOOR(ROUND(ub*tot,6),1)), "")
      ))
    ))
  )
)`;
  sh.getRange('A4').setFormula(f);
  estiloHeader_(sh.getRange('A4:M4'));
  sh.setRowHeight(4, 60);
  sh.getRange('F5:F100').setNumberFormat('0.0%');
  sh.getRange('I5:I100').setNumberFormat('0.0%');
  sh.getRange('K5:K100').setNumberFormat('0%');
  sh.getRange('A5:M100').setHorizontalAlignment('center');

  // Resalta las filas que FIAT mide el mes próximo
  const reglas = sh.getConditionalFormatRules();
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J5<>""').setBackground('#fff2cc').setRanges([sh.getRange('A5:M100')]).build());
  sh.setConditionalFormatRules(reglas);

  sh.setColumnWidths(1, 13, 110);
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

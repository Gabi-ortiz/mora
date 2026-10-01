/**
 * PEDIDOS PDA — base acumulada de adjudicados (SGA)
 *
 * Reemplaza el esquema de "un archivo por mes" (AGOSTO 26, SEPTIEMBRE 26...)
 * por un solo archivo con todas las adjudicaciones acumuladas en BASE.
 *
 * Hojas que arma "Armar / reparar estructura":
 * - BASE: una fila por adjudicación (clave = Acto + Solicitud). Mismas
 *   columnas y orden que la hoja mensual de siempre, más "Acto" al principio
 *   y "Mes" renombrado a "Mes cierre". Desplegables en las columnas de
 *   trabajo para que no aparezcan variantes de nombres (CLARI / CLARISA).
 * - PEGAR SGA: se pega el reporte de SGA tal cual se baja (un acto).
 * - RESUMEN: el tablero del mes, con el mes elegido en B1.
 * - INFORME: evolución mes a mes (acto y cierres) y por responsable.
 * - PRECIOS: V.M por mes y modelo (se carga una vez por mes).
 * - OBJETIVOS: objetivo del mes (llega por circular).
 * - LISTAS: valores de los desplegables.
 *
 * Instalación: ver PEDIDOS_PDA.md.
 */

// --- CONFIGURACIÓN ---
const HOJA_BASE = 'BASE';
const HOJA_PEGAR = 'PEGAR SGA';
const HOJA_RESUMEN = 'RESUMEN';
const HOJA_INFORME = 'INFORME';
const HOJA_PRECIOS = 'PRECIOS';
const HOJA_OBJETIVOS = 'OBJETIVOS';
const HOJA_LISTAS = 'LISTAS';

const FILAS_BASE = 5000;

// Columnas de BASE, en orden. sga = nombre de la columna en el reporte de SGA
// (null = se completa a mano o la pone el script).
const COLUMNAS_BASE = [
  { h: 'Acto', oculta: false, sga: null },
  { h: 'Concesionario', oculta: true, sga: 'Concesionario' },
  { h: 'Concepto de venta', oculta: true, sga: 'Concepto de venta' },
  { h: 'Modelo ahorro', oculta: false, sga: 'Modelo ahorro' },
  { h: 'Descripción de producto', oculta: true, sga: 'Descripción de producto' },
  { h: 'Mes cierre', oculta: false, sga: null },
  { h: 'Solicitud', oculta: false, sga: 'Solicitud' },
  { h: 'Grupo', oculta: false, sga: 'Grupo' },
  { h: 'Orden', oculta: false, sga: 'Orden' },
  { h: 'Apellido, Nombre', oculta: false, sga: 'Apellido, Nombre' },
  { h: 'Responsable', oculta: false, sga: null },
  { h: 'Documento', oculta: true, sga: 'Documento' },
  { h: 'CUIT/CUIL', oculta: true, sga: 'CUIT/CUIL' },
  { h: 'Teléfono', oculta: true, sga: 'Teléfono' },
  { h: 'Localidad', oculta: true, sga: 'Localidad' },
  { h: 'Provincia', oculta: true, sga: 'Provincia' },
  { h: 'CP', oculta: true, sga: 'CP' },
  { h: 'Avance', oculta: false, sga: 'Av.' },
  { h: 'Modalidad', oculta: false, sga: 'Mod. Ganador' },
  { h: 'Llave x llave', oculta: false, sga: null },
  { h: 'Pedido', oculta: false, sga: null },
  { h: 'Carpeta', oculta: false, sga: null },
  { h: 'Observacion', oculta: false, sga: null },
  { h: 'RESPONSABLE', oculta: false, sga: null },
  // Datos de SGA que hoy se borran: se guardan ocultos para el dashboard.
  { h: 'Vendedor SGA', oculta: true, sga: 'Vendedor' },
  { h: 'Monto licitado', oculta: true, sga: 'Monto licitado' },
  { h: 'Mail', oculta: true, sga: 'Mail' },
  { h: 'Dirección', oculta: true, sga: 'Dirección' },
];

const COLUMNAS_NUMERICAS = ['Solicitud', 'Grupo', 'Orden', 'Avance', 'Documento', 'CP', 'Monto licitado'];

// Valores de los desplegables (columna de LISTAS: [encabezado, valores]).
// RESPONSABLE (vendedor) es la única lista que acepta valores nuevos con aviso.
const LISTAS = {
  Responsable: ['CLARISA', 'SERGIO', 'TP', 'CHEXA'],
  Pedido: ['APROBADO', 'PENDIENTE', 'SUSPENDIDO', 'BAJA ADJ'],
  Carpeta: ['APROBADA', 'WEB', 'RECHAZADA'],
  Modalidad: ['Por licitación', 'Por sorteo'],
  'Llave x llave': ['SI'],
};

// Variantes de nombres vistas en los archivos mensuales.
const NORMALIZAR_RESPONSABLE = { CLARI: 'CLARISA', SERGY: 'SERGIO' };

// Archivos mensuales a importar una sola vez con "Importar meses anteriores".
const HISTORICOS = [
  { id: '16IQ8wKAHrRA2Oy2pntrB9ZeZxjTZkdjPj-7kZPYuREY', hoja: 'AGOSTO', acto: 'AGOSTO 26' },
  { id: '1_5SnWB3QvWHHkV2nNZ0kH3ZVhNBNJX0W4ipnlBrvaOI', hoja: 'SEPTIEMBRE', acto: 'SEPTIEMBRE 26' },
];

// Semillas de PRECIOS y OBJETIVOS (tomadas de los RESUMEN de cada mes).
const PRECIOS_INICIALES = {
  'AGOSTO 26': {
    DP1: 38370000, AR2: 30700000, FS1: 38300000, FT3: 45310000, FO1: 29310000,
    TN5: 47250000.01, NC1: 35600000, TV6: 51846000, LI1: 31600000, DT1: 48110000,
    PC5: 38000000.01, CA6: 39180000, FP3: 32833000, MB1: 24096000, TI1: 48964000,
    NT3: 42390000,
  },
  'SEPTIEMBRE 26': {
    NC1: 37550000, AR2: 30700000, FS1: 38300000, FT3: 45310000, FO1: 29310000,
    TN5: 47250000.01, DP1: 42635000, TV6: 51846000, LI1: 31600000, DT1: 48110000,
    PC5: 38000000.01, CA6: 39180000, FP3: 32833000, MB1: 24096000, TI1: 48964000,
    NT3: 42390000,
  },
};
const OBJETIVOS_INICIALES = [['AGOSTO 26', 57], ['SEPTIEMBRE 26', 60]];

const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const ANIO_DESDE = 26;
const ANIO_HASTA = 28;

const COLOR_ENCABEZADO = '#430000'; // el bordó del reporte de SGA

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pedidos')
    .addItem('Agregar acto pegado a BASE', 'agregarActoPegado')
    .addItem('Actualizar informe', 'actualizarInforme')
    .addSeparator()
    .addItem('Armar / reparar estructura', 'armarEstructura')
    .addItem('Importar meses anteriores (una sola vez)', 'importarHistorico')
    .addToUi();
}

/**
 * Al marcar Pedido = APROBADO y Carpeta = APROBADA, completa "Mes cierre" con
 * el mes actual si está vacío. Se puede corregir a mano.
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== HOJA_BASE || e.range.getRow() < 2) return;
  const cPedido = colBase('Pedido');
  const cCarpeta = colBase('Carpeta');
  const c0 = e.range.getColumn();
  const c1 = c0 + e.range.getNumColumns() - 1;
  if (c1 < Math.min(cPedido, cCarpeta) || c0 > Math.max(cPedido, cCarpeta)) return;

  const cMes = colBase('Mes cierre');
  const mesActual = etiquetaMes(new Date().getMonth(), new Date().getFullYear() % 100);
  const fila0 = e.range.getRow();
  const n = e.range.getNumRows();
  const datos = sh.getRange(fila0, 1, n, COLUMNAS_BASE.length).getValues();
  datos.forEach((fila, i) => {
    if (fila[cPedido - 1] === 'APROBADO' && fila[cCarpeta - 1] === 'APROBADA' && fila[cMes - 1] === '') {
      sh.getRange(fila0 + i, cMes).setValue(mesActual);
    }
  });
}

// ---------------------------------------------------------------------------
// Estructura
// ---------------------------------------------------------------------------

function armarEstructura() {
  const ss = SpreadsheetApp.getActive();
  const listas = armarListas(ss);
  armarBase(ss, listas);
  armarPegar(ss);
  armarPrecios(ss);
  armarObjetivos(ss);
  armarResumen(ss, listas);
  actualizarInforme();

  // La hoja vacía que trae el archivo nuevo.
  const vacia = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (vacia && vacia.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(vacia);

  [HOJA_BASE, HOJA_PEGAR, HOJA_RESUMEN, HOJA_INFORME, HOJA_PRECIOS, HOJA_OBJETIVOS, HOJA_LISTAS]
    .forEach((nombre, i) => {
      ss.setActiveSheet(ss.getSheetByName(nombre));
      ss.moveActiveSheet(i + 1);
    });

  ss.setActiveSheet(ss.getSheetByName(HOJA_BASE));
  SpreadsheetApp.getUi().alert('Estructura lista.');
}

function hoja(ss, nombre, posicion) {
  let sh = ss.getSheetByName(nombre);
  const nueva = !sh;
  if (nueva) sh = ss.insertSheet(nombre, Math.min(posicion, ss.getSheets().length));
  return { sh, nueva };
}

function armarListas(ss) {
  const { sh, nueva } = hoja(ss, HOJA_LISTAS, 6);
  const meses = [];
  for (let a = ANIO_DESDE; a <= ANIO_HASTA; a++) {
    for (let m = 0; m < 12; m++) meses.push(etiquetaMes(m, a));
  }
  // Meses se reescribe siempre (orden cronológico, lo usa el INFORME).
  sh.getRange(1, 1, 1, 1).setValue('Meses');
  sh.getRange(2, 1, meses.length, 1).setValues(meses.map((m) => [m]));

  const nombres = Object.keys(LISTAS);
  if (nueva) {
    nombres.forEach((nombre, i) => {
      const vals = LISTAS[nombre];
      sh.getRange(1, i + 2).setValue(nombre);
      sh.getRange(2, i + 2, vals.length, 1).setValues(vals.map((v) => [v]));
    });
    sh.getRange(1, nombres.length + 2).setValue('RESPONSABLE');
  }
  sh.getRange(1, 1, 1, nombres.length + 2)
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.setFrozenRows(1);

  // Rango de cada lista, para los desplegables.
  const rangos = { Meses: sh.getRange('A2:A') };
  const encabezados = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  encabezados.forEach((h, i) => {
    if (h && h !== 'Meses') rangos[h] = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1);
  });
  return rangos;
}

function armarBase(ss, listas) {
  const { sh } = hoja(ss, HOJA_BASE, 0);
  const n = COLUMNAS_BASE.length;
  if (sh.getMaxRows() < FILAS_BASE) sh.insertRowsAfter(sh.getMaxRows(), FILAS_BASE - sh.getMaxRows());
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());

  sh.getRange(1, 1, 1, n).setValues([COLUMNAS_BASE.map((c) => c.h)])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white')
    .setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);

  COLUMNAS_BASE.forEach((c, i) => {
    if (c.oculta) sh.hideColumns(i + 1); else sh.showColumns(i + 1);
  });

  // Desplegables.
  const filas = sh.getMaxRows() - 1;
  const desplegable = (columna, rango, estricto) => {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rango, true)
      .setAllowInvalid(!estricto)
      .build();
    sh.getRange(2, colBase(columna), filas, 1).setDataValidation(regla);
  };
  desplegable('Acto', listas.Meses, true);
  desplegable('Mes cierre', listas.Meses, true);
  desplegable('Responsable', listas.Responsable, true);
  desplegable('Pedido', listas.Pedido, true);
  desplegable('Carpeta', listas.Carpeta, true);
  desplegable('Modalidad', listas.Modalidad, true);
  desplegable('Llave x llave', listas['Llave x llave'], true);
  desplegable('RESPONSABLE', listas.RESPONSABLE, false);

  // Colores de Pedido.
  const rPedido = sh.getRange(2, colBase('Pedido'), filas, 1);
  const color = (texto, fondo) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(texto).setBackground(fondo).setRanges([rPedido]).build();
  const otras = sh.getConditionalFormatRules().filter((r) =>
    !r.getRanges().some((g) => g.getColumn() === colBase('Pedido')));
  sh.setConditionalFormatRules(otras.concat([
    color('APROBADO', '#d9ead3'),
    color('PENDIENTE', '#fff2cc'),
    color('SUSPENDIDO', '#fce5cd'),
    color('BAJA ADJ', '#f4cccc'),
  ]));

  sh.getRange(2, colBase('Observacion'), filas, 1).setWrap(false);
  sh.setColumnWidth(colBase('Observacion'), 320);
  sh.setColumnWidth(colBase('Apellido, Nombre'), 220);

  if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), n).createFilter();
}

function armarPegar(ss) {
  const { sh, nueva } = hoja(ss, HOJA_PEGAR, 1);
  if (nueva) {
    sh.getRange('A1').setNote(
      'Pegar acá el reporte de SGA completo, tal cual se baja (Ctrl+A / Ctrl+C en ' +
      'el Excel y Ctrl+V en A1). Después: menú Pedidos > Agregar acto pegado a BASE.');
  }
  sh.setTabColor('#999999');
}

function armarPrecios(ss) {
  const { sh, nueva } = hoja(ss, HOJA_PRECIOS, 4);
  if (nueva) {
    const filas = [];
    Object.keys(PRECIOS_INICIALES).forEach((mes) => {
      Object.keys(PRECIOS_INICIALES[mes]).forEach((modelo) => {
        filas.push([mes, modelo, PRECIOS_INICIALES[mes][modelo]]);
      });
    });
    sh.getRange(1, 1, 1, 3).setValues([['Mes', 'Modelo', 'V.M']]);
    sh.getRange(2, 1, filas.length, 3).setValues(filas);
  }
  sh.getRange('C:C').setNumberFormat('$ #,##0');
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.setFrozenRows(1);
}

function armarObjetivos(ss) {
  const { sh, nueva } = hoja(ss, HOJA_OBJETIVOS, 5);
  if (nueva) {
    sh.getRange(1, 1, 1, 3).setValues([['Mes', 'Objetivo', 'Circular / nota']]);
    sh.getRange(2, 1, OBJETIVOS_INICIALES.length, 2).setValues(OBJETIVOS_INICIALES);
  }
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// RESUMEN (fórmulas vivas sobre BASE, el mes se elige en B1)
// ---------------------------------------------------------------------------

function armarResumen(ss, listas) {
  const { sh } = hoja(ss, HOJA_RESUMEN, 2);
  const mesElegido = sh.getRange('B1').getValue();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.getRange('A1:B1').setValues([['Mes:', mesElegido || '']]);
  sh.getRange('B1').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(listas.Meses, true).build())
    .setFontWeight('bold').setBackground('#fff2cc');
  sh.getRange('A1').setFontWeight('bold');

  const r = rangosBase();
  const M = '$B$1';
  const aprob = `${r.Pedido},"APROBADO",${r.Carpeta},"APROBADA"`;
  const acto = `${r.Acto},${M}`;

  const bloques = [
    {
      col: 1, titulo: 'ADJUDICADOS DEL ACTO',
      filas: [
        ['Total adjudicados', `=COUNTIFS(${acto})`],
        ['Por licitación', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación")`],
        ['   Bajas (caídas)', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Pedido},"BAJA ADJ")`],
        ['   TP', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Responsable},"TP")`],
        ['   CHEXA', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Responsable},"CHEXA")`],
        ['   Netos licitación', '=B5-B6-B7-B8'],
        ['Por sorteo', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo")`],
        ['   Bajas (caídas)', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.Pedido},"BAJA ADJ")`],
        ['   ADM', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.RESPONSABLE},"ADM")`],
        ['   TP', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.Responsable},"TP")`],
      ],
    },
    {
      col: 4, titulo: 'ESTADO DEL ACTO',
      filas: [
        ['Aprobados (pedido + carpeta)', `=COUNTIFS(${acto},${aprob})`],
        ['Suspendidos', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO")`],
        ['   con CC aprobada', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO",${r.Carpeta},"APROBADA")`],
        ['   con CC ingresada (WEB)', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO",${r.Carpeta},"WEB")`],
        ['Pendientes', `=COUNTIFS(${acto},${r.Pedido},"PENDIENTE")`],
        ['   con CC aprobada', `=COUNTIFS(${acto},${r.Pedido},"PENDIENTE",${r.Carpeta},"APROBADA")`],
        ['Bajas', `=COUNTIFS(${acto},${r.Pedido},"BAJA ADJ")`],
        ['Sin estado', `=COUNTIFS(${acto},${r.Pedido},"")`],
        ['Trabajados', '=B4-E6-E7-E9-E10'],
      ],
    },
    {
      col: 7, titulo: 'CIERRES DEL MES',
      filas: [
        ['Netos del mes (todos los actos)', `=COUNTIFS(${r['Mes cierre']},${M},${aprob})`],
        ['   del acto del mes', `=COUNTIFS(${acto},${r['Mes cierre']},${M},${aprob})`],
        ['   de actos anteriores', '=H4-H5'],
        ['   TP', `=COUNTIFS(${r['Mes cierre']},${M},${aprob},${r.Responsable},"TP")`],
        ['Objetivo', `=IFERROR(VLOOKUP(${M},OBJETIVOS!$A:$B,2,FALSE),"")`],
        ['Proyección (+15%)', '=IF(H8="","",H8*1.15)'],
        ['Faltan para objetivo', '=IF(H8="","",MAX(0,H8-H4))'],
        ['% cumplimiento', '=IF(N(H8)=0,"",H4/H8)'],
        ['VM netos del mes', '=SUM(G16:G)'],
      ],
    },
  ];

  bloques.forEach((b) => {
    sh.getRange(3, b.col, 1, 2).merge().setValue(b.titulo)
      .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
    sh.getRange(4, b.col, b.filas.length, 1).setValues(b.filas.map((f) => [f[0]]));
    sh.getRange(4, b.col + 1, b.filas.length, 1).setFormulas(b.filas.map((f) => [f[1]]))
      .setFontWeight('bold').setHorizontalAlignment('center');
  });
  sh.getRange('H11').setNumberFormat('0%');
  sh.getRange('H12').setNumberFormat('$ #,##0');

  // Por modelo: lista dinámica con los modelos que tienen movimiento en el mes.
  const enc = ['Modelo', 'Adjudicados (acto)', 'Susp. c/CC aprobada', 'Aprobados (acto)',
    'Netos del mes', 'V.M', 'VM netos'];
  sh.getRange(15, 1, 1, enc.length).setValues([enc])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white').setWrap(true);
  const mod = r['Modelo ahorro'];
  const lista = 'A16:A60';
  const mapa = (expr) => `=MAP(${lista},LAMBDA(m,IF(m="","",${expr})))`;
  sh.getRange('A16').setFormula(
    `=IFERROR(SORT(UNIQUE(FILTER(${mod},${mod}<>"",(${r.Acto}=${M})+(${r['Mes cierre']}=${M})))),"")`);
  sh.getRange('B16').setFormula(mapa(`COUNTIFS(${acto},${mod},m)`));
  sh.getRange('C16').setFormula(mapa(`COUNTIFS(${acto},${mod},m,${r.Pedido},"SUSPENDIDO",${r.Carpeta},"APROBADA")`));
  sh.getRange('D16').setFormula(mapa(`COUNTIFS(${acto},${mod},m,${aprob})`));
  sh.getRange('E16').setFormula(mapa(`COUNTIFS(${r['Mes cierre']},${M},${mod},m,${aprob})`));
  sh.getRange('F16').setFormula(mapa(`SUMIFS(PRECIOS!$C:$C,PRECIOS!$A:$A,${M},PRECIOS!$B:$B,m)`));
  sh.getRange('G16').setFormula(`=MAP(E16:E60,F16:F60,LAMBDA(n,p,IF(n="","",n*p)))`);
  sh.getRange('F16:G60').setNumberFormat('$ #,##0');

  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(4, 210);
  sh.setColumnWidth(7, 230);
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// INFORME (mes a mes + por responsable)
// ---------------------------------------------------------------------------

function actualizarInforme() {
  const ss = SpreadsheetApp.getActive();
  const { sh } = hoja(ss, HOJA_INFORME, 3);
  sh.clear();

  const meses = mesesConDatos(ss);
  const r = rangosBase();
  const aprob = `${r.Pedido},"APROBADO",${r.Carpeta},"APROBADA"`;

  sh.getRange('A1').setValue('Evolución por mes').setFontWeight('bold').setFontSize(12);
  const enc = ['Mes', 'Adjudicados (acto)', 'Por licitación', 'Por sorteo', 'Bajas', '% bajas',
    'Aprobados del acto', '% conversión acto', 'Sin cerrar', 'Netos cerrados en el mes',
    '   del acto del mes', '   de actos anteriores', 'Objetivo', '% cumplimiento', 'VM netos del mes'];
  sh.getRange(3, 1, 1, enc.length).setValues([enc])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white').setWrap(true);

  const filas = meses.map((mes, i) => {
    const f = i + 4;
    const A = `$A${f}`;
    return [
      mes,
      `=COUNTIFS(${r.Acto},${A})`,
      `=COUNTIFS(${r.Acto},${A},${r.Modalidad},"Por licitación")`,
      `=COUNTIFS(${r.Acto},${A},${r.Modalidad},"Por sorteo")`,
      `=COUNTIFS(${r.Acto},${A},${r.Pedido},"BAJA ADJ")`,
      `=IF(B${f}=0,"",E${f}/B${f})`,
      `=COUNTIFS(${r.Acto},${A},${aprob})`,
      `=IF(B${f}=0,"",G${f}/B${f})`,
      `=B${f}-E${f}-G${f}`,
      `=COUNTIFS(${r['Mes cierre']},${A},${aprob})`,
      `=COUNTIFS(${r.Acto},${A},${r['Mes cierre']},${A},${aprob})`,
      `=J${f}-K${f}`,
      `=IFERROR(VLOOKUP(${A},OBJETIVOS!$A:$B,2,FALSE),"")`,
      `=IF(N(M${f})=0,"",J${f}/M${f})`,
      `=ARRAYFORMULA(SUMPRODUCT((${r['Mes cierre']}=${A})*(${r.Pedido}="APROBADO")*(${r.Carpeta}="APROBADA")*` +
        `SUMIFS(PRECIOS!$C:$C,PRECIOS!$A:$A,${A},PRECIOS!$B:$B,${r['Modelo ahorro']})))`,
    ];
  });
  if (filas.length) {
    sh.getRange(4, 1, filas.length, enc.length).setFormulas(filas.map((f) => ['', ...f.slice(1)]));
    sh.getRange(4, 1, filas.length, 1).setValues(filas.map((f) => [f[0]])).setFontWeight('bold');
    ['F', 'H', 'N'].forEach((c) => sh.getRange(`${c}4:${c}${3 + filas.length}`).setNumberFormat('0%'));
    sh.getRange(`O4:O${3 + filas.length}`).setNumberFormat('$ #,##0');
  }

  // Por responsable, todos los actos.
  const f0 = 6 + filas.length;
  sh.getRange(f0, 1).setValue('Por responsable (todos los actos)').setFontWeight('bold').setFontSize(12);
  const enc2 = ['Responsable', 'Asignados', 'Aprobados', 'Bajas', 'Sin cerrar', '% conversión'];
  sh.getRange(f0 + 2, 1, 1, enc2.length).setValues([enc2])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  const resp = valoresLista(ss, 'Responsable');
  const filas2 = resp.map((nombre, i) => {
    const f = f0 + 3 + i;
    const A = `$A${f}`;
    return [
      nombre,
      `=COUNTIFS(${r.Responsable},${A})`,
      `=COUNTIFS(${r.Responsable},${A},${aprob})`,
      `=COUNTIFS(${r.Responsable},${A},${r.Pedido},"BAJA ADJ")`,
      `=B${f}-C${f}-D${f}`,
      `=IF(B${f}=0,"",C${f}/B${f})`,
    ];
  });
  if (filas2.length) {
    sh.getRange(f0 + 3, 1, filas2.length, 1).setValues(filas2.map((f) => [f[0]])).setFontWeight('bold');
    sh.getRange(f0 + 3, 2, filas2.length, 5).setFormulas(filas2.map((f) => f.slice(1)));
    sh.getRange(f0 + 3, 6, filas2.length, 1).setNumberFormat('0%');
  }

  sh.setColumnWidth(1, 150);
  sh.setFrozenColumns(1);

  // Si el RESUMEN no tiene mes elegido, poner el último acto.
  const resumen = ss.getSheetByName(HOJA_RESUMEN);
  if (resumen && !resumen.getRange('B1').getValue() && meses.length) {
    const actos = valoresColumnaBase(ss, 'Acto');
    const ultimo = meses.filter((m) => actos.indexOf(m) >= 0).pop();
    if (ultimo) resumen.getRange('B1').setValue(ultimo);
  }
}

/** Meses desde el primer acto hasta el último acto o mes de cierre cargado. */
function mesesConDatos(ss) {
  const todos = [];
  for (let a = ANIO_DESDE; a <= ANIO_HASTA; a++) {
    for (let m = 0; m < 12; m++) todos.push(etiquetaMes(m, a));
  }
  const usados = valoresColumnaBase(ss, 'Acto').concat(valoresColumnaBase(ss, 'Mes cierre'))
    .map((v) => todos.indexOf(v)).filter((i) => i >= 0);
  if (!usados.length) return [];
  return todos.slice(Math.min(...usados), Math.max(...usados) + 1);
}

// ---------------------------------------------------------------------------
// Carga de datos
// ---------------------------------------------------------------------------

/** Toma el reporte pegado en PEGAR SGA y lo agrega a BASE. */
function agregarActoPegado() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const pegar = ss.getSheetByName(HOJA_PEGAR);
  if (!pegar || pegar.getLastRow() === 0) {
    ui.alert(`La hoja "${HOJA_PEGAR}" está vacía. Pegá el reporte de SGA en A1.`);
    return;
  }
  const datos = pegar.getDataRange().getValues();

  const iEnc = datos.findIndex((fila) =>
    fila.indexOf('Solicitud') >= 0 && fila.indexOf('Mod. Ganador') >= 0);
  if (iEnc < 0) {
    ui.alert('No encontré la fila de títulos del reporte (con "Solicitud" y "Mod. Ganador").');
    return;
  }
  const enc = datos[iEnc].map((v) => String(v).trim());
  const iSolicitud = enc.indexOf('Solicitud');
  const filasSga = datos.slice(iEnc + 1).filter((f) => /^\d+$/.test(String(f[iSolicitud]).trim()));

  // Acto: la fecha que aparece en la fila de filtros ("Adjudicados; 15/09/2026").
  let actoDetectado = '';
  datos.slice(0, iEnc).some((fila) => fila.some((v) => {
    const m = String(v).match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (m && /adjudicad/i.test(String(v))) {
      actoDetectado = etiquetaMes(Number(m[2]) - 1, Number(m[3]) % 100);
      return true;
    }
    return false;
  }));

  const resp = ui.prompt('Agregar acto a BASE',
    `${filasSga.length} adjudicados en el reporte.\n` +
    (actoDetectado ? `Acto detectado: ${actoDetectado}. Aceptar para confirmar, ` : 'No pude detectar el acto: ') +
    'o escribí el acto (ej. OCTUBRE 26).', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const acto = (resp.getResponseText().trim().toUpperCase() || actoDetectado);
  if (!esMesValido(acto)) {
    ui.alert(`"${acto}" no es un mes válido. Usá el formato MES AA, ej. OCTUBRE 26.`);
    return;
  }

  const filas = filasSga.map((f) => COLUMNAS_BASE.map((c) => {
    if (c.h === 'Acto') return acto;
    if (!c.sga) return '';
    const i = enc.indexOf(c.sga);
    return i >= 0 ? limpiar(c.h, f[i]) : '';
  }));
  const { agregadas, repetidas } = agregarABase(ss, filas);
  if (agregadas) pegar.clear();
  actualizarInforme();
  ui.alert(`Acto ${acto}: ${agregadas} filas agregadas a BASE` +
    (repetidas ? `, ${repetidas} ya estaban cargadas (no se duplicaron).` : '.'));
}

/** Copia una sola vez los archivos mensuales viejos (ver HISTORICOS). */
function importarHistorico() {
  const ss = SpreadsheetApp.getActive();
  const resumen = [];
  HISTORICOS.forEach((h) => {
    const origen = SpreadsheetApp.openById(h.id).getSheetByName(h.hoja);
    const datos = origen.getDataRange().getValues();
    const enc = datos[0].map((v) => String(v).trim());
    const iSolicitud = enc.indexOf('Solicitud');
    const filas = datos.slice(1)
      .filter((f) => String(f[iSolicitud]).trim() !== '')
      .map((f) => COLUMNAS_BASE.map((c) => {
        if (c.h === 'Acto') return h.acto;
        const i = enc.indexOf(c.h === 'Mes cierre' ? 'Mes' : c.h);
        if (i < 0) return '';
        if (c.h === 'Mes cierre') return mesCierreDesdeTexto(f[i], h.acto);
        return limpiar(c.h, f[i]);
      }));
    const { agregadas, repetidas } = agregarABase(ss, filas);
    resumen.push(`${h.acto}: ${agregadas} agregadas` + (repetidas ? `, ${repetidas} ya estaban` : ''));
  });
  actualizarInforme();
  SpreadsheetApp.getUi().alert(resumen.join('\n'));
}

/** Agrega filas a BASE salteando las que ya existen (mismo Acto + Solicitud). */
function agregarABase(ss, filas) {
  const base = ss.getSheetByName(HOJA_BASE);
  const cActo = colBase('Acto') - 1;
  const cSol = colBase('Solicitud') - 1;
  const clave = (f) => `${f[cActo]}|${f[cSol]}`;
  const existentes = new Set();
  const ultima = ultimaFilaBase(base);
  if (ultima > 1) {
    base.getRange(2, 1, ultima - 1, COLUMNAS_BASE.length).getValues()
      .forEach((f) => existentes.add(clave(f)));
  }
  const nuevas = filas.filter((f) => {
    const k = clave(f);
    if (existentes.has(k)) return false;
    existentes.add(k);
    return true;
  });
  if (nuevas.length) {
    base.getRange(ultima + 1, 1, nuevas.length, COLUMNAS_BASE.length).setValues(nuevas);
  }
  agregarVendedoresALista(ss, nuevas);
  return { agregadas: nuevas.length, repetidas: filas.length - nuevas.length };
}

/** Suma a LISTAS los vendedores (RESPONSABLE) que todavía no están. */
function agregarVendedoresALista(ss, filas) {
  const listas = ss.getSheetByName(HOJA_LISTAS);
  const enc = listas.getRange(1, 1, 1, listas.getLastColumn()).getValues()[0];
  const c = enc.indexOf('RESPONSABLE') + 1;
  if (!c) return;
  const actuales = valoresLista(ss, 'RESPONSABLE');
  const iResp = colBase('RESPONSABLE') - 1;
  const nuevos = [...new Set(filas.map((f) => f[iResp]).filter((v) => v && actuales.indexOf(v) < 0))].sort();
  if (nuevos.length) {
    const todos = actuales.concat(nuevos).sort();
    listas.getRange(2, c, listas.getMaxRows() - 1, 1).clearContent();
    listas.getRange(2, c, todos.length, 1).setValues(todos.map((v) => [v]));
  }
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

function colBase(nombre) {
  const i = COLUMNAS_BASE.findIndex((c) => c.h === nombre);
  if (i < 0) throw new Error(`Columna inexistente en BASE: ${nombre}`);
  return i + 1;
}

function letra(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Rangos abiertos de BASE para usar en fórmulas: r['Pedido'] = BASE!$U$2:$U */
function rangosBase() {
  const r = {};
  COLUMNAS_BASE.forEach((c, i) => {
    const l = letra(i + 1);
    r[c.h] = `${HOJA_BASE}!$${l}$2:$${l}`;
  });
  return r;
}

function etiquetaMes(indice, anio2) {
  return `${MESES[indice]} ${String(anio2).padStart(2, '0')}`;
}

function esMesValido(texto) {
  const m = String(texto).match(/^([A-Z]+) (\d{2})$/);
  return !!m && MESES.indexOf(m[1]) >= 0 && Number(m[2]) >= ANIO_DESDE && Number(m[2]) <= ANIO_HASTA;
}

/** "SEPTIEMBRE" en el archivo de AGOSTO 26 -> "SEPTIEMBRE 26" (o año siguiente si el mes es anterior al acto). */
function mesCierreDesdeTexto(valor, acto) {
  const texto = String(valor).trim().toUpperCase();
  if (!texto) return '';
  if (esMesValido(texto)) return texto;
  const iMes = MESES.indexOf(texto);
  if (iMes < 0) return '';
  const [mesActo, anioActo] = acto.split(' ');
  const anio = Number(anioActo) + (iMes < MESES.indexOf(mesActo) ? 1 : 0);
  return etiquetaMes(iMes, anio);
}

function limpiar(columna, valor) {
  if (valor === null || valor === undefined) return '';
  let v = typeof valor === 'string' ? valor.trim() : valor;
  if (v === '') return '';
  if (COLUMNAS_NUMERICAS.indexOf(columna) >= 0 && typeof v === 'string') {
    const n = Number(v.replace(/,/g, ''));
    if (!isNaN(n)) return n;
  }
  if (columna === 'Responsable') {
    v = String(v).toUpperCase();
    return NORMALIZAR_RESPONSABLE[v] || v;
  }
  if (columna === 'Pedido' || columna === 'Carpeta') return String(v).toUpperCase();
  if (columna === 'Llave x llave') return 'SI'; // en los archivos viejos aparece "x", "1" o "SI"
  return v;
}

/** Última fila con Solicitud cargada (las validaciones no cuentan como datos, pero por las dudas). */
function ultimaFilaBase(base) {
  const vals = base.getRange(1, colBase('Solicitud'), base.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') return i + 1;
  }
  return 1;
}

function valoresColumnaBase(ss, nombre) {
  const base = ss.getSheetByName(HOJA_BASE);
  const ultima = ultimaFilaBase(base);
  if (ultima < 2) return [];
  return base.getRange(2, colBase(nombre), ultima - 1, 1).getValues()
    .map((f) => f[0]).filter((v) => v !== '');
}

function valoresLista(ss, encabezado) {
  const sh = ss.getSheetByName(HOJA_LISTAS);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const c = enc.indexOf(encabezado) + 1;
  if (!c) return [];
  return sh.getRange(2, c, sh.getMaxRows() - 1, 1).getValues()
    .map((f) => f[0]).filter((v) => v !== '');
}

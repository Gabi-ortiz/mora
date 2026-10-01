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
 * - OBJETIVOS / INCENTIVOS: se leen (IMPORTRANGE) del archivo "Objetivos y
 *   señales comerciales - PDA"; no se cargan acá.
 *
 * Los meses ("SEPTIEMBRE 26") se guardan como TEXTO: si no, Sheets los toma
 * como la fecha 26/09 y las fórmulas no los encuentran.
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
const HOJA_INCENTIVOS = 'INCENTIVOS';
const HOJA_AJUSTES = 'AJUSTES';

// Archivo "Objetivos y señales comerciales - PDA" (objetivos, categoría e incentivos).
const ID_ARCHIVO_OBJETIVOS = '1TDMqgJkbkP11dSlUhvJpD5NhKtkj0vvuiu-pPPQ5kYo';

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
  { h: 'RESPONSABLE', oculta: false, sga: 'Of. Cuenta' }, // oficial de cuenta, viene de SGA
  // Datos de SGA que hoy se borran: se guardan ocultos para el dashboard.
  { h: 'Vendedor SGA', oculta: true, sga: 'Vendedor' },
  { h: 'Monto licitado', oculta: true, sga: 'Monto licitado' },
  { h: 'Mail', oculta: true, sga: 'Mail' },
  { h: 'Dirección', oculta: true, sga: 'Dirección' },
];

const COLUMNAS_NUMERICAS = ['Solicitud', 'Grupo', 'Orden', 'Avance', 'Documento', 'CP', 'Monto licitado'];

// Valores de los desplegables (columna de LISTAS: [encabezado, valores]).
// RESPONSABLE (oficial de cuenta) no tiene desplegable: viene de SGA ("Of. Cuenta").
const LISTAS = {
  Responsable: ['CLARISA', 'SERGIO', 'TP', 'CHEXA'],
  Pedido: ['APROBADO', 'PENDIENTE', 'SUSPENDIDO', 'BAJA ADJ'],
  Carpeta: ['APROBADA', 'WEB', 'RECHAZADA', 'REINGRESADA'],
  Modalidad: ['Por licitación', 'Por sorteo'],
  'Llave x llave': ['SI'],
};

// Variantes de nombres vistas en los archivos mensuales.
const NORMALIZAR_RESPONSABLE = { CLARI: 'CLARISA', SERGY: 'SERGIO' };

// Archivos mensuales que se importaron una sola vez con importarHistorico() (ya hecho).
const HISTORICOS = [
  { id: '16IQ8wKAHrRA2Oy2pntrB9ZeZxjTZkdjPj-7kZPYuREY', hoja: 'AGOSTO', acto: 'AGOSTO 26' },
  { id: '1_5SnWB3QvWHHkV2nNZ0kH3ZVhNBNJX0W4ipnlBrvaOI', hoja: 'SEPTIEMBRE', acto: 'SEPTIEMBRE 26' },
];

// Semilla de PRECIOS (tomada de los RESUMEN de cada mes).
// Cierres que no están en BASE, para que los meses de arranque den igual que sus
// RESUMEN mensuales originales: [Mes, Modelo, Netos, VM netos, Nota].
const AJUSTES_INICIALES = [
  ...[['DP1', 7, 229740000.1], ['FO1', 1, 27459000], ['DT1', 2, 96220000], ['MB1', 4, 96384000],
    ['NT3', 1, 42390000]].map(([m, n, vm]) => ['AGOSTO 26', m, n, vm,
    'Cierres en agosto de actos anteriores (archivo de JULIO, no está en BASE). RESUMEN de AGOSTO 26, bloque NETOS ANTERIOR.']),
  ['SEPTIEMBRE 26', 'NC1', 1, 37550000,
    '"+1" cargado a mano en el RESUMEN de SEPTIEMBRE 26 (NETOS ANTERIOR). Confirmar de qué solicitud es.'],
  ['SEPTIEMBRE 26', 'MB1', 1, 24096000,
    '"+1" cargado a mano en el RESUMEN de SEPTIEMBRE 26 (NETOS ANTERIOR). Confirmar de qué solicitud es.'],
];

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
    .addToUi();
  // importarHistorico() ya se usó (agosto y septiembre en BASE): quedó fuera del menú.
  // Si hiciera falta de nuevo, se corre desde el editor de Apps Script (no duplica).
}

/**
 * BASE: al marcar Pedido = APROBADO y Carpeta = APROBADA completa "Mes cierre"
 * con el mes actual (si está vacío). Si "Mes cierre" se escribe a mano, lo
 * pasa al formato MES AA y avisa si no es un mes válido.
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== HOJA_BASE || e.range.getRow() < 2) return;
  const cPedido = colBase('Pedido');
  const cCarpeta = colBase('Carpeta');
  const cMes = colBase('Mes cierre');
  const c0 = e.range.getColumn();
  const c1 = c0 + e.range.getNumColumns() - 1;
  const fila0 = e.range.getRow();
  const n = e.range.getNumRows();

  // Mes cierre escrito a mano: "septiembre 26" -> "SEPTIEMBRE 26"; si no es un mes válido, avisa.
  if (c0 <= cMes && cMes <= c1) {
    const rango = sh.getRange(fila0, cMes, n, 1);
    const vals = rango.getValues().map((f) => [textoMes(f[0])]);
    rango.setValues(vals);
    if (vals.some((f) => f[0] !== '' && !esMesValido(f[0]))) {
      SpreadsheetApp.getActive().toast('Mes cierre: usar el formato MES AA, por ejemplo SEPTIEMBRE 26.',
        'Formato de mes', 8);
    }
  }

  // Al marcar Pedido = APROBADO y Carpeta = APROBADA, completa Mes cierre con el mes actual si está vacío.
  if (c1 < Math.min(cPedido, cCarpeta) || c0 > Math.max(cPedido, cCarpeta)) return;
  const mesActual = etiquetaMes(new Date().getMonth(), new Date().getFullYear() % 100);
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
  armarIncentivos(ss);
  armarAjustes(ss);
  armarResumen(ss, listas);
  actualizarInforme();

  // La hoja vacía que trae el archivo nuevo.
  const vacia = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (vacia && vacia.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(vacia);

  [HOJA_BASE, HOJA_PEGAR, HOJA_RESUMEN, HOJA_INFORME, HOJA_PRECIOS, HOJA_AJUSTES, HOJA_OBJETIVOS,
    HOJA_INCENTIVOS, HOJA_LISTAS]
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
  const nombres = Object.keys(LISTAS);
  if (nueva) {
    nombres.forEach((nombre, i) => {
      const vals = LISTAS[nombre];
      sh.getRange(1, i + 2).setValue(nombre);
      sh.getRange(2, i + 2, vals.length, 1).setValues(vals.map((v) => [v]));
    });
  }
  // Valores nuevos de las listas fijas (p. ej. REINGRESADA) en una hoja ya creada.
  nombres.forEach((nombre, i) => {
    const actuales = sh.getRange(2, i + 2, sh.getMaxRows() - 1, 1).getValues().map((f) => f[0]).filter((v) => v !== '');
    const faltan = LISTAS[nombre].filter((v) => actuales.indexOf(v) < 0);
    if (faltan.length) sh.getRange(actuales.length + 2, i + 2, faltan.length, 1).setValues(faltan.map((v) => [v]));
  });
  // Columnas de versiones anteriores que ya no se usan: vendedores (RESPONSABLE) y las
  // listas de meses calculadas que estaban después de las listas fijas.
  sh.getRange(1, nombres.length + 2, sh.getMaxRows(), Math.max(1, sh.getMaxColumns() - nombres.length - 1)).clear();

  // A: "Meses con datos" = UNIQUE de los meses de BASE (Acto y Mes cierre), el más reciente
  // primero. Es la lista del selector de mes del RESUMEN. Calculada: no se edita.
  const nombresMes = MESES.map((m) => `"${m}"`).join(',');
  sh.getRange('A:A').clear().setNumberFormat('@');
  sh.getRange('A1').setValue('Meses con datos');
  sh.getRange('A2').setFormula(
    `=IFERROR(LET(v,{${HOJA_BASE}!A2:A;${HOJA_BASE}!F2:F},` +
    'u,UNIQUE(FILTER(v,REGEXMATCH(v&"","^[A-Z]+ [0-9]{2}$"))),' +
    `SORT(u,ARRAYFORMULA(DATE(2000+VALUE(RIGHT(u,2)),MATCH(LEFT(u,LEN(u)-3),{${nombresMes}},0),1)),FALSE)),"")`);
  sh.getRange(1, 1, 1, nombres.length + 1)
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.setFrozenRows(1);

  // Rango de cada lista, para los desplegables.
  const rangos = {};
  const encabezados = sh.getRange(1, 1, 1, nombres.length + 1).getValues()[0];
  encabezados.forEach((h, i) => {
    if (h) rangos[h] = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1);
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
  mesesComoTexto(sh.getRange(2, colBase('Acto'), sh.getMaxRows() - 1, 1));
  mesesComoTexto(sh.getRange(2, colBase('Mes cierre'), sh.getMaxRows() - 1, 1));

  // Desplegables.
  const filas = sh.getMaxRows() - 1;
  const desplegable = (columna, rango, estricto) => {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rango, true)
      .setAllowInvalid(!estricto)
      .build();
    sh.getRange(2, colBase(columna), filas, 1).setDataValidation(regla);
  };
  // Acto y Mes cierre los completa el script (al importar / al aprobar): sin desplegable y en
  // gris. Acto avisa si se edita; Mes cierre se puede corregir a mano (onEdit lo normaliza).
  const rActo = sh.getRange(2, colBase('Acto'), filas, 1);
  const rCierre = sh.getRange(2, colBase('Mes cierre'), filas, 1);
  [rActo, rCierre].forEach((rg) => rg.clearDataValidations().setBackground('#f3f3f3').setFontColor(GRIS_TEXTO));
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter((pr) => pr.getDescription() === 'Acto (viene de SGA)').forEach((pr) => pr.remove());
  rActo.protect().setDescription('Acto (viene de SGA)').setWarningOnly(true);
  desplegable('Responsable', listas.Responsable, true);
  desplegable('Pedido', listas.Pedido, true);
  desplegable('Carpeta', listas.Carpeta, true);
  desplegable('Modalidad', listas.Modalidad, true);
  desplegable('Llave x llave', listas['Llave x llave'], true);
  sh.getRange(2, colBase('RESPONSABLE'), filas, 1).clearDataValidations();

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

  // Cuadrícula marcada en toda la tabla (también en las filas vacías, para las que se agreguen).
  sh.getRange(1, 1, sh.getMaxRows(), n)
    .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

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
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['Mes', 'Modelo', 'V.M']]);
    sh.getRange(2, 1, filas.length, 3).setValues(filas);
  }
  mesesComoTexto(sh.getRange(2, 1, sh.getMaxRows() - 1, 1));
  sh.getRange('C:C').setNumberFormat('$ #,##0');
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.setFrozenRows(1);
}

/**
 * OBJETIVOS: Mes | Objetivo | Categoría de PEDIDOS TOTALES (Fiat), leídos del
 * archivo de objetivos. La primera vez hay que hacer clic en A2 > "Permitir acceso".
 */
function armarObjetivos(ss) {
  const { sh } = hoja(ss, HOJA_OBJETIVOS, 5);
  sh.clear();
  sh.getRange(1, 1, 1, 3).setValues([['Mes', 'Objetivo', 'Categoría']])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.getRange('A2').setFormula(
    `=IFERROR(QUERY(IMPORTRANGE("${ID_ARCHIVO_OBJETIVOS}","OBJETIVOS!A2:F"),` +
    `"select Col1, Col5, Col6 where Col4 = 'PEDIDOS TOTALES' and Col3 = 'FIAT'",0),` +
    `"Hacer clic acá y elegir Permitir acceso al archivo de objetivos")`);
  sh.getRange('E1').setValue('Se cargan en el archivo "Objetivos y señales comerciales - PDA" (no editar acá).')
    .setFontStyle('italic');
  sh.setFrozenRows(1);
}

/** INCENTIVOS: Mes | Concepto | A | B | C (Fiat), leídos del archivo de objetivos. */
function armarIncentivos(ss) {
  const { sh } = hoja(ss, HOJA_INCENTIVOS, 6);
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['Mes', 'Concepto', 'A', 'B', 'C']])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.getRange('A2').setFormula(
    `=IFERROR(QUERY(IMPORTRANGE("${ID_ARCHIVO_OBJETIVOS}","INCENTIVOS!A2:G"),` +
    `"select Col1, Col3, Col5, Col6, Col7 where Col2 = 'FIAT'",0),"")`);
  sh.getRange('C:E').setNumberFormat('0.00%');
  sh.getRange('G1').setValue('Se cargan en el archivo "Objetivos y señales comerciales - PDA" (no editar acá).')
    .setFontStyle('italic');
  sh.setFrozenRows(1);
}

/**
 * AJUSTES: cierres que no están en BASE (p. ej. de actos anteriores al primer mes
 * cargado). Suman a los netos y al VM del mes en RESUMEN e INFORME.
 */
function armarAjustes(ss) {
  const { sh, nueva } = hoja(ss, HOJA_AJUSTES, 5);
  if (nueva) {
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(2, 1, AJUSTES_INICIALES.length, 5).setValues(AJUSTES_INICIALES);
  }
  mesesComoTexto(sh.getRange(2, 1, sh.getMaxRows() - 1, 1));
  sh.getRange(1, 1, 1, 5).setValues([['Mes', 'Modelo', 'Netos', 'VM netos', 'Nota']])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.getRange('D:D').setNumberFormat('$ #,##0');
  sh.setColumnWidth(5, 520);
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// RESUMEN (fórmulas vivas sobre BASE, el mes se elige en H1)
// ---------------------------------------------------------------------------

// Celdas del RESUMEN que usan otras fórmulas / el script.
const CELDA_MES_RESUMEN = 'H1';
// Celda oculta con el mes elegido normalizado a "MES AA": Sheets a veces guarda la
// opción del desplegable como fecha (26/09/2026) y las fórmulas no la encontraban.
// Todas las fórmulas del RESUMEN usan esta celda, no H1.
const CELDA_MES_CALC = 'K1';
const CELDA_CATEGORIA_RESUMEN = 'I21';

// Paleta: el bordó del reporte de SGA + grises suaves.
const GRIS_TEXTO = '#666666';
const GRIS_LINEA = '#e0e0e0';
const FONDO_TARJETA = '#f8f4f4';
const FONDO_TOTAL = '#efe7e7';

function armarResumen(ss, listas) {
  const { sh } = hoja(ss, HOJA_RESUMEN, 2);
  // El mes elegido se conserva; B1 es donde estaba el selector en la versión anterior.
  const mesElegido = [CELDA_MES_RESUMEN, 'B1'].map((c) => textoMes(sh.getRange(c).getValue()))
    .find(esMesValido) || '';
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.setFrozenRows(0);
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  sh.setHiddenGridlines(true);

  const r = rangosBase();
  const M = `$${CELDA_MES_CALC.replace(/(\d+)/, '$$$1')}`; // $K$1
  if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  sh.getRange(CELDA_MES_CALC).setFormula(formulaMesNormalizado(`$${CELDA_MES_RESUMEN.replace(/(\d+)/, '$$$1')}`));
  sh.hideColumns(sh.getRange(CELDA_MES_CALC).getColumn());
  const aprob = `${r.Pedido},"APROBADO",${r.Carpeta},"APROBADA"`;
  const acto = `${r.Acto},${M}`;

  // Grilla: A = margen; B..I = 8 columnas iguales.
  sh.setColumnWidth(1, 16);
  for (let c = 2; c <= 9; c++) sh.setColumnWidth(c, 118);
  sh.setColumnWidth(10, 16);

  // Título y selector de mes.
  sh.getRange('B1:E1').merge().setValue('Resumen de pedidos')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR_ENCABEZADO);
  sh.getRange('G1').setValue('Mes').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('H1:I1').merge().setNumberFormat('@').setValue(mesElegido || '')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(listas['Meses con datos'], true).setAllowInvalid(true).build())
    .setFontWeight('bold').setFontSize(12).setBackground('#fff2cc').setHorizontalAlignment('center');
  sh.getRange('B2:I2').merge()
    .setValue('Acto = adjudicados en el acto del mes elegido.   Cierres = aprobados (pedido + carpeta) en el mes, de cualquier acto.')
    .setFontSize(9).setFontColor(GRIS_TEXTO);
  sh.setRowHeight(1, 34);
  sh.setRowHeight(3, 8);

  // --- Bloques (filas 9-18 y 21-29). Valores en E (izquierda) e I (derecha).
  const izq1 = [
    ['Total adjudicados', `=COUNTIFS(${acto})`],
    ['Por licitación', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación")`],
    ['   Bajas (caídas)', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Pedido},"BAJA ADJ")`, true],
    ['   TP', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Responsable},"TP")`, true],
    ['   CHEXA', `=COUNTIFS(${acto},${r.Modalidad},"Por licitación",${r.Responsable},"CHEXA")`, true],
    ['   Netos licitación', '=E10-E11-E12-E13', true],
    ['Por sorteo', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo")`],
    ['   Bajas (caídas)', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.Pedido},"BAJA ADJ")`, true],
    ['   ADM', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.RESPONSABLE},"ADM")`, true],
    ['   TP', `=COUNTIFS(${acto},${r.Modalidad},"Por sorteo",${r.Responsable},"TP")`, true],
  ];
  const der1 = [
    ['Aprobados (pedido + carpeta)', `=COUNTIFS(${acto},${aprob})`],
    ['Suspendidos', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO")`],
    ['   con CC aprobada', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO",${r.Carpeta},"APROBADA")`, true],
    ['   con CC ingresada (WEB)', `=COUNTIFS(${acto},${r.Pedido},"SUSPENDIDO",${r.Carpeta},"WEB")`, true],
    ['Pendientes', `=COUNTIFS(${acto},${r.Pedido},"PENDIENTE")`],
    ['   con CC aprobada', `=COUNTIFS(${acto},${r.Pedido},"PENDIENTE",${r.Carpeta},"APROBADA")`, true],
    ['Bajas', `=COUNTIFS(${acto},${r.Pedido},"BAJA ADJ")`],
    // Como el RESUMEN mensual original: (netos licitación + sorteo) - (susp/pend con CC + bajas).
    ['Trabajados', '=E14+E15-I11-I12-I14-I15'],
    ['Terminados', '=I9+I11+I12+I14'],
    ['Sin terminar', '=E9-I16-I15'],
  ];
  const izq2 = [
    ['Netos del mes (todos los actos)',
      `=COUNTIFS(${r['Mes cierre']},${M},${aprob})+SUMIFS(${HOJA_AJUSTES}!$C:$C,${HOJA_AJUSTES}!$A:$A,${M})`],
    ['   del acto del mes', `=COUNTIFS(${acto},${r['Mes cierre']},${M},${aprob})`, true],
    ['   de actos anteriores', '=E21-E22', true],
    ['   TP', `=COUNTIFS(${r['Mes cierre']},${M},${aprob},${r.Responsable},"TP")`, true],
    ['Objetivo', `=IFERROR(VLOOKUP(${M},OBJETIVOS!$A:$B,2,FALSE),"")`],
    ['Proyección (+15%)', '=IF(E25="","",ROUND(E25*1.15))'],
    ['Faltan para el objetivo', '=IF(E25="","",MAX(0,E25-E21))'],
    ['% cumplimiento', '=IF(N(E25)=0,"",E21/E25)'],
    ['VM netos del mes', '=I33'],
  ];
  const der2 = [
    ['Categoría del mes', '=IFERROR(VLOOKUP(' + M + ',OBJETIVOS!$A:$C,3,FALSE),"")'],
    ['% cumplimiento pedidos', '=E28'],
    ['Bonus cumplimiento pedidos totales', `=${bonusSegunCategoria('CUMPLIMIENTO PEDIDOS', 'N(E28)<1')}`],
    ['Pedidos adicionales al objetivo', '=IF(N(E25)=0,"",MAX(0,E21-E25))'],
    ['Bonus por pedido adicional', `=${bonusSegunCategoria('ADICIONALES', 'N(I24)=0')}`],
  ];
  bloqueResumen(sh, 8, 2, 'ADJUDICADOS DEL ACTO', izq1);
  bloqueResumen(sh, 8, 6, 'ESTADO DEL ACTO', der1);
  bloqueResumen(sh, 20, 2, 'CIERRES DEL MES', izq2);
  bloqueResumen(sh, 20, 6, 'INCENTIVO PEDIDOS', der2);
  sh.getRange('F27:I29').merge().setWrap(true).setVerticalAlignment('top')
    .setValue('El bonus de cumplimiento exige además el 100% del objetivo de suscripciones (no se mide en ' +
      'este archivo). Los % salen de la hoja INCENTIVOS según la categoría.')
    .setFontSize(9).setFontStyle('italic').setFontColor(GRIS_TEXTO);
  ['E28', 'I22'].forEach((c) => sh.getRange(c).setNumberFormat('0%'));
  ['I23', 'I25'].forEach((c) => sh.getRange(c).setNumberFormat('0.00%'));
  sh.getRange('E29').setNumberFormat('$ #,##0');
  sh.setRowHeight(7, 10);
  sh.setRowHeight(19, 14);
  sh.setRowHeight(30, 14);

  // --- Tarjetas (filas 4-6): lo más importante, a la vista.
  const tarjetas = [
    ['B', 'C', 'Adjudicados del acto', '=E9', '="Licitación "&E10&"   ·   Sorteo "&E15', '0'],
    ['D', 'E', 'Aprobados del acto', '=I9', '=IF(N(E9)=0,"","Conversión "&TEXT(I9/E9,"0%"))', '0'],
    ['F', 'G', 'Netos del mes', '=E21', '="Del acto "&E22&"   ·   Anteriores "&E23', '0'],
    ['H', 'I', 'Cumplimiento objetivo', '=E28',
      '=IF(E25="","Sin objetivo cargado","Objetivo "&E25&"   ·   Categoría "&I21)', '0%'],
  ];
  tarjetas.forEach(([c1, c2, titulo, valor, detalle, formato]) => {
    sh.getRange(`${c1}4:${c2}6`).setBackground(FONDO_TARJETA)
      .setBorder(true, null, null, null, null, null, COLOR_ENCABEZADO, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(`${c1}4:${c2}4`).merge().setValue(titulo).setFontSize(9).setFontColor(GRIS_TEXTO)
      .setHorizontalAlignment('center');
    sh.getRange(`${c1}5:${c2}5`).merge().setFormula(valor).setFontSize(22).setFontWeight('bold')
      .setHorizontalAlignment('center').setNumberFormat(formato);
    sh.getRange(`${c1}6:${c2}6`).merge().setFormula(detalle).setFontSize(9).setFontColor(GRIS_TEXTO)
      .setHorizontalAlignment('center');
  });
  sh.setRowHeight(5, 40);

  // Semáforo del cumplimiento (tarjeta y bloque).
  const semaforo = [sh.getRange('H5'), sh.getRange('E28'), sh.getRange('I22')];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1)
      .setFontColor('#38761d').setRanges(semaforo).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.9, 0.9999)
      .setFontColor('#b45f06').setRanges(semaforo).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.9)
      .setFontColor('#cc0000').setRanges(semaforo).build(),
  ]);

  // --- Por modelo (fila 31 en adelante), con TOTAL fijo arriba.
  sh.getRange('B31:I31').merge().setValue('POR MODELO')
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  const enc = ['Modelo', 'Adjudicados (acto)', 'Aprobados (acto)', 'Susp. c/CC aprobada', '% conversión',
    'Netos del mes', 'V.M', 'VM netos'];
  sh.getRange('B32:I32').setValues([enc]).setFontWeight('bold').setFontSize(9).setFontColor(GRIS_TEXTO)
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, null, COLOR_ENCABEZADO, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B32').setHorizontalAlignment('left');
  sh.setRowHeight(32, 32);

  const mod = r['Modelo ahorro'];
  const lista = 'B34:B80';
  const mapa = (expr) => `=MAP(${lista},LAMBDA(m,IF(m="","",${expr})))`;
  const aj = `${HOJA_AJUSTES}!`;
  sh.getRange('B34').setFormula(
    `=IFERROR(LET(u,UNIQUE({IFERROR(FILTER(${mod},${mod}<>"",(${r.Acto}=${M})+(${r['Mes cierre']}=${M})),"");` +
    `IFERROR(FILTER(${aj}B2:B,${aj}A2:A=${M}),"")}),SORT(FILTER(u,u<>""))),"")`);
  sh.getRange('C34').setFormula(mapa(`COUNTIFS(${acto},${mod},m)`));
  sh.getRange('D34').setFormula(mapa(`COUNTIFS(${acto},${mod},m,${aprob})`));
  sh.getRange('E34').setFormula(mapa(`COUNTIFS(${acto},${mod},m,${r.Pedido},"SUSPENDIDO",${r.Carpeta},"APROBADA")`));
  sh.getRange('F34').setFormula('=MAP(C34:C80,D34:D80,LAMBDA(a,b,IF(N(a)=0,"",b/a)))');
  sh.getRange('G34').setFormula(mapa(`COUNTIFS(${r['Mes cierre']},${M},${mod},m,${aprob})+` +
    `SUMIFS(${aj}$C:$C,${aj}$A:$A,${M},${aj}$B:$B,m)`));
  sh.getRange('H34').setFormula(mapa(`SUMIFS(PRECIOS!$C:$C,PRECIOS!$A:$A,${M},PRECIOS!$B:$B,m)`));
  // VM netos = netos de BASE × V.M del mes + VM de los AJUSTES (que traen su propio importe).
  sh.getRange('I34').setFormula(`=MAP(B34:B80,G34:G80,H34:H80,LAMBDA(m,n,p,IF(m="","",` +
    `(n-SUMIFS(${aj}$C:$C,${aj}$A:$A,${M},${aj}$B:$B,m))*p+SUMIFS(${aj}$D:$D,${aj}$A:$A,${M},${aj}$B:$B,m))))`);
  sh.getRange('C34:I80').setHorizontalAlignment('center');
  sh.getRange('B34:I80').setBorder(null, null, null, null, null, true, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange('B33:I33').setValues([['TOTAL', '=SUM(C34:C80)', '=SUM(D34:D80)', '=SUM(E34:E80)',
    '=IF(N(C33)=0,"",D33/C33)', '=SUM(G34:G80)', '', '=SUM(I34:I80)']])
    .setFontWeight('bold').setBackground(FONDO_TOTAL).setHorizontalAlignment('center');
  sh.getRange('B33').setHorizontalAlignment('left');
  sh.getRange('F33:F80').setNumberFormat('0%');
  sh.getRange('H33:I80').setNumberFormat('$ #,##0');
  sh.getRange('I33:I80').setHorizontalAlignment('right');
  sh.getRange('H34:H80').setHorizontalAlignment('right');
}

/** Un bloque del RESUMEN: título en la fila f0, etiqueta en 3 columnas y valor en la 4ta. */
function bloqueResumen(sh, f0, c0, titulo, filas) {
  sh.getRange(f0, c0, 1, 4).merge().setValue(titulo)
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  filas.forEach(([etiqueta, formula, secundaria], i) => {
    const f = f0 + 1 + i;
    sh.getRange(f, c0, 1, 3).merge().setValue(etiqueta)
      .setFontColor(secundaria ? GRIS_TEXTO : '#000000');
    sh.getRange(f, c0 + 3).setFormula(formula).setHorizontalAlignment('center')
      .setFontWeight(secundaria ? 'normal' : 'bold');
    sh.getRange(f, c0, 1, 4)
      .setBorder(null, null, true, null, null, null, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
  });
}

// ---------------------------------------------------------------------------
// INFORME (mes a mes + por responsable)
// ---------------------------------------------------------------------------

function actualizarInforme() {
  const ss = SpreadsheetApp.getActive();
  const { sh } = hoja(ss, HOJA_INFORME, 3);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.setFrozenColumns(0);
  sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 12) sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());

  const meses = mesesConDatos(ss);
  const r = rangosBase();
  const aprob = `${r.Pedido},"APROBADO",${r.Carpeta},"APROBADA"`;

  sh.setColumnWidth(1, 16);
  sh.setColumnWidth(2, 140);
  for (let c = 3; c <= 11; c++) sh.setColumnWidth(c, 104);
  sh.getRange('B1:G1').merge().setValue('Evolución mensual')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR_ENCABEZADO);
  sh.getRange('B2:K2').merge()
    .setValue('Acto: lo adjudicado en el acto de ese mes y cuánto se aprobó.   Cierre: los netos aprobados en ese mes (de cualquier acto) contra el objetivo.')
    .setFontSize(9).setFontColor(GRIS_TEXTO);
  sh.setRowHeight(1, 34);

  // Encabezado en dos niveles.
  sh.getRange('B4:B5').merge().setValue('Mes');
  sh.getRange('C4:F4').merge().setValue('ACTO');
  sh.getRange('G4:K4').merge().setValue('CIERRE DEL MES');
  sh.getRange('B4:K4').setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('C5:K5').setValues([['Adjudicados', 'Bajas', 'Aprobados', '% conversión',
    'Netos', 'Objetivo', '% cumplimiento', 'Categoría', 'VM netos']]);
  sh.getRange('C5:K5').setFontWeight('bold').setFontSize(9).setFontColor(GRIS_TEXTO)
    .setHorizontalAlignment('center')
    .setBorder(null, null, true, null, null, null, COLOR_ENCABEZADO, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('F4:F80').setBorder(null, null, null, true, null, null, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);

  const f1 = 6;
  const filas = meses.map((mes, i) => {
    const f = f1 + i;
    const A = `$B${f}`;
    return [
      mes,
      `=COUNTIFS(${r.Acto},${A})`,
      `=COUNTIFS(${r.Acto},${A},${r.Pedido},"BAJA ADJ")`,
      `=COUNTIFS(${r.Acto},${A},${aprob})`,
      `=IF(N(C${f})=0,"",E${f}/C${f})`,
      `=COUNTIFS(${r['Mes cierre']},${A},${aprob})+SUMIFS(${HOJA_AJUSTES}!$C:$C,${HOJA_AJUSTES}!$A:$A,${A})`,
      `=IFERROR(VLOOKUP(${A},OBJETIVOS!$A:$B,2,FALSE),"")`,
      `=IF(N(H${f})=0,"",G${f}/H${f})`,
      `=IFERROR(VLOOKUP(${A},OBJETIVOS!$A:$C,3,FALSE),"")`,
      // VM netos: por cada modelo con precio en el mes, V.M × netos de ese modelo (+ AJUSTES).
      // (Un SUMIFS con criterio de matriz dentro de SUMPRODUCT usaba un solo precio para todo.)
      `=IFERROR(LET(p,FILTER(PRECIOS!$B$2:$C,PRECIOS!$A$2:$A=${A}),SUM(MAP(INDEX(p,,1),INDEX(p,,2),` +
        `LAMBDA(m,pr,pr*COUNTIFS(${r['Mes cierre']},${A},${r['Modelo ahorro']},m,${aprob}))))),0)` +
        `+SUMIFS(${HOJA_AJUSTES}!$D:$D,${HOJA_AJUSTES}!$A:$A,${A})`,
    ];
  });
  let fTotal = f1;
  if (filas.length) {
    const ult = f1 + filas.length - 1;
    sh.getRange(f1, 3, filas.length, 9).setFormulas(filas.map((f) => f.slice(1)));
    sh.getRange(f1, 2, filas.length, 1).setNumberFormat('@').setValues(filas.map((f) => [f[0]]))
      .setFontWeight('bold');
    sh.getRange(f1, 2, filas.length, 10)
      .setBorder(null, null, true, null, null, true, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
    fTotal = ult + 1;
    sh.getRange(fTotal, 2, 1, 10).setValues([['TOTAL', `=SUM(C${f1}:C${ult})`, `=SUM(D${f1}:D${ult})`,
      `=SUM(E${f1}:E${ult})`, `=IF(N(C${fTotal})=0,"",E${fTotal}/C${fTotal})`, `=SUM(G${f1}:G${ult})`,
      `=SUM(H${f1}:H${ult})`, `=IF(N(H${fTotal})=0,"",G${fTotal}/H${fTotal})`, '', `=SUM(K${f1}:K${ult})`]])
      .setFontWeight('bold').setBackground(FONDO_TOTAL);
    sh.getRange(f1, 3, filas.length + 1, 9).setHorizontalAlignment('center');
    ['C', 'D', 'E', 'G', 'H'].forEach((c) => sh.getRange(`${c}${f1}:${c}${fTotal}`).setNumberFormat('0'));
    sh.getRange(`F${f1}:F${fTotal}`).setNumberFormat('0%');
    sh.getRange(`I${f1}:I${fTotal}`).setNumberFormat('0%');
    sh.getRange(`K${f1}:K${fTotal}`).setNumberFormat('$ #,##0').setHorizontalAlignment('right');

    const cumpl = [sh.getRange(`I${f1}:I${fTotal}`)];
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1)
        .setFontColor('#38761d').setBold(true).setRanges(cumpl).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.9, 0.9999)
        .setFontColor('#b45f06').setBold(true).setRanges(cumpl).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.9)
        .setFontColor('#cc0000').setBold(true).setRanges(cumpl).build(),
    ]);
  }
  sh.getRange('F4:F5').setBorder(null, null, null, true, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID);

  // Por responsable (todos los actos).
  const f0 = fTotal + 3;
  sh.getRange(f0, 2, 1, 6).merge().setValue('Por responsable (todos los actos)')
    .setFontSize(12).setFontWeight('bold').setFontColor(COLOR_ENCABEZADO);
  const enc2 = ['Responsable', 'Asignados', 'Aprobados', 'Bajas', 'En curso', '% conversión'];
  sh.getRange(f0 + 1, 2, 1, 6).setValues([enc2]).setFontWeight('bold').setBackground(COLOR_ENCABEZADO)
    .setFontColor('white').setHorizontalAlignment('center');
  const resp = valoresLista(ss, 'Responsable');
  const filas2 = resp.map((nombre, i) => {
    const f = f0 + 2 + i;
    const A = `$B${f}`;
    return [
      nombre,
      `=COUNTIFS(${r.Responsable},${A})`,
      `=COUNTIFS(${r.Responsable},${A},${aprob})`,
      `=COUNTIFS(${r.Responsable},${A},${r.Pedido},"BAJA ADJ")`,
      `=C${f}-D${f}-E${f}`,
      `=IF(N(C${f})=0,"",D${f}/C${f})`,
    ];
  });
  if (filas2.length) {
    const p1 = f0 + 2;
    const ult = p1 + filas2.length - 1;
    sh.getRange(p1, 2, filas2.length, 1).setValues(filas2.map((f) => [f[0]])).setFontWeight('bold');
    sh.getRange(p1, 3, filas2.length, 5).setFormulas(filas2.map((f) => f.slice(1)));
    sh.getRange(p1, 2, filas2.length, 6)
      .setBorder(null, null, true, null, null, true, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
    const t = ult + 1;
    sh.getRange(t, 2, 1, 6).setValues([['TOTAL', `=SUM(C${p1}:C${ult})`, `=SUM(D${p1}:D${ult})`,
      `=SUM(E${p1}:E${ult})`, `=SUM(F${p1}:F${ult})`, `=IF(N(C${t})=0,"",D${t}/C${t})`]])
      .setFontWeight('bold').setBackground(FONDO_TOTAL);
    sh.getRange(p1, 3, filas2.length + 1, 5).setHorizontalAlignment('center');
    sh.getRange(p1, 3, filas2.length + 1, 4).setNumberFormat('0');
    sh.getRange(p1, 7, filas2.length + 1, 1).setNumberFormat('0%');
  }

  // Si el RESUMEN no tiene mes elegido, poner el último acto.
  const resumen = ss.getSheetByName(HOJA_RESUMEN);
  if (resumen && !textoMes(resumen.getRange(CELDA_MES_RESUMEN).getValue()) && meses.length) {
    const actos = valoresColumnaBase(ss, 'Acto');
    const ultimo = meses.filter((m) => actos.indexOf(m) >= 0).pop();
    if (ultimo) resumen.getRange(CELDA_MES_RESUMEN).setValue(ultimo);
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
  const clave = (f) => `${textoMes(f[cActo])}|${f[cSol]}`;
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
  return { agregadas: nuevas.length, repetidas: filas.length - nuevas.length };
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

/**
 * Fórmula del % de un bonus de INCENTIVOS para el mes de RESUMEN (H1) y la
 * categoría del bloque INCENTIVO PEDIDOS (I21). Busca el concepto por texto ("CUMPLIMIENTO
 * PEDIDOS", "ADICIONALES") porque el nombre cambia un poco mes a mes.
 * Devuelve 0 si no se cumple la condición.
 */
function bonusSegunCategoria(textoConcepto, condicionNoCumple) {
  const abs = (celda) => `$${celda.replace(/(\d+)/, '$$$1')}`;
  const mes = abs(CELDA_MES_CALC);
  const cat = abs(CELDA_CATEGORIA_RESUMEN);
  return `IF(OR(${cat}="",${condicionNoCumple}),0,IFERROR(INDEX(FILTER(${HOJA_INCENTIVOS}!$C$2:$E,` +
    `${HOJA_INCENTIVOS}!$A$2:$A=${mes},REGEXMATCH(UPPER(${HOJA_INCENTIVOS}!$B$2:$B),"${textoConcepto}")),` +
    `1,MATCH(${cat},{"A","B","C"},0)),0))`;
}

/**
 * Fórmula que devuelve el mes de una celda como "MES AA", esté guardado como texto
 * o como la fecha que arma Sheets (26/09/2026 -> "SEPTIEMBRE 26").
 */
function formulaMesNormalizado(celda) {
  const nombres = MESES.map((m) => `"${m}"`).join(',');
  return `=IF(${celda}="","",IF(ISNUMBER(${celda}),CHOOSE(MONTH(${celda}),${nombres})&" "&TEXT(DAY(${celda}),"00"),` +
    `UPPER(TRIM(${celda}))))`;
}

/** Date que Sheets armó con "SEPTIEMBRE 26" (26/09) -> "SEPTIEMBRE 26"; texto -> en mayúsculas. */
function textoMes(v) {
  if (v instanceof Date) return etiquetaMes(v.getMonth(), v.getDate());
  return v === null || v === undefined ? '' : String(v).trim().toUpperCase();
}

/** Pasa un rango a formato texto y corrige los meses que Sheets había convertido en fecha. */
function mesesComoTexto(rango) {
  const vals = rango.getValues();
  const hayFechas = vals.some((f) => f[0] instanceof Date);
  rango.setNumberFormat('@');
  if (hayFechas) rango.setValues(vals.map((f) => [f[0] instanceof Date ? textoMes(f[0]) : f[0]]));
}

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
  if (valor instanceof Date) return textoMes(valor);
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
  const vals = base.getRange(2, colBase(nombre), ultima - 1, 1).getValues().map((f) => f[0]);
  return (nombre === 'Acto' || nombre === 'Mes cierre' ? vals.map(textoMes) : vals).filter((v) => v !== '');
}

function valoresLista(ss, encabezado) {
  const sh = ss.getSheetByName(HOJA_LISTAS);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const c = enc.indexOf(encabezado) + 1;
  if (!c) return [];
  return sh.getRange(2, c, sh.getMaxRows() - 1, 1).getValues()
    .map((f) => f[0]).filter((v) => v !== '');
}

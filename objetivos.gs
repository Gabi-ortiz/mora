/**
 * OBJETIVOS E INCENTIVOS — Plan de ahorro
 *
 * Base única de lo que manda Fiat: objetivos mensuales (carta de objetivos),
 * categoría ABC por trimestre (mora cuota 6) y reglas de incentivo por mes y
 * categoría (señales comerciales). Va en el mismo archivo que senales.gs:
 * cada señal de la hoja SEÑALES se marca "Cargada" cuando se vuelca acá.
 *
 * Hojas (las arma "Armar / reparar estructura", sin borrar lo ya cargado):
 * - OBJETIVOS: una fila por mes + indicador (formato largo).
 * - CATEGORIAS: una fila por trimestre de vigencia.
 * - INCENTIVOS: una fila por mes + concepto, con el % para A / B / C.
 * - MES: vista de un mes elegido (objetivos, categoría vigente e incentivos).
 * - LISTAS_OBJ: valores de los desplegables.
 *
 * Las cartas de objetivos que se suben a la carpeta de señales se cargan
 * solas en OBJETIVOS (cargarCartaObjetivos, llamada desde senales.gs).
 *
 * Instalación: ver OBJETIVOS.md.
 */

// --- CONFIGURACIÓN ---
const HOJA_OBJ = 'OBJETIVOS';
const HOJA_CAT = 'CATEGORIAS';
const HOJA_INC = 'INCENTIVOS';
const HOJA_MES = 'MES';
const HOJA_LISTAS_OBJ = 'LISTAS_OBJ';
const COLOR_ENC = '#430000';

const ENC_OBJ = ['Mes', 'Concesionario', 'Marca', 'Indicador', 'Objetivo', 'Categoría',
  'Nº carta', 'Fecha carta', 'Fecha flujo', '% flujo', 'Link', 'Notas'];
const ENC_CAT = ['Trimestre', 'Desde', 'Hasta', 'Marca', 'Medición', 'Corte A (≤)', 'Corte C (≥)',
  'Valor real', 'Categoría', 'Nº señal', 'Notas'];
const ENC_INC = ['Mes', 'Marca', 'Concepto', 'Condición', 'A', 'B', 'C', 'Base de cálculo',
  'Nº señal', 'Notas'];

const LISTAS_OBJ = {
  Indicador: ['SUSCRIPCIONES', 'PATENTAMIENTOS', 'PEDIDOS TOTALES'],
  Marca: ['FIAT', 'JEEP', 'RAM'],
  Categoría: ['A', 'B', 'C'],
  Concesionario: ['TURIN S.A.'],
};

// Semillas: lo que salió de la carta 09/2026 y de las señales 1595, 1599 y 1603.
const SEMILLA_OBJ = [
  ['SEPTIEMBRE 26', 'TURIN S.A.', 'FIAT', 'SUSCRIPCIONES', 146, 'A', "'09/2026", new Date(2026, 8, 7),
    new Date(2026, 8, 23), 0.65, '', 'Flujo: ingresar el 65% o más de las suscripciones al día del flujo'],
  ['SEPTIEMBRE 26', 'TURIN S.A.', 'FIAT', 'PATENTAMIENTOS', 43, 'A', "'09/2026", new Date(2026, 8, 7), '', '', '', ''],
  ['SEPTIEMBRE 26', 'TURIN S.A.', 'FIAT', 'PEDIDOS TOTALES', 60, 'A', "'09/2026", new Date(2026, 8, 7), '', '', '', ''],
  ['AGOSTO 26', 'TURIN S.A.', 'FIAT', 'PEDIDOS TOTALES', 57, '', '', '', '', '', '',
    'Tomado del RESUMEN de AGOSTO 26 (falta la carta)'],
];
const SEMILLA_CAT = [
  ['JUL-SEP 26', new Date(2026, 6, 1), new Date(2026, 8, 30), 'FIAT', 'Mora cuota 6', '', '', '', 'A', '',
    'Categoría según carta de objetivos 09/2026; falta la señal con los cortes'],
  ['OCT-DIC 26', new Date(2026, 9, 1), new Date(2026, 11, 31), 'FIAT', 'Mora cuota 6', 0.43, 0.51, '', '', '1595',
    'Cierre medición 10/09/2026, vto. cuota 6 MAY-JUN-JUL. Cargar el % real para calcular la categoría'],
];
const SEMILLA_INC = [
  ['SEPTIEMBRE 26', 'FIAT', '2A Bonus cumplimiento pedidos totales', 'Pedidos ≥ 100% del objetivo y suscripciones ≥ 100%',
    0.01, 0.005, 0.0025, 'Pedidos del mes', '1599', ''],
  ['SEPTIEMBRE 26', 'FIAT', '2B Pedidos adicionales al objetivo', 'Pedidos ≥ 100% del objetivo; por cada pedido excedente',
    0.005, 0.0025, 0.001, 'Precio concesionario sin IVA NC1 al 01/09/26', '1599', ''],
  ['SEPTIEMBRE 26', 'FIAT', '2C Extra bonus pedidos', 'Patentamientos ≥ 80% del objetivo (1603, antes 100%)',
    '', '', '', 'Precio concesionario sin IVA del modelo al 01/09/26; % por modelo', '1599 / 1603',
    '% por modelo en la señal 1599. Excluye Cronos 90/10 B90 desde grupo 18002'],
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Objetivos')
    .addItem('Armar / reparar estructura', 'armarObjetivos')
    .addSeparator()
    .addItem('Revisar carpeta de señales ahora', 'revisarCarpeta')
    .addToUi();
  // De una sola vez, ya usadas: quedaron fuera del menú pero se pueden correr desde
  // el editor de Apps Script si hiciera falta: instalarTriggerSenales (la revisión
  // automática cada hora sigue instalada), leerSenalesRegistradas, rehacerExtractos.
}

function armarObjetivos() {
  const ss = SpreadsheetApp.getActive();
  const listas = armarListasObj(ss);
  const obj = tabla(ss, HOJA_OBJ, ENC_OBJ, SEMILLA_OBJ);
  const cat = tabla(ss, HOJA_CAT, ENC_CAT, SEMILLA_CAT);
  const inc = tabla(ss, HOJA_INC, ENC_INC, SEMILLA_INC);
  limpiarIndicadoresPegados(ss, obj);

  // Meses como texto ("SEPTIEMBRE 26"), corrigiendo los que Sheets convirtió en fecha.
  mesesComoTextoObj(obj.getRange(2, 1, obj.getMaxRows() - 1, 1));
  mesesComoTextoObj(inc.getRange(2, 1, inc.getMaxRows() - 1, 1));

  desplegableObj(obj, ENC_OBJ.indexOf('Indicador') + 1, listas.Indicador, false);
  desplegableObj(obj, ENC_OBJ.indexOf('Marca') + 1, listas.Marca, true);
  desplegableObj(obj, ENC_OBJ.indexOf('Categoría') + 1, listas['Categoría'], true);
  desplegableObj(obj, ENC_OBJ.indexOf('Concesionario') + 1, listas.Concesionario, false);
  obj.getRange('J:J').setNumberFormat('0%');
  obj.getRange('H:I').setNumberFormat('dd/mm/yyyy');

  desplegableObj(cat, ENC_CAT.indexOf('Marca') + 1, listas.Marca, true);
  cat.getRange('F:H').setNumberFormat('0.0%');
  cat.getRange('B:C').setNumberFormat('dd/mm/yyyy');
  // Categoría: si hay valor real y cortes, se calcula; si no, queda lo cargado a mano.
  const filasCat = cat.getMaxRows() - 1;
  const cCat = ENC_CAT.indexOf('Categoría') + 1;
  const formulas = [];
  for (let r = 2; r <= cat.getLastRow(); r++) {
    const actual = cat.getRange(r, cCat).getValue();
    if (actual && !cat.getRange(r, cCat).getFormula()) { formulas.push([actual]); continue; }
    formulas.push([`=IF(OR(H${r}="",F${r}="",G${r}=""),"",IF(H${r}<=F${r},"A",IF(H${r}>=G${r},"C","B")))`]);
  }
  if (formulas.length) cat.getRange(2, cCat, formulas.length, 1).setValues(formulas);
  cat.getRange(2, cCat, filasCat, 1).setFontWeight('bold').setHorizontalAlignment('center');

  desplegableObj(inc, ENC_INC.indexOf('Marca') + 1, listas.Marca, true);
  inc.getRange('E:G').setNumberFormat('0.00%');

  armarVistaMes(ss);
  ss.setActiveSheet(ss.getSheetByName(HOJA_MES));
  SpreadsheetApp.getUi().alert('Estructura de objetivos lista.');
}

/** Crea la hoja con encabezados y semilla si no existe; si existe, solo repara encabezados y formato. */
function tabla(ss, nombre, encabezados, semilla) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange('A:A').setNumberFormat('@'); // Mes / Trimestre como texto
    if (semilla.length) sh.getRange(2, 1, semilla.length, encabezados.length).setValues(semilla);
  }
  sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
    .setFontWeight('bold').setBackground(COLOR_ENC).setFontColor('white').setWrap(true);
  sh.setFrozenRows(1);
  return sh;
}

function armarListasObj(ss) {
  let sh = ss.getSheetByName(HOJA_LISTAS_OBJ);
  const nombres = Object.keys(LISTAS_OBJ);
  if (!sh) {
    sh = ss.insertSheet(HOJA_LISTAS_OBJ);
    nombres.forEach((n, i) => {
      sh.getRange(1, i + 1).setValue(n);
      sh.getRange(2, i + 1, LISTAS_OBJ[n].length, 1).setValues(LISTAS_OBJ[n].map((v) => [v]));
    });
  }
  sh.getRange(1, 1, 1, nombres.length).setFontWeight('bold').setBackground(COLOR_ENC).setFontColor('white');
  sh.setFrozenRows(1);
  const rangos = {};
  nombres.forEach((n, i) => { rangos[n] = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1); });
  return rangos;
}

function desplegableObj(sh, col, rango, estricto) {
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(rango, true).setAllowInvalid(!estricto).build());
}

/** Vista de un mes: objetivos, categoría vigente e incentivos. El mes se elige en B1. */
function armarVistaMes(ss) {
  let sh = ss.getSheetByName(HOJA_MES);
  if (!sh) sh = ss.insertSheet(HOJA_MES, 0);
  const elegido = textoMes(sh.getRange('B1').getValue()) || 'SEPTIEMBRE 26';
  sh.clear();
  sh.getRange('B1').setNumberFormat('@');
  sh.getRange('A1:B1').setValues([['Mes:', elegido]]);
  sh.getRange('A1').setFontWeight('bold');
  sh.getRange('B1').setFontWeight('bold').setBackground('#fff2cc')
    .setNote('Formato MES AA, igual que en Pedidos PDA (ej. OCTUBRE 26)');
  sh.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(HOJA_OBJ).getRange('A2:A'), true).setAllowInvalid(true).build());

  // H1 (oculta): el mes de B1 normalizado a "MES AA". Sheets a veces guarda la opción
  // del desplegable como fecha (26/09/2026); las fórmulas usan H1, no B1.
  const nombresMes = MESES_SENAL.map((m) => `"${m}"`).join(',');
  sh.getRange('H1').setFormula(`=IF($B$1="","",IF(ISNUMBER($B$1),CHOOSE(MONTH($B$1),${nombresMes})&" "&` +
    'TEXT(DAY($B$1),"00"),UPPER(TRIM($B$1))))');
  sh.hideColumns(8);

  const titulo = (celda, texto) => sh.getRange(celda).setValue(texto)
    .setFontWeight('bold').setBackground(COLOR_ENC).setFontColor('white');

  titulo('A3', 'OBJETIVOS DEL MES');
  sh.getRange('A4:E4').setValues([['Indicador', 'Marca', 'Objetivo', 'Fecha flujo', '% flujo']]).setFontWeight('bold');
  sh.getRange('A5').setFormula(
    `=IFERROR(FILTER({${HOJA_OBJ}!D2:D,${HOJA_OBJ}!C2:C,${HOJA_OBJ}!E2:E,${HOJA_OBJ}!I2:I,${HOJA_OBJ}!J2:J},` +
    `${HOJA_OBJ}!A2:A=$H$1),"Sin objetivos cargados para este mes")`);
  sh.getRange('D5:D12').setNumberFormat('dd/mm/yyyy');
  sh.getRange('E5:E12').setNumberFormat('0%');

  titulo('A14', 'CATEGORÍA');
  sh.getRange('A15:B15').setValues([['Categoría del mes (carta)', '']]);
  sh.getRange('B15').setFormula(
    `=IFERROR(INDEX(FILTER(${HOJA_OBJ}!F2:F,${HOJA_OBJ}!A2:A=$H$1,${HOJA_OBJ}!F2:F<>""),1),"")`)
    .setFontWeight('bold').setFontSize(14);

  titulo('A17', 'INCENTIVOS DEL MES');
  sh.getRange('A18:F18').setValues([['Concepto', 'Condición', 'A', 'B', 'C', 'Nº señal']]).setFontWeight('bold');
  sh.getRange('A19').setFormula(
    `=IFERROR(FILTER({${HOJA_INC}!C2:C,${HOJA_INC}!D2:D,${HOJA_INC}!E2:E,${HOJA_INC}!F2:F,${HOJA_INC}!G2:G,${HOJA_INC}!I2:I},` +
    `${HOJA_INC}!A2:A=$H$1),"Sin incentivos cargados para este mes")`);
  sh.getRange('C19:E40').setNumberFormat('0.00%');

  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 320);
}

// ---------------------------------------------------------------------------
// Carga automática de cartas de objetivos (la llama senales.gs)
// ---------------------------------------------------------------------------

/**
 * Vuelca una carta interpretada (ver interpretarCarta en senales.gs) en
 * OBJETIVOS. Clave = Mes + Concesionario + Marca + Indicador: si la fila ya
 * existe con el mismo objetivo solo completa lo que falte; si el objetivo
 * cambió, lo actualiza y deja constancia en Notas. Devuelve un resumen.
 */
function cargarCartaObjetivos(ss, carta, link) {
  armarListasObj(ss);
  const sh = ss.getSheetByName(HOJA_OBJ) || tabla(ss, HOJA_OBJ, ENC_OBJ, []);
  const c = (h) => ENC_OBJ.indexOf(h);
  const conc = carta.concesionario || LISTAS_OBJ.Concesionario[0];
  const ultima = sh.getLastRow();
  const datos = ultima > 1 ? sh.getRange(2, 1, ultima - 1, ENC_OBJ.length).getValues() : [];
  const conFlujo = carta.indicadores.some((x) => x.indicador === 'SUSCRIPCIONES')
    ? 'SUSCRIPCIONES' : carta.indicadores[0].indicador;
  const nCarta = carta.nCarta ? `'${carta.nCarta}` : '';

  const nuevas = [];
  const cambios = [];
  const iguales = [];
  const faltan = [];
  const completados = [];
  carta.indicadores.forEach((x) => {
    const flujo = x.indicador === conFlujo;
    const sinNumero = x.objetivo === '';
    const fila = [carta.mes, conc, carta.marca, x.indicador, x.objetivo, carta.categoria, nCarta,
      aFecha(carta.fechaCarta), flujo ? aFecha(carta.fechaFlujo) : '', flujo ? carta.pctFlujo : '',
      link, sinNumero ? `⚠ Completar objetivo (no se pudo leer del PDF de la carta ${carta.nCarta})`
        : carta.deducida ? `⚠ Verificar: deducido del PDF de la carta ${carta.nCarta} (el OCR pegó los números)`
          : `Cargado automáticamente de la carta ${carta.nCarta}`.trim()];
    const i = datos.findIndex((d) => textoMes(d[c('Mes')]) === carta.mes && d[c('Concesionario')] === conc &&
      (d[c('Marca')] === carta.marca || d[c('Marca')] === '') && d[c('Indicador')] === x.indicador);
    if (i < 0) {
      nuevas.push(fila);
      if (sinNumero) faltan.push(x.indicador);
      return;
    }
    const actual = datos[i];
    const antes = actual[c('Objetivo')];
    const combinada = actual.map((v, j) => (v === '' || v === null ? fila[j] : v));
    combinada[c('Mes')] = carta.mes;
    if (sinNumero) {
      faltan.push(x.indicador);
    } else if (antes === '' || antes === null) {
      // Fila que había quedado sin número en una lectura anterior: se completa.
      combinada[c('Objetivo')] = x.objetivo;
      combinada[c('Link')] = link;
      combinada[c('Notas')] = fila[c('Notas')];
      completados.push(`${x.indicador} ${x.objetivo}`);
    } else if (Number(antes) !== x.objetivo) {
      combinada[c('Objetivo')] = x.objetivo;
      combinada[c('Nº carta')] = nCarta || combinada[c('Nº carta')];
      combinada[c('Link')] = link;
      combinada[c('Notas')] = `Actualizado por carta ${carta.nCarta} (antes ${antes})`;
      cambios.push(`${x.indicador} ${antes} → ${x.objetivo}`);
    } else {
      iguales.push(x.indicador);
    }
    sh.getRange(i + 2, 1, 1, ENC_OBJ.length).setValues([combinada]);
  });
  sh.getRange('A:A').setNumberFormat('@');
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, ENC_OBJ.length).setValues(nuevas);

  agregarAListaObj(ss, 'Indicador', carta.indicadores.map((x) => x.indicador));
  agregarAListaObj(ss, 'Concesionario', [conc]);

  const partes = [];
  if (nuevas.length) partes.push(`${nuevas.length} objetivo(s) nuevo(s)`);
  if (completados.length) partes.push(`completados: ${completados.join(', ')}`);
  if (cambios.length) partes.push(`cambiaron: ${cambios.join(', ')}`);
  if (iguales.length) partes.push(`sin cambios: ${iguales.join(', ')}`);
  if (faltan.length) partes.push(`sin número en el PDF (completar si falta): ${faltan.join(', ')}`);
  return `${carta.mes}: ${partes.join('; ')}`;
}

/** Agrega a una columna de LISTAS_OBJ los valores que todavía no están. */
function agregarAListaObj(ss, encabezado, valores) {
  const sh = ss.getSheetByName(HOJA_LISTAS_OBJ);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = enc.indexOf(encabezado) + 1;
  if (!col) return;
  const actuales = sh.getRange(2, col, sh.getMaxRows() - 1, 1).getValues()
    .map((f) => f[0]).filter((v) => v !== '');
  const faltan = [...new Set(valores)].filter((v) => v && actuales.indexOf(v) < 0);
  if (faltan.length) {
    sh.getRange(actuales.length + 2, col, faltan.length, 1).setValues(faltan.map((v) => [v]));
  }
}

/** "07/09/2026" -> Date (o '' si no hay fecha). */
function aFecha(texto) {
  const m = String(texto || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : '';
}

/**
 * Borra lo que dejó una carta leída antes de que el lector separara indicadores
 * pegados por el OCR: filas de OBJETIVOS cuyo indicador son varios juntos (p. ej.
 * "PATENTAMIENTOS PEDIDOS TOTALES"), y ese valor en LISTAS_OBJ.
 */
function limpiarIndicadoresPegados(ss, obj) {
  const listas = ss.getSheetByName(HOJA_LISTAS_OBJ);
  const enc = listas.getRange(1, 1, 1, listas.getLastColumn()).getValues()[0];
  const cInd = enc.indexOf('Indicador') + 1;
  if (!cInd) return;
  const rInd = listas.getRange(2, cInd, listas.getMaxRows() - 1, 1);
  const valores = rInd.getValues().map((f) => String(f[0]).trim()).filter((v) => v !== '');
  const pegado = (v) => separarIndicadores(v, valores.filter((x) => x !== v)).length >= 2;
  // (también detecta "PATENTAMIENTOS PEDIDOS CRONOS PEDIDOS TITANO PEDIDOS TOTALES")
  const malos = valores.filter(pegado);
  if (!malos.length) return;

  const iInd = ENC_OBJ.indexOf('Indicador');
  const iObj = ENC_OBJ.indexOf('Objetivo');
  for (let r = obj.getLastRow(); r >= 2; r--) {
    const fila = obj.getRange(r, 1, 1, ENC_OBJ.length).getValues()[0];
    // Indicador "pegado" = nombre inválido (el número, si tiene, quedó mal asignado): se borra.
    if (malos.indexOf(String(fila[iInd]).trim()) >= 0) obj.deleteRow(r);
  }
  const quedan = valores.filter((v) => malos.indexOf(v) < 0);
  rInd.clearContent();
  rInd.offset(0, 0, quedan.length, 1).setValues(quedan.map((v) => [v]));
}

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
 * - PRECIOS: precios por mes y modelo (se carga una vez por mes). El VM se calcula con Prec CC s/IVA.
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
const HOJA_LICITACIONES = 'Licitaciones'; // la arma el equipo (solicitud en D, pago en N)
const HOJA_OBJETIVOS = 'OBJETIVOS';
const HOJA_LISTAS = 'LISTAS';
const HOJA_INCENTIVOS = 'INCENTIVOS';
const HOJA_AJUSTES = 'AJUSTES';

// Archivo "Objetivos y señales comerciales - PDA" (objetivos, categoría e incentivos).
const ID_ARCHIVO_OBJETIVOS = '1TDMqgJkbkP11dSlUhvJpD5NhKtkj0vvuiu-pPPQ5kYo';

const FILAS_BASE = 5000;

// Columnas de BASE. El script las busca por el nombre del encabezado (ver mapaBase): este orden
// solo se usa al crear BASE. sga = nombre de la columna en el reporte de SGA
// (null = se completa a mano o la pone el script).
const COLUMNAS_BASE = [
  { h: 'Acto', oculta: false, sga: null },
  { h: 'Concesionario', oculta: true, sga: 'Concesionario' },
  { h: 'Concepto de venta', oculta: true, sga: 'Concepto de venta' },
  { h: 'Modelo ahorro', oculta: false, sga: 'Modelo ahorro' },
  { h: 'Modelo PEDIDO', oculta: false, sga: null }, // se carga a mano; es el modelo de los cálculos
  { h: '$ descuento', oculta: false, sga: null }, // importe bonificado (a mano); resta en RESUMEN
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

// Otros nombres con los que el equipo renombró columnas de BASE.
const ALIAS_BASE = { 'Llave x llave': ['LLXLL', 'Llave por llave'] };
let MAPA_BASE = null; // encabezados de BASE leídos en esta ejecución (ver mapaBase)

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
  { id: '1dcy1DYrZi8-eihD3GP4LNYLa2zXXm3O5rfY-Ymz27Qw', hoja: 'JULIO', acto: 'JULIO 26' },
  { id: '16IQ8wKAHrRA2Oy2pntrB9ZeZxjTZkdjPj-7kZPYuREY', hoja: 'AGOSTO', acto: 'AGOSTO 26' },
  { id: '1_5SnWB3QvWHHkV2nNZ0kH3ZVhNBNJX0W4ipnlBrvaOI', hoja: 'SEPTIEMBRE', acto: 'SEPTIEMBRE 26' },
];

// Cierres que no están en BASE, para que los meses den igual que sus RESUMEN mensuales:
// [Mes, Modelo, Netos, VM extra, Nota]. Los netos se valúan con el Prec CC s/IVA del mes de cierre
// (PRECIOS), como cualquier cierre; "VM extra" queda solo para correcciones de importe.
const NOTA_JUNIO = 'Cierres en julio del acto de junio (archivo JUNIO 26.xlsx; junio no está en BASE).';
const AJUSTES_JUNIO_EN_JULIO = [['DP1', 14], ['MB1', 6], ['AR2', 4], ['FO1', 2], ['FT3', 1], ['FP3', 1]]
  .map(([m, n]) => ['JULIO 26', m, n, '', NOTA_JUNIO]);
const AJUSTES_INICIALES = [
  ...AJUSTES_JUNIO_EN_JULIO,
  ['JULIO 26', 'DP1', 1, '', 'Cierres en julio del acto de Mayo (archivo MAYO 26.xlsx).'],
  ['SEPTIEMBRE 26', 'NC1', 1, '',
    '"+1" cargado a mano en el RESUMEN de SEPTIEMBRE 26 (NETOS ANTERIOR). Confirmar de qué solicitud es.'],
];

// Columnas de PRECIOS. El VM (RESUMEN, INFORME, AJUSTES) se calcula con Prec CC s/IVA.
const ENCABEZADOS_PRECIOS = ['Mes', 'Modelo', 'Version', 'Inc pedido', 'Prec publico c/IVA',
  'Prec CC c/IVA', 'Prec CC s/IVA', 'Familia'];
const COL_FAMILIA = 'H'; // PRECIOS: familia del modelo (Cronos, Strada...), la carga el equipo
const COL_PRECIO_VM = 'G'; // Prec CC s/IVA

// VM netos del INFORME de los meses que todavía no tienen Prec CC s/IVA cargado: queda el
// importe que se calculaba con el precio cliente con IVA, como valor fijo y aclarado. Cuando se
// cargan los Prec CC s/IVA del mes, el INFORME vuelve a calcularlo solo.
const VM_INFORME_A_CORREGIR = {
  'JULIO 26': 1855113000,
  'AGOSTO 26': 2006468000,
  'SEPTIEMBRE 26': 2711975000,
};
const NOTA_VM_A_CORREGIR = 'Precio cliente con iva - a corregir';

// Semilla de PRECIOS al crear la hoja: precio cliente con IVA (Prec publico c/IVA).
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
    .addItem('Generar informe de incentivos', 'generarInforme')
    .addSeparator()
    .addItem('Armar / reparar estructura', 'armarEstructura')
    .addToUi();
  // ordenarBase() salió del menú (se sigue usando sola al agregar un acto).
  // importarHistorico() ya se usó (agosto y septiembre en BASE): quedó fuera del menú.
  // Si hiciera falta de nuevo, se corre desde el editor de Apps Script (no duplica).
  try { ajustarFilasResumen(); } catch (err) { /* sin RESUMEN armado todavía */ }
}

/**
 * BASE: al marcar Pedido = APROBADO y Carpeta = APROBADA completa "Mes cierre"
 * con el mes actual (si está vacío). Si "Mes cierre" se escribe a mano, lo
 * pasa al formato MES AA y avisa si no es un mes válido.
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  const nombre = sh.getName();
  // RESUMEN: al cambiar el mes (H1) o algo que cambia los cuadros (AJUSTES, PRECIOS), se muestran solo las
  // filas con modelos.
  if (nombre === HOJA_RESUMEN || nombre === HOJA_AJUSTES || nombre === HOJA_PRECIOS) {
    if (nombre !== HOJA_RESUMEN || e.range.getA1Notation().replace(/:.*/, '') === CELDA_MES_RESUMEN) {
      SpreadsheetApp.flush();
      ajustarFilasResumen();
    }
    return;
  }
  if (nombre !== HOJA_BASE || e.range.getRow() < 2) return;
  edicionBase(e, sh);
  // Columnas que cambian qué modelos aparecen en los cuadros del RESUMEN.
  const c0e = e.range.getColumn();
  const c1e = c0e + e.range.getNumColumns() - 1;
  const afecta = ['Pedido', 'Carpeta', 'Mes cierre', 'Modelo PEDIDO', 'Modelo ahorro', 'Responsable']
    .some((h) => { const c = colBase(h); return c >= c0e && c <= c1e; });
  if (afecta) {
    SpreadsheetApp.flush();
    ajustarFilasResumen();
  }
}

function edicionBase(e, sh) {
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
  const datos = sh.getRange(fila0, 1, n, sh.getLastColumn()).getValues();
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
  if (!verificarColumnasBase()) return;
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

  // El orden de las pestañas solo se acomoda si se creó alguna hoja nueva (si el equipo
  // las reordenó, se respeta).
  if (hojasCreadas) {
    [HOJA_BASE, HOJA_PEGAR, HOJA_RESUMEN, HOJA_INFORME, HOJA_PRECIOS, HOJA_AJUSTES, HOJA_OBJETIVOS,
      HOJA_INCENTIVOS, HOJA_LISTAS]
      .forEach((nombre, i) => {
        ss.setActiveSheet(ss.getSheetByName(nombre));
        ss.moveActiveSheet(i + 1);
      });
  }

  ss.setActiveSheet(ss.getSheetByName(HOJA_BASE));
  SpreadsheetApp.getUi().alert('Estructura lista.');
}

let hojasCreadas = 0;
function hoja(ss, nombre, posicion) {
  let sh = ss.getSheetByName(nombre);
  const nueva = !sh;
  if (nueva) {
    sh = ss.insertSheet(nombre, Math.min(posicion, ss.getSheets().length));
    hojasCreadas++;
  }
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
  // Valores nuevos de las listas que trae el script (p. ej. REINGRESADA) en una hoja ya
  // creada: cada valor se agrega UNA sola vez. Si después se borra a mano, no vuelve.
  const props = PropertiesService.getDocumentProperties();
  const aplicados = JSON.parse(props.getProperty('LISTAS_APLICADAS') || 'null') ||
    // Primera vez con este control: todo lo que ya traía el script se toma como aplicado,
    // salvo lo agregado después (REINGRESADA), para no reponer lo que se haya borrado.
    Object.fromEntries(nombres.map((n) => [n, LISTAS[n].filter((v) => v !== 'REINGRESADA')]));
  if (nueva) nombres.forEach((n) => { aplicados[n] = LISTAS[n].slice(); });
  nombres.forEach((nombre, i) => {
    const yaAplicados = aplicados[nombre] || [];
    const actuales = sh.getRange(2, i + 2, sh.getMaxRows() - 1, 1).getValues().map((f) => f[0]).filter((v) => v !== '');
    const faltan = LISTAS[nombre].filter((v) => yaAplicados.indexOf(v) < 0 && actuales.indexOf(v) < 0);
    if (faltan.length) sh.getRange(actuales.length + 2, i + 2, faltan.length, 1).setValues(faltan.map((v) => [v]));
    aplicados[nombre] = [...new Set(yaAplicados.concat(LISTAS[nombre]))];
  });
  props.setProperty('LISTAS_APLICADAS', JSON.stringify(aplicados));

  // Columnas de versiones anteriores que ya no se usan (solo esas; otras columnas que
  // se agreguen a mano a la derecha se respetan).
  const encActual = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  encActual.forEach((h, i) => {
    if (i > nombres.length && ['RESPONSABLE', 'Meses cierre', 'Meses con datos'].indexOf(h) >= 0) {
      sh.getRange(1, i + 1, sh.getMaxRows(), 1).clear();
    }
  });

  // A: "Meses con datos" = UNIQUE de los meses de BASE (Acto y Mes cierre), el más reciente
  // primero. Es la lista del selector de mes del RESUMEN. Calculada: no se edita.
  const nombresMes = MESES.map((m) => `"${m}"`).join(',');
  sh.getRange('A:A').clear().setNumberFormat('@');
  sh.getRange('A1').setValue('Meses con datos');
  const lActo = letra(colBase('Acto'));
  const lCierre = letra(colBase('Mes cierre'));
  sh.getRange('A2').setFormula(
    `=IFERROR(LET(v,{${HOJA_BASE}!${lActo}2:${lActo};${HOJA_BASE}!${lCierre}2:${lCierre}},` +
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
  const { sh, nueva } = hoja(ss, HOJA_BASE, 0);
  const n = COLUMNAS_BASE.length;
  if (sh.getMaxRows() < FILAS_BASE) sh.insertRowsAfter(sh.getMaxRows(), FILAS_BASE - sh.getMaxRows());
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());

  if (!nueva) {
    // BASE ya armada: NO se toca nada de lo que el equipo cambió (nombres de columnas, columnas
    // ocultas, anchos, casillas / desplegables, colores, bordes). Solo se corrigen los meses que
    // Sheets haya convertido en fecha. Las columnas se buscan por nombre (ver mapaBase), así que
    // se pueden agregar o mover columnas (p. ej. "Modelo PEDIDO") sin romper nada.
    mesesComoTexto(sh.getRange(2, colBase('Acto'), sh.getMaxRows() - 1, 1));
    mesesComoTexto(sh.getRange(2, colBase('Mes cierre'), sh.getMaxRows() - 1, 1));
    if (!sh.getFilter()) sh.getRange(1, 1, sh.getMaxRows(), sh.getLastColumn()).createFilter();
    resaltarLicitaciones(sh);
    return;
  }

  sh.getRange(1, 1, 1, n).setValues([COLUMNAS_BASE.map((c) => c.h)])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white')
    .setWrap(true).setVerticalAlignment('middle');
  MAPA_BASE = null; // recién escritos los encabezados
  sh.setFrozenRows(1);
  COLUMNAS_BASE.forEach((c, i) => {
    if (c.oculta) sh.hideColumns(i + 1); else sh.showColumns(i + 1);
  });
  sh.getRange(2, colBase('Acto'), sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, colBase('Mes cierre'), sh.getMaxRows() - 1, 1).setNumberFormat('@');

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
  [rActo, rCierre].forEach((rg) => rg.setBackground('#f3f3f3').setFontColor(GRIS_TEXTO));
  rActo.protect().setDescription('Acto (viene de SGA)').setWarningOnly(true);
  desplegable('Responsable', listas.Responsable, true);
  desplegable('Pedido', listas.Pedido, true);
  desplegable('Carpeta', listas.Carpeta, true);
  desplegable('Modalidad', listas.Modalidad, true);
  sh.getRange(2, colBase('Llave x llave'), filas, 1).insertCheckboxes();

  // Colores de Pedido.
  const rPedido = sh.getRange(2, colBase('Pedido'), filas, 1);
  const color = (texto, fondo) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(texto).setBackground(fondo).setRanges([rPedido]).build();
  sh.setConditionalFormatRules([
    color('APROBADO', '#d9ead3'),
    color('SUSPENDIDO', '#fce5cd'),
    color('BAJA ADJ', '#f4cccc'),
  ]);

  sh.getRange(2, colBase('Observacion'), filas, 1).setWrap(false);
  sh.setColumnWidth(colBase('Observacion'), 320);
  sh.setColumnWidth(colBase('Apellido, Nombre'), 220);
  sh.getRange(1, 1, sh.getMaxRows(), n)
    .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(1, 1, sh.getMaxRows(), n).createFilter();
  resaltarLicitaciones(sh);
}

/**
 * Resalta la Solicitud de BASE si figura en la hoja Licitaciones (solicitud en D) con
 * "Turin" o "FCA" en la columna N. Solo agrega la regla si no está: no toca los demás formatos
 * condicionales ni los colores que el equipo pinta a mano.
 */
function resaltarLicitaciones(sh) {
  const reglas = sh.getConditionalFormatRules();
  const yaEsta = reglas.some((regla) => {
    const cond = regla.getBooleanCondition();
    return cond && cond.getCriteriaValues().some((v) => String(v).indexOf(`${HOJA_LICITACIONES}!`) >= 0);
  });
  if (yaEsta) return;
  const c = letra(colBase('Solicitud'));
  const lic = (col) => `INDIRECT("${HOJA_LICITACIONES}!${col}:${col}")`;
  const cuenta = (pago) => `COUNTIFS(${lic('D')},$${c}2&"",${lic('N')},"${pago}")`;
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${c}2<>"",${cuenta('Turin')}+${cuenta('FCA')}>0)`)
    .setBackground('#cfe2f3').setBold(true)
    .setRanges([sh.getRange(2, colBase('Solicitud'), sh.getMaxRows() - 1, 1)])
    .build());
  sh.setConditionalFormatRules(reglas);
}

/**
 * Las columnas de BASE se buscan por el nombre del encabezado. Si falta alguna (se borró o se
 * renombró), avisa cuál y devuelve false para no armar fórmulas con columnas equivocadas.
 */
function verificarColumnasBase() {
  MAPA_BASE = null;
  const faltan = mapaBase().faltan;
  if (!faltan.length) return true;
  SpreadsheetApp.getUi().alert('En BASE no encuentro estas columnas: ' + faltan.join(', ') + '.\n' +
    '¿Se borraron o se les cambió el nombre? Volvé a ponerles ese nombre en la fila 1 ' +
    '(el orden no importa y se pueden agregar columnas nuevas).');
  return false;
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
        filas.push([mes, modelo, '', '', PRECIOS_INICIALES[mes][modelo], '', '', '']);
      });
    });
    const n = ENCABEZADOS_PRECIOS.length;
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, n).setValues([ENCABEZADOS_PRECIOS]);
    sh.getRange(2, 1, filas.length, n).setValues(filas);
    sh.getRange('D:D').setNumberFormat('0.0%');
    sh.getRange('E:G').setNumberFormat('$ #,##0');
    sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
    sh.setFrozenRows(1);
  } else if (String(sh.getRange('C1').getValue()).trim() === 'V.M') {
    SpreadsheetApp.getUi().alert('PRECIOS tiene el formato viejo (Mes, Modelo, V.M). Pasar V.M a ' +
      '"Prec publico c/IVA" (columna E) y armar las columnas ' + ENCABEZADOS_PRECIOS.join(', ') + '.');
  }
  // Hoja ya creada: los precios (valores y formato) no se tocan. Solo se corrige un mes
  // que Sheets haya convertido en fecha (26/09/2026 -> "SEPTIEMBRE 26"); el resto queda igual.
  mesesComoTexto(sh.getRange(2, 1, sh.getMaxRows() - 1, 1));
  armarFaltanPrecios(sh);
}

/**
 * PRECIOS!I:J (calculado, no se edita): modelos con netos en un mes (BASE aprobados por Mes
 * cierre, o AJUSTES) que no tienen Prec CC s/IVA cargado para ese mes. Sin ese precio, su VM
 * da 0. Desaparecen solos al cargar el precio.
 */
function armarFaltanPrecios(sh) {
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  sh.getRange('I:J').clearContent();
  sh.getRange('I1:J1').setValues([['Falta precio: Modelo', 'Período']])
    .setFontWeight('bold').setBackground('#b45f06').setFontColor('white');
  const cierre = colBase('Mes cierre');
  const col = (c) => `${HOJA_BASE}!${letra(c)}2:${letra(c)}`;
  const mp = `ARRAYFORMULA(IF(${col(colBase('Modelo PEDIDO'))}<>"",${col(colBase('Modelo PEDIDO'))},` +
    `${col(colBase('Modelo ahorro'))}))`;
  const nombres = MESES.map((m) => `"${m}"`).join(',');
  sh.getRange('I2').setFormula(
    `=LET(b,IFERROR(FILTER({${mp},${col(cierre)}},${col(colBase('Pedido'))}="APROBADO",` +
    `${col(colBase('Carpeta'))}="APROBADA",${col(cierre)}<>""),{"",""}),` +
    `a,IFERROR(FILTER({${HOJA_AJUSTES}!B2:B,${HOJA_AJUSTES}!A2:A},${HOJA_AJUSTES}!C2:C<>"",${HOJA_AJUSTES}!C2:C<>0),{"",""}),` +
    `t,{b;a},u,UNIQUE(FILTER(t,INDEX(t,,1)<>"",INDEX(t,,2)<>"")),` +
    `f,MAP(INDEX(u,,1),INDEX(u,,2),LAMBDA(mo,me,COUNTIFS($B$2:$B,mo,$A$2:$A,me,$${COL_PRECIO_VM}$2:$${COL_PRECIO_VM},"<>")=0)),` +
    `x,FILTER(u,f),p,INDEX(x,,2),` +
    `IFERROR(SORT(x,ARRAYFORMULA(IFERROR(DATE(2000+VALUE(RIGHT(p,2)),MATCH(LEFT(p,LEN(p)-3),{${nombres}},0),1),0)),TRUE,INDEX(x,,1),TRUE),` +
    `"Todos los modelos con netos tienen precio"))`);
  sh.getRange('I2:J').setFontColor('#b45f06');
  sh.setColumnWidth(9, 170);
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
    // La carta lo trae como "PEDIDOS TOTALES" o solo "PEDIDOS" (octubre 2026).
    `"select Col1, Col5, Col6 where (Col4 = 'PEDIDOS TOTALES' or Col4 = 'PEDIDOS') and Col3 = 'FIAT'",0),` +
    `"Hacer clic acá y elegir Permitir acceso al archivo de objetivos")`);
  // Objetivos de pedidos por modelo (p. ej. PEDIDOS CRONOS 27 y PEDIDOS TITANO 4 dentro de los 54).
  sh.getRange(1, 5, 1, 3).setValues([['Mes', 'Indicador', 'Objetivo']])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.getRange('E2').setFormula(
    `=IFERROR(QUERY(IMPORTRANGE("${ID_ARCHIVO_OBJETIVOS}","OBJETIVOS!A2:F"),` +
    `"select Col1, Col4, Col5 where Col4 starts with 'PEDIDOS ' and Col4 <> 'PEDIDOS TOTALES' and Col3 = 'FIAT'",0),"")`);
  sh.getRange('I1').setValue('Se cargan en el archivo "Objetivos y señales comerciales - PDA" (no editar acá).')
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
  if (sh.getRange('D1').getValue() !== 'VM netos' && !nueva) return; // ya armada: no se toca
  sh.getRange(1, 1, 1, 5).setValues([['Mes', 'Modelo', 'Netos', 'VM extra (opcional)', 'Nota']])
    .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
  sh.getRange('C1').setNote('Netos que no están en BASE. Se valúan con el Prec CC s/IVA del mes de PRECIOS ' +
    '(como todo cierre: vale el precio del mes en que se liquida).');
  sh.getRange('D1').setNote('Solo para corregir importes. Vacío en los ajustes de netos: el VM sale de PRECIOS.');
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
const CELDA_MES_CALC = 'Z1'; // oculta; B:M las usan los cuadros
const CELDA_CATEGORIA_RESUMEN = 'M9';

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
  sh.clearNotes(); // clear() no borra las notas ni las alturas de fila del armado anterior
  sh.setRowHeights(1, sh.getMaxRows(), 21);
  sh.clearConditionalFormatRules();
  sh.setFrozenRows(0);
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  sh.setHiddenGridlines(true);

  const r = rangosBase();
  const M = `$${CELDA_MES_CALC.replace(/(\d+)/, '$$$1')}`; // $Z$1
  if (sh.getMaxColumns() < 26) sh.insertColumnsAfter(sh.getMaxColumns(), 26 - sh.getMaxColumns());
  sh.showColumns(1, sh.getMaxColumns()); // la celda oculta antes estaba en K y después en M
  sh.getRange(CELDA_MES_CALC).setFormula(formulaMesNormalizado(`$${CELDA_MES_RESUMEN.replace(/(\d+)/, '$$$1')}`));
  sh.hideColumns(sh.getRange(CELDA_MES_CALC).getColumn());
  const aprob = `${r.Pedido},"APROBADO",${r.Carpeta},"APROBADA"`;
  const acto = `${r.Acto},${M}`;

  // Grilla: A = margen; B..I = 8 columnas iguales.
  sh.setColumnWidth(1, 16);
  // Anchos para ver todo en una pantalla: B (etiquetas) un poco más ancha; C..M iguales
  // (J:M: proyección por familia; M: también PROYECCION del cuadro por modelo).
  sh.setColumnWidth(2, 110);
  for (let c = 3; c <= 16; c++) sh.setColumnWidth(c, 94); // hasta P (proyección por modelo)

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

  // --- Bloques (filas 8-18), uno al lado del otro: ADJUDICADOS DEL ACTO (B:E), CIERRES DEL MES (F:I) e
  // INCENTIVO PEDIDOS (J:M). Valores en E, I y M.
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
  // Cuadros por modelo (más abajo): clientes y TP. El VM netos del mes es la suma de los dos.
  const T = tablasResumen();
  const izq2 = [
    ['Netos del mes (todos los actos)',
      `=COUNTIFS(${r['Mes cierre']},${M},${aprob})+SUMIFS(${HOJA_AJUSTES}!$C:$C,${HOJA_AJUSTES}!$A:$A,${M})`],
    ['   del acto del mes', `=COUNTIFS(${acto},${r['Mes cierre']},${M},${aprob})`, true],
    ['   de actos anteriores', '=I9-I10', true],
    ['   TP', `=COUNTIFS(${r['Mes cierre']},${M},${aprob},${r.Responsable},"TP")`, true],
    ['Objetivo', `=IFERROR(VLOOKUP(${M},OBJETIVOS!$A:$B,2,FALSE),"")`],
    ['Proyección (+15%)', '=IF(I13="","",ROUND(I13*1.15))'],
    ['Faltan para el objetivo', '=IF(I13="","",MAX(0,I13-I9))'],
    ['% cumplimiento', '=IF(N(I13)=0,"",I9/I13)'],
    ['VM netos del mes', `=F${T.clientes.total}+F${T.tp.total}`],
  ];
  const der2 = [
    ['Categoría del mes', '=IFERROR(VLOOKUP(' + M + ',OBJETIVOS!$A:$C,3,FALSE),"")'],
    ['% cumplimiento pedidos', '=I16'],
    ['Bonus cumplimiento pedidos totales', `=${bonusSegunCategoria('CUMPLIMIENTO PEDIDOS', 'N(I16)<1')}`],
    ['Pedidos adicionales al objetivo', '=IF(N(I13)=0,"",MAX(0,I9-I13))'],
    ['Bonus por pedido adicional', `=${bonusSegunCategoria('ADICIONALES', 'N(M12)=0')}`],
  ];
  bloqueResumen(sh, 8, 2, 'ADJUDICADOS DEL ACTO', izq1);
  bloqueResumen(sh, 8, 6, 'CIERRES DEL MES', izq2);
  bloqueResumen(sh, 8, 10, 'INCENTIVO PEDIDOS', der2);
  // Objetivos de pedidos por modelo dentro del total: "Cronos 20/27 · Titano 1/4" (netos del mes
  // cuyo modelo contiene CRONOS / TITANO, contra su objetivo). El modelo es la Version de PRECIOS
  // del Modelo PEDIDO en el mes; si no hay, la descripción de producto de la fila.
  const pp = `${HOJA_PRECIOS}!`;
  const cierres = `IFERROR(FILTER({${r['Modelo PEDIDO']},${r['Descripción de producto']}},` +
    `${r['Mes cierre']}=${M},${r.Pedido}="APROBADO",${r.Carpeta}="APROBADA"),{"",""})`;
  sh.getRange('J14:K14').merge().setValue('   de los cuales, por modelo').setFontColor(GRIS_TEXTO);
  sh.getRange('L14:M14').merge().setFormula(
    `=IFERROR(LET(c,${cierres},` +
    `v,ARRAYFORMULA(XLOOKUP(${M}&"|"&INDEX(c,,1),${pp}$A$2:$A&"|"&${pp}$B$2:$B,${pp}$C$2:$C,"")),` +
    `d,ARRAYFORMULA(IF((INDEX(c,,1)<>"")*(v<>""),v,INDEX(c,,2))),` +
    `TEXTJOIN("  ·  ",TRUE,MAP(FILTER(OBJETIVOS!$F$2:$F,OBJETIVOS!$E$2:$E=${M}),` +
    `FILTER(OBJETIVOS!$G$2:$G,OBJETIVOS!$E$2:$E=${M}),LAMBDA(ind,obj,LET(mod,TRIM(SUBSTITUTE(ind,"PEDIDOS ","")),` +
    `PROPER(mod)&" "&SUMPRODUCT(--ISNUMBER(SEARCH(mod,d)))&"/"&obj))))),"—")`)
    .setHorizontalAlignment('center').setFontColor(GRIS_TEXTO);
  sh.getRange('J14:M14').setBorder(null, null, true, null, null, null, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('J15:M18').merge().setWrap(true).setVerticalAlignment('top')
    .setValue('El bonus de cumplimiento exige además el 100% del objetivo de suscripciones (no se mide en ' +
      'este archivo). Los % salen de la hoja INCENTIVOS según la categoría.')
    .setFontSize(9).setFontStyle('italic').setFontColor(GRIS_TEXTO);
  ['I16', 'M10'].forEach((c) => sh.getRange(c).setNumberFormat('0%'));
  ['M11', 'M13'].forEach((c) => sh.getRange(c).setNumberFormat('0.00%'));
  sh.getRange('I17').setNumberFormat('$ #,##0');
  sh.setRowHeight(7, 10);
  sh.setRowHeight(19, 14);

  // --- Tarjetas (filas 4-6): lo más importante, a la vista.
  const tarjetas = [
    ['B', 'C', 'Adjudicados del acto', '=E9', '="Licitación "&E10&"   ·   Sorteo "&E15', '0'],
    ['D', 'E', 'Aprobados del acto', `=COUNTIFS(${acto},${aprob})`,
      '=IF(N(B5)=0,"","Conversión "&TEXT(D5/B5,"0%"))', '0'],
    ['F', 'G', 'Netos del mes', '=I9', '="Del acto "&I10&"   ·   Anteriores "&I11', '0'],
    ['H', 'I', 'Cumplimiento objetivo', '=I16',
      '=IF(I13="","Sin objetivo cargado","Objetivo "&I13&"   ·   Categoría "&M9)', '0%'],
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

  // Semáforo del cumplimiento (tarjeta y bloques).
  const semaforo = [sh.getRange('H5'), sh.getRange('I16'), sh.getRange('M10')];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1)
      .setFontColor('#38761d').setRanges(semaforo).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.9, 0.9999)
      .setFontColor('#b45f06').setRanges(semaforo).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.9)
      .setFontColor('#cc0000').setRanges(semaforo).build(),
  ]);

  // --- Cuadros: POR FAMILIA y POR MODELO PEDIDO de CLIENTES (sin TP) y POR MODELO PEDIDO de TP, cada uno
  // con su PROYECCION al lado. TP = "TP" en la columna Responsable de BASE. Solo aparecen los modelos con
  // pedidos (netos del mes o, en el mes en curso, en proyección): las filas sobrantes se ocultan
  // (ajustarFilasResumen).
  const enc = ['Modelo', 'Familia', 'Netos del mes', 'V.M', 'VM netos', 'Bonus pedidos', 'Inc pedido',
    'Total incentivos', 'Descuentos', 'Total'];
  const titulo = (fila, texto, encabezados) => {
    const n = encabezados.length;
    sh.getRange(fila, 2, 1, n).merge().setValue(texto)
      .setFontWeight('bold').setBackground(COLOR_ENCABEZADO).setFontColor('white');
    sh.getRange(fila + 1, 2, 1, n).setValues([encabezados]).setFontWeight('bold').setFontSize(9)
      .setFontColor(GRIS_TEXTO).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(null, null, true, null, null, null, COLOR_ENCABEZADO, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(fila + 1, 2).setHorizontalAlignment('left');
    sh.setRowHeight(fila + 1, 32);
  };
  // ancho = cantidad de columnas desde B; pesos / porcentajes = números de columna con ese formato.
  const formatoTabla = (f0, f1, ancho, pesos, porcentajes) => {
    const filas = f1 - f0 + 1;
    sh.getRange(f0, 2, filas, ancho).setHorizontalAlignment('center');
    sh.getRange(f0, 2, filas, 1).setHorizontalAlignment('left');
    sh.getRange(f0 + 1, 2, f1 - f0, ancho)
      .setBorder(null, null, null, null, null, true, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(f0, 2, 1, ancho).setFontWeight('bold').setBackground(FONDO_TOTAL);
    pesos.forEach((c) => sh.getRange(f0, c, filas, 1).setNumberFormat('$ #,##0').setHorizontalAlignment('right'));
    porcentajes.forEach((c) => sh.getRange(f0, c, filas, 1).setNumberFormat('0.00%'));
    sh.getRange(f0, 4, filas, 1).setNumberFormat('#,##0'); // Netos: cantidad, no $
  };

  const mp = formulaModeloCalculo(r);
  const cierreMes = `(${r['Mes cierre']}=${M})*(${r.Pedido}="APROBADO")*(${r.Carpeta}="APROBADA")`;
  const aj = `${HOJA_AJUSTES}!`;
  const pr = `${HOJA_PRECIOS}!`;
  const esTP = `(UPPER(TRIM(${r.Responsable}))="TP")`;
  // PROYECCION: operaciones con Modelo PEDIDO cargado que todavía no son neto (de cualquier acto). Solo se
  // muestra cuando el mes elegido es el mes en curso; esos modelos se suman a la lista.
  const nombresMes = MESES.map((x) => `"${x}"`).join(',');
  const mesEnCurso = `(${M}=CHOOSE(MONTH(TODAY()),${nombresMes})&" "&RIGHT(YEAR(TODAY()),2))`;
  const mped = r['Modelo PEDIDO'];
  // Pendiente = Modelo PEDIDO cargado, todavía no es neto (falta Pedido APROBADO o Carpeta APROBADA)
  // y no es BAJA ADJ. De cualquier acto.
  const pendiente = `(${mped}<>"")*(1-(${r.Pedido}="APROBADO")*(${r.Carpeta}="APROBADA"))*(${r.Pedido}<>"BAJA ADJ")`;
  const fam = `${pr}$${COL_FAMILIA}$2:$${COL_FAMILIA}`;

  // PROYECCION al lado de cada cuadro: Cantidad, Descuentos ($ descuento), Incentivos / bonus (Prec CC
  // s/IVA del mes × (Bonus pedidos + Inc pedido)) e Incentivos esperados (incentivos - descuentos).
  const proyeccion = (fTitulo, fTotal, fUlt, col, texto) => {
    sh.getRange(fTitulo, col, 1, 4).merge().setValue(texto).setFontWeight('bold')
      .setBackground(COLOR_ENCABEZADO).setFontColor('white').setHorizontalAlignment('center').setWrap(true)
      .setVerticalAlignment('middle');
    const encP = sh.getRange(fTitulo + 1, col, 1, 4);
    encP.setValues([['Cantidad', 'Descuentos', 'Incentivos / bonus', 'Incentivos esperados']]).setFontWeight('bold')
      .setFontSize(9).setFontColor(GRIS_TEXTO).setWrap(true).setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(null, null, true, null, null, null, COLOR_ENCABEZADO, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(fTitulo + 1, col).setNote('Operaciones de BASE (de cualquier acto) con Modelo PEDIDO cargado que ' +
      'todavía no son neto (falta Pedido APROBADO o Carpeta APROBADA), sin BAJA ADJ. Solo en el mes en curso.');
    sh.getRange(fTitulo + 1, col + 2).setNote('Prec CC s/IVA del mes × (Bonus pedidos + Inc pedido) de cada ' +
      'operación proyectada.');
    sh.getRange(fTitulo + 1, col + 3).setNote('Incentivos / bonus - Descuentos.');
    const filas = fUlt - fTotal + 1;
    sh.getRange(fTotal, col, filas, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    sh.getRange(fTotal, col + 1, filas, 3).setNumberFormat('$ #,##0').setHorizontalAlignment('right');
    sh.getRange(fTotal, col, 1, 4).setFontWeight('bold').setBackground(FONDO_TOTAL);
    sh.getRange(fTotal + 1, col, fUlt - fTotal, 4)
      .setBorder(null, null, null, null, null, true, GRIS_LINEA, SpreadsheetApp.BorderStyle.SOLID);
    // Totales (vacíos fuera del mes en curso).
    sh.getRange(fTotal, col, 1, 4).setFormulas([[0, 1, 2, 3].map((k) => {
      const l = letra(col + k);
      return `=IF(COUNT(${l}${fTotal + 1}:${l}${fUlt})=0,"",SUM(${l}${fTotal + 1}:${l}${fUlt}))`;
    })]);
  };

  // Un cuadro POR MODELO PEDIDO (B:K) con su PROYECCION (M:P). tp = true: solo TP; false: clientes (sin TP,
  // con los AJUSTES del mes).
  const cuadroModelo = (t, tp, texto, textoProy) => {
    const cond = tp ? esTP : `(1-${esTP})`;
    const f1 = t.total + 1;
    const f2 = t.ultima;
    const rg = (l) => `${l}${f1}:${l}${f2}`;
    const lista = rg('B');
    const mapa = (expr) => `=LET(mp,${mp},MAP(${lista},LAMBDA(m,IF(m="","",${expr}))))`;
    titulo(t.titulo, texto, enc);
    const h = t.titulo + 1;
    sh.getRange(h, 3).setNote('Familia del modelo en PRECIOS (columna Familia).');
    sh.getRange(h, 7).setNote('% de "Bonus cumplimiento pedidos totales" del bloque INCENTIVO PEDIDOS (M11): ' +
      'según la categoría, y 0 si no se llegó al objetivo de pedidos.');
    sh.getRange(h, 8).setNote('Inc pedido del modelo en el mes (hoja PRECIOS).');
    sh.getRange(h, 9).setNote('VM netos × (Bonus pedidos + Inc pedido).');
    sh.getRange(h, 10).setNote('Suma de "$ descuento" de BASE de los netos del mes de ese modelo.');
    sh.getRange(h, 11).setNote('Total incentivos - Descuentos.');
    const ajustes = tp ? '' : `IFERROR(FILTER(${aj}B2:B,${aj}A2:A=${M},${aj}C2:C<>"",${aj}C2:C<>0),"");`;
    sh.getRange(f1, 2).setFormula(
      `=IFERROR(LET(mp,${mp},u,UNIQUE({IFERROR(FILTER(mp,mp<>"",${cierreMes}*${cond}>0),"");${ajustes}` +
      `IF(${mesEnCurso},IFERROR(FILTER(${mped},${pendiente}*${cond}>0),""),"")}),SORT(FILTER(u,u<>""))),"")`);
    // Familia: la del modelo en el mes; si ese mes no la tiene, la de cualquier otro mes.
    sh.getRange(f1, 3).setFormula(`=MAP(${lista},LAMBDA(m,IF(m="","",IFERROR(INDEX(FILTER(${fam},` +
      `${pr}$B$2:$B=m,${fam}<>"",${pr}$A$2:$A=${M}),1),IFERROR(INDEX(FILTER(${fam},${pr}$B$2:$B=m,${fam}<>""),1),` +
      `"Sin familia")))))`);
    const ajNetos = tp ? '' : `+SUMIFS(${aj}$C:$C,${aj}$A:$A,${M},${aj}$B:$B,m)`;
    sh.getRange(f1, 4).setFormula(mapa(`SUMPRODUCT(${cierreMes}*${cond}*(mp=m))${ajNetos}`));
    sh.getRange(f1, 5).setFormula(mapa(`SUMIFS(${pr}$${COL_PRECIO_VM}:$${COL_PRECIO_VM},${pr}$A:$A,${M},${pr}$B:$B,m)`));
    // VM netos = netos del mes × Prec CC s/IVA del mes en que se liquidan (+ VM extra de AJUSTES en clientes).
    const ajVM = tp ? '' : `+SUMIFS(${aj}$D:$D,${aj}$A:$A,${M},${aj}$B:$B,m)`;
    sh.getRange(f1, 6).setFormula(`=MAP(${lista},${rg('D')},${rg('E')},LAMBDA(m,n,p,IF(m="","",n*p${ajVM})))`);
    sh.getRange(f1, 7).setFormula(`=MAP(${lista},LAMBDA(m,IF(m="","",N($M$11))))`);
    sh.getRange(f1, 8).setFormula(mapa(`SUMIFS(${pr}$D:$D,${pr}$A:$A,${M},${pr}$B:$B,m)`));
    sh.getRange(f1, 9).setFormula(`=MAP(${lista},${rg('F')},${rg('G')},${rg('H')},LAMBDA(m,v,b,i,IF(m="","",v*(b+i))))`);
    sh.getRange(f1, 10).setFormula(`=LET(mp,${mp},d,ARRAYFORMULA(IFERROR(--${r['$ descuento']},0)),` +
      `MAP(${lista},LAMBDA(m,IF(m="","",SUMPRODUCT(${cierreMes}*${cond}*(mp=m)*d)))))`);
    sh.getRange(f1, 11).setFormula(`=MAP(${lista},${rg('I')},${rg('J')},LAMBDA(m,t,b,IF(m="","",t-b)))`);
    sh.getRange(t.total, 2, 1, 10).setValues([['TOTAL', '', `=SUM(${rg('D')})`, '', `=SUM(${rg('F')})`, '', '',
      `=SUM(${rg('I')})`, `=SUM(${rg('J')})`, `=SUM(${rg('K')})`]]);
    formatoTabla(t.total, f2, 10, [5, 6, 9, 10, 11], [7, 8]);

    // Proyección (M:P).
    const fuera = `OR(m="",NOT(${mesEnCurso}))`;
    sh.getRange(f1, 13).setFormula(`=MAP(${lista},LAMBDA(m,IF(${fuera},"",` +
      `SUMPRODUCT((${mped}=m)*${pendiente}*${cond}))))`);
    sh.getRange(f1, 14).setFormula(`=LET(d,ARRAYFORMULA(IFERROR(--${r['$ descuento']},0)),MAP(${lista},LAMBDA(m,` +
      `IF(${fuera},"",SUMPRODUCT((${mped}=m)*${pendiente}*${cond}*d)))))`);
    sh.getRange(f1, 15).setFormula(`=MAP(${lista},${rg('M')},LAMBDA(m,q,IF(q="","",q*` +
      `SUMIFS(${pr}$${COL_PRECIO_VM}:$${COL_PRECIO_VM},${pr}$A:$A,${M},${pr}$B:$B,m)*` +
      `(N($M$11)+SUMIFS(${pr}$D:$D,${pr}$A:$A,${M},${pr}$B:$B,m)))))`);
    sh.getRange(f1, 16).setFormula(`=MAP(${lista},${rg('N')},${rg('O')},LAMBDA(m,d,i,IF(OR(m="",AND(d="",i="")),"",` +
      'N(i)-N(d))))');
    proyeccion(t.titulo, t.total, f2, 13, textoProy);
  };

  // --- Por familia (clientes): suma el cuadro por modelo de clientes según su familia. Solo importes.
  const c = T.clientes;
  const fa = T.familia;
  const encFam = ['Familia', 'Modelos', 'Netos del mes', 'VM netos', 'Total incentivos', 'Descuentos',
    'Total'];
  titulo(fa.titulo, 'POR FAMILIA · CLIENTES (sin TP)', encFam);
  const fa1 = fa.total + 1;
  const lf = `B${fa1}:B${fa.ultima}`;
  const cm = (l) => `$${l}$${c.total + 1}:$${l}$${c.ultima}`;
  const suma = (col) => `=MAP(${lf},LAMBDA(f,IF(f="","",SUMIF(${cm('C')},f,${cm(col)}))))`;
  sh.getRange(fa1, 2).setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${cm('C')},${cm('C')}<>""))),"")`);
  sh.getRange(fa1, 3).setFormula(`=MAP(${lf},LAMBDA(f,IF(f="","",TEXTJOIN(" · ",TRUE,FILTER(${cm('B')},${cm('C')}=f)))))`);
  [['D', 4], ['F', 5], ['I', 6], ['J', 7], ['K', 8]].forEach(([origen, col]) => {
    sh.getRange(fa1, col).setFormula(suma(origen));
  });
  sh.getRange(fa.total, 2, 1, 7).setValues([['TOTAL', '', ...['D', 'E', 'F', 'G', 'H'].map((l) =>
    `=SUM(${l}${fa1}:${l}${fa.ultima})`)]]);
  formatoTabla(fa.total, fa.ultima, 7, [5, 6, 7, 8], []);
  sh.getRange(fa1, 3, fa.ultima - fa1 + 1, 1).setHorizontalAlignment('left').setFontSize(9).setFontColor(GRIS_TEXTO);
  // Proyección por familia (J:M): suma la proyección de sus modelos (clientes).
  ['M', 'N', 'O', 'P'].forEach((origen, k) => {
    sh.getRange(fa1, 10 + k).setFormula(`=MAP(${lf},LAMBDA(f,IF(OR(f="",COUNT(${cm(origen)})=0),"",` +
      `SUMIF(${cm('C')},f,${cm(origen)}))))`);
  });
  proyeccion(fa.titulo, fa.total, fa.ultima, 10, 'PROYECCION CLIENTES (sin carpeta aprobada)');

  cuadroModelo(c, false, 'POR MODELO PEDIDO · CLIENTES (sin TP)   (V.M. = Precio CC sin iva)',
    'PROYECCION CLIENTES (sin carpeta aprobada)');
  cuadroModelo(T.tp, true, 'POR MODELO PEDIDO · TP   (V.M. = Precio CC sin iva)',
    'PROYECCION TP (sin carpeta aprobada)');

  // Control: clientes + TP = netos del mes.
  sh.getRange(T.control, 2, 1, 10).merge()
    .setFormula(`="Control: netos clientes "&D${c.total}&"  +  netos TP "&D${T.tp.total}&"  =  "&` +
      `(D${c.total}+D${T.tp.total})&"   ·   Netos del mes: "&I9&IF(D${c.total}+D${T.tp.total}=I9,"   ✔","   ✖ revisar")`)
    .setFontSize(9).setFontColor(GRIS_TEXTO).setHorizontalAlignment('left');
  [fa.titulo - 1, c.titulo - 1, T.tp.titulo - 1, T.control - 1].forEach((f) => sh.setRowHeight(f, 14));
  alinearTodo(sh);
  sh.getRange(T.control, 2).setHorizontalAlignment('left');
  SpreadsheetApp.flush();
  ajustarFilasResumen(sh);
}

/**
 * Filas de los cuadros del RESUMEN (familia, modelo clientes y modelo TP): título, TOTAL, primera y
 * última fila de datos reservada. Las que quedan vacías se ocultan.
 */
function tablasResumen() {
  const familia = { titulo: 20, total: 22, ultima: 33 };
  const clientes = { titulo: 35, total: 37, ultima: 76 };
  const tp = { titulo: 78, total: 80, ultima: 109 };
  return { familia, clientes, tp, control: 111 };
}

/** Oculta las filas vacías de los cuadros del RESUMEN y muestra las que tienen modelo (o familia). */
function ajustarFilasResumen(sh) {
  sh = sh || SpreadsheetApp.getActive().getSheetByName(HOJA_RESUMEN);
  if (!sh) return;
  const T = tablasResumen();
  [T.familia, T.clientes, T.tp].forEach((t) => {
    const f1 = t.total + 1;
    const n = t.ultima - f1 + 1;
    const vals = sh.getRange(f1, 2, n, 1).getDisplayValues();
    let k = 0;
    while (k < n && vals[k][0] !== '') k++;
    const visibles = Math.max(1, k);
    sh.showRows(f1, n);
    if (visibles < n) sh.hideRows(f1 + visibles, n - visibles);
  });
}

/** RESUMEN e INFORME: todo centrado (horizontal y vertical) y con ajuste de texto. */
function alinearTodo(sh) {
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setWrap(true);
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
  sh.setColumnWidth(2, 110);
  for (let c = 3; c <= 11; c++) sh.setColumnWidth(c, 92);
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
      // VM netos: por cada modelo con Prec CC s/IVA en el mes, precio × netos de ese modelo (+ AJUSTES).
      // (Un SUMIFS con criterio de matriz dentro de SUMPRODUCT usaba un solo precio para todo.)
      `=IFERROR(LET(p,FILTER({PRECIOS!$B$2:$B,PRECIOS!$${COL_PRECIO_VM}$2:$${COL_PRECIO_VM}},` +
        `PRECIOS!$A$2:$A=${A},PRECIOS!$${COL_PRECIO_VM}$2:$${COL_PRECIO_VM}<>""),mp,${formulaModeloCalculo(r)},SUM(MAP(INDEX(p,,1),INDEX(p,,2),` +
        `LAMBDA(m,pr,pr*(SUMPRODUCT((${r['Mes cierre']}=${A})*(mp=m)*(${r.Pedido}="APROBADO")*(${r.Carpeta}="APROBADA"))+` +
        `SUMIFS(${HOJA_AJUSTES}!$C:$C,${HOJA_AJUSTES}!$A:$A,${A},${HOJA_AJUSTES}!$B:$B,m)))))),0)` +
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

    // Meses sin Prec CC s/IVA: VM fijo calculado con el precio cliente con IVA, aclarado al costado.
    const conPrecioCC = mesesConPrecioCC(ss);
    meses.forEach((mes, i) => {
      if (!(mes in VM_INFORME_A_CORREGIR) || conPrecioCC.has(mes)) return;
      sh.getRange(f1 + i, 11).setValue(VM_INFORME_A_CORREGIR[mes]);
      sh.getRange(f1 + i, 12).setValue(NOTA_VM_A_CORREGIR)
        .setFontSize(9).setFontStyle('italic').setFontColor('#b45f06');
    });
    sh.setColumnWidth(12, 200);

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
  alinearTodo(sh);
}

/** Meses que tienen al menos un Prec CC s/IVA cargado en PRECIOS. */
function mesesConPrecioCC(ss) {
  const sh = ss.getSheetByName(HOJA_PRECIOS);
  const set = new Set();
  if (!sh || sh.getLastRow() < 2) return set;
  const col = COL_PRECIO_VM.charCodeAt(0) - 64;
  const meses = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const precios = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  meses.forEach((f, i) => { if (precios[i][0] !== '' && precios[i][0] !== null) set.add(textoMes(f[0])); });
  return set;
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
  // Recarga del mismo acto (p. ej. operaciones sumadas a mitad de mes): solo se agregan las
  // solicitudes que no estaban; las ya cargadas no se tocan aunque en SGA hayan cambiado.
  const recarga = valoresColumnaBase(ss, 'Acto').indexOf(acto) >= 0;
  const { agregadas, repetidas, nuevas, fila0 } = agregarABase(ss, filas);
  if (recarga && agregadas) {
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    SpreadsheetApp.getActive().getSheetByName(HOJA_BASE)
      .getRange(fila0, colBase('Solicitud'), agregadas, 1)
      .setNotes(nuevas.map(() => [`Agregada en la recarga del acto ${acto} del ${hoy}`]));
  }
  pegar.clear();
  ordenarBase(true);
  actualizarInforme();
  SpreadsheetApp.flush();
  ajustarFilasResumen();

  const iSol = iConfig('Solicitud');
  const iNom = iConfig('Apellido, Nombre');
  const lista = nuevas.slice(0, 15).map((f) => `• ${f[iSol]} - ${f[iNom]}`).join('\n') +
    (nuevas.length > 15 ? `\n… y ${nuevas.length - 15} más` : '');
  ui.alert(recarga
    ? `Recarga del acto ${acto}: ${agregadas} operación(es) nueva(s) agregada(s); ` +
      `${repetidas} ya estaban y no se modificaron.` + (agregadas ? `\n\n${lista}` : '')
    : `Acto ${acto}: ${agregadas} filas agregadas a BASE` +
      (repetidas ? `, ${repetidas} ya estaban cargadas (no se duplicaron).` : '.'));
}

/** Copia una sola vez los archivos mensuales viejos (ver HISTORICOS). */
function importarHistorico() {
  const ss = SpreadsheetApp.getActive();
  const resumen = HISTORICOS.map((h) => importarArchivoMensual(ss, h));
  actualizarInforme();
  SpreadsheetApp.getUi().alert(resumen.join('\n'));
}

/** Copia a BASE la hoja de un archivo mensual viejo ({id, hoja, acto}). No duplica. */
function importarArchivoMensual(ss, h) {
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
  if (agregadas) ordenarBase(true);
  return `${h.acto}: ${agregadas} agregadas` + (repetidas ? `, ${repetidas} ya estaban` : '');
}

/**
 * Agrega filas a BASE salteando las que ya existen (mismo Acto + Solicitud). Las filas vienen en
 * el orden de COLUMNAS_BASE y se escriben en la columna que tenga cada nombre en BASE; las
 * columnas que agregó el equipo (p. ej. "Modelo PEDIDO") no se tocan.
 */
function agregarABase(ss, filas) {
  const base = ss.getSheetByName(HOJA_BASE);
  const clave = (acto, sol) => `${textoMes(acto)}|${sol}`;
  const existentes = new Set();
  const ultima = ultimaFilaBase(base);
  if (ultima > 1) {
    const actos = base.getRange(2, colBase('Acto'), ultima - 1, 1).getValues();
    const sols = base.getRange(2, colBase('Solicitud'), ultima - 1, 1).getValues();
    actos.forEach((f, i) => existentes.add(clave(f[0], sols[i][0])));
  }
  const iActo = iConfig('Acto');
  const iSol = iConfig('Solicitud');
  const nuevas = filas.filter((f) => {
    const k = clave(f[iActo], f[iSol]);
    if (existentes.has(k)) return false;
    existentes.add(k);
    return true;
  });
  if (nuevas.length) escribirFilasBase(base, ultima + 1, nuevas);
  return { agregadas: nuevas.length, repetidas: filas.length - nuevas.length, nuevas, fila0: ultima + 1 };
}

/** Escribe filas (orden de COLUMNAS_BASE) en BASE, por tramos de columnas seguidas. */
function escribirFilasBase(base, fila0, filas) {
  const cols = COLUMNAS_BASE.map((c, i) => ({ col: colBase(c.h), i })).sort((a, b) => a.col - b.col);
  for (let k = 0; k < cols.length;) {
    let j = k;
    while (j + 1 < cols.length && cols[j + 1].col === cols[j].col + 1) j++;
    const tramo = cols.slice(k, j + 1);
    base.getRange(fila0, tramo[0].col, filas.length, tramo.length)
      .setValues(filas.map((f) => tramo.map((t) => f[t.i])));
    k = j + 1;
  }
}

/**
 * Ordena BASE por Acto en orden cronológico (JULIO 26, AGOSTO 26, SEPTIEMBRE 26...). Dentro de
 * cada acto mantiene el orden que ya tenían las filas. Mueve también las notas (p. ej. las de
 * las recargas). Se corre sola al agregar un acto o una recarga; también está en el menú.
 */
function ordenarBase(silencioso) {
  const ss = SpreadsheetApp.getActive();
  const base = ss.getSheetByName(HOJA_BASE);
  const ultima = ultimaFilaBase(base);
  if (ultima < 3) return;
  const nFilas = ultima - 1;
  const actos = base.getRange(2, colBase('Acto'), nFilas, 1).getValues();
  const orden = (v) => {
    const m = textoMes(v).match(/^([A-Z]+) (\d{2})$/);
    const i = m ? MESES.indexOf(m[1]) : -1;
    return i < 0 ? 99999 : Number(m[2]) * 12 + i;
  };
  // Clave: acto cronológico y, dentro del acto, el orden que ya tenían las filas.
  const claves = actos.map((f, i) => [orden(f[0]) * 100000 + i]);
  if (claves.every((k, i) => i === 0 || k[0] > claves[i - 1][0])) {
    if (!silencioso) SpreadsheetApp.getUi().alert('BASE ya estaba ordenada por acto.');
    return;
  }
  // Se ordena con Range.sort usando una columna auxiliar temporal: así se mueven las filas
  // enteras (valores, casillas, colores, notas) y no se pierde nada de lo que cargó el equipo.
  const colAux = base.getMaxColumns() + 1;
  base.insertColumnAfter(base.getMaxColumns());
  base.getRange(2, colAux, nFilas, 1).setValues(claves);
  base.getRange(2, 1, nFilas, colAux).sort({ column: colAux, ascending: true });
  base.deleteColumn(colAux);
  if (!silencioso) SpreadsheetApp.getUi().alert('BASE ordenada por acto (del más viejo al más nuevo).');
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

/**
 * Número de columna de BASE según el nombre del encabezado (fila 1), no según la posición: el
 * equipo puede agregar, mover o renombrar columnas (ALIAS_BASE) sin romper el script.
 */
function colBase(nombre) {
  if (iConfig(nombre) < 0) throw new Error(`Columna inexistente en BASE: ${nombre}`);
  const c = mapaBase().col[nombre];
  if (!c) {
    throw new Error(`En BASE no encuentro la columna "${nombre}" (¿se borró o se renombró?). ` +
      'Volvé a ponerle ese nombre en la fila 1.');
  }
  return c;
}

/** Posición de la columna en COLUMNAS_BASE (para las filas que arma el script). */
function iConfig(nombre) {
  return COLUMNAS_BASE.findIndex((c) => c.h === nombre);
}

/**
 * Lee los encabezados de BASE una vez por ejecución. Busca cada columna por su nombre, por
 * sus alias y, si no, sin distinguir mayúsculas ni acentos (solo si no es el nombre exacto de
 * otra columna: "Responsable" y "RESPONSABLE" son distintas). Si BASE todavía no existe o no
 * tiene encabezados, usa el orden de COLUMNAS_BASE.
 */
function mapaBase() {
  if (MAPA_BASE) return MAPA_BASE;
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_BASE);
  const enc = sh && sh.getLastColumn() > 0
    ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((v) => String(v).trim())
    : [];
  const col = {};
  if (!enc.some((h) => h !== '')) {
    COLUMNAS_BASE.forEach((c, i) => { col[c.h] = i + 1; });
    MAPA_BASE = { col, faltan: [] };
    return MAPA_BASE;
  }
  const usadas = new Set();
  COLUMNAS_BASE.forEach((c) => {
    const i = [c.h].concat(ALIAS_BASE[c.h] || []).map((n) => enc.indexOf(n)).find((x) => x >= 0);
    if (i !== undefined) { col[c.h] = i + 1; usadas.add(i); }
  });
  const simple = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  COLUMNAS_BASE.forEach((c) => {
    if (col[c.h]) return;
    const i = enc.findIndex((h, k) => !usadas.has(k) && h && simple(h) === simple(c.h) &&
      iConfig(h) < 0);
    if (i >= 0) { col[c.h] = i + 1; usadas.add(i); }
  });
  MAPA_BASE = { col, faltan: COLUMNAS_BASE.filter((c) => !col[c.h]).map((c) => c.h) };
  return MAPA_BASE;
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

/**
 * Modelo con el que se calcula (por fila de BASE, como matriz): Modelo PEDIDO; si está vacío,
 * Modelo ahorro (filas viejas o todavía sin cargar).
 */
function formulaModeloCalculo(r) {
  return `ARRAYFORMULA(IF(${r['Modelo PEDIDO']}<>"",${r['Modelo PEDIDO']},${r['Modelo ahorro']}))`;
}

/** Rangos abiertos de BASE para usar en fórmulas: r['Pedido'] = BASE!$U$2:$U */
function rangosBase() {
  const r = {};
  COLUMNAS_BASE.forEach((c) => {
    const l = letra(colBase(c.h));
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
  if (columna === 'Llave x llave') return true; // casilla; en los archivos viejos aparece "x", "1" o "SI"
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

// ---------------------------------------------------------------------------
// Informe de incentivos (menú Pedidos > Generar informe de incentivos)
// ---------------------------------------------------------------------------
//
// Escribe el período en un único Sheet de informes ("Incentivos pedidos FIAT", ID_INFORME) que va
// acumulando los meses: al generar un mes que ya estaba, se reemplazan sus filas. Lo comparte y manda
// los totales por mail. Reglas (señal 1599 y liquidaciones de Fiat de septiembre 2026):
// - 2A Bonus cumplimiento pedidos totales: % según categoría (INCENTIVOS) × precio público CON IVA
//   del modelo PEDIDO. Solo si netos >= objetivo de pedidos y se cumplió suscripciones.
// - 2B Pedidos adicionales: % según categoría × precio CE sin IVA de NC1, por cada pedido sobre el
//   objetivo (solo si se cumplió el objetivo de pedidos).
// - 2C Extra bonus pedidos: Inc pedido (PRECIOS) del modelo PEDIDO × precio CE sin IVA del modelo
//   base de la familia del pedido (MODELO_BASE_FAMILIA: Cronos = NC1, Titano = DT1...), que es como
//   liquida Fiat; en los meses de MESES_2C_MODELO_PEDIDO, CE sin IVA del propio modelo pedido. Solo si se
//   cumplió patentamientos. Excluye Cronos 90/10 (B90) en el rango de RANGO_B90.
// - Adelanto de margen: precio público SIN IVA del modelo PEDIDO × 14% × 60%. Cruce con las
//   entregas (hoja SGA del archivo de entregas): si el REU (Fecha Ingreso de SGA) es del mismo mes,
//   Fiat no adelanta (paga el 100% del margen con la entrega). Tampoco adelanta las operaciones de
//   otro concesionario.
// - Margen por entrega (40%): unidades con REU en el mes (Fecha Ingreso de SGA). Precio público SIN IVA
//   del modelo entregado (FBS -> MODELOS FBS), lista del mes de la "Fecha factura" de SGA × 14% × 40%
//   (verificado con el detalle de Fiat de septiembre 2026: la fecha que usa Fiat es la de factura).

const ID_INFORME = '14wzzN13YYjYjneehyeZHYQLzFsDlfIXHS2LXo1CX98I';
const NOMBRE_INFORME = 'Incentivos pedidos FIAT';
const ID_ENTREGAS = '1lWVidE-2smtBoAEWC5Z_kEbHXXJ3cpoBGArL5YE96Mw';
const HOJA_SGA = 'SGA';
const CONCESIONARIO_PROPIO = 'TURIN S.A.';
const CARPETA_INFORMES = 'Informes de incentivos';
const MARGEN_COMISIONAL = 0.14;
const ADELANTO_MARGEN = 0.6;
// Cronos 90/10 (B90) sin 2C: grupos desde / hasta según la señal de cada mes (sin mes = valor por defecto).
// Septiembre 2026 (señal 1599): desde 18002, sin tope. Octubre 2026 (señal 1606): 18002 a 18115 inclusive.
const RANGO_B90 = {
  '': [18002, Infinity],
  'OCTUBRE 26': [18002, 18115],
};
const FAMILIAS_IVA_REDUCIDO = ['FIORINO', 'STRADA', 'TORO', 'TITANO'];
// Meses en que Fiat liquida el 2C sobre el precio CE s/IVA del propio modelo pedido (no del modelo base
// de la familia). Octubre 2026: señal 1606, "por el modelo solicitado en cada pedido aceptado".
const MESES_2C_MODELO_PEDIDO = ['OCTUBRE 26'];
// Modelo de ahorro de referencia de cada familia: el 2C se paga a su precio CE sin IVA.
// Verificado con la liquidación de Fiat de septiembre 2026 (Cronos, Titano, Toro, Fiorino) y con el
// archivo de agosto (Strada, Fastback, Pulse, Argo). Si la familia no está, se usa el modelo de ahorro.
const MODELO_BASE_FAMILIA = {
  CRONOS: 'NC1', TITANO: 'DT1', TORO: 'NT3', FIORINO: 'FO1', STRADA: 'FS1', FASTBACK: 'FT3', PULSE: 'FP3',
  ARGO: 'AR2', MOBI: 'MB1',
};

// Hojas del archivo de informes.
const INF_RESUMEN = 'Resumen';
const INF_DETALLE = 'Detalle';
const INF_PERIODO = 'Por período';
const INF_ENTREGAS = 'Entregas';
const MARGEN_ENTREGA = 0.4;
// Versión entregada (texto de FBS) -> código de modelo, para el margen por entrega. Hoja de este archivo:
// las versiones nuevas se agregan solas con el código vacío.
const HOJA_MODELOS_FBS = 'MODELOS FBS';
const MODELOS_FBS_INICIALES = [
  ['ARGO DRIVE 1.3 CVT', 'AC1'], ['ARGO DRIVE 1.3 MT', 'AR2'], ['CRONOS 1.3 DRIVE CVT', 'PC5'],
  ['CRONOS 1.3 LIKE', 'LI1'], ['CRONOS DRIVE 1.3 GSE Bz', 'DP1'], ['CRONOS DRIVE PLUS', 'NC1'],
  ['CRONOS PRECISION 1.3 GSE CVT Bz', 'CA6'], ['DT1 - TITANO FREEDOM MT 4WD', 'DT1'],
  ['FASTBACK ABARTH TURBO 270 AT6', 'FA4'], ['FASTBACK TURBO 270 AT6', 'FT3'], ['FIORINO ENDURANCE 1.3L MT', 'FO1'],
  ['MOBI TREKKING 1.0 MT', 'MB1'], ['MOBI TREKKING 1.0MT5', 'MB1'], ['PULSE ABARTH TURBO 270 AT6', 'UA4'],
  ['PULSE DRIVE 1.3 CVT', 'UL3'], ['PULSE DRIVE 1.3 MT5', 'FP3'], ['STRADA FREEDOM 1.3 CD', 'FS1'],
  ['STRADA VOLCANO CD 1.3 8V CVT', 'VS1'], ['TITANO ENDURANCE MT', 'ET1'], ['TITANO ENDURANCE MT 4X4', 'TI1'],
  ['TITANO FREEDOM PLUS AT8 AWD', 'PT1'], ['TITANO RANCH AT8 AWD', 'AT1'], ['TORO FREEDOM 1.3T AT6 4X2', 'NT3'],
  ['TORO VOLCANO 1,3T AT6 4X2', 'VT3'], ['TT2 TORO VOLCANO 2,2TD AT9 4X4', 'TT1'],
];
const ENC_ENTREGAS = ['Mes (REU)', 'Grupo', 'Orden', 'Cliente', 'Fecha factura (SGA)', 'REU (Fecha ingreso SGA)',
  'Fecha entrega (SGA)', 'Versión entregada (FBS)', 'Código', 'Familia', 'Mes del precio (factura)',
  'Precio público c/IVA', '% IVA', 'Precio público s/IVA', 'Margen por entrega (14% × 40%)', 'Nota'];
const ENC_DETALLE = ['Mes', 'Acto', 'Solicitud', 'Grupo', 'Orden', 'Apellido, Nombre', 'Concesionario',
  'Concepto de venta', 'Modelo ahorro', 'Modelo pedido', 'Familia', 'Precio público c/IVA (pedido)', '% IVA',
  'Precio público s/IVA (pedido)', '% 2A cumplimiento', '2A Bonus cumplimiento', 'Modelo base 2C (familia)',
  'Precio CE s/IVA (modelo base)', '% 2C (Inc pedido)', '2C Extra bonus', 'REU (Fecha ingreso SGA)',
  'Fecha entrega (SGA)', '% adelanto (14% × 60%)', 'Adelanto margen', '$ descuento', 'Nota'];
const ENC_RESUMEN = ['Mes', 'Categoría', 'Objetivo', 'Netos', '% cumpl.', 'Adicionales', 'Suscripciones',
  'Patentamientos', '% 2A', '% 2B', 'CE s/IVA NC1', '2A Bonus cumplimiento', '2B Pedidos adicionales',
  '2C Extra bonus', 'Total incentivos', 'Descuentos', 'Total - descuentos', 'Adelanto margen',
  'Pedidos con adelanto', 'Unidades con REU', 'Margen por entrega (40%)', 'Generado', 'Avisos'];

function generarInforme() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  // 1) Período.
  const meses = mesesConDatos(ss).filter((m) => valoresColumnaBase(ss, 'Mes cierre').indexOf(m) >= 0);
  const hoy = new Date();
  const anterior = etiquetaMes((hoy.getMonth() + 11) % 12, (hoy.getFullYear() - (hoy.getMonth() === 0 ? 1 : 0)) % 100);
  const sugerido = meses.indexOf(anterior) >= 0 ? anterior : meses[meses.length - 1];
  const r1 = ui.prompt('Generar informe de incentivos',
    `Período (mes de cierre). Meses con netos: ${meses.join(', ')}.\n` +
    `Aceptar sin escribir nada = ${sugerido}.`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const mes = textoMes(r1.getResponseText()) || sugerido;
  if (meses.indexOf(mes) < 0) {
    ui.alert(`"${mes}" no tiene netos en BASE. Usá el formato MES AA, por ejemplo SEPTIEMBRE 26.`);
    return;
  }

  // 2) Mails (se recuerdan para la próxima vez).
  const mailsGuardados = props.getProperty('MAILS_INFORME') || '';
  const r2 = ui.prompt('Generar informe de incentivos',
    'Mails a los que se envía, separados por coma.' +
    (mailsGuardados ? `\nAceptar sin escribir nada = ${mailsGuardados}` : ''), ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const mails = (r2.getResponseText().trim() || mailsGuardados).split(/[,;\s]+/)
    .map((m) => m.trim()).filter((m) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m));
  if (!mails.length) {
    ui.alert('No hay ningún mail válido.');
    return;
  }
  props.setProperty('MAILS_INFORME', mails.join(', '));

  // 3) Condiciones que este archivo no mide.
  const suscripciones = ui.alert('Generar informe de incentivos',
    `¿Se cumplió el 100% del objetivo de SUSCRIPCIONES en ${mes}?\n(Si no, el bonus 2A va en $0.)`,
    ui.ButtonSet.YES_NO) === ui.Button.YES;
  const patentamientos = ui.alert('Generar informe de incentivos',
    `¿Se cumplió el 100% del objetivo de PATENTAMIENTOS en ${mes}?\n(Si no, el extra bonus 2C va en $0.)`,
    ui.ButtonSet.YES_NO) === ui.Button.YES;

  const datos = calcularIncentivos(ss, mes, suscripciones, patentamientos);
  const informe = abrirArchivoInforme(props);
  escribirInforme(informe, datos);
  enviarInforme(informe, datos, mails);
  ui.alert(`Informe de ${mes} guardado en "${informe.getName()}" y enviado a ${mails.join(', ')}.\n\n` +
    `${informe.getUrl()}` + (datos.avisos.length ? `\n\nAtención:\n• ${datos.avisos.join('\n• ')}` : ''));
}

/** Calcula el detalle (una fila por neto de BASE) y los totales del período. */
function calcularIncentivos(ss, mes, suscripciones, patentamientos) {
  const avisos = [];
  // PRECIOS del mes y familia de cada modelo.
  const pr = ss.getSheetByName(HOJA_PRECIOS).getDataRange().getValues().slice(1);
  const precio = {};
  const familia = {};
  const pubPorMes = {};
  pr.forEach((f) => {
    const m = textoMes(f[0]);
    const mod = String(f[1]).trim();
    if (!mod) return;
    if (f[7] && (!familia[mod] || m === mes)) familia[mod] = String(f[7]).trim();
    if (m === mes) precio[mod] = { inc: f[3], pub: f[4], ceIva: f[5], ceSin: f[6] };
    if (m) (pubPorMes[m] = pubPorMes[m] || {})[mod] = f[4];
  });
  const num = (v) => (typeof v === 'number' ? v : Number(String(v).replace(/[^\d.-]/g, '')) || 0);
  const ivaDe = (mod) => {
    const p = precio[mod];
    if (p && num(p.ceIva) && num(p.ceSin)) return Math.round((num(p.ceIva) / num(p.ceSin) - 1) * 1000) / 1000;
    return FAMILIAS_IVA_REDUCIDO.indexOf(String(familia[mod] || '').toUpperCase()) >= 0 ? 0.105 : 0.21;
  };

  // Objetivo, categoría e incentivos según categoría.
  const obj = ss.getSheetByName(HOJA_OBJETIVOS).getRange('A2:C').getValues()
    .find((f) => textoMes(f[0]) === mes) || [];
  const objetivo = num(obj[1]);
  const categoria = String(obj[2] || '').trim().toUpperCase();
  if (!objetivo) avisos.push(`${mes} no tiene objetivo de pedidos cargado: 2A y 2B en $0.`);
  const inc = ss.getSheetByName(HOJA_INCENTIVOS).getRange('A2:E').getValues()
    .filter((f) => textoMes(f[0]) === mes);
  const pctCategoria = (patron) => {
    const f = inc.find((x) => patron.test(String(x[1]).toUpperCase()));
    const i = ['A', 'B', 'C'].indexOf(categoria);
    return f && i >= 0 ? num(f[2 + i]) : 0;
  };

  // Netos del mes en BASE.
  const base = ss.getSheetByName(HOJA_BASE);
  const ultima = ultimaFilaBase(base);
  const vals = base.getRange(2, 1, ultima - 1, base.getLastColumn()).getValues();
  const c = (n) => colBase(n) - 1;
  const netos = vals.filter((f) => textoMes(f[c('Mes cierre')]) === mes &&
    f[c('Pedido')] === 'APROBADO' && f[c('Carpeta')] === 'APROBADA');
  const ajustes = ss.getSheetByName(HOJA_AJUSTES).getRange('A2:C').getValues()
    .filter((f) => textoMes(f[0]) === mes && num(f[2]));
  const totalNetos = netos.length + ajustes.reduce((a, f) => a + num(f[2]), 0);
  const cumplePedidos = objetivo > 0 && totalNetos >= objetivo;
  const pct2A = cumplePedidos && suscripciones ? pctCategoria(/CUMPLIMIENTO/) : 0;
  const pct2B = cumplePedidos ? pctCategoria(/ADICIONALES/) : 0;
  if (ajustes.length) {
    avisos.push(`Hay ${ajustes.length} ajuste(s) de ${mes} en AJUSTES: suman a los netos pero no al detalle ` +
      '(no tienen cliente ni modelo pedido).');
  }

  // Cruce con las entregas (SGA): REU en el mismo mes = sin adelanto.
  let reu = {};
  try {
    reu = leerEntregasSGA();
  } catch (e) {
    reu = null;
    avisos.push(`No se pudo leer la hoja SGA del archivo de entregas (${e.message}): el adelanto no descuenta ` +
      'las unidades entregadas en el mes, y no se calcula el margen por entrega.');
  }
  const sga = reu || {};

  const rangoB90 = RANGO_B90[mes] || RANGO_B90[''];
  const faltan = new Set();
  const filas = netos.map((f) => {
    const ahorro = String(f[c('Modelo ahorro')]).trim();
    const pedido = String(f[c('Modelo PEDIDO')]).trim() || ahorro;
    const pp = precio[pedido] || {};
    const fam = familia[pedido] || familia[ahorro] || '';
    const modeloBase = MESES_2C_MODELO_PEDIDO.indexOf(mes) >= 0 ? pedido :
      MODELO_BASE_FAMILIA[String(fam).toUpperCase()] || ahorro;
    const pb = precio[modeloBase] || {};
    if (!num(pp.pub)) faltan.add(`${pedido} (precio público)`);
    if (!num(pb.ceSin)) faltan.add(`${modeloBase} (precio CE sin IVA)`);
    const concepto = String(f[c('Concepto de venta')]);
    const grupo = num(f[c('Grupo')]);
    const b90 = /CRONOS/i.test(fam) && /^\s*90\s*\/\s*10/.test(concepto) && grupo >= rangoB90[0] &&
      grupo <= rangoB90[1];
    const pct2C = !patentamientos || b90 ? 0 : num(pp.inc);
    const conc = String(f[c('Concesionario')]).trim();
    const otroCE = conc !== '' && conc.toUpperCase() !== CONCESIONARIO_PROPIO;
    const ent = sga[claveGrupoOrden(f[c('Grupo')], f[c('Orden')])] || {};
    const reuMismoMes = !!ent.ingreso && etiquetaMes(ent.ingreso.getMonth(), ent.ingreso.getFullYear() % 100) === mes;
    const notas = [];
    if (b90) {
      notas.push(`Cronos 90/10 (B90), grupos ${rangoB90[0]}` +
        (rangoB90[1] === Infinity ? ' en adelante' : ` a ${rangoB90[1]}`) + ': sin 2C');
    }
    if (reuMismoMes) notas.push('REU en el mismo mes: sin adelanto (Fiat paga el 100% del margen con la entrega)');
    if (otroCE) notas.push(`Otro concesionario (${conc}): sin adelanto`);
    if (!String(f[c('Modelo PEDIDO')]).trim()) notas.push('Sin Modelo PEDIDO: se usa el de ahorro');
    if (pp.inc === '' || pp.inc === undefined) notas.push(`${pedido} sin Inc pedido en PRECIOS`);
    return {
      acto: textoMes(f[c('Acto')]), solicitud: f[c('Solicitud')], grupo: f[c('Grupo')], orden: f[c('Orden')],
      nombre: f[c('Apellido, Nombre')], conc, concepto, ahorro, pedido, familia: fam,
      pub: num(pp.pub), ivaPed: ivaDe(pedido), pct2A, modeloBase, ceSinBase: num(pb.ceSin), pct2C,
      reu: ent.ingreso || '', entrega: ent.entrega || '',
      pctAdelanto: reuMismoMes || otroCE ? 0 : MARGEN_COMISIONAL * ADELANTO_MARGEN,
      descuento: num(f[c('$ descuento')]), nota: notas.join('. '),
    };
  });
  if (faltan.size) avisos.push(`Faltan precios de ${mes} en PRECIOS: ${[...faltan].join(', ')}.`);

  // Margen por entrega (40%): unidades con REU en el mes.
  const pedidoDe = {};
  vals.forEach((f) => {
    const k = claveGrupoOrden(f[c('Grupo')], f[c('Orden')]);
    if (k) pedidoDe[k] = String(f[c('Modelo PEDIDO')]).trim() || String(f[c('Modelo ahorro')]).trim();
  });
  const entregas = reu ? calcularEntregasMes(ss, mes, sga, pubPorMes, familia, pedidoDe, avisos) : [];

  const adicionales = cumplePedidos ? totalNetos - objetivo : 0;
  const ceNC1 = num((precio.NC1 || {}).ceSin);
  return {
    mes, categoria, objetivo, totalNetos, cumplePedidos, suscripciones, patentamientos, pct2A, pct2B,
    adicionales, ceNC1, filas, entregas, avisos,
  };
}

/**
 * Unidades con REU (Fecha Ingreso de SGA) en el mes: modelo entregado (FBS -> MODELOS FBS; si no, el
 * Modelo PEDIDO de BASE) y precio público de la lista del mes de la factura (si falta, el del mes más
 * cercano, con nota).
 */
function calcularEntregasMes(ss, mes, sga, pubPorMes, familia, pedidoDe, avisos) {
  const num = (v) => (typeof v === 'number' ? v : Number(String(v).replace(/[^\d.-]/g, '')) || 0);
  const mesDe = (fecha) => (fecha ? etiquetaMes(fecha.getMonth(), fecha.getFullYear() % 100) : '');
  const normal = (t) => String(t || '').toUpperCase().replace(/\s+/g, ' ').trim();
  let fbs = {};
  try {
    fbs = leerEntregasFBS();
  } catch (e) {
    avisos.push(`No se pudo leer la hoja Base (FBS) del archivo de entregas (${e.message}): el modelo entregado ` +
      'se toma del Modelo PEDIDO de BASE.');
  }
  // Equivalencias versión FBS -> código.
  let hm = ss.getSheetByName(HOJA_MODELOS_FBS);
  if (!hm) {
    hm = ss.insertSheet(HOJA_MODELOS_FBS);
    hm.getRange(1, 1, 1, 2).setValues([['Versión entregada (FBS)', 'Código']]).setFontWeight('bold')
      .setBackground(COLOR_ENCABEZADO).setFontColor('white');
    hm.getRange(2, 1, MODELOS_FBS_INICIALES.length, 2).setValues(MODELOS_FBS_INICIALES);
    hm.setFrozenRows(1);
    hm.setColumnWidth(1, 300);
  }
  const mapa = {};
  hm.getDataRange().getValues().slice(1).forEach((f) => {
    if (normal(f[0])) mapa[normal(f[0])] = String(f[1]).trim();
  });
  const familias = [...new Set(Object.values(familia).map((x) => String(x).trim()).filter(String))];
  const famDeTexto = (t) => familias.find((fa) => normal(t).split(' ').indexOf(fa.toUpperCase()) >= 0) || '';
  const ordenes = Object.keys(pubPorMes);
  const pubDe = (m, cod) => {
    if (num((pubPorMes[m] || {})[cod])) return { mes: m, pub: num(pubPorMes[m][cod]) };
    const o = ordenMesInforme(m);
    const otro = ordenes.filter((x) => num((pubPorMes[x] || {})[cod]))
      .sort((a, b) => (Math.abs(ordenMesInforme(a) - o) - Math.abs(ordenMesInforme(b) - o)) ||
        (ordenMesInforme(a) - ordenMesInforme(b)))[0];
    return otro ? { mes: otro, pub: num(pubPorMes[otro][cod]) } : { mes: '', pub: 0 };
  };

  const nuevas = new Set();
  const sinCodigo = [];
  const sinPrecioMes = new Set();
  const filas = Object.keys(sga).filter((k) => mesDe(sga[k].ingreso) === mes).map((k) => {
    const s = sga[k];
    const fb = fbs[k] || {};
    const texto = fb.modelo || '';
    if (texto && !(normal(texto) in mapa)) nuevas.add(texto);
    const cod = mapa[normal(texto)] || pedidoDe[k] || '';
    const mesPrecio = mesDe(s.factura) || mes;
    const p = cod ? pubDe(mesPrecio, cod) : { mes: '', pub: 0 };
    const fam = familia[cod] || famDeTexto(texto);
    const notas = [];
    if (!cod) {
      sinCodigo.push(`${k} ${s.cliente}`);
      notas.push('Sin código de modelo: completarlo en MODELOS FBS');
    } else if (!mapa[normal(texto)]) {
      notas.push(texto ? `Versión "${texto}" sin código en MODELOS FBS: se usa el Modelo PEDIDO` :
        'No está en FBS: se usa el Modelo PEDIDO');
    }
    if (cod && p.mes !== mesPrecio) {
      sinPrecioMes.add(`${cod} ${mesPrecio}`);
      notas.push(p.pub ? `Sin precio de ${cod} en ${mesPrecio}: se usa el de ${p.mes}` : `Sin precio de ${cod}`);
    }
    if (!s.factura) notas.push('Sin fecha de factura en SGA: precio del mes del REU');
    const [g, o] = k.split('-');
    return [mes, Number(g), Number(o), s.cliente, s.factura || '', s.ingreso, s.entrega || '', texto, cod, fam,
      p.mes, p.pub, FAMILIAS_IVA_REDUCIDO.indexOf(String(fam).toUpperCase()) >= 0 ? 0.105 : 0.21, '', '',
      notas.join('. ')];
  }).sort((a, b) => String(a[3]).localeCompare(String(b[3])));

  if (nuevas.size) {
    hm.getRange(hm.getLastRow() + 1, 1, nuevas.size, 2).setValues([...nuevas].map((t) => [t, '']));
    avisos.push(`Versiones entregadas nuevas en ${HOJA_MODELOS_FBS} (completar el código): ${[...nuevas].join(', ')}.`);
  }
  if (sinCodigo.length) avisos.push(`Entregas sin código de modelo (margen 40% en $0): ${sinCodigo.join(', ')}.`);
  if (sinPrecioMes.size) {
    avisos.push(`Margen por entrega: faltan en PRECIOS los precios del mes de factura de ${[...sinPrecioMes].join(', ')} ` +
      '(se usó el mes más cercano).');
  }
  // Entregadas (FBS) en el mes que todavía no están en SGA: se liquidan cuando tengan REU.
  const pendientes = Object.keys(fbs).filter((k) => !sga[k] && mesDe(fbs[k].entrega) === mes)
    .map((k) => `${k} ${fbs[k].cliente}`);
  if (pendientes.length) {
    avisos.push(`Entregadas en ${mes} sin REU todavía (no están en SGA): ${pendientes.join(', ')}.`);
  }
  return filas;
}

/** Hoja SGA del archivo de entregas: grupo-orden -> {ingreso, entrega} (fechas). */
function leerEntregasSGA() {
  const vals = SpreadsheetApp.openById(ID_ENTREGAS).getSheetByName(HOJA_SGA).getDataRange().getValues();
  const iEnc = vals.findIndex((f) => f.map((v) => String(v).trim()).indexOf('Grupo') >= 0);
  if (iEnc < 0) throw new Error('no se encontró la fila de encabezados');
  const enc = vals[iEnc].map((v) => String(v).trim());
  const col = (n) => {
    const i = enc.indexOf(n);
    if (i < 0) throw new Error(`falta la columna "${n}"`);
    return i;
  };
  const g = col('Grupo');
  const o = col('Orden');
  const fi = col('Fecha Ingreso');
  const fe = col('Fecha ent.unidad');
  const ff = enc.indexOf('Fecha factura');
  const rs = enc.indexOf('Razón social');
  const out = {};
  vals.slice(iEnc + 1).forEach((f) => {
    const k = claveGrupoOrden(f[g], f[o]);
    if (k) {
      out[k] = { ingreso: aFechaInforme(f[fi]), entrega: aFechaInforme(f[fe]),
        factura: ff >= 0 ? aFechaInforme(f[ff]) : null, cliente: rs >= 0 ? String(f[rs]).trim() : '' };
    }
  });
  return out;
}

/** Hoja Base (FBS) del archivo de entregas: grupo-orden -> {modelo, entrega, cliente}. */
function leerEntregasFBS() {
  const vals = SpreadsheetApp.openById(ID_ENTREGAS).getSheetByName('Base').getDataRange().getValues();
  const enc = vals[0].map((v) => String(v).trim());
  const col = (n) => {
    const i = enc.indexOf(n);
    if (i < 0) throw new Error(`falta la columna "${n}"`);
    return i;
  };
  const g = col('Grupo');
  const o = col('Orden');
  const m = col('Modelo');
  const fe = col('FechaEntrega');
  const ap = enc.indexOf('Apellido');
  const no = enc.indexOf('Nombre');
  const out = {};
  vals.slice(1).forEach((f) => {
    const k = claveGrupoOrden(f[g], f[o]);
    if (k) {
      out[k] = { modelo: String(f[m]).replace(/\s+/g, ' ').trim(), entrega: aFechaInforme(f[fe]),
        cliente: [ap >= 0 ? String(f[ap]).trim() : '', no >= 0 ? String(f[no]).trim() : ''].filter(String).join(', ') };
    }
  });
  return out;
}

function claveGrupoOrden(g, o) {
  const a = String(g === null || g === undefined ? '' : g).trim();
  const b = String(o === null || o === undefined ? '' : o).trim();
  return /^\d+$/.test(a) && /^\d+$/.test(b) ? `${Number(a)}-${Number(b)}` : '';
}

/** Date, "28/9/2026" o "2026-09-28" -> Date (o null). */
function aFechaInforme(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function ordenMesInforme(m) {
  const x = String(m).match(/^([A-Z]+) (\d{2})$/);
  return x ? Number(x[2]) * 12 + MESES.indexOf(x[1]) : 99999;
}

/** Abre el Sheet de informes (si no existe más, crea uno nuevo en la subcarpeta de informes). */
function abrirArchivoInforme(props) {
  const id = props.getProperty('ID_INFORME') || ID_INFORME;
  try {
    const ss = SpreadsheetApp.openById(id);
    if (!DriveApp.getFileById(id).isTrashed()) return ss;
  } catch (e) { /* se crea uno nuevo */ }
  const nuevo = SpreadsheetApp.create(NOMBRE_INFORME);
  const archivo = DriveApp.getFileById(nuevo.getId());
  const padre = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  if (padre.hasNext()) {
    const carpeta = padre.next();
    const sub = carpeta.getFoldersByName(CARPETA_INFORMES);
    archivo.moveTo(sub.hasNext() ? sub.next() : carpeta.createFolder(CARPETA_INFORMES));
  }
  props.setProperty('ID_INFORME', nuevo.getId());
  return nuevo;
}

/**
 * Hoja del informe con su encabezado. Si el encabezado cambió (columnas nuevas), los meses ya generados se
 * conservan acomodándolos por nombre de columna; si es el formato viejo (sin columna Mes), se rehace vacía.
 */
function hojaInforme(inf, nombre, fila, enc) {
  let sh = inf.getSheetByName(nombre);
  if (!sh) sh = inf.insertSheet(nombre);
  const ancho = sh.getLastColumn();
  const actual = ancho && sh.getLastRow() >= fila ? sh.getRange(fila, 1, 1, ancho).getValues()[0].map(String) : [];
  if (actual.slice(0, enc.length).join('|') === enc.join('|') && actual.length >= enc.length) return sh;
  let datos = [];
  if (/^Mes/.test(actual[0] || '') && sh.getLastRow() > fila) {
    datos = sh.getRange(fila + 1, 1, sh.getLastRow() - fila, ancho).getValues()
      .filter((f) => String(f[0]).trim() !== '')
      .map((f) => enc.map((h) => (actual.indexOf(h) >= 0 ? f[actual.indexOf(h)] : '')));
  }
  sh.clear();
  sh.getRange(fila, 1, 1, enc.length).setValues([enc]).setFontWeight('bold').setBackground(COLOR_ENCABEZADO)
    .setFontColor('white').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setFrozenRows(fila);
  if (datos.length) {
    sh.getRange(fila + 1, 1, datos.length, 1).setNumberFormat('@');
    sh.getRange(fila + 1, 1, datos.length, enc.length).setValues(datos);
  }
  return sh;
}

/** Filas de datos de una hoja del informe (sin las del mes que se va a reemplazar). */
function filasSinMes(sh, fila, ancho, mes) {
  const n = sh.getLastRow() - fila;
  if (n <= 0) return [];
  return sh.getRange(fila + 1, 1, n, ancho).getValues()
    .filter((f) => String(f[0]).trim() !== '' && textoMes(f[0]) !== mes);
}

/** Agrega (o reemplaza) el mes en el Sheet de informes: Detalle, Entregas, Resumen y Por período. */
function escribirInforme(inf, d) {
  const pesos = '$ #,##0';
  if (inf.getName() !== NOMBRE_INFORME) inf.rename(NOMBRE_INFORME);

  // Detalle: una fila por neto de todos los meses generados. Insumos como valores, importes como fórmulas.
  const det = hojaInforme(inf, INF_DETALLE, 1, ENC_DETALLE);
  const nuevas = d.filas.map((f) => [d.mes, f.acto, f.solicitud, f.grupo, f.orden, f.nombre, f.conc, f.concepto,
    f.ahorro, f.pedido, f.familia, f.pub, f.ivaPed, '', f.pct2A, '', f.modeloBase, f.ceSinBase, f.pct2C, '', f.reu,
    f.entrega, f.pctAdelanto, '', f.descuento, f.nota]);
  const todas = filasSinMes(det, 1, ENC_DETALLE.length, d.mes).concat(nuevas)
    .sort((a, b) => ordenMesInforme(textoMes(a[0])) - ordenMesInforme(textoMes(b[0])));
  if (det.getLastRow() > 1) det.getRange(2, 1, det.getLastRow() - 1, ENC_DETALLE.length).clearContent();
  det.getRange('A:B').setNumberFormat('@'); // Mes y Acto como texto (si no, Sheets los convierte en fecha)
  if (todas.length) {
    det.getRange(2, 1, todas.length, ENC_DETALLE.length).setValues(todas.map((f, i) => {
      const r = i + 2;
      const x = f.slice();
      x[0] = textoMes(x[0]);
      x[1] = textoMes(x[1]);
      x[13] = `=L${r}/(1+M${r})`;
      x[15] = `=L${r}*O${r}`;
      x[19] = `=R${r}*S${r}`;
      x[23] = `=N${r}*W${r}`;
      return x;
    }));
    const fin = todas.length + 1;
    ['L', 'N', 'P', 'R', 'T', 'X', 'Y'].forEach((l) => det.getRange(`${l}2:${l}${fin}`).setNumberFormat(pesos));
    ['M', 'O', 'S', 'W'].forEach((l) => det.getRange(`${l}2:${l}${fin}`).setNumberFormat('0.0%'));
    det.getRange(`U2:V${fin}`).setNumberFormat('dd/MM/yyyy');
  }
  det.setColumnWidth(6, 220);
  det.setColumnWidth(8, 200);
  det.setColumnWidth(26, 320);
  if (det.getFilter()) det.getFilter().remove();
  det.getRange(1, 1, Math.max(2, todas.length + 1), ENC_DETALLE.length).createFilter();

  // Entregas: una fila por unidad con REU en el mes (margen por entrega 40%), todos los meses generados.
  const ent = hojaInforme(inf, INF_ENTREGAS, 1, ENC_ENTREGAS);
  const todasEnt = filasSinMes(ent, 1, ENC_ENTREGAS.length, d.mes).concat(d.entregas)
    .sort((a, b) => ordenMesInforme(textoMes(a[0])) - ordenMesInforme(textoMes(b[0])));
  if (ent.getLastRow() > 1) ent.getRange(2, 1, ent.getLastRow() - 1, ENC_ENTREGAS.length).clearContent();
  ent.getRange('A:A').setNumberFormat('@');
  ent.getRange('K:K').setNumberFormat('@');
  if (todasEnt.length) {
    ent.getRange(2, 1, todasEnt.length, ENC_ENTREGAS.length).setValues(todasEnt.map((f, i) => {
      const r = i + 2;
      const x = f.slice();
      x[0] = textoMes(x[0]);
      x[10] = textoMes(x[10]);
      x[13] = `=L${r}/(1+M${r})`;
      x[14] = `=N${r}*${MARGEN_COMISIONAL}*${MARGEN_ENTREGA}`;
      return x;
    }));
    const fin = todasEnt.length + 1;
    ent.getRange(`E2:G${fin}`).setNumberFormat('dd/MM/yyyy');
    ['L', 'N', 'O'].forEach((l) => ent.getRange(`${l}2:${l}${fin}`).setNumberFormat(pesos));
    ent.getRange(`M2:M${fin}`).setNumberFormat('0.0%');
  }
  ent.setColumnWidth(4, 220);
  ent.setColumnWidth(8, 240);
  ent.setColumnWidth(16, 320);
  if (ent.getFilter()) ent.getFilter().remove();
  ent.getRange(1, 1, Math.max(2, todasEnt.length + 1), ENC_ENTREGAS.length).createFilter();

  // Resumen: una fila por mes. Importes con fórmulas sobre el Detalle.
  const res = hojaInforme(inf, INF_RESUMEN, 3, ENC_RESUMEN);
  inf.setActiveSheet(res);
  inf.moveActiveSheet(1);
  res.getRange('A1').setValue(NOMBRE_INFORME).setFontSize(16).setFontWeight('bold').setFontColor(COLOR_ENCABEZADO);
  res.getRange('A2').setValue('Una fila por mes (al generar de nuevo un mes, se reemplaza). Detalle por operación ' +
    'en Detalle, por unidad entregada en Entregas; el resumen de un mes, con el detalle por familia, en "Por período".').setFontSize(9).setFontColor('#666666');
  const siNo = (b) => (b ? 'Sí' : 'No');
  const filaMes = [d.mes, d.categoria || '—', d.objetivo || '', d.totalNetos, '', d.adicionales, siNo(d.suscripciones),
    siNo(d.patentamientos), d.pct2A, d.pct2B, d.ceNC1, '', '', '', '', '', '', '', '', '', '',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), d.avisos.join(' • ')];
  const meses = filasSinMes(res, 3, ENC_RESUMEN.length, d.mes).concat([filaMes])
    .sort((a, b) => ordenMesInforme(textoMes(a[0])) - ordenMesInforme(textoMes(b[0])));
  if (res.getLastRow() > 3) res.getRange(4, 1, res.getLastRow() - 3, ENC_RESUMEN.length).clearContent().setBackground(null)
    .setFontWeight('normal');
  res.getRange('A:A').setNumberFormat('@');
  const D = `'${INF_DETALLE}'!`;
  const E = `'${INF_ENTREGAS}'!`;
  const s = (col, r) => `=SUMIFS(${D}${col}:${col},${D}A:A,$A${r})`;
  res.getRange(4, 1, meses.length, ENC_RESUMEN.length).setValues(meses.map((f, i) => {
    const r = i + 4;
    const x = f.slice();
    x[0] = textoMes(x[0]);
    x[4] = `=IF(N(C${r})>0,D${r}/C${r},"")`;
    x[11] = s('P', r);
    x[12] = `=F${r}*J${r}*K${r}`;
    x[13] = s('T', r);
    x[14] = `=L${r}+M${r}+N${r}`;
    x[15] = s('Y', r);
    x[16] = `=O${r}-P${r}`;
    x[17] = s('X', r);
    x[18] = `=COUNTIFS(${D}A:A,$A${r},${D}W:W,">0")`;
    x[19] = `=COUNTIFS(${E}A:A,$A${r})`;
    x[20] = `=SUMIFS(${E}O:O,${E}A:A,$A${r})`;
    return x;
  }));
  const ultimaRes = meses.length + 3;
  res.getRange(`E4:E${ultimaRes}`).setNumberFormat('0%');
  res.getRange(`I4:J${ultimaRes}`).setNumberFormat('0.00%');
  res.getRange(`K4:R${ultimaRes}`).setNumberFormat(pesos);
  res.getRange(`O4:O${ultimaRes}`).setFontWeight('bold');
  res.getRange(`Q4:Q${ultimaRes}`).setFontWeight('bold');
  res.getRange(`R4:R${ultimaRes}`).setFontWeight('bold').setBackground(FONDO_TOTAL);
  res.getRange(`U4:U${ultimaRes}`).setNumberFormat(pesos).setFontWeight('bold').setBackground(FONDO_TOTAL);
  res.getRange(`W4:W${ultimaRes}`).setFontColor('#b45f06').setFontSize(9).setWrap(true);
  res.setColumnWidth(1, 120);
  for (let col = 2; col <= 21; col++) res.setColumnWidth(col, col >= 11 ? 120 : 90);
  res.setColumnWidth(22, 120);
  res.setColumnWidth(23, 420);
  res.setRowHeight(3, 45);

  // Por período: el resumen de un mes (elegido en B2) con el formato del informe mensual.
  let per = inf.getSheetByName(INF_PERIODO);
  if (!per && inf.getSheetByName('Por familia')) per = inf.getSheetByName('Por familia').setName(INF_PERIODO);
  if (!per) per = inf.insertSheet(INF_PERIODO);
  const mesSel = textoMes(per.getRange('B2').getValue());
  const listaMeses = meses.map((f) => textoMes(f[0]));
  per.clear();
  per.getRange(1, 1, per.getMaxRows(), per.getMaxColumns()).breakApart().clearDataValidations();
  per.getRange('A1').setValue(`${NOMBRE_INFORME} - por período`).setFontSize(16).setFontWeight('bold')
    .setFontColor(COLOR_ENCABEZADO);
  per.getRange('A2').setValue('Período').setFontWeight('bold');
  per.getRange('B2').setNumberFormat('@').setValue(listaMeses.indexOf(mesSel) >= 0 ? mesSel : d.mes)
    .setFontWeight('bold').setBackground('#fff2cc')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(listaMeses, true).build());
  // Dato del mes elegido, leído de la fila del mes en Resumen.
  const R = `'${INF_RESUMEN}'!`;
  const dato = (col) => `INDEX(${R}${col}:${col},MATCH($B$2,${R}A:A,0))`;
  per.getRange('A4:B11').setValues([
    ['Categoría', `=IFERROR(${dato('B')},"")`],
    ['Objetivo de pedidos', `=IFERROR(${dato('C')},"")`],
    ['Netos del mes', `=IFERROR(${dato('D')},"")`],
    ['% cumplimiento pedidos', `=IFERROR(${dato('E')},"")`],
    ['Pedidos adicionales al objetivo', `=IFERROR(${dato('F')},"")`],
    ['Cumplió suscripciones (100%)', `=IFERROR(${dato('G')},"")`],
    ['Cumplió patentamientos (100%)', `=IFERROR(${dato('H')},"")`],
    ['Generado', `=IFERROR(${dato('V')},"")`],
  ]);
  per.getRange('A4:A11').setFontWeight('bold');
  per.getRange('B4:B11').setHorizontalAlignment('left');
  per.getRange('B7').setNumberFormat('0%');
  per.getRange('A13:D13').setValues([['Concepto', 'Condición', 'Base', 'Importe']]).setFontWeight('bold')
    .setBackground(COLOR_ENCABEZADO).setFontColor('white');
  const imp = (col) => `=IFERROR(${dato(col)},0)`;
  per.getRange('A14:D21').setValues([
    ['2A Bonus cumplimiento pedidos totales', `=IFERROR("100% pedidos y suscripciones ("&TEXT(${dato('I')},"0.00%")&")","")`,
      'Precio público c/IVA del modelo pedido', imp('L')],
    ['2B Pedidos adicionales al objetivo', `=IFERROR(${dato('F')}&" pedido(s) × "&TEXT(${dato('J')},"0.00%"),"")`,
      'Precio CE s/IVA de NC1', imp('M')],
    ['2C Extra bonus pedidos', `=IFERROR("100% patentamientos ("&${dato('H')}&")","")`,
      `=IF(OR(${MESES_2C_MODELO_PEDIDO.map((m) => `$B$2="${m}"`).join(',') || 'FALSE'}),` +
        '"Inc pedido del modelo pedido × CE s/IVA del modelo pedido",' +
        '"Inc pedido del modelo pedido × CE s/IVA del modelo base de la familia (Cronos NC1, Titano DT1...)")', imp('N')],
    ['Total incentivos', '', '', '=D14+D15+D16'],
    ['Descuentos otorgados', '', '$ descuento de BASE', imp('P')],
    ['Total incentivos - descuentos', '', '', '=D17-D18'],
    ['Adelanto de margen', `=IFERROR("60% del 14% · "&${dato('S')}&" de "&${dato('D')}&" pedidos (sin REU en el mes)","")`,
      'Precio público s/IVA del modelo pedido', imp('R')],
    ['Margen por entrega (40%)', `=IFERROR("40% del 14% · "&${dato('T')}&" unidades con REU en el mes","")`,
      'Precio público s/IVA del modelo entregado (lista del mes de la factura)', imp('U')],
  ]);
  per.getRange('D14:D21').setNumberFormat(pesos);
  per.getRange('A17:D17').setFontWeight('bold').setBackground(FONDO_TOTAL);
  per.getRange('A19:D19').setFontWeight('bold').setBackground(FONDO_TOTAL);

  // Por familia del mes elegido (pedidos: familia del modelo pedido; entregas: del modelo entregado).
  const encF = ['Familia', 'Netos', '2A', '2C', 'Adelanto margen', 'Pedidos con adelanto', 'Descuentos',
    'Unidades con REU', 'Margen por entrega (40%)'];
  per.getRange(23, 1, 1, encF.length).setValues([encF]).setFontWeight('bold').setBackground(COLOR_ENCABEZADO)
    .setFontColor('white');
  const familias = [...new Set(todas.map((f) => String(f[10] || '').trim() || 'Sin familia')
    .concat(todasEnt.map((f) => String(f[9] || '').trim() || 'Sin familia')))].sort();
  let filaAvisos = 26;
  if (familias.length) {
    const crit = (r) => `IF($A${r}="Sin familia","",$A${r})`;
    per.getRange(24, 1, familias.length, encF.length).setValues(familias.map((fa, i) => {
      const r = i + 24;
      const sf = (col) => `=SUMIFS(${D}${col}:${col},${D}A:A,$B$2,${D}K:K,${crit(r)})`;
      return [fa, `=COUNTIFS(${D}A:A,$B$2,${D}K:K,${crit(r)})`, sf('P'), sf('T'), sf('X'),
        `=COUNTIFS(${D}A:A,$B$2,${D}K:K,${crit(r)},${D}W:W,">0")`, sf('Y'),
        `=COUNTIFS(${E}A:A,$B$2,${E}J:J,${crit(r)})`, `=SUMIFS(${E}O:O,${E}A:A,$B$2,${E}J:J,${crit(r)})`];
    }));
    const t = familias.length + 24;
    per.getRange(t, 1, 1, encF.length).setValues([['TOTAL', `=SUM(B24:B${t - 1})`, `=SUM(C24:C${t - 1})`,
      `=SUM(D24:D${t - 1})`, `=SUM(E24:E${t - 1})`, `=SUM(F24:F${t - 1})`, `=SUM(G24:G${t - 1})`,
      `=SUM(H24:H${t - 1})`, `=SUM(I24:I${t - 1})`]])
      .setFontWeight('bold').setBackground(FONDO_TOTAL);
    per.getRange(24, 3, familias.length + 1, 3).setNumberFormat(pesos);
    per.getRange(24, 7, familias.length + 1, 1).setNumberFormat(pesos);
    per.getRange(24, 9, familias.length + 1, 1).setNumberFormat(pesos);
    filaAvisos = t + 2;
  }
  per.getRange(filaAvisos, 1).setValue('Avisos').setFontWeight('bold');
  per.getRange(filaAvisos + 1, 1).setFormula(`=IFERROR(SUBSTITUTE(${dato('W')}," • ",CHAR(10)),"")`)
    .setFontColor('#b45f06').setWrap(true);
  per.getRange(filaAvisos + 1, 1, 1, 4).merge();
  per.setColumnWidth(1, 260);
  per.setColumnWidth(2, 260);
  per.setColumnWidth(3, 320);
  per.setColumnWidth(4, 140);
  for (let col = 5; col <= 9; col++) per.setColumnWidth(col, 130);

  // Hoja vacía que crea Sheets en un archivo nuevo.
  const vacia = inf.getSheetByName('Hoja 1') || inf.getSheetByName('Sheet1');
  if (vacia && vacia.getLastRow() === 0 && inf.getSheets().length > 1) inf.deleteSheet(vacia);
  SpreadsheetApp.flush();
}

/** Comparte el informe con los destinatarios (lectura) y les manda los totales del mes. */
function enviarInforme(inf, d, mails) {
  const archivo = DriveApp.getFileById(inf.getId());
  mails.forEach((m) => {
    try { archivo.addViewer(m); } catch (e) { /* p. ej. mail sin cuenta Google: igual recibe el link */ }
  });
  const sum = (fn) => d.filas.reduce((a, f) => a + fn(f), 0);
  const t2A = sum((f) => f.pub * f.pct2A);
  const t2B = d.adicionales * d.pct2B * d.ceNC1;
  const t2C = sum((f) => f.ceSinBase * f.pct2C);
  const desc = sum((f) => f.descuento);
  const adelanto = sum((f) => f.pub / (1 + f.ivaPed) * f.pctAdelanto);
  const conAdelanto = d.filas.filter((f) => f.pctAdelanto > 0).length;
  const lineas = [
    ['2A Bonus cumplimiento pedidos totales', t2A], ['2B Pedidos adicionales al objetivo', t2B],
    ['2C Extra bonus pedidos', t2C], ['Total incentivos', t2A + t2B + t2C], ['Descuentos otorgados', desc],
    ['Total incentivos - descuentos', t2A + t2B + t2C - desc],
    [`Adelanto de margen (${conAdelanto} de ${d.filas.length} pedidos)`, adelanto],
    [`Margen por entrega 40% (${d.entregas.length} unidades con REU)`,
      d.entregas.reduce((a, f) => a + f[11] / (1 + f[12]) * MARGEN_COMISIONAL * MARGEN_ENTREGA, 0)],
  ];
  const fmt = (v) => '$ ' + Math.round(Number(v) || 0).toLocaleString('es-AR');
  const filas = lineas.map((f) => `<tr><td style="padding:4px 10px">${f[0]}</td>` +
    `<td style="padding:4px 10px;text-align:right">${fmt(f[1])}</td></tr>`).join('');
  MailApp.sendEmail({
    to: mails.join(','),
    subject: `Incentivos pedidos FIAT - ${d.mes}`,
    htmlBody: `<p>Informe de incentivos de pedidos FIAT de <b>${d.mes}</b>: ${d.totalNetos} netos` +
      (d.objetivo ? ` sobre un objetivo de ${d.objetivo} (categoría ${d.categoria})` : '') + '.</p>' +
      `<table style="border-collapse:collapse">${filas}</table>` +
      `<p><a href="${inf.getUrl()}">Ver el detalle en Google Sheets</a></p>` +
      (d.avisos.length ? `<p style="color:#b45f06">Atención:<br>${d.avisos.join('<br>')}</p>` : ''),
  });
}

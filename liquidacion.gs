/**
 * Liquidación FIAT - PDA (archivo "Incentivos pedidos FIAT").
 *
 * Reproduce y previsiona, mes a mes, la liquidación de Fiat:
 *   Adelanto margen · Recupero adelanto margen · Margen por pedido (60%) · Margen por entrega (40%)
 *   · Bonus pedidos aceptados (2A + 2B) · Bonus pedidos por familia (2C).
 *
 * Lee (sin modificarlos):
 *   - Archivo de pedidos "Pedidos - PDA": BASE, PRECIOS, OBJETIVOS, INCENTIVOS.
 *   - Archivo "Entregas PDA": Base (FBS, se actualiza sola) y SGA (las unidades con REU generado).
 *
 * Hojas de este archivo:
 *   RESUMEN       mes elegido: cada línea de la liquidación de Fiat, lo liquidado (LIQ_FIAT) y la diferencia.
 *   PARAMETROS    una fila por mes (suscripciones / patentamientos cumplidos, margen, adelanto, B90).
 *   FAMILIAS      modelo base del 2C e IVA de cada familia.
 *   MODELOS       versión de entrega (texto de FBS) -> código de modelo.
 *   LIQ_FIAT      lo que liquidó Fiat (se carga a mano) para comparar.
 *   PEDIDOS_LIQ   una fila por neto (todos los meses): 2A, 2C y adelanto de margen.
 *   ENTREGAS_LIQ  una fila por unidad con REU (todos los meses): recupero, margen 60% y 40%.
 *   MESES         datos por mes que arma el script (objetivo, categoría, 2B...). No se edita.
 *
 * Reglas (señal 1599, liquidaciones de Fiat de septiembre 2026):
 *   - Mes del REU = mes de la "Fecha Ingreso" de SGA (coincide con lo que liquida Fiat). Una unidad
 *     entregada (FBS) que todavía no está en SGA queda "PENDIENTE REU".
 *   - Adelanto margen: pedidos del mes SIN REU en el mismo mes. Precio público s/IVA del modelo
 *     pedido, lista del mes del pedido × margen (14%) × adelanto (60%).
 *   - Margen por pedido (60%): unidades con REU en el mes, a precio del mes del REU.
 *   - Recupero adelanto (−60%): unidades con REU en el mes cuyo pedido es de un mes anterior, a
 *     precio actualizado (mes del REU).
 *   - Margen por entrega (40%): unidades con REU en el mes, a precio del mes de entrega.
 *   - 2A: % según categoría × precio público c/IVA del modelo pedido (si cumple pedidos y suscripciones).
 *   - 2B: % según categoría × CE s/IVA de NC1 × pedidos sobre el objetivo.
 *   - 2C: Inc pedido del modelo pedido × CE s/IVA del modelo base de su familia (si cumple
 *     patentamientos; sin Cronos 90/10 B90 desde el grupo indicado).
 *   - Operaciones de otro concesionario (no TURIN): sin adelanto, marcadas.
 */

const ID_PEDIDOS = '1dsvnNpKUmoKDSaZcEv4UsofHWbWK_m0SwKKvr7dkxGM';
const ID_ENTREGAS = '1lWVidE-2smtBoAEWC5Z_kEbHXXJ3cpoBGArL5YE96Mw';
const CONCESIONARIO = 'TURIN S.A.';

const MESES_NOMBRE = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE',
  'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const COLOR_ENC = '#7f0000';
const FONDO_TOT = '#f3f3f3';
const GRIS = '#666666';
const PESOS = '$ #,##0';

const H_RESUMEN = 'RESUMEN';
const H_PARAM = 'PARAMETROS';
const H_FAMILIAS = 'FAMILIAS';
const H_MODELOS = 'MODELOS';
const H_FIAT = 'LIQ_FIAT';
const H_PED = 'PEDIDOS_LIQ';
const H_ENT = 'ENTREGAS_LIQ';
const H_MESES = 'MESES';

// Valores iniciales (se escriben solo al crear cada hoja; después se editan en el archivo).
const PARAM_DEFECTO = { suscripciones: 'Sí', patentamientos: 'Sí', margen: 0.14, adelanto: 0.6, b90Grupo: 18002,
  b90Concepto: '90/10' };
const FAMILIAS_INICIALES = [
  ['Cronos', 'NC1', 0.21], ['Titano', 'DT1', 0.105], ['Toro', 'NT3', 0.105], ['Fiorino', 'FO1', 0.105],
  ['Strada', 'FS1', 0.105], ['Fastback', 'FT3', 0.21], ['Pulse', 'FP3', 0.21], ['Argo', 'AR2', 0.21],
  ['Mobi', 'MB1', 0.21], ['600', 'GO0', 0.21],
];
const MODELOS_INICIALES = [
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
const CONCEPTOS_FIAT = ['Adelanto Margen', 'Recupero Adelanto Margen', 'Márgen Por Pedido (60%)',
  'Margen por Entrega (40%)', 'Bonus Pedidos Aceptados Fiat', 'Bonus Pedidos por familia (2C)'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Liquidación')
    .addItem('Actualizar liquidación', 'actualizarLiquidacion')
    .addItem('Enviar resumen por mail', 'enviarResumen')
    .addSeparator()
    .addItem('Armar / reparar estructura', 'armarLiquidacion')
    .addToUi();
}

// ---------------------------------------------------------------------------
// Estructura
// ---------------------------------------------------------------------------

function armarLiquidacion() {
  const ss = SpreadsheetApp.getActive();
  armarParametros(ss);
  armarFamilias(ss);
  armarModelos(ss);
  armarLiqFiat(ss);
  actualizarLiquidacion(true);
  // Hojas del informe viejo ("Resumen" / "Detalle"), si quedaron: se ocultan, no se borran.
  ['Resumen', 'Detalle'].forEach((n) => {
    const h = ss.getSheetByName(n);
    if (h && n !== H_RESUMEN) h.hideSheet();
  });
}

function hojaLiq(ss, nombre) {
  let sh = ss.getSheetByName(nombre);
  const nueva = !sh;
  if (nueva) sh = ss.insertSheet(nombre);
  return { sh, nueva };
}

function encabezado(sh, fila, valores) {
  sh.getRange(fila, 1, 1, valores.length).setValues([valores]).setFontWeight('bold').setBackground(COLOR_ENC)
    .setFontColor('white').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
}

function armarParametros(ss) {
  const { sh, nueva } = hojaLiq(ss, H_PARAM);
  if (!nueva) return;
  encabezado(sh, 1, ['Mes', 'Cumplió suscripciones (100%)', 'Cumplió patentamientos (100%)', 'Margen comisional',
    '% adelanto margen', 'B90: grupo desde', 'B90: concepto de venta', 'Notas (señales del mes)']);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('D:E').setNumberFormat('0%');
  const siNo = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build();
  sh.getRange('B2:C').setDataValidation(siNo);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(8, 360);
  sh.getRange('A1').setNote('Una fila por mes. Los meses nuevos se agregan solos (copian el mes anterior) al ' +
    'actualizar la liquidación. Categoría, objetivo y % 2A/2B se leen del archivo de pedidos.');
}

function armarFamilias(ss) {
  const { sh, nueva } = hojaLiq(ss, H_FAMILIAS);
  if (!nueva) return;
  encabezado(sh, 1, ['Familia', 'Modelo base 2C', 'IVA']);
  sh.getRange(2, 1, FAMILIAS_INICIALES.length, 3).setValues(FAMILIAS_INICIALES);
  sh.getRange('C:C').setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.getRange('B1').setNote('El 2C de todos los pedidos de la familia se paga al precio CE sin IVA de este modelo ' +
    '(así liquida Fiat).');
}

function armarModelos(ss) {
  const { sh, nueva } = hojaLiq(ss, H_MODELOS);
  if (!nueva) return;
  encabezado(sh, 1, ['Versión de entrega (FBS)', 'Código', 'Version en PRECIOS']);
  sh.getRange(2, 1, MODELOS_INICIALES.length, 2).setValues(MODELOS_INICIALES);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(3, 320);
  sh.getRange('A1').setNote('Las versiones nuevas se agregan solas al actualizar, con el Código vacío: completarlo.');
}

function armarLiqFiat(ss) {
  const { sh, nueva } = hojaLiq(ss, H_FIAT);
  if (!nueva) return;
  encabezado(sh, 1, ['Mes', 'Concepto', 'Importe liquidado', 'Fecha', 'Nota']);
  const sep = 'SEPTIEMBRE 26';
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange(2, 1, 6, 4).setValues([
    [sep, 'Adelanto Margen', 188478807.94, '01/10/2026'],
    [sep, 'Recupero Adelanto Margen', -159658407.21, '01/10/2026'],
    [sep, 'Márgen Por Pedido (60%)', 163410372.44, '01/10/2026'],
    [sep, 'Margen por Entrega (40%)', 108038801.74, '01/10/2026'],
    [sep, 'Bonus Pedidos Aceptados Fiat', 28078984.30, '05/10/2026'],
    [sep, 'Bonus Pedidos por familia (2C)', 68860122.49, 'HES 105880926'],
  ]);
  sh.getRange('C:C').setNumberFormat('$ #,##0.00');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 240);
  const lista = SpreadsheetApp.newDataValidation().requireValueInList(CONCEPTOS_FIAT, true).build();
  sh.getRange('B2:B').setDataValidation(lista);
}

// ---------------------------------------------------------------------------
// Cálculo
// ---------------------------------------------------------------------------

function actualizarLiquidacion(silencioso) {
  const ss = SpreadsheetApp.getActive();
  ['PARAMETROS', 'FAMILIAS', 'MODELOS', 'LIQ_FIAT'].forEach((n) => {
    if (!ss.getSheetByName(n)) ({ PARAMETROS: armarParametros, FAMILIAS: armarFamilias, MODELOS: armarModelos,
      LIQ_FIAT: armarLiqFiat })[n](ss);
  });
  const d = leerDatos(ss);
  const ped = calcularPedidos(d);
  const ent = calcularEntregas(d, ped);
  completarParametros(ss, d);
  escribirMeses(ss, d, ped);
  escribirPedidos(ss, ped.filas);
  escribirEntregas(ss, ent.filas);
  armarResumen(ss, d);
  SpreadsheetApp.flush();
  if (silencioso !== true) {
    SpreadsheetApp.getUi().alert(`Liquidación actualizada: ${ped.filas.length} pedidos netos y ` +
      `${ent.filas.length} entregas.` + (d.avisos.length ? `\n\nAtención:\n• ${d.avisos.join('\n• ')}` : ''));
  }
}

/** Lee todo lo necesario de los archivos de pedidos y entregas y de las hojas de parámetros. */
function leerDatos(ss) {
  const avisos = [];
  const ped = SpreadsheetApp.openById(ID_PEDIDOS);
  const ent = SpreadsheetApp.openById(ID_ENTREGAS);

  // PRECIOS: precio[mes][modelo] = {pub, ceSin, inc, familia}; familia[modelo].
  const precio = {};
  const familiaDe = {};
  ped.getSheetByName('PRECIOS').getDataRange().getValues().slice(1).forEach((f) => {
    const mes = textoMes(f[0]);
    const mod = String(f[1]).trim();
    if (!mes || !mod) return;
    (precio[mes] = precio[mes] || {})[mod] = { pub: num(f[4]), ceSin: num(f[6]), inc: f[3] === '' ? null : num(f[3]),
      version: String(f[2]) };
    if (String(f[7] || '').trim()) familiaDe[mod] = String(f[7]).trim();
  });

  // Objetivo y categoría; % 2A y 2B por categoría.
  const objetivos = {};
  ped.getSheetByName('OBJETIVOS').getRange('A2:C').getValues().forEach((f) => {
    const m = textoMes(f[0]);
    if (m) objetivos[m] = { objetivo: num(f[1]), categoria: String(f[2]).trim().toUpperCase() };
  });
  const incentivos = {};
  ped.getSheetByName('INCENTIVOS').getRange('A2:E').getValues().forEach((f) => {
    const m = textoMes(f[0]);
    if (!m) return;
    const c = String(f[1]).toUpperCase();
    const k = /CUMPLIMIENTO/.test(c) ? '2A' : /ADICIONALES/.test(c) ? '2B' : '';
    if (k) (incentivos[m] = incentivos[m] || {})[k] = { A: num(f[2]), B: num(f[3]), C: num(f[4]) };
  });

  // BASE de pedidos (por nombre de columna).
  const bVals = ped.getSheetByName('BASE').getDataRange().getValues();
  const bh = bVals[0].map((h) => String(h).trim());
  const col = (n, alias) => {
    const i = [n].concat(alias || []).map((x) => bh.indexOf(x)).find((x) => x >= 0);
    if (i === undefined) throw new Error(`En BASE de pedidos no está la columna "${n}".`);
    return i;
  };
  const C = {
    acto: col('Acto'), conc: col('Concesionario'), concepto: col('Concepto de venta'), ahorro: col('Modelo ahorro'),
    pedido: col('Modelo PEDIDO'), desc: bh.indexOf('$ descuento'), cierre: col('Mes cierre'), sol: col('Solicitud'),
    grupo: col('Grupo'), orden: col('Orden'), nombre: col('Apellido, Nombre'), estPed: col('Pedido'),
    carpeta: col('Carpeta'),
  };
  const base = bVals.slice(1).filter((f) => String(f[C.sol]).trim() !== '');

  // Entregas: FBS (Base) y SGA.
  const fbs = {};
  const fVals = ent.getSheetByName('Base').getDataRange().getValues();
  const fh = fVals[0].map((h) => String(h).trim());
  const fi = (n) => fh.indexOf(n);
  fVals.slice(1).forEach((f) => {
    const k = claveGO(f[fi('Grupo')], f[fi('Orden')]);
    if (k) fbs[k] = { entrega: aFecha(f[fi('FechaEntrega')]), modelo: String(f[fi('Modelo')]).replace(/\s+/g, ' ').trim(),
      cliente: `${String(f[fi('Apellido')]).trim()}, ${String(f[fi('Nombre')]).trim()}`, remito: f[fi('Remito')] };
  });
  const sVals = ent.getSheetByName('SGA').getDataRange().getValues();
  const iEnc = sVals.findIndex((f) => f.map(String).indexOf('Grupo') >= 0);
  const sh = sVals[iEnc].map((h) => String(h).trim());
  const si = (n) => sh.indexOf(n);
  const sga = {};
  sVals.slice(iEnc + 1).forEach((f) => {
    const k = claveGO(f[si('Grupo')], f[si('Orden')]);
    if (k) sga[k] = { cliente: String(f[si('Razón social')]).trim(), ingreso: aFecha(f[si('Fecha Ingreso')]),
      entrega: aFecha(f[si('Fecha ent.unidad')]) };
  });

  // Parámetros por mes, familias y equivalencias de modelos.
  const param = {};
  const pVals = ss.getSheetByName(H_PARAM).getDataRange().getValues().slice(1);
  pVals.forEach((f) => {
    const m = textoMes(f[0]);
    if (m) param[m] = { suscripciones: f[1] !== 'No', patentamientos: f[2] !== 'No', margen: num(f[3]) || PARAM_DEFECTO.margen,
      adelanto: num(f[4]) || PARAM_DEFECTO.adelanto, b90Grupo: num(f[5]) || 0, b90Concepto: String(f[6] || '').trim() };
  });
  const familias = {};
  ss.getSheetByName(H_FAMILIAS).getDataRange().getValues().slice(1).forEach((f) => {
    if (String(f[0]).trim()) familias[String(f[0]).trim().toUpperCase()] = { nombre: String(f[0]).trim(),
      base: String(f[1]).trim(), iva: num(f[2]) };
  });
  const modelos = {};
  const mSh = ss.getSheetByName(H_MODELOS);
  mSh.getDataRange().getValues().slice(1).forEach((f) => {
    if (String(f[0]).trim()) modelos[normal(f[0])] = String(f[1]).trim();
  });
  // Modelos sin Familia en PRECIOS: se deduce del texto de la versión (MODELOS o PRECIOS), p. ej. DP1 -> Cronos.
  const famDeTexto = (t) => {
    const F = Object.keys(familias).find((x) => normal(t).split(' ').indexOf(x) >= 0);
    return F ? familias[F].nombre : '';
  };
  mSh.getDataRange().getValues().slice(1).forEach((f) => {
    const cod = String(f[1]).trim();
    if (cod && !familiaDe[cod] && famDeTexto(f[0])) familiaDe[cod] = famDeTexto(f[0]);
  });
  Object.keys(precio).forEach((m) => Object.keys(precio[m]).forEach((cod) => {
    if (!familiaDe[cod] && famDeTexto(precio[m][cod].version)) familiaDe[cod] = famDeTexto(precio[m][cod].version);
  }));
  // Versiones de entrega nuevas: se agregan a MODELOS con el código vacío.
  const nuevas = [...new Set(Object.values(fbs).map((x) => x.modelo))].filter((m) => m && !(normal(m) in modelos));
  if (nuevas.length) {
    mSh.getRange(mSh.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas.map((m) => [m, '']));
    avisos.push(`Versiones de entrega nuevas en MODELOS, sin código: ${nuevas.join(', ')}.`);
  }

  return { precio, familiaDe, objetivos, incentivos, base, C, fbs, sga, param, familias, modelos, avisos };
}

function paramMes(d, mes) {
  return d.param[mes] || Object.assign({}, PARAM_DEFECTO, { suscripciones: true, patentamientos: true });
}

function ivaDe(d, mod) {
  const fam = String(d.familiaDe[mod] || '').toUpperCase();
  return d.familias[fam] ? d.familias[fam].iva : 0.21;
}

/** Netos de cada mes: 2A, 2C y adelanto de margen. */
function calcularPedidos(d) {
  const C = d.C;
  const faltan = new Set();
  const porMes = {};
  const filas = [];
  d.base.forEach((f) => {
    const mes = textoMes(f[C.cierre]);
    if (!mes || f[C.estPed] !== 'APROBADO' || f[C.carpeta] !== 'APROBADA') return;
    porMes[mes] = (porMes[mes] || 0) + 1;
  });
  const reuMes = {}; // grupo-orden -> mes del REU (Fecha Ingreso de SGA)
  Object.keys(d.sga).forEach((k) => { if (d.sga[k].ingreso) reuMes[k] = etiquetaMes(d.sga[k].ingreso); });

  d.base.forEach((f) => {
    const mes = textoMes(f[C.cierre]);
    if (!mes || f[C.estPed] !== 'APROBADO' || f[C.carpeta] !== 'APROBADA') return;
    const p = paramMes(d, mes);
    const obj = d.objetivos[mes] || {};
    const cumple = obj.objetivo > 0 && porMes[mes] >= obj.objetivo;
    const pct = (k) => ((d.incentivos[mes] || {})[k] || {})[obj.categoria] || 0;
    const ahorro = String(f[C.ahorro]).trim();
    const pedido = String(f[C.pedido]).trim() || ahorro;
    const fam = d.familiaDe[pedido] || d.familiaDe[ahorro] || '';
    const famInfo = d.familias[fam.toUpperCase()] || {};
    const mBase = famInfo.base || ahorro;
    const pr = (d.precio[mes] || {})[pedido] || {};
    const pb = (d.precio[mes] || {})[mBase] || {};
    if (!pr.pub) faltan.add(`${mes} ${pedido}`);
    if (!pb.ceSin) faltan.add(`${mes} ${mBase} (CE)`);
    const concepto = String(f[C.concepto]);
    const b90 = /CRONOS/i.test(fam) && p.b90Concepto && concepto.replace(/\s/g, '').indexOf(p.b90Concepto) === 0 &&
      num(f[C.grupo]) >= p.b90Grupo;
    const go = claveGO(f[C.grupo], f[C.orden]);
    const otroCE = String(f[C.conc]).trim().toUpperCase() !== CONCESIONARIO;
    const reuMismoMes = reuMes[go] === mes;
    const notas = [];
    if (b90) notas.push('Cronos 90/10 B90: sin 2C');
    if (otroCE) notas.push(`Otro concesionario (${f[C.conc]}): sin adelanto`);
    if (reuMismoMes) notas.push('REU en el mismo mes: sin adelanto (se paga el 100% en entregas)');
    if (pr.inc === null || pr.inc === undefined) notas.push(`${pedido} sin Inc pedido en PRECIOS`);
    filas.push({
      mes, acto: textoMes(f[C.acto]), sol: f[C.sol], grupo: f[C.grupo], orden: f[C.orden], nombre: f[C.nombre],
      conc: f[C.conc], concepto, ahorro, pedido, fam, pub: pr.pub || 0, iva: ivaDe(d, pedido),
      pct2A: cumple && p.suscripciones ? pct('2A') : 0, mBase, ceBase: pb.ceSin || 0,
      pct2C: !p.patentamientos || b90 ? 0 : (pr.inc || 0), reu: reuMismoMes ? 'Sí' : 'No',
      pctAdelanto: reuMismoMes || otroCE ? 0 : p.margen * p.adelanto, desc: C.desc >= 0 ? num(f[C.desc]) : 0,
      nota: notas.join('. '), go,
    });
  });
  if (faltan.size) d.avisos.push(`Faltan precios en PRECIOS (pedidos): ${[...faltan].join(', ')}.`);
  return { filas, porMes };
}

/** Precio del modelo en el mes; si no está cargado, el del mes más cercano (a igual distancia, el anterior). */
function precioCercano(d, mes, cod) {
  if (((d.precio[mes] || {})[cod] || {}).pub) return { mes, precio: d.precio[mes][cod] };
  const o = ordenMes(mes);
  const otros = Object.keys(d.precio).filter((m) => (d.precio[m][cod] || {}).pub)
    .sort((a, b) => (Math.abs(ordenMes(a) - o) - Math.abs(ordenMes(b) - o)) || (ordenMes(a) - ordenMes(b)));
  return otros.length ? { mes: otros[0], precio: d.precio[otros[0]][cod] } : { mes: '', precio: {} };
}

/** Unidades con REU (SGA) y entregadas sin REU (FBS): recupero, margen 60% y 40%. */
function calcularEntregas(d, ped) {
  const pedPorGO = {};
  // Mes del pedido de cada grupo-orden (netos de BASE, cualquier mes).
  d.base.forEach((f) => {
    const C = d.C;
    const mes = textoMes(f[C.cierre]);
    if (mes) pedPorGO[claveGO(f[C.grupo], f[C.orden])] = { mes, sol: f[C.sol], pedido: String(f[C.pedido]).trim() ||
      String(f[C.ahorro]).trim(), conc: f[C.conc] };
  });
  const faltan = new Set();
  const sinCodigo = new Set();
  const filas = [];
  const claves = new Set(Object.keys(d.sga).concat(Object.keys(d.fbs)));
  claves.forEach((k) => {
    const s = d.sga[k];
    const fb = d.fbs[k] || {};
    const pe = pedPorGO[k] || {};
    const fechaEnt = (s && s.entrega) || fb.entrega || null;
    const mesReu = s && s.ingreso ? etiquetaMes(s.ingreso) : 'PENDIENTE REU';
    const mesEnt = fechaEnt ? etiquetaMes(fechaEnt) : '';
    let codigo = fb.modelo ? (d.modelos[normal(fb.modelo)] || '') : '';
    if (!codigo && pe.pedido) codigo = pe.pedido;
    if (!codigo) sinCodigo.add(fb.modelo || k);
    const p = paramMes(d, mesReu);
    const conReu = !!(s && s.ingreso);
    const notas = [];
    const pReu = precioCercano(d, mesReu, codigo);
    const pEnt = (d.precio[mesEnt] || {})[codigo] && (d.precio[mesEnt] || {})[codigo].pub ? precioCercano(d, mesEnt, codigo)
      : pReu;
    const pr = pReu.precio;
    const prE = pEnt.precio;
    if (conReu && !codigo) notas.push('Sin código de modelo: completarlo en MODELOS (o el Modelo PEDIDO en la BASE)');
    if (conReu && codigo && pReu.mes !== mesReu) {
      faltan.add(`${mesReu} ${codigo}`);
      notas.push(pr.pub ? `Sin precio de ${codigo} en ${mesReu}: se usa el de ${pReu.mes}` : `Sin precio de ${codigo}`);
    }
    const mismoMes = pe.mes && pe.mes === mesReu;
    if (!pe.mes) notas.push('Pedido no está en la BASE (anterior a julio 26): se toma adelanto ya cobrado');
    if (mesEnt && mesReu !== mesEnt && conReu) notas.push(`Entregada en ${mesEnt}, REU en ${mesReu}`);
    if (conReu && mesEnt && pEnt === pReu && mesEnt !== mesReu) {
      notas.push(`Sin precio de ${codigo} en ${mesEnt}: 40% a precio de ${pReu.mes}`);
    }
    const otroCE = pe.conc && String(pe.conc).trim().toUpperCase() !== CONCESIONARIO;
    if (otroCE) notas.push(`Otro concesionario (${pe.conc}): sin adelanto, no hay recupero`);
    filas.push({
      mesReu, mesEnt, grupo: k.split('-')[0], orden: k.split('-')[1], cliente: (s && s.cliente) || fb.cliente || '',
      entrega: fechaEnt, ingreso: s ? s.ingreso : null, modeloTxt: fb.modelo || '', codigo,
      fam: d.familiaDe[codigo] || '', sol: pe.sol || '', mesPed: pe.mes || '', iva: ivaDe(d, codigo),
      pubReu: conReu ? (pr.pub || 0) : 0, pubEnt: conReu ? (prE.pub || 0) : 0, margen: p.margen, adelanto: p.adelanto,
      recupera: conReu && !mismoMes && !otroCE ? 'Sí' : 'No', nota: notas.join('. '),
    });
  });
  if (faltan.size) d.avisos.push(`Faltan precios en PRECIOS (entregas; se usó el mes más cercano): ${[...faltan].join(', ')}.`);
  if (sinCodigo.size) d.avisos.push(`Entregas sin código de modelo (completar en MODELOS): ${[...sinCodigo].join(', ')}.`);
  filas.sort((a, b) => (ordenMes(a.mesReu) - ordenMes(b.mesReu)) || String(a.cliente).localeCompare(b.cliente));
  return { filas };
}

/** Agrega a PARAMETROS los meses con netos o REU que falten (copiando el mes anterior). */
function completarParametros(ss, d) {
  const sh = ss.getSheetByName(H_PARAM);
  const meses = new Set();
  d.base.forEach((f) => { const m = textoMes(f[d.C.cierre]); if (m) meses.add(m); });
  Object.values(d.sga).forEach((s) => { if (s.ingreso) meses.add(etiquetaMes(s.ingreso)); });
  const faltan = [...meses].filter((m) => !d.param[m]).sort((a, b) => ordenMes(a) - ordenMes(b));
  if (!faltan.length) return;
  const filas = faltan.map((m) => {
    const prev = Object.keys(d.param).filter((x) => ordenMes(x) < ordenMes(m)).sort((a, b) => ordenMes(b) - ordenMes(a))[0];
    const p = prev ? d.param[prev] : PARAM_DEFECTO;
    d.param[m] = Object.assign({}, p);
    return [m, 'Sí', 'Sí', p.margen, p.adelanto, p.b90Grupo || PARAM_DEFECTO.b90Grupo, p.b90Concepto || PARAM_DEFECTO.b90Concepto,
      'Agregado automáticamente: revisar con las señales del mes'];
  });
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  const todas = sh.getRange(2, 1, sh.getLastRow() - 1, 8);
  const vals = todas.getValues().sort((a, b) => ordenMes(textoMes(a[0])) - ordenMes(textoMes(b[0])));
  todas.setValues(vals);
}

/** Hoja MESES: datos por mes (objetivo, categoría, netos, 2B). */
function escribirMeses(ss, d, ped) {
  const { sh } = hojaLiq(ss, H_MESES);
  sh.clear();
  const meses = [...new Set(Object.keys(ped.porMes).concat(Object.keys(d.param)))]
    .sort((a, b) => ordenMes(a) - ordenMes(b));
  const filas = meses.map((m) => {
    const o = d.objetivos[m] || {};
    const netos = ped.porMes[m] || 0;
    const cumple = o.objetivo > 0 && netos >= o.objetivo;
    const pct = (k) => ((d.incentivos[m] || {})[k] || {})[o.categoria] || 0;
    const p = paramMes(d, m);
    const adic = cumple ? netos - o.objetivo : 0;
    const ceNC1 = ((d.precio[m] || {}).NC1 || {}).ceSin || 0;
    return [m, o.categoria || '', o.objetivo || '', netos, cumple ? 'Sí' : 'No', p.suscripciones ? 'Sí' : 'No',
      p.patentamientos ? 'Sí' : 'No', cumple && p.suscripciones ? pct('2A') : 0, pct('2B'), adic, ceNC1,
      adic * (cumple ? pct('2B') : 0) * ceNC1];
  });
  encabezado(sh, 1, ['Mes', 'Categoría', 'Objetivo pedidos', 'Netos', 'Cumplió pedidos', 'Cumplió suscripciones',
    'Cumplió patentamientos', '% 2A', '% 2B', 'Pedidos adicionales', 'CE s/IVA NC1', '2B importe']);
  sh.getRange('A:A').setNumberFormat('@');
  if (filas.length) sh.getRange(2, 1, filas.length, 12).setValues(filas);
  sh.getRange('H:I').setNumberFormat('0.00%');
  sh.getRange('K:L').setNumberFormat(PESOS);
  sh.setFrozenRows(1);
}

function escribirPedidos(ss, filas) {
  const { sh } = hojaLiq(ss, H_PED);
  sh.clear();
  const enc = ['Mes', 'Acto', 'Solicitud', 'Grupo', 'Orden', 'Apellido, Nombre', 'Concesionario', 'Concepto de venta',
    'Modelo ahorro', 'Modelo pedido', 'Familia', 'Precio público c/IVA (pedido)', '% IVA', 'Precio público s/IVA',
    '% 2A', '2A Bonus cumplimiento', 'Modelo base 2C', 'Precio CE s/IVA (modelo base)', '% 2C (Inc pedido)',
    '2C Extra bonus', 'REU en el mismo mes', '% adelanto (14% × 60%)', 'Adelanto margen', '$ descuento', 'Nota'];
  encabezado(sh, 1, enc);
  sh.setFrozenRows(1);
  sh.getRange('A:B').setNumberFormat('@');
  if (filas.length) {
    sh.getRange(2, 1, filas.length, enc.length).setValues(filas.map((f, i) => {
      const r = i + 2;
      return [f.mes, f.acto, f.sol, f.grupo, f.orden, f.nombre, f.conc, f.concepto, f.ahorro, f.pedido, f.fam, f.pub,
        f.iva, `=L${r}/(1+M${r})`, f.pct2A, `=L${r}*O${r}`, f.mBase, f.ceBase, f.pct2C, `=R${r}*S${r}`, f.reu,
        f.pctAdelanto, `=N${r}*V${r}`, f.desc, f.nota];
    }));
    ['L', 'N', 'P', 'R', 'T', 'W', 'X'].forEach((l) => sh.getRange(`${l}2:${l}`).setNumberFormat(PESOS));
    ['M', 'O', 'S', 'V'].forEach((l) => sh.getRange(`${l}2:${l}`).setNumberFormat('0.0%'));
  }
  sh.setColumnWidth(6, 220);
  sh.setColumnWidth(25, 320);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(2, filas.length + 1), enc.length).createFilter();
}

function escribirEntregas(ss, filas) {
  const { sh } = hojaLiq(ss, H_ENT);
  sh.clear();
  const enc = ['Mes REU', 'Mes entrega', 'Grupo', 'Orden', 'Cliente', 'Fecha entrega', 'Fecha ingreso (SGA)',
    'Versión entregada (FBS)', 'Código', 'Familia', 'Solicitud', 'Mes del pedido', '% IVA',
    'Precio público c/IVA (mes REU)', 'Precio público c/IVA (mes entrega)', 'Margen comisional', '% adelanto',
    'Recupera adelanto', 'Recupero adelanto margen', 'Margen por pedido (60%)', 'Margen por entrega (40%)', 'Nota'];
  encabezado(sh, 1, enc);
  sh.setFrozenRows(1);
  sh.getRange('A:B').setNumberFormat('@');
  if (filas.length) {
    sh.getRange(2, 1, filas.length, enc.length).setValues(filas.map((f, i) => {
      const r = i + 2;
      return [f.mesReu, f.mesEnt, f.grupo, f.orden, f.cliente, f.entrega || '', f.ingreso || '', f.modeloTxt, f.codigo,
        f.fam, f.sol, f.mesPed, f.iva, f.pubReu, f.pubEnt, f.margen, f.adelanto, f.recupera,
        `=IF(R${r}="Sí",-N${r}/(1+M${r})*P${r}*Q${r},0)`, `=N${r}/(1+M${r})*P${r}*Q${r}`,
        `=O${r}/(1+M${r})*P${r}*(1-Q${r})`, f.nota];
    }));
    sh.getRange('F2:G').setNumberFormat('dd/MM/yyyy');
    ['N', 'O', 'S', 'T', 'U'].forEach((l) => sh.getRange(`${l}2:${l}`).setNumberFormat(PESOS));
    ['M', 'P', 'Q'].forEach((l) => sh.getRange(`${l}2:${l}`).setNumberFormat('0.0%'));
  }
  sh.setColumnWidth(5, 220);
  sh.setColumnWidth(8, 240);
  sh.setColumnWidth(22, 320);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(2, filas.length + 1), enc.length).createFilter();
}

// ---------------------------------------------------------------------------
// RESUMEN (fórmulas: cambia solo al elegir otro mes)
// ---------------------------------------------------------------------------

function armarResumen(ss, d) {
  const { sh } = hojaLiq(ss, H_RESUMEN);
  const mesActual = textoMes(sh.getRange('C3').getValue());
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
  sh.setHiddenGridlines(true);
  const meses = ss.getSheetByName(H_MESES).getRange('A2:A').getValues().map((f) => f[0]).filter(String);
  const sugerido = mesActual && meses.indexOf(mesActual) >= 0 ? mesActual : meses[meses.length - 1] || '';

  sh.getRange('B1').setValue('Liquidación FIAT - PDA').setFontSize(16).setFontWeight('bold').setFontColor(COLOR_ENC);
  sh.getRange('B3').setValue('Mes').setFontWeight('bold');
  sh.getRange('C3').setNumberFormat('@').setValue(sugerido).setFontWeight('bold').setBackground('#fff2cc')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(meses.length ? meses : [''], true).build());
  const M = '$C$3';
  const P = `'${H_PED}'!`;
  const E = `'${H_ENT}'!`;
  const MS = `'${H_MESES}'!`;
  const F = `'${H_FIAT}'!`;
  const dato = (col) => `=IFERROR(INDEX(${MS}${col}:${col},MATCH(${M},${MS}A:A,0)),"")`;
  sh.getRange('E3:L3').setValues([['Categoría', dato('B'), 'Objetivo', dato('C'), 'Netos', dato('D'),
    'Pedidos adicionales', dato('J')]]);
  sh.getRange('E4:L4').setValues([['Cumplió pedidos', dato('E'), 'Suscripciones', dato('F'), 'Patentamientos',
    dato('G'), '', '']]);
  ['E3', 'G3', 'I3', 'K3', 'E4', 'G4', 'I4'].forEach((c) => sh.getRange(c).setFontColor(GRIS));

  const filaEnc = 6;
  sh.getRange(filaEnc, 2, 1, 6).setValues([['Concepto (liquidación Fiat)', 'Calculado', 'Liquidado Fiat',
    'Diferencia', 'Unidades', 'Cómo se calcula']]).setFontWeight('bold').setBackground(COLOR_ENC).setFontColor('white');
  const fiat = (c) => `=IFERROR(SUMIFS(${F}C:C,${F}A:A,${M},${F}B:B,"${c}"),0)`;
  const lineas = [
    ['Adelanto Margen', `=SUMIFS(${P}W:W,${P}A:A,${M})`, fiat('Adelanto Margen'),
      `=COUNTIFS(${P}A:A,${M},${P}V:V,">0")`, 'Pedidos del mes sin REU en el mes: 14% × 60% del precio público s/IVA (lista del mes del pedido)'],
    ['Recupero Adelanto Margen', `=SUMIFS(${E}S:S,${E}A:A,${M})`, fiat('Recupero Adelanto Margen'),
      `=COUNTIFS(${E}A:A,${M},${E}R:R,"Sí")`, 'REU del mes con pedido de un mes anterior: −14% × 60% a precio del mes del REU'],
    ['Márgen Por Pedido (60%)', `=SUMIFS(${E}T:T,${E}A:A,${M})`, fiat('Márgen Por Pedido (60%)'),
      `=COUNTIFS(${E}A:A,${M})`, 'Unidades con REU en el mes: 14% × 60% a precio del mes del REU'],
    ['Margen por Entrega (40%)', `=SUMIFS(${E}U:U,${E}A:A,${M})`, fiat('Margen por Entrega (40%)'),
      `=COUNTIFS(${E}A:A,${M})`, 'Unidades con REU en el mes: 14% × 40% a precio del mes de entrega'],
    ['Bonus Pedidos Aceptados Fiat', `=SUMIFS(${P}P:P,${P}A:A,${M})+IFERROR(INDEX(${MS}L:L,MATCH(${M},${MS}A:A,0)),0)`,
      fiat('Bonus Pedidos Aceptados Fiat'), `=COUNTIFS(${P}A:A,${M})`,
      '2A (% categoría × precio público c/IVA del pedido) + 2B (% categoría × CE s/IVA NC1 × pedidos adicionales)'],
    ['Bonus Pedidos por familia (2C)', `=SUMIFS(${P}T:T,${P}A:A,${M})`, fiat('Bonus Pedidos por familia (2C)'),
      `=COUNTIFS(${P}A:A,${M},${P}T:T,">0")`, 'Inc pedido del modelo pedido × CE s/IVA del modelo base de la familia'],
  ];
  const f0 = filaEnc + 1;
  sh.getRange(f0, 2, lineas.length, 6).setValues(lineas.map((l, i) => {
    const r = f0 + i;
    return [l[0], l[1], l[2], `=C${r}-D${r}`, l[3], l[4]];
  }));
  const ft = f0 + lineas.length;
  sh.getRange(ft, 2, 1, 6).setValues([['TOTAL', `=SUM(C${f0}:C${ft - 1})`, `=SUM(D${f0}:D${ft - 1})`,
    `=C${ft}-D${ft}`, '', '']]).setFontWeight('bold').setBackground(FONDO_TOT);
  sh.getRange(f0, 3, lineas.length + 1, 3).setNumberFormat(PESOS);
  sh.getRange(f0, 7, lineas.length, 1).setFontSize(9).setFontColor(GRIS).setWrap(true);
  sh.getRange(ft + 1, 2).setValue('Fondo de contingencia, incentivos por cuota (mora) y retribución a la suscripción: ' +
    'no se calculan acá.').setFontSize(9).setFontStyle('italic').setFontColor(GRIS);

  // 2C por familia.
  const f2 = ft + 3;
  sh.getRange(f2, 2, 1, 5).setValues([['2C por familia', 'Pedidos con 2C', 'Calculado', 'Liquidado Fiat', 'Diferencia']])
    .setFontWeight('bold').setBackground(COLOR_ENC).setFontColor('white');
  const fams = ss.getSheetByName(H_FAMILIAS).getRange('A2:A').getValues().map((f) => f[0]).filter(String);
  if (fams.length) {
    sh.getRange(f2 + 1, 2, fams.length, 5).setValues(fams.map((fa, i) => {
      const r = f2 + 1 + i;
      return [fa, `=COUNTIFS(${P}A:A,${M},${P}K:K,B${r},${P}T:T,">0")`, `=SUMIFS(${P}T:T,${P}A:A,${M},${P}K:K,B${r})`,
        `=IFERROR(SUMIFS(${F}C:C,${F}A:A,${M},${F}B:B,"Bonus Pedidos por familia (2C)",${F}E:E,B${r}),0)`, `=D${r}-E${r}`];
    }));
    sh.getRange(f2 + 1, 4, fams.length, 3).setNumberFormat(PESOS);
  }
  const f3 = f2 + fams.length + 2;
  sh.getRange(f3, 2).setValue('Entregas pendientes de REU (en FBS, todavía no en SGA)').setFontWeight('bold');
  sh.getRange(f3 + 1, 2).setFormula(`=IFERROR(TEXTJOIN("  ·  ",TRUE,FILTER(${E}C2:C&"-"&${E}D2:D&" "&${E}E2:E&" ("&${E}I2:I&", entrega "&TEXT(${E}F2:F,"dd/mm")&")",${E}A2:A="PENDIENTE REU")),"—")`)
    .setWrap(true).setFontSize(9);
  sh.getRange(f3 + 1, 2, 1, 6).merge();
  sh.setRowHeight(f3 + 1, 60);
  sh.getRange(f2, 2, 1, 5);
  sh.setColumnWidth(1, 16);
  sh.setColumnWidth(2, 250);
  [3, 4, 5].forEach((c) => sh.setColumnWidth(c, 140));
  sh.setColumnWidth(6, 80);
  sh.setColumnWidth(7, 420);
  sh.getRange('E3:L4').setHorizontalAlignment('center');
}

// ---------------------------------------------------------------------------
// Mail
// ---------------------------------------------------------------------------

function enviarResumen() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const res = ss.getSheetByName(H_RESUMEN);
  const r1 = ui.prompt('Enviar resumen por mail', `Mes (Aceptar sin escribir = ${res.getRange('C3').getValue()}).`,
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const mes = textoMes(r1.getResponseText()) || textoMes(res.getRange('C3').getValue());
  res.getRange('C3').setValue(mes);
  const guardados = props.getProperty('MAILS_LIQ') || '';
  const r2 = ui.prompt('Enviar resumen por mail', 'Mails separados por coma.' +
    (guardados ? `\nAceptar sin escribir = ${guardados}` : ''), ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const mails = (r2.getResponseText().trim() || guardados).split(/[,;\s]+/).filter((m) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m));
  if (!mails.length) { ui.alert('No hay ningún mail válido.'); return; }
  props.setProperty('MAILS_LIQ', mails.join(', '));
  SpreadsheetApp.flush();
  const filas = res.getRange('B7:E13').getValues();
  const fmt = (v) => '$ ' + Math.round(Number(v) || 0).toLocaleString('es-AR');
  const tabla = filas.map((f) => `<tr><td style="padding:3px 10px">${f[0]}</td><td style="padding:3px 10px;text-align:right">` +
    `${fmt(f[1])}</td><td style="padding:3px 10px;text-align:right">${f[2] ? fmt(f[2]) : '—'}</td></tr>`).join('');
  const archivo = DriveApp.getFileById(ss.getId());
  mails.forEach((m) => { try { archivo.addViewer(m); } catch (e) { /* sin cuenta Google: igual recibe el mail */ } });
  MailApp.sendEmail({
    to: mails.join(','),
    subject: `Liquidación FIAT - PDA - ${mes}`,
    htmlBody: `<p>Liquidación FIAT - PDA de <b>${mes}</b>.</p><table style="border-collapse:collapse">` +
      '<tr><th style="text-align:left;padding:3px 10px">Concepto</th><th style="padding:3px 10px">Calculado</th>' +
      `<th style="padding:3px 10px">Liquidado Fiat</th></tr>${tabla}</table>` +
      `<p><a href="${ss.getUrl()}">Ver el detalle</a></p>`,
  });
  ui.alert(`Resumen de ${mes} enviado a ${mails.join(', ')}.`);
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

function num(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[^\d,.%-]/g, '');
  const n = Number(s.replace(/[,%]/g, '')) || 0;
  return /%$/.test(s) ? n / 100 : n;
}

function textoMes(v) {
  if (v instanceof Date) return etiquetaMes(new Date(v.getFullYear(), v.getMonth(), v.getDate())).replace(/ \d+$/,
    ` ${String(v.getDate()).padStart(2, '0')}`);
  return v === null || v === undefined ? '' : String(v).trim().toUpperCase();
}

function etiquetaMes(fecha) {
  return `${MESES_NOMBRE[fecha.getMonth()]} ${String(fecha.getFullYear() % 100).padStart(2, '0')}`;
}

function ordenMes(m) {
  const x = String(m).match(/^([A-Z]+) (\d{2})$/);
  return x ? Number(x[2]) * 12 + MESES_NOMBRE.indexOf(x[1]) : 99999;
}

/** "2026-09-28 00:00:00", "28/9/2026" o Date -> Date (o null). */
function aFecha(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return null;
}

function claveGO(g, o) {
  const a = String(g || '').trim();
  const b = String(o || '').trim();
  return /^\d+$/.test(a) && /^\d+$/.test(b) ? `${Number(a)}-${Number(b)}` : '';
}

function normal(s) {
  return String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();
}

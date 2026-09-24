/**
 * CONCILIACIÓN BANCARIA — Banco Macro vs. FBS (cuentas E + O)
 *
 * Hojas de ENTRADA (se pegan los archivos tal cual, desde A1):
 *   - "Extracto"               : extracto de Macro (Excel "Últimos Movimientos" o CSV "Resumen").
 *   - "Mayor E" / "Mayor O"    : exportación a Excel del mayor de FBS de cada cuenta.
 *   - "Pendientes anteriores"  : partidas pendientes del cierre anterior (Sector, Fecha, Concepto, Importe).
 *   - "Reglas"                 : catálogo de conceptos del banco (código causal -> categoría). Editable.
 *   - "Empresas grupo"         : CUIT de las empresas del grupo (sus movimientos son internos).
 *
 * Hojas de SALIDA (se regeneran en cada ejecución):
 *   - "Tablero"                    : saldos FBS, las 4 puntas y el control contra el extracto.
 *   - "Pend. registrar en FBS"     : movimientos del banco sin registro en FBS (puntas 1 y 2).
 *   - "Pend. FBS sin banco"        : registros de FBS que el banco no muestra (puntas 3 y 4).
 *   - "Conciliados"                : cada cruce con el método usado (para auditar).
 *   - "Historial"                  : una fila por mes cerrado, con los saldos y el link a la foto.
 *
 * Criterios de cruce (importe y sentido siempre iguales, tolerancia $0,02), en orden:
 *   1. CUIT/DNI  2. Referencia del banco en el comprobante FBS  3. Apellido
 *   + agrupados del mismo CUIT/apellido y asiento mensual de gastos bancarios
 *   4. Fecha + importe único  + lotes del día (sueldos, VEP)
 *   5. Solo importe a <=10 días (Sugerido)  + combinaciones de 2-3 partidas (revisar)
 * Los pases entre E y O y los registros de FBS que se compensan entre sí no son movimientos.
 *
 * Instalación: Extensiones > Apps Script, pegar este archivo, guardar y recargar la planilla.
 * Menú "Conciliación" > "1. Crear hojas de entrada" (una sola vez), pegar los archivos del mes y
 * "2. Procesar conciliación". Al cerrar el mes: "3. Cerrar mes" guarda una copia completa de la planilla
 * (foto) en la carpeta "Conciliaciones - Historial", la registra en la hoja "Historial" y pasa los
 * pendientes a "Pendientes anteriores".
 */

// ------------------------------------------------------------------ configuración

const HOJA = {
  extracto: 'Extracto',
  mayorE: 'Mayor E',
  mayorO: 'Mayor O',
  anteriores: 'Pendientes anteriores',
  reglas: 'Reglas',
  empresas: 'Empresas grupo',
  tablero: 'Tablero',
  pendBanco: 'Pend. registrar en FBS',
  pendFbs: 'Pend. FBS sin banco',
  conciliados: 'Conciliados',
  historial: 'Historial',
  foto: 'FOTO',
};

const SECTORES = [
  ['S1', 'Depósitos no registrados en contabilidad', '+'],
  ['S2', 'Débitos no registrados en contabilidad', '-'],
  ['S3', 'Depósitos no acreditados en banco', '-'],
  ['S4', 'Cheques / pagos no debitados en banco', '+'],
];

const VENTANA_DIAS = 10;
const VENTANA_SUGERIDO = 10;
const VENTANA_ID = 31;
const TOLERANCIA = 0.02;
const NUM_FMT = '#,##0.00;[Red]-#,##0.00';
const CAUSALES_IMPUESTOS = ['1684', '1685', '1297', '4145'];
const PALABRAS_VACIAS = new Set(['SA', 'SRL', 'SAS', 'DE', 'LA', 'EL', 'BANCO', 'MACRO', 'TRANSF', 'VAR', 'VARIOS',
  'CUO', 'CUOTAS', 'FAC', 'FACTURAS', 'TEF', 'DATANET', 'ING', 'CIRC', 'CERRADO', 'CCERR']);

const EMPRESAS_INICIALES = [['30669400167', 'TURIN SA'], ['30681026386', 'CHEXA SA']];

// Código causal | Sentido (Débito/Crédito/Ambos) | Patrón en concepto (regex, vacío = todos) | Categoría | Descripción
const REGLAS_INICIALES = [
  ["493", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)"],
  ["893", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)"],
  ["4397", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida (solo CUIT)"],
  ["3816", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)"],
  ["4543", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)"],
  ["4544", "Ambos", "", "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)"],
  ["4080", "Ambos", "", "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual"],
  ["4081", "Ambos", "", "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual"],
  ["4091", "Ambos", "", "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual"],
  ["4013", "Ambos", "", "COBRANZA CLIENTES", "CREDIN (crédito inmediato)"],
  ["4016", "Ambos", "", "COBRANZA CLIENTES", "CREDIN (crédito inmediato)"],
  ["2010", "Ambos", "", "COBRANZA CLIENTES", "Transferencia minorista (TRMIN)"],
  ["366", "Ambos", "", "A IDENTIFICAR", "Crédito 'DL1' con CUIT persona"],
  ["4333", "Ambos", "", "A IDENTIFICAR", "Crédito 'Número de Operación'"],
  ["4334", "Ambos", "", "A IDENTIFICAR", "Crédito 'Número de Operación'"],
  ["4218", "Ambos", "", "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'"],
  ["4220", "Ambos", "", "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'"],
  ["2027", "Ambos", "PROVEEDORES", "PAGO PROVEEDORES", "Transferencia Datanet a proveedores"],
  ["2027", "Ambos", "TEF DATANET PR", "COBRANZA EMPRESAS / ASEGURADORAS", "TEF Datanet recibida de empresa"],
  ["4015", "Ambos", "EGRESO", "PAGO PROVEEDORES", "Egreso transferencia inmediata"],
  ["3862", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR)"],
  ["3956", "Ambos", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR) San Cristóbal"],
  ["2024", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Pago AFIP/ARCA vía Datanet"],
  ["2018", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet misma titularidad (M/T)"],
  ["2026", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet BTOB"],
  ["3808", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia desde CHEXA SA"],
  ["4254", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata CHEXA SA"],
  ["4267", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata a CHEXA SA"],
  ["1411", "Ambos", "", "INVERSIONES (FCI)", "Suscripción FCI"],
  ["4361", "Ambos", "", "INVERSIONES (FCI)", "Suscripción FCI"],
  ["1413", "Ambos", "", "INVERSIONES (FCI)", "Rescate FCI"],
  ["4196", "Ambos", "PLAN AHORRO", "PLAN DE AHORRO / TERMINAL", "Débito Plan de Ahorro Chevrolet"],
  ["1445", "Ambos", "", "PLAN DE AHORRO / TERMINAL", "Transferencia recibida de Chevrolet (TRPRO)"],
  ["1952", "Ambos", "", "A IDENTIFICAR", "PP- N/C Deuda Pública - P.Previo"],
  ["1684", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/débitos"],
  ["1685", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/créditos"],
  ["1297", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Retención IIBB SIRCREB"],
  ["4145", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Retención IIBB Córdoba renta financiera"],
  ["4196", "Ambos", "IMP\\. AFIP", "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)"],
  ["3696", "Ambos", "", "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)"],
  ["23", "Ambos", "ARCA|AFIP", "IMPUESTOS Y RETENCIONES", "Débito directo ARCA"],
  ["531", "Ambos", "", "GASTOS BANCARIOS", "Comisión transferencia Datanet + IVA"],
  ["1210", "Ambos", "", "GASTOS BANCARIOS", "Comisión convenio recaudación 61068"],
  ["1211", "Ambos", "", "GASTOS BANCARIOS", "IVA / percepción convenio recaudación 61068"],
  ["17", "Ambos", "", "GASTOS BANCARIOS", "Comisión resumen cuenta + IVA"],
  ["33", "Ambos", "", "GASTOS BANCARIOS", "Comisión mantenimiento cuenta + IVA"],
  ["3872", "Ambos", "", "GASTOS BANCARIOS", "Comisión transferencia"],
  ["1450", "Ambos", "", "GASTOS BANCARIOS", "Comisión transferencia"],
  ["1693", "Ambos", "", "SUELDOS", "Pago de remuneraciones"],
  ["23", "Ambos", "", "DEBITOS AUTOMATICOS", "Débito automático (seguros, peajes, cámara)"],
  ["1735", "Ambos", "", "DEPOSITOS EN EFECTIVO", "Interdepósito en efectivo"],
  ["4061", "Ambos", "", "FINANCIACION / PRESTAMOS", "Transf. 'Presta' Banco del Sol"],
  ["4066", "Ambos", "", "A IDENTIFICAR", "Transf. Varios VAZCAR SAS"],
  ["3862", "Débito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transf. MacrOnline 'CCDO DIST T'"],
  ["3912", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transf. MacrOnline E-set M/T (misma titularidad)"],
  ["3913", "Ambos", "", "PAGO PROVEEDORES", "Transf. MacrOnline E-set D/T (distinta titularidad)"],
  ["4085", "Débito", "", "PAGO PROVEEDORES", "Transferencia inmediata enviada"],
  ["4250", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "CREDIN propio"],
  ["4253", "Ambos", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata propia"],
  ["183", "Crédito", "", "COBRANZA TARJETAS", "Liquidación Payway (tarjetas)"],
  ["183", "Débito", "", "COBRANZA TARJETAS", "Débito liquidación Payway (contracargo/ajuste)"],
  ["14", "Ambos", "", "CHEQUES", "Acreditación cheques depositados (remesas)"],
  ["85", "Ambos", "", "CHEQUES", "Cheque propio pagado por cámara"],
  ["2837", "Ambos", "", "CHEQUES", "Cheque propio pagado por canje interno"],
  ["1232", "Ambos", "", "A IDENTIFICAR", "Crédito con referencia numérica (30 dígitos)"],
  ["16", "Ambos", "IVA", "GASTOS BANCARIOS", "IVA comisión valores al cobro"],
  ["16", "Ambos", "", "GASTOS BANCARIOS", "Comisión administración valores al cobro"],
  ["1724", "Ambos", "", "GASTOS BANCARIOS", "Comisión depósito cheque otra sucursal + IVA"],
  ["574", "Ambos", "", "GASTOS BANCARIOS", "Comisión cheque pagado por clearing + IVA"],
  ["3914", "Ambos", "", "GASTOS BANCARIOS", "Comisión transferencia MacrOnline"],
  ["3734", "Ambos", "", "GASTOS BANCARIOS", "Comisión débito pago remuneraciones"],
  ["3205", "Ambos", "", "GASTOS BANCARIOS", "IVA comisión"],
  ["1212", "Ambos", "", "GASTOS BANCARIOS", "Mantenimiento caja de seguridad"],
  ["1214", "Ambos", "", "GASTOS BANCARIOS", "IVA caja de seguridad"],
  ["1215", "Ambos", "", "GASTOS BANCARIOS", "Percepción IVA caja de seguridad"]
];

// ------------------------------------------------------------------ menú

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Conciliación')
    .addItem('1. Crear hojas de entrada', 'crearHojasEntrada')
    .addItem('2. Procesar conciliación', 'procesarConciliacion')
    .addItem('3. Cerrar mes (guardar foto y pasar pendientes)', 'cerrarMes')
    .addToUi();
}

function crearHojasEntrada() {
  const ss = SpreadsheetApp.getActive();
  const crear = (nombre, filas) => {
    let sh = ss.getSheetByName(nombre);
    if (sh) return sh;
    sh = ss.insertSheet(nombre);
    if (filas && filas.length) {
      sh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
      sh.getRange(1, 1, 1, filas[0].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  crear(HOJA.extracto, [['Pegar acá el extracto de Macro desde A1 (Excel "Últimos Movimientos" o CSV "Resumen")']]);
  crear(HOJA.mayorE, [['Pegar acá la exportación del mayor de FBS de la cuenta E desde A1']]);
  crear(HOJA.mayorO, [['Pegar acá la exportación del mayor de FBS de la cuenta O desde A1']]);
  crear(HOJA.anteriores, [['Sector', 'Fecha', 'Concepto', 'Importe']]);
  crear(HOJA.reglas, [['Código causal', 'Sentido', 'Patrón en concepto', 'Categoría', 'Descripción']].concat(REGLAS_INICIALES));
  crear(HOJA.empresas, [['CUIT', 'Razón social']].concat(EMPRESAS_INICIALES));
  SpreadsheetApp.getUi().alert('Hojas de entrada listas. Pegá el extracto y los mayores E y O, y corré "2. Procesar conciliación".');
}

// ------------------------------------------------------------------ proceso principal (capa de hojas)

function procesarConciliacion() {
  const ss = SpreadsheetApp.getActive();
  const res = procesar_(ss);
  const dif = res.ajustado - res.saldoBanco;
  ss.toast('Diferencia de control: ' + formato_(dif) + (res.avisos.length ? ' — ver avisos en el Tablero' : ''),
    'Conciliación terminada', 10);
}

function procesar_(ss) {
  if (ss.getSheetByName(HOJA.foto)) {
    throw new Error('Esta planilla es una FOTO de un cierre y no se puede reprocesar. Trabajá en la planilla principal.');
  }
  const leer = nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) throw new Error('Falta la hoja "' + nombre + '". Corré "1. Crear hojas de entrada".');
    return sh.getDataRange().getValues();
  };
  const res = conciliarTodo_({
    extracto: leer(HOJA.extracto),
    mayorE: leer(HOJA.mayorE),
    mayorO: leer(HOJA.mayorO),
    anteriores: leer(HOJA.anteriores),
    reglas: leer(HOJA.reglas),
    empresas: leer(HOJA.empresas),
  });
  escribirTablero_(ss, res);
  escribirPendientes_(ss, HOJA.pendBanco, res, ['S1', 'S2']);
  escribirPendientes_(ss, HOJA.pendFbs, res, ['S3', 'S4']);
  escribirConciliados_(ss, res);
  ss.getSheetByName(HOJA.tablero).activate();
  SpreadsheetApp.flush();
  return res;
}

/**
 * Cierre del mes:
 *  1. reprocesa (para que la foto refleje exactamente lo que hay cargado),
 *  2. guarda una COPIA COMPLETA de la planilla (entradas + tablero + pendientes) en la carpeta de historial,
 *     protegida y marcada como foto para que nadie la reprocese,
 *  3. agrega una fila en "Historial" con los saldos y el link a la foto,
 *  4. pasa los pendientes a "Pendientes anteriores" y, si se confirma, vacía extracto y mayores.
 */
function cerrarMes() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const res = procesar_(ss);
  const periodo = res.corte.getFullYear() + '-' + ('0' + (res.corte.getMonth() + 1)).slice(-2);
  const empresa = res.info.empresa || '';
  const dif = redondear_(res.ajustado - res.saldoBanco);

  const hist = hojaHistorial_(ss);
  const previos = hist.getLastRow() > 1 ? hist.getRange(2, 1, hist.getLastRow() - 1, 2).getValues() : [];
  const yaCerrado = previos.some(r => String(r[0]) === periodo && String(r[1]) === empresa);
  let msg = 'Cerrar ' + periodo + (empresa ? ' — ' + empresa : '') + '\n\nDiferencia de control: ' + formato_(dif) +
    '\n\nSe va a guardar una foto de la conciliación en Drive y los pendientes pasan al mes siguiente.';
  if (yaCerrado) msg += '\n\nATENCIÓN: este período ya tiene un cierre en "Historial". Se guarda una foto nueva (versión 2, 3...).';
  if (Math.abs(dif) >= 1000) msg += '\n\nATENCIÓN: la diferencia de control no es cero.';
  if (ui.alert('Cerrar mes', msg + '\n\n¿Continuar?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  // 2. foto
  const carpeta = carpetaHistorial_(ss);
  let nombre = 'Conciliación Macro' + (empresa ? ' - ' + empresa : '') + ' - ' + periodo;
  const version = previos.filter(r => String(r[0]) === periodo && String(r[1]) === empresa).length + 1;
  if (version > 1) nombre += ' (v' + version + ')';
  const archivo = DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);
  marcarComoFoto_(SpreadsheetApp.openById(archivo.getId()), periodo, empresa);

  // 3. historial
  hist.appendRow([periodo, empresa, new Date(), Session.getActiveUser().getEmail(), res.saldoE, res.saldoO,
    res.saldoE + res.saldoO, res.tot.S1, res.tot.S2, res.tot.S3, res.tot.S4, res.ajustado, res.saldoBanco, dif,
    res.apertura.diferencia, archivo.getUrl()]);
  const fila = hist.getLastRow();
  hist.getRange(fila, 3).setNumberFormat('dd/mm/yyyy hh:mm');
  hist.getRange(fila, 5, 1, 11).setNumberFormat(NUM_FMT);
  hist.getRange(fila, 16).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Abrir foto')
    .setLinkUrl(archivo.getUrl()).build());

  // 4. pendientes al mes siguiente
  const n = pasarPendientes_(ss);
  if (ui.alert('Mes cerrado', 'Foto guardada en la carpeta "' + carpeta.getName() + '" y registrada en "Historial".\n' + n +
      ' partidas pasadas a "' + HOJA.anteriores + '".\n\n¿Vaciar las hojas Extracto, Mayor E y Mayor O para cargar el mes siguiente?',
      ui.ButtonSet.YES_NO) === ui.Button.YES) {
    [HOJA.extracto, HOJA.mayorE, HOJA.mayorO].forEach(h => { const sh = ss.getSheetByName(h); if (sh) sh.clearContents(); });
  }
  hist.activate();
}

function hojaHistorial_(ss) {
  let sh = ss.getSheetByName(HOJA.historial);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA.historial);
  const cab = ['Período', 'Empresa', 'Fecha de cierre', 'Usuario', 'Saldo E', 'Saldo O', 'Saldo FBS', 'Dep. no registrados',
    'Déb. no registrados', 'Dep. no acreditados', 'Pagos no debitados', 'Saldo ajustado', 'Saldo extracto',
    'Diferencia control', 'Diferencia apertura', 'Foto'];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

function carpetaHistorial_(ss) {
  const nombre = 'Conciliaciones - Historial';
  const padres = DriveApp.getFileById(ss.getId()).getParents();
  const padre = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
  const it = padre.getFoldersByName(nombre);
  if (it.hasNext()) return it.next();
  try {
    return padre.createFolder(nombre);
  } catch (e) {
    // sin permiso de edición en la carpeta de la planilla: se usa "Mi unidad"
    const raiz = DriveApp.getRootFolder();
    const r = raiz.getFoldersByName(nombre);
    SpreadsheetApp.getActive().toast('No tengo permiso para crear carpetas en "' + padre.getName() +
      '". La foto se guarda en "Mi unidad / ' + nombre + '".', 'Historial', 10);
    return r.hasNext() ? r.next() : raiz.createFolder(nombre);
  }
}

/** Deja la copia como foto: hoja "FOTO" con los datos del cierre y todas las hojas protegidas (solo el dueño edita). */
function marcarComoFoto_(copia, periodo, empresa) {
  const info = copia.insertSheet(HOJA.foto, 0);
  info.getRange(1, 1, 5, 2).setValues([
    ['FOTO DE CIERRE — no modificar', ''],
    ['Período', periodo],
    ['Empresa', empresa],
    ['Fecha de cierre', new Date()],
    ['Cerrado por', Session.getActiveUser().getEmail()],
  ]);
  info.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  info.getRange(4, 2).setNumberFormat('dd/mm/yyyy hh:mm');
  const yo = Session.getEffectiveUser();
  copia.getSheets().forEach(sh => {
    const p = sh.protect().setDescription('Foto de cierre ' + periodo);
    p.addEditor(yo);
    p.removeEditors(p.getEditors().filter(u => u.getEmail() !== yo.getEmail()));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
  const tablero = copia.getSheetByName(HOJA.tablero);
  (tablero || info).activate();
}

function pasarPendientes_(ss) {
  const filas = [['Sector', 'Fecha', 'Concepto', 'Importe']];
  [HOJA.pendBanco, HOJA.pendFbs].forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) throw new Error('Primero corré "2. Procesar conciliación".');
    const v = sh.getDataRange().getValues();
    const h = v[1];   // fila 1 = título, fila 2 = encabezados
    const c = { s: h.indexOf('Sector'), f: h.indexOf('Fecha'), t: h.indexOf('Concepto / Comprobante'), i: h.indexOf('Importe') };
    v.slice(2).forEach(r => { if (r[c.s]) filas.push([r[c.s], r[c.f], r[c.t], r[c.i]]); });
  });
  const sh = ss.getSheetByName(HOJA.anteriores) || ss.insertSheet(HOJA.anteriores);
  sh.clear();
  sh.getRange(1, 1, filas.length, 4).setValues(filas);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.getRange(2, 4, Math.max(filas.length - 1, 1), 1).setNumberFormat(NUM_FMT);
  return filas.length - 1;
}

// ------------------------------------------------------------------ salida

function hojaLimpia_(ss, nombre) {
  let sh = ss.getSheetByName(nombre);
  if (sh) { sh.clear(); sh.clearConditionalFormatRules(); if (sh.getFilter()) sh.getFilter().remove(); }
  else sh = ss.insertSheet(nombre);
  return sh;
}

function escribirTablero_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.tablero);
  const filas = [];
  const estilos = [];   // 'titulo' | 'cab' | 'total' | 'dif' | ''
  const add = (fila, estilo) => { while (fila.length < 3) fila.push(''); filas.push(fila); estilos.push(estilo || ''); };
  add(['CONCILIACIÓN BANCO MACRO' + (res.info.empresa ? ' — ' + res.info.empresa : '') + ' — al ' + fechaTexto_(res.corte)], 'titulo');
  add([]);
  add(['Concepto', 'Importe', 'Partidas'], 'cab');
  add(['Cuenta E  ' + (res.infoE.cuenta || ''), res.saldoE]);
  add(['Cuenta O  ' + (res.infoO.cuenta || ''), res.saldoO]);
  add(['SALDO CONTABLE FBS (E + O)', res.saldoE + res.saldoO], 'total');
  add([]);
  SECTORES.forEach(([s, nombre, signo]) => {
    const hoja = (s === 'S1' || s === 'S2') ? HOJA.pendBanco : HOJA.pendFbs;
    add(['(' + signo + ') ' + nombre + '  — ver "' + hoja + '"', res.tot[s], res.pend[s].length]);
  });
  add([]);
  add(['SALDO CONTABLE AJUSTADO', res.ajustado], 'total');
  add(['SALDO SEGÚN EXTRACTO al ' + fechaTexto_(res.corte), res.saldoBanco], 'total');
  add(['DIFERENCIA (control)', res.ajustado - res.saldoBanco], 'dif');
  add([]);
  add(['Control de apertura (inicio del mes)', 'Importe', ''], 'cab');
  add(['  Saldo inicial FBS (E + O)', res.apertura.fbs]);
  add(['  (+/-) Pendientes anteriores (neto)', res.apertura.pendientes]);
  add(['  Saldo inicial del extracto', res.apertura.banco]);
  add(['DIFERENCIA DE APERTURA', res.apertura.diferencia], 'dif');
  add([]);
  add(['Calidad del cruce — movimientos del extracto del mes', 'Importe', 'Cant.'], 'cab');
  Object.keys(res.metodos).sort().forEach(k => add(['  ' + k, res.metodos[k][1], res.metodos[k][0]]));
  if (res.difGastos !== null) add(['  Diferencia asiento gastos bancarios (banco − FBS)', res.difGastos]);
  add([]);
  add(['Gastos e impuestos bancarios del mes (para el asiento)', 'Importe', ''], 'cab');
  res.gastos.forEach(g => add(['  ' + g[0], g[1]]));
  if (res.avisos.length) {
    add([]);
    add(['Avisos', '', ''], 'cab');
    res.avisos.forEach(a => add(['  ' + a]));
  }
  sh.getRange(1, 1, filas.length, 3).setValues(filas);
  sh.getRange(1, 2, filas.length, 1).setNumberFormat(NUM_FMT);
  estilos.forEach((e, i) => {
    const r = sh.getRange(i + 1, 1, 1, 3);
    if (e === 'titulo') r.setFontWeight('bold').setFontSize(14);
    if (e === 'cab') r.setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    if (e === 'total') r.setFontWeight('bold').setBorder(true, null, true, null, null, null);
    if (e === 'dif') r.setFontWeight('bold').setBackground(Math.abs(filas[i][1]) < 1000 ? '#C6EFCE' : '#FFC7CE');
  });
  sh.setColumnWidth(1, 520); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 80);
}

function escribirPendientes_(ss, nombre, res, sectores) {
  const sh = hojaLimpia_(ss, nombre);
  const cab = ['Sector', 'Cuadro', 'Fecha', 'Días', 'Origen', 'Concepto / Comprobante', 'Importe', 'Categoría',
    'CUIT/DNI', 'Ref. banco / Asiento FBS'];
  const nombres = {};
  SECTORES.forEach(([s, n]) => { nombres[s] = n; });
  const filas = [];
  sectores.forEach(s => res.pend[s].slice().sort((a, b) => a.fecha - b.fecha).forEach(p => filas.push([
    s, nombres[s], p.fecha, dias_(res.corte, p.fecha), p.origen, p.texto, Math.abs(p.importe), p.cat,
    p.cuit || p.dni, p.asiento || p.ref])));
  const total = sectores.map(s => nombres[s] + ': ' + res.pend[s].length + ' partidas, ' + formato_(res.tot[s])).join('   |   ');
  sh.getRange(1, 1).setValue(total).setFontWeight('bold');
  sh.getRange(2, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  if (filas.length) {
    sh.getRange(3, 9, filas.length, 2).setNumberFormat('@');   // CUIT y referencias como texto
    sh.getRange(3, 1, filas.length, cab.length).setValues(filas.map(f => f.map((v, i) => (i >= 8 ? String(v || '') : v))));
    sh.getRange(3, 3, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(3, 7, filas.length, 1).setNumberFormat(NUM_FMT);
  }
  sh.setFrozenRows(2);
  sh.getRange(2, 1, Math.max(filas.length, 1) + 1, cab.length).createFilter();
  [60, 250, 90, 50, 80, 380, 130, 230, 110, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function escribirConciliados_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.conciliados);
  const cab = ['Método', 'Fecha banco', 'Concepto banco', 'Importe banco', 'Origen banco',
    'Fecha FBS', 'Comprobante FBS', 'Importe FBS', 'Asiento FBS', 'Origen FBS'];
  const filas = res.conciliados;
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  if (filas.length) {
    sh.getRange(2, 1, filas.length, cab.length).setValues(filas);
    sh.getRange(2, 2, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 6, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 4, filas.length, 1).setNumberFormat(NUM_FMT);
    sh.getRange(2, 8, filas.length, 1).setNumberFormat(NUM_FMT);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, filas.length + 1, cab.length).createFilter();
  [200, 90, 300, 120, 70, 90, 350, 120, 90, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

// ------------------------------------------------------------------ utilidades

function normalizar_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

function parseNum_(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).replace(/\$/g, '').replace(/\s/g, '');
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s) || /\)$/.test(s)) { neg = true; s = s.replace(/[()]/g, ''); }
  if (s.startsWith('-')) { neg = !neg; s = s.slice(1); }
  if (!/^[\d.,]+$/.test(s)) return null;
  const coma = s.lastIndexOf(','), punto = s.lastIndexOf('.');
  if (coma >= 0 && punto >= 0) {
    s = coma > punto ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (coma >= 0) {
    s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (punto >= 0 && /^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');
  }
  const n = Number(s);
  return isNaN(n) ? null : (neg ? -n : n);
}

function parseFecha_(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v == null ? '' : v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  let a = Number(m[3]);
  if (a < 100) a += 2000;
  return new Date(a, Number(m[2]) - 1, Number(m[1]));
}

function dias_(a, b) { return Math.round(Math.abs(a - b) / 86400000); }
function redondear_(x) { return Math.round(x * 100) / 100; }
function fechaTexto_(d) { return d ? ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear() : ''; }
function formato_(n) {
  const s = Math.abs(n).toFixed(2).split('.');
  return (n < 0 ? '-' : '') + '$ ' + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + s[1];
}

/** Busca la fila de encabezados y devuelve {fila, col: {clave: índice}}. */
function encabezados_(filas, patrones, requeridos) {
  for (let i = 0; i < Math.min(filas.length, 60); i++) {
    const col = {};
    filas[i].forEach((v, j) => {
      const t = normalizar_(v);
      if (!t) return;
      Object.keys(patrones).forEach(k => { if (col[k] === undefined && patrones[k].test(t)) col[k] = j; });
    });
    if (requeridos.every(k => col[k] !== undefined)) return { fila: i, col: col };
  }
  return null;
}

function cuitValido_(c) {
  const pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
  let s = 0;
  for (let i = 0; i < 10; i++) s += Number(c[i]) * pesos[i];
  let r = 11 - (s % 11);
  if (r === 11) r = 0; else if (r === 10) r = 9;
  return Number(c[10]) === r;
}

function extraerCuit_(texto) {
  const re = /\b(20|23|24|27|30|33|34)(\d{8})(\d)\b/g;
  let m;
  while ((m = re.exec(texto)) !== null) {
    const c = m[1] + m[2] + m[3];
    if (cuitValido_(c)) return c;
  }
  return '';
}

function extraerNombre_(concepto) {
  const pats = [/ING TRANSF:(.+?)-\d{11}/, /^TRANSF (.+?) \d{11}/, /TEF DATANET \w+ (.+?)\s+\d{11}/,
    /^CCERR (.+?) \d{11}/, /^TRMIN (.+?) \d{11}/, /^TRPRO (.+?) \d{11}/, /^\d{11} (.+?) DL1/, /Transf\. \w+ (.+?)\s+\d{11}/];
  for (const p of pats) {
    const m = concepto.match(p);
    if (m) return m[1].replace(/\s+/g, ' ').replace(/^[ /,]+|[ /,]+$/g, '');
  }
  return '';
}

function tokens_(texto) {
  const out = new Set();
  (String(texto || '').toUpperCase().match(/[A-ZÑ]{3,}/g) || []).forEach(t => { if (!PALABRAS_VACIAS.has(t)) out.add(t); });
  return out;
}

function comparten_(a, b) { for (const x of a) if (b.has(x)) return true; return false; }

// ------------------------------------------------------------------ lectura de entradas

function leerReglas_(filas) {
  return filas.slice(1).filter(r => String(r[0]).trim()).map(r => {
    const sen = normalizar_(r[1]);
    const pat = String(r[2] || '').trim();
    return {
      cod: String(r[0]).trim(),
      sentido: sen.startsWith('d') ? 'D' : sen.startsWith('c') ? 'C' : null,
      patron: pat && pat !== '(todos)' ? new RegExp(pat) : null,
      cat: String(r[3] || '').trim(),
      desc: String(r[4] || '').trim(),
    };
  });
}

function leerEmpresas_(filas) {
  const e = {};
  filas.slice(1).forEach(r => { const c = String(r[0]).replace(/\D/g, ''); if (c.length === 11) e[c] = String(r[1]); });
  return e;
}

function clasificar_(reglas, empresas, causal, concepto, importe, cuit) {
  if (cuit && empresas[cuit]) return ['TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO', 'Movimiento con ' + empresas[cuit]];
  const sentido = importe < 0 ? 'D' : 'C';
  for (const r of reglas) {
    if (r.cod === causal && (r.sentido === null || r.sentido === sentido) && (!r.patron || r.patron.test(concepto))) {
      return [r.cat, r.desc];
    }
  }
  return ['A IDENTIFICAR', 'Código causal nuevo'];
}

/** Extracto Macro: "Últimos Movimientos" (Importe con signo) o "Resumen" (Débito / Crédito). */
function leerExtracto_(filas, reglas, empresas) {
  const h = encabezados_(filas, {
    fecha: /^fecha$/, ref: /referencia/, causal: /causal/, concepto: /^concepto/, importe: /^importe/,
    debito: /^debito/, credito: /^credito/, saldo: /^saldo/,
  }, ['fecha', 'concepto', 'saldo']);
  if (!h) throw new Error('Extracto: no encuentro los encabezados (Fecha, Concepto, Saldo).');
  const c = h.col, info = {}, movs = [];
  filas.forEach((r, i) => {
    const t0 = String(r[0] || '');
    if (/^Empresa:/i.test(t0)) info.empresa = t0.split(':').slice(1).join(':').trim();
    if (i <= h.fila) return;
    const fecha = parseFecha_(r[c.fecha]);
    if (!fecha) return;
    let deb, cred;
    if (c.importe !== undefined) {
      const imp = parseNum_(r[c.importe]) || 0;
      deb = Math.max(-imp, 0); cred = Math.max(imp, 0);
    } else {
      deb = Math.abs(parseNum_(r[c.debito]) || 0); cred = Math.abs(parseNum_(r[c.credito]) || 0);
    }
    const concepto = String(r[c.concepto] || '').trim();
    const causal = c.causal !== undefined ? String(r[c.causal]).trim() : '';
    const importe = redondear_(cred - deb);
    const cuit = extraerCuit_(concepto);
    const [cat, desc] = clasificar_(reglas, empresas, causal, concepto, importe, cuit);
    movs.push({ fecha, ref: c.ref !== undefined ? String(r[c.ref]).trim() : '', causal, concepto, importe,
      saldo: parseNum_(r[c.saldo]), cat, desc, cuit, nombre: extraerNombre_(concepto) });
  });
  if (!movs.length) throw new Error('Extracto: no hay movimientos con fecha debajo de los encabezados.');
  if (movs[0].fecha > movs[movs.length - 1].fecha) movs.reverse();   // el banco lo entrega del más nuevo al más viejo
  return { info, movs };
}

/** Mayor FBS exportado a Excel: Asiento, Fecha, Referencia, Comentario, Debe, Haber, Saldo. */
function leerMayor_(filas, nombreHoja) {
  const h = encabezados_(filas, {
    asiento: /asiento/, fecha: /^fecha/, ref: /referencia/, com: /comentario|detalle|descripcion|concepto/,
    debe: /^debe/, haber: /^haber/, saldo: /^saldo/,
  }, ['fecha', 'debe', 'haber']);
  if (!h) throw new Error(nombreHoja + ': no encuentro los encabezados (Fecha, Debe, Haber).');
  const c = h.col, info = {}, asientos = [];
  let saldoInicialTexto = null;
  filas.forEach((r, i) => {
    const txt = r.map(v => String(v == null ? '' : v)).join(' ');
    const mc = txt.match(/Cuenta:\s*(.+?)(\s{2,}|$)/);
    if (mc && !info.cuenta) info.cuenta = mc[1].trim();
    if (/saldo inicial/i.test(txt)) {
      const nums = r.map(parseNum_).filter(x => x !== null && typeof x === 'number');
      if (nums.length) saldoInicialTexto = nums[nums.length - 1];
    }
    if (i <= h.fila) return;
    const fecha = parseFecha_(r[c.fecha]);
    const debe = parseNum_(r[c.debe]) || 0, haber = parseNum_(r[c.haber]) || 0;
    const com = c.com !== undefined ? String(r[c.com] || '').trim() : '';
    if (fecha && (debe || haber)) {
      asientos.push({ asiento: c.asiento !== undefined ? String(r[c.asiento]).trim() : '', fecha,
        referencia: c.ref !== undefined ? String(r[c.ref] || '').trim() : '', comentario: com, debe, haber,
        saldo: c.saldo !== undefined ? parseNum_(r[c.saldo]) : null });
    } else if (!fecha && com && asientos.length && !debe && !haber && !/saldo/i.test(txt)) {
      asientos[asientos.length - 1].comentario += ' ' + com;   // comentario partido en dos filas
    }
  });
  asientos.forEach(a => { a.comentario = a.comentario.replace(/\s+/g, ' ').trim(); });
  // saldo inicial: el más repetido de (saldo de la fila - acumulado); si no hay columna Saldo, la fila "Saldo Inicial"
  const cand = {};
  let acum = 0;
  asientos.forEach(a => {
    acum += a.debe - a.haber;
    if (a.saldo !== null) { const k = redondear_(a.saldo - acum).toFixed(2); cand[k] = (cand[k] || 0) + 1; }
  });
  const claves = Object.keys(cand);
  const avisos = [];
  if (claves.length) info.saldoInicial = Number(claves.sort((a, b) => cand[b] - cand[a])[0]);
  else if (saldoInicialTexto !== null) info.saldoInicial = saldoInicialTexto;
  else { info.saldoInicial = 0; avisos.push(nombreHoja + ': no encontré saldo inicial, se tomó 0.'); }
  info.saldoFinal = redondear_(info.saldoInicial + asientos.reduce((s, a) => s + a.debe - a.haber, 0));
  return { info, asientos, avisos };
}

// ------------------------------------------------------------------ partidas

function partidaBanco_(m, origen) {
  const cuit = m.cuit || extraerCuit_(m.concepto);
  return { lado: 'BANCO', origen: origen || 'Mes', fecha: m.fecha, importe: redondear_(m.importe), texto: m.concepto,
    ref: String(m.ref || ''), cuit, dni: cuit ? cuit.slice(2, 10) : '', nombre: tokens_(m.nombre || m.concepto),
    cat: m.cat || '', causal: m.causal || '', match: null, metodo: '' };
}

function partidaFbs_(a, origen) {
  const com = a.comentario;
  let cuit = '', dni = '';
  const m = (com + ' ' + (a.referencia || '')).match(/TR ?(\d{2})[ -]?(\d{8})[ -]?(\d)\b/);
  if (m) { cuit = m[1] + m[2] + m[3]; dni = m[2]; }
  else {
    const d = com.match(/\b(?:SX|TR) ?(\d{7,8})\b/);
    if (d) dni = ('0' + d[1]).slice(-8);
  }
  const partes = com.split('/').map(p => p.trim());
  return { lado: 'FBS', origen: origen || 'Mes', fecha: a.fecha, importe: redondear_(a.debe - a.haber), texto: com,
    ref: a.referencia || '', asiento: a.asiento || '', cuit, dni,
    nombre: partes.length > 1 ? tokens_(partes[partes.length - 1]) : new Set(), cat: '', causal: '', match: null, metodo: '' };
}

/** Une E y O descartando los pases entre ambas cuentas (mismo asiento e importe con signo opuesto). */
function movimientosFbs_(asientosE, asientosO) {
  const neto = new Map(), primero = new Map();
  asientosE.concat(asientosO).forEach(a => {
    const k = a.asiento + '|' + a.comentario + '|' + redondear_(Math.abs(a.debe - a.haber));
    neto.set(k, (neto.get(k) || 0) + a.debe - a.haber);
    if (!primero.has(k)) primero.set(k, a);
  });
  const out = [];
  neto.forEach((v, k) => {
    if (Math.abs(v) > 0.005) {
      const a = Object.assign({}, primero.get(k));
      a.debe = Math.max(v, 0); a.haber = Math.max(-v, 0);
      out.push(a);
    }
  });
  return out;
}

/** Hoja "Pendientes anteriores": Sector (S1..S4), Fecha, Concepto, Importe (positivo). */
function leerAnteriores_(filas, desde) {
  const out = [];
  filas.slice(1).forEach(r => {
    const sec = String(r[0] || '').trim().toUpperCase();
    const imp = parseNum_(r[3]);
    if (!/^S[1-4]$/.test(sec) || imp === null) return;
    const fecha = parseFecha_(r[1]) || new Date(desde.getFullYear(), desde.getMonth(), desde.getDate() - 1);
    const texto = String(r[2] || '');
    const v = Math.abs(imp);
    if (sec === 'S1' || sec === 'S2') {
      out.push(partidaBanco_({ fecha, concepto: texto, importe: sec === 'S1' ? v : -v, ref: '', nombre: '' }, 'Arrastre'));
    } else {
      out.push(partidaFbs_({ fecha, comentario: texto, referencia: '', debe: sec === 'S3' ? v : 0, haber: sec === 'S4' ? v : 0 }, 'Arrastre'));
    }
  });
  return out;
}

// ------------------------------------------------------------------ cruce

function unir_(banco, fbs, metodo) {
  banco.forEach(b => { b.match = fbs; b.metodo = metodo; });
  fbs.forEach(f => { f.match = banco; f.metodo = metodo; });
}

function primeraCombinacion_(cands, n, objetivo) {
  const idx = [];
  const buscar = (inicio, suma) => {
    if (idx.length === n) return Math.abs(suma - objetivo) <= TOLERANCIA;
    for (let i = inicio; i <= cands.length - (n - idx.length); i++) {
      idx.push(i);
      if (buscar(i + 1, suma + cands[i].importe)) return true;
      idx.pop();
    }
    return false;
  };
  return buscar(0, 0) ? idx.map(i => cands[i]) : null;
}

function pasada_(banco, fbs, criterio, metodo, ventana, unico) {
  const libres = new Map();
  fbs.forEach(f => {
    if (f.match !== null) return;
    const k = Math.round(f.importe);
    if (!libres.has(k)) libres.set(k, []);
    libres.get(k).push(f);
  });
  banco.slice().sort((a, b) => a.fecha - b.fecha).forEach(b => {
    if (b.match !== null) return;
    const k = Math.round(b.importe), cands = [];
    [k - 1, k, k + 1].forEach(kk => (libres.get(kk) || []).forEach(f => {
      if (f.match === null && Math.abs(f.importe - b.importe) <= TOLERANCIA && dias_(b.fecha, f.fecha) <= ventana && criterio(b, f)) cands.push(f);
    }));
    if (!cands.length || (unico && cands.length > 1)) return;
    let mejor = cands[0];
    cands.forEach(f => {
      const d1 = dias_(b.fecha, f.fecha), d0 = dias_(b.fecha, mejor.fecha);
      if (d1 < d0 || (d1 === d0 && Math.abs(f.importe - b.importe) < Math.abs(mejor.importe - b.importe))) mejor = f;
    });
    unir_([b], [mejor], metodo);
  });
}

function agrupados_(banco, fbs) {
  [[banco, fbs, true], [fbs, banco, false]].forEach(([uno, varios, esBanco]) => {
    uno.forEach(u => {
      if (u.match !== null || !(u.dni || u.nombre.size)) return;
      const cands = varios.filter(v => v.match === null && dias_(u.fecha, v.fecha) <= VENTANA_DIAS &&
        (v.importe > 0) === (u.importe > 0) && ((u.dni && v.dni === u.dni) || comparten_(u.nombre, v.nombre))).slice(0, 12);
      for (let n = 2; n <= Math.min(cands.length, 6); n++) {
        const combo = primeraCombinacion_(cands, n, u.importe);
        if (combo) {
          if (esBanco) unir_([u], combo, 'Agrupado CUIT/DNI/Nombre'); else unir_(combo, [u], 'Agrupado CUIT/DNI/Nombre');
          break;
        }
      }
    });
  });
}

function combinaciones_(banco, fbs) {
  [[banco, fbs, true], [fbs, banco, false]].forEach(([uno, varios, esBanco]) => {
    uno.slice().sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe)).forEach(u => {
      if (u.match !== null || u.origen !== 'Mes') return;
      const cands = varios.filter(v => v.match === null && v.origen === 'Mes' && dias_(u.fecha, v.fecha) <= 7)
        .sort((a, b) => dias_(u.fecha, a.fecha) - dias_(u.fecha, b.fecha)).slice(0, 40);
      for (let n = 2; n <= 3; n++) {
        const combo = primeraCombinacion_(cands, n, u.importe);
        if (combo) {
          if (esBanco) unir_([u], combo, 'Combinación (revisar)'); else unir_(combo, [u], 'Combinación (revisar)');
          break;
        }
      }
    });
  });
}

function claveLote_(f) {
  const m = f.texto.match(/Nº\s*(.+)$/);
  return m ? m[1].trim() : f.texto.split(/[\s/-]/)[0];
}

function compensacionesFbs_(fbs, mismaClave) {
  const grupos = new Map();
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes') return;
    const k = f.fecha.getTime() + '|' + (mismaClave ? claveLote_(f) : '');
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(f);
  });
  grupos.forEach(dia => {
    dia.slice().sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe)).forEach(u => {
      if (u.match !== null) return;
      const cands = dia.filter(v => v.match === null && (v.importe > 0) !== (u.importe > 0)).slice(0, 25);
      for (let n = 1; n <= 4; n++) {
        const combo = primeraCombinacion_(cands, n, -u.importe);
        if (combo) {
          [u].concat(combo).forEach(x => { x.match = []; x.metodo = 'Compensa dentro de FBS'; });
          break;
        }
      }
    });
  });
}

function reversionesFbs_(fbs) {
  const grupos = new Map();
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes') return;
    const k = f.texto + '|' + Math.abs(f.importe);
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(f);
  });
  grupos.forEach(g => {
    const pos = g.filter(f => f.importe > 0), neg = g.filter(f => f.importe < 0);
    for (let i = 0; i < Math.min(pos.length, neg.length); i++) {
      pos[i].match = []; neg[i].match = [];
      pos[i].metodo = neg[i].metodo = 'Compensa dentro de FBS';
    }
  });
}

function lotes_(banco, fbs) {
  const gb = new Map(), gf = new Map();
  const push = (m, k, x) => { if (!m.has(k)) m.set(k, { fecha: x.fecha, signo: x.importe > 0, items: [] }); m.get(k).items.push(x); };
  banco.forEach(b => { if (b.match === null && b.origen === 'Mes') push(gb, b.fecha.getTime() + '|' + b.causal + '|' + (b.importe > 0), b); });
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes') return;
    push(gf, f.fecha.getTime() + '|' + claveLote_(f) + '|' + (f.importe > 0), f);
    push(gf, f.fecha.getTime() + '|*|' + (f.importe > 0), f);
  });
  Array.from(gb.values()).sort((a, b) => a.fecha - b.fecha).forEach(g => {
    if (g.items.some(b => b.match !== null)) return;
    const total = g.items.reduce((s, b) => s + b.importe, 0);
    for (const h of gf.values()) {
      if (h.signo === g.signo && dias_(h.fecha, g.fecha) <= 3 && h.items.every(f => f.match === null) &&
        Math.abs(h.items.reduce((s, f) => s + f.importe, 0) - total) <= TOLERANCIA && (g.items.length > 1 || h.items.length > 1)) {
        unir_(g.items, h.items, 'Lote (suma del día)');
        break;
      }
    }
  });
}

function gastosAsiento_(banco, fbs) {
  const b = banco.filter(x => x.match === null && x.origen === 'Mes' && (x.cat === 'GASTOS BANCARIOS' || CAUSALES_IMPUESTOS.includes(x.causal)));
  const f = fbs.filter(x => x.match === null && /GASTOS BANCARIOS \d{2}\/\d{4}/.test(x.texto.toUpperCase()));
  if (!b.length || !f.length) return null;
  const dif = redondear_(b.reduce((s, x) => s + x.importe, 0) - f.reduce((s, x) => s + x.importe, 0));
  unir_(b, f, 'Asiento gastos bancarios (dif. ' + formato_(dif) + ')');
  return dif;
}

function conciliar_(banco, fbs) {
  const mismoId = (b, f) => (b.cuit && b.cuit === f.cuit) || (b.dni && b.dni === f.dni);
  const porRef = (b, f) => b.ref.length >= 5 && (f.texto + ' ' + f.ref).replace(/\./g, '').indexOf(b.ref) >= 0;
  const porNombre = (b, f) => comparten_(b.nombre, f.nombre);
  pasada_(banco, fbs, mismoId, '1. CUIT/DNI', VENTANA_ID, false);
  pasada_(banco, fbs, porRef, '2. Referencia', 60, false);
  pasada_(banco, fbs, porNombre, '3. Nombre', 15, false);
  agrupados_(banco, fbs);
  const difGastos = gastosAsiento_(banco, fbs);
  compensacionesFbs_(fbs, true);
  pasada_(banco, fbs, () => true, '4. Fecha + importe', 0, true);
  lotes_(banco, fbs);
  pasada_(banco, fbs, () => true, '5. Sugerido (solo importe)', VENTANA_SUGERIDO, false);
  pasada_(banco, fbs, b => /^(TRANSFERENCIA ENTRE|INVERSIONES)/.test(b.cat), '5. Sugerido (importe en el mes)', 31, true);
  combinaciones_(banco, fbs);
  compensacionesFbs_(fbs, false);
  reversionesFbs_(fbs);
  return difGastos;
}

function sector_(p) {
  if (p.lado === 'BANCO') return p.importe > 0 ? 'S1' : 'S2';
  return p.importe > 0 ? 'S3' : 'S4';
}

function gastosBancarios_(movs) {
  const g = { com: 0, iva: 0, perc: 0, sircreb: 0, ret: 0, imp: 0 };
  movs.forEach(m => {
    const t = m.concepto.toUpperCase(), d = -m.importe;
    if (m.causal === '1684' || m.causal === '1685') g.imp += d;
    else if (m.causal === '1297') g.sircreb += d;
    else if (m.causal === '4145') g.ret += d;
    else if (m.cat === 'GASTOS BANCARIOS') {
      if (t.includes('PERCEP') || t.includes('IVA_PER')) g.perc += d; else if (t.includes('IVA')) g.iva += d; else g.com += d;
    }
  });
  return [['Comisiones', g.com], ['IVA 21%', g.iva], ['Percepciones IVA', g.perc], ['Percepciones IIBB (SIRCREB)', g.sircreb],
    ['Retenciones IIBB rentas financieras', g.ret], ['Imp. créditos y débitos (total)', g.imp],
    ['    Crédito computable 33%', g.imp * 0.33], ['    Imp. créditos y débitos (gasto 67%)', g.imp * 0.67]];
}

/** Todo el proceso sobre matrices de valores (sin tocar hojas): se puede probar fuera de Sheets. */
function conciliarTodo_(entrada) {
  const reglas = leerReglas_(entrada.reglas), empresas = leerEmpresas_(entrada.empresas);
  const ext = leerExtracto_(entrada.extracto, reglas, empresas);
  const e = leerMayor_(entrada.mayorE, HOJA.mayorE), o = leerMayor_(entrada.mayorO, HOJA.mayorO);
  const movs = ext.movs;
  const desde = new Date(movs[0].fecha.getFullYear(), movs[0].fecha.getMonth(), 1);
  const corte = movs[movs.length - 1].fecha;
  const avisos = e.avisos.concat(o.avisos);

  // controles de los archivos
  let quiebres = 0;
  for (let i = 1; i < movs.length; i++) {
    if (movs[i].saldo !== null && movs[i - 1].saldo !== null && Math.abs(movs[i - 1].saldo + movs[i].importe - movs[i].saldo) > 0.005) quiebres++;
  }
  if (quiebres) avisos.push('Extracto: ' + quiebres + ' filas donde el saldo no sigue a la anterior (¿faltan movimientos?).');
  const nuevos = movs.filter(m => m.desc === 'Código causal nuevo');
  if (nuevos.length) avisos.push('Extracto: ' + nuevos.length + ' movimientos con código causal sin regla (' +
    Array.from(new Set(nuevos.map(m => m.causal))).join(', ') + '). Agregalos en "' + HOJA.reglas + '".');

  const banco = movs.map(m => partidaBanco_(m, 'Mes'));
  const fbs = movimientosFbs_(e.asientos, o.asientos).map(a => partidaFbs_(a, 'Mes'));
  const anteriores = leerAnteriores_(entrada.anteriores || [], desde);
  anteriores.forEach(p => (p.lado === 'BANCO' ? banco : fbs).push(p));
  // control de apertura: saldo FBS inicial + pendientes anteriores debe dar el saldo inicial del banco
  const neto = anteriores.reduce((x, p) => x + (p.lado === 'BANCO' ? p.importe : -p.importe), 0);
  const saldoInicialBanco = redondear_(movs[0].saldo - movs[0].importe);
  const apertura = { fbs: redondear_(e.info.saldoInicial + o.info.saldoInicial), pendientes: redondear_(neto),
    banco: saldoInicialBanco };
  apertura.diferencia = redondear_(apertura.fbs + apertura.pendientes - apertura.banco);
  if (Math.abs(apertura.diferencia) >= 1000) {
    avisos.unshift('La apertura no cierra por ' + formato_(apertura.diferencia) + ': faltan o sobran partidas en "' +
      HOJA.anteriores + '" (pendientes del cierre anterior). Esa misma diferencia se arrastra al control del mes.');
  }
  const difGastos = conciliar_(banco, fbs);

  const pend = { S1: [], S2: [], S3: [], S4: [] };
  banco.concat(fbs).forEach(p => { if (p.match === null) pend[sector_(p)].push(p); });
  const tot = {};
  Object.keys(pend).forEach(s => { tot[s] = redondear_(pend[s].reduce((x, p) => x + Math.abs(p.importe), 0)); });
  const saldoE = e.info.saldoFinal, saldoO = o.info.saldoFinal;
  const saldoBanco = movs[movs.length - 1].saldo;
  const ajustado = redondear_(saldoE + saldoO + tot.S1 - tot.S2 - tot.S3 + tot.S4);

  const metodos = {};
  banco.forEach(p => {
    if (p.origen !== 'Mes') return;
    const k = p.match !== null ? p.metodo.split(' (dif')[0] : 'Sin conciliar';
    if (!metodos[k]) metodos[k] = [0, 0];
    metodos[k][0]++; metodos[k][1] += Math.abs(p.importe);
  });

  const conciliados = [], vistos = new Set();
  banco.forEach(b => {
    if (b.match === null || vistos.has(b.match)) return;
    const grupoF = b.match;
    vistos.add(grupoF);
    const grupoB = grupoF.length ? grupoF[0].match : [b];
    for (let i = 0; i < Math.max(grupoB.length, grupoF.length); i++) {
      const x = grupoB[i], y = grupoF[i];
      conciliados.push([b.metodo, x ? x.fecha : '', x ? x.texto : '', x ? x.importe : '', x ? x.origen : '',
        y ? y.fecha : '', y ? y.texto : '', y ? y.importe : '', y ? y.asiento : '', y ? y.origen : '']);
    }
  });
  fbs.forEach(f => { if (Array.isArray(f.match) && !f.match.length) conciliados.push([f.metodo, '', '', '', '', f.fecha, f.texto, f.importe, f.asiento, f.origen]); });

  return { info: ext.info, infoE: e.info, infoO: o.info, corte, saldoE, saldoO, saldoBanco, pend, tot, ajustado, apertura,
    metodos, difGastos, gastos: gastosBancarios_(movs), conciliados, avisos, banco, fbs };
}

if (typeof module !== 'undefined') module.exports = { conciliarTodo_, leerExtracto_, leerMayor_, leerReglas_, leerEmpresas_, parseNum_, parseFecha_, REGLAS_INICIALES, EMPRESAS_INICIALES };

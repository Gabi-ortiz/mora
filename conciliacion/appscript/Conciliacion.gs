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
  pendO: 'O sin confirmar',
  pendFbs: 'E sin cruzar',
  pendBanco: 'Banco sin registrar',
  conciliados: 'E conciliado',
  analisisO: 'Análisis O',
  historial: 'Historial',
  parametros: 'Parámetros',
  viejas: ['Pend. registrar en FBS', 'Pend. FBS sin banco', 'Conciliados'],   // nombres de versiones anteriores
  foto: 'FOTO',
};

const SECTORES = [
  ['S1', 'Depósitos no registrados en contabilidad', '+'],
  ['S2', 'Débitos no registrados en contabilidad', '-'],
  ['S3', 'Depósitos no acreditados en banco', '-'],
  ['S4', 'Cheques / pagos no debitados en banco', '+'],
];

const VENTANA_DIAS = 10;
const VENTANA_ID = 31;
const TOLERANCIA = 0.02;
const NUM_FMT = '#,##0.00;[Red]-#,##0.00';
const CAUSALES_IMPUESTOS = ['1684', '1685', '1297', '4145', '1479', '1972'];   // van al asiento mensual de gastos
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
  ["1215", "Ambos", "", "GASTOS BANCARIOS", "Percepción IVA caja de seguridad"],
  ["5", "Ambos", "", "GASTOS BANCARIOS", "Intereses adelanto en cuenta corriente + IVA"],
  ["1802", "Ambos", "", "GASTOS BANCARIOS", "Comisión cheque consulta + IVA"],
  ["1479", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Impuesto de sellos Córdoba (DGR)"],
  ["1972", "Ambos", "", "IMPUESTOS Y RETENCIONES", "Retención IIBB Tucumán"],
  ["4196", "Ambos", "PLAN", "PLAN DE AHORRO / TERMINAL", "Débito plan de ahorro (Chevrolet, Fiat, etc.)"],
  ["*", "Crédito", "LIQ COMER (PRISMA|CABAL|PAYWAY)", "COBRANZA TARJETAS", "Liquidación de tarjetas (Prisma, Cabal, Payway)"],
  ["*", "Crédito", "CCERR TARJETA NAR", "COBRANZA TARJETAS", "Liquidación Tarjeta Naranja (circuito cerrado)"],
  ["*", "Crédito", "TARJETA NARANJA", "COBRANZA TARJETAS", "Liquidación Tarjeta Naranja (Datanet)"]
];

// ------------------------------------------------------------------ menú

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Conciliación')
    .addItem('1. Crear hojas de entrada', 'crearHojasEntrada')
    .addItem('2. Procesar conciliación', 'procesarConciliacion')
    .addItem('3. Cerrar mes (guardar foto y pasar pendientes)', 'cerrarMes')
    .addSeparator()
    .addItem('Agregar reglas nuevas del script', 'agregarReglasNuevas')
    .addToUi();
}

function crearParametros_(ss) {
  if (ss.getSheetByName(HOJA.parametros)) return;
  const sh = ss.insertSheet(HOJA.parametros);
  sh.getRange(1, 1, 4, 2).setValues([
    ['Parámetro', 'Valor'],
    ['Modo de cruce', 'ESTRICTO'],
    ['ESTRICTO', 'Sin CUIT / referencia / nombre / fecha escrita en el comprobante, solo cruza si coincide la fecha y el importe no es redondo.'],
    ['INTERMEDIO', 'Además cruza importes redondos cuando hay un único candidato de cada lado (mismo día o hasta 3 días) y transferencias propias / FCI contra pases de E. Todo queda marcado "(revisar)".'],
  ]);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.getRange(2, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ESTRICTO', 'INTERMEDIO']).build());
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 700);
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
  crearParametros_(ss);
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
    parametros: ss.getSheetByName(HOJA.parametros) ? leer(HOJA.parametros) : null,
  });
  HOJA.viejas.forEach(n => { const v = ss.getSheetByName(n); if (v) ss.deleteSheet(v); });
  escribirTablero_(ss, res);
  escribirAnalisisO_(ss, res);
  escribirPendientes_(ss, HOJA.pendO, res, res.listas.O,
    'Paso 1 — Cuenta O: líneas que no netean (no confirmadas). Van a Depósitos no acreditados / Cheques no debitados.');
  escribirConciliados_(ss, res);
  escribirPendientes_(ss, HOJA.pendFbs, res, res.listas.E,
    'Paso 2 — Cuenta E que no se cruzó ni con pendientes anteriores ni con el extracto (posibles errores de registración).');
  escribirPendientes_(ss, HOJA.pendBanco, res, res.listas.B,
    'Movimientos del extracto (y pendientes anteriores del banco) que no están en la cuenta E.');
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
  [HOJA.pendO, HOJA.pendFbs, HOJA.pendBanco].forEach(nombre => {
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
    let detalle;
    if (s === 'S1' || s === 'S2') detalle = 'ver "' + HOJA.pendBanco + '"';
    else {
      const o = res.pend[s].filter(p => res.listas.O.indexOf(p) >= 0);
      detalle = 'O sin confirmar: ' + o.length + ' por ' + formato_(o.reduce((x, p) => x + Math.abs(p.importe), 0)) +
        ' / E sin cruzar: ' + (res.pend[s].length - o.length);
    }
    add(['(' + signo + ') ' + nombre + '  — ' + detalle, res.tot[s], res.pend[s].length]);
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

function observacion_(p) {
  if (p.alerta) return p.alerta;
  if (p.origen === 'Arrastre') return 'Pendiente de meses anteriores' + (p.cuenta === 'O' ? ' (sigue sin confirmar en O)' : '');
  if (p.lado === 'BANCO') return p.cat === 'A IDENTIFICAR' ? 'Concepto del banco sin identificar' : 'Registrar en FBS';
  if (/Diferencia asiento gastos/.test(p.texto)) return 'Revisar asiento de gastos bancarios';
  if (p.cuenta === 'O') return 'No confirmado (cuenta O)';
  if (p.cuenta === 'E') return 'Pendiente en cuenta E: posible error de registración, revisar';
  return '';
}

function escribirPendientes_(ss, nombre, res, items, titulo) {
  const sh = hojaLimpia_(ss, nombre);
  const cab = ['Sector', 'Cuadro', 'Fecha', 'Días', 'Origen', 'Concepto / Comprobante', 'Importe', 'Categoría',
    'CUIT/DNI', 'Ref. banco / Asiento FBS', 'Cuenta FBS', 'Observación', 'Cruces encontrados'];
  const nombres = {};
  SECTORES.forEach(([s, n]) => { nombres[s] = n; });
  const filas = items.slice().sort((a, b) => sector_(a).localeCompare(sector_(b)) || a.fecha - b.fecha).map(p => [
    sector_(p), nombres[sector_(p)], p.fecha, dias_(res.corte, p.fecha), p.origen === 'Arrastre' ? 'Mes anterior' : 'Mes',
    p.texto, Math.abs(p.importe), p.cat, p.cuit || p.dni, p.asiento || p.ref, p.cuenta || '', observacion_(p), p.cruces || '']);
  const porSector = {};
  items.forEach(p => { const k = sector_(p); porSector[k] = porSector[k] || [0, 0]; porSector[k][0]++; porSector[k][1] += Math.abs(p.importe); });
  const total = Object.keys(porSector).sort().map(k => nombres[k] + ': ' + porSector[k][0] + ' partidas, ' + formato_(porSector[k][1])).join('   |   ');
  sh.getRange(1, 1).setValue(titulo + '   —   ' + (total || 'sin partidas')).setFontWeight('bold');
  sh.getRange(2, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  if (filas.length) {
    sh.getRange(3, 9, filas.length, 2).setNumberFormat('@');   // CUIT y referencias como texto
    sh.getRange(3, 13, filas.length, 1).setWrap(true);
    sh.getRange(3, 1, filas.length, cab.length).setValues(filas.map(f => f.map((v, i) => (i >= 8 ? String(v || '') : v))));
    sh.getRange(3, 3, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(3, 7, filas.length, 1).setNumberFormat(NUM_FMT);
  }
  sh.setFrozenRows(2);
  sh.getRange(2, 1, Math.max(filas.length, 1) + 1, cab.length).createFilter();
  [60, 230, 90, 50, 90, 360, 130, 200, 110, 150, 80, 280, 420].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function escribirConciliados_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.conciliados);
  const cab = ['Cruce', 'Fecha banco', 'Concepto banco / pendiente anterior', 'Importe banco', 'Origen banco',
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
  // primero las reglas con patrón (el patrón más largo gana), después las de todo el código; "*" = cualquier código
  const ordenadas = reglas.filter(r => r.patron).sort((a, b) => b.patron.source.length - a.patron.source.length)
    .concat(reglas.filter(r => !r.patron));
  for (const conPatron of [true, false]) {
    for (const r of ordenadas) {
      if (!!r.patron === conPatron && (r.cod === causal || r.cod === '*') && (r.sentido === null || r.sentido === sentido) &&
        (!r.patron || r.patron.test(concepto))) return [r.cat, r.desc];
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
    cat: m.cat || '', causal: m.causal || '', match: null, metodo: '', cruces: '' };
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
  const fr = a.fecha ? fechaEnTexto_(com, a.fecha) : null;
  return { lado: 'FBS', origen: origen || 'Mes', fecha: a.fecha, importe: redondear_(a.debe - a.haber), texto: com,
    fechaRef: fr || null, cruces: '',
    ref: a.referencia || '', asiento: a.asiento || '', cuenta: a.cuenta || '', alerta: '', cuit, dni,
    nombre: partes.length > 1 ? tokens_(partes[partes.length - 1]) : new Set(), cat: '', causal: '', match: null, metodo: '' };
}

/** Une E y O descartando los pases entre ambas cuentas (mismo asiento e importe con signo opuesto). */
function movimientosFbs_(asientosE, asientosO) {
  const neto = new Map(), primero = new Map(), porCuenta = new Map();
  [[asientosE, 'E'], [asientosO, 'O']].forEach(([lista, cta]) => lista.forEach(a => {
    const k = a.asiento + '|' + a.comentario + '|' + redondear_(Math.abs(a.debe - a.haber));
    neto.set(k, (neto.get(k) || 0) + a.debe - a.haber);
    if (!primero.has(k)) primero.set(k, a);
    if (!porCuenta.has(k)) porCuenta.set(k, { E: 0, O: 0 });
    porCuenta.get(k)[cta] += a.debe - a.haber;
  }));
  const out = [];
  neto.forEach((v, k) => {
    if (Math.abs(v) > 0.005) {
      const a = Object.assign({}, primero.get(k));
      const c = porCuenta.get(k);
      a.debe = Math.max(v, 0); a.haber = Math.max(-v, 0);
      a.cuenta = Math.abs(c.E) > 0.005 && Math.abs(c.O) > 0.005 ? 'E+O' : Math.abs(c.E) > 0.005 ? 'E' : 'O';
      out.push(a);
    }
  });
  return out;
}

// ------------------------------------------------------------------ análisis de la cuenta O y armado de la cuenta E

/** Clave para netear dentro de la cuenta O: número de comprobante (RC / RM) o número de liquidación de tarjeta. */
function claveO_(texto) {
  const t = String(texto || '').replace(/\s+/g, ' ').trim().toUpperCase();
  let m = t.match(/^(RC-[A-Z]-\d{4}-\d{8}|RM-\d+)/);
  if (m) return m[1];
  m = t.match(/^(\d{4,})\s*-\s*LIQ/);
  if (m && !/GASTOS/.test(t)) return 'LIQ ' + m[1];
  return t;
}

/** Busca dentro de "lista" combinaciones que se anulan (1 contra 1..4 del signo contrario). Devuelve las que quedan. */
function netear_(lista) {
  const libres = lista.slice();
  const usar = x => { x.match = []; x.metodo = 'Confirmado en O'; };
  libres.slice().sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe)).forEach(u => {
    if (u.match !== null) return;
    const cands = libres.filter(v => v.match === null && v !== u && (v.importe > 0) !== (u.importe > 0)).slice(0, 25);
    for (let n = 1; n <= 4; n++) {
      const combo = primeraCombinacion_(cands, n, -u.importe);
      if (combo) { [u].concat(combo).forEach(usar); break; }
    }
  });
  let resto = libres.filter(x => x.match === null);
  if (resto.length > 1 && Math.abs(resto.reduce((x, p) => x + p.importe, 0)) <= TOLERANCIA) { resto.forEach(usar); resto = []; }
  return resto;
}

/**
 * Paso 1 del procedimiento: se analiza la cuenta O sola. Cada comprobante (o liquidación) se registra y después se
 * confirma en la misma cuenta; lo que netea a cero está confirmado. Lo que queda es "no confirmado" y va directo a
 * Depósitos no acreditados / Cheques no debitados, sin cruzarlo con el banco. Los pendientes anteriores de FBS con
 * comprobante (RC / RM / liquidación) participan del neteo: si este mes se confirmaron, desaparecen.
 */
function analisisO_(asientosO, anterioresFbs) {
  const lineas = asientosO.map(a => { const p = partidaFbs_(a, 'Mes'); p.cuenta = 'O'; return p; });
  const grupos = new Map();
  const agregar = p => { const k = claveO_(p.texto); if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(p); };
  lineas.forEach(agregar);
  const usadosAnteriores = [];
  anterioresFbs.forEach(p => { if (/^(RC-|RM-|LIQ )/.test(claveO_(p.texto)) && grupos.has(claveO_(p.texto))) { p.cuenta = 'O'; agregar(p); usadosAnteriores.push(p); } });
  const pendientes = [], resumen = [];
  grupos.forEach((g, clave) => {
    const resto = netear_(g);
    resto.forEach(p => pendientes.push(p));
    const debe = g.filter(p => p.importe > 0).reduce((x, p) => x + p.importe, 0);
    const haber = -g.filter(p => p.importe < 0).reduce((x, p) => x + p.importe, 0);
    resumen.push({ clave, lineas: g.length, debe, haber, neto: redondear_(debe - haber), pendientes: resto.length,
      detallePend: resto.map(p => fechaTexto_(p.fecha) + ' ' + (p.importe > 0 ? 'debe ' : 'haber ') + formato_(Math.abs(p.importe))).join(' | '),
      desde: g.reduce((d, p) => (p.fecha < d ? p.fecha : d), g[0].fecha), texto: g[0].texto,
      anteriores: g.filter(p => p.origen === 'Arrastre').length });
  });
  return { pendientes, resumen, usadosAnteriores };
}

/**
 * Paso 2: la cuenta E se cruza con el banco. Las líneas de una misma liquidación de tarjeta (la "Confirmación de
 * Valores Agrupados" del asiento que confirma la liquidación en O, y el asiento con el número de liquidación) se
 * juntan en un solo neto, que es lo que acredita el banco.
 */
function partidasE_(asientosE, asientosO) {
  const liqPorAsiento = new Map(), numeros = new Set();
  asientosO.forEach(a => {
    const k = claveO_(a.comentario);
    if (k.startsWith('LIQ ')) { liqPorAsiento.set(String(a.asiento), k); numeros.add(k.slice(4)); }
  });
  const liqDe = a => {
    if (liqPorAsiento.has(String(a.asiento))) return liqPorAsiento.get(String(a.asiento));
    const k = claveO_(a.comentario);
    if (k.startsWith('LIQ ')) return k;
    const d = String(a.comentario || '').trim().match(/^(\d{4,})\b/);
    if (d) {
      const num = Array.from(numeros).filter(n => d[1].startsWith(n)).sort((x, y) => y.length - x.length)[0];
      if (num) return 'LIQ ' + num;
    }
    return null;
  };
  // operaciones con el mismo número al inicio del comentario (ej. venta de cheques "25082026-Vta Cheques" y
  // "25082026" con los gastos): el banco acredita el neto
  const opDe = a => {
    const t = String(a.comentario || '').trim();
    const m = t.match(/^(\d{6,9})(\s*-|$)/);
    return m && !/GASTOS|LIQ/i.test(t) ? 'OP ' + m[1] : null;
  };
  const cuentaOp = new Map();
  asientosE.forEach(a => { const k = opDe(a); if (k) cuentaOp.set(k, (cuentaOp.get(k) || 0) + 1); });
  const sueltas = [], liqs = new Map();
  asientosE.forEach(a => {
    const op = opDe(a);
    const k = liqDe(a) || (op && cuentaOp.get(op) > 1 ? op : null);
    if (!k) { const p = partidaFbs_(a, 'Mes'); p.cuenta = 'E'; sueltas.push(p); return; }
    if (!liqs.has(k)) liqs.set(k, []);
    liqs.get(k).push(a);
  });
  liqs.forEach((g, k) => {
    const neto = g.reduce((x, a) => x + a.debe - a.haber, 0);
    if (Math.abs(neto) <= TOLERANCIA) return;   // la liquidación se anula dentro de E
    const ult = g.reduce((m, a) => (a.fecha > m.fecha ? a : m), g[0]);
    const nombre = k.startsWith('OP ') ? 'Operación ' + k.slice(3) : 'Liquidación tarjeta ' + k.slice(4);
    const p = partidaFbs_({ fecha: ult.fecha, comentario: nombre + ' (neto cuenta E, ' + g.length + ' líneas)',
      referencia: '', asiento: Array.from(new Set(g.map(a => a.asiento))).join(' / '),
      debe: Math.max(neto, 0), haber: Math.max(-neto, 0) }, 'Mes');
    p.cuenta = 'E';
    if (k.startsWith('OP ')) p.fechaRef = fechaEnTexto_(k.slice(3), ult.fecha);   // el número suele ser la fecha
    sueltas.push(p);
  });
  return sueltas;
}

function escribirAnalisisO_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.analisisO);
  const cab = ['Estado', 'Comprobante / liquidación', 'Desde', 'Líneas', 'De meses anteriores', 'Debe', 'Haber', 'Neto',
    'Líneas sin confirmar', 'Cuáles quedan sin confirmar', 'Detalle'];
  const filas = res.analisisO.slice().sort((a, b) => (b.pendientes > 0) - (a.pendientes > 0) || a.desde - b.desde)
    .map(g => [g.pendientes ? 'Pendiente' : 'Confirmado', g.clave, g.desde, g.lineas, g.anteriores, g.debe, g.haber, g.neto,
      g.pendientes, g.detallePend, g.texto]);
  const pend = res.analisisO.filter(g => g.pendientes).length;
  sh.getRange(1, 1).setValue('Cuenta O: ' + res.analisisO.length + ' comprobantes/liquidaciones, ' + (res.analisisO.length - pend) +
    ' confirmados (netean a cero) y ' + pend + ' con líneas pendientes').setFontWeight('bold');
  sh.getRange(2, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  if (filas.length) {
    sh.getRange(3, 1, filas.length, cab.length).setValues(filas);
    sh.getRange(3, 3, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(3, 6, filas.length, 3).setNumberFormat(NUM_FMT);
  }
  sh.setFrozenRows(2);
  sh.getRange(2, 1, filas.length + 1, cab.length).createFilter();
  [90, 200, 90, 60, 80, 130, 130, 130, 80, 260, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w));
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
    out[out.length - 1].sectorOriginal = sec;
  });
  return out;
}

// ------------------------------------------------------------------ cruce

let CRUCES_ = [];   // registro de todos los cruces de la corrida (para la hoja "E conciliado")

function unir_(banco, fbs, metodo) {
  banco.forEach(b => { b.match = fbs; b.metodo = metodo; });
  fbs.forEach(f => { f.match = banco; f.metodo = metodo; });
  CRUCES_.push({ metodo, banco, fbs });
}

/** Registros de FBS que se anulan entre sí (no pasan por el banco). */
function compensar_(items, metodo) {
  items.forEach(x => { x.match = []; x.metodo = metodo; });
  CRUCES_.push({ metodo, banco: [], fbs: items });
}

/** Días entre dos partidas; si el comentario del recibo trae una fecha (ej. "... 16/04/2026"), también se usa. */
function distancia_(a, b) {
  let d = dias_(a.fecha, b.fecha);
  if (a.fechaRef) d = Math.min(d, dias_(a.fechaRef, b.fecha));
  if (b.fechaRef) d = Math.min(d, dias_(a.fecha, b.fechaRef));
  return d;
}

/** Última fecha dd/mm o dd/mm/aaaa escrita en un comentario (sin tomar "30052/9" ni "7/2026"). */
function fechaEnTexto_(texto, base) {
  const re = /(?:^|[^\d\/])(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?(?![\d\/])/g;
  let m, ult = null;
  while ((m = re.exec(texto)) !== null) {
    const d = Number(m[1]), mes = Number(m[2]);
    if (d < 1 || d > 31 || mes < 1 || mes > 12) continue;
    let a = m[3] ? Number(m[3]) : base.getFullYear();
    if (a < 100) a += 2000;
    if (!m[3] && mes > base.getMonth() + 1) a -= 1;
    ult = new Date(a, mes - 1, d);
  }
  if (!ult) {
    // "Nº 4080226", "Nº 21082026", "Nº 210820261", "25082026-Vta Cheques": día + mes + año pegados
    const c = String(texto).match(/(?:Nº\s*|^)(\d{6,9})(?!\d)/);
    if (c) {
      const t = c[1];
      for (const ld of [2, 1]) {
        const d = Number(t.slice(0, ld)), mes = Number(t.slice(ld, ld + 2)), resto = t.slice(ld + 2);
        let a = resto.startsWith('20') && resto.length >= 4 ? Number(resto.slice(0, 4)) : Number(resto.slice(-2));
        if (a < 100) a += 2000;
        const f = new Date(a, mes - 1, d);
        if (d >= 1 && d <= 31 && mes >= 1 && mes <= 12 && a === base.getFullYear() && dias_(f, base) <= 62) { ult = f; break; }
      }
    }
  }
  return ult;
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
      if (f.match === null && Math.abs(f.importe - b.importe) <= TOLERANCIA && distancia_(b, f) <= ventana && criterio(b, f)) cands.push(f);
    }));
    if (!cands.length || (unico && cands.length > 1)) return;
    let mejor = cands[0];
    cands.forEach(f => {
      const d1 = distancia_(b, f), d0 = distancia_(b, mejor);
      if (d1 < d0 || (d1 === d0 && Math.abs(f.importe - b.importe) < Math.abs(mejor.importe - b.importe))) mejor = f;
    });
    unir_([b], [mejor], metodo);
  });
}

function agrupados_(banco, fbs) {
  [[banco, fbs, true], [fbs, banco, false]].forEach(([uno, varios, esBanco]) => {
    uno.forEach(u => {
      if (u.match !== null || !(u.dni || u.nombre.size)) return;
      const cands = varios.filter(v => v.match === null && distancia_(u, v) <= VENTANA_DIAS &&
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
          compensar_([u].concat(combo), 'Compensa dentro de FBS');
          break;
        }
      }
    });
  });
}

function reversionesFbs_(fbs) {
  // anulaciones: "Anulacion As ..." contra el registro anulado (del mes o de un pendiente anterior), mismo importe opuesto
  fbs.filter(a => a.match === null && /ANULA/i.test(a.texto)).forEach(a => {
    const o = fbs.find(x => x !== a && x.match === null && Math.abs(x.importe + a.importe) <= TOLERANCIA && dias_(x.fecha, a.fecha) <= 60);
    if (o) compensar_([a, o], 'Compensa dentro de FBS (anulación)');
  });
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
      compensar_([pos[i], neg[i]], 'Compensa dentro de FBS (reversión)');
    }
  });
}

function lotes_(banco, fbs) {
  const gb = new Map(), gf = new Map();
  const push = (m, k, x) => { if (!m.has(k)) m.set(k, { fecha: x.fecha, signo: x.importe > 0, clave: k.split('|')[1], items: [] }); m.get(k).items.push(x); };
  banco.forEach(b => { if (b.match === null && b.origen === 'Mes') push(gb, b.fecha.getTime() + '|' + b.causal + '|' + (b.importe > 0), b); });
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes') return;
    push(gf, f.fecha.getTime() + '|' + claveLote_(f) + '|' + (f.importe > 0), f);
    push(gf, f.fecha.getTime() + '|*|' + (f.importe > 0), f);
    if (f.asiento) push(gf, f.fecha.getTime() + '|as:' + f.asiento + '|' + (f.importe > 0), f);   // varias líneas de un mismo asiento
  });
  Array.from(gb.values()).sort((a, b) => a.fecha - b.fecha).forEach(g => {
    if (g.items.some(b => b.match !== null)) return;
    const total = g.items.reduce((s, b) => s + b.importe, 0);
    for (const h of gf.values()) {
      if (h.signo === g.signo && dias_(h.fecha, g.fecha) <= 3 && h.items.every(f => f.match === null) &&
        Math.abs(h.items.reduce((s, f) => s + f.importe, 0) - total) <= TOLERANCIA && (g.items.length > 1 || h.items.length > 1) &&
        (h.clave !== '*' || especifico_(total))) {
        unir_(g.items, h.items, 'Lote (suma del día)');
        break;
      }
    }
  });
}

/** Un movimiento del banco = neto de un asiento FBS con varias líneas (ej. Confirmación de Valores Agrupados
 *  menos comisiones, o venta de cheques: cheques individuales menos comisiones). */
function asientosNetos_(banco, fbs) {
  const grupos = new Map();
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes' || !f.asiento) return;
    if (!grupos.has(f.asiento)) grupos.set(f.asiento, []);
    grupos.get(f.asiento).push(f);
  });
  grupos.forEach(g => {
    if (g.length < 2 || g.some(f => f.match !== null)) return;
    const neto = g.reduce((x, f) => x + f.importe, 0);
    let mejor = null;
    banco.forEach(b => {
      if (b.match !== null || b.origen !== 'Mes' || Math.abs(b.importe - neto) > TOLERANCIA || dias_(b.fecha, g[0].fecha) > 7 ||
        !especifico_(b.importe)) return;
      if (!mejor || dias_(b.fecha, g[0].fecha) < dias_(mejor.fecha, g[0].fecha)) mejor = b;
    });
    if (mejor) unir_([mejor], g, 'Asiento FBS neto (tarjetas / venta de cheques)');
  });
}

/** Marca pendientes de FBS que parecen duplicados de otro registro (mismo importe y mismo detalle). */
function marcarDuplicados_(fbs) {
  // solo comprobantes con formato "COMPROBANTE / detalle" (recibos RC, órdenes de pago RM, etc.)
  const detalle = f => f.texto.split(' / ').slice(1).join(' / ').replace(/\s+/g, ' ').trim().toUpperCase();
  fbs.forEach(f => {
    if (f.match !== null || f.origen !== 'Mes' || !detalle(f)) return;
    const otro = fbs.find(x => x !== f && x.origen === 'Mes' && x.asiento !== f.asiento && Math.abs(x.importe - f.importe) <= TOLERANCIA &&
      detalle(x) === detalle(f) && dias_(x.fecha, f.fecha) <= 31);
    if (otro) f.alerta = 'Posible duplicado de ' + (otro.texto.split('/')[0].trim() || 'asiento ' + otro.asiento) +
      ' (asiento ' + otro.asiento + ')';
  });
}

/** Liquidaciones de tarjeta (neto de la cuenta E) contra acreditaciones de tarjetas del banco: se admite hasta
 *  $1 de diferencia por redondeos; la diferencia queda como partida a ajustar para no esconderla en el control. */
function liquidacionesTarjeta_(banco, fbs) {
  const ajustes = [];
  fbs.filter(f => f.match === null && /^Liquidación tarjeta/.test(f.texto)).forEach(f => {
    let mejor = null;
    banco.forEach(b => {
      if (b.match !== null || b.origen !== 'Mes' || b.cat !== 'COBRANZA TARJETAS' || (b.importe > 0) !== (f.importe > 0) ||
        Math.abs(b.importe - f.importe) > 1 || dias_(b.fecha, f.fecha) > 10) return;
      if (!mejor || dias_(b.fecha, f.fecha) < dias_(mejor.fecha, f.fecha)) mejor = b;
    });
    if (!mejor) return;
    const dif = redondear_(mejor.importe - f.importe);
    unir_([mejor], [f], 'Liquidación de tarjeta' + (Math.abs(dif) > 0.005 ? ' (dif. ' + formato_(dif) + ')' : ''));
    if (Math.abs(dif) > 0.005) {
      const a = partidaFbs_({ fecha: f.fecha, comentario: 'Diferencia de redondeo ' + f.texto.split(' (')[0].toLowerCase() + ' - ajustar',
        referencia: '', asiento: f.asiento, debe: Math.max(-dif, 0), haber: Math.max(dif, 0) }, 'Mes');
      a.cuenta = 'E';
      ajustes.push(a);
    }
  });
  return ajustes;
}

function gastosAsiento_(banco, fbs) {
  const b = banco.filter(x => x.match === null && x.origen === 'Mes' && (x.cat === 'GASTOS BANCARIOS' || CAUSALES_IMPUESTOS.includes(x.causal)));
  // el asiento mensual "Gastos bancarios MM/AAAA" (no las "Liq ... gastos bancarios", que son otros registros)
  const f = fbs.filter(x => x.match === null && x.origen === 'Mes' && /GASTOS BANCARIOS \d{2}\/\d{4}/.test(x.texto.toUpperCase()) &&
    !/LIQ/.test(x.texto.toUpperCase()));
  if (!b.length || !f.length) return null;
  const dif = redondear_(b.reduce((s, x) => s + x.importe, 0) - f.reduce((s, x) => s + x.importe, 0));
  unir_(b, f, 'Asiento gastos bancarios (dif. ' + formato_(dif) + ')');
  let ajuste = null;
  if (Math.abs(dif) > TOLERANCIA) {
    // la diferencia queda como partida pendiente para que el control no la esconda
    ajuste = partidaFbs_({ fecha: f[0].fecha, comentario: 'Diferencia asiento gastos bancarios (FBS vs. banco) - revisar',
      referencia: '', asiento: f[0].asiento, debe: Math.max(-dif, 0), haber: Math.max(dif, 0) }, 'Mes');
  }
  return { dif, ajuste };
}

function conciliar_(banco, fbs) {
  pasada_(banco, fbs, MISMO_ID_, '1. CUIT/DNI', VENTANA_ID, false);
  pasada_(banco, fbs, POR_REF_, '2. Referencia', 60, false);
  pasada_(banco, fbs, POR_NOMBRE_, '3. Nombre', 15, false);
  pasada_(banco, fbs, (b, f) => !!f.fechaRef && dias_(f.fechaRef, b.fecha) === 0, '3. Fecha escrita en el comprobante + importe', 60, true);
  agrupados_(banco, fbs);
  const gastos = gastosAsiento_(banco, fbs);
  compensacionesFbs_(fbs, true);
  // regla: sin CUIT / referencia / nombre, solo se cruza si coincide la fecha y el importe es específico (no redondo)
  pasada_(banco, fbs, b => especifico_(b.importe), '4. Fecha + importe específico', 0, true);
  if (CRUZAR_REDONDO_UNICO_DIA) unicoDelDia_(banco, fbs);
  if (CRUZAR_REDONDO_UNICO_DIA >= 2) {
    // transferencias propias / FCI contra pases, fondos o registros "BANCO ..." de E, hasta 3 días, candidato único
    pasada_(banco, fbs, (b, f) => /^(TRANSFERENCIA ENTRE|INVERSIONES)/.test(b.cat) && /PASE|FONDO|BANCO|RESCATE|SUSCRIP|FCI/i.test(f.texto),
      '4c. Transferencia propia / FCI contra pase de E (revisar)', 3, true);
  }
  if (CRUZAR_REDONDO_UNICO_DIA >= 3) ventanaUnica_(banco, fbs, 3);
  lotes_(banco, fbs);
  asientosNetos_(banco, fbs);
  const ajustesTarjeta = liquidacionesTarjeta_(banco, fbs);
  reversionesFbs_(fbs);
  if (gastos && gastos.ajuste) fbs.push(gastos.ajuste);   // se agrega al final para que ninguna pasada la cruce
  ajustesTarjeta.forEach(a => fbs.push(a));
  marcarDuplicados_(fbs);
  return gastos ? gastos.dif : null;
}

/** Importe "específico": tiene centavos o no termina en 00 (123.852,68 sí; 500.000 o 71.100 no). */
function especifico_(x) {
  const c = Math.round(Math.abs(x) * 100);
  return c % 100 !== 0 || (c / 100) % 100 !== 0;
}

/** Importe redondo: se cruza solo si ese día hay UN movimiento del banco y UN registro de E con ese importe. */
let CRUZAR_REDONDO_UNICO_DIA = 0;   // 0 = ESTRICTO, 3 = INTERMEDIO (se lee de la hoja "Parámetros")
function unicoDelDia_(banco, fbs) {
  const clave = p => p.fecha.getTime() + '|' + p.importe.toFixed(2);
  const cb = new Map(), cf = new Map();
  banco.forEach(b => { if (b.origen === 'Mes') cb.set(clave(b), (cb.get(clave(b)) || []).concat([b])); });
  fbs.forEach(f => { if (f.origen === 'Mes') cf.set(clave(f), (cf.get(clave(f)) || []).concat([f])); });
  cb.forEach((bs, k) => {
    const fs = cf.get(k) || [];
    if (bs.length === 1 && fs.length === 1 && bs[0].match === null && fs[0].match === null) unir_(bs, fs, '4b. Fecha + importe redondo, único en el día (revisar)');
  });
}

/** Importe (aunque sea redondo) único en ambos lados dentro de +-dias. */
function ventanaUnica_(banco, fbs, d) {
  banco.filter(b => b.match === null && b.origen === 'Mes').forEach(b => {
    const fs = fbs.filter(f => f.match === null && Math.abs(f.importe - b.importe) <= TOLERANCIA && distancia_(b, f) <= d);
    if (fs.length !== 1) return;
    const bs = banco.filter(x => x.match === null && Math.abs(x.importe - b.importe) <= TOLERANCIA && distancia_(x, fs[0]) <= d);
    if (bs.length === 1) unir_([b], fs, '4d. Importe único en +-' + d + ' días (revisar)');
  });
}

const MISMO_ID_ = (b, f) => (b.cuit && b.cuit === f.cuit) || (b.dni && b.dni === f.dni);
const POR_REF_ = (b, f) => b.ref.length >= 5 && (f.texto + ' ' + f.ref).replace(/\./g, '').indexOf(b.ref) >= 0;
const POR_NOMBRE_ = (b, f) => comparten_(b.nombre, f.nombre);

/**
 * Paso 2a: la cuenta E contra los pendientes de la conciliación anterior.
 *  - Movimientos del banco que el mes pasado no estaban registrados y este mes se registraron en E.
 *  - Pendientes anteriores de FBS que este mes se revierten en E (mismo importe, signo contrario).
 */
function cruzarAnteriores_(eItems, antBanco, antFbs) {
  const v = 400;
  pasada_(antBanco, eItems, MISMO_ID_, 'Pendiente anterior: CUIT/DNI', v, false);
  pasada_(antBanco, eItems, POR_REF_, 'Pendiente anterior: referencia', v, false);
  pasada_(antBanco, eItems, POR_NOMBRE_, 'Pendiente anterior: nombre', v, false);
  pasada_(antBanco, eItems, (b, f) => !!f.fechaRef && dias_(f.fechaRef, b.fecha) === 0 && especifico_(b.importe),
    'Pendiente anterior: fecha del comentario + importe', v, true);
  antFbs.forEach(a => {
    if (a.match !== null) return;
    const f = eItems.find(x => x.match === null && Math.abs(x.importe + a.importe) <= TOLERANCIA &&
      (claveO_(x.texto) === claveO_(a.texto) || comparten_(x.nombre, a.nombre) || /ANULA|REVERS/i.test(x.texto)));
    if (f) compensar_([a, f], 'Pendiente anterior: revertido en E');
  });
}

/**
 * Marca con qué se cruza cada línea de O sin confirmar (pendiente anterior, cuenta E, extracto) y cada
 * movimiento del banco sin registrar que parece un recibo de O sin confirmar. Solo marca: no resuelve nada.
 */
function marcarCruces_(oPend, anteriores, eItems, bancoMes) {
  const desc = p => fechaTexto_(p.fecha) + ' ' + p.texto.slice(0, 45) + ' ' + formato_(p.importe);
  // apellido del comprobante FBS ("RC-... / ... / APELLIDO, NOMBRE") presente del otro lado, o 2 palabras en común
  const apellido = p => { const t = p.texto.split(' / '); return t.length > 1 ? (t[t.length - 1].toUpperCase().match(/[A-ZÑ]{3,}/) || [''])[0] : ''; };
  const mismoNombre = (a, b) => {
    const comunes = Array.from(a.nombre).filter(x => b.nombre.has(x));
    return comunes.length >= 2 || [a, b].some(x => x.lado === 'FBS' && apellido(x) && comunes.indexOf(apellido(x)) >= 0);
  };
  const parecido = (a, b) => (a.cuit && a.cuit === b.cuit) || (a.dni && a.dni === b.dni) || mismoNombre(a, b) ||
    (a.lado === 'FBS' && b.lado === 'FBS' && claveO_(a.texto) === claveO_(b.texto));
  const mismoSigno = (a, b) => Math.abs(a.importe - b.importe) <= TOLERANCIA;
  const cualquierSigno = (a, b) => Math.abs(Math.abs(a.importe) - Math.abs(b.importe)) <= TOLERANCIA;
  oPend.forEach(p => {
    const marcas = [];
    const ant = anteriores.find(x => x !== p && (x.lado === 'BANCO' ? mismoSigno(x, p) : cualquierSigno(x, p)) && parecido(x, p));
    if (ant) marcas.push('Pendiente anterior ' + ant.sectorOriginal + ': ' + desc(ant) + (ant.match !== null ? ' [ya resuelto]' : ''));
    const e = eItems.find(x => mismoSigno(x, p) && parecido(x, p));
    if (e) marcas.push('Cuenta E asiento ' + e.asiento + ': ' + desc(e) + (e.match !== null ? ' [conciliado]' : ' [sin cruzar]'));
    let b = bancoMes.find(x => mismoSigno(x, p) && distancia_(x, p) <= 15 && parecido(x, p)), soloImporte = false;
    if (!b && especifico_(p.importe)) { b = bancoMes.find(x => x.match === null && mismoSigno(x, p) && distancia_(x, p) === 0); soloImporte = !!b; }
    if (b) marcas.push('Banco' + (soloImporte ? ' (misma fecha e importe específico)' : '') + ': ' + desc(b) +
      (b.match !== null ? ' [conciliado con E]' : ' [sin registrar]'));
    p.cruces = marcas.join(' | ');
  });
  bancoMes.forEach(b => {
    if (b.match !== null) return;
    const o = oPend.find(x => mismoSigno(x, b) && distancia_(b, x) <= 15 && parecido(b, x));
    if (o) b.cruces = 'Posible recibo de O sin confirmar: ' + desc(o);
  });
  // sugerencias que NO se cruzan (regla: importe redondo o sin dato que lo identifique): misma fecha e importe
  bancoMes.forEach(b => {
    if (b.match !== null) return;
    const cands = eItems.filter(f => f.match === null && mismoSigno(f, b) && distancia_(b, f) === 0);
    if (cands.length !== 1) return;
    const f = cands[0], txt = 'Sugerencia (no cruzado, importe sin dato que lo identifique): ';
    b.cruces = (b.cruces ? b.cruces + ' | ' : '') + txt + 'E asiento ' + f.asiento + ' ' + desc(f);
    f.cruces = (f.cruces ? f.cruces + ' | ' : '') + txt + 'banco ' + desc(b);
  });
}

function sector_(p) {
  if (p.lado === 'BANCO') return p.importe > 0 ? 'S1' : 'S2';
  return p.importe > 0 ? 'S3' : 'S4';
}

function gastosBancarios_(movs) {
  const g = { com: 0, iva: 0, iva105: 0, perc: 0, int: 0, sircreb: 0, ret: 0, tuc: 0, sellos: 0, imp: 0 };
  movs.forEach(m => {
    const t = m.concepto.toUpperCase(), d = -m.importe;
    if (m.causal === '1684' || m.causal === '1685') g.imp += d;
    else if (m.causal === '1297') g.sircreb += d;
    else if (m.causal === '4145') g.ret += d;
    else if (m.causal === '1972') g.tuc += d;
    else if (m.causal === '1479') g.sellos += d;
    else if (m.cat === 'GASTOS BANCARIOS') {
      if (t.includes('PERCEP') || t.includes('IVA_PER')) g.perc += d;
      else if (t.includes('IVA')) { if (m.causal === '5') g.iva105 += d; else g.iva += d; }
      else if (m.causal === '5') g.int += d;
      else g.com += d;
    }
  });
  const filas = [['Comisiones', g.com], ['IVA 21%', g.iva], ['Intereses', g.int], ['IVA 10,5%', g.iva105],
    ['Percepciones IVA', g.perc], ['Percepciones IIBB (SIRCREB)', g.sircreb], ['Retenciones IIBB rentas financieras', g.ret],
    ['Percepciones IIBB Tucumán', g.tuc], ['DGR Sellos Córdoba', g.sellos], ['Imp. créditos y débitos (total)', g.imp],
    ['    Crédito computable 33%', g.imp * 0.33], ['    Imp. créditos y débitos (gasto 67%)', g.imp * 0.67]];
  const total = g.com + g.iva + g.int + g.iva105 + g.perc + g.sircreb + g.ret + g.tuc + g.sellos + g.imp;
  return filas.filter(f => Math.abs(f[1]) > 0.004).concat([['TOTAL (debe coincidir con el asiento de FBS)', total]]);
}

/** Agrega a la hoja "Reglas" las reglas del script que todavía no están (no toca las existentes). */
function agregarReglasNuevas() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(HOJA.reglas);
  if (!sh) throw new Error('Falta la hoja "' + HOJA.reglas + '". Corré "1. Crear hojas de entrada".');
  const clave = r => [String(r[0]).trim(), normalizar_(r[1]).charAt(0), String(r[2] || '').trim()].join('|');
  const existentes = new Set(sh.getDataRange().getValues().slice(1).map(clave));
  const nuevas = REGLAS_INICIALES.filter(r => !existentes.has(clave(r)));
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 5).setValues(nuevas);
  SpreadsheetApp.getUi().alert(nuevas.length + ' reglas nuevas agregadas a "' + HOJA.reglas + '".');
}

/** Todo el proceso sobre matrices de valores (sin tocar hojas): se puede probar fuera de Sheets. */
function conciliarTodo_(entrada, redondoUnico) {
  if (redondoUnico !== undefined) CRUZAR_REDONDO_UNICO_DIA = redondoUnico;
  else if (entrada.parametros) {
    const fila = entrada.parametros.find(r => /modo de cruce/i.test(String(r[0])));
    CRUZAR_REDONDO_UNICO_DIA = fila && /INTERMEDIO/i.test(String(fila[1])) ? 3 : 0;
  }
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

  CRUCES_ = [];
  const bancoMes = movs.map(m => partidaBanco_(m, 'Mes'));
  const anteriores = leerAnteriores_(entrada.anteriores || [], desde);
  const anterioresBanco = anteriores.filter(p => p.lado === 'BANCO');
  const anterioresFbs = anteriores.filter(p => p.lado === 'FBS');
  // PASO 1: cuenta O sola, neteando por comprobante / liquidación junto con los pendientes anteriores
  const ao = analisisO_(o.asientos, anterioresFbs);
  // PASO 2: cuenta E -> primero contra los pendientes anteriores, después contra el extracto
  const eItems = partidasE_(e.asientos, o.asientos);
  const antFbsRestantes = anterioresFbs.filter(p => ao.usadosAnteriores.indexOf(p) < 0);
  cruzarAnteriores_(eItems, anterioresBanco, antFbsRestantes);
  const banco = bancoMes.concat(anterioresBanco);
  const fbs = eItems.concat(antFbsRestantes);
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
  marcarCruces_(ao.pendientes, anteriores, eItems, bancoMes);

  const pend = { S1: [], S2: [], S3: [], S4: [] };
  banco.concat(fbs).forEach(p => { if (p.match === null) pend[sector_(p)].push(p); });
  ao.pendientes.forEach(p => pend[sector_(p)].push(p));   // no confirmados en O: directo a pendientes
  const tot = {};
  Object.keys(pend).forEach(s => { tot[s] = redondear_(pend[s].reduce((x, p) => x + Math.abs(p.importe), 0)); });
  const fbsPend = pend.S3.concat(pend.S4).filter(p => p.origen === 'Mes');
  const enE = fbsPend.filter(p => p.cuenta === 'E' && !/Diferencia asiento gastos/.test(p.texto));
  const dup = fbsPend.filter(p => p.alerta);
  if (dup.length) avisos.push(dup.length + ' registros de FBS parecen duplicados (ver columna Observación en "' + HOJA.pendFbs + '").');
  const conMarca = ao.pendientes.filter(p => p.cruces).length;
  if (conMarca) avisos.push(conMarca + ' líneas de O sin confirmar tienen un posible cruce (columna "Cruces encontrados" en "' + HOJA.pendO + '").');
  if (enE.length) avisos.push(enE.length + ' registros de la cuenta E no se encontraron en el banco, por ' + formato_(enE.reduce((x, p) => x + Math.abs(p.importe), 0)) +
    ': según el procedimiento son posibles errores de registración, revisar.');
  const noConf = pend.S3.concat(pend.S4).filter(p => p.cuenta === 'O');
  if (noConf.length) avisos.push(noConf.length + ' líneas de la cuenta O sin confirmar, por ' +
    formato_(noConf.reduce((x, p) => x + Math.abs(p.importe), 0)) + ' (detalle en "' + HOJA.pendO + '").');
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

  const conciliados = [];
  CRUCES_.forEach(c => {
    for (let i = 0; i < Math.max(c.banco.length, c.fbs.length, 1); i++) {
      const x = c.banco[i], y = c.fbs[i];
      conciliados.push([c.metodo, x ? x.fecha : '', x ? x.texto : '', x ? x.importe : '', x ? x.origen : '',
        y ? y.fecha : '', y ? y.texto : '', y ? y.importe : '', y ? y.asiento : '', y ? (y.origen === 'Arrastre' ? 'Pendiente anterior' : 'Cuenta ' + (y.cuenta || 'E')) : '']);
    }
  });
  const listas = { O: ao.pendientes, E: fbs.filter(p => p.match === null), B: banco.filter(p => p.match === null) };

  return { info: ext.info, infoE: e.info, infoO: o.info, corte, saldoE, saldoO, saldoBanco, pend, tot, ajustado, apertura, analisisO: ao.resumen, listas,
    metodos, difGastos, gastos: gastosBancarios_(movs), conciliados, avisos, banco, fbs };
}

if (typeof module !== 'undefined') module.exports = { conciliarTodo_, leerExtracto_, leerMayor_, leerReglas_, leerEmpresas_, parseNum_, parseFecha_, REGLAS_INICIALES, EMPRESAS_INICIALES };

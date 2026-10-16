/**
 * CONCILIACIÓN BANCARIA — Banco Macro vs. FBS (cuentas E + O)
 *
 * Hojas de ENTRADA (se pegan los archivos tal cual, desde A1):
 *   - "E. bancario"            : extracto de Macro (Excel "Últimos Movimientos" o CSV "Resumen").
 *   - "Base"                   : reporte de FBS con las cuentas E y O (DETPLAN, DETComenta, DETComentaAux, ...).
 *   - "Pend mes anterior"      : partidas pendientes del cierre anterior (Sector, Fecha, Concepto, Importe).
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
  extractoAlt: ['E. bancario', 'Extracto bancario'],
  base: 'Base',
  mayorE: 'Mayor E',
  mayorO: 'Mayor O',
  anteriores: 'Pendientes anteriores',
  anterioresAlt: ['Pendientes en trabajo', 'Pend mes anterior'],
  reglas: 'Reglas',
  empresas: 'Empresas grupo',
  tablero: 'Tablero',
  cuatro: 'Conciliación',
  pendO: 'O sin confirmar',
  pendFbs: 'E sin cruzar',
  pendBanco: 'Banco sin registrar',
  conciliados: 'E conciliado',
  analisisO: 'Análisis O',
  revisar: 'Para revisar',
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
const CAUSALES_IMPUESTOS = ['1684', '1685', '1297', '4145', '1479', '1972',   // Macro: van al asiento mensual de gastos
  '907176', '907376', '917376', '907173', '905130',                             // Galicia: imp. déb./créd., IIBB, sellos
  '4637', '4633', '1384', '566'];                                               // Santander: imp. déb./créd., SIRCREB, Tucumán
const PALABRAS_VACIAS = new Set(['SA', 'SRL', 'SAS', 'DE', 'LA', 'EL', 'BANCO', 'MACRO', 'TRANSF', 'VAR', 'VARIOS',
  'CUO', 'CUOTAS', 'FAC', 'FACTURAS', 'TEF', 'DATANET', 'ING', 'CIRC', 'CERRADO', 'CCERR']);

const EMPRESAS_INICIALES = [['30669400167', 'TURIN SA'], ['30681026386', 'CHEXA SA']];

// Código causal | Sentido (Débito/Crédito/Ambos) | Patrón en concepto (regex, vacío = todos) | Categoría | Descripción
/** Galicia: "Código causal" = número de la columna "Concepto" del Excel de Office Banking (ej. "907173 - ING. BRUTOS S/ CRED"). */
const REGLAS_GALICIA = [
  ["917151", "Crédito", "", "COBRANZA CLIENTES", "Transferencia de terceros"],
  ["917123", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Servicio pago a proveedores (cobro de empresas)"],
  ["917187", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Pago a proveedores factura (cobro de empresas)"],
  ["917193", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Transferencias cash proveedores (cobro de empresas)"],
  ["917115", "Crédito", "", "COBRANZA TARJETAS", "Acreditamiento Prisma comercios"],
  ["917102", "Crédito", "", "DEPOSITOS EN EFECTIVO", "Depósito en efectivo"],
  ["917302", "Crédito", "", "DEPOSITOS EN EFECTIVO", "Depósito efectivo en autoservicio"],
  ["917195", "Crédito", "", "CHEQUES", "Gestión de cheque / echeq depositado"],
  ["917169", "Crédito", "", "CHEQUES", "Gestión de cheque"],
  ["917118", "Crédito", "", "CHEQUES", "Acreditamiento canje Galicia"],
  ["907281", "Débito", "", "CHEQUES", "Rechazo de cheque depositado"],
  ["907282", "Débito", "", "CHEQUES", "Compra / cesión de cheques"],
  ["907364", "Débito", "", "CHEQUES", "Echeq propio debitado (Galicia)"],
  ["907365", "Débito", "", "CHEQUES", "Echeq propio debitado (48 hs)"],
  ["917133", "Crédito", "", "FINANCIACION / PRESTAMOS", "Crédito por descuento de documentos"],
  ["907751", "Débito", "", "GASTOS BANCARIOS", "Interés nominal (descuento de documentos)"],
  ["917138", "Crédito", "", "INVERSIONES (FCI)", "Rescate FIMA"],
  ["907237", "Débito", "", "INVERSIONES (FCI)", "Suscripción FIMA"],
  ["917152", "Crédito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia de cuenta propia"],
  ["917154", "Crédito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transfer cash misma titularidad"],
  ["917312", "Crédito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia cuentas propias"],
  ["907268", "Débito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata a cuenta propia"],
  ["907297", "Débito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia cuentas propias"],
  ["907316", "Débito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia cuentas propias"],
  ["907178", "Débito", "", "PAGO PROVEEDORES", "Transferencia a terceros"],
  ["907179", "Débito", "", "PAGO PROVEEDORES", "Transferencia a terceros"],
  ["907232", "Débito", "", "PAGO PROVEEDORES", "Transferencia inmediata a proveedores"],
  ["907269", "Débito", "", "PAGO PROVEEDORES", "Transferencia inmediata a proveedores"],
  ["907104", "Débito", "", "PAGO PROVEEDORES", "Pago de servicios"],
  ["907154", "Débito", "", "PAGO PROVEEDORES", "Pago tarjeta Visa empresa"],
  ["907213", "Débito", "", "DEBITOS AUTOMATICOS", "Débito automático de servicios"],
  ["907139", "Débito", "", "SUELDOS", "Acreditamiento de haberes"],
  ["907255", "Débito", "", "IMPUESTOS Y RETENCIONES", "Transferencia AFIP (VEP)"],
  ["907176", "Débito", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 sobre débitos"],
  ["907376", "Débito", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 sobre créditos"],
  ["917376", "Crédito", "", "IMPUESTOS Y RETENCIONES", "Devolución impuesto Ley 25.413"],
  ["907173", "Débito", "", "IMPUESTOS Y RETENCIONES", "Ingresos brutos s/ créditos (SIRCREB / Tucumán)"],
  ["905130", "Débito", "", "IMPUESTOS Y RETENCIONES", "Impuesto de sellos"],
  ["907171", "Débito", "", "GASTOS BANCARIOS", "IVA 21% sobre comisiones"],
  ["907307", "Débito", "", "GASTOS BANCARIOS", "IVA 10,5% sobre intereses"],
  ["907118", "Débito", "", "GASTOS BANCARIOS", "Intereses sobre saldos deudores"],
  ["907138", "Débito", "", "GASTOS BANCARIOS", "Comisión transferencia entre bancos"],
  ["907188", "Débito", "", "GASTOS BANCARIOS", "Comisión depósito de cheque"],
  ["907189", "Débito", "", "GASTOS BANCARIOS", "Comisión rechazo de cheques"],
  ["907291", "Débito", "", "GASTOS BANCARIOS", "Comisión depósito de cheque en otra sucursal"],
  ["907394", "Débito", "", "GASTOS BANCARIOS", "Comisión servicio de cuenta"],
];

/** Reglas iniciales según el banco de la planilla (fila "Cuenta bancaria" de Parámetros). */
function reglasDelBanco_(nombreBanco) {
  const n = String(nombreBanco || '');
  if (/galicia/i.test(n)) return REGLAS_GALICIA;
  if (/santander|r[ií]o\b/i.test(n)) return REGLAS_SANTANDER;
  if (/patagonia/i.test(n)) return REGLAS_PATAGONIA;
  if (/c[oó]rdoba|bancor/i.test(n)) return REGLAS_CORDOBA;
  return REGLAS_INICIALES;
}

/**
 * La hoja Reglas tiene que tener las reglas del banco de "Cuenta bancaria". Si no tiene ninguna (ej. quedó con las de
 * Macro al crear la planilla): si todas las que hay son las de fábrica de otro banco, se reemplazan; si alguien agregó
 * reglas propias, se suman las del banco sin borrar nada.
 */
function asegurarReglasDelBanco_(ss, banco) {
  const sh = ss.getSheetByName(HOJA.reglas);
  if (!sh) return;
  const clave = r => String(r[0]).trim() + '|' + String(r[2] || '').trim();
  const delBanco = reglasDelBanco_(banco);
  const hay = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => String(r[0]).trim()) : [];
  const claves = new Set(hay.map(clave));
  if (delBanco.some(r => claves.has(clave(r)))) return;
  const deFabrica = new Set([].concat(REGLAS_INICIALES, REGLAS_GALICIA, REGLAS_SANTANDER, REGLAS_PATAGONIA, REGLAS_CORDOBA).map(clave));
  if (hay.every(r => deFabrica.has(clave(r))) && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  const desde = hay.every(r => deFabrica.has(clave(r))) ? 2 : sh.getLastRow() + 1;
  sh.getRange(desde, 1, delBanco.length, 5).setValues(delBanco);
}

/** Banco de Córdoba: "Concepto" es el tipo de operación (sin código) y "Descripcion" trae CUIT y nombre. */
const REGLAS_CORDOBA = [
  ["*", "Crédito", "^CR[EÉ]DITO TRANSFERENCIA CAJERO", "COBRANZA CLIENTES", "Crédito transferencia cajero / homebanking"],
  ["*", "Crédito", "^TRANSF\\.? ?RECIBIDA", "COBRANZA CLIENTES", "Transferencia recibida"],
  ["*", "Crédito", "^TRANSF\\.? ?CREDITO HB", "COBRANZA CLIENTES", "Transferencia crédito homebanking"],
  ["*", "Crédito", "^CREDIN", "COBRANZA CLIENTES", "Credin (transferencia recibida)"],
  ["*", "Crédito", "^CREDITO PARA TRANSFERENCIA INTERNA MODO", "COBRANZA CLIENTES", "Transferencia MODO"],
  ["*", "Crédito", "^TRANSF\\. MODO", "COBRANZA CLIENTES", "Transferencia MODO"],
  ["*", "Crédito", "^DEP[OÓ]SITO", "DEPOSITOS EN EFECTIVO", "Depósito"],
  ["*", "Crédito", "^ACREDITACION A COMER", "COBRANZA EMPRESAS / ASEGURADORAS", "Acreditación a comercio (UNA)"],
  ["*", "Crédito", "^LIQUIDACI[OÓ]N DE PR[EÉ]STAMOS", "COBRANZA EMPRESAS / ASEGURADORAS", "Liquidación de préstamos comercios aliados"],
  ["*", "Crédito", "^CR[EÉ]DITO PROVEEDORES", "COBRANZA EMPRESAS / ASEGURADORAS", "Crédito proveedores (Ministerio)"],
  ["*", "Crédito", "^TRANSF\\.CRED\\. IB PLAN", "COBRANZA EMPRESAS / ASEGURADORAS", "Pago a proveedores (Interbanking)"],
  ["*", "Débito", "^TRANSFERENCIA HOMEBANKING", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia homebanking (pase)"],
  ["*", "Débito", "^PAGO AUTOMATICO DE SERVICIOS", "DEBITOS AUTOMATICOS", "Pago automático de servicios"],
  ["*", "Débito", "^DEBITO PAGOS COMARB", "IMPUESTOS Y RETENCIONES", "Pago IIBB Convenio Multilateral"],
  ["*", "Débito", "^IMPUESTO SOBRE OPERACIONES", "GASTOS BANCARIOS", "Impuesto Ley 25.413 (débitos y créditos)"],
  ["*", "Débito", "^RECAUDACIONES SIRCREB", "GASTOS BANCARIOS", "Recaudación IIBB SIRCREB"],
  ["*", "Débito", "^GASTOS DE MANTENIMIENTO", "GASTOS BANCARIOS", "Gastos de mantenimiento"],
  ["*", "Débito", "^IVA", "GASTOS BANCARIOS", "IVA (comisiones y percepción)"],
  ["*", "Débito", "^COMISI", "GASTOS BANCARIOS", "Comisiones"],
];

/** Patagonia: el Excel de "Movimientos de Cuenta" no trae código; las reglas van por la descripción (código "*"). */
const REGLAS_PATAGONIA = [
  ["*", "Crédito", "^CREDITO POR TRANSFERENCIA", "COBRANZA CLIENTES", "Crédito por transferencia"],
  ["*", "Crédito", "^TRANSFERENCIA ENTRE CUENTAS", "COBRANZA CLIENTES", "Transferencia entre cuentas (de terceros)"],
  ["*", "Crédito", "^TRANSFERENCIA D/C", "COBRANZA CLIENTES", "Transferencia D/C recibida"],
  ["*", "Débito", "^TRANSFERENCIA D/C", "PAGO PROVEEDORES", "Transferencia D/C enviada"],
  ["*", "Crédito", "^DEPOSITO CHEQUE", "CHEQUES", "Depósito de cheque"],
  ["*", "Crédito", "^DEPOSITO EN EFECTIVO", "DEPOSITOS EN EFECTIVO", "Depósito en efectivo por caja"],
  ["*", "Crédito", "^CREDITO INTERPYME", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Crédito Interpyme"],
  ["*", "Crédito", "^PROMOCION PRODUCTOS", "OTROS CREDITOS DEL BANCO", "Promoción productos Patagonia"],
  ["*", "Débito", "^TRANSF\\. TERCEROS", "PAGO PROVEEDORES", "Transferencia a terceros (e-bank)"],
  ["*", "Débito", "^TRANSF\\. MEP", "PAGO PROVEEDORES", "Transferencia MEP a otros bancos"],
  ["*", "Débito", "^TRANSF\\. PROP", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia a cuenta propia (e-bank)"],
  ["*", "Débito", "^DEBITO P/ACREDITAC\\.DE SUELDOS", "SUELDOS", "Acreditación de sueldos"],
  ["*", "Débito", "^DEBITO AUTOMATICO", "DEBITOS AUTOMATICOS", "Débito automático"],
  ["*", "Débito", "^COMISION POR CONVENIOS", "PAGO PROVEEDORES", "Comisión por convenios (se registra aparte)"],
  ["*", "Débito", "^IMP\\.DB/CR BANCARIOS", "GASTOS BANCARIOS", "Impuesto Ley 25.413 (débitos y créditos)"],
  ["*", "Débito", "^RET\\.IIBB SIRCREB", "GASTOS BANCARIOS", "Retención IIBB SIRCREB"],
  ["*", "Débito", "^RET\\.IIBB TUCUMAN", "GASTOS BANCARIOS", "Retención IIBB Tucumán"],
  ["*", "Débito", "^IMPUESTO A LOS SELLOS", "GASTOS BANCARIOS", "Impuesto de sellos"],
  ["*", "Débito", "^INTERES", "GASTOS BANCARIOS", "Intereses (acuerdo)"],
  ["*", "Débito", "^IVA", "GASTOS BANCARIOS", "IVA (comisiones e intereses)"],
  ["*", "Débito", "^COM", "GASTOS BANCARIOS", "Comisiones"],
  ["*", "Crédito", "^BONIFICACION COMISION", "GASTOS BANCARIOS", "Bonificación de comisiones"],
  ["*", "Crédito", "^IVA", "GASTOS BANCARIOS", "Devolución de IVA"],
];

/** Santander Río: "Cod. Operativo" del Excel de Últimos Movimientos. */
const REGLAS_SANTANDER = [
  ["4805", "Crédito", "", "COBRANZA CLIENTES", "Transferencia recibida"],
  ["4806", "Crédito", "", "COBRANZA CLIENTES", "Transferencia recibida"],
  ["2767", "Crédito", "", "COBRANZA CLIENTES", "Transferencia recibida (DEBIN / credin)"],
  ["3413", "Crédito", "", "COBRANZA CLIENTES", "Transferencia recibida desde CVU (billetera)"],
  ["1430", "Crédito", "", "COBRANZA CLIENTES", "Transferencia ctas mobile banking"],
  ["1253", "Crédito", "", "COBRANZA CLIENTES", "Crédito transferencia online banking empresas"],
  ["1257", "Crédito", "", "COBRANZA CLIENTES", "Crédito transferencia online banking"],
  ["2377", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Pago a proveedores recibido (cobro de empresas)"],
  ["3002", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Transferencia pagos a terceros (cobro de empresas)"],
  ["1970", "Crédito", "", "COBRANZA EMPRESAS / ASEGURADORAS", "Servicios de pago (aseguradoras)"],
  ["2602", "Crédito", "", "COBRANZA TARJETAS", "Acreditación comercio Payway (Prisma)"],
  ["2604", "Crédito", "", "COBRANZA TARJETAS", "Acreditación comercio Fiserv (First Data)"],
  ["2571", "Débito", "", "COBRANZA TARJETAS", "Débito comercio Fiserv (First Data)"],
  ["2574", "Débito", "", "COBRANZA TARJETAS", "Débito comercio Payway (Prisma)"],
  ["3372", "Débito", "", "COBRANZA TARJETAS", "Débito comercio Getnet"],
  ["1397", "Crédito", "", "FINANCIACION / PRESTAMOS", "Liquidación de préstamo"],
  ["2327", "Débito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Pago a cuentas propias (Interbanking)"],
  ["2393", "Crédito", "", "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Cobro desde cuentas propias (Interbanking)"],
  ["4713", "Débito", "", "PAGO PROVEEDORES", "Pago Interbanking B2B"],
  ["1852", "Débito", "", "PAGO PROVEEDORES", "Servicio Interbanking (abono)"],
  ["4712", "Débito", "", "IMPUESTOS Y RETENCIONES", "Pago AFIP por Interbanking (VEP)"],
  ["4637", "Débito", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 sobre créditos"],
  ["4633", "Débito", "", "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 sobre débitos"],
  ["1384", "Débito", "", "IMPUESTOS Y RETENCIONES", "Recaudación IIBB SIRCREB"],
  ["566", "Débito", "", "IMPUESTOS Y RETENCIONES", "Adelanto IIBB Tucumán"],
  ["3254", "Débito", "", "GASTOS BANCARIOS", "IVA 21% sobre comisiones"],
  ["2872", "Débito", "", "GASTOS BANCARIOS", "Comisión pago servicio Interbanking"],
  ["2960", "Débito", "", "GASTOS BANCARIOS", "Comisión por servicio de cuenta"],
];

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

function onOpen() { crearMenu(); }

/** Menú de la planilla. Con la biblioteca, la planilla de cada banco llama a Motor.crearMenu() en su onOpen. */
function crearMenu() {
  SpreadsheetApp.getUi().createMenu('Conciliación')
    .addItem('1. Crear hojas de entrada', 'crearHojasEntrada')
    .addItem('2. Procesar conciliación', 'procesarConciliacion')
    .addItem('3. Cerrar mes (guardar foto y pasar pendientes)', 'cerrarMes')
    .addSeparator()
    .addItem('Revisar pendientes (con la base al día)', 'revisarPendientes')
    .addItem('Agregar reglas nuevas del script', 'agregarReglasNuevas')
    .addToUi();
}

function crearParametros_(ss) {
  const filas = [
    ['Parámetro', 'Valor'],
    ['Cuenta E', 'COMPLETAR'],
    ['Cuenta O', 'COMPLETAR'],
    ['Saldo inicial cuenta E', ''],
    ['Saldo inicial cuenta O', ''],
    ['Cuenta bancaria', 'COMPLETAR'],
    ['Base FBS (link o ID)', ''],
    ['Hoja de la base', 'MOVIMIENTOS'],
    ['Carpeta de cierres (link o ID)', ''],
    ['Sin cuenta O', 'Si la empresa tiene la confirmación de valores al día (ej. Chexa), poné NO en "Cuenta O": se concilia solo la cuenta E.'],
    ['Base FBS', 'Link de la planilla base de la empresa (la que se actualiza todos los días). Con la base, los movimientos del mes y los saldos iniciales de E y O se toman solos de ahí; no hace falta la hoja "Base".'],
    ['Carpeta de cierres', 'Carpeta de Drive donde "Cerrar mes" guarda la foto de cada cierre. Vacía: "Conciliaciones - Historial" junto a esta planilla.'],
    ['Saldos iniciales', 'Con la hoja "Base": saldo final de E y de O de la conciliación anterior. Al cerrar el mes se completan solos.'],
  ];
  let sh = ss.getSheetByName(HOJA.parametros);
  if (sh) {   // agrega los parámetros que falten sin tocar los existentes
    const hay = sh.getDataRange().getValues().map(r => String(r[0]));
    filas.filter(f => hay.indexOf(f[0]) < 0).forEach(f => sh.appendRow(f));
    return;
  }
  sh = ss.insertSheet(HOJA.parametros);
  sh.getRange(1, 1, filas.length, 2).setValues(filas);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.getRange(4, 2, 2, 1).setNumberFormat(NUM_FMT);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 700);
}

/** Primera hoja existente con datos entre varios nombres posibles (o la primera que exista). */
function hojaDe_(ss, nombres) {
  const existentes = nombres.map(n => ss.getSheetByName(n)).filter(Boolean);
  return existentes.find(sh => sh.getLastRow() > 1) || existentes[0] || null;
}
const nombresExtracto_ = () => [HOJA.extracto].concat(HOJA.extractoAlt);
const nombresAnteriores_ = () => HOJA.anterioresAlt.concat([HOJA.anteriores]);

/** Borra hojas de entrada vacías que quedaron de la versión con mayores (si se trabaja con la hoja Base). */
function limpiarHojasSobrantes_(ss) {
  const vacia = n => { const sh = ss.getSheetByName(n); return sh && sh.getLastRow() <= 1 ? sh : null; };
  const conDatos = n => { const sh = ss.getSheetByName(n); return sh && sh.getLastRow() > 1; };
  const borrar = [];
  if (conDatos(HOJA.base)) borrar.push(vacia(HOJA.mayorE), vacia(HOJA.mayorO));
  if (HOJA.extractoAlt.some(conDatos)) borrar.push(vacia(HOJA.extracto));
  if (HOJA.anterioresAlt.some(n => ss.getSheetByName(n))) borrar.push(vacia(HOJA.anteriores));
  borrar.filter(Boolean).forEach(sh => ss.deleteSheet(sh));
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
  // solo crea lo que falta. FBS se lee de la base de la empresa (link en Parámetros); los mayores E y O ya no se usan
  if (!hojaDe_(ss, nombresExtracto_())) crear(HOJA.extractoAlt[0], [['Pegar acá el extracto del banco desde A1, con la fila de encabezados (Macro: Excel "Últimos Movimientos" o CSV "Resumen"; Galicia: Excel de Office Banking)']]);
  if (!hojaDe_(ss, nombresAnteriores_())) crear(HOJA.anterioresAlt[0], [COLS_PEND]);
  limpiarHojasSobrantes_(ss);
  // el banco define las reglas iniciales: se pregunta la primera vez (cuando todavía no hay Parámetros)
  if (!ss.getSheetByName(HOJA.parametros)) {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('Banco de esta planilla', 'Escribí el nombre de la cuenta (ej. "Banco Macro", "Banco Galicia"):', ui.ButtonSet.OK);
    crearParametros_(ss);
    const nombre = String(r.getResponseText() || '').trim();
    if (nombre) {
      const par = ss.getSheetByName(HOJA.parametros), v = par.getDataRange().getValues();
      const i = v.findIndex(x => /^cuenta bancaria$/i.test(String(x[0])));
      if (i >= 0) par.getRange(i + 1, 2).setValue(nombre);
    }
  }
  SpreadsheetApp.flush();
  const banco = paramDe_(ss.getSheetByName(HOJA.parametros).getDataRange().getValues(), '^cuenta bancaria$');
  crear(HOJA.reglas, [['Código causal', 'Sentido', 'Patrón en concepto', 'Categoría', 'Descripción']].concat(reglasDelBanco_(banco)));
  asegurarReglasDelBanco_(ss, banco);
  crear(HOJA.empresas, [['CUIT', 'Razón social']].concat(EMPRESAS_INICIALES));
  crearParametros_(ss);
  SpreadsheetApp.getUi().alert('Hojas de entrada listas. Cargá el link de la base FBS de la empresa en "' + HOJA.parametros + '", el extracto y los pendientes anteriores ' +
    'y los saldos iniciales en "' + HOJA.parametros + '", y corré "2. Procesar conciliación".');
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
  limpiarHojasSobrantes_(ss);
  const hojaExtracto = (hojaDe_(ss, nombresExtracto_()) || { getName: () => HOJA.extracto }).getName();
  const hojaAnteriores = (hojaDe_(ss, nombresAnteriores_()) || { getName: () => HOJA.anteriores }).getName();
  const parametros = ss.getSheetByName(HOJA.parametros) ? leer(HOJA.parametros) : null;
  const extracto = leer(hojaExtracto), empresas = leer(HOJA.empresas);
  asegurarReglasDelBanco_(ss, paramDe_(parametros, '^cuenta bancaria$'));
  const reglas = leer(HOJA.reglas);
  const idBase = idDeLink_(paramDe_(parametros, '^base fbs'));
  let base = null, saldosIniciales = null, avisosBase = [], tardios = null, ctxBase = null, hastaBase = null;
  if (idBase) {
    // base de la empresa en otra planilla: se trae solo lo de las cuentas E y O del mes del extracto
    const movs = leerExtracto_(extracto, leerReglas_(reglas), leerEmpresas_(empresas)).movs;
    if (!movs.length) throw new Error('El extracto está vacío.');
    const desde = new Date(movs[0].fecha.getFullYear(), movs[0].fecha.getMonth(), 1);
    const ult = movs[movs.length - 1].fecha;
    const hasta = new Date(ult.getFullYear(), ult.getMonth() + 1, 0);
    const r = leerBaseExterna_(idBase, paramDe_(parametros, '^hoja de la base') || 'MOVIMIENTOS', cuentasDe_(parametros), desde, hasta);
    base = r.filas; saldosIniciales = r.saldos; ctxBase = r.ctx || null; hastaBase = hasta;
    // movimientos de meses cerrados cargados o modificados después del cierre anterior
    const cierre = { E: parseNum_(paramDe_(parametros, '^saldo inicial cuenta e')), O: parseNum_(paramDe_(parametros, '^saldo inicial cuenta o')) };
    const foto = leerFotoBase_(ss);
    if (r.ctx && (foto || cierre.E !== null || cierre.O !== null)) {
      const dif = { E: cierre.E === null ? 0 : r.saldos.E - cierre.E, O: cierre.O === null ? 0 : r.saldos.O - cierre.O };
      if (foto || Math.abs(dif.E) > 1 || Math.abs(dif.O) > 1) {
        const t = movimientosTardios_(r.ctx, desde, dif, foto);
        if (t.filas.length > 1) { base = unirFilas_(base, t.filas); tardios = t.total; }
      }
    }
    if (ss.getSheetByName(HOJA.base) && ss.getSheetByName(HOJA.base).getLastRow() > 1) {
      avisosBase.push('Se usó la base externa de "' + HOJA.parametros + '"; la hoja "' + HOJA.base + '" de esta planilla se ignoró (se puede borrar).');
    }
  } else if (ss.getSheetByName(HOJA.base)) {
    base = leer(HOJA.base);
  } else if (!ss.getSheetByName(HOJA.mayorE)) {
    throw new Error('Falta la base de FBS: cargá el link de la base de la empresa en "' + HOJA.parametros + '" (fila "Base FBS").');
  }
  const res = conciliarTodo_({
    extracto, base, saldosIniciales, avisosBase, tardios,
    mayorE: base ? null : leer(HOJA.mayorE),
    mayorO: base ? null : leer(HOJA.mayorO),
    anteriores: leer(hojaAnteriores),
    reglas, empresas, parametros,
  });
  res.baseCtx = ctxBase; res.hastaBase = hastaBase;
  HOJA.viejas.forEach(n => { const v = ss.getSheetByName(n); if (v) ss.deleteSheet(v); });
  escribirCuatroPuntas_(ss, res);
  escribirTablero_(ss, res);
  escribirRevisar_(ss, res);
  escribirAnalisisO_(ss, res);
  escribirPendientes_(ss, HOJA.pendO, res, res.listas.O,
    'Paso 1 — Cuenta O: líneas que no netean (no confirmadas). Van a Depósitos no acreditados / Cheques no debitados.');
  escribirConciliados_(ss, res);
  escribirPendientes_(ss, HOJA.pendFbs, res, res.listas.E,
    'Paso 2 — Cuenta E que no se cruzó ni con pendientes anteriores ni con el extracto (posibles errores de registración).');
  escribirPendientes_(ss, HOJA.pendBanco, res, res.listas.B,
    'Movimientos del extracto (y pendientes anteriores del banco) que no están en la cuenta E.');
  // orden de las pestañas: primero lo que se lee
  [HOJA.cuatro, HOJA.tablero, HOJA.revisar].forEach((n, i) => { const h = ss.getSheetByName(n); if (h) { ss.setActiveSheet(h); ss.moveActiveSheet(i + 1); } });
  ss.getSheetByName(HOJA.cuatro).activate();
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
  const empresa = res.empresa || res.info.empresa || '';
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
  let nombre = 'Conciliación ' + String(res.cuentaBanco || 'Banco Macro').replace(/^banco\s+/i, '') + (empresa ? ' - ' + empresa : '') + ' - ' + periodo;
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

  // 4. pendientes al mes siguiente (y los saldos finales de E y O pasan a ser los iniciales)
  const n = pasarPendientes_(ss, res.corte);
  const par = ss.getSheetByName(HOJA.parametros);
  // foto de E y O en la base al cierre: el mes siguiente detecta lo que se cargue o cambie después con fecha de meses cerrados
  if (par) {   // fecha del último cierre: desde ahí "Revisar pendientes" busca movimientos nuevos en la base
    const v = par.getDataRange().getValues();
    let i = v.findIndex(r => /^ultimo cierre/i.test(normalizar_(r[0])));
    if (i < 0) { par.appendRow(['Último cierre', '']); i = par.getLastRow() - 1; }
    par.getRange(i + 1, 2).setValue(res.hastaBase || new Date(res.corte.getFullYear(), res.corte.getMonth() + 1, 0)).setNumberFormat('dd/mm/yyyy');
  }
  if (res.baseCtx) {
    try { guardarFotoBase_(ss, res.baseCtx, res.hastaBase); } catch (e) {
      ui.alert('No se pudo guardar la foto de la base (' + e.message + '). El mes siguiente se controla solo por saldo.');
    }
  }
  if (par) {
    const v = par.getDataRange().getValues();
    v.forEach((r, i) => {
      if (/^saldo inicial cuenta e$/i.test(String(r[0]))) par.getRange(i + 1, 2).setValue(res.saldoE);
      if (/^saldo inicial cuenta o$/i.test(String(r[0]))) par.getRange(i + 1, 2).setValue(res.saldoO);
    });
  }
  if (ui.alert('Mes cerrado', 'Foto guardada en la carpeta "' + carpeta.getName() + '" y registrada en "Historial".\n' + n +
      ' partidas pasadas a "' + HOJA.anterioresAlt[0] + '".\n\n¿Vaciar la hoja del extracto para cargar el mes siguiente? (FBS se lee de la base, no hay que borrar nada más)',
      ui.ButtonSet.YES_NO) === ui.Button.YES) {
    [HOJA.extracto].concat(HOJA.extractoAlt).concat(res.baseCtx ? [] : [HOJA.mayorE, HOJA.mayorO, HOJA.base])
      .forEach(h => { const sh = ss.getSheetByName(h); if (sh) sh.clearContents(); });
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
  const par = ss.getSheetByName(HOJA.parametros);
  const link = par ? paramDe_(par.getDataRange().getValues(), '^carpeta de cierres') : '';
  const idCarpeta = String(link || '').match(/folders\/([a-zA-Z0-9_-]{10,})/) || String(link || '').trim().match(/^([a-zA-Z0-9_-]{10,})$/);
  if (idCarpeta) {
    try { return DriveApp.getFolderById(idCarpeta[1]); } catch (e) {
      throw new Error('No puedo abrir la carpeta de cierres de "' + HOJA.parametros + '": revisá el link y que tengas permiso de edición.');
    }
  }
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

const COLS_PEND = ['Sector', 'Fecha', 'Concepto', 'Importe', 'ID', 'Días', 'Estado', 'Responsable', 'Comentario', 'Revisión automática', 'Revisado el'];
const ESTADOS_PEND = ['Pendiente', 'En gestión', 'Resuelto en FBS', 'Justificado', 'Error de la lista'];

/** Clave estable de un pendiente (sector, fecha, concepto, importe) para conservar estado y comentarios entre cierres. */
function clavePend_(sector, fecha, concepto, importe) {
  const f = parseFecha_(fecha);
  return [String(sector).trim().toUpperCase(), f ? fechaTexto_(f) : '', normalizar_(concepto).replace(/\s+/g, ' '),
    Math.abs(parseNum_(importe) || 0).toFixed(2)].join('|');
}

/** Agrega el ID a cada fila (con #2, #3 si se repite). filas sin encabezado: [sector, fecha, concepto, importe]. */
function idsPend_(filas) {
  const vistos = {};
  return filas.map(r => { const k = clavePend_(r[0], r[1], r[2], r[3]); vistos[k] = (vistos[k] || 0) + 1; return vistos[k] > 1 ? k + '#' + vistos[k] : k; });
}

/**
 * Arma "Pendientes en trabajo" con los pendientes del cierre: Sector, Fecha, Concepto, Importe (lo que lee el proceso)
 * y columnas para trabajarlos (Estado, Responsable, Comentario). Lo que ya estaba en la lista conserva su estado y
 * comentarios; lo resuelto no pasa.
 */
function pasarPendientes_(ss, corte) {
  const nuevas = [];
  [HOJA.pendO, HOJA.pendFbs, HOJA.pendBanco].forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) throw new Error('Primero corré "2. Procesar conciliación".');
    const v = sh.getDataRange().getValues();
    const h = v[1];   // fila 1 = título, fila 2 = encabezados
    const c = { s: h.indexOf('Sector'), f: h.indexOf('Fecha'), t: h.indexOf('Concepto / Comprobante'), i: h.indexOf('Importe') };
    v.slice(2).forEach(r => { if (r[c.s]) nuevas.push([r[c.s], r[c.f], r[c.t], r[c.i]]); });
  });
  let sh = hojaDe_(ss, nombresAnteriores_());
  // lo trabajado en la lista anterior (por ID)
  const previo = {};
  if (sh && sh.getLastRow() > 1) {
    const v = sh.getDataRange().getValues(), h = v[0].map(String);
    const ix = n => h.indexOf(n);
    const ids = ix('ID') >= 0 ? v.slice(1).map(r => String(r[ix('ID')])) : idsPend_(v.slice(1));
    v.slice(1).forEach((r, i) => {
      if (ix('Estado') < 0) return;
      previo[ids[i]] = [r[ix('Estado')], r[ix('Responsable')], r[ix('Comentario')]];
    });
  }
  // se escribe siempre en "Pendientes en trabajo"; la lista vieja ("Pend mes anterior") ya quedó absorbida y se borra
  const destino = ss.getSheetByName(HOJA.anterioresAlt[0]);
  if (destino && sh && destino.getName() !== sh.getName()) { ss.deleteSheet(sh); sh = destino; }
  if (!sh) sh = destino || ss.insertSheet(HOJA.anterioresAlt[0]);
  if (sh.getName() !== HOJA.anterioresAlt[0]) sh.setName(HOJA.anterioresAlt[0]);
  nombresAnteriores_().forEach(n => { const otra = ss.getSheetByName(n); if (otra && otra.getName() !== sh.getName()) ss.deleteSheet(otra); });
  const ids = idsPend_(nuevas);
  const filas = nuevas.map((r, i) => {
    const p = previo[ids[i]] || [];
    const f = parseFecha_(r[1]);
    const estado = p[0] && !/^Resuelto/i.test(p[0]) ? p[0] : 'Pendiente';
    // con signo: un pendiente anterior en negativo (comisión de una liquidación en S3) sigue en negativo
    return [r[0], r[1], r[2], parseNum_(r[3]) || 0, ids[i], f && corte ? dias_(corte, f) : '', estado, p[1] || '', p[2] || '', '', ''];
  });
  escribirPendTrabajo_(sh, filas);
  return filas.length;
}

function escribirPendTrabajo_(sh, filas) {
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  sh.getRange(1, 1, 1, COLS_PEND.length).setValues([COLS_PEND]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  if (filas.length) {
    sh.getRange(2, 1, filas.length, COLS_PEND.length).setValues(filas);
    sh.getRange(2, 2, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 4, filas.length, 1).setNumberFormat(NUM_FMT);
    sh.getRange(2, 11, filas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(2, 7, filas.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_PEND, true).setAllowInvalid(false).build());
    sh.getRange(2, 10, filas.length, 1).setWrap(true);
  }
  sh.getRange(1, 1, Math.max(filas.length, 1) + 1, COLS_PEND.length).createFilter();
  [55, 85, 360, 120, 60, 50, 120, 110, 260, 380, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.hideColumns(5);   // ID
}

// ------------------------------------------------------------------ salida

function hojaLimpia_(ss, nombre) {
  let sh = ss.getSheetByName(nombre);
  if (sh) { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart().clearNote(); sh.clear(); sh.clearConditionalFormatRules(); if (sh.getFilter()) sh.getFilter().remove(); }
  else sh = ss.insertSheet(nombre);
  return sh;
}

/**
 * Conciliación con el formato de "4 puntas" que usa administración: saldo contable arriba, los cuatro cuadros
 * (izquierda depósitos, derecha cheques / débitos; arriba no registrados en contabilidad, abajo no pasados por el
 * banco) y abajo el saldo ajustado contra el extracto. Columnas: Fecha | Concepto | Importe | (separador) | Fecha | Concepto | Importe.
 * Devuelve { filas, estilos, notas } con estilos 'titulo' | 'cuadro' | 'cab' | 'item' | 'total' | 'saldo' | 'dif' | ''.
 */
function armarCuatroPuntas_(res) {
  const filas = [], estilos = [], notas = [];
  const add = (f, e, n) => { while (f.length < 7) f.push(''); filas.push(f); estilos.push(e || ''); notas.push(n || ['', '']); };
  const orden = l => l.slice().sort((a, b) => a.fecha - b.fecha || Math.abs(b.importe) - Math.abs(a.importe));
  const nota = p => (p ? observacion_(p) + (p.cruces ? '\n' + p.cruces : '') : '');
  add(['Saldo Contabilidad al', '', res.corte, '', '', 'E', res.saldoE], 'titulo');
  add(['', res.cuentaBanco || 'Banco Macro', res.saldoE + res.saldoO, '', '', 'O', res.saldoO], 'titulo');
  add([]);
  [['S1', 'Depósitos No Registrados en Contabilidad', 'S2', 'Cheques No Registrados en Contabilidad'],
    ['S3', 'Depósitos No Acreditados en BANCO', 'S4', 'Cheques No Debitados en BANCO']].forEach(([si, ni, sd, nd]) => {
    // una partida agrupada (liquidación + comisiones, operación) se muestra con sus líneas originales
    const expandir = l => [].concat.apply([], orden(l).map(p => (p.lineas && p.lineas.length > 1
      ? p.lineas.slice().sort((x, y) => (p.importe < 0 ? x.importe - y.importe : y.importe - x.importe)).map((x, i) => ({ fecha: x.fecha, texto: x.texto, valor: p.importe < 0 ? -x.importe : x.importe, p: i ? null : p }))
      : [{ fecha: p.fecha, texto: p.texto, valor: valorSector_(p), p }])));
    const iz = expandir(res.pend[si]), de = expandir(res.pend[sd]);
    add([ni, '', '', '', nd, '', ''], 'cuadro');
    add(['Fecha', 'Concepto', 'Importe', '', 'Fecha', 'Concepto', 'Importe'], 'cab');
    for (let i = 0; i < Math.max(iz.length, de.length); i++) {
      const a = iz[i], b = de[i];
      add([a ? a.fecha : '', a ? a.texto : '', a ? a.valor : '', '', b ? b.fecha : '', b ? b.texto : '', b ? b.valor : ''],
        'item', [nota(a && a.p), nota(b && b.p)]);
    }
    add(['', res.pend[si].length + ' partidas', res.tot[si], '', '', res.pend[sd].length + ' partidas', res.tot[sd]], 'total');
    add([]);
  });
  add(['Saldo Contabilidad AJUSTADO al', res.corte, res.ajustado], 'saldo');
  add(['Saldo Extracto al', res.corte, res.saldoBanco], 'saldo');
  add(['Diferencia', '', redondear_(res.ajustado - res.saldoBanco)], 'dif');
  return { filas, estilos, notas };
}

function escribirCuatroPuntas_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.cuatro);
  const { filas, estilos, notas } = armarCuatroPuntas_(res);
  sh.getRange(1, 1, filas.length, 7).setValues(filas);
  sh.getRange(1, 1, filas.length, 7).setFontFamily('Arial').setFontSize(10);
  [3, 7].forEach(c => sh.getRange(1, c, filas.length, 1).setNumberFormat('#,##0.00;(#,##0.00)'));
  [1, 5].forEach(c => sh.getRange(4, c, filas.length - 3, 1).setNumberFormat('dd/mm/yyyy'));
  sh.getRange(1, 3).setNumberFormat('dd/mm/yyyy');
  estilos.forEach((e, i) => {
    const f = i + 1;
    if (e === 'titulo') sh.getRange(f, 1, 1, 7).setFontWeight('bold');
    if (e === 'cuadro') [1, 5].forEach(c => sh.getRange(f, c, 1, 3).merge().setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM));
    if (e === 'cab') [1, 5].forEach(c => sh.getRange(f, c, 1, 3).setFontWeight('bold').setBackground('#D9E1F2'));
    if (e === 'total') [3, 7].forEach(c => sh.getRange(f, c).setFontWeight('bold').setBorder(true, true, true, true, null, null));
    if (e === 'saldo') { sh.getRange(f, 1, 1, 3).setFontWeight('bold'); sh.getRange(f, 2).setNumberFormat('dd/mm/yyyy').setBorder(true, true, true, true, null, null); }
    if (e === 'dif') sh.getRange(f, 1, 1, 3).setFontWeight('bold').setBackground(Math.abs(filas[i][2]) < 1000 ? '#C6EFCE' : '#FFC7CE');
  });
  // el motivo y la pista de cada partida quedan como nota del concepto (pasar el mouse)
  const n = notas.map(x => [x[0]]), m = notas.map(x => [x[1]]);
  sh.getRange(1, 2, filas.length, 1).setNotes(n);
  sh.getRange(1, 6, filas.length, 1).setNotes(m);
  [85, 380, 130, 30, 85, 380, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(2);
}

function escribirTablero_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.tablero);
  const filas = [];
  const estilos = [];   // 'titulo' | 'cab' | 'total' | 'dif' | ''
  const add = (fila, estilo) => { while (fila.length < 3) fila.push(''); filas.push(fila); estilos.push(estilo || ''); };
  const emp = res.empresa || res.info.empresa;
  add(['CONCILIACIÓN ' + String(res.cuentaBanco || 'Banco Macro').toUpperCase() + (emp ? ' — ' + emp : '') + ' — al ' + fechaTexto_(res.corte)], 'titulo');
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
  if (p.tardio) return 'Movimiento con fecha ' + fechaTexto_(p.fecha) + ' (mes ya cerrado) cargado o modificado en FBS después del cierre';
  if (p.desfasado) return 'Confirmado con fecha ' + fechaTexto_(p.fecha) + ' pero el recibo se registró el ' + fechaTexto_(p.fechaRef2) +
    ': queda en S4 y se cancela solo el mes próximo, cuando entra el ingreso en O';
  if (p.fechaRef2 && CORTE_ && p.fechaRef2 > CORTE_) return 'Cargado después del cierre (' + fechaTexto_(p.fechaRef2) + ') con fecha del mes';
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
    p.texto, valorSector_(p), p.cat, p.cuit || p.dni, p.asiento || p.ref, p.cuenta || '', observacion_(p), p.cruces || '']);
  const porSector = {};
  items.forEach(p => { const k = sector_(p); porSector[k] = porSector[k] || [0, 0]; porSector[k][0]++; porSector[k][1] += valorSector_(p); });
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

/**
 * Lista de trabajo: qué hacer con cada partida del mes, agrupado por acción y ordenado por importe.
 * Devuelve filas [tipo, col1..col6] donde tipo = 'titulo' | 'resumen' | 'seccion' | 'sub' | 'cab' | 'item' | ''.
 */
function armarRevisar_(res) {
  const out = [];
  const fila = (tipo, ...v) => { while (v.length < 6) v.push(''); out.push([tipo].concat(v)); };
  const suma = l => l.reduce((x, p) => x + Math.abs(p.importe), 0);
  const mes = l => l.filter(p => p.origen === 'Mes').sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe));
  const B = mes(res.listas.B), E = mes(res.listas.E);
  const O = mes(res.listas.O).filter(p => !p.desfasado), D = mes(res.listas.O).filter(p => p.desfasado);
  const revisar = res.conciliados.filter(c => /revisar/i.test(c[0]));
  const ant = [].concat(res.listas.B, res.listas.O, res.listas.E).filter(p => p.origen === 'Arrastre');
  const signo = p => (p.importe < 0 ? 'Débito' : 'Crédito');

  fila('titulo', 'PARA REVISAR — ' + fechaTexto_(res.corte) + '   (diferencia de control ' + formato_(res.ajustado - res.saldoBanco) + ')');
  fila('');
  fila('cab', 'Qué hay que hacer', 'Partidas', 'Importe', '', '', '');
  fila('resumen', '1. Registrar en FBS: movimientos del banco del mes que no están en FBS', B.length, suma(B));
  fila('resumen', '2. Confirmar en FBS: registros de la cuenta O sin confirmar', O.length, suma(O));
  fila('resumen', '3. Corregir en FBS: registros de la cuenta E que no aparecen en el banco', E.length, suma(E));
  fila('resumen', '4. Validar: cruces automáticos marcados "(revisar)"', revisar.length, '');
  fila('resumen', '5. Pendientes de meses anteriores que siguen abiertos', ant.length, suma(ant));
  fila('resumen', '6. Desfasados: confirmados en el mes y registrados después del cierre (no hay que hacer nada)', D.length, suma(D));

  const seccion = (titulo, items, motivo, pista) => {
    fila('');
    fila('seccion', titulo);
    if (!items.length) { fila('', 'Sin partidas.'); return; }
    fila('cab', 'Fecha', 'Concepto / comprobante', 'Importe', 'Qué pasa', 'Pista para cruzar', '');
    items.forEach(p => fila('item', p.fecha, p.texto, p.importe, motivo(p), pista(p) || '', ''));
  };

  // 1. banco sin registrar, agrupado por tipo de movimiento
  fila('');
  fila('seccion', '1. REGISTRAR EN FBS — el banco muestra el movimiento y FBS no lo tiene');
  if (!B.length) fila('', 'Sin partidas.');
  const porCat = new Map();
  B.forEach(p => { const k = p.cat || 'OTROS'; porCat.set(k, (porCat.get(k) || []).concat([p])); });
  Array.from(porCat.entries()).sort((a, b) => suma(b[1]) - suma(a[1])).forEach(([cat, l]) => {
    fila('sub', cat + ' — ' + l.length + ' partidas', '', suma(l));
    fila('cab', 'Fecha', 'Concepto del banco', 'Importe', 'Qué pasa', 'Pista para cruzar', '');
    l.forEach(p => fila('item', p.fecha, p.texto, p.importe, signo(p) + ' sin registrar', p.cruces || (p.cuit ? 'CUIT ' + p.cuit : '')));
  });

  seccion('2. CONFIRMAR EN FBS — registrado en la cuenta O pero sin confirmar', O,
    p => observacion_(p), p => p.cruces);
  seccion('3. CORREGIR EN FBS — registrado en la cuenta E y no aparece en el banco', E,
    p => observacion_(p).replace('Pendiente en cuenta E: ', ''), p => p.cruces);

  fila('');
  fila('seccion', '4. VALIDAR — cruces automáticos por importe (sin CUIT, referencia ni nombre)');
  if (!revisar.length) fila('', 'Sin partidas.');
  else {
    fila('cab', 'Fecha banco', 'Concepto del banco', 'Importe', 'Cruzado con (FBS)', 'Criterio', '');
    revisar.forEach(c => fila('item', c[1], c[2], c[3], (c[6] || '') + (c[8] ? ' (asiento ' + c[8] + ')' : ''), c[0]));
  }

  fila('');
  fila('seccion', '5. PENDIENTES DE MESES ANTERIORES — resumen por antigüedad (detalle en las hojas de pendientes)');
  fila('cab', 'Antigüedad', 'Partidas', 'Importe', '', '', '');
  [['Hasta 30 días', 0, 30], ['31 a 90 días', 31, 90], ['91 a 365 días', 91, 365], ['Más de un año', 366, 1e9]].forEach(([n, d0, d1]) => {
    const l = ant.filter(p => { const d = dias_(res.corte, p.fecha); return d >= d0 && d <= d1; });
    fila('item', n, l.length, suma(l));
  });

  seccion('6. DESFASADOS — confirmados con fecha del mes, el recibo se registró después del cierre (quedan en S4 y se cancelan solos el mes próximo)', D,
    p => observacion_(p), p => p.cruces);
  return out;
}

function escribirRevisar_(ss, res) {
  const sh = hojaLimpia_(ss, HOJA.revisar);
  const filas = armarRevisar_(res);
  sh.getRange(1, 1, filas.length, 6).setValues(filas.map(f => f.slice(1)));
  const fmt = [];
  filas.forEach((f, i) => {
    const r = sh.getRange(i + 1, 1, 1, 6);
    if (f[0] === 'titulo') r.setFontWeight('bold').setFontSize(14);
    if (f[0] === 'seccion') r.setFontWeight('bold').setFontSize(12).setBackground('#1F4E78').setFontColor('#FFFFFF');
    if (f[0] === 'sub') r.setFontWeight('bold').setBackground('#DDEBF7');
    if (f[0] === 'cab') r.setFontWeight('bold').setBackground('#F2F2F2');
    if (f[0] === 'resumen') sh.getRange(i + 1, 1).setFontWeight('bold');
    if (f[0] === 'item' || f[0] === 'resumen' || f[0] === 'sub') fmt.push(i + 1);
  });
  fmt.forEach(r => sh.getRange(r, 3).setNumberFormat(NUM_FMT));
  sh.getRange(1, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
  filas.forEach((f, i) => { if (f[0] === 'resumen' || f[0] === 'sub') sh.getRange(i + 1, 1).setNumberFormat('@'); });
  sh.getRange(1, 1, filas.length, 6).setWrap(true).setVerticalAlignment('top');
  [95, 380, 130, 300, 380, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
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
  const pats = [/\s-\s+De\s+(.+?)\s*\//i, /^(?:Pago a proveedores recibido|Transferencia pagos a terceros|Servicios de pago) - (.+?)\s{2,}/i,
    /ING TRANSF:(.+?)-\d{11}/, /^TRANSF (.+?) \d{11}/, /TEF DATANET \w+ (.+?)\s+\d{11}/,
    /^CCERR (.+?) \d{11}/, /^TRMIN (.+?) \d{11}/, /^TRPRO (.+?) \d{11}/, /^\d{11} (.+?) DL1/, /Transf\. \w+ (.+?)\s+\d{11}/,
    // Patagonia (al final, para no pisar a Macro): "... VAR 20323224545  LEANDRO JESUS GRASSI", "... 30500007539 MAPFRE ARG SEGUR"
    /^(?:CREDITO POR|TRANSFERENCIA|TRANSF\.).*?\d{11}\s+([A-ZÑ][A-ZÑ .,]+?)\s*$/i,
    // Córdoba: "... 20340699379-ERIC MAXIMILIANO HERRE-VAR"
    /\d{11}\s*-\s*(.+?)\s*-\s*(?:VAR|CUO|FAC)/i];
  for (const p of pats) {
    const m = concepto.match(p);
    if (m) return m[1].replace(/\s+/g, ' ').replace(/^[ /,]+|[ /,]+$/g, '');
  }
  return '';
}

function tokens_(texto) {
  const out = new Set();
  (String(texto || '').toUpperCase().replace(/&/g, 'Ñ').match(/[A-ZÑ]{3,}/g) || []).forEach(t => { if (!PALABRAS_VACIAS.has(t)) out.add(t); });
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
      patron: pat && pat !== '(todos)' ? new RegExp(pat, 'i') : null,   // sin distinguir mayúsculas (Córdoba: "Recaudaciones SIRCREB")
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
    fecha: /^fecha$/, ref: /referencia|^numero de comprobante/, causal: /causal|^cod\.? operativo/, concepto: /^concepto/, importe: /^importe|^monto/,
    debito: /^debito/, credito: /^credito/, saldo: /^saldo/, descripcion: /^descripcion/,
    ley1: /^leyendas? adicional(es)?\s*1/, ley2: /^leyendas? adicional(es)?\s*2/, ley3: /^leyendas? adicional(es)?\s*3/,
    ley4: /^leyendas? adicional(es)?\s*4/,
  }, ['fecha', 'saldo']);
  if (!h || (h.col.concepto === undefined && h.col.descripcion === undefined)) {
    throw new Error('Extracto: no encuentro los encabezados (Fecha, Concepto o Descripción, Saldo).');
  }
  const c = h.col, info = {}, movs = [];
  // Santander: el Excel trae primero "Movimientos del Día" (sin saldo, todavía no contabilizados) y después
  // "Últimos Movimientos"; solo se toma el segundo bloque
  let delDia = false;
  filas.forEach((r, i) => {
    const t0 = String(r[0] || '');
    if (/^Empresa:/i.test(t0)) info.empresa = t0.split(':').slice(1).join(':').trim();
    if (/^movimientos del d[ií]a/i.test(t0)) delDia = true;
    if (/^[úu]ltimos movimientos/i.test(t0)) delDia = false;
    if (i <= h.fila || delDia) return;
    const fecha = parseFecha_(r[c.fecha]);
    if (!fecha) return;
    // filas sueltas con solo una fecha (ej. fecha y hora de la descarga al pie)
    if (!String(r[c.concepto !== undefined ? c.concepto : c.descripcion] || '').trim()) return;
    let deb, cred;
    if (c.importe !== undefined) {
      const imp = parseNum_(r[c.importe]) || 0;
      deb = Math.max(-imp, 0); cred = Math.max(imp, 0);
    } else {
      deb = Math.abs(parseNum_(r[c.debito]) || 0); cred = Math.abs(parseNum_(r[c.credito]) || 0);
    }
    let concepto, causal, nombre;
    if (c.descripcion !== undefined) {
      // Galicia (Office Banking): Descripción + leyendas (nombre, CUIT, motivo, banco); el código de "Concepto" es el causal
      const ley = ['ley1', 'ley2', 'ley3', 'ley4'].map(k => (c[k] !== undefined ? String(r[c[k]] == null ? '' : r[c[k]]).trim() : ''));
      // Patagonia: sin leyendas ni código; la Referencia trae el CUIT y el nombre ("30500007539 MAPFRE ARG SEGUR")
      const sinLey = c.concepto === undefined && ['ley1', 'ley2', 'ley3', 'ley4'].every(k => c[k] === undefined);
      const extra = sinLey && c.ref !== undefined ? [String(r[c.ref] == null ? '' : r[c.ref]).trim()] : [];
      concepto = [String(r[c.descripcion] || '').trim()].concat(ley.filter(x => x && x !== '0'), extra).join(' ').replace(/\s+/g, ' ').trim();
      const cod = c.concepto !== undefined ? String(r[c.concepto] || '').match(/^\s*(\d{4,})/) : null;
      // Córdoba: "Concepto" es el tipo de operación (texto, sin código) y va adelante de la descripción
      if (c.concepto !== undefined && !cod && String(r[c.concepto] || '').trim()) {
        concepto = (String(r[c.concepto]).trim() + ' ' + concepto).replace(/\s+/g, ' ').trim();
      }
      causal = cod ? cod[1] : (c.causal !== undefined ? String(r[c.causal]).trim() : '');
      nombre = ley[0] && !/^\d+$/.test(ley[0]) ? ley[0] : extraerNombre_(concepto);
    } else {
      concepto = String(r[c.concepto] || '').trim();
      causal = c.causal !== undefined ? String(r[c.causal]).trim() : '';
      nombre = extraerNombre_(concepto);
    }
    const importe = redondear_(cred - deb);
    const cuit = extraerCuit_(concepto);
    const [cat, desc] = clasificar_(reglas, empresas, causal, concepto, importe, cuit);
    const ref = c.ref !== undefined ? String(r[c.ref] == null ? '' : r[c.ref]).trim() : '';
    movs.push({ fecha, ref: ref === '0' ? '' : ref, causal, concepto, importe,
      saldo: parseNum_(r[c.saldo]), cat, desc, cuit, nombre });
  });
  if (!movs.length) throw new Error('Extracto: no hay movimientos con fecha debajo de los encabezados.');
  if (movs[0].fecha > movs[movs.length - 1].fecha) movs.reverse();   // el banco lo entrega del más nuevo al más viejo
  // si el extracto trae más de un mes (ej. Córdoba bajado con junio y julio), se concilia el último mes
  const ult = movs[movs.length - 1].fecha, mes = m => m.fecha.getFullYear() * 12 + m.fecha.getMonth();
  if (mes(movs[0]) !== mes(movs[movs.length - 1])) {
    info.mesesAnteriores = movs.filter(m => mes(m) !== mes(movs[movs.length - 1])).length;
    return { info, movs: movs.filter(m => mes(m) === ult.getFullYear() * 12 + ult.getMonth()) };
  }
  return { info, movs };
}

/**
 * Reporte de la base (hoja "Base"): una fila por renglón de asiento, con DETPLAN = cuenta del mayor
 * (1103012 = O, 1103013 = E) y DETComentaAux con el número y la fecha del valor ("Nº 994833-BANCO MACRO(1103012)-14/07/2026").
 */
// ------------------------------------------------------------------ base FBS de la empresa (otra planilla)

function paramDe_(parametros, nombre) {
  const f = (parametros || []).find(r => new RegExp(nombre, 'i').test(String(r[0])));
  return f ? f[1] : null;
}

/** Cuentas a conciliar según Parámetros. O = null si "Cuenta O" está en NO (solo se concilia la E). */
function cuentasDe_(parametros) {
  const limpio = v => String(v).replace(/\.0+$/, '').trim();
  const pE = paramDe_(parametros, '^cuenta e$'), pO = paramDe_(parametros, '^cuenta o$');
  const sinO = pO !== null && (String(pO).trim() === '' || /^no$/i.test(String(pO).trim()));
  const E = limpio(pE === null ? '' : pE), O = sinO ? null : limpio(pO === null ? '' : pO);
  // sin valores por defecto: una planilla nueva no puede quedar leyendo las cuentas de otro banco
  if (!/^\d+$/.test(E) || (O !== null && !/^\d+$/.test(O))) {
    throw new Error('Completá "Cuenta E" y "Cuenta O" en "' + HOJA.parametros + '" con los números de cuenta de FBS de este banco ' +
      '(ej. Galicia: 1103098 y 1103097). Si no se usa la cuenta O, poné NO.');
  }
  return { E, O };
}

function idDeLink_(v) {
  const t = String(v || '').trim();
  if (!t) return '';
  const m = t.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  return m ? m[1] : (/^[a-zA-Z0-9_-]{20,}$/.test(t) ? t : '');
}

/** Letra de columna (1 -> A, 27 -> AA). */
function letra_(n) { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }

function fechaGviz_(d) { return "date '" + d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) + "'"; }

/** Respuesta JSON de la consulta de Google Sheets (gviz) -> filas con encabezado; las fechas vuelven como Date. */
function parseGviz_(texto) {
  const json = JSON.parse(texto.slice(texto.indexOf('{'), texto.lastIndexOf('}') + 1));
  if (json.status === 'error') throw new Error((json.errors || []).map(e => e.detailed_message || e.message).join('; '));
  const cols = json.table.cols;
  const valor = c => {
    if (!c || c.v === null || c.v === undefined) return '';
    if (typeof c.v === 'string') {
      const d = c.v.match(/^Date\((\d+),(\d+),(\d+)/);
      if (d) return new Date(Number(d[1]), Number(d[2]), Number(d[3]));
    }
    return c.v;
  };
  return [cols.map(c => c.label || c.id)].concat((json.table.rows || []).map(r => cols.map((_, i) => valor(r.c[i]))));
}

function consultaGviz_(id, hoja, consulta) {
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/gviz/tq?tqx=out:json&headers=1&sheet=' +
    encodeURIComponent(hoja) + '&tq=' + encodeURIComponent(consulta);
  const r = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200 || !/setResponse/.test(r.getContentText().slice(0, 200))) {
    throw new Error('consulta a la base: respuesta ' + r.getResponseCode());
  }
  return parseGviz_(r.getContentText());
}

/**
 * Trae de la base de la empresa los movimientos de las cuentas E y O entre "desde" y "hasta" (sin los eliminados) y
 * calcula el saldo inicial de cada cuenta sumando todo lo anterior a "desde". Consulta filtrada (rápida); si falla,
 * lee la hoja completa.
 */
function leerBaseExterna_(id, hoja, ctas, desde, hasta) {
  let sh;
  try { sh = SpreadsheetApp.openById(id).getSheetByName(hoja); } catch (e) {
    throw new Error('No pude abrir la base FBS (' + id + '): revisá el link en "' + HOJA.parametros + '" y que tengas acceso.');
  }
  if (!sh) throw new Error('La base FBS no tiene la hoja "' + hoja + '".');
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normalizar_);
  const col = re => { const i = cab.findIndex(h => re.test(h)); return i < 0 ? null : letra_(i + 1); };
  const L = { id: col(/^detmovid$/), plan: col(/^detplan$/), fecha: col(/^detfecha$/), debe: col(/^detdebe$/), haber: col(/^dethaber$/), estado: col(/^_estado$/),
    asiento: col(/^detmovnro$/) };
  if (!L.plan || !L.fecha || !L.debe || !L.haber) throw new Error('La base FBS no tiene las columnas DETPLAN, DETFECHA, DETDEBE y DETHABER.');
  const cuentas = [ctas.E].concat(ctas.O ? [ctas.O] : []);
  const enCuentas = '(' + cuentas.map(c => L.plan + ' = ' + Number(c)).join(' or ') + ')';
  const vigente = L.estado ? ' and not (' + L.estado + " contains 'ELIMIN')" : '';
  try {
    const filas = consultaGviz_(id, hoja, 'select * where ' + enCuentas + ' and ' + L.fecha + ' >= ' + fechaGviz_(desde) +
      ' and ' + L.fecha + ' <= ' + fechaGviz_(hasta) + vigente);
    const sum = consultaGviz_(id, hoja, 'select ' + L.plan + ', sum(' + L.debe + '), sum(' + L.haber + ') where ' + enCuentas +
      ' and ' + L.fecha + ' < ' + fechaGviz_(desde) + vigente + ' group by ' + L.plan);
    const saldos = { E: 0, O: 0 };
    sum.slice(1).forEach(r => { const k = String(r[0]) === String(Number(ctas.E)) ? 'E' : 'O'; saldos[k] = redondear_(Number(r[1] || 0) - Number(r[2] || 0)); });
    return { filas, saldos, ctx: { id, hoja, L, enCuentas, vigente, ctas } };
  } catch (e) {
    // sin consulta: se lee la hoja entera y se filtra acá (más lento)
    const todo = sh.getDataRange().getValues();
    const h = encabezados_(todo, { plan: /^detplan$/, fecha: /^detfecha$/, debe: /^detdebe$/, haber: /^dethaber$/, estado: /^_estado$/ }, ['plan', 'fecha', 'debe', 'haber']);
    const c = h.col, filas = [todo[h.fila]], saldos = { E: 0, O: 0 };
    todo.slice(h.fila + 1).forEach(r => {
      const plan = String(r[c.plan]).replace(/\.0+$/, '').trim();
      const k = plan === ctas.E ? 'E' : (ctas.O && plan === ctas.O ? 'O' : null);
      const f = parseFecha_(r[c.fecha]);
      if (!k || !f || (c.estado !== undefined && /ELIMIN/i.test(String(r[c.estado])))) return;
      if (f < desde) saldos[k] += (parseNum_(r[c.debe]) || 0) - (parseNum_(r[c.haber]) || 0);
      else if (f <= hasta) filas.push(r);
    });
    return { filas, saldos: { E: redondear_(saldos.E), O: redondear_(saldos.O) } };
  }
}

/**
 * Movimientos de meses ya cerrados que se cargaron, cambiaron o borraron en FBS después del cierre anterior (ej. la
 * confirmación en E de un recibo hecha después de cerrar el mes, con fecha de ese mes). Se suman al mes que se concilia
 * para que crucen contra los pendientes anteriores; el saldo inicial sigue siendo el del cierre.
 *  - Con la foto del cierre anterior (hoja oculta "_Foto FBS", la guarda "Cerrar mes"): diferencia exacta por DETMOVID.
 *  - Sin foto (primer mes): por cada cuenta cuyo saldo en la base no coincide con el del cierre, busca el movimiento
 *    anterior al mes por exactamente esa diferencia.
 * Devuelve filas (con encabezado, columna extra _TARDIO) y el total por cuenta.
 */
function movimientosTardios_(ctx, desde, dif, foto) {
  const L = ctx.L, total = { E: 0, O: 0 };
  const cuentaDe = plan => (String(Number(plan)) === String(Number(ctx.ctas.E)) ? 'E' : 'O');
  let filas = null;
  const sumar = rows => rows.slice(1).forEach(r => {
    if (!filas) filas = [rows[0].concat(['_TARDIO'])];
    filas.push(r.concat(['SI']));
  });
  if (foto && L.id) {
    const actual = consultaGviz_(ctx.id, ctx.hoja, 'select ' + [L.id, L.plan, L.debe, L.haber, L.fecha].join(', ') + ' where ' + ctx.enCuentas +
      ' and ' + L.fecha + ' < ' + fechaGviz_(desde) + ctx.vigente);
    const vistos = new Set(), nuevos = [];
    actual.slice(1).forEach(r => {
      const id = String(r[0]), imp = redondear_(Number(r[2] || 0) - Number(r[3] || 0)), f = foto.get(id);
      vistos.add(id);
      if (!f || f.cuenta !== String(Number(r[1])) || Math.abs(f.importe - imp) > 0.005) {
        nuevos.push(id);
        if (f) foto.set(id, Object.assign({}, f, { cambiado: true }));
      }
    });
    for (let i = 0; i < nuevos.length; i += 40) {
      sumar(consultaGviz_(ctx.id, ctx.hoja, 'select * where ' + nuevos.slice(i, i + 40).map(x => L.id + ' = ' + x).join(' or ')));
    }
    // lo que estaba en la foto y ya no está (o cambió): se revierte la versión vieja
    const quitados = [];
    foto.forEach((f, id) => { if ((!vistos.has(id) || f.cambiado) && Math.abs(f.importe) > 0.005) quitados.push([id, f]); });
    if (quitados.length) {
      if (!filas) filas = [['DETMOVID', 'DETMOVNRO', 'DETFECHA', 'DETPLAN', 'DETDEBE', 'DETHABER', 'DETComenta', '_TARDIO']];
      const h = filas[0].map(normalizar_), fila = () => filas[0].map(() => '');
      quitados.forEach(([id, f]) => {
        const r = fila();
        r[h.indexOf('detmovid')] = id; r[h.indexOf('detfecha')] = f.fecha; r[h.indexOf('detplan')] = Number(f.cuenta);
        r[h.indexOf('detdebe')] = Math.max(-f.importe, 0); r[h.indexOf('dethaber')] = Math.max(f.importe, 0);
        r[h.indexOf('detcomenta')] = 'Movimiento ' + id + ' del ' + fechaTexto_(f.fecha) + ' borrado o cambiado en FBS después del cierre (se revierte)';
        if (h.indexOf('detmovnro') >= 0) r[h.indexOf('detmovnro')] = 'foto ' + id;
        r[r.length - 1] = 'SI';
        filas.push(r);
      });
    }
  } else {
    ['E', 'O'].forEach(k => {
      const d = redondear_(dif[k] || 0), cuenta = ctx.ctas[k];
      if (!cuenta || Math.abs(d) <= 1) return;
      const cond = d > 0 ? L.debe + ' = ' + d : L.haber + ' = ' + (-d);
      const r = consultaGviz_(ctx.id, ctx.hoja, 'select * where ' + L.plan + ' = ' + Number(cuenta) + ' and ' + L.fecha + ' < ' +
        fechaGviz_(desde) + ' and ' + cond + ctx.vigente + (L.id ? ' order by ' + L.id + ' desc' : '') + ' limit 1');
      if (r.length > 1) { sumar(r); return; }
      const mesAnt = new Date(desde.getFullYear(), desde.getMonth() - 1, 1);
      const delMes = ' where ' + L.plan + ' = ' + Number(cuenta) + ' and ' + L.fecha + ' >= ' + fechaGviz_(mesAnt) + ' and ' + L.fecha +
        ' < ' + fechaGviz_(desde) + ctx.vigente;
      // un asiento entero del mes cerrado (ej. el asiento de sueldos del 30/06 con una línea por empleado, cargado en O
      // después del cierre), o dos asientos, que suman la diferencia
      if (L.asiento) {
        const g = consultaGviz_(ctx.id, ctx.hoja, 'select ' + L.asiento + ', sum(' + L.debe + ') - sum(' + L.haber + ')' + delMes +
          ' group by ' + L.asiento);
        const asientos = g.slice(1).map(x => ({ nro: x[0], importe: redondear_(Number(x[1] || 0)) })).filter(x => x.nro !== '' && x.nro !== null);
        let combo = null;
        for (let n = 1; n <= 2 && !combo; n++) combo = primeraCombinacion_(asientos.slice().reverse().slice(0, 300), n, d);
        if (combo) {
          const val = x => (typeof x === 'number' || /^\d+$/.test(String(x)) ? Number(x) : "'" + x + "'");
          sumar(consultaGviz_(ctx.id, ctx.hoja, 'select *' + delMes + ' and (' + combo.map(x => L.asiento + ' = ' + val(x.nro)).join(' or ') + ')'));
          return;
        }
      }
      // varios movimientos sueltos: la combinación de hasta 3 movimientos del mes cerrado, del mismo sentido
      const cands = consultaGviz_(ctx.id, ctx.hoja, 'select * where ' + L.plan + ' = ' + Number(cuenta) + ' and ' + L.fecha + ' >= ' +
        fechaGviz_(mesAnt) + ' and ' + L.fecha + ' < ' + fechaGviz_(desde) + ' and ' + (d > 0 ? L.debe : L.haber) + ' > 0' + ctx.vigente +
        (L.id ? ' order by ' + L.id + ' desc' : '') + ' limit 150');
      if (cands.length < 3) return;
      const h = cands[0].map(normalizar_), ide = h.indexOf('detdebe'), iha = h.indexOf('dethaber');
      const items = cands.slice(1).map(row => ({ row, importe: redondear_(Number(row[ide] || 0) - Number(row[iha] || 0)) }))
        .filter(x => Math.abs(x.importe) <= Math.abs(d) + 0.01);
      for (let n = 2; n <= 3; n++) {
        const combo = primeraCombinacion_(items.slice(0, 60), n, d);
        if (combo) { sumar([cands[0]].concat(combo.map(x => x.row))); return; }
      }
    });
  }
  if (filas) {
    const h = filas[0].map(normalizar_), ip = h.indexOf('detplan'), id = h.indexOf('detdebe'), ih = h.indexOf('dethaber');
    filas.slice(1).forEach(r => { total[cuentaDe(r[ip])] += Number(r[id] || 0) - Number(r[ih] || 0); });
  }
  return { filas: filas || [], total: { E: redondear_(total.E), O: redondear_(total.O) } };
}

/** Une las filas del mes con las tardías (mismas columnas por nombre; se agrega _TARDIO). */
function unirFilas_(filas, extra) {
  if (!extra.length) return filas;
  const cab = filas[0].concat(['_TARDIO']), h = cab.map(normalizar_), he = extra[0].map(normalizar_);
  return [cab].concat(filas.slice(1).map(r => r.concat([''])),
    extra.slice(1).map(r => h.map(c => { const i = he.indexOf(c); return i < 0 ? '' : r[i]; })));
}

const HOJA_FOTO_FBS = '_Foto FBS';

/** Foto de las cuentas E y O al cierre: DETMOVID, cuenta, importe (debe - haber) y fecha de todo hasta fin de mes. */
function guardarFotoBase_(ss, ctx, hasta) {
  const L = ctx.L;
  if (!L.id) return 0;
  const r = consultaGviz_(ctx.id, ctx.hoja, 'select ' + [L.id, L.plan, L.debe, L.haber, L.fecha].join(', ') + ' where ' + ctx.enCuentas +
    ' and ' + L.fecha + ' <= ' + fechaGviz_(hasta) + ctx.vigente);
  const filas = [['DETMOVID', 'Cuenta', 'Importe', 'Fecha', 'Cierre al ' + fechaTexto_(hasta)]].concat(r.slice(1).map(x =>
    [String(x[0]), String(Number(x[1])), redondear_(Number(x[2] || 0) - Number(x[3] || 0)), x[4], '']));
  let sh = ss.getSheetByName(HOJA_FOTO_FBS);
  if (sh) sh.clearContents(); else sh = ss.insertSheet(HOJA_FOTO_FBS);
  sh.getRange(1, 1, filas.length, 5).setValues(filas);
  sh.getRange(2, 1, Math.max(filas.length - 1, 1), 2).setNumberFormat('@');
  sh.hideSheet();
  return filas.length - 1;
}

function leerFotoBase_(ss) {
  const sh = ss.getSheetByName(HOJA_FOTO_FBS);
  if (!sh || sh.getLastRow() < 2) return null;
  const m = new Map();
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => m.set(String(r[0]), { cuenta: String(r[1]), importe: Number(r[2]), fecha: r[3] }));
  return m;
}

function leerBase_(filas, cuentaE, cuentaO) {
  const h = encabezados_(filas, {
    plan: /^detplan$/, asiento: /^detmovnro$/, fecha: /^detfecha$/, ref: /^detref$/, debe: /^detdebe$/, haber: /^dethaber$/,
    com: /^detcomenta$/, aux: /^detcomentaaux$/, id: /^detmovid$/, cc: /centrocosto/, estado: /^_estado$/, tardio: /^_tardio$/,
  }, ['plan', 'fecha', 'debe', 'haber', 'com']);
  if (!h) throw new Error('Base: no encuentro las columnas DETPLAN, DETFECHA, DETDEBE, DETHABER, DETComenta.');
  const c = h.col, out = { E: [], O: [] };
  filas.slice(h.fila + 1).forEach(r => {
    const plan = String(r[c.plan]).replace(/\.0+$/, '').trim();
    const cta = plan === String(cuentaE) ? 'E' : plan === String(cuentaO) ? 'O' : null;
    const fecha = parseFecha_(r[c.fecha]);
    if (!cta || !fecha) return;
    if (c.estado !== undefined && /ELIMIN/i.test(String(r[c.estado]))) return;   // borrado en el ERP
    const aux = c.aux !== undefined ? r[c.aux] : '';
    out[cta].push({ asiento: String(r[c.asiento] || '').replace(/\.0+$/, ''), fecha, referencia: String(r[c.ref] || '').trim(),
      comentario: String(r[c.com] || '').replace(/\s+/g, ' ').trim(), debe: parseNum_(r[c.debe]) || 0, haber: parseNum_(r[c.haber]) || 0,
      aux: aux instanceof Date ? aux : (typeof aux === 'number' ? String(Math.round(aux)) : String(aux || '').trim()),
      id: c.id !== undefined ? String(r[c.id] || '').replace(/\.0+$/, '') : '', cuentaBancoPropia: String(cuentaO),
      cc: c.cc !== undefined ? String(r[c.cc] || '').trim() : '', tardio: c.tardio !== undefined && r[c.tardio] === 'SI' });
  });
  return out;
}

/**
 * Datos del valor en DETComentaAux. Formatos vistos:
 *   "Nº 20/7/1153-BANCO MACRO(1103012)-21/07/2026"  -> fecha del valor 20/07 (la transferencia; 11:53 es la hora)
 *   "Nº 994833-BANCO MACRO(1103012)-14/07/2026"     -> número = Nro. de referencia del banco
 *   "Nº 1200058-BANCO COMAFI(1103074)-17/07/2026"   -> valor de OTRO banco
 *   "I-14-7" / "E-0172026" / "13072026" / "64274939" -> ingreso / egreso con fecha, fecha, nro. de cheque
 * La fecha del final es la fecha de carga del recibo (se usa como segunda opción).
 */
/** Centros de costo que cargan en el Nº del valor la fecha del banco sin separador: "2206" = 22/06, "607" = 6/07, "10626" = 1/06/26. */
const CC_NUMERO_ES_FECHA = /DOME TALLER/i;

function fechaCorta_(nro, base) {
  // "96" = 9/6, "2206" = 22/06, "19066" = 19/06 (dígito de más), "5062026" = 5/06/2026, "22062026" = 22/06/2026
  const opciones = { 2: [[1, 1, 0]], 3: [[1, 2, 0], [2, 1, 0]], 4: [[2, 2, 0], [1, 2, 1]], 5: [[1, 2, 2], [2, 1, 2], [2, 2, -1]],
    7: [[1, 2, 4]], 8: [[2, 2, 4]] }[nro.length] || [];
  for (const [ld, lm, la] of opciones) {
    if (la >= 0 && ld + lm + la !== nro.length) continue;
    const d = Number(nro.slice(0, ld)), mes = Number(nro.slice(ld, ld + lm));
    if (d < 1 || d > 31 || mes < 1 || mes > 12) continue;
    let a = la === 2 ? 2000 + Number(nro.slice(ld + lm)) : la === 4 ? Number(nro.slice(ld + lm)) : base.getFullYear();
    let f = new Date(a, mes - 1, d);
    if (la <= 0 && f > base) f = new Date(a - 1, mes - 1, d);
    if (f <= base && dias_(f, base) <= 180) return f;
  }
  return null;
}

function datosAux_(aux, base, cc) {
  const vacio = { numero: '', fecha: null, fecha2: null, banco: '', cuentaBanco: '' };
  if (!aux) return vacio;
  if (aux instanceof Date) return Object.assign(vacio, { fecha: parseFecha_(aux) });
  const t = String(aux).trim();
  const out = Object.assign({}, vacio);
  if (parseFecha_(t)) return Object.assign(out, { fecha: parseFecha_(t) });   // la celda trae directamente una fecha
  const m = t.match(/^N[º°o]?\s*(.*?)-\s*([A-Z][A-Z .]+?)\s*-?\s*\((\d+)\)\s*-\s*(\d{1,2}\/\d{1,2}\/\d{2,4})?\s*$/i);   // también \"BANCO MACRO-(1103015)\"
  const dm = (txt) => {
    let x = txt.match(/^(\d{1,2})\s*[\/-]\s*(\d{1,2})(?!\d)/);
    if (x) {
      const d = Number(x[1]), mes = Number(x[2]);
      if (d >= 1 && d <= 31 && mes >= 1 && mes <= 12) {
        let a = base.getFullYear(); if (mes > base.getMonth() + 1) a -= 1;
        return new Date(a, mes - 1, d);
      }
    }
    return /^\d{6,8}$/.test(txt) ? fechaEnTexto_(txt, base) : null;
  };
  if (m) {
    const nro = m[1].replace(/\s+/g, '');
    out.banco = m[2].trim(); out.cuentaBanco = m[3];
    out.fecha2 = m[4] ? parseFecha_(m[4]) : null;
    out.fecha = dm(nro);
    if (/^\d{5,}$/.test(nro)) out.numero = nro;
    if (!out.fecha && cc && CC_NUMERO_ES_FECHA.test(cc) && /^\d{2,5}$/.test(nro)) {
      out.fecha = fechaCorta_(nro, out.fecha2 || base);
      if (out.fecha) out.numero = '';
    }
    if (!out.fecha) out.fecha = out.fecha2;
    return out;
  }
  const ie = t.match(/^[IE]-\s*(.+)$/i);
  if (ie) { out.fecha = dm(ie[1].replace(/\s/g, '')); return out; }
  if (/^\d{5,}$/.test(t)) { out.numero = t; out.fecha = /^\d{7,8}$/.test(t) ? fechaEnTexto_(t, base) : null; }
  return out;
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
  const ax = datosAux_(a.aux, a.fecha || new Date(), a.cc);
  const fr = ax.fecha || (a.fecha ? fechaEnTexto_(com, a.fecha) : null);
  const otroBanco = ax.cuentaBanco && a.cuentaBancoPropia && ax.cuentaBanco !== String(a.cuentaBancoPropia) ? ax.banco : '';
  return { lado: 'FBS', origen: origen || 'Mes', fecha: a.fecha, importe: redondear_(a.debe - a.haber), texto: com,
    fechaRef: fr || null, fechaRef2: ax.fecha2 && (!fr || dias_(ax.fecha2, fr) > 0) ? ax.fecha2 : null, cruces: '', valor: ax.numero,
    otroBanco, aux: a.aux instanceof Date ? fechaTexto_(a.aux) : (a.aux || ''),
    ref: a.referencia || '', asiento: a.asiento || '', cuenta: a.cuenta || '', cc: a.cc || '', tardio: !!a.tardio,
    alerta: otroBanco ? 'El valor figura en otro banco (' + otroBanco + ' ' + ax.cuentaBanco + '): no corresponde a esta cuenta' : '', cuit, dni,
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
  // pendiente anterior cargado sin centavos (ej. 456.048.014 en la lista y 456.048.014,26 en FBS): mismo comprobante,
  // hasta $1 de diferencia
  resto.filter(a => a.origen === 'Arrastre').forEach(a => {
    if (a.match !== null) return;
    const b = resto.find(x => x.match === null && x.origen === 'Mes' && (x.importe > 0) !== (a.importe > 0) &&
      Math.abs(x.importe + a.importe) <= 1 && x.texto.replace(/\s+/g, ' ').trim() === a.texto.replace(/\s+/g, ' ').trim());
    if (b) [a, b].forEach(x => { x.match = []; x.metodo = 'Confirmado en O (pendiente anterior con importe redondeado)'; });
  });
  return resto.filter(x => x.match === null);
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
 * Pase de O a E (haber O / debe E en el mismo asiento) cuyo ingreso en O no está en el mes ni en los pendientes
 * anteriores. Dos casos:
 *  - el recibo se cargó después del cierre (fecha de carga del DETComentaAux posterior al corte) y se confirmó con
 *    fecha del mes: el debe de E se cruza con el banco y el haber de O queda en S4 ("desfasado") hasta que el mes
 *    próximo aparezca el ingreso en O con el mismo comprobante;
 *  - si no, el recibo ya se había dado por acreditado antes: el pase no cambia el saldo E+O, así que las dos líneas
 *    se anulan entre sí y no se cruzan con el banco.
 * Devuelve las líneas de E que quedan para cruzar.
 */
function pasesSinOrigen_(ao, asientosE) {
  const usadas = new Set();
  ao.pendientes.forEach(p => {
    if (p.origen !== 'Mes' || p.importe >= 0) return;
    if (p.fechaRef2 && CORTE_ && p.fechaRef2 > CORTE_) {
      // confirmado con fecha del mes pero el ingreso en O se registró después del cierre: el pase queda en S4
      // hasta que el mes próximo entre el ingreso (con el mismo comprobante) y se netee en O
      p.desfasado = true;
      return;
    }
    const i = asientosE.findIndex((a, j) => !usadas.has(j) && String(a.asiento) === String(p.asiento) &&
      a.comentario === p.texto && Math.abs(a.debe - a.haber + p.importe) < 0.005);
    if (i < 0) return;
    usadas.add(i);
    const eLinea = partidaFbs_(asientosE[i], 'Mes');
    eLinea.cuenta = 'E';
    compensar_([p, eLinea], 'Pase O→E sin ingreso en O en el mes (se anulan)');
    const r = ao.resumen.find(x => x.clave === claveO_(p.texto));
    if (r) { r.pendientes--; r.detallePend = (r.pendientes ? r.detallePend + ' — ' : '') + 'pase a E sin ingreso en O en el mes: se anula con E'; }
  });
  ao.pendientes = ao.pendientes.filter(p => p.match === null);
  ao.pasesSinOrigen = usadas.size;
  return asientosE.filter((a, j) => !usadas.has(j));
}

/**
 * Registros de E que son el pase de un movimiento de O cargado antes (mismo comentario, mismo importe y sentido): el
 * banco lo movió en la fecha de O (ej. pago CIG FIAT del 13/07 en O, pasado a E el 23/07). Esa fecha también vale para
 * los cruces.
 */
function fechaEnO_(eItems, asientosO) {
  const antes = new Map();
  asientosO.forEach(a => {
    const k = String(a.comentario).trim() + '|' + redondear_(a.debe - a.haber).toFixed(2);
    if (!antes.has(k) || a.fecha < antes.get(k)) antes.set(k, a.fecha);
  });
  eItems.forEach(p => {
    const f = antes.get(String(p.texto).trim() + '|' + p.importe.toFixed(2));
    if (f && f < p.fecha) p.fechaO = f;
  });
}

/**
 * Paso 2: la cuenta E se cruza con el banco. Las líneas de una misma liquidación de tarjeta (la "Confirmación de
 * Valores Agrupados" del asiento que confirma la liquidación en O, y el asiento con el número de liquidación) se
 * juntan en un solo neto, que es lo que acredita el banco.
 */
function partidasE_(asientosE, asientosO) {
  const liqPorAsiento = new Map(), numeros = new Set(), fechaLiq = new Map();
  asientosO.forEach(a => {
    const k = claveO_(a.comentario);
    if (k.startsWith('LIQ ')) {
      liqPorAsiento.set(String(a.asiento), k); numeros.add(k.slice(4));
      // fecha de la liquidación escrita en el comentario ("12658-liq.- MC. 07/7"): es el día en que acredita el banco
      if (!fechaLiq.has(k)) { const f = fechaEnTexto_(String(a.comentario).replace(/^\d+/, ''), a.fecha); if (f) fechaLiq.set(k, f); }
    }
  });
  const liqDe = a => {
    if (liqPorAsiento.has(String(a.asiento))) return liqPorAsiento.get(String(a.asiento));
    const k = claveO_(a.comentario);
    if (k.startsWith('LIQ ')) return k;
    const d = String(a.comentario || '').trim().match(/^(\d{4,})\b/);
    if (d) {
      const num = Array.from(numeros).filter(n => d[1].startsWith(n)).sort((x, y) => y.length - x.length)[0];
      if (num) return 'LIQ ' + num;
      // comisiones de la liquidación: se cargan en el asiento inmediato siguiente con solo un número en el comentario,
      // que a veces no coincide con el de la liquidación (ej. liq 45956 en el asiento 2410279, comisión "75956" en el 2410280)
      const prev = String(Number(a.asiento) - 1);
      if (/^\d+(\.0+)?$/.test(String(a.comentario).trim()) && liqPorAsiento.has(prev)) return liqPorAsiento.get(prev);
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
    // el concepto es el de las líneas originales (no se inventa un texto); el detalle queda en p.lineas
    const textos = Array.from(new Set(g.map(a => String(a.comentario).trim())));
    const p = partidaFbs_({ fecha: ult.fecha, comentario: textos.join(' + '),
      referencia: '', asiento: Array.from(new Set(g.map(a => a.asiento))).join(' / '),
      debe: Math.max(neto, 0), haber: Math.max(-neto, 0) }, 'Mes');
    p.cuenta = 'E';
    p.liq = k;
    if (fechaLiq.has(k)) p.fechaLiq = fechaLiq.get(k);
    if (g.length > 1) p.lineas = g.map(a => ({ fecha: a.fecha, texto: String(a.comentario).trim(), importe: redondear_(a.debe - a.haber), asiento: a.asiento }));
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
    // el importe va en positivo; uno negativo invierte el sentido (ej. la comisión de una liquidación listada en S3)
    const v = sec === 'S1' || sec === 'S3' ? imp : -imp;
    if (sec === 'S1' || sec === 'S2') {
      out.push(partidaBanco_({ fecha, concepto: texto, importe: v, ref: '', nombre: '' }, 'Arrastre'));
    } else {
      out.push(partidaFbs_({ fecha, comentario: texto, referencia: '', debe: Math.max(v, 0), haber: Math.max(-v, 0) }, 'Arrastre'));
    }
    out[out.length - 1].sectorOriginal = sec;
  });
  return out;
}

// ------------------------------------------------------------------ cruce

let CRUCES_ = [];
let CORTE_ = null;   // fecha de corte del extracto de la corrida   // registro de todos los cruces de la corrida (para la hoja "E conciliado")

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
  if (a.fechaRef2) d = Math.min(d, dias_(a.fechaRef2, b.fecha));
  if (b.fechaRef2) d = Math.min(d, dias_(a.fecha, b.fechaRef2));
  if (a.fechaO) d = Math.min(d, dias_(a.fechaO, b.fecha));
  if (b.fechaO) d = Math.min(d, dias_(a.fecha, b.fechaO));
  return d;
}

/** Última fecha dd/mm o dd/mm/aaaa escrita en un comentario (sin tomar "30052/9" ni "7/2026"). */
function fechaEnTexto_(texto, base) {
  // "3/7/-/2026" también vale (barra suelta después del mes)
  const re = /(?:^|[^\d\/])(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?(?!\d|\/\d)/g;
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
    // "Nº 17-7" / "Nº 13-07-26": día y mes separados por guión después de Nº
    const g = String(texto).match(/Nº\s*(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?(?!\d)/);
    if (g && Number(g[1]) <= 31 && Number(g[2]) >= 1 && Number(g[2]) <= 12) {
      let a = g[3] ? Number(g[3]) : base.getFullYear(); if (a < 100) a += 2000;
      if (!g[3] && Number(g[2]) > base.getMonth() + 1) a -= 1;
      ult = new Date(a, Number(g[2]) - 1, Number(g[1]));
    }
  }
  if (!ult) {
    // "07-7-26-dep e cheq": día-mes-año con guiones al principio del comentario
    const g = String(texto).match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})(?!\d)/);
    if (g) {
      let a = Number(g[3]); if (a < 100) a += 2000;
      const f = new Date(a, Number(g[2]) - 1, Number(g[1]));
      if (Number(g[1]) <= 31 && Number(g[2]) >= 1 && Number(g[2]) <= 12 && dias_(f, base) <= 62) ult = f;
    }
  }
  if (!ult) {
    // "7130269950-DP CHEQUES" (Galicia): mes + día + número de boleta de 7 dígitos
    const g = String(texto).match(/^([1-9])(\d{2})\d{7}-/);
    if (g) {
      const f = new Date(base.getFullYear(), Number(g[1]) - 1, Number(g[2]));
      if (Number(g[2]) >= 1 && Number(g[2]) <= 31 && dias_(f, base) <= 62) ult = f;
    }
  }
  if (!ult) {
    // "Nº 4080226", "Nº 21082026", "Nº 210820261", "25082026-Vta Cheques", "14726-dep. echeq": día + mes + año pegados
    const c = String(texto).match(/(?:Nº\s*|^)(\d{5,9})(?!\d)/);
    if (c) {
      const t = c[1];
      busca: for (const [ld, lm] of [[2, 2], [1, 2], [2, 1], [1, 1]]) {
        const d = Number(t.slice(0, ld)), mes = Number(t.slice(ld, ld + lm)), resto = t.slice(ld + lm);
        if (!resto) continue;
        let a = resto.startsWith('20') && resto.length >= 4 ? Number(resto.slice(0, 4)) : Number(resto.slice(-2));
        if (a < 100) a += 2000;
        const f = new Date(a, mes - 1, d);
        if (d >= 1 && d <= 31 && mes >= 1 && mes <= 12 && a === base.getFullYear() && dias_(f, base) <= 62) { ult = f; break busca; }
      }
    }
  }
  return ult;
}

function primeraCombinacion_(cands, n, objetivo, tolerancia) {
  const idx = [];
  const buscar = (inicio, suma) => {
    if (idx.length === n) return Math.abs(suma - objetivo) <= (tolerancia || TOLERANCIA);
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
    if (!cands.length) return;
    if (unico && cands.length > 1) {
      // candidatos intercambiables: mismo importe y la misma cantidad de movimientos iguales en el banco ese día. Valen los
      // valores de Dome Taller con la misma fecha del valor (dos de $145.000 del 22/06 y dos transferencias de $145.000 el
      // 22/06) y registros con el mismo detalle, del día del banco (tres cheques de $400.000 de "02/07-vta ch" y tres
      // créditos del 02/07). Recibos de personas distintas no.
      const dia = f => (f.fechaRef || f.fecha).getTime();
      const detalle = f => (f.texto.split(' / ').slice(1).join(' / ') || f.texto).replace(/\s+/g, ' ').trim().toUpperCase();
      const iguales = cands.every(f => Math.abs(f.importe - cands[0].importe) < 0.005 &&
        (f.cc && CC_NUMERO_ES_FECHA.test(f.cc) ? dia(f) === dia(cands[0]) : distancia_(b, f) === 0 && detalle(f) === detalle(cands[0])));
      const gemelos = banco.filter(x => x.match === null && Math.abs(x.importe - b.importe) < 0.005 && x.fecha.getTime() === b.fecha.getTime());
      if (!iguales || gemelos.length !== cands.length) return;
    }
    let mejor = cands[0];
    cands.forEach(f => {
      const d1 = distancia_(b, f), d0 = distancia_(b, mejor);
      if (d1 < d0 || (d1 === d0 && Math.abs(f.importe - b.importe) < Math.abs(mejor.importe - b.importe))) mejor = f;
    });
    unir_([b], [mejor], metodo);
  });
}

/** "ECHEQ 48 HS. NRO. 558" en el banco contra "RM-388142 / GERARDO JUAN GIECO Nº 558." en FBS. */
function MISMO_NUMERO_(b, f) {
  const nb = b.texto.match(/\bNRO\.?:?\s*(\d{1,8})\b/i), nf = f.texto.match(/Nº\s*(\d{1,8})\.?\s*$/);
  return !!nb && !!nf && Number(nb[1]) === Number(nf[1]);
}

/**
 * Varias líneas de un lado que suman un movimiento (o un grupo del día con el mismo código) del otro, todas con la misma
 * fecha: la del banco, la escrita en el comentario de FBS o la del movimiento original en O. Ej.: un crédito por
 * descuento de documentos = varias líneas "02/07-vta ch"; una línea "Nº 3/7/-/2026" = tres transferencias del 03/07;
 * los tres pagos CIG FIAT del 13/07 = cuatro pases de O del 13/07. Importes redondos: solo un movimiento (no un grupo),
 * contra líneas con la fecha escrita y hasta 3 líneas.
 */
function combinacionesFechaEscrita_(banco, fbs) {
  const fechasFbs = f => [f.fechaRef, f.fechaO].filter(Boolean).map(x => x.getTime());
  const unidades = [];
  const grupos = new Map();
  banco.forEach(b => {
    if (b.match !== null || b.origen !== 'Mes') return;
    unidades.push([b]);
    const k = b.fecha.getTime() + '|' + b.causal + '|' + (b.importe > 0);
    grupos.set(k, (grupos.get(k) || []).concat([b]));
  });
  grupos.forEach(g => { if (g.length > 1) unidades.push(g); });
  const total = l => redondear_(l.reduce((x, p) => x + p.importe, 0));
  // banco (uno o un grupo del día) = varias líneas de FBS
  unidades.sort((x, y) => Math.abs(total(y)) - Math.abs(total(x))).forEach(u => {
    if (u.some(b => b.match !== null)) return;
    const t = total(u), d = u[0].fecha.getTime(), esp = especifico_(t);
    if (u.length > 1 && !esp) return;   // grupos del día solo con total específico
    const cands = fbs.filter(f => f.match === null && f.origen === 'Mes' && (f.importe > 0) === (t > 0) &&
      Math.abs(f.importe) <= Math.abs(t) + TOLERANCIA && (esp ? fechasFbs(f) : [f.fechaRef && f.fechaRef.getTime()]).indexOf(d) >= 0).slice(0, 25);
    for (let n = u.length > 1 ? 1 : 2; n <= (esp ? 5 : 3) && cands.length >= n; n++) {
      const combo = primeraCombinacion_(cands, n, t);
      if (combo) { unir_(u, combo, 'Suma de registros con la misma fecha (escrita o de O) (revisar)'); break; }
    }
  });
  // una línea de FBS con fecha escrita = varios movimientos del banco de ese día
  fbs.slice().sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe)).forEach(f => {
    if (f.match !== null || f.origen !== 'Mes') return;
    const ds = fechasFbs(f), esp = especifico_(f.importe);
    if (!ds.length || (!esp && !f.fechaRef)) return;
    const cands = banco.filter(b => b.match === null && b.origen === 'Mes' && (b.importe > 0) === (f.importe > 0) &&
      Math.abs(b.importe) < Math.abs(f.importe) && ds.indexOf(b.fecha.getTime()) >= 0).slice(0, 25);
    // importe redondo (ej. RM por 2.420.000 = comisión 2.000.000 + IVA 420.000 del día): solo de a dos
    for (let n = 2; n <= (esp ? 4 : 2) && cands.length >= n; n++) {
      const combo = primeraCombinacion_(cands, n, f.importe);
      if (combo) { unir_(combo, [f], 'Suma de movimientos del día escrito en el comprobante (revisar)'); break; }
    }
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
    if (otro && !f.alerta) f.alerta = 'Posible duplicado de ' + (otro.texto.split('/')[0].trim() || 'asiento ' + otro.asiento) +
      ' (asiento ' + otro.asiento + ')';
  });
}

/** Liquidaciones de tarjeta (neto de la cuenta E) contra acreditaciones de tarjetas del banco: se admite hasta
 *  $1 de diferencia por redondeos; la diferencia queda como partida a ajustar para no esconderla en el control. */
function liquidacionesTarjeta_(banco, fbs) {
  const ajustes = [];
  fbs.filter(f => f.match === null && f.liq && f.liq.startsWith('LIQ ')).forEach(f => {
    // el banco acredita la liquidación el día escrito en su comentario: si está, tiene que ser ese día sí o sí
    // movimiento de tarjeta: por categoría o por concepto (ej. Naranja cobrada por Datanet)
    const esTarjeta = b => /TARJETA/.test(b.cat) || /TARJETA|PAYWAY|PRISMA|CABAL|NARANJA/i.test(b.texto);
    const diaOk = b => (f.fechaLiq ? dias_(b.fecha, f.fechaLiq) === 0 : dias_(b.fecha, f.fecha) <= 10);
    let mejor = null;
    banco.forEach(b => {
      if (b.match !== null || b.origen !== 'Mes' || !esTarjeta(b) || (b.importe > 0) !== (f.importe > 0) ||
        Math.abs(b.importe - f.importe) > 1 || !diaOk(b)) return;
      if (!mejor || dias_(b.fecha, f.fecha) < dias_(mejor.fecha, f.fecha)) mejor = b;
    });
    if (!mejor) {
      // la liquidación se acredita en varios movimientos del mismo día y con el mismo concepto, que pueden incluir
      // débitos de la tarjeta (ej. liq 9341 = 4 PAYWAY del 27/07; liq 12658 = 3 créditos y 1 débito PAYWAY del 07/07)
      const porDia = new Map();
      banco.forEach(b => {
        if (b.match !== null || b.origen !== 'Mes' || !esTarjeta(b) || !diaOk(b)) return;
        const k = b.fecha.getTime() + '|' + normalizar_(b.texto);
        porDia.set(k, (porDia.get(k) || []).concat([b]));
      });
      const dias = Array.from(porDia.keys()).sort((x, y) => dias_(porDia.get(x)[0].fecha, f.fecha) - dias_(porDia.get(y)[0].fecha, f.fecha));
      let combo = null;
      for (const k of dias) {
        const cands = porDia.get(k).slice(0, 15);
        for (let n = 2; n <= Math.min(6, cands.length) && !combo; n++) combo = primeraCombinacion_(cands, n, f.importe, 1);
        if (combo) break;
      }
      if (!combo) return;
      mejor = { importe: redondear_(combo.reduce((x, b) => x + b.importe, 0)), grupo: combo };
    }
    const dif = redondear_(mejor.importe - f.importe);
    const nombre = 'Liquidación de tarjeta' + (mejor.grupo ? ' (' + mejor.grupo.length + ' movimientos de tarjeta del día)' : '');
    unir_(mejor.grupo || [mejor], [f], nombre + (Math.abs(dif) > 0.005 ? ' (dif. ' + formato_(dif) + ')' : ''));
    if (Math.abs(dif) > 0.005) {
      const a = partidaFbs_({ fecha: f.fecha, comentario: 'Diferencia de redondeo liquidación tarjeta ' + f.liq.slice(4) + ' - ajustar',
        referencia: '', asiento: f.asiento, debe: Math.max(-dif, 0), haber: Math.max(dif, 0) }, 'Mes');
      a.cuenta = 'E';
      ajustes.push(a);
    }
  });
  return ajustes;
}

/**
 * Liquidación de tarjeta con el mismo número en el banco y en FBS (Santander: "Acreditacion a comercio fiserv ...
 * nro.liq. 00340801/..." y "340801-LIQ.- VISA 07/8"), aunque la fecha escrita en FBS no sea la del banco. Hasta $1 de
 * diferencia: queda una línea de ajuste.
 */
function liquidacionPorNumero_(banco, fbs) {
  const ajustes = [];
  const nroBanco = b => { const m = b.texto.match(/nro\.?\s*liq\.?\s*0*(\d{4,})/i); return m ? m[1] : null; };
  // "340801-LIQ.- VISA 07/8"; también "338370-arreglo liq. mal. 338370 02/06" (cualquier número largo si dice liq)
  const nrosFbs = f => (/liq/i.test(f.texto) ? f.texto.match(/\d{5,}/g) || [] : []).map(x => x.replace(/^0+/, ''));
  fbs.forEach(f => {
    if (f.match !== null || f.importe <= 0) return;
    const nros = nrosFbs(f);
    if (!nros.length) return;
    const cands = banco.filter(b => b.match === null && b.importe > 0 && nros.indexOf(nroBanco(b)) >= 0 &&
      Math.abs(b.importe - f.importe) <= 1 && dias_(b.fecha, f.fecha) <= 60);
    if (cands.length !== 1) return;
    const dif = redondear_(cands[0].importe - f.importe);
    unir_(cands, [f], 'Liquidación de tarjeta: mismo número de liquidación' + (Math.abs(dif) > 0.005 ? ' (dif. ' + formato_(dif) + ')' : ''));
    if (Math.abs(dif) > 0.005) {
      const a = partidaFbs_({ fecha: f.fecha, comentario: 'Diferencia de redondeo liquidación tarjeta ' + nros[0] + ' - ajustar',
        referencia: '', asiento: f.asiento, debe: Math.max(dif, 0), haber: Math.max(-dif, 0) }, 'Mes');
      a.cuenta = 'E';
      ajustes.push(a);
    }
  });
  // el crédito del banco = líneas de FBS con ese número de liquidación + una Confirmación de Valores Agrupados sin número
  // (ej. 293.855,49 = 343.000 confirmados en junio - 49.144,51 "338370-arreglo liq. mal")
  banco.forEach(b => {
    const n = b.match === null && b.importe > 0 ? nroBanco(b) : null;
    if (!n) return;
    const L = fbs.filter(f => f.match === null && nrosFbs(f).indexOf(n) >= 0);
    if (!L.length) return;
    const resto = redondear_(b.importe - L.reduce((x, f) => x + f.importe, 0));
    if (Math.abs(resto) <= TOLERANCIA) { unir_([b], L, 'Liquidación de tarjeta: mismo número de liquidación'); return; }
    const g = fbs.filter(f => f.match === null && L.indexOf(f) < 0 && Math.abs(f.importe - resto) <= TOLERANCIA &&
      /CONFIRMACI|AGRUPADOS|LIQ/i.test(f.texto) && dias_(f.fecha, b.fecha) <= 60);
    if (g.length === 1) unir_([b], L.concat(g), 'Liquidación de tarjeta: mismo número de liquidación + confirmación de valores (revisar)');
  });
  return ajustes;
}

/** Mismo CUIT / nombre e importe con centavos de diferencia (hasta $1): ej. "veronesi" 298.280 y el recibo 298.280,13. */
function difCentavos_(banco, fbs, ventana) {
  const ajustes = [];
  banco.slice().sort((a, b) => a.fecha - b.fecha).forEach(b => {
    if (b.match !== null) return;
    const cands = fbs.filter(f => f.match === null && f.origen === 'Mes' && (f.importe > 0) === (b.importe > 0) &&
      Math.abs(f.importe - b.importe) > 0.005 && Math.abs(f.importe - b.importe) <= 1 && distancia_(b, f) <= ventana &&
      (MISMO_ID_(b, f) || POR_NOMBRE_(b, f)));
    if (cands.length !== 1) return;
    const f = cands[0], dif = redondear_(b.importe - f.importe);
    unir_([b], [f], 'Mismo CUIT o nombre, importe con centavos de diferencia (dif. ' + formato_(dif) + ')');
    const a = partidaFbs_({ fecha: f.fecha, comentario: 'Diferencia de redondeo ' + f.texto.slice(0, 40) + ' - ajustar',
      referencia: '', asiento: f.asiento, debe: Math.max(dif, 0), haber: Math.max(-dif, 0) }, 'Mes');
    a.cuenta = 'E';
    ajustes.push(a);
  });
  return ajustes;
}

function gastosAsiento_(banco, fbs) {
  // el IVA de una comisión que no va al asiento de gastos (Patagonia "COMISION POR CONVENIOS" 2.000.000 + IVA 420.000 del
  // mismo día) se cruza junto con la comisión, no con el asiento
  const ivaDeOtra = x => /^IVA/i.test(x.texto) && banco.some(c => c !== x && c.cat !== 'GASTOS BANCARIOS' && /^COMISION/i.test(c.texto) &&
    c.fecha.getTime() === x.fecha.getTime() && Math.abs(c.importe * 0.21 - x.importe) <= 0.02);
  const b = banco.filter(x => x.match === null && x.origen === 'Mes' && (x.cat === 'GASTOS BANCARIOS' || CAUSALES_IMPUESTOS.includes(x.causal)) &&
    !ivaDeOtra(x));
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
  // las liquidaciones de tarjeta van primero y solo contra movimientos de tarjeta del día de la liquidación
  const ajustesTarjeta = liquidacionPorNumero_(banco, fbs).concat(liquidacionesTarjeta_(banco, fbs));
  pasada_(banco, fbs, MISMO_ID_, '1. CUIT/DNI', VENTANA_ID, false);
  pasada_(banco, fbs, POR_REF_, '2. Referencia', 60, false);
  pasada_(banco, fbs, POR_NOMBRE_, '3. Nombre', 15, false);
  pasada_(banco, fbs, (b, f) => !!f.fechaRef && dias_(f.fechaRef, b.fecha) === 0, '3. Fecha escrita en el comprobante + importe', 60, true);
  // depósitos de cheques: el comentario trae la fecha del depósito y el banco lo acredita hasta 5 días después (fines de semana)
  // (importes redondos solo si es un cheque / depósito / canje)
  pasada_(banco, fbs, (b, f) => !!f.fechaRef && b.fecha >= f.fechaRef && dias_(f.fechaRef, b.fecha) <= 5 &&
    (especifico_(b.importe) || /CHEQ|CANJE|BOLETA|DEP/i.test(b.texto + ' ' + f.texto)),
    '3. Fecha escrita en el comprobante + importe, acreditado hasta 5 días después', 60, true);
  pasada_(banco, fbs, MISMO_NUMERO_, '3. Número de cheque + importe', 45, true);
  combinacionesFechaEscrita_(banco, fbs);
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
  if (CRUZAR_REDONDO_UNICO_DIA >= 3) { ventanaUnica_(banco, fbs, 3); mismaCantidad_(banco, fbs, 3); }
  lotes_(banco, fbs);
  difCentavos_(banco, fbs, 400).forEach(a => ajustesTarjeta.push(a));
  importeGrandeUnico_(banco, fbs, '');
  asientosNetos_(banco, fbs);
  reversionesFbs_(fbs);
  if (gastos && gastos.ajuste) fbs.push(gastos.ajuste);   // se agrega al final para que ninguna pasada la cruce
  ajustesTarjeta.forEach(a => fbs.push(a));
  marcarDuplicados_(fbs);
  return gastos ? gastos.dif : null;
}

/** Importe "específico": tiene centavos o no es múltiplo de $1.000 (123.852,68 o 7.234.300 sí; 500.000 o 5.800.000 no). */
function especifico_(x) {
  const c = Math.round(Math.abs(x) * 100);
  return c % 100 !== 0 || (c / 100) % 1000 !== 0;
}

/** Importe redondo: se cruza solo si ese día hay UN movimiento del banco y UN registro de E con ese importe. */
let CRUZAR_REDONDO_UNICO_DIA = 3;   // cruces de importe redondo con candidato único (hasta 3 días), siempre marcados "(revisar)"
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

/**
 * Mismo importe repetido: si dentro de +-dias hay la misma cantidad en el banco y en FBS (ej. tres débitos de $5.000.000
 * el 13/07 y tres registros del 10/07 y 13/07), se cruzan en orden de fecha. El grupo no puede abarcar más de 7 días.
 */
function mismaCantidad_(banco, fbs, d) {
  banco.filter(b => b.match === null && b.origen === 'Mes').forEach(b => {
    if (b.match !== null) return;
    const igual = x => x.match === null && Math.abs(x.importe - b.importe) <= TOLERANCIA;
    let bs = [b], fs = [];
    for (let vuelta = 0; vuelta < 4; vuelta++) {
      const nf = fbs.filter(f => igual(f) && bs.some(x => distancia_(x, f) <= d));
      const nb = banco.filter(x => igual(x) && x.origen === 'Mes' && nf.some(f => distancia_(x, f) <= d));
      if (nf.length === fs.length && nb.length === bs.length) break;
      fs = nf; bs = nb.length ? nb : bs;
    }
    if (bs.length < 2 || bs.length !== fs.length) return;
    // importe redondo: solo registros que no son de personas (no recibos) y con el mismo detalle ("BANCO GALICIA Nº ...")
    const detalle = f => (f.texto.split(' / ').slice(1).join(' / ') || f.texto).replace(/Nº.*$/, '').replace(/\s+/g, ' ').trim().toUpperCase();
    // (o pagos RM de banco a banco: "RM-388981 / BANCO PATAGONIA Nº ..." y "RM-388967 / BANCO GALICIA Nº ...")
    if (!especifico_(b.importe) && fs.some(f => /^RC-/.test(f.texto) ||
      (detalle(f) !== detalle(fs[0]) && !(/^RM-/.test(f.texto) && /^BANCO /i.test(detalle(f)))))) return;
    const fechas = bs.map(x => x.fecha.getTime()).concat(fs.map(f => f.fecha.getTime()));
    if ((Math.max.apply(null, fechas) - Math.min.apply(null, fechas)) / 864e5 > 7) return;
    const orden = (x, y) => x.fecha - y.fecha;
    fs.sort(orden);
    bs.sort(orden).forEach(x => {
      const f = fs.filter(y => y.match === null).sort((p, q) => distancia_(x, p) - distancia_(x, q))[0];
      unir_([x], [f], '4d. Mismo importe, misma cantidad en +-' + d + ' días (revisar)');
    });
  });
}

/** Importes grandes (desde $10.000.000) que aparecen una sola vez en cada lado dentro del mes, aunque las fechas difieran. */
function importeGrandeUnico_(banco, fbs, prefijo) {
  banco.filter(b => b.match === null && Math.abs(b.importe) >= 1e7).forEach(b => {
    const igual = x => x.match === null && Math.abs(x.importe - b.importe) <= TOLERANCIA;
    const fs = fbs.filter(f => igual(f) && distancia_(b, f) <= 31);
    if (fs.length !== 1) return;
    if (banco.filter(x => igual(x) && distancia_(x, fs[0]) <= 31).length !== 1) return;
    unir_([b], fs, prefijo + 'Importe grande único en el mes (revisar fechas)');
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
const POR_REF_ = (b, f) => b.ref.length >= 5 && ((f.valor && f.valor === b.ref) ||
  (f.texto + ' ' + f.ref + ' ' + (f.aux || '')).replace(/\./g, '').indexOf(b.ref) >= 0);
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
  // registros tardíos de meses cerrados: suelen regularizar un pendiente; mismo importe específico, hasta 5 días
  pasada_(antBanco, eItems, (b, f) => f.tardio && especifico_(b.importe), 'Pendiente anterior: registro tardío + importe', 5, true);
  // valores de Dome Taller: el Nº es la fecha de la transferencia ("Nº 96" = 9/6), aunque el importe sea redondo
  pasada_(antBanco, eItems, (b, f) => !!f.cc && CC_NUMERO_ES_FECHA.test(f.cc) && !!f.fechaRef && dias_(f.fechaRef, b.fecha) <= 3,
    'Pendiente anterior: valor de Dome Taller (fecha del Nº) + importe', v, true);
  // importe con centavos, un solo candidato en +-15 días (ej. descuento de documento del 25/06 y "01072026-G. de Echeq")
  pasada_(antBanco, eItems, b => Math.round(Math.abs(b.importe) * 100) % 100 !== 0,
    'Pendiente anterior: importe con centavos, único en +-15 días (revisar)', 15, true);
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

/**
 * Limpieza entre cuadros (paso 5 del instructivo anterior, lo que hace administración a mano): una línea de O sin
 * confirmar y un movimiento del banco sin registrar que son la misma operación se sacan de los dos cuadros.
 * Mismos criterios que el cruce con E (CUIT, referencia, nombre, fecha del valor, importe específico en la fecha),
 * y órdenes de pago con varias líneas (mismo RM) contra un solo débito del banco.
 */
/**
 * Pendientes de FBS de signo contrario (S3 contra S4), mismo importe y comentario que coincide (misma palabra o número
 * de comprobante): se anulan entre sí y no quedan pendientes. Ej. "Rv Degra pase a Chexa" en el debe de O y en el haber
 * de E; "Canje de Valor - BANCO MACRO-liq. naranja fix" contra "13072026-liq. naranja fix 13/7".
 */
/** S1 ↔ S2: una anulación / reversa del banco contra el movimiento que anula (ej. "ANULACION DEBITOS ACRED.HABERES" de
 *  $1.068.330 del 30/06 y el débito de haberes por el mismo importe del 01/07). */
function netearPendientesBanco_(items) {
  const anula = p => /ANULAC|REVERS|DEVOLUC|RECHAZ|CONTRASIENTO/i.test(p.texto);
  items.filter(anula).forEach(a => {
    if (a.match !== null) return;
    const k = tokens_(a.texto.replace(/ANULACI\S*|REVERS\S*|DEVOLUCI\S*|RECHAZ\S*|CONTRASIENTO/gi, ''));
    const b = items.filter(x => x !== a && x.match === null && Math.abs(x.importe + a.importe) <= TOLERANCIA &&
      dias_(x.fecha, a.fecha) <= 45 && comparten_(k, tokens_(x.texto)))
      .sort((x, y) => dias_(x.fecha, a.fecha) - dias_(y.fecha, a.fecha))[0];
    if (!b) return;
    const metodo = 'Neteo S1 ↔ S2: anulación del banco y el movimiento anulado';
    [a, b].forEach(x => { x.match = []; x.metodo = metodo; });
    CRUCES_.push({ metodo, banco: [a, b], fbs: [] });
  });
}

function netearPendientesFbs_(items) {
  const genericas = new Set(['CANJE', 'VALOR', 'VALORES', 'LIQ', 'CONFIRMACI', 'AGRUPADOS', 'PASE', 'DIFERENCIA', 'AJUSTAR']);
  const claves = p => {
    const t = [p.texto].concat((p.lineas || []).map(l => l.texto)).join(' ');
    const out = new Set();
    tokens_(t).forEach(x => { if (!genericas.has(x)) out.add(x); });
    (t.match(/\d{5,}/g) || []).forEach(n => out.add(n));
    (t.match(/Nº\s*\d+/g) || []).forEach(n => out.add('Nº' + n.replace(/\D/g, '')));   // "RM-389384 / FORD ... Nº 562." y "RM-389888 / ... Nº 562."
    return out;
  };
  const libres = items.filter(p => !/^Diferencia (de redondeo|asiento gastos)/.test(p.texto))
    .map(p => ({ p, k: claves(p) }))
    .sort((a, b) => Math.abs(b.p.importe) - Math.abs(a.p.importe));
  libres.forEach(a => {
    if (a.p.match !== null || a.p.importe <= 0) return;
    const cands = libres.filter(b => b.p.match === null && b.p.importe < 0 && Math.abs(a.p.importe + b.p.importe) <= TOLERANCIA &&
      (a.p.texto === b.p.texto || comparten_(a.k, b.k)));
    if (!cands.length) return;
    const cerca = b => (String(a.p.asiento) === String(b.p.asiento) ? -1 : dias_(a.p.fecha, b.p.fecha));
    const b = cands.sort((x, y) => cerca(x) - cerca(y))[0];
    compensar_([a.p, b.p], 'Neteo S3 ↔ S4: mismo importe y comentario');
  });
}

function limpiezaO_(banco, oPend) {
  const b = banco.filter(x => x.match === null), o = oPend.filter(x => x.match === null && !x.desfasado);
  pasada_(b, o, MISMO_ID_, 'Limpieza O sin confirmar ↔ banco: CUIT/DNI', VENTANA_ID, false);
  pasada_(b, o, POR_REF_, 'Limpieza O sin confirmar ↔ banco: referencia', 60, false);
  pasada_(b, o, POR_NOMBRE_, 'Limpieza O sin confirmar ↔ banco: nombre', 15, false);
  pasada_(b, o, (x, f) => !!f.fechaRef && dias_(f.fechaRef, x.fecha) === 0, 'Limpieza O sin confirmar ↔ banco: fecha del valor + importe', 60, true);
  pasada_(b, o, x => especifico_(x.importe), 'Limpieza O sin confirmar ↔ banco: fecha + importe específico', 0, true);
  importeGrandeUnico_(b, o, 'Limpieza O sin confirmar ↔ banco: ');
  // varias líneas de O (mismo comprobante o mismo asiento) contra uno o varios débitos del banco del mismo día y causal
  const grupos = new Map();
  o.forEach(f => {
    if (f.match !== null) return;
    [claveO_(f.texto) + '|' + f.fecha.getTime(), 'as:' + f.asiento].forEach(k => grupos.set(k + '|' + (f.importe > 0), (grupos.get(k + '|' + (f.importe > 0)) || []).concat([f])));
  });
  const gb = new Map();
  const cz = y => y.causal || y.cat || y.texto.slice(0, 20);   // Patagonia no trae código: se agrupa por categoría
  b.forEach(y => { if (y.origen === 'Mes') { const k = y.fecha.getTime() + '|' + cz(y) + '|' + (y.importe > 0); gb.set(k, (gb.get(k) || []).concat([y])); } });
  grupos.forEach(g => {
    if (g.length < 2 || g.some(f => f.match !== null)) return;
    const total = g.reduce((x, f) => x + f.importe, 0);
    const x = b.find(y => y.match === null && Math.abs(y.importe - total) <= TOLERANCIA && dias_(y.fecha, g[0].fecha) <= 3);
    if (x) { unir_([x], g, 'Limpieza O sin confirmar ↔ banco: comprobante en varias líneas'); return; }
    for (const bs of gb.values()) {
      if (bs.length > 1 && bs.every(y => y.match === null) && dias_(bs[0].fecha, g[0].fecha) <= 3 &&
        Math.abs(bs.reduce((a, y) => a + y.importe, 0) - total) <= TOLERANCIA) { unir_(bs, g, 'Limpieza O sin confirmar ↔ banco: lote del día'); return; }
    }
    // ej. sueldos: un asiento de O con todas las transferencias contra los débitos del mismo concepto en los días siguientes
    const causales = new Set(b.filter(y => y.match === null && y.origen === 'Mes' && (y.importe > 0) === (total > 0)).map(cz));
    for (const c of causales) {
      const bs = b.filter(y => y.match === null && y.origen === 'Mes' && cz(y) === c && (y.importe > 0) === (total > 0) &&
        y.fecha >= g[0].fecha && dias_(y.fecha, g[0].fecha) <= 10);
      if (bs.length > 1 && Math.abs(bs.reduce((a, y) => a + y.importe, 0) - total) <= TOLERANCIA) {
        unir_(bs, g, 'Limpieza O sin confirmar ↔ banco: lote (mismo concepto en varios días)'); return;
      }
      // no todos: los débitos del concepto que suman el asiento (ej. 4 débitos de sueldos del 06 al 13/07 y otro del 16/07)
      for (let n = 2; n <= Math.min(6, bs.length - 1); n++) {
        const combo = primeraCombinacion_(bs.slice(0, 15), n, total);
        if (combo) { unir_(combo, g, 'Limpieza O sin confirmar ↔ banco: lote (mismo concepto en varios días)'); return; }
      }
    }
  });
  if (CRUZAR_REDONDO_UNICO_DIA) unicoDelDia_(b, o);
}

function sector_(p) {
  // un pendiente anterior queda en el sector donde estaba (ej. la comisión de una liquidación listada en negativo en S3)
  if (p.origen === 'Arrastre' && p.sectorOriginal) return p.sectorOriginal;
  if (p.lado === 'BANCO') return p.importe > 0 ? 'S1' : 'S2';
  return p.importe > 0 ? 'S3' : 'S4';
}

/** Importe de la partida en su sector: positivo, salvo un pendiente anterior que estaba en negativo. */
function valorSector_(p) {
  const s = sector_(p);
  return redondear_(s === 'S1' || s === 'S3' ? p.importe : -p.importe);
}

function gastosBancarios_(movs) {
  const g = { com: 0, iva: 0, iva105: 0, perc: 0, int: 0, sircreb: 0, ret: 0, tuc: 0, sellos: 0, imp: 0 };
  // el IVA de una comisión que se registra aparte (Patagonia "COMISION POR CONVENIOS") no va al asiento de gastos
  const ivaDeOtra = x => /^IVA/i.test(x.concepto) && movs.some(c => c !== x && c.cat !== 'GASTOS BANCARIOS' && /^COMISION/i.test(c.concepto) &&
    c.fecha.getTime() === x.fecha.getTime() && Math.abs(c.importe * 0.21 - x.importe) <= 0.02);
  movs.filter(m => !ivaDeOtra(m)).forEach(m => {
    const t = m.concepto.toUpperCase(), d = -m.importe;
    if (['1684', '1685', '907176', '907376', '917376', '4637', '4633'].includes(m.causal)) g.imp += d;
    else if (m.causal === '1384') g.sircreb += d;   // Santander
    else if (m.causal === '566') g.tuc += d;        // Santander
    else if (m.causal === '907173') { if (t.includes('TUCUMAN')) g.tuc += d; else g.sircreb += d; }
    else if (m.causal === '905130') g.sellos += d;
    else if (m.causal === '907307') g.iva105 += d;
    else if (m.causal === '907118' || m.causal === '907751') g.int += d;
    else if (m.causal === '1297') g.sircreb += d;
    else if (m.causal === '4145') g.ret += d;
    else if (m.causal === '1972') g.tuc += d;
    else if (m.causal === '1479') g.sellos += d;
    else if (m.cat === 'GASTOS BANCARIOS') {
      // Patagonia: sin código, por la descripción
      if (t.startsWith('IMP.DB/CR') || t.startsWith('IMPUESTO SOBRE OPERACIONES')) g.imp += d;
      else if (t.includes('SIRCREB')) g.sircreb += d;
      else if (t.includes('IIBB TUCUMAN')) g.tuc += d;
      else if (t.includes('SELLOS')) g.sellos += d;
      else if (t.startsWith('INTERES')) g.int += d;
      else if (t.includes('PERCEP') || t.includes('IVA_PER')) g.perc += d;
      else if (t.includes('IVA')) { if (m.causal === '5') g.iva105 += d; else g.iva += d; }
      else if (m.causal === '5') g.int += d;
      else g.com += d;
    }
  });
  // IVA sin distinguir en la descripción (Patagonia "IVA ALICUOTA GENERAL"): la línea igual al 10,5% de los intereses es IVA 10,5%
  if (g.int && !g.iva105) {
    const m = movs.find(x => x.cat === 'GASTOS BANCARIOS' && /^IVA/i.test(x.concepto) && Math.abs(-x.importe - redondear_(g.int * 0.105)) <= 0.02);
    if (m) { g.iva105 = -m.importe; g.iva += m.importe; }
  }
  const filas = [['Comisiones', g.com], ['IVA 21%', g.iva], ['Intereses', g.int], ['IVA 10,5%', g.iva105],
    ['Percepciones IVA', g.perc], ['Percepciones IIBB (SIRCREB)', g.sircreb], ['Retenciones IIBB rentas financieras', g.ret],
    ['Percepciones IIBB Tucumán', g.tuc], ['Impuesto de sellos', g.sellos], ['Imp. créditos y débitos (total)', g.imp],
    ['    Crédito computable 33%', g.imp * 0.33], ['    Imp. créditos y débitos (gasto 67%)', g.imp * 0.67]];
  const total = g.com + g.iva + g.int + g.iva105 + g.perc + g.sircreb + g.ret + g.tuc + g.sellos + g.imp;
  return filas.filter(f => Math.abs(f[1]) > 0.004).concat([['TOTAL (debe coincidir con el asiento de FBS)', total]]);
}

/** Agrega a la hoja "Reglas" las reglas del script que todavía no están (no toca las existentes). */
// ------------------------------------------------------------------ revisar pendientes entre cierres

/**
 * Revisa la lista "Pendientes en trabajo" contra la base FBS al día (movimientos cargados después del último cierre,
 * más los tardíos de meses cerrados) y escribe en "Revisión automática" lo que encuentra. No cambia Sector, Fecha,
 * Concepto ni Importe (eso lo sigue usando el proceso del mes): solo informa y, si quedó resuelto, sugiere el estado.
 */
function revisarPendientes() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaDe_(ss, nombresAnteriores_());
  if (!sh || sh.getLastRow() < 2) throw new Error('No hay pendientes para revisar.');
  const parametros = ss.getSheetByName(HOJA.parametros) ? ss.getSheetByName(HOJA.parametros).getDataRange().getValues() : null;
  const idBase = idDeLink_(paramDe_(parametros, '^base fbs'));
  if (!idBase) throw new Error('Falta el link de la base FBS en "' + HOJA.parametros + '".');
  let cierre = parseFecha_(paramDe_(parametros, '^.ltimo cierre'));
  if (!cierre) {
    const hist = ss.getSheetByName(HOJA.historial);
    const ult = hist && hist.getLastRow() > 1 ? String(hist.getRange(hist.getLastRow(), 1).getValue()) : '';
    const m = ult.match(/^(\d{4})-(\d{2})$/);
    if (m) cierre = new Date(Number(m[1]), Number(m[2]), 0);
  }
  if (!cierre) throw new Error('No sé la fecha del último cierre: cargala en "' + HOJA.parametros + '" (fila "Último cierre").');
  const ctas = cuentasDe_(parametros);
  const hoy = new Date();
  // movimientos de E y O después del cierre + los tardíos (meses cerrados cambiados después de la foto)
  const desde = new Date(cierre.getFullYear(), cierre.getMonth(), cierre.getDate() + 1);
  const r = leerBaseExterna_(idBase, paramDe_(parametros, '^hoja de la base') || 'MOVIMIENTOS', ctas, desde, new Date(hoy.getFullYear() + 1, 0, 1));
  let filas = r.filas;
  if (r.ctx) {
    const foto = leerFotoBase_(ss);
    if (foto) { const t = movimientosTardios_(r.ctx, desde, {}, foto); if (t.filas.length > 1) filas = unirFilas_(filas, t.filas); }
  }
  const b = leerBase_(filas, ctas.E, ctas.O || '__sin_cuenta_O__');
  const movs = b.E.map(a => Object.assign(a, { cuenta: 'E' })).concat(b.O.map(a => Object.assign(a, { cuenta: 'O' })));
  const v = sh.getDataRange().getValues();
  const pend = v.slice(1).map(x => [x[0], x[1], x[2], x[3]]);
  const res = revisarPendientesCore_(pend, movs, cierre);
  // pendientes del banco sin rastro: ¿están registrados en la cuenta de otro banco?
  if (r.ctx) {
    const sinRastro = res.filter((x, i) => !x.texto && /^S[12]$/.test(String(pend[i][0]).trim().toUpperCase()) && especifico_(parseNum_(pend[i][3]) || 0));
    if (sinRastro.length) otrasCuentas_(r.ctx, pend, res, desde);
  }
  // escribir
  const h = v[0].map(String), cRev = h.indexOf('Revisión automática'), cFecha = h.indexOf('Revisado el'), cEst = h.indexOf('Estado');
  if (cRev < 0) throw new Error('La hoja "' + sh.getName() + '" no tiene las columnas de trabajo: se agregan solas en el próximo "Cerrar mes".');
  const rev = res.map(x => [x.texto || 'Sin novedades en FBS']), fec = res.map(() => [hoy]);
  sh.getRange(2, cRev + 1, rev.length, 1).setValues(rev);
  sh.getRange(2, cFecha + 1, rev.length, 1).setValues(fec).setNumberFormat('dd/mm/yyyy hh:mm');
  const colores = res.map(x => [x.tipo === 'resuelto' ? '#C6EFCE' : x.tipo === 'posible' ? '#FFEB9C' : null]);
  sh.getRange(2, cRev + 1, rev.length, 1).setBackgrounds(colores);
  if (cEst >= 0) {
    const est = v.slice(1).map((x, i) => [res[i].tipo === 'resuelto' && (!x[cEst] || x[cEst] === 'Pendiente') ? 'Resuelto en FBS' : x[cEst]]);
    sh.getRange(2, cEst + 1, est.length, 1).setValues(est);
  }
  const n = { resuelto: 0, posible: 0 };
  res.forEach(x => { if (n[x.tipo] !== undefined) n[x.tipo]++; });
  sh.activate();
  SpreadsheetApp.getUi().alert('Revisar pendientes', 'Revisé ' + res.length + ' pendientes contra FBS (movimientos desde el ' + fechaTexto_(desde) +
    ').\n\n' + n.resuelto + ' resueltos (verde)\n' + n.posible + ' con un posible cruce para mirar (amarillo)\n\nVer columna "Revisión automática".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Núcleo (sin hojas): pend = [[sector, fecha, concepto, importe]], movs = líneas de E y O cargadas después del cierre
 * (formato de leerBase_ con .cuenta). Devuelve por pendiente { tipo: 'resuelto' | 'posible' | '', texto }.
 */
function revisarPendientesCore_(pend, movs, cierre) {
  CRUCES_ = [];
  CORTE_ = null;
  const out = pend.map(() => ({ tipo: '', texto: '' }));
  const items = pend.map((r, i) => {
    const sec = String(r[0] || '').trim().toUpperCase(), imp = parseNum_(r[3]) || 0;   // negativo: invierte el sentido
    const fecha = parseFecha_(r[1]) || cierre, texto = String(r[2] || '');
    let p = null;
    const v = sec === 'S1' || sec === 'S3' ? imp : -imp;
    if (sec === 'S1' || sec === 'S2') p = partidaBanco_({ fecha, concepto: texto, importe: v, ref: '', nombre: '' }, 'Arrastre');
    else if (sec === 'S3' || sec === 'S4') p = partidaFbs_({ fecha, comentario: texto, referencia: '', debe: Math.max(v, 0), haber: Math.max(-v, 0) }, 'Arrastre');
    if (p) { p.idx = i; p.sec = sec; }
    return p;
  });
  const banco = items.filter(p => p && p.lado === 'BANCO'), fbsPend = items.filter(p => p && p.lado === 'FBS');
  const nuevos = movs.map(a => { const p = partidaFbs_(a, 'Mes'); p.cuenta = a.cuenta; return p; });
  const desc = f => f.texto + ' (cuenta ' + f.cuenta + ', asiento ' + f.asiento + ', ' + fechaTexto_(f.fecha) + (f.tardio ? ', mes ya cerrado' : '') + ')';
  // 1. banco sin registrar (S1 / S2): ¿se registró después del cierre?
  const v = 400;
  pasada_(banco, nuevos, MISMO_ID_, 'CUIT/DNI', v, false);
  pasada_(banco, nuevos, POR_REF_, 'referencia', v, false);
  pasada_(banco, nuevos, POR_NOMBRE_, 'nombre', 60, false);
  pasada_(banco, nuevos, (x, f) => !!f.fechaRef && dias_(f.fechaRef, x.fecha) === 0 && especifico_(x.importe), 'fecha del valor + importe', v, true);
  pasada_(banco, nuevos, (x, f) => f.tardio && especifico_(x.importe), 'registro tardío + importe', 5, true);
  banco.forEach(x => {
    if (x.match === null) return;
    // CUIT, referencia o nombre: resuelto; solo fecha e importe: posible (a confirmar)
    const seguro = /^(CUIT|referencia|nombre)/.test(x.metodo);
    out[x.idx] = { tipo: seguro ? 'resuelto' : 'posible', texto: (seguro ? 'Registrado en FBS por ' : 'Posible registro en FBS (por ') + x.metodo +
      (seguro ? ': ' : ', confirmar): ') + x.match.map(desc).join(' + ') };
  });
  // importe específico único después del cierre: posible, para mirar
  pasada_(banco.filter(x => x.match === null), nuevos, x => especifico_(x.importe), 'importe', 400, true);
  banco.forEach(x => {
    if (x.match !== null && !out[x.idx].texto) out[x.idx] = { tipo: 'posible', texto: 'Posible registro (mismo importe, sin CUIT ni nombre): ' + x.match.map(desc).join(' + ') };
  });
  // 2. FBS sin banco (S3 / S4): ¿se anuló o se confirmó en FBS?
  fbsPend.forEach(p => {
    const clave = claveO_(p.texto), comp = /^(RC-|RM-|LIQ )/.test(clave);
    const rev = nuevos.find(f => f.match === null && Math.abs(f.importe + p.importe) <= TOLERANCIA &&
      ((comp && claveO_(f.texto) === clave) || comparten_(f.nombre, p.nombre) || /ANULA|REVERS/i.test(f.texto)));
    if (!rev) return;
    rev.match = [p];
    const pase = rev.cuenta === 'O' && nuevos.some(f => f !== rev && f.cuenta === 'E' && String(f.asiento) === String(rev.asiento) && Math.abs(f.importe + rev.importe) <= TOLERANCIA);
    out[p.idx] = pase
      ? { tipo: 'posible', texto: 'Confirmado en FBS (pase de O a E) el ' + fechaTexto_(rev.fecha) + ', asiento ' + rev.asiento + ': se va a cruzar con el banco al procesar el mes' }
      : { tipo: 'resuelto', texto: 'Revertido o anulado en FBS: ' + desc(rev) };
  });
  // 3. pendientes que se anulan entre sí
  const libresB = banco.filter(x => x.match === null && !out[x.idx].texto);
  libresB.forEach(x => {
    if (x.match !== null || x.importe <= 0) return;
    const y = libresB.find(z => z !== x && z.match === null && Math.abs(z.importe + x.importe) <= TOLERANCIA &&
      (MISMO_ID_(x, z) || comparten_(x.nombre, z.nombre) || normalizar_(x.texto) === normalizar_(z.texto)));
    if (!y) return;
    x.match = [y]; y.match = [x];
    out[x.idx] = out[y.idx] = { tipo: 'resuelto', texto: 'Se anula con el otro pendiente del banco: ' + fechaTexto_(x.fecha) + ' ' + x.texto + ' / ' + fechaTexto_(y.fecha) + ' ' + y.texto };
  });
  const libresF = fbsPend.filter(p => !out[p.idx].texto);
  netearPendientesFbs_(libresF);
  libresF.forEach(p => {
    if (p.match === null) return;
    const otro = CRUCES_.filter(c => c.fbs.indexOf(p) >= 0).map(c => c.fbs.find(q => q !== p))[0];
    out[p.idx] = { tipo: 'resuelto', texto: 'Se anula con el otro pendiente de FBS (mismo importe y comentario)' + (otro ? ': ' + fechaTexto_(otro.fecha) + ' ' + otro.texto : '') };
  });
  return out;
}

/** S1/S2 sin rastro en E/O: busca en la base el mismo importe registrado en la cuenta de otro banco después del pendiente. */
function otrasCuentas_(ctx, pend, res, desde) {
  const L = ctx.L;
  const sh = SpreadsheetApp.openById(ctx.id).getSheetByName(ctx.hoja);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(normalizar_);
  const col = re => { const i = cab.findIndex(x => re.test(x)); return i < 0 ? null : letra_(i + 1); };
  const nombre = col(/^cuentanombre$/), com = col(/^detcomenta$/), asi = col(/^detmovnro$/);
  const pedidos = [];
  res.forEach((x, i) => {
    const sec = String(pend[i][0]).trim().toUpperCase(), imp = Math.abs(parseNum_(pend[i][3]) || 0);
    if (!x.texto && /^S[12]$/.test(sec) && especifico_(imp)) pedidos.push({ i, sec, imp, fecha: parseFecha_(pend[i][1]) });
  });
  const propias = [ctx.ctas.E, ctx.ctas.O].filter(Boolean).map(c => L.plan + ' <> ' + Number(c)).join(' and ');
  for (let k = 0; k < pedidos.length; k += 30) {
    const grupo = pedidos.slice(k, k + 30);
    const conds = grupo.map(g => (g.sec === 'S1' ? L.debe : L.haber) + ' = ' + g.imp).join(' or ');
    const sel = [L.plan, nombre, asi, L.fecha, L.debe, L.haber, com].filter(Boolean);
    const filas = consultaGviz_(ctx.id, ctx.hoja, 'select ' + sel.join(', ') + ' where ' + L.plan + ' >= 1103000 and ' + L.plan + ' < 1105000 and ' +
      propias + ' and (' + conds + ')' + ctx.vigente);
    const h = filas[0].map(normalizar_);
    const ix = n => h.indexOf(n);
    grupo.forEach(g => {
      const hit = filas.slice(1).find(r => {
        const imp = g.sec === 'S1' ? Number(r[ix('detdebe')]) : Number(r[ix('dethaber')]);
        const f = r[ix('detfecha')] instanceof Date ? r[ix('detfecha')] : parseFecha_(r[ix('detfecha')]);
        return Math.abs(imp - g.imp) <= TOLERANCIA && f && (!g.fecha || f >= new Date(g.fecha.getTime() - 5 * 86400000));
      });
      if (hit) res[g.i] = { tipo: 'posible', texto: 'Posible registro en otra cuenta: ' + (ix('cuentanombre') >= 0 ? hit[ix('cuentanombre')] : hit[0]) +
        ' (' + hit[ix('detplan')] + (ix('detmovnro') >= 0 ? ', asiento ' + hit[ix('detmovnro')] : '') + ', ' + fechaTexto_(parseFecha_(hit[ix('detfecha')])) +
        (ix('detcomenta') >= 0 ? ', ' + hit[ix('detcomenta')] : '') + ')' };
    });
  }
}

function agregarReglasNuevas() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(HOJA.reglas);
  if (!sh) throw new Error('Falta la hoja "' + HOJA.reglas + '". Corré "1. Crear hojas de entrada".');
  const clave = r => [String(r[0]).trim(), normalizar_(r[1]).charAt(0), String(r[2] || '').trim()].join('|');
  const existentes = new Set(sh.getDataRange().getValues().slice(1).map(clave));
  const par = ss.getSheetByName(HOJA.parametros);
  const banco = par ? paramDe_(par.getDataRange().getValues(), '^cuenta bancaria$') : '';
  const nuevas = reglasDelBanco_(banco).filter(r => !existentes.has(clave(r)));
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 5).setValues(nuevas);
  SpreadsheetApp.getUi().alert(nuevas.length + ' reglas nuevas agregadas a "' + HOJA.reglas + '".');
}

/** Todo el proceso sobre matrices de valores (sin tocar hojas): se puede probar fuera de Sheets. */
function conciliarTodo_(entrada, redondoUnico) {
  // un solo modo de cruce (el que antes se llamaba INTERMEDIO); el parámetro solo lo usan las pruebas
  CRUZAR_REDONDO_UNICO_DIA = redondoUnico !== undefined ? redondoUnico : 3;
  const reglas = leerReglas_(entrada.reglas), empresas = leerEmpresas_(entrada.empresas);
  const ext = leerExtracto_(entrada.extracto, reglas, empresas);
  const movs = ext.movs;
  const param = nombre => { const f = (entrada.parametros || []).find(r => new RegExp(nombre, 'i').test(String(r[0]))); return f ? f[1] : null; };
  let e, o;
  if (entrada.base) {
    const ctas = cuentasDe_(entrada.parametros), sinO = !ctas.O;
    const cuentaO = sinO ? '__sin_cuenta_O__' : ctas.O;
    const b = leerBase_(entrada.base, ctas.E, cuentaO);
    let siE = parseNum_(param('^saldo inicial cuenta e'));
    let siO = sinO ? 0 : parseNum_(param('^saldo inicial cuenta o'));
    if (entrada.saldosIniciales) {
      // base externa. Manda el saldo del cierre anterior (Parámetros): es el que se usó para armar los pendientes
      // anteriores. La base suma todo lo anterior al mes como está HOY, así que incluye lo que se cargó o modificó en FBS
      // con fecha de meses ya cerrados después del cierre; esa diferencia se informa. Sin saldo en Parámetros (primer
      // mes) se usa el de la base.
      const avisoSaldo = [], tard = entrada.tardios || { E: 0, O: 0 };
      const comparar = (n, cargado, calculado) => {
        if (cargado === null) return;
        const resto = redondear_(calculado - cargado - tard[n]);
        if (Math.abs(resto) > 1) avisoSaldo.push('cuenta ' + n + ': cierre anterior ' + formato_(cargado) + ', base hoy ' + formato_(calculado) +
          (tard[n] ? ', registros tardíos ' + formato_(tard[n]) : '') + ' (sin explicar ' + formato_(resto) + ')');
      };
      comparar('E', siE, entrada.saldosIniciales.E);
      if (!sinO) comparar('O', siO, entrada.saldosIniciales.O);
      if (tard.E || tard.O) entrada.avisosBase = (entrada.avisosBase || []).concat(['Se sumaron a este mes movimientos de meses ya cerrados ' +
        'cargados o modificados en FBS después del cierre (E ' + formato_(tard.E) + ', O ' + formato_(tard.O) + '): se cruzan contra los pendientes ' +
        'anteriores (ver "Movimiento con fecha ... (mes ya cerrado)" en las hojas de detalle).']);
      if (avisoSaldo.length) entrada.avisosBase = (entrada.avisosBase || []).concat(['La base no coincide con el saldo del cierre anterior: ' +
        avisoSaldo.join('; ') + '. Se usó el saldo del cierre anterior.']);
      if (siE === null) siE = entrada.saldosIniciales.E;
      if (siO === null) siO = sinO ? 0 : entrada.saldosIniciales.O;
    }
    const mk = (lista, si, n) => ({ asientos: lista, avisos: [], info: { cuenta: n, saldoInicial: si || 0,
      saldoFinal: redondear_((si || 0) + lista.reduce((x, a) => x + a.debe - a.haber, 0)) } });
    e = mk(b.E, siE, 'E (' + ctas.E + ')');
    o = mk(b.O, siO, sinO ? 'O (no se usa)' : 'O (' + cuentaO + ')');
    if (siE === null || siO === null) {
      // sin saldos iniciales: se deducen de la conciliación anterior (saldo inicial del banco - pendientes anteriores)
      const antes = leerAnteriores_(entrada.anteriores || [], new Date(movs[0].fecha.getFullYear(), movs[0].fecha.getMonth(), 1));
      const netoAnt = antes.reduce((x, p) => x + (p.lado === 'BANCO' ? p.importe : -p.importe), 0);
      const total = redondear_(movs[0].saldo - movs[0].importe - netoAnt);
      e.info.saldoInicial = total; o.info.saldoInicial = 0;
      e.info.saldoFinal = redondear_(total + b.E.reduce((x, a) => x + a.debe - a.haber, 0));
      o.info.saldoFinal = redondear_(b.O.reduce((x, a) => x + a.debe - a.haber, 0));
      e.info.cuenta += ' — saldo inicial E+O deducido';
      e.avisos.push('Faltan los saldos iniciales de E y O en "' + HOJA.parametros + '": se dedujeron de la conciliación anterior ' +
        '(saldo inicial del banco menos pendientes anteriores), así que el control de apertura no aplica. Cargalos para controlar la apertura.');
    }
  } else {
    e = leerMayor_(entrada.mayorE, HOJA.mayorE); o = leerMayor_(entrada.mayorO, HOJA.mayorO);
  }
  const desde = new Date(movs[0].fecha.getFullYear(), movs[0].fecha.getMonth(), 1);
  const corte = movs[movs.length - 1].fecha;
  const avisos = (entrada.avisosBase || []).concat(e.avisos, o.avisos);

  // controles de los archivos
  let quiebres = 0;
  for (let i = 1; i < movs.length; i++) {
    if (movs[i].saldo !== null && movs[i - 1].saldo !== null && Math.abs(movs[i - 1].saldo + movs[i].importe - movs[i].saldo) > 0.005) quiebres++;
  }
  if (quiebres) avisos.push('Extracto: ' + quiebres + ' filas donde el saldo no sigue a la anterior (¿faltan movimientos?).');
  if (ext.info.mesesAnteriores) avisos.push('Extracto: trae ' + ext.info.mesesAnteriores + ' movimientos de meses anteriores; ' +
    'se concilia solo el último mes.');
  const nuevos = movs.filter(m => m.desc === 'Código causal nuevo');
  if (nuevos.length) avisos.push('Extracto: ' + nuevos.length + ' movimientos con código causal sin regla (' +
    Array.from(new Set(nuevos.map(m => m.causal))).join(', ') + '). Agregalos en "' + HOJA.reglas + '".');

  CRUCES_ = [];
  CORTE_ = corte;
  const bancoMes = movs.map(m => partidaBanco_(m, 'Mes'));
  const anteriores = leerAnteriores_(entrada.anteriores || [], desde);
  const anterioresBanco = anteriores.filter(p => p.lado === 'BANCO');
  const anterioresFbs = anteriores.filter(p => p.lado === 'FBS');
  // PASO 1: cuenta O sola, neteando por comprobante / liquidación junto con los pendientes anteriores
  const ao = analisisO_(o.asientos, anterioresFbs);
  const asientosE = pasesSinOrigen_(ao, e.asientos);
  // PASO 2: cuenta E -> primero contra los pendientes anteriores, después contra el extracto
  const eItems = partidasE_(asientosE, o.asientos);
  fechaEnO_(eItems, o.asientos);
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
    avisos.unshift('La apertura no cierra por ' + formato_(apertura.diferencia) + ': revisá los saldos iniciales de E y O' +
      (entrada.base ? ' en "' + HOJA.parametros + '"' : '') + ' (tienen que ser los saldos al cierre del mes anterior) y los pendientes ' +
      'anteriores. Esa misma diferencia se arrastra al control del mes.');
  }
  const difGastos = conciliar_(banco, fbs);
  limpiezaO_(banco, ao.pendientes);
  netearPendientesFbs_(fbs.filter(p => p.match === null).concat(ao.pendientes.filter(p => p.match === null && !p.desfasado)));
  netearPendientesBanco_(banco.filter(p => p.match === null));
  // pendientes anteriores del banco que recién quedan libres después de netear FBS (ej. tarjeta cargada, anulada y
  // vuelta a cargar con la fecha del crédito escrita: "04050405-CREDITO X TARJETA 04/05")
  pasada_(banco.filter(p => p.match === null && p.origen === 'Arrastre'), fbs.filter(p => p.match === null && p.origen === 'Mes'),
    (b, f) => !!f.fechaRef && dias_(f.fechaRef, b.fecha) === 0 && especifico_(b.importe), 'Pendiente anterior: fecha del comentario + importe', 400, true);
  ao.pendientes = ao.pendientes.filter(p => p.match === null);
  marcarCruces_(ao.pendientes, anteriores, eItems, bancoMes);

  const pend = { S1: [], S2: [], S3: [], S4: [] };
  banco.concat(fbs).forEach(p => { if (p.match === null) pend[sector_(p)].push(p); });
  ao.pendientes.forEach(p => pend[sector_(p)].push(p));   // no confirmados en O: directo a pendientes
  const tot = {};
  Object.keys(pend).forEach(s => { tot[s] = redondear_(pend[s].reduce((x, p) => x + valorSector_(p), 0)); });
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
    // las partidas agrupadas se muestran con sus líneas originales
    const lineasFbs = [].concat.apply([], c.fbs.map(f => (f.lineas && f.lineas.length > 1
      ? f.lineas.slice().sort((x, y) => Math.abs(y.importe) - Math.abs(x.importe)).map(l => Object.assign({}, f, l)) : [f])));
    for (let i = 0; i < Math.max(c.banco.length, lineasFbs.length, 1); i++) {
      const x = c.banco[i], y = lineasFbs[i];
      conciliados.push([c.metodo, x ? x.fecha : '', x ? x.texto : '', x ? x.importe : '', x ? x.origen : '',
        y ? y.fecha : '', y ? y.texto : '', y ? y.importe : '', y ? y.asiento : '', y ? (y.origen === 'Arrastre' ? 'Pendiente anterior' : 'Cuenta ' + (y.cuenta || 'E')) : '']);
    }
  });
  const listas = { O: ao.pendientes, E: fbs.filter(p => p.match === null), B: banco.filter(p => p.match === null) };

  let empresa = '';
  if (entrada.base) {
    const ie = (entrada.base[0] || []).map(normalizar_).indexOf('empresa');
    if (ie >= 0) { const f = entrada.base.slice(1).find(r => String(r[ie] || '').trim()); empresa = f ? String(f[ie]).trim() : ''; }
  }
  return { empresa, cuentaBanco: param('^cuenta bancaria') || 'Banco Macro', info: ext.info, infoE: e.info, infoO: o.info, corte, saldoE, saldoO, saldoBanco, pend, tot, ajustado, apertura, analisisO: ao.resumen, listas,
    metodos, difGastos, gastos: gastosBancarios_(movs), conciliados, avisos, banco, fbs };
}

if (typeof module !== 'undefined') module.exports = { REGLAS_CORDOBA, REGLAS_PATAGONIA, REGLAS_SANTANDER, REGLAS_GALICIA, reglasDelBanco_, revisarPendientesCore_, idsPend_, clavePend_, parseGviz_, cuentasDe_, idDeLink_, letra_, armarCuatroPuntas_, armarRevisar_, conciliarTodo_, leerExtracto_, leerMayor_, leerReglas_, leerEmpresas_, parseNum_, parseFecha_, REGLAS_INICIALES, EMPRESAS_INICIALES };

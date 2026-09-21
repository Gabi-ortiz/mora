/**
 * TABLERO DE CONTROL DE MORA — v4
 *
 * Problema que resuelve: la columna "Avance" de BASE es LIVE y se mueve
 * sola con el tiempo (un plan que hoy está en avance=10 puede pasar a
 * avance=11 en cualquier momento del mes). Si el cuadro de "Cuota 9 AGOSTO"
 * se recalcula tarde, deja de coincidir con lo que fue realmente el corte.
 *
 * Solución: separar la FOTO (snapshot diario, inmutable, guardado en la
 * hoja "Historial_Mora") del TABLERO (que lee de ese historial, no de la
 * base en vivo). Así elegís qué fecha mirar y el número queda fijo para
 * siempre, sin depender de agarrar el día exacto del corte.
 *
 * v3: cada fila del historial queda etiquetada con MesAnalisis (mes en que
 * se saca la foto) y PeriodoReal (mes atrasado que realmente se evalúa).
 *
 * v4: el trigger corre todos los días, pero SOLO escribe una foto nueva si
 * BASE (que se completa con IMPORTRANGE desde el archivo real) tuvo cambios
 * en los datos relevantes (Avance, Estado, C2..C14) desde la última vez.
 * Si no hubo cambios, no duplica nada en Historial_Mora — solo deja
 * constancia en la hoja "Log_Snapshot" de que se revisó y no había novedad.
 *
 * Instalación:
 * 1) Extensiones > Apps Script, pegar este código.
 * 2) Correr una vez "instalarTriggerDiario" para que el snapshot se chequee
 *    solo todos los días. También queda en el menú "Mora".
 * 3) Menú "Mora" > "Sacar foto ahora" para forzar la revisión/snapshot
 *    manualmente cuando quieras.
 * 4) Menú "Mora" > "Actualizar tablero" para refrescar la hoja "Tablero"
 *    con las 5 cuotas, usando la última foto disponible de cada avance.
 * 5) Menú "Mora" > "Actualizar detalle de planes" para refrescar la hoja
 *    "Detalle_Planes": un listado (uno por plan) con filtro nativo de
 *    Sheets por Avance y por clasificación de mora, para poder ver a qué
 *    planes puntuales corresponde cada número del Tablero.
 */

// --- CONFIGURACIÓN ---
const HOJA_BASE = 'BASE';
const HOJA_HISTORIAL = 'Historial_Mora';
const HOJA_TABLERO = 'Tablero';
const HOJA_LOG = 'Log_Snapshot';
const HOJA_DETALLE = 'Detalle_Planes';
const PROP_ULTIMO_HASH = 'ULTIMO_HASH_BASE';

const COL_AVANCE = 14; // N
const COL_ESTADO = 15; // O
const COL_C2 = 18;     // R = primera cuota del rango (C2)

// Columnas identificatorias del plan, para el detalle por plan.
const COL_SOLICITUD = 2;  // B
const COL_GRUPO = 3;      // C
const COL_CLIENTE = 8;    // H — "NyAP-Razon Social"
const COL_TELEFONO = 9;   // I
const COL_VENDEDOR = 16;  // P
const COL_SUPERVISOR = 17; // Q

// Qué cuotas te interesa ver en el tablero (avance a mirar = cuota+1)
const CUOTAS_TABLERO = [3, 5, 7, 9, 12];

// Objetivos de mora / incentivo por cuota (Jul-Sep 2026, esquema "80/20").
// Por ahora se aplica el mismo objetivo a toda la cartera, sin distinguir
// tipo de plan (70/30, 80/20, 90/10, etc.) — ver README.
// t1/i1: mora por debajo de t1 -> incentivo i1.
// t2/i2: mora entre t1 y t2 (inclusive) -> incentivo i2.
// Mora por encima de t2 -> incentivo 0.
const OBJETIVOS_INCENTIVO = {
  3: { t1: 0.35, i1: 0.0035, t2: 0.41, i2: 0.0020 },
  5: { t1: 0.30, i1: 0.0065, t2: 0.36, i2: 0.0045 },
  7: { t1: 0.32, i1: 0.0070, t2: 0.38, i2: 0.0040 },
  9: { t1: 0.45, i1: 0.0070, t2: 0.51, i2: 0.0040 },
  12: { t1: 0.48, i1: 0.0080, t2: 0.54, i2: 0.0040 },
};

/**
 * Devuelve { franja, incentivo } para una cuota y un % de mora dados,
 * según OBJETIVOS_INCENTIVO. Si la cuota no tiene objetivo configurado,
 * o no hay % de mora (fila "sin datos"), devuelve franja/incentivo vacíos.
 */
function calcularIncentivo(cuota, pctMora) {
  const obj = OBJETIVOS_INCENTIVO[cuota];
  if (!obj || typeof pctMora !== 'number') return { franja: '', incentivo: '' };

  if (pctMora < obj.t1) {
    return { franja: '< ' + (obj.t1 * 100) + '%', incentivo: obj.i1 };
  }
  if (pctMora <= obj.t2) {
    return { franja: (obj.t1 * 100) + ' a ' + (obj.t2 * 100) + '%', incentivo: obj.i2 };
  }
  return { franja: '> ' + (obj.t2 * 100) + '%', incentivo: 0 };
}

const MESES_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
  'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

const ENCABEZADOS_HISTORIAL = [
  'Fecha', 'MesAnalisis', 'Avance', 'CarteraTotal', 'CarteraActiva',
  'PagosAdjudicados', 'PagosAhorristas', 'TotalPagos',
  'Impagos(MoraIrregular)', 'Rescindidos', 'ResIrreImp', 'PctMora',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mora')
    .addItem('Sacar foto ahora', 'snapshotDiario')
    .addItem('Actualizar tablero', 'actualizarTablero')
    .addItem('Actualizar detalle de planes', 'actualizarDetallePlanes')
    .addItem('Instalar snapshot automático diario', 'instalarTriggerDiario')
    .addToUi();
}

/** Crea un trigger que corre snapshotDiario() todos los días (madrugada). */
function instalarTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'snapshotDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('snapshotDiario').timeBased().everyDays(1).atHour(4).create();
  SpreadsheetApp.getUi().alert('Listo: se revisa automáticamente todos los días ~4am (solo guarda foto nueva si hubo cambios en la base).');
}

/** "Septiembre 2026" para la fecha dada. */
function nombreMesAnio(date) {
  return MESES_ES[date.getMonth()] + ' ' + date.getFullYear();
}

/** Mes calendario anterior al de "date", como "Agosto 2026". */
function nombreMesAnioAnterior(date) {
  const anterior = new Date(date.getFullYear(), date.getMonth() - 1, 1);
  return nombreMesAnio(anterior);
}

/** Solo el nombre del mes anterior (sin año) al de "fechaFoto", ej. "Agosto". */
function soloMesAnterior(fechaFoto) {
  const d = (fechaFoto instanceof Date) ? fechaFoto : new Date(fechaFoto);
  const anterior = new Date(d.getFullYear(), d.getMonth() - 1, 1);
  return MESES_ES[anterior.getMonth()];
}

/**
 * Hash (huella digital) de los datos que realmente importan para la mora:
 * Avance, Estado y las columnas C2..C14 de cada fila. Ignora a propósito
 * teléfonos, vendedor, etc. — si eso cambia no cuenta como "cambio real".
 */
function calcularHashBase(datos) {
  const offsetAvance = COL_AVANCE - 1;
  const offsetEstado = COL_ESTADO - 1;
  const offsetC2 = COL_C2 - 1;

  const relevantes = datos.map(function (f) {
    return [f[offsetAvance], f[offsetEstado]].concat(f.slice(offsetC2, offsetC2 + 13));
  });

  const texto = JSON.stringify(relevantes);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ((b < 0 ? b + 256 : b)).toString(16).padStart(2, '0');
  }).join('');
}

function registrarLog(ss, mensaje) {
  let hoja = ss.getSheetByName(HOJA_LOG);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_LOG);
    hoja.getRange(1, 1, 1, 2).setValues([['FechaHora', 'Resultado']]);
  }
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  hoja.appendRow([ahora, mensaje]);
}

/**
 * Revisa BASE. Si los datos relevantes cambiaron desde la última corrida,
 * agrega una foto por cada "Avance" presente a Historial_Mora. Si no
 * cambió nada, no escribe nada en Historial_Mora (solo deja constancia en
 * Log_Snapshot).
 */
function snapshotDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(HOJA_BASE);
  if (!hojaBase) throw new Error('No encuentro la hoja "' + HOJA_BASE + '"');

  const lastRow = hojaBase.getLastRow();
  const lastCol = Math.max(COL_C2 + 13, hojaBase.getLastColumn()); // hasta C14
  const datos = hojaBase.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const hashActual = calcularHashBase(datos);
  const props = PropertiesService.getScriptProperties();
  const hashAnterior = props.getProperty(PROP_ULTIMO_HASH);

  if (hashActual === hashAnterior) {
    registrarLog(ss, 'Sin cambios en BASE (no se generó foto nueva)');
    return;
  }

  const hojaHist = obtenerOCrearHistorial(ss);
  const avances = Array.from(new Set(
    datos.map(function (f) { return f[COL_AVANCE - 1]; })
      .filter(function (v) { return typeof v === 'number' && v >= 2; })
  )).sort(function (a, b) { return a - b; });

  const hoy = new Date();
  const hoyStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const mesAnalisis = nombreMesAnio(hoy);

  // borrar fotos previas de HOY para no duplicar si se corre 2 veces el mismo día
  borrarFotosDeFecha(hojaHist, hoyStr);

  const filasNuevas = avances.map(function (avance) {
    const r = calcularPorAvance(datos, avance);
    return [
      hoyStr, mesAnalisis, avance, r.carteraTotal, r.carteraActiva,
      r.pagosAdjudicados, r.pagosAhorristas, r.totalPagos,
      r.impagos, r.rescindidos, r.resIrreImp, r.pctMora,
    ];
  });

  if (filasNuevas.length > 0) {
    hojaHist.getRange(hojaHist.getLastRow() + 1, 1, filasNuevas.length, ENCABEZADOS_HISTORIAL.length)
      .setValues(filasNuevas);
  }

  props.setProperty(PROP_ULTIMO_HASH, hashActual);
  registrarLog(ss, 'Actualizado: ' + filasNuevas.length + ' avances guardados (hubo cambios en BASE)');
}

function obtenerOCrearHistorial(ss) {
  let hoja = ss.getSheetByName(HOJA_HISTORIAL);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_HISTORIAL);
    hoja.getRange(1, 1, 1, ENCABEZADOS_HISTORIAL.length).setValues([ENCABEZADOS_HISTORIAL]);
  }
  return hoja;
}

function borrarFotosDeFecha(hojaHist, fechaStr) {
  const lastRow = hojaHist.getLastRow();
  if (lastRow < 2) return;
  const valores = hojaHist.getRange(2, 1, lastRow - 1, 1).getValues();
  // borrar de abajo hacia arriba para no romper los índices
  for (let i = valores.length - 1; i >= 0; i--) {
    if (valores[i][0] === fechaStr) {
      hojaHist.deleteRow(i + 2);
    }
  }
}

/**
 * Calcula los agregados de mora para un valor de "Avance" puntual,
 * a partir de la matriz completa de BASE (ya leída con getValues()).
 */
function calcularPorAvance(datos, avance) {
  const ncols = avance - 2; // C2..C(avance-1)
  const offsetC2 = COL_C2 - 1;
  const offsetAvance = COL_AVANCE - 1;
  const offsetEstado = COL_ESTADO - 1;

  let carteraTotal = 0, rescindidos = 0, irregulares = 0;
  let pagosAdjudicados = 0, pagosAhorristas = 0;

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[offsetAvance] !== avance) continue;
    carteraTotal++;

    const rango = fila.slice(offsetC2, offsetC2 + ncols);
    const tieneR = rango.some(function (v) { return v === 'R'; });
    const tieneI = rango.some(function (v) { return v === 'I'; });
    const todoP = rango.every(function (v) { return v === 'P'; });

    if (tieneR) rescindidos++;
    if (tieneI) irregulares++;
    if (todoP) {
      const estado = fila[offsetEstado];
      if (estado === 'Adjudicado') pagosAdjudicados++;
      else if (estado === 'Ahorrista') pagosAhorristas++;
    }
  }

  const carteraActiva = carteraTotal - rescindidos;
  const totalPagos = pagosAdjudicados + pagosAhorristas;
  const resIrreImp = rescindidos + irregulares;
  const pctMora = carteraTotal > 0 ? resIrreImp / carteraTotal : 0;

  return {
    carteraTotal: carteraTotal, carteraActiva: carteraActiva,
    pagosAdjudicados: pagosAdjudicados, pagosAhorristas: pagosAhorristas,
    totalPagos: totalPagos, impagos: irregulares, rescindidos: rescindidos,
    resIrreImp: resIrreImp, pctMora: pctMora,
  };
}

/**
 * Refresca la hoja "Tablero": una fila por cuota (3,5,7,9,12) con la última
 * foto disponible en el historial para el avance correspondiente
 * (avance = cuota + 1).
 */
function actualizarTablero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHist = ss.getSheetByName(HOJA_HISTORIAL);
  if (!hojaHist || hojaHist.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Todavía no hay fotos en "' + HOJA_HISTORIAL + '". Corré primero "Sacar foto ahora".');
    return;
  }

  let hojaTab = ss.getSheetByName(HOJA_TABLERO);
  if (!hojaTab) hojaTab = ss.insertSheet(HOJA_TABLERO);
  hojaTab.clear();

  const encabezados = ['Cuota', 'Fecha foto', 'Período real', 'Avance',
    'Cartera total', 'Cartera activa', 'Pagos Adjudicados', 'Pagos Ahorristas',
    'TOTAL PAGOS', 'Impagos', 'Rescindidos', 'res+irre+imp', '% de mora',
    'Franja objetivo', '% Incentivo'];
  hojaTab.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  const hist = hojaHist.getRange(2, 1, hojaHist.getLastRow() - 1, ENCABEZADOS_HISTORIAL.length).getValues();

  const filas = CUOTAS_TABLERO.map(function (cuota) {
    const avanceObjetivo = cuota + 1;
    // última foto (fecha más reciente) para ese avance
    const candidatas = hist.filter(function (f) { return f[2] === avanceObjetivo; });
    if (candidatas.length === 0) {
      return [cuota, '(sin datos)', '', avanceObjetivo, '', '', '', '', '', '', '', '', '', '', ''];
    }
    candidatas.sort(function (a, b) { return new Date(b[0]) - new Date(a[0]); });
    const r = candidatas[0];
    // r: [Fecha, MesAnalisis, Avance, CarteraTotal, CarteraActiva,
    //     PagosAdj, PagosAhorr, TotalPagos, Impagos, Rescindidos, ResIrreImp, PctMora]
    // "Período real" se muestra como "Cuota N MES" (ej. "Cuota 3 AGOSTO") en vez
    // de la fecha cruda, para que no se confunda con el valor de "Avance".
    const etiquetaPeriodo = 'Cuota ' + cuota + ' ' + soloMesAnterior(r[0]).toUpperCase();
    const pctMora = r[11];
    const inc = calcularIncentivo(cuota, pctMora);
    return [cuota, r[0], etiquetaPeriodo, r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], pctMora,
      inc.franja, inc.incentivo];
  });

  hojaTab.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
  hojaTab.getRange(2, 13, filas.length, 1).setNumberFormat('0.00%');
  hojaTab.getRange(2, 15, filas.length, 1).setNumberFormat('0.00%');
}

/**
 * Clasifica UNA fila de BASE según la misma regla de "mes atrasado" que usa
 * calcularPorAvance (rango C2..C(avance-1)), pero devolviendo una sola
 * etiqueta por plan en vez de un conteo agregado:
 * - 'Rescindido': hay al menos una 'R' en el rango.
 * - 'Mora irregular': no hay 'R', pero hay al menos una 'I' en el rango.
 * - 'Pagado al día': todas las columnas del rango son 'P'.
 * - 'Sin cuotas para analizar': avance=2 (cuota 1, todavía no hay rango que
 *   mirar — evita el caso trivial de un rango vacío "cumpliendo" todoP).
 * - 'Sin clasificar': combinación rara (huecos en el rango, etc.), para no
 *   perder la fila silenciosamente.
 * El Estado (Ahorrista/Adjudicado/Rescindido/Renunciado/Cancelado) es un
 * dato aparte de BASE y se devuelve tal cual — un plan puede figurar como
 * "Renunciado" en Estado y "Pagado al día" en esta clasificación a la vez.
 */
function clasificarFila(fila) {
  const offsetAvance = COL_AVANCE - 1;
  const offsetC2 = COL_C2 - 1;
  const avance = fila[offsetAvance];

  if (typeof avance !== 'number' || avance < 2) return null;

  const ncols = avance - 2; // C2..C(avance-1)
  if (ncols <= 0) return 'Sin cuotas para analizar';

  const rango = fila.slice(offsetC2, offsetC2 + ncols);
  const tieneR = rango.some(function (v) { return v === 'R'; });
  const tieneI = rango.some(function (v) { return v === 'I'; });
  const todoP = rango.length > 0 && rango.every(function (v) { return v === 'P'; });

  if (tieneR) return 'Rescindido';
  if (tieneI) return 'Mora irregular';
  if (todoP) return 'Pagado al día';
  return 'Sin clasificar';
}

/**
 * Refresca la hoja "Detalle_Planes": un listado, un plan por fila, leído en
 * vivo de BASE (no del historial — acá interesa quiénes son HOY, no una
 * foto vieja), con la clasificación de mora de cada uno y un filtro nativo
 * de Sheets ya armado para poder filtrar por Avance y por Clasificación/
 * Estado (Mora irregular, Rescindido, Renunciado, etc.).
 */
function actualizarDetallePlanes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(HOJA_BASE);
  if (!hojaBase) throw new Error('No encuentro la hoja "' + HOJA_BASE + '"');

  const lastRow = hojaBase.getLastRow();
  const lastCol = Math.max(COL_C2 + 13, hojaBase.getLastColumn());
  const datos = hojaBase.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const offsetSolicitud = COL_SOLICITUD - 1;
  const offsetGrupo = COL_GRUPO - 1;
  const offsetCliente = COL_CLIENTE - 1;
  const offsetTelefono = COL_TELEFONO - 1;
  const offsetVendedor = COL_VENDEDOR - 1;
  const offsetSupervisor = COL_SUPERVISOR - 1;
  const offsetAvance = COL_AVANCE - 1;
  const offsetEstado = COL_ESTADO - 1;

  const encabezados = ['Solicitud', 'Grupo', 'Cliente', 'Teléfono', 'Vendedor',
    'Supervisor', 'Avance', 'Cuota analizada', 'Estado', 'Clasificación mora'];

  const filas = [];
  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    const avance = fila[offsetAvance];
    const clasificacion = clasificarFila(fila);
    if (clasificacion === null) continue; // sin Avance numérico válido

    filas.push([
      fila[offsetSolicitud], fila[offsetGrupo], fila[offsetCliente],
      fila[offsetTelefono], fila[offsetVendedor], fila[offsetSupervisor],
      avance, avance - 1, fila[offsetEstado], clasificacion,
    ]);
  }

  let hojaDet = ss.getSheetByName(HOJA_DETALLE);
  if (!hojaDet) hojaDet = ss.insertSheet(HOJA_DETALLE);
  hojaDet.clear();

  hojaDet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  if (filas.length > 0) {
    hojaDet.getRange(2, 1, filas.length, encabezados.length).setValues(filas);
  }
  hojaDet.setFrozenRows(1);

  const rangoCompleto = hojaDet.getRange(1, 1, filas.length + 1, encabezados.length);
  const filtroExistente = hojaDet.getFilter();
  if (filtroExistente) filtroExistente.remove();
  rangoCompleto.createFilter();
}

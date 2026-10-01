/**
 * LECTOR DE SEÑALES COMERCIALES — alerta de archivos nuevos en la carpeta
 *
 * Revisa la carpeta de Drive donde se suben las señales comerciales (y
 * cartas de objetivos, tablas de cuotas, etc.), registra cada archivo en la
 * hoja "SEÑALES" con número, marca, tipo y mes detectados del título, y manda
 * un mail cuando aparece algo nuevo. La hoja "FALTANTES" lista los números de
 * señal que no están en la carpeta (la numeración es correlativa entre Fiat,
 * Jeep y RAM, así que no todos los faltantes son necesariamente de Fiat Plan).
 *
 * Instalación: ver SENALES.md.
 */

// --- CONFIGURACIÓN ---
const CARPETA_SENALES = '103O9LnAq7tljmQpwdFl7dUlwkecehcd_';
const HOJA_SENALES = 'SEÑALES';
const HOJA_FALTANTES = 'FALTANTES';
// Vacío = el mail de quien instala el trigger. Se pueden poner varios separados por coma.
const MAIL_AVISO = '';
// Para FALTANTES: no mirar números más viejos que este (la señal madre 1065 es de 2021).
const NUMERO_DESDE = 1500;

const ENCABEZADOS = ['Detectada', 'Fecha archivo', 'Nº', 'Marca', 'Tipo', 'Mes ref.',
  'Título', 'Link', 'Subido por', 'Estado', 'Notas', 'ID'];
const ESTADOS = ['Pendiente', 'Leída', 'Cargada'];

const MESES_SENAL = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const MESES_CORTOS = { ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5, JUL: 6, AGO: 7, SEP: 8, SEPT: 8, OCT: 9, NOV: 10, DIC: 11 };

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Señales')
    .addItem('Revisar carpeta ahora', 'revisarCarpeta')
    .addItem('Instalar revisión automática (cada hora)', 'instalarTriggerSenales')
    .addToUi();
}

function instalarTriggerSenales() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'revisarCarpeta')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisarCarpeta').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Listo: la carpeta se revisa sola cada hora y avisa por mail si hay archivos nuevos.');
}

/** Registra los archivos nuevos de la carpeta (y subcarpetas) y avisa por mail. */
function revisarCarpeta() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaSenales(ss);
  const ultima = sh.getLastRow();
  const conocidos = new Set(ultima > 1
    ? sh.getRange(2, ENCABEZADOS.indexOf('ID') + 1, ultima - 1, 1).getValues().map((f) => f[0])
    : []);
  const primeraVez = conocidos.size === 0;

  const nuevos = [];
  recorrerCarpeta(DriveApp.getFolderById(CARPETA_SENALES), (archivo) => {
    if (conocidos.has(archivo.getId())) return;
    nuevos.push(filaDesdeArchivo(archivo));
  });
  if (!nuevos.length) return;

  nuevos.sort((a, b) => a[1] - b[1]);
  sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, ENCABEZADOS.length).setValues(nuevos);
  sh.getRange(2, 1, sh.getLastRow() - 1, ENCABEZADOS.length).sort([{ column: 2, ascending: false }]);
  actualizarFaltantes(ss);
  avisarPorMail(ss, nuevos, primeraVez);
}

function recorrerCarpeta(carpeta, fn) {
  const archivos = carpeta.getFiles();
  while (archivos.hasNext()) fn(archivos.next());
  const sub = carpeta.getFolders();
  while (sub.hasNext()) recorrerCarpeta(sub.next(), fn);
}

function filaDesdeArchivo(archivo) {
  const titulo = archivo.getName();
  const d = interpretarTitulo(titulo, archivo.getDateCreated());
  let autor = '';
  try { autor = archivo.getOwner() ? archivo.getOwner().getEmail() : ''; } catch (e) { /* unidades compartidas */ }
  return [new Date(), archivo.getDateCreated(), d.numero, d.marca, d.tipo, d.mes,
    titulo, archivo.getUrl(), autor, 'Pendiente', '', archivo.getId()];
}

/** Saca número, marca, tipo y mes del nombre del archivo. */
function interpretarTitulo(titulo, fechaArchivo) {
  const t = titulo.toUpperCase();
  const num = t.match(/N\s*[°º]?\s*(\d{3,4})\b/);
  const marca = /\bJEEP\b/.test(t) && /\bRAM\b/.test(t) ? 'JEEP / RAM'
    : /\bJEEP\b/.test(t) ? 'JEEP' : /\bRAM\b/.test(t) ? 'RAM' : /\bFIAT\b/.test(t) ? 'FIAT' : '';
  let tipo = 'Otro';
  if (/CARTA DE OBJETIVOS|OBJETIVOS/.test(t)) tipo = 'Carta de objetivos';
  else if (/SE[ÑN]AL/.test(t)) tipo = 'Señal comercial';
  else if (/CUOTAS/.test(t)) tipo = 'Tabla de cuotas';
  else if (/OPCIONALES/.test(t)) tipo = 'Opcionales';
  else if (/CAMBIO DE MODELO/.test(t)) tipo = 'Lista cambio de modelo';

  let mes = '';
  const anio = (t.match(/\b20(\d{2})\b/) || t.match(/'(\d{2})\b/) ||
    t.match(/\b(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[A-Z]*\s+(\d{2})\b/) || [])[1] ||
    (fechaArchivo ? String(fechaArchivo.getFullYear() % 100) : '');
  const iLargo = MESES_SENAL.findIndex((m) => new RegExp(`\\b${m}\\b`).test(t));
  if (iLargo >= 0) {
    mes = MESES_SENAL[iLargo] + (anio ? ` ${anio}` : '');
  } else {
    const corto = Object.keys(MESES_CORTOS).find((m) => new RegExp(`\\b${m}\\b`).test(t));
    if (corto) mes = MESES_SENAL[MESES_CORTOS[corto]] + (anio ? ` ${anio}` : '');
  }
  return { numero: num ? Number(num[1]) : '', marca, tipo, mes };
}

function hojaSenales(ss) {
  let sh = ss.getSheetByName(HOJA_SENALES);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA_SENALES, 0);
  sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS])
    .setFontWeight('bold').setBackground('#430000').setFontColor('white');
  sh.setFrozenRows(1);
  const filas = sh.getMaxRows() - 1;
  const cEstado = ENCABEZADOS.indexOf('Estado') + 1;
  sh.getRange(2, cEstado, filas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  const rEstado = sh.getRange(2, cEstado, filas, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pendiente')
      .setBackground('#fff2cc').setRanges([rEstado]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cargada')
      .setBackground('#d9ead3').setRanges([rEstado]).build(),
  ]);
  sh.getRange('A:B').setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidth(ENCABEZADOS.indexOf('Título') + 1, 420);
  sh.hideColumns(ENCABEZADOS.indexOf('ID') + 1);
  return sh;
}

/** Números de señal entre NUMERO_DESDE y el último que no están en la carpeta. */
function actualizarFaltantes(ss) {
  const sh = ss.getSheetByName(HOJA_SENALES);
  const cNum = ENCABEZADOS.indexOf('Nº') + 1;
  const nums = sh.getRange(2, cNum, sh.getLastRow() - 1, 1).getValues()
    .map((f) => f[0]).filter((n) => typeof n === 'number' && n >= NUMERO_DESDE);
  let f = ss.getSheetByName(HOJA_FALTANTES);
  if (!f) f = ss.insertSheet(HOJA_FALTANTES);
  f.clear();
  f.getRange(1, 1, 1, 2).setValues([['Nº faltante', 'Nota']])
    .setFontWeight('bold').setBackground('#430000').setFontColor('white');
  if (!nums.length) return;
  const presentes = new Set(nums);
  const faltan = [];
  for (let n = Math.min(...nums); n <= Math.max(...nums); n++) {
    if (!presentes.has(n)) faltan.push([n, '']);
  }
  if (faltan.length) f.getRange(2, 1, faltan.length, 2).setValues(faltan);
}

function avisarPorMail(ss, nuevos, primeraVez) {
  const destino = MAIL_AVISO || Session.getEffectiveUser().getEmail();
  if (!destino) return;
  const iTit = ENCABEZADOS.indexOf('Título');
  const iLink = ENCABEZADOS.indexOf('Link');
  const iTipo = ENCABEZADOS.indexOf('Tipo');
  const asunto = primeraVez
    ? `Señales: se registraron ${nuevos.length} archivos de la carpeta`
    : `Señales: ${nuevos.length} archivo(s) nuevo(s) en la carpeta`;
  const items = nuevos.map((f) =>
    `<li><b>${f[iTipo]}</b> — <a href="${f[iLink]}">${f[iTit]}</a></li>`).join('');
  MailApp.sendEmail({
    to: destino,
    subject: asunto,
    htmlBody: `<p>${asunto}:</p><ul>${items}</ul>` +
      `<p><a href="${ss.getUrl()}">Abrir el registro de señales</a></p>`,
  });
}

/**
 * LECTOR DE SEÑALES COMERCIALES — alerta y lectura de la carpeta
 *
 * Revisa la carpeta de Drive donde se suben las señales comerciales y las
 * cartas de objetivos, registra cada archivo en la hoja "SEÑALES" y manda un
 * mail cuando aparece algo nuevo. Además LEE EL CONTENIDO (PDF / imagen con
 * OCR, PowerPoint):
 * - Carta de objetivos: carga sola los objetivos en la hoja OBJETIVOS (ver
 *   cargarCartaObjetivos en objetivos.gs), agrega indicadores nuevos a
 *   LISTAS_OBJ y la marca "Cargada".
 * - Señal comercial: detecta si afecta objetivos / incentivos / categoría /
 *   flujo / mora y guarda un extracto. Los % de incentivo se cargan a mano
 *   (cada señal tiene un formato distinto y un número mal leído afecta la
 *   liquidación).
 * La hoja "FALTANTES" lista los números de señal que no están en la carpeta
 * (la numeración es correlativa entre Fiat, Jeep y RAM).
 *
 * Requiere el servicio avanzado "Drive API" activado (Apps Script > Servicios).
 * Va en el archivo de objetivos, junto con objetivos.gs. Instalación: ver SENALES.md.
 */

// --- CONFIGURACIÓN ---
const CARPETA_SENALES = '103O9LnAq7tljmQpwdFl7dUlwkecehcd_';
const HOJA_SENALES = 'SEÑALES';
const HOJA_FALTANTES = 'FALTANTES';
// Destinatario del aviso. Se pueden poner varios separados por coma.
const MAIL_AVISO = 'gortiz@grupoantun.com.ar';
// Para FALTANTES: no mirar números más viejos que este (la señal madre 1065 es de 2021).
const NUMERO_DESDE = 1500;
// Cuántos archivos leer por corrida (Apps Script corta a los 6 minutos).
const MAX_LECTURAS_POR_CORRIDA = 12;

// 'Afecta' y 'Extracto' van al final para no correr las columnas de hojas ya creadas.
const ENCABEZADOS = ['Detectada', 'Fecha archivo', 'Nº', 'Marca', 'Tipo', 'Mes ref.',
  'Título', 'Link', 'Subido por', 'Estado', 'Notas', 'ID', 'Afecta', 'Extracto'];
const ESTADOS = ['Pendiente', 'Leída', 'Cargada'];
const TIPOS_SIN_LECTURA = ['Tabla de cuotas', 'Opcionales', 'Lista cambio de modelo'];

// Temas que se buscan en el texto de las señales.
const TEMAS = [
  { tema: 'Objetivos', re: /OBJETIVO/ },
  { tema: 'Incentivos', re: /BONUS|INCENTIVO/ },
  { tema: 'Categoría', re: /CATEGOR[IÍ]A/ },
  { tema: 'Flujo', re: /FLUJO/ },
  { tema: 'Mora / permanencia', re: /MORA|PERMANENCIA/ },
  { tema: 'Pedidos', re: /PEDIDO/ },
];

const MESES_SENAL = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const MESES_CORTOS = { ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5, JUL: 6, AGO: 7, SEP: 8, SEPT: 8, OCT: 9, NOV: 10, DIC: 11 };

// El menú lo arma onOpen() en objetivos.gs (un solo onOpen por proyecto).

function instalarTriggerSenales() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'revisarCarpeta')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisarCarpeta').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Listo: la carpeta se revisa sola cada hora y avisa por mail si hay archivos nuevos.');
}

/**
 * Registra los archivos nuevos de la carpeta (y subcarpetas), lee el contenido
 * de los que todavía no se leyeron y avisa por mail.
 */
function revisarCarpeta() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaSenales(ss);
  const col = columnasSenales(sh);
  const ultima = sh.getLastRow();
  const conocidos = new Set(ultima > 1
    ? sh.getRange(2, col.ID, ultima - 1, 1).getValues().map((f) => f[0])
    : []);
  const primeraVez = conocidos.size === 0;

  const nuevos = [];
  recorrerCarpeta(DriveApp.getFolderById(CARPETA_SENALES), (archivo) => {
    if (conocidos.has(archivo.getId())) return;
    nuevos.push(filaDesdeArchivo(archivo));
  });
  if (nuevos.length) {
    nuevos.sort((a, b) => a[1] - b[1]);
    sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, ENCABEZADOS.length).setValues(nuevos);
    sh.getRange(2, 1, sh.getLastRow() - 1, ENCABEZADOS.length).sort([{ column: 2, ascending: false }]);
    actualizarFaltantes(ss);
  }

  const leidos = leerPendientes(ss, sh, col);
  const idsNuevos = new Set(nuevos.map((f) => f[ENCABEZADOS.indexOf('ID')]));
  // En la primera corrida no se avisa por cada señal vieja: un solo mail con el total.
  if (nuevos.length || leidos.some((l) => idsNuevos.has(l.id))) {
    avisarPorMail(ss, sh, col, idsNuevos, leidos, primeraVez);
  }
}

/** Para usar desde el menú: lee el contenido de lo ya registrado, sin esperar archivos nuevos. */
function leerSenalesRegistradas() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaSenales(ss);
  const col = columnasSenales(sh);
  const leidos = leerPendientes(ss, sh, col);
  const quedan = filasSinLeer(sh, col).length;
  const cartas = leidos.filter((l) => l.carta);
  SpreadsheetApp.getUi().alert(
    `Leídos: ${leidos.length}` +
    (cartas.length ? `\nCartas de objetivos cargadas: ${cartas.map((c) => c.carta.mes).join(', ')}` : '') +
    (quedan ? `\nQuedan ${quedan} por leer: volvé a correrlo (o esperá a la revisión automática).` : ''));
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
    titulo, archivo.getUrl(), autor, 'Pendiente', '', archivo.getId(), '', ''];
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

// ---------------------------------------------------------------------------
// Lectura de contenido
// ---------------------------------------------------------------------------

/** Filas (número de fila) que todavía no se leyeron y se pueden leer. */
function filasSinLeer(sh, col) {
  const ultima = sh.getLastRow();
  if (ultima < 2) return [];
  const datos = sh.getRange(2, 1, ultima - 1, sh.getLastColumn()).getValues();
  const filas = [];
  datos.forEach((f, i) => {
    if (f[col.Afecta - 1] !== '') return;
    if (TIPOS_SIN_LECTURA.indexOf(f[col.Tipo - 1]) >= 0) return;
    filas.push(i + 2);
  });
  return filas;
}

/** Lee hasta MAX_LECTURAS_POR_CORRIDA archivos pendientes. Devuelve lo leído. */
function leerPendientes(ss, sh, col) {
  const leidos = [];
  if (typeof Drive === 'undefined') {
    Logger.log('Falta activar el servicio avanzado "Drive API": no se lee el contenido.');
    return leidos;
  }
  const inicio = Date.now();
  filasSinLeer(sh, col).slice(0, MAX_LECTURAS_POR_CORRIDA).forEach((fila) => {
    if (Date.now() - inicio > 4.5 * 60 * 1000) return; // margen antes del corte de 6 min
    const get = (c) => sh.getRange(fila, col[c]).getValue();
    const set = (c, v) => sh.getRange(fila, col[c]).setValue(v);
    const id = get('ID');
    const res = { id, fila, titulo: get('Título'), link: get('Link'), carta: null, temas: [] };
    let texto = '';
    try {
      texto = textoDeArchivo(DriveApp.getFileById(id));
    } catch (e) {
      set('Afecta', '-');
      set('Extracto', `No se pudo leer: ${e.message}`);
      leidos.push(res);
      return;
    }
    if (!texto.trim()) {
      set('Afecta', '-');
      set('Extracto', 'Sin texto legible');
      leidos.push(res);
      return;
    }

    const carta = interpretarCarta(texto);
    if (carta) {
      res.carta = carta;
      const resumen = cargarCartaObjetivos(ss, carta, res.link);
      set('Tipo', 'Carta de objetivos');
      if (!get('Marca') && carta.marca) set('Marca', carta.marca);
      set('Mes ref.', carta.mes);
      set('Afecta', 'Objetivos');
      set('Extracto', carta.indicadores.map((x) => `${x.indicador} ${x.objetivo}`).join(' · ') +
        (carta.categoria ? ` · Categoría ${carta.categoria}` : '') +
        (carta.fechaFlujo ? ` · Flujo ${carta.fechaFlujo} al ${Math.round(carta.pctFlujo * 100)}%` : ''));
      set('Estado', 'Cargada');
      set('Notas', resumen);
      res.temas = ['Objetivos'];
      res.resumen = resumen;
    } else {
      const t = texto.toUpperCase();
      res.temas = TEMAS.filter((x) => x.re.test(t)).map((x) => x.tema);
      set('Afecta', res.temas.length ? res.temas.join(', ') : '-');
      set('Extracto', extracto(texto));
    }
    leidos.push(res);
  });
  return leidos;
}

/** Texto de un PDF / imagen (conversión a Doc con OCR) o de un PowerPoint (conversión a Slides). */
function textoDeArchivo(archivo) {
  const mime = archivo.getMimeType();
  const esPpt = /presentation|powerpoint/.test(mime);
  const esLegible = esPpt || mime === MimeType.PDF || /^image\//.test(mime) || mime === MimeType.GOOGLE_DOCS;
  if (!esLegible) return '';
  if (mime === MimeType.GOOGLE_DOCS) return DocumentApp.openById(archivo.getId()).getBody().getText();

  const destino = esPpt ? MimeType.GOOGLE_SLIDES : MimeType.GOOGLE_DOCS;
  const nombre = `tmp lectura señal - ${archivo.getName()}`;
  let idTmp;
  if (Drive.Files.create) { // Drive API v3
    idTmp = Drive.Files.create({ name: nombre, mimeType: destino }, archivo.getBlob(),
      { ocrLanguage: 'es' }).id;
  } else { // Drive API v2
    idTmp = Drive.Files.insert({ title: nombre, mimeType: destino }, archivo.getBlob(),
      { ocr: true, ocrLanguage: 'es' }).id;
  }
  try {
    if (!esPpt) return DocumentApp.openById(idTmp).getBody().getText();
    const partes = [];
    SlidesApp.openById(idTmp).getSlides().forEach((s) => s.getPageElements().forEach((el) => {
      const tipo = el.getPageElementType();
      if (tipo === SlidesApp.PageElementType.SHAPE) partes.push(el.asShape().getText().asString());
      if (tipo === SlidesApp.PageElementType.TABLE) {
        const tabla = el.asTable();
        for (let r = 0; r < tabla.getNumRows(); r++) {
          const fila = [];
          for (let c = 0; c < tabla.getNumColumns(); c++) fila.push(tabla.getCell(r, c).getText().asString().trim());
          partes.push(fila.join(' '));
        }
      }
    }));
    return partes.join('\n');
  } finally {
    DriveApp.getFileById(idTmp).setTrashed(true);
  }
}

/**
 * Resumen legible de una señal: los temas del "REF:" (1- Liquidación flujo,
 * 2.a Bonus cumplimiento...) en una lista corta + el primer dato con % o $.
 * Limpia lo que ensucia el OCR: encabezados con letras separadas
 * ("T O D O S L O S ..."), mails, "®" y el texto de cortesía.
 */
function extracto(texto) {
  const t = limpiarTexto(texto);
  const corte = /\b(Nos dirigimos|Informamos|Por la presente|Se informa|Estimad[oa]s?|Buenos Aires|Les recordamos|Les informamos)\b/i;

  // Temas: después de "REF" si existe; si no, desde el principio.
  const iRef = t.search(/\bREF\b\s*[:.]?/i);
  let cabecera = iRef >= 0 ? t.slice(iRef).replace(/^REF\s*[:.]?\s*/i, '') : t;
  const iCorte = cabecera.search(corte);
  cabecera = cabecera.slice(0, iCorte > 0 ? iCorte : 300).slice(0, 300);
  const vistos = new Set();
  const temas = cabecera
    .split(/\s(?=\d{1,2}(?:\.?[a-zA-Z])?\s*[-–.)]?\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ]{2})/)
    .map((x) => x.replace(/^\d{1,2}(?:\.?[a-zA-Z])?\s*[-–.)]?\s*/, '').replace(/[\s\-–.:,]+$/, '').trim())
    .filter((x) => x.length > 4)
    .filter((x) => {
      // Clave por las primeras letras: "Actualización extra bonus pedido(s)..." cuenta como el mismo tema.
      const k = x.toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g, '').slice(0, 20);
      if (vistos.has(k)) return false;
      vistos.add(k);
      return true;
    })
    .slice(0, 6)
    .map((x) => enOracion(x.length > 90 ? `${x.slice(0, 88).replace(/\s\S*$/, '')}…` : x));

  // Dato clave: primera oración del cuerpo con un % o un $.
  const iCuerpo = t.search(corte);
  const cuerpo = iCuerpo >= 0 ? t.slice(iCuerpo) : t;
  let dato = (cuerpo.split(/(?<=[.;])\s+/).find((o) => /%|\$/.test(o) && o.length < 300) || '')
    .replace(/^(Nos dirigimos a ustedes para informarles que|Por la presente le?s? informamos que|Se informa que|Les informamos que|Informamos)\s*/i, '')
    .trim();
  if (dato) dato = dato.charAt(0).toUpperCase() + dato.slice(1);
  const k = (x) => x.toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g, '').slice(0, 30);
  if (temas.some((x) => k(x) === k(dato))) dato = '';

  let res = temas.length ? temas.map((x) => `• ${x}`).join('  ') : enOracion(cuerpo.slice(0, 200));
  if (dato) res += `  |  Dato: ${enOracion(dato)}`;
  return res.slice(0, 400);
}

function limpiarTexto(texto) {
  return texto
    .replace(/[®•▪●]/g, ' ')
    .replace(/(?:\b\p{L}\s+){4,}\p{L}\b/gu, ' ')       // "T O D O S L O S ..." del OCR
    .replace(/(?:\s(?:a|y))?\s\S+@\S+/g, ' ')            // mails (y el "a ... y a ..." que los rodea)
    .replace(/\s+(?:y|a)\s*(?=,)/g, '')
    .replace(/A todos los concesionarios de la red\s+\w+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Siglas que quedan en mayúscula al pasar un texto a "oración".
const SIGLAS = ['ABC', 'NPS', 'REU', 'TP', 'RAM', 'JEEP', 'FIAT', 'IVA', 'CC', 'CE', 'SGA', 'CTA', 'JUL', 'SEPT', 'AGO', 'OCT', 'DIC'];

/** "BONUS CUMPLIMIENTO PEDIDOS TOTALES" -> "Bonus cumplimiento pedidos totales" (solo si viene todo en mayúsculas). */
function enOracion(x) {
  const letras = x.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g, '');
  if (!letras || letras !== letras.toUpperCase()) return x;
  const lower = x.toLowerCase().replace(/\S+/g, (w) => {
    const W = w.toUpperCase();
    const limpio = W.replace(/[^A-ZÁÉÍÓÚÑ0-9]/g, '');
    return SIGLAS.indexOf(limpio) >= 0 || /\d/.test(w) ? W : w;
  });
  return lower.charAt(0).toUpperCase() + lower.slice(1);
}

/**
 * Vuelve a leer las señales (no las cartas) para rehacer Afecta y Extracto con
 * el formato actual. Lee de a MAX_LECTURAS_POR_CORRIDA: repetir hasta terminar.
 */
function rehacerExtractos() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaSenales(ss);
  const col = columnasSenales(sh);
  const props = PropertiesService.getDocumentProperties();
  if (!props.getProperty('REHACER_EXTRACTOS')) {
    const ultima = sh.getLastRow();
    if (ultima > 1) {
      const tipos = sh.getRange(2, col.Tipo, ultima - 1, 1).getValues();
      tipos.forEach((f, i) => {
        if (f[0] !== 'Carta de objetivos') sh.getRange(i + 2, col.Afecta).setValue('');
      });
    }
    props.setProperty('REHACER_EXTRACTOS', 'en curso');
  }
  leerPendientes(ss, sh, col);
  const quedan = filasSinLeer(sh, col).length;
  if (!quedan) props.deleteProperty('REHACER_EXTRACTOS');
  SpreadsheetApp.getUi().alert(quedan
    ? `Quedan ${quedan} por rehacer: volvé a correrlo.`
    : 'Listo: extractos rehechos.');
}

/**
 * Si el texto es una carta de objetivos, devuelve sus datos; si no, null.
 * Formato visto (Carta Nº 09/2026): fecha arriba, "CARTA DE OBJETIVOS Nº 09/2026",
 * concesionario, "objetivos definidos para el mes de SEPTIEMBRE 2026",
 * "CATEGORÍA A", una línea por indicador con su número y "El flujo es el día
 * 23/09/2026 ... el 65% o más".
 */
function interpretarCarta(texto) {
  const t = texto.replace(/\r/g, '');
  const T = t.toUpperCase();
  if (!/CARTA\s+DE\s+OBJETIVOS/.test(T)) return null;

  const mMes = T.match(/MES\s+DE\s+([A-ZÁÉÍÓÚ]+)\s*(?:DE\s+)?(\d{4})/);
  if (!mMes || MESES_SENAL.indexOf(mMes[1]) < 0) return null;
  const mes = `${mMes[1]} ${mMes[2].slice(2)}`;

  const nCarta = (T.match(/OBJETIVOS\s*N\s*[º°O.]?\s*(\d{1,2}\s*\/\s*\d{4})/) || [])[1];
  const mFlujo = T.match(/FLUJO[^0-9]{0,40}(\d{2}\/\d{2}\/\d{4})/);
  const fechas = (T.match(/\d{2}\/\d{2}\/\d{4}/g) || []);
  const fechaCarta = fechas.find((f) => !mFlujo || f !== mFlujo[1]) || '';
  const mPct = T.match(/(\d{1,3})\s*%\s*O\s+M[AÁ]S/);
  const conc = (t.match(/Concesionario:?\s*\n?\s*([A-ZÁÉÍÓÚÑ0-9 .&]+?S\.?\s?A\.?)/i) || [])[1];
  const marca = /\bJEEP\b/.test(T) ? 'JEEP' : /\bRAM\b/.test(T) ? 'RAM' : /\bFIAT\b/.test(T) ? 'FIAT' : '';

  // Bloque de objetivos: entre "objetivos definidos..." y "El flujo" (o el saludo).
  const ini = T.search(/OBJETIVOS\s+DEFINIDOS/);
  let fin = T.search(/EL\s+FLUJO/);
  if (fin < 0) fin = T.search(/SALUDAMOS/);
  const bloque = T.slice(ini < 0 ? 0 : ini, fin < 0 ? T.length : fin)
    .replace(/^[^\n]*\n/, ''); // saca la línea "objetivos definidos para el mes de ..."

  const categoria = (bloque.match(/CATEGOR[IÍ]A\s*:?\s*\n?\s*([ABC])\b/) || [])[1] || '';
  const sinCategoria = bloque.replace(/CATEGOR[IÍ]A\s*:?\s*\n?\s*[ABC]\b/, '\n');

  // 1) "SUSCRIPCIONES   146" en la misma línea.
  const indicadores = [];
  const reLinea = /^[ \t]*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .]{3,}?)[ \t]*:?[ \t]+(\d{1,6})[ \t]*$/gm;
  let m;
  while ((m = reLinea.exec(sinCategoria)) !== null) {
    indicadores.push({ indicador: normalizarIndicador(m[1]), objetivo: Number(m[2]) });
  }
  // 2) Si el OCR separó la tabla: primero los nombres y después los números, en el mismo orden.
  if (!indicadores.length) {
    const lineas = sinCategoria.split('\n').map((l) => l.trim()).filter(Boolean);
    const nombres = lineas.filter((l) => /^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .]{3,}$/.test(l));
    const numeros = lineas.filter((l) => /^\d{1,6}$/.test(l)).map(Number);
    if (nombres.length && nombres.length === numeros.length) {
      nombres.forEach((n, i) => indicadores.push({ indicador: normalizarIndicador(n), objetivo: numeros[i] }));
    }
  }
  if (!indicadores.length) return null;

  return {
    mes, marca, categoria, indicadores,
    concesionario: conc ? conc.replace(/\s+/g, ' ').trim().toUpperCase() : '',
    nCarta: nCarta ? nCarta.replace(/\s/g, '') : '',
    fechaCarta,
    fechaFlujo: mFlujo ? mFlujo[1] : '',
    pctFlujo: mPct ? Number(mPct[1]) / 100 : '',
  };
}

function normalizarIndicador(nombre) {
  return nombre.replace(/\s+/g, ' ').replace(/[ .:]+$/, '').trim().toUpperCase();
}

// ---------------------------------------------------------------------------
// Hojas
// ---------------------------------------------------------------------------

function hojaSenales(ss) {
  let sh = ss.getSheetByName(HOJA_SENALES);
  if (!sh) {
    sh = ss.insertSheet(HOJA_SENALES, 0);
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
  }
  // Encabezados siempre al día (agrega Afecta / Extracto en hojas creadas con la versión anterior).
  if (sh.getMaxColumns() < ENCABEZADOS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), ENCABEZADOS.length - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS])
    .setFontWeight('bold').setBackground('#430000').setFontColor('white');
  sh.setFrozenRows(1);
  sh.setColumnWidth(ENCABEZADOS.indexOf('Extracto') + 1, 420);
  return sh;
}

/** Número de columna de cada encabezado de SEÑALES: col.Tipo, col.ID, ... */
function columnasSenales() {
  const col = {};
  ENCABEZADOS.forEach((h, i) => { col[h.replace(/[^A-Za-zÁÉÍÓÚáéíóúñÑ]/g, '')] = i + 1; col[h] = i + 1; });
  return col;
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

// ---------------------------------------------------------------------------
// Mail
// ---------------------------------------------------------------------------

function avisarPorMail(ss, sh, col, idsNuevos, leidos, primeraVez) {
  const destino = MAIL_AVISO || Session.getEffectiveUser().getEmail();
  if (!destino) return;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');

  if (primeraVez) {
    MailApp.sendEmail({
      to: destino,
      subject: `Señales: se registraron ${idsNuevos.size} archivos de la carpeta`,
      htmlBody: `<p>Se registraron ${idsNuevos.size} archivos de la carpeta de señales.</p>` +
        `<p><a href="${ss.getUrl()}">Abrir el registro de señales</a></p>`,
    });
    return;
  }

  const porId = {};
  leidos.forEach((l) => { porId[l.id] = l; });
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, ENCABEZADOS.length).getValues()
    .filter((f) => idsNuevos.has(f[col.ID - 1]));

  const cartas = [];
  const relevantes = [];
  const otros = [];
  datos.forEach((f) => {
    const l = porId[f[col.ID - 1]];
    const link = `<a href="${f[col.Link - 1]}">${esc(f[col.Título - 1])}</a>`;
    if (l && l.carta) {
      cartas.push(`<li>${link}<br><b>${esc(l.carta.mes)}</b>: ${esc(f[col.Extracto - 1])}<br><i>${esc(l.resumen)}</i></li>`);
    } else if (f[col.Afecta - 1] && f[col.Afecta - 1] !== '-') {
      relevantes.push(`<li>${link}<br><b>Afecta: ${esc(f[col.Afecta - 1])}</b><br>${esc(f[col.Extracto - 1])}</li>`);
    } else {
      otros.push(`<li><b>${esc(f[col.Tipo - 1])}</b> — ${link}</li>`);
    }
  });

  let html = '';
  if (cartas.length) html += `<h3>Cartas de objetivos cargadas automáticamente</h3><ul>${cartas.join('')}</ul>` +
    '<p>Revisá que los números coincidan con el PDF.</p>';
  if (relevantes.length) html += `<h3>Señales que afectan objetivos / incentivos</h3><ul>${relevantes.join('')}</ul>` +
    '<p>Los % de incentivo se cargan a mano en la hoja INCENTIVOS; después marcá la señal como "Cargada".</p>';
  if (otros.length) html += `<h3>Otros archivos</h3><ul>${otros.join('')}</ul>`;
  const asunto = cartas.length
    ? `Señales: carta de objetivos ${cartas.length > 1 ? 'cargadas' : 'cargada'} + ${datos.length - cartas.length} archivo(s) nuevo(s)`
    : `Señales: ${datos.length} archivo(s) nuevo(s) en la carpeta`;
  MailApp.sendEmail({
    to: destino,
    subject: asunto,
    htmlBody: html + `<p><a href="${ss.getUrl()}">Abrir el archivo de objetivos</a></p>`,
  });
}

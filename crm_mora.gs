/**
 * CRM DE GESTIÓN DE MORA — v1
 *
 * Web App (Apps Script) para que cada responsable trabaje su cartera al día
 * y deje registro de cada intento de contacto, y para que el/los
 * supervisor/es tengan control total (usuarios, asignaciones, estados) y un
 * tablero de avance.
 *
 * Va en el MISMO proyecto de Apps Script que tablero_mora.gs: reusa sus
 * constantes (HOJA_BASE, COL_AVANCE, COL_ESTADO, COL_C2, CUOTAS_TABLERO), así
 * que las posiciones de columnas de BASE se configuran en un solo lugar.
 *
 * BASE llega por IMPORTRANGE y no se escribe nunca. Todo lo del CRM
 * (usuarios, casos, gestiones, auditoría) vive en hojas CRM_* de un archivo
 * aparte (CRM_ID_ARCHIVO_DATOS), ligado al plan por Solicitud, para que
 * nadie con acceso a la planilla del tablero pueda ver ni tocar esos datos.
 *
 * Instalación: ver sección "CRM" del README.
 */

// --- CONFIGURACIÓN CRM ---
// Archivo donde se guardan las hojas CRM_* (el ID es lo que está entre /d/ y
// /edit en la URL). BASE se sigue leyendo de la planilla a la que está
// pegado este script.
const CRM_ID_ARCHIVO_DATOS = '1p_jxBNqhtRCax3qak0PlN-uEh-772HkkJ4VW0NOgZ_Y';
const CRM_HOJA_USUARIOS = 'CRM_Usuarios';
const CRM_HOJA_CASOS = 'CRM_Casos';
const CRM_HOJA_GESTIONES = 'CRM_Gestiones';
const CRM_HOJA_AUDITORIA = 'CRM_Auditoria';
const CRM_HOJA_OBJETIVOS = 'CRM_Objetivos';

// Columnas de BASE que usa el CRM además de las de tablero_mora.gs.
const CRM_COL_RESPONSABLE = 1;   // A — EZE / FABIO / SANTI
const CRM_COL_SOLICITUD = 2;     // B
const CRM_COL_GRUPO = 3;         // C
const CRM_COL_ORDEN = 4;         // D
const CRM_COL_MODELO = 5;        // E
const CRM_COL_SOBREPAUTA = 6;    // F — viene de la planilla de licitaciones
const CRM_COL_EN_CONDICIONES = 7; // G — viene de la planilla de licitaciones
const CRM_COL_CLIENTE = 8;       // H
const CRM_COL_TELEFONO = 9;      // I
const CRM_COL_TELEFONO_ALT = 10; // J
const CRM_COL_DOCUMENTO = 11;    // K
const CRM_COL_VENDEDOR = 16;     // P
const CRM_COL_SUPERVISOR_VTA = 17; // Q
const CRM_COL_FORMA_PAGO = 31;   // AE
const CRM_COL_TIPO_PLAN = 32;    // AF
const CRM_COL_SCORING = 33;      // AG — la carga el equipo; tiene que seguir en BASE
const CRM_COL_PRIMERA_NOTA = 34; // AH en adelante: seguimiento viejo en texto libre (ver crmImportarNotasBase)
const CRM_CANT_CUOTAS = 13;      // C2..C14
// Columnas de AH+ que NO son seguimiento de mora sino datos de la etapa de
// licitación (AO "ESTADO AGOSTO", AP "AGOSTO"): se muestran en la ficha como
// datos de licitación y no entran en el historial de gestiones.
const CRM_COLS_LICITACION = [41, 42]; // AO, AP

const CRM_ROL_SUPERVISOR = 'SUPERVISOR';
const CRM_ROL_RESPONSABLE = 'RESPONSABLE';

// Acceso: mail @grupoantun.com.ar + clave propia del CRM.
const CRM_DOMINIO = 'grupoantun.com.ar';  // único dominio habilitado
const CRM_CLAVE_INICIAL = 'Turin3800';    // clave de alta y de blanqueo; obliga a cambiarla
const CRM_CLAVE_MIN = 8;
const CRM_SESION_SEG = 6 * 60 * 60;       // sesión de 6 h (máximo de CacheService)
const CRM_MAX_INTENTOS = 5;               // intentos fallidos antes de bloquear 15 min
const CRM_BLOQUEO_SEG = 15 * 60;
const CRM_HASH_ITERACIONES = 500;

// Usuarios que se cargan la primera vez que se corre crmInicializar(),
// todos con CRM_CLAVE_INICIAL y cambio de clave obligatorio.
const CRM_USUARIOS_INICIALES = [
  { email: 'grietschi@grupoantun.com.ar', nombre: 'Rietschi Guillermo', rol: CRM_ROL_SUPERVISOR, alias: '' },
  { email: 'gortiz@grupoantun.com.ar', nombre: 'Ortiz Gabriel', rol: CRM_ROL_SUPERVISOR, alias: '' },
  { email: 'sgodoy@grupoantun.com.ar', nombre: 'Godoy Santiago', rol: CRM_ROL_RESPONSABLE, alias: 'SANTI' },
  { email: 'evaca@grupoantun.com.ar', nombre: 'Vaca Ezequiel', rol: CRM_ROL_RESPONSABLE, alias: 'EZE' },
  { email: 'faguero@grupoantun.com.ar', nombre: 'Aguero Fabio', rol: CRM_ROL_RESPONSABLE, alias: 'FABIO' },
];

// Estado "Estado" de BASE que saca al plan de la cartera activa (se ven en
// la vista "Rescindidos / bajas" pero no entran en la cola de trabajo).
const CRM_ESTADOS_BAJA = ['Rescindido', 'Renunciado', 'Cancelado'];

// Listas del formulario de gestión (a ajustar juntos).
const CRM_ESTADOS_CASO = ['Sin gestionar', 'Contactando', 'Contactado', 'Promesa de pago',
  'Pagó (a verificar)', 'Negativa de pago', 'Inubicable', 'Derivado'];
const CRM_CANALES = ['Llamada', 'WhatsApp', 'SMS', 'Mail', 'Presencial', 'Sin contacto (solo actualización)'];
const CRM_RESULTADOS = ['Atendió', 'No atendió', 'Buzón / apagado', 'Mensaje enviado',
  'Respondió mensaje', 'Número erróneo', 'N/A'];
const CRM_RESULTADOS_EFECTIVOS = ['Atendió', 'Respondió mensaje'];
// Canal con el que quedan en CRM_Gestiones las notas importadas de BASE.
// Van sin FechaHora, así no cuentan como gestiones del mes ni en la actividad.
const CRM_CANAL_IMPORTADO = 'Importado de planilla';
// Motivos de no pago: los códigos que ya usa el equipo (col. "ACT MOTIVO NO PAGO").
const CRM_MOTIVOS_NO_PAGO = ['',
  'P/E - Problemas económicos',
  'S/C - Sin contacto',
  'M/V - Mala venta',
  'T/U - Toma de usado',
  'P/P - Problemas personales',
  'A/C - Área comercial',
  'TR - Transferencia',
  'T/P - Título propio',
  'A/A - Área administrativa',
  'S/G - Sin gestión',
  'C/C - Cambio de concesionario',
];

// Situación de cada plan según BASE (ver crmSituacion). "orden" define la
// prioridad en la cola de trabajo; "detalle" es la explicación que ve el
// operador en la leyenda.
const CRM_SITUACIONES = {
  P1: { orden: 1, etiqueta: 'P1 · Atrasado', detalle: 'Tiene cuotas vencidas impagas y está en un avance que se mide el mes próximo.' },
  P2: { orden: 2, etiqueta: 'P2 · Cuota del mes', detalle: 'Lo vencido está pago, falta la cuota de este mes, y está en un avance que se mide el mes próximo.' },
  P3: { orden: 3, etiqueta: 'P3 · Atrasado', detalle: 'Tiene cuotas vencidas impagas, en el resto de los avances.' },
  P4: { orden: 4, etiqueta: 'P4 · Cuota del mes', detalle: 'Lo vencido está pago, falta la cuota de este mes, en el resto de los avances.' },
  AL_DIA: { orden: 5, etiqueta: 'Al día' },
  RESCINDIDO: { orden: 8, etiqueta: 'Rescindido' },
  BAJA: { orden: 9, etiqueta: 'Baja' },
};

const CRM_ENC_USUARIOS = ['Email', 'Nombre', 'Rol', 'AliasBase', 'Activo',
  'ClaveHash', 'ClaveSal', 'DebeCambiarClave'];
const CRM_ENC_CASOS = ['Solicitud', 'ResponsableEmail', 'EstadoCaso', 'UltimaGestion',
  'UltimoCanal', 'UltimoResultado', 'ProximoContacto', 'PromesaFecha', 'PromesaMonto',
  'MotivoNoPago', 'ActualizadoPor'];
const CRM_ENC_GESTIONES = ['FechaHora', 'UsuarioEmail', 'UsuarioNombre', 'Solicitud', 'Canal',
  'Resultado', 'EstadoCaso', 'MotivoNoPago', 'PromesaFecha', 'PromesaMonto',
  'ProximoContacto', 'Nota', 'AvanceAlMomento', 'SituacionAlMomento'];
const CRM_ENC_AUDITORIA = ['FechaHora', 'UsuarioEmail', 'Accion', 'Detalle'];
// Objetivos que el supervisor le pone a un responsable (o a todos, '*'):
// un grupo de casos (por avance y prioridad) a gestionar entre dos fechas.
const CRM_ENC_OBJETIVOS = ['Id', 'Titulo', 'ResponsableEmail', 'Avances', 'Situaciones', 'Meta',
  'Desde', 'Hasta', 'CreadoPor', 'Activo'];

// ---------------------------------------------------------------------------
// Web App
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('crm_index').evaluate()
    .setTitle('CRM Mora')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Crea las hojas CRM_* (si no existen) y carga los usuarios iniciales con la
 * clave inicial. Se puede correr más de una vez: no duplica usuarios, no
 * borra datos y no toca las claves de los usuarios que ya existen.
 */
function crmInicializar() {
  const ss = crmDatos_();
  const movidas = crmMoverHojasViejas_(ss);
  const hoja = crmHoja_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS);
  // Por si la hoja se creó con una versión anterior (sin columnas de clave).
  hoja.getRange(1, 1, 1, CRM_ENC_USUARIOS.length).setValues([CRM_ENC_USUARIOS]).setFontWeight('bold');
  crmHoja_(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS);
  crmHoja_(ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES);
  crmHoja_(ss, CRM_HOJA_AUDITORIA, CRM_ENC_AUDITORIA);
  crmHoja_(ss, CRM_HOJA_OBJETIVOS, CRM_ENC_OBJETIVOS);

  const usuarios = crmLeerUsuarios_(ss);
  const existentes = usuarios.map(function (u) { return u.email; });
  CRM_USUARIOS_INICIALES.forEach(function (u) {
    if (existentes.indexOf(u.email) >= 0) return;
    hoja.appendRow([u.email, u.nombre, u.rol, u.alias, true].concat(crmClaveInicial_()));
    existentes.push(u.email);
  });
  // Usuarios cargados sin clave (versión anterior o a mano en la hoja).
  usuarios.forEach(function (u, i) {
    if (!u.tieneClave) hoja.getRange(i + 2, 6, 1, 3).setValues([crmClaveInicial_()]);
  });
  crmAuditar_(ss, Session.getEffectiveUser().getEmail(), 'Inicializar CRM', 'Hojas y usuarios iniciales verificados');
  SpreadsheetApp.getUi().alert('CRM listo. Los datos se guardan en el archivo "' + ss.getName() + '".' +
    (movidas.length ? '\n\nSe copiaron ahí las hojas ' + movidas.join(', ') +
      ' que estaban en esta planilla: ya las podés borrar de acá.' : '') +
    '\n\nSi todavía no lo hiciste, publicalo: Implementar > Nueva implementación > Aplicación web (ver README).');
}

/**
 * Si las hojas CRM_* se habían creado en la planilla del tablero (versión
 * anterior), copia su contenido al archivo de datos, solo cuando allá la
 * hoja todavía está vacía. No borra nada: devuelve los nombres copiados
 * para avisar que se pueden eliminar a mano.
 */
function crmMoverHojasViejas_(ssDatos) {
  const ssVieja = SpreadsheetApp.getActiveSpreadsheet();
  if (ssVieja.getId() === ssDatos.getId()) return [];
  const movidas = [];
  [CRM_HOJA_USUARIOS, CRM_HOJA_CASOS, CRM_HOJA_GESTIONES, CRM_HOJA_AUDITORIA].forEach(function (nombre) {
    const vieja = ssVieja.getSheetByName(nombre);
    if (!vieja || vieja.getLastRow() < 2) return;
    const nueva = ssDatos.getSheetByName(nombre);
    if (nueva && nueva.getLastRow() > 1) return; // ya tiene datos: no pisar
    const valores = vieja.getRange(1, 1, vieja.getLastRow(), vieja.getLastColumn()).getValues();
    const destino = nueva || ssDatos.insertSheet(nombre);
    destino.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
    destino.getRange(1, 1, 1, valores[0].length).setFontWeight('bold');
    destino.setFrozenRows(1);
    movidas.push(nombre);
  });
  return movidas;
}

/**
 * Copia UNA vez el seguimiento viejo en texto libre de BASE (columnas AH en
 * adelante) a CRM_Gestiones, una fila por celda con contenido, con canal
 * CRM_CANAL_IMPORTADO. Salta los planes que ya tienen notas importadas, así
 * que se puede volver a correr (por ejemplo, para planes nuevos) sin
 * duplicar. Después de importar, esas columnas de BASE se pueden borrar:
 * la ficha muestra lo importado.
 */
function crmImportarNotasBase() {
  const ctx = crmContexto_();
  const yaImportados = {};
  crmLeerObjetos_(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES).forEach(function (g) {
    if (g.Canal === CRM_CANAL_IMPORTADO) yaImportados[String(g.Solicitud)] = true;
  });

  const filas = [];
  let planes = 0;
  ctx.planes.forEach(function (p) {
    // También las de licitación (AO/AP): así no se pierden si se borran esas
    // columnas de BASE; la ficha las muestra aparte, no como seguimiento.
    const notas = p.notasBase.concat(p.licitacion);
    if (yaImportados[p.solicitud] || !notas.length) return;
    planes++;
    notas.forEach(function (n) {
      filas.push(['', '', 'Planilla BASE', p.solicitud, CRM_CANAL_IMPORTADO, '', '', '', '', '', '',
        '[' + n.columna + '] ' + n.valor, p.avance, p.situacion.codigo]);
    });
  });

  if (filas.length) {
    const hoja = crmHoja_(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES);
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, CRM_ENC_GESTIONES.length).setValues(filas);
  }
  crmAuditar_(ctx.ss, Session.getEffectiveUser().getEmail(), 'Importar notas de BASE',
    filas.length + ' notas de ' + planes + ' planes');
  SpreadsheetApp.getUi().alert('Importación lista: ' + filas.length + ' notas de ' + planes + ' planes.' +
    (Object.keys(yaImportados).length ? '\n(' + Object.keys(yaImportados).length + ' planes ya estaban importados y se saltearon.)' : '') +
    '\n\nRevisá algunas fichas en el CRM antes de borrar las columnas AH en adelante de BASE (la AG, Scoring, se queda).');
}

/** Archivo de datos del CRM (hojas CRM_*). */
function crmDatos_() {
  return SpreadsheetApp.openById(CRM_ID_ARCHIVO_DATOS);
}

// ---------------------------------------------------------------------------
// Lógica de negocio (funciones puras, sin acceso a hojas)
// ---------------------------------------------------------------------------

/**
 * Situación de UN plan de BASE, mirando lo que Fiat va a medir el mes que
 * viene. Un plan hoy en avance N tiene cargadas C2..C(N-1) y la cuota del
 * mes en curso es C(N) (vacía o 'I' hasta que paga). El mes que viene ese
 * plan pasa a avance N+1 y, si N es una de CUOTAS_TABLERO, Fiat mide
 * C2..C(N): o sea las cuotas ya vencidas MÁS la del mes en curso.
 *
 * - RESCINDIDO: alguna 'R' en C2..C(N), o Estado = Rescindido.
 * - BAJA: Estado en CRM_ESTADOS_BAJA (Renunciado, Cancelado).
 * - P1 / P3: alguna 'I' en C2..C(N-1) (cuotas ya vencidas impagas).
 *   P1 si el avance es de los que mide Fiat, P3 si no.
 * - P2 / P4: vencidas al día pero la cuota del mes (C(N)) todavía no es 'P'.
 * - AL_DIA: todo pago, incluida la cuota del mes.
 * Devuelve null si la fila no tiene un Avance válido.
 */
function crmSituacion(fila) {
  const avance = fila[COL_AVANCE - 1];
  if (typeof avance !== 'number' || avance < 2) return null;
  const estado = String(fila[COL_ESTADO - 1] || '').trim();

  const offsetC2 = COL_C2 - 1;
  const vencidas = fila.slice(offsetC2, offsetC2 + avance - 2);     // C2..C(N-1)
  const cuotaMes = avance - 2 < CRM_CANT_CUOTAS ? fila[offsetC2 + avance - 2] : 'P'; // C(N)
  const medidoPorFiat = CUOTAS_TABLERO.indexOf(avance) >= 0;

  let codigo;
  if (estado === 'Rescindido' || vencidas.concat([cuotaMes]).some(function (v) { return v === 'R'; })) {
    codigo = 'RESCINDIDO';
  } else if (CRM_ESTADOS_BAJA.indexOf(estado) >= 0) {
    codigo = 'BAJA';
  } else if (vencidas.some(function (v) { return v === 'I'; })) {
    codigo = medidoPorFiat ? 'P1' : 'P3';
  } else if (cuotaMes !== 'P') {
    codigo = medidoPorFiat ? 'P2' : 'P4';
  } else {
    codigo = 'AL_DIA';
  }

  return {
    codigo: codigo,
    etiqueta: CRM_SITUACIONES[codigo].etiqueta,
    orden: CRM_SITUACIONES[codigo].orden,
    medidoPorFiat: medidoPorFiat,
    cuotasAtrasadas: vencidas.filter(function (v) { return v === 'I'; }).length,
    cuotaMesPaga: cuotaMes === 'P',
  };
}

/** true si la situación es parte de la cartera activa (cola de trabajo). */
function crmEsActiva(codigo) {
  return codigo !== 'RESCINDIDO' && codigo !== 'BAJA';
}

// ---------------------------------------------------------------------------
// API para la Web App (google.script.run)
// ---------------------------------------------------------------------------

/** Datos iniciales: quién soy y las listas del formulario. */
function crmInicio(token) {
  const u = crmUsuarioActual_(token);
  const ss = crmDatos_();
  return {
    usuario: u,
    esSupervisor: u.rol === CRM_ROL_SUPERVISOR,
    // Solo responsables activos: a un supervisor no se le asignan planes.
    responsables: u.rol === CRM_ROL_SUPERVISOR
      ? crmLeerUsuarios_(ss).filter(function (x) { return x.activo && x.rol === CRM_ROL_RESPONSABLE; })
        .map(function (x) { return { email: x.email, nombre: x.nombre }; })
      : [],
    listas: {
      estadosCaso: CRM_ESTADOS_CASO, canales: CRM_CANALES, resultados: CRM_RESULTADOS,
      motivos: CRM_MOTIVOS_NO_PAGO, situaciones: CRM_SITUACIONES,
    },
    cuotasFiat: CUOTAS_TABLERO,
  };
}

/**
 * Cartera visible para el usuario: el responsable ve solo la suya, el
 * supervisor ve todo. Incluye rescindidos/bajas (la vista los separa).
 */
function crmListarCartera(token) {
  const u = crmUsuarioActual_(token);
  const ctx = crmContexto_();
  const hoy = crmHoyStr();
  return ctx.planes
    .filter(function (p) { return crmPuedeVer_(u, p); })
    .map(function (p) {
      const c = ctx.casos[p.solicitud] || {};
      const g = ctx.gestionesMes[p.solicitud] || { intentos: 0, efectivos: 0 };
      return {
        solicitud: p.solicitud, cliente: p.cliente, telefono: p.telefono, grupo: p.grupo,
        orden: p.orden, avance: p.avance, estadoBase: p.estado, cuotas: p.cuotas,
        situacion: p.situacion, responsableEmail: p.responsableEmail,
        responsableNombre: p.responsableNombre,
        estadoCaso: c.estadoCaso || 'Sin gestionar',
        ultimaGestion: c.ultimaGestion || '', ultimoCanal: c.ultimoCanal || '',
        ultimoResultado: c.ultimoResultado || '', proximoContacto: c.proximoContacto || '',
        promesaFecha: c.promesaFecha || '', promesaMonto: c.promesaMonto || '',
        intentosMes: g.intentos, efectivosMes: g.efectivos,
        agendaHoy: (!!c.proximoContacto && c.proximoContacto <= hoy) ||
          (!!c.promesaFecha && c.promesaFecha <= hoy && !p.situacion.cuotaMesPaga),
      };
    });
}

/** Ficha completa de un plan: datos de BASE, notas viejas y gestiones. */
function crmFichaPlan(token, solicitud) {
  const u = crmUsuarioActual_(token);
  const ctx = crmContexto_();
  const p = crmBuscarPlan_(ctx, solicitud);
  if (!crmPuedeVer_(u, p)) throw new Error('No tenés acceso a este plan.');

  const delPlan = crmLeerObjetos_(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES)
    .filter(function (g) { return String(g.Solicitud) === p.solicitud; });

  // Historial viejo: si ya se importó, sale de CRM_Gestiones (sirve aunque
  // se borren las columnas de BASE); si no, de las columnas AH+ en vivo.
  const importadas = delPlan.filter(function (g) { return g.Canal === CRM_CANAL_IMPORTADO; });
  if (importadas.length) {
    const notas = importadas.map(function (g) {
      const m = String(g.Nota).match(/^\[(.*?)\] ([\s\S]*)$/);
      return m ? { columna: m[1], valor: m[2] } : { columna: '', valor: String(g.Nota) };
    });
    // Las columnas de licitación (si se importaron con una versión anterior)
    // no son seguimiento: van al bloque de licitación, solo si BASE ya no las tiene.
    p.notasBase = notas.filter(function (n) { return !crmEsNotaLicitacion_(n.columna); });
    if (!p.licitacion.length) p.licitacion = notas.filter(function (n) { return crmEsNotaLicitacion_(n.columna); });
  }

  const gestiones = delPlan
    .filter(function (g) { return g.Canal !== CRM_CANAL_IMPORTADO; })
    .map(function (g) {
      return {
        fechaHora: crmFmt(g.FechaHora, 'yyyy-MM-dd HH:mm'), usuario: g.UsuarioNombre || g.UsuarioEmail,
        canal: g.Canal, resultado: g.Resultado, estadoCaso: g.EstadoCaso, motivo: g.MotivoNoPago,
        promesaFecha: crmFmt(g.PromesaFecha), promesaMonto: g.PromesaMonto,
        proximoContacto: crmFmt(g.ProximoContacto), nota: g.Nota,
      };
    })
    .reverse();

  return {
    plan: p,
    caso: ctx.casos[p.solicitud] || { estadoCaso: 'Sin gestionar' },
    gestiones: gestiones,
  };
}

/**
 * Registra un intento de contacto / actualización sobre un plan y actualiza
 * el estado vigente del caso.
 */
function crmRegistrarGestion(token, solicitud, datos) {
  const u = crmUsuarioActual_(token);
  const ctx = crmContexto_();
  const p = crmBuscarPlan_(ctx, solicitud);
  if (!crmPuedeVer_(u, p)) throw new Error('No tenés acceso a este plan.');
  if (CRM_CANALES.indexOf(datos.canal) < 0) throw new Error('Canal inválido.');
  if (CRM_ESTADOS_CASO.indexOf(datos.estadoCaso) < 0) throw new Error('Estado inválido.');
  if (datos.estadoCaso === 'Promesa de pago' && !datos.promesaFecha) {
    throw new Error('Para "Promesa de pago" cargá la fecha prometida.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ahora = new Date();
    const hojaG = crmHoja_(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES);
    hojaG.appendRow([
      ahora, u.email, u.nombre, p.solicitud, datos.canal, datos.resultado || '',
      datos.estadoCaso, datos.motivo || '', crmFecha(datos.promesaFecha), datos.promesaMonto || '',
      crmFecha(datos.proximoContacto), datos.nota || '', p.avance, p.situacion.codigo,
    ]);
    crmActualizarCaso_(ctx.ss, p.solicitud, {
      EstadoCaso: datos.estadoCaso,
      UltimaGestion: ahora,
      UltimoCanal: datos.canal,
      UltimoResultado: datos.resultado || '',
      ProximoContacto: crmFecha(datos.proximoContacto),
      PromesaFecha: crmFecha(datos.promesaFecha),
      PromesaMonto: datos.promesaMonto || '',
      MotivoNoPago: datos.motivo || '',
      ActualizadoPor: u.email,
    });
  } finally {
    lock.releaseLock();
  }
  return crmFichaPlan(token, solicitud);
}

/** Supervisor: reasigna uno o varios planes a un responsable. */
function crmReasignar(token, solicitudes, email) {
  const u = crmExigirSupervisor_(token);
  const ss = crmDatos_();
  const destino = crmLeerUsuarios_(ss).filter(function (x) {
    return x.email === email && x.activo && x.rol === CRM_ROL_RESPONSABLE;
  })[0];
  if (!destino) throw new Error('Solo se puede asignar a un responsable activo.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    solicitudes.forEach(function (s) {
      crmActualizarCaso_(ss, String(s), { ResponsableEmail: email, ActualizadoPor: u.email });
    });
    crmAuditar_(ss, u.email, 'Reasignar', solicitudes.length + ' plan(es) → ' + email + ': ' + solicitudes.join(', '));
  } finally {
    lock.releaseLock();
  }
  return solicitudes.length;
}

/** Supervisor: lista de usuarios. */
function crmListarUsuarios(token) {
  crmExigirSupervisor_(token);
  return crmLeerUsuarios_(crmDatos_());
}

/** Supervisor: alta o modificación de un usuario (clave = email). */
function crmGuardarUsuario(token, datos) {
  const u = crmExigirSupervisor_(token);
  const ss = crmDatos_();
  const email = String(datos.email || '').trim().toLowerCase();
  if (!crmEmailValido_(email)) throw new Error('El email tiene que ser @' + CRM_DOMINIO + '.');
  if ([CRM_ROL_SUPERVISOR, CRM_ROL_RESPONSABLE].indexOf(datos.rol) < 0) throw new Error('Rol inválido.');

  const usuarios = crmLeerUsuarios_(ss);
  const quedanSupervisores = usuarios.filter(function (x) {
    return x.email !== email && x.activo && x.rol === CRM_ROL_SUPERVISOR;
  }).length + (datos.activo && datos.rol === CRM_ROL_SUPERVISOR ? 1 : 0);
  if (quedanSupervisores === 0) throw new Error('Tiene que quedar al menos un supervisor activo.');

  const fila = [email, datos.nombre || email, datos.rol,
    String(datos.alias || '').trim().toUpperCase(), !!datos.activo];
  const hoja = crmHoja_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS);
  const idx = usuarios.map(function (x) { return x.email; }).indexOf(email);
  // Modificar no toca la clave; un alta nace con la clave inicial.
  if (idx >= 0) hoja.getRange(idx + 2, 1, 1, fila.length).setValues([fila]);
  else hoja.appendRow(fila.concat(crmClaveInicial_()));
  if (idx >= 0 && !fila[4]) crmCerrarSesiones_(email);
  crmAuditar_(ss, u.email, idx >= 0 ? 'Modificar usuario' : 'Alta usuario', JSON.stringify(fila));
  return crmLeerUsuarios_(ss);
}

/** Supervisor: números para el tablero. */
function crmTableroSupervisor(token) {
  crmExigirSupervisor_(token);
  const ctx = crmContexto_();
  const hoy = crmHoyStr();

  // Por responsable
  const porResp = {};
  crmLeerUsuarios_(ctx.ss).forEach(function (x) {
    if (x.rol === CRM_ROL_RESPONSABLE && x.activo) porResp[x.email] = crmFilaResp_(x.nombre);
  });
  const sinAsignar = crmFilaResp_('(sin asignar)');

  ctx.planes.forEach(function (p) {
    if (p.responsableEmail && !porResp[p.responsableEmail]) {
      porResp[p.responsableEmail] = crmFilaResp_(p.responsableNombre);
    }
    const r = p.responsableEmail ? porResp[p.responsableEmail] : sinAsignar;
    const cod = p.situacion.codigo;
    if (!crmEsActiva(cod)) { r.bajas++; return; }
    const c = ctx.casos[p.solicitud] || {};
    const g = ctx.gestionesMes[p.solicitud];
    r.activa++;
    if (cod === 'P1') r.p1++;
    if (cod === 'P2') r.p2++;
    if (cod === 'P3' || cod === 'P4') r.p34++;
    if (cod === 'P1' || cod === 'P2') {
      if (g) r.prioGestionados++; else r.prioSinGestion++;
    }
    if (g) {
      r.intentosMes += g.intentos;
      r.efectivosMes += g.efectivos;
      if (cod === 'AL_DIA') r.regularizadosConGestion++;
    }
    if (c.estadoCaso === 'Promesa de pago' && c.promesaFecha) {
      if (c.promesaFecha >= hoy) r.promesasVigentes++;
      else if (cod !== 'AL_DIA') r.promesasVencidas++;
    }
  });
  const filas = Object.keys(porResp).map(function (k) { return porResp[k]; });
  if (sinAsignar.activa + sinAsignar.bajas > 0) filas.push(sinAsignar);

  // Proyección de lo que va a medir Fiat el mes que viene, cuota por cuota,
  // con la misma definición de cartera que el Tablero (todas las filas de
  // ese avance, rescindidos incluidos).
  const proyeccion = CUOTAS_TABLERO.map(function (cuota) {
    const delAvance = ctx.planes.filter(function (p) { return p.avance === cuota; });
    const cuenta = function (codigos) {
      return delAvance.filter(function (p) { return codigos.indexOf(p.situacion.codigo) >= 0; }).length;
    };
    const total = delAvance.length;
    const rescindidos = cuenta(['RESCINDIDO']);
    const atrasados = cuenta(['P1']);
    const mesPendiente = cuenta(['P2']);
    return {
      cuota: cuota, avanceHoy: cuota, carteraTotal: total, rescindidos: rescindidos,
      atrasados: atrasados, mesPendiente: mesPendiente,
      pctPiso: total ? (rescindidos + atrasados) / total : 0,
      pctSiNadiePaga: total ? (rescindidos + atrasados + mesPendiente) / total : 0,
    };
  });

  // Actividad de los últimos 14 días, por día y responsable.
  const desde = new Date();
  desde.setDate(desde.getDate() - 13);
  const desdeStr = crmFmt(desde);
  const actividad = {};
  crmLeerObjetos_(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES).forEach(function (g) {
    const dia = crmFmt(g.FechaHora);
    if (dia < desdeStr) return;
    const quien = g.UsuarioNombre || g.UsuarioEmail;
    actividad[dia] = actividad[dia] || {};
    actividad[dia][quien] = (actividad[dia][quien] || 0) + 1;
  });

  return { porResponsable: filas, proyeccion: proyeccion, actividad: actividad, hoy: hoy };
}

// ---------------------------------------------------------------------------
// Objetivos del supervisor
// ---------------------------------------------------------------------------

/**
 * Objetivos vigentes. El responsable ve los suyos y los de "todos" que no
 * vencieron; el supervisor ve todos los activos. El avance de cada uno lo
 * calcula la pantalla con la cartera que ya tiene cargada (casos que
 * cumplen el filtro y cuántos tienen una gestión desde la fecha "Desde").
 */
function crmListarObjetivos(token) {
  const u = crmUsuarioActual_(token);
  const hoy = crmHoyStr();
  return crmLeerObjetos_(crmDatos_(), CRM_HOJA_OBJETIVOS, CRM_ENC_OBJETIVOS)
    .filter(function (o) { return crmBool_(o.Activo); })
    .map(function (o) {
      return {
        id: String(o.Id), titulo: String(o.Titulo), responsableEmail: String(o.ResponsableEmail).toLowerCase(),
        avances: String(o.Avances || '').split(',').filter(String).map(Number),
        situaciones: String(o.Situaciones || '').split(',').filter(String),
        meta: Number(o.Meta) || 100, desde: crmFmt(o.Desde), hasta: crmFmt(o.Hasta),
        creadoPor: String(o.CreadoPor), vencido: crmFmt(o.Hasta) < hoy,
      };
    })
    .filter(function (o) {
      return u.rol === CRM_ROL_SUPERVISOR || (!o.vencido && (o.responsableEmail === '*' || o.responsableEmail === u.email));
    });
}

/** Supervisor: alta de un objetivo. */
function crmGuardarObjetivo(token, datos) {
  const u = crmExigirSupervisor_(token);
  const ss = crmDatos_();
  const titulo = String(datos.titulo || '').trim();
  if (!titulo) throw new Error('Poné un título para el objetivo.');
  const resp = String(datos.responsableEmail || '').toLowerCase();
  if (resp !== '*' && !crmLeerUsuarios_(ss).some(function (x) { return x.email === resp && x.activo && x.rol === CRM_ROL_RESPONSABLE; })) {
    throw new Error('Elegí un responsable activo (o "Todos").');
  }
  const avances = (datos.avances || []).map(Number).filter(function (n) { return n >= 2 && n <= 20; });
  const situaciones = (datos.situaciones || []).filter(function (x) { return ['P1', 'P2', 'P3', 'P4'].indexOf(x) >= 0; });
  if (!situaciones.length) throw new Error('Elegí al menos una prioridad.');
  const meta = Math.min(100, Math.max(1, Number(datos.meta) || 100));
  if (!datos.desde || !datos.hasta || datos.hasta < datos.desde) throw new Error('Revisá las fechas: "hasta" tiene que ser igual o posterior a "desde".');

  const fila = [Utilities.getUuid().slice(0, 8), titulo, resp, avances.join(','), situaciones.join(','), meta,
    crmFecha(datos.desde), crmFecha(datos.hasta), u.email, true];
  crmHoja_(ss, CRM_HOJA_OBJETIVOS, CRM_ENC_OBJETIVOS).appendRow(fila);
  crmAuditar_(ss, u.email, 'Alta objetivo', titulo + ' → ' + resp);
  return crmListarObjetivos(token);
}

/** Supervisor: da de baja un objetivo (queda en la hoja como inactivo). */
function crmBajaObjetivo(token, id) {
  const u = crmExigirSupervisor_(token);
  const ss = crmDatos_();
  const hoja = crmHoja_(ss, CRM_HOJA_OBJETIVOS, CRM_ENC_OBJETIVOS);
  const filas = crmLeerObjetos_(ss, CRM_HOJA_OBJETIVOS, CRM_ENC_OBJETIVOS);
  for (let i = 0; i < filas.length; i++) {
    if (String(filas[i].Id) === String(id)) {
      hoja.getRange(i + 2, CRM_ENC_OBJETIVOS.indexOf('Activo') + 1).setValue(false);
      crmAuditar_(ss, u.email, 'Baja objetivo', String(filas[i].Titulo));
    }
  }
  return crmListarObjetivos(token);
}

// ---------------------------------------------------------------------------
// Acceso: login con mail @grupoantun.com.ar + clave
// ---------------------------------------------------------------------------

/** Valida mail y clave; devuelve un token de sesión. */
function crmLogin(email, clave) {
  email = String(email || '').trim().toLowerCase();
  if (!crmEmailValido_(email)) throw new Error('Ingresá con tu mail @' + CRM_DOMINIO + '.');
  const cache = CacheService.getScriptCache();
  const claveIntentos = 'int_' + email;
  const intentos = Number(cache.get(claveIntentos) || 0);
  if (intentos >= CRM_MAX_INTENTOS) {
    throw new Error('Demasiados intentos fallidos. Esperá 15 minutos o pedí a un supervisor que blanquee tu clave.');
  }

  const ss = crmDatos_();
  const fila = crmFilaUsuario_(ss, email);
  const ok = fila && fila.activo && fila.hash && crmHashClave_(clave, fila.sal) === fila.hash;
  if (!ok) {
    cache.put(claveIntentos, String(intentos + 1), CRM_BLOQUEO_SEG);
    throw new Error('Mail o clave incorrectos.');
  }
  cache.remove(claveIntentos);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('ses_' + token, JSON.stringify({ email: email, t: Date.now() }), CRM_SESION_SEG);
  crmAuditar_(ss, email, 'Ingreso', '');
  return { token: token, debeCambiarClave: fila.debeCambiar };
}

function crmLogout(token) {
  if (token) CacheService.getScriptCache().remove('ses_' + token);
}

/** Cambio de clave del propio usuario (obligatorio tras alta o blanqueo). */
function crmCambiarClave(token, actual, nueva) {
  const u = crmUsuarioActual_(token, true);
  const ss = crmDatos_();
  const fila = crmFilaUsuario_(ss, u.email);
  if (crmHashClave_(actual, fila.sal) !== fila.hash) throw new Error('La clave actual no es correcta.');
  nueva = String(nueva || '');
  if (nueva.length < CRM_CLAVE_MIN) throw new Error('La clave nueva tiene que tener al menos ' + CRM_CLAVE_MIN + ' caracteres.');
  if (nueva === CRM_CLAVE_INICIAL || nueva === actual) throw new Error('La clave nueva tiene que ser distinta de la actual.');
  const sal = Utilities.getUuid();
  crmHoja_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS).getRange(fila.fila, 6, 1, 3)
    .setValues([[crmHashClave_(nueva, sal), sal, false]]);
  crmAuditar_(ss, u.email, 'Cambio de clave', '');
  return true;
}

/** Supervisor: vuelve la clave de un usuario a la inicial y obliga a cambiarla. */
function crmBlanquearClave(token, email) {
  const u = crmExigirSupervisor_(token);
  const ss = crmDatos_();
  const fila = crmFilaUsuario_(ss, String(email || '').toLowerCase());
  if (!fila) throw new Error('Usuario inexistente.');
  crmHoja_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS).getRange(fila.fila, 6, 1, 3).setValues([crmClaveInicial_()]);
  CacheService.getScriptCache().remove('int_' + fila.email);
  crmCerrarSesiones_(fila.email);
  crmAuditar_(ss, u.email, 'Blanqueo de clave', fila.email);
  return crmLeerUsuarios_(ss);
}

function crmEmailValido_(email) {
  return new RegExp('^[^@\\s]+@' + CRM_DOMINIO.replace(/\./g, '\\.') + '$').test(email);
}

/** [hash, sal, debeCambiar] de la clave inicial, con sal nueva. */
function crmClaveInicial_() {
  const sal = Utilities.getUuid();
  return [crmHashClave_(CRM_CLAVE_INICIAL, sal), sal, true];
}

function crmHashClave_(clave, sal) {
  let h = sal + '|' + String(clave || '');
  for (let i = 0; i < CRM_HASH_ITERACIONES; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + sal, Utilities.Charset.UTF_8)
      .map(function (b) { return ((b < 0 ? b + 256 : b)).toString(16).padStart(2, '0'); }).join('');
  }
  return h;
}

/** Fila cruda de CRM_Usuarios (con hash), o null. Nunca sale al navegador. */
function crmFilaUsuario_(ss, email) {
  const filas = crmLeerObjetos_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS);
  for (let i = 0; i < filas.length; i++) {
    if (String(filas[i].Email).trim().toLowerCase() === email) {
      return {
        fila: i + 2, email: email, activo: crmBool_(filas[i].Activo),
        hash: String(filas[i].ClaveHash || ''), sal: String(filas[i].ClaveSal || ''),
        debeCambiar: crmBool_(filas[i].DebeCambiarClave),
      };
    }
  }
  return null;
}

/**
 * Las sesiones viven en CacheService (no se pueden listar), así que para
 * cortar las de un usuario se marca una "época": cualquier token de ese
 * mail emitido antes deja de valer (ver crmUsuarioActual_).
 */
function crmCerrarSesiones_(email) {
  CacheService.getScriptCache().put('epoca_' + email, String(Date.now()), CRM_SESION_SEG);
}

function crmBool_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

// ---------------------------------------------------------------------------
// Acceso a datos
// ---------------------------------------------------------------------------

/** Lee BASE + CRM_Casos + gestiones del mes y arma los planes resueltos. */
function crmContexto_() {
  const ss = crmDatos_();
  const hojaBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_BASE);
  if (!hojaBase) throw new Error('No encuentro la hoja "' + HOJA_BASE + '"');
  const valores = hojaBase.getDataRange().getValues();
  const encabezados = valores[0];

  const usuarios = crmLeerUsuarios_(ss);
  const porAlias = {};
  const porEmail = {};
  usuarios.forEach(function (x) {
    porEmail[x.email] = x;
    if (x.alias) porAlias[x.alias] = x;
  });

  const casos = {};
  crmLeerObjetos_(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS).forEach(function (c) {
    casos[String(c.Solicitud)] = {
      responsableEmail: String(c.ResponsableEmail || '').toLowerCase(),
      estadoCaso: c.EstadoCaso || 'Sin gestionar',
      ultimaGestion: crmFmt(c.UltimaGestion, 'yyyy-MM-dd HH:mm'),
      ultimoCanal: c.UltimoCanal, ultimoResultado: c.UltimoResultado,
      proximoContacto: crmFmt(c.ProximoContacto), promesaFecha: crmFmt(c.PromesaFecha),
      promesaMonto: c.PromesaMonto, motivo: c.MotivoNoPago,
    };
  });

  const mes = crmFmt(new Date(), 'yyyy-MM');
  const gestionesMes = {};
  crmLeerObjetos_(ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES).forEach(function (g) {
    if (crmFmt(g.FechaHora, 'yyyy-MM') !== mes) return;
    const k = String(g.Solicitud);
    gestionesMes[k] = gestionesMes[k] || { intentos: 0, efectivos: 0 };
    gestionesMes[k].intentos++;
    if (CRM_RESULTADOS_EFECTIVOS.indexOf(g.Resultado) >= 0) gestionesMes[k].efectivos++;
  });

  const planes = [];
  for (let i = 1; i < valores.length; i++) {
    const f = valores[i];
    const solicitud = String(f[CRM_COL_SOLICITUD - 1] || '').trim();
    if (!solicitud) continue;
    const situacion = crmSituacion(f);
    if (!situacion) continue;

    const caso = casos[solicitud];
    const alias = String(f[CRM_COL_RESPONSABLE - 1] || '').trim().toUpperCase();
    const resp = (caso && caso.responsableEmail && porEmail[caso.responsableEmail]) || porAlias[alias];

    planes.push({
      solicitud: solicitud,
      grupo: crmFmt(f[CRM_COL_GRUPO - 1]), orden: crmFmt(f[CRM_COL_ORDEN - 1]),
      modelo: crmFmt(f[CRM_COL_MODELO - 1]), sobrepauta: crmFmt(f[CRM_COL_SOBREPAUTA - 1]),
      enCondiciones: crmFmt(f[CRM_COL_EN_CONDICIONES - 1]),
      licitacion: crmColumnasBase_(encabezados, f, CRM_COLS_LICITACION),
      cliente: crmFmt(f[CRM_COL_CLIENTE - 1]), telefono: crmFmt(f[CRM_COL_TELEFONO - 1]),
      telefonoAlt: crmFmt(f[CRM_COL_TELEFONO_ALT - 1]), documento: crmFmt(f[CRM_COL_DOCUMENTO - 1]),
      avance: f[COL_AVANCE - 1], estado: crmFmt(f[COL_ESTADO - 1]),
      vendedor: crmFmt(f[CRM_COL_VENDEDOR - 1]), supervisorVenta: crmFmt(f[CRM_COL_SUPERVISOR_VTA - 1]),
      formaPago: crmFmt(f[CRM_COL_FORMA_PAGO - 1]), tipoPlan: crmFmt(f[CRM_COL_TIPO_PLAN - 1]),
      scoring: crmFmt(f[CRM_COL_SCORING - 1]),
      cuotas: f.slice(COL_C2 - 1, COL_C2 - 1 + CRM_CANT_CUOTAS).map(function (v) { return String(v || ''); }),
      notasBase: crmNotasBase_(encabezados, f),
      situacion: situacion,
      aliasBase: alias,
      responsableEmail: resp ? resp.email : '',
      responsableNombre: resp ? resp.nombre : (alias || '(sin asignar)'),
    });
  }
  return { ss: ss, planes: planes, casos: casos, gestionesMes: gestionesMes };
}

/** Columnas de texto libre de BASE (AH en adelante) que tengan algo, sin las de licitación. */
function crmNotasBase_(encabezados, fila) {
  const cols = [];
  for (let c = CRM_COL_PRIMERA_NOTA; c <= fila.length; c++) {
    if (CRM_COLS_LICITACION.indexOf(c) < 0) cols.push(c);
  }
  return crmColumnasBase_(encabezados, fila, cols);
}

/** [{columna: 'Título (col XX)', valor}] de las columnas (1-based) dadas que tengan algo. */
function crmColumnasBase_(encabezados, fila, columnas) {
  const out = [];
  columnas.forEach(function (c) {
    const v = fila[c - 1];
    if (v === '' || v === null || v === undefined) return;
    out.push({ columna: crmEtiquetaColumna_(encabezados, c), valor: crmFmt(v) });
  });
  return out;
}

function crmEtiquetaColumna_(encabezados, c) {
  const letra = crmLetraColumna_(c);
  const titulo = String(encabezados[c - 1] || '').trim();
  return titulo ? titulo + ' (col ' + letra + ')' : 'Col ' + letra;
}

/** true si la etiqueta de una nota importada corresponde a una columna de licitación. */
function crmEsNotaLicitacion_(etiqueta) {
  return CRM_COLS_LICITACION.some(function (c) {
    return String(etiqueta).slice(-(' (col ' + crmLetraColumna_(c) + ')').length) === ' (col ' + crmLetraColumna_(c) + ')';
  });
}

/** 34 → "AH". */
function crmLetraColumna_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function crmBuscarPlan_(ctx, solicitud) {
  const p = ctx.planes.filter(function (x) { return x.solicitud === String(solicitud); })[0];
  if (!p) throw new Error('No encuentro la solicitud ' + solicitud + ' en BASE.');
  return p;
}

function crmPuedeVer_(u, plan) {
  return u.rol === CRM_ROL_SUPERVISOR || plan.responsableEmail === u.email;
}

/**
 * Usuario de la sesión (token que devolvió crmLogin). Exige usuario activo
 * y con la clave ya cambiada, salvo permitirClaveInicial (solo para
 * crmCambiarClave). Los errores de sesión empiezan con "SESION:" para que
 * la pantalla vuelva al login.
 */
function crmUsuarioActual_(token, permitirClaveInicial) {
  const cache = CacheService.getScriptCache();
  const ses = token ? JSON.parse(cache.get('ses_' + token) || 'null') : null;
  const epoca = ses ? Number(cache.get('epoca_' + ses.email) || 0) : 0;
  if (!ses || ses.t < epoca) throw new Error('SESION: tu sesión venció, volvé a ingresar.');
  const email = ses.email;
  const u = crmLeerUsuarios_(crmDatos_())
    .filter(function (x) { return x.email === email && x.activo; })[0];
  if (!u || !crmEmailValido_(u.email)) throw new Error('SESION: tu usuario no está habilitado en el CRM.');
  if (u.debeCambiarClave && !permitirClaveInicial) throw new Error('SESION: tenés que cambiar la clave.');
  cache.put('ses_' + token, JSON.stringify(ses), CRM_SESION_SEG); // renueva
  return u;
}

function crmExigirSupervisor_(token) {
  const u = crmUsuarioActual_(token);
  if (u.rol !== CRM_ROL_SUPERVISOR) throw new Error('Solo un supervisor puede hacer esto.');
  return u;
}

function crmLeerUsuarios_(ss) {
  return crmLeerObjetos_(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS).map(function (x) {
    return {
      email: String(x.Email || '').trim().toLowerCase(), nombre: String(x.Nombre || ''),
      rol: String(x.Rol || '').trim().toUpperCase(), alias: String(x.AliasBase || '').trim().toUpperCase(),
      activo: crmBool_(x.Activo),
      debeCambiarClave: crmBool_(x.DebeCambiarClave),
      tieneClave: !!x.ClaveHash,
    };
  }).filter(function (x) { return x.email; });
}

/** Actualiza (o crea) la fila de CRM_Casos de una solicitud con los campos dados. */
function crmActualizarCaso_(ss, solicitud, campos) {
  const hoja = crmHoja_(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS);
  const ultima = hoja.getLastRow();
  const ids = ultima > 1 ? hoja.getRange(2, 1, ultima - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const idx = ids.indexOf(String(solicitud));

  let fila;
  if (idx >= 0) {
    fila = hoja.getRange(idx + 2, 1, 1, CRM_ENC_CASOS.length).getValues()[0];
  } else {
    fila = CRM_ENC_CASOS.map(function () { return ''; });
    fila[0] = String(solicitud);
  }
  Object.keys(campos).forEach(function (k) {
    const col = CRM_ENC_CASOS.indexOf(k);
    if (col >= 0) fila[col] = campos[k];
  });
  if (idx >= 0) hoja.getRange(idx + 2, 1, 1, fila.length).setValues([fila]);
  else hoja.appendRow(fila);
}

function crmAuditar_(ss, email, accion, detalle) {
  crmHoja_(ss, CRM_HOJA_AUDITORIA, CRM_ENC_AUDITORIA).appendRow([new Date(), email, accion, detalle]);
}

/** Devuelve la hoja, creándola con encabezados si no existe. */
function crmHoja_(ss, nombre, encabezados) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

/** Filas de una hoja CRM como objetos {Encabezado: valor}. */
function crmLeerObjetos_(ss, nombre, encabezados) {
  const hoja = crmHoja_(ss, nombre, encabezados);
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, encabezados.length).getValues().map(function (f) {
    const o = {};
    encabezados.forEach(function (h, i) { o[h] = f[i]; });
    return o;
  });
}

function crmFilaResp_(nombre) {
  return {
    nombre: nombre, activa: 0, p1: 0, p2: 0, p34: 0, bajas: 0, prioGestionados: 0,
    prioSinGestion: 0, intentosMes: 0, efectivosMes: 0, regularizadosConGestion: 0,
    promesasVigentes: 0, promesasVencidas: 0,
  };
}

// --- Fechas: google.script.run no transporta Date, todo viaja como texto ---

function crmFmt(v, patron) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), patron || 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function crmHoyStr() {
  return crmFmt(new Date());
}

/** 'yyyy-MM-dd' (del input date) → Date a mediodía, o '' si viene vacío. */
function crmFecha(s) {
  if (!s) return '';
  const p = String(s).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12);
}

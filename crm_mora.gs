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
 * BASE llega por IMPORTRANGE y no se escribe nunca: todo lo del CRM vive en
 * hojas propias (CRM_*), ligado al plan por Solicitud.
 *
 * Instalación: ver sección "CRM" del README.
 */

// --- CONFIGURACIÓN CRM ---
const CRM_HOJA_USUARIOS = 'CRM_Usuarios';
const CRM_HOJA_CASOS = 'CRM_Casos';
const CRM_HOJA_GESTIONES = 'CRM_Gestiones';
const CRM_HOJA_AUDITORIA = 'CRM_Auditoria';

// Columnas de BASE que usa el CRM además de las de tablero_mora.gs.
const CRM_COL_RESPONSABLE = 1;   // A — EZE / FABIO / SANTI
const CRM_COL_SOLICITUD = 2;     // B
const CRM_COL_GRUPO = 3;         // C
const CRM_COL_ORDEN = 4;         // D
const CRM_COL_MODELO = 5;        // E
const CRM_COL_SOBREPAUTA = 6;    // F
const CRM_COL_CLIENTE = 8;       // H
const CRM_COL_TELEFONO = 9;      // I
const CRM_COL_TELEFONO_ALT = 10; // J
const CRM_COL_DOCUMENTO = 11;    // K
const CRM_COL_VENDEDOR = 16;     // P
const CRM_COL_SUPERVISOR_VTA = 17; // Q
const CRM_COL_FORMA_PAGO = 31;   // AE
const CRM_COL_TIPO_PLAN = 32;    // AF
const CRM_COL_SCORING = 33;      // AG
const CRM_COL_PRIMERA_NOTA = 34; // AH en adelante: notas / seguimiento en texto libre
const CRM_CANT_CUOTAS = 13;      // C2..C14

const CRM_ROL_SUPERVISOR = 'SUPERVISOR';
const CRM_ROL_RESPONSABLE = 'RESPONSABLE';

// Usuarios que se cargan la primera vez que se corre crmInicializar().
// El que corre la inicialización queda además como SUPERVISOR.
const CRM_USUARIOS_INICIALES = [
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
const CRM_MOTIVOS_NO_PAGO = ['', 'Problemas económicos', 'Olvido', 'Esperaba adjudicar',
  'Disconforme con aumento de cuota', 'Problema con débito / medio de pago',
  'Garantías rechazadas', 'Quiere renunciar', 'Otro'];

// Situación de cada plan según BASE (ver crmSituacion). "orden" define la
// prioridad en la cola de trabajo.
const CRM_SITUACIONES = {
  P1: { orden: 1, etiqueta: 'P1 · Atrasada (mide Fiat)' },
  P2: { orden: 2, etiqueta: 'P2 · Cuota del mes (mide Fiat)' },
  P3: { orden: 3, etiqueta: 'P3 · Atrasada' },
  P4: { orden: 4, etiqueta: 'P4 · Cuota del mes' },
  AL_DIA: { orden: 5, etiqueta: 'Al día' },
  RESCINDIDO: { orden: 8, etiqueta: 'Rescindido' },
  BAJA: { orden: 9, etiqueta: 'Baja' },
};

const CRM_ENC_USUARIOS = ['Email', 'Nombre', 'Rol', 'AliasBase', 'Activo'];
const CRM_ENC_CASOS = ['Solicitud', 'ResponsableEmail', 'EstadoCaso', 'UltimaGestion',
  'UltimoCanal', 'UltimoResultado', 'ProximoContacto', 'PromesaFecha', 'PromesaMonto',
  'MotivoNoPago', 'ActualizadoPor'];
const CRM_ENC_GESTIONES = ['FechaHora', 'UsuarioEmail', 'UsuarioNombre', 'Solicitud', 'Canal',
  'Resultado', 'EstadoCaso', 'MotivoNoPago', 'PromesaFecha', 'PromesaMonto',
  'ProximoContacto', 'Nota', 'AvanceAlMomento', 'SituacionAlMomento'];
const CRM_ENC_AUDITORIA = ['FechaHora', 'UsuarioEmail', 'Accion', 'Detalle'];

// ---------------------------------------------------------------------------
// Web App
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('crm_index').evaluate()
    .setTitle('CRM Mora')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Crea las hojas CRM_* (si no existen) y carga los usuarios iniciales. Quien
 * lo corre queda como SUPERVISOR. Se puede correr más de una vez: no
 * duplica usuarios ni borra datos.
 */
function crmInicializar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  crmHoja(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS);
  crmHoja(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS);
  crmHoja(ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES);
  crmHoja(ss, CRM_HOJA_AUDITORIA, CRM_ENC_AUDITORIA);

  const existentes = crmLeerUsuarios(ss).map(function (u) { return u.email; });
  const nuevos = CRM_USUARIOS_INICIALES.slice();
  const yo = Session.getEffectiveUser().getEmail().toLowerCase();
  if (yo) nuevos.unshift({ email: yo, nombre: yo, rol: CRM_ROL_SUPERVISOR, alias: '' });

  const hoja = ss.getSheetByName(CRM_HOJA_USUARIOS);
  nuevos.forEach(function (u) {
    if (existentes.indexOf(u.email) >= 0) return;
    hoja.appendRow([u.email, u.nombre, u.rol, u.alias, true]);
    existentes.push(u.email);
  });
  crmAuditar(ss, yo, 'Inicializar CRM', 'Hojas y usuarios iniciales verificados');
  SpreadsheetApp.getUi().alert('CRM listo. Ahora publicalo: Implementar > Nueva implementación > Aplicación web (ver README).');
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
function crmInicio() {
  const u = crmUsuarioActual();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    usuario: u,
    esSupervisor: u.rol === CRM_ROL_SUPERVISOR,
    responsables: u.rol === CRM_ROL_SUPERVISOR
      ? crmLeerUsuarios(ss).filter(function (x) { return x.activo; })
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
function crmListarCartera() {
  const u = crmUsuarioActual();
  const ctx = crmContexto();
  const hoy = crmHoyStr();
  return ctx.planes
    .filter(function (p) { return crmPuedeVer(u, p); })
    .map(function (p) {
      const c = ctx.casos[p.solicitud] || {};
      const g = ctx.gestionesMes[p.solicitud] || { intentos: 0, efectivos: 0 };
      return {
        solicitud: p.solicitud, cliente: p.cliente, telefono: p.telefono, grupo: p.grupo,
        orden: p.orden, avance: p.avance, estadoBase: p.estado,
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
function crmFichaPlan(solicitud) {
  const u = crmUsuarioActual();
  const ctx = crmContexto();
  const p = crmBuscarPlan(ctx, solicitud);
  if (!crmPuedeVer(u, p)) throw new Error('No tenés acceso a este plan.');

  const gestiones = crmLeerObjetos(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES)
    .filter(function (g) { return String(g.Solicitud) === p.solicitud; })
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
function crmRegistrarGestion(solicitud, datos) {
  const u = crmUsuarioActual();
  const ctx = crmContexto();
  const p = crmBuscarPlan(ctx, solicitud);
  if (!crmPuedeVer(u, p)) throw new Error('No tenés acceso a este plan.');
  if (CRM_CANALES.indexOf(datos.canal) < 0) throw new Error('Canal inválido.');
  if (CRM_ESTADOS_CASO.indexOf(datos.estadoCaso) < 0) throw new Error('Estado inválido.');
  if (datos.estadoCaso === 'Promesa de pago' && !datos.promesaFecha) {
    throw new Error('Para "Promesa de pago" cargá la fecha prometida.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ahora = new Date();
    const hojaG = crmHoja(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES);
    hojaG.appendRow([
      ahora, u.email, u.nombre, p.solicitud, datos.canal, datos.resultado || '',
      datos.estadoCaso, datos.motivo || '', crmFecha(datos.promesaFecha), datos.promesaMonto || '',
      crmFecha(datos.proximoContacto), datos.nota || '', p.avance, p.situacion.codigo,
    ]);
    crmActualizarCaso(ctx.ss, p.solicitud, {
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
  return crmFichaPlan(solicitud);
}

/** Supervisor: reasigna uno o varios planes a un responsable. */
function crmReasignar(solicitudes, email) {
  const u = crmExigirSupervisor();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destino = crmLeerUsuarios(ss).filter(function (x) { return x.email === email && x.activo; })[0];
  if (!destino) throw new Error('Usuario destino inexistente o inactivo.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    solicitudes.forEach(function (s) {
      crmActualizarCaso(ss, String(s), { ResponsableEmail: email, ActualizadoPor: u.email });
    });
    crmAuditar(ss, u.email, 'Reasignar', solicitudes.length + ' plan(es) → ' + email + ': ' + solicitudes.join(', '));
  } finally {
    lock.releaseLock();
  }
  return solicitudes.length;
}

/** Supervisor: lista de usuarios. */
function crmListarUsuarios() {
  crmExigirSupervisor();
  return crmLeerUsuarios(SpreadsheetApp.getActiveSpreadsheet());
}

/** Supervisor: alta o modificación de un usuario (clave = email). */
function crmGuardarUsuario(datos) {
  const u = crmExigirSupervisor();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = String(datos.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+$/.test(email)) throw new Error('Email inválido.');
  if ([CRM_ROL_SUPERVISOR, CRM_ROL_RESPONSABLE].indexOf(datos.rol) < 0) throw new Error('Rol inválido.');

  const usuarios = crmLeerUsuarios(ss);
  const quedanSupervisores = usuarios.filter(function (x) {
    return x.email !== email && x.activo && x.rol === CRM_ROL_SUPERVISOR;
  }).length + (datos.activo && datos.rol === CRM_ROL_SUPERVISOR ? 1 : 0);
  if (quedanSupervisores === 0) throw new Error('Tiene que quedar al menos un supervisor activo.');

  const fila = [email, datos.nombre || email, datos.rol,
    String(datos.alias || '').trim().toUpperCase(), !!datos.activo];
  const hoja = crmHoja(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS);
  const idx = usuarios.map(function (x) { return x.email; }).indexOf(email);
  if (idx >= 0) hoja.getRange(idx + 2, 1, 1, fila.length).setValues([fila]);
  else hoja.appendRow(fila);
  crmAuditar(ss, u.email, idx >= 0 ? 'Modificar usuario' : 'Alta usuario', JSON.stringify(fila));
  return crmLeerUsuarios(ss);
}

/** Supervisor: números para el tablero. */
function crmTableroSupervisor() {
  crmExigirSupervisor();
  const ctx = crmContexto();
  const hoy = crmHoyStr();

  // Por responsable
  const porResp = {};
  crmLeerUsuarios(ctx.ss).forEach(function (x) {
    if (x.rol === CRM_ROL_RESPONSABLE && x.activo) porResp[x.email] = crmFilaResp(x.nombre);
  });
  const sinAsignar = crmFilaResp('(sin asignar)');

  ctx.planes.forEach(function (p) {
    if (p.responsableEmail && !porResp[p.responsableEmail]) {
      porResp[p.responsableEmail] = crmFilaResp(p.responsableNombre);
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
  crmLeerObjetos(ctx.ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES).forEach(function (g) {
    const dia = crmFmt(g.FechaHora);
    if (dia < desdeStr) return;
    const quien = g.UsuarioNombre || g.UsuarioEmail;
    actividad[dia] = actividad[dia] || {};
    actividad[dia][quien] = (actividad[dia][quien] || 0) + 1;
  });

  return { porResponsable: filas, proyeccion: proyeccion, actividad: actividad, hoy: hoy };
}

// ---------------------------------------------------------------------------
// Acceso a datos
// ---------------------------------------------------------------------------

/** Lee BASE + CRM_Casos + gestiones del mes y arma los planes resueltos. */
function crmContexto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = ss.getSheetByName(HOJA_BASE);
  if (!hojaBase) throw new Error('No encuentro la hoja "' + HOJA_BASE + '"');
  const valores = hojaBase.getDataRange().getValues();
  const encabezados = valores[0];

  const usuarios = crmLeerUsuarios(ss);
  const porAlias = {};
  const porEmail = {};
  usuarios.forEach(function (x) {
    porEmail[x.email] = x;
    if (x.alias) porAlias[x.alias] = x;
  });

  const casos = {};
  crmLeerObjetos(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS).forEach(function (c) {
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
  crmLeerObjetos(ss, CRM_HOJA_GESTIONES, CRM_ENC_GESTIONES).forEach(function (g) {
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
      cliente: crmFmt(f[CRM_COL_CLIENTE - 1]), telefono: crmFmt(f[CRM_COL_TELEFONO - 1]),
      telefonoAlt: crmFmt(f[CRM_COL_TELEFONO_ALT - 1]), documento: crmFmt(f[CRM_COL_DOCUMENTO - 1]),
      avance: f[COL_AVANCE - 1], estado: crmFmt(f[COL_ESTADO - 1]),
      vendedor: crmFmt(f[CRM_COL_VENDEDOR - 1]), supervisorVenta: crmFmt(f[CRM_COL_SUPERVISOR_VTA - 1]),
      formaPago: crmFmt(f[CRM_COL_FORMA_PAGO - 1]), tipoPlan: crmFmt(f[CRM_COL_TIPO_PLAN - 1]),
      scoring: crmFmt(f[CRM_COL_SCORING - 1]),
      cuotas: f.slice(COL_C2 - 1, COL_C2 - 1 + CRM_CANT_CUOTAS).map(function (v) { return String(v || ''); }),
      notasBase: crmNotasBase(encabezados, f),
      situacion: situacion,
      aliasBase: alias,
      responsableEmail: resp ? resp.email : '',
      responsableNombre: resp ? resp.nombre : (alias || '(sin asignar)'),
    });
  }
  return { ss: ss, planes: planes, casos: casos, gestionesMes: gestionesMes };
}

/** Columnas de texto libre de BASE (AH en adelante) que tengan algo. */
function crmNotasBase(encabezados, fila) {
  const notas = [];
  for (let c = CRM_COL_PRIMERA_NOTA - 1; c < fila.length; c++) {
    const v = fila[c];
    if (v === '' || v === null) continue;
    notas.push({ columna: String(encabezados[c] || '').trim() || ('Col ' + (c + 1)), valor: crmFmt(v) });
  }
  return notas;
}

function crmBuscarPlan(ctx, solicitud) {
  const p = ctx.planes.filter(function (x) { return x.solicitud === String(solicitud); })[0];
  if (!p) throw new Error('No encuentro la solicitud ' + solicitud + ' en BASE.');
  return p;
}

function crmPuedeVer(u, plan) {
  return u.rol === CRM_ROL_SUPERVISOR || plan.responsableEmail === u.email;
}

function crmUsuarioActual() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) {
    throw new Error('No pude identificar tu cuenta de Google. La app tiene que estar publicada ' +
      'desde una cuenta del mismo dominio que los usuarios (ver README).');
  }
  const u = crmLeerUsuarios(SpreadsheetApp.getActiveSpreadsheet())
    .filter(function (x) { return x.email === email && x.activo; })[0];
  if (!u) throw new Error('El usuario ' + email + ' no está habilitado en el CRM. Pedíselo a un supervisor.');
  return u;
}

function crmExigirSupervisor() {
  const u = crmUsuarioActual();
  if (u.rol !== CRM_ROL_SUPERVISOR) throw new Error('Solo un supervisor puede hacer esto.');
  return u;
}

function crmLeerUsuarios(ss) {
  return crmLeerObjetos(ss, CRM_HOJA_USUARIOS, CRM_ENC_USUARIOS).map(function (x) {
    return {
      email: String(x.Email || '').trim().toLowerCase(), nombre: String(x.Nombre || ''),
      rol: String(x.Rol || '').trim().toUpperCase(), alias: String(x.AliasBase || '').trim().toUpperCase(),
      activo: x.Activo === true || String(x.Activo).toUpperCase() === 'TRUE',
    };
  }).filter(function (x) { return x.email; });
}

/** Actualiza (o crea) la fila de CRM_Casos de una solicitud con los campos dados. */
function crmActualizarCaso(ss, solicitud, campos) {
  const hoja = crmHoja(ss, CRM_HOJA_CASOS, CRM_ENC_CASOS);
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

function crmAuditar(ss, email, accion, detalle) {
  crmHoja(ss, CRM_HOJA_AUDITORIA, CRM_ENC_AUDITORIA).appendRow([new Date(), email, accion, detalle]);
}

/** Devuelve la hoja, creándola con encabezados si no existe. */
function crmHoja(ss, nombre, encabezados) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

/** Filas de una hoja CRM como objetos {Encabezado: valor}. */
function crmLeerObjetos(ss, nombre, encabezados) {
  const hoja = crmHoja(ss, nombre, encabezados);
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, encabezados.length).getValues().map(function (f) {
    const o = {};
    encabezados.forEach(function (h, i) { o[h] = f[i]; });
    return o;
  });
}

function crmFilaResp(nombre) {
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

/**
 * AINHOA TAREAS — Backend completo
 * Este archivo va como Tareas.gs dentro de tu mismo proyecto de Apps Script
 * (el mismo donde está Código.gs). Reemplazá TODO lo que tengas en tu
 * Tareas.gs actual por este contenido completo.
 *
 * Recordá que en Código.gs tenés que tener estas líneas, antes de
 * "return json({ error: 'Accion no reconocida' });":
 *
 *   if (action === 'listarTareas') return listarTareasHandler(ss, e);
 *   if (action === 'crearTarea') return crearTareaHandler(ss, e);
 *   if (action === 'tomarTarea') return tomarTareaHandler(ss, e);
 *   if (action === 'marcarHecha') return marcarHechaHandler(ss, e);
 *   if (action === 'dejarPosta') return dejarPostaHandler(ss, e);
 *   if (action === 'agregarComentario') return agregarComentarioHandler(ss, e);
 *   if (action === 'checkPin') return checkPinHandler(ss, e);
 */

function generarIdTarea_(ss) {
  return 'TSK-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function listarTareasHandler(ss, e) {
  const persona = e.parameter.persona;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const filas = datos.slice(1);

  const tareas = filas.map(fila => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = fila[i]);
    if (obj.FechaLimite) obj.FechaLimite = fechaComoTexto_(obj.FechaLimite);
    if (obj.FechaCompletada) obj.FechaCompletada = fechaHoraComoTexto_(obj.FechaCompletada);
    if (obj.TimerInicio) obj.TimerInicio = fechaHoraComoTexto_(obj.TimerInicio);
    return obj;
  }).filter(t => t.ID);

  if (persona === 'Angie') {
    return json({ ok: true, tareas: tareas });
  }

  const filtradas = tareas.filter(t =>
    t.Responsable === persona || t.Tipo === 'General'
  );
  return json({ ok: true, tareas: filtradas });
}

function crearTareaHandler(ss, e) {
  const sheet = ss.getSheetByName('Tareas');
  const id = generarIdTarea_(ss);

  sheet.appendRow([
    id,
    e.parameter.titulo || '',
    e.parameter.descripcion || '',
    e.parameter.categoria || '',
    e.parameter.tipo || 'General',
    e.parameter.responsable || '',
    'Pendiente',
    e.parameter.prioridad || 'Media',
    e.parameter.fechaLimite || '',
    e.parameter.esRecurrente || 'NO',
    e.parameter.diaRecurrencia || '',
    e.parameter.requiereFoto || 'NO',
    '',
    e.parameter.checklist || '[]',
    '',
    'NO',
    '',
    '',
    0,
    '',
    e.parameter.duracionEstimada || '',
    e.parameter.creadoPor || ''
  ]);

  return json({ ok: true, id: id });
}

function editarTareaHandler(ss, e) {
  const id = e.parameter.id;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  const campos = {
    Titulo: e.parameter.titulo,
    Descripcion: e.parameter.descripcion,
    Categoria: e.parameter.categoria,
    Prioridad: e.parameter.prioridad,
    FechaLimite: e.parameter.fechaLimite,
    DuracionEstimada: e.parameter.duracionEstimada
  };
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      Object.keys(campos).forEach(campo => {
        if (campos[campo] !== undefined && campos[campo] !== '') {
          const col = headers.indexOf(campo);
          if (col > -1) sheet.getRange(i + 1, col + 1).setValue(campos[campo]);
        }
      });
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function eliminarTareaHandler(ss, e) {
  const id = e.parameter.id;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.deleteRow(i + 1);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function tomarTareaHandler(ss, e) {
  const id = e.parameter.id;
  const persona = e.parameter.persona;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  const colResponsable = headers.indexOf('Responsable');

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colResponsable + 1).setValue(persona);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function marcarHechaHandler(ss, e) {
  const id = e.parameter.id;
  const persona = e.parameter.persona;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  const colEstado = headers.indexOf('Estado');
  const colHechaPor = headers.indexOf('HechaPor');
  const colFechaCompletada = headers.indexOf('FechaCompletada');

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colEstado + 1).setValue('Hecha');
      sheet.getRange(i + 1, colHechaPor + 1).setValue(persona);
      sheet.getRange(i + 1, colFechaCompletada + 1).setValue(new Date());
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function dejarPostaHandler(ss, e) {
  const id = e.parameter.id;
  const nota = e.parameter.nota;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  const colComentario = headers.indexOf('Comentario');
  const colEsPosta = headers.indexOf('EsPosta');

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colComentario + 1).setValue(nota);
      sheet.getRange(i + 1, colEsPosta + 1).setValue('SI');
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function agregarComentarioHandler(ss, e) {
  const id = e.parameter.id;
  const texto = e.parameter.texto;
  const sheet = ss.getSheetByName('Tareas');
  const datos = sheet.getDataRange().getValues();
  const headers = datos[0];
  const colID = headers.indexOf('ID');
  const colComentario = headers.indexOf('Comentario');

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colComentario + 1).setValue(texto);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function checkPinHandler(ss, e) {
  var persona = e.parameter.persona;
  var pin = e.parameter.pin;
  var configSheet = ss.getSheetByName('Config');
  var configData = configSheet.getDataRange().getValues();
  var campoBuscado = 'PIN_' + persona;
  var pinGuardado = '';
  for (var i = 1; i < configData.length; i++) {
    if (String(configData[i][0]).trim() === campoBuscado) {
      pinGuardado = String(configData[i][1]).trim();
      break;
    }
  }
  var ok = pinGuardado && pinGuardado === String(pin).trim();
  return json({ ok: ok });
}

/* ======================= TIMER ======================= */
function iniciarTimerHandler(ss, e) {
  var id = e.parameter.id;
  var sheet = ss.getSheetByName('Tareas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colInicio = headers.indexOf('TimerInicio');
  var colEstadoTimer = headers.indexOf('TimerEstado');
  var colEstado = headers.indexOf('Estado');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colInicio + 1).setValue(new Date().toISOString());
      sheet.getRange(i + 1, colEstadoTimer + 1).setValue('Corriendo');
      sheet.getRange(i + 1, colEstado + 1).setValue('En curso');
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function pausarTimerHandler(ss, e) {
  var id = e.parameter.id;
  var sheet = ss.getSheetByName('Tareas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colInicio = headers.indexOf('TimerInicio');
  var colAcum = headers.indexOf('TimerAcumuladoSeg');
  var colEstadoTimer = headers.indexOf('TimerEstado');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      var inicioStr = datos[i][colInicio];
      var acumulado = parseInt(datos[i][colAcum]) || 0;
      if (inicioStr) {
        var inicio = new Date(inicioStr);
        acumulado += Math.floor((new Date() - inicio) / 1000);
      }
      sheet.getRange(i + 1, colAcum + 1).setValue(acumulado);
      sheet.getRange(i + 1, colInicio + 1).setValue('');
      sheet.getRange(i + 1, colEstadoTimer + 1).setValue('Pausado');
      return json({ ok: true, acumuladoSeg: acumulado });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

function finalizarConTiempoHandler(ss, e) {
  var id = e.parameter.id;
  var persona = e.parameter.persona;
  var tiempoManual = e.parameter.tiempoManual;
  var sheet = ss.getSheetByName('Tareas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colInicio = headers.indexOf('TimerInicio');
  var colAcum = headers.indexOf('TimerAcumuladoSeg');
  var colEstadoTimer = headers.indexOf('TimerEstado');
  var colTiempoReal = headers.indexOf('TiempoRealMin');
  var colEstado = headers.indexOf('Estado');
  var colHechaPor = headers.indexOf('HechaPor');
  var colFechaCompletada = headers.indexOf('FechaCompletada');

  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      var minutos;
      if (tiempoManual) {
        minutos = parseInt(tiempoManual);
      } else {
        var acumulado = parseInt(datos[i][colAcum]) || 0;
        var inicioStr = datos[i][colInicio];
        if (inicioStr) {
          acumulado += Math.floor((new Date() - new Date(inicioStr)) / 1000);
        }
        minutos = Math.round(acumulado / 60);
      }
      sheet.getRange(i + 1, colTiempoReal + 1).setValue(minutos);
      sheet.getRange(i + 1, colInicio + 1).setValue('');
      sheet.getRange(i + 1, colEstadoTimer + 1).setValue('Detenido');
      sheet.getRange(i + 1, colEstado + 1).setValue('Hecha');
      sheet.getRange(i + 1, colHechaPor + 1).setValue(persona);
      sheet.getRange(i + 1, colFechaCompletada + 1).setValue(new Date());
      return json({ ok: true, minutos: minutos });
    }
  }
  return json({ ok: false, error: 'Tarea no encontrada' });
}

/* ======================= ESTADÍSTICAS (solo Angie) ======================= */
function estadisticasHandler(ss, e) {
  var sheet = ss.getSheetByName('Tareas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var filas = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    return obj;
  }).filter(function(t) { return t.ID && t.Estado === 'Hecha' && t.TiempoRealMin; });

  var porTitulo = {};
  filas.forEach(function(t) {
    var key = t.Titulo;
    if (!porTitulo[key]) porTitulo[key] = { titulo: key, categoria: t.Categoria, cantidad: 0, sumaReal: 0, sumaEstimada: 0 };
    porTitulo[key].cantidad++;
    porTitulo[key].sumaReal += parseInt(t.TiempoRealMin) || 0;
    porTitulo[key].sumaEstimada += parseInt(t.DuracionEstimada) || 0;
  });
  var resultado = Object.keys(porTitulo).map(function(k) {
    var x = porTitulo[k];
    return {
      titulo: x.titulo, categoria: x.categoria, cantidad: x.cantidad,
      promedioReal: Math.round(x.sumaReal / x.cantidad),
      promedioEstimado: x.cantidad ? Math.round(x.sumaEstimada / x.cantidad) : 0
    };
  });
  return json({ ok: true, estadisticas: resultado });
}

/* ======================= NOTAS / POST-ITS ======================= */
function listarNotasHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Notas');
  if (!sheet) return json({ ok: true, notas: [] });
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var notas = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    return obj;
  }).filter(function(n) { return n.ID && (n.Visibilidad === 'Todas' || n.Visibilidad === persona || n.Autor === persona); });
  return json({ ok: true, notas: notas });
}

function crearNotaHandler(ss, e) {
  var sheet = ss.getSheetByName('Notas');
  var id = 'NOTE-' + Utilities.getUuid().split('-')[0].toUpperCase();
  sheet.appendRow([
    id,
    e.parameter.texto || '',
    e.parameter.autor || '',
    e.parameter.importancia || 'Media',
    e.parameter.visibilidad || 'Todas',
    new Date(),
    '[]',
    '[]'
  ]);
  return json({ ok: true, id: id });
}

function marcarNotaLeidaHandler(ss, e) {
  var id = e.parameter.id;
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Notas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colLeido = headers.indexOf('LeidoPor');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      var leido = [];
      try { leido = JSON.parse(datos[i][colLeido] || '[]'); } catch (err) { leido = []; }
      if (leido.indexOf(persona) === -1) leido.push(persona);
      sheet.getRange(i + 1, colLeido + 1).setValue(JSON.stringify(leido));
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Nota no encontrada' });
}

function responderNotaHandler(ss, e) {
  var id = e.parameter.id;
  var persona = e.parameter.persona;
  var texto = e.parameter.texto;
  var sheet = ss.getSheetByName('Notas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colResp = headers.indexOf('Respuestas');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      var resp = [];
      try { resp = JSON.parse(datos[i][colResp] || '[]'); } catch (err) { resp = []; }
      resp.push({ autor: persona, texto: texto, fecha: new Date().toISOString() });
      sheet.getRange(i + 1, colResp + 1).setValue(JSON.stringify(resp));
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'Nota no encontrada' });
}

/* ======================= CALENDARIO / EVENTOS ======================= */
function listarEventosHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Eventos');
  if (!sheet) return json({ ok: true, eventos: [] });
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var eventos = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    if (obj.Fecha) obj.Fecha = fechaComoTexto_(obj.Fecha);
    return obj;
  }).filter(function(ev) { return ev.ID && (ev.Visibilidad === 'Todas' || ev.Visibilidad === persona || ev.CreadoPor === persona); });
  return json({ ok: true, eventos: eventos });
}

function crearEventoHandler(ss, e) {
  var sheet = ss.getSheetByName('Eventos');
  var id = 'EVT-' + Utilities.getUuid().split('-')[0].toUpperCase();
  sheet.appendRow([
    id,
    e.parameter.titulo || '',
    e.parameter.fecha || '',
    e.parameter.visibilidad || 'Todas',
    e.parameter.creadoPor || ''
  ]);
  return json({ ok: true, id: id });
}

/* ======================= JORNADA ======================= */
function iniciarJornadaHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Jornadas');
  var tz = Session.getScriptTimeZone();
  var hoy = new Date();
  var fechaStr = Utilities.formatDate(hoy, tz, 'yyyy-MM-dd');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colPersona = headers.indexOf('Persona');
  var colFecha = headers.indexOf('Fecha');
  var colSalida = headers.indexOf('HoraSalida');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colPersona] === persona && datos[i][colFecha] === fechaStr && !datos[i][colSalida]) {
      return json({ ok: false, error: 'Ya hay una jornada abierta hoy' });
    }
  }
  var id = 'JOR-' + Utilities.getUuid().split('-')[0].toUpperCase();
  sheet.appendRow([id, persona, fechaStr, hoy.toISOString(), '', '']);
  return json({ ok: true, id: id, horaEntrada: hoy.toISOString() });
}

function finalizarJornadaHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Jornadas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colPersona = headers.indexOf('Persona');
  var colSalida = headers.indexOf('HoraSalida');
  var colEntrada = headers.indexOf('HoraEntrada');
  var colDuracion = headers.indexOf('DuracionMin');
  for (var i = datos.length - 1; i >= 1; i--) {
    if (datos[i][colPersona] === persona && !datos[i][colSalida]) {
      var ahora = new Date();
      var entrada = new Date(datos[i][colEntrada]);
      var minutos = Math.round((ahora - entrada) / 60000);
      sheet.getRange(i + 1, colSalida + 1).setValue(ahora.toISOString());
      sheet.getRange(i + 1, colDuracion + 1).setValue(minutos);
      return json({ ok: true, minutos: minutos });
    }
  }
  return json({ ok: false, error: 'No hay jornada abierta' });
}

function fechaComoTexto_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}
function fechaHoraComoTexto_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.toISOString();
  }
  return v;
}
function normalizarHeader_(h) {
  return String(h || '').trim().toLowerCase()
    .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u');
}

// FUNCIÓN DE UN SOLO USO — recorre la pestaña Tareas y, si encuentra dos o
// más filas con el mismo ID (por ejemplo TSK-0005 repetido, de antes de que
// generarIdTarea_ generara IDs únicos), le asigna un ID nuevo a cada
// duplicado a partir del segundo. Correr una sola vez desde el editor de
// Apps Script: elegir esta función en el desplegable de arriba y Ejecutar ▶.
function corregirIdsDuplicadosTareas() {
  var ss = SpreadsheetApp.openById('1y7gcnhXDbqA17pHO1-5QODP_-YOi_Vc-6qWCGkq6EWM');
  var sheet = ss.getSheetByName('Tareas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var vistos = {};
  var corregidos = 0;
  for (var i = 1; i < datos.length; i++) {
    var id = datos[i][colID];
    if (!id) continue;
    if (vistos[id]) {
      var nuevoId = 'TSK-' + Utilities.getUuid().split('-')[0].toUpperCase();
      sheet.getRange(i + 1, colID + 1).setValue(nuevoId);
      corregidos++;
      Logger.log('Fila ' + (i + 1) + ': ' + id + ' -> ' + nuevoId);
    } else {
      vistos[id] = true;
    }
  }
  Logger.log('IDs duplicados corregidos: ' + corregidos);
}

function listarJornadasHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Jornadas');
  if (!sheet) return json({ ok: true, jornadas: [] });
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var jornadas = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    obj.Fecha = fechaComoTexto_(obj.Fecha);
    obj.HoraEntrada = obj.HoraEntrada ? fechaHoraComoTexto_(obj.HoraEntrada) : obj.HoraEntrada;
    obj.HoraSalida = obj.HoraSalida ? fechaHoraComoTexto_(obj.HoraSalida) : obj.HoraSalida;
    return obj;
  }).filter(function(j) { return j.ID && (persona === 'Angie' || j.Persona === persona); });
  return json({ ok: true, jornadas: jornadas });
}

function crearJornadaManualHandler(ss, e) {
  var persona = e.parameter.persona;
  var fecha = e.parameter.fecha;
  var horaEntrada = e.parameter.horaEntrada;
  var horaSalida = e.parameter.horaSalida;
  var sheet = ss.getSheetByName('Jornadas');
  var id = 'JOR-' + Utilities.getUuid().split('-')[0].toUpperCase();
  var entradaISO = new Date(fecha + 'T' + horaEntrada + ':00').toISOString();
  var salidaISO = horaSalida ? new Date(fecha + 'T' + horaSalida + ':00').toISOString() : '';
  var minutos = horaSalida ? Math.round((new Date(salidaISO) - new Date(entradaISO)) / 60000) : '';
  sheet.appendRow([id, persona, fecha, entradaISO, salidaISO, minutos]);
  return json({ ok: true, id: id });
}

function editarJornadaHandler(ss, e) {
  var id = e.parameter.id;
  var fecha = e.parameter.fecha;
  var horaEntrada = e.parameter.horaEntrada;
  var horaSalida = e.parameter.horaSalida;
  var sheet = ss.getSheetByName('Jornadas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colFecha = headers.indexOf('Fecha');
  var colEntrada = headers.indexOf('HoraEntrada');
  var colSalida = headers.indexOf('HoraSalida');
  var colDuracion = headers.indexOf('DuracionMin');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      var f = fecha || datos[i][colFecha];
      var entradaISO = horaEntrada ? new Date(f + 'T' + horaEntrada + ':00').toISOString() : datos[i][colEntrada];
      var salidaISO = horaSalida ? new Date(f + 'T' + horaSalida + ':00').toISOString() : datos[i][colSalida];
      var minutos = (entradaISO && salidaISO) ? Math.round((new Date(salidaISO) - new Date(entradaISO)) / 60000) : '';
      sheet.getRange(i + 1, colFecha + 1).setValue(f);
      sheet.getRange(i + 1, colEntrada + 1).setValue(entradaISO);
      sheet.getRange(i + 1, colSalida + 1).setValue(salidaISO);
      sheet.getRange(i + 1, colDuracion + 1).setValue(minutos);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'No encontrada' });
}

function eliminarJornadaHandler(ss, e) {
  var id = e.parameter.id;
  var sheet = ss.getSheetByName('Jornadas');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.deleteRow(i + 1);
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'No encontrada' });
}

/* ======================= REGISTRO DE PRODUCCIÓN ======================= */
function crearRegistroHandler(ss, e) {
  var sheet = ss.getSheetByName('Registros');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[normalizarHeader_(h)] = i; });
  var id = 'REG-' + Utilities.getUuid().split('-')[0].toUpperCase();
  var fila = new Array(headers.length).fill('');
  var valores = {
    id: id,
    persona: e.parameter.persona || '',
    fecha: e.parameter.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    area: e.parameter.area || '',
    cantidad: e.parameter.cantidad || '',
    unidad: e.parameter.unidad || ''
  };
  Object.keys(valores).forEach(function(k) { if (idx[k] !== undefined) fila[idx[k]] = valores[k]; });
  sheet.appendRow(fila);
  return json({ ok: true, id: id });
}

function listarRegistrosHandler(ss, e) {
  var persona = e.parameter.persona;
  var sheet = ss.getSheetByName('Registros');
  if (!sheet) return json({ ok: true, registros: [] });
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[normalizarHeader_(h)] = i; });
  var registros = datos.slice(1).map(function(fila) {
    return {
      ID: idx.id !== undefined ? fila[idx.id] : '',
      Persona: idx.persona !== undefined ? fila[idx.persona] : '',
      Fecha: idx.fecha !== undefined ? fechaComoTexto_(fila[idx.fecha]) : '',
      Area: idx.area !== undefined ? fila[idx.area] : '',
      Cantidad: idx.cantidad !== undefined ? fila[idx.cantidad] : '',
      Unidad: idx.unidad !== undefined ? fila[idx.unidad] : ''
    };
  }).filter(function(r) { return r.ID && (persona === 'Angie' || r.Persona === persona); });
  return json({ ok: true, registros: registros });
}

/* ======================= BANDEJA DE ENTRADA ======================= */
function crearBandejaHandler(ss, e) {
  var sheet = ss.getSheetByName('Bandeja');
  var id = 'BAN-' + Utilities.getUuid().split('-')[0].toUpperCase();
  sheet.appendRow([id, e.parameter.texto || '', e.parameter.autor || '', new Date(), 'NO']);
  return json({ ok: true, id: id });
}

function listarBandejaHandler(ss, e) {
  var sheet = ss.getSheetByName('Bandeja');
  if (!sheet) return json({ ok: true, items: [] });
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var items = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    return obj;
  }).filter(function(b) { return b.ID && b.Procesado !== 'SI'; });
  return json({ ok: true, items: items });
}

function procesarBandejaHandler(ss, e) {
  var id = e.parameter.id;
  var sheet = ss.getSheetByName('Bandeja');
  var datos = sheet.getDataRange().getValues();
  var headers = datos[0];
  var colID = headers.indexOf('ID');
  var colProc = headers.indexOf('Procesado');
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][colID] === id) {
      sheet.getRange(i + 1, colProc + 1).setValue('SI');
      return json({ ok: true });
    }
  }
  return json({ ok: false, error: 'No encontrado' });
}

/* ======================= DASHBOARD (solo Angie) ======================= */
function dashboardHandler(ss, e) {
  var tz = Session.getScriptTimeZone();
  var hoyStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var tareasSheet = ss.getSheetByName('Tareas');
  var datos = tareasSheet.getDataRange().getValues();
  var headers = datos[0];
  var tareas = datos.slice(1).map(function(fila) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = fila[i]; });
    return obj;
  }).filter(function(t) { return t.ID; });

  var pendientes = tareas.filter(function(t) { return t.Estado !== 'Hecha'; }).length;
  var atrasadas = tareas.filter(function(t) { return t.Estado !== 'Hecha' && t.FechaLimite && String(t.FechaLimite).slice(0,10) < hoyStr; }).length;
  var sinResp = tareas.filter(function(t) { return !t.Responsable && t.Estado !== 'Hecha'; }).length;
  var completadasHoy = tareas.filter(function(t) {
    if (t.Estado !== 'Hecha' || !t.FechaCompletada) return false;
    return Utilities.formatDate(new Date(t.FechaCompletada), tz, 'yyyy-MM-dd') === hoyStr;
  }).length;

  var haceUnaSemana = new Date(); haceUnaSemana.setDate(haceUnaSemana.getDate() - 7);
  var completadasSemana = tareas.filter(function(t) {
    return t.Estado === 'Hecha' && t.FechaCompletada && new Date(t.FechaCompletada) > haceUnaSemana;
  }).length;

  var trabajando = [];
  var jornadasSheet = ss.getSheetByName('Jornadas');
  if (jornadasSheet) {
    var jdatos = jornadasSheet.getDataRange().getValues();
    var jheaders = jdatos[0];
    var colPersona = jheaders.indexOf('Persona');
    var colFecha = jheaders.indexOf('Fecha');
    var colSalida = jheaders.indexOf('HoraSalida');
    var colEntrada = jheaders.indexOf('HoraEntrada');
    for (var i = 1; i < jdatos.length; i++) {
      if (jdatos[i][colFecha] === hoyStr && !jdatos[i][colSalida]) {
        trabajando.push({ persona: jdatos[i][colPersona], desde: jdatos[i][colEntrada] });
      }
    }
  }

  var metaSemanal = 0;
  var configSheet = ss.getSheetByName('Config');
  if (configSheet) {
    var cdatos = configSheet.getDataRange().getValues();
    for (var c = 1; c < cdatos.length; c++) {
      if (String(cdatos[c][0]).trim() === 'MetaSemanalTareas') { metaSemanal = parseInt(cdatos[c][1]) || 0; break; }
    }
  }

  return json({
    ok: true, pendientes: pendientes, atrasadas: atrasadas, sinResponsable: sinResp,
    completadasHoy: completadasHoy, completadasSemana: completadasSemana,
    metaSemanal: metaSemanal, trabajando: trabajando
  });
}

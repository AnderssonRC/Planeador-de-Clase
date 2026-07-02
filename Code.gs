/**
 * API del Planeador de Clases — Res Cogitans
 * ---------------------------------------------------------------
 * Este script YA NO sirve el HTML (doGet ya no usa HtmlService).
 * Ahora es una API JSON pura que el frontend (alojado en GitHub
 * Pages) consulta por fetch(). Los datos siguen viviendo en esta
 * misma hoja de cálculo — no cambia nada de cómo se guardan.
 *
 * Lecturas (GET)  -> ?action=getMaterias
 *                     ?action=getActividades&grado=Once
 * Escrituras (POST) -> body JSON: { action: 'guardarActividad', ... }
 *                       (enviado con Content-Type: text/plain para
 *                       evitar el preflight CORS que Apps Script no maneja)
 *
 * Después de pegar este código:
 * 1) Ejecuta prepararColumnas() una vez (si no lo has hecho ya).
 * 2) Implementar → Administrar implementaciones → editar → Nueva versión.
 * 3) Copia la URL que termina en /exec y pégala en API_URL dentro
 *    del Index.html del frontend.
 * 4) Verifica que la implementación tenga:
 *      Ejecutar como: Yo (tu cuenta)
 *      Quién tiene acceso: Cualquiera
 */

function doGet(e) {
  var action = e.parameter.action;
  try {
    var resultado;
    if (action === 'getMaterias') {
      resultado = getMaterias();
    } else if (action === 'getActividades') {
      resultado = getActividades(e.parameter.grado);
    } else {
      resultado = { ok: false, error: 'Acción GET no reconocida: ' + action };
    }
    return jsonOut(resultado);
  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var resultado;
    if (action === 'guardarActividad') {
      resultado = guardarActividad(body);
    } else if (action === 'actualizarEstado') {
      resultado = actualizarEstado(body.grado, body.id, body.estado);
    } else if (action === 'eliminarActividad') {
      resultado = eliminarActividad(body.grado, body.id);
    } else if (action === 'editarActividad') {
      resultado = editarActividad(body);
    } else {
      resultado = { ok: false, error: 'Acción POST no reconocida: ' + action };
    }
    return jsonOut(resultado);
  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * MIGRACIÓN — ejecutar UNA SOLA VEZ desde el editor de Apps Script:
 * Selecciona "prepararColumnas" en el menú de funciones y pulsa Ejecutar.
 * Agrega los encabezados Objetivo (col I) y PalabrasClave (col J)
 * a las hojas Noveno, Décimo y Once si no existen todavía.
 */
function prepararColumnas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Noveno', 'Décimo', 'Once'].forEach(function(nombre) {
    const hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    if (hoja.getMaxColumns() < 10) {
      hoja.insertColumnsAfter(hoja.getMaxColumns(), 10 - hoja.getMaxColumns());
    }
    if (!hoja.getRange(1, 9).getValue()) hoja.getRange(1, 9).setValue('Objetivo');
    if (!hoja.getRange(1, 10).getValue()) hoja.getRange(1, 10).setValue('PalabrasClave');
  });
}

function getMaterias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Materias');
  const datos = hoja.getDataRange().getValues();
  const materias = [];
  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][0]) continue;
    materias.push({
      nombre: datos[i][0],
      color: datos[i][1],
      grados: datos[i][2]
    });
  }
  return materias;
}

function getActividades(grado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(grado);
  if (!hoja) return [];
  const datos = hoja.getDataRange().getValues();
  const actividades = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '' || datos[i][0] === null || datos[i][0] === undefined) continue;
    const fecha = datos[i][1];
    let fechaStr = '';
    if (fecha instanceof Date) {
      const y = fecha.getFullYear();
      const m = String(fecha.getMonth() + 1).padStart(2, '0');
      const d = String(fecha.getDate()).padStart(2, '0');
      fechaStr = y + '-' + m + '-' + d;
    } else {
      fechaStr = String(fecha).trim().split('T')[0];
    }
    actividades.push({
      id: datos[i][0],
      fecha: fechaStr,
      materia: datos[i][2],
      titulo: datos[i][3],
      descripcion: datos[i][4],
      estado: datos[i][5],
      archivos: datos[i][7],
      objetivo: datos[i][8] || '',
      keywords: datos[i][9] || ''
    });
  }
  return actividades;
}

/**
 * Genera un ID único: el mayor ID numérico existente + 1.
 * (Evita duplicados que producía getLastRow() tras eliminar filas.)
 */
function siguienteId(hoja) {
  const datos = hoja.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < datos.length; i++) {
    const n = Number(datos[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function escribirFila(hoja, id, datos) {
  hoja.appendRow([
    id,
    datos.fecha,
    datos.materia,
    datos.titulo,
    datos.descripcion,
    datos.estado,
    false,
    datos.archivos || '',
    datos.objetivo || '',
    datos.keywords || ''
  ]);
  const fila = hoja.getLastRow();
  hoja.getRange(fila, 2).setNumberFormat('@STRING@');
  hoja.getRange(fila, 7).insertCheckboxes();
}

function guardarActividad(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(datos.grado);
  if (!hoja) return { ok: false, error: 'Grado no encontrado' };
  const id = siguienteId(hoja);
  escribirFila(hoja, id, datos);
  return { ok: true, id: id };
}

function actualizarEstado(grado, id, estado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(grado);
  if (!hoja) return { ok: false, error: 'Grado no encontrado' };
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) == String(id)) {
      hoja.getRange(i + 1, 6).setValue(estado);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ID no encontrado' };
}

function eliminarActividad(grado, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(grado);
  if (!hoja) return { ok: false, error: 'Grado no encontrado' };
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) == String(id)) {
      hoja.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ID no encontrado' };
}

function editarActividad(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gradoOriginal = String(datos.gradoOriginal).trim();
    const gradoNuevo = String(datos.grado).trim();

    const hojaOriginal = ss.getSheetByName(gradoOriginal);
    if (!hojaOriginal) return { ok: false, error: 'Hoja no encontrada: ' + gradoOriginal };

    const filas = hojaOriginal.getDataRange().getValues();
    let filaEncontrada = -1;
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) == String(datos.id)) {
        filaEncontrada = i + 1;
        break;
      }
    }
    if (filaEncontrada === -1) return { ok: false, error: 'ID no encontrado: ' + datos.id };

    if (gradoOriginal === gradoNuevo) {
      hojaOriginal.getRange(filaEncontrada, 2).setValue(datos.fecha);
      hojaOriginal.getRange(filaEncontrada, 2).setNumberFormat('@STRING@');
      hojaOriginal.getRange(filaEncontrada, 3).setValue(datos.materia);
      hojaOriginal.getRange(filaEncontrada, 4).setValue(datos.titulo);
      hojaOriginal.getRange(filaEncontrada, 5).setValue(datos.descripcion);
      hojaOriginal.getRange(filaEncontrada, 6).setValue(datos.estado);
      hojaOriginal.getRange(filaEncontrada, 8).setValue(datos.archivos || '');
      hojaOriginal.getRange(filaEncontrada, 9).setValue(datos.objetivo || '');
      hojaOriginal.getRange(filaEncontrada, 10).setValue(datos.keywords || '');
      return { ok: true, id: datos.id };
    } else {
      hojaOriginal.deleteRow(filaEncontrada);
      const hojaNueva = ss.getSheetByName(gradoNuevo);
      if (!hojaNueva) return { ok: false, error: 'Hoja nueva no encontrada: ' + gradoNuevo };
      const nuevoId = siguienteId(hojaNueva);
      escribirFila(hojaNueva, nuevoId, datos);
      return { ok: true, id: nuevoId };
    }
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

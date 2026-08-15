/**
 * KARAOKE - Backend en Google Apps Script
 *
 * INSTALACIÓN (una sola vez):
 * 1. Crea un Google Sheet nuevo con DOS hojas (pestañas):
 *
 *    Hoja "Catalogo" con esta fila de cabecera en la fila 1:
 *      ID | Titulo | Artista | Idioma | Categoria | Imagen_URL | Enlace_Letra
 *    (Idioma debe ser exactamente "ES" o "EXT" en cada fila. Categoria: "General"
 *    o "Infantil". Imagen_URL y Enlace_Letra son opcionales.)
 *    Pega aquí el contenido de la pestaña "Catalogo" de catalogo_karaoke.xlsx
 *    (o de tu Excel, siempre que respete estas columnas).
 *
 *    Hoja "Solicitudes" con esta fila de cabecera en la fila 1:
 *      IdSolicitud | Timestamp | Titulo | Artista | Solicitante | Estado | Imagen_URL
 *    (déjala vacía, se rellena sola cuando la gente pide canciones)
 *
 * 2. En el Sheet: Extensiones > Apps Script. Borra el contenido de
 *    Code.gs que aparece por defecto y pega TODO este archivo.
 *
 * 3. Pulsa "Implementar" > "Nueva implementación":
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 *    Copia la URL que termina en /exec. Esa es la URL que hay que
 *    pegar en CONFIG.API_URL dentro de index.html y admin.html.
 *
 * 4. Si más adelante cambias el código, tienes que "Gestionar
 *    implementaciones" > editar > Nueva versión, para que se actualice
 *    la URL ya publicada.
 */

var SHEET_CATALOGO = 'Catalogo';
var SHEET_SOLICITUDES = 'Solicitudes';

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'catalogo') return respondJSON(leerCatalogo());
  if (action === 'solicitudes') return respondJSON(leerSolicitudes());
  return respondJSON({ error: 'Accion no reconocida' });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;

  if (action === 'nueva_solicitud') {
    crearSolicitud(body);
    return respondJSON({ ok: true });
  }
  if (action === 'marcar_cantada') {
    cambiarEstado(body.idSolicitud, 'Cantada');
    return respondJSON({ ok: true });
  }
  if (action === 'eliminar_solicitud') {
    eliminarSolicitud(body.idSolicitud);
    return respondJSON({ ok: true });
  }
  if (action === 'nueva_cancion') {
    crearCancion(body);
    return respondJSON({ ok: true });
  }
  if (action === 'reemplazar_catalogo') {
    var total = reemplazarCatalogo(body);
    return respondJSON({ ok: true, total: total });
  }
  return respondJSON({ error: 'Accion no reconocida' });
}

function respondJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hojaComoObjetos(nombreHoja) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var filas = [];
  for (var i = 1; i < values.length; i++) {
    var fila = {};
    var vacia = true;
    for (var j = 0; j < headers.length; j++) {
      fila[headers[j]] = values[i][j];
      if (values[i][j] !== '' && values[i][j] !== null) vacia = false;
    }
    if (!vacia) {
      fila._row = i + 1; // fila real en el sheet (1-indexed), útil para editar/borrar
      filas.push(fila);
    }
  }
  return filas;
}

function leerCatalogo() {
  return hojaComoObjetos(SHEET_CATALOGO);
}

function leerSolicitudes() {
  var filas = hojaComoObjetos(SHEET_SOLICITUDES);
  // Las más recientes primero
  filas.sort(function (a, b) {
    return new Date(b.Timestamp) - new Date(a.Timestamp);
  });
  return filas;
}

function crearSolicitud(body) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SOLICITUDES);
  var idSolicitud = new Date().getTime() + '-' + Math.floor(Math.random() * 10000);
  sheet.appendRow([
    idSolicitud,
    new Date(),
    body.titulo,
    body.artista,
    body.solicitante,
    'Pendiente',
    body.imagenUrl || ''
  ]);
}

function cambiarEstado(idSolicitud, nuevoEstado) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SOLICITUDES);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(idSolicitud)) {
      sheet.getRange(i + 1, 6).setValue(nuevoEstado); // columna 6 = Estado
      break;
    }
  }
}

function eliminarSolicitud(idSolicitud) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SOLICITUDES);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(idSolicitud)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function crearCancion(body) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGO);
  var nuevoId = new Date().getTime();
  sheet.appendRow([
    nuevoId,
    body.titulo,
    body.artista,
    body.idioma || 'ES',
    body.categoria || 'General',
    body.imagenUrl || '',
    body.enlaceLetra || ''
  ]);
}

// Sustituye TODO el contenido de la hoja Catalogo (menos la cabecera) por
// el array de canciones recibido. Se usa para sincronizar de una vez el
// catálogo completo desde el Excel, sin copiar/pegar a mano.
function reemplazarCatalogo(body) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOGO);
  var canciones = body.canciones || [];

  var filasActuales = sheet.getLastRow();
  if (filasActuales > 1) {
    sheet.getRange(2, 1, filasActuales - 1, sheet.getLastColumn()).clearContent();
  }
  if (canciones.length === 0) return 0;

  var filas = canciones.map(function (c) {
    return [
      c.ID, c.Titulo, c.Artista, c.Idioma, c.Categoria,
      c.Imagen_URL || '', c.Enlace_Letra || ''
    ];
  });
  sheet.getRange(2, 1, filas.length, 7).setValues(filas);
  return filas.length;
}

// Funciones Comunes de Acceso a Datos
// =========================================================================

/**
 * @summary Abre y devuelve la hoja de cálculo principal por su ID.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} El objeto de la hoja de cálculo.
 */
function getSpreadsheet() {
  // Asume que ID_INVENTARIO está definido globalmente
  return SpreadsheetApp.openById(ID_INVENTARIO);
}

/**
 * @summary Obtiene una hoja por su nombre. Si no existe, la crea con cabeceras predefinidas (específico para la hoja de usuarios).
 * @param {string} sheetName El nombre de la hoja a obtener o crear.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto de la hoja.
 */
function getSheetByName(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Definir las cabeceras si la hoja es de usuarios
    if (sheetName === HOJA_USUARIO) {
      sheet.appendRow([
        'Id',
        'NombreCompleto',
        'userName',
        'password',
        'CDE',
        'Email',
        'Estado_User',
        'Rol',
        'Fecha de Registro',
      ]);
    }
  }
  return sheet;
}

/**
 * @summary Obtiene los datos de una hoja y los filtra para devolver solo las filas con estado "Activo".
 * @param {string} sheetName El nombre de la hoja.
 * @returns {string} Una cadena JSON con los datos de las filas activas.
 */
function getInventoryDataForSheet(sheetName) {
  Logger.log('getInventoryDataForSheet called with sheetName: ' + sheetName);
  const sheet = getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let estadoColumn = 'Estado_User'; // Por defecto para usuarios
    if (sheetName !== HOJA_USUARIO) estadoColumn = 'Estado'; // Asumir 'Estado' para otras hojas

    if (row[headers.indexOf(estadoColumn)] === 'Activo') {
      const item = {};
      for (let j = 0; j < headers.length; j++) {
        item[headers[j]] = row[j];
      }
      items.push(item);
    }
  }
  return JSON.stringify({ data: items });
}

/**
 * @summary Busca en una hoja el número de fila que corresponde a un ID único.
 * @param {string|number} id El identificador único del registro a buscar.
 * @param {string} sheetName El nombre de la hoja donde se realizará la búsqueda.
 * @returns {number} El número de la fila (base 1) si se encuentra, o -1 si no.
 */
function findRowById(id, sheetName) {
  const sheet = getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('Id');

  if (idColumnIndex === -1) return -1;

  for (let i = 1; i < data.length; i++) {
    // Usar == para permitir que coincida 'string' con 'number' si es necesario
    if (String(data[i][idColumnIndex]) === String(id)) {
      return i + 1; // +1 porque getRange() es 1-based
    }
  }
  return -1; // No encontrado
}

/**
 * @summary Obtiene la información completa de un registro por su ID en cualquier hoja.
 * @param {string|number} id El ID a buscar.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {object|null} Un objeto con los datos del registro o null si no lo encuentra.
 */
function getUsuario(id, sheetName) {
  const sheet = getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idColumnIndex = headers.indexOf('Id');
  if (idColumnIndex === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idColumnIndex] == id) {
      const info = {};
      headers.forEach((header, j) => {
        info[header] = data[i][j];
      });
      return info;
    }
  }
  return null;
}


/**
 * @summary Formatea una fecha en formato ISO a una cadena de fecha local (español).
 * @param {string} isoString La fecha en formato ISO.
 * @returns {string} La fecha formateada, o una cadena vacía si la entrada es nula.
 */
function formatDate(isoString) {
  if (!isoString) return '';
  const date = new Date(isoString);
  return date.toLocaleDateString('es-ES');
}
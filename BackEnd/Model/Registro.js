// BackEnd/Registro.js

/**
 * @summary Obtiene la hoja de 'Registro de Usuario'. Si no existe, la crea con las cabeceras adecuadas.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto de la hoja de registro.
 */
function getRegistroSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(getHojasConfig().REGISTRO.nombre);
  if (!sheet) {
    sheet = ss.insertSheet(getHojasConfig().REGISTRO.nombre);
    sheet.appendRow(["Id", "Fecha", "User Name", "Registro"]);
  }
  return sheet;
}
/**
 * @summary Obtiene todos los datos de la hoja de 'Registro de Usuario' en formato JSON.
 * @returns {string} Una cadena JSON que contiene todos los registros de la hoja.
 */
function getRegistroData() {
  var sheet = getRegistroSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Extraer encabezados
  if (!headers || headers.length === 0) {
    return JSON.stringify({ data: [] });
  }
  var items = data.map(function(row) {
    var item = {};
    headers.forEach(function(header, index) {
      item[header] = row[index];
    });
    return item;
  });
  return JSON.stringify({ data: items }); // Devolver el array de objetos directamente
}

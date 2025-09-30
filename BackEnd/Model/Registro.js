// BackEnd/Registro.js

/**
 * @summary Obtiene la hoja de 'Registro de Usuario'. Si no existe, la crea con las cabeceras adecuadas.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto de la hoja de registro.
 */
function getRegistroSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(getHojasConfig().REGISTRO);
  if (!sheet) {
    sheet = ss.insertSheet(getHojasConfig().REGISTRO);
    sheet.appendRow(['Id', 'Fecha', 'User Name', 'Registro']);
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
  var headers = data[0];
  var items = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var item = {};
    for (var j = 0; j < headers.length; j++) {
      item[headers[j]] = row[j];
    }
    items.push(item);
  }
  return JSON.stringify({ data: items });
}
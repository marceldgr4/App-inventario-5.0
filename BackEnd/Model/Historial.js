// BackEnd/Historial.js

/**
 * @summary Obtiene la hoja de 'Historial'. Si no existe, la crea con las cabeceras adecuadas.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto de la hoja de historial.
 */
function getHistorialSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(getHojasConfig().HISTORIAL);
  if (!sheet) {
    sheet = ss.insertSheet(getHojasConfig().HISTORIAL);
    sheet.appendRow([
      "Id",
      "ProductoId",
      "Fecha y Hora",
      "Producto",
      "Programa",
      "Unidades Anteriores",
      "Unidades Nuevas",
      "Estado",
      "Usuario",
      "Fecha de Retiro",
      "Cantidad",
    ]);
  }
  return sheet;
}

/**
 * @summary Obtiene todos los registros de la hoja de 'Historial' y los formatea como un arreglo de objetos.
 * @returns {Array<object>} Un arreglo de objetos, donde cada objeto es un registro del historial.
 */
function getHistorialDataActual() {
  var sheet = getHistorialSheet();
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return [];

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
  return items;
}

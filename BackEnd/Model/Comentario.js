// BackEnd/Comentario.js

/** @summary Obtiene el objeto de la hoja de 'Comentarios'. */
function getComentarioSheet() {
  return getSheet(getHojasConfig().COMENTARIOS.nombre);
}
/** @summary Obtiene todos los datos de la hoja de 'Comentarios' en formato JSON. */
function getComentarioData() {
  return JSON.stringify({ data: _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre) });
}
/** @summary Obtiene la información de un comentario específico por su ID. */
function getComentarioInfo(id) {
  return getInfo(id, getHojasConfig().COMENTARIOS.nombre);
}
/** @summary Agrega un nuevo comentario a la hoja de 'Comentarios'. */
function agregarComentario(data) {
  return agregar(data, getHojasConfig().COMENTARIOS.nombre);
}

/**
 * @summary Función auxiliar para obtener datos de la hoja 'Comentarios'.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {Array<object>} Los datos de la hoja.
 * @private
 */
function _getComentarioData(sheetName) {
  return _getInventoryDataForSheet(sheetName);
}

function guardarComentarioEnSheet(productoId, producto, programa, comentario, usuario) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTARIO); // 📌 pon aquí el ID real de tu spreadsheet
    const sheet = ss.getSheetByName(HOJA_COMENTARIOS);      // 📌 asegúrate que esta hoja exista

    const fecha = new Date();

     // --- Obtener el último Id ---
    let lastRow = sheet.getLastRow();
    let newId = 1; // valor por defecto si es el primer comentario

    if (lastRow > 1) { // asumiendo que la fila 1 son los encabezados
      const lastId = sheet.getRange(lastRow, 1).getValue(); // Columna A = Id
      if (!isNaN(lastId)) {
        newId = Number(lastId) + 1;
      }
    }

    // --- Agregar nueva fila ---
    sheet.appendRow([
      newId,                // Id secuencial
      productoId,
      producto,
      programa,
      fecha,
      comentario,
      usuario
    ]);

    return JSON.stringify({ success: true, id: newId });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}

// 👇 Función para generar un Id único si no lo tienes
function generarIdUnico() {
  return new Date().getTime(); // timestamp
}

// BackEnd/Comentario.js

/** @summary Obtiene el objeto de la hoja de 'Comentarios'. */
function getComentarioSheet() {
  return getSheet(getHojasConfig().COMENTARIOS);
}
/** @summary Obtiene todos los datos de la hoja de 'Comentarios' en formato JSON. */
function getComentarioData() {
  return JSON.stringify({ data: _getInventoryDataForSheet(getHojasConfig().COMENTARIOS) });
}
/** @summary Obtiene la información de un comentario específico por su ID. */
function getComentarioInfo(id) {
  return getInfo(id, getHojasConfig().COMENTARIOS);
}
/** @summary Agrega un nuevo comentario a la hoja de 'Comentarios'. */
function agregarComentario(data) {
  return agregar(data, getHojasConfig().COMENTARIOS);
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
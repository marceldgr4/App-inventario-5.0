// BackEnd/Decoracion.js

/** @summary Obtiene el objeto de la hoja de 'Decoracion'. */
function getDecoracionSheet() {
  return getSheet(getHojasConfig().DECORACION);
}
/** @summary Obtiene todos los datos de la hoja de 'Decoracion' en formato JSON. */
function getDecoracionData() {
  return JSON.stringify({ data: _getInventoryDataForSheet(getHojasConfig().DECORACION) });
}
/** @summary Obtiene la información de un artículo de decoración específico por su ID. */
function getDecoracionInfo(id) {
  return getInfo(id, getHojasConfig().DECORACION);
}
/** @summary Agrega un nuevo artículo a la hoja de 'Decoracion'. */
function agregarDecoracion(data) {
  return agregar(data, getHojasConfig().DECORACION);
}
/** @summary Actualiza un artículo existente en la hoja de 'Decoracion'. */
function actualizarDecoracion(data) {
  return actualizar(data, getHojasConfig().DECORACION);
}
/** @summary Elimina (desactiva) un artículo de la hoja de 'Decoracion'. */
function eliminarDecoracion(id) {
  return eliminar(id, getHojasConfig().DECORACION);
}
/** @summary Retira unidades de un artículo de la hoja de 'Decoracion'. */
function retirarDecoracion(id, unidades) {
  return retirar(id, unidades, getHojasConfig().DECORACION);
}

/**
 * @summary Función auxiliar para obtener datos de la hoja 'Decoracion'.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {Array<object>} Los datos de la hoja.
 * @private
 */
function _getDecoracionData(sheetName) {
  return _getInventoryDataForSheet(sheetName);
}
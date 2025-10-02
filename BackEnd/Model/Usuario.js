// BackEnd/Usuario.js

/** @summary Obtiene el objeto de la hoja de 'Usuario'. */
function getUsuarioSheet() {
  return getSheet(getHojasConfig().USUARIO);
}
/** @summary Obtiene todos los datos de la hoja de 'Usuario' en formato JSON. */
function getUsuarioData() {
  return JSON.stringify({
    data: _getInventoryDataForSheet(getHojasConfig().USUARIO),
  });
}
/** @summary Obtiene la información de un usuario específico por su ID. */
function getUsuarioInfo(id) {
  return getInfo(id, getHojasConfig().USUARIO);
}
/** @summary Agrega un nuevo usuario a la hoja de 'Usuario'. */
function agregarUsuario(data) {
  return agregar(data, getHojasConfig().USUARIO);
}
/** @summary Actualiza un usuario existente en la hoja de 'Usuario'. */
function actualizarUsuario(data) {
  return actualizar(data, getHojasConfig().USUARIO);
}
/** @summary Elimina (desactiva) un usuario de la hoja de 'Usuario'. */
function eliminarUsuario(id) {
  return eliminar(id, getHojasConfig().USUARIO);
}

/**
 * @summary Función auxiliar para obtener datos de la hoja 'Usuario'.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {Array<object>} Los datos de la hoja.
 * @private
 */
function _getUsuarioData(sheetName) {
  return _getInventoryDataForSheet(sheetName);
}

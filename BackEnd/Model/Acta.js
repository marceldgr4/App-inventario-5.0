// BackEnd/Acta.js

/**
 * @summary Obtiene todos los datos de la hoja de Acta.
 * @returns {string} Una cadena JSON que contiene los datos de la hoja solicitada.
 */
function getActaData() {
  return JSON.stringify({
    data: _getInventoryDataForSheet(getHojasConfig().ACTA),
  });
}

/**
 * @summary Obtiene la información de un producto de Acta por su ID.
 * @param {string|number} id El ID del producto.
 * @returns {object|null} Un objeto con la información del producto o null si no se encuentra.
 */
function getActaInfo(id) {
  return getInfo(id, getHojasConfig().ACTA);
}

/**
 * @summary Agrega un nuevo producto a la hoja de Acta.
 * @param {object} data Objeto con los datos del nuevo producto.
 * @returns {string} Resultado de la operación en formato JSON.
 */
function agregarProductoActa(data) {
  return agregar(data, getHojasConfig().ACTA);
}

/**
 * @summary Actualiza un producto existente en la hoja de Acta.
 * @param {object} data Objeto con los datos a actualizar, incluyendo el ID.
 * @returns {string} Resultado de la operación en formato JSON.
 */
function actualizarProductoActa(data) {
  return actualizar(data, getHojasConfig().ACTA);
}

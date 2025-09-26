// BackEnd/Decoracion.js

/** @summary Obtiene el objeto de la hoja de 'Decoracion'. */
function getDecoracionSheet() {
  return getSheet(getHojasConfig().DECORACION.nombre);
}
/** @summary Obtiene todos los datos de la hoja de 'Decoracion' en formato JSON. */
function getDecoracionData() {
  Logger.log("getDecoracionData: Solicitando datos de Decoracion.");
  try {
    const data = _getInventoryDataForSheet(getHojasConfig().DECORACION.nombre);
    Logger.log(`getDecoracionData: Se obtuvieron ${data.length} registros de Decoracion.`);
   return JSON.stringify({ data: data });
  } catch (e) {
    Logger.log(`ERROR: getDecoracionData - Fallo al obtener datos de Decoracion: ${e.message}`);
    return JSON.stringify({ data: [] });
  }
}
/** @summary Obtiene la información de un artículo de decoración específico por su ID. */
function getDecoracionInfo(id) {
  return getInfo(id, getHojasConfig().DECORACION.nombre);
}
/** @summary Agrega un nuevo artículo a la hoja de 'Decoracion'. */
function agregarDecoracion(data) {
  return agregar(data, getHojasConfig().DECORACION.nombre);
}
/** @summary Actualiza un artículo existente en la hoja de 'Decoracion'. */
function actualizarDecoracion(data) {
  return actualizar(data, getHojasConfig().DECORACION.nombre);
}
/** @summary Elimina (desactiva) un artículo de la hoja de 'Decoracion'. */
function eliminarDecoracion(id) {
  return eliminar(id, getHojasConfig().DECORACION.nombre);
}
/** @summary Retira unidades de un artículo de la hoja de 'Decoracion'. */
function retirarDecoracion(id, unidades) {
  return retirar(id, unidades, getHojasConfig().DECORACION.nombre);
}


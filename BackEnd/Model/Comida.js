/** @summary Obtiene el objeto de la hoja de 'Comida'. */
function getComidaSheet() {
  return getSheet(getHojasConfig().COMIDA.nombre); // Asegurarse de usar .nombre si getHojasConfig().COMIDA retorna un objeto
}
/** @summary Obtiene todos los datos de la hoja de 'Comida' en formato JSON. */
function getComidaData() {
  Logger.log("getComidaData: Solicitando datos de comida.");
  try {
    const data = _getInventoryDataForSheet(getHojasConfig().COMIDA.nombre);
    Logger.log(
      `getComidaData: Se obtuvieron ${data.length} registros de comida.`
    );
    return JSON.stringify({ data: data });
  } catch (e) {
    Logger.log(
      `ERROR: getComidaData - Fallo al obtener datos de comida: ${e.message}`
    );
    return JSON.stringify({ data: [], error: e.message });
  }
}
/** @summary Obtiene la información de un producto de comida específico por su ID. */
function getComidaInfo(id) {
  return getInfo(id, getHojasConfig().COMIDA.nombre);
}
/** @summary Agrega un nuevo producto a la hoja de 'Comida'. */
function agregarComida(data) {
  return agregar(data, getHojasConfig().COMIDA.nombre);
}
/** @summary Actualiza un producto existente en la hoja de 'Comida'. */
function actualizarComida(data) {
  return actualizar(data, getHojasConfig().COMIDA.nombre);
}
/** @summary Elimina (desactiva) un producto de la hoja de 'Comida'. */
function eliminarComida(id) {
  return eliminar(id, getHojasConfig().COMIDA.nombre);
}
/** @summary Retira unidades de un producto de la hoja de 'Comida'. */
function retirarComida(id, unidades) {
  return retirar(id, unidades, getHojasConfig().COMIDA.nombre);
}

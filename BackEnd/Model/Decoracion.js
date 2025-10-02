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
    Logger.log(
      `getDecoracionData: Se obtuvieron ${data.length} registros de Decoracion.`
    );
    return JSON.stringify({ data: data });
  } catch (e) {
    Logger.log(
      `ERROR: getDecoracionData - Fallo al obtener datos de Decoracion: ${e.message}`
    );
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
// ✅ Nueva función genérica para manejar retiros
function retirarProductoGenerico(sheetName, id, numUnidadesRetirar) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf("Id");
  const ingresosIndex = headers.indexOf("Ingresos");
  const salidasIndex = headers.indexOf("Salidas");
  const disponiblesIndex = headers.indexOf("UnidadesDisponibles");

  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return JSON.stringify({
      success: false,
      message: `No se encontró el producto con ID ${id} en ${sheetName}.`,
    });
  }

  const ingresos =
    parseFloat(sheet.getRange(rowIndex, ingresosIndex + 1).getValue()) || 0;
  const salidas =
    parseFloat(sheet.getRange(rowIndex, salidasIndex + 1).getValue()) || 0;

  let disponibles = parseFloat(
    sheet.getRange(rowIndex, disponiblesIndex + 1).getValue()
  );
  if (isNaN(disponibles)) {
    disponibles = ingresos - salidas;
  }

  if (numUnidadesRetirar > disponibles) {
    return JSON.stringify({
      success: false,
      message: `Error: No se pueden retirar ${numUnidadesRetirar} unidad(es). Solo hay ${disponibles} disponibles.`,
    });
  }

  const nuevasSalidas = salidas + numUnidadesRetirar;
  sheet.getRange(rowIndex, salidasIndex + 1).setValue(nuevasSalidas);

  const nuevosDisponibles = ingresos - nuevasSalidas;
  sheet.getRange(rowIndex, disponiblesIndex + 1).setValue(nuevosDisponibles);

  return JSON.stringify({
    success: true,
    message: `Se retiraron ${numUnidadesRetirar} unidad(es) correctamente.`,
    id: id,
    disponiblesAntes: disponibles,
    disponiblesDespues: nuevosDisponibles,
  });
}

function agregarComentarioDecoracion(idProducto, comentario) {
  try {
    const sheet = getDecoracionSheet();
    const data = sheet.getDataRange().getValues();

    // Buscar fila por ID en la columna A
    const rowIndex = data.findIndex((r) => r[0] == idProducto);
    if (rowIndex === -1) {
      return JSON.stringify({
        success: false,
        error: "Producto no encontrado",
      });
    }

    // 👉 índice de la columna "Comentarios" (ajústalo según tu hoja)
    const comentariosColIndex = 9;

    // Guardar solo el comentario en la columna de la hoja Decoracion
    sheet.getRange(rowIndex + 1, comentariosColIndex + 1).setValue(comentario);

    return JSON.stringify({
      success: true,
      producto: data[rowIndex][1], // nombre producto (columna B)
      comentario: comentario,
    });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.message });
  }
}

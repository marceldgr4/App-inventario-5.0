/**
 * @summary Obtiene los comentarios pendientes de confirmación para el usuario actual y un origen específico.
 * @param {string} origen - El origen para filtrar (ej. "Inventario comida").
 * @returns {string} JSON con los datos de las notificaciones pendientes.
 */
function getNotificacionesPendientesUsuario(origen) {
  try {
    const currentUserEmail = Session.getActiveUser()?.getEmail();
    if (!currentUserEmail) {
      return JSON.stringify({ data: [] });
    }
    let allComments = _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre);
    let pendientes = allComments.filter(item =>
      item.Usuario === currentUserEmail &&
      item.Respuesta && String(item.Respuesta).trim() !== '' &&
      (item['Borrado por Usuario'] === false || !item['Borrado por Usuario']) &&
      (item['RespuestaConfirmada'] === false || !item['RespuestaConfirmada'])
    );
    if (origen) {
      pendientes = pendientes.filter(item => item.Origen === origen);
    }
    return JSON.stringify({ data: pendientes });
  } catch (e) {
    Logger.log(`ERROR en getNotificacionesPendientesUsuario: ${e.stack}`);
    return JSON.stringify({ data: [], error: e.message });
  }
}

/**
 * @summary Obtiene los comentarios leídos (confirmados) para el usuario actual y un origen específico.
 * @param {string} origen - El origen para filtrar (ej. "Inventario comida").
 * @returns {string} JSON con los datos de las notificaciones leídas.
 */
function getNotificacionesLeidasUsuario(origen) {
  try {
    const currentUserEmail = Session.getActiveUser()?.getEmail();
    if (!currentUserEmail) {
      return JSON.stringify({ data: [] });
    }
    let allComments = _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre);
    let leidas = allComments.filter(item =>
      item.Usuario === currentUserEmail &&
      item.Respuesta && String(item.Respuesta).trim() !== '' &&
      (item['Borrado por Usuario'] === false || !item['Borrado por Usuario']) &&
      (item['RespuestaConfirmada'] === true)
    );
    if (origen) {
      leidas = leidas.filter(item => item.Origen === origen);
    }
    return JSON.stringify({ data: leidas });
  } catch (e) {
    Logger.log(`ERROR en getNotificacionesLeidasUsuario: ${e.stack}`);
    return JSON.stringify({ data: [], error: e.message });
  }
}
/**
 * @file BackEnd/Model/Comentario.js
 * @summary Centraliza toda la lógica de negocio para la gestión de comentarios.
 */

/**
 * @summary Obtiene la hoja de 'Comentarios'. Si no existe, la crea y se asegura de que tenga todas las columnas necesarias.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} El objeto de la hoja de cálculo.
 */
function getComentarioSheet() {
  const sheetName = getHojasConfig().COMENTARIOS.nombre;
  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  let sheet = ss.getSheetByName(sheetName);

  const expectedHeaders = [
    "Id", "ProductoId", "Producto", "Programa", "Fecha del Comentario", 
    "Comentario", "Usuario", "Origen", "Leido", "Respuesta", "Fecha De Respuesta",
    "Borrado por Usuario", "Borrado por Admin", "RespuestaConfirmada"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(expectedHeaders);
    Logger.log(`Hoja ${sheetName} creada con encabezados.`);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    let headerChanged = false;
    expectedHeaders.forEach(expectedHeader => {
      if (!headers.includes(expectedHeader)) {
        sheet.getRange(1, headers.length + 1).setValue(expectedHeader);
        headers.push(expectedHeader); // Add to our in-memory copy
        headerChanged = true;
      }
    });
    if (headerChanged) {
      Logger.log(`Columnas actualizadas en la hoja ${sheetName}.`);
    }
  }
  return sheet;
}

/**
 * @summary Obtiene todos los datos de la hoja de 'Comentarios' para la vista de administrador.
 * @returns {string} JSON con los datos.
 */
function getComentarioData() {
  try {
    const data = _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre);
    return JSON.stringify({ data: data });
  } catch (e) {
    Logger.log(`ERROR en getComentarioData: ${e.stack}`);
    return JSON.stringify({ data: [], error: e.message });
  }
}

/**
 * @summary Obtiene las notificaciones de comentarios para el usuario actual, opcionalmente filtradas por origen.
 * @description Devuelve los comentarios hechos por el usuario actual que han recibido una respuesta.
 * @param {string} origen - El origen para filtrar (ej. "Inventario comida").
 * @returns {string} JSON con los datos de las notificaciones.
 */
function getMisNotificacionesDeComentarios(origen) {
  try {
    const currentUserEmail = Session.getActiveUser()?.getEmail();
    if (!currentUserEmail) {
      return JSON.stringify({ data: [] }); // No hay usuario, no hay notificaciones
    }

    let allComments = _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre);

    let userNotifications = allComments.filter(item => 
        item.Usuario === currentUserEmail && // Comentario hecho por el usuario actual
        item.Respuesta && String(item.Respuesta).trim() !== '' && // Que tenga una respuesta
        (item['Borrado por Usuario'] === false || !item['Borrado por Usuario']) // Y que no esté borrado por el usuario
    );

    // Filtrar por origen si se proporciona
    if (origen) {
      userNotifications = userNotifications.filter(item => item.Origen === origen);
    }

    return JSON.stringify({ data: userNotifications });
  } catch (e) {
    Logger.log(`ERROR en getMisNotificacionesDeComentarios: ${e.stack}`);
    return JSON.stringify({ data: [], error: e.message });
  }
}


/**
 * @summary Obtiene la información de un comentario específico por su ID.
 * @param {string|number} id El ID del comentario.
 * @returns {string} Un objeto JSON con la información del comentario.
 */
function getComentarioInfo(id) {
  return getInfo(id, getHojasConfig().COMENTARIOS.nombre);
}

/**
 * @summary Actualiza la columna de comentarios de un producto en su hoja de inventario original.
 * @param {string} sheetName - El nombre de la hoja de inventario (ej. "Inventario comida").
 * @param {string|number} productoId - El ID del producto a actualizar.
 * @param {string} comentario - El nuevo texto del comentario.
 */
function actualizarComentarioEnHojaOrigen(sheetName, productoId, comentario) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`La hoja "${sheetName}" no fue encontrada.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const comentarioColIdx = headers.indexOf("COMENTARIOS");

    if (comentarioColIdx === -1) {
      Logger.log(`La hoja "${sheetName}" no tiene una columna "COMENTARIOS". No se actualizará.`);
      return; // No es un error fatal, simplemente no se puede actualizar.
    }

    const rowIndex = findRowById(productoId, sheetName);
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, comentarioColIdx + 1).setValue(comentario);
      Logger.log(`Comentario actualizado para el producto ID ${productoId} en la hoja ${sheetName}.`);
    } else {
      Logger.log(`No se encontró el producto con ID ${productoId} en la hoja ${sheetName} para actualizar el comentario.`);
    }
  } catch (e) {
    Logger.log(`Error en actualizarComentarioEnHojaOrigen: ${e.stack}`);
    // No relanzamos el error para no detener el proceso principal de registro de comentarios.
  }
}

/**
 * @summary Registra un nuevo comentario. Esta es la función principal y única para crear comentarios.
 * @param {object} data - Objeto con los datos del comentario.
 * @returns {string} Un objeto JSON indicando el éxito o fracaso.
 */
function registrarNuevoComentario(data) {
  const { sheetName, productoId, comentario, producto: productoDesdeCliente } = data;
  const activeUser = Session.getActiveUser();
  const usuario = activeUser ? activeUser.getEmail() : "Usuario desconocido";

  if (!sheetName || !productoId || !comentario) {
    return JSON.stringify({ success: false, message: "Faltan datos (sheetName, productoId, comentario)." });
  }

  try {
    actualizarComentarioEnHojaOrigen(sheetName, productoId, comentario);
    const productoInfo = getProductInfoGenerico(productoId, sheetName);
    
    const producto = productoInfo ? productoInfo.PRODUCTO : (productoDesdeCliente || "N/A");
    const programa = productoInfo ? productoInfo.PROGRAMA : sheetName;

    const commentSheet = getComentarioSheet();
    const newId = getNextId(commentSheet);

    commentSheet.appendRow([
      newId,          // Id
      productoId,     // ProductoId
      producto,       // Producto
      programa,       // Programa
      new Date(),     // Fecha del Comentario
      comentario,     // Comentario
      usuario,        // Usuario
      sheetName,      // Origen
      false,          // Leido
      "",             // Respuesta
      "",             // Fecha De Respuesta
      false,          // Borrado por Usuario
      false,          // Borrado por Admin
      false           // RespuestaConfirmada
    ]);

    return JSON.stringify({ success: true, message: "Comentario guardado exitosamente." });

  } catch (e) {
    Logger.log(`Error en registrarNuevoComentario: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Guarda la respuesta a un comentario existente.
 * @param {string|number} comentarioId - El ID del comentario que se está respondiendo.
 * @param {string} respuesta - El texto de la respuesta.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function responderComentario(comentarioId, respuesta) {
  if (!comentarioId || !respuesta) {
    return JSON.stringify({ success: false, message: "Faltan datos (comentarioId, respuesta)." });
  }

  try {
    const sheet = getComentarioSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    const respuestaColIdx = headers.indexOf("RESPUESTA");
    const fechaRespuestaColIdx = headers.indexOf("FECHA DE RESPUESTA");
    const confirmadaColIdx = headers.indexOf("RESPUESTACONFIRMADA");

    if (respuestaColIdx === -1 || fechaRespuestaColIdx === -1 || confirmadaColIdx === -1) {
      return JSON.stringify({ success: false, message: "No se encontraron las columnas necesarias (Respuesta, Fecha De Respuesta, RespuestaConfirmada)." });
    }

    const rowIndex = findRowById(comentarioId, sheet.getName());
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, respuestaColIdx + 1).setValue(respuesta);
      sheet.getRange(rowIndex, fechaRespuestaColIdx + 1).setValue(new Date());
      sheet.getRange(rowIndex, confirmadaColIdx + 1).setValue(false); // Marcar como no confirmada
      return JSON.stringify({ success: true, message: "Respuesta guardada exitosamente." });
    } else {
      return JSON.stringify({ success: false, message: "No se encontró el comentario para guardar la respuesta." });
    }
  } catch (e) {
    Logger.log(`Error en responderComentario: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Confirma que el usuario ha leído la respuesta a un comentario.
 * @param {string|number} comentarioId - El ID del comentario a confirmar.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function confirmarRespuesta(comentarioId) {
  if (!comentarioId) {
    return JSON.stringify({ success: false, message: "Falta el ID del comentario." });
  }

  try {
    const sheet = getComentarioSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    const confirmadaColIdx = headers.indexOf("RESPUESTACONFIRMADA");

    if (confirmadaColIdx === -1) {
      return JSON.stringify({ success: false, message: "No se encontró la columna 'RespuestaConfirmada'." });
    }

    const rowIndex = findRowById(comentarioId, sheet.getName());
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, confirmadaColIdx + 1).setValue(true);
      return JSON.stringify({ success: true, message: "Respuesta confirmada." });
    } else {
      return JSON.stringify({ success: false, message: "No se encontró el comentario." });
    }
  } catch (e) {
    Logger.log(`Error en confirmarRespuesta: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Marca un comentario como borrado por un administrador.
 * @param {string|number} comentarioId - El ID del comentario a marcar.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function marcarComentarioBorradoAdmin(comentarioId) {
  if (!comentarioId) {
    return JSON.stringify({ success: false, message: "Falta el ID del comentario." });
  }

  try {
    const sheet = getComentarioSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const borradoAdminColIdx = headers.indexOf("Borrado por Admin");

    if (borradoAdminColIdx === -1) {
      return JSON.stringify({ success: false, message: "No se encontró la columna 'Borrado por Admin'." });
    }

    const rowIndex = findRowById(comentarioId, sheet.getName());
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, borradoAdminColIdx + 1).setValue(true);
      return JSON.stringify({ success: true, message: "Comentario marcado como borrado." });
    } else {
      return JSON.stringify({ success: false, message: "No se encontró el comentario." });
    }
  } catch (e) {
    Logger.log(`Error en marcarComentarioBorradoAdmin: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Marca un comentario como borrado por un usuario.
 * @param {string|number} comentarioId - El ID del comentario a marcar.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function marcarComentarioBorradoUsuario(comentarioId) {
  if (!comentarioId) {
    return JSON.stringify({ success: false, message: "Falta el ID del comentario." });
  }

  try {
    const sheet = getComentarioSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const borradoUsuarioColIdx = headers.indexOf("Borrado por Usuario");

    if (borradoUsuarioColIdx === -1) {
      return JSON.stringify({ success: false, message: "No se encontró la columna 'Borrado por Usuario'." });
    }

    const rowIndex = findRowById(comentarioId, sheet.getName());
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, borradoUsuarioColIdx + 1).setValue(true);
      return JSON.stringify({ success: true, message: "Comentario marcado como borrado." });
    } else {
      return JSON.stringify({ success: false, message: "No se encontró el comentario." });
    }
  } catch (e) {
    Logger.log(`Error en marcarComentarioBorradoUsuario: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Marca uno o más comentarios como leídos.
 * @param {Array<string|number>} ids - Arreglo de IDs a marcar.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function marcarComentariosComoLeidos(ids) {
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return JSON.stringify({ success: false, message: "Se requiere un arreglo de IDs." });
  }

  try {
    const sheet = getComentarioSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const idIndex = headers.indexOf("ID");
    const leidoIndex = headers.indexOf("LEIDO");

    if (idIndex === -1 || leidoIndex === -1) {
      return JSON.stringify({ success: false, message: "No se encontraron las columnas 'Id' o 'Leido'." });
    }

    let actualizados = 0;
    const idsStr = ids.map(String);
    for (let i = 1; i < data.length; i++) {
      if (idsStr.includes(String(data[i][idIndex]))) {
        sheet.getRange(i + 1, leidoIndex + 1).setValue(true);
        actualizados++;
      }
    }

    return actualizados > 0
      ? JSON.stringify({ success: true, message: `${actualizados} comentario(s) marcado(s) como leído(s).` })
      : JSON.stringify({ success: false, message: "No se encontraron comentarios con los IDs proporcionados." });

  } catch (e) {
    Logger.log(`Error en marcarComentariosComoLeidos: ${e.stack}`);
    return JSON.stringify({ success: false, message: `Error del servidor: ${e.message}` });
  }
}

/**
 * @summary Elimina uno o más comentarios de la hoja central.
 * @param {Array<string|number>} ids - Arreglo de IDs a eliminar.
 * @returns {string} JSON indicando éxito o fracaso.
 */
function eliminarComentarios(ids) {
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return JSON.stringify({ success: false, message: "Se requiere un arreglo de IDs." });
  }
  return eliminarRegistrosPorIds(ids, getHojasConfig().COMENTARIOS.nombre);
}
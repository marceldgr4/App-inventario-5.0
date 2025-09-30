/** @summary Obtiene el objeto de la hoja de 'Papeleria'. */
function getPapeleriaSheet() {
  return getSheet(getHojasConfig().PAPELERIA.nombre);
}

/** @summary Obtiene todos los datos de la hoja de 'Papeleria' en formato JSON. */
function getPapeleriaData() {
  Logger.log("getPapeleriaData: Solicitando datos de papelería.");
  try {
    const data = _getInventoryDataForSheet(getHojasConfig().PAPELERIA.nombre);
    Logger.log(`getPapeleriaData: Se obtuvieron ${data.length} registros de papelería.`);
    return JSON.stringify({ success: true, data: data });
  } catch (e) {
    Logger.log(`ERROR: getPapeleriaData - Fallo al obtener datos de papelería: ${e.message}`);
    return JSON.stringify({ success: false, data: [], error: e.message });
  }
}

/** @summary Obtiene la información de un artículo de papelería específico por su ID. */
function getPapeleriaInfo(id) {
  return getInfo(id, getHojasConfig().PAPELERIA.nombre);
}

/** @summary Agrega un nuevo artículo a la hoja de 'Papeleria'. */
function agregarPapeleria(data) {
  return agregar(data, getHojasConfig().PAPELERIA.nombre);
}

/** @summary Agrega un producto con imagen (se guarda en Drive y se almacena el enlace). */
function agregarProductoConImagenDesdeArchivoPapeleria(productData, fileData) {
  try {
    const folderId = ID_PAPELERIA_IMG; // ⚠️ Reemplaza con el ID de tu carpeta en Drive
    const folder = DriveApp.getFolderById(folderId);

    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.base64Data),
      fileData.mimeType,
      fileData.fileName
    );

    const file = folder.createFile(blob);
    const imageUrl = file.getUrl();

    // Añadir el campo imagen
    productData.imagenAgregar = imageUrl;

    return agregarPapeleria(productData);
  } catch (e) {
    Logger.log("Error al guardar imagen: " + e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

/** @summary Actualiza un artículo existente en la hoja de 'Papeleria'. */
function actualizarPapeleria(data) {
  return actualizar(data, getHojasConfig().PAPELERIA.nombre);
}

/** @summary Retira unidades de un artículo de la hoja de 'Papeleria'. */
function retirarPapeleria(id, unidades) {
  return retirar(id, unidades, getHojasConfig().PAPELERIA.nombre);
}

/** @summary Elimina (desactiva) un artículo de la hoja de 'Papeleria'. */
function eliminarPapeleria(id) {
  return eliminar(id, getHojasConfig().PAPELERIA.nombre);
}

/** @summary Agregar comentario a un producto de Papelería. */
function agregarComentarioPapeleria(id, comentario) {
  try {
    const sheet = getPapeleriaSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Busca columnas
    const idIndex = headers.indexOf("Id") + 1;
    const productoIndex = headers.indexOf("PRODUCTO") + 1;
    const comentariosIndex = headers.indexOf("COMENTARIOS") + 1;

    if (idIndex === 0 || productoIndex === 0 || comentariosIndex === 0) {
      throw new Error("No se encontró columna Id, PRODUCTO o COMENTARIOS en la hoja.");
    }

    // Busca la fila que coincide con el Id
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex - 1]) === String(id)) {
        const producto = data[i][productoIndex - 1] || "Desconocido";
        const usuario = Session.getActiveUser().getEmail() || "Usuario anónimo";
        const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

        const cell = sheet.getRange(i + 1, comentariosIndex);
        const oldValue = cell.getValue();

        // Nuevo comentario formateado
        const nuevoComentario = `[${fecha}] ${usuario} sobre ${producto}: ${comentario}`;

        // Concatenar si ya había comentarios
        const valorFinal = oldValue ? oldValue + "\n" + nuevoComentario : nuevoComentario;

        cell.setValue(valorFinal);

        return JSON.stringify({
          success: true,
          message: "Comentario agregado correctamente.",
          id: id,
          producto: producto,
          usuario: usuario,
          fecha: fecha,
          comentario: comentario
        });
      }
    }

    throw new Error("Producto no encontrado con ID: " + id);

  } catch (e) {
    Logger.log("Error en agregarComentarioPapeleria: " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}


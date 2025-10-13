// =================================================================
// --- FUNCIONES DE GOOGLE SHEETS ---
// =================================================================

/**
 * Obtiene el siguiente ID autoincremental para una hoja de cálculo.
 * @param {Sheet} sheet La hoja de cálculo de Google Sheets.
 * @returns {number} El siguiente ID.
 */
function _obtenerSiguienteId(sheet) {
  if (!sheet) return 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;
  const lastId = parseInt(sheet.getRange(lastRow, 1).getValue());
  return isNaN(lastId) ? 1 : lastId + 1;
}

/**
 * Formatea una fecha a una cadena de texto estándar.
 * @param {Date|string} fecha La fecha a formatear.
 * @returns {string} La fecha formateada.
 */
function _formatearFecha(fecha) {
  const zonaHoraria = Session.getScriptTimeZone();
  return Utilities.formatDate(
    new Date(fecha),
    zonaHoraria,
    "yyyy-MM-dd HH:mm:ss"
  );
}

/**
 * Registra la información de un acta en la hoja de cálculo correspondiente.
 * @param {Sheet} hoja La hoja de cálculo donde se registrará el acta.
 * @param {object} data Los datos del acta a registrar.
 * @param {string} fileUrl La URL del archivo PDF del acta en Drive.
 * @param {string} userEmail El email del usuario que realiza la carga.
 * @returns {number} El ID del nuevo registro.
 */
/**
 * Registra la información de un acta en la hoja de cálculo correspondiente.
 * @param {Sheet} hoja La hoja de cálculo donde se registrará el acta.
 * @param {object} data Los datos del acta a registrar.
 * @param {string} fileUrl La URL del archivo PDF del acta en Drive.
 * @param {string} userId El ID del usuario activo (para la columna Id_Usuario).
 * @returns {number} El ID del nuevo registro.
 */
function _registrarActaEnHoja(hoja, data, fileUrl, userId) {
  const nuevoId = _obtenerSiguienteId(hoja);
  const fechaEntrega = _formatearFecha(data.fechaEntregaAgregar);
  const fechaCarga = _formatearFecha(data.fechaCargaPDFAgregar);

  // 🛑 ORDEN CORREGIDO según la lista de columnas:
  // [Id, Id_Usuario, Usuario, Fecha de entrega, Producto, Programa, Cantidad, Link, Ciudad, Nombre Compelto, fecha de carga]
  hoja.appendRow([
    nuevoId,                           // 1. Id
    userId || "",                      // 2. Id_Usuario (ID del usuario activo)
    data.nombreUsuarioAgregar,         // 3. Usuario
    fechaEntrega,                      // 4. Fecha de entrega
    data.productoAgregar,              // 5. Producto
    data.programaAgregar,              // 6. Programa
    data.cantidadAgregar,              // 7. Cantidad
    fileUrl,                           // 8. Link
    data.ciudadAgregar,                // 9. Ciudad
    data.nombreCompletoAgregar,        // 10. Nombre Compelto
    fechaCarga,                        // 11. fecha de carga
  ]);

  return nuevoId;
}

function subirActaConPDF(data) {
  const activeUser = getActiveUser();
  if (!activeUser) {
    throw new Error("Usuario no autenticado para subir acta.");
  }

  const fileData = {
    base64: data.base64PDF,
    name: data.nombreArchivo,
    type: "application/pdf",
  };

  const file = _subirArchivoADrive(ID_DRIVE_PDF, fileData);
  const fileUrl = file.getUrl();

  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  const hoja = ss.getSheetByName(HOJA_ACTA);
  if (!hoja) {
    throw new Error(`La hoja "${HOJA_ACTA}" no existe.`);
  }
  const nuevoId = _registrarActaEnHoja(hoja, data, fileUrl, activeUser.email);
 
  _registrarHistorialModificacion(
    nuevoId,
    `Acta: ${data.nombreArchivo}`,
    data.programaAgregar,
    null,
    data.cantidadAgregar,
    `Subida de acta: ${data.nombreArchivo}`,
    activeUser.name || activeUser.email,
    new Date(data.fechaEntregaAgregar),
    data.cantidadAgregar
  );

  return JSON.stringify({
    success: true,
    fileUrl: fileUrl,
    message: "Acta subida y registrada correctamente.",
  });
}

function _subirArchivoADrive(folderId, fileData) {
  if (!folderId) {
    throw new Error("El ID de la carpeta de Drive no puede ser nulo.");
  }
  
  const folder = DriveApp.getFolderById(folderId);
  if (!folder) {
    throw new Error(`Carpeta de Drive no encontrada con ID: ${folderId}`);
  }

  // 1. Decodificar el contenido Base64 a bytes
  const bytes = Utilities.base64Decode(fileData.base64);

  // 2. Crear un Blob (objeto de datos binarios grandes)
  const blob = Utilities.newBlob(bytes, fileData.type, fileData.name);

  // 3. Subir el Blob a la carpeta
  const file = folder.createFile(blob);
  
  return file;
}
/**
 * Obtiene la hoja de Historial. Si no existe, la crea con los encabezados correctos.
 * @returns {Sheet|null} El objeto de la hoja de cálculo o null si hay un error.
 */
function getHistorialModificacionesSheet() {
  const nombreHojaCorrecto = HOJA_HISTORIAL;

  Logger.log(
    `INFO: getHistorialModificacionesSheet - Intentando obtener/crear hoja de historial con el nombre: '${nombreHojaCorrecto}'`
  );

  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  let sheet = ss.getSheetByName(nombreHojaCorrecto);

  // Si la hoja no existe, la crea.
  if (!sheet) {
    Logger.log(
      `INFO: getHistorialModificacionesSheet - La hoja '${nombreHojaCorrecto}' no existe. Intentando crearla.`
    );
    try {
      sheet = ss.insertSheet(nombreHojaCorrecto);
      // Añade los encabezados a la nueva hoja.
      sheet.appendRow([
        "ID Historial",
        "ID Producto",
        "Fecha y Hora",
        "Producto",
        "Programa",
        "Unidades Anteriores",
        "Unidades Nuevas",
        "Acción/Estado",
        "Usuario",
        "Fecha de Entrega/Retiro",
        "Cantidad Entregada/Retirada",
        "Origen",
      ]);
    } catch (e) {
      return null;
    }
  }
  return sheet;
}

/**
 * @summary Obtiene el último ID de la hoja de historial de forma OPTIMIZADA.
 * @description Lee únicamente el valor de la celda del ID en la última fila,
 * que es el método más rápido para hojas que crecen secuencialmente.
 * @returns {number} El último ID numérico encontrado, o 0 si no hay datos.
 */
function _obtenerUltimoIdHistorialMod() {
  const sheet = getHistorialModificacionesSheet();
  const lastRow = sheet.getLastRow();

  // Si la hoja solo tiene la fila del encabezado (o está vacía), no hay IDs.
  const firstDataRow = (sheet.getFrozenRows() || 0) + 1;
  if (lastRow < firstDataRow) {
    return 0;
  }

  // Obtenemos el ID de la primera columna en la última fila.
  const lastIdValue = sheet.getRange(lastRow, 1).getValue();
  const lastId = parseInt(lastIdValue);

  // Si el valor no es un número (por ejemplo, está vacío o es texto), devolvemos 0.
  // De lo contrario, devolvemos el ID encontrado.
  return isNaN(lastId) ? 0 : lastId;
}

function _registrarHistorialModificacion(
  productoId,
  producto,
  programa,
  unidadesAnteriores,
  unidadesNuevas,
  accionEstado,
  usuario,
  entregaFecha,
  entregaCantidad,
  origen
) {
  try {
    var historialSheet = getHistorialModificacionesSheet();
    var nuevoIdHistorial = _obtenerUltimoIdHistorialMod() + 1;
    var now = new Date();

    // Asegura que los valores numéricos sean tratados como números, evitando errores.
    unidadesAnteriores = parseFloat(unidadesAnteriores);
    if (isNaN(unidadesAnteriores)) unidadesAnteriores = 0;

    unidadesNuevas = parseFloat(unidadesNuevas);
    if (isNaN(unidadesNuevas)) unidadesNuevas = 0;

    entregaCantidad = parseFloat(entregaCantidad);
    if (isNaN(entregaCantidad)) entregaCantidad = null;

    // Agrega una fila a la hoja de historial con todos los detalles de la modificación.
    historialSheet.appendRow([
      nuevoIdHistorial,
      productoId,
      now, // Fecha y hora del registro de historial
      producto || "",
      programa || "",
      unidadesAnteriores,
      unidadesNuevas,
      accionEstado || "",
      usuario || "Sistema", // Usuario que realiza la acción
      entregaFecha instanceof Date ? entregaFecha : null,
      entregaCantidad,
      origen || "",
    ]);
  } catch (e) {
    Logger.log(
      `Error al registrar historial de modificación para producto ID ${productoId}: ${e.toString()} Stack: ${
        e.stack
      }`
    );
  }
}

/**
 * Agrega un nuevo producto a la hoja de Decoración, subiendo una imagen si se proporciona.
 * @param {object} productData - Objeto con los datos del producto a agregar.
 * @param {object} fileData - Objeto opcional con los datos de la imagen (base64, mimeType, fileName).
 * @returns {string} - JSON con el resultado de la operación.
 */
function agregarDecoracion(productData, fileData) {
  try {
    if (fileData && fileData.base64Data) {
      Logger.log(
        `Subiendo imagen a Decoración (carpeta ID: ${ID_DECORACION_IMG})`
      );

      const driveFileData = {
        base64: fileData.base64Data,
        name: fileData.fileName,
        type: fileData.mimeType,
      };

      const file = _subirArchivoADrive(ID_DECORACION_IMG, driveFileData);
      productData.imagenAgregar = file.getUrl();
      Logger.log(
        `Imagen de Decoración subida. URL: ${productData.imagenAgregar}`
      );
    }

    // La función agregarProductoGenerico no está definida en este archivo,
    // pero se asume que existe en el proyecto.
    return agregarProductoGenerico(productData, HOJA_DECORACION);
  } catch (e) {
    Logger.log(`Error en agregarDecoracion: ${e.message}\nStack: ${e.stack}`);
    return JSON.stringify({
      success: false,
      message: `Error al agregar la decoración: ${e.message}`,
    });
  }
}

/**
 * @summary Obtiene productos de comida próximos a vencer o ya vencidos.
 * @description Busca en la hoja de comida productos cuya fecha de vencimiento esté dentro de los próximos 30 días o ya haya pasado.
 * @returns {Array<object>} Un array de objetos, donde cada objeto representa un producto y contiene su nombre, días restantes y estado.
 */
function getProductosProximosAVencer() {
  try {
    const sheetName = getHojasConfig().COMIDA.nombre;
    const sheet = getSheet(sheetName);
    if (!sheet) {
      Logger.log(
        `getProductosProximosAVencer: No se encontró la hoja '${sheetName}'.`
      );
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // No hay datos además del encabezado

    const headers = data.shift().map((h) => h.toString().trim().toUpperCase());
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0); // Para comparar solo fechas

    const colIndexProducto = headers.indexOf("PRODUCTO");
    const colIndexVencimiento = headers.indexOf("FECHA DE VENCIMIENTO");
    const colIndexEstado = headers.indexOf("ESTADO");

    if (
      colIndexVencimiento === -1 ||
      colIndexProducto === -1 ||
      colIndexEstado === -1
    ) {
      Logger.log(
        `getProductosProximosAVencer: Faltan columnas requeridas (PRODUCTO, FECHA DE VENCIMIENTO, ESTADO) en la hoja '${sheetName}'.`
      );
      return [];
    }

    const productos = data
      .filter(
        (row) =>
          row[colIndexEstado] &&
          row[colIndexEstado].toLowerCase() === "activo" &&
          row[colIndexVencimiento]
      )
      .map((row) => {
        const fechaVencimiento = new Date(row[colIndexVencimiento]);
        if (isNaN(fechaVencimiento.getTime())) return null;

        const diffTime = fechaVencimiento.getTime() - hoy.getTime();
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        return {
          producto: row[colIndexProducto],
          diasRestantes: diffDays,
          estado: diffDays < 0 ? "Vencido" : "Próximo a vencer",
        };
      })
      .filter((p) => p && p.diasRestantes <= 30); // Filtra productos que vencen en 30 días o ya vencieron.

    return productos;
  } catch (e) {
    Logger.log(`Error en getProductosProximosAVencer: ${e.message}`);
    return [];
  }
}

/**
 * @summary Obtiene comentarios filtrados por el nombre de la hoja de origen.
 * @description Lee todos los comentarios de la hoja 'Comentarios' y devuelve solo aquellos que pertenecen a la hoja especificada.
 * @param {string} sheetName - El nombre de la hoja para la cual se solicitan los comentarios (ej. "Inventario comida").
 * @returns {Array<object>} Un array de objetos de comentarios.
 */
function getComentariosPorHoja(sheetName) {
  try {
    const comentariosSheet = getSheet(HOJA_COMENTARIOS);
    if (!comentariosSheet) {
      Logger.log(
        `getComentariosPorHoja: No se encontró la hoja de comentarios '${HOJA_COMENTARIOS}'.`
      );
      return [];
    }

    const data = comentariosSheet.getDataRange().getValues();
    if (data.length < 2) return []; // No hay datos además del encabezado

    const headers = data.shift().map((h) => h.toString().trim().toUpperCase());
    const colIndexSheetName = headers.indexOf("SHEETNAME");

    if (colIndexSheetName === -1) {
      Logger.log(
        `getComentariosPorHoja: No se encontró la columna 'SheetName' en la hoja de comentarios.`
      );
      return []; // Si no hay columna para filtrar, devuelve vacío para evitar errores.
    }

    const comentariosFiltrados = data
      .filter((row) => row[colIndexSheetName] === sheetName)
      .map((row) => {
        const comentarioObj = {};
        headers.forEach(
          (header, index) => (comentarioObj[header] = row[index])
        );
        return comentarioObj;
      });

    return comentariosFiltrados;
  } catch (e) {
    Logger.log(`Error en getComentariosPorHoja: ${e.message}`);
    return [];
  }
}

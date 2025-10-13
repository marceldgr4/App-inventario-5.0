// BackEnd/Acta.js

/**
 * @summary Obtiene todos los datos de la hoja de Acta.
 * @returns {string} Una cadena JSON que contiene los datos de la hoja solicitada.
 */
function getActaData() {
  return JSON.stringify({
    data: _getInventoryDataForSheet(getHojasConfig().ACTA.nombre),
  });
}

/**
 * @summary Obtiene la información de un acta por su ID.
 * @param {string|number} id El ID del acta.
 * @returns {object|null} Un objeto con la información del acta o null si no se encuentra.
 */
function getActaInfo(id) {
  return getInfo(id, getHojasConfig().ACTA.nombre);
}

/**
 * @summary Agrega un nuevo acta a la hoja de Acta.
 * @param {object} data Objeto con los datos del nuevo acta.
 * @returns {string} Resultado de la operación en formato JSON.
 */
function agregarProductoActa(data) {
  return agregar(data, getHojasConfig().ACTA.nombre);
}

/**
 * @summary Actualiza un acta existente en la hoja de Acta.
 * @param {object} data Objeto con los datos a actualizar, incluyendo el ID.
 * @returns {string} Resultado de la operación en formato JSON.
 */
function actualizarProductoActa(data) {
  // ✅ CORRECCIÓN: Se quitó el .nombre que estaba mal colocado
  return actualizar(data, getHojasConfig().ACTA.nombre);
}

/**
 * @summary Sube un acta con su PDF asociado a Google Drive y registra la información en la hoja.
 * @param {object} data Objeto con los datos del acta y el PDF en base64.
 * @returns {string} Resultado de la operación en formato JSON.
 */
function subirActaConPDF(data) {
  const activeUser = getActiveUser();
  if (!activeUser) throw new Error('Usuario no autenticado para subir acta.');
  
  // Subir PDF a Drive
  const folder = DriveApp.getFolderById(ID_DRIVE_PDF);
  const contentType = 'application/pdf';
  const bytes = Utilities.base64Decode(data.base64PDF);
  const blob = Utilities.newBlob(bytes, contentType, data.nombreArchivo);
  const file = folder.createFile(blob);
  
  // Acceder a la hoja
  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  const hoja = ss.getSheetByName(HOJA_ACTA);
  if (!hoja) throw new Error('La hoja "Acta" no existe.');
  
  // Generar nuevo ID
  const ultimaFila = hoja.getLastRow();
  let nuevoId = 1;
  if (ultimaFila >= 2) {
    const lastId = parseInt(hoja.getRange(ultimaFila, 1).getValue());
    nuevoId = isNaN(lastId) ? 1 : lastId + 1;
  }
  
  // Formatear fechas
  const zonaHoraria = Session.getScriptTimeZone();
  const fechaEntrega = Utilities.formatDate(
    new Date(data.fechaEntregaAgregar),
    zonaHoraria,
    'yyyy-MM-dd HH:mm:ss'
  );
  const fechaCarga = Utilities.formatDate(
    new Date(data.fechaCargaPDFAgregar),
    zonaHoraria,
    'yyyy-MM-dd HH:mm:ss'
  );
  
  // ✅ CORRECCIÓN: Orden ajustado según las columnas reales de la hoja
  // Columnas: Id | Id_Usuario | Usuario | Fecha de entrega | Producto | Programa | Cantidad | Link | Ciudad | Nombre Compelto | fecha de carga
  hoja.appendRow([
    nuevoId,                           // Id
    activeUser.id || '',               // Id_Usuario (si existe)
    data.nombreUsuarioAgregar,         // Usuario
    fechaEntrega,                      // Fecha de entrega
    data.productoAgregar,              // Producto
    data.programaAgregar,              // Programa
    data.cantidadAgregar,              // Cantidad
    file.getUrl(),                     // Link
    data.ciudadAgregar,                // Ciudad
    data.nombreCompletoAgregar,        // Nombre Compelto (así está en la hoja)
    fechaCarga                         // fecha de carga
  ]);
  
  // Registrar en historial
  _registrarHistorialModificacion(
    nuevoId,
    `Acta: ${data.nombreArchivo}`,
    data.programaAgregar,
    null,
    data.cantidadAgregar,
    `Subida de acta: ${data.nombreArchivo}`,
    activeUser.name || activeUser.email,
    new Date(data.fechaEntregaAgregar),
    data.cantidadAgregar,
    HOJA_ACTA
  );
  
  return JSON.stringify({
    success: true,
    fileUrl: file.getUrl(),
    message: 'Acta subida y registrada correctamente.',
  });
}
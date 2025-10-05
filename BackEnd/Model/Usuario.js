//=======================================
//--- Funciones Específicas de Usuario ---
//=======================================

/**
 * @summary Obtiene el objeto de la hoja de 'Usuario'.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} El objeto de la hoja de usuarios.
 */
function getUsuarioSheet() {
  // Asume que getSheetByName(HOJA_USUARIO) está en un archivo Comunes.js
  return getSheetByName(HOJA_USUARIO);
}

/**
 * @summary Obtiene los datos de todos los usuarios activos de la hoja 'Usuario'.
 * @returns {string} Una cadena JSON con los datos de los usuarios.
 */
function getUsuarioData() {
  // Asume que getInventoryDataForSheet(HOJA_USUARIO) está en un archivo Comunes.js
  return getInventoryDataForSheet(HOJA_USUARIO);
}

/**
 * @summary Obtiene la información completa de un usuario específico por su ID.
 * @param {string|number} id El ID del usuario a buscar.
 * @returns {object|null} Un objeto con la información del usuario o null si no se encuentra.
 */
function getUsuarioInfo(id) {
  // Asume que getProductInfoGenerico está definido globalmente
  return getProductInfoGenerico(id, HOJA_USUARIO);
}

/**
 * @summary Prepara los datos de un nuevo usuario y llama a la función principal para agregarlo.
 * @description Wrapper para agregar un usuario, mapeando los campos del cliente.
 * @param {object} data El objeto con los datos del nuevo usuario desde el cliente.
 * @returns {string} El resultado de la operación de agregado.
 */
function agregarUsuario(data) {
  const newData = {
    nombreCompletoAgregar: data.nombreCompletoAgregar,
    userNameAgregar: data.userNameAgregar,
    passwordAgregar: data.passwordAgregar,
    cdeAgregar: data.cdeAgregar,
    // Se usa 'emilAgregar' del cliente, y se mapea a 'emailAgregar' para la implementación
    emailAgregar: data.emilAgregar,
    rolAgregar: data.rolAgregar,
  };
  // Llama a la implementación principal (ver abajo)
  return _ejecutarAgregarUsuario(newData, HOJA_USUARIO);
}

/**
 * @summary Prepara los datos de un usuario a actualizar y llama a la función principal.
 * @description Wrapper para actualizar un usuario, mapeando los campos del cliente.
 * @param {object} data El objeto con los datos a actualizar del usuario, incluyendo su ID.
 * @returns {string} El resultado de la operación de actualización.
 */
function actualizarUsuario(data) {
  const updateData = {
    idEditar: data.idEditar,
    nombreCompletoEditar: data.nombreCompletoEditar,
    userNameEditar: data.userNameEditar,
    passwordEditar: data.passwordEditar,
    emailEditar: data.emilEditar, // Usamos 'emilEditar' por compatibilidad con el cliente
    rolEditar: data.rolEditar,
  };
  // Llama a la implementación principal (ver abajo)
  return _ejecutarActualizarUsuario(updateData, HOJA_USUARIO);
}

/**
 * @summary Llama a la función principal para eliminar (desactivar) un usuario.
 * @param {string|number} idToDelete El ID del usuario a desactivar.
 * @returns {string} El resultado de la operación de eliminación.
 */
function eliminarUsuario(idToDelete) {
  // Llama a la implementación principal (ver abajo)
  return _ejecutarEliminarUsuario(idToDelete, HOJA_USUARIO);
}

//=========================================================
//--- Implementaciones CRUD (Funciones Internas de Lógica) ---
//=========================================================

/**
 * @summary Función principal para agregar un nuevo usuario a la hoja.
 * @param {object} data Objeto con los datos del usuario a agregar.
 * @param {string} sheetName El nombre de la hoja de destino.
 * @returns {string} Un mensaje de éxito o un objeto JSON de error.
 * @private
 */
function _ejecutarAgregarUsuario(data, sheetName) {
  const sheet = getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Validación de campos obligatorios
  if (
    !data.nombreCompletoAgregar ||
    !data.userNameAgregar ||
    !data.passwordAgregar ||
    !data.emailAgregar ||
    !data.rolAgregar
  ) {
    return JSON.stringify({ success: false, message: 'Por favor, complete todos los campos obligatorios.' });
  }

  const newRowData = [];
  const lastRow = sheet.getLastRow();
  let newId = 1;

  // Lógica para generar un nuevo ID autoincremental
  const idHeaderIndex = headers.indexOf('Id');
  if (idHeaderIndex !== -1) {
    if (lastRow >= ((sheet.getFrozenRows() || 0) + 1)) {
      const ids = sheet.getRange((sheet.getFrozenRows() || 0) + 1, idHeaderIndex + 1, lastRow - (sheet.getFrozenRows() || 0), 1)
        .getValues().flat().map(id => parseInt(id)).filter(id => !isNaN(id));
      if (ids.length > 0) newId = Math.max(...ids) + 1;
    }
  }

  // Mapeo de datos a las columnas correspondientes
  newRowData[headers.indexOf('Id')] = newId;
  newRowData[headers.indexOf('NombreCompleto')] = data.nombreCompletoAgregar;
  newRowData[headers.indexOf('userName')] = data.userNameAgregar;
  newRowData[headers.indexOf('password')] = data.passwordAgregar;
  newRowData[headers.indexOf('CDE')] = data.cdeAgregar || '';
  newRowData[headers.indexOf('Email')] = data.emailAgregar;
  newRowData[headers.indexOf('Estado_User')] = 'Activo';

  // Normalización del Rol
  let rolParaGuardar = data.rolAgregar;
  if (rolParaGuardar && typeof rolParaGuardar === 'string') {
    if (rolParaGuardar.toLowerCase() === 'usuario') {
      rolParaGuardar = ROL_USUARIO;
    } else if (rolParaGuardar.toLowerCase() === 'admin') {
      rolParaGuardar = ROL_ADMIN;
    }
  }
  newRowData[headers.indexOf('Rol')] = rolParaGuardar;
  newRowData[headers.indexOf('Fecha de Registro')] = new Date();

  sheet.appendRow(newRowData);

  // Registrar la acción en el historial (Asume que getActiveUser() y _registrarHistorialModificacion están definidos)
  const activeUser = getActiveUser();
  const performingUser = activeUser ? (activeUser.name || activeUser.email) : 'Sistema';
  _registrarHistorialModificacion(
    newId,
    data.userNameAgregar || data.nombreCompletoAgregar,
    'Gestión de Usuarios',
    null, null,
    `Usuario agregado: ${data.userNameAgregar || data.nombreCompletoAgregar}`,
    performingUser,
    new Date(), null
  );

  return `Usuario agregado exitosamente a ${sheetName}.`;
}

/**
 * @summary Función principal para actualizar los datos de un usuario existente.
 * @param {object} data Objeto con los datos a actualizar, debe contener 'idEditar'.
 * @param {string} sheetName El nombre de la hoja donde está el usuario.
 * @returns {string} Un mensaje de éxito/error o un objeto JSON.
 * @private
 */
function _ejecutarActualizarUsuario(data, sheetName) {
  const sheet = getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Asume que findRowById(id, sheetName) está en Comunes.js
  const rowIndex = findRowById(data.idEditar, sheetName);

  if (rowIndex > 0) {
    let cambiosRealizados = false;
    let nombreOriginal = sheet.getRange(rowIndex, headers.indexOf('NombreCompleto') + 1).getValue();

    // Función auxiliar para actualizar un campo
    const updateField = (fieldKey, headerName) => {
      const headerIndex = headers.indexOf(headerName);
      if (data[fieldKey] !== undefined && headerIndex > -1) {
        if (headerName === 'password' && data[fieldKey] === "") return;

        sheet.getRange(rowIndex, headerIndex + 1).setValue(data[fieldKey]);
        cambiosRealizados = true;
      }
    };

    updateField('nombreCompletoEditar', 'NombreCompleto');
    updateField('userNameEditar', 'userName');
    updateField('passwordEditar', 'password');
    updateField('emailEditar', 'Email');
    updateField('cdeEditar', 'CDE');

    // Actualización y Normalización del Rol
    if (data.rolEditar !== undefined && headers.indexOf('Rol') > -1) {
      let rolParaActualizar = data.rolEditar;
      if (rolParaActualizar && typeof rolParaActualizar === 'string') {
        if (rolParaActualizar.toLowerCase() === 'usuario') {
          rolParaActualizar = ROL_USUARIO;
        } else if (rolParaActualizar.toLowerCase() === 'admin') {
          rolParaActualizar = ROL_ADMIN;
        }
      }
      sheet.getRange(rowIndex, headers.indexOf('Rol') + 1).setValue(rolParaActualizar);
      cambiosRealizados = true;
    }

    if (cambiosRealizados) {
      const activeUser = getActiveUser();
      const performingUser = activeUser ? (activeUser.name || activeUser.email) : 'Sistema';
      _registrarHistorialModificacion(
        data.idEditar,
        data.userNameEditar || nombreOriginal,
        'Gestión de Usuarios',
        null, null,
        `Usuario actualizado: ${data.userNameEditar || nombreOriginal}`,
        performingUser,
        new Date(), null
      );
      return `Usuario modificado exitosamente a ${sheetName}.`;
    } else {
      return JSON.stringify({ success: false, message: `No se realizaron cambios para el usuario ID ${data.idEditar}.` });
    }

  } else {
    return JSON.stringify({ success: false, message: `No se encontró el Usuario con el ID ${data.idEditar} en ${sheetName}.` });
  }
}

/**
 * @summary Realiza una "eliminación lógica" de un usuario cambiando su estado a "Desactivado".
 * @param {string|number} idToDelete El ID del usuario a desactivar.
 * @param {string} sheetName El nombre de la hoja de usuarios.
 * @returns {string} Un mensaje indicando el resultado de la operación.
 * @private
 */
function _ejecutarEliminarUsuario(idToDelete, sheetName) {
  const sheet = getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = findRowById(idToDelete, sheetName);
  const activeUserPerformingAction = Session.getActiveUser().getEmail();

  if (rowIndex > 0) {
    // Cambiar estado a 'Desactivado'
    sheet
      .getRange(rowIndex, headers.indexOf('Estado_User') + 1)
      .setValue('Desactivado');

    const nombreCompletoDelUsuarioEliminado = sheet
      .getRange(rowIndex, headers.indexOf('NombreCompleto') + 1)
      .getValue();

    // Registrar en historial
    _registrarHistorialModificacion(
      idToDelete,
      nombreCompletoDelUsuarioEliminado,
      'Gestión de Usuarios',
      null,
      null,
      `Usuario Desactivado: ${nombreCompletoDelUsuarioEliminado} (ID: ${idToDelete})`,
      activeUserPerformingAction,
      new Date(),
      null
    );
    return `Usuario ${nombreCompletoDelUsuarioEliminado} (ID: ${idToDelete}) desactivado exitosamente.`;
  } else {
    return `No se encontró el Usuario con el ID ${idToDelete} en ${sheetName}.`;
  }
}
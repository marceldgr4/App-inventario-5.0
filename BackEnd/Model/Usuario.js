/**
 * Usuarios.gs
 * Backend Apps Script para módulo Usuarios
 */

/** Obtener todos los usuarios (objeto JS { success, data }) */
function getUsuarioData() {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    const raw = _getInventoryDataForSheet(sheetName); // array de objetos
    // Normalizar fechas a ISO strings para evitar problemas de serialización
    const data = (raw || []).map(item => {
      const out = {};
      for (const k in item) {
        const v = item[k];
        try {
          if (Object.prototype.toString.call(v) === '[object Date]') {
            out[k] = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
          } else {
            out[k] = v;
          }
        } catch (e) {
          out[k] = v;
        }
      }
      return out;
    });
    return { success: true, data };
  } catch (e) {
    Logger.log('ERROR getUsuarioData: ' + e.stack);
    return { success: false, data: [], error: e.message };
  }
}

/** Alias (compatibilidad) */
/** Alias (compatibilidad) */
function fetchUsuarioData() {
  return getUsuarioData();
}

/** Compatibilidad: devolver JSON string (algunos módulos esperan stringified) */
function getUsuariosJSON() {
  try {
    const resp = getUsuarioData();
    return JSON.stringify(resp);
  } catch (e) {
    Logger.log('ERROR getUsuariosJSON: ' + e.stack);
    return JSON.stringify({ success: false, data: [], error: e.message });
  }
}

/** Obtener info de un usuario por ID (retorna objeto JS o null) */
function getUsuarioInfo(id) {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    return getInfo(id, sheetName); // asume getInfo() existe
  } catch (e) {
    Logger.log('ERROR getUsuarioInfo: ' + e.stack);
    return null;
  }
}

/** Agregar usuario */
function agregarUsuario(data) {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    Logger.log('agregarUsuario: llamada recibida. payload=' + JSON.stringify(data));
    // Compatibilidad: el frontend envía { nombreCompleto, userName, password, cde, email, rol }
    // mientras que la implementación interna espera campos con sufijo "Agregar".
    const payload = Object.assign({}, data);
    if (payload.nombreCompleto && !payload.nombreCompletoAgregar) payload.nombreCompletoAgregar = payload.nombreCompleto;
    if (payload.userName && !payload.userNameAgregar) payload.userNameAgregar = payload.userName;
    if (payload.password && !payload.passwordAgregar) payload.passwordAgregar = payload.password;
    if (payload.cde && !payload.cdeAgregar) payload.cdeAgregar = payload.cde;
    if (payload.email && !payload.emailAgregar) payload.emailAgregar = payload.email;
    if (payload.rol && !payload.rolAgregar) payload.rolAgregar = payload.rol;

    const result = _ejecutarAgregarUsuario(payload, sheetName);
    Logger.log('agregarUsuario: resultado=' + JSON.stringify(result));
    // Retornar objeto (no string) para facilitar consumo desde google.script.run
    return result;
  } catch (e) {
    Logger.log('ERROR agregarUsuario: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/** Actualizar usuario */
function actualizarUsuario(data) {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    const result = _ejecutarActualizarUsuario(data, sheetName);
    return result;
  } catch (e) {
    Logger.log('ERROR actualizarUsuario: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/** Eliminar (desactivar) usuario */
function eliminarUsuario(id) {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    Logger.log('eliminarUsuario: llamada recibida. id=' + id);
    const result = _ejecutarEliminarUsuario(id, sheetName);
    Logger.log('eliminarUsuario: resultado=' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('ERROR eliminarUsuario: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/** Compatibilidad: devolver directamente la lista de usuarios (usado por la UI) */
function getUsuarios() {
  try {
    const sheetName = getHojasConfig().USUARIO.nombre;
    Logger.log('getUsuarios: llamada recibida. Hoja=' + sheetName);
    const raw = _getInventoryDataForSheet(sheetName) || [];
    // Convertir fechas a strings ISO
    const data = raw.map(item => {
      const out = {};
      for (const k in item) {
        const v = item[k];
        if (Object.prototype.toString.call(v) === '[object Date]') {
          out[k] = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        } else {
          out[k] = v;
        }
      }
      return out;
    });
    Logger.log('getUsuarios: registros leidos=' + data.length);
    return data;
  } catch (e) {
    Logger.log('ERROR getUsuarios: ' + e.stack);
    return [];
  }
}

/** Compatibilidad: wrapper para llamadas desde la UI que envían {id, nombreCompleto, userName, ...} */
function editarUsuario(data) {
  try {
    Logger.log('editarUsuario: llamada recibida. payload=' + JSON.stringify(data));
    // Mapear a la forma interna esperada por _ejecutarActualizarUsuario
    const payload = {};
    if (data.id !== undefined) payload.idEditar = data.id;
    if (data.nombreCompleto !== undefined) payload.nombreCompletoEditar = data.nombreCompleto;
    if (data.userName !== undefined) payload.userNameEditar = data.userName;
    if (data.password !== undefined) payload.passwordEditar = data.password;
    if (data.cde !== undefined) payload.cdeEditar = data.cde;
    if (data.email !== undefined) payload.emailEditar = data.email;
    if (data.rol !== undefined) payload.rolEditar = data.rol;

    const sheetName = getHojasConfig().USUARIO.nombre;
    const result = _ejecutarActualizarUsuario(payload, sheetName);
    Logger.log('editarUsuario: resultado=' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('ERROR editarUsuario: ' + e.stack);
    return { success: false, error: e.message };
  }
}

/* ===========================
   Implementaciones internas
   =========================== */

/** Agregar Usuario */
function _ejecutarAgregarUsuario(data, sheetName) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Hoja no encontrada: ' + sheetName };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Validación obligatoria
    if (!data?.nombreCompletoAgregar || !data?.userNameAgregar || !data?.passwordAgregar || !data?.emailAgregar || !data?.rolAgregar) {
      return { success: false, message: 'Por favor complete todos los campos obligatorios.' };
    }

    // Generar nuevo ID
    const idHeaderIndex = headers.indexOf('Id');
    let newId = 1;
    if (idHeaderIndex !== -1 && sheet.getLastRow() >= 2) {
      const idValues = sheet.getRange(2, idHeaderIndex + 1, sheet.getLastRow() - 1, 1).getValues().flat()
        .map(v => parseInt(v, 10)).filter(n => !isNaN(n));
      if (idValues.length) newId = Math.max(...idValues) + 1;
    }

    // Construcción de fila
    const newRow = new Array(headers.length).fill('');
    const setIfHeader = (headerName, value) => {
      const idx = headers.indexOf(headerName);
      if (idx > -1) newRow[idx] = value;
    };

    setIfHeader('Id', newId);
    setIfHeader('NombreCompleto', data.nombreCompletoAgregar);
    setIfHeader('userName', data.userNameAgregar);
    setIfHeader('password', data.passwordAgregar);
    setIfHeader('CDE', data.cdeAgregar || '');
    setIfHeader('Email', data.emailAgregar);
    setIfHeader('Estado_User', 'Activo');

    // Rol
    let rol = data.rolAgregar;
    if (rol?.toLowerCase() === 'usuario') rol = typeof ROL_USUARIO !== 'undefined' ? ROL_USUARIO : 'Usuario';
    if (rol?.toLowerCase().startsWith('admin')) rol = typeof ROL_ADMIN !== 'undefined' ? ROL_ADMIN : 'Administrador';
    setIfHeader('Rol', rol);

    setIfHeader('Fecha de Registro', new Date());

    sheet.appendRow(newRow);

    // Registrar historial
    const active = getActiveUser();
    const performingUser = active?.getEmail ? active.getEmail() : (active?.email || 'Sistema');
    _registrarHistorialModificacion(
      newId,
      data.userNameAgregar,
      'Gestión de Usuarios',
      null, null,
      `Usuario agregado: ${data.userNameAgregar}`,
      performingUser,
      new Date(),
      null,
      sheetName
    );

    return { success: true, message: `Usuario agregado exitosamente.`, id: newId };
  } catch (e) {
    Logger.log('_ejecutarAgregarUsuario ERROR: ' + e.stack);
    return { success: false, message: 'Error al agregar usuario: ' + e.message };
  }
}

/** Actualizar Usuario */
function _ejecutarActualizarUsuario(data, sheetName) {
  try {
    if (!data?.idEditar) return { success: false, message: 'Falta idEditar.' };

    const sheet = getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowIndex = findRowById(data.idEditar, sheetName);
    if (rowIndex <= 0) return { success: false, message: `No se encontró el Usuario con el ID ${data.idEditar}.` };

    let cambios = false;

    const updateIf = (fieldKey, headerName, allowEmpty) => {
      const idx = headers.indexOf(headerName);
      if (idx === -1) return;
      if (data[fieldKey] === undefined) return;
      if (!allowEmpty && !data[fieldKey]) return;
      sheet.getRange(rowIndex, idx + 1).setValue(data[fieldKey]);
      cambios = true;
    };

    updateIf('nombreCompletoEditar', 'NombreCompleto', false);
    updateIf('userNameEditar', 'userName', false);
    updateIf('passwordEditar', 'password', true);
    updateIf('emailEditar', 'Email', false);
    updateIf('cdeEditar', 'CDE', true);

    if (data.rolEditar !== undefined && headers.indexOf('Rol') > -1) {
      let rol = data.rolEditar;
      if (rol?.toLowerCase() === 'usuario') rol = typeof ROL_USUARIO !== 'undefined' ? ROL_USUARIO : 'Usuario';
      if (rol?.toLowerCase().startsWith('admin')) rol = typeof ROL_ADMIN !== 'undefined' ? ROL_ADMIN : 'Administrador';
      sheet.getRange(rowIndex, headers.indexOf('Rol') + 1).setValue(rol);
      cambios = true;
    }

    if (cambios) {
      const active = getActiveUser();
      const performingUser = active?.getEmail ? active.getEmail() : (active?.email || 'Sistema');
      _registrarHistorialModificacion(
        data.idEditar,
        data.userNameEditar,
        'Gestión de Usuarios',
        null, null,
        `Usuario actualizado: ${data.userNameEditar}`,
        performingUser,
        new Date(),
        null
      );
      return { success: true, message: 'Usuario actualizado exitosamente.' };
    }

    return { success: false, message: `No se realizaron cambios para el usuario ID ${data.idEditar}.` };
  } catch (e) {
    Logger.log('_ejecutarActualizarUsuario ERROR: ' + e.stack);
    return { success: false, message: 'Error al actualizar usuario: ' + e.message };
  }
}

/** Eliminar Usuario (lógico) */
function _ejecutarEliminarUsuario(idToDelete, sheetName) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Hoja no encontrada.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowIndex = findRowById(idToDelete, sheetName);
    if (rowIndex <= 0) return { success: false, message: `No se encontró el Usuario con el ID ${idToDelete}.` };

    const estadoIdx = headers.indexOf('Estado_User');
    if (estadoIdx > -1) sheet.getRange(rowIndex, estadoIdx + 1).setValue('Desactivado');

    const nombreCompleto = headers.indexOf('NombreCompleto') > -1
      ? sheet.getRange(rowIndex, headers.indexOf('NombreCompleto') + 1).getValue()
      : '';

    const active = getActiveUser();
    const performingUser = active?.getEmail ? active.getEmail() : (active?.email || 'Sistema');

    _registrarHistorialModificacion(
      idToDelete,
      nombreCompleto,
      'Gestión de Usuarios',
      null, null,
      `Usuario Desactivado: ${nombreCompleto} (ID: ${idToDelete})`,
      performingUser,
      new Date(),
      null
    );

    return { success: true, message: `Usuario ${nombreCompleto} desactivado exitosamente.` };
  } catch (e) {
    Logger.log('_ejecutarEliminarUsuario ERROR: ' + e.stack);
    return { success: false, message: 'Error al desactivar usuario: ' + e.message };
  }
}

// Funciones de Perfil de Usuario
// =========================================================================

/**
 * @summary Obtiene los datos del perfil del usuario actualmente logueado.
 * @description Usa la sesión activa para encontrar al usuario en la hoja 'Usuario' y devuelve sus datos públicos.
 * No devuelve la contraseña por seguridad.
 * @returns {string} Un objeto JSON con el estado de la operación y los datos del perfil.
 */
function getMiPerfilData() {
  // Asume que getActiveUser() está definido globalmente
  const activeUserSession = getActiveUser();
  if (!activeUserSession || !activeUserSession.email) {
    return JSON.stringify({ success: false, message: "No se pudo identificar al usuario activo." });
  }

  try {
    // Asume que getSheetByName(HOJA_USUARIO) está en Comunes.js
    const sheet = getSheetByName(HOJA_USUARIO);
    if (!sheet) {
      return JSON.stringify({ success: false, message: `Hoja "${HOJA_USUARIO}" no encontrada.` });
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailColIdx = headers.indexOf('Email');
    const nombreColIdx = headers.indexOf('NombreCompleto');
    const userNameColIdx = headers.indexOf('userName');

    if (emailColIdx === -1 || nombreColIdx === -1 || userNameColIdx === -1) {
      return JSON.stringify({ success: false, message: "Columnas requeridas no encontradas en la hoja de usuarios." });
    }

    for (let i = 1; i < data.length; i++) {
      // Comparación por Email del usuario activo
      if (data[i][emailColIdx] === activeUserSession.email) {
        return JSON.stringify({
          success: true,
          data: {
            // Asegúrate de que las claves 'nombreCompleto' y 'userName' coincidan con el frontend
            nombreCompleto: data[i][nombreColIdx],
            userName: data[i][userNameColIdx],
          }
        });
      }
    }
    return JSON.stringify({ success: false, message: "Usuario no encontrado en la hoja." });
  } catch (e) {
    Logger.log('Error en getMiPerfilData: ' + e.toString());
    return JSON.stringify({ success: false, message: 'Error al obtener datos del perfil: ' + e.message });
  }
}

/**
 * @summary Actualiza el perfil del usuario actualmente logueado.
 * @description Permite cambiar el nombre completo, nombre de usuario y, opcionalmente, la contraseña.
 * @param {object} data Objeto que contiene `nombreCompleto`, `userName` y, opcionalmente, `newPassword`.
 * @returns {string} Un objeto JSON con el resultado de la operación.
 */
function actualizarMiPerfil(data) {
  // Asume que getActiveUser() y las constantes están definidas globalmente
  const activeUserSession = getActiveUser();
  if (!activeUserSession || !activeUserSession.email) {
    return JSON.stringify({ success: false, message: "No se pudo identificar al usuario activo para la actualización." });
  }

  // Validación de datos recibidos
  if (!data || !data.nombreCompleto || !data.userName) {
    return JSON.stringify({ success: false, message: "Nombre completo y nombre de usuario son requeridos." });
  }

  let newPassword = data.newPassword ? data.newPassword.trim() : "";
  let passwordChanged = false;

  if (newPassword !== "") {
    if (!/^[a-zA-Z]/.test(newPassword)) {
      return JSON.stringify({ success: false, message: "La nueva contraseña debe comenzar con una letra." });
    }
    passwordChanged = true;
  }

  try {
    const sheet = getSheetByName(HOJA_USUARIO);
    if (!sheet) {
      return JSON.stringify({ success: false, message: `Hoja "${HOJA_USUARIO}" no encontrada.` });
    }
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    const idColIdx = headers.indexOf('Id');
    const emailColIdx = headers.indexOf('Email');
    const nombreColIdx = headers.indexOf('NombreCompleto');
    const userNameColIdx = headers.indexOf('userName');
    const passwordColIdx = headers.indexOf('password');

    if (idColIdx === -1 || emailColIdx === -1 || nombreColIdx === -1 || userNameColIdx === -1 || passwordColIdx === -1) {
      return JSON.stringify({ success: false, message: "Columnas críticas no encontradas para actualizar perfil." });
    }

    let rowIndex = -1;
    let userIdForHistory = null;

    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][emailColIdx] === activeUserSession.email) {
        rowIndex = i + 1; // getRange es 1-based
        userIdForHistory = sheetData[i][idColIdx];
        break;
      }
    }

    if (rowIndex === -1) {
      return JSON.stringify({ success: false, message: "Usuario no encontrado para actualizar." });
    }

    // Actualizar Hoja de Cálculo
    sheet.getRange(rowIndex, nombreColIdx + 1).setValue(data.nombreCompleto);
    sheet.getRange(rowIndex, userNameColIdx + 1).setValue(data.userName);
    if (passwordChanged) {
      sheet.getRange(rowIndex, passwordColIdx + 1).setValue(newPassword);
    }

    // Actualizar PropertiesService si el nombre cambió
    if (activeUserSession.name !== data.nombreCompleto) {
      const userPropsString = PropertiesService.getUserProperties().getProperty(CLAVE_PROPIEDAD_USUARIO);
      let userProps = userPropsString ? JSON.parse(userPropsString) : {};

      userProps.name = data.nombreCompleto;
      PropertiesService.getUserProperties().setProperty(CLAVE_PROPIEDAD_USUARIO, JSON.stringify(userProps));
    }

    // Registrar en Historial
    let actionDetail = "Perfil actualizado.";
    if (passwordChanged) {
      actionDetail = "Perfil actualizado (contraseña cambiada).";
    }
    _registrarHistorialModificacion(
      userIdForHistory,
      data.userName,
      'Perfil de Usuario',
      null, null,
      actionDetail,
      activeUserSession.name || activeUserSession.email,
      new Date(),
      null
    );

    return JSON.stringify({ success: true, message: "Perfil actualizado exitosamente." });

  } catch (e) {
    Logger.log('Error en actualizarMiPerfil: ' + e.toString() + " Stack: " + e.stack);
    return JSON.stringify({ success: false, message: 'Error al actualizar el perfil: ' + e.message });
  }
}
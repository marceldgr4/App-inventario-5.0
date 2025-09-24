// =================================================================
// --- FUNCIONES DE VERIFICACIÓN DE PERMISOS ---
// =================================================================

/**
 * @summary Obtiene las páginas permitidas para un rol, usando la caché para acelerar.
 * @param {string} userRole El rol del usuario ('Admin', 'Usuario').
 * @returns {Array<string>} Un array de las páginas permitidas.
 */
function getAllowedPagesFromCache(userRole) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `PAGINAS_ROL_${userRole}`;
  
  const cachedPages = cache.get(cacheKey);
  if (cachedPages) {
    // Si está en caché, la devolvemos directamente. ¡Esto es súper rápido!
    Logger.log(`Obteniendo permisos para rol '${userRole}' desde la CACHÉ.`);
    return JSON.parse(cachedPages);
  } else {
    // Si no está en caché, la obtenemos de la constante y la guardamos para la próxima vez.
    Logger.log(`Obteniendo permisos para rol '${userRole}' desde la CONSTANTE y guardando en caché.`);
    const pages = PAGINAS_POR_ROL[userRole] || [];
    // Guardar en caché por 6 horas (21600 segundos)
    cache.put(cacheKey, JSON.stringify(pages), 21600);
    return pages;
  }
}

function isPageAllowedForUser(pageName, userRole) {
  // Primero, se asegura de que se haya proporcionado un rol.
  if (!userRole) {
    Logger.log('isPageAllowedForUser: No se proporcionó rol de usuario.');
    return false;
  }
  // Obtiene la lista de páginas permitidas para el rol del usuario desde el objeto PAGINAS_POR_ROL.
  const allowedPagesForRole = PAGINAS_POR_ROL[userRole];
  if (!allowedPagesForRole) {
    Logger.log(
      `isPageAllowedForUser: Rol desconocido o sin páginas definidas: ${userRole}`
    );
    return false;
  }
  // Adicionalmente, verifica si la página existe en la lista maestra PAGES_PERMITIDAS.
  if (!PAGES_PERMITIDAS.includes(pageName)) {
    Logger.log(
      `isPageAllowedForUser: La página '${pageName}' no está en la lista global PAGES_PERMITIDAS.`
    );
    return false;
  }
  // Finalmente, comprueba si la página está en la lista de páginas permitidas para ese rol.
  const isAllowed = allowedPagesForRole.includes(pageName);
  Logger.log(
    `isPageAllowedForUser: Verificando acceso a página '${pageName}', Rol '${userRole}', Permitido: ${isAllowed}`
  );
  return isAllowed;
}

function getActiveUser() {
  const userProperties = PropertiesService.getUserProperties();
  const userJson = userProperties.getProperty(CLAVE_PROPIEDAD_USUARIO);

  if (!userJson) {
    return null;
  }

  try {
    return JSON.parse(userJson);
  } catch (e) {
    Logger.log(`Error al parsear JSON de usuario activo: ${e}. Datos: ${userJson}`);
    // Si el JSON está corrupto, lo eliminamos para evitar errores futuros.
    userProperties.deleteProperty(CLAVE_PROPIEDAD_USUARIO);
    return null;
  }
}

/**
 * Cierra la sesión del usuario actual borrando sus datos de PropertiesService.
 */
function clearAllUserProperties() {
  PropertiesService.getUserProperties().deleteAllProperties();
  Logger.log('Todas las propiedades de usuario han sido borradas.');
}

function logAuthAction(userName, action) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(HOJA_REGISTRO_INICIO_SESION);
    if (!sheet) {
      Logger.log(
        `Advertencia: La hoja de registro '${HOJA_REGISTRO_INICIO_SESION}' no existe. No se pudo registrar la acción.`
      );
      return;
    }
    const newId = _obtenerSiguienteId(sheet);
    sheet.appendRow([newId, new Date(), userName, action]);
    Logger.log(
      `Acción de autenticación registrada: ID ${newId}, Usuario: ${userName}, Acción: ${action}`
    );
  } catch (e) {
    Logger.log(
      `Error al registrar acción de autenticación para ${userName}: ${e.message}`
    );
  }
}


function logout() {
  const activeUser = getActiveUser();
  if (activeUser && activeUser.name) {
    // Registra la acción de cierre de sesión en la hoja de cálculo.
    logAction(activeUser.name, 'Cierre de sesión');
  } else {
    logAction(
      'Desconocido',
      'Cierre de sesión (sin usuario activo en properties)'
    );
  }
  // Borra todas las propiedades guardadas para este usuario.
  clearAllUserProperties();
  // Devuelve la URL de la página de Login para que el cliente pueda redirigir.
  return getScriptUrl() + '?page=Login';
}

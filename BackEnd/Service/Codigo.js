// =================================================================
// --- RUTAS PRINCIPALES ---
// =================================================================

function doGet(e) {
  // En la arquitectura SPA, el parámetro 'page' es manejado por el cliente.
  // El servidor solo necesita verificar la autenticación y servir la página correcta.

  const activeUser = getActiveUser();
  if (!activeUser) {
    Logger.log('Usuario no autenticado. Sirviendo página de Login.');
    return processSignin();
  }
  
  Logger.log(`Usuario activo: ${activeUser.name} (Rol: ${activeUser.rol}). Sirviendo la aplicación principal.`);
  // Si el usuario está autenticado, siempre servimos la página maestra 'Index'.
  // El cliente se encargará de cargar el contenido correcto (ej. Home, Articulos, etc.).
  return loadPage('Index', activeUser);
}

/**
 * Función de utilidad que devuelve la URL de despliegue del script.
 * @returns {string} - La URL de la aplicación web.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

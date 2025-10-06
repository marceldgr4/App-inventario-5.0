// =================================================================
// --- FUNCIONES DE VISTAS ---
// =================================================================

function _crearRespuestaHtml(template, pageTitle) {
  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setTitle(pageTitle)
    .setFaviconUrl(URL_FAVICON)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _mostrarPaginaError(titulo, mensaje, esCritico = false) {
    let link = `<a href="${getScriptUrl()}?page=Index">Ir a la página de inicio</a>`;
    if (esCritico) {
        link = `<a href="${getScriptUrl()}?page=Login">Intentar iniciar sesión de nuevo</a>`;
    }
    const html = `<h1>${titulo}</h1><p>${mensaje}</p><p>${link}</p>`;
    return HtmlService.createHtmlOutput(html)
        .setTitle(titulo)
        .setFaviconUrl(URL_FAVICON);
}

function _redirigirA(pageName) {
    return HtmlService.createHtmlOutput(
        `<script>window.top.location.href = "${getScriptUrl()}?page=${pageName}";</script>`
    );
}

function loadPage(nombrePagina, usuarioActivo, accessMessage = null) {
  try {
    if (!PAGES_PERMITIDAS.includes(nombrePagina)) {
      Logger.log(`Intento de cargar página no válida: "${nombrePagina}"`);
      return _mostrarPaginaError('Página no encontrada', `La página '${nombrePagina}' no existe.`);
    }

    const filePath = `View/${nombrePagina}`;

    const plantilla = HtmlService.createTemplateFromFile(filePath);

    plantilla.usuarioActivo = usuarioActivo;
    plantilla.paginasDisponiblesParaUsuario = [];
    if (usuarioActivo && usuarioActivo.rol) {
      plantilla.paginasDisponiblesParaUsuario = PAGINAS_POR_ROL[usuarioActivo.rol] || [];
    }
    plantilla.accessMessage = accessMessage;

    return _crearRespuestaHtml(plantilla, `Inventario - ${nombrePagina}`);
  } catch (error) {
    Logger.log(`Error severo al cargar la página "${nombrePagina}": ${error.stack}`);
    let errorMessage = `Ocurrió un error al cargar la página: ${nombrePagina}.`;
    if (usuarioActivo && usuarioActivo.rol === ROL_ADMIN) {
      errorMessage += ` Detalles: ${error.message}`;
    }
    return _mostrarPaginaError('Error del Sistema', errorMessage, true);
  }
}

/**
 * @summary Carga el contenido HTML de una subpágina para la arquitectura SPA.
 * @param {string} pageName El nombre de la página de contenido a cargar (ej. 'Home').
 * @returns {string} El contenido HTML del archivo solicitado.
 * @throws {Error} Si el usuario no está autenticado, no tiene permisos o el archivo no existe.
 */
function loadContent(pageName) {
  const activeUser = getActiveUser();
  if (!activeUser) {
    throw new Error('Sesión expirada. Por favor, recargue la página para iniciar sesión.');
  }

  if (!isPageAllowedForUser(pageName, activeUser.rol)) {
    throw new Error(`Acceso denegado a la página '${pageName}'.`);
  }

  // Mapea el nombre de la página al archivo de contenido correspondiente.
  // Por convención, el contenido de 'Home' está en 'Home_content.html'.
  let contentFile;
  if (
<<<<<<< Updated upstream
   
    pageName === 'Articulos' ||
    pageName === 'Categorias' ||
    pageName === 'Usuarios' ||
    pageName === 'Proveedores' ||
    pageName === 'Clientes' ||
    pageName === 'Ventas' ||
    pageName === 'Compras' ||
    pageName === 'Reportes'||
    pageName === 'Historial' ||
    pageName === 'Comida'||
    pageName === 'Decoracion'||
    pageName === 'Papeleria'

  ){
=======
    pageName === "Articulos" ||
    pageName === "Categorias" ||
    pageName === "Usuarios" ||
    pageName === "Dashboard" ||
    pageName === "Clientes" ||
    pageName === "Ventas" ||
    pageName === "Compras" ||
    pageName === "Reportes" ||
    pageName === "Historial" ||
    pageName === "Comida" ||
    pageName === "Decoracion" ||
    pageName === "Papeleria" ||
    pageName === "Comentario"
  ) {
>>>>>>> Stashed changes
    contentFile = `View/Page/${pageName}`;
  } else {
    contentFile = `View/${pageName}_content`;
  }

  try {
    const template = HtmlService.createTemplateFromFile(contentFile);
    template.usuarioActivo = activeUser;
    return template.evaluate().getContent();
  } catch (e) {
    Logger.log(`Error al cargar contenido para '${pageName}' desde '${contentFile}': ${e.toString()}`);
    throw new Error(`No se pudo cargar el contenido para la página '${pageName}'. El archivo podría no existir.`);
  }
}


function include(filename) {
  const cleanFilename = filename.replace('.html', '');
  return HtmlService.createHtmlOutputFromFile(cleanFilename).getContent();
}

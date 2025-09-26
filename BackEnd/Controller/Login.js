/**
 * @summary Registra una acción específica de un usuario en una hoja de cálculo de registro.
 * @description Si la hoja de registro no existe, la crea automáticamente con las cabeceras necesarias.
 * Genera un ID autoincremental para cada nueva entrada de registro.
 * @param {string} nombreUsuario El nombre del usuario que realiza la acción.
 * @param {string} accion La descripción de la acción realizada (ej. "Inicio de sesión fallido").
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [hojaCalculo] Opcional. El objeto Spreadsheet para evitar reabrirlo.
 */
function logAction(nombreUsuario, accion, hojaCalculo) {
  try {
    const ss = hojaCalculo || SpreadsheetApp.openById(ID_INVENTARIO);
    let hojaRegistro = ss.getSheetByName(HOJA_REGISTRO_USUARIO);

    if (!hojaRegistro) {
      hojaRegistro = ss.insertSheet(HOJA_REGISTRO_USUARIO);
      hojaRegistro.appendRow(['ID', 'Fecha', 'Usuario', 'Acción']); // encabezados
      hojaRegistro.setFrozenRows(1);
    }

    const lastRow = hojaRegistro.getLastRow();
    let nextId = 1;

    if (lastRow >= 1 && hojaRegistro.getRange(1, 1).getValue() === 'ID') {
      // Si hay encabezado 'ID'
      if (lastRow > 1) {
        // Si hay datos además del encabezado
        const lastIdCell = hojaRegistro.getRange(lastRow, 1).getValue();
        if (typeof lastIdCell === 'number' && !isNaN(lastIdCell)) {
          nextId = lastIdCell + 1;
        } else {
          // Si la última celda no es un número, buscar el último ID numérico
          const ids = hojaRegistro
            .getRange(2, 1, lastRow - 1, 1)
            .getValues()
            .flat()
            .filter(id => typeof id === 'number');
          if (ids.length > 0) nextId = Math.max(...ids) + 1;
        }
      }
    } else if (lastRow >= 1) {
      // No hay encabezado 'ID' o está mal
      const lastIdCell = hojaRegistro.getRange(lastRow, 1).getValue();
      if (typeof lastIdCell === 'number' && !isNaN(lastIdCell)) {
        nextId = lastIdCell + 1;
      }
    }

    hojaRegistro.appendRow([nextId, new Date(), nombreUsuario, accion]);
  } catch (error) {
    console.error(
      "Error al registrar la acción '" +
        accion +
        "' para el usuario '" +
        nombreUsuario +
        "': " +
        error.toString()
    );
    Logger.log(
      'Error en logAction: ' + error.message + ' Stack: ' + error.stack
    );
  }
}
//======================================
// --- MANEJO DE SESIÓN DE USUARIO  ---
//======================================
/**
 * @summary Obtiene una propiedad almacenada para el usuario actual.
 * @description Utiliza PropertiesService para recuperar un valor asociado a una clave, específico para el usuario que ejecuta el script.
 * @param {string} clave La clave de la propiedad a obtener.
 * @returns {object|null} El valor de la propiedad parseado como JSON, o null si no se encuentra o hay un error.
 */
function getUserProperty(clave) {
  try {
    const propiedadesUsuario = PropertiesService.getUserProperties();
    const valor = propiedadesUsuario.getProperty(clave);
    return valor ? JSON.parse(valor) : null;
  } catch (e) {
    Logger.log(
      "Error al obtener propiedad de usuario '" + clave + "': " + e.toString()
    );
    return null;
  }
}

/**
 * @summary Establece una propiedad para el usuario actual.
 * @description Utiliza PropertiesService para guardar un par clave-valor para el usuario que ejecuta el script. El valor se guarda como una cadena JSON.
 * @param {string} clave La clave de la propiedad a establecer.
 * @param {object} valor El valor (objeto) a almacenar.
 */
function setUserProperty(clave, valor) {
  try {
    const propiedadesUsuario = PropertiesService.getUserProperties();
    propiedadesUsuario.setProperty(clave, JSON.stringify(valor));
  } catch (e) {
    Logger.log(
      "Error al establecer propiedad de usuario '" +
        clave +
        "': " +
        e.toString()
    );
  }
}

/**
 * @summary Elimina una propiedad del usuario actual.
 * @param {string} clave La clave de la propiedad a eliminar.
 */
function deleteUserProperty(clave) {
  try {
    const propiedadesUsuario = PropertiesService.getUserProperties();
    propiedadesUsuario.deleteProperty(clave);
  } catch (e) {
    Logger.log(
      "Error al eliminar propiedad de usuario '" + clave + "': " + e.toString()
    );
  }
}

/**
 * @summary Elimina todas las propiedades almacenadas para el usuario actual.
 * @description Útil para una limpieza completa o un cierre de sesión forzado.
 */
function clearAllUserProperties() {
  try {
    const propiedadesUsuario = PropertiesService.getUserProperties();
    propiedadesUsuario.deleteAllProperties();
    Logger.log('Todas las propiedades de usuario han sido limpiadas.');
  } catch (e) {
    Logger.log(
      'Error al limpiar todas las propiedades de usuario: ' + e.toString()
    );
  }
}

/**
 * @summary Guarda los datos del usuario que ha iniciado sesión en las propiedades del usuario.
 * @description Esto establece la sesión activa para el usuario.
 * @param {object} usuario Un objeto que contiene los detalles del usuario (Email, Id, NombreCompleto, Rol).
 */
function setActiveUser(usuario) {
  if (
    usuario &&
    usuario.Email &&
    usuario.Id &&
    usuario.NombreCompleto &&
    usuario.Rol
  ) {
    const userData = {
      email: usuario.Email,
      id: usuario.Id,
      name: usuario.NombreCompleto,
      rol: usuario.Rol,
    };
    setUserProperty(CLAVE_PROPIEDAD_USUARIO, userData);
    Logger.log('Usuario activo establecido: ' + JSON.stringify(userData));
  } else {
    Logger.logError(
      'setActiveUser: Objeto de usuario no válido o incompleto: ' +
        JSON.stringify(usuario)
    );
    console.error('setActiveUser: Objeto de usuario no válido:', usuario);
  }
}

/**
 * @summary Obtiene los datos del usuario actualmente activo de las propiedades de la sesión.
 * @returns {object|null} El objeto de datos del usuario activo, o null si no hay sesión activa.
 */
function getActiveUser() {
  const user = getUserProperty(CLAVE_PROPIEDAD_USUARIO);
  return user;
}

/**
 * @summary Cierra la sesión del usuario activo eliminando sus datos de las propiedades.
 */
function clearActiveUser() {
  deleteUserProperty(CLAVE_PROPIEDAD_USUARIO);
  Logger.log('Usuario activo limpiado de PropertiesService.');
}

// --- CONSTANTES DE RUTAS Y PLANTILLAS ---
const PATH_LOGIN_TEMPLATE = 'View/Login';

// --- LÓGICA DE LOGIN ---

/**
 * @summary Prepara y devuelve el contenido HTML para la página de inicio de sesión.
 * @param {string} [mensaje] Un mensaje opcional para mostrar en la página (ej. "Contraseña incorrecta").
 * @returns {GoogleAppsScript.HTML.HtmlOutput} El objeto HTML listo para ser servido.
 */
function processSignin(mensaje) {
  const plantilla = HtmlService.createTemplateFromFile(PATH_LOGIN_TEMPLATE);
  plantilla.mensajeMensaje = mensaje || '';
  return plantilla
    .evaluate()
    .setFaviconUrl(URL_FAVICON)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1.0')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setTitle('Login Inventario EX')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * @summary Manejador genérico para solicitudes POST. Devuelve un error ya que el flujo de la app no lo usa de forma genérica.
 * @param {object} e El objeto del evento de la solicitud POST.
 * @returns {GoogleAppsScript.Content.TextOutput} Una respuesta JSON de error.
 */
function doPost(e) {
  Logger.log('Solicitud POST recibida. Contenido: ' + JSON.stringify(e));
  return ContentService.createTextOutput(
    JSON.stringify({
      status: 'error',
      message:
        'Las solicitudes POST no son compatibles de forma genérica. Usa funciones específicas expuestas vía google.script.run.',
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * @summary Valida las credenciales de un usuario contra la hoja de 'Usuarios'.
 * @description Busca al usuario, compara la contraseña, verifica que el estado sea 'activo',
 * y si todo es correcto, establece la sesión del usuario y registra la acción.
 * @param {string} nombreUsuario El nombre de usuario proporcionado.
 * @param {string} contrasena La contraseña proporcionada.
 * @returns {object} Un objeto con el estado del inicio de sesión (`status`), un mensaje,
 * y opcionalmente la URL a la que redirigir y el rol del usuario.
 */
function loginCheck(nombreUsuario, contrasena) {
  Logger.log('Intento de inicio de sesión para usuario: ' + nombreUsuario);

  try {
    const hojaCalculo = SpreadsheetApp.openById(ID_INVENTARIO);
    const hoja = hojaCalculo.getSheetByName(HOJA_USUARIO);
    if (!hoja) {
      Logger.logError(
        "loginCheck: No se encontró la hoja de usuarios '" + HOJA_USUARIO + "'."
      );
      return {
        status: false,
        message:
          'Error del sistema: No se encontró la configuración de usuarios.',
      };
    }

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      // Debe haber al menos encabezados y una fila de datos
      Logger.log(
        "loginCheck: No hay usuarios registrados en la hoja '" +
          HOJA_USUARIO +
          "'."
      );
      return { status: false, message: 'No hay usuarios registrados.' };
    }

    const encabezados = datos[0].map(h => String(h).trim().toLowerCase());

    // Mapeo de nombres de propiedad deseados a los nombres de columna en la hoja
    const MAPA_COLUMNAS = {
      id: 'id',
      userName: 'username',
      password: 'password',
      email: 'email',
      estado: 'estado_user',
      rol: 'rol',
      nombreCompleto: 'nombrecompleto',
    };

    const indices = {};
    const columnasFaltantes = [];

    for (const prop in MAPA_COLUMNAS) {
      const nombreColumna = MAPA_COLUMNAS[prop];
      const index = encabezados.indexOf(nombreColumna);
      indices[prop] = index;
      if (index === -1) {
        columnasFaltantes.push(nombreColumna);
      }
    }

    if (columnasFaltantes.length > 0) {
      const mensajeError = `loginCheck: Faltan las siguientes columnas requeridas en la hoja de usuarios: ${columnasFaltantes.join(
        ', '
      )}. Encabezados encontrados: ${encabezados.join(', ')}`;
      Logger.logError(mensajeError);
      return {
        status: false,
        message:
          'Error en la configuración del sistema de usuarios. Contacte al administrador.',
      };
    }

    const usuarios = datos.slice(1);
    const filaUsuario = usuarios.find(
      fila =>
        fila[indices.userName] &&
        typeof fila[indices.userName] === 'string' &&
        fila[indices.userName].trim() === nombreUsuario &&
        fila[indices.password] &&
        String(fila[indices.password]) === String(contrasena) // Comparar como strings por si las contraseñas son numéricas
    );

    if (!filaUsuario) {
      logAction(
        nombreUsuario,
        'Intento de inicio de sesión fallido: Usuario/Contraseña incorrectos',
        hojaCalculo
      );
      return { status: false, message: 'Usuario o contraseña incorrectos.' };
    }

    if (String(filaUsuario[indices.estado]).trim().toLowerCase() !== 'activo') {
      logAction(
        nombreUsuario,
        'Intento de inicio de sesión fallido: Cuenta inactiva',
        hojaCalculo
      );
      return {
        status: false,
        message: 'Tu cuenta está inactiva. Contacta al administrador.',
      };
    }

    // Asegurar que las propiedades cruciales para setActiveUser estén presentes y con el nombre esperado
    const usuarioParaSesion = {
      Email: filaUsuario[indices.email],
      Id: filaUsuario[indices.id],
      NombreCompleto: filaUsuario[indices.nombreCompleto],
      Rol: filaUsuario[indices.rol],
    };

    if (
      !usuarioParaSesion.Email ||
      !usuarioParaSesion.Id ||
      !usuarioParaSesion.NombreCompleto ||
      !usuarioParaSesion.Rol
    ) {
      Logger.logError(
        'loginCheck: Datos de usuario incompletos después de encontrar la fila. ' +
          JSON.stringify(usuarioParaSesion)
      );
      logAction(
        nombreUsuario,
        'Error de inicio de sesión: datos de usuario incompletos',
        hojaCalculo
      );
      return {
        status: false,
        message:
          'Error al obtener los detalles del usuario. Contacte al administrador.',
      };
    }

    setActiveUser(usuarioParaSesion); // Establece el usuario en PropertiesService
    logAction(nombreUsuario, 'Inicio de sesión exitoso', hojaCalculo);

    return {
      status: true,
      message: 'Inicio de sesión correcto',
      // Redirige a Home. doGet se encargará de verificar si Home es accesible para este rol.
      page: getScriptUrl() + '?page=Index',
      rol: usuarioParaSesion.Rol, // Devolver el rol para posible uso en el cliente
    };
  } catch (error) {
    Logger.logError(
      "Error durante la verificación de inicio de sesión para '" +
        nombreUsuario +
        "': " +
        error.toString() +
        '\nStack: ' +
        error.stack
    );
    // Intentar abrir la hoja de cálculo solo si no se ha hecho ya, para loguear el error.
    let ssForLog;
    try {
      ssForLog = SpreadsheetApp.openById(ID_INVENTARIO);
    } catch (e) {
      Logger.logError(
        'No se pudo abrir el Spreadsheet para loguear el error de login: ' +
          e.toString()
      );
    }
    if (ssForLog) {
      logAction(
        nombreUsuario || 'Desconocido',
        `Error crítico de inicio de sesión: ${error.message}`,
        ssForLog
      );
    }
    return {
      status: false,
      message:
        'Ocurrió un error inesperado durante el inicio de sesión. Por favor, inténtalo de nuevo más tarde.',
    };
  }
}
//=====================================
// --- RECUPERACIÓN DE CONTRASEÑA ---
//=======================================
/**
 * @summary Gestiona la solicitud de recuperación de contraseña de un usuario.
 * @description Busca al usuario por su correo electrónico. Si el usuario existe y está activo,
 * le envía un correo con sus credenciales. Si la cuenta está inactiva, le notifica de ello.
 * @param {string} correo El correo electrónico del usuario que solicita la recuperación.
 * @returns {object} Un objeto con el estado (`status`) de la solicitud y un mensaje para el usuario.
 */
function recoverPassword(correo) {
  Logger.log(
    'Solicitud de recuperación de contraseña para el correo: ' + correo
  );
  try {
    const hojaCalculo = SpreadsheetApp.openById(ID_INVENTARIO);
    const hoja = hojaCalculo.getSheetByName(HOJA_USUARIO);
    if (!hoja) {
      Logger.log('recoverPassword: No se encontró la hoja de usuarios.');
      return {
        status: false,
        message: '⚠️ No se encontró la hoja de usuarios.',
      };
    }
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      // Encabezados + al menos un usuario
      Logger.log('recoverPassword: No hay datos de usuarios en la hoja.');
      return { status: false, message: '⚠️ No hay datos de usuarios.' };
    }

    const encabezados = datos[0].map(h => String(h).trim().toLowerCase());
    const usuarios = datos.slice(1);

    // Mapeo consistente con loginCheck
    const MAPA_COLUMNAS = {
      email: 'email',
      userName: 'username',
      password: 'password',
      estado: 'estado_user',
      nombreCompleto: 'nombrecompleto',
    };

    const indices = {};
    const columnasFaltantes = [];

    for (const prop in MAPA_COLUMNAS) {
      const nombreColumna = MAPA_COLUMNAS[prop];
      const index = encabezados.indexOf(nombreColumna);
      indices[prop] = index;
      if (index === -1) {
        columnasFaltantes.push(nombreColumna);
      }
    }

    if (columnasFaltantes.length > 0) {
      Logger.logError(
        `recoverPassword: Error en la estructura de la hoja. Faltan columnas: ${columnasFaltantes.join(
          ', '
        )}`
      );
      return {
        status: false,
        message: '⚠️ Error en la estructura de la hoja de datos de usuario.',
      };
    }

    const filaUsuario = usuarios.find(
      fila =>
        fila[indices.email] &&
        String(fila[indices.email]).trim().toLowerCase() ===
          String(correo).trim().toLowerCase()
    );

    if (!filaUsuario) {
      Logger.log('recoverPassword: Correo no registrado: ' + correo);
      logAction(
        correo,
        'Intento de recuperación de contraseña fallido: Correo no registrado',
        hojaCalculo
      );
      return { status: false, message: '⚠️ Correo no registrado.' };
    }

    const nombreDelUsuario =
      filaUsuario[indices.nombreCompleto] || filaUsuario[indices.userName]; // Usar NombreCompleto si existe, sino userName
    const contrasena = filaUsuario[indices.password];
    const estado = String(filaUsuario[indices.estado]).trim().toLowerCase();

    const opcionesCorreo = {
      to: correo,
      subject: '',
      htmlBody: '', // Usar htmlBody para un formato más amigable
    };

    // =================================================================================================
    // ADVERTENCIA DE SEGURIDAD CRÍTICA
    // Enviar contraseñas en texto plano por correo electrónico es una VULNERABILIDAD GRAVE.
    // 1. Las contraseñas NUNCA deben almacenarse en texto plano. Deben ser "hasheadas" usando un algoritmo seguro (ej. bcrypt).
    // 2. La recuperación de contraseña NUNCA debe enviar la contraseña actual. El procedimiento correcto es
    //    enviar un enlace de un solo uso con un token de tiempo limitado que permita al usuario ESTABLECER UNA NUEVA contraseña.
    // Este código se mantiene para no romper la funcionalidad existente, pero DEBE SER REEMPLAZADO.
    // =================================================================================================
    if (estado !== 'activo') {
      opcionesCorreo.subject =
        '⚠️ Intento de recuperación de contraseña: Cuenta Inactiva';
      opcionesCorreo.htmlBody = `
        <p>Hola ${nombreDelUsuario || 'usuario'},</p>
        <p>Hemos recibido una solicitud para recuperar la contraseña de tu cuenta asociada a este correo electrónico.</p>
        <p>Actualmente, <strong>tu cuenta se encuentra inactiva</strong>. Por favor, contacta con el administrador del sistema para reactivar tu cuenta.</p>
        <p>Saludos,<br>El equipo de Soporte del Inventario</p>
        <hr><p style="font-size:0.8em; color:grey;">Este es un correo automático, no es necesario responder.</p>`;
      MailApp.sendEmail(opcionesCorreo);
      logAction(
        correo,
        'Intento de recuperación de contraseña: Cuenta inactiva',
        hojaCalculo
      );
      return {
        status: false,
        message:
          '⚠️ Tu cuenta está inactiva. Se ha enviado una notificación a tu correo con más detalles.',
      };
    }

    opcionesCorreo.subject =
      '🔑 Recuperación de Credenciales de Acceso - Inventario';
    opcionesCorreo.htmlBody = `
      <p>Hola ${nombreDelUsuario || 'usuario'},</p>
      <p>Hemos recibido una solicitud para recuperar tus credenciales de acceso al sistema de inventario.</p>
      <p>Aquí están tus datos:</p>
      <ul>
        <li><strong>Nombre de Usuario:</strong> ${
          filaUsuario[indices.nombreUsuario]
        }</li>
        <li><strong>Contraseña:</strong> ${contrasena}</li>
      </ul>
      <p>Te recomendamos cambiar tu contraseña después de iniciar sesión si crees que tu cuenta pudo haber sido comprometida o si esta es una contraseña temporal.</p>
      <p>Puedes acceder al sistema aquí: <a href="${getScriptUrl()}?page=Login">${getScriptUrl()}?page=Login</a></p>
      <p>Saludos,<br>El equipo de Soporte del Inventario</p>
      <hr><p style="font-size:0.8em; color:grey;">Este es un correo automático, no es necesario responder.</p>`;
    MailApp.sendEmail(opcionesCorreo);
    logAction(
      correo,
      'Recuperación de contraseña exitosa (correo enviado)',
      hojaCalculo
    );
    return {
      status: true,
      message:
        '📧 Se ha enviado un correo electrónico con tus credenciales a ' +
        correo +
        '.',
    };
  } catch (error) {
    Logger.logError(
      "Error en recoverPassword para '" +
        correo +
        "': " +
        error.toString() +
        '\nStack: ' +
        error.stack
    );
    // No loguear la acción aquí porque podría ser el mismo error que impidió abrir la hoja.
    return {
      status: false,
      message: `⚠️ Ocurrió un error al procesar tu solicitud de recuperación: ${error.message}. Intenta de nuevo más tarde.`,
    };
  }
}

/**
 * @summary Devuelve el usuario activo guardado en la sesión.
 * @returns {string|null} El usuario en JSON o null si no hay sesión.
 */
function getCurrentUser() {
  const usuario = getActiveUser(); // 👈 ya lo tienes implementado
  return usuario ? JSON.stringify(usuario) : null;
}

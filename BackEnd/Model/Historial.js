// =================================================================
// --- MODELO DE HISTORIAL ---
// =================================================================

const HistorialModel = {};

/**
 * @summary Obtiene la hoja 'Historial'. Si no existe, la crea con sus cabeceras.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Objeto de la hoja de historial.
 * @private
 */
function getHistorialSheet() {
  try {
    Logger.log('Intentando obtener spreadsheet...');
    const ss = getSpreadsheet();
    Logger.log('Spreadsheet obtenida');

    Logger.log('Buscando hoja de historial...');
    let sheet = ss.getSheetByName(getHojasConfig().HISTORIAL.nombre);
    Logger.log(`Resultado búsqueda de hoja: ${sheet ? 'encontrada' : 'no encontrada'}`);

    if (!sheet) {
      Logger.log('Creando nueva hoja de historial...');
      sheet = ss.insertSheet(getHojasConfig().HISTORIAL.nombre);
      sheet.appendRow([
        "Id",
        "ProductoId",
        "Fecha y Hora",
        "Producto",
        "Programa",
        "Unidades Anteriores",
        "Unidades Nuevas",
        "Estado",
        "Usuario",
        "Fecha de Retiro",
        "Cantidad",
        "Origen",
      ]);
      Logger.log('Nueva hoja creada con cabeceras');
    }

    return sheet;
  } catch (error) {
    Logger.log(`Error en getHistorialSheet: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw new Error(`Error al acceder a la hoja de Historial: ${error.message}`);
  }
}

/**
 * @summary Obtiene todos los registros de la hoja 'Historial' y los formatea como arreglo de objetos.
 * @returns {Array<object>} Arreglo de registros del historial.
 */
HistorialModel.obtenerDatos = function () {
  try {
    Logger.log('Iniciando HistorialModel.obtenerDatos...');
    const sheet = getHistorialSheet();

    if (!sheet) {
      Logger.log('No se pudo obtener la hoja de historial');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log(`Filas totales en la hoja: ${data.length}`);

    if (data.length < 2) {
      Logger.log('La hoja está vacía o solo contiene cabeceras');
      return [];
    }

    const headers = data[0];
    Logger.log('Cabeceras encontradas: ' + headers.join(', '));

    const items = data.slice(1).map(row => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row[index];
      });
      return record;
    });

    Logger.log(`Registros procesados: ${items.length}`);
    return items;
  } catch (error) {
    Logger.log(`Error en HistorialModel.obtenerDatos: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    throw new Error(`Error al obtener datos del historial: ${error.message}`);
  }
};

// =================================================================
// --- CONTROLADOR DE HISTORIAL ---
// =================================================================

/**
 * @summary Endpoint para obtener todos los registros del historial.
 * @returns {Array<object>} Registros del historial.
 */
function getHistorialData() {
  try {
    Logger.log('getHistorialData - Iniciando...');

    const activeUser = getActiveUser();
    Logger.log('getHistorialData - Usuario activo: ' + JSON.stringify(activeUser));

    if (!activeUser) {
      throw new Error('Usuario no autenticado');
    }

    // Verificar permisos
    if (!isPageAllowedForUser("Historial", activeUser.rol)) {
      throw new Error('No tienes permisos para ver el historial');
    }

    // Obtener datos del modelo
    const data = HistorialModel.obtenerDatos();
    Logger.log(`getHistorialData - Datos obtenidos: ${data.length} registros`);

    // Devolver en formato JSON igual que Registro.js
    return JSON.stringify({ data: data });

  } catch (error) {
    Logger.log(`Error en getHistorialData: ${error.message}`);
    throw new Error(`Error al obtener datos del historial: ${error.message}`);
  }
}

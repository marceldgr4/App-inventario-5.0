/**
 * @fileoverview Contiene funciones de utilidad para las operaciones CRUD,
 * como la búsqueda de filas, el cálculo de fechas y la obtención de hojas.
 * Estas funciones son utilizadas por otros módulos para interactuar con Google Sheets.
 * @summary Se conecta a la hoja de cálculo de Google y obtiene una hoja específica por su nombre.
 * @param {string} name El nombre exacto de la hoja que se desea obtener (ej. "Articulos").
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} El objeto de la hoja si se encuentra, o null si no existe.
 */
function getSheetByName(name) {
  Logger.log(
    `DEBUG: getSheetByName - Intentando abrir Spreadsheet con ID: ${ID_INVENTARIO}`
  );
  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  Logger.log(`DEBUG: getSheetByName - Buscando hoja con nombre: '${name}'`);
  const sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`DEBUG: getSheetByName - Hoja '${name}' encontrada.`);
  } else {
    Logger.log(`DEBUG: getSheetByName - Hoja '${name}' NO encontrada.`);
  }
  return sheet;
}

/**
 * @summary Busca en una hoja específica el número de fila que corresponde a un Id único usando TextFinder (Optimizado).
 * @param {string|number} id El identificador único del registro que se está buscando.
 * @param {string} sheetName El nombre de la hoja donde se realizará la búsqueda.
 * @returns {number} El número de la fila (base 1) si se encuentra una coincidencia. Devuelve -1 si no se encuentra.
 */
/**
 * @summary Busca en una hoja específica el número de fila que corresponde a un Id único usando TextFinder (Optimizado).
 * @param {string|number} id El identificador único del registro que se está buscando.
 * @param {string} sheetName El nombre de la hoja donde se realizará la búsqueda.
 * @returns {number} El número de la fila (base 1) si se encuentra una coincidencia. Devuelve -1 si no se encuentra.
 */
function findRowById(id, sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) return -1;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idIndex = headers.findIndex((header) => header.toUpperCase() === "ID");

  if (idIndex === -1) {
    Logger.log(
      `findRowById: No se encontró la columna de ID en la hoja "${sheetName}".`
    );
    return -1;
  }

  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return -1;

  const range = sheet.getRange(2, idIndex + 1, numRows, 1);
  const textFinder = range.createTextFinder(String(id)).matchEntireCell(true);
  const found = textFinder.findNext();
  return found ? found.getRow() : -1;
}

/**
 * @summary Calcula cuántos días han transcurrido desde una fecha de ingreso hasta el día de hoy.
 * @param {Date} fechaIngreso El objeto de fecha que representa cuándo ingresó el producto.
 * @returns {number} El número de días transcurridos. Devuelve 0 si la fecha no es válida.
 */
function calcularTiempoEnStorage(fechaIngreso) {
  if (!(fechaIngreso instanceof Date) || isNaN(fechaIngreso.getTime())) {
    return 0;
  }
  const hoy = new Date();
  const utcHoy = Date.UTC(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  const utcFechaIngreso = Date.UTC(
    fechaIngreso.getFullYear(),
    fechaIngreso.getMonth(),
    fechaIngreso.getDate()
  );
  const diffMs = utcHoy - utcFechaIngreso;
  const diffDias = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  return diffDias >= 0 ? diffDias : 0;
}

/**
 * @summary Función de utilidad para convertir un número de columna (ej. 1, 27) a su letra correspondiente en la hoja de cálculo (ej. 'A', 'AA').
 * @param {number} column El número de la columna (base 1).
 * @returns {string} La letra o letras correspondientes a la columna.
 */
function columnToLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function _getInventoryDataForSheet(sheetName) {
  Logger.log(
    `_getInventoryDataForSheet: Obteniendo datos de la hoja '${sheetName}'`
  );
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(
      `_getInventoryDataForSheet: La hoja '${sheetName}' no fue encontrada.`
    );
    return [];
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length < 2) {
    Logger.log(
      `_getInventoryDataForSheet: No hay datos en la hoja '${sheetName}' (solo cabeceras o vacía).`
    );
    return [];
  }

  const headers = values[0].map((header) => header.trim());
  const idIndex = headers.indexOf("Id");
  const fechaIngresoIndex = headers.indexOf("FECHA DE INGRESO");

  if (idIndex === -1) {
    Logger.log(
      `_getInventoryDataForSheet: No se encontró la columna 'Id' en la hoja '${sheetName}'.`
    );
    return [];
  }

  const data = values
    .slice(1)
    .map((row, index) => {
      const obj = {};
      let hasData = false;
      headers.forEach((header, i) => {
        let value = row[i];
        if (value !== "" && value !== null && value !== undefined) {
          hasData = true;
        }
        if (header.toLowerCase().includes("fecha") && value) {
          let date = value instanceof Date ? value : new Date(value);
          if (!isNaN(date.getTime())) {
            obj[header] = date.toISOString();
          } else {
            obj[header] = value; // Mantener el valor original si no es una fecha válida
          }
        } else {
          obj[header] = value;
        }
      });

      if (!obj["Id"]) {
        obj["Id"] = `temp_${index + 2}`;
      }

      if (fechaIngresoIndex !== -1) {
        const fechaIngreso = obj["FECHA DE INGRESO"];
        if (fechaIngreso && fechaIngreso instanceof Date) {
          obj["Tiempo en Storage"] = calcularTiempoEnStorage(fechaIngreso);
        } else {
          obj["Tiempo en Storage"] = 0;
        }
      }

      return hasData ? obj : null;
    })
    .filter((obj) => obj !== null);

  Logger.log(
    `_getInventoryDataForSheet: Se procesaron ${data.length} registros de la hoja '${sheetName}'.`
  );
  return data;
}

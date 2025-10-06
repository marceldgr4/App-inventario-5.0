/**
 * @summary Agrega un producto con una imagen opcional a la hoja de 'Articulos'.
 * @description Esta función maneja la subida de un archivo de imagen a Google Drive,
 * obtiene su URL, y luego registra todos los datos del producto, incluyendo la URL de la imagen,
 * en la hoja de cálculo. También calcula un ID único y establece fórmulas dinámicas.
 * @param {object} productData Objeto con los datos del producto (productoAgregar, programaAgregar, etc.).
 * @param {object} [fileData] Opcional. Objeto con los datos del archivo de imagen (base64Data, fileName, mimeType).
 * @returns {string} Una cadena JSON con el resultado de la operación (éxito/error, mensaje, nuevo ID y URL de la imagen).
 */
function agregarProductoConImagenDesdeArchivo(productData, fileData) {
  try {
    let imageUrl = '';

    if (
      fileData &&
      fileData.base64Data &&
      fileData.fileName &&
      fileData.mimeType
    ) {
      const folder = DriveApp.getFolderById(ID_DRIVE_IMG);
      const decodedData = Utilities.base64Decode(
        fileData.base64Data,
        Utilities.Charset.UTF_8
      );
      const blob = Utilities.newBlob(
        decodedData,
        fileData.mimeType,
        fileData.fileName
      );

      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
      Logger.log(
        'Imagen subida a Drive, URL: ' +
          imageUrl +
          ' (Archivo ID: ' +
          file.getId() +
          ')'
      );

      try {
        // Intentar establecer los permisos para que cualquiera con el enlace pueda ver
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
        Logger.log(
          'Permisos de compartición del archivo actualizados exitosamente para: ' +
            file.getName()
        );
      } catch (sharingError) {
        // Si falla, solo registrar una advertencia pero no detener la ejecución
        Logger.log(
          "ADVERTENCIA: No se pudieron establecer los permisos de compartición para el archivo '" +
            file.getName() +
            "'. Error: " +
            sharingError.toString() +
            '. La URL podría no ser públicamente accesible. Continuando con el registro en la hoja...'
        );
      }
    }

    const sheet = getSheetByName(HOJA_ARTICULOS);
    if (!sheet) {
      throw new Error(
        `La hoja especificada "${HOJA_ARTICULOS}" no fue encontrada.`
      );
    }
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const newRowValues = new Array(headers.length).fill('');

    // Calcular nuevo ID
    const newId = _obtenerSiguienteIdNumerico(sheet, headers);
    const idHeaderIndex = headers.indexOf('Id');
    if (idHeaderIndex !== -1) {
        newRowValues[idHeaderIndex] = newId;
    }

    // Mapeo de datos a columnas
    const columnMapping = {
      PRODUCTO: productData.productoAgregar,
      PROGRAMA: productData.programaAgregar,
      'FECHA DE INGRESO': new Date(),
      Ingresos: parseFloat(productData.ingresosAgregar) || 0,
      Salidas: 0,
      Estado: 'Activo',
      'Tiempo en Storage': 0,
      Imagen: imageUrl,
      'Entregas fecha': '',
      'Entregas cantidad': 0,
    };

    headers.forEach((header, index) => {
      if (columnMapping.hasOwnProperty(header)) {
        newRowValues[index] = columnMapping[header];
      }
    });

    // Fórmula para 'Unidades disponibles'
    const ingresosHeaderIdx = headers.indexOf('Ingresos');
    const salidasHeaderIdx = headers.indexOf('Salidas');
    const unidadesDispHeaderIdx = headers.indexOf('Unidades disponibles');
    const targetRowNumForFormula = sheet.getLastRow() + 1;

    if (
      unidadesDispHeaderIdx !== -1 &&
      ingresosHeaderIdx !== -1 &&
      salidasHeaderIdx !== -1
    ) {
      const ingresosColLetter = columnToLetter(ingresosHeaderIdx);
      const salidasColLetter = columnToLetter(salidasHeaderIdx);
      newRowValues[
        unidadesDispHeaderIdx
      ] = `=IF(ISNUMBER(${ingresosColLetter}${targetRowNumForFormula}),${ingresosColLetter}${targetRowNumForFormula},0)-IF(ISNUMBER(${salidasColLetter}${targetRowNumForFormula}),${salidasColLetter}${targetRowNumForFormula},0)`;
    } else if (unidadesDispHeaderIdx !== -1) {
      newRowValues[unidadesDispHeaderIdx] =
        parseFloat(productData.ingresosAgregar) || 0;
    }

    sheet.appendRow(newRowValues);
    const appendedRowIndex = sheet.getLastRow();

    // Actualizar "Tiempo en Storage"
    const tiempoStorageHeaderIdx = headers.indexOf('Tiempo en Storage');
    const fechaIngresoHeaderIdx = headers.indexOf('FECHA DE INGRESO');
    if (tiempoStorageHeaderIdx !== -1 && fechaIngresoHeaderIdx !== -1) {
      const fechaIngresoValue = sheet
        .getRange(appendedRowIndex, fechaIngresoHeaderIdx + 1)
        .getValue();
      if (fechaIngresoValue instanceof Date) {
        const tiempoEnStorage = calcularTiempoEnStorage(fechaIngresoValue);
        sheet
          .getRange(appendedRowIndex, tiempoStorageHeaderIdx + 1)
          .setValue(tiempoEnStorage);
      }
    }

    const activeUser = getActiveUser();
    logAction(
      activeUser ? activeUser.name : 'Usuario no identificado',
      `Producto agregado con imagen: ${productData.productoAgregar} a ${HOJA_ARTICULOS}`
    );

    return JSON.stringify({
      success: true,
      message: `Producto "${productData.productoAgregar}" agregado con imagen a ${HOJA_ARTICULOS} exitosamente.`,
      newId: newId,
      imageUrl: imageUrl,
    });
  } catch (e) {
    Logger.log(
      `Error en agregarProductoConImagenDesdeArchivo: ${e.toString()}\nStack: ${
        e.stack || 'No disponible'
      }`
    );
    return JSON.stringify({
      success: false,
      message: `Error al agregar producto con imagen: ${e.message}`,
    });
  }
}

/**
 * @summary Convierte un índice de columna (base 0) a su letra correspondiente en la hoja de cálculo (A, B, C...).
 * @param {number} column El índice de la columna basado en 0.
 * @returns {string} La letra o letras de la columna.
 */
function columnToLetter(column) {
  let temp,
    letter = '';
  let colNum = column + 1; // Convierte de base 0 a base 1 para el cálculo
  while (colNum > 0) {
    temp = (colNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colNum = (colNum - temp - 1) / 26;
  }
  return letter;
}


/**
 * @summary Calcula el siguiente ID numérico secuencial para una nueva fila en una hoja.
 * @description Busca el valor máximo en la columna 'Id' y le suma 1.
 * Maneja correctamente hojas con filas congeladas y sin datos.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja en la que se buscará el último ID.
 * @param {Array<string>} headers Un arreglo con los encabezados de la hoja.
 * @returns {number} El siguiente ID numérico disponible.
 * @private
 */
function _obtenerSiguienteIdNumerico(sheet, headers) {
  let newId = 1;
  const idHeaderName = 'Id';
  const idHeaderIndex = headers.indexOf(idHeaderName);

  if (idHeaderIndex !== -1) {
    const lastDataRow = sheet.getLastRow();
    const firstDataRowIndex = (sheet.getFrozenRows() || 0) + 1;

    if (lastDataRow >= firstDataRowIndex) {
      const ids = sheet
        .getRange(
          firstDataRowIndex,
          idHeaderIndex + 1,
          lastDataRow - firstDataRowIndex + 1,
          1
        )
        .getValues()
        .flat()
        .map(id => parseInt(id))
        .filter(id => !isNaN(id) && id !== null && id !== '');
      if (ids.length > 0) {
        newId = Math.max(...ids) + 1;
      }
    }
  } else {
    Logger.log(
      `ADVERTENCIA: La columna de encabezado "${idHeaderName}" no se encontró en la hoja "${sheet.getName()}". Se usará el ID inicial 1.`
    );
  }
  return newId;
}

/**
 * @summary Agrega un producto de papelería con una imagen opcional a su hoja correspondiente.
 * @param {object} productData Los datos del producto.
 * @param {object} [fileData] Los datos del archivo de imagen (opcional).
 * @returns {string} Un JSON con el resultado de la operación.
 */
function agregarPapeleriaConImagenDesdeArchivo(productData, fileData) {
  try {
    let imageUrl = '';
    const activeUser = getActiveUser();
    const usuarioLogueado = activeUser ? activeUser.name : 'Sistema';

    if (
      fileData &&
      fileData.base64Data &&
      fileData.fileName &&
      fileData.mimeType
    ) {
      const folder = DriveApp.getFolderById(ID_PAPELERIA_IMG);
      const decodedData = Utilities.base64Decode(
        fileData.base64Data,
        Utilities.Charset.UTF_8
      );
      const blob = Utilities.newBlob(
        decodedData,
        fileData.mimeType,
        fileData.fileName
      );
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
      Logger.log(
        `Imagen de papelería subida a Drive, URL: ${imageUrl} (Archivo ID: ${file.getId()})`
      );
      try {
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
      } catch (sharingError) {
        Logger.log(
          `ADVERTENCIA: Permisos de compartición para archivo de papelería '${file.getName()}' fallaron. Error: ${sharingError.toString()}`
        );
      }
    }

    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(HOJA_PAPELERIA);
    if (!sheet) {
      throw new Error(
        `La hoja especificada "${HOJA_PAPELERIA}" no fue encontrada.`
      );
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const newRowValues = new Array(headers.length).fill('');
    const newId = _obtenerSiguienteIdNumerico(sheet, headers);

    const dataParaHoja = {
      Id: newId,
      PRODUCTO: productData.productoAgregar,
      PROGRAMA: productData.programaAgregar || '',
      'FECHA DE INGRESO': new Date(),
      Ingresos: parseFloat(productData.ingresosAgregar) || 0,
      Salidas: 0,
      'Tiempo en Storage': 0,
      Imagen: imageUrl,
      Estado: 'Activo',
    };

    headers.forEach((header, index) => {
      if (dataParaHoja.hasOwnProperty(header)) {
        newRowValues[index] = dataParaHoja[header];
      }
    });

    const ingresosIdx = headers.indexOf('Ingresos');
    const salidasIdx = headers.indexOf('Salidas');
    const unidadesDispIdx = headers.indexOf('Unidades disponibles');
    if (unidadesDispIdx !== -1) {
      if (ingresosIdx !== -1 && salidasIdx !== -1) {
        const targetRowNumForFormula = sheet.getLastRow() + 1;
        const ingresosColLetter = columnToLetter(ingresosIdx);
        const salidasColLetter = columnToLetter(salidasIdx);
        newRowValues[
          unidadesDispIdx
        ] = `=IF(ISNUMBER(${ingresosColLetter}${targetRowNumForFormula}),${ingresosColLetter}${targetRowNumForFormula},0)-IF(ISNUMBER(${salidasColLetter}${targetRowNumForFormula}),${salidasColLetter}${targetRowNumForFormula},0)`;
      } else if (dataParaHoja['Ingresos'] !== undefined) {
        newRowValues[unidadesDispIdx] = dataParaHoja['Ingresos'];
      }
    }

    sheet.appendRow(newRowValues);
    const appendedRowIndex = sheet.getLastRow();
    // ... (lógica para recalcular Tiempo en Storage si es necesario) ...

    logAction(
      usuarioLogueado,
      `Agregado papelería con imagen: ${productData.productoAgregar} (ID: ${newId}) a ${HOJA_PAPELERIA}`
    );

    return JSON.stringify({
      success: true,
      message: `Producto de papelería ("${productData.productoAgregar}") se ha agregado a ${HOJA_PAPELERIA} exitosamente.`,
      newId: newId,
      imageUrl: imageUrl,
    });
  } catch (e) {
    Logger.log(
      `Error en agregarPapeleriaConImagenDesdeArchivo: ${e.toString()}\nStack: ${
        e.stack || 'No disponible'
      }`
    );
    return JSON.stringify({
      success: false,
      message: `Error al agregar producto de papelería con imagen: ${e.message}`,
    });
  }
}

/**
 * @summary Agrega un nuevo producto a la hoja de 'Papeleria'. Esta versión no maneja la subida de archivos.
 * @param {object} productData Los datos del producto, puede incluir una URL de imagen existente en `imagenAgregar`.
 * @returns {string} Un JSON con el resultado de la operación.
 */
function agregarPapeleria(productData) {
  try {
    const activeUser = getActiveUser();
    const usuarioLogueado = activeUser ? activeUser.name : 'Sistema';
    const imageUrl = productData.imagenAgregar || '';

    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(HOJA_PAPELERIA);
    if (!sheet) {
      throw new Error(
        `La hoja especificada "${HOJA_PAPELERIA}" no fue encontrada.`
      );
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const newRowValues = new Array(headers.length).fill('');
    const newId = _obtenerSiguienteIdNumerico(sheet, headers);

    const dataParaHoja = {
      Id: newId,
      PRODUCTO: productData.productoAgregar,
      PROGRAMA: productData.programaAgregar || '',
      'FECHA DE INGRESO': new Date(),
      Ingresos: parseFloat(productData.ingresosAgregar) || 0,
      Salidas: 0,
      'Tiempo en Storage': 0,
      Imagen: imageUrl,
      Estado: 'Activo',
    };

    headers.forEach((header, index) => {
      if (dataParaHoja.hasOwnProperty(header)) {
        newRowValues[index] = dataParaHoja[header];
      }
    });

    const ingresosIdx = headers.indexOf('Ingresos');
    const salidasIdx = headers.indexOf('Salidas');
    const unidadesDispIdx = headers.indexOf('Unidades disponibles');
    if (unidadesDispIdx !== -1) {
      if (ingresosIdx !== -1 && salidasIdx !== -1) {
        const targetRowNumForFormula = sheet.getLastRow() + 1;
        const ingresosColLetter = columnToLetter(ingresosIdx);
        const salidasColLetter = columnToLetter(salidasIdx);
        newRowValues[
          unidadesDispIdx
        ] = `=IF(ISNUMBER(${ingresosColLetter}${targetRowNumForFormula}),${ingresosColLetter}${targetRowNumForFormula},0)-IF(ISNUMBER(${salidasColLetter}${targetRowNumForFormula}),${salidasColLetter}${targetRowNumForFormula},0)`;
      } else if (dataParaHoja['Ingresos'] !== undefined) {
        newRowValues[unidadesDispIdx] = dataParaHoja['Ingresos'];
      }
    }

    sheet.appendRow(newRowValues);
    // ... (lógica de Tiempo en Storage si aplica) ...

    logAction(
      usuarioLogueado,
      `Agregado producto de papelería: ${productData.productoAgregar} (ID: ${newId}) a ${HOJA_PAPELERIA}`
    );

    return JSON.stringify({
      success: true,
      message: `Producto de papelería ("${productData.productoAgregar}") se ha agregado a ${HOJA_PAPELERIA} exitosamente.`,
      newId: newId,
      imageUrl: imageUrl,
    });
  } catch (e) {
    Logger.log(
      `Error en agregarPapeleria: ${e.toString()}\nStack: ${
        e.stack || 'No disponible'
      }`
    );
    return JSON.stringify({
      success: false,
      message: `Error al agregar producto de papelería: ${e.message}`,
    });
  }
}

/**
 * @summary Agrega un nuevo producto de decoración, manejando opcionalmente la subida de una imagen.
 * @description Función unificada que sube una imagen a Drive si se proporciona, y luego añade los
 * datos del producto a la hoja 'Decoracion'. Incluye manejo de errores para permisos de archivo.
 * @param {object} productData Objeto con los datos del producto de decoración.
 * @param {object} [fileData] Opcional. Objeto con los datos del archivo de imagen.
 * @returns {string} Un JSON con el resultado de la operación.
 */
function agregarDecoracion(productData, fileData) {
  try {
    let imageUrl = '';
    const activeUser = getActiveUser();
    const usuarioLogueado = activeUser ? activeUser.name : 'Sistema';

    if (fileData && fileData.base64Data) {
      const folder = DriveApp.getFolderById(ID_DECORACION_IMG);
      const bytes = Utilities.base64Decode(fileData.base64Data);
      const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.fileName);
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingError) {
        Logger.log(`ADVERTENCIA: Permisos de compartición para '${file.getName()}' fallaron. Error: ${sharingError.toString()}`);
      }
    }

    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(HOJA_DECORACION);
    if (!sheet) {
      throw new Error(`La hoja especificada "${HOJA_DECORACION}" no fue encontrada.`);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Encabezados leídos de la hoja de Decoración: " + JSON.stringify(headers));

    const newRowValues = new Array(headers.length).fill('');
    const newId = _obtenerSiguienteIdNumerico(sheet, headers);

    const dataParaHoja = {
      'Id': newId,
      'PRODUCTO': productData.productoAgregar,
      'PRECIO': parseFloat(productData.precioAgregar) || 0,
      'TIPO': productData.tipoAgregar,
      'FECHA DE INGRESO': new Date(),
      'Ingresos': parseInt(productData.ingresosAgregar) || 0,
      'Salidas': 0,
      'Comentarios': productData.comentariosAgregar,
      'FECHA DE ACTUALIZACION': '',
      'Imagen': imageUrl,
      'Estado': 'Activo'
    };

    headers.forEach((header, index) => {
      if (dataParaHoja.hasOwnProperty(header)) {
        newRowValues[index] = dataParaHoja[header];
      }
    });

    const ingresosIdx = headers.indexOf('Ingresos');
    const salidasIdx = headers.indexOf('Salidas');
    const unidadesDispIdx = headers.indexOf('Unidades disponibles');
    if (unidadesDispIdx !== -1) {
      if (ingresosIdx !== -1 && salidasIdx !== -1) {
        const targetRowNum = sheet.getLastRow() + 1;
        const ingresosCol = columnToLetter(ingresosIdx);
        const salidasCol = columnToLetter(salidasIdx);
        newRowValues[unidadesDispIdx] = `=IF(ISNUMBER(${ingresosCol}${targetRowNum}),${ingresosCol}${targetRowNum},0)-IF(ISNUMBER(${salidasCol}${targetRowNum}),${salidasCol}${targetRowNum},0)`;
      } else {
        newRowValues[unidadesDispIdx] = parseInt(productData.ingresosAgregar) || 0;
      }
    }

    sheet.appendRow(newRowValues);
    logAction(usuarioLogueado, `Agregado Decoración: ${productData.productoAgregar} (ID: ${newId})`);

    return JSON.stringify({
      success: true,
      message: `Producto "${productData.productoAgregar}" agregado a Decoración exitosamente.`
    });

  } catch (e) {
    Logger.log(`ERROR CRÍTICO en agregarDecoracion: ${e.toString()}\nStack: ${e.stack}`);
    return JSON.stringify({
      success: false,
      message: `Error en el servidor al agregar decoración: ${e.message}`
    });
  }
}
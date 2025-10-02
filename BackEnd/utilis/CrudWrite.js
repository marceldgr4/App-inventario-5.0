/**
 * @fileoverview Contiene las funciones de ESCRITURA (Create, Update, Delete) para la base de datos en Google Sheets.
 */

/**
 * @summary Función genérica para añadir un nuevo registro (producto, usuario, etc.) a cualquier hoja.
 * Se encarga de generar un ID único, añadir fechas, calcular fórmulas y registrar la acción en el historial.
 * @param {object} data Objeto que contiene la información del nuevo registro (ej. { productoAgregar: 'Tornillos', ingresosAgregar: 50 }).
 * @param {string} sheetName El nombre de la hoja donde se agregará el nuevo registro.
 * @returns {string} Una cadena JSON que indica el éxito o fracaso de la operación, incluyendo el nuevo ID si fue exitosa.
 */
function agregarProductoGenerico(data, sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    return JSON.stringify({
      success: false,
      message: `Hoja "${sheetName}" no encontrada.`,
    });
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastDataRow = sheet.getLastRow();
  const firstDataRowIndex = (sheet.getFrozenRows() || 0) + 1;

  const idx = {};
  [
    "Id",
    "PRODUCTO",
    "FECHA DE INGRESO",
    "Ingresos",
    "Salidas",
    "Estado",
    "Unidades disponibles",
    "PROGRAMA",
    "Tiempo en Storage",
    "Imagen",
    "Entregas fecha",
    "Entregas cantidad",
    "TIPO",
    "PRECIO",
    "Comentarios",
    "FECHA DE ACTUALIZACION",
    "UBICACION",
    "FECHA DE VENCIMIENTO",
    "ESTADO DEL PRODUCTO",
    "COMENTARIOS",
    "NombreCompleto",
    "UserName",
    "Rol",
    "Fecha de Registro",
    "Comentario",
    "Autor",
    "Fecha",
    "Password",
    "CDE",
    "Email",
  ].forEach((col) => {
    idx[col] = headers.indexOf(col);
  });

  let newId = 1;
  if (idx["Id"] !== -1 && lastDataRow >= firstDataRowIndex) {
    const ids = sheet
      .getRange(
        firstDataRowIndex,
        idx["Id"] + 1,
        lastDataRow - firstDataRowIndex + 1,
        1
      )
      .getValues()
      .flat()
      .map(Number)
      .filter((id) => !isNaN(id));
    if (ids.length > 0) newId = Math.max(...ids) + 1;
  }

  const newRowData = Array(headers.length).fill("");
  if (idx["Id"] !== -1) newRowData[idx["Id"]] = newId;
  if (idx["PRODUCTO"] !== -1)
    newRowData[idx["PRODUCTO"]] = data.productoAgregar;
  if (idx["FECHA DE INGRESO"] !== -1)
    newRowData[idx["FECHA DE INGRESO"]] = new Date();
  if (idx["Ingresos"] !== -1)
    newRowData[idx["Ingresos"]] = parseFloat(data.ingresosAgregar) || 0;
  if (idx["Salidas"] !== -1) newRowData[idx["Salidas"]] = 0;
  if (idx["Estado"] !== -1) newRowData[idx["Estado"]] = "Activo";

  if (idx["Unidades disponibles"] !== -1) {
    let formula = "";
    const rowNum = sheet.getLastRow() + 1;
    if (sheetName === HOJA_DECORACION) {
      formula = `=IF(ISNUMBER(F${rowNum}); F${rowNum}; 0) - IF(ISNUMBER(G${rowNum});G${rowNum};0)`;
    } else if (sheetName === HOJA_COMIDA) {
      formula = `=IF(ISNUMBER(G${rowNum}); G${rowNum}; 0) - IF(ISNUMBER(H${rowNum}); H${rowNum}; 0)`;
    } else if (idx["Ingresos"] !== -1 && idx["Salidas"] !== -1) {
      const ingresosCol = columnToLetter(idx["Ingresos"] + 1);
      const salidasCol = columnToLetter(idx["Salidas"] + 1);
      formula = `=IF(ISNUMBER(${ingresosCol}${rowNum}); ${ingresosCol}${rowNum}; 0) - IF(ISNUMBER(${salidasCol}${rowNum}); ${salidasCol}${rowNum}; 0)`;
    }
    newRowData[idx["Unidades disponibles"]] =
      formula || parseFloat(data.ingresosAgregar) || 0;
  }

  // Lógica específica por hoja para poblar la nueva fila
  switch (sheetName) {
    case HOJA_ARTICULOS:
      if (idx["PROGRAMA"] !== -1)
        newRowData[idx["PROGRAMA"]] = data.programaAgregar || "";
      if (idx["Tiempo en Storage"] !== -1)
        newRowData[idx["Tiempo en Storage"]] = 0;
      if (idx["Imagen"] !== -1)
        newRowData[idx["Imagen"]] = data.imagenAgregar || "";
      if (idx["Entregas fecha"] !== -1) newRowData[idx["Entregas fecha"]] = "";
      if (idx["Entregas cantidad"] !== -1)
        newRowData[idx["Entregas cantidad"]] = 0;
      break;
    case HOJA_PAPELERIA:
      if (idx["Tiempo en Storage"] !== -1)
        newRowData[idx["Tiempo en Storage"]] = 0;
      if (idx["Imagen"] !== -1)
        newRowData[idx["Imagen"]] = data.imagenAgregar || "";
      break;
    case HOJA_DECORACION:
      if (idx["TIPO"] !== -1) newRowData[idx["TIPO"]] = data.tipoAgregar || "";
      if (idx["PRECIO"] !== -1)
        newRowData[idx["PRECIO"]] = data.precioAgregar || "";
      if (idx["Comentarios"] !== -1)
        newRowData[idx["Comentarios"]] = data.comentariosAgregar || "";
      if (idx["Imagen"] !== -1)
        newRowData[idx["Imagen"]] = data.imagenAgregar || "";
      if (idx["FECHA DE ACTUALIZACION"] !== -1)
        newRowData[idx["FECHA DE ACTUALIZACION"]] = new Date();
      break;
    case HOJA_COMIDA:
      if (idx["PRECIO"] !== -1)
        newRowData[idx["PRECIO"]] = data.precioAgregar || "";
      if (idx["UBICACION"] !== -1)
        newRowData[idx["UBICACION"]] = data.ubicacionAgregar || "";
      if (idx["FECHA DE VENCIMIENTO"] !== -1)
        newRowData[idx["FECHA DE VENCIMIENTO"]] = data.fechaVencimientoAgregar
          ? new Date(data.fechaVencimientoAgregar)
          : "";
      if (idx["ESTADO DEL PRODUCTO"] !== -1)
        newRowData[idx["ESTADO DEL PRODUCTO"]] = data.estadoAgregar || "ok";
      if (idx["COMENTARIOS"] !== -1)
        newRowData[idx["COMENTARIOS"]] = data.comentariosAgregar || "";
      if (idx["Entregas fecha"] !== -1) newRowData[idx["Entregas fecha"]] = "";
      if (idx["Entregas cantidad"] !== -1)
        newRowData[idx["Entregas cantidad"]] = 0;
      break;
    case HOJA_USUARIO:
      if (idx["NombreCompleto"] !== -1)
        newRowData[idx["NombreCompleto"]] = data.nombreCompletoAgregar;
      if (idx["UserName"] !== -1)
        newRowData[idx["UserName"]] = data.userNameAgregar;
      if (idx["Rol"] !== -1) newRowData[idx["Rol"]] = data.rolAgregar;
      if (idx["Fecha de Registro"] !== -1)
        newRowData[idx["Fecha de Registro"]] = new Date();
      break;
    case HOJA_COMENTARIOS:
      if (idx["Comentario"] !== -1)
        newRowData[idx["Comentario"]] = data.comentarioAgregar || "";
      const activeUserForComment = getActiveUser();
      if (idx["Autor"] !== -1)
        newRowData[idx["Autor"]] = activeUserForComment
          ? activeUserForComment.name || activeUserForComment.email
          : "Anónimo";
      if (idx["Fecha"] !== -1) newRowData[idx["Fecha"]] = new Date();
      break;
  }

  try {
    sheet.appendRow(newRowData);
    const appendedRowIndex = sheet.getLastRow();

    if (idx["Tiempo en Storage"] !== -1 && idx["FECHA DE INGRESO"] !== -1) {
      const fechaIngresoVal = sheet
        .getRange(appendedRowIndex, idx["FECHA DE INGRESO"] + 1)
        .getValue();
      if (fechaIngresoVal instanceof Date) {
        sheet
          .getRange(appendedRowIndex, idx["Tiempo en Storage"] + 1)
          .setValue(calcularTiempoEnStorage(fechaIngresoVal));
      }
    }

    const activeUser = getActiveUser();
    const usuario = activeUser
      ? activeUser.name || activeUser.email
      : "Sistema";

    // Preparar datos para el historial de forma modular
    let nombreItemHistorial,
      programaHistorial,
      cantidadHistorial,
      accionEstadoHistorial;

    switch (sheetName) {
      case HOJA_USUARIO:
        nombreItemHistorial =
          data.userNameAgregar || data.nombreCompletoAgregar;
        accionEstadoHistorial = `Usuario agregado: ${nombreItemHistorial}`;
        cantidadHistorial = null;
        programaHistorial = null;
        break;
      case HOJA_COMENTARIOS:
        nombreItemHistorial = `Comentario ID ${newId}`;
        accionEstadoHistorial = `Comentario agregado en ${sheetName}`;
        cantidadHistorial = null;
        programaHistorial = null;
        break;
      default:
        nombreItemHistorial = data.productoAgregar;
        programaHistorial =
          sheetName === HOJA_ARTICULOS ? data.programaAgregar : null;
        cantidadHistorial = data.ingresosAgregar
          ? parseFloat(data.ingresosAgregar)
          : null;
        accionEstadoHistorial = `Agregado: ${
          data.productoAgregar || "ítem"
        } en ${sheetName}`;
        break;
    }

    _registrarHistorialModificacion(
      newId,
      nombreItemHistorial,
      programaHistorial,
      0,
      cantidadHistorial,
      accionEstadoHistorial,
      usuario,
      new Date(),
      cantidadHistorial
    );
    return JSON.stringify({
      success: true,
      message: `Producto agregado exitosamente en ${sheetName}.`,
      newId: newId,
      imageUrl: data.imagenAgregar || "",
    });
  } catch (e) {
    Logger.log(
      `Error en agregarProductoGenerico para ${sheetName}: ${e.message}\nStack: ${e.stack}`
    );
    return JSON.stringify({
      success: false,
      message: `Error al agregar producto a ${sheetName}: ${e.message}`,
    });
  }
}

/**
 * @summary Modifica un registro existente en una hoja específica de forma OPTIMIZADA.
 * @description Lee toda la fila de datos una vez, la modifica en memoria y la escribe de vuelta en una sola operación para maximizar el rendimiento.
 * @param {object} data Objeto que contiene los datos a actualizar del formulario. Debe incluir 'idEditar'.
 * @param {string} sheetName El nombre de la hoja donde se encuentra el registro a modificar.
 * @returns {string} Una cadena JSON que indica el éxito o fracaso de la operación.
 */
function actualizarProductoGenerico(data, sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    return JSON.stringify({
      success: false,
      message: `Hoja "${sheetName}" no encontrada.`,
    });
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIndex = findRowById(data.idEditar, sheetName);

  if (rowIndex > 0) {
    try {
      const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
      const rowData = rowRange.getValues()[0];

      const activeUser = getActiveUser();
      const usuario = activeUser
        ? activeUser.name || activeUser.email
        : "Sistema";

      // Índices de columnas para fácil acceso
      const productoIdx = headers.indexOf("PRODUCTO");
      const programaIdx = headers.indexOf("PROGRAMA");
      const unidadesDispIdx = headers.indexOf("Unidades disponibles");
      const ingresosIdx = headers.indexOf("Ingresos");
      const salidasIdx = headers.indexOf("Salidas");
      const fechaIngresoIdx = headers.indexOf("FECHA DE INGRESO");

      const productoActualOriginal = rowData[productoIdx];
      const programaActualOriginal = rowData[programaIdx] || "";
      const unidadesOriginalesEnFila =
        parseFloat(rowData[unidadesDispIdx]) || 0;

      // Lógica para actualizar los ingresos
      const cantidadAgregadaDesdeFormulario =
        parseFloat(data.ingresosEditar) || 0;
      if (cantidadAgregadaDesdeFormulario > 0) {
        if (unidadesOriginalesEnFila === 0) {
          // Si el producto estaba agotado, se reinicia
          rowData[ingresosIdx] = cantidadAgregadaDesdeFormulario;
          rowData[salidasIdx] = 0;
          rowData[fechaIngresoIdx] = new Date();
        } else {
          // Si ya había unidades, se suman
          const ingresosActuales = parseFloat(rowData[ingresosIdx]) || 0;
          rowData[ingresosIdx] =
            ingresosActuales + cantidadAgregadaDesdeFormulario;
        }
      }

      if (data.productoEditar !== undefined)
        rowData[productoIdx] = data.productoEditar;

      // Lógica específica por hoja para actualizar campos
      switch (sheetName) {
        case HOJA_ARTICULOS:
          if (data.programaEditar !== undefined)
            rowData[programaIdx] = data.programaEditar;
          if (data.imagenEditar !== undefined)
            rowData[headers.indexOf("Imagen")] = data.imagenEditar;
          break;
        case HOJA_PAPELERIA:
          if (data.imagenEditar !== undefined)
            rowData[headers.indexOf("Imagen")] = data.imagenEditar;
          break;
        case HOJA_DECORACION:
          if (data.tipoEditar !== undefined)
            rowData[headers.indexOf("TIPO")] = data.tipoEditar;
          if (data.precioEditar !== undefined)
            rowData[headers.indexOf("PRECIO")] = data.precioEditar;
          if (data.comentariosEditar !== undefined)
            rowData[headers.indexOf("Comentarios")] = data.comentariosEditar;
          if (data.fechaActualizacionEditar) {
            rowData[headers.indexOf("FECHA DE ACTUALIZACION")] = new Date(
              data.fechaActualizacionEditar
            );
          }
          break;
        case HOJA_COMIDA:
          if (data.precioEditar !== undefined)
            rowData[headers.indexOf("PRECIO")] = data.precioEditar;
          if (data.ubicacionEditar !== undefined)
            rowData[headers.indexOf("UBICACION")] = data.ubicacionEditar;
          if (data.estadoEditar !== undefined)
            rowData[headers.indexOf("ESTADO DEL PRODUCTO")] = data.estadoEditar;
          if (data.fechaVencimientoEditar !== undefined)
            rowData[headers.indexOf("FECHA DE VENCIMIENTO")] = new Date(
              data.fechaVencimientoEditar
            );
          if (data.comentariosEditar !== undefined)
            rowData[headers.indexOf("COMENTARIOS")] = data.comentariosEditar;
          break;
        case HOJA_USUARIO:
          if (data.nombreCompletoEditar !== undefined)
            rowData[headers.indexOf("NombreCompleto")] =
              data.nombreCompletoEditar;
          if (data.userNameEditar !== undefined)
            rowData[headers.indexOf("UserName")] = data.userNameEditar;
          if (data.passwordEditar)
            rowData[headers.indexOf("Password")] = data.passwordEditar; // Solo si se provee una nueva
          if (data.cdeEditar !== undefined)
            rowData[headers.indexOf("CDE")] = data.cdeEditar;
          if (data.emailEditar !== undefined)
            rowData[headers.indexOf("Email")] = data.emailEditar;
          if (data.rolEditar !== undefined)
            rowData[headers.indexOf("Rol")] = data.rolEditar;
          break;
      }

      // Actualizar 'Tiempo en Storage' basado en la fecha de ingreso que ahora está en 'rowData'
      const tiempoStorageIdx = headers.indexOf("Tiempo en Storage");
      const fechaIngresoVal = rowData[fechaIngresoIdx];
      if (fechaIngresoVal instanceof Date) {
        rowData[tiempoStorageIdx] = calcularTiempoEnStorage(fechaIngresoVal);
      }

      let formula = "";
      if (sheetName === HOJA_DECORACION) {
        formula = `=IF(ISNUMBER(E${rowIndex}); E${rowIndex}; 0) - IF(ISNUMBER(F${rowIndex}); F${rowIndex};0)`;
      } else if (sheetName === HOJA_COMIDA) {
        formula = `=IF(ISNUMBER(G${rowIndex}); G${rowIndex}; 0) - IF(ISNUMBER(H${rowIndex}); H${rowIndex}; 0)`;
      } else if (ingresosIdx !== -1 && salidasIdx !== -1) {
        const ingresosColLetter = columnToLetter(ingresosIdx + 1);
        const salidasColLetter = columnToLetter(salidasIdx + 1);
        formula = `=IF(ISNUMBER(${ingresosColLetter}${rowIndex}); ${ingresosColLetter}${rowIndex}; 0) - IF(ISNUMBER(${salidasColLetter}${rowIndex});${salidasColLetter}${rowIndex}; 0)`;
      }

      if (formula) {
        rowData[unidadesDispIdx] = formula;
      }

      rowRange.setValues([rowData]);
      SpreadsheetApp.flush(); // Forzamos la actualización para que las fórmulas se calculen.

      // Leemos el valor final de las unidades DESPUÉS de que la fórmula se ha calculado.
      const unidadesNuevasCalculadas =
        parseFloat(sheet.getRange(rowIndex, unidadesDispIdx + 1).getValue()) ||
        0;

      _registrarHistorialModificacion(
        data.idEditar,
        data.productoEditar || productoActualOriginal,
        sheetName === HOJA_ARTICULOS && data.programaEditar !== undefined
          ? data.programaEditar
          : programaActualOriginal,
        unidadesOriginalesEnFila,
        unidadesNuevasCalculadas,
        `Modificado en ${sheetName} (Agregado por formulario: ${cantidadAgregadaDesdeFormulario})`,
        usuario,
        null,
        null
      );

      return JSON.stringify({
        success: true,
        message: `Producto actualizado exitosamente en ${sheetName}.`,
      });
    } catch (e) {
      Logger.log(
        `Error al actualizar producto ID ${
          data.idEditar
        } en ${sheetName}: ${e.toString()}`
      );
      return JSON.stringify({
        success: false,
        message: `Error del servidor: ${e.message}`,
      });
    }
  }

  return JSON.stringify({
    success: false,
    message: `No se encontró el producto con el ID proporcionado en ${sheetName}.`,
  });
}

/**
 * @summary Realiza una "eliminación lógica" de un producto. Cambia el estado del producto a "Desactivado" en lugar de borrar la fila.
 * @param {string|number} idToDelete El ID del registro que se va a desactivar.
 * @param {string} sheetName El nombre de la hoja donde se encuentra el registro.
 * @returns {string} Un mensaje de éxito indicando que el producto fue desactivado.
 * @throws {Error} Lanza un error si no se encuentra el producto a eliminar o si ocurre un problema durante el proceso.
 */
function eliminarProductoGenerico(idToDelete, sheetName) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) {
      return JSON.stringify({
        success: false,
        message: `Hoja "${sheetName}" no encontrada.`,
      });
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const rowIndex = findRowById(idToDelete, sheetName);

    const activeUser = getActiveUser();
    const usuario = activeUser
      ? activeUser.email || activeUser.name || "Usuario Desconocido"
      : "Sistema";

    if (rowIndex > 0) {
      const productoCell = sheet.getRange(
        rowIndex,
        headers.indexOf("PRODUCTO") + 1
      );
      const producto = productoCell
        ? productoCell.getValue().toString()
        : "Producto Desconocido";

      const estadoIdx = headers.indexOf("Estado");
      if (estadoIdx !== -1) {
        sheet.getRange(rowIndex, estadoIdx + 1).setValue("Desactivado");
      } else {
        Logger.log(
          `Advertencia: Columna 'Estado' no encontrada en ${sheetName} al intentar desactivar producto ID ${idToDelete}.`
        );
      }

      const programaIdx = headers.indexOf("PROGRAMA");
      const programa =
        programaIdx !== -1
          ? (
              sheet.getRange(rowIndex, programaIdx + 1).getValue() || ""
            ).toString()
          : "";

      let unidadesAnteriores = 0;
      // ... (lógica para obtener unidadesAnteriores) ...
      const unidadesDispIdx = headers.indexOf("Unidades disponibles");
      if (unidadesDispIdx !== -1) {
        const unidadesDispCell = sheet.getRange(rowIndex, unidadesDispIdx + 1);
        unidadesAnteriores =
          parseFloat(unidadesDispCell ? unidadesDispCell.getValue() : 0) || 0;
      } else {
        const ingresosIdx = headers.indexOf("Ingresos");
        const salidasIdx = headers.indexOf("Salidas");
        const ingresos =
          ingresosIdx !== -1
            ? parseFloat(
                sheet.getRange(rowIndex, ingresosIdx + 1).getValue()
              ) || 0
            : 0;
        const salidas =
          salidasIdx !== -1
            ? parseFloat(sheet.getRange(rowIndex, salidasIdx + 1).getValue()) ||
              0
            : 0;
        unidadesAnteriores = ingresos - salidas;
      }

      _registrarHistorialModificacion(
        idToDelete,
        producto,
        programa,
        unidadesAnteriores,
        unidadesAnteriores, // Las unidades "nuevas" son las mismas si solo se desactiva
        "Eliminado (Desactivado) de " + sheetName,
        usuario,
        null,
        null
      );

      let successMessageText;
      if (sheetName === HOJA_ARTICULOS) {
        successMessageText = `Producto ID ${idToDelete} (${producto}) eliminado exitosamente de Articulos.`;
      } else {
        successMessageText = `Producto ID ${idToDelete} (${producto}) eliminado (desactivado) exitosamente de ${sheetName}.`;
      }

      return JSON.stringify({ success: true, message: successMessageText });
    }

    throw new Error(
      `No se encontró el producto con el ID ${idToDelete} en ${sheetName} para eliminar.`
    );
  } catch (error) {
    const errorMessage = `Error en eliminarProductoGenerico para ${sheetName}, ID ${idToDelete}: ${error.toString()}`;
    Logger.log(`${errorMessage}\nStack: ${error.stack}`);
    throw new Error(
      `Error al procesar la eliminación del producto ID ${idToDelete} en ${sheetName}: ${error.message}`
    );
  }
}

/**
 * @summary Gestiona la salida (retiro) de unidades de un producto del inventario.
 * @param {string|number} idProducto El ID del producto del cual se retirarán unidades.
 * @param {number} unidadesRetirar La cantidad de unidades a retirar.
 * @param {string} sheetName El nombre de la hoja donde se encuentra el producto.
 * @returns {string} Una cadena JSON con el resultado de la operación (éxito o fracaso con mensaje).
 */
function retirarProductoGenerico(idProducto, unidadesRetirar, sheetName) {
  let effectiveId = idProducto;
  if (typeof idProducto === "object" && idProducto !== null) {
    if (idProducto.id) effectiveId = idProducto.id;
    else if (idProducto.ID) effectiveId = idProducto.ID;
    else if (idProducto.Id) effectiveId = idProducto.Id;
    else if (Object.keys(idProducto).length > 0)
      effectiveId = idProducto[Object.keys(idProducto)[0]];
  }

  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) {
      return JSON.stringify({
        success: false,
        message: `Error: Hoja "${sheetName}" no encontrada.`,
      });
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const rowIndex = findRowById(effectiveId, sheetName);

    const activeUser = getActiveUser();
    const usuario = activeUser
      ? activeUser.email || activeUser.name || "Usuario Desconocido"
      : "Sistema";

    if (rowIndex > 0) {
      const salidasIndex = headers.indexOf("Salidas");
      const ingresosIndex = headers.indexOf("Ingresos");
      const productoIndex = headers.indexOf("PRODUCTO");
      const programaIndex = headers.indexOf("PROGRAMA"); // Puede no existir en todas las hojas
      const unidadesDispIdx = headers.indexOf("Unidades disponibles");

      if (productoIndex === -1) {
        return JSON.stringify({
          success: false,
          message: "Error: Columna 'PRODUCTO' no encontrada.",
        });
      }
      if (
        ingresosIndex === -1 ||
        salidasIndex === -1 ||
        unidadesDispIdx === -1
      ) {
        return JSON.stringify({
          success: false,
          message:
            "Error: Columnas 'Ingresos', 'Salidas', o 'Unidades disponibles' no encontradas.",
        });
      }

      const productoCell = sheet.getRange(rowIndex, productoIndex + 1);
      const producto = productoCell
        ? productoCell.getValue().toString()
        : "Producto Desconocido";

      const unidadesDispCell = sheet.getRange(rowIndex, unidadesDispIdx + 1);
      const unidadesDisponiblesAntes =
        parseFloat(unidadesDispCell ? unidadesDispCell.getValue() : 0) || 0;

      const numUnidadesRetirar = parseFloat(unidadesRetirar);
      if (isNaN(numUnidadesRetirar) || numUnidadesRetirar <= 0) {
        return JSON.stringify({
          success: false,
          message: "La cantidad a retirar debe ser un número mayor que cero.",
        });
      }
      if (numUnidadesRetirar > unidadesDisponiblesAntes) {
        return JSON.stringify({
          success: false,
          message: `No se pueden retirar ${numUnidadesRetirar} unidad(es). Solo hay ${unidadesDisponiblesAntes} disponibles.`,
        });
      }

      const salidasActualesCell = sheet.getRange(rowIndex, salidasIndex + 1);
      const salidasActuales =
        parseFloat(salidasActualesCell ? salidasActualesCell.getValue() : 0) ||
        0;
      const nuevasSalidas = salidasActuales + numUnidadesRetirar;
      sheet.getRange(rowIndex, salidasIndex + 1).setValue(nuevasSalidas);

      const entregasFechaIdx = headers.indexOf("Entregas fecha");
      const entregasCantidadIdx = headers.indexOf("Entregas cantidad");
      if (entregasFechaIdx !== -1) {
        sheet.getRange(rowIndex, entregasFechaIdx + 1).setValue(new Date());
      }
      if (entregasCantidadIdx !== -1) {
        const entregasCantidadCell = sheet.getRange(
          rowIndex,
          entregasCantidadIdx + 1
        );
        const entregasCantidadActual =
          parseFloat(
            entregasCantidadCell ? entregasCantidadCell.getValue() : 0
          ) || 0;
        sheet
          .getRange(rowIndex, entregasCantidadIdx + 1)
          .setValue(entregasCantidadActual + numUnidadesRetirar);
      }

      SpreadsheetApp.flush(); // Importante si 'Unidades disponibles' es una fórmula

      const unidadesDispActualizadasCell = sheet.getRange(
        rowIndex,
        unidadesDispIdx + 1
      );
      const unidadesDisponiblesDespues =
        parseFloat(
          unidadesDispActualizadasCell
            ? unidadesDispActualizadasCell.getValue()
            : 0
        ) || 0;

      const programa =
        programaIndex !== -1
          ? sheet.getRange(rowIndex, programaIndex + 1).getValue() || ""
          : "";

      _registrarHistorialModificacion(
        effectiveId,
        producto,
        programa,
        unidadesDisponiblesAntes,
        unidadesDisponiblesDespues,
        `Retiro de ${numUnidadesRetirar} unidades desde ${sheetName}`,
        usuario,
        new Date(),
        numUnidadesRetirar
      );

      return JSON.stringify({
        success: true,
        message: `${numUnidadesRetirar} unidad(es) del producto "${producto}" (ID: ${effectiveId}) se retiraron. Quedan ${unidadesDisponiblesDespues}.`,
      });
    }

    return JSON.stringify({
      success: false,
      message: `Error: Producto con ID ${effectiveId} no encontrado.`,
    });
  } catch (error) {
    Logger.log(
      `Error en retirarProductoGenerico para ${sheetName}, ID ${effectiveId}, Unidades ${unidadesRetirar}: ${error.toString()}\nStack: ${
        error.stack
      }`
    );
    return JSON.stringify({
      success: false,
      message: `Error al procesar el retiro: ${error.message}`,
    });
  }
}

/**
 * @summary Orquesta la adición de un producto que incluye una imagen. Sube el archivo a Drive y luego añade el registro a la hoja.
 * @param {object} productData Los datos del producto (nombre, cantidad, etc.).
 * @param {object} fileData La información del archivo de imagen, incluyendo `base64Data`, `mimeType` y `fileName`.
 * @returns {string} El resultado JSON de la función `agregarProductoGenerico`, indicando éxito o fracaso.
 */
function agregarProductoConImagenDesdeArchivo(productData, fileData) {
  try {
    const folderId = ID_DRIVE_IMG;
    const folder = DriveApp.getFolderById(folderId);

    const bytes = Utilities.base64Decode(fileData.base64Data);
    const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.fileName);

    const file = folder.createFile(blob);
    productData.imagenAgregar = file.getUrl();

    const result = agregarProductoGenerico(productData, HOJA_ARTICULOS);

    return result; // Devuelve el JSON de éxito/error de agregarProductoGenerico
  } catch (e) {
    Logger.log(
      "Error en agregarProductoConImagenDesdeArchivo: " +
        e.message +
        " Stack: " +
        e.stack
    );
    return JSON.stringify({
      success: false,
      message: "Error al subir la imagen y agregar el producto: " + e.message,
    });
  }
}

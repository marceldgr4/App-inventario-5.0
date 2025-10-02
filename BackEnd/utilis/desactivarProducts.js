function deactivateProducts(ids, sheetName) {
  // 🚨 CAMBIO CLAVE: Usar el nombre de la hoja pasado por el cliente (sheetName)
  // Si por alguna razón sheetName no viene, usamos 'Inventario' como respaldo.
  const hojaAUsar = sheetName;
  const ss = SpreadsheetApp.openById(
    "1lWsyJLZTOZDbIeAcagvj7wCPIp9vD_METhdkiHQBzEU"
  );

  // Usamos hojaAUsar para obtener la hoja.
  const sheet = ss.getSheetByName(hojaAUsar);

  // 🚨 Nueva validación: Si la hoja no existe, devuelve un error claro.
  if (!sheet) {
    return JSON.stringify({
      success: false,
      message: `Hoja de cálculo '${hojaAUsar}' no encontrada. Revise el nombre.`,
    });
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Se usa 'Estado' y 'ID' como están definidos en su código.
  const estadoIndex = headers.indexOf("Estado");
  const idIndex = headers.indexOf("ID");

  if (estadoIndex === -1 || idIndex === -1) {
    return JSON.stringify({
      success: false,
      message: "Encabezado 'Estado' o 'ID' no encontrado en la hoja.",
    });
  }

  const updatedRows = [];
  ids.forEach((idToDeactivate) => {
    let rowFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] == idToDeactivate) {
        const row = data[i];
        // CAMBIO SOLICITADO: Cambia el estado a 'Desactivado'
        row[estadoIndex] = "Desactivado";
        updatedRows.push({ rowNumber: i + 1, rowData: row });
        rowFound = true;
        break;
      }
    }
    if (!rowFound) {
      console.warn(`Producto con ID ${idToDeactivate} no encontrado.`);
    }
  });

  if (updatedRows.length > 0) {
    // Aplica los cambios a toda la hoja
    const range = sheet.getRange(2, 1, data.length - 1, headers.length);
    range.setValues(data.slice(1));
    return JSON.stringify({
      success: true,
      message: `Se han desactivado ${updatedRows.length} producto(s) exitosamente de la hoja ${hojaAUsar}.`,
    });
  } else {
    return JSON.stringify({
      success: false,
      message: "No se encontraron productos para dar de baja.",
    });
  }
}

function desactivarNotificacionesComida(ids) {
  const sheetName = getHojasConfig().COMIDA.nombre;
  const sheet = getSheet(sheetName);
  if (!sheet) {
    return JSON.stringify({
      success: false,
      message: "No se pudo encontrar la hoja de cálculo.",
    });
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const estadoIndex = headers.indexOf("Estado");
  const idIndex = headers.indexOf("Id");
  const productoIndex = headers.indexOf("PRODUCTO");
  const unidadesIndex = headers.indexOf("Unidades disponibles");
  const programaIndex = headers.indexOf("PROGRAMA");

  if (estadoIndex === -1 || idIndex === -1 || productoIndex === -1) {
    return JSON.stringify({
      success: false,
      message: "Encabezado 'Estado', 'Id' o 'PRODUCTO' no encontrado.",
    });
  }

  const activeUser = getActiveUser();
  const usuario = activeUser
    ? activeUser.email || activeUser.name || "Usuario Desconocido"
    : "Sistema";

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = String(row[idIndex]);
    if (ids.includes(rowId)) {
      // Don't deactivate if already deactivated
      if (row[estadoIndex] === "Desactivado") {
        continue;
      }

      row[estadoIndex] = "Desactivado";
      updated++;

      const producto = row[productoIndex] || "Producto Desconocido";
      const programa = programaIndex !== -1 ? row[programaIndex] || "" : "";
      const unidadesAnteriores =
        unidadesIndex !== -1 ? parseFloat(row[unidadesIndex]) || 0 : 0;

      _registrarHistorialModificacion(
        rowId,
        producto,
        programa,
        unidadesAnteriores,
        unidadesAnteriores,
        "Eliminado (Desactivado) de " + sheetName,
        usuario,
        null,
        null
      );
    }
  }

  if (updated > 0) {
    sheet
      .getRange(2, 1, data.length - 1, headers.length)
      .setValues(data.slice(1));
    return JSON.stringify({
      success: true,
      message: `${updated} producto(s) desactivados.`,
    });
  } else {
    return JSON.stringify({
      success: false,
      message:
        "No se encontraron productos a desactivar o ya estaban desactivados.",
    });
  }
}

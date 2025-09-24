function deactivateProducts(ids) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const estadoIndex = headers.indexOf("Estado");
  const idIndex = headers.indexOf("ID");
  
  if (estadoIndex === -1 || idIndex === -1) {
    return JSON.stringify({ success: false, message: "Encabezado 'Estado' o 'ID' no encontrado." });
  }

  const updatedRows = [];
  ids.forEach(idToDeactivate => {
    let rowFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] == idToDeactivate) {
        const row = data[i];
        row[estadoIndex] = "inactivo"; // Cambia el estado a "inactivo"
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
    const range = sheet.getRange(2, 1, data.length - 1, headers.length);
    range.setValues(data.slice(1));
    return JSON.stringify({ success: true, message: `Se han desactivado ${updatedRows.length} producto(s) exitosamente.` });
  } else {
    return JSON.stringify({ success: false, message: "No se encontraron productos para dar de baja." });
  }
}
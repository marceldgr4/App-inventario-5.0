// BackEnd/Inventario.js

function getInventarioData() {
  Logger.log("getInventarioData: Iniciando la obtención de datos.");
  try {
    const data = _getInventoryDataForSheet(getHojasConfig().ARTICULO.nombre);
    Logger.log("getInventarioData: Registros obtenidos: " + data.length);
    return JSON.stringify({ data: data }); // 👈 devolver siempre string JSON
  } catch (e) {
    Logger.log("getInventarioData: Error al obtener datos. " + e.message);
    return JSON.stringify({ data: [], error: e.message });
  }
}

function getInventarioInfo(id) {
  try {
    return getProductInfoGenerico(id, getHojasConfig().ARTICULO.nombre);
  } catch (e) {
    Logger.log("getInventarioInfo: Error al obtener info. " + e.message);
    return null;
  }
}

function agregarProductoInventario(data) {
  try {
    const result = agregar(data, getHojasConfig().ARTICULO.nombre);
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("agregarProductoInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

function actualizarProductoInventario(data) {
  try {
    const result = actualizar(data, getHojasConfig().ARTICULO.nombre);
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("actualizarProductoInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

function eliminarProductoInventario(id) {
  try {
    const result = eliminar(id, getHojasConfig().ARTICULO.nombre);
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("eliminarProductoInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

function retirarProductoInventario(id, unidades) {
  try {
    const sheetName = getHojasConfig().ARTICULO.nombre;
    const result = retirar(id, unidades, sheetName); // tu lógica existente

    // --- Actualizar columna Entregas fecha ---
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data && data.length > 0) {
      const headers = data[0].map((h) =>
        String(h || "")
          .trim()
          .toLowerCase()
          .replace(/\s/g, "")
      );
      const idIdx = headers.findIndex(
        (h) => h === "id" || h === "identificador"
      );
      // buscar nombres variantes:
      const entregasIdx = headers.findIndex(
        (h) => h.replace(/\s/g, "").toLowerCase() === "entregasfecha"
      );

      if (idIdx !== -1 && entregasIdx !== -1) {
        const rowIndex = data.findIndex(
          (r, idx) => idx > 0 && String(r[idIdx]) === String(id)
        );
        if (rowIndex !== -1) {
          sheet.getRange(rowIndex + 1, entregasIdx + 1).setValue(new Date());
        }
      }
    }

    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("retirarProductoInventario error: " + e);
    return JSON.stringify({ success: false, error: e.message });
  }
}

function agregarComentarioInventario(id, comentario) {
  try {
    const result = agregarComentarioGenerico(
      getHojasConfig().ARTICULO.nombre,
      id,
      comentario
    );
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("agregarComentarioInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

// Archivo: Inventario.js
function getArticulosViejos() {
  try {
    const productosViejos = getProductosViejos(getHojasConfig().ARTICULO.nombre);
    Logger.log(
      `getArticulosViejos: Devolviendo ${productosViejos.length} artículos viejos.`
    );
    return JSON.stringify(productosViejos);
  } catch (e) {
    Logger.log("getArticulosViejos: Error al obtener datos. " + e.message);
    return JSON.stringify([]);
  }
}

function deactivateProducts(ids) {
  try {
    if (!ids || !Array.isArray(ids) || ids.length === 0)
      return JSON.stringify({ success: false, message: "No hay ids" });

    const sheetName = getHojasConfig().ARTICULO.nombre;
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)
      return JSON.stringify({
        success: false,
        message: "Hoja no encontrada: " + sheetName,
      });

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map((h) =>
      String(h || "")
        .trim()
        .toLowerCase()
    );
    const idIdx = headers.findIndex(
      (h) => h === "id" || h.indexOf("id") !== -1
    );
    const estadoIdx = headers.findIndex((h) => h === "estado");

    if (idIdx === -1)
      return JSON.stringify({
        success: false,
        message: "Columna Id no encontrada.",
      });
    if (estadoIdx === -1)
      return JSON.stringify({
        success: false,
        message: "Columna Estado no encontrada.",
      });

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = String(row[idIdx]);
      if (ids.map(String).indexOf(rowId) !== -1) {
        sheet.getRange(i + 1, estadoIdx + 1).setValue("eliminado");
        updated++;
      }
    }

    return JSON.stringify({
      success: true,
      message: `${updated} producto(s) dados de baja.`,
    });
  } catch (e) {
    Logger.log("deactivateProducts error: " + e);
    return JSON.stringify({ success: false, message: e.message || String(e) });
  }
}

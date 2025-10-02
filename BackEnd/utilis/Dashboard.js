// BackEnd/Dashboard.js

/**
 * @summary Recopila y procesa datos de múltiples hojas para construir el objeto de datos del dashboard.
 * @returns {object} Un objeto que contiene los datos procesados para cada sección del dashboard (inventario, comida, etc.).
 */
function obtenerDatosDashboard() {
  const ss = SpreadsheetApp.openById(ID_INVENTARIO);
  if (!ss) {
    Logger.log(
      "Error: No se pudo abrir la hoja de cálculo con el ID: " + ID_INVENTARIO
    );
    return {
      error:
        "No se pudo acceder a la hoja de cálculo. Revise los registros de secuencia de comandos para obtener más detalles.",
    };
  }
  return {
    inventario: _getInventarioDataForDashboard(
      ss.getSheetByName(HOJA_ARTICULOS)
    ),
    comida: _getComidaDataForDashboardWithMonthlySpending(
      ss.getSheetByName(HOJA_COMIDA)
    ),
    decoracion: _getDecoracionDataForDashboardWithMonthlySpending(
      ss.getSheetByName(HOJA_DECORACION)
    ),
    papeleria: _getPapeleriaDataForDashboard(ss.getSheetByName(HOJA_PAPELERIA)),
  };
}

/**
 * @summary Procesa los datos de la hoja 'Articulos' para obtener las métricas del dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet El objeto de la hoja 'Articulos'.
 * @returns {object} Un objeto con datos agregados sobre el inventario.
 * @private
 */
function _getInventarioDataForDashboard(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxProducto = headers.indexOf("PRODUCTO");
  const idxUnidades = headers.indexOf("Unidades disponibles");
  const idxPrograma = headers.indexOf("PROGRAMA");
  const idxTiempoStorage = headers.indexOf("Tiempo en Storage");

  if (
    idxProducto === -1 ||
    idxUnidades === -1 ||
    idxPrograma === -1 ||
    idxTiempoStorage === -1
  )
    return {};

  const inventarioData = {
    totalProductos: 0,
    unidadesAgotadas: 0,
    unidadesDisponibles: 0,
    programas: {},
    productosPorPrograma: [],
    masDeOchoMeses: 0,
  };

  const todosLosProgramas = new Set();

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][idxProducto];
    const unidades = parseInt(data[i][idxUnidades]) || 0;
    const programa = data[i][idxPrograma];
    const tiempoStorage = parseFloat(data[i][idxTiempoStorage]) || 0;

    if (producto && programa) {
      todosLosProgramas.add(programa);
      inventarioData.totalProductos++;

      if (unidades > 0) {
        inventarioData.unidadesDisponibles++;
        if (tiempoStorage >= 8) {
          inventarioData.masDeOchoMeses++;
        }
        inventarioData.programas[programa] =
          (inventarioData.programas[programa] || 0) + 1;
      } else {
        inventarioData.unidadesAgotadas++;
        if (!(programa in inventarioData.programas)) {
          inventarioData.programas[programa] = 0;
        }
      }

      inventarioData.productosPorPrograma.push({
        producto,
        unidades,
        programa,
      });
    }
  }

  todosLosProgramas.forEach((programa) => {
    if (!(programa in inventarioData.programas)) {
      inventarioData.programas[programa] = 0;
    }
  });

  inventarioData.programList = Array.from(todosLosProgramas);
  return inventarioData;
}

/**
 * @summary Procesa los datos de la hoja 'Comida' para obtener las métricas del dashboard, incluyendo el gasto mensual.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet El objeto de la hoja 'Comida'.
 * @returns {object} Un objeto con datos agregados sobre la comida.
 * @private
 */
function _getComidaDataForDashboardWithMonthlySpending(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxProducto = headers.indexOf("PRODUCTO");
  const idxPrecio = headers.indexOf("PRECIO");
  const idxUnidades = headers.indexOf("Unidades disponibles");
  const idxFechaIngreso = headers.indexOf("FECHA DE INGRESO");

  if (
    idxProducto === -1 ||
    idxPrecio === -1 ||
    idxUnidades === -1 ||
    idxFechaIngreso === -1
  )
    return {};

  const comidaData = {
    totalProductos: 0,
    unidadesDisponibles: 0.0,
    unidadesAgotadas: 0,
    productos: [],
    gastoMensual: {},
  };

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][idxProducto];
    const precio = parseFloat(data[i][idxPrecio]) || 0;
    const unidades = parseInt(data[i][idxUnidades]) || 0;
    const fechaIngreso = data[i][idxFechaIngreso];

    if (producto) {
      comidaData.totalProductos++;
      if (unidades > 0.0) comidaData.unidadesDisponibles++;
      else comidaData.unidadesAgotadas++;
      comidaData.productos.push({ producto, unidades });

      if (fechaIngreso instanceof Date) {
        const ano = fechaIngreso.getFullYear();
        const mes = fechaIngreso.getMonth() + 1;
        const anoMes = `${ano}-${mes < 10 ? "0" + mes : mes}`;
        comidaData.gastoMensual[anoMes] =
          (comidaData.gastoMensual[anoMes] || 0) + precio;
      }
    }
  }
  return comidaData;
}

/**
 * @summary Procesa los datos de la hoja 'Decoracion' para obtener las métricas del dashboard, incluyendo el gasto mensual.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet El objeto de la hoja 'Decoracion'.
 * @returns {object} Un objeto con datos agregados sobre la decoración.
 * @private
 */
function _getDecoracionDataForDashboardWithMonthlySpending(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxProducto = headers.indexOf("PRODUCTO");
  const idxPrecio = headers.indexOf("PRECIO");
  const idxUnidades = headers.indexOf("Unidades disponibles");
  const idxCategoria = headers.indexOf("TIPO");
  const idxFechaIngreso = headers.indexOf("FECHA DE INGRESO");

  if (
    idxProducto === -1 ||
    idxPrecio === -1 ||
    idxUnidades === -1 ||
    idxCategoria === -1 ||
    idxFechaIngreso === -1
  )
    return {};

  const decoracionData = {
    totalProductos: 0,
    unidadesDisponibles: 0,
    unidadesAgotadas: 0,
    categorias: {},
    productos: [],
    gastoMensual: {},
  };

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][idxProducto];
    const precio = parseFloat(data[i][idxPrecio]) || 0;
    const unidades = parseInt(data[i][idxUnidades]) || 0;
    const categoria = data[i][idxCategoria];
    const fechaIngreso = data[i][idxFechaIngreso];

    if (producto) {
      decoracionData.totalProductos++;
      if (unidades > 0) {
        decoracionData.unidadesDisponibles++;
      } else {
        decoracionData.unidadesAgotadas++;
      }
      decoracionData.categorias[categoria] =
        (decoracionData.categorias[categoria] || 0) + unidades;
      decoracionData.productos.push({ producto, unidades });

      if (fechaIngreso instanceof Date) {
        const ano = fechaIngreso.getFullYear();
        const mes = fechaIngreso.getMonth() + 1;
        const anoMes = `${ano}-${mes < 10 ? "0" + mes : mes}`;
        decoracionData.gastoMensual[anoMes] =
          (decoracionData.gastoMensual[anoMes] || 0) + precio;
      }
    }
  }

  decoracionData.categoriaList = Object.keys(decoracionData.categorias);
  return decoracionData;
}

/**
 * @summary Procesa los datos de la hoja 'Papeleria' para obtener las métricas del dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet El objeto de la hoja 'Papeleria'.
 * @returns {object} Un objeto con datos agregados sobre la papelería.
 * @private
 */
function _getPapeleriaDataForDashboard(sheet) {
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxProducto = headers.indexOf("PRODUCTO");
  const idxUnidades = headers.indexOf("Unidades disponibles");

  if (idxProducto === -1 || idxUnidades === -1) return {};

  const papeleriaData = {
    totalProductos: 0,
    unidadesDisponibles: 0,
    unidadesAgotadas: 0,
    productos: [],
  };

  for (let i = 1; i < data.length; i++) {
    const producto = data[i][idxProducto];
    const unidades = parseInt(data[i][idxUnidades]) || 0;

    if (producto) {
      papeleriaData.totalProductos++;

      if (unidades > 0) {
        papeleriaData.unidadesDisponibles++;
      } else {
        papeleriaData.unidadesAgotadas++;
      }

      papeleriaData.productos.push({ producto, unidades });
    }
  }

  return papeleriaData;
}

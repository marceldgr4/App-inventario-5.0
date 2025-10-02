function getProductosProximosAVencer(sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const productosProximos = [];
  const fechaActual = new Date(); // 🔑 Columnas Ajustadas a su lista

  const idIndex = headers.indexOf("Id"); // Antes buscaba 'Id', es correcto
  const productoIndex = headers.indexOf("PRODUCTO");
  const fechaVencimientoIndex = headers.indexOf("FECHA DE VENCIMIENTO");
  const unidadesIndex = headers.indexOf("Unidades disponibles"); // ✅ AJUSTADO // ✅ Estado flexible

  let estadoIndex = headers.indexOf("Estado del Producto"); // ✅ AJUSTADO (Usa "Estado del Producto")
  if (estadoIndex === -1) {
    estadoIndex = headers.indexOf("Estado");
  }

  if (
    fechaVencimientoIndex === -1 ||
    productoIndex === -1 ||
    unidadesIndex === -1
  ) {
    Logger.log(
      `getProductosProximosAVencer: Columnas requeridas no encontradas en ${sheetName}`
    );
    return [];
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = idIndex !== -1 ? row[idIndex] : i;
    const producto = row[productoIndex];
    const fechaVencimientoValue = row[fechaVencimientoIndex];
    const estadoProducto =
      estadoIndex !== -1 ? row[estadoIndex] || "Activo" : "Activo";

    if (String(estadoProducto).toLowerCase() !== "activo") {
      continue;
    } // Lógica de cálculo de días (omitido por brevedad)

    if (fechaVencimientoValue && fechaVencimientoValue instanceof Date) {
      const fechaVencimiento = fechaVencimientoValue;
      const utcHoy = Date.UTC(
        fechaActual.getFullYear(),
        fechaActual.getMonth(),
        fechaActual.getDate()
      );
      const utcVencimiento = Date.UTC(
        fechaVencimiento.getFullYear(),
        fechaVencimiento.getMonth(),
        fechaVencimiento.getDate()
      );
      diasRestantes = Math.ceil(
        (utcVencimiento - utcHoy) / (1000 * 60 * 60 * 24)
      ); // 🚨 FILTRO DE 30 DÍAS: Si quedan más de 30 días, saltar este producto
      if (diasRestantes > 30) {
        continue;
      }

      productosProximos.push({
        id: id || "N/A",
        producto: producto,
        diasRestantes: diasRestantes,
        estado: estadoProducto,
        unidadesDisponibles: unidadesIndex !== -1 ? row[unidadesIndex] : 0, // ✅ CLAVE CORREGIDA
      });
    }
  }

  return productosProximos;
}

/**
 * @summary Obtiene toda la información de una fila (un producto) basado en su ID y la devuelve como un objeto.
 * @description Utiliza `findRowById` para ser más eficiente que leer toda la hoja.
 * @param {string|number} id El ID del producto a consultar.
 * @param {string} sheetName El nombre de la hoja donde buscar.
 * @returns {object | null} Un objeto con los datos del producto (cabeceras como claves) o null si no se encuentra.
 */
function getProductInfoGenerico(id, sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`getProductInfoGenerico: Hoja "${sheetName}" no encontrada.`);
    return null;
  }

  const rowIndex = findRowById(id, sheetName);
  if (rowIndex === -1) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const info = {};
  headers.forEach((header, index) => {
    info[header] = rowData[index];
  });

  return info;
}

/**
 * @summary Escanea la hoja de inventario para encontrar productos con más de 8 meses de antigüedad.
 * @param {string} sheetName El nombre de la hoja a escanear.
 * @returns {Array<object>} Un arreglo de objetos, donde cada objeto representa un producto que cumple el criterio.
 */
function getProductosViejos(sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const productosViejos = [];
  const hoy = new Date();

  const idIndex = headers.indexOf("Id");
  const fechaIngresoIndex = headers.indexOf("FECHA DE INGRESO");
  const productoIndex = headers.indexOf("PRODUCTO");
  const unidadesIndex = headers.indexOf("Unidades disponibles");
  const estadoIndex = headers.indexOf("Estado");

  if (
    fechaIngresoIndex === -1 ||
    productoIndex === -1 ||
    unidadesIndex === -1
  ) {
    Logger.log(
      `getProductosViejos: Columnas requeridas (FECHA DE INGRESO, PRODUCTO, Unidades disponibles) no encontradas en ${sheetName}`
    );
    return [];
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fechaIngresoValue = row[fechaIngresoIndex];
    const id = idIndex !== -1 ? row[idIndex] : null;
    const producto = row[productoIndex];
    const unidades = parseInt(row[unidadesIndex]) || 0;
    const estado = estadoIndex !== -1 ? row[estadoIndex] || "Activo" : "Activo";

    if (estado.toLowerCase() !== "activo" || unidades <= 0) {
      continue;
    }

    if (fechaIngresoValue) {
      let fechaIngreso;
      if (fechaIngresoValue instanceof Date) {
        fechaIngreso = fechaIngresoValue;
      } else {
        fechaIngreso = new Date(fechaIngresoValue);
      }

      if (isNaN(fechaIngreso.getTime())) {
        continue;
      }

      let months;
      months = (hoy.getFullYear() - fechaIngreso.getFullYear()) * 12;
      months -= fechaIngreso.getMonth();
      months += hoy.getMonth();
      if (months >= 0 && hoy.getDate() < fechaIngreso.getDate()) {
        if (
          hoy.getMonth() === fechaIngreso.getMonth() &&
          hoy.getFullYear() === fechaIngreso.getFullYear()
        ) {
        } else {
          months--;
        }
      }

      if (months >= 8) {
        const programaIndex = headers.indexOf("PROGRAMA");
        const programa = programaIndex !== -1 ? row[programaIndex] : "N/A";
        const diffTime = Math.abs(hoy.getTime() - fechaIngreso.getTime());
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        productosViejos.push({
          id: id,
          producto: producto,
          programa: programa,
          fechadeingreso: fechaIngresoValue,
          tiempoenstorage: diffDays,
          tiempoEnMeses: months < 0 ? 0 : months,
        });
      }
    }
  }
  return productosViejos;
}

/**
 * @summary Escanea una hoja para encontrar productos próximos a vencer o vencidos, con información completa.
 * @param {string} sheetName El nombre de la hoja a escanear.
 * @returns {Array<object>} Un arreglo de objetos con información detallada de los productos.
 */
function getProductosProximosAVencerCompleto(sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(
      `getProductosProximosAVencerCompleto: Hoja "${sheetName}" no encontrada.`
    );
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const productos = [];
  const fechaActual = new Date(); // 🔑 Búsqueda de índices usando los nombres EXACTOS de su lista de columnas

  const idIndex = headers.indexOf("Id");
  const productoIndex = headers.indexOf("PRODUCTO");
  const fechaVencimientoIndex = headers.indexOf("FECHA DE VENCIMIENTO");
  const ubicacionIndex = headers.indexOf("UBICACION");
  const unidadesIndex = headers.indexOf("Unidades disponibles");
  const estadoGeneralIndex = headers.indexOf("Estado"); // Usado para filtrar 'Activo' // El estado descriptivo (ESTADO DEL PRODUCTO) es opcional en este punto
  const estadoProductoIndex = headers.indexOf("Estado del Producto"); // Comprobación de índices críticos. Si faltan, retorna lista vacía.
  if (
    fechaVencimientoIndex === -1 ||
    productoIndex === -1 ||
    idIndex === -1 ||
    unidadesIndex === -1 ||
    estadoGeneralIndex === -1
  ) {
    const missingCols = [];
    if (idIndex === -1) missingCols.push("Id");
    if (productoIndex === -1) missingCols.push("PRODUCTO");
    if (fechaVencimientoIndex === -1) missingCols.push("FECHA DE VENCIMIENTO");
    if (unidadesIndex === -1) missingCols.push("Unidades disponibles");
    if (estadoGeneralIndex === -1) missingCols.push("Estado");
    Logger.log(
      `[ERROR CRÍTICO] Columnas faltantes: ${missingCols.join(
        ", "
      )} en la hoja ${sheetName}.`
    );
    return [];
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i]; // 1. FILTRO DE ESTADO: Solo procesar productos con Estado = 'Activo'
    const estadoGeneral = row[estadoGeneralIndex];
    if (String(estadoGeneral).toLowerCase() !== "activo") {
      continue;
    } // 2. CÁLCULO DE DÍAS RESTANTES

    const fechaVencimientoValue = row[fechaVencimientoIndex];
    let diasRestantes = "Sin fecha";

    if (fechaVencimientoValue && fechaVencimientoValue instanceof Date) {
      const fechaVencimiento = fechaVencimientoValue;
      const utcHoy = Date.UTC(
        fechaActual.getFullYear(),
        fechaActual.getMonth(),
        fechaActual.getDate()
      );
      const utcVencimiento = Date.UTC(
        fechaVencimiento.getFullYear(),
        fechaVencimiento.getMonth(),
        fechaVencimiento.getDate()
      );
      diasRestantes = Math.ceil(
        (utcVencimiento - utcHoy) / (1000 * 60 * 60 * 24)
      );
    } // 3. FILTRO DE VENCIMIENTO: Solo incluir productos que tengan una fecha válida (diasRestantes es un número)

    if (typeof diasRestantes === "number") {
      productos.push({
        id: row[idIndex],
        producto: row[productoIndex],
        diasRestantes: diasRestantes,
        estadoProducto:
          estadoProductoIndex !== -1 ? row[estadoProductoIndex] : "N/A",
        ubicacion: ubicacionIndex !== -1 ? row[ubicacionIndex] : "N/A",
        unidadesDisponibles: unidadesIndex !== -1 ? row[unidadesIndex] : "N/A",
      });
    }
  }

  return productos;
}

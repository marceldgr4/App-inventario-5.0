
function getProductosProximosAVencer(sheetName) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const productosProximos = [];
  const fechaActual = new Date();

  // 🔑 Columnas
  const idIndex = headers.indexOf("Id");                // Columna Id
  const productoIndex = headers.indexOf("PRODUCTO");    // Columna Producto
  const fechaVencimientoIndex = headers.indexOf("FECHA DE VENCIMIENTO");

  // ✅ Estado flexible
  let estadoIndex = headers.indexOf("Estado");
  if (estadoIndex === -1) {
    estadoIndex = headers.indexOf("ESTADO DEL PRODUCTO");
  }

  if (fechaVencimientoIndex === -1 || productoIndex === -1) {
    Logger.log(`getProductosProximosAVencer: Columnas requeridas no encontradas en ${sheetName}`);
    return [];
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = idIndex !== -1 ? row[idIndex] : i; // 👈 ahora sí tomamos el valor del Id
    const producto = row[productoIndex];
    const fechaVencimientoValue = row[fechaVencimientoIndex];
    const estadoProducto = estadoIndex !== -1 ? (row[estadoIndex] || "Activo") : "Activo";

    if (String(estadoProducto).toLowerCase() !== "activo") {
      continue;
    }

    if (fechaVencimientoValue) {
      const fechaVencimiento = fechaVencimientoValue instanceof Date
        ? fechaVencimientoValue
        : new Date(fechaVencimientoValue);

      if (isNaN(fechaVencimiento.getTime())) continue;

      const utcHoy = Date.UTC(fechaActual.getFullYear(), fechaActual.getMonth(), fechaActual.getDate());
      const utcVencimiento = Date.UTC(fechaVencimiento.getFullYear(), fechaVencimiento.getMonth(), fechaVencimiento.getDate());
      const diferenciaMs = utcVencimiento - utcHoy;
      const diasRestantes = Math.ceil(diferenciaMs / (1000 * 60 * 60 * 24));

      productosProximos.push({
        id: id || "N/A",   // 👈 ahora siempre se envía Id
        producto: producto,
        diasRestantes: diasRestantes,
        estado: estadoProducto,
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

  if (fechaIngresoIndex === -1 || productoIndex === -1 || unidadesIndex === -1) {
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
        if (hoy.getMonth() === fechaIngreso.getMonth() && hoy.getFullYear() === fechaIngreso.getFullYear()) {
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
    Logger.log(`getProductosProximosAVencerCompleto: Hoja "${sheetName}" no encontrada.`);
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toUpperCase());
  const productos = [];
  const fechaActual = new Date();

  const idIndex = headers.indexOf("ID");
  const productoIndex = headers.indexOf("PRODUCTO");
  const fechaVencimientoIndex = headers.indexOf("FECHA DE VENCIMIENTO");
  const estadoProductoIndex = headers.indexOf("ESTADO DEL PRODUCTO");
  const ubicacionIndex = headers.indexOf("UBICACION");
  const unidadesIndex = headers.indexOf("UNIDADES DISPONIBLES");
  const estadoGeneralIndex = headers.indexOf("ESTADO");

  if (fechaVencimientoIndex === -1 || productoIndex === -1 || idIndex === -1) {
    Logger.log(`Columnas requeridas (ID, PRODUCTO, FECHA DE VENCIMIENTO) no encontradas en ${sheetName}`);
    return [];
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const estadoGeneral = estadoGeneralIndex !== -1 ? row[estadoGeneralIndex] : "Activo";
    if (String(estadoGeneral).toLowerCase() !== "activo") {
      continue;
    }

    const fechaVencimientoValue = row[fechaVencimientoIndex];
    let diasRestantes = "Sin fecha";

    if (fechaVencimientoValue) {
      const fechaVencimiento = new Date(fechaVencimientoValue);
      if (!isNaN(fechaVencimiento.getTime())) {
        const utcHoy = Date.UTC(fechaActual.getFullYear(), fechaActual.getMonth(), fechaActual.getDate());
        const utcVencimiento = Date.UTC(fechaVencimiento.getFullYear(), fechaVencimiento.getMonth(), fechaVencimiento.getDate());
        diasRestantes = Math.ceil((utcVencimiento - utcHoy) / (1000 * 60 * 60 * 24));
      }
    }

    // Solo incluir productos que tienen fecha de vencimiento y están próximos a vencer o vencidos
    if (typeof diasRestantes === 'number') {
      productos.push({
        id: row[idIndex],
        producto: row[productoIndex],
        diasRestantes: diasRestantes,
        estadoProducto: estadoProductoIndex !== -1 ? row[estadoProductoIndex] : "N/A",
        ubicacion: ubicacionIndex !== -1 ? row[ubicacionIndex] : "N/A",
        unidades: unidadesIndex !== -1 ? row[unidadesIndex] : "N/A",
      });
    }
  }

  return productos;
}
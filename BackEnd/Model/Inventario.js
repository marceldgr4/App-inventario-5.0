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
    const result = retirar(id, unidades, getHojasConfig().ARTICULO.nombre);
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("retirarProductoInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

function agregarComentarioInventario(id, comentario) {
  try {
    const result = agregarComentario(id, comentario, getHojasConfig().ARTICULO.nombre);
    return JSON.stringify({ success: true, result: result });
  } catch (e) {
    Logger.log("agregarComentarioInventario: Error " + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

// Archivo: Inventario.js
function getArticulosViejos() {
  try {
    const rawData = getProductosViejos(getHojasConfig().ARTICULO.nombre); // Esta función debería devolver los datos
    
    // Normalizar los datos: convertir las claves a minúsculas
    const normalizedData = rawData.map(item => {
      const newItem = {};
      for (const key in item) {
        // Convertir la clave a minúsculas y reemplazar espacios
        const newKey = key.toLowerCase().replace(/ /g, ''); 
        newItem[newKey] = item[key];
      }
      return newItem;
    });

    Logger.log("getArticulosViejos: Data returned after normalization: " + JSON.stringify(normalizedData));
    return JSON.stringify(normalizedData);
    // La función getProductosViejos ya devuelve un array de objetos con las claves correctas.
    // No es necesario volver a procesarlo ni usar JSON.stringify.
    const productosViejos = getProductosViejos(getHojasConfig().ARTICULO.nombre);
    Logger.log(`getArticulosViejos: Devolviendo ${productosViejos.length} artículos viejos.`);
    return productosViejos;
  } catch (e) {
    Logger.log("getArticulosViejos: Error al obtener datos. " + e.message);
    return JSON.stringify([]);
    return [];
  }
}
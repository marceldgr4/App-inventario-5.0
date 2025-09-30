
// ===============================
// 📌 BackEnd/Comentario.js
// Manejo centralizado de comentarios
// ===============================

/** @summary Obtiene el objeto de la hoja de 'Comentarios'. */
function getComentarioSheet() {
  return getSheet(getHojasConfig().COMENTARIOS.nombre);
}

/** @summary Obtiene todos los datos de la hoja de 'Comentarios' en formato JSON. */
function getComentarioData() {
  return JSON.stringify({ data: _getInventoryDataForSheet(getHojasConfig().COMENTARIOS.nombre) });
}

/** @summary Obtiene la información de un comentario específico por su ID. */
function getComentarioInfo(id) {
  return getInfo(id, getHojasConfig().COMENTARIOS.nombre);
}

function agregarComentarioGenerico(sheetName, productoId, comentario) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    if (!ss) return JSON.stringify({ success: false, message: "No se pudo acceder a la hoja de cálculo principal." });
    // Buscar hoja por nombre tolerante (case-insensitive, quitar espacios)
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const target = sheetName.toString().toLowerCase().replace(/\s/g,'');
      const sheets = ss.getSheets();
      for (let s of sheets) {
        if (s.getName().toLowerCase().replace(/\s/g,'') === target) { sheet = s; break; }
      }
    }
    if (!sheet) return JSON.stringify({ success:false, message: `Hoja origen '${sheetName}' no encontrada.` });

    const data = sheet.getDataRange().getValues();
    if (!data || data.length === 0) return JSON.stringify({ success:false, message: 'Hoja sin datos.' });

    // Normalizar encabezados
    const headers = data[0].map(h => String(h || '').trim().toLowerCase());
    const findHeaderIndex = (names) => {
      for (let name of names) {
        const idx = headers.indexOf(name);
        if (idx !== -1) return idx;
      }
      // si no encontró exacto, busca si el header contiene la palabra
      for (let i = 0; i < headers.length; i++) {
        if (names.some(n => headers[i].indexOf(n) !== -1)) return i;
      }
      return -1;
    };

    const idIndex = findHeaderIndex(['id', 'identificador']);
    if (idIndex === -1) return JSON.stringify({ success:false, message: 'Columna Id no encontrada.' });

    const rowIndex = data.findIndex((r, idx) => idx > 0 && String(r[idIndex]) === String(productoId));
    if (rowIndex === -1) return JSON.stringify({ success:false, message: `Producto ${productoId} no encontrado en ${sheet.getName()}` });

    const productoIndex = findHeaderIndex(['producto', 'nombre']);
    const programaIndex = findHeaderIndex(['programa']);
    const comentariosIndex = findHeaderIndex(['comentarios', 'comentario']);

    const producto = productoIndex !== -1 ? data[rowIndex][productoIndex] : '';
    const programa = programaIndex !== -1 ? data[rowIndex][programaIndex] : '';
    const usuario = (Session && Session.getActiveUser) ? Session.getActiveUser().getEmail() : 'Anónimo';
    const fecha = new Date();

    // Guardar en hoja central de comentarios
    guardarComentarioEnSheet(productoId, producto, programa, comentario, usuario);

    // actualizar columna COMENTARIOS si existe
    if (comentariosIndex !== -1) {
      sheet.getRange(rowIndex + 1, comentariosIndex + 1).setValue(comentario);
    }

    return JSON.stringify({ success: true, message: 'Comentario guardado.' });
  } catch (err) {
    Logger.log('agregarComentarioGenerico error: ' + err);
    return JSON.stringify({ success: false, message: err.message || String(err) });
  }
}

function guardarComentarioEnSheet(productoId, producto, programa, comentario, usuario) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const hojasConfig = getHojasConfig();

    // --- 1. Guardar en la hoja central de comentarios ---
    const commentSheetName = hojasConfig.COMENTARIOS.nombre;
    const commentSheet = ss.getSheetByName(commentSheetName);
    if (!commentSheet) throw new Error('Hoja de comentarios no encontrada: ' + commentSheetName);
    
    const lastRow = commentSheet.getLastRow();
    let newId = 1;
    if (lastRow >= 1) {
      const lastId = commentSheet.getRange(lastRow, 1).getValue();
      if (!isNaN(lastId) && lastId !== '') newId = Number(lastId) + 1;
    }
    commentSheet.appendRow([newId, productoId, producto, programa, new Date(), comentario, usuario]);

    // --- 2. Actualizar la columna de comentarios en la hoja de origen ---
    const programaKey = programa.toUpperCase().replace(/\s/g, '');
    const sourceSheetConfig = Object.values(hojasConfig).find(config => config.nombre.toUpperCase().replace(/\s/g, '').includes(programaKey));
    
    if (sourceSheetConfig && sourceSheetConfig.nombre) {
      const sourceSheet = ss.getSheetByName(sourceSheetConfig.nombre);
      if (sourceSheet) {
        const data = sourceSheet.getDataRange().getValues();
        const headers = data[0].map(h => String(h || '').trim().toLowerCase());
        
        const idIndex = headers.indexOf('id');
        const comentariosIndex = headers.indexOf('comentarios');

        if (idIndex !== -1 && comentariosIndex !== -1) {
          // Buscar la fila correcta (+1 porque los datos son un array base 0, pero las filas de la hoja son base 1)
          const rowIndex = data.slice(1).findIndex(r => String(r[idIndex]) === String(productoId)) + 2;
          if (rowIndex > 1) { // si es > 1, significa que encontró la fila (el índice sería >= 0)
            sourceSheet.getRange(rowIndex, comentariosIndex + 1).setValue(comentario);
            Logger.log(`Comentario actualizado en la hoja '${sourceSheetConfig.nombre}' para el producto ID ${productoId}.`);
          } else {
            Logger.log(`WARN: No se encontró el producto con ID ${productoId} en la hoja '${sourceSheetConfig.nombre}'.`);
          }
        } else {
          Logger.log(`WARN: No se encontró la columna 'Id' o 'Comentarios' en la hoja '${sourceSheetConfig.nombre}'.`);
        }
      } else {
        Logger.log(`WARN: No se pudo encontrar la hoja de origen '${sourceSheetConfig.nombre}'.`);
      }
    } else {
      Logger.log(`WARN: No se encontró configuración de hoja para el programa '${programa}'.`);
    }

    return JSON.stringify({ success: true, message: 'Comentario guardado y actualizado.' });
  } catch (e) {
    Logger.log('guardarComentarioEnSheet error: ' + e.stack);
    return JSON.stringify({ success: false, message: e.message || String(e) });
  }
}

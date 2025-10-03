// BackEnd/GoogleDrive.js

/**
 * @summary Orquesta la adición de un producto con una imagen. Sube el archivo a Drive, obtiene la URL y luego llama a la función de agregar.
 * @param {object} productData Los datos textuales del producto.
 * @param {object} fileData La información del archivo de imagen (base64, mimeType, fileName).
 * @returns {string} El resultado de la operación de agregado en formato JSON.
 */
function agregarProductoConImagenDesdeArchivo(productData, fileData) {
  try {
    const folderId = ID_DRIVE_IMG;
    const folder = DriveApp.getFolderById(folderId);

    const bytes = Utilities.base64Decode(fileData.base64Data);
    const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.fileName);

    const file = folder.createFile(blob);
    const imageUrl = file.getUrl();

    productData.imagenAgregar = imageUrl;

    const result = agregarProductoGenerico(productData, HOJA_ARTICULOS);

    return result;
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

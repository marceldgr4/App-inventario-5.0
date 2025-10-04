// =================================================================
// --- SECCIÓN DE CONSTANTES ---
// =================================================================

const URL_FAVICON =
  "https://www.intouchcx.com/wp-content/themes/intouchcx/assets/favicon/favicon-16x16.png";
//ID de la pagina de la hoja de cálculo de Google Sheets que funciona como base de datos principal.
const ID_INVENTARIO = "1lWsyJLZTOZDbIeAcagvj7wCPIp9vD_METhdkiHQBzEU";
//ID la ubicacion Drive donde se almacenarán los archivos subidos.
const ID_DRIVE_PDF = "1yIl44eNcBWQVxV8CjoI84bBqdfpIs1LB"; // Carpeta actas escaneados
const ID_PAPELERIA_IMG = "1E8H7vg2eWWmHzSZCGPIEqicc1EqTie33";
const ID_DRIVE_IMG = "17cHe5AQwuWlwXbNClpSN1FW7vZGJK4eN";
const ID_DECORACION_IMG = "1mxaXcJuKdbSOD-NW9yugZf-iGrhc0jed";

//Nombres exactos de las hojas (pestañas) dentro del Google Sheet.
const HOJA_USUARIO = "Usuario";
const HOJA_ARTICULOS = "Inventario";
const HOJA_HISTORIAL = "Historial_Modificaciones";
const HOJA_COMENTARIOS = "Comentarios";
const HOJA_DECORACION = "Inventario decoracion";
const HOJA_COMIDA = "Inventario comida";
const HOJA_PAPELERIA = "Papeleria";
const HOJA_REGISTRO_USUARIO = "Registro de Inicio de Sesion";
const HOJA_ACTA = "Acta";

/**
 * @summary Objeto de configuración de hojas
 */
function getHojasConfig() {
  return {
    ARTICULO: { nombre: HOJA_ARTICULOS },
    DECORACION: { nombre: HOJA_DECORACION },
    COMIDA: { nombre: HOJA_COMIDA },
    PAPELERIA: { nombre: HOJA_PAPELERIA },
    HISTORIAL: { nombre: HOJA_HISTORIAL },
    COMENTARIOS: { nombre: HOJA_COMENTARIOS },
    USUARIO: { nombre: HOJA_USUARIO },
    REGISTRO: { nombre: HOJA_REGISTRO_USUARIO },
    ACTA: { nombre: HOJA_ACTA },
  };
}

const PAGES_PERMITIDAS = [
  "Home",
  "Articulos",
  "Index",
  "Acta",
  "Comida",
  "Decoracion",
  "Papeleria",
  "Comentario",
  "Historial",
  "Login",
  "Dashboard",
  "Registro",
  "Usuario",
  "Perfil",
];

// =================================================================
// --- CONSTANTES PARA CONTROL DE ACCESO POR ROLES (RBAC) ---
// =================================================================

// Define los nombres de los roles de usuario.
const ROL_ADMIN = "Admin";
const ROL_USUARIO = "Usuario";

// Define las páginas permitidas para cada rol.
const PAGINAS_POR_ROL = {
  [ROL_ADMIN]: [
    "Home",
    "Articulos",
    "Acta",
    "Comida",
    "Decoracion",
    "Papeleria",
    "Comentario",
    "Historial",
    "Login",
    "Dashboard",
    "Registro",
    "Usuario",
    "Perfil",
  ],
  [ROL_USUARIO]: [
    "Articulos",
    "Comida",
    "Decoracion",
    "Papeleria",
    "Home",
    "Acta",
    "Perfil",
    "Dashboard",
  ],
};

// Clave de PropertiesService
const CLAVE_PROPIEDAD_USUARIO = "USUARIO_ACTIVO";
//=======================================
/** ====== Utilidades Genéricas ====== */
//=======================================

function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(ID_INVENTARIO);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`DEBUG: getSheetByName - Hoja '${sheetName}' NO encontrada.`);
    }
    return sheet;
  } catch (e) {
    Logger.log(
      `ERROR: getSheetByName - No se pudo abrir Spreadsheet con ID: ${ID_INVENTARIO} o la hoja '${sheetName}'. Error: ${e.message}`
    );
    return null;
  }
}

function getData(sheetName) {
  switch (sheetName) {
    case HOJA_HISTORIAL:
      return JSON.stringify({ data: getHistorialDataActual() });
    case HOJA_REGISTRO_USUARIO:
      return JSON.stringify({ data: getRegistroData() });
    default:
      return JSON.stringify({ data: _getInventoryDataForSheet(sheetName) });
  }
}

// Wrappers genéricos
function getInfo(id, sheetName) {
  return getProductInfoGenerico(id, sheetName);
}
function agregar(data, sheetName) {
  return agregarProductoGenerico(data, sheetName);
}
function actualizar(data, sheetName) {
  return actualizarProductoGenerico(data, sheetName);
}
function eliminar(id, sheetName) {
  return eliminarProductoGenerico(id, sheetName);
}
function retirar(id, unidades, sheetName) {
  return retirarProductoGenerico(id, unidades, sheetName);
}
function agregarComentario(id, comentario, sheetName) {
  // Wrapper para compatibilidad con llamadas antiguas. Construye el objeto de datos esperado.
  const data = {
    productoId: id,
    comentario: comentario,
    sheetName: sheetName
  };
  return registrarNuevoComentario(data); // Llama a la nueva función unificada.
}

function _getInventoryDataForSheet(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map((h) => h.toString().trim());
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every((cell) => cell === "")) continue;

    const item = {};
    for (let j = 0; j < headers.length; j++) {
      item[headers[j]] = row[j];
    }
    items.push(item);
  }
  Logger.log("_getInventoryDataForSheet: Filas cargadas -> " + items.length);
  return items;
}
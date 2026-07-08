// ============================================================
//  Marcaje Casino Melinka · AquaChile
//  Google Apps Script – Web App backend
//
//  INSTRUCCIONES DE DESPLIEGUE:
//  1. Abre Google Sheets con la planilla de marcajes.
//  2. Menú: Extensiones → Apps Script.
//  3. Pega este código en Code.gs y guarda (Ctrl+S).
//  4. Haz clic en "Implementar" → "Nueva implementación".
//  5. Tipo: Aplicación web.
//     · Ejecutar como: Yo (tu cuenta Google).
//     · Quién tiene acceso: Cualquier persona (anónimo).
//  6. Copia la URL que aparece en "URL de la aplicación web".
//  7. Pégala en index.html → CONFIG.GAS_ENDPOINT.
// ============================================================

const SHEET_NAME = null; // null = primera hoja
const HEADER = ["fecha","hora","tipo","rut","nombre","centroBase","servicioExterno","centroCasino","timestamp"];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet   = getSheet();
    ensureHeader(sheet);
    if (isDuplicate(sheet, payload)) {
      return jsonResponse({ status: "duplicado" });
    }
    sheet.appendRow([
      payload.fecha          || "",
      payload.hora           || "",
      payload.tipo           || "",
      payload.rut            || "",
      payload.nombre         || "",
      payload.centroBase     || "",
      payload.servicioExterno ? "SI" : "NO",
      payload.centroCasino   || "",
      payload.timestamp      || new Date().toISOString()
    ]);
    return jsonResponse({ status: "ok" });
  } catch (err) {
    return jsonResponse({ status: "error", mensaje: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ status: "ok", servicio: "Marcaje Casino Melinka" });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADER);
}

function isDuplicate(sheet, payload) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const data = sheet.getRange(2, 1, last - 1, HEADER.length).getValues();
  for (const row of data) {
    const mismaFecha  = row[0] === payload.fecha;
    const mismoTipo   = row[2] === payload.tipo;
    const mismoRut    = payload.rut && row[3] === payload.rut;
    const mismoNombre = !payload.rut && row[4] === payload.nombre;
    if (mismaFecha && mismoTipo && (mismoRut || mismoNombre)) return true;
  }
  return false;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

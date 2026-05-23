const RUNTIME_CACHE = {
  configMap: null,
  globalProtocolItems: null,
  rolloutProtocolItems: {},
  sheetSnapshots: {}
};

function getSheetSnapshot(sheetName, options) {
  const key = sheetName + '::' + (options && options.ensureFn ? options.ensureFn.name : 'raw');
  if (RUNTIME_CACHE.sheetSnapshots[key]) return RUNTIME_CACHE.sheetSnapshots[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) {
    const empty = { sheet: sheet, headers: [], data: [] };
    RUNTIME_CACHE.sheetSnapshots[key] = empty;
    return empty;
  }
  let headers = null;
  if (options && options.ensureFn) {
    headers = options.ensureFn(sheet);
  } else {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
  }
  const rowCount = sheet.getLastRow();
  const colCount = headers.length || sheet.getLastColumn();
  const data = rowCount > 0 ? sheet.getRange(1, 1, rowCount, colCount).getValues() : [];
  const snapshot = { sheet: sheet, headers: headers, data: data };
  RUNTIME_CACHE.sheetSnapshots[key] = snapshot;
  return snapshot;
}
